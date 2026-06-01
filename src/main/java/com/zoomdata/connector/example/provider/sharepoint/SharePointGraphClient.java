/**
 * Copyright (C) insightsoftware 2026. All rights reserved.
 *
 * GraphServiceClient factory with per-connection caching.
 *
 * Builds at most one client per (tenantId, clientId, secret-hash, authority)
 * tuple for the pod's lifetime. Reuse matters for two reasons:
 *
 *  1. The MSAL4J token cache lives inside the ClientSecretCredential.
 *     A fresh credential = fresh cache = an OAuth round-trip on every
 *     Composer call. Caching the credential keeps tokens warm.
 *
 *  2. GraphServiceClient wraps an OkHttp client with a thread pool and
 *     connection pool. Creating one per call leaked file descriptors and
 *     threads under sustained load.
 *
 * Rotating CLIENT_SECRET requires a pod restart — the cache is keyed on a
 * hash of the secret, so a new value yields a new client, but the old one
 * stays cached. For a v1 EDC where one Composer connection ≈ one tenant
 * the cache rarely grows past 1-2 entries; we log a warning past 32 in
 * case of mass rotation or misconfiguration.
 */
package com.zoomdata.connector.example.provider.sharepoint;

import com.azure.identity.ClientSecretCredential;
import com.azure.identity.ClientSecretCredentialBuilder;
import com.microsoft.graph.serviceclient.GraphServiceClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.net.URI;
import java.net.URISyntaxException;
import java.util.Objects;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ConcurrentMap;

public class SharePointGraphClient {

    private static final Logger log = LoggerFactory.getLogger(SharePointGraphClient.class);

    private static final String DEFAULT_AUTHORITY = "https://login.microsoftonline.com";
    private static final String[] DEFAULT_SCOPES = new String[]{"https://graph.microsoft.com/.default"};
    private static final int CACHE_SIZE_WARN_THRESHOLD = 32;

    private final ConcurrentMap<String, GraphServiceClient> cache = new ConcurrentHashMap<>();

    public GraphServiceClient build(String tenantId, String clientId, String clientSecret, String authority) {
        if (tenantId == null || tenantId.isEmpty()) {
            throw new IllegalArgumentException("TENANT_ID is required");
        }
        if (clientId == null || clientId.isEmpty()) {
            throw new IllegalArgumentException("CLIENT_ID is required");
        }
        if (clientSecret == null || clientSecret.isEmpty()) {
            throw new IllegalArgumentException("CLIENT_SECRET is required");
        }

        String authorityHost = (authority != null && !authority.isEmpty()) ? authority : DEFAULT_AUTHORITY;
        String key = cacheKey(tenantId, clientId, clientSecret, authorityHost);

        return cache.computeIfAbsent(key, k -> {
            int size = cache.size();
            if (size >= CACHE_SIZE_WARN_THRESHOLD) {
                log.warn("Graph client cache has {} entries (tenant={} clientId={}). "
                        + "Mass secret rotation or many distinct tenants? Restart the pod to flush.",
                        size, tenantId, clientId);
            } else {
                log.info("Building new GraphServiceClient (tenant={}, clientId={}, cache size now {})",
                        tenantId, clientId, size + 1);
            }
            ClientSecretCredential credential = new ClientSecretCredentialBuilder()
                    .clientId(clientId)
                    .tenantId(tenantId)
                    .clientSecret(clientSecret)
                    .authorityHost(authorityHost)
                    .build();
            return new GraphServiceClient(credential, DEFAULT_SCOPES);
        });
    }

    /**
     * Cache key includes a hash of the secret so a rotated secret produces a
     * different key and triggers a rebuild. The secret itself is never
     * stored in the key — only its hash, alongside the (already public)
     * tenant + client + authority.
     */
    private String cacheKey(String tenantId, String clientId, String secret, String authority) {
        return tenantId + "|" + clientId + "|" + authority
                + "|" + Integer.toHexString(Objects.hash(tenantId, clientId, secret, authority));
    }

    /** Test/diagnostic helper. */
    int cacheSize() {
        return cache.size();
    }

    /**
     * Convert a SharePoint site URL like
     *   https://contoso.sharepoint.com/sites/sales
     * into Graph's path-based site identifier:
     *   contoso.sharepoint.com:/sites/sales:
     *
     * Validates that the host looks like a SharePoint Online hostname so
     * misconfigured connections fail fast with a clear error rather than
     * a cryptic Graph response.
     */
    public String siteIdFromUrl(String siteUrl) {
        if (siteUrl == null || siteUrl.isEmpty()) {
            throw new IllegalArgumentException("SITE_URL is required");
        }
        try {
            URI uri = new URI(siteUrl);
            String host = uri.getHost();
            String path = uri.getPath();
            if (host == null) {
                throw new IllegalArgumentException("SITE_URL must be a full https URL with a hostname");
            }
            if (!host.toLowerCase().endsWith(".sharepoint.com")) {
                throw new IllegalArgumentException("SITE_URL host must be a *.sharepoint.com domain: " + host);
            }
            // Normalise: trim trailing slash; null path becomes root.
            if (path != null && path.endsWith("/") && path.length() > 1) {
                path = path.substring(0, path.length() - 1);
            }
            if (path == null || path.isEmpty() || "/".equals(path)) {
                return host;
            }
            return host + ":" + path + ":";
        } catch (URISyntaxException e) {
            throw new IllegalArgumentException("Invalid SITE_URL: " + siteUrl, e);
        }
    }
}
