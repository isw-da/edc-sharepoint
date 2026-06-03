/**
 * Copyright (C) insightsoftware 2026. All rights reserved.
 *
 * GraphServiceClient factory with per-connection caching + configurable
 * HTTP timeouts.
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
 *
 * v1.1 (M1 audit fix): connect/read timeouts are now read from
 * framework.properties (sharepoint.connection.timeout.sec /
 * sharepoint.read.timeout.sec). Previously declared but unread.
 */
package com.zoomdata.connector.example.provider.sharepoint;

import com.azure.core.credential.TokenRequestContext;
import com.azure.identity.ClientSecretCredential;
import com.azure.identity.ClientSecretCredentialBuilder;
import com.microsoft.graph.core.authentication.AzureIdentityAccessTokenProvider;
import com.microsoft.graph.serviceclient.GraphServiceClient;
import com.microsoft.kiota.authentication.BaseBearerTokenAuthenticationProvider;
import com.microsoft.kiota.http.KiotaClientFactory;
import okhttp3.OkHttpClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.net.URI;
import java.net.URISyntaxException;
import java.time.Duration;
import java.util.Objects;
import java.util.Properties;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ConcurrentMap;

public class SharePointGraphClient {

    private static final Logger log = LoggerFactory.getLogger(SharePointGraphClient.class);

    private static final String DEFAULT_AUTHORITY = "https://login.microsoftonline.com";
    private static final String GRAPH_SCOPE = "https://graph.microsoft.com/.default";
    private static final String[] DEFAULT_SCOPES = new String[]{GRAPH_SCOPE};
    private static final int CACHE_SIZE_WARN_THRESHOLD = 32;

    // Hardcoded defaults that match framework.properties baked-in values.
    // Used when the properties file is missing or unreadable.
    private static final long DEFAULT_CONNECT_TIMEOUT_SEC = 30L;
    private static final long DEFAULT_READ_TIMEOUT_SEC = 60L;

    private final ConcurrentMap<String, GraphServiceClient> cache = new ConcurrentHashMap<>();

    // Credentials cached separately, keyed identically, so the raw-REST
    // Workbook path (bearerToken) reuses the SAME ClientSecretCredential —
    // and therefore the same MSAL token cache — as the typed SDK path. A
    // separate credential would mean a second OAuth round-trip per connection.
    private final ConcurrentMap<String, ClientSecretCredential> credentialCache = new ConcurrentHashMap<>();

    // Resolved once at instance construction from framework.properties.
    private final long connectTimeoutSec;
    private final long readTimeoutSec;

    // Shared OkHttp client for raw Workbook REST calls (the Kiota typed layer
    // can't read 2D workbook value arrays — see SharePointWorkbookReader).
    // Plain client with our configured timeouts; bearer header added per call.
    private final OkHttpClient rawHttpClient;

    public SharePointGraphClient() {
        Properties props = loadFrameworkProperties();
        this.connectTimeoutSec = parseLong(props.getProperty("sharepoint.connection.timeout.sec"),
                DEFAULT_CONNECT_TIMEOUT_SEC);
        this.readTimeoutSec = parseLong(props.getProperty("sharepoint.read.timeout.sec"),
                DEFAULT_READ_TIMEOUT_SEC);
        this.rawHttpClient = buildRawHttpClient(connectTimeoutSec, readTimeoutSec);
        log.info("SharePoint Graph client configured: connectTimeout={}s, readTimeout={}s",
                connectTimeoutSec, readTimeoutSec);
    }

    /**
     * Test-only constructor that bypasses the framework.properties load so
     * unit tests can pin timeouts deterministically.
     */
    SharePointGraphClient(long connectTimeoutSec, long readTimeoutSec) {
        this.connectTimeoutSec = connectTimeoutSec;
        this.readTimeoutSec = readTimeoutSec;
        this.rawHttpClient = buildRawHttpClient(connectTimeoutSec, readTimeoutSec);
    }

    private static OkHttpClient buildRawHttpClient(long connectSec, long readSec) {
        return new OkHttpClient.Builder()
                .connectTimeout(Duration.ofSeconds(connectSec))
                .readTimeout(Duration.ofSeconds(readSec))
                .writeTimeout(Duration.ofSeconds(readSec))
                .build();
    }

    /** Shared OkHttp client for raw Graph REST (Workbook reads). */
    public OkHttpClient rawHttpClient() {
        return rawHttpClient;
    }

    /** GET a Graph URL with the given bearer token; true on a 2xx response.
     *  Used to validate a drive container (e.g. /users/{upn}/drive) reachably
     *  without the typed SDK. */
    public boolean ping(String graphUrl, String bearerToken) {
        okhttp3.Request req = new okhttp3.Request.Builder().url(graphUrl)
                .header("Authorization", "Bearer " + bearerToken)
                .header("Accept", "application/json").get().build();
        try (okhttp3.Response resp = rawHttpClient.newCall(req).execute()) {
            return resp.code() >= 200 && resp.code() < 300;
        } catch (Exception e) {
            log.warn("Graph ping failed for {}: {}", graphUrl, e.getMessage());
            return false;
        }
    }

    /**
     * Acquire a Graph bearer token for the given credentials, reusing the
     * cached ClientSecretCredential (and its MSAL token cache). Synchronous —
     * azure-identity caches and refreshes the token internally, so repeated
     * calls within a token's lifetime do not hit the network.
     */
    public String bearerToken(String tenantId, String clientId, String clientSecret, String authority) {
        if (tenantId == null || tenantId.isEmpty()) throw new IllegalArgumentException("TENANT_ID is required");
        if (clientId == null || clientId.isEmpty()) throw new IllegalArgumentException("CLIENT_ID is required");
        if (clientSecret == null || clientSecret.isEmpty()) throw new IllegalArgumentException("CLIENT_SECRET is required");
        String authorityHost = (authority != null && !authority.isEmpty()) ? authority : DEFAULT_AUTHORITY;
        ClientSecretCredential cred = credential(tenantId, clientId, clientSecret, authorityHost);
        return cred.getTokenSync(new TokenRequestContext().addScopes(GRAPH_SCOPE)).getToken();
    }

    /** Get-or-build a ClientSecretCredential, cached by the same key as the SDK client. */
    private ClientSecretCredential credential(String tenantId, String clientId, String clientSecret, String authorityHost) {
        String key = cacheKey(tenantId, clientId, clientSecret, authorityHost);
        return credentialCache.computeIfAbsent(key, k -> new ClientSecretCredentialBuilder()
                .clientId(clientId)
                .tenantId(tenantId)
                .clientSecret(clientSecret)
                .authorityHost(authorityHost)
                .build());
    }

    private Properties loadFrameworkProperties() {
        Properties p = new Properties();
        try (InputStream in = Thread.currentThread().getContextClassLoader()
                .getResourceAsStream("framework.properties")) {
            if (in != null) {
                p.load(in);
            } else {
                log.warn("framework.properties not found on classpath; using built-in timeout defaults");
            }
        } catch (IOException e) {
            log.warn("Failed to read framework.properties ({}); using built-in defaults", e.getMessage());
        }
        return p;
    }

    private static long parseLong(String value, long fallback) {
        if (value == null || value.trim().isEmpty()) return fallback;
        try {
            return Long.parseLong(value.trim());
        } catch (NumberFormatException e) {
            return fallback;
        }
    }

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
            ClientSecretCredential credential = credential(tenantId, clientId, clientSecret, authorityHost);

            // v1.1: build the OkHttp client ourselves so we can apply our
            // configured timeouts. KiotaClientFactory.create() returns a
            // builder pre-populated with the standard Kiota interceptor
            // chain (auth, retry, redirect, parameter sanitisation, etc.)
            // — we only override timeouts on top.
            OkHttpClient httpClient = KiotaClientFactory.create()
                    .connectTimeout(Duration.ofSeconds(connectTimeoutSec))
                    .readTimeout(Duration.ofSeconds(readTimeoutSec))
                    .writeTimeout(Duration.ofSeconds(readTimeoutSec))
                    .build();

            // 1-arg constructor uses default Graph scopes (.default) and
            // default observability options — exactly what we want.
            AzureIdentityAccessTokenProvider tokenProvider =
                    new AzureIdentityAccessTokenProvider(credential);
            BaseBearerTokenAuthenticationProvider authProvider =
                    new BaseBearerTokenAuthenticationProvider(tokenProvider);

            return new GraphServiceClient(authProvider, httpClient);
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

    /** Test/diagnostic helper for the resolved timeouts. */
    long connectTimeoutSec() { return connectTimeoutSec; }
    long readTimeoutSec() { return readTimeoutSec; }

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

    /**
     * Human-readable slug for a site, used to prefix collection names in
     * multi-site connections. Last non-empty path segment (e.g. "sales" for
     * .../sites/sales); falls back to the host's first DNS label for a root
     * site. Lowercased, non-alphanumeric runs collapsed to a single "_", so
     * the slug never contains "__" (the collection-name delimiter).
     */
    public String siteSlugFromUrl(String siteUrl) {
        if (siteUrl == null || siteUrl.isEmpty()) {
            throw new IllegalArgumentException("site URL is required");
        }
        try {
            URI uri = new URI(siteUrl);
            String host = uri.getHost();
            String path = uri.getPath();
            String raw = null;
            if (path != null) {
                String[] segs = path.split("/");
                for (int i = segs.length - 1; i >= 0; i--) {
                    if (!segs[i].isEmpty()) { raw = segs[i]; break; }
                }
            }
            if (raw == null) {
                raw = (host != null && host.contains(".")) ? host.substring(0, host.indexOf('.')) : host;
            }
            if (raw == null || raw.isEmpty()) raw = "site";
            return raw.replaceAll("[^A-Za-z0-9]+", "_").toLowerCase();
        } catch (URISyntaxException e) {
            throw new IllegalArgumentException("Invalid site URL: " + siteUrl, e);
        }
    }
}
