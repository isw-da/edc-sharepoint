/**
 * Copyright (C) insightsoftware 2026. All rights reserved.
 *
 * Thin wrapper around GraphServiceClient that produces an authenticated
 * client per connection. Caches nothing: token refresh is handled by
 * azure-identity's ClientSecretCredential, so a fresh GraphServiceClient
 * per request is cheap and avoids cross-tenant state.
 */
package com.zoomdata.connector.example.provider.sharepoint;

import com.azure.identity.ClientSecretCredential;
import com.azure.identity.ClientSecretCredentialBuilder;
import com.microsoft.graph.serviceclient.GraphServiceClient;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.net.URI;
import java.net.URISyntaxException;

public class SharePointGraphClient {

    private static final Logger log = LoggerFactory.getLogger(SharePointGraphClient.class);

    private static final String DEFAULT_AUTHORITY = "https://login.microsoftonline.com";
    private static final String[] DEFAULT_SCOPES = new String[]{"https://graph.microsoft.com/.default"};

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

        ClientSecretCredential credential = new ClientSecretCredentialBuilder()
                .clientId(clientId)
                .tenantId(tenantId)
                .clientSecret(clientSecret)
                .authorityHost(authorityHost)
                .build();

        return new GraphServiceClient(credential, DEFAULT_SCOPES);
    }

    /**
     * Convert a SharePoint site URL like
     *   https://contoso.sharepoint.com/sites/sales
     * into Graph's path-based site identifier:
     *   contoso.sharepoint.com:/sites/sales:
     *
     * This identifier can be passed directly to bySiteId(...) and Graph
     * resolves it to the canonical site ID server-side.
     */
    public String siteIdFromUrl(String siteUrl) {
        if (siteUrl == null || siteUrl.isEmpty()) {
            throw new IllegalArgumentException("SITE_URL is required");
        }
        try {
            URI uri = new URI(siteUrl);
            String host = uri.getHost();
            String path = uri.getPath();
            if (host == null || path == null) {
                throw new IllegalArgumentException("SITE_URL must be a full https URL");
            }
            // Trim trailing slash on path; keep leading slash
            if (path.endsWith("/") && path.length() > 1) {
                path = path.substring(0, path.length() - 1);
            }
            if (path.isEmpty() || "/".equals(path)) {
                // Root site of the tenant
                return host;
            }
            return host + ":" + path + ":";
        } catch (URISyntaxException e) {
            throw new IllegalArgumentException("Invalid SITE_URL: " + siteUrl, e);
        }
    }
}
