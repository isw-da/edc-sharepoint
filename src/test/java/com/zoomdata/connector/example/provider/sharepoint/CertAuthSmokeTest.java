/**
 * Copyright (C) insightsoftware 2026. All rights reserved.
 *
 * Validates AUTH_MODE=CLIENT_CERTIFICATE end to end: builds a GraphAuth with
 * a PEM (private key + cert), acquires a token via the certificate credential,
 * and makes a real Graph call (/sites/root) with it.
 *
 * Run:
 *   mvn -q test-compile exec:java \
 *     -Dexec.mainClass=com.zoomdata.connector.example.provider.sharepoint.CertAuthSmokeTest \
 *     -Dexec.classpathScope=test
 *
 * Env: SP_TENANT_ID, SP_CLIENT_ID, SP_CERT_PEM_FILE (path to key+cert PEM)
 */
package com.zoomdata.connector.example.provider.sharepoint;

import java.nio.file.Files;
import java.nio.file.Paths;

public class CertAuthSmokeTest {

    public static void main(String[] args) throws Exception {
        String tenantId = env("SP_TENANT_ID");
        String clientId = env("SP_CLIENT_ID");
        String pemFile = env("SP_CERT_PEM_FILE");
        String pem = new String(Files.readAllBytes(Paths.get(pemFile)));

        System.out.println("=== Certificate-auth smoke test (AUTH_MODE=CLIENT_CERTIFICATE) ===");
        System.out.println("Tenant: " + tenantId + "  Client: " + clientId);
        System.out.println("PEM:    " + pemFile + " (" + pem.length() + " chars)");
        System.out.println();

        SharePointGraphClient gf = new SharePointGraphClient();
        SharePointGraphClient.GraphAuth auth = new SharePointGraphClient.GraphAuth(
                tenantId, clientId, SharePointGraphClient.AuthMode.CLIENT_CERTIFICATE,
                null, pem, null);

        System.out.println("--- 1. Acquire token via certificate credential ---");
        String token = gf.bearerToken(auth);
        System.out.println("token acquired: " + (token != null && token.length() > 20)
                + " (" + (token == null ? 0 : token.length()) + " chars)");
        if (token == null || token.isEmpty()) { System.err.println("FAIL: no token"); System.exit(2); }

        System.out.println();
        System.out.println("--- 2. Real Graph call with the cert-derived token ---");
        boolean ok = gf.ping("https://graph.microsoft.com/v1.0/sites/root", token);
        System.out.println("GET /sites/root with cert token -> 2xx: " + ok);
        if (!ok) { System.err.println("FAIL: Graph call rejected the cert token"); System.exit(3); }

        System.out.println();
        System.out.println("=== Certificate-auth smoke test passed ===");
    }

    private static String env(String name) {
        String v = System.getenv(name);
        if (v == null || v.isEmpty()) { System.err.println("Missing env var: " + name); System.exit(99); }
        return v;
    }
}
