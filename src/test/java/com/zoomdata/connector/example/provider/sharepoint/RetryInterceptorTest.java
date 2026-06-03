/**
 * Copyright (C) insightsoftware 2026. All rights reserved.
 *
 * Deterministic test for GraphRetryInterceptor using a local JDK HttpServer
 * (no extra test dependency). Proves: a 429 with Retry-After is retried and
 * the eventual 200 is returned; a persistent 500 is NOT retried; the
 * Retry-After delay is honoured.
 *
 * Run:
 *   mvn -q test-compile exec:java \
 *     -Dexec.mainClass=com.zoomdata.connector.example.provider.sharepoint.RetryInterceptorTest \
 *     -Dexec.classpathScope=test
 */
package com.zoomdata.connector.example.provider.sharepoint;

import com.sun.net.httpserver.HttpServer;
import okhttp3.OkHttpClient;
import okhttp3.Request;
import okhttp3.Response;

import java.net.InetSocketAddress;
import java.util.concurrent.atomic.AtomicInteger;

public class RetryInterceptorTest {

    public static void main(String[] args) throws Exception {
        boolean ok = true;
        ok &= retriesThenSucceeds();
        ok &= doesNotRetry500();
        System.out.println(ok ? "\n=== Retry interceptor test passed ===" : "\n=== Retry interceptor test FAILED ===");
        if (!ok) System.exit(1);
    }

    private static boolean retriesThenSucceeds() throws Exception {
        AtomicInteger hits = new AtomicInteger(0);
        HttpServer server = HttpServer.create(new InetSocketAddress("127.0.0.1", 0), 0);
        server.createContext("/x", ex -> {
            int n = hits.incrementAndGet();
            byte[] body;
            if (n == 1) { // first hit: throttle for 1s
                ex.getResponseHeaders().add("Retry-After", "1");
                body = "throttled".getBytes();
                ex.sendResponseHeaders(429, body.length);
            } else {       // second hit: success
                body = "ok".getBytes();
                ex.sendResponseHeaders(200, body.length);
            }
            ex.getResponseBody().write(body);
            ex.close();
        });
        server.start();
        int port = server.getAddress().getPort();
        OkHttpClient c = new OkHttpClient.Builder().addInterceptor(new GraphRetryInterceptor(3)).build();
        long t0 = System.currentTimeMillis();
        try (Response r = c.newCall(new Request.Builder().url("http://127.0.0.1:" + port + "/x").build()).execute()) {
            long elapsed = System.currentTimeMillis() - t0;
            boolean pass = r.code() == 200 && hits.get() == 2 && elapsed >= 900;
            System.out.printf("A. 429+Retry-After then 200: final=%d hits=%d elapsed=%dms -> %s%n",
                    r.code(), hits.get(), elapsed, pass ? "PASS" : "FAIL");
            server.stop(0);
            return pass;
        }
    }

    private static boolean doesNotRetry500() throws Exception {
        AtomicInteger hits = new AtomicInteger(0);
        HttpServer server = HttpServer.create(new InetSocketAddress("127.0.0.1", 0), 0);
        server.createContext("/y", ex -> {
            hits.incrementAndGet();
            byte[] body = "err".getBytes();
            ex.sendResponseHeaders(500, body.length); // 500 is not in {429,503,504}
            ex.getResponseBody().write(body);
            ex.close();
        });
        server.start();
        int port = server.getAddress().getPort();
        OkHttpClient c = new OkHttpClient.Builder().addInterceptor(new GraphRetryInterceptor(3)).build();
        try (Response r = c.newCall(new Request.Builder().url("http://127.0.0.1:" + port + "/y").build()).execute()) {
            boolean pass = r.code() == 500 && hits.get() == 1; // not retried
            System.out.printf("B. persistent 500 not retried: final=%d hits=%d -> %s%n",
                    r.code(), hits.get(), pass ? "PASS" : "FAIL");
            server.stop(0);
            return pass;
        }
    }
}
