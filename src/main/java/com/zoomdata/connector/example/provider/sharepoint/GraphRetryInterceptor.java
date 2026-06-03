/**
 * Copyright (C) insightsoftware 2026. All rights reserved.
 *
 * OkHttp interceptor that retries throttled / transient Graph responses on the
 * raw-REST reader path (Workbook + file readers). The typed Graph SDK path
 * already has Kiota's retry handler; this gives the raw OkHttp calls the same
 * protection: retry on 429 / 503 / 504, honour the Retry-After header when
 * present, otherwise exponential backoff. Logs each retry (logs-only
 * telemetry — no metrics endpoint by design).
 */
package com.zoomdata.connector.example.provider.sharepoint;

import okhttp3.Interceptor;
import okhttp3.Response;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;

public class GraphRetryInterceptor implements Interceptor {

    private static final Logger log = LoggerFactory.getLogger(GraphRetryInterceptor.class);
    private static final long MAX_BACKOFF_MS = 30_000L;

    private final int maxRetries;

    public GraphRetryInterceptor(int maxRetries) {
        this.maxRetries = Math.max(0, maxRetries);
    }

    @Override
    public Response intercept(Chain chain) throws IOException {
        Response response = chain.proceed(chain.request());
        int attempt = 0;
        while (isRetryable(response.code()) && attempt < maxRetries) {
            long waitMs = backoffMillis(response, attempt);
            int code = response.code();
            response.close(); // release the connection before sleeping/retrying
            try {
                Thread.sleep(waitMs);
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                break;
            }
            attempt++;
            log.info("Graph HTTP {} — retry {}/{} after {}ms", code, attempt, maxRetries, waitMs);
            response = chain.proceed(chain.request());
        }
        return response;
    }

    private static boolean isRetryable(int code) {
        return code == 429 || code == 503 || code == 504;
    }

    /**
     * Honour Retry-After (delta seconds, or an HTTP-date) when present;
     * otherwise exponential backoff 1s, 2s, 4s ... capped at 30s.
     */
    private long backoffMillis(Response response, int attempt) {
        String retryAfter = response.header("Retry-After");
        if (retryAfter != null && !retryAfter.isEmpty()) {
            try {
                long secs = Long.parseLong(retryAfter.trim());
                return Math.min(secs * 1000L, MAX_BACKOFF_MS);
            } catch (NumberFormatException notSeconds) {
                try {
                    long epochMs = java.time.ZonedDateTime
                            .parse(retryAfter, java.time.format.DateTimeFormatter.RFC_1123_DATE_TIME)
                            .toInstant().toEpochMilli();
                    long delta = epochMs - System.currentTimeMillis();
                    if (delta > 0) return Math.min(delta, MAX_BACKOFF_MS);
                } catch (Exception ignore) {
                    // fall through to exponential backoff
                }
            }
        }
        long exp = 1000L * (1L << attempt); // 1s, 2s, 4s, ...
        return Math.min(exp, MAX_BACKOFF_MS);
    }
}
