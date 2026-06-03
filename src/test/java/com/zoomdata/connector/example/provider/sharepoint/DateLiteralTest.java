/**
 * Copyright (C) insightsoftware 2026. All rights reserved.
 *
 * Unit test for SharePointComputeTask.odataDateLiteral — the v1.1 date filter
 * pushdown hardening. Confirms epoch-millis -> unquoted ISO 8601 UTC, and
 * ISO/quoted-ISO inputs pass through unquoted (so Graph accepts the dateTime
 * comparison instead of a rejected quoted string).
 *
 * Run:
 *   mvn -q test-compile exec:java \
 *     -Dexec.mainClass=com.zoomdata.connector.example.provider.sharepoint.DateLiteralTest \
 *     -Dexec.classpathScope=test
 */
package com.zoomdata.connector.example.provider.sharepoint;

public class DateLiteralTest {

    public static void main(String[] args) {
        boolean ok = true;

        // 2026-01-01T00:00:00Z == 1767225600000 ms
        ok &= check("epoch millis -> ISO UTC", "1767225600000", "2026-01-01T00:00:00Z");
        // a non-midnight instant: 2026-07-15T09:00:00Z == 1784106000000
        ok &= check("epoch millis (non-midnight)", "1784106000000", "2026-07-15T09:00:00Z");
        // already ISO -> passthrough
        ok &= check("ISO passthrough", "2026-03-04T12:30:00Z", "2026-03-04T12:30:00Z");
        // quoted ISO -> quotes stripped
        ok &= check("quoted ISO -> unquoted", "'2026-03-04T12:30:00Z'", "2026-03-04T12:30:00Z");

        System.out.println(ok ? "\n=== DateLiteral test passed ===" : "\n=== DateLiteral test FAILED ===");
        if (!ok) System.exit(1);
    }

    private static boolean check(String label, String in, String expected) {
        String got = SharePointComputeTask.odataDateLiteral(in);
        boolean pass = expected.equals(got);
        System.out.printf("%-28s %-22s -> %-22s %s%n", label, in, got, pass ? "PASS" : "FAIL(exp " + expected + ")");
        return pass;
    }
}
