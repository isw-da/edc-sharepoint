/**
 * Copyright (C) insightsoftware 2026. All rights reserved.
 *
 * Unit test for TtlGridCache — counts loader invocations to prove:
 *   - TTL <= 0 (disabled): loads every call
 *   - TTL > 0: a repeat within the window is served from cache (loader not
 *     called again); after expiry the loader runs again.
 *
 * Run:
 *   mvn -q test-compile exec:java \
 *     -Dexec.mainClass=com.zoomdata.connector.example.provider.sharepoint.TtlGridCacheTest \
 *     -Dexec.classpathScope=test
 */
package com.zoomdata.connector.example.provider.sharepoint;

import java.util.Collections;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

public class TtlGridCacheTest {

    public static void main(String[] args) throws Exception {
        boolean ok = true;

        // Disabled cache (TTL 0): loads every time.
        TtlGridCache off = new TtlGridCache(0, 10);
        AtomicInteger c1 = new AtomicInteger();
        off.get("k", () -> grid(c1));
        off.get("k", () -> grid(c1));
        boolean a = (c1.get() == 2) && !off.enabled();
        System.out.printf("A. disabled -> loads each call: loads=%d -> %s%n", c1.get(), a ? "PASS" : "FAIL");
        ok &= a;

        // Enabled cache (TTL 2s): repeat within window served from cache.
        TtlGridCache on = new TtlGridCache(2, 10);
        AtomicInteger c2 = new AtomicInteger();
        on.get("k", () -> grid(c2));
        on.get("k", () -> grid(c2));
        on.get("k", () -> grid(c2));
        boolean b = (c2.get() == 1) && on.enabled();
        System.out.printf("B. enabled -> repeat from cache: loads=%d (expect 1) -> %s%n", c2.get(), b ? "PASS" : "FAIL");
        ok &= b;

        // After TTL expiry: loads again.
        Thread.sleep(2200);
        on.get("k", () -> grid(c2));
        boolean c = (c2.get() == 2);
        System.out.printf("C. after expiry -> reloads: loads=%d (expect 2) -> %s%n", c2.get(), c ? "PASS" : "FAIL");
        ok &= c;

        System.out.println(ok ? "\n=== TtlGridCache test passed ===" : "\n=== TtlGridCache test FAILED ===");
        if (!ok) System.exit(1);
    }

    private static List<List<Object>> grid(AtomicInteger counter) {
        counter.incrementAndGet();
        return Collections.singletonList(Collections.singletonList((Object) "x"));
    }
}
