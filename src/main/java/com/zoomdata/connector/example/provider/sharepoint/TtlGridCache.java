/**
 * Copyright (C) insightsoftware 2026. All rights reserved.
 *
 * Small pod-lifetime TTL cache for parsed grids (Excel ranges / file contents).
 *
 * Rationale: aggregation pushdown to Graph is impossible (Graph ignores
 * $apply), so caching is the remaining lever for repeated reads. A Composer
 * source action typically describes then fetches the same entity, and
 * dashboard refreshes re-read it; a short TTL collapses those into one Graph
 * round-trip. It caches the SAME parsed grids the readers already materialise
 * in memory, so it adds no new large buffers (file *content* streams are not
 * cached — only the resulting bounded grid).
 *
 * Opt-in: TTL <= 0 disables it entirely (default). Deliberately tolerates a
 * benign load race (two callers may both load a cold key) and bounds memory
 * by clearing when maxEntries is exceeded — adequate for a soft cache over a
 * handful of configured files.
 *
 * Time is read via System.currentTimeMillis(); staleness is intentional and
 * bounded by the TTL.
 */
package com.zoomdata.connector.example.provider.sharepoint;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ConcurrentMap;
import java.util.function.Supplier;

public class TtlGridCache {

    private static final Logger log = LoggerFactory.getLogger(TtlGridCache.class);

    private final long ttlMs;
    private final int maxEntries;
    private final ConcurrentMap<String, Entry> map = new ConcurrentHashMap<>();

    public TtlGridCache(long ttlSeconds, int maxEntries) {
        this.ttlMs = ttlSeconds * 1000L;
        this.maxEntries = maxEntries;
    }

    public boolean enabled() {
        return ttlMs > 0;
    }

    /**
     * Return the cached grid for {@code key} if fresh, else load via
     * {@code loader}, cache, and return. When disabled, always loads.
     */
    public List<List<Object>> get(String key, Supplier<List<List<Object>>> loader) {
        if (ttlMs <= 0) {
            return loader.get();
        }
        long now = System.currentTimeMillis();
        Entry e = map.get(key);
        if (e != null && (now - e.timestamp) < ttlMs) {
            log.debug("Grid cache hit: {}", key);
            return e.grid;
        }
        List<List<Object>> grid = loader.get();
        if (map.size() >= maxEntries) {
            log.debug("Grid cache full ({} entries) — clearing", map.size());
            map.clear();
        }
        map.put(key, new Entry(grid, now));
        return grid;
    }

    /** Test/diagnostic. */
    int size() { return map.size(); }

    private static final class Entry {
        final List<List<Object>> grid;
        final long timestamp;
        Entry(List<List<Object>> grid, long timestamp) { this.grid = grid; this.timestamp = timestamp; }
    }
}
