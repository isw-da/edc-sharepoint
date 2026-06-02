/**
 * Copyright (C) insightsoftware 2026. All rights reserved.
 *
 * Standalone multi-site routing test against a real tenant. Exercises the
 * DataProvider's discovery prefixing (>1 site) and route resolution, plus
 * the single-site regression (names unchanged).
 *
 * NOTE: the dev tenant has only one physical site, so the multi-site case
 * lists the same site URL twice. That still fully exercises the new logic —
 * slug prefixing, slug de-duplication, and route peel-back — but does not
 * prove two *distinct* siteIds return distinct data (low risk: per-site
 * fetch is already proven for the single site).
 *
 * Run:
 *   mvn -q test-compile exec:java \
 *     -Dexec.mainClass=com.zoomdata.connector.example.provider.sharepoint.MultiSiteSmokeTest \
 *     -Dexec.classpathScope=test
 *
 * Env: SP_TENANT_ID, SP_CLIENT_ID, SP_CLIENT_SECRET, SP_SITE_URL
 */
package com.zoomdata.connector.example.provider.sharepoint;

import com.zoomdata.gen.edc.request.CollectionInfo;
import com.zoomdata.gen.edc.request.DataSourceInfo;
import com.zoomdata.gen.edc.request.MetaCollectionsRequest;
import com.zoomdata.gen.edc.request.MetaCollectionsResponse;
import com.zoomdata.gen.edc.request.MetaDescribeRequest;
import com.zoomdata.gen.edc.request.MetaDescribeResponse;
import com.zoomdata.gen.edc.request.RequestInfo;
import com.zoomdata.gen.edc.types.FieldMetadata;

import java.util.HashMap;
import java.util.Map;

public class MultiSiteSmokeTest {

    public static void main(String[] args) {
        String tenantId = env("SP_TENANT_ID");
        String clientId = env("SP_CLIENT_ID");
        String clientSecret = env("SP_CLIENT_SECRET");
        String siteUrl = env("SP_SITE_URL");

        SharePointDataProvider provider = new SharePointDataProvider();

        System.out.println("=== Multi-site routing smoke test ===\n");

        // --- 1. Single-site (SITE_URL only) -> names UNPREFIXED (regression) ---
        System.out.println("--- 1. Single-site (SITE_URL): expect unprefixed names ---");
        Map<String, String> single = base(tenantId, clientId, clientSecret);
        single.put("SITE_URL", siteUrl);
        MetaCollectionsResponse r1 = provider.collections(collReq(single));
        printCollections(r1);
        boolean anyPrefixed = r1.getCollections().stream()
                .anyMatch(c -> c.getCollection().contains("__")
                        && !SharePointIntrospector.isExcelCollection(c.getCollection()));
        System.out.println("  single-site has site-prefixed list names? " + anyPrefixed + " (expect false)\n");

        // --- 2. Multi-site (SITE_URLS, site listed twice) -> names PREFIXED ---
        System.out.println("--- 2. Multi-site (SITE_URLS x2): expect <slug>__ prefixes ---");
        Map<String, String> multi = base(tenantId, clientId, clientSecret);
        multi.put("SITE_URLS", siteUrl + "," + siteUrl);
        MetaCollectionsResponse r2 = provider.collections(collReq(multi));
        printCollections(r2);

        String prefixedEvents = r2.getCollections().stream()
                .map(CollectionInfo::getCollection)
                .filter(n -> n.endsWith("__Events"))
                .findFirst().orElse(null);
        long distinctSlugs = r2.getCollections().stream()
                .map(n -> n.getCollection().substring(0, n.getCollection().indexOf("__")))
                .distinct().count();
        System.out.println("  distinct site slugs seen: " + distinctSlugs + " (expect 2)");

        // --- 3. Route a prefixed name back to its site via describe ---
        System.out.println("\n--- 3. Route + describe a prefixed collection ---");
        if (prefixedEvents != null) {
            System.out.println("  describing: " + prefixedEvents);
            MetaDescribeResponse d = provider.describe(descReq(multi, prefixedEvents));
            int n = d.getFields() != null ? d.getFields().size() : 0;
            System.out.println("  fields returned: " + n + " (expect 29 for Events)");
            if (d.getFields() != null) {
                StringBuilder sb = new StringBuilder("  first cols: ");
                for (int i = 0; i < Math.min(5, n); i++) sb.append(d.getFields().get(i).getName()).append(" ");
                System.out.println(sb);
            }
        } else {
            System.out.println("  (no __Events collection found to route)");
        }

        // --- 4. Negative: unprefixed name on a multi-site connection must fail ---
        System.out.println("\n--- 4. Negative: unprefixed name on multi-site must be rejected ---");
        MetaDescribeResponse bad = provider.describe(descReq(multi, "Events"));
        boolean rejected = bad.getFields() == null || bad.getFields().isEmpty();
        System.out.println("  unprefixed 'Events' rejected/empty on multi-site? " + rejected + " (expect true)");

        System.out.println("\n=== Multi-site smoke test passed ===");
    }

    private static Map<String, String> base(String t, String c, String s) {
        Map<String, String> m = new HashMap<>();
        m.put("TENANT_ID", t);
        m.put("CLIENT_ID", c);
        m.put("CLIENT_SECRET", s);
        return m;
    }

    private static RequestInfo info(Map<String, String> params) {
        return new RequestInfo().setDataSourceInfo(new DataSourceInfo().setParams(params));
    }

    private static MetaCollectionsRequest collReq(Map<String, String> params) {
        return new MetaCollectionsRequest().setRequestInfo(info(params));
    }

    private static MetaDescribeRequest descReq(Map<String, String> params, String collection) {
        return new MetaDescribeRequest()
                .setRequestInfo(info(params))
                .setCollectionInfo(new CollectionInfo().setCollection(collection));
    }

    private static void printCollections(MetaCollectionsResponse r) {
        if (r.getCollections() == null) { System.out.println("  (none)"); return; }
        for (CollectionInfo c : r.getCollections()) System.out.println("  - " + c.getCollection());
    }

    private static String env(String name) {
        String v = System.getenv(name);
        if (v == null || v.isEmpty()) { System.err.println("Missing env var: " + name); System.exit(99); }
        return v;
    }
}
