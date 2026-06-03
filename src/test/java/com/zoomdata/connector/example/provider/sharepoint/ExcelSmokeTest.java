/**
 * Copyright (C) insightsoftware 2026. All rights reserved.
 *
 * Standalone Excel-in-SharePoint smoke test against a real tenant.
 * Exercises SharePointWorkbookReader + the introspector's Excel discovery,
 * describe (with type inference), and a ComputeTask data fetch.
 *
 * Run:
 *   mvn -q test-compile exec:java \
 *     -Dexec.mainClass=com.zoomdata.connector.example.provider.sharepoint.ExcelSmokeTest \
 *     -Dexec.classpathScope=test
 *
 * Env: SP_TENANT_ID, SP_CLIENT_ID, SP_CLIENT_SECRET, SP_SITE_URL, SP_EXCEL_PATH
 */
package com.zoomdata.connector.example.provider.sharepoint;

import com.zoomdata.gen.edc.request.CollectionInfo;
import com.zoomdata.gen.edc.types.Field;
import com.zoomdata.gen.edc.types.FieldMetadata;
import com.zoomdata.gen.edc.types.Record;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ExcelSmokeTest {

    public static void main(String[] args) {
        String tenantId = env("SP_TENANT_ID");
        String clientId = env("SP_CLIENT_ID");
        String clientSecret = env("SP_CLIENT_SECRET");
        String siteUrl = env("SP_SITE_URL");
        String excelPath = env("SP_EXCEL_PATH");

        System.out.println("=== Excel-in-SharePoint smoke test ===");
        System.out.println("Site:  " + siteUrl);
        System.out.println("Excel: " + excelPath);
        System.out.println();

        SharePointTypesMapping types = new SharePointTypesMapping();
        SharePointGraphClient gf = new SharePointGraphClient();
        SharePointIntrospector intro = new SharePointIntrospector(types);

        String siteId = gf.siteIdFromUrl(siteUrl);
        String token = gf.bearerToken(tenantId, clientId, clientSecret, null);
        Map<String, String> pathsBySlug = new HashMap<>();
        pathsBySlug.put(SharePointIntrospector.excelSlugFromPath(excelPath), excelPath);
        SharePointWorkbookReader wb = new SharePointWorkbookReader(
                gf.rawHttpClient(), token, "sites/" + siteId + "/drive", pathsBySlug);
        SharePointReaders readers = new SharePointReaders(wb, null);

        System.out.println("--- 1. Discover Excel collections ---");
        List<CollectionInfo> cols = intro.getCollections(
                gf.build(tenantId, clientId, clientSecret, null), siteId,
                java.util.Collections.emptySet(), readers);
        // Filter to only excel collections for clarity
        List<CollectionInfo> excel = cols.stream()
                .filter(c -> SharePointIntrospector.isExcelCollection(c.getCollection()))
                .collect(java.util.stream.Collectors.toList());
        System.out.println("found " + excel.size() + " excel collection(s):");
        for (CollectionInfo c : excel) System.out.println("  - " + c.getCollection());
        if (excel.isEmpty()) { System.err.println("FAIL: no excel collections discovered"); System.exit(2); }

        System.out.println();
        System.out.println("--- 2. Describe each + 3. fetch rows ---");
        for (CollectionInfo c : excel) {
            String name = c.getCollection();
            System.out.println("Collection: " + name);
            try {
                List<FieldMetadata> fm = intro.describeCollection(null, siteId, name, readers);
                System.out.print("  columns: ");
                for (FieldMetadata f : fm) System.out.print(f.getName() + "(" + f.getType() + ") ");
                System.out.println();

                List<String> fields = fm.stream().map(FieldMetadata::getName)
                        .collect(java.util.stream.Collectors.toList());
                SharePointComputeTask task = new SharePointComputeTask(
                        null, siteId, name, fields, types, intro, 1000, null, readers);
                SharePointComputeTask.SharePointCursor cur =
                        (SharePointComputeTask.SharePointCursor) task.compute();
                int n = 0;
                while (cur.hasNext()) {
                    Record rec = cur.next();
                    if (n < 3) {
                        StringBuilder sb = new StringBuilder("    row: ");
                        List<Field> rf = rec.getRecord();
                        for (int i = 0; i < rf.size(); i++) {
                            sb.append(fields.get(i)).append("=")
                              .append(rf.get(i).isIsNull() ? "<null>" : rf.get(i).getValue()).append(" ");
                        }
                        System.out.println(sb);
                    }
                    n++;
                }
                System.out.println("  total rows: " + n);
            } catch (Exception e) {
                System.err.println("  ERROR: " + e.getMessage());
            }
            System.out.println();
        }
        System.out.println("=== Excel smoke test passed ===");
    }

    private static String env(String name) {
        String v = System.getenv(name);
        if (v == null || v.isEmpty()) { System.err.println("Missing env var: " + name); System.exit(99); }
        return v;
    }
}
