/**
 * Copyright (C) insightsoftware 2026. All rights reserved.
 *
 * OneDrive data-path smoke test. Exercises the exact OneDrive code path:
 * a file reader rooted at users/{upn}/drive, siteId == null (so List
 * discovery is skipped), through discover -> describe -> fetch.
 *
 * Run:
 *   mvn -q test-compile exec:java \
 *     -Dexec.mainClass=com.zoomdata.connector.example.provider.sharepoint.OneDriveSmokeTest \
 *     -Dexec.classpathScope=test
 *
 * Env: SP_TENANT_ID, SP_CLIENT_ID, SP_CLIENT_SECRET, SP_UPN, SP_FILE
 */
package com.zoomdata.connector.example.provider.sharepoint;

import com.zoomdata.gen.edc.request.CollectionInfo;
import com.zoomdata.gen.edc.types.Field;
import com.zoomdata.gen.edc.types.FieldMetadata;
import com.zoomdata.gen.edc.types.Record;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class OneDriveSmokeTest {

    public static void main(String[] args) {
        String tenantId = env("SP_TENANT_ID");
        String clientId = env("SP_CLIENT_ID");
        String clientSecret = env("SP_CLIENT_SECRET");
        String upn = env("SP_UPN");
        String filePath = env("SP_FILE");

        System.out.println("=== OneDrive data-path smoke test ===");
        System.out.println("UPN:  " + upn);
        System.out.println("File: " + filePath);
        System.out.println();

        SharePointTypesMapping types = new SharePointTypesMapping();
        SharePointGraphClient gf = new SharePointGraphClient();
        SharePointIntrospector intro = new SharePointIntrospector(types);

        String token = gf.bearerToken(tenantId, clientId, clientSecret, null);
        String driveResourcePath = "users/" + upn + "/drive";
        Map<String, String> pathsBySlug = new HashMap<>();
        pathsBySlug.put(SharePointIntrospector.fileSlugFromPath(filePath), filePath);
        SharePointFileReader fr = new SharePointFileReader(gf.rawHttpClient(), token, driveResourcePath, pathsBySlug);
        SharePointReaders readers = new SharePointReaders(null, fr);

        // 1. discover — siteId null (no Lists), client null (unused for files)
        System.out.println("--- 1. Discover (siteId=null -> no Lists, just the file) ---");
        List<CollectionInfo> cols = intro.getCollections(null, null, java.util.Collections.emptySet(), readers);
        System.out.println("collections: ");
        for (CollectionInfo c : cols) System.out.println("  - " + c.getCollection());
        if (cols.isEmpty()) { System.err.println("FAIL: no collections"); System.exit(2); }
        String name = cols.get(0).getCollection();

        // 2. describe
        System.out.println();
        System.out.println("--- 2. Describe " + name + " ---");
        List<FieldMetadata> fm = intro.describeCollection(null, null, name, readers);
        System.out.print("  columns: ");
        for (FieldMetadata f : fm) System.out.print(f.getName() + "(" + f.getType() + ") ");
        System.out.println();

        // 3. fetch
        System.out.println();
        System.out.println("--- 3. Fetch rows ---");
        List<String> fields = fm.stream().map(FieldMetadata::getName).collect(java.util.stream.Collectors.toList());
        SharePointComputeTask task = new SharePointComputeTask(
                null, null, name, fields, types, intro, 1000, null, readers);
        SharePointComputeTask.SharePointCursor cur =
                (SharePointComputeTask.SharePointCursor) task.compute();
        int n = 0;
        while (cur.hasNext()) {
            Record r = cur.next();
            if (n < 5) {
                StringBuilder sb = new StringBuilder("  row: ");
                List<Field> rf = r.getRecord();
                for (int i = 0; i < rf.size(); i++)
                    sb.append(fields.get(i)).append("=").append(rf.get(i).isIsNull() ? "<null>" : rf.get(i).getValue()).append(" ");
                System.out.println(sb);
            }
            n++;
        }
        System.out.println("  total rows: " + n);
        if (n == 0) { System.err.println("FAIL: no rows"); System.exit(3); }
        System.out.println();
        System.out.println("=== OneDrive smoke test passed ===");
    }

    private static String env(String name) {
        String v = System.getenv(name);
        if (v == null || v.isEmpty()) { System.err.println("Missing env var: " + name); System.exit(99); }
        return v;
    }
}
