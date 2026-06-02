/**
 * Copyright (C) insightsoftware 2026. All rights reserved.
 *
 * Standalone smoke test that exercises SharePointGraphClient and
 * SharePointIntrospector against a real tenant. Reads credentials from
 * environment variables; prints discovery results and field metadata.
 *
 * Run via Maven exec:
 *   mvn -q exec:java \
 *     -Dexec.mainClass=com.zoomdata.connector.example.provider.sharepoint.SharePointSmokeTest \
 *     -Dexec.classpathScope=compile
 *
 * Env vars expected:
 *   SP_TENANT_ID, SP_CLIENT_ID, SP_CLIENT_SECRET, SP_SITE_URL
 */
package com.zoomdata.connector.example.provider.sharepoint;

import com.microsoft.graph.serviceclient.GraphServiceClient;
import com.zoomdata.connector.example.framework.async.Cursor;
import com.zoomdata.gen.edc.request.CollectionInfo;
import com.zoomdata.gen.edc.types.Field;
import com.zoomdata.gen.edc.types.FieldMetadata;
import com.zoomdata.gen.edc.types.Record;
import com.zoomdata.gen.edc.types.ResponseMetadata;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;

public class SharePointSmokeTest {

    public static void main(String[] args) {
        String tenantId = env("SP_TENANT_ID");
        String clientId = env("SP_CLIENT_ID");
        String clientSecret = env("SP_CLIENT_SECRET");
        String siteUrl = env("SP_SITE_URL");

        System.out.println("=== SharePoint EDC smoke test ===");
        System.out.println("Tenant:   " + tenantId);
        System.out.println("Client:   " + clientId);
        System.out.println("Site URL: " + siteUrl);
        System.out.println();

        SharePointTypesMapping types = new SharePointTypesMapping();
        SharePointGraphClient gf = new SharePointGraphClient();
        SharePointIntrospector intro = new SharePointIntrospector(types);

        System.out.println("--- 1. Build Graph client ---");
        GraphServiceClient client;
        String siteId;
        try {
            client = gf.build(tenantId, clientId, clientSecret, null);
            siteId = gf.siteIdFromUrl(siteUrl);
            System.out.println("siteId resolved: " + siteId);
        } catch (Exception e) {
            System.err.println("FAIL build/resolve: " + e.getMessage());
            e.printStackTrace();
            System.exit(1);
            return;
        }

        System.out.println();
        System.out.println("--- 2. Validate connection (call /sites/{id}) ---");
        boolean ok = intro.validateConnection(client, siteId);
        System.out.println("validateConnection: " + ok);
        if (!ok) {
            System.err.println("FAIL: cannot reach site");
            System.exit(2);
        }

        System.out.println();
        System.out.println("--- 3. Enumerate collections (Lists) ---");
        List<CollectionInfo> collections;
        try {
            collections = intro.getCollections(client, siteId, Collections.emptySet(), SharePointReaders.NONE);
            System.out.println("found " + collections.size() + " collection(s):");
            for (CollectionInfo c : collections) {
                System.out.println("  - " + c.getCollection() + "  (schema=" + c.getSchema() + ")");
            }
        } catch (Exception e) {
            System.err.println("FAIL getCollections: " + e.getMessage());
            e.printStackTrace();
            System.exit(3);
            return;
        }

        if (collections.isEmpty()) {
            System.out.println("(no collections to describe; smoke test ends here)");
            return;
        }

        System.out.println();
        System.out.println("--- 4. Describe each collection ---");
        for (CollectionInfo c : collections) {
            System.out.println("Describing: " + c.getCollection());
            try {
                List<FieldMetadata> fields = intro.describeCollection(client, siteId, c.getCollection(), SharePointReaders.NONE);
                System.out.println("  " + fields.size() + " column(s):");
                for (FieldMetadata fm : fields) {
                    String label = fm.getFieldParams() != null ? fm.getFieldParams().getFieldLabel() : null;
                    System.out.printf("    %-30s %-12s %s%n", fm.getName(), fm.getType(),
                            label != null ? "(" + label + ")" : "");
                }
            } catch (Exception e) {
                System.err.println("  WARN describe failed: " + e.getMessage());
            }
        }

        System.out.println();
        System.out.println("--- 5. Data fetch via ComputeTask (first collection with no filters) ---");
        CollectionInfo first = collections.get(0);
        try {
            List<FieldMetadata> meta = intro.describeCollection(client, siteId, first.getCollection(), SharePointReaders.NONE);
            List<String> fields = meta.stream().map(FieldMetadata::getName).collect(Collectors.toList());
            System.out.println("Fetching: " + first.getCollection() + " (" + fields.size() + " fields)");

            SharePointComputeTask task = new SharePointComputeTask(
                    client, siteId, first.getCollection(), fields,
                    types, intro, 100, null, null);
            Cursor cursor = task.compute();

            List<ResponseMetadata> respMeta = cursor.getMetadata();
            System.out.println("response metadata: " + respMeta.size() + " columns");

            int rowCount = 0;
            List<Record> sampleRows = new ArrayList<>();
            while (cursor.hasNext() && rowCount < 5) {
                sampleRows.add(cursor.next());
                rowCount++;
            }
            while (cursor.hasNext()) {
                cursor.next();
                rowCount++;
            }
            System.out.println("total rows fetched: " + rowCount);
            for (int i = 0; i < sampleRows.size(); i++) {
                Record r = sampleRows.get(i);
                System.out.print("  row " + (i + 1) + ":");
                List<Field> rf = r.getRecord();
                for (int j = 0; j < Math.min(rf.size(), 6); j++) {
                    String name = j < fields.size() ? fields.get(j) : "?";
                    String val = rf.get(j).isIsNull() ? "<null>" : rf.get(j).getValue();
                    if (val != null && val.length() > 20) val = val.substring(0, 20) + "...";
                    System.out.print(" " + name + "=" + val);
                }
                if (rf.size() > 6) System.out.print(" ...");
                System.out.println();
            }
        } catch (Exception e) {
            System.err.println("FAIL ComputeTask: " + e.getMessage());
            e.printStackTrace();
            System.exit(4);
        }

        System.out.println();
        System.out.println("=== Smoke test passed ===");
    }

    private static String env(String name) {
        String v = System.getenv(name);
        if (v == null || v.isEmpty()) {
            System.err.println("Missing required env var: " + name);
            System.exit(99);
        }
        return v;
    }
}
