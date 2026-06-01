/**
 * Copyright (C) insightsoftware 2026. All rights reserved.
 *
 * Schema discovery for the SharePoint EDC.
 *
 * Each visible SharePoint List in the configured site becomes one EDC
 * collection. Columns are discovered via Graph's /lists/{id}/columns
 * endpoint and mapped to Thrift FieldType via SharePointTypesMapping.
 *
 * Excel-in-SharePoint tables are supported via INCLUDE_EXCEL: each entry
 * is a path to an xlsx file in the site's default drive; every named
 * Workbook table in the file becomes its own collection.
 */
package com.zoomdata.connector.example.provider.sharepoint;

import com.microsoft.graph.models.ColumnDefinition;
import com.microsoft.graph.models.ColumnDefinitionCollectionResponse;
import com.microsoft.graph.models.DriveItem;
import com.microsoft.graph.models.ListCollectionResponse;
import com.microsoft.graph.serviceclient.GraphServiceClient;
import com.zoomdata.connector.example.framework.common.Meta;
import com.zoomdata.gen.edc.request.CollectionInfo;
import com.zoomdata.gen.edc.types.FieldMetadata;
import com.zoomdata.gen.edc.types.FieldParams;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

public class SharePointIntrospector {

    private static final Logger log = LoggerFactory.getLogger(SharePointIntrospector.class);

    static final String EXCEL_PREFIX = "excel__";

    private final SharePointTypesMapping typesMapping;

    public SharePointIntrospector(SharePointTypesMapping typesMapping) {
        this.typesMapping = typesMapping;
    }

    public boolean validateConnection(GraphServiceClient client, String siteId) {
        try {
            client.sites().bySiteId(siteId).get();
            return true;
        } catch (Exception e) {
            log.error("Site validation failed for {}: {}", siteId, e.getMessage());
            return false;
        }
    }

    public List<CollectionInfo> getCollections(GraphServiceClient client, String siteId,
                                                Set<String> includeLists, List<String> includeExcel) {
        List<CollectionInfo> collections = new ArrayList<>();

        // SharePoint Lists
        try {
            ListCollectionResponse lists = client.sites().bySiteId(siteId).lists().get();
            if (lists != null && lists.getValue() != null) {
                for (com.microsoft.graph.models.List list : lists.getValue()) {
                    if (Boolean.TRUE.equals(list.getList() != null ? list.getList().getHidden() : null)) {
                        continue;
                    }
                    String name = list.getDisplayName();
                    if (name == null || name.isEmpty()) continue;
                    if (includeLists != null && !includeLists.isEmpty() && !includeLists.contains(name)) {
                        continue;
                    }
                    CollectionInfo ci = new CollectionInfo();
                    ci.setCollection(name);
                    ci.setSchema("default");
                    collections.add(ci);
                    log.debug("Discovered List: {} (id={})", name, list.getId());
                }
            }
        } catch (Exception e) {
            log.error("Failed to enumerate Lists at site {}: {}", siteId, e.getMessage());
        }

        // Excel tables (only if INCLUDE_EXCEL configured)
        if (includeExcel != null && !includeExcel.isEmpty()) {
            for (String path : includeExcel) {
                try {
                    collections.addAll(discoverExcelTables(client, siteId, path));
                } catch (Exception e) {
                    log.warn("Excel discovery failed for path '{}': {}", path, e.getMessage());
                }
            }
        }

        return collections;
    }

    public List<FieldMetadata> describeCollection(GraphServiceClient client, String siteId,
                                                   String collectionName) {
        if (collectionName.startsWith(EXCEL_PREFIX)) {
            return describeExcelTable(client, siteId, collectionName);
        }
        return describeList(client, siteId, collectionName);
    }

    private List<FieldMetadata> describeList(GraphServiceClient client, String siteId, String listName) {
        String listId = resolveListId(client, siteId, listName);
        if (listId == null) {
            throw new RuntimeException("List not found: " + listName);
        }

        ColumnDefinitionCollectionResponse cols = client.sites().bySiteId(siteId)
                .lists().byListId(listId).columns().get();
        if (cols == null || cols.getValue() == null) {
            return Collections.emptyList();
        }

        List<FieldMetadata> metadata = new ArrayList<>();
        for (ColumnDefinition col : cols.getValue()) {
            if (Boolean.TRUE.equals(col.getHidden())) continue;
            String columnGroup = col.getColumnGroup();
            if (columnGroup != null && columnGroup.startsWith("_Hidden")) continue;

            String name = col.getName();
            if (name == null || name.isEmpty()) continue;
            // Skip read-only system columns that have no user-visible value
            if (Boolean.TRUE.equals(col.getReadOnly())
                    && (name.startsWith("_") || "ContentType".equals(name) || "Attachments".equals(name))) {
                continue;
            }

            String typeKey = resolveColumnTypeKey(col);
            Meta meta = typesMapping.metaForType(typeKey);

            FieldMetadata fm = new FieldMetadata();
            fm.setName(name);
            fm.setType(meta.getThriftType());
            FieldParams params = new FieldParams();
            params.setFieldName(name);
            if (col.getDisplayName() != null && !col.getDisplayName().isEmpty()) {
                params.setFieldLabel(col.getDisplayName());
            }
            fm.setFieldParams(params);
            metadata.add(fm);
            log.debug("List column {} -> type {} (kind: {})", name, meta.getThriftType(), typeKey);
        }
        return metadata;
    }

    /**
     * Determine which SharePoint column-kind discriminator is set on a
     * ColumnDefinition and return the lowercase key used by
     * SharePointTypesMapping. Defaults to "text" when no discriminator
     * matches (Graph occasionally returns columns with no kind set).
     */
    private String resolveColumnTypeKey(ColumnDefinition col) {
        if (col.getText() != null) return "text";
        if (col.getNumber() != null) return "number";
        if (col.getCurrency() != null) return "currency";
        if (col.getDateTime() != null) return "datetime";
        if (col.getBoolean() != null) return "boolean";
        if (col.getChoice() != null) return "choice";
        if (col.getLookup() != null) return "lookup";
        if (col.getPersonOrGroup() != null) return "personorgroup";
        if (col.getHyperlinkOrPicture() != null) return "hyperlinkorpicture";
        if (col.getCalculated() != null) return "calculated";
        if (col.getThumbnail() != null) return "thumbnail";
        if (col.getTerm() != null) return "term";
        return "text";
    }

    /**
     * Resolve a List displayName to its GUID identifier.
     */
    public String resolveListId(GraphServiceClient client, String siteId, String listName) {
        ListCollectionResponse lists = client.sites().bySiteId(siteId).lists().get();
        if (lists == null || lists.getValue() == null) return null;
        for (com.microsoft.graph.models.List list : lists.getValue()) {
            if (listName.equals(list.getDisplayName())) {
                return list.getId();
            }
        }
        return null;
    }

    // ----- Excel support (v1.1) -----------------------------------------
    //
    // v1 stub. Excel-in-SharePoint via Graph's Workbook API is on the
    // roadmap; the connection parameter INCLUDE_EXCEL is accepted now so
    // that configurations don't need to change later. Discovery returns
    // nothing; the ComputeTask will refuse Excel collections with a clear
    // error.

    private List<CollectionInfo> discoverExcelTables(GraphServiceClient client, String siteId, String path) {
        log.warn("Excel discovery requested for '{}' but Excel support is not implemented in v1. "
                + "Skipping; use SharePoint Lists in v1.", path);
        return Collections.emptyList();
    }

    private List<FieldMetadata> describeExcelTable(GraphServiceClient client, String siteId,
                                                    String collectionName) {
        throw new RuntimeException("Excel-in-SharePoint is not implemented in v1. "
                + "Use SharePoint Lists, or wait for the v1.1 Workbook API integration.");
    }

    public DriveItem resolveDriveItemByPath(GraphServiceClient client, String siteId, String path) {
        log.warn("Drive path resolution is not implemented in v1 (path={})", path);
        return null;
    }

    static String excelSlugFromPath(String path) {
        String trimmed = path.replaceAll("^/+", "").replaceAll("\\.xlsx$", "");
        return trimmed.replaceAll("[^A-Za-z0-9]+", "_").toLowerCase();
    }

    static ExcelRef parseExcelCollection(String collectionName) {
        if (!collectionName.startsWith(EXCEL_PREFIX)) return null;
        // Format: excel__<file-slug>__<table-name>
        String rest = collectionName.substring(EXCEL_PREFIX.length());
        int idx = rest.lastIndexOf("__");
        if (idx < 0) return null;
        ExcelRef ref = new ExcelRef();
        ref.fileSlug = rest.substring(0, idx);
        ref.tableName = rest.substring(idx + 2);
        // The path is reconstructed by the caller; we keep slug for lookup.
        return ref;
    }

    static class ExcelRef {
        String fileSlug;
        String tableName;
        String path;
    }

    /** Parse a comma-separated allowlist into a Set, returning empty for null/blank input. */
    public static Set<String> parseAllowlist(String csv) {
        if (csv == null || csv.trim().isEmpty()) return Collections.emptySet();
        return new HashSet<>(Arrays.asList(csv.split("\\s*,\\s*")));
    }

    /** Parse a comma-separated list of Excel paths. */
    public static List<String> parsePaths(String csv) {
        if (csv == null || csv.trim().isEmpty()) return Collections.emptyList();
        return Arrays.asList(csv.split("\\s*,\\s*"));
    }
}
