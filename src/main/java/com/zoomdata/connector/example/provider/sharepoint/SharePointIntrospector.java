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
import com.microsoft.graph.models.ListCollectionResponse;
import com.microsoft.graph.serviceclient.GraphServiceClient;
import com.zoomdata.connector.example.framework.common.Meta;
import com.zoomdata.gen.edc.request.CollectionInfo;
import com.zoomdata.gen.edc.types.FieldMetadata;
import com.zoomdata.gen.edc.types.FieldParams;
import com.zoomdata.gen.edc.types.FieldType;
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

    // v1.1: three Excel surfaces, each its own collection-name prefix.
    // Format: <prefix><file-slug>__<entity>, where <file-slug> never
    // contains "__" (excelSlugFromPath collapses non-alnum runs to one "_"),
    // so the entity (real table/sheet/range name) may itself contain "__".
    static final String PREFIX_TABLE = "excel_table__";
    static final String PREFIX_SHEET = "excel_sheet__";
    static final String PREFIX_RANGE = "excel_range__";

    // v1.1: file collections (CSV/TSV/JSON). One collection per file:
    // file__<file-slug>. The file slug (fileSlugFromPath) collapses non-alnum
    // runs to a single "_" so it never contains "__".
    static final String PREFIX_FILE = "file__";

    enum ExcelSurface { TABLE, SHEET, RANGE }

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
                                                Set<String> includeLists, SharePointReaders readers) {
        List<CollectionInfo> collections = new ArrayList<>();

        // SharePoint Lists — follow @odata.nextLink across pages. Skipped
        // entirely for drive-only containers (OneDrive: siteId == null), which
        // have no SharePoint lists.
        if (siteId != null) try {
            int doclibsSkipped = 0;
            for (com.microsoft.graph.models.List list : iterateAllLists(client, siteId)) {
                com.microsoft.graph.models.ListInfo info = list.getList();
                if (Boolean.TRUE.equals(info != null ? info.getHidden() : null)) {
                    continue;
                }
                String name = list.getDisplayName();
                if (name == null || name.isEmpty()) continue;
                if (includeLists != null && !includeLists.isEmpty() && !includeLists.contains(name)) {
                    continue;
                }
                // v1.1 (M3 audit fix): document libraries appear here but
                // their items are DriveItem-shaped, not classic ListItem
                // field bags. Surfacing them as collections would return
                // mostly-null rows from /lists/{id}/items because Graph
                // doesn't populate the field-set that way for libraries.
                // Skip with an explicit log; the v1.1 INCLUDE_FILES path
                // is the right way to read files from a library.
                if (isDocumentLibrary(info)) {
                    doclibsSkipped++;
                    log.debug("Skipping document library '{}' (template={}); use INCLUDE_FILES to read files instead",
                            name, info != null ? info.getTemplate() : "?");
                    continue;
                }
                CollectionInfo ci = new CollectionInfo();
                ci.setCollection(name);
                ci.setSchema("default");
                collections.add(ci);
                log.debug("Discovered List: {} (id={})", name, list.getId());
            }
            if (doclibsSkipped > 0) {
                log.info("Skipped {} document library/libraries at site {} — use INCLUDE_FILES " +
                        "to expose files as collections (v1.1).", doclibsSkipped, siteId);
            }
        } catch (Exception e) {
            log.error("Failed to enumerate Lists at site {}: {}", siteId, e.getMessage());
        }

        // Excel surfaces (only if INCLUDE_EXCEL configured)
        if (readers != null && readers.workbook != null) {
            for (String path : readers.workbook.configuredPaths()) {
                try {
                    collections.addAll(discoverExcelCollections(readers.workbook, path));
                } catch (Exception e) {
                    log.warn("Excel discovery failed for path '{}': {}", path, e.getMessage());
                }
            }
        }

        // File collections (CSV/TSV/JSON) — one collection per INCLUDE_FILES path.
        if (readers != null && readers.files != null) {
            for (String path : readers.files.configuredPaths()) {
                try {
                    CollectionInfo ci = discoverFileCollection(path);
                    if (ci != null) collections.add(ci);
                } catch (Exception e) {
                    log.warn("File discovery failed for path '{}': {}", path, e.getMessage());
                }
            }
        }

        return collections;
    }

    public List<FieldMetadata> describeCollection(GraphServiceClient client, String siteId,
                                                   String collectionName,
                                                   SharePointReaders readers) {
        if (isExcelCollection(collectionName)) {
            return describeExcelCollection(readers != null ? readers.workbook : null, collectionName);
        }
        if (isFileCollection(collectionName)) {
            return describeFileCollection(readers != null ? readers.files : null, collectionName);
        }
        return describeList(client, siteId, collectionName);
    }

    private List<FieldMetadata> describeList(GraphServiceClient client, String siteId, String listName) {
        com.microsoft.graph.models.List target = findListByName(client, siteId, listName);
        if (target == null) {
            throw new RuntimeException("List not found: " + listName);
        }
        // v1.1 (M3): refuse describe/fetch on document libraries — the
        // List API can't return their data in a useful shape. Operator gets
        // a clear error rather than mystery null rows. INCLUDE_FILES is
        // the documented path for library content.
        if (isDocumentLibrary(target.getList())) {
            throw new RuntimeException("'" + listName + "' is a document library, not a List. "
                    + "Use INCLUDE_FILES to expose files from this library as collections.");
        }
        String listId = target.getId();

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
     * Resolve a List displayName to its GUID identifier. Walks all pages of
     * the /lists response — tenants with many sites/lists routinely exceed
     * the default ~200 item page and previous single-page lookups would
     * silently fail for the (N>200)th list.
     */
    public String resolveListId(GraphServiceClient client, String siteId, String listName) {
        com.microsoft.graph.models.List l = findListByName(client, siteId, listName);
        return l != null ? l.getId() : null;
    }

    private com.microsoft.graph.models.List findListByName(GraphServiceClient client, String siteId, String listName) {
        for (com.microsoft.graph.models.List list : iterateAllLists(client, siteId)) {
            if (listName.equals(list.getDisplayName())) {
                return list;
            }
        }
        return null;
    }

    /**
     * Detect whether a List is actually a document library. Graph's /lists
     * endpoint includes both classic Lists and libraries; their items have
     * different shapes and need different code paths.
     */
    private static boolean isDocumentLibrary(com.microsoft.graph.models.ListInfo info) {
        if (info == null) return false;
        String template = info.getTemplate();
        return template != null && template.equalsIgnoreCase("documentLibrary");
    }

    /**
     * Walk every page of /sites/{id}/lists, materialising all entries into a
     * single list. Uses @odata.nextLink via withUrl(). Caller iterates the
     * returned list normally.
     */
    private List<com.microsoft.graph.models.List> iterateAllLists(GraphServiceClient client, String siteId) {
        List<com.microsoft.graph.models.List> all = new ArrayList<>();
        ListCollectionResponse page = client.sites().bySiteId(siteId).lists().get();
        int pageCount = 0;
        while (page != null) {
            pageCount++;
            if (page.getValue() != null) all.addAll(page.getValue());
            String nextLink = page.getOdataNextLink();
            if (nextLink == null || nextLink.isEmpty()) break;
            if (pageCount > 50) {
                log.warn("List pagination exceeded 50 pages at site {} — capping to avoid runaway", siteId);
                break;
            }
            page = client.sites().bySiteId(siteId).lists().withUrl(nextLink).get();
        }
        log.debug("Walked {} page(s) of lists at site {}, {} entries total", pageCount, siteId, all.size());
        return all;
    }

    // ----- Excel support (v1.1, via Workbook API / SharePointWorkbookReader) -----
    //
    // Each xlsx in INCLUDE_EXCEL contributes up to three kinds of collection:
    //   excel_table__<slug>__<tableName>   — every Excel table (header row)
    //   excel_sheet__<slug>__<sheetName>   — every visible sheet's used-range
    //   excel_range__<slug>__<rangeName>   — every visible named Range
    // Cell values come back as a 2D array via the raw reader; we never touch
    // the Kiota Json type (which cannot represent the array).

    private List<CollectionInfo> discoverExcelCollections(SharePointWorkbookReader wb, String path) {
        List<CollectionInfo> out = new ArrayList<>();
        SharePointWorkbookReader.FileRef f = wb.resolveFile(path);
        if (f == null) return out; // resolveFile already logged why
        String slug = excelSlugFromPath(path);
        for (String t : wb.listTables(f)) out.add(excelCi(ExcelSurface.TABLE, slug, t));
        for (String s : wb.listWorksheets(f)) out.add(excelCi(ExcelSurface.SHEET, slug, s));
        for (String r : wb.listNamedRanges(f)) out.add(excelCi(ExcelSurface.RANGE, slug, r));
        log.info("Excel '{}': discovered {} collection(s)", path, out.size());
        return out;
    }

    private CollectionInfo excelCi(ExcelSurface surface, String slug, String entity) {
        CollectionInfo ci = new CollectionInfo();
        ci.setCollection(excelCollectionName(surface, slug, entity));
        ci.setSchema("default");
        return ci;
    }

    private List<FieldMetadata> describeExcelCollection(SharePointWorkbookReader wb, String collectionName) {
        if (wb == null) {
            throw new RuntimeException("Excel collection '" + collectionName + "' requested but no "
                    + "workbook reader is available (INCLUDE_EXCEL not configured on this connection?)");
        }
        ExcelColl c = parseExcelCollection(collectionName);
        if (c == null) throw new RuntimeException("Invalid Excel collection name: " + collectionName);
        String path = wb.pathForSlug(c.fileSlug);
        if (path == null) {
            throw new RuntimeException("Excel file for slug '" + c.fileSlug
                    + "' is not in this connection's INCLUDE_EXCEL list");
        }
        SharePointWorkbookReader.FileRef f = wb.resolveFile(path);
        if (f == null) throw new RuntimeException("Excel file not found: " + path);
        List<List<Object>> grid = fetchExcelGrid(wb, f, c);
        List<String> names = excelColumnNames(grid, c);
        List<List<Object>> data = excelDataRows(grid, c.surface);
        List<FieldMetadata> meta = new ArrayList<>();
        for (int i = 0; i < names.size(); i++) {
            FieldMetadata fm = new FieldMetadata();
            fm.setName(names.get(i));
            fm.setType(inferExcelType(data, i));
            FieldParams params = new FieldParams();
            params.setFieldName(names.get(i));
            fm.setFieldParams(params);
            meta.add(fm);
        }
        return meta;
    }

    /** Fetch the 2D grid for an Excel collection (used by describe + ComputeTask). */
    static List<List<Object>> fetchExcelGrid(SharePointWorkbookReader wb,
                                             SharePointWorkbookReader.FileRef f, ExcelColl c) {
        switch (c.surface) {
            case TABLE: return wb.getTableRange(f, c.entity);
            case SHEET: return wb.getUsedRange(f, c.entity);
            case RANGE: return wb.getNamedRange(f, c.entity);
            default: return Collections.emptyList();
        }
    }

    /**
     * Column names for an Excel collection.
     *  TABLE / SHEET: first grid row is the header.
     *  RANGE: no header — single column takes the range name, multi-column
     *         takes rangeName_1..rangeName_N.
     */
    static List<String> excelColumnNames(List<List<Object>> grid, ExcelColl c) {
        List<String> names = new ArrayList<>();
        if (grid == null || grid.isEmpty()) return names;
        if (c.surface == ExcelSurface.RANGE) {
            int width = 0;
            for (List<Object> row : grid) width = Math.max(width, row.size());
            if (width == 1) {
                names.add(c.entity);
            } else {
                for (int i = 1; i <= width; i++) names.add(c.entity + "_" + i);
            }
            return names;
        }
        // TABLE / SHEET: header row.
        List<Object> header = grid.get(0);
        int i = 0;
        for (Object h : header) {
            i++;
            String name = h != null ? String.valueOf(h).trim() : "";
            names.add(name.isEmpty() ? "column_" + i : name);
        }
        return names;
    }

    /** Data rows (header dropped for TABLE/SHEET; all rows for RANGE). */
    static List<List<Object>> excelDataRows(List<List<Object>> grid, ExcelSurface surface) {
        if (grid == null || grid.isEmpty()) return Collections.emptyList();
        if (surface == ExcelSurface.RANGE) return grid;
        return grid.subList(1, grid.size());
    }

    /**
     * Infer a Thrift FieldType for one Excel column by sampling its data
     * cells. Numbers stay numeric (INTEGER if all integral, else DOUBLE);
     * anything mixed or non-numeric falls back to STRING.
     */
    static FieldType inferExcelType(List<List<Object>> dataRows, int colIdx) {
        boolean any = false, allIntegral = true, numeric = true;
        for (List<Object> row : dataRows) {
            if (colIdx >= row.size()) continue;
            Object v = row.get(colIdx);
            if (v == null) continue;
            any = true;
            if (v instanceof Long || v instanceof Integer) {
                // integral
            } else if (v instanceof Double || v instanceof Float) {
                allIntegral = false;
            } else {
                numeric = false;
                break;
            }
        }
        if (!any || !numeric) return FieldType.STRING;
        return allIntegral ? FieldType.INTEGER : FieldType.DOUBLE;
    }

    static boolean isExcelCollection(String name) {
        return name != null && (name.startsWith(PREFIX_TABLE)
                || name.startsWith(PREFIX_SHEET) || name.startsWith(PREFIX_RANGE));
    }

    static String excelCollectionName(ExcelSurface surface, String fileSlug, String entity) {
        String prefix = surface == ExcelSurface.TABLE ? PREFIX_TABLE
                : surface == ExcelSurface.SHEET ? PREFIX_SHEET : PREFIX_RANGE;
        return prefix + fileSlug + "__" + entity;
    }

    static ExcelColl parseExcelCollection(String collectionName) {
        ExcelSurface surface;
        String prefix;
        if (collectionName.startsWith(PREFIX_TABLE)) { surface = ExcelSurface.TABLE; prefix = PREFIX_TABLE; }
        else if (collectionName.startsWith(PREFIX_SHEET)) { surface = ExcelSurface.SHEET; prefix = PREFIX_SHEET; }
        else if (collectionName.startsWith(PREFIX_RANGE)) { surface = ExcelSurface.RANGE; prefix = PREFIX_RANGE; }
        else return null;
        String rest = collectionName.substring(prefix.length());
        int idx = rest.indexOf("__");
        if (idx < 0) return null;
        ExcelColl c = new ExcelColl();
        c.surface = surface;
        c.fileSlug = rest.substring(0, idx);
        c.entity = rest.substring(idx + 2);
        return c;
    }

    static String excelSlugFromPath(String path) {
        String trimmed = path.replaceAll("^/+", "").replaceAll("\\.xlsx$", "");
        return trimmed.replaceAll("[^A-Za-z0-9]+", "_").toLowerCase();
    }

    static class ExcelColl {
        ExcelSurface surface;
        String fileSlug;
        String entity;
    }

    // ----- File collections (CSV/TSV/JSON, v1.1) ------------------------

    private CollectionInfo discoverFileCollection(String path) {
        if (!SharePointFileReader.isSupported(path)) {
            log.warn("INCLUDE_FILES path '{}' is not a supported type (csv/tsv/json); skipping", path);
            return null;
        }
        CollectionInfo ci = new CollectionInfo();
        ci.setCollection(fileCollectionName(fileSlugFromPath(path)));
        ci.setSchema("default");
        log.debug("Discovered file collection for '{}'", path);
        return ci;
    }

    private List<FieldMetadata> describeFileCollection(SharePointFileReader fr, String collectionName) {
        if (fr == null) {
            throw new RuntimeException("File collection '" + collectionName + "' requested but no file "
                    + "reader is available (INCLUDE_FILES not configured on this connection?)");
        }
        String slug = parseFileCollection(collectionName);
        if (slug == null) throw new RuntimeException("Invalid file collection name: " + collectionName);
        String path = fr.pathForSlug(slug);
        if (path == null) {
            throw new RuntimeException("File for slug '" + slug
                    + "' is not in this connection's INCLUDE_FILES list");
        }
        return schemaFromHeaderGrid(fr.readGrid(path));
    }

    static boolean isFileCollection(String name) {
        return name != null && name.startsWith(PREFIX_FILE);
    }

    static String fileCollectionName(String fileSlug) {
        return PREFIX_FILE + fileSlug;
    }

    /** Returns the file slug, or null if not a file collection name. */
    static String parseFileCollection(String collectionName) {
        return isFileCollection(collectionName)
                ? collectionName.substring(PREFIX_FILE.length()) : null;
    }

    static String fileSlugFromPath(String path) {
        return path.replaceAll("^/+", "").replaceAll("[^A-Za-z0-9]+", "_").toLowerCase();
    }

    // ----- Shared header-grid helpers (files + future) ------------------

    /**
     * Column names from a header-row grid (row 0 = header). Blank headers
     * become column_N. Shared by file collections; Excel tables/sheets use
     * the equivalent logic in excelColumnNames.
     */
    static List<String> headerRowNames(List<List<Object>> grid) {
        List<String> names = new ArrayList<>();
        if (grid == null || grid.isEmpty()) return names;
        List<Object> header = grid.get(0);
        for (int i = 0; i < header.size(); i++) {
            Object h = header.get(i);
            String nm = h != null ? String.valueOf(h).trim() : "";
            names.add(nm.isEmpty() ? "column_" + (i + 1) : nm);
        }
        return names;
    }

    /** Data rows of a header-row grid (everything after row 0). */
    static List<List<Object>> headerGridData(List<List<Object>> grid) {
        if (grid == null || grid.size() <= 1) return Collections.emptyList();
        return grid.subList(1, grid.size());
    }

    static List<FieldMetadata> schemaFromHeaderGrid(List<List<Object>> grid) {
        List<String> names = headerRowNames(grid);
        List<List<Object>> data = headerGridData(grid);
        List<FieldMetadata> meta = new ArrayList<>();
        for (int i = 0; i < names.size(); i++) {
            FieldMetadata fm = new FieldMetadata();
            fm.setName(names.get(i));
            fm.setType(inferExcelType(data, i));
            FieldParams p = new FieldParams();
            p.setFieldName(names.get(i));
            fm.setFieldParams(p);
            meta.add(fm);
        }
        return meta;
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
