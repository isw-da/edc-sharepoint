/**
 * Copyright (C) insightsoftware 2026. All rights reserved.
 *
 * Executes a single fetch against a SharePoint List (or Excel table) and
 * returns rows as Thrift Records.
 *
 * Pushdown to Graph: OData $filter, $select, $top for Lists. Excel tables
 * fetch raw and let the QE filter. AND-composed EDC filters become a
 * conjunction; unsupported filter types are skipped and applied by the QE.
 */
package com.zoomdata.connector.example.provider.sharepoint;

import com.microsoft.graph.models.FieldValueSet;
import com.microsoft.graph.models.ListItem;
import com.microsoft.graph.models.ListItemCollectionResponse;
import com.microsoft.graph.serviceclient.GraphServiceClient;
import com.zoomdata.connector.example.framework.async.Cursor;
import com.zoomdata.connector.example.framework.async.IComputeTask;
import com.zoomdata.gen.edc.filter.Filter;
import com.zoomdata.gen.edc.filter.FilterFunction;
import com.zoomdata.gen.edc.types.Field;
import com.zoomdata.gen.edc.types.FieldType;
import com.zoomdata.gen.edc.types.Record;
import com.zoomdata.gen.edc.types.ResponseMetadata;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.NoSuchElementException;
import java.util.stream.Collectors;

public class SharePointComputeTask implements IComputeTask {

    private static final Logger log = LoggerFactory.getLogger(SharePointComputeTask.class);

    private final GraphServiceClient client;
    private final String siteId;
    private final String collectionName;
    private final List<String> requestedFields;
    private final SharePointTypesMapping typesMapping;
    private final SharePointIntrospector introspector;
    private final int fetchSize;
    private final List<Filter> filters;
    private final SharePointReaders readers; // workbook + file readers; either may be null

    private volatile boolean cancelled = false;
    private volatile double progress = 0.0;
    private List<Record> records;
    private List<ResponseMetadata> metadata;
    private List<String> resolvedFieldOrder;

    public SharePointComputeTask(GraphServiceClient client, String siteId,
                                  String collectionName, List<String> requestedFields,
                                  SharePointTypesMapping typesMapping,
                                  SharePointIntrospector introspector,
                                  int fetchSize, List<Filter> filters,
                                  SharePointReaders readers) {
        this.client = client;
        this.siteId = siteId;
        this.collectionName = collectionName;
        this.requestedFields = requestedFields;
        this.typesMapping = typesMapping;
        this.introspector = introspector;
        this.fetchSize = fetchSize;
        this.filters = filters;
        this.readers = readers;
    }

    @Override
    public Cursor compute() {
        try {
            if (SharePointIntrospector.isExcelCollection(collectionName)) {
                return computeExcel();
            }
            if (SharePointIntrospector.isFileCollection(collectionName)) {
                return computeFile();
            }
            return computeList();
        } catch (Exception e) {
            log.error("SharePoint fetch failed for '{}': {}", collectionName, e.getMessage(), e);
            throw new RuntimeException("SharePoint query failed: " + e.getMessage(), e);
        }
    }

    @Override
    public double progress() { return progress; }

    @Override
    public void cancel() { this.cancelled = true; }

    @Override
    public void close() { /* no resources to release */ }

    // ----- Lists --------------------------------------------------------

    private Cursor computeList() {
        String listId = introspector.resolveListId(client, siteId, collectionName);
        if (listId == null) {
            throw new RuntimeException("List not found: " + collectionName);
        }

        // Resolve field order: explicit request → fields from describe → all
        List<String> fields = resolveFieldOrder();
        this.resolvedFieldOrder = new ArrayList<>(fields);

        String selectExpr = "fields($select=" + String.join(",", fields) + ")";
        String filterExpr = buildODataFilter();
        int top = (fetchSize > 0 && fetchSize <= 5000) ? fetchSize : 5000;

        log.info("Fetching List '{}' (id={}) top={} select={} filter={}",
                collectionName, listId, top, selectExpr, filterExpr);

        ListItemCollectionResponse resp = client.sites().bySiteId(siteId)
                .lists().byListId(listId).items().get(config -> {
                    config.queryParameters.expand = new String[]{selectExpr};
                    config.queryParameters.top = top;
                    if (filterExpr != null) {
                        config.queryParameters.filter = filterExpr;
                    }
                    // Required when filtering on non-indexed columns
                    config.headers.add("Prefer", "HonorNonIndexedQueriesWarningMayFailRandomly");
                });

        records = new ArrayList<>();
        if (resp != null && resp.getValue() != null) {
            int idx = 0;
            int total = resp.getValue().size();
            for (ListItem item : resp.getValue()) {
                if (cancelled) break;
                FieldValueSet fv = item.getFields();
                Map<String, Object> raw = fv != null && fv.getAdditionalData() != null
                        ? fv.getAdditionalData() : Collections.emptyMap();
                records.add(rowToRecord(raw, fields));
                idx++;
                progress = ((double) idx / Math.max(total, 1)) * 100.0;
            }
        }
        metadata = buildMetadata(fields, true);
        progress = 100.0;
        log.info("List '{}' returned {} rows", collectionName, records.size());
        return new SharePointCursor(records, metadata);
    }

    private List<String> resolveFieldOrder() {
        if (requestedFields != null && !requestedFields.isEmpty()
                && !(requestedFields.size() == 1 && "*".equals(requestedFields.get(0)))) {
            return requestedFields;
        }
        // Wildcard: introspect to get all fields
        return introspector.describeCollection(client, siteId, collectionName, readers)
                .stream().map(fm -> fm.getName()).collect(Collectors.toList());
    }

    // ----- Excel (v1.1, via Workbook API) -------------------------------

    private Cursor computeExcel() {
        SharePointWorkbookReader workbook = readers != null ? readers.workbook : null;
        if (workbook == null) {
            throw new RuntimeException("Excel collection '" + collectionName + "' requested but no "
                    + "workbook reader available (INCLUDE_EXCEL not configured on this connection?)");
        }
        SharePointIntrospector.ExcelColl c = SharePointIntrospector.parseExcelCollection(collectionName);
        if (c == null) throw new RuntimeException("Invalid Excel collection name: " + collectionName);
        String path = workbook.pathForSlug(c.fileSlug);
        if (path == null) {
            throw new RuntimeException("Excel file for slug '" + c.fileSlug
                    + "' is not in this connection's INCLUDE_EXCEL list");
        }
        SharePointWorkbookReader.FileRef f = workbook.resolveFile(path);
        if (f == null) throw new RuntimeException("Excel file not found: " + path);

        List<List<Object>> grid = SharePointIntrospector.fetchExcelGrid(workbook, f, c);
        List<String> allColumns = SharePointIntrospector.excelColumnNames(grid, c);
        List<List<Object>> dataRows = SharePointIntrospector.excelDataRows(grid, c.surface);
        return gridCursor(allColumns, dataRows, "Excel");
    }

    // ----- Files (v1.1, CSV/TSV/JSON) -----------------------------------

    private Cursor computeFile() {
        SharePointFileReader fr = readers != null ? readers.files : null;
        if (fr == null) {
            throw new RuntimeException("File collection '" + collectionName + "' requested but no "
                    + "file reader available (INCLUDE_FILES not configured on this connection?)");
        }
        String slug = SharePointIntrospector.parseFileCollection(collectionName);
        if (slug == null) throw new RuntimeException("Invalid file collection name: " + collectionName);
        String path = fr.pathForSlug(slug);
        if (path == null) {
            throw new RuntimeException("File for slug '" + slug
                    + "' is not in this connection's INCLUDE_FILES list");
        }
        List<List<Object>> grid = fr.readGrid(path);
        List<String> allColumns = SharePointIntrospector.headerRowNames(grid);
        List<List<Object>> dataRows = SharePointIntrospector.headerGridData(grid);
        return gridCursor(allColumns, dataRows, "File");
    }

    /**
     * Shared row-building for grid-shaped sources (Excel + files): honour the
     * QE's requested field subset/order, map to grid column indices, render
     * cells as strings, and attach response metadata.
     */
    private Cursor gridCursor(List<String> allColumns, List<List<Object>> dataRows, String kind) {
        List<String> fields = (requestedFields != null && !requestedFields.isEmpty()
                && !(requestedFields.size() == 1 && "*".equals(requestedFields.get(0))))
                ? requestedFields : allColumns;
        this.resolvedFieldOrder = new ArrayList<>(fields);

        int[] colIndex = new int[fields.size()];
        for (int i = 0; i < fields.size(); i++) {
            colIndex[i] = allColumns.indexOf(fields.get(i));
        }

        int cap = (fetchSize > 0 && fetchSize <= 100000) ? fetchSize : 100000;
        records = new ArrayList<>();
        int total = dataRows.size();
        for (int r = 0; r < total && r < cap && !cancelled; r++) {
            List<Object> row = dataRows.get(r);
            Record record = new Record();
            List<Field> out = new ArrayList<>(fields.size());
            for (int ci : colIndex) {
                Field field = new Field();
                Object v = (ci >= 0 && ci < row.size()) ? row.get(ci) : null;
                if (v == null) {
                    field.setIsNull(true);
                    field.setValue("");
                } else {
                    field.setValue(String.valueOf(v));
                }
                out.add(field);
            }
            record.setRecord(out);
            records.add(record);
            progress = ((double) (r + 1) / Math.max(total, 1)) * 100.0;
        }
        metadata = buildMetadata(fields, true);
        progress = 100.0;
        log.info("{} '{}' returned {} row(s) ({} column(s))", kind, collectionName, records.size(), fields.size());
        return new SharePointCursor(records, metadata);
    }

    // ----- Filter pushdown ----------------------------------------------

    private String buildODataFilter() {
        if (filters == null || filters.isEmpty()) return null;
        List<Filter> flat = flattenFilters(filters);
        List<String> clauses = new ArrayList<>();
        for (Filter f : flat) {
            String clause = convertFilter(f);
            if (clause != null) clauses.add(clause);
        }
        return clauses.isEmpty() ? null : String.join(" and ", clauses);
    }

    private List<Filter> flattenFilters(List<Filter> input) {
        List<Filter> out = new ArrayList<>();
        for (Filter f : input) {
            if (f.getType() == FilterFunction.AND && f.getFilterAND() != null
                    && f.getFilterAND().getFilters() != null) {
                out.addAll(flattenFilters(f.getFilterAND().getFilters()));
            } else {
                out.add(f);
            }
        }
        return out;
    }

    private String convertFilter(Filter f) {
        if (f == null || f.getType() == null) return null;
        switch (f.getType()) {
            case EQ:
                return f.getFilterEQ() != null
                        ? path(f.getFilterEQ().getPath()) + " eq " + odataValue(f.getFilterEQ().getValue().getValue(), f.getFilterEQ().getType())
                        : null;
            case GT:
                return f.getFilterGT() != null
                        ? path(f.getFilterGT().getPath()) + " gt " + odataValue(f.getFilterGT().getValue().getValue(), f.getFilterGT().getType())
                        : null;
            case GE:
                return f.getFilterGE() != null
                        ? path(f.getFilterGE().getPath()) + " ge " + odataValue(f.getFilterGE().getValue().getValue(), f.getFilterGE().getType())
                        : null;
            case LT:
                return f.getFilterLT() != null
                        ? path(f.getFilterLT().getPath()) + " lt " + odataValue(f.getFilterLT().getValue().getValue(), f.getFilterLT().getType())
                        : null;
            case LE:
                return f.getFilterLE() != null
                        ? path(f.getFilterLE().getPath()) + " le " + odataValue(f.getFilterLE().getValue().getValue(), f.getFilterLE().getType())
                        : null;
            case CONTAINS:
                return f.getFilterCONTAINS() != null
                        ? "contains(" + path(f.getFilterCONTAINS().getPath()) + ","
                            + odataValue(f.getFilterCONTAINS().getValue().getValue(), FieldType.STRING) + ")"
                        : null;
            case STARTS_WITH:
                return f.getFilterSTARTS_WITH() != null
                        ? "startswith(" + path(f.getFilterSTARTS_WITH().getPath()) + ","
                            + odataValue(f.getFilterSTARTS_WITH().getValue().getValue(), FieldType.STRING) + ")"
                        : null;
            case ENDS_WITH:
                return f.getFilterENDS_WITH() != null
                        ? "endswith(" + path(f.getFilterENDS_WITH().getPath()) + ","
                            + odataValue(f.getFilterENDS_WITH().getValue().getValue(), FieldType.STRING) + ")"
                        : null;
            case IS_NULL:
                return f.getFilterISNULL() != null
                        ? path(f.getFilterISNULL().getPath()) + " eq null"
                        : null;
            default:
                log.debug("Unsupported filter for OData pushdown: {}", f.getType());
                return null;
        }
    }

    /** Items API requires fields/ prefix when filtering by list-item field. */
    private String path(String column) {
        return "fields/" + column;
    }

    private String odataValue(String value, FieldType type) {
        if (value == null) return "null";
        if (type == FieldType.INTEGER || type == FieldType.DOUBLE) {
            return value;
        }
        if (type == FieldType.DATE) {
            return odataDateLiteral(value);
        }
        return "'" + value.replace("'", "''") + "'";
    }

    /**
     * Format a date filter value as an OData dateTime literal (UNQUOTED ISO
     * 8601 in UTC), e.g. fields/DueDate ge 2026-01-01T00:00:00Z. Graph rejects
     * quoted-string comparisons on dateTime list columns, which is why v1.0
     * date filters were weak. Composer sends DATE filter values as epoch
     * milliseconds; we also accept an already-ISO value and pass it through.
     */
    static String odataDateLiteral(String value) {
        String v = value.trim();
        try {
            long ms = Long.parseLong(v);
            return java.time.Instant.ofEpochMilli(ms)
                    .atZone(java.time.ZoneOffset.UTC)
                    .format(java.time.format.DateTimeFormatter.ofPattern("yyyy-MM-dd'T'HH:mm:ss'Z'"));
        } catch (NumberFormatException notEpoch) {
            // Already an ISO 8601 string (possibly quoted) — strip quotes, keep as-is.
            if (v.length() >= 2 && v.startsWith("'") && v.endsWith("'")) {
                v = v.substring(1, v.length() - 1);
            }
            return v;
        }
    }

    // ----- Record building ----------------------------------------------

    private Record rowToRecord(Map<String, Object> raw, List<String> fields) {
        Record record = new Record();
        List<Field> out = new ArrayList<>();
        for (String f : fields) {
            Object v = raw.get(f);
            Field field = new Field();
            if (v == null) {
                field.setIsNull(true);
                field.setValue("");
            } else {
                field.setValue(String.valueOf(v));
            }
            out.add(field);
        }
        record.setRecord(out);
        return record;
    }

    private List<ResponseMetadata> buildMetadata(List<String> fields, boolean useDescribe) {
        List<ResponseMetadata> out = new ArrayList<>();
        Map<String, FieldType> declared = new HashMap<>();
        if (useDescribe) {
            try {
                introspector.describeCollection(client, siteId, collectionName, readers)
                        .forEach(fm -> declared.put(fm.getName(), fm.getType()));
            } catch (Exception e) {
                log.debug("Type lookup for metadata failed; falling back to STRING: {}", e.getMessage());
            }
        }
        for (String f : fields) {
            ResponseMetadata rm = new ResponseMetadata();
            rm.setName(f);
            rm.setType(declared.getOrDefault(f, FieldType.STRING));
            out.add(rm);
        }
        return out;
    }

    /** In-memory cursor over the fetched records. */
    static class SharePointCursor implements Cursor {
        private final List<Record> records;
        private final List<ResponseMetadata> metadata;
        private final Iterator<Record> iterator;

        SharePointCursor(List<Record> records, List<ResponseMetadata> metadata) {
            this.records = records;
            this.metadata = metadata;
            this.iterator = records.iterator();
        }

        @Override public List<ResponseMetadata> getMetadata() { return metadata; }
        @Override public boolean hasNextBatch() { return false; }
        @Override public boolean hasNext() { return iterator.hasNext(); }
        @Override public Record next() {
            if (!hasNext()) throw new NoSuchElementException();
            return iterator.next();
        }
    }
}
