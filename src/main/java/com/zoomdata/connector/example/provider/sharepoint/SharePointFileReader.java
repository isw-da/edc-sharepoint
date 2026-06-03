/**
 * Copyright (C) insightsoftware 2026. All rights reserved.
 *
 * Reads tabular data files (CSV / TSV / JSON) from a SharePoint drive via
 * the Graph drive content API, returning a 2D grid (first row = column
 * names, rest = data) just like SharePointWorkbookReader does for Excel.
 *
 * Scope (v1.1):
 *   - CSV / TSV: header row + data rows. Column type is decided per column
 *     (all-integral -> Long, all-numeric -> Double, else String) so the
 *     downstream type inference and record rendering behave like Excel.
 *   - JSON: a top-level array of flat objects, OR newline-delimited objects
 *     (NDJSON), OR a single object. Object keys become columns (union,
 *     first-seen order). Nested objects/arrays in a value are stringified.
 *
 * One instance per EDC request, built only when INCLUDE_FILES is set;
 * shares the OkHttp client + bearer token with the rest of the connector.
 */
package com.zoomdata.connector.example.provider.sharepoint;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.MappingIterator;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.dataformat.csv.CsvMapper;
import com.fasterxml.jackson.dataformat.csv.CsvParser;
import com.fasterxml.jackson.dataformat.csv.CsvSchema;
import okhttp3.OkHttpClient;
import okhttp3.Request;
import okhttp3.Response;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collection;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;

public class SharePointFileReader {

    private static final Logger log = LoggerFactory.getLogger(SharePointFileReader.class);
    private static final ObjectMapper JSON = new ObjectMapper();
    private static final CsvMapper CSV = new CsvMapper();
    private static final String GRAPH_BASE = "https://graph.microsoft.com/v1.0";

    private final OkHttpClient http;
    private final String bearerToken;
    // "sites/{siteId}/drive" (SharePoint) or "users/{upn}/drive" (OneDrive).
    private final String driveResourcePath;
    private final Map<String, String> pathsBySlug;

    public SharePointFileReader(OkHttpClient http, String bearerToken, String driveResourcePath,
                                Map<String, String> pathsBySlug) {
        this.http = http;
        this.bearerToken = bearerToken;
        this.driveResourcePath = driveResourcePath;
        this.pathsBySlug = pathsBySlug != null ? pathsBySlug : java.util.Collections.emptyMap();
    }

    public String pathForSlug(String slug) {
        return pathsBySlug.get(slug);
    }

    public Collection<String> configuredPaths() {
        return pathsBySlug.values();
    }

    /** True if the path has an extension this reader can parse. */
    public static boolean isSupported(String path) {
        String e = ext(path);
        return e.equals("csv") || e.equals("tsv") || e.equals("json");
    }

    /**
     * Download + parse the file at the given drive path into a 2D grid.
     * First inner list is the header (column names); the rest are data rows.
     */
    public List<List<Object>> readGrid(String path) {
        String e = ext(path);
        try (InputStream in = download(path)) {
            switch (e) {
                case "csv": return parseDelimited(in, ',');
                case "tsv": return parseDelimited(in, '\t');
                case "json": return parseJson(in);
                default:
                    throw new RuntimeException("Unsupported file type ." + e + " for " + path
                            + " (supported: csv, tsv, json)");
            }
        } catch (IOException ex) {
            throw new RuntimeException("Failed to read file " + path + ": " + ex.getMessage(), ex);
        }
    }

    // ----- download -----------------------------------------------------

    private InputStream download(String path) throws IOException {
        String normalised = path.startsWith("/") ? path.substring(1) : path;
        String url = GRAPH_BASE + "/" + driveResourcePath
                + "/root:/" + encodeDrivePath(normalised) + ":/content";
        Request req = new Request.Builder().url(url)
                .header("Authorization", "Bearer " + bearerToken)
                .get().build();
        Response resp = http.newCall(req).execute(); // OkHttp follows the content redirect
        int code = resp.code();
        if (code < 200 || code >= 300) {
            String body = resp.body() != null ? resp.body().string() : "";
            resp.close();
            log.debug("Drive content {} -> HTTP {} body={}", url, code, truncate(body, 200));
            if (code == 404) throw new IOException("file not found");
            throw new IOException("drive content request failed (HTTP " + code + ")");
        }
        if (resp.body() == null) {
            resp.close();
            throw new IOException("empty response body");
        }
        return resp.body().byteStream();
    }

    // ----- CSV / TSV ----------------------------------------------------

    private List<List<Object>> parseDelimited(InputStream in, char sep) throws IOException {
        // WRAP_AS_ARRAY: read each CSV row as a String[] without a predeclared
        // column schema (an empty schema otherwise rejects "extra" columns
        // against its 0 columns). First row is treated as the header below.
        CsvSchema schema = CsvSchema.emptySchema().withColumnSeparator(sep);
        List<String[]> rows;
        try (MappingIterator<String[]> it = CSV.readerFor(String[].class)
                .with(schema).with(CsvParser.Feature.WRAP_AS_ARRAY).readValues(in)) {
            rows = it.readAll();
        }
        List<List<Object>> grid = new ArrayList<>();
        if (rows.isEmpty()) return grid;

        // Header row (kept as strings = column names).
        String[] header = rows.get(0);
        List<Object> headerRow = new ArrayList<>();
        for (String h : header) headerRow.add(h);
        grid.add(headerRow);
        int cols = header.length;

        // Decide each column's type from ALL data cells (column-level, so a
        // single non-numeric cell keeps the whole column STRING — avoids the
        // leading-zero/ID coercion footgun).
        boolean[] colAllInt = new boolean[cols];
        boolean[] colAllNum = new boolean[cols];
        boolean[] colAny = new boolean[cols];
        for (int c = 0; c < cols; c++) { colAllInt[c] = true; colAllNum[c] = true; }
        for (int r = 1; r < rows.size(); r++) {
            String[] row = rows.get(r);
            for (int c = 0; c < cols; c++) {
                String v = c < row.length ? row[c] : null;
                if (v == null || v.isEmpty()) continue;
                colAny[c] = true;
                // A leading-zero integer (e.g. "00123") is almost certainly a
                // code/ID/ZIP, not a number — coercing it to 123 would silently
                // drop the leading zeros. Force the whole column to STRING.
                if (isLeadingZeroInt(v)) { colAllInt[c] = false; colAllNum[c] = false; continue; }
                if (!isLong(v)) colAllInt[c] = false;
                if (!isDouble(v)) colAllNum[c] = false;
            }
        }

        for (int r = 1; r < rows.size(); r++) {
            String[] row = rows.get(r);
            List<Object> out = new ArrayList<>(cols);
            for (int c = 0; c < cols; c++) {
                String v = c < row.length ? row[c] : null;
                if (v == null || v.isEmpty()) { out.add(null); continue; }
                if (colAny[c] && colAllInt[c]) out.add(Long.parseLong(v.trim()));
                else if (colAny[c] && colAllNum[c]) out.add(Double.parseDouble(v.trim()));
                else out.add(v);
            }
            grid.add(out);
        }
        return grid;
    }

    private static boolean isLong(String s) {
        try { Long.parseLong(s.trim()); return true; } catch (NumberFormatException e) { return false; }
    }

    /** True for "00123"-style values: leading zero, all digits, length > 1. */
    private static boolean isLeadingZeroInt(String s) {
        String t = s.trim();
        if (t.length() < 2 || t.charAt(0) != '0') return false;
        for (int i = 0; i < t.length(); i++) if (!Character.isDigit(t.charAt(i))) return false;
        return true;
    }

    private static boolean isDouble(String s) {
        try { Double.parseDouble(s.trim()); return true; } catch (NumberFormatException e) { return false; }
    }

    // ----- JSON ---------------------------------------------------------

    private List<List<Object>> parseJson(InputStream in) throws IOException {
        // readValues handles a top-level array of objects AND newline-delimited
        // / whitespace-separated objects (NDJSON) AND a single object.
        List<Map<String, Object>> records = new ArrayList<>();
        @SuppressWarnings("unchecked")
        MappingIterator<Map<String, Object>> it =
                (MappingIterator<Map<String, Object>>) (MappingIterator<?>)
                        JSON.readerFor(Map.class).readValues(in);
        try {
            while (it.hasNext()) records.add(it.next());
        } finally {
            it.close();
        }

        // Column order: union of keys, first-seen order.
        LinkedHashSet<String> columns = new LinkedHashSet<>();
        for (Map<String, Object> rec : records) columns.addAll(rec.keySet());

        List<List<Object>> grid = new ArrayList<>();
        List<Object> header = new ArrayList<>(columns);
        grid.add(header);
        for (Map<String, Object> rec : records) {
            List<Object> row = new ArrayList<>(columns.size());
            for (String col : columns) row.add(flatten(rec.get(col)));
            grid.add(row);
        }
        return grid;
    }

    /** Keep scalars; stringify nested objects/arrays so a cell stays atomic. */
    private Object flatten(Object v) {
        if (v == null) return null;
        if (v instanceof Map || v instanceof List) {
            try { return JSON.writeValueAsString(v); } catch (Exception e) { return String.valueOf(v); }
        }
        if (v instanceof Integer) return ((Integer) v).longValue(); // normalise to Long like CSV
        return v; // Long, Double, Boolean, String
    }

    // ----- helpers ------------------------------------------------------

    private static String ext(String path) {
        String p = path;
        int slash = p.lastIndexOf('/');
        if (slash >= 0) p = p.substring(slash + 1);
        int dot = p.lastIndexOf('.');
        return dot >= 0 ? p.substring(dot + 1).toLowerCase() : "";
    }

    private static String encodeDrivePath(String path) {
        String[] parts = path.split("/");
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < parts.length; i++) {
            if (i > 0) sb.append("/");
            try {
                sb.append(java.net.URLEncoder.encode(parts[i], "UTF-8").replace("+", "%20"));
            } catch (Exception e) {
                sb.append(parts[i]);
            }
        }
        return sb.toString();
    }

    private static String truncate(String s, int n) {
        if (s == null) return "";
        return s.length() <= n ? s : s.substring(0, n) + "...";
    }
}
