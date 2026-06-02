/**
 * Copyright (C) insightsoftware 2026. All rights reserved.
 *
 * Reads Excel-in-SharePoint data via the Microsoft Graph Workbook API,
 * using RAW authenticated REST + Jackson rather than the Kiota typed SDK.
 *
 * Why raw REST and not the typed client:
 *   The Graph Java SDK models workbook cell data (WorkbookRange.getValues(),
 *   WorkbookTableRow.getValues()) as com.microsoft.graph.models.Json, which
 *   only exposes getAdditionalData() (a Map for JSON *objects*). Workbook
 *   values arrive as a 2D JSON *array* — [[header...],[row...]] — which that
 *   type cannot represent. Reading the array through the SDK is unreliable.
 *   The Workbook REST surface is small and stable, so a raw GET + Jackson
 *   parse of the `values` array is both simpler and more robust.
 *
 * One instance is built per EDC request (in SharePointDataProvider.context())
 * and carries: the shared OkHttp client, a bearer token minted once for the
 * request, the resolved siteId, and the INCLUDE_EXCEL slug->path map. It is
 * only constructed when a connection actually configures Excel paths, so
 * List-only connections never mint a Workbook token.
 *
 * Surfaces exposed (each becomes one EDC collection):
 *   - tables       : every named Excel table (header row = column names)
 *   - worksheets   : used-range of each visible sheet (header row = names)
 *   - named ranges : each visible named Range (no header; synthetic names)
 */
package com.zoomdata.connector.example.provider.sharepoint;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import okhttp3.OkHttpClient;
import okhttp3.Request;
import okhttp3.Response;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;

public class SharePointWorkbookReader {

    private static final Logger log = LoggerFactory.getLogger(SharePointWorkbookReader.class);
    private static final ObjectMapper MAPPER = new ObjectMapper();
    private static final String GRAPH_BASE = "https://graph.microsoft.com/v1.0";

    private final OkHttpClient http;
    private final String bearerToken;
    private final String siteId;
    private final Map<String, String> pathsBySlug;

    public SharePointWorkbookReader(OkHttpClient http, String bearerToken, String siteId,
                                    Map<String, String> pathsBySlug) {
        this.http = http;
        this.bearerToken = bearerToken;
        this.siteId = siteId;
        this.pathsBySlug = pathsBySlug != null ? pathsBySlug : Collections.emptyMap();
    }

    /** Resolve an INCLUDE_EXCEL file slug back to its configured path. */
    public String pathForSlug(String slug) {
        return pathsBySlug.get(slug);
    }

    /** Reference to a resolved drive item (an xlsx file). */
    public static final class FileRef {
        final String driveId;
        final String itemId;
        FileRef(String driveId, String itemId) { this.driveId = driveId; this.itemId = itemId; }
    }

    /**
     * Resolve a path within the site's default drive to {driveId, itemId}.
     * Returns null (with a warning) if the file is missing, so discovery can
     * skip a mis-typed path without failing the whole connection.
     */
    public FileRef resolveFile(String path) {
        String normalised = path.startsWith("/") ? path.substring(1) : path;
        // GET /sites/{siteId}/drive/root:/{path}  (path segment-encoded)
        String url = GRAPH_BASE + "/sites/" + encodePathLiteral(siteId)
                + "/drive/root:/" + encodeDrivePath(normalised)
                + "?$select=id,parentReference,file";
        try {
            JsonNode item = getJson(url);
            if (item == null) return null;
            String itemId = text(item, "id");
            JsonNode parent = item.get("parentReference");
            String driveId = parent != null ? text(parent, "driveId") : null;
            if (itemId == null || driveId == null) {
                log.warn("Excel file '{}' resolved but missing id/driveId", path);
                return null;
            }
            if (item.get("file") == null) {
                log.warn("Path '{}' is not a file (folder?); skipping", path);
                return null;
            }
            return new FileRef(driveId, itemId);
        } catch (FileNotFound e) {
            log.warn("Excel file not found at path '{}' in site drive; skipping", path);
            return null;
        }
    }

    private String workbookBase(FileRef f) {
        return GRAPH_BASE + "/drives/" + encodePathLiteral(f.driveId)
                + "/items/" + encodePathLiteral(f.itemId) + "/workbook";
    }

    /** Names of every Excel table in the workbook. */
    public List<String> listTables(FileRef f) {
        return namesFrom(workbookBase(f) + "/tables?$select=name&$top=200", "table");
    }

    /** Names of every visible worksheet. */
    public List<String> listWorksheets(FileRef f) {
        List<String> out = new ArrayList<>();
        JsonNode resp = getJsonQuiet(workbookBase(f) + "/worksheets?$select=name,visibility&$top=200");
        if (resp == null) return out;
        JsonNode value = resp.get("value");
        if (value != null && value.isArray()) {
            for (JsonNode n : value) {
                String vis = text(n, "visibility");
                if (vis != null && !vis.equalsIgnoreCase("Visible")) continue;
                String name = text(n, "name");
                if (name != null && !name.isEmpty()) out.add(name);
            }
        }
        return out;
    }

    /** Names of every visible named Range (type == Range). */
    public List<String> listNamedRanges(FileRef f) {
        List<String> out = new ArrayList<>();
        JsonNode resp = getJsonQuiet(workbookBase(f) + "/names?$select=name,type,visible&$top=200");
        if (resp == null) return out;
        JsonNode value = resp.get("value");
        if (value != null && value.isArray()) {
            for (JsonNode n : value) {
                String type = text(n, "type");
                JsonNode visible = n.get("visible");
                if (visible != null && visible.isBoolean() && !visible.asBoolean()) continue;
                if (type != null && !type.equalsIgnoreCase("Range")) continue;
                String name = text(n, "name");
                if (name != null && !name.isEmpty()) out.add(name);
            }
        }
        return out;
    }

    private List<String> namesFrom(String url, String what) {
        List<String> out = new ArrayList<>();
        JsonNode resp = getJsonQuiet(url);
        if (resp == null) return out;
        JsonNode value = resp.get("value");
        if (value != null && value.isArray()) {
            for (JsonNode n : value) {
                String name = text(n, "name");
                if (name != null && !name.isEmpty()) out.add(name);
            }
        }
        return out;
    }

    /**
     * 2D cell grid. For tables and worksheet used-ranges the first inner list
     * is the header row; for named ranges there is no header.
     */
    public List<List<Object>> getTableRange(FileRef f, String tableName) {
        return gridFromValues(workbookBase(f) + "/tables/" + encodePathSegment(tableName)
                + "/range?$select=values");
    }

    public List<List<Object>> getUsedRange(FileRef f, String sheetName) {
        return gridFromValues(workbookBase(f) + "/worksheets/" + encodePathSegment(sheetName)
                + "/usedRange?$select=values");
    }

    public List<List<Object>> getNamedRange(FileRef f, String rangeName) {
        return gridFromValues(workbookBase(f) + "/names/" + encodePathSegment(rangeName)
                + "/range?$select=values");
    }

    /** Parse the `values` 2D array from a range/usedRange response. */
    private List<List<Object>> gridFromValues(String url) {
        JsonNode resp = getJsonQuiet(url);
        List<List<Object>> grid = new ArrayList<>();
        if (resp == null) return grid;
        JsonNode values = resp.get("values");
        if (values == null || !values.isArray()) return grid;
        for (JsonNode row : values) {
            List<Object> cells = new ArrayList<>();
            if (row.isArray()) {
                for (JsonNode cell : row) {
                    cells.add(scalar(cell));
                }
            }
            grid.add(cells);
        }
        return grid;
    }

    /** Convert a Jackson cell node to a Java scalar preserving number/string/bool. */
    private static Object scalar(JsonNode cell) {
        if (cell == null || cell.isNull()) return null;
        if (cell.isNumber()) {
            // Preserve integral vs fractional so type inference can distinguish.
            if (cell.isIntegralNumber()) return cell.longValue();
            return cell.doubleValue();
        }
        if (cell.isBoolean()) return cell.booleanValue();
        return cell.asText();
    }

    // ----- HTTP plumbing ------------------------------------------------

    private static class FileNotFound extends RuntimeException {}

    private JsonNode getJson(String url) {
        Request req = new Request.Builder().url(url)
                .header("Authorization", "Bearer " + bearerToken)
                .header("Accept", "application/json")
                .get().build();
        try (Response resp = http.newCall(req).execute()) {
            int code = resp.code();
            String body = resp.body() != null ? resp.body().string() : "";
            if (code == 404) throw new FileNotFound();
            if (code < 200 || code >= 300) {
                // Don't echo the body (may carry tenant/identity detail) — log
                // it at debug only, throw a short message. Mirrors the
                // sanitised-error pattern from the v1.0.1 audit fix.
                log.debug("Graph workbook call {} -> HTTP {} body={}", url, code, truncate(body, 300));
                throw new RuntimeException("Graph workbook request failed (HTTP " + code + ")");
            }
            return MAPPER.readTree(body);
        } catch (IOException e) {
            throw new RuntimeException("Graph workbook request I/O error: " + e.getMessage(), e);
        }
    }

    /** Same as getJson but returns null instead of throwing FileNotFound. */
    private JsonNode getJsonQuiet(String url) {
        try {
            return getJson(url);
        } catch (FileNotFound e) {
            return null;
        }
    }

    private static String text(JsonNode node, String field) {
        JsonNode v = node.get(field);
        return v != null && !v.isNull() ? v.asText() : null;
    }

    private static String truncate(String s, int n) {
        if (s == null) return "";
        return s.length() <= n ? s : s.substring(0, n) + "...";
    }

    /** Encode a value used as a single OData path literal (siteId, driveId, itemId). */
    private static String encodePathLiteral(String v) {
        // siteId/driveId/itemId contain commas, exclamation marks etc. Encode
        // reserved chars but keep the value as one path segment.
        return v.replace(" ", "%20");
    }

    /** Encode a name used after a colon path (table/sheet/range names). */
    private static String encodePathSegment(String v) {
        try {
            return URLEncoder.encode(v, StandardCharsets.UTF_8.name()).replace("+", "%20");
        } catch (Exception e) {
            return v;
        }
    }

    /** Encode a drive path like "folder/sub/file.xlsx", preserving slashes. */
    private static String encodeDrivePath(String path) {
        String[] parts = path.split("/");
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < parts.length; i++) {
            if (i > 0) sb.append("/");
            sb.append(encodePathSegment(parts[i]));
        }
        return sb.toString();
    }
}
