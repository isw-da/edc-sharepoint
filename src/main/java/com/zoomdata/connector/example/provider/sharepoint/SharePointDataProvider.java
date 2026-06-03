/**
 * Copyright (C) insightsoftware 2026. All rights reserved.
 *
 * SharePoint EDC data provider. Connects Composer / Simba Intelligence to
 * SharePoint Online via Microsoft Graph using an app registration
 * (client-credentials flow).
 *
 * One connection scopes to one SharePoint site (SITE_URL). All visible
 * Lists in that site are exposed as EDC collections; Excel-in-SharePoint
 * tables are also exposed when INCLUDE_EXCEL is populated.
 */
package com.zoomdata.connector.example.provider.sharepoint;

import com.microsoft.graph.serviceclient.GraphServiceClient;
import com.zoomdata.connector.example.framework.annotation.Connector;
import com.zoomdata.connector.example.framework.api.AbstractDataProvider;
import com.zoomdata.connector.example.framework.api.IDescriptionProvider;
import com.zoomdata.connector.example.framework.async.IComputeTaskFactory;
import com.zoomdata.connector.example.framework.provider.serverdescription.GenericDescriptionProvider;
import com.zoomdata.gen.edc.filter.Filter;
import com.zoomdata.gen.edc.request.CollectionInfo;
import com.zoomdata.gen.edc.request.DataReadRequest;
import com.zoomdata.gen.edc.request.ExecuteCommandRequest;
import com.zoomdata.gen.edc.request.ExecuteCommandResponse;
import com.zoomdata.gen.edc.request.ExecuteException;
import com.zoomdata.gen.edc.request.MetaCollectionsResponse;
import com.zoomdata.gen.edc.request.MetaDescribeRequest;
import com.zoomdata.gen.edc.request.MetaDescribeResponse;
import com.zoomdata.gen.edc.request.MetaDescribeSchemaRequest;
import com.zoomdata.gen.edc.request.MetaDescribeSchemaResponse;
import com.zoomdata.gen.edc.request.MetaSchemasRequest;
import com.zoomdata.gen.edc.request.MetaSchemasResponse;
import com.zoomdata.gen.edc.request.MetaCollectionsRequest;
import com.zoomdata.gen.edc.request.RequestInfo;
import com.zoomdata.gen.edc.request.SampleRequest;
import com.zoomdata.gen.edc.request.SampleResponse;
import com.zoomdata.gen.edc.request.Schema;
import com.zoomdata.gen.edc.request.ServerInfoRequest;
import com.zoomdata.gen.edc.request.ServerInfoResponse;
import com.zoomdata.gen.edc.request.ValidateCollectionRequest;
import com.zoomdata.gen.edc.request.ValidateCollectionResponse;
import com.zoomdata.gen.edc.request.ValidateSourceRequest;
import com.zoomdata.gen.edc.request.ValidateSourceResponse;
import com.zoomdata.gen.edc.types.Field;
import com.zoomdata.gen.edc.types.FieldMetadata;
import com.zoomdata.gen.edc.types.SampleField;
import com.zoomdata.gen.edc.types.SampleRecord;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.autoconfigure.condition.ConditionalOnProperty;

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

import static com.zoomdata.connector.example.common.utils.metadatabuilders.ResponseInfoBuilder.ok;
import static com.zoomdata.connector.example.common.utils.metadatabuilders.ResponseInfoBuilder.serverError;
import static com.zoomdata.connector.example.framework.provider.serverdescription.connectionparameters.impl.PasswordConnectionParameter.PasswordConnectionParameterBuilder.passwordParameter;
import static com.zoomdata.connector.example.framework.provider.serverdescription.connectionparameters.impl.StringConnectionParameter.StringConnectionParameterBuilder.stringParameter;
import static com.zoomdata.connector.example.provider.sharepoint.SharePointDataProvider.CONNECTION_TYPE;

// One connector per server process: the Zoomdata framework only auto-selects
// a provider when exactly one @Connector bean exists; with two it demands a
// CONNECTOR_TYPE param on every request (breaking connections that lack it).
// So each provider is gated by the `edc.connector` property (env EDC_CONNECTOR).
// Default (unset) = SharePoint only, preserving v1 behaviour and existing
// connections. The OneDrive pod sets EDC_CONNECTOR=SHAREPOINT_ONEDRIVE.
@Connector(CONNECTION_TYPE)
@ConditionalOnProperty(name = "edc.connector", havingValue = "SHAREPOINT", matchIfMissing = true)
public class SharePointDataProvider extends AbstractDataProvider {

    private static final Logger log = LoggerFactory.getLogger(SharePointDataProvider.class);

    protected static final String CONNECTION_TYPE = "SHAREPOINT";

    protected static final String PARAM_TENANT_ID = "TENANT_ID";
    protected static final String PARAM_CLIENT_ID = "CLIENT_ID";
    protected static final String PARAM_CLIENT_SECRET = "CLIENT_SECRET";
    protected static final String PARAM_SITE_URL = "SITE_URL";   // single-site (v1) — alias for SITE_URLS
    protected static final String PARAM_SITE_URLS = "SITE_URLS";  // v1.1 multi-site, comma-separated
    protected static final String PARAM_INCLUDE_LISTS = "INCLUDE_LISTS";
    protected static final String PARAM_INCLUDE_EXCEL = "INCLUDE_EXCEL";
    protected static final String PARAM_INCLUDE_FILES = "INCLUDE_FILES";  // v1.1 CSV/TSV/JSON
    protected static final String PARAM_AUTHORITY = "AUTHORITY";
    protected static final String PARAM_AUTH_MODE = "AUTH_MODE";          // v1.1: CLIENT_SECRET|CLIENT_CERTIFICATE|MANAGED_IDENTITY
    protected static final String PARAM_CLIENT_CERT_PEM = "CLIENT_CERT_PEM"; // v1.1: PEM (key+cert) for CLIENT_CERTIFICATE

    // protected so the OneDrive subclass (SHAREPOINT_ONEDRIVE) can reuse them.
    protected final SharePointTypesMapping typesMapping = new SharePointTypesMapping();
    protected final SharePointFeatures features = new SharePointFeatures();
    protected final SharePointGraphClient graphFactory = new SharePointGraphClient();
    protected final SharePointIntrospector introspector = new SharePointIntrospector(typesMapping);

    @Override
    public ValidateSourceResponse pingSource(ValidateSourceRequest request) {
        try {
            ConnContext ctx = context(request.getRequestInfo());
            // Validate every configured container; fail with the first unreachable one.
            for (SiteCtx s : ctx.sites) {
                if (!validateContainer(ctx, s)) {
                    return new ValidateSourceResponse(
                            serverError("Failed to connect to container: " + s.siteUrl));
                }
            }
            return new ValidateSourceResponse(ok());
        } catch (Exception e) {
            log.error("pingSource failed: {}", e.getMessage());
            return new ValidateSourceResponse(serverError(safeMessage(e)));
        }
    }

    @Override
    public ValidateCollectionResponse pingCollection(ValidateCollectionRequest request) {
        try {
            ConnContext ctx = context(request.getRequestInfo());
            String collectionName = request.getCollectionInfo().getCollection();
            List<CollectionInfo> collections = allCollections(ctx);
            boolean found = collections.stream().anyMatch(c -> c.getCollection().equals(collectionName));
            return new ValidateCollectionResponse(found ? ok() : serverError("Collection not found: " + collectionName));
        } catch (Exception e) {
            log.error("pingCollection failed: {}", e.getMessage());
            return new ValidateCollectionResponse(serverError(safeMessage(e)));
        }
    }

    @Override
    public ServerInfoResponse info(ServerInfoRequest request) {
        Map<String, String> all = features.getAllFeatures();
        List<String> keys = request.getKeys();
        Map<String, String> result;
        if (keys != null && keys.size() == 1 && "*".equals(keys.get(0))) {
            result = new HashMap<>(all);
        } else if (keys != null) {
            result = new HashMap<>();
            for (String k : keys) result.put(k, all.getOrDefault(k, "UNKNOWN"));
        } else {
            result = new HashMap<>(all);
        }
        return new ServerInfoResponse(result, ok());
    }

    @Override
    public ExecuteCommandResponse executeCommand(ExecuteCommandRequest request) {
        return new ExecuteCommandResponse(ok());
    }

    @Override
    public MetaSchemasResponse schemas(MetaSchemasRequest request) {
        return new MetaSchemasResponse(Collections.singletonList("default"), ok());
    }

    @Override
    public MetaCollectionsResponse collections(MetaCollectionsRequest request) {
        try {
            ConnContext ctx = context(request.getRequestInfo());
            List<CollectionInfo> collections = allCollections(ctx);
            return new MetaCollectionsResponse(collections, ok());
        } catch (Exception e) {
            log.error("collections failed: {}", e.getMessage());
            return new MetaCollectionsResponse(Collections.emptyList(), serverError(safeMessage(e)));
        }
    }

    @Override
    public MetaDescribeResponse describe(MetaDescribeRequest request) {
        try {
            ConnContext ctx = context(request.getRequestInfo());
            String collectionName = request.getCollectionInfo().getCollection();
            Resolved r = resolve(ctx, collectionName);
            List<FieldMetadata> fields = introspector.describeCollection(ctx.client, r.siteId, r.entity, r.readers);
            return new MetaDescribeResponse(fields, ok());
        } catch (Exception e) {
            log.error("describe failed: {}", e.getMessage());
            return new MetaDescribeResponse(Collections.emptyList(), serverError(safeMessage(e)));
        }
    }

    @Override
    public MetaDescribeSchemaResponse describeSchemas(MetaDescribeSchemaRequest request) {
        try {
            ConnContext ctx = context(request.getRequestInfo());
            List<CollectionInfo> all = allCollections(ctx);

            List<CollectionInfo> requested = request.getCollections();
            if (requested != null && !requested.isEmpty()) {
                List<String> requestedNames = requested.stream()
                        .map(CollectionInfo::getCollection).collect(Collectors.toList());
                all = all.stream().filter(c -> requestedNames.contains(c.getCollection()))
                        .collect(Collectors.toList());
            }

            List<CollectionInfo> enriched = new ArrayList<>();
            for (CollectionInfo ci : all) {
                try {
                    Resolved r = resolve(ctx, ci.getCollection());
                    List<FieldMetadata> fields = introspector.describeCollection(
                            ctx.client, r.siteId, r.entity, r.readers);
                    CollectionInfo e2 = new CollectionInfo();
                    e2.setCollection(ci.getCollection());
                    e2.setSchema("default");
                    e2.setFields(fields);
                    enriched.add(e2);
                } catch (Exception ex) {
                    log.warn("Failed to describe '{}': {}", ci.getCollection(), ex.getMessage());
                }
            }

            Schema schema = new Schema("default");
            schema.setCollections(enriched);
            List<Schema> schemas = Collections.singletonList(schema);
            log.info("describeSchemas returning {} collections", enriched.size());
            return new MetaDescribeSchemaResponse(schemas, ok());
        } catch (Exception e) {
            log.error("describeSchemas failed: {}", e.getMessage());
            return new MetaDescribeSchemaResponse(Collections.emptyList(), serverError(safeMessage(e)));
        }
    }

    @Override
    public SampleResponse sample(SampleRequest request) {
        try {
            ConnContext ctx = context(request.getRequestInfo());
            String collectionName = request.getCollectionInfo().getCollection();
            Resolved r = resolve(ctx, collectionName);
            List<FieldMetadata> fieldMeta = introspector.describeCollection(ctx.client, r.siteId, r.entity, r.readers);
            List<String> fieldNames = fieldMeta.stream().map(FieldMetadata::getName).collect(Collectors.toList());

            SharePointComputeTask task = new SharePointComputeTask(
                    ctx.client, r.siteId, r.entity, fieldNames,
                    typesMapping, introspector, 10, null, r.readers);
            SharePointComputeTask.SharePointCursor cursor =
                    (SharePointComputeTask.SharePointCursor) task.compute();

            List<SampleRecord> samples = new ArrayList<>();
            int count = 0;
            while (cursor.hasNext() && count < 10) {
                com.zoomdata.gen.edc.types.Record record = cursor.next();
                List<SampleField> sampleFields = new ArrayList<>();
                List<Field> recFields = record.getRecord();
                for (int i = 0; i < recFields.size(); i++) {
                    Field f = recFields.get(i);
                    String name = (i < fieldNames.size()) ? fieldNames.get(i) : "field_" + i;
                    SampleField sf = new SampleField(name);
                    if (i < fieldMeta.size()) sf.setType(fieldMeta.get(i).getType());
                    if (f.isIsNull()) sf.setIsNull(true);
                    else sf.setValue(f.getValue());
                    sampleFields.add(sf);
                }
                samples.add(new SampleRecord(sampleFields));
                count++;
            }
            return new SampleResponse(samples, ok());
        } catch (Exception e) {
            log.error("sample failed: {}", e.getMessage());
            return new SampleResponse(Collections.emptyList(), serverError(safeMessage(e)));
        }
    }

    @Override
    protected IComputeTaskFactory createComputeTaskFactory(DataReadRequest request, int fetchSize)
            throws ExecuteException {
        try {
            ConnContext ctx = context(request.getRequestInfo());

            String collectionName;
            if (request.getStructured() != null && request.getStructured().getCollectionInfo() != null) {
                collectionName = request.getStructured().getCollectionInfo().getCollection();
            } else {
                throw new IllegalArgumentException("No collection info in request");
            }
            Resolved r = resolve(ctx, collectionName);

            List<String> requestedFields = extractFieldsFromRequest(request);

            List<Filter> filters = null;
            if (request.getStructured() != null) {
                if (request.getStructured().getRawDataRequest() != null
                        && request.getStructured().getRawDataRequest().getFilters() != null) {
                    filters = request.getStructured().getRawDataRequest().getFilters();
                } else if (request.getStructured().getAggDataRequest() != null
                        && request.getStructured().getAggDataRequest().getFilters() != null) {
                    filters = request.getStructured().getAggDataRequest().getFilters();
                }
            }

            return new SharePointComputeTaskFactory(ctx.client, r.siteId, r.entity,
                    requestedFields, typesMapping, introspector, fetchSize, filters, r.readers);
        } catch (Exception e) {
            throw new ExecuteException("Failed to create compute task: " + safeMessage(e));
        }
    }

    /**
     * Build a message safe to return to Composer / surface in admin UI.
     *
     * Graph SDK exceptions can carry the request body, Authorization
     * header fragments, the failing principal, or tenant identifiers in
     * their message. Returning those verbatim to a Composer admin (and
     * thereby into Composer's audit logs) leaks secrets and identity
     * detail unnecessarily.
     *
     * Our own IllegalArgumentExceptions are intended for the operator and
     * are passed through. Everything else is reduced to a category +
     * status code where one is recoverable; full stack lands in the pod
     * log only.
     */
    private String safeMessage(Throwable e) {
        if (e == null) return "SharePoint connector error";
        log.error("Connector error (full detail): {}", e.toString(), e);
        if (e instanceof IllegalArgumentException) {
            // Our own validation — message is operator-written, safe to surface.
            return e.getMessage() != null ? e.getMessage() : "Invalid argument";
        }
        // Try to extract a Graph status code without echoing the raw message.
        String className = e.getClass().getSimpleName();
        String msg = e.getMessage();
        if (msg != null) {
            // Conservative pattern: surface a 3-digit status if present at the start.
            java.util.regex.Matcher m = java.util.regex.Pattern.compile("\\b([45][0-9]{2})\\b").matcher(msg);
            if (m.find()) {
                return "SharePoint Graph error (HTTP " + m.group(1) + "); see pod logs for detail";
            }
        }
        return "SharePoint connector internal error (" + className + "); see pod logs for detail";
    }

    @Override
    protected IDescriptionProvider createDescriptionProvider() {
        return new GenericDescriptionProvider(CONNECTION_TYPE)
                // Connector icon shown in the Composer / SI connection-type
                // picker. Original SharePoint-teal mark (not Microsoft's
                // trademarked logo) at classpath root. Leading slash =
                // resolve from classpath root regardless of caller package.
                .svgIcon("/sharepoint-icon.svg")
                .addParameters(
                        stringParameter(PARAM_AUTH_MODE)
                                .description("CLIENT_SECRET (default), CLIENT_CERTIFICATE, or MANAGED_IDENTITY"))
                .addParameters(
                        stringParameter(PARAM_TENANT_ID)
                                .description("Azure AD tenant ID (GUID or domain). Required for CLIENT_SECRET/CLIENT_CERTIFICATE."))
                .addParameters(
                        stringParameter(PARAM_CLIENT_ID)
                                .description("App registration client (application) ID. Required except for system-assigned MANAGED_IDENTITY."))
                .addParameters(
                        passwordParameter(PARAM_CLIENT_SECRET)
                                .description("App registration client secret. Required for AUTH_MODE=CLIENT_SECRET."))
                .addParameters(
                        passwordParameter(PARAM_CLIENT_CERT_PEM)
                                .description("PEM contents (private key + certificate). Required for AUTH_MODE=CLIENT_CERTIFICATE."))
                .addParameters(
                        stringParameter(PARAM_SITE_URL)
                                .description("Single SharePoint site URL, e.g. "
                                        + "https://contoso.sharepoint.com/sites/sales. Use this OR SITE_URLS."))
                .addParameters(
                        stringParameter(PARAM_SITE_URLS)
                                .description("Comma-separated SharePoint site URLs for a multi-site connection. "
                                        + "When >1 site, collection names are prefixed with the site slug. "
                                        + "Takes precedence over SITE_URL."))
                .addParameters(
                        stringParameter(PARAM_INCLUDE_LISTS)
                                .description("Comma-separated allowlist of List displayNames. Empty = all visible Lists."))
                .addParameters(
                        stringParameter(PARAM_INCLUDE_FILES)
                                .description("Comma-separated paths to CSV/TSV/JSON files in the site drive "
                                        + "(e.g. /data/agents.csv, /exports/savings.json). Each becomes a collection."))
                .addParameters(
                        stringParameter(PARAM_INCLUDE_EXCEL)
                                .description("Comma-separated paths to .xlsx files in the site drive. Empty = no Excel."))
                .addParameters(
                        stringParameter(PARAM_AUTHORITY)
                                .description("Azure AD authority host override (sovereign clouds). Default https://login.microsoftonline.com"))
                .minVersion("1.0")
                .maxVersion("1.0");
    }

    private List<String> extractFieldsFromRequest(DataReadRequest request) {
        if (request.getStructured() == null) return null;
        com.zoomdata.gen.edc.request.StructuredRequest sr = request.getStructured();

        if (sr.getFieldMetadata() != null && !sr.getFieldMetadata().isEmpty()) {
            log.info("fieldMetadata present ({} fields: {}), returning all fields",
                    sr.getFieldMetadata().size(), sr.getFieldMetadata().keySet());
        }
        if (sr.getRawDataRequest() != null
                && sr.getRawDataRequest().getFields() != null
                && !sr.getRawDataRequest().getFields().isEmpty()) {
            return sr.getRawDataRequest().getFields();
        }
        if (sr.getStatsDataRequest() != null
                && sr.getStatsDataRequest().getStatFields() != null) {
            List<String> fields = new ArrayList<>();
            for (com.zoomdata.gen.edc.request.StatField sf : sr.getStatsDataRequest().getStatFields()) {
                fields.add(sf.getField());
            }
            return fields;
        }
        return null;
    }

    // ----- Connection context + multi-site routing ----------------------

    /**
     * Bundle of the GraphServiceClient + per-site contexts + parsed include
     * lists for one EDC call. Built once per request from the connection
     * params.
     *
     * Multi-site (v1.1): SITE_URLS is a comma-separated list; SITE_URL is the
     * single-site alias (if both set, SITE_URLS wins and SITE_URL is ignored
     * with a warning). When more than one site is configured, collection
     * names are prefixed with "<site-slug>__" to disambiguate; single-site
     * connections keep the exact v1 names.
     */
    private ConnContext context(RequestInfo info) {
        // Auth params are read loosely; SharePointGraphClient.credential()
        // enforces what each AUTH_MODE requires (e.g. MANAGED_IDENTITY needs
        // no secret/tenant). Default mode = CLIENT_SECRET (v1 behaviour).
        String authMode = param(info, PARAM_AUTH_MODE, false);
        String tenantId = param(info, PARAM_TENANT_ID, false);
        String clientId = param(info, PARAM_CLIENT_ID, false);
        String clientSecret = param(info, PARAM_CLIENT_SECRET, false);
        String clientCertPem = param(info, PARAM_CLIENT_CERT_PEM, false);
        String authority = param(info, PARAM_AUTHORITY, false);
        String includeLists = param(info, PARAM_INCLUDE_LISTS, false);
        String includeExcel = param(info, PARAM_INCLUDE_EXCEL, false);
        String includeFiles = param(info, PARAM_INCLUDE_FILES, false);

        SharePointGraphClient.GraphAuth auth = new SharePointGraphClient.GraphAuth(
                tenantId, clientId, SharePointGraphClient.GraphAuth.parseMode(authMode),
                clientSecret, clientCertPem, authority);

        // Container entries (site URLs for SharePoint; user UPNs for OneDrive)
        // and the per-entry SiteCtx are produced by overridable seams.
        List<String> entries = resolveContainerEntries(info);

        GraphServiceClient client = graphFactory.build(auth);

        ConnContext ctx = new ConnContext();
        ctx.client = client;
        ctx.includeLists = SharePointIntrospector.parseAllowlist(includeLists);

        Map<String, String> excelPathsBySlug = new HashMap<>();
        for (String p : SharePointIntrospector.parsePaths(includeExcel)) {
            excelPathsBySlug.put(SharePointIntrospector.excelSlugFromPath(p), p);
        }
        Map<String, String> filePathsBySlug = new HashMap<>();
        for (String p : SharePointIntrospector.parsePaths(includeFiles)) {
            filePathsBySlug.put(SharePointIntrospector.fileSlugFromPath(p), p);
        }

        // Build the container list up front so we know if any is drive-only
        // (OneDrive), which needs a token for the drive-reachability ping even
        // when no Excel/files are configured.
        Set<String> usedSlugs = new java.util.HashSet<>();
        List<SiteCtx> containers = new ArrayList<>();
        boolean anyDriveOnly = false;
        for (String entry : entries) {
            SiteCtx s = newContainer(entry);
            String unique = s.siteSlug;
            int n = 2;
            while (!usedSlugs.add(unique)) unique = s.siteSlug + "_" + (n++);
            s.siteSlug = unique;
            if (s.siteId == null) anyDriveOnly = true;
            containers.add(s);
        }

        // Mint one token shared across all per-container readers + drive pings
        // (same creds, same MSAL cache).
        boolean needToken = !excelPathsBySlug.isEmpty() || !filePathsBySlug.isEmpty() || anyDriveOnly;
        String token = needToken ? graphFactory.bearerToken(auth) : null;
        ctx.token = token;

        for (SiteCtx s : containers) {
            SharePointWorkbookReader wb = excelPathsBySlug.isEmpty() ? null
                    : new SharePointWorkbookReader(graphFactory.rawHttpClient(), token, s.driveResourcePath,
                            excelPathsBySlug, graphFactory.gridCache());
            SharePointFileReader fr = filePathsBySlug.isEmpty() ? null
                    : new SharePointFileReader(graphFactory.rawHttpClient(), token, s.driveResourcePath,
                            filePathsBySlug, graphFactory.gridCache());
            s.readers = new SharePointReaders(wb, fr);
        }
        ctx.sites = containers;
        ctx.multiSite = ctx.sites.size() > 1;
        return ctx;
    }

    /**
     * Resolve the container entries for this connection. Base = SharePoint
     * sites from SITE_URLS / SITE_URL. Overridden by the OneDrive provider to
     * read user UPNs.
     */
    protected List<String> resolveContainerEntries(RequestInfo info) {
        String siteUrls = param(info, PARAM_SITE_URLS, false);
        String siteUrl = param(info, PARAM_SITE_URL, false);
        if (siteUrls != null && !siteUrls.trim().isEmpty()) {
            if (siteUrl != null && !siteUrl.isEmpty()) {
                log.warn("Both SITE_URLS and SITE_URL are set; using SITE_URLS and ignoring SITE_URL");
            }
            return SharePointIntrospector.parsePaths(siteUrls);
        }
        if (siteUrl != null && !siteUrl.isEmpty()) {
            return Collections.singletonList(siteUrl);
        }
        throw new IllegalArgumentException("One of SITE_URL or SITE_URLS is required");
    }

    /**
     * Build a container (SiteCtx) for one entry. Base = a SharePoint site:
     * siteId set (enables List discovery), driveResourcePath = sites/{id}/drive.
     * The OneDrive provider overrides this to a user drive (siteId == null,
     * driveResourcePath = users/{upn}/drive). Readers are attached later.
     */
    protected SiteCtx newContainer(String entry) {
        SiteCtx s = new SiteCtx();
        s.siteUrl = entry;
        s.siteId = graphFactory.siteIdFromUrl(entry);
        s.siteSlug = graphFactory.siteSlugFromUrl(entry);
        s.driveResourcePath = "sites/" + s.siteId + "/drive";
        return s;
    }

    /**
     * Validate one container is reachable. Base validates the SharePoint site;
     * the OneDrive provider validates the user's drive. Override as needed.
     */
    protected boolean validateContainer(ConnContext ctx, SiteCtx s) {
        if (s.siteId != null) {
            return introspector.validateConnection(ctx.client, s.siteId);
        }
        // Drive-only container: ping the drive directly.
        return graphFactory.ping("https://graph.microsoft.com/v1.0/" + s.driveResourcePath, ctx.token);
    }

    /**
     * Discover collections across all configured sites, prefixing each name
     * with "<site-slug>__" only in multi-site connections (single-site keeps
     * the v1 names unchanged).
     */
    private List<CollectionInfo> allCollections(ConnContext ctx) {
        List<CollectionInfo> out = new ArrayList<>();
        for (SiteCtx s : ctx.sites) {
            List<CollectionInfo> siteCols = introspector.getCollections(
                    ctx.client, s.siteId, ctx.includeLists, s.readers);
            for (CollectionInfo ci : siteCols) {
                if (ctx.multiSite) {
                    CollectionInfo p = new CollectionInfo();
                    p.setCollection(s.siteSlug + "__" + ci.getCollection());
                    p.setSchema("default");
                    out.add(p);
                } else {
                    out.add(ci);
                }
            }
        }
        return out;
    }

    /**
     * Route a (possibly site-prefixed) collection name back to its site and
     * the underlying entity collection name. Single-site: the name passes
     * through unchanged. Multi-site: peel the first "<slug>__" segment (site
     * slugs never contain "__"), match it to a configured site, and hand the
     * remainder to the per-site code path.
     */
    private Resolved resolve(ConnContext ctx, String collectionName) {
        Resolved r = new Resolved();
        if (!ctx.multiSite) {
            SiteCtx s = ctx.sites.get(0);
            r.siteId = s.siteId;
            r.readers = s.readers;
            r.entity = collectionName;
            return r;
        }
        int i = collectionName.indexOf("__");
        if (i < 0) {
            throw new IllegalArgumentException("Collection '" + collectionName
                    + "' is not site-qualified in a multi-site connection");
        }
        String slug = collectionName.substring(0, i);
        String entity = collectionName.substring(i + 2);
        for (SiteCtx s : ctx.sites) {
            if (s.siteSlug.equals(slug)) {
                r.siteId = s.siteId;
                r.readers = s.readers;
                r.entity = entity;
                return r;
            }
        }
        throw new IllegalArgumentException("No configured site matches slug '" + slug + "'");
    }

    protected String param(RequestInfo info, String name, boolean required) {
        if (info.getDataSourceInfo() == null || info.getDataSourceInfo().getParams() == null) {
            if (required) throw new IllegalArgumentException("Missing connection parameters");
            return null;
        }
        String value = info.getDataSourceInfo().getParams().get(name);
        if (required && (value == null || value.isEmpty())) {
            throw new IllegalArgumentException("Missing required parameter: " + name);
        }
        return value;
    }

    protected static class ConnContext {
        GraphServiceClient client;
        Set<String> includeLists;
        List<SiteCtx> sites;
        boolean multiSite;
        String token; // shared Graph token (null when not needed)
    }

    /**
     * Per-container resolved state. One per configured site (SharePoint) or
     * user drive (OneDrive). siteId is null for drive-only containers, which
     * disables List discovery; driveResourcePath addresses the drive for the
     * Excel/file readers and the reachability ping.
     */
    protected static class SiteCtx {
        String siteUrl;          // original entry (site URL or user UPN) — for messages
        String siteSlug;         // collection-name prefix in multi-container connections
        String siteId;           // SharePoint site id, or null for OneDrive
        String driveResourcePath; // "sites/{id}/drive" or "users/{upn}/drive"
        SharePointReaders readers; // workbook + file readers; either may be null
    }

    /** Result of routing a collection name to its site + entity. */
    private static class Resolved {
        String siteId;
        SharePointReaders readers;
        String entity;
    }
}
