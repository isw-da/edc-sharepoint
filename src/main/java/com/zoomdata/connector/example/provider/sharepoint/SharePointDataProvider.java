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

@Connector(CONNECTION_TYPE)
public class SharePointDataProvider extends AbstractDataProvider {

    private static final Logger log = LoggerFactory.getLogger(SharePointDataProvider.class);

    protected static final String CONNECTION_TYPE = "SHAREPOINT";

    private static final String PARAM_TENANT_ID = "TENANT_ID";
    private static final String PARAM_CLIENT_ID = "CLIENT_ID";
    private static final String PARAM_CLIENT_SECRET = "CLIENT_SECRET";
    private static final String PARAM_SITE_URL = "SITE_URL";
    private static final String PARAM_INCLUDE_LISTS = "INCLUDE_LISTS";
    private static final String PARAM_INCLUDE_EXCEL = "INCLUDE_EXCEL";
    private static final String PARAM_AUTHORITY = "AUTHORITY";

    private final SharePointTypesMapping typesMapping = new SharePointTypesMapping();
    private final SharePointFeatures features = new SharePointFeatures();
    private final SharePointGraphClient graphFactory = new SharePointGraphClient();
    private final SharePointIntrospector introspector = new SharePointIntrospector(typesMapping);

    @Override
    public ValidateSourceResponse pingSource(ValidateSourceRequest request) {
        try {
            ConnContext ctx = context(request.getRequestInfo());
            boolean ok = introspector.validateConnection(ctx.client, ctx.siteId);
            return new ValidateSourceResponse(ok ? ok() : serverError("Failed to connect to SharePoint site"));
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
            List<CollectionInfo> collections = introspector.getCollections(
                    ctx.client, ctx.siteId, ctx.includeLists, ctx.includeExcel);
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
            List<CollectionInfo> collections = introspector.getCollections(
                    ctx.client, ctx.siteId, ctx.includeLists, ctx.includeExcel);
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
            List<FieldMetadata> fields = introspector.describeCollection(ctx.client, ctx.siteId, collectionName);
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
            List<CollectionInfo> all = introspector.getCollections(ctx.client, ctx.siteId,
                    ctx.includeLists, ctx.includeExcel);

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
                    List<FieldMetadata> fields = introspector.describeCollection(
                            ctx.client, ctx.siteId, ci.getCollection());
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
            List<FieldMetadata> fieldMeta = introspector.describeCollection(ctx.client, ctx.siteId, collectionName);
            List<String> fieldNames = fieldMeta.stream().map(FieldMetadata::getName).collect(Collectors.toList());

            SharePointComputeTask task = new SharePointComputeTask(
                    ctx.client, ctx.siteId, collectionName, fieldNames,
                    typesMapping, introspector, 10, null);
            if (ctx.excelPathsBySlug != null && !ctx.excelPathsBySlug.isEmpty()) {
                task.setExcelPathsBySlug(ctx.excelPathsBySlug);
            }
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

            return new SharePointComputeTaskFactory(ctx.client, ctx.siteId, collectionName,
                    requestedFields, typesMapping, introspector, fetchSize, filters, ctx.excelPathsBySlug);
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
                .addParameters(
                        stringParameter(PARAM_TENANT_ID)
                                .isRequired(true)
                                .description("Azure AD tenant ID (GUID or domain, e.g. contoso.onmicrosoft.com)"))
                .addParameters(
                        stringParameter(PARAM_CLIENT_ID)
                                .isRequired(true)
                                .description("App registration client (application) ID"))
                .addParameters(
                        passwordParameter(PARAM_CLIENT_SECRET)
                                .isRequired(true)
                                .description("App registration client secret"))
                .addParameters(
                        stringParameter(PARAM_SITE_URL)
                                .isRequired(true)
                                .description("SharePoint site URL, e.g. https://contoso.sharepoint.com/sites/sales"))
                .addParameters(
                        stringParameter(PARAM_INCLUDE_LISTS)
                                .description("Comma-separated allowlist of List displayNames. Empty = all visible Lists."))
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

    // ----- Connection context -------------------------------------------

    /**
     * Bundle of the GraphServiceClient + resolved site identifier + parsed
     * include lists for one EDC call. Built once per request from the
     * connection params.
     */
    private ConnContext context(RequestInfo info) {
        String tenantId = param(info, PARAM_TENANT_ID, true);
        String clientId = param(info, PARAM_CLIENT_ID, true);
        String clientSecret = param(info, PARAM_CLIENT_SECRET, true);
        String siteUrl = param(info, PARAM_SITE_URL, true);
        String authority = param(info, PARAM_AUTHORITY, false);
        String includeLists = param(info, PARAM_INCLUDE_LISTS, false);
        String includeExcel = param(info, PARAM_INCLUDE_EXCEL, false);

        GraphServiceClient client = graphFactory.build(tenantId, clientId, clientSecret, authority);
        String siteId = graphFactory.siteIdFromUrl(siteUrl);

        ConnContext ctx = new ConnContext();
        ctx.client = client;
        ctx.siteId = siteId;
        ctx.includeLists = SharePointIntrospector.parseAllowlist(includeLists);
        ctx.includeExcel = SharePointIntrospector.parsePaths(includeExcel);
        ctx.excelPathsBySlug = new HashMap<>();
        for (String p : ctx.includeExcel) {
            ctx.excelPathsBySlug.put(SharePointIntrospector.excelSlugFromPath(p), p);
        }
        return ctx;
    }

    private String param(RequestInfo info, String name, boolean required) {
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

    private static class ConnContext {
        GraphServiceClient client;
        String siteId;
        Set<String> includeLists;
        List<String> includeExcel;
        Map<String, String> excelPathsBySlug;
    }
}
