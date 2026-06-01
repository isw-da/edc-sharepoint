/**
 * Copyright (C) insightsoftware 2026. All rights reserved.
 */
package com.zoomdata.connector.example.provider.sharepoint;

import com.microsoft.graph.serviceclient.GraphServiceClient;
import com.zoomdata.connector.example.framework.async.IComputeTask;
import com.zoomdata.connector.example.framework.async.IComputeTaskFactory;
import com.zoomdata.gen.edc.filter.Filter;

import java.util.List;
import java.util.Map;

public class SharePointComputeTaskFactory implements IComputeTaskFactory {

    private final GraphServiceClient client;
    private final String siteId;
    private final String collectionName;
    private final List<String> requestedFields;
    private final SharePointTypesMapping typesMapping;
    private final SharePointIntrospector introspector;
    private final int fetchSize;
    private final List<Filter> filters;
    private final Map<String, String> excelPathsBySlug;
    private final String rawQuery;

    public SharePointComputeTaskFactory(GraphServiceClient client, String siteId,
                                         String collectionName, List<String> requestedFields,
                                         SharePointTypesMapping typesMapping,
                                         SharePointIntrospector introspector,
                                         int fetchSize, List<Filter> filters,
                                         Map<String, String> excelPathsBySlug) {
        this.client = client;
        this.siteId = siteId;
        this.collectionName = collectionName;
        this.requestedFields = requestedFields;
        this.typesMapping = typesMapping;
        this.introspector = introspector;
        this.fetchSize = fetchSize;
        this.filters = filters;
        this.excelPathsBySlug = excelPathsBySlug;

        StringBuilder sb = new StringBuilder("SharePoint ").append(collectionName);
        if (filters != null && !filters.isEmpty()) sb.append(" filter(...)");
        if (requestedFields != null && !requestedFields.isEmpty()) {
            sb.append(" select(").append(String.join(",", requestedFields)).append(")");
        }
        this.rawQuery = sb.toString();
    }

    @Override
    public IComputeTask create() {
        SharePointComputeTask task = new SharePointComputeTask(
                client, siteId, collectionName, requestedFields,
                typesMapping, introspector, fetchSize, filters);
        if (excelPathsBySlug != null && !excelPathsBySlug.isEmpty()) {
            task.setExcelPathsBySlug(excelPathsBySlug);
        }
        return task;
    }

    @Override public String getRawQuery() { return rawQuery; }
    @Override public int getFetchSize() { return fetchSize; }
}
