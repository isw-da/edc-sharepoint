/**
 * Copyright (C) insightsoftware 2026. All rights reserved.
 */
package com.zoomdata.connector.example.provider.sharepoint;

import com.microsoft.graph.serviceclient.GraphServiceClient;
import com.zoomdata.connector.example.framework.async.IComputeTask;
import com.zoomdata.connector.example.framework.async.IComputeTaskFactory;
import com.zoomdata.gen.edc.filter.Filter;

import java.util.List;

public class SharePointComputeTaskFactory implements IComputeTaskFactory {

    private final GraphServiceClient client;
    private final String siteId;
    private final String collectionName;
    private final List<String> requestedFields;
    private final SharePointTypesMapping typesMapping;
    private final SharePointIntrospector introspector;
    private final int fetchSize;
    private final List<Filter> filters;
    private final SharePointWorkbookReader workbook; // null for List-only connections
    private final String rawQuery;

    public SharePointComputeTaskFactory(GraphServiceClient client, String siteId,
                                         String collectionName, List<String> requestedFields,
                                         SharePointTypesMapping typesMapping,
                                         SharePointIntrospector introspector,
                                         int fetchSize, List<Filter> filters,
                                         SharePointWorkbookReader workbook) {
        this.client = client;
        this.siteId = siteId;
        this.collectionName = collectionName;
        this.requestedFields = requestedFields;
        this.typesMapping = typesMapping;
        this.introspector = introspector;
        this.fetchSize = fetchSize;
        this.filters = filters;
        this.workbook = workbook;

        StringBuilder sb = new StringBuilder("SharePoint ").append(collectionName);
        if (filters != null && !filters.isEmpty()) sb.append(" filter(...)");
        if (requestedFields != null && !requestedFields.isEmpty()) {
            sb.append(" select(").append(String.join(",", requestedFields)).append(")");
        }
        this.rawQuery = sb.toString();
    }

    @Override
    public IComputeTask create() {
        return new SharePointComputeTask(
                client, siteId, collectionName, requestedFields,
                typesMapping, introspector, fetchSize, filters, workbook);
    }

    @Override public String getRawQuery() { return rawQuery; }
    @Override public int getFetchSize() { return fetchSize; }
}
