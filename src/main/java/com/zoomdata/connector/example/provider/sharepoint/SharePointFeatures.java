/**
 * Copyright (C) insightsoftware 2026. All rights reserved.
 */
package com.zoomdata.connector.example.provider.sharepoint;

import com.google.common.collect.ImmutableMap;
import com.zoomdata.connector.example.framework.api.IFeatures;

import java.util.Map;

public class SharePointFeatures implements IFeatures {

    private final Map<String, String> features =
            ImmutableMap.<String, String>builder()
                    .put("REQUEST.SEND_METADATA", "true")
                    .put("REQUEST.TYPE", "STRUCTURED")
                    .put("FEATURE.LIVE_SOURCE", "true")
                    .put("FEATURE.REFRESHABLE", "true")
                    .put("FEATURE.RAW_DATA_ONLY", "false")
                    .put("FEATURE.SUPPORTS_SCHEMA", "true")
                    .put("FEATURE.OFFSET", "false")
                    .put("FEATURE.FAST_DISTINCT_VALUES", "false")
                    .put("FEATURE.GROUP_BY_TIME", "false")
                    .put("FEATURE.GROUP_BY_TIME.GROUP_BY_UNIX_TIME", "false")
                    .put("FEATURE.MULTI_GROUP_SUPPORT", "false")
                    .put("FEATURE.SUPPORT_OPTIMIZED_READ", "false")
                    .put("FEATURE.TEXT_SEARCH", "false")
                    .put("FEATURE.WILDCARD_FILTERS", "true")
                    .put("FEATURE.WILDCARD_FILTERS.CASE_SENSITIVE", "true")
                    .put("FEATURE.WILDCARD_FILTERS.CASE_INSENSITIVE", "true")
                    .put("FEATURE.DISTINCT_COUNT", "false")
                    .put("FEATURE.DISTINCT_COUNT.DISTINCT_COUNT_ONLY_ONE", "false")
                    .put("FEATURE.PARTITION", "false")
                    .put("FEATURE.SUPPORTS_MULTI_VALUED", "false")
                    .put("FEATURE.SUPPORTS_NESTED", "false")
                    .put("FEATURE.SUPPORTED_BY_SPARKIT", "false")
                    .put("FEATURE.CUSTOM_QUERY", "false")
                    .put("FEATURE.PAGING_AND_SORTING", "false")
                    .put("FEATURE.PAGING_AND_SORTING.AGGREGATED", "false")
                    .put("FEATURE.LV_METRIC", "false")
                    .put("FEATURE.PERCENTILES", "false")
                    .put("FEATURE.HISTOGRAM", "false")
                    .build();

    @Override
    public Map<String, String> getAllFeatures() {
        return features;
    }
}
