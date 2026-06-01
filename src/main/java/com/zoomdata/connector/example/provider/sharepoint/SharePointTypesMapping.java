/**
 * Copyright (C) insightsoftware 2026. All rights reserved.
 *
 * Maps SharePoint ColumnDefinition column types to Composer Thrift FieldType.
 *
 * SharePoint columns expose their type via the ColumnDefinition discriminator
 * fields (text, number, dateTime, boolean, choice, lookup, personOrGroup,
 * hyperlinkOrPicture, calculated, currency, etc). The introspector resolves
 * which discriminator is set on each column and passes a lowercase key here.
 */
package com.zoomdata.connector.example.provider.sharepoint;

import com.google.common.collect.ImmutableMap;
import com.zoomdata.connector.example.framework.api.ITypesMapping;
import com.zoomdata.connector.example.framework.common.Meta;
import com.zoomdata.connector.example.framework.common.ThriftTypeFunction;
import com.zoomdata.gen.edc.types.FieldType;

import java.util.Map;

public class SharePointTypesMapping implements ITypesMapping {

    private final Meta defaultMeta = metaString();

    private final Map<String, Meta> typesMapping =
            ImmutableMap.<String, Meta>builder()
                    // SharePoint scalar column kinds
                    .put("text", metaString())
                    .put("number", metaDouble())
                    .put("currency", metaDouble())
                    .put("boolean", metaString())
                    .put("datetime", metaTimestamp())
                    .put("date", metaTimestamp())
                    .put("choice", metaString())
                    .put("lookup", metaString())
                    .put("personorgroup", metaString())
                    .put("hyperlinkorpicture", metaString())
                    .put("calculated", metaString())
                    .put("thumbnail", metaString())
                    .put("term", metaString())
                    .put("contenttypeinfo", metaString())
                    // Excel Workbook column kinds (inferred from sample cell values)
                    .put("string", metaString())
                    .put("integer", metaInt())
                    .put("double", metaDouble())
                    .put("long", metaInt())
                    .build();

    private static Meta metaInt() {
        return Meta.from(FieldType.INTEGER, ThriftTypeFunction.GET_INTEGER);
    }

    private static Meta metaDouble() {
        return Meta.from(FieldType.DOUBLE, ThriftTypeFunction.GET_DOUBLE);
    }

    private static Meta metaString() {
        return Meta.from(FieldType.STRING, ThriftTypeFunction.GET_STRING);
    }

    private static Meta metaTimestamp() {
        return Meta.from(FieldType.DATE, ThriftTypeFunction.GET_TIMESTAMP);
    }

    @Override
    public Meta metaForType(String type) {
        if (type == null) return defaultMeta;
        return typesMapping.getOrDefault(type.toLowerCase(), defaultMeta);
    }
}
