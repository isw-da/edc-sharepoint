/**
 * Copyright (C) Zoomdata, Inc. 2012-2017. All rights reserved.
 */
package com.zoomdata.connector.example.framework.common;

import com.zoomdata.gen.edc.types.FieldType;

public class Meta {
    private FieldType thriftType;
    private ThriftTypeFunction rsFunction;

    public Meta(FieldType thriftType, ThriftTypeFunction rsFunction) {
        this.thriftType = thriftType;
        this.rsFunction = rsFunction;
    }

    public FieldType getThriftType() {
        return thriftType;
    }

    public void setThriftType(FieldType thriftType) {
        this.thriftType = thriftType;
    }

    public ThriftTypeFunction getRsFunction() {
        return rsFunction;
    }

    public void setRsFunction(ThriftTypeFunction rsFunction) {
        this.rsFunction = rsFunction;
    }

    public static Meta from(FieldType thriftType, ThriftTypeFunction rsFunction) {
        return new Meta(thriftType, rsFunction);
    }
}
