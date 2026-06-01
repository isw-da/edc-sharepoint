/**
 * Copyright (C) Zoomdata, Inc. 2012-2017. All rights reserved.
 */
package com.zoomdata.connector.example.framework.common;

import com.google.common.collect.ImmutableMap;

import java.util.Map;
import java.util.Objects;

public final class SQLConnectionPoolKey {
    private final String jdbcUrl;
    private final String username;
    private final String password;
    private final Map<String, String> customParams;

    public SQLConnectionPoolKey(String jdbcUrl, String username, String password) {
        this(jdbcUrl, username, password, null);
    }

    public SQLConnectionPoolKey(String jdbcUrl, String username, String password, Map<String, String> customParams) {
        this.jdbcUrl = jdbcUrl;
        this.username = username;
        this.password = password;
        this.customParams = customParams == null ? ImmutableMap.of() : ImmutableMap.copyOf(customParams);
    }

    public String getJdbcUrl() { return jdbcUrl; }
    public String getUsername() { return username; }
    public String getPassword() { return password; }
    public Map<String, String> getCustomParams() { return customParams; }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        SQLConnectionPoolKey that = (SQLConnectionPoolKey) o;
        return Objects.equals(jdbcUrl, that.jdbcUrl)
                && Objects.equals(username, that.username)
                && Objects.equals(password, that.password)
                && Objects.equals(customParams, that.customParams);
    }

    @Override
    public int hashCode() {
        return Objects.hash(jdbcUrl, username, password, customParams);
    }

    @Override
    public String toString() {
        return "SQLConnectionPoolKey{jdbcUrl='" + jdbcUrl + "', username='" + username + "'}";
    }
}
