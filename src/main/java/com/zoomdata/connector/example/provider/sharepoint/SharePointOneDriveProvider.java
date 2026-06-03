/**
 * Copyright (C) insightsoftware 2026. All rights reserved.
 *
 * OneDrive for Business EDC provider (subStorageType SHAREPOINT_ONEDRIVE).
 *
 * A thin subclass of SharePointDataProvider. OneDrive is "files on a user's
 * drive": it reuses the Excel (Workbook) and CSV/JSON file readers unchanged,
 * just rooted at /users/{upn}/drive instead of /sites/{id}/drive. It has no
 * SharePoint Lists, so List discovery is skipped (siteId == null).
 *
 * Only three things differ from the SharePoint provider, all expressed as
 * overrides: the connection params (USER_UPNS instead of SITE_URLS, no
 * INCLUDE_LISTS), how a container is built (a user drive), and the connector
 * type/icon in the description. Everything else — routing, multi-container
 * prefixing, describe, sample, fetch — is inherited.
 */
package com.zoomdata.connector.example.provider.sharepoint;

import com.zoomdata.connector.example.framework.annotation.Connector;
import com.zoomdata.connector.example.framework.api.IDescriptionProvider;
import com.zoomdata.connector.example.framework.provider.serverdescription.GenericDescriptionProvider;
import com.zoomdata.gen.edc.request.RequestInfo;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.autoconfigure.condition.ConditionalOnProperty;

import java.util.Collections;
import java.util.List;

import static com.zoomdata.connector.example.framework.provider.serverdescription.connectionparameters.impl.PasswordConnectionParameter.PasswordConnectionParameterBuilder.passwordParameter;
import static com.zoomdata.connector.example.framework.provider.serverdescription.connectionparameters.impl.StringConnectionParameter.StringConnectionParameterBuilder.stringParameter;
import static com.zoomdata.connector.example.provider.sharepoint.SharePointOneDriveProvider.ONEDRIVE_TYPE;

// Active only when EDC_CONNECTOR=SHAREPOINT_ONEDRIVE — runs as its own pod so
// each server process hosts exactly one connector (see SharePointDataProvider).
@Connector(ONEDRIVE_TYPE)
@ConditionalOnProperty(name = "edc.connector", havingValue = "SHAREPOINT_ONEDRIVE")
public class SharePointOneDriveProvider extends SharePointDataProvider {

    protected static final String ONEDRIVE_TYPE = "SHAREPOINT_ONEDRIVE";

    private static final Logger log = LoggerFactory.getLogger(SharePointOneDriveProvider.class);

    private static final String PARAM_USER_UPN = "USER_UPN";    // single-user alias
    private static final String PARAM_USER_UPNS = "USER_UPNS";  // multi-user, comma-separated

    /** Container entries are user UPNs rather than site URLs. */
    @Override
    protected List<String> resolveContainerEntries(RequestInfo info) {
        String upns = param(info, PARAM_USER_UPNS, false);
        String upn = param(info, PARAM_USER_UPN, false);
        if (upns != null && !upns.trim().isEmpty()) {
            if (upn != null && !upn.isEmpty()) {
                log.warn("Both USER_UPNS and USER_UPN are set; using USER_UPNS and ignoring USER_UPN");
            }
            return SharePointIntrospector.parsePaths(upns);
        }
        if (upn != null && !upn.isEmpty()) {
            return Collections.singletonList(upn);
        }
        throw new IllegalArgumentException("One of USER_UPN or USER_UPNS is required");
    }

    /**
     * A OneDrive container: the user's drive. No siteId (so List discovery is
     * skipped); driveResourcePath = users/{upn}/drive; slug = the UPN local
     * part (the bit before '@'), slugified.
     */
    @Override
    protected SiteCtx newContainer(String entry) {
        SiteCtx s = new SiteCtx();
        s.siteUrl = entry;            // the UPN, used in messages
        s.siteId = null;              // no SharePoint site -> no Lists
        s.driveResourcePath = "users/" + entry + "/drive";
        s.siteSlug = upnSlug(entry);
        return s;
    }

    private static String upnSlug(String upn) {
        String local = upn.contains("@") ? upn.substring(0, upn.indexOf('@')) : upn;
        String slug = local.replaceAll("[^A-Za-z0-9]+", "_").toLowerCase();
        return slug.isEmpty() ? "user" : slug;
    }

    @Override
    protected IDescriptionProvider createDescriptionProvider() {
        return new GenericDescriptionProvider(ONEDRIVE_TYPE)
                .svgIcon("/sharepoint-icon.svg")
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
                        stringParameter(PARAM_USER_UPN)
                                .description("Single user principal name, e.g. jane@contoso.com. Use this OR USER_UPNS."))
                .addParameters(
                        stringParameter(PARAM_USER_UPNS)
                                .description("Comma-separated user UPNs for a multi-user connection. When >1, "
                                        + "collection names are prefixed with the user slug. Takes precedence over USER_UPN."))
                .addParameters(
                        stringParameter(PARAM_INCLUDE_FILES)
                                .description("Comma-separated paths to CSV/TSV/JSON files in the user's OneDrive "
                                        + "(e.g. /Reports/agents.csv). Each becomes a collection."))
                .addParameters(
                        stringParameter(PARAM_INCLUDE_EXCEL)
                                .description("Comma-separated paths to .xlsx files in the user's OneDrive. Empty = no Excel."))
                .addParameters(
                        stringParameter(PARAM_AUTHORITY)
                                .description("Azure AD authority host override (sovereign clouds). Default https://login.microsoftonline.com"))
                .minVersion("1.0")
                .maxVersion("1.0");
    }
}
