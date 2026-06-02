/**
 * Copyright (C) insightsoftware 2026. All rights reserved.
 *
 * Per-site bundle of the data-source readers a connection needs beyond
 * SharePoint Lists: the Workbook reader (Excel) and the file reader
 * (CSV/JSON). Either may be null when that source kind isn't configured
 * (no INCLUDE_EXCEL / no INCLUDE_FILES). Grouping them keeps the
 * introspector / ComputeTask signatures stable as more reader kinds are
 * added (e.g. OneDrive reuses these).
 */
package com.zoomdata.connector.example.provider.sharepoint;

public final class SharePointReaders {

    final SharePointWorkbookReader workbook; // null when no INCLUDE_EXCEL
    final SharePointFileReader files;        // null when no INCLUDE_FILES

    public SharePointReaders(SharePointWorkbookReader workbook, SharePointFileReader files) {
        this.workbook = workbook;
        this.files = files;
    }

    static final SharePointReaders NONE = new SharePointReaders(null, null);
}
