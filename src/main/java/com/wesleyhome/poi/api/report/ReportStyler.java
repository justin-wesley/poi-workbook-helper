package com.wesleyhome.poi.api.report;

import com.wesleyhome.poi.api.CellStyler;

public interface ReportStyler {

    void createStyles(CellStyler reset);

    String getDescriptionStyleName();

    String getReportTitleStyleName();
}
