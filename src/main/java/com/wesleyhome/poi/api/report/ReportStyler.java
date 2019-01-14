package com.wesleyhome.poi.api.report;

import com.wesleyhome.poi.api.CellStyler;

import java.util.List;

public interface ReportStyler {

    void createStyles(CellStyler reset);

    List<String> getRowStyles(boolean isEven);

    String getHeaderStyleName();

    String getDescriptionStyleName();
}
