package com.wesleyhome.poi.api.report;

import com.wesleyhome.poi.api.CellStyler;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.function.Function;

public class DefaultReportStyler implements ReportStyler {

    protected static String REPORT_TITLE = "REPORT_TITLE";
    protected static String REPORT_DESCRIPTION_DETAILS = "REPORT_DESCRIPTION_DETAILS";

    @Override
    public String getDescriptionStyleName() {
        return REPORT_DESCRIPTION_DETAILS;
    }

    @Override
    public String getReportTitleStyleName() {
        return REPORT_TITLE;
    }

    @Override
    public void createStyles(CellStyler cellStyler) {
        cellStyler.applyIf(true, createReportTitleStyle()).reset()
            .applyIf(true, createReportDetailsStyle()).reset()
            .reset();
    }

    private Function<CellStyler, CellStyler> createReportDetailsStyle() {
        return cs -> cs.withItalicFont().withFontSize(10).as(REPORT_DESCRIPTION_DETAILS);
    }

    protected Function<CellStyler, CellStyler> createReportTitleStyle() {
        return cs -> cs.withBoldFont()
            .withFontSize(28)
            .as(REPORT_TITLE);
    }
}
