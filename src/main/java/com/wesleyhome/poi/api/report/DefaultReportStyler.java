package com.wesleyhome.poi.api.report;

import com.wesleyhome.poi.api.CellStyler;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.Arrays;
import java.util.List;
import java.util.function.Function;

public class DefaultReportStyler implements ReportStyler {

    private static String REPORT_TITLE = "REPORT_TITLE";
    private static String REPORT_DESCRIPTION_DETAILS = "REPORT_DESCRIPTION_DETAILS";
    private static String COLUMN_HEADER = "COLUMN_HEADER";
    private static String EVEN_ROW = "EVEN_ROW";
    private static String ODD_ROW = "ODD_ROW";

    @Override
    public List<String> getRowStyles(boolean isEven) {
        return Arrays.asList(isEven ? EVEN_ROW : ODD_ROW);
    }

    @Override
    public String getHeaderStyleName() {
        return COLUMN_HEADER;
    }

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
            .applyIf(true, createColumnHeaderStyle()).reset()
            .applyIf(true, createEvenRowStyle()).reset()
            .applyIf(true, createOddRowStyle());
    }

    private Function<CellStyler, CellStyler> createReportDetailsStyle() {
        return cs -> cs.withItalicFont().withFontSize(10).as(REPORT_DESCRIPTION_DETAILS);
    }

    protected Function<CellStyler, CellStyler> createOddRowStyle() {
        return cs -> cs.withBackgroundColor(oddRowBackgroundColor())
            .withFontColor(oddRowFontColor())
            .applyIf(isOddRowBold(), CellStyler::withBoldFont)
            .as(ODD_ROW);
    }

    protected Function<CellStyler, CellStyler> createEvenRowStyle() {
        return cs -> cs.withBackgroundColor(evenRowBackgroundColor())
            .withFontColor(evenRowFontColor())
            .applyIf(isEvenRowBold(), CellStyler::withBoldFont)
            .as(EVEN_ROW);
    }

    protected Function<CellStyler, CellStyler> createColumnHeaderStyle() {
        return cs -> cs.withBackgroundColor(headerBackgroundColor())
            .withFontColor(headerFontColor())
            .withHorizontalAlignment(HorizontalAlignment.CENTER)
            .applyIf(isHeaderBold(), CellStyler::withBoldFont)
            .as(COLUMN_HEADER);
    }

    protected Function<CellStyler, CellStyler> createReportTitleStyle() {
        return cs -> cs.withBoldFont()
            .withFontSize(28)
            .as(REPORT_TITLE);
    }

    public IndexedColors headerBackgroundColor() {
        return IndexedColors.DARK_BLUE;
    }

    public IndexedColors headerFontColor() {
        return IndexedColors.WHITE;
    }

    public boolean isHeaderBold() {
        return true;
    }

    public IndexedColors oddRowBackgroundColor() {
        return IndexedColors.WHITE;
    }

    public IndexedColors oddRowFontColor() {
        return IndexedColors.BLACK;
    }

    public boolean isEvenRowBold() {
        return false;
    }

    public IndexedColors evenRowBackgroundColor() {
        return IndexedColors.LIGHT_CORNFLOWER_BLUE;
    }

    public IndexedColors evenRowFontColor() {
        return IndexedColors.BLACK;
    }

    public boolean isOddRowBold() {
        return false;
    }

}
