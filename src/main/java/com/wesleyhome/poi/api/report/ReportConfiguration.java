package com.wesleyhome.poi.api.report;

import com.wesleyhome.poi.api.CellGenerator;
import com.wesleyhome.poi.api.CellStyler;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;
import java.util.stream.Collectors;

import static com.wesleyhome.poi.api.internal.DefaultWorkbookGenerator.BuiltinStyles.*;

public interface ReportConfiguration<T> {

    String REPORT_DESCRIPTION = "REPORT_DESCRIPTION";
    String COLUMN_HEADER = "COLUMN_HEADER";
    String EVEN_ROW = "EVEN_ROW";
    String ODD_ROW = "ODD_ROW";

    SortedMap<String, ColumnConfiguration<T>> columns();

    String getReportSheetName();

    String getReportDescription();

    IndexedColors headerBackgroundColor();

    IndexedColors headerFontColor();

    boolean isHeaderBold();

    IndexedColors evenRowBackgroundColor();

    IndexedColors evenRowFontColor();

    boolean isEvenRowBold();

    IndexedColors oddRowBackgroundColor();

    IndexedColors oddRowFontColor();

    boolean isOddRowBold();

    default ColumnConfiguration<T> getColumnConfiguration(String columnIdentifier) {
        return columns().get(columnIdentifier);
    }

    default Map<String, String> columnHeaders() {
        return columns().values()
            .stream()
            .collect(Collectors.toMap(ColumnConfiguration::getColumnName, ColumnConfiguration::getColumnHeader));
    }

    default CellGenerator applyStyleAndValueToCell(CellGenerator cellGenerator, ColumnConfiguration<T> columnConfiguration, Object transformedValue) {
        int rowNum = cellGenerator.rowNum();
        boolean isEven = rowNum % 2 == 0;
        List<String> styles = new ArrayList<>();
        styles.add(isEven ? EVEN_ROW : ODD_ROW);
        ColumnType columnType = columnConfiguration.getColumnType();
        if(ColumnType.DERIVED.equals(columnType) && transformedValue != null){
            if (transformedValue instanceof Date || transformedValue instanceof LocalDateTime || transformedValue instanceof LocalDate) {
                columnType = ColumnType.DATE;
            } else if(transformedValue instanceof Integer){
                columnType = ColumnType.INTEGER;
            } else if(transformedValue instanceof Number) {
                columnType = ColumnType.DECIMAL;
            } else if(transformedValue instanceof Boolean) {
                columnType = ColumnType.BOOLEAN;
            } else{
                columnType = ColumnType.TEXT;
            }
        }
        switch (columnType) {
            case DECIMAL:
                styles.add(NUMERIC);
                break;
            case DATE:
                styles.add(DATE);
                break;
            case INTEGER:
                styles.add(INTEGER);
                break;
            case TEXT:
            case URL:
            case FORMULA:
            case CURRENCY:
            case FRACTION:
            case ACCOUNTING:
            case SCI_NOTATION:
            case BOOLEAN:

            case DERIVED:
                break;
        }

        CellGenerator cg = cellGenerator
            .usingStyles(styles);
        return cg
            .havingValue(transformedValue);
    }

    default void createStyles(CellStyler cellStyler) {
        cellStyler
            .isBold()
            .withFontSize(36)
            .as(REPORT_DESCRIPTION)
            .reset()
            .withBackgroundColor(headerBackgroundColor())
            .withFontColor(headerFontColor())
            .withHorizontalAlignment(HorizontalAlignment.CENTER)
            .applyIf(isHeaderBold(), CellStyler::isBold)
            .as(COLUMN_HEADER)
            .reset()
            .withBackgroundColor(evenRowBackgroundColor())
            .withFontColor(evenRowFontColor())
            .applyIf(isEvenRowBold(), CellStyler::isBold)
            .as(EVEN_ROW)
            .reset()
            .withBackgroundColor(oddRowBackgroundColor())
            .withFontColor(oddRowFontColor())
            .applyIf(isOddRowBold(), CellStyler::isBold)
            .as(ODD_ROW)
        ;
    }

    default String getHeaderStyle() {
        return COLUMN_HEADER;
    }

    default String getDescriptionStyle() {
        return REPORT_DESCRIPTION;
    }
}
