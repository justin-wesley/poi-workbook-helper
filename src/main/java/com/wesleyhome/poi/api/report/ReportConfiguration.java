package com.wesleyhome.poi.api.report;

import com.wesleyhome.poi.api.CellGenerator;
import com.wesleyhome.poi.api.CellStyler;
import com.wesleyhome.poi.api.internal.DefaultWorkbookGenerator;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;
import java.util.stream.Collectors;

import static com.wesleyhome.poi.api.internal.DefaultWorkbookGenerator.BuiltinStyles.*;
import static com.wesleyhome.poi.api.internal.DefaultWorkbookGenerator.BuiltinStyles.DATE;

public interface ReportConfiguration<T> {

    String COLUMN_HEADER = "COLUMN_HEADER";
    String EVEN_ROW = "EVEN_ROW";
    String ODD_ROW = "ODD_ROW";

    SortedMap<String, ColumnConfiguration<T>> columns();

    default ColumnConfiguration<T> getColumnConfiguration(String columnIdentifier) {
        return columns().get(columnIdentifier);
    }

    String getReportSheetName();

    default Map<String, String> columnHeaders() {
        return columns().values()
            .stream()
            .collect(Collectors.toMap(ColumnConfiguration::getColumnName, ColumnConfiguration::getColumnHeader));
    }

    IndexedColors headerBackgroundColor();

    IndexedColors headerFontColor();

    boolean isHeaderBold();

    IndexedColors evenRowBackgroundColor();

    IndexedColors evenRowFontColor();

    boolean isEvenRowBold();

    IndexedColors oddRowBackgroundColor();

    IndexedColors oddRowFontColor();

    boolean isOddRowBold();

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
}
