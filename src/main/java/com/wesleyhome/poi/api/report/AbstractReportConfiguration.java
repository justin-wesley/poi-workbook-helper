package com.wesleyhome.poi.api.report;

import com.wesleyhome.poi.api.CellGenerator;
import com.wesleyhome.poi.api.CellStyler;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.SortedMap;

import static com.wesleyhome.poi.api.internal.DefaultWorkbookGenerator.BuiltinStyles.*;

public abstract class AbstractReportConfiguration<T> implements ReportConfiguration<T> {
    private final SafeLazyInitializer<String> sheetNameInitializer;
    private final SafeLazyInitializer<String> reportDescriptionInitializer;
    private final SafeLazyInitializer<ReportStyler> reportStylerInitializer;
    private final SafeLazyInitializer<SortedMap<String, ColumnConfiguration<T>>> columnInitializer;

    public AbstractReportConfiguration() {
        this.columnInitializer = new SafeLazyInitializer<SortedMap<String, ColumnConfiguration<T>>>() {
            @Override
            protected SortedMap<String, ColumnConfiguration<T>> safeInitialize() {
                return initializeColumns();
            }
        };
        this.reportStylerInitializer = new SafeLazyInitializer<ReportStyler>() {
            @Override
            protected ReportStyler safeInitialize() {
                return initializeReportStyler();
            }
        };
        this.sheetNameInitializer = new SafeLazyInitializer<String>() {
            @Override
            protected String safeInitialize() {
                return initializeSheetName();
            }
        };
        this.reportDescriptionInitializer = new SafeLazyInitializer<String>() {
            @Override
            protected String safeInitialize() {
                return initializeReportDescription();
            }
        };
    }

    protected abstract String initializeReportDescription();

    protected abstract String initializeSheetName();

    protected abstract ReportStyler initializeReportStyler();

    protected abstract SortedMap<String, ColumnConfiguration<T>> initializeColumns();

    @Override
    public final SortedMap<String, ColumnConfiguration<T>> columns() {
        return this.columnInitializer.get();
    }

    public final ReportStyler getReportStyler() {
        return reportStylerInitializer.get();
    }

    @Override
    public final String getReportSheetName() {
        return this.sheetNameInitializer.get();
    }

    @Override
    public final String getReportDescription() {
        return this.reportDescriptionInitializer.get();
    }

    @Override
    public String getHeaderStyleName() {
        return getReportStyler().getHeaderStyleName();
    }

    @Override
    public String getDescriptionStyleName() {
        return getReportStyler().getDescriptionStyleName();
    }

    public CellGenerator applyStyleAndValueToCell(CellGenerator cellGenerator, ColumnConfiguration<T> columnConfiguration, Object transformedValue) {
        int rowNum = cellGenerator.rowNum();
        boolean isEven = rowNum % 2 == 0;
        List<String> styles = new ArrayList<>(getReportStyler().getRowStyles(isEven));
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

    public void createStyles(CellStyler cellStyler) {
        getReportStyler().createStyles(cellStyler.reset());
    }




}
