package com.wesleyhome.poi.api.report;

import com.wesleyhome.poi.api.CellGenerator;
import com.wesleyhome.poi.api.CellStyler;
import com.wesleyhome.poi.api.internal.TableConfiguration;

import java.sql.Timestamp;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;

import static com.wesleyhome.poi.api.internal.DefaultWorkbookGenerator.BuiltinStyles.*;

public abstract class AbstractReportConfiguration<T> implements ReportConfiguration<T> {
    private final SafeLazyInitializer<String> sheetNameInitializer;
    private final SafeLazyInitializer<String> reportDescriptionInitializer;
    private final SafeLazyInitializer<ReportStyler> reportStylerInitializer;
    private final SafeLazyInitializer<SortedSet<ColumnConfiguration<T>>> columnInitializer;
    private final SafeLazyInitializer<String> reportTitleInitializer;
    private final SafeLazyInitializer<TableConfiguration> tableConfigurationInitializer;

    public AbstractReportConfiguration() {
        this.columnInitializer = new SafeLazyInitializer<SortedSet<ColumnConfiguration<T>>>() {
            @Override
            protected SortedSet<ColumnConfiguration<T>> safeInitialize() {
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
        this.reportTitleInitializer = new SafeLazyInitializer<String>() {
            @Override
            protected String safeInitialize() {
                return initializeReportTitle();
            }
        };
        this.reportDescriptionInitializer = new SafeLazyInitializer<String>() {
            @Override
            protected String safeInitialize() {
                return initializeReportDescription();
            }
        };
        this.tableConfigurationInitializer = new SafeLazyInitializer<TableConfiguration>() {
            @Override
            protected TableConfiguration safeInitialize() {
                return initializeTableConfiguration();
            }
        };
    }

    @Override
    public ColumnConfiguration<T> getColumnConfiguration(String propertyName) {
        return columns()
            .stream()
            .filter(c->c.getPropertyName().equals(propertyName))
            .findFirst()
            .orElseThrow(()->new IllegalArgumentException("Cannot find Property Name "+propertyName));
    }

    @Override
    public TableConfiguration getTableConfiguration() {
        return this.tableConfigurationInitializer.get();
    }

    protected TableConfiguration initializeTableConfiguration() {
        return unsupported("TableConfiguration");
    }

    @Override
    public String getReportTitle() {
        return this.reportTitleInitializer.get();
    }

    protected String initializeReportTitle() {
        return unsupported("ReportTitle");
    }

    protected abstract String initializeReportDescription();

    protected abstract String initializeSheetName();

    protected abstract ReportStyler initializeReportStyler();

    protected abstract SortedSet<ColumnConfiguration<T>> initializeColumns();

    @Override
    public final SortedSet<ColumnConfiguration<T>> columns() {
        return this.columnInitializer.get();
    }

    public final ReportStyler getReportStyler() {
        return reportStylerInitializer.get();
    }

    @Override
    public String getReportSheetName() {
        return this.sheetNameInitializer.get();
    }

    @Override
    public String getReportTitleStyleName() {
        return getReportStyler().getReportTitleStyleName();
    }

    @Override
    public final String getReportDescriptionDetail() {
        return hasReportDescriptionDetails() ? this.reportDescriptionInitializer.get() : null;
    }

    @Override
    public boolean hasReportDescriptionDetails() {
        String reportDescription = this.reportDescriptionInitializer.get();
        return reportDescription != null && !Objects.equals(reportDescription, this.reportTitleInitializer.get());
    }

    @Override
    public String getReportDescriptionDetailStyleName() {
        return getReportStyler().getDescriptionStyleName();
    }

    public CellGenerator applyStyleAndValueToCell(CellGenerator cellGenerator, ColumnConfiguration<T> columnConfiguration, Object transformedValue) {
        List<String> styles = new ArrayList<>();
        ColumnType columnType = columnConfiguration.getColumnType();
        if (ColumnType.DERIVED.equals(columnType) && transformedValue != null) {
            if (transformedValue instanceof Timestamp || transformedValue instanceof LocalDateTime) {
                columnType = ColumnType.TIMESTAMP;
            } else if (transformedValue instanceof Date || transformedValue instanceof LocalDate) {
                columnType = ColumnType.DATE;
            } else if (transformedValue instanceof Integer) {
                columnType = ColumnType.INTEGER;
            } else if (transformedValue instanceof Number) {
                columnType = ColumnType.DECIMAL;
            } else if (transformedValue instanceof Boolean) {
                columnType = ColumnType.BOOLEAN;
            } else {
                columnType = ColumnType.TEXT;
            }
        }
        switch (columnType) {
            case DECIMAL:
                styles.add(NUMERIC);
                break;
            case TIMESTAMP:
                styles.add(DATE_TIME);
                break;
            case DATE:
                styles.add(DATE);
                break;
            case INTEGER:
                styles.add(INTEGER);
                break;
            case ZIP:
                styles.add(ZIP);
                break;
            case URL:
                styles.add(URL);
                break;
            case TEXT:
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

    private <T> T unsupported(String property) {
        throw new UnsupportedOperationException(String.format("You must override get%1$s or initialize%1$s", property));
    }


}
