package com.wesleyhome.poi.api.report;

import com.google.common.base.Function;
import com.wesleyhome.poi.api.report.annotations.ReportColumn;
import org.checkerframework.checker.nullness.qual.Nullable;

public class DerivedFormatter implements Function<Object, Object> {

    private final ReportColumn reportColumn;

    public DerivedFormatter(ReportColumn reportColumn) {
        this.reportColumn = reportColumn;
    }

    @Nullable
    @Override
    public Object apply(@Nullable Object input) {
        return null;
    }
}
