package com.wesleyhome.poi.api.report.annotations;

import com.wesleyhome.poi.api.report.ColumnType;
import com.wesleyhome.poi.api.report.IdentityFormatter;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STTotalsRowFunction;

import java.lang.annotation.*;
import java.util.function.Function;

@Target({ElementType.FIELD, ElementType.METHOD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ReportColumn {

    String NULL = "__NULL__";

    String columnHeader() default NULL;

    String columnIdentifier() default NULL;

    ColumnType columnType() default ColumnType.DERIVED;

    TotalsRowFunction totalFunction() default TotalsRowFunction.F_NONE;

    Class<? extends Function<?, ?>> formatter() default IdentityFormatter.class;
}
