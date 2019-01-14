package com.wesleyhome.poi.api.report.annotations;

import com.wesleyhome.poi.api.report.DefaultReportStyler;
import com.wesleyhome.poi.api.report.ReportStyler;

import java.lang.annotation.*;

@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Report {

    String NULL = "__NULL__";

    String sheetName() default NULL;

    String description() default NULL;

    Class<? extends ReportStyler> styler() default DefaultReportStyler.class;

}
