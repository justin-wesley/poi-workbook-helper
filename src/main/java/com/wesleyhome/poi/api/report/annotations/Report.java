package com.wesleyhome.poi.api.report.annotations;

import com.wesleyhome.poi.api.TableStyle;
import com.wesleyhome.poi.api.report.DefaultReportStyler;
import com.wesleyhome.poi.api.report.ReportStyler;

import java.lang.annotation.*;

import static com.wesleyhome.poi.api.TableStyle.TABLE_STYLE_MEDIUM_6;

@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Report {

    String NULL = "__NULL__";

    String sheetName() default NULL;

    String title() default NULL;

    String description() default NULL;

    Class<? extends ReportStyler> styler() default DefaultReportStyler.class;

    TableStyle tableStyle() default TABLE_STYLE_MEDIUM_6;
}
