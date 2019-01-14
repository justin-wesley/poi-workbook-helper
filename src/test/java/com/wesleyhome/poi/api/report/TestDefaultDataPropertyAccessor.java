package com.wesleyhome.poi.api.report;

import com.wesleyhome.poi.api.report.annotations.Report;
import com.wesleyhome.poi.api.report.annotations.ReportColumn;
import lombok.Builder;
import lombok.Data;

import java.time.LocalDate;
import java.time.LocalDateTime;

@Data
@Builder
@Report
class TestDefaultDataPropertyAccessor {

    @ReportColumn
    private Integer id;
    @ReportColumn
    private String name;
    @ReportColumn
    private LocalDate birthdate;
    @ReportColumn
    private LocalDateTime now;
    @ReportColumn
    private boolean old;
    @ReportColumn
    private Double averageScore;
    @ReportColumn
    private Double bankBalance;
    @ReportColumn
    private String website;
}
