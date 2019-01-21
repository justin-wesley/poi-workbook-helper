package com.wesleyhome.poi.api.report;

import com.wesleyhome.poi.api.report.annotations.Report;
import com.wesleyhome.poi.api.report.annotations.ReportColumn;
import com.wesleyhome.poi.api.report.annotations.TotalsRowFunction;
import lombok.Builder;
import lombok.Data;

import java.time.LocalDate;
import java.time.LocalDateTime;

@Data
@Builder
@Report
class TestDefaultDataPropertyAccessor {

    @ReportColumn(columnIdentifier = "A")
    private Integer id;
    @ReportColumn(columnIdentifier = "B")
    private String name;
    @ReportColumn(columnIdentifier = "C")
    private LocalDate birthdate;
    @ReportColumn(columnIdentifier = "D")
    private LocalDateTime now;
    @ReportColumn(columnIdentifier = "E")
    private boolean old;
    @ReportColumn(columnIdentifier = "F", totalFunction = TotalsRowFunction.F_SUM)
    private Double averageScore;
    @ReportColumn(columnIdentifier = "G", totalFunction = TotalsRowFunction.F_AVERAGE)
    private Double bankBalance;
    @ReportColumn(columnIdentifier = "H")
    private String website;
}
