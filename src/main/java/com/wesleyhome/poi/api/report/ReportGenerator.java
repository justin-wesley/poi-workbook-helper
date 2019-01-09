package com.wesleyhome.poi.api.report;

import org.apache.poi.ss.usermodel.Workbook;

public interface ReportGenerator<T> {

    Workbook generateWorkbook(Iterable<T> data, ReportConfiguration<T> reportConfiguration);
}
