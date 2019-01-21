package com.wesleyhome.poi.api.report;

import org.apache.poi.ss.usermodel.Workbook;
import org.checkerframework.checker.units.qual.K;

import java.util.Map;
import java.util.SortedMap;

public interface ReportGenerator {

    <T> ReportGenerator applyReport(Iterable<T> data, ReportConfiguration<T> reportConfiguration);

//    <K extends Comparable<K>> Workbook generateWorkbooks(SortedMap<K, Iterable<?>> data, SortedMap<K, ReportConfiguration<?>> reportConfiguration);
//
//    <T> Workbook generateWorkbook(Iterable<T> data, ReportConfiguration<T> reportConfiguration);

    Workbook create();
}
