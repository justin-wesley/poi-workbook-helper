package com.wesleyhome.poi.api.report;

import org.apache.commons.lang3.RandomStringUtils;
import org.apache.commons.lang3.RandomUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import java.awt.*;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import static org.apache.commons.lang3.RandomUtils.nextInt;

class ReportGeneratorTest {

    private static final String URLS[] = {
        "http://www.amazon.com",
        "http://www.cnn.com",
        "http://www.google.com"
    };

    @Test
    void testDefaultDataPropertyAccessor() throws Exception{
        List<TestDefaultDataPropertyAccessor> records = dpa(nextInt(20, 41));
        ReportConfiguration<TestDefaultDataPropertyAccessor> reportConfiguration = new AnnotatedReportConfiguration<>(TestDefaultDataPropertyAccessor.class);
        ReportGenerator reportGenerator = new DefaultReportGenerator().applyReport(records, reportConfiguration);

        Path testReportDirectory = Paths.get("target", "test", "reports");
        Files.createDirectories(testReportDirectory);
        Path reportPath = Files.createTempFile(testReportDirectory, "DefaultReport", ".xlsx");

        try(OutputStream os = Files.newOutputStream(reportPath); Workbook workbook = reportGenerator.create()){
            workbook.write(os);
        }
        System.out.println(reportPath.toAbsolutePath());
        Desktop.getDesktop().open(reportPath.toFile());
    }

    @Test
    void testMultipleSheetsDefaultDataPropertyAccessor() throws Exception{
        ReportConfiguration<TestDefaultDataPropertyAccessor> reportConfiguration = new AnnotatedReportConfiguration<TestDefaultDataPropertyAccessor>(TestDefaultDataPropertyAccessor.class){
            @Override
            public String getReportSheetName() {
                return super.getReportSheetName()+RandomStringUtils.randomAlphabetic(3);
            }
        };
        ReportGenerator reportGenerator = new DefaultReportGenerator()
            .applyReport(this.dpa(nextInt(20, 41)), reportConfiguration)
            .applyReport(this.dpa(nextInt(20, 41)), reportConfiguration);
        Path testReportDirectory = Paths.get("target", "test", "reports");
        Files.createDirectories(testReportDirectory);
        Path reportPath = Files.createTempFile(testReportDirectory, "DefaultReport", ".xlsx");
        try(OutputStream os = Files.newOutputStream(reportPath); Workbook workbook = reportGenerator.create()){
            workbook.write(os);
        }
        System.out.println(reportPath.toAbsolutePath());
        Desktop.getDesktop().open(reportPath.toFile());
    }

    private List<TestDefaultDataPropertyAccessor> dpa(int numberOfRecords) {
        return IntStream.rangeClosed(1, numberOfRecords)
            .boxed()
            .map(this::createDpa)
            .collect(Collectors.toList());
    }

    private TestDefaultDataPropertyAccessor createDpa(Integer id) {
        LocalDateTime now = LocalDateTime.now();
        LocalDate localDate = now.toLocalDate();
        LocalDate birthdate = localDate.minusYears(nextInt(0, 100));
        LocalDate oldDate = localDate.minusYears(30);
        boolean isOld = oldDate.isAfter(birthdate);
        return TestDefaultDataPropertyAccessor.builder()
            .id(id)
            .name(RandomStringUtils.randomAlphabetic(6,30))
            .birthdate(birthdate)
            .now(now)
            .old(isOld)
            .averageScore(RandomUtils.nextDouble(56d, 300d))
            .bankBalance(RandomUtils.nextDouble(100d, 100_000d))
            .website(URLS[nextInt(0, 2)])
            .build();
    }


}