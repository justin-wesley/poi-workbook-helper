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

class ReportGeneratorTest {

    private ReportGenerator<TestDefaultDataPropertyAccessor> defaultDataReportGenerator = new DefaultReportGenerator<>();
    private static final String URLS[] = {
        "http://www.amazon.com",
        "http://www.cnn.com",
        "http://www.google.com"
    };

    @Test
    void testDefaultDataPropertyAccessor() throws Exception{
        List<TestDefaultDataPropertyAccessor> records = dpa(20);
        Path testReportDirectory = Paths.get("target", "test", "reports");
        Files.createDirectories(testReportDirectory);
        Path reportPath = Files.createTempFile(testReportDirectory, "DefaultReport", ".xlsx");
        ReportConfiguration<TestDefaultDataPropertyAccessor> reportConfiguration = new AnnotatedReportConfiguration<>(TestDefaultDataPropertyAccessor.class);
        try(OutputStream os = Files.newOutputStream(reportPath); Workbook workbook = defaultDataReportGenerator.generateWorkbook(records, reportConfiguration)){
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
        LocalDate birthdate = LocalDate.now().minusYears(RandomUtils.nextInt(0, 100));
        LocalDate oldDate = birthdate.minusYears(30);
        boolean isOld = oldDate.isAfter(birthdate);
        return TestDefaultDataPropertyAccessor.builder()
            .id(id)
            .name(RandomStringUtils.randomPrint(6,30))
            .birthdate(birthdate)
            .now(LocalDateTime.now())
            .old(isOld)
            .averageScore(RandomUtils.nextDouble(56d, 300d))
            .bankBalance(RandomUtils.nextDouble(100d, 100_000d))
            .website(URLS[RandomUtils.nextInt(0, 2)])
            .build();
    }


}