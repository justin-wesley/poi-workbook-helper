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
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import static org.apache.commons.lang3.RandomUtils.nextInt;

class ReportGeneratorTest {

    private static final String URLS[] = {
        "https://www.amazon.com/s/ref=nb_sb_noss_1?url=search-alias%3Daps&field-keywords=",
        "https://www.bing.com/search?q=",
        "https://www.google.com/search?q="
    };

    private static final String SEARCH_TERMS[] = {
        "nintendo switch",
        "fidget spinner",
        "laptop",
        "headphones",
        "fitbit",
        "ps4",
        "external hard drive",
        "bluetooth headphones",
        "instant pot",
        "kindle",
        "tv",
        "fire stick",
        "tablet",
        "micro sd card",
        "ipad",
        "ssd",
        "roku",
        "toilet paper",
        "backpack",
        "wireless headphones",
        "game of thrones",
        "iphone 7 case",
        "books",
        "monitor",
        "bluetooth speakers",
        "essential oils",
        "pop socket",
        "air fryer",
        "paper towels",
        "desk",
        "iphone charger",
        "water bottle",
        "doctor who",
        "xbox one",
        "iphone 7 plus case",
        "switch",
        "alexa",
        "harry potter",
        "coffee maker",
        "hdmi cable",
        "lego",
        "shoes",
        "kindle fire",
        "iphone x case",
        "wireless mouse",
        "printer",
        "iphone 6 case",
        "iphone 6",
        "iphone 7",
        "gaming mouse",
        "office chair",
        "star wars",
        "shower curtain",
        "mouse pad",
        "iphone 6s case",
        "gift card",
        "keyboard",
        "yoga mat",
        "gaming chair",
        "sd card",
        "earbuds",
        "echo",
        "socks",
        "xbox one controller",
        "apple watch",
        "protein powder",
        "camera",
        "solar eclipse glasses",
        "mouse",
        "drone",
        "eclipse glasses",
        "projector",
        "iphone",
        "iphone 8 plus case",
        "mattress",
        "playstation 4",
        "microwave",
        "coffee",
        "dash cam",
        "smart watch",
        "fidget cube",
        "vacuum cleaner",
        "laptops",
        "nerf guns",
        "bluetooth earbuds",
        "ps4 games",
        "router",
        "echo dot",
        "fan",
        "gtx 1070",
        "computer desk",
        "ps4 controller",
        "blender",
        "wireless earbuds",
        "iphone 6s",
        "snes classic",
        "prime video",
        "hard drive",
        "luggage",
        "gtx 1080"
    };

    @Test
    void testDefaultDataPropertyAccessor() throws Exception {
        List<TestDefaultDataPropertyAccessor> records = dpa(nextInt(20, 41));
        ReportConfiguration<TestDefaultDataPropertyAccessor> reportConfiguration = new AnnotatedReportConfiguration<>(TestDefaultDataPropertyAccessor.class);
        ReportGenerator reportGenerator = new DefaultReportGenerator().applyReport(records, reportConfiguration);

        Path testReportDirectory = Paths.get("target", "test", "reports");
        Files.createDirectories(testReportDirectory);
        Path reportPath = Files.createTempFile(testReportDirectory, "DefaultReport", ".xlsx");

        try (OutputStream os = Files.newOutputStream(reportPath); Workbook workbook = reportGenerator.create()) {
            workbook.write(os);
        }
        System.out.println(reportPath.toAbsolutePath());
        Desktop.getDesktop().open(reportPath.toFile());
    }

    @Test
    void testMultipleSheetsDefaultDataPropertyAccessor() throws Exception {
        ReportConfiguration<TestDefaultDataPropertyAccessor> reportConfiguration = new AnnotatedReportConfiguration<TestDefaultDataPropertyAccessor>(TestDefaultDataPropertyAccessor.class) {
            @Override
            public String getReportSheetName() {
                return super.getReportSheetName() + RandomStringUtils.randomAlphabetic(3);
            }
        };
        ReportGenerator reportGenerator = new DefaultReportGenerator()
            .applyReport(this.dpa(nextInt(20, 41)), reportConfiguration)
            .applyReport(this.dpa(nextInt(20, 41)), reportConfiguration);
        Path testReportDirectory = Paths.get("target", "test", "reports");
        Files.createDirectories(testReportDirectory);
        Path reportPath = Files.createTempFile(testReportDirectory, "DefaultReport", ".xlsx");
        try (OutputStream os = Files.newOutputStream(reportPath); Workbook workbook = reportGenerator.create()) {
            workbook.write(os);
        }
        System.out.println(reportPath.toAbsolutePath());
        Desktop.getDesktop().open(reportPath.toFile());
    }

    @Test
    void testMultipleTablesDefaultDataPropertyAccessor() throws Exception {
        AtomicBoolean flagged = new AtomicBoolean(true);
        ReportConfiguration<TestDefaultDataPropertyAccessor> reportConfiguration = new AnnotatedReportConfiguration<TestDefaultDataPropertyAccessor>(TestDefaultDataPropertyAccessor.class) {
        };
        ReportGenerator reportGenerator = new DefaultReportGenerator()
            .applyReport(this.dpa(nextInt(20, 41)), new AnnotatedReportConfiguration<>(TestDefaultDataPropertyAccessor.class))
            .applyReport(this.dpa(nextInt(20, 41)), new AnnotatedReportConfiguration<TestDefaultDataPropertyAccessor>(TestDefaultDataPropertyAccessor.class) {

                @Override
                protected String initializeReportTitle() {
                    return null;
                }

                @Override
                protected String initializeReportDescription() {
                    return null;
                }
            });
        Path testReportDirectory = Paths.get("target", "test", "reports");
        Files.createDirectories(testReportDirectory);
        Path reportPath = Files.createTempFile(testReportDirectory, "DefaultReport", ".xlsx");
        try (OutputStream os = Files.newOutputStream(reportPath); Workbook workbook = reportGenerator.create()) {
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
            .name(RandomStringUtils.randomAlphabetic(6, 30))
            .birthdate(birthdate)
            .now(now)
            .old(isOld)
            .averageScore(RandomUtils.nextDouble(56d, 300d))
            .bankBalance(RandomUtils.nextDouble(100d, 100_000d))
            .searchEngineSite(URLS[nextInt(0, URLS.length)])
            .searchTerm(SEARCH_TERMS[nextInt(0, SEARCH_TERMS.length)])
            .build();
    }


}