package com.wesleyhome.poi.api;

import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import static com.wesleyhome.poi.api.assertions.WorkbookAssert.assertThat;

class WorkbookGeneratorTest {

    @Test
    void createBasicWorkbook() {

        assertThat(this.getBasicWorkbook()).hasSheet("Sheet1")
            .cell("A1").hasValue("I am awesome")
            .cell("B1").hasValue("You are awesome")
            .cell("A2").hasValue("So are you!!!")
            .cell("B2").hasValue("So am I!!!")
            .cell("A3").hasValue("A3");
    }

    private Workbook getBasicWorkbook() {
        return WorkbookGenerator
            .create()
            .nextCell()
            .havingValue("I am awesome")
            .nextCell()
            .havingValue("You are awesome")
            .nextRow()
            .nextCell()
            .havingValue("So are you!!!")
            .nextCell()
            .havingValue("So am I!!!")
            .cell("A3")
            .havingValue("A3")
            .createWorkbook();
    }


//    @Test
//    void create() {
//        Workbook actual = WorkbookGenerator.create()
//            .withSheetName("TestSheet").withStartRow(2)
//                .createNextRow().withStartColumn("B")
//                    .withNextCell().havingValue("Name").withBackgroundColor(IndexedColors.DARK_BLUE).withFontColor(IndexedColors.WHITE).isBold().as("HEADER")
//                    .withNextCell().havingValue("Birthday").usingStyle("HEADER")
//                    .withNextCell().havingValue("Age").usingStyle("HEADER")
//                    .withNextCell().havingValue("Salary").usingStyle("HEADER")
//                    .withNextCell().havingValue("Hours in Service").usingStyle("HEADER")
//                .createNextRow()
//                    .withNextCell().havingValue("Justin").withNoBackgroundColor().withFontColor(IndexedColors.VIOLET).isBold().as("ODD_ROW")
//                    .withNextCell().havingValue(LocalDate.of(1976, Month.JUNE, 20)).usingStyle("ODD_ROW").withDateFormat()
//                    .withNextCell().havingValue(42).usingStyle("ODD_ROW").withIntegerFormat()
//                    .withNextCell().havingValue(142000D).usingStyle("ODD_ROW").withCurrencyFormat()
//                    .withNextCell().havingValue(352123.54).usingStyle("ODD_ROW").withNumericStyle()
//            .createNewSheet("New Sheet")
//                .createNextRow()
//                    .withNextCell().havingValue("Awesome!!!")
//            .createWorkbook();
//    }
}