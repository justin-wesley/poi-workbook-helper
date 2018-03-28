package com.wesleyhome.poi.api;

import com.wesleyhome.poi.api.assertions.WorkbookAssert;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import java.util.function.Supplier;

import static com.wesleyhome.poi.api.assertions.WorkbookAssert.assertThat;
import static org.assertj.core.api.Assertions.assertThat;

class WorkbookGeneratorTest {

    @Test
    void createBasicWorkbookWithOneDefinedCell(){
        Supplier<Workbook> supplier = () -> getWorkbook();

        assertThat(supplier).hasSheet("sheet0").cell("A1").hasValue("I am awesome").cell("A2");
    }

    private Workbook getWorkbook() {
        return WorkbookGenerator
            .create()
            .cell("A1")
            .havingValue("I am awesome")
            .cell("A2")
            .havingValue("So are you!!!")
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