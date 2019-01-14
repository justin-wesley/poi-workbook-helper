package com.wesleyhome.poi.api;

import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;

import static com.wesleyhome.poi.api.assertions.WorkbookAssert.assertThat;
import static org.apache.poi.ss.usermodel.IndexedColors.*;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFTableStyleInfo;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
class WorkbookGeneratorTest {

//    @Test
//    void createTableInWorkbook() throws IOException {
//        try (Workbook wb = new XSSFWorkbook()) {
//            XSSFSheet sheet = (XSSFSheet) wb.createSheet();
//
//            // Create
//            XSSFTable table = sheet.createTable();
//            table.setName("Test");
//            table.setDisplayName("Test_Table");
//
//            // For now, create the initial style in a low-level way
//            table.getCTTable().addNewTableStyleInfo();
//            table.getCTTable().getTableStyleInfo().setName("TableStyleMedium2");
//
//            // Style the table
//            XSSFTableStyleInfo style = (XSSFTableStyleInfo) table.getStyle();
//            style.setName("TableStyleMedium2");
//            style.setShowColumnStripes(false);
//            style.setShowRowStripes(true);
//            style.setFirstColumn(false);
//            style.setLastColumn(false);
//            style.setShowRowStripes(true);
//            style.setShowColumnStripes(true);
//
//            // Set the values for the table
//            XSSFRow row;
//            XSSFCell cell;
//            for (int i = 0; i < 3; i++) {
//                // Create row
//                row = sheet.createRow(i);
//                for (int j = 0; j < 3; j++) {
//                    // Create cell
//                    cell = row.createCell(j);
//                    if (i == 0) {
//                        cell.setCellValue("Column" + (j + 1));
//                    } else {
//                        cell.setCellValue((i + 1) * (j + 1));
//                    }
//                }
//            }
//            // Create the columns
//            table.createColumn("Column 1");
//            table.createColumn("Column 2");
//            table.createColumn("Column 3");
//
//            // Set which area the table should be placed in
//            AreaReference reference = wb.getCreationHelper().createAreaReference(
//                new CellReference(0, 0), new CellReference(2, 2));
//
//            table.setCellReferences(reference);
//
//            // Save
//            try (FileOutputStream fileOut = new FileOutputStream("ooxml-table.xlsx")) {
//                wb.write(fileOut);
//            }
//        }
//    }

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
        SheetGenerator sheetGenerator = WorkbookGenerator
            .create()
            .generateStyles(styler -> styler.withBackgroundColor(DARK_BLUE)
                .withFontColor(WHITE)
                .as("base")
                .withBackgroundColor(RED)
                .as("red"));
        return sheetGenerator
            .nextCell()
            .havingValue("I am awesome")
            .usingStyle("base")
            .nextCell()
            .havingValue("You are awesome")
            .usingStyle("red_base")
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
//                    .withNextCell().havingValue("Name").withBackgroundColor(IndexedColors.DARK_BLUE).withFontColor(IndexedColors.WHITE).withBoldFont().as("HEADER")
//                    .withNextCell().havingValue("Birthday").usingStyle("HEADER")
//                    .withNextCell().havingValue("Age").usingStyle("HEADER")
//                    .withNextCell().havingValue("Salary").usingStyle("HEADER")
//                    .withNextCell().havingValue("Hours in Service").usingStyle("HEADER")
//                .createNextRow()
//                    .withNextCell().havingValue("Justin").withNoBackgroundColor().withFontColor(IndexedColors.VIOLET).withBoldFont().as("ODD_ROW")
//                    .withNextCell().havingValue(LocalDate.of(1976, Month.JUNE, 20)).usingStyle("ODD_ROW").withDateTimeFormat()
//                    .withNextCell().havingValue(42).usingStyle("ODD_ROW").withIntegerFormat()
//                    .withNextCell().havingValue(142000D).usingStyle("ODD_ROW").withCurrencyFormat()
//                    .withNextCell().havingValue(352123.54).usingStyle("ODD_ROW").withNumericStyle()
//            .createNewSheet("New Sheet")
//                .createNextRow()
//                    .withNextCell().havingValue("Awesome!!!")
//            .createWorkbook();
//    }
}