package com.wesleyhome.poi.api;

import java.awt.*;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STTotalsRowFunction;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STTotalsRowFunction.Enum;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.impl.STTotalsRowFunctionImpl;

/**
 * Demonstrates how to create a simple table using Apache POI.
 */
public class CreateTable {

    public static void main(String[] args) throws IOException {

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet();

            // Set which area the table should be placed in
            AreaReference reference = wb.getCreationHelper().createAreaReference(
                new CellReference(0, 0), new CellReference(2, 2));

            // Create
            XSSFTable table = sheet.createTable(reference);

            table.setStyleName("TableStyleMedium6");
            table.setName("Test");
            table.setDisplayName("Test_Table");

//            table.sty
            // For now, create the initial style in a low-level way
//            table.getCTTable().addNewTableStyleInfo();
//            table.getCTTable().getTableStyleInfo().setName("TableStyleMedium2");

            // Style the table
            XSSFTableStyleInfo style = (XSSFTableStyleInfo) table.getStyle();
            style.setName("TableStyleMedium6");
            style.setShowColumnStripes(false);
            style.setShowRowStripes(true);
            table.getCTTable().setTotalsRowShown(true);
            table.getCTTable().setTotalsRowCount(1);
            table.getCTTable().addNewAutoFilter();
            boolean hasTotalsRow = table.isHasTotalsRow();
//            style.setFirstColumn(false);
//            style.setLastColumn(false);
//            style.setShowRowStripes(true);
//            style.setShowColumnStripes(true);

            // Set the values for the table
            XSSFRow row;
            XSSFCell cell;
            for (int i = 0; i < 3; i++) {
                // Create row
                row = sheet.createRow(i);
                for (int j = 0; j < 3; j++) {
                    // Create cell
                    cell = row.createCell(j);
                    if (i == 0) {
                        cell.setCellValue("Column" + (j + 1));
                    } else {
                        cell.setCellValue((i + 1.0) * (j + 1.0));
                    }
                }
            }
            // Save
            Path path = Paths.get("ooxml-table.xlsx");
            try (OutputStream fileOut = Files.newOutputStream(path)) {
                wb.write(fileOut);
            }
            Desktop.getDesktop().open(path.toFile());
        }
    }
}