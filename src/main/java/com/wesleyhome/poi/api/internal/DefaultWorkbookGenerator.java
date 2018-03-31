package com.wesleyhome.poi.api.internal;

import com.wesleyhome.poi.api.CellGenerator;
import com.wesleyhome.poi.api.RowGenerator;
import com.wesleyhome.poi.api.SheetGenerator;
import com.wesleyhome.poi.api.WorkbookGenerator;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class DefaultWorkbookGenerator implements WorkbookGenerator {

    private final CellStyleManager cellStyleManager;
    private ExtendedMap<String, DefaultSheetGenerator> sheets;
    private SheetGenerator workingSheet;

    public DefaultWorkbookGenerator(){
        sheets = new ExtendedTreeMap<>();
        this.cellStyleManager = new CellStyleManager();
    }

    @Override
    public SheetGenerator sheet(String sheetName) {
        String safeSheetName = WorkbookUtil.createSafeSheetName(sheetName);
        workingSheet = sheets.getOrDefault(safeSheetName, () -> new DefaultSheetGenerator(this, safeSheetName));
        return workingSheet;
    }

    @Override
    public Workbook createWorkbook() {
        Workbook workbook = new SXSSFWorkbook();
        sheets.values()
            .forEach(sheetGen -> {
                sheetGen.applySheet(workbook);
            });
        return workbook;
    }

    @Override
    public RowGenerator nextRow() {
        return workingSheet.nextRow();
    }

    @Override
    public RowGenerator row(int rowNum) {
        return workingSheet.row(rowNum);
    }

    @Override
    public CellGenerator nextCell() {
        return workingSheet.nextCell();
    }

    @Override
    public CellGenerator cell(int columnNumber, int rowNumber) {
        return workingSheet.cell(columnNumber, rowNumber);
    }

    @Override
    public CellStyleManager cellStyleManager() {
        return cellStyleManager;
    }
}
