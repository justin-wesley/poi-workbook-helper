package com.wesleyhome.poi.api.internal;

import com.wesleyhome.poi.api.CellGenerator;
import com.wesleyhome.poi.api.SheetGenerator;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.ToString;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;

import java.util.stream.Stream;

@EqualsAndHashCode
@ToString
public class Table {
    private int startRow;
    private int startCol;
    private int endRow;
    private int endCol;

    public Table(CellGenerator currentCell) {
        this.startRow = currentCell.rowNum();
        this.startCol = currentCell.columnNum();
    }

    public void setEndCell(CellGenerator cell) {
        this.endRow = cell.rowNum();
        this.endCol = cell.columnNum();
    }

    public boolean isValid() {
        return Stream.of(startRow, startCol, endRow, endCol).allMatch(i->i>=0);
    }

    public AreaReference getAreaReference(boolean withTotalRow, Workbook wb) {
        return wb.getCreationHelper().createAreaReference(new CellReference(startRow, startCol),new CellReference(withTotalRow ? endRow + 1 : endRow, endCol));
    }
}
