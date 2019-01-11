package com.wesleyhome.poi.api;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.extension.ExtendWith;
import org.mockito.Mock;
import org.mockito.junit.jupiter.MockitoExtension;

import static org.assertj.core.api.Assertions.*;
import static org.mockito.Mockito.*;
@ExtendWith(MockitoExtension.class)
class CellNameHelperTest {

    @Mock
    private Sheet sheet;
    @Mock
    private Row row;
    @Mock
    private Cell cell;

    @Test
    void convertToCellGoodValue(){
        when(sheet.getRow(9999)).thenReturn(row);
        when(row.getCell(CellReference.convertColStringToIndex("AA"))).thenReturn(cell);
        Cell aa10000 = CellNameHelper.convertToCell(sheet, "aa10000");
        assertThat(cell).isSameAs(aa10000);
    }

}