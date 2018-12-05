package com.wesleyhome.poi.api;

import com.wesleyhome.poi.api.creator.WorkbookCreator;

import java.util.Iterator;
import java.util.function.BiConsumer;
import java.util.function.BiFunction;
import java.util.function.Consumer;
import java.util.function.Function;

public interface SheetGenerator extends WorkbookCreator {

    SheetGenerator withSheetName(String test_sheet);

    /**
     * @param rowNum (0-based)
     * @return
     */
    SheetGenerator withStartRow(int rowNum);

    default SheetGenerator generateSheet(Function<SheetGenerator, SheetGenerator> sheetGeneratorFunction) {
        return sheetGeneratorFunction.apply(this);
    }

    default <T> SheetGenerator generateRows(Iterable<T> iterable, BiFunction<RowGenerator, T, RowGenerator> rowGeneratorConsumer) {
        RowGenerator rg = this.row();
        Iterator<T> i = iterable.iterator();
        while(i.hasNext()) {
            T t = i.next();
            rg = rowGeneratorConsumer.apply(rg, t);
            if(i.hasNext()){
                rg = rg.nextRow();
            }
        }
        return rg.sheet();
    }

    default SheetGenerator generateRow(Function<RowGenerator, RowGenerator> rowGeneratorFunction) {
        return rowGeneratorFunction.apply(this.row()).sheet();
    }

    default SheetGenerator generateCell(Function<CellGenerator, CellGenerator> cellGeneratorFunction) {
        return cellGeneratorFunction.apply(this.cell()).sheet();
    }

    SheetGenerator autosize(int columnNum);

    SheetGenerator hide(int columnNum);
}
