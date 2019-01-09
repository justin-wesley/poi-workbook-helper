package com.wesleyhome.poi.api;

import com.wesleyhome.poi.api.creator.WorkbookCreator;

import java.util.function.BiFunction;
import java.util.function.Function;

import static org.apache.poi.ss.util.CellReference.convertColStringToIndex;

public interface RowGenerator extends WorkbookCreator {
    default RowGenerator withStartColumn(String columnName) {
        return withStartColumn(convertColStringToIndex(columnName));
    }


    RowGenerator withStartColumn(int columnNum);

    CellGenerator cell(int columnNum);

    RowGenerator height(float rowHeight);

    int rowNum();

    default <T> RowGenerator generateCells(Iterable<T> iterable, BiFunction<CellGenerator, T, CellGenerator> rowGeneratorConsumer) {
        return GeneratorHelper.iterate(this.cell(), iterable, rowGeneratorConsumer, WorkbookCreator::nextCell).row();
    }

    default RowGenerator generateRow(Function<RowGenerator, RowGenerator> rowGeneratorFunction) {
        return rowGeneratorFunction.apply(this);
    }

    default <T> RowGenerator generateRows(Iterable<T> iterable, BiFunction<RowGenerator, T, RowGenerator> rowGeneratorConsumer) {
        return GeneratorHelper.iterate(this, iterable, rowGeneratorConsumer, WorkbookCreator::nextRow);
    }

    default RowGenerator generateCell(Function<CellGenerator, CellGenerator> cellGeneratorFunction) {
        return cellGeneratorFunction.apply(this.cell()).row();
    }
}
