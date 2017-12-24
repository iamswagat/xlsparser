package org.acme.commercial;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

public class CustomDataFormatter extends DataFormatter {

    /**
     * Return the cell type.
     *
     * @param cell
     * @param evaluator
     * @return the cell type
     */
    private CellType getCellType(Cell cell, FormulaEvaluator evaluator) {
        CellType cellType = cell.getCellTypeEnum();
        if (cellType == CellType.FORMULA) {
            cellType = evaluator.evaluateFormulaCellEnum(cell);
        }
        return cellType;
    }
}
