package com.zeki.merger.service.data;

import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.List;

/**
 * Low-level Apache POI cell reading utilities.
 *
 * Extracted to eliminate the repeated cellStr/cellDouble patterns found in
 * DataReader, ProcreancesComparator, and EtatPublicGenerator.
 */
public class DataExtractor {

    private final DataFormatter formatter = new DataFormatter();

    /** Reads a cell value as trimmed String. Returns "" on null. */
    public String extractString(Row row, int col, FormulaEvaluator evaluator) {
        if (row == null) return "";
        Cell cell = row.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (cell == null) return "";
        try {
            return formatter.formatCellValue(cell, evaluator).trim();
        } catch (Exception e) {
            return "";
        }
    }

    /** Reads a cell value as double. Returns 0.0 on null or parse error. */
    public double extractDouble(Row row, int col, FormulaEvaluator evaluator) {
        if (row == null) return 0.0;
        Cell cell = row.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (cell == null) return 0.0;
        try {
            CellType type = cell.getCellType() == CellType.FORMULA
                ? cell.getCachedFormulaResultType()
                : cell.getCellType();
            if (type == CellType.NUMERIC) return cell.getNumericCellValue();
            return com.zeki.merger.trf.model.ConsolidationRow
                .parseFrenchDouble(formatter.formatCellValue(cell, evaluator).trim());
        } catch (Exception e) {
            return 0.0;
        }
    }

    /** Reads multiple string columns from a row in one call. */
    public List<String> extractStrings(Row row, FormulaEvaluator evaluator, int... columns) {
        List<String> result = new ArrayList<>(columns.length);
        for (int col : columns) {
            result.add(extractString(row, col, evaluator));
        }
        return result;
    }

    /** Reads a string from a specific (rowIndex, colIndex) position on the sheet. */
    public String extractSheetCell(Sheet sheet, int rowIndex, int colIndex,
                                   FormulaEvaluator evaluator) {
        Row row = sheet.getRow(rowIndex);
        return extractString(row, colIndex, evaluator);
    }
}
