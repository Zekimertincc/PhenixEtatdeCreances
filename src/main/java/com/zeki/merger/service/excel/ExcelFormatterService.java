package com.zeki.merger.service.excel;

import com.zeki.merger.trf.model.ConsolidationRow;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;

import java.time.LocalDateTime;
import java.util.Locale;

/**
 * Stateless helpers for reading and writing Excel cell values.
 *
 * Extracted from TrfSheetWriter and EtatPublicGenerator to eliminate
 * code duplication — both classes had identical writeValue() and fmtPdf() methods.
 */
public class ExcelFormatterService {

    /**
     * Writes a typed value (Double, Number, Boolean, LocalDateTime, String) into a cell.
     * Strings that look like French-formatted numbers are coerced to doubles.
     */
    public void writeValue(XSSFCell cell, Object val,
                           XSSFCellStyle defaultStyle, XSSFCellStyle dateStyle) {
        if (val instanceof Double d) {
            cell.setCellValue(d);
            cell.setCellStyle(defaultStyle);
            return;
        }
        if (val instanceof Number n) {
            cell.setCellValue(n.doubleValue());
            cell.setCellStyle(defaultStyle);
            return;
        }
        if (val instanceof Boolean b) {
            cell.setCellValue(b);
            cell.setCellStyle(defaultStyle);
            return;
        }
        if (val instanceof LocalDateTime ldt) {
            cell.setCellValue(ldt);
            cell.setCellStyle(dateStyle);
            return;
        }
        if (val instanceof String s && !s.isBlank()) {
            String stripped = s.replaceAll("[€$£¥₺]", "")
                               .replaceAll("\\p{Z}", "")
                               .trim();
            if (!stripped.isEmpty() && !stripped.equals("-")
                    && stripped.matches("[-+]?[\\d.,]+")) {
                cell.setCellValue(ConsolidationRow.parseFrenchDouble(s));
                cell.setCellStyle(defaultStyle);
                return;
            }
            cell.setCellValue(s);
            cell.setCellStyle(defaultStyle);
            return;
        }
        cell.setCellStyle(defaultStyle);
    }

    /** Formats a cell value for plain-text output (used by PDF and log display). */
    public String formatForDisplay(Object val) {
        if (val == null) return "";
        if (val instanceof Double d)        return String.format(Locale.FRANCE, "%,.2f", d);
        if (val instanceof Number n)        return String.format(Locale.FRANCE, "%,.2f", n.doubleValue());
        if (val instanceof LocalDateTime t) return t.toLocalDate().toString();
        if (val instanceof Boolean b)       return b ? "OUI" : "NON";
        return val.toString().trim();
    }

    /**
     * Reads a string value from a cell, returning "" on null or blank.
     */
    public String readString(Row row, int col, DataFormatter fmt, FormulaEvaluator eval) {
        if (row == null) return "";
        Cell cell = row.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (cell == null) return "";
        return fmt.formatCellValue(cell, eval).trim();
    }

    /**
     * Reads a numeric value from a cell. Falls back to French-number parsing for strings.
     */
    public double readDouble(Row row, int col, DataFormatter fmt, FormulaEvaluator eval) {
        if (row == null) return 0.0;
        Cell cell = row.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (cell == null) return 0.0;
        CellType type = cell.getCellType() == CellType.FORMULA
            ? cell.getCachedFormulaResultType()
            : cell.getCellType();
        if (type == CellType.NUMERIC) return cell.getNumericCellValue();
        return ConsolidationRow.parseFrenchDouble(fmt.formatCellValue(cell, eval).trim());
    }

    /** Converts a 0-based column index to an Excel column letter (A, B, …, Z, AA, …). */
    public static String columnLetter(int idx) {
        if (idx < 26) return String.valueOf((char) ('A' + idx));
        return String.valueOf((char) ('A' + idx / 26 - 1)) + (char) ('A' + idx % 26);
    }
}
