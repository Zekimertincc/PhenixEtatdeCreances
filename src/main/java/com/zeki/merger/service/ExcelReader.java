package com.zeki.merger.service;

import com.zeki.merger.AppConfig;
import com.zeki.merger.model.CreanceRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

/**
 * Reads an Excel file with Apache POI and returns:
 * <ul>
 *   <li>the header row (row 0) as a list of strings</li>
 *   <li>all data rows where column S (index 18) is non-empty, skipping row 0</li>
 * </ul>
 */
public class ExcelReader {

    /**
     * Returns the header row from the first sheet (row index 0).
     */
    public List<String> readHeader(File excelFile) throws IOException {
        try (Workbook wb = openWorkbook(excelFile)) {
            Sheet sheet = wb.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) return List.of();

            DataFormatter fmt = new DataFormatter();
            List<String> headers = new ArrayList<>();
            for (int c = 0; c < headerRow.getLastCellNum(); c++) {
                Cell cell = headerRow.getCell(c);
                headers.add(cell != null ? fmt.formatCellValue(cell).trim() : "");
            }
            return headers;
        }
    }

    /**
     * Reads every data row (starting at index 1) whose value in column
     * {@link AppConfig#FILTER_COLUMN_INDEX} is non-blank, and wraps each in a
     * {@link CreanceRow} tagged with {@code companyName}.
     */
    public List<CreanceRow> readFiltered(String companyName, File excelFile) throws IOException {
        List<CreanceRow> rows = new ArrayList<>();

        try (Workbook wb = openWorkbook(excelFile)) {
            Sheet sheet = wb.getSheetAt(0);
            DataFormatter fmt = new DataFormatter();
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {   // row 0 = header
                Row row = sheet.getRow(r);
                if (row == null) continue;

                Cell filterCell = row.getCell(AppConfig.FILTER_COLUMN_INDEX);
                if (!hasRealData(filterCell)) continue;

                List<Object> values = extractRowValues(row, fmt, evaluator);
                rows.add(new CreanceRow(companyName, values, r));
            }
        }

        if (rows.isEmpty()) {
            System.out.println("[" + companyName + "] SKIPPED - no data in column S");
        }
        return rows;
    }

    // -------------------------------------------------------------------------

    private Workbook openWorkbook(File file) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        return file.getName().toLowerCase().endsWith(".xls")
            ? new HSSFWorkbook(fis)
            : new XSSFWorkbook(fis);
    }

    /**
     * Returns true only when the cell contains a genuinely meaningful value:
     * - non-blank STRING that is not empty/whitespace
     * - NUMERIC that is not zero (zero = formula placeholder with no real data)
     * - BOOLEAN (any boolean is real data)
     *
     * This prevents financial templates where column S is pre-filled with
     * SUM/IF formulas that evaluate to 0 when no data is entered from being
     * treated as having data.
     */
    /**
     * Values that appear in column S as template sub-headers (not real data).
     * Case-insensitive comparison is used.
     */
    private static final java.util.Set<String> HEADER_LABELS = java.util.Set.of(
        "lieu", "location", "place"
    );

    /** Returns true when the Excel format string contains a known currency symbol. */
    private boolean isCurrencyFormat(String formatString) {
        if (formatString == null) return false;
        return formatString.contains("€")
            || formatString.contains("$")
            || formatString.contains("£")
            || formatString.contains("¥")
            || formatString.contains("₺");
    }

    private boolean hasRealData(Cell cell) {
        if (cell == null) return false;

        CellType type = cell.getCellType() == CellType.FORMULA
            ? cell.getCachedFormulaResultType()
            : cell.getCellType();

        return switch (type) {
            case BLANK   -> false;
            case BOOLEAN -> true;
            case NUMERIC -> cell.getNumericCellValue() != 0.0;
            case STRING  -> {
                String val = cell.getStringCellValue().trim();
                yield !val.isEmpty() && !HEADER_LABELS.contains(val.toLowerCase());
            }
            default      -> false;
        };
    }

    private static final java.util.Set<Integer> NUMERIC_STRING_COLS = java.util.Set.of(
        8, 9, 12, 13, 14, 18, 20, 21, 22, 23, 24
    );

    private Object coerceStringToDouble(String raw) {
        String cleaned = raw.replaceAll("[€$£\\s]", "")
                            .replace(".", "")
                            .replace(",", ".");
        try {
            return Double.parseDouble(cleaned);
        } catch (NumberFormatException e) {
            return raw;
        }
    }

    private List<Object> extractRowValues(Row row, DataFormatter fmt, FormulaEvaluator evaluator) {
        int lastCell = row.getLastCellNum();
        List<Object> values = new ArrayList<>(lastCell);

        for (int c = 0; c < lastCell; c++) {
            Cell cell = row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            if (cell == null) {
                values.add("");
                continue;
            }

            // Evaluate formula cells so we get the computed result, never the formula text
            CellType type;
            double   numericResult  = 0;
            String   stringResult   = "";
            boolean  booleanResult  = false;

            if (cell.getCellType() == CellType.FORMULA) {
                try {
                    org.apache.poi.ss.usermodel.CellValue cv = evaluator.evaluate(cell);
                    type = cv.getCellType();
                    switch (type) {
                        case NUMERIC -> numericResult = cv.getNumberValue();
                        case STRING  -> stringResult  = cv.getStringValue();
                        case BOOLEAN -> booleanResult = cv.getBooleanValue();
                        default      -> {}
                    }
                } catch (Exception e) {
                    // Formula evaluation failed (e.g. external references) — treat as blank
                    values.add("");
                    continue;
                }
            } else {
                type = cell.getCellType();
                switch (type) {
                    case NUMERIC -> numericResult = cell.getNumericCellValue();
                    case STRING  -> stringResult  = cell.getStringCellValue();
                    case BOOLEAN -> booleanResult = cell.getBooleanCellValue();
                    default      -> {}
                }
            }

            switch (type) {
                case NUMERIC -> {
                    if (DateUtil.isCellDateFormatted(cell)) {
                        values.add(DateUtil.getJavaDate(numericResult)
                            .toInstant()
                            .atZone(java.time.ZoneId.systemDefault())
                            .toLocalDateTime());
                    } else if (isCurrencyFormat(cell.getCellStyle().getDataFormatString())) {
                        // Preserve € symbol — format the evaluated numeric value
                        values.add(fmt.formatCellValue(cell, evaluator));
                    } else {
                        values.add(numericResult);
                    }
                }
                case BOOLEAN -> values.add(booleanResult);
                case STRING  -> values.add(
                    NUMERIC_STRING_COLS.contains(c) ? coerceStringToDouble(stringResult) : stringResult
                );
                case BLANK   -> values.add("");
                default      -> values.add(""); // ERROR veya bilinmeyen — boş bırak
            }
        }
        return values;
    }
}
