package com.zeki.merger.service;

import com.zeki.merger.AppConfig;
import com.zeki.merger.model.CreanceRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

public class ExcelReader {

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
     * CONSOLIDER — original behaviour: col S non-empty = include row.
     */
    public List<CreanceRow> readFiltered(String companyName, File excelFile) throws IOException {
        List<CreanceRow> rows = new ArrayList<>();
        try (Workbook wb = openWorkbook(excelFile)) {
            Sheet sheet = wb.getSheet(AppConfig.CREANCES_SHEET_NAME);
            if (sheet == null) sheet = wb.getSheetAt(0);
            DataFormatter fmt = new DataFormatter();
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

            for (int r = 16; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                Cell filterCell = row.getCell(AppConfig.FILTER_COLUMN_INDEX);
                if (!hasRealData(filterCell)) continue;
                rows.add(new CreanceRow(companyName, extractRowValues(row, fmt, evaluator), r));
            }
        }
        if (rows.isEmpty()) System.out.println("[" + companyName + "] SKIPPED - no data in column S");
        return rows;
    }

    /**
     * TOUS LES DOSSIERS — no Lieu filter, optional date range on col C (Remis Le).
     */
    public List<CreanceRow> readFilteredTous(String companyName, File excelFile,
                                              LocalDate dateDebut, LocalDate dateFin) throws IOException {
        List<CreanceRow> rows = new ArrayList<>();
        try (Workbook wb = openWorkbook(excelFile)) {
            Sheet sheet = wb.getSheet(AppConfig.CREANCES_SHEET_NAME);
            if (sheet == null) sheet = wb.getSheetAt(0);
            DataFormatter fmt = new DataFormatter();
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

            for (int r = 16; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                // Row must have some real data (col S or fallback col J)
                Cell filterCell = row.getCell(AppConfig.FILTER_COLUMN_INDEX);
                if (!hasRealData(filterCell)) {
                    Cell fallback = row.getCell(9);
                    if (!hasRealData(fallback)) continue;
                }

                // Date range filter on col C (index 2 = Remis Le)
                if (dateDebut != null || dateFin != null) {
                    LocalDate remis = extractDate(row.getCell(2));
                    if (remis == null) continue;
                    if (dateDebut != null && remis.isBefore(dateDebut)) continue;
                    if (dateFin   != null && remis.isAfter(dateFin))   continue;
                }

                rows.add(new CreanceRow(companyName, extractRowValues(row, fmt, evaluator), r));
            }
        }
        if (rows.isEmpty()) System.out.println("[" + companyName + "] SKIPPED - no rows matched");
        return rows;
    }

    // -------------------------------------------------------------------------

    private Workbook openWorkbook(File file) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        return file.getName().toLowerCase().endsWith(".xls")
            ? new HSSFWorkbook(fis)
            : new XSSFWorkbook(fis);
    }

    private static final java.util.Set<String> HEADER_LABELS = java.util.Set.of("lieu", "location", "place");

    private boolean hasRealData(Cell cell) {
        if (cell == null) return false;
        CellType type = cell.getCellType() == CellType.FORMULA
            ? cell.getCachedFormulaResultType() : cell.getCellType();
        return switch (type) {
            case BLANK   -> false;
            case BOOLEAN -> true;
            case NUMERIC -> cell.getNumericCellValue() != 0.0;
            case STRING  -> {
                String val = cell.getStringCellValue().trim();
                yield !val.isEmpty() && !HEADER_LABELS.contains(val.toLowerCase());
            }
            default -> false;
        };
    }

    private LocalDate extractDate(Cell cell) {
        if (cell == null) return null;
        try {
            CellType type = cell.getCellType() == CellType.FORMULA
                ? cell.getCachedFormulaResultType() : cell.getCellType();
            if (type == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                return DateUtil.getJavaDate(cell.getNumericCellValue())
                    .toInstant().atZone(java.time.ZoneId.systemDefault()).toLocalDate();
            }
            if (type == CellType.STRING) {
                String s = cell.getStringCellValue().trim();
                for (String pat : new String[]{"dd/MM/yyyy", "d/MM/yyyy", "dd/M/yyyy"}) {
                    try { return LocalDate.parse(s, DateTimeFormatter.ofPattern(pat)); }
                    catch (Exception ignored) {}
                }
            }
        } catch (Exception ignored) {}
        return null;
    }

    private static final java.util.Set<Integer> NUMERIC_STRING_COLS = java.util.Set.of(
        8, 9, 12, 13, 14, 18, 20, 21, 22, 23, 24
    );

    private Object coerceStringToDouble(String raw) {
        String cleaned = raw.replaceAll("[€$£\\s]", "").replace(".", "").replace(",", ".");
        try { return Double.parseDouble(cleaned); }
        catch (NumberFormatException e) { return raw; }
    }

    private List<Object> extractRowValues(Row row, DataFormatter fmt, FormulaEvaluator evaluator) {
        int lastCell = row.getLastCellNum();
        List<Object> values = new ArrayList<>(lastCell);
        for (int c = 0; c < lastCell; c++) {
            Cell cell = row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            if (cell == null) { values.add(""); continue; }

            CellType type;
            double  numericResult = 0;
            String  stringResult  = "";
            boolean boolResult    = false;

            if (cell.getCellType() == CellType.FORMULA) {
                try {
                    org.apache.poi.ss.usermodel.CellValue cv = evaluator.evaluate(cell);
                    type = cv.getCellType();
                    switch (type) {
                        case NUMERIC -> numericResult = cv.getNumberValue();
                        case STRING  -> stringResult  = cv.getStringValue();
                        case BOOLEAN -> boolResult    = cv.getBooleanValue();
                        default -> {}
                    }
                } catch (Exception e) { values.add(""); continue; }
            } else {
                type = cell.getCellType();
                switch (type) {
                    case NUMERIC -> numericResult = cell.getNumericCellValue();
                    case STRING  -> stringResult  = cell.getStringCellValue();
                    case BOOLEAN -> boolResult    = cell.getBooleanCellValue();
                    default -> {}
                }
            }

            switch (type) {
                case NUMERIC -> {
                    if (DateUtil.isCellDateFormatted(cell)) {
                        values.add(DateUtil.getJavaDate(numericResult)
                            .toInstant().atZone(java.time.ZoneId.systemDefault()).toLocalDateTime());
                    } else {
                        values.add(numericResult);
                    }
                }
                case BOOLEAN -> values.add(boolResult);
                case STRING  -> values.add(
                    NUMERIC_STRING_COLS.contains(c) ? coerceStringToDouble(stringResult) : stringResult);
                case BLANK   -> values.add("");
                default      -> values.add("");
            }
        }
        return values;
    }
}
