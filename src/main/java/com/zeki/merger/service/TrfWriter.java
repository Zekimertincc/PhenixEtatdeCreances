package com.zeki.merger.service;

import com.zeki.merger.model.CreanceRow;
import com.zeki.merger.trf.model.ConsolidationRow;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.*;

/**
 * Writes grouped company rows to a single XLSX file in TRF format.
 *
 * Sheet structure:
 *   Row 0   — TRF header (26 columns)
 *   Per company:
 *     [Company header row]  — company name in column B, light-blue background
 *     [Empty spacer row]
 *     [Data rows]           — CLIENT in col 0, NBRE blank, source cols (skip 5) in cols 2–25
 *     [Total row]           — "Total <company>" in col A, SUM formula for each numeric col
 *
 * Column mapping (0-based):
 *   TRF col 0 = CLIENT  (company name, injected by app)
 *   TRF col 1 = NBRE    (left blank)
 *   TRF cols 2–25       = source cols 0..N skipping source index 5, in order
 */
public class TrfWriter {

    private static final int TOTAL_COLS = 26;

    private static final String[] HEADERS = {
        "CLIENT",
        "NBRE",
        "V/REF",
        "REMIS LE",
        "ANCIENNETE",
        "N/REF",
        "DEBITEUR",
        "CREANCE PRINCIPALE",
        "RECOUVRE ET FACTURE",
        "ETAT",
        "CLOTURE",
        "PENALITES",
        "Extraction du départements de la colonne E",
        "Transformation de la colonne L en nombre",
        "CONDITION de calcul de formule 2 :France ou 1 :Export",
        "DONT EN ATTENTE DE FACTURATION",
        "Lieu",
        "Frais de procédure",
        "Recouvré total",
        "Déjà facturé",
        "dépuis le début",
        "Commissions",
        "Pénalits",
        "SOMMES CZ PENIX",
        "MONTANT A FACTURER TTC",
        "SOMMES A REVERSER"
    };

    public void write(Map<String, List<CreanceRow>> groupedRows, File outputFile)
            throws IOException {

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Consolidation");

            XSSFCellStyle headerStyle        = buildHeaderStyle(wb);
            XSSFCellStyle companyHeaderStyle = buildCompanyHeaderStyle(wb);
            XSSFCellStyle dataStyle          = buildDataStyle(wb);
            XSSFCellStyle dateStyle          = buildDateStyle(wb, dataStyle);
            XSSFCellStyle totalStyle         = buildTotalStyle(wb);

            // ---- row 0: TRF header ----
            XSSFRow headerRow = sheet.createRow(0);
            for (int c = 0; c < HEADERS.length; c++) {
                XSSFCell cell = headerRow.createCell(c);
                cell.setCellValue(HEADERS[c]);
                cell.setCellStyle(headerStyle);
            }

            int rowIdx = 1;

            for (Map.Entry<String, List<CreanceRow>> entry : groupedRows.entrySet()) {
                String           company = entry.getKey();
                List<CreanceRow> rows    = entry.getValue();

                if (rows == null || rows.isEmpty()) continue;

                // ---- company header row ----
                XSSFRow compRow = sheet.createRow(rowIdx++);
                for (int c = 0; c < TOTAL_COLS; c++) {
                    XSSFCell cell = compRow.createCell(c);
                    cell.setCellStyle(companyHeaderStyle);
                }
                compRow.getCell(1).setCellValue(company); // column B

                // ---- spacer ----
                sheet.createRow(rowIdx++);

                // ---- data rows ----
                int dataStartRow = rowIdx; // 0-based, used for SUM range
                Set<Integer> numericTrfCols = new HashSet<>();

                for (CreanceRow cr : rows) {
                    XSSFRow row = sheet.createRow(rowIdx++);

                    // TRF col 0: CLIENT
                    XSSFCell clientCell = row.createCell(0);
                    clientCell.setCellValue(company);
                    clientCell.setCellStyle(dataStyle);

                    // TRF col 1: NBRE — left blank
                    XSSFCell nbreCell = row.createCell(1);
                    nbreCell.setCellStyle(dataStyle);

                    // TRF cols 2–25: source cols, skipping source index 5
                    List<Object> values = cr.getCellValues();
                    int trfCol = 2;
                    for (int s = 0; s < values.size() && trfCol < TOTAL_COLS; s++) {
                        if (s == 5) continue;
                        Object val = values.get(s);
                        XSSFCell cell = row.createCell(trfCol);
                        writeValue(cell, val, dataStyle, dateStyle);
                        if (isNumericValue(val)) numericTrfCols.add(trfCol);
                        trfCol++;
                    }
                    // fill remaining cols with border style
                    for (; trfCol < TOTAL_COLS; trfCol++) {
                        row.createCell(trfCol).setCellStyle(dataStyle);
                    }
                }

                int dataEndRow = rowIdx - 1; // 0-based, inclusive

                // ---- total row ----
                XSSFRow totalRow = sheet.createRow(rowIdx++);

                XSSFCell totalLabelCell = totalRow.createCell(0);
                totalLabelCell.setCellValue("Total " + company);
                totalLabelCell.setCellStyle(totalStyle);

                totalRow.createCell(1).setCellStyle(totalStyle); // NBRE — blank

                for (int c = 2; c < TOTAL_COLS; c++) {
                    XSSFCell cell = totalRow.createCell(c);
                    cell.setCellStyle(totalStyle);
                    if (numericTrfCols.contains(c)) {
                        // Excel rows are 1-based; POI rowIdx is 0-based
                        String colLetter = colLetter(c);
                        String formula = "SUM(" + colLetter + (dataStartRow + 1)
                            + ":" + colLetter + (dataEndRow + 1) + ")";
                        cell.setCellFormula(formula);
                    }
                }
            }

            // ---- auto-size all columns ----
            for (int c = 0; c < TOTAL_COLS; c++) {
                sheet.autoSizeColumn(c);
                int w = sheet.getColumnWidth(c);
                sheet.setColumnWidth(c, Math.min(w + 512, 20_000));
            }

            sheet.createFreezePane(0, 1);

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                wb.write(fos);
            }
        }
    }

    // -------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------

    /** Converts a 0-based column index (0–25) to its Excel letter (A–Z). */
    private static String colLetter(int colIndex) {
        if (colIndex < 26) return String.valueOf((char) ('A' + colIndex));
        // Two-letter fallback (not needed for ≤25 but kept for safety)
        return String.valueOf((char) ('A' + colIndex / 26 - 1))
             + (char) ('A' + colIndex % 26);
    }

    private void writeValue(XSSFCell cell, Object val,
                            XSSFCellStyle defaultStyle, XSSFCellStyle dateStyle) {
        if (val instanceof Double d)            { cell.setCellValue(d);              cell.setCellStyle(defaultStyle); return; }
        if (val instanceof Number n)            { cell.setCellValue(n.doubleValue()); cell.setCellStyle(defaultStyle); return; }
        if (val instanceof Boolean b)           { cell.setCellValue(b);              cell.setCellStyle(defaultStyle); return; }
        if (val instanceof LocalDateTime ldt)   { cell.setCellValue(ldt);            cell.setCellStyle(dateStyle);    return; }
        if (val instanceof String s && !s.isBlank()) {
            String stripped = s.replaceAll("[€$£¥₺  \\s]", "");
            if (!stripped.isEmpty() && stripped.matches("[-+]?[\\d.,]+")) {
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

    private boolean isNumericValue(Object val) {
        if (val instanceof Number) return true;
        if (val instanceof String s && !s.isBlank()) {
            String stripped = s.replaceAll("[€$£¥₺  \\s]", "");
            return !stripped.isEmpty() && stripped.matches("[-+]?[\\d.,]+");
        }
        return false;
    }

    // -------------------------------------------------------------------------
    // Style builders
    // -------------------------------------------------------------------------

    /** Bold white text on dark-blue (#1F4E79) background — TRF global header. */
    private XSSFCellStyle buildHeaderStyle(XSSFWorkbook wb) {
        XSSFCellStyle s = wb.createCellStyle();
        XSSFFont f = wb.createFont();
        f.setBold(true);
        f.setColor(new XSSFColor(new byte[]{(byte) 0xFF, (byte) 0xFF, (byte) 0xFF}, null));
        f.setFontHeightInPoints((short) 10);
        s.setFont(f);
        s.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 0x1F, (byte) 0x4E, (byte) 0x79}, null));
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        s.setWrapText(true);
        return s;
    }

    /** Bold dark text, light-blue (#BDD7EE) background — company group header. */
    private XSSFCellStyle buildCompanyHeaderStyle(XSSFWorkbook wb) {
        XSSFCellStyle s = wb.createCellStyle();
        XSSFFont f = wb.createFont();
        f.setBold(true);
        f.setFontHeightInPoints((short) 11);
        s.setFont(f);
        s.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 0xBD, (byte) 0xD7, (byte) 0xEE}, null));
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        return s;
    }

    /** No background, thin grey borders on all four sides. */
    private XSSFCellStyle buildDataStyle(XSSFWorkbook wb) {
        XSSFCellStyle s = wb.createCellStyle();
        XSSFColor borderColor = new XSSFColor(new byte[]{(byte) 0xD9, (byte) 0xD9, (byte) 0xD9}, null);
        s.setBorderTop(BorderStyle.THIN);
        s.setBorderBottom(BorderStyle.THIN);
        s.setBorderLeft(BorderStyle.THIN);
        s.setBorderRight(BorderStyle.THIN);
        s.setTopBorderColor(borderColor);
        s.setBottomBorderColor(borderColor);
        s.setLeftBorderColor(borderColor);
        s.setRightBorderColor(borderColor);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        return s;
    }

    /** Data style + dd/MM/yyyy date format. */
    private XSSFCellStyle buildDateStyle(XSSFWorkbook wb, XSSFCellStyle base) {
        XSSFCellStyle s = wb.createCellStyle();
        s.cloneStyleFrom(base);
        s.setDataFormat(wb.getCreationHelper().createDataFormat().getFormat("dd/MM/yyyy"));
        return s;
    }

    /** Bold, light-yellow (#FFF2CC) background — totals row. */
    private XSSFCellStyle buildTotalStyle(XSSFWorkbook wb) {
        XSSFCellStyle s = wb.createCellStyle();
        XSSFFont f = wb.createFont();
        f.setBold(true);
        f.setFontHeightInPoints((short) 10);
        s.setFont(f);
        s.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 0xFF, (byte) 0xF2, (byte) 0xCC}, null));
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        XSSFColor borderColor = new XSSFColor(new byte[]{(byte) 0xD9, (byte) 0xD9, (byte) 0xD9}, null);
        s.setBorderTop(BorderStyle.THIN);
        s.setBorderBottom(BorderStyle.THIN);
        s.setBorderLeft(BorderStyle.THIN);
        s.setBorderRight(BorderStyle.THIN);
        s.setTopBorderColor(borderColor);
        s.setBottomBorderColor(borderColor);
        s.setLeftBorderColor(borderColor);
        s.setRightBorderColor(borderColor);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        return s;
    }
}
