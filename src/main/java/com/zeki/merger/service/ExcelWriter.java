package com.zeki.merger.service;

import com.zeki.merger.model.CreanceRow;
import com.zeki.merger.trf.model.ConsolidationRow;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * Writes grouped company rows to a single XLSX file in the 26-column
 * ConsolidationGenerale format (sheet "Consolidation", columns A–Z).
 *
 * Layout per company:
 *   Row 0      — 26-column header (written once at the very top)
 *   [Company header]  A=blank, B=company name, light-blue background
 *   [Spacer row]
 *   [Data rows]       A=company name, B–V=mapped source data, W–Z=Excel formulas
 */
public class ExcelWriter {

    private static final int TOTAL_COLS = 26; // A–Z

    private static final String[] CONSO_HEADERS = {
        "CLIENT", "No Client", "V/REF", "REMIS LE", "ANCIENNETE", "N/REF",
        "DEBITEUR", "CREANCE PRINCIPALE ", "RECOUVRE ET FACTURE", "ETAT", "CLOTURE",
        "PENALITES", "Extratction du départements de la colonne E",
        "Transformation de la colonne L en nombre",
        "CONDITION de calcul de formule 2 :France ou 1 :Export",
        "DONT EN ATTENTE DE FACTURATION", "Lieu", "Frais de procédure",
        "Recouvré total", "Déjà facturé", "dépuis le début", "Commissions",
        "Commisions TTC", "SOMMES CZ PENIX", "MONTANT A FACTURER TTC", "SOMMES A REVERSER "
    };

    // Money columns in the target layout (0-based)
    private static final Set<Integer> MONEY_COLS = Set.of(
        7, 8, 11, 15, 17, 18, 19, 20, 21, 22, 23, 24
    );

    public void write(Map<String, List<CreanceRow>> groupedRows, File outputFile)
            throws IOException {

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Consolidation");

            XSSFCellStyle headerStyle  = buildHeaderStyle(wb);
            XSSFCellStyle compStyle    = buildCompanyHeaderStyle(wb);
            XSSFCellStyle dataStyle    = buildDataStyle(wb);
            XSSFCellStyle moneyStyle   = buildMoneyStyle(wb, dataStyle);
            XSSFCellStyle dateStyle    = buildDateStyle(wb, dataStyle);

            int rowIdx = 0;

            // Row 0: global 26-column header
            XSSFRow hdr = sheet.createRow(rowIdx++);
            for (int c = 0; c < TOTAL_COLS; c++) {
                XSSFCell cell = hdr.createCell(c);
                cell.setCellValue(CONSO_HEADERS[c]);
                cell.setCellStyle(headerStyle);
            }

            for (Map.Entry<String, List<CreanceRow>> entry : groupedRows.entrySet()) {
                String           company = entry.getKey();
                List<CreanceRow> rows    = entry.getValue();

                if (rows == null || rows.isEmpty()) {
                    System.out.println("[" + company + "] SKIPPED - no data in column S");
                    continue;
                }

                // Company header row: col A blank, col B = company name
                XSSFRow compRow = sheet.createRow(rowIdx++);
                for (int c = 0; c < TOTAL_COLS; c++) {
                    compRow.createCell(c).setCellStyle(compStyle);
                }
                compRow.getCell(1).setCellValue(company);

                // Spacer row
                sheet.createRow(rowIdx++);

                // Data rows
                for (CreanceRow cr : rows) {
                    int excelRow = rowIdx + 1; // 1-based Excel row number
                    XSSFRow row = sheet.createRow(rowIdx++);
                    List<Object> src = cr.getCellValues();

                    // A = company name
                    writeCell(row, 0, company, dataStyle, dateStyle, moneyStyle);
                    // B = src[0]  NBRE
                    writeCell(row, 1, get(src, 0), dataStyle, dateStyle, moneyStyle);
                    // C = src[1]  V/REF
                    writeCell(row, 2, get(src, 1), dataStyle, dateStyle, moneyStyle);
                    // D = src[2]  REMIS LE (date)
                    writeCell(row, 3, get(src, 2), dataStyle, dateStyle, moneyStyle);
                    // E = src[3]  ANCIENNETE
                    writeCell(row, 4, get(src, 3), dataStyle, dateStyle, moneyStyle);
                    // F = src[5]  N/REF  (src[4] = NOMBRE DE FACTURES → skipped)
                    writeCell(row, 5, get(src, 5), dataStyle, dateStyle, moneyStyle);
                    // G = src[6]  DEBITEUR
                    writeCell(row, 6, get(src, 6), dataStyle, dateStyle, moneyStyle);
                    // H = src[7]  CREANCE PRINCIPALE (money)
                    writeCell(row, 7, get(src, 7), dataStyle, dateStyle, moneyStyle);
                    // I = src[8]  RECOUVRE ET FACTURE (money)
                    writeCell(row, 8, get(src, 8), dataStyle, dateStyle, moneyStyle);
                    // J = src[9]  ETAT
                    writeCell(row, 9, get(src, 9), dataStyle, dateStyle, moneyStyle);
                    // K = src[10] CLOTURE
                    writeCell(row, 10, get(src, 10), dataStyle, dateStyle, moneyStyle);
                    // L = src[13] TOTAUX PENALITES  (src[11]=PENALITES, src[12]=INDEMNITES → skipped)
                    writeCell(row, 11, get(src, 13), dataStyle, dateStyle, moneyStyle);
                    // M = src[14] Extraction département
                    writeCell(row, 12, get(src, 14), dataStyle, dateStyle, moneyStyle);
                    // N = src[15] Transformation colonne L
                    writeCell(row, 13, get(src, 15), dataStyle, dateStyle, moneyStyle);
                    // O = src[16] CONDITION
                    writeCell(row, 14, get(src, 16), dataStyle, dateStyle, moneyStyle);
                    // P = src[17] DONT EN ATTENTE DE FACTURATION (money)
                    writeCell(row, 15, get(src, 17), dataStyle, dateStyle, moneyStyle);
                    // Q = src[18] Lieu
                    writeCell(row, 16, get(src, 18), dataStyle, dateStyle, moneyStyle);
                    // R = src[19] Frais de procédure (money)
                    writeCell(row, 17, get(src, 19), dataStyle, dateStyle, moneyStyle);
                    // S = src[20] Recouvré total (money)
                    writeCell(row, 18, get(src, 20), dataStyle, dateStyle, moneyStyle);
                    // T = src[21] Déjà facturé (money)
                    writeCell(row, 19, get(src, 21), dataStyle, dateStyle, moneyStyle);
                    // U = src[22] depuis le début (money)
                    writeCell(row, 20, get(src, 22), dataStyle, dateStyle, moneyStyle);
                    // V = src[23] Commissions HT (money)
                    writeCell(row, 21, get(src, 23), dataStyle, dateStyle, moneyStyle);

                    // W = Commisions TTC = V * 1.2
                    setFormula(row, 22, "ROUND(V" + excelRow + "*1.2,2)", moneyStyle);
                    // X = SOMMES CZ PHENIX
                    setFormula(row, 23,
                        "IF(Q" + excelRow + "=\"AG\",ROUND(P" + excelRow + "*1.2,2)"
                        + ",IF(Q" + excelRow + "=\"CL\",0,IF(Q" + excelRow + "=\"NA\",0,\"\")))",
                        moneyStyle);
                    // Y = MONTANT A FACTURER TTC
                    setFormula(row, 24,
                        "IF(ISNUMBER(V" + excelRow + "),ROUND((V" + excelRow
                        + "+R" + excelRow + ")*1.2,2),\"\")",
                        moneyStyle);
                    // Z = SOMMES A REVERSER
                    setFormula(row, 25,
                        "IF(ISBLANK(V" + excelRow + "),\"\",IF(AND(ISNUMBER(V" + excelRow
                        + "),X" + excelRow + ">=Y" + excelRow + "),X" + excelRow
                        + "-Y" + excelRow + ",\"RAS\"))",
                        dataStyle);
                }
            }

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

    private Object get(List<Object> list, int index) {
        if (index < 0 || index >= list.size()) return "";
        Object v = list.get(index);
        return v != null ? v : "";
    }

    private void writeCell(XSSFRow row, int col, Object val,
                            XSSFCellStyle dataStyle, XSSFCellStyle dateStyle,
                            XSSFCellStyle moneyStyle) {
        XSSFCell cell = row.createCell(col);
        if (val instanceof LocalDateTime ldt) {
            cell.setCellValue(ldt);
            cell.setCellStyle(dateStyle);
            return;
        }
        XSSFCellStyle style = MONEY_COLS.contains(col) ? moneyStyle : dataStyle;
        writeValue(cell, val, style, dateStyle);
    }

    private void setFormula(XSSFRow row, int col, String formula, XSSFCellStyle style) {
        XSSFCell cell = row.createCell(col);
        cell.setCellFormula(formula);
        cell.setCellStyle(style);
    }

    private void writeValue(XSSFCell cell, Object val,
                             XSSFCellStyle defaultStyle, XSSFCellStyle dateStyle) {
        if (val instanceof Double d) {
            double rounded = Math.round(d * 100.0) / 100.0;
            cell.setCellValue(rounded);
            cell.setCellStyle(defaultStyle);
            return;
        }
        if (val instanceof Number n) {
            double rounded = Math.round(n.doubleValue() * 100.0) / 100.0;
            cell.setCellValue(rounded);
            cell.setCellStyle(defaultStyle);
            return;
        }
        if (val instanceof Boolean b)         { cell.setCellValue(b);               cell.setCellStyle(defaultStyle); return; }
        if (val instanceof LocalDateTime ldt) { cell.setCellValue(ldt);             cell.setCellStyle(dateStyle);    return; }
        if (val instanceof String str && !str.isBlank()) {
            String stripped = str.replaceAll("[€$£¥₺]", "").replaceAll("\\p{Z}", "").trim();
            if (!stripped.isEmpty() && !stripped.equals("-")
                    && stripped.matches("[-+]?[\\d.,]+")) {
                double parsed = ConsolidationRow.parseFrenchDouble(str);
                double rounded = Math.round(parsed * 100.0) / 100.0;
                cell.setCellValue(rounded);
                cell.setCellStyle(defaultStyle);
                return;
            }
            cell.setCellValue(str);
            cell.setCellStyle(defaultStyle);
            return;
        }
        cell.setCellStyle(defaultStyle);
    }

    // -------------------------------------------------------------------------
    // Style builders
    // -------------------------------------------------------------------------

    private XSSFCellStyle buildHeaderStyle(XSSFWorkbook wb) {
        XSSFCellStyle s = wb.createCellStyle();
        XSSFFont f = wb.createFont();
        f.setBold(true);
        f.setColor(new XSSFColor(new byte[]{(byte)0xFF,(byte)0xFF,(byte)0xFF}, null));
        f.setFontHeightInPoints((short) 10);
        s.setFont(f);
        s.setFillForegroundColor(
            new XSSFColor(new byte[]{(byte)0x1F,(byte)0x4E,(byte)0x79}, null));
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        s.setWrapText(true);
        return s;
    }

    private XSSFCellStyle buildCompanyHeaderStyle(XSSFWorkbook wb) {
        XSSFCellStyle s = wb.createCellStyle();
        XSSFFont f = wb.createFont();
        f.setBold(true);
        f.setFontHeightInPoints((short) 11);
        s.setFont(f);
        s.setFillForegroundColor(
            new XSSFColor(new byte[]{(byte)0xBD,(byte)0xD7,(byte)0xEE}, null));
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        return s;
    }

    private XSSFCellStyle buildDataStyle(XSSFWorkbook wb) {
        XSSFCellStyle s = wb.createCellStyle();
        XSSFColor bc = new XSSFColor(new byte[]{(byte)0xD9,(byte)0xD9,(byte)0xD9}, null);
        s.setBorderTop(BorderStyle.THIN);    s.setTopBorderColor(bc);
        s.setBorderBottom(BorderStyle.THIN); s.setBottomBorderColor(bc);
        s.setBorderLeft(BorderStyle.THIN);   s.setLeftBorderColor(bc);
        s.setBorderRight(BorderStyle.THIN);  s.setRightBorderColor(bc);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        return s;
    }

    private XSSFCellStyle buildMoneyStyle(XSSFWorkbook wb, XSSFCellStyle base) {
        XSSFCellStyle s = wb.createCellStyle();
        s.cloneStyleFrom(base);
        s.setDataFormat(wb.createDataFormat().getFormat("#,##0.00"));
        return s;
    }

    private XSSFCellStyle buildDateStyle(XSSFWorkbook wb, XSSFCellStyle base) {
        XSSFCellStyle s = wb.createCellStyle();
        s.cloneStyleFrom(base);
        s.setDataFormat(wb.getCreationHelper().createDataFormat().getFormat("dd/MM/yyyy"));
        return s;
    }
}
