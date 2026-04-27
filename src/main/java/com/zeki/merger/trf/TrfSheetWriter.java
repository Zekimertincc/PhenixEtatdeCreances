package com.zeki.merger.trf;

import com.zeki.merger.trf.model.ClientSummary;
import com.zeki.merger.trf.model.ConsolidationRow;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;

/**
 * Writes the three-sheet TRF workbook:
 * <ol>
 *   <li>"Consolidation" — verbatim copy of the source sheet</li>
 *   <li>"Feuil1"        — one summary row per client</li>
 *   <li>"TRF"           — main transfer document with virements sections</li>
 * </ol>
 */
public class TrfSheetWriter {

    // -------------------------------------------------------------------------
    // Public entry point
    // -------------------------------------------------------------------------

    public void write(List<ConsolidationRow> allRows,
                      List<ClientSummary>    summaries,
                      File                   outputFile) throws IOException {

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Styles s = new Styles(wb);

            writeConsolidationSheet(wb, allRows, s);
            writeFeuil1Sheet(wb, summaries, s);
            writeTrfSheet(wb, summaries, s);

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                wb.write(fos);
            }
        }
    }

    // Columns that carry monetary amounts (0-based) — used to apply #,##0.00 format
    private static final Set<Integer> MONEY_COLS =
        Set.of(7, 8, 11, 15, 17, 18, 19, 20, 21, 22, 23, 24, 25);

    private static boolean isMoneyCol(int c) { return MONEY_COLS.contains(c); }

    // =========================================================================
    // Sheet 1 — "Consolidation" (verbatim copy)
    // =========================================================================

    private void writeConsolidationSheet(XSSFWorkbook wb,
                                          List<ConsolidationRow> rows,
                                          Styles s) {
        XSSFSheet sheet = wb.createSheet("Consolidation");

        int rowIdx = 0;
        for (ConsolidationRow cr : rows) {
            XSSFRow row = sheet.createRow(rowIdx++);
            List<Object> vals = cr.getValues();

            for (int c = 0; c < vals.size(); c++) {
                XSSFCell cell = row.createCell(c);
                XSSFCellStyle cellStyle;
                if (cr.isHeaderRow()) {
                    cellStyle = s.headerDark;
                } else if (cr.isTotalRow()) {
                    cellStyle = isMoneyCol(c) ? s.totalMoneyStyle : s.totalStyle;
                } else {
                    cellStyle = isMoneyCol(c) ? s.moneyStyle : s.dataStyle;
                }
                writeValue(cell, vals.get(c), cellStyle, s.dateStyle);
            }
        }

        autoSize(sheet, 26);
        sheet.createFreezePane(0, 1);
    }

    // =========================================================================
    // Sheet 2 — "Feuil1" (summary per client)
    // =========================================================================

    private static final String[] FEUIL1_HEADERS = {
        "CLIENT", "NBRE", "CREANCE PRINCIPALE", "RECOUVRE ET FACTURE",
        "PENALITES", "DONT EN ATTENTE", "Frais procédure",
        "Recouvré total", "Déjà facturé", "Depuis le début",
        "Commissions", "Pénalits", "SOMMES CZ PHENIX",
        "MONTANT A FACTURER TTC", "SOMMES A REVERSER"
    };
    private static final int F1_COLS = FEUIL1_HEADERS.length;

    private void writeFeuil1Sheet(XSSFWorkbook wb,
                                   List<ClientSummary> summaries,
                                   Styles s) {
        XSSFSheet sheet = wb.createSheet("Feuil1");
        int rowIdx = 0;

        // Header
        XSSFRow hdr = sheet.createRow(rowIdx++);
        for (int c = 0; c < F1_COLS; c++) {
            XSSFCell cell = hdr.createCell(c);
            cell.setCellValue(FEUIL1_HEADERS[c]);
            cell.setCellStyle(s.headerDark);
        }

        int dataStart = rowIdx;

        // Data rows
        for (ClientSummary cs : summaries) {
            XSSFRow row = sheet.createRow(rowIdx++);
            int c = 0;
            num(row, c++, cs.getClientName(),            s.textStyle);
            num(row, c++, cs.getClientCode(),            s.textStyle);
            num(row, c++, cs.getCreancePrincipale(),     s.moneyStyle);
            num(row, c++, cs.getRecouvreEtFacture(),     s.moneyStyle);
            num(row, c++, cs.getPenalites(),             s.moneyStyle);
            num(row, c++, cs.getDontEnAttente(),         s.moneyStyle);
            num(row, c++, cs.getFraisProcedure(),        s.moneyStyle);
            num(row, c++, cs.getRecouvreTotol(),         s.moneyStyle);
            num(row, c++, cs.getDejaFacture(),           s.moneyStyle);
            num(row, c++, cs.getDepuisLeDebut(),         s.moneyStyle);
            num(row, c++, cs.getCommissions(),           s.moneyStyle);
            num(row, c++, cs.getPenalits(),              s.moneyStyle);
            num(row, c++, cs.getSommesCzPhenix(),        s.moneyStyle);
            num(row, c++, cs.getMontantAFacturerTtc(),   s.moneyStyle);
            num(row, c  , cs.getSommesAReverserSrc(),    s.moneyStyle);
        }

        int dataEnd = rowIdx - 1;

        // Totaux row
        XSSFRow totRow = sheet.createRow(rowIdx);
        txt(totRow, 0, "TOTAUX", s.totalStyle);
        txt(totRow, 1, "",       s.totalStyle);
        for (int c = 2; c < F1_COLS; c++) {
            XSSFCell cell = totRow.createCell(c);
            cell.setCellStyle(s.totalMoneyStyle);
            cell.setCellFormula("SUM(" + col(c) + (dataStart + 1)
                + ":" + col(c) + (dataEnd + 1) + ")");
        }

        autoSize(sheet, F1_COLS);
        sheet.createFreezePane(0, 1);
    }

    // =========================================================================
    // Sheet 3 — "TRF" (main transfer document)
    // =========================================================================

    private static final String[] TRF_HEADERS = {
        "CLIENT",                        // A  0
        "ENCAISSEMENTS CZ PHENIX",       // B  1
        "MONTANT A FACTURER TTC",        // C  2
        "NOUS DOIT précédemment",        // D  3
        "NOUS DOIT MAINTENANT",          // E  4
        "SOMMES A REVERSER AU FINAL",    // F  5
        "ENCAISSEMENTS PAR COMPENSATION",// G  6
        "NOUS DOIT APRES FACTURATION",   // H  7
        "ETAT DE COMPENSATIONS",         // I  8
        "VIREMENTS",                     // J  9
        "CHEQUES",                       // K  10
        "CODE CLIENT"                    // L  11
    };
    private static final int TRF_COLS = TRF_HEADERS.length;

    private void writeTrfSheet(XSSFWorkbook wb,
                                List<ClientSummary> summaries,
                                Styles s) {
        XSSFSheet sheet = wb.createSheet("TRF");
        int rowIdx = 0;

        // ---- Header row -------------------------------------------------
        XSSFRow hdr = sheet.createRow(rowIdx++);
        for (int c = 0; c < TRF_COLS; c++) {
            XSSFCell cell = hdr.createCell(c);
            cell.setCellValue(TRF_HEADERS[c]);
            cell.setCellStyle(s.headerDark);
        }

        int dataStart = rowIdx; // first client row (Excel row = rowIdx+1)

        // ---- One row per client -----------------------------------------
        for (ClientSummary cs : summaries) {
            int excelRow = rowIdx + 1; // 1-based Excel row
            XSSFRow row = sheet.createRow(rowIdx++);

            // A: CLIENT
            txt(row, 0, cs.getClientName(), s.textStyle);

            // B: ENCAISSEMENTS CZ PHENIX (plain number)
            dbl(row, 1, cs.getSommesCzPhenix(), s.moneyStyle);

            // C: MONTANT A FACTURER TTC (plain number)
            dbl(row, 2, cs.getMontantAFacturerTtc(), s.moneyStyle);

            // D: NOUS DOIT précédemment (plain number)
            dbl(row, 3, cs.getNousDoit_Prec(), s.moneyStyle);

            // E: NOUS DOIT MAINTENANT = C + D
            formula(row, 4, "C" + excelRow + "+D" + excelRow, s.moneyStyle);

            if (cs.isNonCompensation()) {
                // F: SOMMES A REVERSER = B (all encaissements returned)
                formula(row, 5, "B" + excelRow, s.moneyStyle);
                // G: ENCAISSEMENTS PAR COMPENSATION = 0
                dbl(row, 6, 0.0, s.moneyStyle);
                // H: NOUS DOIT APRES FACTURATION = E (full invoice still owed)
                formula(row, 7, "E" + excelRow, s.moneyStyle);
            } else {
                // F: SOMMES A REVERSER = MAX(0, B - E)
                formula(row, 5, "MAX(0,B" + excelRow + "-E" + excelRow + ")", s.moneyStyle);
                // G: ENCAISSEMENTS PAR COMPENSATION = MIN(B, MAX(0,E))
                formula(row, 6, "MIN(B" + excelRow + ",MAX(0,E" + excelRow + "))", s.moneyStyle);
                // H: NOUS DOIT APRES FACTURATION = MAX(0, E - G)
                formula(row, 7, "MAX(0,E" + excelRow + "-G" + excelRow + ")", s.moneyStyle);
            }

            // I: ETAT DE COMPENSATIONS (computed text)
            txt(row, 8, cs.getEtatCompensations(), s.textStyle);

            // J: VIREMENTS = F (client gets money back)
            formula(row, 9, "F" + excelRow, s.moneyStyle);

            // K: CHEQUES (filled manually, 0 initial)
            dbl(row, 10, cs.getCheques(), s.moneyStyle);

            // L: CODE CLIENT
            txt(row, 11, cs.getClientCode(), s.textStyle);
        }

        int dataEnd = rowIdx - 1; // last client row index (0-based)

        // ---- TOTAUX row -------------------------------------------------
        int totExcelRow = rowIdx + 1;
        XSSFRow totRow = sheet.createRow(rowIdx++);
        txt(totRow, 0, "TOTAUX", s.totalStyle);
        for (int c = 1; c < TRF_COLS; c++) {
            XSSFCell cell = totRow.createCell(c);
            cell.setCellStyle(c == 8 || c == 11 ? s.totalStyle : s.totalMoneyStyle);
            if (c != 8 && c != 11) { // skip text columns
                cell.setCellFormula("SUM(" + col(c) + (dataStart + 1)
                    + ":" + col(c) + (dataEnd + 1) + ")");
            }
        }

        // Blank separator
        sheet.createRow(rowIdx++);

        // ---- Virements sections below -----------------------------------
        rowIdx = writeVirementsSection    (sheet, summaries, rowIdx, s);
        rowIdx = writeManuellesSection    (sheet, summaries, rowIdx, s);
        rowIdx = writeNonCompSection      (sheet, summaries, rowIdx, s);
        rowIdx = writeCompPartielleSection(sheet, summaries, rowIdx, s);
        rowIdx = writeDebiteursSection    (sheet, summaries, rowIdx, s);

        autoSize(sheet, TRF_COLS);
        sheet.createFreezePane(0, 1);
    }

    // -------------------------------------------------------------------------
    // TRF bottom sections
    // -------------------------------------------------------------------------

    /** VIREMENTS CLIENTS — non-NonComp, SOMMES A REVERSER > 0, IBAN present */
    private int writeVirementsSection(XSSFSheet sheet, List<ClientSummary> summaries,
                                       int rowIdx, Styles s) {
        List<ClientSummary> list = summaries.stream()
            .filter(ClientSummary::needsAutoVirement)
            .collect(Collectors.toList());

        rowIdx = writeSectionHeader(sheet, rowIdx, "VIREMENTS CLIENTS", s);

        // Sub-header
        String[] sh = {"CLIENT", "IBAN", "BIC", "MONTANT"};
        rowIdx = writeSubHeader(sheet, rowIdx, sh, s);

        int dataStart = rowIdx;
        for (ClientSummary cs : list) {
            XSSFRow row = sheet.createRow(rowIdx++);
            txt(row, 0, cs.getClientName(),       s.textStyle);
            txt(row, 1, cs.getIban(),             s.textStyle);
            txt(row, 2, cs.getBic(),              s.textStyle);
            dbl(row, 3, cs.getSommesAReverserFinal(), s.moneyStyle);
        }

        if (!list.isEmpty()) {
            XSSFRow tot = sheet.createRow(rowIdx++);
            txt(tot, 0, "TOTAL VIREMENTS", s.totalStyle);
            txt(tot, 1, "", s.totalStyle);
            txt(tot, 2, "", s.totalStyle);
            XSSFCell tc = tot.createCell(3);
            tc.setCellStyle(s.totalMoneyStyle);
            tc.setCellFormula("SUM(D" + (dataStart + 1) + ":D" + rowIdx + ")");
        }

        sheet.createRow(rowIdx++); // blank
        return rowIdx;
    }

    /** VIREMENTS MANUELLES — non-NonComp, SOMMES A REVERSER > 0, IBAN absent */
    private int writeManuellesSection(XSSFSheet sheet, List<ClientSummary> summaries,
                                       int rowIdx, Styles s) {
        List<ClientSummary> list = summaries.stream()
            .filter(ClientSummary::needsManualVirement)
            .collect(Collectors.toList());

        if (list.isEmpty()) return rowIdx;

        rowIdx = writeSectionHeader(sheet, rowIdx, "VIREMENTS MANUELLES", s);
        rowIdx = writeSubHeader(sheet, rowIdx, new String[]{"CLIENT", "MONTANT"}, s);

        int dataStart = rowIdx;
        for (ClientSummary cs : list) {
            XSSFRow row = sheet.createRow(rowIdx++);
            txt(row, 0, cs.getClientName(),       s.textStyle);
            dbl(row, 1, cs.getSommesAReverserFinal(), s.moneyStyle);
        }

        XSSFRow tot = sheet.createRow(rowIdx++);
        txt(tot, 0, "TOTAL MANUELLES", s.totalStyle);
        XSSFCell tc = tot.createCell(1);
        tc.setCellStyle(s.totalMoneyStyle);
        tc.setCellFormula("SUM(B" + (dataStart + 1) + ":B" + rowIdx + ")");

        sheet.createRow(rowIdx++);
        return rowIdx;
    }

    /** NON COMP — clients where NonComp = "OUI" */
    private int writeNonCompSection(XSSFSheet sheet, List<ClientSummary> summaries,
                                     int rowIdx, Styles s) {
        List<ClientSummary> list = summaries.stream()
            .filter(ClientSummary::isNonCompensation)
            .collect(Collectors.toList());

        if (list.isEmpty()) return rowIdx;

        rowIdx = writeSectionHeader(sheet, rowIdx, "NON COMP", s);
        rowIdx = writeSubHeader(sheet, rowIdx,
            new String[]{"CLIENT", "ENCAISSEMENTS CZ PHENIX", "NOUS DOIT APRES FACTURATION"}, s);

        int dataStart = rowIdx;
        for (ClientSummary cs : list) {
            XSSFRow row = sheet.createRow(rowIdx++);
            txt(row, 0, cs.getClientName(),             s.textStyle);
            dbl(row, 1, cs.getSommesCzPhenix(),         s.moneyStyle);
            dbl(row, 2, cs.getNousDoit_ApreFacturation(), s.moneyStyle);
        }

        XSSFRow tot = sheet.createRow(rowIdx++);
        txt(tot, 0, "TOTAL NON COMP", s.totalStyle);
        XSSFCell tc1 = tot.createCell(1);
        tc1.setCellStyle(s.totalMoneyStyle);
        tc1.setCellFormula("SUM(B" + (dataStart + 1) + ":B" + rowIdx + ")");
        XSSFCell tc2 = tot.createCell(2);
        tc2.setCellStyle(s.totalMoneyStyle);
        tc2.setCellFormula("SUM(C" + (dataStart + 1) + ":C" + rowIdx + ")");

        sheet.createRow(rowIdx++);
        return rowIdx;
    }

    /** COMP PARTIELLE — partial compensation applied, client still owes remainder */
    private int writeCompPartielleSection(XSSFSheet sheet, List<ClientSummary> summaries,
                                           int rowIdx, Styles s) {
        List<ClientSummary> list = summaries.stream()
            .filter(ClientSummary::isPartiallyCompensated)
            .collect(Collectors.toList());

        if (list.isEmpty()) return rowIdx;

        rowIdx = writeSectionHeader(sheet, rowIdx, "COMP PARTIELLE", s);
        rowIdx = writeSubHeader(sheet, rowIdx,
            new String[]{"CLIENT", "COMP APPLIQUÉE", "RESTE NOUS DEVOIR"}, s);

        int dataStart = rowIdx;
        for (ClientSummary cs : list) {
            XSSFRow row = sheet.createRow(rowIdx++);
            txt(row, 0, cs.getClientName(),                  s.textStyle);
            dbl(row, 1, cs.getEncaissementsParCompensation(), s.moneyStyle);
            dbl(row, 2, cs.getNousDoit_ApreFacturation(),    s.moneyStyle);
        }

        XSSFRow tot = sheet.createRow(rowIdx++);
        txt(tot, 0, "TOTAL COMP PARTIELLE", s.totalStyle);
        XSSFCell tc1 = tot.createCell(1);
        tc1.setCellStyle(s.totalMoneyStyle);
        tc1.setCellFormula("SUM(B" + (dataStart + 1) + ":B" + rowIdx + ")");
        XSSFCell tc2 = tot.createCell(2);
        tc2.setCellStyle(s.totalMoneyStyle);
        tc2.setCellFormula("SUM(C" + (dataStart + 1) + ":C" + rowIdx + ")");

        sheet.createRow(rowIdx++);
        return rowIdx;
    }

    /** DEBITEURS — no encaissements this period but still owe Phénix */
    private int writeDebiteursSection(XSSFSheet sheet, List<ClientSummary> summaries,
                                       int rowIdx, Styles s) {
        List<ClientSummary> list = summaries.stream()
            .filter(ClientSummary::isDebtor)
            .collect(Collectors.toList());

        if (list.isEmpty()) return rowIdx;

        rowIdx = writeSectionHeader(sheet, rowIdx, "DEBITEURS", s);
        rowIdx = writeSubHeader(sheet, rowIdx,
            new String[]{"CLIENT", "NOUS DOIT APRES FACTURATION"}, s);

        int dataStart = rowIdx;
        for (ClientSummary cs : list) {
            XSSFRow row = sheet.createRow(rowIdx++);
            txt(row, 0, cs.getClientName(),               s.textStyle);
            dbl(row, 1, cs.getNousDoit_ApreFacturation(), s.moneyStyle);
        }

        XSSFRow tot = sheet.createRow(rowIdx++);
        txt(tot, 0, "TOTAL DEBITEURS", s.totalStyle);
        XSSFCell tc = tot.createCell(1);
        tc.setCellStyle(s.totalMoneyStyle);
        tc.setCellFormula("SUM(B" + (dataStart + 1) + ":B" + rowIdx + ")");

        return rowIdx;
    }

    // =========================================================================
    // Layout helpers for bottom sections
    // =========================================================================

    private int writeSectionHeader(XSSFSheet sheet, int rowIdx, String title, Styles s) {
        XSSFRow row = sheet.createRow(rowIdx++);
        XSSFCell cell = row.createCell(0);
        cell.setCellValue(title);
        cell.setCellStyle(s.sectionHeader);
        return rowIdx;
    }

    private int writeSubHeader(XSSFSheet sheet, int rowIdx, String[] headers, Styles s) {
        XSSFRow row = sheet.createRow(rowIdx++);
        for (int c = 0; c < headers.length; c++) {
            XSSFCell cell = row.createCell(c);
            cell.setCellValue(headers[c]);
            cell.setCellStyle(s.subHeader);
        }
        return rowIdx;
    }

    // =========================================================================
    // Cell writing helpers
    // =========================================================================

    private void txt(XSSFRow row, int col, String val, XSSFCellStyle style) {
        XSSFCell cell = row.createCell(col);
        cell.setCellValue(val != null ? val : "");
        cell.setCellStyle(style);
    }

    private void dbl(XSSFRow row, int col, double val, XSSFCellStyle style) {
        XSSFCell cell = row.createCell(col);
        cell.setCellValue(val);
        cell.setCellStyle(style);
    }

    /** Dual-dispatch helper that routes string or double to the right cell type. */
    private void num(XSSFRow row, int col, Object val, XSSFCellStyle style) {
        XSSFCell cell = row.createCell(col);
        if (val instanceof Double d) {
            cell.setCellValue(d);
        } else if (val instanceof Number n) {
            cell.setCellValue(n.doubleValue());
        } else if (val instanceof String s && !s.isBlank()) {
            double d = ConsolidationRow.parseFrenchDouble(s);
            if (d != 0.0) { cell.setCellValue(d); } else { cell.setCellValue(s); }
        } else {
            cell.setCellValue(val != null ? val.toString() : "");
        }
        cell.setCellStyle(style);
    }

    private void formula(XSSFRow row, int col, String f, XSSFCellStyle style) {
        XSSFCell cell = row.createCell(col);
        cell.setCellFormula(f);
        cell.setCellStyle(style);
    }

    private void writeValue(XSSFCell cell, Object val,
                             XSSFCellStyle defStyle, XSSFCellStyle dateStyle) {
        if (val instanceof Double d)            { cell.setCellValue(d);             cell.setCellStyle(defStyle);   return; }
        if (val instanceof Number n)            { cell.setCellValue(n.doubleValue()); cell.setCellStyle(defStyle); return; }
        if (val instanceof Boolean b)           { cell.setCellValue(b);             cell.setCellStyle(defStyle);   return; }
        if (val instanceof LocalDateTime ldt)   { cell.setCellValue(ldt);           cell.setCellStyle(dateStyle);  return; }
        if (val instanceof String s && !s.isBlank()) {
            double d = ConsolidationRow.parseFrenchDouble(s);
            if (d != 0.0) { cell.setCellValue(d); cell.setCellStyle(defStyle); return; }
            cell.setCellValue(s);
            cell.setCellStyle(defStyle);
            return;
        }
        cell.setCellStyle(defStyle);
    }

    // =========================================================================
    // Utilities
    // =========================================================================

    /** Converts 0-based column index to Excel column letter(s). */
    private static String col(int idx) {
        if (idx < 26) return String.valueOf((char) ('A' + idx));
        return String.valueOf((char) ('A' + idx / 26 - 1))
             + (char) ('A' + idx % 26);
    }

    private void autoSize(XSSFSheet sheet, int numCols) {
        for (int c = 0; c < numCols; c++) {
            sheet.autoSizeColumn(c);
            int w = sheet.getColumnWidth(c);
            sheet.setColumnWidth(c, Math.min(w + 512, 20_000));
        }
    }

    // =========================================================================
    // Styles inner class
    // =========================================================================

    private static class Styles {
        final XSSFCellStyle headerDark;
        final XSSFCellStyle dataStyle;
        final XSSFCellStyle dateStyle;
        final XSSFCellStyle totalStyle;
        final XSSFCellStyle totalMoneyStyle;
        final XSSFCellStyle moneyStyle;
        final XSSFCellStyle textStyle;
        final XSSFCellStyle sectionHeader;
        final XSSFCellStyle subHeader;

        Styles(XSSFWorkbook wb) {
            DataFormat df = wb.createDataFormat();
            short moneyFmt = df.getFormat("#,##0.00");

            // --- dark blue header (sheet-level headers) ---
            headerDark = wb.createCellStyle();
            {
                XSSFFont f = wb.createFont();
                f.setBold(true);
                f.setColor(new XSSFColor(new byte[]{(byte)0xFF,(byte)0xFF,(byte)0xFF}, null));
                f.setFontHeightInPoints((short) 10);
                headerDark.setFont(f);
                headerDark.setFillForegroundColor(
                    new XSSFColor(new byte[]{(byte)0x1F,(byte)0x4E,(byte)0x79}, null));
                headerDark.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                headerDark.setVerticalAlignment(VerticalAlignment.CENTER);
                headerDark.setWrapText(true);
                setBorder(headerDark, wb);
            }

            // --- default data text ---
            dataStyle = wb.createCellStyle();
            dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            setBorder(dataStyle, wb);

            // --- date data ---
            dateStyle = wb.createCellStyle();
            dateStyle.cloneStyleFrom(dataStyle);
            dateStyle.setDataFormat(wb.getCreationHelper().createDataFormat()
                .getFormat("dd/MM/yyyy"));

            // --- total row (bold, yellow) ---
            totalStyle = wb.createCellStyle();
            {
                XSSFFont f = wb.createFont();
                f.setBold(true);
                f.setFontHeightInPoints((short) 10);
                totalStyle.setFont(f);
                totalStyle.setFillForegroundColor(
                    new XSSFColor(new byte[]{(byte)0xFF,(byte)0xF2,(byte)0xCC}, null));
                totalStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                totalStyle.setVerticalAlignment(VerticalAlignment.CENTER);
                setBorder(totalStyle, wb);
            }

            // --- total row with money format ---
            totalMoneyStyle = wb.createCellStyle();
            totalMoneyStyle.cloneStyleFrom(totalStyle);
            totalMoneyStyle.setDataFormat(moneyFmt);

            // --- money number format ---
            moneyStyle = wb.createCellStyle();
            moneyStyle.cloneStyleFrom(dataStyle);
            moneyStyle.setDataFormat(moneyFmt);

            // --- plain text data ---
            textStyle = wb.createCellStyle();
            textStyle.cloneStyleFrom(dataStyle);

            // --- section header (light blue) ---
            sectionHeader = wb.createCellStyle();
            {
                XSSFFont f = wb.createFont();
                f.setBold(true);
                f.setFontHeightInPoints((short) 11);
                sectionHeader.setFont(f);
                sectionHeader.setFillForegroundColor(
                    new XSSFColor(new byte[]{(byte)0x9D,(byte)0xC3,(byte)0xE6}, null));
                sectionHeader.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                sectionHeader.setVerticalAlignment(VerticalAlignment.CENTER);
            }

            // --- sub-header (light grey) ---
            subHeader = wb.createCellStyle();
            {
                XSSFFont f = wb.createFont();
                f.setBold(true);
                f.setFontHeightInPoints((short) 9);
                subHeader.setFont(f);
                subHeader.setFillForegroundColor(
                    new XSSFColor(new byte[]{(byte)0xD9,(byte)0xD9,(byte)0xD9}, null));
                subHeader.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                subHeader.setVerticalAlignment(VerticalAlignment.CENTER);
                setBorder(subHeader, wb);
            }
        }

        private static void setBorder(XSSFCellStyle s, XSSFWorkbook wb) {
            XSSFColor bc = new XSSFColor(new byte[]{(byte)0xD9,(byte)0xD9,(byte)0xD9}, null);
            s.setBorderTop   (BorderStyle.THIN); s.setTopBorderColor   (bc);
            s.setBorderBottom(BorderStyle.THIN); s.setBottomBorderColor(bc);
            s.setBorderLeft  (BorderStyle.THIN); s.setLeftBorderColor  (bc);
            s.setBorderRight (BorderStyle.THIN); s.setRightBorderColor (bc);
        }
    }
}
