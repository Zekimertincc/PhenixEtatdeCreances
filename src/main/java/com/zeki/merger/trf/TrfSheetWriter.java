package com.zeki.merger.trf;

import com.zeki.merger.trf.model.ClientSummary;
import com.zeki.merger.trf.model.ConsolidationRow;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;

/**
 * Writes the four-sheet TRF workbook:
 *   Consolidation — verbatim source data with header + per-client SUBTOTAL rows
 *   Feuil1        — one summary row per client (26-column structure)
 *   TRF           — main transfer document
 *   Feuil3        — empty (tab required by reference format)
 */
public class TrfSheetWriter {

    // -------------------------------------------------------------------------
    // Shared column definitions
    // -------------------------------------------------------------------------

    private static final String[] CONSO_HEADERS = {
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
        "Extratction du départements de la colonne E",
        "Transformation de la colonne L en nombre",
        "CONDITION de calcul de formule 2 :France ou 1 :Export",
        "DONT EN ATTENTE DE FACTURATION",
        "Lieu",
        "Frais de procédure",
        "Recouvré total",
        "Déjà facturé",
        "Depuis le début",
        "Commissions",
        "Pénalits",
        "SOMMES CZ PHENIX",
        "MONTANT A FACTURER TTC",
        "SOMMES A REVERSER"
    };
    private static final int CONSO_COLS = CONSO_HEADERS.length; // 26

    /** Columns that carry monetary values (0-based), used for #,##0.00 and SUBTOTAL. */
    private static final Set<Integer> MONEY_COLS =
        Set.of(7, 8, 11, 15, 17, 18, 19, 20, 21, 22, 23, 24, 25);

    private static boolean isMoneyCol(int c) { return MONEY_COLS.contains(c); }

    // =========================================================================
    // Public entry point
    // =========================================================================

    public void write(List<ConsolidationRow> allRows,
                      List<ClientSummary>    summaries,
                      File                   outputFile) throws IOException {

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Styles s = new Styles(wb);

            writeConsolidationSheet(wb, allRows, s);
            writeFeuil1Sheet(wb, summaries, s);
            writeTrfSheet(wb, summaries, s);
            wb.createSheet("Feuil3"); // required empty tab

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                wb.write(fos);
            }
        }
    }

    // =========================================================================
    // Sheet 1 — "Consolidation"
    // =========================================================================

    private void writeConsolidationSheet(XSSFWorkbook wb,
                                          List<ConsolidationRow> rows,
                                          Styles s) {
        XSSFSheet sheet = wb.createSheet("Consolidation");
        int rowIdx = 0;

        // Fixed 26-column header
        XSSFRow hdr = sheet.createRow(rowIdx++);
        for (int c = 0; c < CONSO_COLS; c++) {
            XSSFCell cell = hdr.createCell(c);
            cell.setCellValue(CONSO_HEADERS[c]);
            cell.setCellStyle(s.headerDark);
        }

        String currentClient = null;
        int    groupStartRow = -1;

        for (ConsolidationRow cr : rows) {
            if (cr.isHeaderRow()) continue; // skip source header; we wrote our own

            List<Object> vals = cr.getValues();
            String colA = vals.isEmpty() ? "" : strOf(vals.get(0));

            if (!colA.isEmpty()) {
                if (!colA.equals(currentClient)) {
                    // Flush previous client's subtotal
                    if (currentClient != null && rowIdx > groupStartRow) {
                        rowIdx = writeConsoSubtotal(sheet, rowIdx, currentClient, groupStartRow, s);
                    }
                    currentClient = colA;
                    groupStartRow = rowIdx; // first data row of this group
                }
            } else {
                // Blank col A — company header or spacer: flush current client first
                if (currentClient != null && rowIdx > groupStartRow) {
                    rowIdx = writeConsoSubtotal(sheet, rowIdx, currentClient, groupStartRow, s);
                    currentClient = null;
                    groupStartRow = -1;
                }
            }

            // Write source row
            XSSFRow row = sheet.createRow(rowIdx++);
            for (int c = 0; c < vals.size(); c++) {
                XSSFCell cell = row.createCell(c);
                writeValue(cell, vals.get(c),
                           isMoneyCol(c) ? s.moneyStyle : s.dataStyle, s.dateStyle);
            }
        }

        // Flush final client
        if (currentClient != null && rowIdx > groupStartRow) {
            writeConsoSubtotal(sheet, rowIdx, currentClient, groupStartRow, s);
        }

        autoSize(sheet, CONSO_COLS);
        sheet.createFreezePane(0, 1);
    }

    /**
     * Writes a "Total [clientName]" row at rowIdx using SUBTOTAL(9,...) for money columns.
     * Range covers [groupStartRow+1 .. rowIdx] in Excel (1-based).
     */
    private int writeConsoSubtotal(XSSFSheet sheet, int rowIdx,
                                    String clientName, int groupStartRow,
                                    Styles s) {
        XSSFRow row = sheet.createRow(rowIdx);
        row.createCell(0).setCellValue("Total " + clientName);
        row.getCell(0).setCellStyle(s.totalStyle);

        for (int c = 1; c < CONSO_COLS; c++) {
            XSSFCell cell = row.createCell(c);
            if (MONEY_COLS.contains(c)) {
                String letter = col(c);
                // Excel rows are 1-based: first data = groupStartRow+1, last data = rowIdx
                cell.setCellFormula("SUBTOTAL(9," + letter + (groupStartRow + 1)
                    + ":" + letter + rowIdx + ")");
                cell.setCellStyle(s.totalMoneyStyle);
            } else {
                cell.setCellStyle(s.totalStyle);
            }
        }
        return rowIdx + 1;
    }

    // =========================================================================
    // Sheet 2 — "Feuil1" (26-column summary, one row per client)
    // =========================================================================

    private void writeFeuil1Sheet(XSSFWorkbook wb,
                                   List<ClientSummary> summaries,
                                   Styles s) {
        XSSFSheet sheet = wb.createSheet("Feuil1");
        int rowIdx = 0;

        // Same 26-column header as Consolidation
        XSSFRow hdr = sheet.createRow(rowIdx++);
        for (int c = 0; c < CONSO_COLS; c++) {
            XSSFCell cell = hdr.createCell(c);
            cell.setCellValue(CONSO_HEADERS[c]);
            cell.setCellStyle(s.headerDark);
        }

        int dataStart = rowIdx;

        for (ClientSummary cs : summaries) {
            XSSFRow row = sheet.createRow(rowIdx++);

            // Col 0 (A): CLIENT name
            txt(row, 0, cs.getClientName(), s.textStyle);
            // Col 1 (B): client code (in NBRE position)
            txt(row, 1, cs.getClientCode(), s.textStyle);
            // Cols 2–6: blank
            for (int c = 2; c <= 6; c++) row.createCell(c).setCellStyle(s.dataStyle);

            // Money columns mapped to their Consolidation indices
            dbl(row, 7,  cs.getCreancePrincipale(),   s.moneyStyle);
            dbl(row, 8,  cs.getRecouvreEtFacture(),    s.moneyStyle);
            row.createCell(9).setCellStyle(s.dataStyle);   // ETAT (blank)
            row.createCell(10).setCellStyle(s.dataStyle);  // CLOTURE (blank)
            dbl(row, 11, cs.getPenalites(),            s.moneyStyle);
            row.createCell(12).setCellStyle(s.dataStyle);
            row.createCell(13).setCellStyle(s.dataStyle);
            row.createCell(14).setCellStyle(s.dataStyle);
            dbl(row, 15, cs.getDontEnAttente(),        s.moneyStyle);
            row.createCell(16).setCellStyle(s.dataStyle);  // Lieu (blank)
            dbl(row, 17, cs.getFraisProcedure(),       s.moneyStyle);
            dbl(row, 18, cs.getRecouvreTotol(),        s.moneyStyle);
            dbl(row, 19, cs.getDejaFacture(),          s.moneyStyle);
            dbl(row, 20, cs.getDepuisLeDebut(),        s.moneyStyle);
            dbl(row, 21, cs.getCommissions(),          s.moneyStyle);
            dbl(row, 22, cs.getPenalits(),             s.moneyStyle);
            dbl(row, 23, cs.getSommesCzPhenix(),       s.moneyStyle);
            dbl(row, 24, cs.getMontantAFacturerTtc(),  s.moneyStyle);
            dbl(row, 25, cs.getSommesAReverserSrc(),   s.moneyStyle);
        }

        int dataEnd = rowIdx - 1;

        // TOTAUX row — SUM for each money column
        XSSFRow totRow = sheet.createRow(rowIdx);
        txt(totRow, 0, "TOTAUX", s.totalStyle);
        txt(totRow, 1, "",       s.totalStyle);
        for (int c = 2; c < CONSO_COLS; c++) {
            XSSFCell cell = totRow.createCell(c);
            if (MONEY_COLS.contains(c)) {
                cell.setCellStyle(s.totalMoneyStyle);
                cell.setCellFormula("SUM(" + col(c) + (dataStart + 1)
                    + ":" + col(c) + (dataEnd + 1) + ")");
            } else {
                cell.setCellStyle(s.totalStyle);
            }
        }

        autoSize(sheet, CONSO_COLS);
        sheet.createFreezePane(0, 1);
    }

    // =========================================================================
    // Sheet 3 — "TRF"
    // =========================================================================

    private static final String[] TRF_HEADERS = {
        "",                              // A  0 — filled with dynamic label below
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
        LocalDate today = LocalDate.now();
        String mmyy = String.format("%02d", today.getMonthValue())
                      + "/" + String.valueOf(today.getYear()).substring(2);

        XSSFRow hdr = sheet.createRow(rowIdx++);
        XSSFCell colAHdr = hdr.createCell(0);
        colAHdr.setCellValue("CLIENTS EN FACTURATION " + mmyy);
        colAHdr.setCellStyle(s.headerDark);
        for (int c = 1; c < TRF_COLS; c++) {
            XSSFCell cell = hdr.createCell(c);
            cell.setCellValue(TRF_HEADERS[c]);
            cell.setCellStyle(s.headerDark);
        }

        int dataStart = rowIdx;

        // ---- One row per client -----------------------------------------
        for (ClientSummary cs : summaries) {
            int excelRow = rowIdx + 1; // 1-based
            XSSFRow row = sheet.createRow(rowIdx++);

            txt(row, 0, cs.getClientName(),         s.textStyle);
            dbl(row, 1, cs.getSommesCzPhenix(),     s.moneyStyle);
            dbl(row, 2, cs.getMontantAFacturerTtc(), s.moneyStyle);
            dbl(row, 3, cs.getNousDoit_Prec(),      s.moneyStyle);

            // E = C + D
            formula(row, 4, "C" + excelRow + "+D" + excelRow, s.moneyStyle);

            if (cs.isNonCompensation()) {
                // F = B (return all encaissements)
                formula(row, 5, "B" + excelRow, s.moneyStyle);
                // G = 0
                dbl(row, 6, 0.0, s.moneyStyle);
                // H = E - G
                formula(row, 7, "E" + excelRow + "-G" + excelRow, s.moneyStyle);
            } else {
                // F = IF(B=0,0,IF(B<E,0,B-E))
                formula(row, 5,
                    "IF(B" + excelRow + "=0,0,IF(B" + excelRow + "<E" + excelRow
                    + ",0,B" + excelRow + "-E" + excelRow + "))",
                    s.moneyStyle);
                // G = IF(B=0,0,IF(B>E,E,B))
                formula(row, 6,
                    "IF(B" + excelRow + "=0,0,IF(B" + excelRow + ">E" + excelRow
                    + ",E" + excelRow + ",B" + excelRow + "))",
                    s.moneyStyle);
                // H = E - G
                formula(row, 7, "E" + excelRow + "-G" + excelRow, s.moneyStyle);
            }

            // I: ETAT DE COMPENSATIONS
            txt(row, 8, cs.isNonCompensation() ? "NON COMP" : cs.getEtatCompensations(),
                s.textStyle);

            // J: VIREMENTS — "OUI" if client receives money back, else blank
            txt(row, 9, cs.getSommesAReverserFinal() > 0.005 ? "OUI" : "", s.textStyle);

            // K: CHEQUES
            dbl(row, 10, cs.getCheques(), s.moneyStyle);

            // L: CODE CLIENT
            txt(row, 11, cs.getClientCode(), s.textStyle);
        }

        int dataEnd = rowIdx - 1;

        // ---- TOTAUX row -------------------------------------------------
        XSSFRow totRow = sheet.createRow(rowIdx++);
        txt(totRow, 0, "TOTAUX", s.totalStyle);
        for (int c = 1; c < TRF_COLS; c++) {
            XSSFCell cell = totRow.createCell(c);
            boolean textCol = (c == 8 || c == 9 || c == 11);
            cell.setCellStyle(textCol ? s.totalStyle : s.totalMoneyStyle);
            if (!textCol) {
                cell.setCellFormula("SUM(" + col(c) + (dataStart + 1)
                    + ":" + col(c) + (dataEnd + 1) + ")");
            }
        }

        sheet.createRow(rowIdx++); // blank separator

        // ---- Bottom sections --------------------------------------------
        rowIdx = writeVirementsSection    (sheet, summaries, rowIdx, s);
        rowIdx = writeManuellesSection    (sheet, summaries, rowIdx, s);
        rowIdx = writeNonCompSection      (sheet, summaries, rowIdx, s);
        rowIdx = writeCompPartielleSection(sheet, summaries, rowIdx, s);
        rowIdx = writeDebiteursSection    (sheet, summaries, rowIdx, s);

        autoSize(sheet, TRF_COLS);
        sheet.createFreezePane(0, 1);
    }

    // =========================================================================
    // TRF bottom sections
    // =========================================================================

    /** VIREMENTS CLIENTS — all clients where sommesAReverserFinal > 0. */
    private int writeVirementsSection(XSSFSheet sheet, List<ClientSummary> summaries,
                                       int rowIdx, Styles s) {
        List<ClientSummary> list = summaries.stream()
            .filter(cs -> cs.getSommesAReverserFinal() > 0.005)
            .collect(Collectors.toList());

        rowIdx = writeSectionHeader(sheet, rowIdx, "VIREMENTS CLIENTS", s);
        rowIdx = writeSubHeader(sheet, rowIdx, new String[]{"CLIENT", "IBAN", "BIC", "MONTANT"}, s);

        int dataStart = rowIdx;
        for (ClientSummary cs : list) {
            XSSFRow row = sheet.createRow(rowIdx++);
            txt(row, 0, cs.getClientName(),           s.textStyle);
            txt(row, 1, cs.getIban(),                 s.textStyle);
            txt(row, 2, cs.getBic(),                  s.textStyle);
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

        sheet.createRow(rowIdx++);
        return rowIdx;
    }

    /** VIREMENTS MANUELLES — sommesAReverserFinal > 0 but no IBAN. */
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
            txt(row, 0, cs.getClientName(),           s.textStyle);
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

    /** NON COMP — clients where nonCompensation = true. */
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
            txt(row, 0, cs.getClientName(),               s.textStyle);
            dbl(row, 1, cs.getSommesCzPhenix(),           s.moneyStyle);
            dbl(row, 2, cs.getNousDoit_ApreFacturation(), s.moneyStyle);
        }
        XSSFRow tot = sheet.createRow(rowIdx++);
        txt(tot, 0, "TOTAL NON COMP", s.totalStyle);
        XSSFCell tc1 = tot.createCell(1); tc1.setCellStyle(s.totalMoneyStyle);
        tc1.setCellFormula("SUM(B" + (dataStart + 1) + ":B" + rowIdx + ")");
        XSSFCell tc2 = tot.createCell(2); tc2.setCellStyle(s.totalMoneyStyle);
        tc2.setCellFormula("SUM(C" + (dataStart + 1) + ":C" + rowIdx + ")");

        sheet.createRow(rowIdx++);
        return rowIdx;
    }

    /** COMP PARTIELLE — partial compensation applied but client still owes remainder. */
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
            txt(row, 0, cs.getClientName(),                   s.textStyle);
            dbl(row, 1, cs.getEncaissementsParCompensation(), s.moneyStyle);
            dbl(row, 2, cs.getNousDoit_ApreFacturation(),     s.moneyStyle);
        }
        XSSFRow tot = sheet.createRow(rowIdx++);
        txt(tot, 0, "TOTAL COMP PARTIELLE", s.totalStyle);
        XSSFCell tc1 = tot.createCell(1); tc1.setCellStyle(s.totalMoneyStyle);
        tc1.setCellFormula("SUM(B" + (dataStart + 1) + ":B" + rowIdx + ")");
        XSSFCell tc2 = tot.createCell(2); tc2.setCellStyle(s.totalMoneyStyle);
        tc2.setCellFormula("SUM(C" + (dataStart + 1) + ":C" + rowIdx + ")");

        sheet.createRow(rowIdx++);
        return rowIdx;
    }

    /** DEBITEURS — no encaissements this period but still owe Phénix. */
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
        XSSFCell tc = tot.createCell(1); tc.setCellStyle(s.totalMoneyStyle);
        tc.setCellFormula("SUM(B" + (dataStart + 1) + ":B" + rowIdx + ")");

        return rowIdx;
    }

    // =========================================================================
    // Section layout helpers
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

    private void formula(XSSFRow row, int col, String f, XSSFCellStyle style) {
        XSSFCell cell = row.createCell(col);
        cell.setCellFormula(f);
        cell.setCellStyle(style);
    }

    private void writeValue(XSSFCell cell, Object val,
                             XSSFCellStyle defStyle, XSSFCellStyle dateStyle) {
        if (val instanceof Double d)           { cell.setCellValue(d);              cell.setCellStyle(defStyle);  return; }
        if (val instanceof Number n)           { cell.setCellValue(n.doubleValue()); cell.setCellStyle(defStyle); return; }
        if (val instanceof Boolean b)          { cell.setCellValue(b);              cell.setCellStyle(defStyle);  return; }
        if (val instanceof LocalDateTime ldt)  { cell.setCellValue(ldt);            cell.setCellStyle(dateStyle); return; }
        if (val instanceof String str && !str.isBlank()) {
            String stripped = str.replaceAll("[€$£¥₺  \\s]", "");
            if (!stripped.isEmpty() && stripped.matches("[-+]?[\\d.,]+")) {
                cell.setCellValue(ConsolidationRow.parseFrenchDouble(str));
                cell.setCellStyle(defStyle);
                return;
            }
            cell.setCellValue(str);
            cell.setCellStyle(defStyle);
            return;
        }
        cell.setCellStyle(defStyle);
    }

    private static String strOf(Object v) {
        if (v == null) return "";
        String s = v.toString().trim();
        // Coerced doubles (e.g. "CLIENT" stored as 0.0) → blank
        try { Double.parseDouble(s); return ""; } catch (NumberFormatException ignored) {}
        return s;
    }

    // =========================================================================
    // Utilities
    // =========================================================================

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
    // Styles
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
                setBorder(headerDark);
            }

            dataStyle = wb.createCellStyle();
            dataStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            setBorder(dataStyle);

            dateStyle = wb.createCellStyle();
            dateStyle.cloneStyleFrom(dataStyle);
            dateStyle.setDataFormat(wb.getCreationHelper().createDataFormat()
                .getFormat("dd/MM/yyyy"));

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
                setBorder(totalStyle);
            }

            totalMoneyStyle = wb.createCellStyle();
            totalMoneyStyle.cloneStyleFrom(totalStyle);
            totalMoneyStyle.setDataFormat(moneyFmt);

            moneyStyle = wb.createCellStyle();
            moneyStyle.cloneStyleFrom(dataStyle);
            moneyStyle.setDataFormat(moneyFmt);

            textStyle = wb.createCellStyle();
            textStyle.cloneStyleFrom(dataStyle);

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
                setBorder(subHeader);
            }
        }

        private void setBorder(XSSFCellStyle s) {
            XSSFColor bc = new XSSFColor(new byte[]{(byte)0xD9,(byte)0xD9,(byte)0xD9}, null);
            s.setBorderTop   (BorderStyle.THIN); s.setTopBorderColor   (bc);
            s.setBorderBottom(BorderStyle.THIN); s.setBottomBorderColor(bc);
            s.setBorderLeft  (BorderStyle.THIN); s.setLeftBorderColor  (bc);
            s.setBorderRight (BorderStyle.THIN); s.setRightBorderColor (bc);
        }
    }
}
