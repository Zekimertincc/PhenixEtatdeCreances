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
 * Writes the four-sheet TRF workbook matching TRF_04_2026 reference format.
 *
 * Consolidation: source data with deleted cols (F=CLOTURE, M+N=Extraction/Transformation)
 *                + formula cols S/T/U/V + SUBTOTAL rows per client
 * Feuil1:        one summary row per client + TOTAUX
 * TRF:           main transfer document, exact column labels from reference,
 *                bottom sections side-by-side (left = VIREMENTS, right = CHEQUES/NON COMP/etc.)
 * Feuil3:        empty tab
 */
public class TrfSheetWriter {

    // =========================================================================
    // Consolidation sheet — output column layout (22 cols)
    // Source cols SKIPPED: 5=CLOTURE, 12=Extraction dept, 13=Transformation L
    // Formula cols added at end: S=18, T=19, U=20, V=21
    // =========================================================================
    private static final String[] CONSO_HEADERS = {
            "CLIENT",                                           // A  0
            "No Client",                                        // B  1
            "V/REF",                                            // C  2
            "REMIS LE",                                         // D  3
            "ANCIENNETE",                                       // E  4
            "N/REF",                                            // F  5
            "DEBITEUR",                                         // G  6
            "CREANCE PRINCIPALE ",                              // H  7
            "RECOUVRE ET FACTURE",                              // I  8
            "ETAT",                                             // J  9
            "PENALITES",                                        // K  10
            "DONT EN ATTENTE DE FACTURATION",                   // L  11
            "Lieu",                                             // M  12
            "Frais de proc\u00e9dure",                          // N  13
            "Recouv\u00e9 total",                               // O  14
            "D\u00e9j\u00e0 factur\u00e9",                      // P  15
            "d\u00e9puis le d\u00e9but",                        // Q  16
            "Commissions",                                      // R  17
            "Commisions TTC",                                   // S  18  formula =R*1.2
            "SOMMES CZ PENIX",                                  // T  19  formula =IF(M="AG",L,...)
            "MONTANT A FACTURER TTC",                           // U  20  formula =(T+N)*1.2
            "SOMMES A REVERSER "                                // V  21  formula
    };
    private static final int CONSO_COLS = CONSO_HEADERS.length; // 22

    // Source column indices to SKIP when writing Consolidation
    private static final Set<Integer> SKIP_SRC = Set.of(5, 12, 13);
    // Last source column that contains raw data (V=21 in source = Commissions HT)
    private static final int SRC_LAST = 21;

    // Output column indices for formula columns
    private static final int O_LIEU    = 12; // M — Lieu ("AG"/"CL"/"NA")
    private static final int O_DONT    = 11; // L — DONT EN ATTENTE
    private static final int O_FRAIS   = 13; // N — Frais de proc.
    private static final int O_COMMHT  = 17; // R — Commissions HT
    private static final int O_COMMTTC = 18; // S — formula
    private static final int O_CZPHEN  = 19; // T — formula
    private static final int O_MONTANT = 20; // U — formula
    private static final int O_REVERS  = 21; // V — formula

    // Output columns that get SUBTOTAL in total rows
    // H(7), I(8), K(10), L(11), N(13), O(14), P(15), Q(16), S(18), U(20), V(21)
    private static final Set<Integer> SUBTOTAL_COLS = Set.of(7, 8, 10, 11, 13, 14, 15, 16, 18, 20, 21);

    // =========================================================================
    // Feuil1 headers (same structure, col B = client code)
    // =========================================================================
    private static final String[] FEUIL1_HEADERS = CONSO_HEADERS; // reuse same 22 cols
    private static final Set<Integer> FEUIL1_MONEY = Set.of(2,7,8,10,11,13,14,15,16,17,18,19,20,21);

    // =========================================================================
    // TRF sheet headers — exact labels from TRF_04_2026 reference
    // =========================================================================
    private static final String[] TRF_HEADERS = {
            "",                                      // A  0  dynamic "CLIENTS EN FACTURATION MM/YY"
            "ENCAISSEMENTS CZ PHENIX",               // B  1
            "MONTANT A FACTURER TTC",                // C  2
            "Mais NOUS DOIT pr\u00ecedamment ",      // D  3
            "NOUS DOIT MAINTENANT (C+D)",            // E  4
            "SOMMES A REVERSER AU FINAL (B-F)",      // F  5
            "ENCAISSEMENTS PAR COMPENSATION",        // G  6
            "NOUS DOIT APRES FACTURATION",           // H  7
            "ETAT DE COMPENSATIONS",                 // I  8
            "VIREMENTS",                             // J  9
            "CHEQUES",                               // K  10
            "CODE CLIENT"                            // L  11
    };
    private static final int TRF_COLS = TRF_HEADERS.length;

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
            wb.createSheet("Feuil3");
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                wb.write(fos);
            }
        }
    }

    // =========================================================================
    // Sheet 1 — Consolidation
    // =========================================================================
    private void writeConsolidationSheet(XSSFWorkbook wb,
                                         List<ConsolidationRow> rows,
                                         Styles s) {
        XSSFSheet sheet = wb.createSheet("Consolidation");
        int rowIdx = 0;

        // Header
        XSSFRow hdr = sheet.createRow(rowIdx++);
        for (int c = 0; c < CONSO_COLS; c++) {
            XSSFCell cell = hdr.createCell(c);
            cell.setCellValue(CONSO_HEADERS[c]);
            cell.setCellStyle(s.headerDark);
        }

        String currentClient = null;
        int    groupStartRow = -1;

        for (ConsolidationRow cr : rows) {
            if (cr.isHeaderRow()) continue;

            List<Object> src = cr.getValues();
            String colA = src.isEmpty() ? "" : strOf(src.get(0));

            if (!colA.isEmpty()) {
                if (!colA.equals(currentClient)) {
                    if (currentClient != null && rowIdx > groupStartRow)
                        rowIdx = writeConsoSubtotal(sheet, rowIdx, currentClient, groupStartRow, s);
                    currentClient = colA;
                    groupStartRow = rowIdx;
                }
            } else {
                if (currentClient != null && rowIdx > groupStartRow) {
                    rowIdx = writeConsoSubtotal(sheet, rowIdx, currentClient, groupStartRow, s);
                    currentClient = null;
                    groupStartRow = -1;
                }
            }

            XSSFRow row = sheet.createRow(rowIdx++);
            int outCol = 0;
            for (int srcCol = 0; srcCol <= SRC_LAST && srcCol < src.size(); srcCol++) {
                if (SKIP_SRC.contains(srcCol)) continue;
                XSSFCell cell = row.createCell(outCol);
                if (outCol == 6) { // DEBITEUR always string
                    Object v = src.get(srcCol);
                    cell.setCellValue(v != null ? v.toString() : "");
                    cell.setCellStyle(s.dataStyle);
                } else {
                    writeValue(cell, src.get(srcCol),
                            SUBTOTAL_COLS.contains(outCol) ? s.moneyStyle : s.dataStyle,
                            s.dateStyle);
                }
                outCol++;
            }
            // Pad to O_COMMHT if source was short
            while (outCol <= O_COMMHT) {
                row.createCell(outCol++).setCellStyle(s.dataStyle);
            }

            // Formula columns — only for real data rows (not blank spacer rows)
            if (!colA.isEmpty()) {
                int exR = rowIdx; // 1-based Excel row (rowIdx already incremented)
                writeFormulaCols(row, exR, s);
            }
        }

        if (currentClient != null && rowIdx > groupStartRow)
            writeConsoSubtotal(sheet, rowIdx, currentClient, groupStartRow, s);

        autoSize(sheet, CONSO_COLS);
        sheet.createFreezePane(0, 1);
    }

    /** Writes formula cells S, T, U, V for a data row at 1-based excelRow. */
    private void writeFormulaCols(XSSFRow row, int exR, Styles s) {
        String M = col(O_LIEU);
        String L = col(O_DONT);
        String N = col(O_FRAIS);
        String R = col(O_COMMHT);
        String S = col(O_COMMTTC);
        String T = col(O_CZPHEN);
        String U = col(O_MONTANT);
        String V = col(O_REVERS);

        // S = Commisions TTC = R * 1.2
        formula(row, O_COMMTTC, R + exR + "*1.2", s.moneyStyle);

        // T = SOMMES CZ PHENIX = IF(M="AG",L,IF(M="CL",0,IF(M="NA",0,"")))
        formula(row, O_CZPHEN,
                "IF(" + M + exR + "=\"AG\"," + L + exR +
                        ",IF(" + M + exR + "=\"CL\",0,IF(" + M + exR + "=\"NA\",0,\"\")))",
                s.moneyStyle);

        // U = MONTANT A FACTURER TTC = IF(ISNUMBER(T),(T+N)*1.2,"")
        formula(row, O_MONTANT,
                "IF(ISNUMBER(" + T + exR + "),(" + T + exR + "+" + N + exR + ")*1.2,\"\")",
                s.moneyStyle);

        // V = SOMMES A REVERSER = IF(ISBLANK(T),"",IF(AND(ISNUMBER(T),T>=U),T-U,"RAS"))
        formula(row, O_REVERS,
                "IF(ISBLANK(" + T + exR + "),\"\","
                        + "IF(AND(ISNUMBER(" + T + exR + ")," + T + exR + ">=" + U + exR + "),"
                        + T + exR + "-" + U + exR + ",\"RAS\"))",
                s.moneyStyle);
    }

    private int writeConsoSubtotal(XSSFSheet sheet, int rowIdx,
                                   String clientName, int groupStartRow, Styles s) {
        XSSFRow row = sheet.createRow(rowIdx);
        XSSFCell nameCell = row.createCell(0);
        nameCell.setCellValue("Total " + clientName);
        nameCell.setCellStyle(s.totalStyle);

        for (int c = 1; c < CONSO_COLS; c++) {
            XSSFCell cell = row.createCell(c);
            if (SUBTOTAL_COLS.contains(c)) {
                String letter = col(c);
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
    // Sheet 2 — Feuil1
    // =========================================================================
    private void writeFeuil1Sheet(XSSFWorkbook wb, List<ClientSummary> summaries, Styles s) {
        XSSFSheet sheet = wb.createSheet("Feuil1");
        int rowIdx = 0;

        XSSFRow hdr = sheet.createRow(rowIdx++);
        for (int c = 0; c < CONSO_COLS; c++) {
            XSSFCell cell = hdr.createCell(c);
            cell.setCellValue(FEUIL1_HEADERS[c]);
            cell.setCellStyle(s.headerDark);
        }

        int dataStart = rowIdx;
        for (ClientSummary cs : summaries) {
            XSSFRow row = sheet.createRow(rowIdx++);
            txt(row,  0, cs.getClientName(),          s.textStyle);
            txt(row,  1, cs.getClientCode(),           s.textStyle);
            dbl(row,  2, cs.getCommissionTtc(),        s.moneyStyle);
            for (int c = 3; c <= 6; c++) row.createCell(c).setCellStyle(s.dataStyle);
            dbl(row,  7, cs.getCreancePrincipale(),    s.moneyStyle);
            dbl(row,  8, cs.getRecouvreEtFacture(),    s.moneyStyle);
            row.createCell(9).setCellStyle(s.dataStyle);
            dbl(row, 10, cs.getPenalites(),             s.moneyStyle);
            dbl(row, 11, cs.getDontEnAttente(),         s.moneyStyle);
            row.createCell(12).setCellStyle(s.dataStyle);
            dbl(row, 13, cs.getFraisProcedure(),        s.moneyStyle);
            dbl(row, 14, cs.getRecouvreTotol(),         s.moneyStyle);
            dbl(row, 15, cs.getDejaFacture(),           s.moneyStyle);
            dbl(row, 16, cs.getDepuisLeDebut(),         s.moneyStyle);
            dbl(row, 17, cs.getCommissions(),           s.moneyStyle);
            dbl(row, 18, cs.getCommissionTtc(),         s.moneyStyle);
            dbl(row, 19, cs.getSommesCzPhenix(),        s.moneyStyle);
            dbl(row, 20, cs.getMontantAFacturerTtc(),   s.moneyStyle);
            dbl(row, 21, cs.getSommesAReverserSrc(),    s.moneyStyle);
        }

        int dataEnd = rowIdx - 1;
        XSSFRow totRow = sheet.createRow(rowIdx);
        txt(totRow, 0, "TOTAUX", s.totalStyle);
        txt(totRow, 1, "",       s.totalStyle);
        for (int c = 2; c < CONSO_COLS; c++) {
            XSSFCell cell = totRow.createCell(c);
            if (FEUIL1_MONEY.contains(c)) {
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
    // Sheet 3 — TRF
    // =========================================================================
    private void writeTrfSheet(XSSFWorkbook wb, List<ClientSummary> summaries, Styles s) {
        XSSFSheet sheet = wb.createSheet("TRF");
        int rowIdx = 0;

        LocalDate today = LocalDate.now();
        String mmyy = String.format("%02d", today.getMonthValue())
                + "/" + String.valueOf(today.getYear()).substring(2);

        // Header row
        XSSFRow hdr = sheet.createRow(rowIdx++);
        txt(hdr, 0, "CLIENTS EN FACTURATION " + mmyy, s.headerDark);
        for (int c = 1; c < TRF_COLS; c++) {
            XSSFCell cell = hdr.createCell(c);
            cell.setCellValue(TRF_HEADERS[c]);
            cell.setCellStyle(s.headerDark);
        }

        double[] totals = new double[TRF_COLS];

        // Data rows
        for (ClientSummary cs : summaries) {
            XSSFRow row = sheet.createRow(rowIdx++);
            txt(row,  0, cs.getClientName(),                   s.textStyle);
            dbl(row,  1, cs.getSommesCzPhenix(),               s.moneyStyle);
            dbl(row,  2, cs.getMontantAFacturerTtc(),          s.moneyStyle);
            dbl(row,  3, cs.getNousDoit_Prec(),                s.moneyStyle);
            dbl(row,  4, cs.getNousDoit_Maintenant(),          s.moneyStyle);
            dbl(row,  5, cs.getSommesAReverserFinal(),         s.moneyStyle);
            dbl(row,  6, cs.getEncaissementsParCompensation(), s.moneyStyle);
            dbl(row,  7, cs.getNousDoit_ApreFacturation(),     s.moneyStyle);

            totals[1] += cs.getSommesCzPhenix();
            totals[2] += cs.getMontantAFacturerTtc();
            totals[3] += cs.getNousDoit_Prec();
            totals[4] += cs.getNousDoit_Maintenant();
            totals[5] += cs.getSommesAReverserFinal();
            totals[6] += cs.getEncaissementsParCompensation();
            totals[7] += cs.getNousDoit_ApreFacturation();

            String etat = cs.isNonCompensation() ? "NON COMP" : cs.getEtatCompensations();
            txt(row, 8, etat, s.textStyle);

            if (cs.isNonCompensation()) {
                txt(row, 9,  cs.needsManualViremementNonComp() ? "Manuelle" : "", s.textStyle);
                txt(row, 10, cs.isNonCompWithoutIban()         ? "OUI"      : "", s.textStyle);
            } else {
                boolean isVrt = etat.startsWith("Comp VRT");
                boolean isCb  = etat.startsWith("Comp CB");
                txt(row, 9,  (isVrt && cs.getSommesAReverserFinal() > 0.005) ? "OUI" : "", s.textStyle);
                txt(row, 10, (isCb  && cs.getSommesAReverserFinal() > 0.005) ? "OUI" : "", s.textStyle);
            }
            txt(row, 11, cs.getClientCode(), s.textStyle);
        }

        // TOTAUX row
        XSSFRow totRow = sheet.createRow(rowIdx++);
        txt(totRow, 0, "TOTAUX", s.totalStyle);
        for (int c = 1; c < TRF_COLS; c++) {
            if (c >= 8) totRow.createCell(c).setCellStyle(s.totalStyle);
            else        dbl(totRow, c, totals[c], s.totalMoneyStyle);
        }

        sheet.createRow(rowIdx++); // blank

        // Summary rows: nombre de dossiers, total à reverser, sommes dues
        XSSFRow nbRow = sheet.createRow(rowIdx++);
        txt(nbRow, 0, "Nombre de dossiers", s.textStyle);
        dbl(nbRow, 1, summaries.size(),      s.dataStyle);
        txt(nbRow, 9, "Vir\u00e9 le", s.textStyle);

        XSSFRow revRow = sheet.createRow(rowIdx++);
        txt(revRow, 3, "Total \u00e0 reverser ", s.textStyle);
        dbl(revRow, 4, totals[5], s.moneyStyle);

        XSSFRow dueRow = sheet.createRow(rowIdx++);
        txt(dueRow, 3, "Sommes dues ", s.textStyle);
        dbl(dueRow, 4, totals[7], s.moneyStyle);

        sheet.createRow(rowIdx++); // blank before bottom sections

        // Bottom sections side-by-side
        rowIdx = writeBottomSections(sheet, summaries, rowIdx, today, s);

        autoSize(sheet, TRF_COLS);
        sheet.createFreezePane(0, 1);
    }

    /**
     * Bottom sections layout (mirrors TRF_04 reference):
     *   Cols A-C (left):  VIREMENTS CLIENTS auto + SOUS TOTAL 1
     *                     VIREMENTS MANUELLES    + SOUS TOTAL 2
     *                     TOTAUX VIREMENTS
     *   Cols E-H (right): CHEQUES / NON COMP / COMP PART / DEBITEURS
     *   (all written in parallel rows)
     */
    private int writeBottomSections(XSSFSheet sheet, List<ClientSummary> summaries,
                                    int startRow, LocalDate today, Styles s) {
        List<ClientSummary> autoVirt  = summaries.stream().filter(ClientSummary::needsAutoVirement).collect(Collectors.toList());
        List<ClientSummary> manVirt   = summaries.stream().filter(ClientSummary::needsManualVirement).collect(Collectors.toList());
        List<ClientSummary> cheques   = summaries.stream().filter(ClientSummary::needsCheque).collect(Collectors.toList());
        List<ClientSummary> nonComp   = summaries.stream().filter(ClientSummary::isNonCompensation).collect(Collectors.toList());
        List<ClientSummary> compPart  = summaries.stream().filter(ClientSummary::isPartiallyCompensated).collect(Collectors.toList());
        List<ClientSummary> debiteurs = summaries.stream().filter(ClientSummary::isDebtor).collect(Collectors.toList());

        String virtDate = String.format("%02d/%02d/%d",
                today.getDayOfMonth(), today.getMonthValue(), today.getYear());
        String prevDate = String.format("%02d/%02d/%d",
                today.minusMonths(1).getDayOfMonth(),
                today.minusMonths(1).getMonthValue(),
                today.minusMonths(1).getYear());

        int row = startRow;

        // ── Section 1 header ─────────────────────────────────────────────
        // Left: "VIREMENTS CLIENTS DD/MM/YYYY" | count  Right: "CHEQUES" | count
        XSSFRow h1 = sheet.createRow(row++);
        txt(h1, 0, "VIREMENTS CLIENTS " + virtDate, s.sectionHeader);
        txt(h1, 1, "", s.sectionHeader);
        txt(h1, 2, "SOMME A REVERSER", s.sectionHeader);
        dbl(h1, 3, autoVirt.size(), s.sectionHeader);
        if (!cheques.isEmpty()) {
            txt(h1, 4, "CHEQUES", s.sectionHeader);
            txt(h1, 5, "SOMME A REVERSER", s.sectionHeader);
            dbl(h1, 6, cheques.size(), s.sectionHeader);
        }

        // Auto virements (left) + cheques (right) side by side
        int maxR1 = Math.max(autoVirt.size(), cheques.size());
        double autoTotal    = 0;
        double chequesTotal = 0;
        for (int i = 0; i < maxR1; i++) {
            XSSFRow r = sheet.createRow(row++);
            if (i < autoVirt.size()) {
                ClientSummary cs = autoVirt.get(i);
                txt(r, 0, cs.getClientName(),            s.textStyle);
                txt(r, 1, cs.getIban(),                  s.textStyle);
                dbl(r, 2, cs.getSommesAReverserFinal(),  s.moneyStyle);
                autoTotal += cs.getSommesAReverserFinal();
            }
            if (i < cheques.size()) {
                ClientSummary cs = cheques.get(i);
                txt(r, 4, cs.getClientName(),            s.textStyle);
                dbl(r, 5, cs.getSommesAReverserFinal(),  s.moneyStyle);
                chequesTotal += cs.getSommesAReverserFinal();
            }
        }

        // SOUS TOTAL 1 (left) | CHEQUES TOTAL (right)
        XSSFRow st1 = sheet.createRow(row++);
        txt(st1, 0, "SOUS TOTAL 1", s.totalStyle);
        txt(st1, 1, "", s.totalStyle);
        dbl(st1, 2, autoTotal, s.totalMoneyStyle);
        if (!cheques.isEmpty()) {
            txt(st1, 4, "TOTAL", s.totalStyle);
            dbl(st1, 5, chequesTotal, s.totalMoneyStyle);
        }

        sheet.createRow(row++); // blank

        // ── Section 2 header ─────────────────────────────────────────────
        // Left: VIREMENTS MANUELLES  Right: NON COMP
        XSSFRow h2 = sheet.createRow(row++);
        txt(h2, 0, "VIREMENTS MANUELLES " + prevDate, s.sectionHeader);
        if (!nonComp.isEmpty()) {
            txt(h2, 4, "NON COMP",  s.sectionHeader);
            txt(h2, 5, "Nous doit", s.sectionHeader);
            dbl(h2, 6, nonComp.size(), s.sectionHeader);
            txt(h2, 7, "Date de virement effectu\u00e9", s.subHeader);
            txt(h2, 8, "Montant envoy\u00e9", s.subHeader);
        }

        int maxR2 = Math.max(manVirt.size(), nonComp.size());
        double manTotal     = 0;
        double nonCompTotal = 0;
        for (int i = 0; i < maxR2; i++) {
            XSSFRow r = sheet.createRow(row++);
            if (i < manVirt.size()) {
                ClientSummary cs = manVirt.get(i);
                txt(r, 0, cs.getClientName(),           s.textStyle);
                txt(r, 1, cs.getIban(),                 s.textStyle);
                dbl(r, 2, cs.getSommesAReverserFinal(), s.moneyStyle);
                manTotal += cs.getSommesAReverserFinal();
            }
            if (i < nonComp.size()) {
                ClientSummary cs = nonComp.get(i);
                txt(r, 4, cs.getClientName(),               s.textStyle);
                dbl(r, 5, cs.getNousDoit_ApreFacturation(), s.moneyStyle);
                txt(r, 6, cs.getIban().isBlank() ? "Par cheque" : "Par Vrt", s.textStyle);
                nonCompTotal += cs.getNousDoit_ApreFacturation();
            }
        }

        // SOUS TOTAL 2 (left) | NON COMP TOTAL (right)
        XSSFRow st2 = sheet.createRow(row++);
        txt(st2, 0, "SOUS TOTAL 2", s.totalStyle);
        txt(st2, 1, "", s.totalStyle);
        dbl(st2, 2, manTotal, s.totalMoneyStyle);
        if (!nonComp.isEmpty()) {
            txt(st2, 4, "TOTAL", s.totalStyle);
            dbl(st2, 5, nonCompTotal, s.totalMoneyStyle);
        }

        sheet.createRow(row++); // blank

        // ── TOTAUX VIREMENTS (left) | COMP PARTIELLE (right) ─────────────
        XSSFRow totVirt = sheet.createRow(row++);
        txt(totVirt, 0, "TOTAUX VIREMENTS", s.totalStyle);
        txt(totVirt, 1, "", s.totalStyle);
        dbl(totVirt, 2, autoTotal + manTotal, s.totalMoneyStyle);
        if (!compPart.isEmpty()) {
            txt(totVirt, 4, "COMP PART", s.sectionHeader);
            txt(totVirt, 5, "NOUS DOIT", s.sectionHeader);
        }

        double cpTotal = 0;
        for (int i = 0; i < compPart.size(); i++) {
            ClientSummary cs = compPart.get(i);
            XSSFRow r = (i == 0) ? (XSSFRow) sheet.getRow(row - 1) : sheet.createRow(row++);
            txt(r, 4, cs.getClientName(),               s.textStyle);
            dbl(r, 5, cs.getNousDoit_ApreFacturation(), s.moneyStyle);
            cpTotal += cs.getNousDoit_ApreFacturation();
        }
        if (!compPart.isEmpty()) {
            XSSFRow cpTot = sheet.createRow(row++);
            txt(cpTot, 4, "TOTAL", s.totalStyle);
            dbl(cpTot, 5, cpTotal, s.totalMoneyStyle);
        }

        sheet.createRow(row++); // blank

        // ── DEBITEURS (right) ─────────────────────────────────────────────
        if (!debiteurs.isEmpty()) {
            XSSFRow debHdr = sheet.createRow(row++);
            txt(debHdr, 4, "DEBITEURS", s.sectionHeader);
            txt(debHdr, 5, "NOUS DOIT", s.sectionHeader);
            dbl(debHdr, 6, debiteurs.size(), s.sectionHeader);

            double debTotal = 0;
            for (ClientSummary cs : debiteurs) {
                XSSFRow r = sheet.createRow(row++);
                txt(r, 4, cs.getClientName(),               s.textStyle);
                dbl(r, 5, cs.getNousDoit_ApreFacturation(), s.moneyStyle);
                debTotal += cs.getNousDoit_ApreFacturation();
            }
            XSSFRow debTot = sheet.createRow(row++);
            txt(debTot, 4, "TOTAL", s.totalStyle);
            dbl(debTot, 5, debTotal, s.totalMoneyStyle);
        }

        return row;
    }

    // =========================================================================
    // Helpers
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

    private void writeValue(XSSFCell cell, Object val, XSSFCellStyle defStyle, XSSFCellStyle dateStyle) {
        if (val instanceof Double d)          { cell.setCellValue(d);              cell.setCellStyle(defStyle);  return; }
        if (val instanceof Number n)          { cell.setCellValue(n.doubleValue()); cell.setCellStyle(defStyle); return; }
        if (val instanceof Boolean b)         { cell.setCellValue(b);              cell.setCellStyle(defStyle);  return; }
        if (val instanceof LocalDateTime ldt) { cell.setCellValue(ldt);            cell.setCellStyle(dateStyle); return; }
        if (val instanceof String str && !str.isBlank()) {
            String c = str.replaceAll("[€$£¥₺]", "").replaceAll("\\p{Z}", "").trim();
            if (!c.isEmpty() && !c.equals("-") && c.matches("[-+]?[\\d.,]+")) {
                cell.setCellValue(ConsolidationRow.parseFrenchDouble(str));
                cell.setCellStyle(defStyle); return;
            }
            cell.setCellValue(str);
            cell.setCellStyle(defStyle); return;
        }
        cell.setCellStyle(defStyle);
    }

    private static String strOf(Object v) {
        if (v == null) return "";
        String s = v.toString().trim();
        try { Double.parseDouble(s); return ""; } catch (NumberFormatException ignored) {}
        return s;
    }

    private static String col(int idx) {
        if (idx < 26) return String.valueOf((char) ('A' + idx));
        return String.valueOf((char) ('A' + idx / 26 - 1)) + (char) ('A' + idx % 26);
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
    static class Styles {
        final XSSFCellStyle headerDark, dataStyle, dateStyle;
        final XSSFCellStyle totalStyle, totalMoneyStyle;
        final XSSFCellStyle moneyStyle, textStyle;
        final XSSFCellStyle sectionHeader, subHeader;

        Styles(XSSFWorkbook wb) {
            DataFormat df = wb.createDataFormat();
            short moneyFmt = df.getFormat("#,##0.00");

            headerDark = wb.createCellStyle();
            XSSFFont hf = wb.createFont(); hf.setBold(true);
            hf.setColor(new XSSFColor(new byte[]{(byte)0xFF,(byte)0xFF,(byte)0xFF}, null));
            hf.setFontHeightInPoints((short)10);
            headerDark.setFont(hf);
            headerDark.setFillForegroundColor(new XSSFColor(new byte[]{(byte)0x1F,(byte)0x4E,(byte)0x79}, null));
            headerDark.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerDark.setVerticalAlignment(VerticalAlignment.CENTER);
            headerDark.setWrapText(true); setBorder(headerDark);

            dataStyle = wb.createCellStyle();
            dataStyle.setVerticalAlignment(VerticalAlignment.CENTER); setBorder(dataStyle);

            dateStyle = wb.createCellStyle(); dateStyle.cloneStyleFrom(dataStyle);
            dateStyle.setDataFormat(wb.getCreationHelper().createDataFormat().getFormat("dd/MM/yyyy"));

            totalStyle = wb.createCellStyle();
            XSSFFont tf = wb.createFont(); tf.setBold(true); tf.setFontHeightInPoints((short)10);
            totalStyle.setFont(tf);
            totalStyle.setFillForegroundColor(new XSSFColor(new byte[]{(byte)0xFF,(byte)0xF2,(byte)0xCC}, null));
            totalStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            totalStyle.setVerticalAlignment(VerticalAlignment.CENTER); setBorder(totalStyle);

            totalMoneyStyle = wb.createCellStyle(); totalMoneyStyle.cloneStyleFrom(totalStyle);
            totalMoneyStyle.setDataFormat(moneyFmt);

            moneyStyle = wb.createCellStyle(); moneyStyle.cloneStyleFrom(dataStyle);
            moneyStyle.setDataFormat(moneyFmt);

            textStyle = wb.createCellStyle(); textStyle.cloneStyleFrom(dataStyle);

            sectionHeader = wb.createCellStyle();
            XSSFFont sf = wb.createFont(); sf.setBold(true); sf.setFontHeightInPoints((short)11);
            sectionHeader.setFont(sf);
            sectionHeader.setFillForegroundColor(new XSSFColor(new byte[]{(byte)0x9D,(byte)0xC3,(byte)0xE6}, null));
            sectionHeader.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            sectionHeader.setVerticalAlignment(VerticalAlignment.CENTER);

            subHeader = wb.createCellStyle();
            XSSFFont shf = wb.createFont(); shf.setBold(true); shf.setFontHeightInPoints((short)9);
            subHeader.setFont(shf);
            subHeader.setFillForegroundColor(new XSSFColor(new byte[]{(byte)0xD9,(byte)0xD9,(byte)0xD9}, null));
            subHeader.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            subHeader.setVerticalAlignment(VerticalAlignment.CENTER); setBorder(subHeader);
        }

        private void setBorder(XSSFCellStyle s) {
            XSSFColor bc = new XSSFColor(new byte[]{(byte)0xD9,(byte)0xD9,(byte)0xD9}, null);
            s.setBorderTop(BorderStyle.THIN);    s.setTopBorderColor(bc);
            s.setBorderBottom(BorderStyle.THIN); s.setBottomBorderColor(bc);
            s.setBorderLeft(BorderStyle.THIN);   s.setLeftBorderColor(bc);
            s.setBorderRight(BorderStyle.THIN);  s.setRightBorderColor(bc);
        }
    }
}