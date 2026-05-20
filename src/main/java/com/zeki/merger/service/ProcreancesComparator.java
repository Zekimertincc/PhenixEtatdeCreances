package com.zeki.merger.service;

import com.zeki.merger.trf.DataReader;
import com.zeki.merger.trf.model.ConsolidationRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.function.BiConsumer;
import java.util.stream.Collectors;

public class ProcreancesComparator {

    private static final double TOLERANCE = 0.05;

    // PROCREANCES columns (0-based)
    private static final int PC_CODE = 1;
    private static final int PC_NOM  = 2;
    private static final int PC_HONO = 5; // Hono. TTC

    // ConsolidationGenerale columns (0-based)
    private static final int CS_NAME        = 0;
    private static final int CS_CODE_COL    = 1;
    private static final int CS_COMM_FEUIL1 = 2;  // Feuil1 sheet — col C = Commission TTC
    private static final int CS_COMM_CONSO  = 24; // Consolidation sheet — col Y = Commission HT → ×1.2 for TTC

    private static final String[] COL_HEADERS = {
        "CLIENT", "N° CLIENT",
        "PROC Hono.TTC", "CONSO Commission TTC", "DIFF"
    };

    private static double round2(double v) { return Math.round(v * 100.0) / 100.0; }

    // =========================================================================
    // Public entry point
    // =========================================================================

    public File compare(File procFile, File consoFile, File outputFolder,
                        BiConsumer<Double, String> progress) throws Exception {

        progress.accept(0.0, "Comparaison PROCREANCES vs ConsolidationGenerale");

        // 1. Read PROCREANCES
        progress.accept(0.1, "Lecture " + procFile.getName() + "...");
        Map<String, double[]> procSums = new LinkedHashMap<>();
        Map<String, String[]> procMeta = new LinkedHashMap<>();

        try (Workbook wb = openWorkbook(procFile)) {
            Sheet sheet = wb.getSheetAt(0);
            DataFormatter    fmt = new DataFormatter();
            FormulaEvaluator ev  = wb.getCreationHelper().createFormulaEvaluator();
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                String code = cellStr(row, PC_CODE, fmt, ev);
                String name = cellStr(row, PC_NOM,  fmt, ev);
                if (code.isBlank() || name.isBlank()) continue;
                if (name.startsWith("Total")) continue;

                String key = DataReader.normalize(name);
                procSums.computeIfAbsent(key, k -> new double[1]);
                procMeta.putIfAbsent(key, new String[]{name, code});
                procSums.get(key)[0] += cellDouble(row, PC_HONO, fmt, ev);
            }
        }
        progress.accept(0.2, procSums.size() + " clients lus depuis PROCREANCES.");

        // 2. Read ConsolidationGenerale
        progress.accept(0.3, "Lecture " + consoFile.getName() + "...");
        Map<String, double[]> consoSums = new LinkedHashMap<>();
        Map<String, String[]> consoMeta = new LinkedHashMap<>();

        try (Workbook wb = openWorkbook(consoFile)) {
            Sheet sheet = wb.getSheet("Feuil1");
            boolean isFeuil1 = sheet != null;
            if (sheet == null) sheet = wb.getSheet("Consolidation");
            if (sheet == null) sheet = wb.getSheetAt(0);
            int colComm = isFeuil1 ? CS_COMM_FEUIL1 : CS_COMM_CONSO;

            DataFormatter    fmt = new DataFormatter();
            FormulaEvaluator ev  = wb.getCreationHelper().createFormulaEvaluator();

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                String name = cellStr(row, CS_NAME, fmt, ev);
                if (name.isBlank() || name.startsWith("Total") || name.startsWith("TOTAUX")) continue;

                String key  = DataReader.normalize(name);
                String code = cellStr(row, CS_CODE_COL, fmt, ev);
                double comm = cellDouble(row, colComm, fmt, ev);
                // When reading from Consolidation sheet, col Y(24) is HT — multiply by 1.2 for TTC
                if (!isFeuil1) comm = round2(comm * 1.2);

                if (consoSums.containsKey(key)) {
                    consoSums.get(key)[0] += comm;
                } else {
                    consoSums.put(key, new double[]{comm});
                    consoMeta.put(key, new String[]{name, code});
                }
            }
        }
        progress.accept(0.5, consoSums.size() + " clients lus depuis ConsolidationGenerale.");

        // 3. Match and compare
        List<DiffRow>          allRows       = new ArrayList<>();
        List<UnmatchedProcRow> unmatchedProc = new ArrayList<>();
        Set<String>            matchedConso  = new HashSet<>();

        for (Map.Entry<String, double[]> e : procSums.entrySet()) {
            String   procKey = e.getKey();
            double[] pSums   = e.getValue();
            String[] pMeta   = procMeta.get(procKey);

            double[] cSums = consoSums.get(procKey);
            String   cKey  = procKey;
            if (cSums == null) {
                for (Map.Entry<String, double[]> ce : consoSums.entrySet()) {
                    String k = ce.getKey();
                    if (procKey.contains(k) || k.contains(procKey)) {
                        cSums = ce.getValue(); cKey = k; break;
                    }
                }
            }
            if (cSums == null) {
                unmatchedProc.add(new UnmatchedProcRow(pMeta[0], pMeta[1], round2(pSums[0])));
                continue;
            }
            matchedConso.add(cKey);

            double pH = round2(pSums[0]), cH = round2(cSums[0]);
            double diffH = round2(pH - cH);
            boolean discrep = Math.abs(diffH) > TOLERANCE;

            String[] cMeta     = consoMeta.get(cKey);
            String consoName   = cMeta != null ? cMeta[0] : pMeta[0];
            String clientCode  = !pMeta[1].isBlank() ? pMeta[1]
                                 : (cMeta != null ? cMeta[1] : "");

            allRows.add(new DiffRow(consoName, clientCode, pH, cH, diffH, discrep));
        }

        List<UnmatchedConsoRow> unmatchedConso = consoSums.entrySet().stream()
            .filter(e -> !matchedConso.contains(e.getKey()))
            .map(e -> {
                String[] m  = consoMeta.get(e.getKey());
                double[] cs = e.getValue();
                return new UnmatchedConsoRow(m[0], m[1], round2(cs[0]));
            })
            .collect(Collectors.toList());

        List<DiffRow> discrepancies = allRows.stream()
            .filter(DiffRow::hasDiscrepancy)
            .collect(Collectors.toList());

        ComparisonResult result = new ComparisonResult(
            allRows, discrepancies, unmatchedProc, unmatchedConso);

        // 4. Log results
        progress.accept(0.7, String.format(
            "%d clients comparés — %d écarts | %d non appariés",
            allRows.size(), discrepancies.size(),
            unmatchedProc.size() + unmatchedConso.size()));

        if (!discrepancies.isEmpty()) {
            progress.accept(0.7, "── Écarts ──");
            for (DiffRow dr : discrepancies) {
                progress.accept(0.75, String.format(
                    "⚠ %-25s  Hono: %+9.2f", dr.clientName(), dr.diff()));
            }
        }
        if (!unmatchedProc.isEmpty()) {
            progress.accept(0.8, "── Non appariés (PROCREANCES) ──");
            for (UnmatchedProcRow r : unmatchedProc)
                progress.accept(0.8, "  " + r.name() + (r.code().isBlank() ? "" : " (" + r.code() + ")"));
        }
        if (!unmatchedConso.isEmpty()) {
            progress.accept(0.8, "── Non appariés (Conso) ──");
            for (UnmatchedConsoRow r : unmatchedConso)
                progress.accept(0.8, "  " + r.name() + (r.code().isBlank() ? "" : " (" + r.code() + ")"));
        }

        // 5. Write report
        progress.accept(0.9, "Écriture du rapport Excel...");
        File report = writeReport(result, outputFolder);
        progress.accept(1.0, "→ Rapport: " + report.getName());
        return report;
    }

    // =========================================================================
    // Excel report
    // =========================================================================

    private File writeReport(ComparisonResult result, File outputFolder) throws IOException {
        String ts = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd_HH-mm"));
        File outFile = new File(outputFolder,
            "comparison_PROCREANCES_vs_CONSO_" + ts + ".xlsx");

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            ReportStyles s = new ReportStyles(wb);
            writeMatchedSheet(wb.createSheet("Récapitulatif"), result.allRows(),       s, false);
            writeMatchedSheet(wb.createSheet("Écarts"),        result.discrepancies(), s, true);
            writeNonApparieSheet(wb.createSheet("Non appariés"),
                result.unmatchedProcreances(), result.unmatchedConso(), s);
            try (FileOutputStream fos = new FileOutputStream(outFile)) {
                wb.write(fos);
            }
        }
        return outFile;
    }

    private void writeMatchedSheet(XSSFSheet sheet, List<DiffRow> rows,
                                   ReportStyles s, boolean withSummary) {
        XSSFRow hdr = sheet.createRow(0);
        for (int c = 0; c < COL_HEADERS.length; c++) {
            XSSFCell cell = hdr.createCell(c);
            cell.setCellValue(COL_HEADERS[c]);
            cell.setCellStyle(s.header);
        }

        List<DiffRow> sorted = rows.stream()
            .sorted(Comparator.comparing(DiffRow::hasDiscrepancy).reversed()
                .thenComparing(DiffRow::clientName))
            .collect(Collectors.toList());

        int rowIdx = 1;
        for (DiffRow dr : sorted) {
            XSSFRow row = sheet.createRow(rowIdx++);
            str(row, 0, dr.clientName(),  s.text);
            str(row, 1, dr.clientCode(),  s.text);
            num(row, 2, dr.procHono(),    s.money);
            num(row, 3, dr.consoCommTtc(), s.money);
            numDiff(row, 4, dr.diff(),    s);
        }

        if (withSummary && !rows.isEmpty()) {
            rowIdx++;
            XSSFRow sumRow = sheet.createRow(rowIdx);
            str(sumRow, 0, "TOTAUX", s.totalText);
            double tPH = 0, tCH = 0, tD = 0;
            for (DiffRow dr : rows) {
                tPH += dr.procHono();
                tCH += dr.consoCommTtc();
                tD  += dr.diff();
            }
            num(sumRow, 2, round2(tPH), s.totalMoney);
            num(sumRow, 3, round2(tCH), s.totalMoney);
            num(sumRow, 4, round2(tD),  s.totalMoney);
        }

        sheet.setColumnWidth(0, 28 * 256);
        sheet.setColumnWidth(1, 10 * 256);
        for (int c = 2; c < COL_HEADERS.length; c++) sheet.setColumnWidth(c, 14 * 256);
        sheet.createFreezePane(0, 1);
    }

    private void numDiff(XSSFRow row, int col, double diff, ReportStyles s) {
        XSSFCell cell = row.createCell(col);
        cell.setCellValue(diff);
        if (Math.abs(diff) < TOLERANCE)  cell.setCellStyle(s.money);
        else if (diff > 0)               cell.setCellStyle(s.ecartGreen);
        else                             cell.setCellStyle(s.ecartRed);
    }

    private void writeNonApparieSheet(XSSFSheet sheet,
                                      List<UnmatchedProcRow>  unmatchedProc,
                                      List<UnmatchedConsoRow> unmatchedConso,
                                      ReportStyles s) {
        int rowIdx = 0;

        // Table 1 — PROCREANCES side
        str(sheet.createRow(rowIdx++), 0, "Dans PROCREANCES, absent de Conso", s.sectionLabel);
        String[] procCols = {"CLIENT", "N° CLIENT", "Hono.TTC"};
        XSSFRow t1hdr = sheet.createRow(rowIdx++);
        for (int c = 0; c < procCols.length; c++) {
            XSSFCell cell = t1hdr.createCell(c);
            cell.setCellValue(procCols[c]);
            cell.setCellStyle(s.header);
        }
        if (unmatchedProc.isEmpty()) {
            str(sheet.createRow(rowIdx++), 0, "(aucun)", s.text);
        } else {
            for (UnmatchedProcRow r : unmatchedProc) {
                XSSFRow row = sheet.createRow(rowIdx++);
                str(row, 0, r.name(),    s.text);
                str(row, 1, r.code(),    s.text);
                num(row, 2, r.honoTtc(), s.money);
            }
        }

        rowIdx++;

        // Table 2 — Conso side
        str(sheet.createRow(rowIdx++), 0, "Dans Conso, absent de PROCREANCES", s.sectionLabel);
        String[] consoCols = {"CLIENT", "N° CLIENT", "Commission TTC"};
        XSSFRow t2hdr = sheet.createRow(rowIdx++);
        for (int c = 0; c < consoCols.length; c++) {
            XSSFCell cell = t2hdr.createCell(c);
            cell.setCellValue(consoCols[c]);
            cell.setCellStyle(s.header);
        }
        if (unmatchedConso.isEmpty()) {
            str(sheet.createRow(rowIdx++), 0, "(aucun)", s.text);
        } else {
            for (UnmatchedConsoRow r : unmatchedConso) {
                XSSFRow row = sheet.createRow(rowIdx++);
                str(row, 0, r.name(),    s.text);
                str(row, 1, r.code(),    s.text);
                num(row, 2, r.commTtc(), s.money);
            }
        }

        sheet.setColumnWidth(0, 28 * 256);
        sheet.setColumnWidth(1, 10 * 256);
        sheet.setColumnWidth(2, 14 * 256);
    }

    // =========================================================================
    // Cell helpers
    // =========================================================================

    private void str(XSSFRow row, int col, String val, XSSFCellStyle style) {
        XSSFCell cell = row.createCell(col);
        cell.setCellValue(val != null ? val : "");
        cell.setCellStyle(style);
    }

    private void num(XSSFRow row, int col, double val, XSSFCellStyle style) {
        XSSFCell cell = row.createCell(col);
        cell.setCellValue(val);
        cell.setCellStyle(style);
    }

    private String cellStr(Row row, int col, DataFormatter fmt, FormulaEvaluator eval) {
        Cell cell = row.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (cell == null) return "";
        return fmt.formatCellValue(cell, eval).trim();
    }

    private double cellDouble(Row row, int col, DataFormatter fmt, FormulaEvaluator eval) {
        Cell cell = row.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (cell == null) return 0.0;
        CellType type = cell.getCellType() == CellType.FORMULA
            ? cell.getCachedFormulaResultType() : cell.getCellType();
        if (type == CellType.NUMERIC) return cell.getNumericCellValue();
        return ConsolidationRow.parseFrenchDouble(fmt.formatCellValue(cell, eval).trim());
    }

    private Workbook openWorkbook(File file) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        return file.getName().toLowerCase().endsWith(".xls")
            ? new HSSFWorkbook(fis)
            : new XSSFWorkbook(fis);
    }

    // =========================================================================
    // Styles
    // =========================================================================

    private static class ReportStyles {
        final XSSFCellStyle header, text, money, ecartGreen, ecartRed,
                            sectionLabel, totalText, totalMoney;

        ReportStyles(XSSFWorkbook wb) {
            DataFormat df  = wb.createDataFormat();
            short moneyFmt = df.getFormat("#,##0.00");

            XSSFFont white = wb.createFont();
            white.setBold(true);
            white.setColor(new XSSFColor(new byte[]{(byte)0xFF,(byte)0xFF,(byte)0xFF}, null));

            header = wb.createCellStyle();
            header.setFont(white);
            header.setFillForegroundColor(
                new XSSFColor(new byte[]{(byte)0x1F,(byte)0x4E,(byte)0x79}, null));
            header.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            header.setVerticalAlignment(VerticalAlignment.CENTER);
            header.setWrapText(true);

            text = wb.createCellStyle();
            text.setVerticalAlignment(VerticalAlignment.CENTER);

            money = wb.createCellStyle();
            money.cloneStyleFrom(text);
            money.setDataFormat(moneyFmt);

            XSSFFont greenFont = wb.createFont();
            greenFont.setBold(true);
            greenFont.setColor(new XSSFColor(new byte[]{(byte)0x27,(byte)0x62,(byte)0x21}, null));
            ecartGreen = wb.createCellStyle();
            ecartGreen.setFont(greenFont);
            ecartGreen.setFillForegroundColor(
                new XSSFColor(new byte[]{(byte)0xC6,(byte)0xEF,(byte)0xCE}, null));
            ecartGreen.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            ecartGreen.setDataFormat(moneyFmt);
            ecartGreen.setVerticalAlignment(VerticalAlignment.CENTER);

            XSSFFont redFont = wb.createFont();
            redFont.setBold(true);
            redFont.setColor(new XSSFColor(new byte[]{(byte)0x9C,(byte)0x00,(byte)0x06}, null));
            ecartRed = wb.createCellStyle();
            ecartRed.setFont(redFont);
            ecartRed.setFillForegroundColor(
                new XSSFColor(new byte[]{(byte)0xFF,(byte)0xC7,(byte)0xCE}, null));
            ecartRed.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            ecartRed.setDataFormat(moneyFmt);
            ecartRed.setVerticalAlignment(VerticalAlignment.CENTER);

            XSSFFont sectionFont = wb.createFont();
            sectionFont.setBold(true);
            sectionFont.setFontHeightInPoints((short)11);
            sectionLabel = wb.createCellStyle();
            sectionLabel.setFont(sectionFont);
            sectionLabel.setVerticalAlignment(VerticalAlignment.CENTER);

            XSSFFont boldFont = wb.createFont();
            boldFont.setBold(true);
            totalText = wb.createCellStyle();
            totalText.setFont(boldFont);
            totalText.setVerticalAlignment(VerticalAlignment.CENTER);

            totalMoney = wb.createCellStyle();
            totalMoney.setFont(boldFont);
            totalMoney.setDataFormat(moneyFmt);
            totalMoney.setVerticalAlignment(VerticalAlignment.CENTER);
        }
    }
}
