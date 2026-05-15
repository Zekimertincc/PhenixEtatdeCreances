package com.zeki.merger.service;

import com.zeki.merger.trf.DataReader;
import com.zeki.merger.trf.model.ConsolidationRow;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.function.BiConsumer;

/**
 * Compares Controle_Facturation.xlsx (sheet "Controle") against
 * ConsolidationGenerale.xlsx (Feuil1 if present, else Consolidation sheet).
 * Checks col 8 (TOTAL TTC) of Contrôle vs col 24 (MONTANT A FACTURER TTC) of Conso.
 */
public class ConsoControleComparator {

    private static final double TOLERANCE = 0.05;

    // Controle sheet columns (0-based)
    private static final int CT_NOM       = 0;  // Clients
    private static final int CT_TOTAL_TTC = 8;  // TOTAL TTC

    // Conso / Feuil1 columns (0-based)
    private static final int CS_NAME        = 0;
    private static final int CS_MONTANT_TTC = 24; // MONTANT A FACTURER TTC

    private static final String[] OUT_HEADERS = {
        "CLIENT", "CONTROLE TOTAL TTC", "CONSO MONTANT TTC", "DIFF"
    };

    // =========================================================================
    // Public entry point
    // =========================================================================

    public File compare(File controleFile, File consoFile,
                        File outputFolder,
                        BiConsumer<Double, String> progress) throws Exception {

        progress.accept(0.0, "Comparaison Contrôle vs Consolidation");

        // 1. Read Controle_Facturation
        progress.accept(0.1, "Lecture " + controleFile.getName() + "...");
        Map<String, Double>  ctSums  = new LinkedHashMap<>();
        Map<String, String>  ctNames = new LinkedHashMap<>();

        try (Workbook wb = new XSSFWorkbook(new FileInputStream(controleFile))) {
            Sheet sheet = wb.getSheet("Controle");
            if (sheet == null) sheet = wb.getSheetAt(0);
            DataFormatter    fmt = new DataFormatter();
            FormulaEvaluator ev  = wb.getCreationHelper().createFormulaEvaluator();

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                String name = cellStr(row, CT_NOM, fmt, ev);
                if (name.isBlank()) continue;
                double totalTtc = cellDouble(row, CT_TOTAL_TTC, fmt, ev);
                String key = DataReader.normalize(name);
                ctSums.merge(key, totalTtc, Double::sum);
                ctNames.putIfAbsent(key, name);
            }
        }
        progress.accept(0.3, ctSums.size() + " clients lus depuis Contrôle.");

        // 2. Read ConsolidationGenerale (Feuil1 preferred)
        progress.accept(0.4, "Lecture " + consoFile.getName() + "...");
        Map<String, Double>  csSums  = new LinkedHashMap<>();
        Map<String, String>  csNames = new LinkedHashMap<>();

        try (Workbook wb = new XSSFWorkbook(new FileInputStream(consoFile))) {
            Sheet sheet = wb.getSheet("Feuil1");
            if (sheet == null) sheet = wb.getSheet("Consolidation");
            if (sheet == null) sheet = wb.getSheetAt(0);
            DataFormatter    fmt = new DataFormatter();
            FormulaEvaluator ev  = wb.getCreationHelper().createFormulaEvaluator();

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                String name = cellStr(row, CS_NAME, fmt, ev);
                if (name.isBlank() || name.startsWith("Total") || name.startsWith("TOTAUX")) continue;
                double montant = cellDouble(row, CS_MONTANT_TTC, fmt, ev);
                String key = DataReader.normalize(name);
                csSums.merge(key, montant, Double::sum);
                csNames.putIfAbsent(key, name);
            }
        }
        progress.accept(0.6, csSums.size() + " clients lus depuis Consolidation.");

        // 3. Match and compare
        progress.accept(0.7, "Comparaison...");
        List<Object[]> resultRows = new ArrayList<>();
        double totalCtSum = 0, totalCsSum = 0, totalDiff = 0;
        int ecarts = 0;

        Set<String> matched = new HashSet<>();

        for (Map.Entry<String, Double> e : ctSums.entrySet()) {
            String key  = e.getKey();
            double ctVal = e.getValue();

            Double csVal = csSums.get(key);
            if (csVal == null) {
                for (Map.Entry<String, Double> ce : csSums.entrySet()) {
                    String k = ce.getKey();
                    if (key.contains(k) || k.contains(key)) {
                        csVal = ce.getValue(); key = k; break;
                    }
                }
            }
            double csResolved = csVal != null ? csVal : 0.0;
            if (csVal != null) matched.add(key);

            double diff = Math.round((ctVal - csResolved) * 100.0) / 100.0;
            boolean hasEcart = Math.abs(diff) > TOLERANCE;
            if (hasEcart) ecarts++;

            String displayName = ctNames.getOrDefault(e.getKey(), e.getKey());
            resultRows.add(new Object[]{displayName, ctVal, csResolved, diff, hasEcart});
            totalCtSum += ctVal;
            totalCsSum += csResolved;
            totalDiff  += diff;
        }

        // Unmatched from Conso (not in Controle)
        for (Map.Entry<String, Double> e : csSums.entrySet()) {
            if (!matched.contains(e.getKey())) {
                String displayName = csNames.getOrDefault(e.getKey(), e.getKey());
                double csVal = e.getValue();
                double diff  = Math.round((0.0 - csVal) * 100.0) / 100.0;
                resultRows.add(new Object[]{displayName + " [Conso only]", 0.0, csVal, diff, true});
                totalCsSum += csVal;
                totalDiff  += diff;
                ecarts++;
            }
        }

        progress.accept(0.8, String.format(
            "%d clients comparés — %d écart(s)", resultRows.size(), ecarts));

        // 4. Write output
        progress.accept(0.9, "Écriture du rapport...");
        File report = writeReport(resultRows, totalCtSum, totalCsSum, totalDiff, outputFolder);
        progress.accept(1.0, "→ Rapport: " + report.getName());
        return report;
    }

    // =========================================================================
    // Excel output
    // =========================================================================

    private File writeReport(List<Object[]> rows,
                              double totalCt, double totalCs, double totalDiff,
                              File outputFolder) throws IOException {
        String ts = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd_HH-mm"));
        File outFile = new File(outputFolder, "controle_vs_conso_" + ts + ".xlsx");

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Styles s = new Styles(wb);
            XSSFSheet sheet = wb.createSheet("Contrôle vs Conso");

            // Header
            XSSFRow hdr = sheet.createRow(0);
            for (int c = 0; c < OUT_HEADERS.length; c++) {
                XSSFCell cell = hdr.createCell(c);
                cell.setCellValue(OUT_HEADERS[c]);
                cell.setCellStyle(s.header);
            }

            // Data rows
            int rowIdx = 1;
            for (Object[] r : rows) {
                String  name     = (String)  r[0];
                double  ctVal    = (Double)  r[1];
                double  csVal    = (Double)  r[2];
                double  diff     = (Double)  r[3];
                boolean hasEcart = (Boolean) r[4];

                XSSFRow row = sheet.createRow(rowIdx++);
                XSSFCellStyle lineStyle = hasEcart ? s.redLine : s.greenLine;
                XSSFCellStyle moneyLine = hasEcart ? s.redMoney : s.greenMoney;

                XSSFCell nameCell = row.createCell(0);
                nameCell.setCellValue(name);
                nameCell.setCellStyle(lineStyle);

                XSSFCell c1 = row.createCell(1); c1.setCellValue(ctVal); c1.setCellStyle(moneyLine);
                XSSFCell c2 = row.createCell(2); c2.setCellValue(csVal); c2.setCellStyle(moneyLine);
                XSSFCell c3 = row.createCell(3); c3.setCellValue(diff); c3.setCellStyle(moneyLine);
            }

            // TOTAUX row
            XSSFRow totRow = sheet.createRow(rowIdx);
            XSSFCell lbl = totRow.createCell(0); lbl.setCellValue("TOTAUX"); lbl.setCellStyle(s.total);
            XSSFCell t1 = totRow.createCell(1); t1.setCellValue(totalCt);   t1.setCellStyle(s.totalMoney);
            XSSFCell t2 = totRow.createCell(2); t2.setCellValue(totalCs);   t2.setCellStyle(s.totalMoney);
            XSSFCell t3 = totRow.createCell(3); t3.setCellValue(totalDiff); t3.setCellStyle(s.totalMoney);

            for (int c = 0; c < OUT_HEADERS.length; c++) {
                sheet.autoSizeColumn(c);
                sheet.setColumnWidth(c, Math.min(sheet.getColumnWidth(c) + 512, 20_000));
            }
            sheet.createFreezePane(0, 1);

            try (FileOutputStream fos = new FileOutputStream(outFile)) {
                wb.write(fos);
            }
        }
        return outFile;
    }

    // =========================================================================
    // Cell helpers
    // =========================================================================

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

    // =========================================================================
    // Styles
    // =========================================================================

    private static class Styles {
        final XSSFCellStyle header, greenLine, redLine, greenMoney, redMoney, total, totalMoney;

        Styles(XSSFWorkbook wb) {
            DataFormat df  = wb.createDataFormat();
            short moneyFmt = df.getFormat("#,##0.00");

            XSSFFont whiteFont = wb.createFont();
            whiteFont.setBold(true);
            whiteFont.setColor(new XSSFColor(new byte[]{(byte)0xFF,(byte)0xFF,(byte)0xFF}, null));

            header = wb.createCellStyle();
            header.setFont(whiteFont);
            header.setFillForegroundColor(
                new XSSFColor(new byte[]{(byte)0x1F,(byte)0x4E,(byte)0x79}, null));
            header.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            header.setVerticalAlignment(VerticalAlignment.CENTER);

            greenLine = wb.createCellStyle();
            greenLine.setFillForegroundColor(
                new XSSFColor(new byte[]{(byte)0xC6,(byte)0xEF,(byte)0xCE}, null));
            greenLine.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            greenLine.setVerticalAlignment(VerticalAlignment.CENTER);

            redLine = wb.createCellStyle();
            redLine.setFillForegroundColor(
                new XSSFColor(new byte[]{(byte)0xFF,(byte)0xC7,(byte)0xCE}, null));
            redLine.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            redLine.setVerticalAlignment(VerticalAlignment.CENTER);

            greenMoney = wb.createCellStyle();
            greenMoney.cloneStyleFrom(greenLine);
            greenMoney.setDataFormat(moneyFmt);

            redMoney = wb.createCellStyle();
            redMoney.cloneStyleFrom(redLine);
            redMoney.setDataFormat(moneyFmt);

            XSSFFont boldFont = wb.createFont();
            boldFont.setBold(true);

            total = wb.createCellStyle();
            total.setFont(boldFont);
            total.setFillForegroundColor(
                new XSSFColor(new byte[]{(byte)0xFF,(byte)0xF2,(byte)0xCC}, null));
            total.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            total.setVerticalAlignment(VerticalAlignment.CENTER);

            totalMoney = wb.createCellStyle();
            totalMoney.cloneStyleFrom(total);
            totalMoney.setDataFormat(moneyFmt);
        }
    }
}
