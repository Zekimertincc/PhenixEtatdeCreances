package com.zeki.merger.service;

import com.zeki.merger.AppConfig;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.Normalizer;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;
import java.util.function.BiConsumer;

/**
 * For each company discovered by FolderScanner:
 *   1. Locate the "Espace partagé" subfolder (accent-insensitive).
 *   2. Inside it, find or create the "Etat des créances" subfolder.
 *   3. Read fixed header rows and data rows from the company's "Créances" sheet.
 *   4. Write L_ETAT_DE_CREANCES_[company].xlsx into that subfolder.
 */
public class EtatPublicGenerator {

    private static final String[] OUTPUT_HEADERS = {
        "NOMBRE", "V/REF", "REMIS LE", "ANCIENNETE", "N/REF",
        "DEBITEUR", "CREANCE PRINCIPALE", "RECOUVRE",
        "DONT EN ATTENTE DE FACTURATION", "ETAT", "CLOTURE"
    };

    // Source column indices (0-based in Créances sheet) → output cols 0–10
    private static final int[] SOURCE_COL_MAP = {0, 1, 2, 3, 5, 6, 7, 8, 17, 9, 10};

    private static final int OUT_COLS         = OUTPUT_HEADERS.length;
    private static final int OUT_COL_DEBITEUR = 5;   // "TOTAUX :" label
    private static final int OUT_COL_CREANCE  = 6;   // CREANCE PRINCIPALE → SUM
    private static final int OUT_COL_RECOUVRE = 7;   // RECOUVRE → SUM
    private static final int OUT_COL_ATTENTE  = 8;   // DONT EN ATTENTE → SUM

    private final FolderScanner scanner = new FolderScanner();

    // -------------------------------------------------------------------------
    // Entry point
    // -------------------------------------------------------------------------

    public void generate(File rootFolder, BiConsumer<Double, String> progress) throws Exception {
        progress.accept(0.02, "Scanning: " + rootFolder.getAbsolutePath());

        List<FolderScanner.CompanyFile> companyFiles = scanner.scan(rootFolder);
        if (companyFiles.isEmpty()) {
            progress.accept(1.0, "No company files found in: " + rootFolder.getAbsolutePath());
            return;
        }
        progress.accept(0.05, "Found " + companyFiles.size() + " company file(s).");

        int total = companyFiles.size();
        int done = 0, skipped = 0;

        for (int i = 0; i < total; i++) {
            FolderScanner.CompanyFile cf = companyFiles.get(i);
            double prog = 0.05 + 0.90 * (double) (i + 1) / total;
            progress.accept(prog, "[" + (i + 1) + "/" + total + "] " + cf.companyName());

            try {
                File companyDir = new File(rootFolder, cf.companyName());

                Optional<File> espacePartageDir = findEspacePartageDir(companyDir);
                if (espacePartageDir.isEmpty()) {
                    progress.accept(prog,
                        "  SKIP: no 'Espace partagé' subfolder in " + cf.companyName());
                    skipped++;
                    continue;
                }

                File destDir = resolveEtatCreancesDir(espacePartageDir.get());

                File outputFile = new File(destDir,
                    AppConfig.ETAT_PUBLIC_FILENAME_PREFIX + sanitize(cf.companyName()) + ".xlsx");

                generateForClient(cf.companyName(), cf.excelFile(), outputFile);
                progress.accept(prog, "  → " + outputFile.getAbsolutePath());
                done++;
            } catch (Exception e) {
                progress.accept(prog, "  ERROR: " + e.getMessage());
                skipped++;
            }
        }

        progress.accept(1.0, "Done. Generated: " + done
            + (skipped > 0 ? "  |  skipped/errors: " + skipped : ""));
    }

    // -------------------------------------------------------------------------
    // Directory resolution
    // -------------------------------------------------------------------------

    /**
     * Finds the first direct subdirectory of {@code companyDir} whose
     * normalised name contains both "espace" and "partage".
     */
    private Optional<File> findEspacePartageDir(File companyDir) {
        File[] subDirs = companyDir.listFiles(File::isDirectory);
        if (subDirs == null) return Optional.empty();
        return Arrays.stream(subDirs)
            .filter(d -> {
                String n = normalize(d.getName());
                return n.contains("espace") && n.contains("partage");
            })
            .findFirst();
    }

    /**
     * Inside {@code espacePartageDir}, finds an existing subdirectory whose
     * normalised name contains "etat" and "cr" (same rule as FolderScanner), or
     * creates a new one with the exact canonical name "Etat des créances".
     */
    private File resolveEtatCreancesDir(File espacePartageDir) {
        File[] subDirs = espacePartageDir.listFiles(File::isDirectory);
        if (subDirs != null) {
            for (File d : subDirs) {
                String n = normalize(d.getName());
                if (n.contains("etat") && n.contains("cr")) return d;
            }
        }
        File created = new File(espacePartageDir, "Etat des créances");
        created.mkdir();
        return created;
    }

    // -------------------------------------------------------------------------
    // Per-client logic
    // -------------------------------------------------------------------------

    private void generateForClient(String companyName, File sourceFile, File outputFile)
            throws IOException {

        try (Workbook srcWb = openWorkbook(sourceFile)) {
            Sheet src = srcWb.getSheet(AppConfig.CREANCES_SHEET_NAME);
            if (src == null) {
                throw new IOException("Sheet \"" + AppConfig.CREANCES_SHEET_NAME
                    + "\" not found in " + sourceFile.getName());
            }

            DataFormatter fmt  = new DataFormatter();
            FormulaEvaluator eval = srcWb.getCreationHelper().createFormulaEvaluator();

            String company     = getCellString(src, 3, 7, fmt, eval);
            String addressLine1 = getCellString(src, 4, 7, fmt, eval);
            String addressLine2 = getCellString(src, 5, 7, fmt, eval);
            String contactName  = getCellString(src, 7, 7, fmt, eval);
            String codeClient   = getCellString(src, 12, 0, fmt, eval);

            // Data rows: row 16 onwards, stop when NBRE (col 0) is blank
            List<Object[]> dataRows = new ArrayList<>();
            for (int r = 16; r <= src.getLastRowNum(); r++) {
                Row row = src.getRow(r);
                if (row == null) break;
                Cell nbreCell = row.getCell(0, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (nbreCell == null || fmt.formatCellValue(nbreCell).isBlank()) break;

                Object[] out = new Object[OUT_COLS];
                for (int oc = 0; oc < OUT_COLS; oc++) {
                    Cell cell = row.getCell(SOURCE_COL_MAP[oc],
                        Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    out[oc] = readCellValue(cell, fmt, eval);
                }
                dataRows.add(out);
            }

            writeOutput(company, addressLine1, addressLine2,
                contactName, codeClient, dataRows, outputFile);
        }
    }

    // -------------------------------------------------------------------------
    // Writer
    // -------------------------------------------------------------------------

    private void writeOutput(String companyName, String addr1, String addr2,
                              String contactName, String codeClient,
                              List<Object[]> dataRows, File outputFile) throws IOException {

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Etat Public");

            XSSFCellStyle plain = buildPlainStyle(wb);
            XSSFCellStyle hdr   = buildHeaderStyle(wb);
            XSSFCellStyle data  = buildDataStyle(wb);
            XSSFCellStyle date  = buildDateStyle(wb, data);
            XSSFCellStyle total = buildTotalStyle(wb);

            int rowIdx = 0;

            // Company info block (rows 0–7)
            putString(sheet.createRow(rowIdx++), 0, companyName,  plain);
            putString(sheet.createRow(rowIdx++), 0, addr1,        plain);
            putString(sheet.createRow(rowIdx++), 0, addr2,        plain);
            sheet.createRow(rowIdx++);                             // blank
            putString(sheet.createRow(rowIdx++), 0, contactName,  plain);
            sheet.createRow(rowIdx++);                             // blank
            putString(sheet.createRow(rowIdx++), 0, codeClient,   plain);
            sheet.createRow(rowIdx++);                             // blank separator

            // Header row (row index 8)
            XSSFRow headerRow = sheet.createRow(rowIdx++);
            for (int c = 0; c < OUT_COLS; c++) {
                XSSFCell cell = headerRow.createCell(c);
                cell.setCellValue(OUTPUT_HEADERS[c]);
                cell.setCellStyle(hdr);
            }

            // Data rows (start at row index 9)
            int dataStartRow = rowIdx;
            for (Object[] dr : dataRows) {
                XSSFRow row = sheet.createRow(rowIdx++);
                for (int c = 0; c < OUT_COLS; c++) {
                    writeValue(row.createCell(c), dr[c], data, date);
                }
            }
            int dataEndRow = rowIdx - 1;

            // TOTAUX row
            XSSFRow totRow = sheet.createRow(rowIdx);
            for (int c = 0; c < OUT_COLS; c++) {
                totRow.createCell(c).setCellStyle(total);
            }
            XSSFCell lbl = totRow.createCell(OUT_COL_DEBITEUR);
            lbl.setCellValue("TOTAUX :");
            lbl.setCellStyle(total);

            if (!dataRows.isEmpty()) {
                for (int c : new int[]{OUT_COL_CREANCE, OUT_COL_RECOUVRE, OUT_COL_ATTENTE}) {
                    XSSFCell sc = totRow.createCell(c);
                    sc.setCellStyle(total);
                    sc.setCellFormula("SUM(" + colLetter(c) + (dataStartRow + 1)
                        + ":" + colLetter(c) + (dataEndRow + 1) + ")");
                }
            }

            // Presentation
            for (int c = 0; c < OUT_COLS; c++) {
                sheet.autoSizeColumn(c);
                sheet.setColumnWidth(c, Math.min(sheet.getColumnWidth(c) + 512, 25_000));
            }
            sheet.createFreezePane(0, 9);

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                wb.write(fos);
            }
        }
    }

    // -------------------------------------------------------------------------
    // Reading helpers
    // -------------------------------------------------------------------------

    private String getCellString(Sheet sheet, int r, int c,
                                  DataFormatter fmt, FormulaEvaluator eval) {
        Row row = sheet.getRow(r);
        if (row == null) return "";
        Cell cell = row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (cell == null) return "";
        return fmt.formatCellValue(cell, eval).trim();
    }

    private Object readCellValue(Cell cell, DataFormatter fmt, FormulaEvaluator eval) {
        if (cell == null) return "";
        CellType type = cell.getCellType() == CellType.FORMULA
            ? cell.getCachedFormulaResultType()
            : cell.getCellType();
        return switch (type) {
            case NUMERIC -> DateUtil.isCellDateFormatted(cell)
                ? DateUtil.getJavaDate(cell.getNumericCellValue())
                    .toInstant().atZone(java.time.ZoneId.systemDefault()).toLocalDateTime()
                : cell.getNumericCellValue();
            case BOOLEAN -> cell.getBooleanCellValue();
            case STRING  -> cell.getStringCellValue();
            default      -> fmt.formatCellValue(cell, eval);
        };
    }

    private Workbook openWorkbook(File file) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        return file.getName().toLowerCase().endsWith(".xls")
            ? new HSSFWorkbook(fis)
            : new XSSFWorkbook(fis);
    }

    // -------------------------------------------------------------------------
    // Write helpers
    // -------------------------------------------------------------------------

    private void putString(XSSFRow row, int col, String value, XSSFCellStyle style) {
        XSSFCell cell = row.createCell(col);
        cell.setCellValue(value != null ? value : "");
        cell.setCellStyle(style);
    }

    private void writeValue(XSSFCell cell, Object val,
                             XSSFCellStyle def, XSSFCellStyle dateStyle) {
        if (val instanceof Double d) {
            cell.setCellValue(d);
            cell.setCellStyle(def);
        } else if (val instanceof Boolean b) {
            cell.setCellValue(b);
            cell.setCellStyle(def);
        } else if (val instanceof LocalDateTime ldt) {
            cell.setCellValue(ldt);
            cell.setCellStyle(dateStyle);
        } else {
            cell.setCellValue(val != null ? val.toString() : "");
            cell.setCellStyle(def);
        }
    }

    private static String colLetter(int idx) {
        if (idx < 26) return String.valueOf((char) ('A' + idx));
        return String.valueOf((char) ('A' + idx / 26 - 1)) + (char) ('A' + idx % 26);
    }

    // -------------------------------------------------------------------------
    // Style builders
    // -------------------------------------------------------------------------

    private XSSFCellStyle buildPlainStyle(XSSFWorkbook wb) {
        XSSFCellStyle s = wb.createCellStyle();
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        return s;
    }

    private XSSFCellStyle buildHeaderStyle(XSSFWorkbook wb) {
        XSSFCellStyle s = wb.createCellStyle();
        XSSFFont f = wb.createFont();
        f.setBold(true);
        f.setColor(new XSSFColor(new byte[]{(byte) 0xFF, (byte) 0xFF, (byte) 0xFF}, null));
        f.setFontHeightInPoints((short) 10);
        s.setFont(f);
        s.setFillForegroundColor(
            new XSSFColor(new byte[]{(byte) 0x1F, (byte) 0x4E, (byte) 0x79}, null));
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        return s;
    }

    private XSSFCellStyle buildDataStyle(XSSFWorkbook wb) {
        XSSFCellStyle s = wb.createCellStyle();
        XSSFColor bc = new XSSFColor(new byte[]{(byte) 0xD9, (byte) 0xD9, (byte) 0xD9}, null);
        s.setBorderTop(BorderStyle.THIN);
        s.setBorderBottom(BorderStyle.THIN);
        s.setBorderLeft(BorderStyle.THIN);
        s.setBorderRight(BorderStyle.THIN);
        s.setTopBorderColor(bc);
        s.setBottomBorderColor(bc);
        s.setLeftBorderColor(bc);
        s.setRightBorderColor(bc);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        return s;
    }

    private XSSFCellStyle buildDateStyle(XSSFWorkbook wb, XSSFCellStyle base) {
        XSSFCellStyle s = wb.createCellStyle();
        s.cloneStyleFrom(base);
        s.setDataFormat(wb.getCreationHelper().createDataFormat().getFormat("dd/MM/yyyy"));
        return s;
    }

    private XSSFCellStyle buildTotalStyle(XSSFWorkbook wb) {
        XSSFCellStyle s = wb.createCellStyle();
        XSSFFont f = wb.createFont();
        f.setBold(true);
        f.setFontHeightInPoints((short) 10);
        s.setFont(f);
        s.setFillForegroundColor(
            new XSSFColor(new byte[]{(byte) 0xFF, (byte) 0xF2, (byte) 0xCC}, null));
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        XSSFColor bc = new XSSFColor(new byte[]{(byte) 0xD9, (byte) 0xD9, (byte) 0xD9}, null);
        s.setBorderTop(BorderStyle.THIN);
        s.setBorderBottom(BorderStyle.THIN);
        s.setBorderLeft(BorderStyle.THIN);
        s.setBorderRight(BorderStyle.THIN);
        s.setTopBorderColor(bc);
        s.setBottomBorderColor(bc);
        s.setLeftBorderColor(bc);
        s.setRightBorderColor(bc);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        return s;
    }

    // -------------------------------------------------------------------------
    // Utility
    // -------------------------------------------------------------------------

    private static String normalize(String s) {
        return Normalizer.normalize(s, Normalizer.Form.NFD)
            .replaceAll("\\p{M}", "")
            .toLowerCase();
    }

    private static String sanitize(String name) {
        return name.replaceAll("[\\\\/:*?\"<>|]", "_");
    }
}
