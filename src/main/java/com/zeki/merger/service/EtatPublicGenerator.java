package com.zeki.merger.service;

import com.itextpdf.kernel.colors.ColorConstants;
import com.itextpdf.kernel.colors.DeviceRgb;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.properties.UnitValue;
import com.zeki.merger.AppConfig;
import com.zeki.merger.trf.model.ConsolidationRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.Normalizer;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.function.BiConsumer;

/**
 * For each company discovered by FolderScanner:
 *   1. Locate the destination folder via resolveDestDir().
 *   2. Read fixed header rows and data rows from the company's "Créances" sheet.
 *   3. Write L_ETAT_DE_CREANCES_[company].xlsx + .pdf into that folder.
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
    private static final int OUT_COL_DEBITEUR = 5;
    private static final int OUT_COL_CREANCE  = 6;
    private static final int OUT_COL_RECOUVRE = 7;
    private static final int OUT_COL_ATTENTE  = 8;

    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ofPattern("dd/MM/yyyy");

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
                File destDir    = resolveDestDir(companyDir);
                if (destDir == null) {
                    progress.accept(prog, "  SKIP: cannot resolve destination for " + cf.companyName());
                    skipped++;
                    continue;
                }

                // Delete old L_ETAT files before writing the new one
                File[] oldFiles = destDir.listFiles(f -> {
                    String lower = f.getName().toLowerCase();
                    return f.isFile() && lower.startsWith("l_etat")
                        && (lower.endsWith(".xlsx") || lower.endsWith(".xls") || lower.endsWith(".pdf"));
                });
                if (oldFiles != null) {
                    for (File old : oldFiles) {
                        if (old.delete()) progress.accept(prog, "  Deleted old: " + old.getName());
                    }
                }

                String baseName  = AppConfig.ETAT_PUBLIC_FILENAME_PREFIX + sanitize(cf.companyName());
                File   outputFile = new File(destDir, baseName + ".xlsx");
                File   pdfFile    = new File(destDir, baseName + ".pdf");

                generateForClient(cf.companyName(), cf.excelFile(), outputFile, pdfFile);
                progress.accept(prog, "  → " + outputFile.getName() + " + PDF");
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
     * Resolves the output directory for a company using a 3-level fallback:
     * 1. "Espace partagé" subfolder → "Etat des créances" inside it
     * 2. A direct "Etat des créances"-like subfolder in the company dir
     * 3. Creates "Etat des créances" directly inside the company dir
     */
    private File resolveDestDir(File companyDir) {
        File[] subDirs = companyDir.listFiles(File::isDirectory);
        if (subDirs == null) return null;
        for (File d : subDirs) {
            String n = normalize(d.getName());
            if (n.contains("espace") && n.contains("partage")) {
                return resolveEtatCreancesDir(d);
            }
        }
        for (File d : subDirs) {
            String n = normalize(d.getName());
            if (n.contains("etat") && n.contains("cr")) return d;
        }
        File created = new File(companyDir, "Etat des créances");
        created.mkdir();
        return created;
    }

    /**
     * Inside a parent folder, finds or creates an "Etat des créances" subdirectory.
     */
    private File resolveEtatCreancesDir(File parentDir) {
        File[] subDirs = parentDir.listFiles(File::isDirectory);
        if (subDirs != null) {
            for (File d : subDirs) {
                String n = normalize(d.getName());
                if (n.contains("etat") && n.contains("cr")) return d;
            }
        }
        File created = new File(parentDir, "Etat des créances");
        created.mkdir();
        return created;
    }

    // -------------------------------------------------------------------------
    // Per-client logic
    // -------------------------------------------------------------------------

    private void generateForClient(String companyName, File sourceFile,
                                    File outputFile, File pdfFile) throws Exception {

        try (Workbook srcWb = openWorkbook(sourceFile)) {
            Sheet src = srcWb.getSheet(AppConfig.CREANCES_SHEET_NAME);
            if (src == null) {
                throw new IOException("Sheet \"" + AppConfig.CREANCES_SHEET_NAME
                    + "\" not found in " + sourceFile.getName());
            }

            DataFormatter    fmt  = new DataFormatter();
            FormulaEvaluator eval = srcWb.getCreationHelper().createFormulaEvaluator();

            String company      = getCellString(src, 3, 7, fmt, eval);
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
            writePdf(company, addressLine1, addressLine2,
                contactName, codeClient, dataRows, pdfFile);
        }
    }

    // -------------------------------------------------------------------------
    // XLSX writer
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
    // PDF writer (iText7, landscape A4)
    // -------------------------------------------------------------------------

    private void writePdf(String companyName, String addr1, String addr2,
                           String contactName, String codeClient,
                           List<Object[]> dataRows, File pdfFile) throws Exception {

        PdfDocument pdfDoc = new PdfDocument(new PdfWriter(pdfFile.getAbsolutePath()));
        try (Document doc = new Document(pdfDoc, PageSize.A4.rotate())) {

            doc.setMargins(20, 20, 20, 20);

            // Company info block
            doc.add(new Paragraph(companyName).setBold().setFontSize(11));
            if (!addr1.isBlank())       doc.add(new Paragraph(addr1).setFontSize(9));
            if (!addr2.isBlank())       doc.add(new Paragraph(addr2).setFontSize(9));
            if (!contactName.isBlank()) doc.add(new Paragraph(contactName).setFontSize(9));
            if (!codeClient.isBlank())
                doc.add(new Paragraph("Code client : " + codeClient).setFontSize(9));
            doc.add(new Paragraph(" "));

            // Table — A4 landscape ~841 pt wide minus 40 pt margins = ~801 pt
            float[] colWidths = {38, 58, 58, 58, 58, 130, 78, 78, 92, 58, 58};
            Table table = new Table(UnitValue.createPointArray(colWidths))
                .useAllAvailableWidth();

            // Header row — dark-blue background, white bold text
            DeviceRgb darkBlue = new DeviceRgb(0x1F, 0x4E, 0x79);
            for (String h : OUTPUT_HEADERS) {
                table.addHeaderCell(
                    new com.itextpdf.layout.element.Cell()
                        .add(new Paragraph(h).setBold().setFontSize(7.5f)
                            .setFontColor(ColorConstants.WHITE))
                        .setBackgroundColor(darkBlue)
                        .setPadding(3));
            }

            // Data rows — alternating white / light-grey
            DeviceRgb white     = new DeviceRgb(0xFF, 0xFF, 0xFF);
            DeviceRgb lightGrey = new DeviceRgb(0xF2, 0xF2, 0xF2);

            double totCreance = 0, totRecouvre = 0, totAttente = 0;

            for (int i = 0; i < dataRows.size(); i++) {
                Object[]  dr    = dataRows.get(i);
                DeviceRgb color = (i % 2 == 0) ? white : lightGrey;
                for (int c = 0; c < OUT_COLS; c++) {
                    table.addCell(new com.itextpdf.layout.element.Cell()
                        .add(new Paragraph(fmtPdf(dr[c])).setFontSize(7.5f))
                        .setBackgroundColor(color)
                        .setPadding(2));
                    if (dr[c] instanceof Number n) {
                        double v = n.doubleValue();
                        if      (c == OUT_COL_CREANCE)  totCreance  += v;
                        else if (c == OUT_COL_RECOUVRE) totRecouvre += v;
                        else if (c == OUT_COL_ATTENTE)  totAttente  += v;
                    }
                }
            }

            // TOTAUX row — yellow background, bold
            DeviceRgb yellow = new DeviceRgb(0xFF, 0xF2, 0xCC);
            for (int c = 0; c < OUT_COLS; c++) {
                String text = "";
                if      (c == OUT_COL_DEBITEUR) text = "TOTAUX :";
                else if (c == OUT_COL_CREANCE)  text = fmt2(totCreance);
                else if (c == OUT_COL_RECOUVRE) text = fmt2(totRecouvre);
                else if (c == OUT_COL_ATTENTE)  text = fmt2(totAttente);
                table.addCell(new com.itextpdf.layout.element.Cell()
                    .add(new Paragraph(text).setBold().setFontSize(7.5f))
                    .setBackgroundColor(yellow)
                    .setPadding(2));
            }

            doc.add(table);
        }
    }

    private String fmtPdf(Object val) {
        if (val == null)                    return "";
        if (val instanceof Double d)        return fmt2(d);
        if (val instanceof Number n)        return fmt2(n.doubleValue());
        if (val instanceof LocalDateTime t) return t.toLocalDate().format(DATE_FMT);
        if (val instanceof Boolean b)       return b ? "OUI" : "NON";
        return val.toString().trim();
    }

    private static String fmt2(double v) {
        return String.format(Locale.FRANCE, "%,.2f", v);
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
    // XLSX write helpers
    // -------------------------------------------------------------------------

    private void putString(XSSFRow row, int col, String value, XSSFCellStyle style) {
        XSSFCell cell = row.createCell(col);
        cell.setCellValue(value != null ? value : "");
        cell.setCellStyle(style);
    }

    private void writeValue(XSSFCell cell, Object val,
                             XSSFCellStyle def, XSSFCellStyle dateStyle) {
        if (val instanceof Double d)            { cell.setCellValue(d);               cell.setCellStyle(def);       return; }
        if (val instanceof Number n)            { cell.setCellValue(n.doubleValue());  cell.setCellStyle(def);       return; }
        if (val instanceof Boolean b)           { cell.setCellValue(b);               cell.setCellStyle(def);       return; }
        if (val instanceof LocalDateTime ldt)   { cell.setCellValue(ldt);             cell.setCellStyle(dateStyle); return; }
        if (val instanceof String s && !s.isBlank()) {
            String stripped = s.replaceAll("[€$£¥₺  \\s]", "");
            if (!stripped.isEmpty() && stripped.matches("[-+]?[\\d.,]+")) {
                cell.setCellValue(ConsolidationRow.parseFrenchDouble(s));
                cell.setCellStyle(def);
                return;
            }
            cell.setCellValue(s);
            cell.setCellStyle(def);
            return;
        }
        cell.setCellStyle(def);
    }

    private static String colLetter(int idx) {
        if (idx < 26) return String.valueOf((char) ('A' + idx));
        return String.valueOf((char) ('A' + idx / 26 - 1)) + (char) ('A' + idx % 26);
    }

    // -------------------------------------------------------------------------
    // XLSX style builders
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
            .replaceAll("\\p{InCombiningDiacriticalMarks}+", "")
            .toLowerCase();
    }

    private static String sanitize(String name) {
        return name.replaceAll("[\\\\/:*?\"<>|]", "_");
    }
}
