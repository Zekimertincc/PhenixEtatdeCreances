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
import org.apache.poi.ss.usermodel.HorizontalAlignment;
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
    private static final int OUT_COL_DEBITEUR  = 5;
    private static final int OUT_COL_CREANCE   = 6;
    private static final int OUT_COL_RECOUVRE  = 7;
    private static final int OUT_COL_ATTENTE   = 8;
    private static final int OUT_COL_CLOTURE   = 10;
    private static final int OUT_COL_ANCIENNETE = 3;
    private static final int OUT_COL_NREF       = 4;

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
                    if (!f.isFile()) return false;
                    String lower = f.getName().toLowerCase();
                    boolean isExcelOrPdf = lower.endsWith(".xlsx") || lower.endsWith(".xls") || lower.endsWith(".pdf");
                    boolean isEtatFile   = lower.startsWith("l_etat") || lower.startsWith("l'etat")
                            || lower.startsWith("letat") || lower.contains("etat de creances")
                            || lower.contains("etat des creances");
                    return isExcelOrPdf && isEtatFile;
                });
                if (oldFiles != null) {
                    for (File old : oldFiles) {
                        if (old.delete()) progress.accept(prog, "  Deleted old: " + old.getName());
                    }
                }

                String baseName  = AppConfig.ETAT_PUBLIC_FILENAME_PREFIX.replace("_", " ").trim()
                        + " " + sanitize(cf.companyName());
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
                // CLOTURE: blank source value → show "-" explicitly
                if (out[OUT_COL_CLOTURE] == null
                        || out[OUT_COL_CLOTURE].toString().isBlank()) {
                    out[OUT_COL_CLOTURE] = "-";
                }

                dataRows.add(out);
            }

            // Sort by NOMBRE (output col 0) descending
            dataRows.sort((a, b) -> {
                Object va = a[0];
                Object vb = b[0];
                if (va instanceof Number na && vb instanceof Number nb) {
                    return Double.compare(nb.doubleValue(), na.doubleValue()); // nb/na → descending
                }
                return String.valueOf(vb).compareTo(String.valueOf(va)); // ters
            });

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

            // Print setup: landscape A4, fit all columns on 1 page wide
            sheet.getPrintSetup().setPaperSize(PrintSetup.A4_PAPERSIZE);
            sheet.getPrintSetup().setLandscape(true);
            sheet.getPrintSetup().setFitWidth((short) 1);
            sheet.getPrintSetup().setFitHeight((short) 0); // 0 = unlimited height pages
            sheet.setAutobreaks(true);
            sheet.setFitToPage(true);
            // Margins in inches
            sheet.setMargin(Sheet.TopMargin,    0.5);
            sheet.setMargin(Sheet.BottomMargin, 0.5);
            sheet.setMargin(Sheet.LeftMargin,   0.5);
            sheet.setMargin(Sheet.RightMargin,  0.5);

            XSSFCellStyle plain  = buildPlainStyle(wb);
            XSSFCellStyle hdr    = buildHeaderStyle(wb);
            XSSFCellStyle data   = buildDataStyle(wb);
            XSSFCellStyle center = buildCenterStyle(wb, data);
            XSSFCellStyle money  = buildMoneyStyle(wb, data);
            XSSFCellStyle date   = buildDateStyle(wb, data);
            XSSFCellStyle total  = buildTotalStyle(wb);
            XSSFCellStyle yellow    = buildYellowStyle(wb, data);
            XSSFCellStyle yellowHdr = buildYellowHeaderStyle(wb);

            int rowIdx = 0;

            // Company info block (rows 0–7) — left side
            putString(sheet.createRow(rowIdx++), 0, companyName,  plain);
            putString(sheet.createRow(rowIdx++), 0, addr1,        plain);
            putString(sheet.createRow(rowIdx++), 0, addr2,        plain);
            sheet.createRow(rowIdx++);                             // blank
            putString(sheet.createRow(rowIdx++), 0, contactName,  plain);
            sheet.createRow(rowIdx++);                             // blank
            putString(sheet.createRow(rowIdx++), 0, codeClient,   plain);
            sheet.createRow(rowIdx++);                             // blank separator

            // Cabinet Phénix info — right side (cols 6-9, rows 0-4)
            try {
                java.io.InputStream logoStream = getClass().getResourceAsStream("/phenix.png");
                if (logoStream == null)
                    logoStream = getClass().getResourceAsStream("/com/zeki/merger/phenix_logo.png");
                if (logoStream != null) {
                    byte[] logoBytes = logoStream.readAllBytes();
                    logoStream.close();
                    int pictureIdx = wb.addPicture(logoBytes, XSSFWorkbook.PICTURE_TYPE_PNG);
                    org.apache.poi.ss.usermodel.Drawing<?> drawing = sheet.createDrawingPatriarch();
                    org.apache.poi.xssf.usermodel.XSSFClientAnchor anchor =
                            new org.apache.poi.xssf.usermodel.XSSFClientAnchor(
                                    740520, 0, 2867760, 102600, 3, 0, 5, 6);
                    drawing.createPicture(anchor, pictureIdx);
                }
            } catch (Exception ignored) {}

            XSSFCellStyle adresStyle = wb.createCellStyle();
            adresStyle.cloneStyleFrom(plain);
            adresStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.RIGHT);

            XSSFRow r0 = sheet.getRow(4); if (r0 == null) r0 = sheet.createRow(4);
            XSSFRow r1 = sheet.getRow(5); if (r1 == null) r1 = sheet.createRow(5);
            XSSFRow r2 = sheet.getRow(6); if (r2 == null) r2 = sheet.createRow(6);

            putString(r0, 6, "1, rue de Stockholm — 75008 PARIS", adresStyle);
            putString(r1, 6, "Tél. : +33 (0)1 53 20 12 76", adresStyle);
            putString(r2, 6, "contact@cabinetphenix.fr  |  www.cabinetphenix.fr", adresStyle);

            // Merge adres cells across cols 6-9
            sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(4, 4, 6, 9));
            sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(5, 5, 6, 9));
            sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(6, 6, 6, 9));

            // Header row (row index 8)
            XSSFRow headerRow = sheet.createRow(rowIdx++);
            for (int c = 0; c < OUT_COLS; c++) {
                XSSFCell cell = headerRow.createCell(c);
                cell.setCellValue(OUTPUT_HEADERS[c]);
                cell.setCellStyle((c == 1 || c == OUT_COL_NREF) ? yellowHdr : hdr);
            }

            // Data rows (start at row index 9)
            int dataStartRow = rowIdx;
            for (Object[] dr : dataRows) {
                XSSFRow row = sheet.createRow(rowIdx++);
                for (int c = 0; c < OUT_COLS; c++) {
                    boolean isMoneyCol   = (c == OUT_COL_CREANCE || c == OUT_COL_RECOUVRE
                            || c == OUT_COL_ATTENTE);
                    boolean isCenterCol  = (c == 0 || c == 1 || c == OUT_COL_ANCIENNETE || c == OUT_COL_NREF);
                    boolean isYellowCol  = (c == 1 || c == OUT_COL_NREF); // V/REF and N/REF
                    Object val = dr[c];
                    if (isMoneyCol && (val == null || val.toString().isBlank())) val = 0.0;
                    XSSFCellStyle colStyle = isYellowCol ? yellow
                            : isMoneyCol ? money
                            : isCenterCol ? center
                            : data;
                    writeValue(row.createCell(c), val, colStyle, date);
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

            // Header: 2-column table — client info left, Cabinet Phénix right
            Table headerTable = new Table(UnitValue.createPercentArray(new float[]{55, 45}))
                    .useAllAvailableWidth()
                    .setMarginBottom(8);

            // Left cell — client info
            com.itextpdf.layout.element.Cell leftCell =
                    new com.itextpdf.layout.element.Cell()
                    .setBorder(com.itextpdf.layout.borders.Border.NO_BORDER)
                    .setPadding(0);
            leftCell.add(new Paragraph(companyName).setBold().setFontSize(11));
            if (!addr1.isBlank())       leftCell.add(new Paragraph(addr1).setFontSize(9));
            if (!addr2.isBlank())       leftCell.add(new Paragraph(addr2).setFontSize(9));
            if (!contactName.isBlank()) leftCell.add(new Paragraph(contactName).setFontSize(9));
            if (!codeClient.isBlank())  leftCell.add(new Paragraph("Code client : " + codeClient).setFontSize(9));

            // Right cell — Cabinet Phénix logo + adres
            com.itextpdf.layout.element.Cell rightCell =
                    new com.itextpdf.layout.element.Cell()
                    .setBorder(com.itextpdf.layout.borders.Border.NO_BORDER)
                    .setPadding(0)
                    .setTextAlignment(com.itextpdf.layout.properties.TextAlignment.RIGHT);

            // Logo
            try {
                java.io.InputStream logoStream = getClass().getResourceAsStream("/phenix.png");
                if (logoStream == null)
                    logoStream = getClass().getResourceAsStream("/com/zeki/merger/phenix_logo.png");
                if (logoStream != null) {
                    com.itextpdf.io.image.ImageData imgData =
                            com.itextpdf.io.image.ImageDataFactory.create(logoStream.readAllBytes());
                    com.itextpdf.layout.element.Image logo =
                            new com.itextpdf.layout.element.Image(imgData)
                            .setWidth(120)
                            .setHorizontalAlignment(
                                    com.itextpdf.layout.properties.HorizontalAlignment.RIGHT);
                    rightCell.add(logo);
                    logoStream.close();
                }
            } catch (Exception ignored) {}

            rightCell.add(new Paragraph("1, rue de Stockholm — 75008 PARIS")
                    .setFontSize(8).setMarginTop(4));
            rightCell.add(new Paragraph("Tél. : +33 (0)1 53 20 12 76")
                    .setFontSize(8));
            rightCell.add(new Paragraph("contact@cabinetphenix.fr  |  www.cabinetphenix.fr")
                    .setFontSize(8));

            headerTable.addCell(leftCell);
            headerTable.addCell(rightCell);
            doc.add(headerTable);

            // Table — A4 landscape ~841 pt wide minus 40 pt margins = ~801 pt
            float[] colWidths = {50, 52, 68, 50, 60, 136, 90, 76, 110, 72, 37};
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
                    boolean centerCol = (c == 0 || c == OUT_COL_ANCIENNETE);
                    boolean rightCol  = (c == OUT_COL_CREANCE || c == OUT_COL_RECOUVRE || c == OUT_COL_ATTENTE);
                    com.itextpdf.layout.properties.TextAlignment ta = rightCol
                            ? com.itextpdf.layout.properties.TextAlignment.RIGHT
                            : centerCol
                            ? com.itextpdf.layout.properties.TextAlignment.CENTER
                            : com.itextpdf.layout.properties.TextAlignment.LEFT;
                    table.addCell(new com.itextpdf.layout.element.Cell()
                            .add(new Paragraph(fmtPdf(dr[c], c)).setFontSize(7.5f)
                                    .setTextAlignment(ta))
                            .setTextAlignment(ta)
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
                com.itextpdf.layout.properties.TextAlignment ta =
                        com.itextpdf.layout.properties.TextAlignment.LEFT;
                if      (c == OUT_COL_DEBITEUR) { text = "TOTAUX :"; }
                else if (c == OUT_COL_CREANCE)  { text = fmtMoney(totCreance);  ta = com.itextpdf.layout.properties.TextAlignment.RIGHT; }
                else if (c == OUT_COL_RECOUVRE) { text = fmtMoney(totRecouvre); ta = com.itextpdf.layout.properties.TextAlignment.RIGHT; }
                else if (c == OUT_COL_ATTENTE)  { text = fmtMoney(totAttente);  ta = com.itextpdf.layout.properties.TextAlignment.RIGHT; }
                table.addCell(new com.itextpdf.layout.element.Cell()
                        .add(new Paragraph(text).setBold().setFontSize(7.5f).setTextAlignment(ta))
                        .setBackgroundColor(yellow)
                        .setPadding(2));
            }

            doc.add(table);
        }
    }

    private String fmtPdf(Object val, int colIndex) {
        if (val == null) return "";
        // NOMBRE (0), V/REF (1), ANCIENNETE (3), N/REF (4) — integer, no decimals
        if (colIndex == 0 || colIndex == 1 || colIndex == OUT_COL_ANCIENNETE || colIndex == OUT_COL_NREF) {
            if (val instanceof Number n) return String.valueOf(n.intValue());
            return val.toString().trim();
        }
        // Money columns — millier format: 1 000,00
        if (colIndex == OUT_COL_CREANCE || colIndex == OUT_COL_RECOUVRE || colIndex == OUT_COL_ATTENTE) {
            if (val instanceof Number n) return fmtMoney(n.doubleValue());
        }
        if (val instanceof Double d)        return fmt2(d);
        if (val instanceof Number n)        return fmt2(n.doubleValue());
        if (val instanceof LocalDateTime t) return t.toLocalDate().format(DATE_FMT);
        if (val instanceof Boolean b)       return b ? "OUI" : "-";
        return val.toString().trim();
    }

    private static String fmtMoney(double v) {
        java.text.NumberFormat nf = java.text.NumberFormat.getNumberInstance(java.util.Locale.FRANCE);
        nf.setMinimumFractionDigits(2);
        nf.setMaximumFractionDigits(2);
        // Replace non-breaking space (used by Locale.FRANCE as thousands separator)
        // with regular space so iText renders it correctly
        return nf.format(v).replace('\u00A0', ' ').replace('\u202F', ' ');
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
            s = s.trim();
            String stripped = s
                    .replaceAll("[€$£¥₺]", "")
                    .replaceAll("\\p{Z}", "")
                    .trim();
            if (stripped.isEmpty() || stripped.equals("-")) {
                cell.setCellStyle(def);
                return;
            }
            if (stripped.matches("[-+]?[\\d.,]+")) {
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

    private XSSFCellStyle buildCenterStyle(XSSFWorkbook wb, XSSFCellStyle base) {
        XSSFCellStyle s = wb.createCellStyle();
        s.cloneStyleFrom(base);
        s.setAlignment(HorizontalAlignment.CENTER);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        return s;
    }

    private XSSFCellStyle buildMoneyStyle(XSSFWorkbook wb, XSSFCellStyle base) {
        XSSFCellStyle s = wb.createCellStyle();
        s.cloneStyleFrom(base);
        s.setDataFormat(wb.getCreationHelper().createDataFormat().getFormat("#,##0.00"));
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        return s;
    }

    private XSSFCellStyle buildDateStyle(XSSFWorkbook wb, XSSFCellStyle base) {
        XSSFCellStyle s = wb.createCellStyle();
        s.cloneStyleFrom(base);
        s.setDataFormat(wb.getCreationHelper().createDataFormat().getFormat("dd/MM/yyyy"));
        s.setVerticalAlignment(VerticalAlignment.CENTER);
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
        s.setDataFormat(wb.getCreationHelper().createDataFormat().getFormat("#,##0.00")); // ← bunu ekle
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
        return name.replaceAll("[\\\\/:*?\"<>|_]", " ").trim().replaceAll("\\s+", " ");
    }

    /** V/REF and N/REF data cells — yellow background (#FFFF00), same borders as data. */
    private XSSFCellStyle buildYellowStyle(XSSFWorkbook wb, XSSFCellStyle base) {
        XSSFCellStyle s = wb.createCellStyle();
        s.cloneStyleFrom(base);
        s.setFillForegroundColor(
            new XSSFColor(new byte[]{(byte)0xFF,(byte)0xFF,(byte)0x00}, null));
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setAlignment(HorizontalAlignment.CENTER);
        return s;
    }

    /** V/REF and N/REF header cells — yellow background, bold black font. */
    private XSSFCellStyle buildYellowHeaderStyle(XSSFWorkbook wb) {
        XSSFCellStyle s = wb.createCellStyle();
        XSSFFont f = wb.createFont();
        f.setBold(true);
        f.setFontHeightInPoints((short) 10);
        f.setColor(new XSSFColor(new byte[]{(byte)0x00,(byte)0x00,(byte)0x00}, null));
        s.setFont(f);
        s.setFillForegroundColor(
            new XSSFColor(new byte[]{(byte)0xFF,(byte)0xFF,(byte)0x00}, null));
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        s.setAlignment(HorizontalAlignment.CENTER);
        return s;
    }
}