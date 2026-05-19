package com.zeki.merger.service;

import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.properties.UnitValue;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.function.BiConsumer;

/**
 * Reads "Facture en préparation" sheet from each company Excel file
 * and exports it as PDF in the same directory.
 * Equivalent to legacy VBA: Enregistrement_Facture_Format_Pdf
 */
public class FacturePdfService {

    private final FolderScanner scanner = new FolderScanner();

    public List<String> apply(File rootFolder, BiConsumer<Double, String> progress) throws Exception {
        List<String> log = new ArrayList<>();
        List<FolderScanner.CompanyFile> companies = scanner.scan(rootFolder);
        if (companies.isEmpty()) {
            progress.accept(1.0, "Aucun dossier trouvé.");
            return log;
        }
        int total = companies.size();
        for (int i = 0; i < total; i++) {
            FolderScanner.CompanyFile cf = companies.get(i);
            double prog = 0.05 + 0.95 * (i + 1.0) / total;
            String result;
            try {
                result = processCompany(cf.excelFile());
            } catch (Exception e) {
                result = "ERREUR: " + e.getMessage();
            }
            log.add(cf.companyName() + " → " + result);
            progress.accept(prog, "[" + (i+1) + "/" + total + "] " + cf.companyName() + " → " + result);
        }
        progress.accept(1.0, "Génération PDF terminée (" + total + " dossiers).");
        return log;
    }

    private String processCompany(File excelFile) throws Exception {
        try (Workbook wb = openWorkbook(excelFile)) {
            Sheet sheet = wb.getSheet("Facture en préparation");
            if (sheet == null) {
                for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                    if (wb.getSheetName(i).toLowerCase().contains("facture")) {
                        sheet = wb.getSheetAt(i); break;
                    }
                }
            }
            if (sheet == null) return "sheet 'Facture en préparation' introuvable";

            String pdfName = excelFile.getName().replaceAll("\\.(xlsx?|xls)$", "") + "_facture.pdf";
            File pdfFile = new File(excelFile.getParent(), pdfName);
            exportSheetToPdf(sheet, wb, pdfFile);
            return "PDF → " + pdfName;
        }
    }

    private void exportSheetToPdf(Sheet sheet, Workbook wb, File pdfFile) throws Exception {
        DataFormatter fmt = new DataFormatter();
        FormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator();

        int firstRow = sheet.getFirstRowNum();
        int lastRow  = sheet.getLastRowNum();
        int maxCols  = 0;
        for (int r = firstRow; r <= lastRow; r++) {
            Row row = sheet.getRow(r);
            if (row != null && row.getLastCellNum() > maxCols)
                maxCols = row.getLastCellNum();
        }
        if (maxCols == 0) maxCols = 7;

        try (PdfWriter writer = new PdfWriter(pdfFile);
             PdfDocument pdf = new PdfDocument(writer);
             Document doc = new Document(pdf)) {

            doc.setMargins(20, 20, 20, 20);
            com.itextpdf.layout.element.Table table =
                new com.itextpdf.layout.element.Table(UnitValue.createPercentArray(maxCols))
                    .useAllAvailableWidth();

            for (int r = firstRow; r <= lastRow; r++) {
                Row row = sheet.getRow(r);
                for (int c = 0; c < maxCols; c++) {
                    String val = "";
                    if (row != null) {
                        org.apache.poi.ss.usermodel.Cell cell =
                            row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                        if (cell != null) {
                            try { val = fmt.formatCellValue(cell, ev); } catch (Exception ignored) {}
                        }
                    }
                    table.addCell(new com.itextpdf.layout.element.Cell()
                        .add(new Paragraph(val).setFontSize(8))
                        .setPadding(2));
                }
            }
            doc.add(table);
        }
    }

    private Workbook openWorkbook(File file) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        return file.getName().toLowerCase().endsWith(".xls")
            ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
    }
}
