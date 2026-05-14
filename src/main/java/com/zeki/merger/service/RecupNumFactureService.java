package com.zeki.merger.service;

import com.zeki.merger.trf.DataReader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;
import java.util.*;
import java.util.function.BiConsumer;

/**
 * Reads client → N° facture mapping from RecupNumFacture.xlsx,
 * then writes each facture number to D13 of the "Facture en préparation" sheet
 * in each company's Créances Excel file.
 */
public class RecupNumFactureService {

    private final FolderScanner scanner = new FolderScanner();

    // =========================================================================
    // Public entry point
    // =========================================================================

    public List<String> apply(File recupFile, File rootFolder,
                               BiConsumer<Double, String> progress) throws Exception {
        List<String> log = new ArrayList<>();

        progress.accept(0.05, "Lecture " + recupFile.getName() + "...");
        Map<String, String> factureMap = readFactureMap(recupFile);
        progress.accept(0.10, factureMap.size() + " numéros de facture lus.");

        List<FolderScanner.CompanyFile> companies = scanner.scan(rootFolder);
        if (companies.isEmpty()) {
            progress.accept(1.0, "Aucun dossier trouvé dans: " + rootFolder.getName());
            return log;
        }

        int total = companies.size();
        for (int i = 0; i < total; i++) {
            FolderScanner.CompanyFile cf = companies.get(i);
            double prog = 0.10 + 0.90 * (i + 1.0) / total;

            String entry;
            try {
                entry = processCompany(cf.excelFile(), factureMap);
            } catch (Exception e) {
                entry = "ERREUR: " + e.getMessage();
            }

            log.add(cf.companyName() + " → " + entry);
            progress.accept(prog, "[" + (i + 1) + "/" + total + "] "
                + cf.companyName() + " → " + entry);
        }

        progress.accept(1.0, "Récupération numéros de facture terminée (" + total + " dossiers).");
        return log;
    }

    // =========================================================================
    // Read RecupNumFacture.xlsx
    // =========================================================================

    private Map<String, String> readFactureMap(File file) throws IOException {
        Map<String, String> map = new LinkedHashMap<>();
        try (Workbook wb = openWorkbook(file)) {
            Sheet sheet = wb.getSheet("Feuil1");
            if (sheet == null) sheet = wb.getSheetAt(0);
            DataFormatter    fmt = new DataFormatter();
            FormulaEvaluator ev  = wb.getCreationHelper().createFormulaEvaluator();

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                String name = cellStr(row, 0, fmt, ev); // col A = CLIENT
                if (name.isBlank()) break;              // stop at first blank row
                String numFacture = cellStr(row, 1, fmt, ev); // col B = N° facture
                if (!numFacture.isBlank()) {
                    map.put(DataReader.normalize(name), numFacture);
                }
            }
        }
        return map;
    }

    // =========================================================================
    // Process one company file
    // =========================================================================

    private String processCompany(File excelFile,
                                   Map<String, String> factureMap) throws IOException {
        // Load into memory so we can write back to the same file
        byte[] bytes = Files.readAllBytes(excelFile.toPath());

        try (Workbook wb = openWorkbookFromBytes(bytes, excelFile.getName())) {
            DataFormatter    fmt = new DataFormatter();
            FormulaEvaluator ev  = wb.getCreationHelper().createFormulaEvaluator();

            // Sheet "Créances" → H4 (row 3, col 7)
            Sheet creances = wb.getSheet("Créances");
            if (creances == null) creances = findSheetLike(wb, "creance");
            if (creances == null) return "sheet 'Créances' introuvable";

            String nomClient = sheetCell(creances, 3, 7, fmt, ev); // H4
            if (nomClient.isBlank()) return "H4 vide";

            // Match in factureMap (exact, then partial)
            String norm       = DataReader.normalize(nomClient);
            String numFacture = factureMap.get(norm);
            if (numFacture == null) {
                for (Map.Entry<String, String> e : factureMap.entrySet()) {
                    if (norm.contains(e.getKey()) || e.getKey().contains(norm)) {
                        numFacture = e.getValue();
                        break;
                    }
                }
            }
            if (numFacture == null) return "'" + nomClient + "' → eşleşme bulunamadı";

            // Sheet "Facture en préparation" → D13 (row 12, col 3)
            Sheet facture = wb.getSheet("Facture en préparation");
            if (facture == null) facture = findSheetLike(wb, "facture");
            if (facture == null) return "'" + nomClient + "' → sheet 'Facture en préparation' yok";

            Row row12 = facture.getRow(12);
            if (row12 == null) row12 = facture.createRow(12);
            Cell d13 = row12.getCell(3);
            if (d13 == null) d13 = row12.createCell(3);
            d13.setCellValue(numFacture);

            try (FileOutputStream fos = new FileOutputStream(excelFile)) {
                wb.write(fos);
            }
            return "D13 = " + numFacture;
        }
    }

    // =========================================================================
    // Helpers
    // =========================================================================

    private String sheetCell(Sheet sheet, int r, int c,
                               DataFormatter fmt, FormulaEvaluator ev) {
        Row row = sheet.getRow(r);
        if (row == null) return "";
        Cell cell = row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (cell == null) return "";
        return fmt.formatCellValue(cell, ev).trim();
    }

    private String cellStr(Row row, int col, DataFormatter fmt, FormulaEvaluator ev) {
        Cell cell = row.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (cell == null) return "";
        return fmt.formatCellValue(cell, ev).trim();
    }

    private Sheet findSheetLike(Workbook wb, String keyword) {
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            if (DataReader.normalize(wb.getSheetName(i)).contains(keyword)) {
                return wb.getSheetAt(i);
            }
        }
        return null;
    }

    private Workbook openWorkbook(File file) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        return file.getName().toLowerCase().endsWith(".xls")
            ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
    }

    private Workbook openWorkbookFromBytes(byte[] bytes, String fileName) throws IOException {
        ByteArrayInputStream bis = new ByteArrayInputStream(bytes);
        return fileName.toLowerCase().endsWith(".xls")
            ? new HSSFWorkbook(bis) : new XSSFWorkbook(bis);
    }
}
