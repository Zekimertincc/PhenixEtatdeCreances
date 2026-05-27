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

    public List<String> apply(File recupFile, File rootFolder, File tableauFile,
                              BiConsumer<Double, String> progress) throws Exception {
        List<String> log = new ArrayList<>();

        progress.accept(0.05, "Lecture " + recupFile.getName() + "...");
        Map<String, String> factureMap = readFactureMap(recupFile);
        progress.accept(0.10, factureMap.size() + " numéros de facture lus.");

        // Read soldes from Tableau de bord "Soldes" sheet
        Map<String, double[]> soldeMap = new LinkedHashMap<>(); // norm name → [solde, isNonComp]
        if (tableauFile != null && tableauFile.exists()) {
            soldeMap = readSoldeMap(tableauFile);
            progress.accept(0.13, soldeMap.size() + " soldes clients lus depuis Tableau de bord.");
        }

        List<FolderScanner.CompanyFile> companies = scanner.scan(rootFolder);
        if (companies.isEmpty()) {
            progress.accept(1.0, "Aucun dossier trouvé dans: " + rootFolder.getName());
            return log;
        }

        int total = companies.size();
        final Map<String, double[]> finalSoldeMap = soldeMap;
        for (int i = 0; i < total; i++) {
            FolderScanner.CompanyFile cf = companies.get(i);
            double prog = 0.15 + 0.85 * (i + 1.0) / total;

            String entry;
            try {
                entry = processCompany(cf.excelFile(), factureMap, finalSoldeMap);
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
    // Read Tableau de bord "Soldes" sheet
    // Returns: norm(clientName) → [solde, isNonComp (1.0 or 0.0)]
    // =========================================================================

    private Map<String, double[]> readSoldeMap(File tableauFile) throws IOException {
        Map<String, double[]> map = new LinkedHashMap<>();
        try (Workbook wb = openWorkbook(tableauFile)) {
            Sheet sheet = wb.getSheet("Soldes");
            if (sheet == null) return map;
            DataFormatter    fmt = new DataFormatter();
            FormulaEvaluator ev  = wb.getCreationHelper().createFormulaEvaluator();

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                String name = cellStr(row, 0, fmt, ev); // col A = CLIENT
                if (name.isBlank()) continue;
                double solde     = numericVal(row, 2);  // col C = solde
                double nonComp   = numericVal(row, 3);  // col D = -1 if non-comp
                if (solde > 0.005) {
                    map.put(DataReader.normalize(name), new double[]{solde, nonComp == -1.0 ? 1.0 : 0.0});
                }
            }
        }
        return map;
    }

    private double numericVal(Row row, int col) {
        Cell cell = row.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (cell == null) return 0.0;
        CellType ct = cell.getCellType() == CellType.FORMULA
                ? cell.getCachedFormulaResultType() : cell.getCellType();
        if (ct == CellType.NUMERIC) return cell.getNumericCellValue();
        return 0.0;
    }

    // =========================================================================
    // Process one company file
    // =========================================================================

    private String processCompany(File excelFile,
                                  Map<String, String> factureMap,
                                  Map<String, double[]> soldeMap) throws IOException {
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

            // Match in factureMap — exact first, then best partial
            String norm       = DataReader.normalize(nomClient);
            String numFacture = factureMap.get(norm);
            if (numFacture == null) {
                String bestKey = null;
                int bestLen = 0;
                for (Map.Entry<String, String> e : factureMap.entrySet()) {
                    String k = e.getKey();
                    if (norm.contains(k) || k.contains(norm)) {
                        int len = Math.min(norm.length(), k.length());
                        if (len > bestLen) { bestLen = len; bestKey = k; }
                    }
                }
                if (bestKey != null) numFacture = factureMap.get(bestKey);
            }
            if (numFacture == null) return "'" + nomClient + "' → eşleşme bulunamadı";

            // Sheet "Facture en préparation"
            Sheet facture = wb.getSheet("Facture en préparation");
            if (facture == null) facture = findSheetLike(wb, "facture");
            if (facture == null) return "'" + nomClient + "' → sheet 'Facture en préparation' yok";

            // Write facture number → D13 (row 12, col 3)
            Row row12 = facture.getRow(12);
            if (row12 == null) row12 = facture.createRow(12);
            Cell d13 = row12.getCell(3);
            if (d13 == null) d13 = row12.createCell(3);
            d13.setCellValue(numFacture);

            // Write solde client → I or J row (×-1) if client has outstanding balance
            String soldeInfo = "";
            double[] soldeEntry = findSolde(norm, soldeMap);
            if (soldeEntry != null) {
                double solde = soldeEntry[0];
                double negSolde = -solde;

                // COMP ve NON COMP için her ikisi de J satırına yaz
                // I satırı formül içeriyor (C30-C40), ezilmemeli
                // J satırı = "factures en retard" = önceki borç → buraya yazılır
                String targetMarker = "J";
                int targetRow = -1;
                for (int r = 0; r <= Math.min(facture.getLastRowNum(), 200); r++) {
                    Row fr = facture.getRow(r);
                    if (fr == null) continue;
                    Cell c0 = fr.getCell(0, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    if (c0 == null) continue;
                    String val = fmt.formatCellValue(c0, ev).trim();
                    // Match "I", "I=A-H", "J", "J=...", etc.
                    if (val.equals(targetMarker) || val.startsWith(targetMarker + "=")
                            || val.startsWith(targetMarker + " ")) {
                        targetRow = r;
                        break;
                    }
                }

                if (targetRow >= 0) {
                    Row fr = facture.getRow(targetRow);
                    if (fr == null) fr = facture.createRow(targetRow);
                    Cell valCell = fr.getCell(2, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    if (valCell == null) valCell = fr.createCell(2);
                    valCell.setCellValue(solde); // pozitif yaz — J formülü C45>0 → "factures impayées"
                    soldeInfo = String.format(" | Solde → J=+%.2f", solde);
                }
            }

            try (FileOutputStream fos = new FileOutputStream(excelFile)) {
                wb.write(fos);
            }
            return "D13 = " + numFacture + soldeInfo;
        }
    }

    private double[] findSolde(String norm, Map<String, double[]> soldeMap) {
        double[] v = soldeMap.get(norm);
        if (v != null) return v;
        for (Map.Entry<String, double[]> e : soldeMap.entrySet()) {
            String k = e.getKey();
            if (norm.contains(k) || k.contains(norm)) return e.getValue();
        }
        return null;
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