package com.zeki.merger.service;

import com.zeki.merger.trf.DataReader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;
import java.time.LocalDate;
import java.util.*;
import java.util.function.BiConsumer;

/**
 * Mis à jour Listing Client — Dernier dossier arrivé
 *
 * For each company folder:
 *   1. Opens the company's Etat de Créances Excel
 *   2. Reads all dates from Créances sheet col C (REMIS LE) — row 17 onwards
 *   3. Finds the MAX date
 *   4. Writes it to the Listing Cabinet (col 19 = "Dossier remis") for that client
 *
 * Listing columns (0-based):
 *   2 = NOM CLIENT, 3 = CODE CLIENT, 19 = Dossier remis (Dernier dossier arrivé)
 */
public class MisAJourListingService {

    private static final int L_NAME    = 2;
    private static final int L_CODE    = 3;
    private static final int L_DERNIER = 19; // "Dossier remis" = dernier dossier arrivé

    private final FolderScanner scanner = new FolderScanner();

    // =========================================================================
    // Entry point
    // =========================================================================

    public List<String> apply(File listingFile, File rootFolder,
                              BiConsumer<Double, String> progress) throws Exception {
        List<String> log = new ArrayList<>();

        // 1. Load listing into memory
        progress.accept(0.05, "Lecture " + listingFile.getName() + "...");
        byte[] listingBytes = Files.readAllBytes(listingFile.toPath());

        // Build name → row index map from listing
        Map<String, Integer> nameToRow = new LinkedHashMap<>();
        Map<String, Integer> codeToRow = new LinkedHashMap<>();
        try (Workbook wb = openWorkbookFromBytes(listingBytes, listingFile.getName())) {
            Sheet sheet = wb.getSheet("Feuil1");
            if (sheet == null) sheet = wb.getSheetAt(0);
            DataFormatter fmt = new DataFormatter();
            FormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator();
            for (int r = 2; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                String name = cellStr(row, L_NAME, fmt, ev);
                String code = cellStr(row, L_CODE, fmt, ev);
                if (!name.isBlank()) nameToRow.put(DataReader.normalize(name), r);
                if (!code.isBlank()) codeToRow.put(DataReader.normalize(code), r);
            }
        }
        progress.accept(0.10, nameToRow.size() + " clients dans le Listing.");

        // 2. Scan company folders
        List<FolderScanner.CompanyFile> companies = scanner.scan(rootFolder);
        if (companies.isEmpty()) {
            progress.accept(1.0, "Aucun dossier trouvé dans: " + rootFolder.getName());
            return log;
        }

        // 3. For each company, find max date in Créances col C
        Map<Integer, LocalDate> updates = new LinkedHashMap<>(); // listingRow → maxDate
        int total = companies.size();
        for (int i = 0; i < total; i++) {
            FolderScanner.CompanyFile cf = companies.get(i);
            double prog = 0.10 + 0.60 * (i + 1.0) / total;

            try {
                LocalDate maxDate = readMaxDate(cf.excelFile());
                if (maxDate == null) {
                    log.add(cf.companyName() + " → aucune date trouvée");
                    progress.accept(prog, "[" + (i+1) + "/" + total + "] " + cf.companyName() + " → aucune date");
                    continue;
                }

                // Match in listing
                Integer listingRow = findListingRow(cf.excelFile(), nameToRow, codeToRow);
                if (listingRow == null) {
                    log.add(cf.companyName() + " → " + maxDate + " (non trouvé dans Listing)");
                    progress.accept(prog, "[" + (i+1) + "/" + total + "] " + cf.companyName() + " → non trouvé Listing");
                    continue;
                }

                updates.put(listingRow, maxDate);
                log.add(cf.companyName() + " → " + maxDate + " → Listing row " + (listingRow + 1));
                progress.accept(prog, "[" + (i+1) + "/" + total + "] " + cf.companyName() + " → " + maxDate);

            } catch (Exception e) {
                log.add(cf.companyName() + " → ERREUR: " + e.getMessage());
            }
        }

        // 4. Write all updates to Listing in one pass
        if (!updates.isEmpty()) {
            progress.accept(0.75, "Mise à jour du Listing (" + updates.size() + " clients)...");
            writeToListing(listingFile, listingBytes, updates);
            progress.accept(1.0, "Listing mis à jour — " + updates.size() + " dernier(s) dossier(s) écrits.");
        } else {
            progress.accept(1.0, "Aucune mise à jour effectuée.");
        }

        return log;
    }

    // =========================================================================
    // Read max date from Créances col C (REMIS LE)
    // =========================================================================

    private LocalDate readMaxDate(File excelFile) throws IOException {
        try (Workbook wb = openWorkbook(excelFile)) {
            Sheet creances = wb.getSheet("Créances");
            if (creances == null) {
                for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                    if (DataReader.normalize(wb.getSheetName(i)).contains("creance")) {
                        creances = wb.getSheetAt(i); break;
                    }
                }
            }
            if (creances == null) return null;

            LocalDate maxDate = null;
            // Data starts at row 17 (index 16), col C = index 2
            for (int r = 16; r <= creances.getLastRowNum(); r++) {
                Row row = creances.getRow(r);
                if (row == null) continue;
                Cell cell = row.getCell(2, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (cell == null) continue;

                CellType ct = cell.getCellType() == CellType.FORMULA
                        ? cell.getCachedFormulaResultType() : cell.getCellType();
                if (ct != CellType.NUMERIC) continue;
                if (!DateUtil.isCellDateFormatted(cell)) continue;

                try {
                    LocalDate d = cell.getLocalDateTimeCellValue().toLocalDate();
                    if (maxDate == null || d.isAfter(maxDate)) maxDate = d;
                } catch (Exception ignored) {}
            }
            return maxDate;
        }
    }

    // =========================================================================
    // Find listing row for a company
    // =========================================================================

    private Integer findListingRow(File excelFile,
                                   Map<String, Integer> nameToRow,
                                   Map<String, Integer> codeToRow) throws IOException {
        try (Workbook wb = openWorkbook(excelFile)) {
            DataFormatter fmt = new DataFormatter();
            FormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator();
            Sheet creances = wb.getSheet("Créances");
            if (creances == null) return null;

            // Try code from A13
            String a13 = sheetCell(creances, 12, 0, fmt, ev);
            String codeKey = a13.length() >= 6
                    ? DataReader.normalize(a13.substring(a13.length() - 6))
                    : DataReader.normalize(a13);
            Integer row = codeToRow.get(codeKey);
            if (row != null) return row;

            // Try name from H4
            String name = sheetCell(creances, 3, 7, fmt, ev);
            if (!name.isBlank()) {
                String normName = DataReader.normalize(name);
                row = nameToRow.get(normName);
                if (row != null) return row;
                // Partial match
                for (Map.Entry<String, Integer> e : nameToRow.entrySet()) {
                    if (normName.contains(e.getKey()) || e.getKey().contains(normName))
                        return e.getValue();
                }
            }
            return null;
        }
    }

    // =========================================================================
    // Write dates to listing file
    // =========================================================================

    private void writeToListing(File listingFile, byte[] originalBytes,
                                Map<Integer, LocalDate> updates) throws IOException {
        try (Workbook wb = openWorkbookFromBytes(originalBytes, listingFile.getName())) {
            Sheet sheet = wb.getSheet("Feuil1");
            if (sheet == null) sheet = wb.getSheetAt(0);

            // Date cell style
            CellStyle dateStyle = wb.createCellStyle();
            DataFormat df = wb.createDataFormat();
            dateStyle.setDataFormat(df.getFormat("dd/MM/yyyy"));

            for (Map.Entry<Integer, LocalDate> entry : updates.entrySet()) {
                int rowIdx = entry.getKey();
                LocalDate date = entry.getValue();
                Row row = sheet.getRow(rowIdx);
                if (row == null) row = sheet.createRow(rowIdx);
                Cell cell = row.getCell(L_DERNIER);
                if (cell == null) cell = row.createCell(L_DERNIER);
                cell.setCellValue(java.util.Date.from(
                        date.atStartOfDay(java.time.ZoneId.systemDefault()).toInstant()));
                cell.setCellStyle(dateStyle);
            }

            try (FileOutputStream fos = new FileOutputStream(listingFile)) {
                wb.write(fos);
            }
        }
    }

    // =========================================================================
    // Helpers
    // =========================================================================

    private String sheetCell(Sheet sheet, int r, int c, DataFormatter fmt, FormulaEvaluator ev) {
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