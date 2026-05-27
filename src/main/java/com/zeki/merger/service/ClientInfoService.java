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
 * Reads client info from the Listing file (getTrfListing()), then writes
 * an "Infos" sheet to each company's Créances Excel file.
 *
 * Listing columns (0-based, rows from index 2):
 *   2=name, 3=code, 5=adresse, 6=CP, 7=ville, 8=mail, 9=tel, 21=IBAN, 22=BIC, 24=commercial
 *
 * Company Excel: A13 (row12,col0) last-6-chars = client code (H4 name fallback).
 * Output sheet "Infos": 10 rows A:B, text style for CP/IBAN/BIC/Téléphone.
 */
public class ClientInfoService {

    private static final int L_NAME       = 2;
    private static final int L_CODE       = 3;
    private static final int L_ADRESSE    = 5;
    private static final int L_CP         = 6;
    private static final int L_VILLE      = 7;
    private static final int L_MAIL       = 8;
    private static final int L_TEL        = 9;
    private static final int L_IBAN       = 21;
    private static final int L_BIC        = 22;
    private static final int L_COMMERCIAL = 24;
    private static final int L_TVA        = 26; // N° TVA Intracommunautaire (col 27, 0-based=26)

    private static final String INFOS_SHEET = "Infos";

    private final FolderScanner scanner = new FolderScanner();

    // =========================================================================
    // Public entry point
    // =========================================================================

    public List<String> apply(File listingFile, File rootFolder,
                              BiConsumer<Double, String> progress) throws Exception {
        return apply(listingFile, rootFolder, null, progress);
    }

    public List<String> apply(File listingFile, File rootFolder, File procreancesFile,
                              BiConsumer<Double, String> progress) throws Exception {
        List<String> log = new ArrayList<>();

        progress.accept(0.05, "Lecture " + listingFile.getName() + "...");
        Map<String, ClientData> byCode = new LinkedHashMap<>();
        Map<String, ClientData> byName = new LinkedHashMap<>();
        readListing(listingFile, byCode, byName);
        progress.accept(0.10, byCode.size() + " clients lus depuis le Listing.");

        // Read TVA from Procreances CSV if provided
        Map<String, String> tvaByCode = new LinkedHashMap<>(); // norm(code) → TVA
        Map<String, String> tvaByName = new LinkedHashMap<>(); // norm(name) → TVA
        if (procreancesFile != null && procreancesFile.exists()) {
            readTvaFromCsv(procreancesFile, tvaByCode, tvaByName);
            progress.accept(0.13, tvaByCode.size() + " N° TVA lus depuis Procréances.");
        }

        List<FolderScanner.CompanyFile> companies = scanner.scan(rootFolder);
        if (companies.isEmpty()) {
            progress.accept(1.0, "Aucun dossier trouvé dans: " + rootFolder.getName());
            return log;
        }

        int total = companies.size();
        for (int i = 0; i < total; i++) {
            FolderScanner.CompanyFile cf = companies.get(i);
            double prog = 0.15 + 0.85 * (i + 1.0) / total;

            String entry;
            try {
                entry = processCompany(cf.excelFile(), byCode, byName, tvaByCode, tvaByName);
            } catch (Exception e) {
                entry = "ERREUR: " + e.getMessage();
            }

            log.add(cf.companyName() + " → " + entry);
            progress.accept(prog, "[" + (i + 1) + "/" + total + "] "
                    + cf.companyName() + " → " + entry);
        }

        progress.accept(1.0, "Info Clients terminée (" + total + " dossiers).");
        return log;
    }

    // =========================================================================
    // Read Listing
    // =========================================================================

    private void readListing(File file,
                             Map<String, ClientData> byCode,
                             Map<String, ClientData> byName) throws IOException {
        try (Workbook wb = openWorkbook(file)) {
            Sheet sheet = wb.getSheet("Feuil1");
            if (sheet == null) sheet = wb.getSheetAt(0);
            DataFormatter    fmt = new DataFormatter();
            FormulaEvaluator ev  = wb.getCreationHelper().createFormulaEvaluator();

            for (int r = 2; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                String name = cellStr(row, L_NAME, fmt, ev);
                if (name.isBlank()) continue;
                String code = cellStr(row, L_CODE, fmt, ev);

                ClientData cd = new ClientData(
                        name,
                        code,
                        cellStr(row, L_ADRESSE,    fmt, ev),
                        cellStr(row, L_CP,         fmt, ev),
                        cellStr(row, L_VILLE,      fmt, ev),
                        cellStr(row, L_MAIL,       fmt, ev),
                        cellStr(row, L_TEL,        fmt, ev),
                        cellStr(row, L_IBAN,       fmt, ev),
                        cellStr(row, L_BIC,        fmt, ev),
                        cellStr(row, L_COMMERCIAL, fmt, ev),
                        cellStr(row, L_TVA,        fmt, ev)
                );

                if (!code.isBlank()) byCode.put(DataReader.normalize(code), cd);
                byName.put(DataReader.normalize(name), cd);
            }
        }
    }

    // =========================================================================
    // Read TVA from Procreances CSV
    // CSV: col 0 = code client, col 1 = nom, col 5 = N° TVA intracommunautaire
    // =========================================================================

    private void readTvaFromCsv(File csvFile,
                                Map<String, String> byCode,
                                Map<String, String> byName) {
        try (java.io.BufferedReader br = new java.io.BufferedReader(
                new java.io.InputStreamReader(new FileInputStream(csvFile),
                        java.nio.charset.StandardCharsets.UTF_8))) {
            String line;
            boolean first = true;
            while ((line = br.readLine()) != null) {
                if (first) { first = false; continue; } // skip header
                String[] cols = line.split(";", -1);
                if (cols.length < 6) continue;
                String code = cols[0].trim();
                String name = cols[1].trim();
                String tva  = cols[5].trim();
                if (tva.isBlank()) continue;
                if (!code.isBlank()) byCode.put(DataReader.normalize(code), tva);
                if (!name.isBlank()) byName.put(DataReader.normalize(name), tva);
            }
        } catch (Exception e) {
            System.err.println("[ClientInfoService] TVA CSV read error: " + e.getMessage());
        }
    }

    // =========================================================================
    // Process one company file
    // =========================================================================

    private String processCompany(File excelFile,
                                  Map<String, ClientData> byCode,
                                  Map<String, ClientData> byName,
                                  Map<String, String> tvaByCode,
                                  Map<String, String> tvaByName) throws IOException {
        byte[] bytes = Files.readAllBytes(excelFile.toPath());

        try (Workbook wb = openWorkbookFromBytes(bytes, excelFile.getName())) {
            DataFormatter    fmt = new DataFormatter();
            FormulaEvaluator ev  = wb.getCreationHelper().createFormulaEvaluator();

            Sheet creances = wb.getSheet("Créances");
            if (creances == null) creances = findSheetLike(wb, "creance");
            if (creances == null) return "sheet 'Créances' introuvable";

            String a13Raw = sheetCell(creances, 12, 0, fmt, ev);
            String codeKey = a13Raw.length() >= 6
                    ? DataReader.normalize(a13Raw.substring(a13Raw.length() - 6))
                    : DataReader.normalize(a13Raw);

            ClientData cd = byCode.get(codeKey);
            if (cd == null) {
                String nomClient = sheetCell(creances, 3, 7, fmt, ev);
                if (!nomClient.isBlank()) {
                    String normName = DataReader.normalize(nomClient);
                    cd = byName.get(normName);
                    if (cd == null) {
                        for (Map.Entry<String, ClientData> e : byName.entrySet()) {
                            if (normName.contains(e.getKey()) || e.getKey().contains(normName)) {
                                cd = e.getValue(); break;
                            }
                        }
                    }
                }
            }
            if (cd == null) return "'" + a13Raw + "' → aucune correspondance trouvée";

            // Find TVA — first from Listing, then from Procreances CSV
            String tva = (cd.tva != null && !cd.tva.isBlank()) ? cd.tva : null;
            if (tva == null) tva = tvaByCode.get(codeKey);
            if (tva == null) tva = tvaByName.get(DataReader.normalize(cd.name));
            if (tva == null) {
                String normName = DataReader.normalize(cd.name);
                for (Map.Entry<String, String> e : tvaByName.entrySet()) {
                    if (normName.contains(e.getKey()) || e.getKey().contains(normName)) {
                        tva = e.getValue(); break;
                    }
                }
            }

            // Rebuild Infos sheet
            int existingIdx = wb.getSheetIndex(INFOS_SHEET);
            if (existingIdx >= 0) wb.removeSheetAt(existingIdx);
            Sheet infos = wb.createSheet(INFOS_SHEET);

            CellStyle textStyle = wb.createCellStyle();
            DataFormat df = wb.createDataFormat();
            textStyle.setDataFormat(df.getFormat("@"));

            writeInfoRow(infos, 0, "Infos clients", cd.name,       null,      wb);
            writeInfoRow(infos, 1, "Adresse",        cd.adresse,    null,      wb);
            writeInfoRow(infos, 2, "CP",             cd.cp,         textStyle, wb);
            writeInfoRow(infos, 3, "Ville",          cd.ville,      null,      wb);
            writeInfoRow(infos, 4, "Mails",          cd.mail,       null,      wb);
            writeInfoRow(infos, 5, "Conditions",     "",            null,      wb);
            writeInfoRow(infos, 6, "Iban",           cd.iban,       textStyle, wb);
            writeInfoRow(infos, 7, "Bic",            cd.bic,        textStyle, wb);
            writeInfoRow(infos, 8, "Téléphone",      cd.tel,        textStyle, wb);
            writeInfoRow(infos, 9, "Commercial",     cd.commercial, null,      wb);
            // TVA — only if found
            if (tva != null && !tva.isBlank()) {
                writeInfoRow(infos, 10, "N° TVA", tva, textStyle, wb);
            }

            try (FileOutputStream fos = new FileOutputStream(excelFile)) {
                wb.write(fos);
            }
            String tvaInfo = (tva != null && !tva.isBlank()) ? " | TVA=" + tva : " | TVA non trouvée";
            return "Infos créée → " + cd.name + tvaInfo;
        }
    }

    // =========================================================================
    // Helpers
    // =========================================================================

    private void writeInfoRow(Sheet sheet, int rowIdx, String label, String value,
                              CellStyle valueStyle, Workbook wb) {
        Row row = sheet.createRow(rowIdx);
        row.createCell(0).setCellValue(label);
        Cell valCell = row.createCell(1);
        valCell.setCellValue(value != null ? value : "");
        if (valueStyle != null) valCell.setCellStyle(valueStyle);
    }

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

    // =========================================================================
    // Data model
    // =========================================================================

    private record ClientData(
            String name, String code, String adresse, String cp, String ville,
            String mail, String tel, String iban, String bic, String commercial, String tva
    ) {}
}