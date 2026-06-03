package com.zeki.merger.service;

import com.zeki.merger.trf.DataReader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;
import java.util.*;
import java.util.function.BiConsumer;

public class ClientInfoService {

    private static final String INFOS_SHEET = "Infos";
    private final FolderScanner scanner = new FolderScanner();

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

        Map<String, String> tvaByCode = new LinkedHashMap<>();
        Map<String, String> tvaByName = new LinkedHashMap<>();
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
    // Listing okuma — rawStr ile, DataFormatter YOK
    // Kolonlar (0-bazlı): 2=name, 3=code, 5=adresse, 6=cp, 7=ville,
    //                     8=mail, 9=tel, 21=iban, 22=bic, 23=libelle, 24=tva
    // =========================================================================

    private void readListing(File file,
                             Map<String, ClientData> byCode,
                             Map<String, ClientData> byName) throws IOException {
        try (Workbook wb = openWorkbook(file)) {
            Sheet sheet = wb.getSheet("Feuil1");
            if (sheet == null) sheet = wb.getSheetAt(0);

            for (int r = 2; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                String name = rawStr(row, 2);
                if (name.isBlank()) continue;

                ClientData cd = new ClientData(
                        name,
                        rawStr(row, 3),   // code
                        rawStr(row, 5),   // adresse
                        rawStr(row, 6),   // cp
                        rawStr(row, 7),   // ville
                        rawStr(row, 8),   // mail
                        rawStr(row, 9),   // tel
                        rawStr(row, 21),  // iban
                        rawStr(row, 22),  // bic
                        rawStr(row, 23),  // libelle (commercial)
                        rawStr(row, 24)   // tva — col Y (index 24)
                );

                String code = cd.code();
                if (!code.isBlank()) byCode.put(DataReader.normalize(code), cd);
                byName.put(DataReader.normalize(name), cd);
            }
        }
    }

    /**
     * Hücreyi ham string olarak okur.
     * DataFormatter KULLANILMAZ — numeric hücreleri boş döndürdüğü için.
     */
    private String rawStr(Row row, int colIdx) {
        if (row == null) return "";
        Cell cell = row.getCell(colIdx, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (cell == null) return "";
        CellType type = cell.getCellType();
        if (type == CellType.FORMULA) type = cell.getCachedFormulaResultType();
        switch (type) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                double v = cell.getNumericCellValue();
                if (v == Math.floor(v) && !Double.isInfinite(v))
                    return String.valueOf((long) v);
                return String.valueOf(v);
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return "";
        }
    }

    // =========================================================================
    // TVA — Procreances CSV
    // =========================================================================

    private void readTvaFromCsv(File csvFile,
                                Map<String, String> byCode,
                                Map<String, String> byName) {
        try (BufferedReader br = new BufferedReader(
                new InputStreamReader(new FileInputStream(csvFile),
                        java.nio.charset.StandardCharsets.UTF_8))) {
            String line;
            boolean first = true;
            while ((line = br.readLine()) != null) {
                if (first) { first = false; continue; }
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
    // Şirket dosyasını işle
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

            // Client kodu: A13'ün son 6 karakteri
            String a13Raw = sheetCell(creances, 12, 0, fmt, ev);
            String codeKey = a13Raw.length() >= 6
                    ? DataReader.normalize(a13Raw.substring(a13Raw.length() - 6))
                    : DataReader.normalize(a13Raw);

            ClientData cd = byCode.get(codeKey);
            if (cd == null) {
                // Fallback: H4'ten isim ile tam eşleşme
                String nomClient = sheetCell(creances, 3, 7, fmt, ev);
                if (!nomClient.isBlank())
                    cd = byName.get(DataReader.normalize(nomClient));
            }
            if (cd == null) return "'" + a13Raw + "' → aucune correspondance trouvée";

            // TVA
            String tva = (cd.tva() != null && !cd.tva().isBlank()) ? cd.tva() : null;
            if (tva == null) tva = tvaByCode.get(codeKey);
            if (tva == null) tva = tvaByName.get(DataReader.normalize(cd.name()));

            // Infos sheet sil + yeniden oluştur
            int idx = wb.getSheetIndex(INFOS_SHEET);
            if (idx >= 0) wb.removeSheetAt(idx);
            Sheet infos = wb.createSheet(INFOS_SHEET);

            CellStyle textStyle = wb.createCellStyle();
            textStyle.setDataFormat(wb.createDataFormat().getFormat("@"));

            writeRow(infos, 0,  "Infos clients", cd.name(),    null,      wb);
            writeRow(infos, 1,  "Adresse",       cd.adresse(), null,      wb);
            writeRow(infos, 2,  "CP",            cd.cp(),      textStyle, wb);
            writeRow(infos, 3,  "Ville",         cd.ville(),   null,      wb);
            writeRow(infos, 4,  "Mail",          cd.mail(),    null,      wb);
            writeRow(infos, 5,  "Téléphone",     cd.tel(),     textStyle, wb);
            writeRow(infos, 6,  "IBAN",          cd.iban(),    textStyle, wb);
            writeRow(infos, 7,  "BIC",           cd.bic(),     textStyle, wb);
            writeRow(infos, 8,  "Commercial",    cd.libelle(), null,      wb);
            writeRow(infos, 9,  "Conditions",    "",           null,      wb);
            if (tva != null && !tva.isBlank())
                writeRow(infos, 10, "N° TVA",    tva,          textStyle, wb);

            // Facture en préparation col D rows 5-8'e adres yaz
            writeFactureAddress(wb, cd);

            try (FileOutputStream fos = new FileOutputStream(excelFile)) {
                wb.write(fos);
            }

            String tvaInfo = (tva != null && !tva.isBlank()) ? " | TVA=" + tva : " | TVA non trouvée";
            return "Infos créée → " + cd.name() + tvaInfo;
        }
    }

    // =========================================================================
    // Facture en préparation col D rows 5-8'e adres yaz (PDF için)
    // =========================================================================

    private void writeFactureAddress(Workbook wb, ClientData cd) {
        Sheet fep = wb.getSheet("Facture en préparation");
        if (fep == null) return;
        String cpVille = ((cd.cp() != null && !cd.cp().isBlank()) ? cd.cp() + " " : "")
                + (cd.ville() != null ? cd.ville() : "");
        String[] values = {
                cd.name()    != null ? cd.name()    : "",
                cd.adresse() != null ? cd.adresse() : "",
                cpVille.trim(),
                cd.libelle() != null ? cd.libelle() : ""
        };
        for (int i = 0; i < values.length; i++) {
            int rowIdx = 4 + i;  // row 4-7 (0-bazlı)
            Row row = fep.getRow(rowIdx);
            if (row == null) row = fep.createRow(rowIdx);
            Cell cell = row.getCell(3);
            if (cell == null) cell = row.createCell(3);
            cell.setCellValue(values[i].trim());
        }
    }

    // =========================================================================
    // Helpers
    // =========================================================================

    private void writeRow(Sheet sheet, int rowIdx, String label, String value,
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

    private Sheet findSheetLike(Workbook wb, String keyword) {
        for (int i = 0; i < wb.getNumberOfSheets(); i++)
            if (DataReader.normalize(wb.getSheetName(i)).contains(keyword))
                return wb.getSheetAt(i);
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

    private record ClientData(
            String name, String code, String adresse, String cp, String ville,
            String mail, String tel, String iban, String bic, String libelle, String tva
    ) {}
}