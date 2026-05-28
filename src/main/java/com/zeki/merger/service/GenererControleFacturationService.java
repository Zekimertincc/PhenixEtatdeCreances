package com.zeki.merger.service;

import com.zeki.merger.trf.DataReader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.function.BiConsumer;
import java.util.stream.Collectors;

public class GenererControleFacturationService {

    private final FolderScanner scanner = new FolderScanner();

    public File apply(File rootFolder, File outputFolder, File recupFile, File tableauFile,
                      BiConsumer<Double, String> progress) throws Exception {
        List<FolderScanner.CompanyFile> companies = scanner.scan(rootFolder);
        if (companies.isEmpty()) {
            progress.accept(1.0, "Aucun dossier trouvé.");
            return null;
        }

        // Filter to only clients present in recupFile (if provided)
        if (recupFile != null && recupFile.exists()) {
            Set<String> recupNames = readRecupNames(recupFile);
            if (!recupNames.isEmpty()) {
                companies = companies.stream()
                        .filter(cf -> recupNames.contains(DataReader.normalize(cf.companyName())))
                        .collect(Collectors.toList());
                progress.accept(0.02, recupNames.size() + " clients dans recup → " + companies.size() + " dossiers filtrés.");
            }
        }

        Map<String, Double> soldeMap = new LinkedHashMap<>();
        if (tableauFile != null && tableauFile.exists()) {
            soldeMap = readSoldeMap(tableauFile);
            progress.accept(0.04, soldeMap.size() + " soldes lus depuis Tableau de bord.");
        }

        int total = companies.size();
        List<Object[]> rows = new ArrayList<>();
        final Map<String, Double> finalSoldeMap = soldeMap;

        for (int i = 0; i < total; i++) {
            FolderScanner.CompanyFile cf = companies.get(i);
            double prog = 0.05 + 0.85 * (i + 1.0) / total;
            try {
                Object[] row = extractRow(cf, finalSoldeMap);
                if (row != null) rows.add(row);
                progress.accept(prog, "[" + (i+1) + "/" + total + "] " + cf.companyName());
            } catch (Exception e) {
                progress.accept(prog, "[" + (i+1) + "/" + total + "] ERREUR " + cf.companyName() + ": " + e.getMessage());
            }
        }

        progress.accept(0.92, "Écriture Controle_Facturation.xlsx...");
        File out = writeOutput(rows, outputFolder);
        progress.accept(1.0, "✓ Controle_Facturation.xlsx généré — " + rows.size() + " clients.");
        return out;
    }

    private Object[] extractRow(FolderScanner.CompanyFile cf, Map<String, Double> soldeMap) throws IOException {
        try (Workbook wb = openWorkbook(cf.excelFile())) {
            Sheet creances = wb.getSheet("Créances");
            String nomClient = cf.companyName();
            if (creances != null) {
                Row r3 = creances.getRow(3);
                if (r3 != null) {
                    Cell h4 = r3.getCell(7);
                    if (h4 != null && !h4.toString().isBlank()) nomClient = h4.toString().trim();
                }
            }

            Sheet facture = wb.getSheet("Facture en préparation");
            if (facture == null) {
                for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                    if (wb.getSheetName(i).toLowerCase().contains("facture")) {
                        facture = wb.getSheetAt(i); break;
                    }
                }
            }
            if (facture == null) return null;

            DataFormatter fmt = new DataFormatter();
            FormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator();

            int ligneDuA = -1;
            for (int r = 0; r < Math.min(facture.getLastRowNum(), 200); r++) {
                Row row = facture.getRow(r);
                if (row == null) continue;
                Cell cell = row.getCell(0, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (cell != null && "A".equals(fmt.formatCellValue(cell, ev).trim())) {
                    ligneDuA = r; break;
                }
            }
            if (ligneDuA < 0) return null;

            double ag   = numVal(facture, ligneDuA,   2, fmt, ev);
            double cl   = numVal(facture, ligneDuA+1, 2, fmt, ev);
            double agcl = numVal(facture, ligneDuA+2, 2, fmt, ev);

            int ligneDuD = -1;
            for (int r = ligneDuA+3; r < Math.min(facture.getLastRowNum(), 200); r++) {
                Row row = facture.getRow(r);
                if (row == null) continue;
                Cell cell = row.getCell(0, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (cell != null && "D".equals(fmt.formatCellValue(cell, ev).trim())) {
                    ligneDuD = r; break;
                }
            }
            if (ligneDuD < 0) return null;

            double comsHt   = numVal(facture, ligneDuD,   2, fmt, ev);
            double prodHt   = numVal(facture, ligneDuD+1, 2, fmt, ev);
            double totalHt  = numVal(facture, ligneDuD+2, 2, fmt, ev);
            double tva      = numVal(facture, ligneDuD+3, 2, fmt, ev);
            double totalTtc = numVal(facture, ligneDuD+4, 2, fmt, ev);

            String norm = DataReader.normalize(nomClient);
            Double solde = soldeMap.get(norm);
            if (solde == null) {
                for (Map.Entry<String, Double> e : soldeMap.entrySet()) {
                    String k = e.getKey();
                    if (norm.contains(k) || k.contains(norm)) { solde = e.getValue(); break; }
                }
            }
            double soldePrecedent = solde != null ? solde : 0.0;
            return new Object[]{nomClient, ag, cl, agcl, soldePrecedent, comsHt, prodHt, totalHt, tva, totalTtc};
        }
    }

    private double numVal(Sheet sheet, int rowIdx, int colIdx,
                          DataFormatter fmt, FormulaEvaluator ev) {
        Row row = sheet.getRow(rowIdx);
        if (row == null) return 0.0;
        Cell cell = row.getCell(colIdx, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (cell == null) return 0.0;
        CellType type = cell.getCellType() == CellType.FORMULA
                ? cell.getCachedFormulaResultType() : cell.getCellType();
        if (type == CellType.NUMERIC) return cell.getNumericCellValue();
        try {
            return Double.parseDouble(fmt.formatCellValue(cell, ev).replace(",", ".").replace(" ", ""));
        } catch (NumberFormatException e) {
            return 0.0;
        }
    }

    private File writeOutput(List<Object[]> rows, File outputFolder) throws IOException {
        String ts = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd_HH-mm"));
        File out = new File(outputFolder, "Controle_Facturation_" + ts + ".xlsx");

        String[] headers = {"CLIENT", "AG", "CL", "AG+CL", "SOLDES PRÉCÉDENTS", "COMS HT", "PROD HT", "TOTAL HT", "TVA", "TOTAL TTC"};

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Controle");
            DataFormat df = wb.createDataFormat();
            short moneyFmt = df.getFormat("#,##0.00");

            XSSFCellStyle headerStyle = wb.createCellStyle();
            XSSFFont whiteFont = wb.createFont();
            whiteFont.setBold(true);
            whiteFont.setColor(new XSSFColor(new byte[]{(byte)0xFF, (byte)0xFF, (byte)0xFF}, null));
            headerStyle.setFont(whiteFont);
            headerStyle.setFillForegroundColor(new XSSFColor(new byte[]{(byte)0x1F, (byte)0x4E, (byte)0x79}, null));
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            XSSFCellStyle moneyStyle = wb.createCellStyle();
            moneyStyle.setDataFormat(moneyFmt);

            XSSFRow hdr = sheet.createRow(0);
            for (int c = 0; c < headers.length; c++) {
                XSSFCell cell = hdr.createCell(c);
                cell.setCellValue(headers[c]);
                cell.setCellStyle(headerStyle);
            }

            for (int r = 0; r < rows.size(); r++) {
                Object[] data = rows.get(r);
                XSSFRow row = sheet.createRow(r + 1);
                row.createCell(0).setCellValue((String) data[0]);
                for (int c = 1; c < data.length; c++) {
                    XSSFCell cell = row.createCell(c);
                    cell.setCellValue((Double) data[c]);
                    cell.setCellStyle(moneyStyle);
                }
            }

            for (int c = 0; c < headers.length; c++) sheet.autoSizeColumn(c);

            try (FileOutputStream fos = new FileOutputStream(out)) { wb.write(fos); }
        }
        return out;
    }

    private Map<String, Double> readSoldeMap(File tableauFile) throws IOException {
        Map<String, Double> map = new LinkedHashMap<>();
        try (Workbook wb = openWorkbook(tableauFile)) {
            Sheet sheet = wb.getSheet("Soldes");
            if (sheet == null) return map;
            DataFormatter fmt = new DataFormatter();
            FormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator();
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                Cell nameCell = row.getCell(0, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (nameCell == null) continue;
                String name = fmt.formatCellValue(nameCell, ev).trim();
                if (name.isBlank()) continue;
                Cell soldeCell = row.getCell(2, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (soldeCell == null) continue;
                CellType ct = soldeCell.getCellType() == CellType.FORMULA
                        ? soldeCell.getCachedFormulaResultType() : soldeCell.getCellType();
                if (ct == CellType.NUMERIC) {
                    map.put(DataReader.normalize(name), soldeCell.getNumericCellValue());
                }
            }
        }
        return map;
    }

    private Set<String> readRecupNames(File recupFile) throws IOException {
        Set<String> names = new HashSet<>();
        try (Workbook wb = openWorkbook(recupFile)) {
            Sheet sheet = wb.getSheet("Feuil1");
            if (sheet == null) sheet = wb.getSheetAt(0);
            DataFormatter fmt = new DataFormatter();
            FormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator();
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                Cell cell = row.getCell(0, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (cell == null) break;
                String name = fmt.formatCellValue(cell, ev).trim();
                if (name.isBlank()) break;
                names.add(DataReader.normalize(name));
            }
        }
        return names;
    }

    private Workbook openWorkbook(File file) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        return file.getName().toLowerCase().endsWith(".xls")
                ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
    }
}