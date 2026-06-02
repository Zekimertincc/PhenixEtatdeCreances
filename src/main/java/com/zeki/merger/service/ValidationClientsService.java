package com.zeki.merger.service;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.util.function.BiConsumer;

/**
 * Validation des clients — month-end closing operation.
 *
 * For each company's état de créances, on the "Créances" sheet:
 * - Find rows where S (Lieu, col 19) = "AG" AND R (col 18) > 0
 * - For those rows: write U's computed value (=I+R) into I (col 9) as a plain number
 * - Then zero out R (col 18), S (col 19), T (col 20) — preserving cell format
 */
public class ValidationClientsService {

    private final FolderScanner scanner = new FolderScanner();

    public List<String> apply(File rootFolder, BiConsumer<Double, String> progress) throws Exception {
        List<String> log = new ArrayList<>();

        List<FolderScanner.CompanyFile> companies = scanner.scan(rootFolder);
        if (companies.isEmpty()) {
            progress.accept(1.0, "Aucun dossier trouvé dans: " + rootFolder.getName());
            return log;
        }

        int total = companies.size();
        for (int i = 0; i < total; i++) {
            FolderScanner.CompanyFile cf = companies.get(i);
            double prog = 0.05 + 0.95 * (i + 1.0) / total;
            String entry;
            try {
                entry = processCompany(cf.excelFile());
            } catch (Exception e) {
                entry = "ERREUR: " + e.getMessage();
            }
            log.add(cf.companyName() + " → " + entry);
            progress.accept(prog, "[" + (i + 1) + "/" + total + "] " + cf.companyName() + " → " + entry);
        }

        progress.accept(1.0, "Validation terminée (" + total + " dossiers).");
        return log;
    }

    private String processCompany(File excelFile) throws IOException {
        byte[] bytes = Files.readAllBytes(excelFile.toPath());

        try (Workbook wb = openWorkbookFromBytes(bytes, excelFile.getName())) {
            Sheet sheet = wb.getSheet("Créances");
            if (sheet == null) return "sheet 'Créances' introuvable";

            DataFormatter fmt = new DataFormatter();
            FormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator();

            // Find header row — look for row containing "RECOUVRE ET FACTURE" in col I (9)
            int headerRow = -1;
            for (int r = 0; r <= Math.min(sheet.getLastRowNum(), 30); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                Cell cell = row.getCell(8, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL); // col I = index 8
                if (cell != null && "RECOUVRE ET FACTURE".equalsIgnoreCase(fmt.formatCellValue(cell, ev).trim())) {
                    headerRow = r;
                    break;
                }
            }
            if (headerRow < 0) return "ligne d'en-tête introuvable";

            int modifiedCount = 0;

            for (int r = headerRow + 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                // Col S = index 18 (Lieu), Col R = index 17, Col T = index 19
                Cell sCell = row.getCell(18, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL); // S = Lieu
                Cell rCell = row.getCell(17, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL); // R = Dont en attente
                Cell iCell = row.getCell(8,  Row.MissingCellPolicy.RETURN_BLANK_AS_NULL); // I = Recouvré et facturé
                Cell uCell = row.getCell(20, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL); // U = Recouvré total

                if (sCell == null || rCell == null) continue;

                String lieu = fmt.formatCellValue(sCell, ev).trim();
                Cell tCell = row.getCell(19, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL); // T = Frais de procédure
                double rVal = numericVal(rCell, ev);
                double tVal = tCell != null ? numericVal(tCell, ev) : 0.0;

                boolean validLieu = "AG".equalsIgnoreCase(lieu)
                        || "CL".equalsIgnoreCase(lieu)
                        || "NA".equalsIgnoreCase(lieu);
                // R veya T'den biri doluysa işlem yap
                if (!validLieu || (rVal <= 0.0 && tVal <= 0.0)) continue;

                // Compute U = I + R
                double iVal = iCell != null ? numericVal(iCell, ev) : 0.0;
                double uVal = uCell != null ? numericVal(uCell, ev) : (iVal + rVal);

                // Write U's value into I — as plain number, preserving cell format
                if (iCell == null) iCell = row.createCell(8, CellType.NUMERIC);
                iCell.setCellValue(uVal);

                // Zero out R (col 17), S (col 18), T (col 19) — preserve format
                zeroCellValue(row, 17, wb);
                zeroCellValue(row, 18, wb);
                zeroCellValue(row, 19, wb);

                modifiedCount++;
            }

            if (modifiedCount == 0) return "aucune ligne AG/CL/NA avec montant trouvée";

            try (FileOutputStream fos = new FileOutputStream(excelFile)) {
                wb.write(fos);
            }
            return modifiedCount + " ligne(s) validée(s)";
        }
    }

    private void zeroCellValue(Row row, int colIdx, Workbook wb) {
        Cell cell = row.getCell(colIdx, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (cell == null) cell = row.createCell(colIdx);
        // S column (index 18) is a string cell — clear it with empty string, not 0
        if (colIdx == 18) {
            cell.setCellValue("");
        } else {
            cell.setCellType(CellType.NUMERIC);
            cell.setCellValue(0.0);
        }
    }

    private double numericVal(Cell cell, FormulaEvaluator ev) {
        if (cell == null) return 0.0;
        CellType ct = cell.getCellType() == CellType.FORMULA
                ? cell.getCachedFormulaResultType() : cell.getCellType();
        if (ct == CellType.NUMERIC) return cell.getNumericCellValue();
        return 0.0;
    }

    private Workbook openWorkbookFromBytes(byte[] bytes, String fileName) throws IOException {
        ByteArrayInputStream bis = new ByteArrayInputStream(bytes);
        return fileName.toLowerCase().endsWith(".xls")
                ? new HSSFWorkbook(bis) : new XSSFWorkbook(bis);
    }
}
