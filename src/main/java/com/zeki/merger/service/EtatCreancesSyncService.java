package com.zeki.merger.service;

import com.zeki.merger.db.DatabaseManager;
import com.zeki.merger.model.CreanceRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.function.BiConsumer;

public class EtatCreancesSyncService {

    private final DatabaseManager db;
    private final ExcelReader    reader  = new ExcelReader();
    private final FolderScanner  scanner = new FolderScanner();

    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ofPattern("yyyy-MM-dd");

    // Col indices in Créances sheet (0-based, row 15 = header, data from row 16)
    private static final int COL_CREANCE_PRINCIPALE = 7;   // H
    private static final int COL_RECOUVRE_TOTAL     = 20;  // U
    private static final int COL_COMMISSIONS        = 23;  // X
    private static final int COL_ETAT               = 9;   // J
    private static final int COL_DATE               = 2;   // C  (REMIS LE)
    private static final int COL_FILTER             = 9;   // J  (ETAT — always filled for real rows)

    public EtatCreancesSyncService(DatabaseManager db) {
        this.db = db;
    }

    public void syncAll(File rootFolder, BiConsumer<Double, String> progress) throws Exception {
        List<FolderScanner.CompanyFile> companies = scanner.scan(rootFolder);
        if (companies.isEmpty()) {
            log(progress, 1.0, "Aucune société trouvée.");
            return;
        }
        int total = companies.size();
        for (int i = 0; i < total; i++) {
            FolderScanner.CompanyFile cf = companies.get(i);
            double pct = (double) i / total;
            log(progress, pct, "Sync : " + cf.companyName());
            try {
                syncCompany(cf);
                log(progress, pct, "  ✓ " + cf.companyName());
            } catch (Exception e) {
                log(progress, pct, "  ✗ " + cf.companyName() + " — " + e.getMessage());
            }
        }
        log(progress, 1.0, "Synchronisation terminée. " + total + " sociétés.");
    }

    public void syncCompany(FolderScanner.CompanyFile cf) throws Exception {
        if (db == null) return;

        // 1. Write raw rows (existing behaviour)
        List<CreanceRow> rows = reader.readFiltered(cf.companyName(), cf.excelFile());
        long companyId = db.upsertCompany(cf.companyName(), cf.excelFile().getAbsolutePath());
        db.replaceCreanceRows(companyId, rows);

        // 2. Compute and persist summary
        computeAndSaveSummary(companyId, cf);
    }

    // -------------------------------------------------------------------------
    // Summary computation — reads Créances sheet directly with POI
    // -------------------------------------------------------------------------

    private void computeAndSaveSummary(long companyId, FolderScanner.CompanyFile cf)
            throws Exception {
        File file = cf.excelFile();
        try (FileInputStream fis = new FileInputStream(file);
             Workbook wb = file.getName().toLowerCase().endsWith(".xls")
                     ? new HSSFWorkbook(fis)
                     : new XSSFWorkbook(fis)) {

            Sheet sheet = wb.getSheet("Créances");
            if (sheet == null) sheet = wb.getSheetAt(0);

            DataFormatter    fmt = new DataFormatter();
            FormulaEvaluator ev  = wb.getCreationHelper().createFormulaEvaluator();

            // --- Meta rows (fixed positions) ---
            String codeClient   = extractCodeClient(sheet, fmt);
            String responsable  = extractResponsable(sheet, fmt);

            // --- Data rows (row 16+ = index 16+) ---
            int    nbDossiers  = 0;
            int    nbSoldes    = 0;
            int    nbGestion   = 0;
            int    nbIrr       = 0;
            int    nbArj       = 0;
            double creanceTot  = 0.0;
            double recouvTot   = 0.0;
            double commTot     = 0.0;
            LocalDate lastDate = null;

            for (int r = 16; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                // Skip rows where filter column (S) is blank
                Cell filterCell = row.getCell(COL_FILTER);
                if (!hasValue(filterCell, fmt, ev)) continue;

                nbDossiers++;

                String etat = getString(row, COL_ETAT, fmt, ev).trim().toLowerCase();
                if      (etat.startsWith("soldé") || etat.equals("solde"))           nbSoldes++;
                else if (etat.contains("gestion") || etat.contains("cours"))         nbGestion++;
                else if (etat.startsWith("irr"))                                     nbIrr++;
                else if (etat.startsWith("arj") || etat.startsWith("arb"))           nbArj++;

                creanceTot += getDouble(row, COL_CREANCE_PRINCIPALE, ev);
                recouvTot  += getDouble(row, COL_RECOUVRE_TOTAL,     ev);
                commTot    += getDouble(row, COL_COMMISSIONS,         ev);

                LocalDate d = getDate(row, COL_DATE, ev);
                if (d != null && (lastDate == null || d.isAfter(lastDate))) {
                    lastDate = d;
                }
            }

            int nbAutres = nbDossiers - nbSoldes - nbGestion - nbIrr - nbArj;

            db.upsertCompanySummary(
                    companyId, codeClient, responsable,
                    nbDossiers, nbSoldes, nbGestion, nbIrr, nbArj, Math.max(0, nbAutres),
                    creanceTot, recouvTot, commTot,
                    lastDate != null ? lastDate.format(DATE_FMT) : null);
        }
    }

    // -------------------------------------------------------------------------
    // Meta extraction
    // -------------------------------------------------------------------------

    /** Row 12 (index 12), col A: "Code Client : 940135" */
    private String extractCodeClient(Sheet sheet, DataFormatter fmt) {
        Row row = sheet.getRow(12);
        if (row == null) return "";
        Cell cell = row.getCell(0);
        if (cell == null) return "";
        String raw = fmt.formatCellValue(cell).trim();
        int colon = raw.indexOf(':');
        return colon >= 0 ? raw.substring(colon + 1).trim() : raw;
    }

    /** Row 7 (index 7), col H: responsable name */
    private String extractResponsable(Sheet sheet, DataFormatter fmt) {
        Row row = sheet.getRow(7);
        if (row == null) return "";
        Cell cell = row.getCell(7);
        if (cell == null) return "";
        return fmt.formatCellValue(cell).trim();
    }

    // -------------------------------------------------------------------------
    // Cell helpers
    // -------------------------------------------------------------------------

    private boolean hasValue(Cell cell, DataFormatter fmt, FormulaEvaluator ev) {
        if (cell == null) return false;
        try {
            CellType type = cell.getCellType() == CellType.FORMULA
                    ? ev.evaluate(cell).getCellType() : cell.getCellType();
            if (type == CellType.BLANK) return false;
            if (type == CellType.NUMERIC) return cell.getNumericCellValue() != 0.0;
            if (type == CellType.STRING)  return !cell.getStringCellValue().trim().isEmpty();
            return true;
        } catch (Exception e) {
            return false;
        }
    }

    private String getString(Row row, int col, DataFormatter fmt, FormulaEvaluator ev) {
        Cell cell = row.getCell(col);
        if (cell == null) return "";
        try { return fmt.formatCellValue(cell, ev).trim(); } catch (Exception e) { return ""; }
    }

    private double getDouble(Row row, int col, FormulaEvaluator ev) {
        Cell cell = row.getCell(col);
        if (cell == null) return 0.0;
        try {
            CellType type = cell.getCellType() == CellType.FORMULA
                    ? ev.evaluate(cell).getCellType() : cell.getCellType();
            if (type == CellType.NUMERIC) return cell.getNumericCellValue();
            String s = cell.toString().trim().replaceAll("[^0-9.,\\-]", "")
                    .replace(",", ".");
            return s.isEmpty() ? 0.0 : Double.parseDouble(s);
        } catch (Exception e) { return 0.0; }
    }

    private LocalDate getDate(Row row, int col, FormulaEvaluator ev) {
        Cell cell = row.getCell(col);
        if (cell == null) return null;
        try {
            if (cell.getCellType() == CellType.NUMERIC
                    || (cell.getCellType() == CellType.FORMULA
                        && ev.evaluate(cell).getCellType() == CellType.NUMERIC)) {
                return cell.getLocalDateTimeCellValue().toLocalDate();
            }
        } catch (Exception ignored) {}
        return null;
    }

    private void log(BiConsumer<Double, String> cb, double pct, String msg) {
        if (cb != null) cb.accept(pct, msg);
    }
}
