package com.zeki.merger.service;

import com.zeki.merger.AppConfig;
import com.zeki.merger.db.DatabaseManager;
import com.zeki.merger.model.CreanceRow;

import java.io.File;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.function.BiConsumer;

/**
 * Orchestrates the full pipeline:
 * scan → read → group by company → write.
 *
 * Rows are collected into a {@link LinkedHashMap} keyed by company name so that
 * insertion order (alphabetical, from {@link FolderScanner}) is preserved and
 * the writer can emit one group block per company.
 * Companies with zero matching rows are never added to the map, so the writer
 * skips them automatically.
 */
public class MergeService {

    private final FolderScanner   scanner   = new FolderScanner();
    private final ExcelReader     reader    = new ExcelReader();
    private final ExcelWriter     writer    = new ExcelWriter();
    private final TrfWriter       trfWriter = new TrfWriter();
    private final DatabaseManager db;

    public MergeService(DatabaseManager db) {
        this.db = db;
    }

    public File merge(File rootFolder,
                      File outputFolder,
                      BiConsumer<Double, String> progressCallback) throws Exception {

        log(progressCallback, 0.00, "Scanning: " + rootFolder.getAbsolutePath());

        List<FolderScanner.CompanyFile> companyFiles = scanner.scan(rootFolder);

        if (companyFiles.isEmpty()) {
            log(progressCallback, 1.00,
                "No matching files found. Check that company sub-folders contain a " +
                "folder whose name includes \"etat\" and \"cr\", with an Excel file " +
                "starting with \"" + AppConfig.FILE_PREFIX + "\".");
            return null;
        }

        log(progressCallback, 0.05,
            "Found " + companyFiles.size() + " company file(s).");

        // LinkedHashMap preserves the alphabetical insertion order from FolderScanner.
        Map<String, List<CreanceRow>> groupedRows = new LinkedHashMap<>();
        int total   = companyFiles.size();
        int skipped = 0;
        int totalRowCount = 0;

        for (int i = 0; i < total; i++) {
            FolderScanner.CompanyFile cf = companyFiles.get(i);
            double progress = 0.05 + 0.85 * (double) (i + 1) / total;

            log(progressCallback, progress,
                "[" + (i + 1) + "/" + total + "] " + cf.companyName()
                + "  →  " + cf.excelFile().getName());

            try {
                List<CreanceRow> rows = reader.readFiltered(cf.companyName(), cf.excelFile());
                if (rows.isEmpty()) {
                    log(progressCallback, progress,
                        "[" + cf.companyName() + "] SKIPPED - no data in column S");
                } else {
                    groupedRows.put(cf.companyName(), rows);
                    totalRowCount += rows.size();
                    log(progressCallback, progress,
                        "     " + rows.size() + " row(s) matched (column "
                        + AppConfig.FILTER_COLUMN_LABEL + " non-empty).");
                    if (db != null) {
                        try {
                            long cid = db.upsertCompany(cf.companyName(),
                                                        cf.excelFile().getAbsolutePath());
                            db.replaceCreanceRows(cid, rows);
                        } catch (Exception dbEx) {
                            log(progressCallback, progress, "  [DB] " + dbEx.getMessage());
                        }
                    }
                }
            } catch (Exception e) {
                skipped++;
                log(progressCallback, progress,
                    "     ERROR: " + e.getMessage() + " — file skipped.");
            }
        }

        // ---- Explicit pre-filter: remove every company that has zero matching rows.
        // This is the single, authoritative place where the exclusion is enforced.
        // ExcelWriter must never receive a company with an empty row list.
        groupedRows.entrySet().removeIf(e -> {
            if (e.getValue() == null || e.getValue().isEmpty()) {
                log(progressCallback, 0.88,
                    "[" + e.getKey() + "] SKIPPED - no data in column S");
                return true;
            }
            return false;
        });

        if (groupedRows.isEmpty()) {
            log(progressCallback, 1.00,
                "No rows matched the filter across all companies. Output not written.");
            return null;
        }

        log(progressCallback, 0.90,
            "Total rows: " + totalRowCount
            + " across " + groupedRows.size() + " company/ies"
            + (skipped > 0 ? "  |  errors: " + skipped : "")
            + ".  Writing output…");

        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd_HH-mm-ss"));
        String outputFilename = AppConfig.OUTPUT_FILENAME.replace(".xlsx", "_" + timestamp + ".xlsx");
        File outputFile = new File(outputFolder, outputFilename);
        writer.write(groupedRows, outputFile);

        log(progressCallback, 1.00,
            "Done.  Output: " + outputFile.getAbsolutePath());

        return outputFile;
    }

    // -------------------------------------------------------------------------

    public File exportTrf(File rootFolder,
                          File outputFolder,
                          BiConsumer<Double, String> progressCallback) throws Exception {

        log(progressCallback, 0.00, "TRF export — scanning: " + rootFolder.getAbsolutePath());

        List<FolderScanner.CompanyFile> companyFiles = scanner.scan(rootFolder);

        if (companyFiles.isEmpty()) {
            log(progressCallback, 1.00,
                "No matching files found. Check that company sub-folders contain a " +
                "folder whose name includes \"etat\" and \"cr\", with an Excel file " +
                "starting with \"" + AppConfig.FILE_PREFIX + "\".");
            return null;
        }

        log(progressCallback, 0.05,
            "Found " + companyFiles.size() + " company file(s).");

        Map<String, List<CreanceRow>> groupedRows = new LinkedHashMap<>();
        int total   = companyFiles.size();
        int skipped = 0;
        int totalRowCount = 0;

        for (int i = 0; i < total; i++) {
            FolderScanner.CompanyFile cf = companyFiles.get(i);
            double progress = 0.05 + 0.85 * (double) (i + 1) / total;

            log(progressCallback, progress,
                "[" + (i + 1) + "/" + total + "] " + cf.companyName()
                + "  →  " + cf.excelFile().getName());

            try {
                List<CreanceRow> rows = reader.readFiltered(cf.companyName(), cf.excelFile());
                if (rows.isEmpty()) {
                    log(progressCallback, progress,
                        "[" + cf.companyName() + "] SKIPPED - no data in column S");
                } else {
                    groupedRows.put(cf.companyName(), rows);
                    totalRowCount += rows.size();
                    log(progressCallback, progress,
                        "     " + rows.size() + " row(s) matched.");
                }
            } catch (Exception e) {
                skipped++;
                log(progressCallback, progress,
                    "     ERROR: " + e.getMessage() + " — file skipped.");
            }
        }

        groupedRows.entrySet().removeIf(e -> e.getValue() == null || e.getValue().isEmpty());

        if (groupedRows.isEmpty()) {
            log(progressCallback, 1.00,
                "No rows matched the filter across all companies. Output not written.");
            return null;
        }

        log(progressCallback, 0.90,
            "Total rows: " + totalRowCount
            + " across " + groupedRows.size() + " company/ies"
            + (skipped > 0 ? "  |  errors: " + skipped : "")
            + ".  Writing TRF output…");

        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd_HH-mm-ss"));
        String outputFilename = AppConfig.TRF_OUTPUT_FILENAME.replace(".xlsx", "_" + timestamp + ".xlsx");
        File outputFile = new File(outputFolder, outputFilename);
        trfWriter.write(groupedRows, outputFile);

        log(progressCallback, 1.00,
            "Done.  Output: " + outputFile.getAbsolutePath());

        return outputFile;
    }

    // -------------------------------------------------------------------------

    private void log(BiConsumer<Double, String> cb, double progress, String message) {
        cb.accept(progress, message);
    }
}
