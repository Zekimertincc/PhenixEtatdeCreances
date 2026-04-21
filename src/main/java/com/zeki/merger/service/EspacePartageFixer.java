package com.zeki.merger.service;

import com.zeki.merger.AppConfig;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.Normalizer;
import java.util.function.BiConsumer;

/**
 * Reads CorrespondanceClient-EspacePartage.xlsx, ensures every EspacePartagé
 * path (column C) ends with {@link AppConfig#ETAT_CREANCES_SUFFIX}, and writes
 * the result either in-place or as a _fixed sibling depending on
 * {@link AppConfig#FIX_OVERWRITE}.
 */
public class EspacePartageFixer {

    private static final int COL_ESPACE_PARTAGE = 2; // C (0-based)
    private static final String SENTINEL = "_END_";

    /** Canonical suffix stripped of accents and lowercased, used for comparison only. */
    private static final String SUFFIX_NORMALIZED =
        normalize(AppConfig.ETAT_CREANCES_SUFFIX.replace("\\", "").strip());

    public File fix(File rootFolder, BiConsumer<Double, String> progress) throws IOException {
        File source = new File(rootFolder, AppConfig.ESPACE_PARTAGE_FILENAME);
        if (!source.exists()) {
            throw new IOException("File not found: " + source.getAbsolutePath());
        }

        progress.accept(0.05, "Reading " + source.getName() + "…");

        Workbook wb;
        try (FileInputStream fis = new FileInputStream(source)) {
            wb = new XSSFWorkbook(fis);
        }

        Sheet sheet = wb.getSheet("Correspondance");
        if (sheet == null) {
            wb.close();
            throw new IOException("Sheet \"Correspondance\" not found in " + source.getName());
        }

        int fixed = 0;
        int total = sheet.getLastRowNum(); // approximate for progress

        for (int r = 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;

            Cell cellC = row.getCell(COL_ESPACE_PARTAGE);
            String raw = cellC != null && cellC.getCellType() == CellType.STRING
                ? cellC.getStringCellValue()
                : (cellC != null ? new DataFormatter().formatCellValue(cellC) : "");

            if (SENTINEL.equals(raw.trim())) {
                progress.accept(0.8, "Reached sentinel _END_ at row " + (r + 1));
                break;
            }
            if (raw.isBlank()) continue;

            if (!endsWithSuffix(raw)) {
                String updated = raw.stripTrailing() + AppConfig.ETAT_CREANCES_SUFFIX;
                if (cellC == null) cellC = row.createCell(COL_ESPACE_PARTAGE);
                cellC.setCellValue(updated);
                fixed++;
                progress.accept(0.1 + 0.7 * r / Math.max(total, 1),
                    "Row " + (r + 1) + ": appended suffix → " + updated);
            }
        }

        progress.accept(0.85, fixed + " path(s) updated. Saving…");

        File dest = AppConfig.FIX_OVERWRITE
            ? source
            : new File(source.getParent(),
                source.getName().replace(".xlsx", "_fixed.xlsx"));

        try (FileOutputStream fos = new FileOutputStream(dest)) {
            wb.write(fos);
        }
        wb.close();

        progress.accept(1.0, "Saved → " + dest.getAbsolutePath());
        return dest;
    }

    // -------------------------------------------------------------------------

    private static boolean endsWithSuffix(String path) {
        String tail = path.stripTrailing();
        // Extract the last segment after the final backslash (if any)
        int slash = tail.lastIndexOf('\\');
        String lastSegment = slash >= 0 ? tail.substring(slash + 1) : tail;
        return normalize(lastSegment).equals(SUFFIX_NORMALIZED);
    }

    /** Lower-cases and strips diacritics so accent variants match. */
    private static String normalize(String s) {
        String decomposed = Normalizer.normalize(s, Normalizer.Form.NFD);
        return decomposed.replaceAll("\\p{M}", "").toLowerCase();
    }
}
