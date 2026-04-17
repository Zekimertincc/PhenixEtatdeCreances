package com.zeki.merger.service;

import java.io.File;
import java.text.Normalizer;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.List;
import java.util.Optional;

/**
 * Walks the root folder, finds every company sub-directory, and resolves the
 * Excel file to process inside the "Etat des créances" sub-folder (or any
 * accent/case variation of that name).
 */
public class FolderScanner {

    public record CompanyFile(String companyName, File excelFile) {}

    /**
     * Scans {@code rootFolder} for company sub-directories and returns one
     * {@link CompanyFile} per company where a matching Excel file was found.
     * Results are sorted by company name for reproducible output ordering.
     */
    public List<CompanyFile> scan(File rootFolder) {
        List<CompanyFile> results = new ArrayList<>();

        File[] companies = rootFolder.listFiles(File::isDirectory);
        if (companies == null) return results;

        Arrays.sort(companies, Comparator.comparing(File::getName, String.CASE_INSENSITIVE_ORDER));

        for (File companyDir : companies) {
            findExcelFile(companyDir)
                .ifPresent(f -> results.add(new CompanyFile(companyDir.getName(), f)));
        }
        return results;
    }

    // -------------------------------------------------------------------------

    private Optional<File> findExcelFile(File companyDir) {
        // Step 1 — find the "etat … creances" subfolder (accent/case-insensitive)
        File[] subDirs = companyDir.listFiles(File::isDirectory);
        if (subDirs == null) return Optional.empty();

        Optional<File> etcDir = Arrays.stream(subDirs)
            .filter(d -> isEtatCreancesFolder(d.getName()))
            .findFirst();

        if (etcDir.isEmpty()) return Optional.empty();

        // Step 2 — find an .xlsx/.xls file whose name starts with "etat"
        File[] candidates = etcDir.get().listFiles(f ->
            f.isFile()
            && normalize(f.getName()).startsWith("etat")
            && (f.getName().endsWith(".xlsx") || f.getName().endsWith(".xls"))
        );

        if (candidates == null || candidates.length == 0) return Optional.empty();

        // If multiple files match, pick the most recently modified one.
        return Arrays.stream(candidates)
            .max(Comparator.comparingLong(File::lastModified));
    }

    /**
     * Returns true when the folder name, after stripping accents and
     * lower-casing, contains both "etat" and "cr" — which covers:
     *   "Etat des créances", "état de créances", "Etat de creances", etc.
     */
    private boolean isEtatCreancesFolder(String name) {
        String n = normalize(name);
        return n.contains("etat") && n.contains("cr");
    }

    /**
     * Strips diacritical marks (accents) from {@code s} and returns it
     * lower-cased, so comparisons are accent- and case-insensitive.
     * e.g. "Érat des Créances" → "etat des creances"
     */
    private String normalize(String s) {
        // NFD decomposes characters + combining marks (accents) into separate code points;
        // the regex then removes every combining mark category.
        return Normalizer.normalize(s, Normalizer.Form.NFD)
            .replaceAll("\\p{InCombiningDiacriticalMarks}", "")
            .toLowerCase();
    }
}
