package com.zeki.merger.trf;

import com.zeki.merger.trf.model.ClientInfo;
import com.zeki.merger.trf.model.ClientSummary;
import com.zeki.merger.trf.model.ConsolidationRow;

import java.io.File;
import java.time.LocalDate;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.function.BiConsumer;

/**
 * Top-level service that:
 * <ol>
 *   <li>Locates the three input files in the given folder (by name, case-insensitive).</li>
 *   <li>Reads all source data via {@link DataReader}.</li>
 *   <li>Computes per-client summaries via {@link TrfCalculator}.</li>
 *   <li>Writes the TRF workbook via {@link TrfSheetWriter}.</li>
 * </ol>
 *
 * Expected input file names (case-insensitive):
 * <ul>
 *   <li>ConsolidationGenerale.xlsx</li>
 *   <li>LISTING_CABINET_PHENIX_pour_ZEKI.xls</li>
 *   <li>Tableau_de_bord_facturation.xlsx</li>
 * </ul>
 *
 * Output file: {@code TRF_MM_YYYY.xlsx} where MM/YYYY = current month.
 */
public class TrfGeneratorService {

    private static final String CONSOLIDATION_NAME = "consolidation";
    private static final String LISTING_NAME        = "listing";
    private static final String TABLEAU_BORD_NAME   = "tableau";

    private final DataReader     reader    = new DataReader();
    private final TrfCalculator  calculator = new TrfCalculator();
    private final TrfSheetWriter writer    = new TrfSheetWriter();

    /**
     * Generates the TRF file from files found inside {@code inputFolder}.
     *
     * @param inputFolder   folder that contains the three source Excel files
     * @param outputFolder  destination folder for the generated TRF_MM_YYYY.xlsx
     * @param progress      callback {@code (0..1, message)} for UI feedback
     * @return the written output file
     */
    public File generate(File inputFolder, File outputFolder,
                         BiConsumer<Double, String> progress) throws Exception {

        log(progress, 0.00, "TRF Generator — scanning: " + inputFolder.getAbsolutePath());

        // ---- Locate input files ------------------------------------------
        File consolidationFile = findFile(inputFolder, CONSOLIDATION_NAME);
        File listingFile       = findFile(inputFolder, LISTING_NAME);
        File tableauFile       = findFile(inputFolder, TABLEAU_BORD_NAME);

        if (consolidationFile == null) throw new IllegalArgumentException(
            "ConsolidationGenerale file not found in: " + inputFolder.getAbsolutePath());
        if (listingFile == null) throw new IllegalArgumentException(
            "LISTING_CABINET_PHENIX file not found in: " + inputFolder.getAbsolutePath());
        if (tableauFile == null) throw new IllegalArgumentException(
            "Tableau_de_bord_facturation file not found in: " + inputFolder.getAbsolutePath());

        log(progress, 0.05, "Found:  " + consolidationFile.getName());
        log(progress, 0.05, "Found:  " + listingFile.getName());
        log(progress, 0.05, "Found:  " + tableauFile.getName());

        // ---- Read input files -------------------------------------------
        log(progress, 0.10, "Reading ConsolidationGenerale…");
        List<ConsolidationRow> allRows = reader.readAllConsolidationRows(consolidationFile);
        log(progress, 0.30, "  → " + allRows.size() + " rows read (incl. header)");

        log(progress, 0.35, "Reading Listing…");
        Map<String, ClientInfo> clientInfoMap = reader.readClientInfoMap(listingFile);
        log(progress, 0.45, "  → " + clientInfoMap.size() + " client entries");

        log(progress, 0.50, "Reading Tableau de Bord…");
        Map<String, Double> balances = reader.readPreviousBalances(tableauFile);
        log(progress, 0.55, "  → " + balances.size() + " balance entries");

        // ---- Calculate summaries ----------------------------------------
        log(progress, 0.60, "Computing TRF summaries…");
        List<ClientSummary> summaries = calculator.buildClientSummaries(allRows, clientInfoMap, balances);
        log(progress, 0.75, "  → " + summaries.size() + " client(s) included");

        if (summaries.isEmpty()) throw new IllegalStateException(
            "No clients found in ConsolidationGenerale. "
            + "Check that the 'Consolidation' sheet has data rows with client names in column A.");

        // Log which clients had no match in Listing or Tableau de Bord
        summaries.forEach(cs -> {
            if (cs.getClientCode().isBlank())
                log(progress, 0.75, "  [WARN] No Listing match for: " + cs.getClientName());
            if (cs.getNousDoit_Prec() == 0.0 && cs.getMontantAFacturerTtc() > 0)
                log(progress, 0.75, "  [INFO] No previous balance for: " + cs.getClientName());
        });

        // ---- Write output -----------------------------------------------
        LocalDate now      = LocalDate.now();
        String    month    = String.format("%02d", now.getMonthValue());
        String    year     = String.valueOf(now.getYear());
        String    outName  = "TRF_" + month + "_" + year + ".xlsx";
        File      outFile  = new File(outputFolder, outName);

        log(progress, 0.80, "Writing " + outName + "…");
        writer.write(allRows, summaries, outFile);
        log(progress, 1.00, "Done.  Output: " + outFile.getAbsolutePath());

        return outFile;
    }

    // -------------------------------------------------------------------------
    // File finder — searches folder (then one level down) case-insensitively
    // -------------------------------------------------------------------------

    private File findFile(File folder, String normalizedKeyword) {
        // Direct children first
        File[] files = folder.listFiles();
        if (files != null) {
            for (File f : files) {
                if (f.isFile() && normalize(f.getName()).contains(normalizedKeyword)) return f;
            }
        }
        // One level deeper
        if (files != null) {
            for (File sub : files) {
                if (!sub.isDirectory()) continue;
                File[] subFiles = sub.listFiles();
                if (subFiles == null) continue;
                for (File f : subFiles) {
                    if (f.isFile() && normalize(f.getName()).contains(normalizedKeyword)) return f;
                }
            }
        }
        return null;
    }

    private static String normalize(String s) {
        return java.text.Normalizer.normalize(s.toLowerCase(), java.text.Normalizer.Form.NFD)
            .replaceAll("\\p{M}", "")
            .replaceAll("[^a-z0-9]", "");
    }

    private void log(BiConsumer<Double, String> cb, double p, String msg) {
        cb.accept(p, msg);
    }
}
