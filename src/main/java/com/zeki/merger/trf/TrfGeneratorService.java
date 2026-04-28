package com.zeki.merger.trf;

import com.zeki.merger.db.DatabaseManager;
import com.zeki.merger.trf.model.ClientInfo;
import com.zeki.merger.trf.model.ClientSummary;
import com.zeki.merger.trf.model.ConsolidationRow;

import java.io.File;
import java.time.LocalDate;
import java.util.List;
import java.util.Map;
import java.util.function.BiConsumer;

/**
 * Generates the TRF workbook from three explicitly-provided input files.
 * Output file: {@code TRF_MM_YYYY.xlsx} in the given output folder.
 */
public class TrfGeneratorService {

    private final DataReader      reader     = new DataReader();
    private final TrfCalculator   calculator = new TrfCalculator();
    private final TrfSheetWriter  writer     = new TrfSheetWriter();
    private final DatabaseManager db;

    public TrfGeneratorService(DatabaseManager db) {
        this.db = db;
    }

    /**
     * @param consoFile    ConsolidationGenerale Excel file
     * @param listingFile  Listing Cabinet Phénix Excel file
     * @param tableauFile  Tableau de Bord Excel file
     * @param outputFolder destination folder for TRF_MM_YYYY.xlsx
     * @param progress     callback (0..1, message) for UI feedback
     * @return the written output file
     */
    public File generate(File consoFile, File listingFile, File tableauFile,
                         File outputFolder, BiConsumer<Double, String> progress) throws Exception {

        log(progress, 0.00, "TRF Generator — starting");
        log(progress, 0.02, "  Consolidation : " + consoFile.getAbsolutePath());
        log(progress, 0.02, "  Listing       : " + listingFile.getAbsolutePath());
        log(progress, 0.02, "  Tableau       : " + tableauFile.getAbsolutePath());

        // ---- Read input files -------------------------------------------
        log(progress, 0.10, "Reading ConsolidationGenerale…");
        List<ConsolidationRow> allRows = reader.readAllConsolidationRows(consoFile);
        log(progress, 0.30, "  → " + allRows.size() + " rows read (incl. header)");

        log(progress, 0.35, "Reading Listing…");
        Map<String, ClientInfo> clientInfoMap = reader.readClientInfoMap(listingFile,
            msg -> log(progress, 0.35, msg));
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
            + "Check that the Consolidation sheet has data rows with client names in column A.");

        // Persist summaries to local DB
        if (db != null) {
            for (ClientSummary cs : summaries) {
                try {
                    long cid = db.upsertCompany(cs.getClientName(), null);
                    db.replaceTrfSummary(cid, cs);
                } catch (Exception dbEx) {
                    log(progress, 0.78, "  [DB] " + dbEx.getMessage());
                }
            }
        }

        // Warn about unmatched clients
        summaries.forEach(cs -> {
            if (cs.getClientCode().isBlank())
                log(progress, 0.75, "  [WARN] No Listing match for: " + cs.getClientName());
            if (cs.getNousDoit_Prec() == 0.0 && cs.getMontantAFacturerTtc() > 0)
                log(progress, 0.75, "  [INFO] No previous balance for: " + cs.getClientName());
        });

        // ---- Write output -----------------------------------------------
        LocalDate now     = LocalDate.now();
        String    outName = "TRF_" + String.format("%02d", now.getMonthValue())
                            + "_" + now.getYear() + ".xlsx";
        File      outFile = new File(outputFolder, outName);

        log(progress, 0.80, "Writing " + outName + "…");
        writer.write(allRows, summaries, outFile);
        log(progress, 1.00, "Done.  Output: " + outFile.getAbsolutePath());

        return outFile;
    }

    private void log(BiConsumer<Double, String> cb, double p, String msg) {
        cb.accept(p, msg);
    }
}
