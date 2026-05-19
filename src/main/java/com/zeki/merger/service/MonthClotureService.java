package com.zeki.merger.service;

import com.zeki.merger.db.DatabaseManager;
import com.zeki.merger.db.TrfHistoryRecord;
import com.zeki.merger.db.TrfMonthRecord;
import com.zeki.merger.trf.TrfGeneratorService;
import com.zeki.merger.trf.model.ClientSummary;

import java.io.File;
import java.util.List;

public class MonthClotureService {

    private final DatabaseManager     db;
    private final TrfGeneratorService trfGenerator;

    public MonthClotureService(DatabaseManager db, TrfGeneratorService trfGenerator) {
        this.db          = db;
        this.trfGenerator = trfGenerator;
    }

    public void cloturerMois(int year, int month,
                             File consoFile, File listingFile, File tableauFile) throws Exception {
        saveMonth(year, month, "closed", consoFile, listingFile, tableauFile);
    }

    public void saveOpenMonth(int year, int month,
                              File consoFile, File listingFile, File tableauFile) throws Exception {
        saveMonth(year, month, "open", consoFile, listingFile, tableauFile);
    }

    private void saveMonth(int year, int month, String status,
                           File consoFile, File listingFile, File tableauFile) throws Exception {
        List<ClientSummary> summaries = trfGenerator.generateSummaries(consoFile, listingFile, tableauFile);

        double totalMontant  = summaries.stream().mapToDouble(ClientSummary::getMontantAFacturerTtc).sum();
        double totalNousDoit = summaries.stream().mapToDouble(ClientSummary::getNousDoit_Prec).sum();

        db.insertOrUpdateTrfMonth(year, month, status, summaries.size(), totalMontant, totalNousDoit);
        long monthId = db.getTrfMonthId(year, month);

        db.deleteTrfHistory(monthId);
        for (ClientSummary cs : summaries) {
            db.insertTrfHistory(monthId, cs);
        }
    }

    public List<TrfMonthRecord> getAllMonths() {
        return db.getAllTrfMonths();
    }

    public List<TrfHistoryRecord> getHistoryForMonth(long monthId) {
        return db.getTrfHistoryForMonth(monthId);
    }
}
