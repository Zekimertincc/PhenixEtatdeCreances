package com.zeki.merger.service;

import com.zeki.merger.db.DatabaseManager;
import com.zeki.merger.model.CreanceRow;

import java.io.File;
import java.util.List;
import java.util.function.BiConsumer;

public class EtatCreancesSyncService {

    private final DatabaseManager db;
    private final ExcelReader    reader  = new ExcelReader();
    private final FolderScanner  scanner = new FolderScanner();

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
        List<CreanceRow> rows = reader.readFiltered(cf.companyName(), cf.excelFile());
        long companyId = db.upsertCompany(cf.companyName(), cf.excelFile().getAbsolutePath());
        db.replaceCreanceRows(companyId, rows);
    }

    private void log(BiConsumer<Double, String> cb, double pct, String msg) {
        if (cb != null) cb.accept(pct, msg);
    }
}
