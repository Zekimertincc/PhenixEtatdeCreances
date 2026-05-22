package com.zeki.merger.service;

import com.zeki.merger.service.FolderScanner;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.function.BiConsumer;

public class EspacePartageCleanerService {

    private final FolderScanner scanner;

    public EspacePartageCleanerService(FolderScanner scanner) {
        this.scanner = scanner;
    }

    public List<String> clean(File rootFolder, BiConsumer<Double, String> progress) {
        List<String> log = new ArrayList<>();
        List<FolderScanner.CompanyFile> companies = scanner.scan(rootFolder);

        if (companies.isEmpty()) {
            progress.accept(1.0, "Aucun dossier trouvé.");
            return log;
        }

        int total = companies.size();
        int deleted = 0;

        for (int i = 0; i < total; i++) {
            FolderScanner.CompanyFile cf = companies.get(i);
            double prog = 0.05 + 0.95 * (i + 1.0) / total;

            File espacePartage = findEspacePartage(cf.excelFile().getParentFile());
            if (espacePartage == null || !espacePartage.exists()) {
                progress.accept(prog, cf.companyName() + " → Espace partagé introuvable");
                continue;
            }

            File[] files = espacePartage.listFiles(f ->
                f.isFile() && (
                    f.getName().toLowerCase().endsWith(".pdf")  ||
                    f.getName().toLowerCase().endsWith(".xls")  ||
                    f.getName().toLowerCase().endsWith(".xlsx")
                )
            );

            if (files == null || files.length == 0) {
                progress.accept(prog, cf.companyName() + " → Rien à supprimer");
                continue;
            }

            int count = 0;
            for (File f : files) {
                if (f.delete()) {
                    count++;
                    deleted++;
                    log.add("Supprimé: " + f.getAbsolutePath());
                } else {
                    log.add("ÉCHEC suppression: " + f.getAbsolutePath());
                }
            }
            progress.accept(prog, cf.companyName() + " → " + count + " fichier(s) supprimé(s)");
        }

        progress.accept(1.0, "Nettoyage terminé — " + deleted + " fichier(s) supprimé(s) au total.");
        return log;
    }

    private File findEspacePartage(File companyDir) {
        if (companyDir == null) return null;
        File[] subs = companyDir.listFiles(f ->
            f.isDirectory() && f.getName().toLowerCase().contains("espace")
        );
        return (subs != null && subs.length > 0) ? subs[0] : null;
    }
}
