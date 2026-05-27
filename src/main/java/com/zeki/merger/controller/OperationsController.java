package com.zeki.merger.controller;

import com.zeki.merger.AppPreferences;
import com.zeki.merger.service.ClientInfoService;
import com.zeki.merger.service.ConsoControleComparator;
import com.zeki.merger.service.EspacePartageFixer;
import com.zeki.merger.service.EspacePartageCleanerService;
import com.zeki.merger.service.EtatCreancesSyncService;
import com.zeki.merger.service.EtatPublicGenerator;
import com.zeki.merger.service.FacturePdfService;
import com.zeki.merger.service.GenererControleFacturationService;
import com.zeki.merger.service.MergeService;
import com.zeki.merger.service.MisAJourListingService;
import com.zeki.merger.service.ProcreancesComparator;
import com.zeki.merger.service.RecupNumFactureService;
import com.zeki.merger.trf.TrfGeneratorService;
import javafx.application.Platform;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.scene.control.*;
import javafx.scene.layout.*;

import java.awt.Desktop;
import java.io.File;
import java.util.concurrent.ExecutorService;
import java.util.function.Consumer;

public class OperationsController {

    private final MergeService            mergeService;
    private final EspacePartageFixer      espacePartageFixer;
    private final EtatPublicGenerator     etatPublicGenerator;
    private final TrfGeneratorService     trfGeneratorService;
    private final ProcreancesComparator   procreancesComparator;
    private final ConsoControleComparator consoControleComparator;
    private final RecupNumFactureService  recupNumFactureService;
    private final EtatCreancesSyncService syncService;
    private final FacturePdfService                 facturePdfService      = new FacturePdfService();
    private final MisAJourListingService             misAJourListingService = new MisAJourListingService();
    private final GenererControleFacturationService genControleService      = new GenererControleFacturationService();
    private final ClientInfoService                  clientInfoService      = new ClientInfoService();
    private final ProgressBar             progressBar;
    private final HBox                    statusBar;
    private final Label                   statusLabel;
    private final Button                  openFileBtn;
    private final TextArea                logArea;
    private final Consumer<String>        log;
    private final Runnable                onDashboardRefresh;
    private final ExecutorService         executor;

    private File   lastOutputFile;
    private Button trfBtn;
    private Button etatBtn;
    private Button cmpBtn;
    private Button fixBtn;
    private Button syncDbBtn;
    private Button recupBtn;
    private Button runActionBtn;
    private Button controleBtn;
    private Button factureBtn;
    private Button factureClientBtn;
    private Button genControleBtn;
    private Button nettoyerBtn;
    private Button recupInfoClientsBtn;
    private Button misAJourListingBtn;

    public OperationsController(MergeService mergeService,
                                EspacePartageFixer espacePartageFixer,
                                EtatPublicGenerator etatPublicGenerator,
                                TrfGeneratorService trfGeneratorService,
                                ProcreancesComparator procreancesComparator,
                                ConsoControleComparator consoControleComparator,
                                RecupNumFactureService recupNumFactureService,
                                EtatCreancesSyncService syncService,
                                ProgressBar progressBar,
                                HBox statusBar,
                                Label statusLabel,
                                Button openFileBtn,
                                TextArea logArea,
                                Consumer<String> log,
                                Runnable onDashboardRefresh,
                                ExecutorService executor) {
        this.mergeService            = mergeService;
        this.espacePartageFixer      = espacePartageFixer;
        this.etatPublicGenerator     = etatPublicGenerator;
        this.trfGeneratorService     = trfGeneratorService;
        this.procreancesComparator   = procreancesComparator;
        this.consoControleComparator = consoControleComparator;
        this.recupNumFactureService  = recupNumFactureService;
        this.syncService             = syncService;
        this.progressBar             = progressBar;
        this.statusBar               = statusBar;
        this.statusLabel             = statusLabel;
        this.openFileBtn             = openFileBtn;
        this.logArea                 = logArea;
        this.log                     = log;
        this.onDashboardRefresh      = onDashboardRefresh;
        this.executor                = executor;
    }

    public void buildButtons(GridPane actionsGrid) {
        trfBtn           = createActionBtn("Générer TRF",               "Calcul virements et compensations",       "action-card-primary", e -> generateTrf());
        etatBtn          = createActionBtn("États publics",              "Exporter vers Espace Partagé",            "action-card",         e -> generateEtatPublic());
        cmpBtn           = createActionBtn("Comparer fichiers",          "Contrôle PROCRÉANCES",                    "action-card",         e -> compareProcreances());
        fixBtn           = createActionBtn("Corriger espaces",           "Mise à jour Espace Partagé",               "action-card",         e -> fixPaths());
        syncDbBtn        = createActionBtn("Sync toutes sociétés",       "Synchroniser toutes les sociétés",         "action-card",         e -> syncDatabase());
        recupBtn         = createActionBtn("Récup. n° factures",         "Depuis Dropbox",                           "action-card",         e -> recupNumFacture());
        controleBtn      = createActionBtn("Contrôle Facturation",       "Comparer Contrôle vs Consolidation",       "action-card",         e -> compareConsoControle());
        factureBtn       = createActionBtn("Générer factures PDF",       "Export → nos dossiers",                    "action-card",         e -> genererFacturesPdf(FacturePdfService.Mode.OWN));
        factureClientBtn = createActionBtn("Factures → Espace partagé", "Export → espace partagé client",           "action-card",         e -> genererFacturesPdf(FacturePdfService.Mode.CLIENT));
        misAJourListingBtn   = createActionBtn("Mis à jour Listing",    "Dernier dossier arrivé → Listing client",  "action-card",         e -> misAJourListing());
        recupInfoClientsBtn  = createActionBtn("Récup. Info Clients",   "TVA + Infos → Etat de créances",           "action-card",         e -> recupInfoClients());
        genControleBtn   = createActionBtn("Générer Contrôle Fact.",     "Produit Controle_Facturation.xlsx",        "action-card",         e -> genererControleFacturation());
        nettoyerBtn      = createActionBtn("Nettoyer Espace Partagé",   "Supprimer PDF/XLS des espaces partagés",   "action-card-danger",  e -> nettoyerEspacePartage());
        runActionBtn     = createActionBtn("▶  CONSOLIDER",              "Lire les états → ConsolidationGénérale",   "consolider-card",     e -> run());

        Label opsLabel = new Label("OPÉRATIONS");
        opsLabel.getStyleClass().add("section-label");
        GridPane.setColumnSpan(opsLabel, 2);

        Label factLabel = new Label("FACTURATION");
        factLabel.getStyleClass().add("section-label");
        GridPane.setColumnSpan(factLabel, 2);

        actionsGrid.add(opsLabel,       0, 0);
        actionsGrid.add(trfBtn,         0, 1);
        actionsGrid.add(etatBtn,        1, 1);
        actionsGrid.add(cmpBtn,         0, 2);
        actionsGrid.add(fixBtn,         1, 2);
        actionsGrid.add(syncDbBtn,      0, 3);
        actionsGrid.add(nettoyerBtn,    1, 3);

        actionsGrid.add(factLabel,           0, 4);
        actionsGrid.add(recupBtn,            0, 5);
        actionsGrid.add(genControleBtn,      1, 5);
        actionsGrid.add(controleBtn,         0, 6);
        actionsGrid.add(factureBtn,          1, 6);
        actionsGrid.add(factureClientBtn,    0, 7);
        actionsGrid.add(misAJourListingBtn,  1, 7);
        actionsGrid.add(recupInfoClientsBtn, 0, 8);

        actionsGrid.add(runActionBtn,        0, 9);
        GridPane.setColumnSpan(runActionBtn, 2);
    }

    public void openFile() {
        if (lastOutputFile != null && lastOutputFile.exists()) {
            try { Desktop.getDesktop().open(lastOutputFile); }
            catch (Exception e) { log.accept("Cannot open file: " + e.getMessage()); }
        }
    }

    private void generateTrf() {
        String consoPath   = AppPreferences.getTrfConso();
        String listingPath = AppPreferences.getTrfListing();
        String tableauPath = AppPreferences.getTrfTableau();
        String outputPath  = AppPreferences.getOutputFolder();

        if (consoPath.isEmpty() || listingPath.isEmpty() || tableauPath.isEmpty()) {
            log.accept("ERROR: Configurez les trois fichiers TRF avant de générer."); return;
        }
        File consoFile    = new File(consoPath);
        File listingFile  = new File(listingPath);
        File tableauFile  = new File(tableauPath);
        File outputFolder = new File(outputPath);

        if (!consoFile.exists())         { log.accept("ERROR: Fichier introuvable — " + consoPath);   return; }
        if (!listingFile.exists())       { log.accept("ERROR: Fichier introuvable — " + listingPath); return; }
        if (!tableauFile.exists())       { log.accept("ERROR: Fichier introuvable — " + tableauPath); return; }
        if (!outputFolder.isDirectory()) { log.accept("ERROR: Dossier sortie introuvable — " + outputPath); return; }

        setAllButtonsDisabled(true);
        statusBar.setVisible(false);
        progressBar.setProgress(0);
        logArea.clear();
        lastOutputFile = null;

        executor.submit(() -> {
            try {
                File result = trfGeneratorService.generate(consoFile, listingFile, tableauFile, outputFolder,
                        (prog, msg) -> Platform.runLater(() -> { progressBar.setProgress(prog); log.accept(msg); }));
                Platform.runLater(() -> {
                    if (result != null) {
                        lastOutputFile = result;
                        statusLabel.setText("TRF Output: " + result.getAbsolutePath());
                        openFileBtn.setVisible(true);
                        statusBar.setVisible(true);
                        onDashboardRefresh.run();
                    }
                    setAllButtonsDisabled(false);
                });
            } catch (Exception e) {
                Platform.runLater(() -> { log.accept("FATAL: " + e.getMessage()); setAllButtonsDisabled(false); });
            }
        });
    }

    private void generateEtatPublic() {
        File rootFolder = new File(AppPreferences.getMergeRoot());
        if (!rootFolder.isDirectory()) {
            log.accept("ERROR: Dossier source introuvable — " + rootFolder.getAbsolutePath()); return;
        }
        setAllButtonsDisabled(true);
        statusBar.setVisible(false);
        progressBar.setProgress(0);
        logArea.clear();
        lastOutputFile = null;

        executor.submit(() -> {
            try {
                etatPublicGenerator.generate(rootFolder,
                        (prog, msg) -> Platform.runLater(() -> { progressBar.setProgress(prog); log.accept(msg); }));
                Platform.runLater(() -> {
                    statusLabel.setText("Etat Public files written to EspacePartagé paths.");
                    openFileBtn.setVisible(false);
                    statusBar.setVisible(true);
                    setAllButtonsDisabled(false);
                });
            } catch (Exception e) {
                Platform.runLater(() -> { log.accept("FATAL: " + e.getMessage()); setAllButtonsDisabled(false); });
            }
        });
    }

    private void fixPaths() {
        File rootFolder = new File(AppPreferences.getMergeRoot());
        if (!rootFolder.isDirectory()) {
            log.accept("ERROR: Dossier source introuvable — " + rootFolder.getAbsolutePath()); return;
        }
        setAllButtonsDisabled(true);
        statusBar.setVisible(false);
        progressBar.setProgress(0);
        logArea.clear();
        lastOutputFile = null;

        executor.submit(() -> {
            try {
                File result = espacePartageFixer.fix(rootFolder,
                        (progress, msg) -> Platform.runLater(() -> { progressBar.setProgress(progress); log.accept(msg); }));
                Platform.runLater(() -> {
                    lastOutputFile = result;
                    statusLabel.setText("Saved: " + result.getAbsolutePath());
                    openFileBtn.setVisible(true);
                    statusBar.setVisible(true);
                    setAllButtonsDisabled(false);
                });
            } catch (Exception e) {
                Platform.runLater(() -> { log.accept("FATAL: " + e.getMessage()); setAllButtonsDisabled(false); });
            }
        });
    }

    private void run() {
        File rootFolder   = new File(AppPreferences.getMergeRoot());
        File outputFolder = new File(AppPreferences.getOutputFolder());

        if (!rootFolder.isDirectory()) {
            log.accept("ERROR: Dossier source introuvable — " + rootFolder.getAbsolutePath()); return;
        }
        if (!outputFolder.isDirectory()) {
            log.accept("ERROR: Dossier sortie introuvable — " + outputFolder.getAbsolutePath()); return;
        }
        setAllButtonsDisabled(true);
        statusBar.setVisible(false);
        progressBar.setProgress(0);
        logArea.clear();
        lastOutputFile = null;

        executor.submit(() -> {
            try {
                File result = mergeService.merge(rootFolder, outputFolder,
                        (progress, msg) -> Platform.runLater(() -> { progressBar.setProgress(progress); log.accept(msg); }));
                Platform.runLater(() -> {
                    if (result != null) {
                        lastOutputFile = result;
                        statusLabel.setText("Output: " + result.getAbsolutePath());
                        openFileBtn.setVisible(true);
                        statusBar.setVisible(true);
                        onDashboardRefresh.run();
                    }
                    setAllButtonsDisabled(false);
                });
            } catch (Exception e) {
                Platform.runLater(() -> { log.accept("FATAL: " + e.getMessage()); setAllButtonsDisabled(false); });
            }
        });
    }

    private void compareProcreances() {
        String procPath   = AppPreferences.getProcreancesPath();
        String consoPath  = AppPreferences.getTrfConso();
        String outputPath = AppPreferences.getOutputFolder();

        if (procPath.isEmpty() || consoPath.isEmpty()) {
            log.accept("ERROR: Configurez Export PROCREANCES et ConsolidationGénérale avant de comparer."); return;
        }
        File procFile     = new File(procPath);
        File consoFile    = new File(consoPath);
        File outputFolder = new File(outputPath);

        if (!procFile.exists())          { log.accept("ERROR: Fichier introuvable — " + procPath);  return; }
        if (!consoFile.exists())         { log.accept("ERROR: Fichier introuvable — " + consoPath); return; }
        if (!outputFolder.isDirectory()) { log.accept("ERROR: Dossier sortie introuvable — " + outputPath); return; }

        setAllButtonsDisabled(true);
        statusBar.setVisible(false);
        progressBar.setProgress(0);
        logArea.clear();
        lastOutputFile = null;

        executor.submit(() -> {
            try {
                File report = procreancesComparator.compare(procFile, consoFile, outputFolder,
                        (prog, msg) -> Platform.runLater(() -> { progressBar.setProgress(prog); log.accept(msg); }));
                Platform.runLater(() -> {
                    setAllButtonsDisabled(false);
                    if (report != null) {
                        lastOutputFile = report;
                        statusLabel.setText("Rapport: " + report.getAbsolutePath());
                        openFileBtn.setVisible(true);
                        statusBar.setVisible(true);
                        try { Desktop.getDesktop().open(report); } catch (Exception ignored) {}
                    }
                });
            } catch (Exception ex) {
                Platform.runLater(() -> { log.accept("FATAL: " + ex.getMessage()); setAllButtonsDisabled(false); });
            }
        });
    }

    private void compareConsoControle() {
        String controlePath = AppPreferences.getControlePath();
        String consoPath    = AppPreferences.getTrfConso();
        String outputPath   = AppPreferences.getOutputFolder();

        if (controlePath.isEmpty() || consoPath.isEmpty()) {
            log.accept("ERROR: Configurez Contrôle Facturation et ConsolidationGénérale avant de comparer."); return;
        }
        File controleFile  = new File(controlePath);
        File consoFile     = new File(consoPath);
        File outputFolder  = new File(outputPath);

        if (!controleFile.exists())      { log.accept("ERROR: Fichier introuvable — " + controlePath);  return; }
        if (!consoFile.exists())         { log.accept("ERROR: Fichier introuvable — " + consoPath);      return; }
        if (!outputFolder.isDirectory()) { log.accept("ERROR: Dossier sortie introuvable — " + outputPath); return; }

        setAllButtonsDisabled(true);
        statusBar.setVisible(false);
        progressBar.setProgress(0);
        logArea.clear();
        lastOutputFile = null;

        executor.submit(() -> {
            try {
                File report = consoControleComparator.compare(controleFile, consoFile, outputFolder,
                        (prog, msg) -> Platform.runLater(() -> { progressBar.setProgress(prog); log.accept(msg); }));
                Platform.runLater(() -> {
                    setAllButtonsDisabled(false);
                    if (report != null) {
                        lastOutputFile = report;
                        statusLabel.setText("Rapport: " + report.getAbsolutePath());
                        openFileBtn.setVisible(true);
                        statusBar.setVisible(true);
                        try { Desktop.getDesktop().open(report); } catch (Exception ignored) {}
                    }
                });
            } catch (Exception ex) {
                Platform.runLater(() -> { log.accept("FATAL: " + ex.getMessage()); setAllButtonsDisabled(false); });
            }
        });
    }

    private void recupNumFacture() {
        String recupPath   = AppPreferences.getRecupFacturePath();
        String rootPath    = AppPreferences.getMergeRoot();
        String tableauPath = AppPreferences.getTableauBordPath();

        if (recupPath.isEmpty()) {
            log.accept("ERROR: Configurez le fichier Récup. Num Facture avant de lancer."); return;
        }
        File recupFile  = new File(recupPath);
        File rootFolder = new File(rootPath);

        if (!recupFile.exists())       { log.accept("ERROR: Fichier introuvable — " + recupPath); return; }
        if (!rootFolder.isDirectory()) { log.accept("ERROR: Dossier source introuvable — " + rootPath); return; }

        File tableauFile = (tableauPath != null && !tableauPath.isBlank()) ? new File(tableauPath) : null;
        if (tableauFile != null && !tableauFile.exists()) {
            log.accept("AVERT: Tableau de bord introuvable — " + tableauPath + ". Soldes non appliqués.");
            tableauFile = null;
        }
        final File finalTableauFile = tableauFile;

        setAllButtonsDisabled(true);
        statusBar.setVisible(false);
        progressBar.setProgress(0);
        logArea.clear();
        lastOutputFile = null;

        executor.submit(() -> {
            try {
                java.util.List<String> logLines = recupNumFactureService.apply(recupFile, rootFolder, finalTableauFile,
                        (prog, msg) -> Platform.runLater(() -> { progressBar.setProgress(prog); log.accept(msg); }));
                Platform.runLater(() -> {
                    logLines.forEach(log::accept);
                    statusLabel.setText("Récup. Factures terminée — " + logLines.size() + " dossier(s).");
                    openFileBtn.setVisible(false);
                    statusBar.setVisible(true);
                    setAllButtonsDisabled(false);
                });
            } catch (Exception ex) {
                Platform.runLater(() -> { log.accept("FATAL: " + ex.getMessage()); setAllButtonsDisabled(false); });
            }
        });
    }

    private void recupInfoClients() {
        String listingPath     = AppPreferences.getTrfListing();
        String rootPath        = AppPreferences.getMergeRoot();
        String procreancesPath = AppPreferences.getProcreancesPath();
        if (listingPath.isBlank()) { log.accept("ERROR: Fichier Listing non configuré."); return; }
        File listingFile     = new File(listingPath);
        File rootFolder      = new File(rootPath);
        File procreancesFile = (procreancesPath != null && !procreancesPath.isBlank()) ? new File(procreancesPath) : null;
        if (!listingFile.exists())     { log.accept("ERROR: Listing introuvable."); return; }
        if (!rootFolder.isDirectory()) { log.accept("ERROR: Dossier source introuvable."); return; }

        setAllButtonsDisabled(true);
        progressBar.setProgress(-1);
        log.accept("Récup. Info Clients...");
        final File finalProc = procreancesFile;
        new Thread(() -> {
            try {
                java.util.List<String> lines = clientInfoService.apply(listingFile, rootFolder, finalProc,
                        (p, msg) -> Platform.runLater(() -> { progressBar.setProgress(p); log.accept(msg); }));
                Platform.runLater(() -> {
                    progressBar.setProgress(1.0);
                    statusLabel.setText("Info Clients terminée — " + lines.size() + " dossier(s).");
                    statusBar.setVisible(true);
                    statusBar.setManaged(true);
                    setAllButtonsDisabled(false);
                });
            } catch (Exception e) {
                Platform.runLater(() -> { log.accept("ERREUR: " + e.getMessage()); setAllButtonsDisabled(false); });
            }
        }, "recup-info-clients-thread").start();
    }

    private void syncDatabase() {
        File root = new File(AppPreferences.getMergeRoot());
        if (!root.isDirectory()) {
            log.accept("ERROR: Dossier source introuvable. Configurez le chemin."); return;
        }
        syncDbBtn.setDisable(true);
        log.accept("Synchronisation DB en cours…");
        executor.submit(() -> {
            try {
                syncService.syncAll(root, (pct, msg) ->
                        Platform.runLater(() -> { progressBar.setProgress(pct); log.accept(msg); }));
            } catch (Exception e) {
                Platform.runLater(() -> log.accept("ERREUR sync : " + e.getMessage()));
            } finally {
                Platform.runLater(() -> {
                    syncDbBtn.setDisable(false);
                    onDashboardRefresh.run();
                });
            }
        });
    }

    private void setAllButtonsDisabled(boolean disabled) {
        if (trfBtn              != null) trfBtn.setDisable(disabled);
        if (etatBtn             != null) etatBtn.setDisable(disabled);
        if (cmpBtn              != null) cmpBtn.setDisable(disabled);
        if (fixBtn              != null) fixBtn.setDisable(disabled);
        if (controleBtn         != null) controleBtn.setDisable(disabled);
        if (recupBtn            != null) recupBtn.setDisable(disabled);
        if (syncDbBtn           != null) syncDbBtn.setDisable(disabled);
        if (factureBtn          != null) factureBtn.setDisable(disabled);
        if (factureClientBtn    != null) factureClientBtn.setDisable(disabled);
        if (genControleBtn      != null) genControleBtn.setDisable(disabled);
        if (nettoyerBtn         != null) nettoyerBtn.setDisable(disabled);
        if (recupInfoClientsBtn != null) recupInfoClientsBtn.setDisable(disabled);
        if (misAJourListingBtn  != null) misAJourListingBtn.setDisable(disabled);
        if (runActionBtn        != null) runActionBtn.setDisable(disabled);
    }

    private void misAJourListing() {
        String listingPath = AppPreferences.getTrfListing();
        String rootPath    = AppPreferences.getMergeRoot();
        if (listingPath.isBlank()) { log.accept("ERROR: Fichier Listing non configuré."); return; }
        File listingFile = new File(listingPath);
        File rootFolder  = new File(rootPath);
        if (!listingFile.exists())     { log.accept("ERROR: Listing introuvable — " + listingPath); return; }
        if (!rootFolder.isDirectory()) { log.accept("ERROR: Dossier source introuvable."); return; }

        setAllButtonsDisabled(true);
        progressBar.setProgress(-1);
        log.accept("Mis à jour Listing — Dernier dossier arrivé...");
        new Thread(() -> {
            try {
                java.util.List<String> lines = misAJourListingService.apply(listingFile, rootFolder,
                        (p, msg) -> Platform.runLater(() -> { progressBar.setProgress(p); log.accept(msg); }));
                Platform.runLater(() -> {
                    progressBar.setProgress(1.0);
                    statusLabel.setText("Listing mis à jour — " + lines.size() + " dossier(s) traités.");
                    statusBar.setVisible(true);
                    statusBar.setManaged(true);
                    setAllButtonsDisabled(false);
                });
            } catch (Exception e) {
                Platform.runLater(() -> { log.accept("ERREUR: " + e.getMessage()); setAllButtonsDisabled(false); });
            }
        }, "mis-a-jour-listing-thread").start();
    }

    private void genererFacturesPdf(FacturePdfService.Mode mode) {
        String rootPath  = AppPreferences.getMergeRoot();
        String recupPath = AppPreferences.getRecupFacturePath();
        File rootFolder  = new File(rootPath);
        if (!rootFolder.isDirectory()) {
            log.accept("ERROR: Dossier source non configuré."); return;
        }
        File recupFile = (recupPath != null && !recupPath.isBlank()) ? new File(recupPath) : null;
        if (recupFile != null && !recupFile.exists()) {
            log.accept("AVERT: Récup. Num Facture introuvable — " + recupPath + ". Génération sans filtre.");
            recupFile = null;
        }
        final File finalRecupFile = recupFile;
        setAllButtonsDisabled(true);
        progressBar.setProgress(-1);
        String modeLabel = mode == FacturePdfService.Mode.OWN ? "nos dossiers" : "espace partagé client";
        log.accept("Génération PDF → " + modeLabel + "...");
        new Thread(() -> {
            try {
                java.util.List<String> lines = facturePdfService.apply(rootFolder, finalRecupFile, mode,
                        (p, msg) -> Platform.runLater(() -> { progressBar.setProgress(p); log.accept(msg); }));
                Platform.runLater(() -> {
                    progressBar.setProgress(1.0);
                    statusLabel.setText("Factures PDF générées [" + modeLabel + "] — " + lines.size() + " dossier(s).");
                    statusBar.setVisible(true);
                    statusBar.setManaged(true);
                    setAllButtonsDisabled(false);
                });
            } catch (Exception e) {
                Platform.runLater(() -> { log.accept("ERREUR: " + e.getMessage()); setAllButtonsDisabled(false); });
            }
        }, "facture-pdf-thread").start();
    }

    private void nettoyerEspacePartage() {
        File rootFolder = new File(AppPreferences.getMergeRoot());
        if (!rootFolder.isDirectory()) {
            log.accept("ERROR: Dossier source introuvable — " + rootFolder.getAbsolutePath()); return;
        }

        Alert confirm = new Alert(Alert.AlertType.CONFIRMATION);
        confirm.setTitle("Confirmation");
        confirm.setHeaderText("Nettoyer les Espaces Partagés");
        confirm.setContentText(
                "Cette action supprimera tous les fichiers PDF, XLS et XLSX\n" +
                        "dans les dossiers 'Espace partagé' de chaque client.\n\n" +
                        "Cette action est irréversible. Continuer ?");
        confirm.showAndWait().ifPresent(btn -> {
            if (btn != ButtonType.OK) return;

            setAllButtonsDisabled(true);
            statusBar.setVisible(false);
            progressBar.setProgress(0);
            logArea.clear();
            lastOutputFile = null;

            executor.submit(() -> {
                try {
                    java.util.List<String> lines = new EspacePartageCleanerService(new com.zeki.merger.service.FolderScanner())
                            .clean(rootFolder, (prog, msg) -> Platform.runLater(() -> {
                                progressBar.setProgress(prog);
                                log.accept(msg);
                            }));
                    Platform.runLater(() -> {
                        lines.forEach(log::accept);
                        statusLabel.setText("Nettoyage terminé — " + lines.size() + " suppression(s).");
                        openFileBtn.setVisible(false);
                        statusBar.setVisible(true);
                        setAllButtonsDisabled(false);
                    });
                } catch (Exception ex) {
                    Platform.runLater(() -> { log.accept("FATAL: " + ex.getMessage()); setAllButtonsDisabled(false); });
                }
            });
        });
    }

    private void genererControleFacturation() {
        String rootPath   = AppPreferences.getMergeRoot();
        String outputPath = AppPreferences.getOutputFolder();
        String recupPath  = AppPreferences.getRecupFacturePath();
        File rootFolder   = new File(rootPath);
        File outputFolder = new File(outputPath.isBlank() ? rootPath : outputPath);
        if (!rootFolder.isDirectory()) {
            log.accept("ERROR: Dossier source non configuré."); return;
        }
        File recupFile = (recupPath != null && !recupPath.isBlank()) ? new File(recupPath) : null;
        if (recupFile != null && !recupFile.exists()) {
            log.accept("AVERT: Récup. Num Facture introuvable — " + recupPath + ". Génération sans filtre.");
            recupFile = null;
        }
        final File finalRecupFile = recupFile;
        setAllButtonsDisabled(true);
        progressBar.setProgress(0);
        new Thread(() -> {
            try {
                File out = genControleService.apply(rootFolder, outputFolder, finalRecupFile,
                        (p, msg) -> Platform.runLater(() -> { progressBar.setProgress(p); log.accept(msg); }));
                Platform.runLater(() -> {
                    progressBar.setProgress(1.0);
                    if (out != null) openFile(out);
                    setAllButtonsDisabled(false);
                });
            } catch (Exception e) {
                Platform.runLater(() -> { log.accept("ERREUR: " + e.getMessage()); setAllButtonsDisabled(false); });
            }
        }, "gen-controle-thread").start();
    }

    private void openFile(File f) {
        if (f != null && f.exists()) {
            try { Desktop.getDesktop().open(f); }
            catch (Exception e) { log.accept("Cannot open file: " + e.getMessage()); }
        }
    }

    private Button createActionBtn(String name, String desc, String styleClass,
                                   EventHandler<ActionEvent> handler) {
        String titleClass, subtitleClass;
        switch (styleClass) {
            case "action-card-primary" -> { titleClass = "action-card-title-primary"; subtitleClass = "action-card-subtitle-primary"; }
            case "action-card-danger"  -> { titleClass = "action-card-title-danger";  subtitleClass = "action-card-subtitle-danger"; }
            case "consolider-card"     -> { titleClass = "consolider-title";          subtitleClass = "consolider-subtitle"; }
            default                    -> { titleClass = "action-card-title";         subtitleClass = "action-card-subtitle"; }
        }
        Label lName = new Label(name);
        lName.getStyleClass().add(titleClass);
        Label lDesc = new Label(desc);
        lDesc.getStyleClass().add(subtitleClass);
        VBox vb = new VBox(2, lName, lDesc);
        Button btn = new Button();
        btn.setGraphic(vb);
        btn.getStyleClass().add(styleClass);
        btn.setMaxWidth(Double.MAX_VALUE);
        btn.setMaxHeight(Double.MAX_VALUE);
        btn.setOnAction(handler);
        return btn;
    }
}