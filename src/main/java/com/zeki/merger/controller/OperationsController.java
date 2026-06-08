package com.zeki.merger.controller;

import com.zeki.merger.AppPreferences;
import com.zeki.merger.service.*;
import com.zeki.merger.ui.AccuseReceptionDialog;
import com.zeki.merger.ui.DateRangeDialog;
import com.zeki.merger.ui.DateRangeDialog.DateRange;
import com.zeki.merger.ui.FacturationMailDialog;
import javafx.stage.Stage;
import com.zeki.merger.trf.TrfGeneratorService;
import javafx.application.Platform;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
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
    private Button validationClientsBtn;
    private Button recupInfoClientsBtn;
    private Button misAJourListingBtn;
    private Button accuseReceptionBtn;
    private Button facturationMailBtn;
    private Button tousLesDocsiersBtn;

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
        runActionBtn       = createActionBtn("▶  CONSOLIDER",     "Lire les états → ConsolidationGénérale", "consolider-card", e -> showConfirmDialog("Consolider", "Lit tous les états de créances et génère la ConsolidationGénérale.xlsx.", new String[]{"Dossier racine"}, new boolean[]{!AppPreferences.getMergeRoot().isBlank()}, this::run));
        tousLesDocsiersBtn = createActionBtn("Tous les dossiers",   "Conso sans filtre Lieu — choisir une période",      "action-card",         e -> openTousLesDossiers());
        trfBtn           = createActionBtn("Générer TRF",                "Calcul virements et compensations",       "action-card-primary", e -> showConfirmDialog("Générer TRF", "Calcule les virements et compensations depuis la ConsolidationGénérale.", new String[]{"Dossier racine", "ConsolidationGénérale", "Listing Cabinet", "Tableau de bord"}, new boolean[]{!AppPreferences.getMergeRoot().isBlank(), !AppPreferences.getTrfConso().isBlank(), !AppPreferences.getTrfListing().isBlank(), !AppPreferences.getTrfTableau().isBlank()}, this::generateTrf));
        etatBtn          = createActionBtn("États publics",              "Exporter vers Espace Partagé",            "action-card",         e -> showConfirmDialog("États publics", "Exporte les états publics vers l'Espace Partagé de chaque client.", new String[]{"Dossier racine"}, new boolean[]{!AppPreferences.getMergeRoot().isBlank()}, this::generateEtatPublic));
        cmpBtn           = createActionBtn("Comparer fichiers",          "Contrôle PROCRÉANCES",                    "action-card",         e -> showConfirmDialog("Comparer fichiers", "Compare le fichier Procréances avec la ConsolidationGénérale.", new String[]{"Dossier racine", "Procréances", "ConsolidationGénérale"}, new boolean[]{!AppPreferences.getMergeRoot().isBlank(), !AppPreferences.getProcreancesPath().isBlank(), !AppPreferences.getTrfConso().isBlank()}, this::compareProcreances));
        misAJourListingBtn = createActionBtn("Mis à jour Listing",       "Dernier dossier arrivé → Listing client", "action-card",         e -> showConfirmDialog("Mis à jour Listing", "Met à jour le Listing Cabinet avec les infos du dernier dossier arrivé.", new String[]{"Dossier racine", "Listing Cabinet"}, new boolean[]{!AppPreferences.getMergeRoot().isBlank(), !AppPreferences.getTrfListing().isBlank()}, this::misAJourListing));

        recupBtn         = createActionBtn("Récup. n° factures",         "Depuis RecupNumFacture",                  "action-card",         e -> showConfirmDialog("Récup. n° factures", "Lit les numéros de facture et les écrit dans chaque dossier client.", new String[]{"Dossier racine", "RecupNumFacture (requis)", "Tableau de bord (optionnel)"}, new boolean[]{!AppPreferences.getMergeRoot().isBlank(), !AppPreferences.getRecupFacturePath().isBlank(), true}, this::recupNumFacture));
        genControleBtn   = createActionBtn("Générer Contrôle Fact.",     "Produit Controle_Facturation.xlsx",       "action-card",         e -> showConfirmDialog("Générer Contrôle Fact.", "Produit le fichier Controle_Facturation.xlsx depuis les données consolidées.", new String[]{"Dossier racine", "Listing Cabinet"}, new boolean[]{!AppPreferences.getMergeRoot().isBlank(), !AppPreferences.getTrfListing().isBlank()}, this::genererControleFacturation));
        controleBtn      = createActionBtn("Contrôle Facturation",       "Comparer Contrôle vs Consolidation",      "action-card",         e -> showConfirmDialog("Contrôle Facturation", "Compare le fichier Contrôle Facturation avec la ConsolidationGénérale.", new String[]{"Dossier racine", "Contrôle Facturation xlsx", "ConsolidationGénérale"}, new boolean[]{!AppPreferences.getMergeRoot().isBlank(), !AppPreferences.getControlePath().isBlank(), !AppPreferences.getTrfConso().isBlank()}, this::compareConsoControle));
        factureBtn       = createActionBtn("Générer factures PDF",       "Export → nos dossiers",                   "action-card",         e -> showConfirmDialog("Générer factures PDF", "Génère les PDFs de facturation dans nos dossiers.", new String[]{"Dossier racine", "RecupNumFacture (optionnel)"}, new boolean[]{!AppPreferences.getMergeRoot().isBlank(), true}, () -> genererFacturesPdf(FacturePdfService.Mode.OWN)));
        factureClientBtn = createActionBtn("Factures → Espace partagé", "Export → espace partagé client",          "action-card",         e -> showConfirmDialog("Factures → Espace partagé", "Copie les factures PDF vers l'espace partagé de chaque client.", new String[]{"Dossier racine", "RecupNumFacture (optionnel)"}, new boolean[]{!AppPreferences.getMergeRoot().isBlank(), true}, () -> genererFacturesPdf(FacturePdfService.Mode.CLIENT)));
        validationClientsBtn = createActionBtn("Validation des clients", "Clôture mensuelle AG → I, reset R/S/T", "action-card", e -> showConfirmDialog("Validation des clients", "Transfère Recouvré total (U) vers Recouvré et Facturé (I) pour les lignes AG, puis remet à zéro les colonnes R, S, T.", new String[]{"Dossier racine"}, new boolean[]{!AppPreferences.getMergeRoot().isBlank()}, this::validationClients));

        recupInfoClientsBtn = createActionBtn("Récup. Info Clients",     "TVA + Infos → Etat de créances",          "action-card",         e -> showConfirmDialog("Récup. Info Clients", "Récupère les informations TVA et coordonnées depuis le Listing.", new String[]{"Dossier racine", "Listing Cabinet"}, new boolean[]{!AppPreferences.getMergeRoot().isBlank(), !AppPreferences.getTrfListing().isBlank()}, this::recupInfoClients));
        accuseReceptionBtn  = createActionBtn("Accusés de réception",    "Créer drafts mail avec état en PJ",       "action-card",         e -> openAccuseReceptionDialog());
        facturationMailBtn  = createActionBtn("Facturation mails",       "Envoyer les factures par mail",           "action-card",         e -> openFacturationMailDialog());
        syncDbBtn        = createActionBtn("Sync sociétés",              "Synchroniser toutes les sociétés",        "action-card",         e -> showConfirmDialog("Sync sociétés", "Synchronise toutes les sociétés dans la base de données locale.", new String[]{"Dossier racine"}, new boolean[]{!AppPreferences.getMergeRoot().isBlank()}, this::syncDatabase));
        fixBtn           = createActionBtn("Corriger espaces",           "Mise à jour Espace Partagé",              "action-card",         e -> showConfirmDialog("Corriger espaces", "Corrige les chemins et met à jour les fichiers dans l'Espace Partagé.", new String[]{"Dossier racine"}, new boolean[]{!AppPreferences.getMergeRoot().isBlank()}, this::fixPaths));
        nettoyerBtn      = createActionBtn("Nettoyer Espace Partagé",   "Supprimer PDF/XLS des espaces partagés",  "action-card-danger",  e -> showConfirmDialog("Nettoyer Espace Partagé", "Supprime définitivement les PDF et XLS des espaces partagés. Action irréversible.", new String[]{"Dossier racine"}, new boolean[]{!AppPreferences.getMergeRoot().isBlank()}, this::nettoyerEspacePartage));

        Label quotidienLabel = new Label("OPÉRATIONS QUOTIDIENNES");
        quotidienLabel.getStyleClass().add("section-label");
        GridPane.setColumnSpan(quotidienLabel, 2);

        Label factLabel = new Label("FACTURATION");
        factLabel.getStyleClass().add("section-label");
        GridPane.setColumnSpan(factLabel, 2);

        Label utilLabel = new Label("UTILITAIRES");
        utilLabel.getStyleClass().add("section-label");
        GridPane.setColumnSpan(utilLabel, 2);

        actionsGrid.add(quotidienLabel,    0, 0);
        actionsGrid.add(runActionBtn,         0, 1);
        actionsGrid.add(tousLesDocsiersBtn,   1, 1);
        actionsGrid.add(trfBtn,            0, 2);
        actionsGrid.add(etatBtn,           1, 2);
        actionsGrid.add(cmpBtn,            0, 3);
        actionsGrid.add(misAJourListingBtn,1, 3);

        actionsGrid.add(factLabel,         0, 4);
        actionsGrid.add(recupBtn,          0, 5);
        actionsGrid.add(genControleBtn,    1, 5);
        actionsGrid.add(controleBtn,       0, 6);
        actionsGrid.add(factureBtn,        1, 6);
        actionsGrid.add(factureClientBtn,  0, 7); GridPane.setColumnSpan(factureClientBtn, 2);
        actionsGrid.add(validationClientsBtn, 0, 8); GridPane.setColumnSpan(validationClientsBtn, 2);

        actionsGrid.add(utilLabel,           0, 9);
        actionsGrid.add(recupInfoClientsBtn, 0, 10);
        actionsGrid.add(syncDbBtn,           1, 10);
        actionsGrid.add(fixBtn,              0, 11);
        actionsGrid.add(nettoyerBtn,         1, 11);
        actionsGrid.add(accuseReceptionBtn,  0, 12); GridPane.setColumnSpan(accuseReceptionBtn, 2);
        actionsGrid.add(facturationMailBtn,  0, 13); GridPane.setColumnSpan(facturationMailBtn, 2);
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
        statusBar.setManaged(false);
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
                        openFileBtn.setManaged(true);
                        statusBar.setVisible(true);
                        statusBar.setManaged(true);
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
        statusBar.setManaged(false);
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
                    openFileBtn.setManaged(false);
                    statusBar.setVisible(true);
                    statusBar.setManaged(true);
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
        statusBar.setManaged(false);
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
                    openFileBtn.setManaged(true);
                    statusBar.setVisible(true);
                    statusBar.setManaged(true);
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
        if (!rootFolder.isDirectory()) { log.accept("ERROR: Dossier source introuvable — " + rootFolder.getAbsolutePath()); return; }
        if (!outputFolder.isDirectory()) { log.accept("ERROR: Dossier sortie introuvable — " + outputFolder.getAbsolutePath()); return; }
        setAllButtonsDisabled(true);
        statusBar.setVisible(false); statusBar.setManaged(false);
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
                        openFileBtn.setVisible(true); openFileBtn.setManaged(true);
                        statusBar.setVisible(true); statusBar.setManaged(true);
                        onDashboardRefresh.run();
                    }
                    setAllButtonsDisabled(false);
                });
            } catch (Exception e) {
                Platform.runLater(() -> { log.accept("FATAL: " + e.getMessage()); setAllButtonsDisabled(false); });
            }
        });
    }

    private void openTousLesDossiers() {
        Stage owner = (Stage) tousLesDocsiersBtn.getScene().getWindow();
        new DateRangeDialog(owner, "Tous les dossiers — Période")
            .showAndWait()
            .ifPresent(range -> {
                File rootFolder   = new File(AppPreferences.getMergeRoot());
                File outputFolder = new File(AppPreferences.getOutputFolder());
                if (!rootFolder.isDirectory()) { log.accept("ERROR: Dossier source introuvable — " + rootFolder.getAbsolutePath()); return; }
                if (!outputFolder.isDirectory()) { log.accept("ERROR: Dossier sortie introuvable — " + outputFolder.getAbsolutePath()); return; }
                setAllButtonsDisabled(true);
                statusBar.setVisible(false); statusBar.setManaged(false);
                progressBar.setProgress(0);
                logArea.clear();
                lastOutputFile = null;
                executor.submit(() -> {
                    try {
                        File result = mergeService.mergeTous(rootFolder, outputFolder,
                            range.dateDebut(), range.dateFin(),
                            (progress, msg) -> Platform.runLater(() -> { progressBar.setProgress(progress); log.accept(msg); }));
                        Platform.runLater(() -> {
                            if (result != null) {
                                lastOutputFile = result;
                                statusLabel.setText("Output: " + result.getAbsolutePath());
                                openFileBtn.setVisible(true); openFileBtn.setManaged(true);
                                statusBar.setVisible(true); statusBar.setManaged(true);
                                onDashboardRefresh.run();
                            }
                            setAllButtonsDisabled(false);
                        });
                    } catch (Exception e) {
                        Platform.runLater(() -> { log.accept("FATAL: " + e.getMessage()); setAllButtonsDisabled(false); });
                    }
                });
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
        statusBar.setManaged(false);
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
                        openFileBtn.setManaged(true);
                        statusBar.setVisible(true);
                        statusBar.setManaged(true);
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
        statusBar.setManaged(false);
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
                        openFileBtn.setManaged(true);
                        statusBar.setVisible(true);
                        statusBar.setManaged(true);
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
        statusBar.setManaged(false);
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
                    openFileBtn.setManaged(false);
                    statusBar.setVisible(true);
                    statusBar.setManaged(true);
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
        if (misAJourListingBtn     != null) misAJourListingBtn.setDisable(disabled);
        if (validationClientsBtn   != null) validationClientsBtn.setDisable(disabled);
        if (runActionBtn           != null) runActionBtn.setDisable(disabled);
        if (tousLesDocsiersBtn     != null) tousLesDocsiersBtn.setDisable(disabled);
        if (accuseReceptionBtn     != null) accuseReceptionBtn.setDisable(disabled);
        if (facturationMailBtn     != null) facturationMailBtn.setDisable(disabled);
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

        Dialog<java.time.LocalDate> dateDialog = new Dialog<>();
        dateDialog.setTitle("Date de facturation");
        dateDialog.setHeaderText("Choisissez la date à imprimer sur les factures :");

        DatePicker datePicker = new DatePicker(java.time.LocalDate.now());
        datePicker.setPromptText("jj/mm/aaaa");
        datePicker.setConverter(new javafx.util.StringConverter<java.time.LocalDate>() {
            private final java.time.format.DateTimeFormatter fmt =
                    java.time.format.DateTimeFormatter.ofPattern("dd/MM/yyyy");
            @Override public String toString(java.time.LocalDate d) { return d != null ? fmt.format(d) : ""; }
            @Override public java.time.LocalDate fromString(String s) {
                return (s != null && !s.isBlank()) ? java.time.LocalDate.parse(s, fmt) : null;
            }
        });

        VBox content = new VBox(8, new Label("Date :"), datePicker);
        content.setPadding(new Insets(10));
        dateDialog.getDialogPane().setContent(content);
        dateDialog.getDialogPane().getButtonTypes().addAll(ButtonType.OK, ButtonType.CANCEL);
        dateDialog.setResultConverter(bt -> bt == ButtonType.OK ? datePicker.getValue() : null);

        java.util.Optional<java.time.LocalDate> result = dateDialog.showAndWait();
        if (result.isEmpty() || result.get() == null) return;
        final java.time.LocalDate chosenDate = result.get();

        setAllButtonsDisabled(true);
        progressBar.setProgress(-1);
        String modeLabel = mode == FacturePdfService.Mode.OWN ? "nos dossiers" : "espace partagé client";
        log.accept("Génération PDF → " + modeLabel + "...");
        new Thread(() -> {
            try {
                java.util.List<String> lines = facturePdfService.apply(rootFolder, finalRecupFile, mode,
                        chosenDate,
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
            statusBar.setManaged(false);
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
                        openFileBtn.setManaged(false);
                        statusBar.setVisible(true);
                        statusBar.setManaged(true);
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
                String tableauPath = AppPreferences.getTableauBordPath();
                File tableauFile = (tableauPath != null && !tableauPath.isBlank()) ? new File(tableauPath) : null;
                File out = genControleService.apply(rootFolder, outputFolder, finalRecupFile, tableauFile,
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

    private void validationClients() {
        File rootFolder = new File(AppPreferences.getMergeRoot());
        if (!rootFolder.isDirectory()) {
            log.accept("ERROR: Dossier source introuvable — " + rootFolder.getAbsolutePath()); return;
        }
        setAllButtonsDisabled(true);
        statusBar.setVisible(false);
        statusBar.setManaged(false);
        progressBar.setProgress(0);
        logArea.clear();

        executor.submit(() -> {
            try {
                java.util.List<String> lines = new ValidationClientsService().apply(rootFolder,
                        (prog, msg) -> Platform.runLater(() -> { progressBar.setProgress(prog); log.accept(msg); }));
                Platform.runLater(() -> {
                    lines.forEach(log::accept);
                    statusLabel.setText("Validation terminée — " + lines.size() + " dossier(s).");
                    openFileBtn.setVisible(false);
                    openFileBtn.setManaged(false);
                    statusBar.setVisible(true);
                    statusBar.setManaged(true);
                    setAllButtonsDisabled(false);
                });
            } catch (Exception ex) {
                Platform.runLater(() -> { log.accept("FATAL: " + ex.getMessage()); setAllButtonsDisabled(false); });
            }
        });
    }

    private void openAccuseReceptionDialog() {
        new AccuseReceptionDialog(
                (Stage) accuseReceptionBtn.getScene().getWindow(),
                log
        ).show();
    }

    private void openFacturationMailDialog() {
        new FacturationMailDialog(
                (Stage) facturationMailBtn.getScene().getWindow(),
                log
        ).show();
    }

    private void showConfirmDialog(String title, String description,
                                   String[] fileNames, boolean[] filePresent,
                                   Runnable onConfirm) {
        Dialog<ButtonType> dlg = new Dialog<>();
        dlg.setTitle(title);
        dlg.setHeaderText(null);

        ButtonType lancerType = new ButtonType("Lancer", ButtonBar.ButtonData.OK_DONE);
        ButtonType cancelType = new ButtonType("Annuler", ButtonBar.ButtonData.CANCEL_CLOSE);
        dlg.getDialogPane().getButtonTypes().addAll(lancerType, cancelType);

        VBox content = new VBox(10);
        content.setPrefWidth(360);
        content.getStyleClass().add("confirm-dialog-content");

        Label descLabel = new Label(description);
        descLabel.setWrapText(true);
        descLabel.getStyleClass().add("confirm-dialog-desc");
        content.getChildren().add(descLabel);

        Label reqLabel = new Label("FICHIERS REQUIS");
        reqLabel.getStyleClass().add("confirm-dialog-section");
        content.getChildren().add(reqLabel);

        boolean allPresent = true;
        for (int i = 0; i < fileNames.length; i++) {
            boolean ok = filePresent[i];
            if (!ok) allPresent = false;
            HBox row = new HBox(8);
            row.setAlignment(Pos.CENTER_LEFT);
            Label icon = new Label(ok ? "✓" : "✗");
            icon.getStyleClass().add(ok ? "confirm-file-ok" : "confirm-file-missing");
            Label name = new Label(fileNames[i]);
            name.getStyleClass().add("confirm-file-name");
            row.getChildren().addAll(icon, name);
            content.getChildren().add(row);
        }

        dlg.getDialogPane().setContent(content);

        javafx.scene.Node lancerBtn = dlg.getDialogPane().lookupButton(lancerType);
        lancerBtn.setDisable(!allPresent);

        dlg.showAndWait().ifPresent(btn -> {
            if (btn == lancerType) onConfirm.run();
        });
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