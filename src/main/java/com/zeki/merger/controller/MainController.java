package com.zeki.merger.controller;

import com.zeki.merger.AppPreferences;
import com.zeki.merger.db.DatabaseManager;
import com.zeki.merger.db.TrfHistoryRecord;
import com.zeki.merger.db.TrfMonthRecord;
import com.zeki.merger.service.ClientInfoService;
import com.zeki.merger.service.ConsoControleComparator;
import com.zeki.merger.service.EspacePartageFixer;
import com.zeki.merger.service.EtatCreancesSyncService;
import com.zeki.merger.service.EtatPublicGenerator;
import com.zeki.merger.service.MergeService;
import com.zeki.merger.service.MonthClotureService;
import com.zeki.merger.service.ProcreancesComparator;
import com.zeki.merger.service.RecupNumFactureService;
import com.zeki.merger.trf.TrfGeneratorService;
import javafx.application.Platform;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.chart.BarChart;
import javafx.scene.chart.CategoryAxis;
import javafx.scene.chart.NumberAxis;
import javafx.scene.chart.XYChart;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Modality;
import javafx.stage.Stage;

import java.awt.Desktop;
import java.io.File;
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.time.format.TextStyle;
import java.util.List;
import java.util.Locale;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class MainController {

    // =========================================================================
    // FXML — sidebar nav
    // =========================================================================

    @FXML private Button navOperations;
    @FXML private Button navDashboard;
    @FXML private Button navHistorique;
    @FXML private Button navConfig;
    @FXML private Button navLogFull;
    @FXML private Label  pageTitle;

    // =========================================================================
    // FXML — pages
    // =========================================================================

    @FXML private VBox pageOperations;
    @FXML private VBox pageDashboard;
    @FXML private VBox pageHistorique;
    @FXML private VBox pageConfig;

    // =========================================================================
    // FXML — Opérations page
    // =========================================================================

    @FXML private FlowPane    badgesBox;
    @FXML private Label       missingFilesLabel;
    @FXML private GridPane    actionsGrid;
    @FXML private ProgressBar progressBar;
    @FXML private TextArea    logArea;
    @FXML private HBox        statusBar;
    @FXML private Label       statusLabel;
    @FXML private Button      openFileBtn;

    // =========================================================================
    // FXML — Configuration page
    // =========================================================================

    @FXML private VBox configFormBox;

    // =========================================================================
    // Action buttons (Opérations grid — created programmatically)
    // =========================================================================

    private Button trfBtn;
    private Button etatBtn;
    private Button cmpBtn;
    private Button fixBtn;
    private Button syncDbBtn;
    private Button recupBtn;
    private Button runActionBtn;
    private Button controleBtn;

    // =========================================================================
    // Services
    // =========================================================================

    private final MergeService            mergeService           = new MergeService(DatabaseManager.getInstance());
    private final EspacePartageFixer      espacePartageFixer     = new EspacePartageFixer();
    private final EtatPublicGenerator     etatPublicGenerator    = new EtatPublicGenerator();
    private final TrfGeneratorService     trfGeneratorService    = new TrfGeneratorService(DatabaseManager.getInstance());
    private final ProcreancesComparator   procreancesComparator  = new ProcreancesComparator();
    private final ConsoControleComparator consoControleComparator = new ConsoControleComparator();
    private final RecupNumFactureService  recupNumFactureService = new RecupNumFactureService();
    private final ClientInfoService       clientInfoService      = new ClientInfoService();
    private final EtatCreancesSyncService syncService            = new EtatCreancesSyncService(DatabaseManager.getInstance());
    private final MonthClotureService     monthClotureService    = new MonthClotureService(DatabaseManager.getInstance(), trfGeneratorService);

    // =========================================================================
    // Background executor + helpers
    // =========================================================================

    private final ExecutorService executor = Executors.newSingleThreadExecutor(r -> {
        Thread t = new Thread(r, "merge-worker");
        t.setDaemon(true);
        return t;
    });

    private static final DateTimeFormatter TIME_FMT = DateTimeFormatter.ofPattern("HH:mm:ss");

    private File    lastOutputFile;
    private List<Button> navButtons;
    private List<VBox>   pages;

    // Config page state
    private String[] configPaths;
    private Label[]  configPathLabels;

    // Dashboard selection state
    private Label selectedCompanyItem = null;

    // =========================================================================
    // JavaFX lifecycle
    // =========================================================================

    @FXML
    public void initialize() {
        navButtons = List.of(navOperations, navDashboard, navHistorique, navConfig);
        pages      = List.of(pageOperations, pageDashboard, pageHistorique, pageConfig);

        progressBar.setProgress(0);
        statusBar.setVisible(false);

        navOperations.getStyleClass().add("nav-active");

        refreshFileBadges();
        buildOperationsButtons();
        buildConfigPage();

        if (DatabaseManager.getInstance().getAllTrfMonths().isEmpty()) {
            appendLog("[Dashboard] Base de données vide — chargement automatique...");
            refreshDashboardData();
        }
    }

    // =========================================================================
    // Navigation
    // =========================================================================

    @FXML private void showOperations() {
        switchPage(pageOperations, navOperations, "Opérations");
    }

    @FXML private void showDashboard() {
        switchPage(pageDashboard, navDashboard, "Dashboard");
        loadDashboard();
    }

    @FXML private void showHistorique() {
        switchPage(pageHistorique, navHistorique, "Historique");
        loadHistorique();
    }

    @FXML private void showConfig() {
        switchPage(pageConfig, navConfig, "Configuration");
        buildConfigPage();
    }

    @FXML private void showLogFull() {
        Stage s = new Stage();
        s.setTitle("Log complet");
        TextArea ta = new TextArea(logArea.getText());
        ta.setEditable(false);
        ta.setStyle("-fx-control-inner-background:#1A1A1A;-fx-text-fill:#00FF88;"
                  + "-fx-font-family:'Courier New',monospace;-fx-font-size:12px;");
        s.setScene(new Scene(new BorderPane(ta), 800, 500));
        s.show();
    }

    private void switchPage(VBox target, Button navBtn, String title) {
        for (int i = 0; i < pages.size(); i++) {
            VBox p = pages.get(i);
            p.setVisible(p == target);
            p.setManaged(p == target);
        }
        for (Button b : navButtons) b.getStyleClass().remove("nav-active");
        if (navBtn != null) navBtn.getStyleClass().add("nav-active");
        pageTitle.setText(title);
    }

    private void refreshDashboardIfActive() {
        if (pageDashboard != null && pageDashboard.isVisible()) loadDashboard();
    }

    // =========================================================================
    // Opérations page — button grid
    // =========================================================================

    private void buildOperationsButtons() {
        trfBtn       = createActionBtn("Générer TRF",          "Calcul virements et compensations",      "action-card-primary", e -> generateTrf());
        etatBtn      = createActionBtn("États publics",         "Exporter vers Espace Partagé",           "action-card",         e -> generateEtatPublic());
        cmpBtn       = createActionBtn("Comparer fichiers",     "Contrôle PROCRÉANCES",                   "action-card",         e -> compareProcreances());
        fixBtn       = createActionBtn("Corriger espaces",      "Mise à jour Espace Partagé",              "action-card",         e -> fixPaths());
        syncDbBtn    = createActionBtn("Sync toutes sociétés",  "Synchroniser toutes les sociétés",        "action-card",         e -> syncDatabase());
        recupBtn     = createActionBtn("Récup. n° factures",    "Depuis Dropbox",                          "action-card",         e -> recupNumFacture());
        controleBtn  = createActionBtn("Contrôle Facturation",  "Comparer Contrôle vs Consolidation",      "action-card",         e -> compareConsoControle());
        runActionBtn = createActionBtn("▶  CONSOLIDER",         "Lire les états → ConsolidationGénérale",  "consolider-card",     e -> run());

        actionsGrid.add(trfBtn,    0, 0);
        actionsGrid.add(etatBtn,   1, 0);
        actionsGrid.add(cmpBtn,    0, 1);
        actionsGrid.add(fixBtn,    1, 1);
        actionsGrid.add(syncDbBtn, 0, 2);
        actionsGrid.add(recupBtn,  1, 2);

        GridPane.setColumnSpan(runActionBtn, 2);
        actionsGrid.add(runActionBtn, 0, 3);

        controleBtn.setVisible(false);
        controleBtn.setManaged(false);
    }

    // =========================================================================
    // Opérations page — file config dialog
    // =========================================================================

    @FXML
    private void openFileConfig() {
        String[] paths = {
            AppPreferences.getMergeRoot(),
            AppPreferences.getOutputFolder(),
            AppPreferences.getTrfConso(),
            AppPreferences.getTrfListing(),
            AppPreferences.getTrfTableau(),
            AppPreferences.getProcreancesPath(),
            AppPreferences.getControlePath(),
            AppPreferences.getRecupFacturePath()
        };
        String[]  labels = {"Dossier source", "Dossier de sortie", "ConsolidationGénérale",
                             "Listing Cabinet Phénix", "Tableau de Bord", "Export PROCREANCES",
                             "Contrôle Facturation", "Récup. Num Facture"};
        boolean[] isDir  = {true, true, false, false, false, false, false, false};
        String[]  exts   = {null, null, "xlsx", "xlsx", "xlsx", "xls", "xlsx", "xlsx"};

        Stage dialog = new Stage();
        dialog.initModality(Modality.APPLICATION_MODAL);
        dialog.initOwner(badgesBox.getScene().getWindow());
        dialog.setTitle("Configuration des fichiers");
        dialog.setResizable(false);

        VBox root = new VBox(10);
        root.setPadding(new Insets(20));
        root.setPrefWidth(700);

        Label[] pathLabels = new Label[paths.length];

        for (int i = 0; i < paths.length; i++) {
            final int idx = i;
            HBox row = new HBox(8);
            row.setAlignment(Pos.CENTER_LEFT);
            Label lbl = new Label(labels[i] + ":");
            lbl.setMinWidth(200);
            lbl.setStyle("-fx-font-weight:bold;-fx-font-family:'Courier New',monospace;");
            pathLabels[i] = new Label();
            pathLabels[i].setMaxWidth(Double.MAX_VALUE);
            HBox.setHgrow(pathLabels[i], Priority.ALWAYS);
            updatePathLabel(pathLabels[i], paths[i], isDir[i]);
            Button changeBtn = new Button(paths[i].isEmpty() ? "Choisir" : "Changer");
            changeBtn.getStyleClass().add("secondary-btn");
            changeBtn.setOnAction(ev -> {
                File chosen = isDir[idx]
                    ? dialogPickDirectory(dialog, labels[idx], paths[idx])
                    : dialogPickFile(dialog, labels[idx], paths[idx], exts[idx]);
                if (chosen != null) {
                    paths[idx] = chosen.getAbsolutePath();
                    updatePathLabel(pathLabels[idx], paths[idx], isDir[idx]);
                    changeBtn.setText("Changer");
                }
            });
            row.getChildren().addAll(lbl, pathLabels[i], changeBtn);
            root.getChildren().add(row);
        }

        HBox footer = new HBox(8);
        footer.setAlignment(Pos.CENTER_RIGHT);
        footer.setPadding(new Insets(10, 0, 0, 0));
        Button cancelBtn = new Button("Annuler");
        cancelBtn.getStyleClass().add("secondary-btn");
        cancelBtn.setOnAction(ev -> dialog.close());
        Button saveBtn = new Button("Enregistrer");
        saveBtn.getStyleClass().add("run-btn");
        saveBtn.setOnAction(ev -> {
            AppPreferences.setMergeRoot(paths[0]);
            AppPreferences.setOutputFolder(paths[1]);
            AppPreferences.setTrfConso(paths[2]);
            AppPreferences.setTrfListing(paths[3]);
            AppPreferences.setTrfTableau(paths[4]);
            AppPreferences.setProcreancesPath(paths[5]);
            AppPreferences.setControlePath(paths[6]);
            AppPreferences.setRecupFacturePath(paths[7]);
            dialog.close();
            refreshFileBadges();
        });
        footer.getChildren().addAll(cancelBtn, saveBtn);
        root.getChildren().add(footer);

        Scene scene = new Scene(root);
        if (!badgesBox.getScene().getStylesheets().isEmpty()) {
            scene.getStylesheets().addAll(badgesBox.getScene().getStylesheets());
        }
        dialog.setScene(scene);
        dialog.showAndWait();
    }

    private void refreshFileBadges() {
        badgesBox.getChildren().clear();
        int missing = 0;
        missing += addBadge("Dossier source",        AppPreferences.getMergeRoot(),        true);
        missing += addBadge("Dossier sortie",         AppPreferences.getOutputFolder(),     true);
        missing += addBadge("ConsolidationGénérale",  AppPreferences.getTrfConso(),         false);
        missing += addBadge("Listing",                AppPreferences.getTrfListing(),       false);
        missing += addBadge("Tableau de bord",        AppPreferences.getTrfTableau(),       false);
        missing += addBadge("PROCREANCES",            AppPreferences.getProcreancesPath(),  false);
        missing += addBadge("Contrôle Fact.",         AppPreferences.getControlePath(),     false);
        missing += addBadge("Récup Factures",         AppPreferences.getRecupFacturePath(), false);
        if (missing > 0) {
            missingFilesLabel.setText(missing + " fichier(s) manquant(s)");
            missingFilesLabel.setVisible(true);
            missingFilesLabel.setManaged(true);
        } else {
            missingFilesLabel.setVisible(false);
            missingFilesLabel.setManaged(false);
        }
    }

    private int addBadge(String label, String path, boolean isDirectory) {
        boolean ok = !path.isEmpty()
            && (isDirectory ? new File(path).isDirectory() : new File(path).exists());
        Label badge = new Label(label + (ok ? " ✓" : " ✗"));
        badge.getStyleClass().add(ok ? "badge-ok" : "badge-missing");
        badgesBox.getChildren().add(badge);
        return ok ? 0 : 1;
    }

    // =========================================================================
    // Dashboard page
    // =========================================================================

    private void loadDashboard() {
        pageDashboard.getChildren().clear();
        pageDashboard.setPadding(new Insets(16, 24, 12, 24));
        pageDashboard.setSpacing(12);

        List<TrfMonthRecord> months = monthClotureService.getAllMonths();
        TrfMonthRecord latest = months.isEmpty() ? null : months.get(0);

        long   activeSocietes = latest != null ? latest.nbClients()    : 0;
        double totalMontant   = latest != null ? latest.totalMontant() : 0;
        double totalNousDoit  = latest != null ? latest.totalNousDoit(): 0;
        String dernierTrf     = latest != null
            ? String.format("%02d/%d", latest.month(), latest.year()) : "—";

        Button refreshBtn = new Button("↻  Actualiser les données");
        refreshBtn.getStyleClass().add("save-btn");
        refreshBtn.setOnAction(e -> refreshDashboardData());

        Region kpiSpacer = new Region();
        HBox.setHgrow(kpiSpacer, Priority.ALWAYS);
        HBox kpiHeader = new HBox(kpiSpacer, refreshBtn);
        kpiHeader.setAlignment(Pos.CENTER_RIGHT);
        pageDashboard.getChildren().add(kpiHeader);

        HBox kpiRow = new HBox(12,
            kpiCard("Sociétés actives",   String.valueOf(activeSocietes)),
            kpiCard("Montant à facturer", String.format("%.2f €", totalMontant)),
            kpiCard("Nous doit",          String.format("%.2f €", totalNousDoit)),
            kpiCard("Dernier TRF",        dernierTrf)
        );
        pageDashboard.getChildren().add(kpiRow);

        List<TrfHistoryRecord> clients = latest != null
            ? monthClotureService.getHistoryForMonth(latest.id()) : List.of();

        VBox leftPanel = new VBox(4);
        leftPanel.setPrefWidth(190);
        leftPanel.setMinWidth(190);
        leftPanel.setMaxWidth(190);
        leftPanel.setPadding(new Insets(0, 8, 0, 0));

        Label societeLabel = new Label("Sociétés");
        societeLabel.getStyleClass().add("files-card-title");
        leftPanel.getChildren().add(societeLabel);

        VBox companyListBox = new VBox(2);
        ScrollPane companyScroll = new ScrollPane(companyListBox);
        companyScroll.setFitToWidth(true);
        companyScroll.setStyle("-fx-background-color:transparent;-fx-border-color:transparent;");
        VBox.setVgrow(companyScroll, Priority.ALWAYS);
        leftPanel.getChildren().add(companyScroll);

        VBox rightPanel = new VBox(10);
        HBox.setHgrow(rightPanel, Priority.ALWAYS);
        rightPanel.setPadding(new Insets(0, 0, 0, 16));

        HBox bodyRow = new HBox(0, leftPanel, rightPanel);
        VBox.setVgrow(bodyRow, Priority.ALWAYS);
        pageDashboard.getChildren().add(bodyRow);

        selectedCompanyItem = null;
        for (TrfHistoryRecord r : clients) {
            Label item = new Label(etatDot(r.etat()) + "  " + r.clientName());
            item.setMaxWidth(Double.MAX_VALUE);
            item.getStyleClass().add("company-list-item");
            item.setOnMouseClicked(e -> {
                if (selectedCompanyItem != null)
                    selectedCompanyItem.getStyleClass().remove("company-list-selected");
                selectedCompanyItem = item;
                item.getStyleClass().add("company-list-selected");
                showClientDetail(r, rightPanel);
            });
            companyListBox.getChildren().add(item);
        }

        if (!clients.isEmpty()) {
            Label first = (Label) companyListBox.getChildren().get(0);
            selectedCompanyItem = first;
            first.getStyleClass().add("company-list-selected");
            showClientDetail(clients.get(0), rightPanel);
        }
    }

    private void refreshDashboardData() {
        String consoPath   = AppPreferences.getTrfConso();
        String listingPath = AppPreferences.getTrfListing();
        String tableauPath = AppPreferences.getTrfTableau();

        if (consoPath.isEmpty()) {
            appendLog("[Dashboard] ConsolidationGénérale non configurée — ouvrez Configuration.");
            return;
        }
        if (listingPath.isEmpty()) {
            appendLog("[Dashboard] Listing Cabinet Phénix non configuré — ouvrez Configuration.");
            return;
        }
        if (tableauPath.isEmpty()) {
            appendLog("[Dashboard] Tableau de bord non configuré — ouvrez Configuration.");
            return;
        }
        File consoFile   = new File(consoPath);
        File listingFile = new File(listingPath);
        File tableauFile = new File(tableauPath);

        if (!consoFile.exists()) {
            appendLog("[Dashboard] Fichier introuvable : " + consoPath);
            return;
        }
        if (!listingFile.exists()) {
            appendLog("[Dashboard] Fichier introuvable : " + listingPath);
            return;
        }
        if (!tableauFile.exists()) {
            appendLog("[Dashboard] Fichier introuvable : " + tableauPath);
            return;
        }

        appendLog("[Dashboard] Chargement des données...");
        LocalDate now = LocalDate.now();

        executor.submit(() -> {
            try {
                monthClotureService.saveOpenMonth(now.getYear(), now.getMonthValue(),
                    consoFile, listingFile, tableauFile);
                List<TrfHistoryRecord> loaded = monthClotureService
                    .getHistoryForMonth(DatabaseManager.getInstance()
                        .getTrfMonthId(now.getYear(), now.getMonthValue()));
                Platform.runLater(() -> {
                    appendLog("[Dashboard] ✓ " + loaded.size() + " sociétés chargées");
                    loadDashboard();
                });
            } catch (Exception e) {
                Platform.runLater(() -> appendLog("[Dashboard] ERREUR : " + e.getMessage()));
            }
        });
    }

    private VBox kpiCard(String label, String value) {
        Label lbl = new Label(label);
        lbl.getStyleClass().add("kpi-label");
        Label val = new Label(value);
        val.getStyleClass().add("kpi-value");
        VBox card = new VBox(4, lbl, val);
        card.getStyleClass().add("kpi-card");
        HBox.setHgrow(card, Priority.ALWAYS);
        return card;
    }

    private void showClientDetail(TrfHistoryRecord r, VBox detailPanel) {
        detailPanel.getChildren().clear();

        String etatText = r.etat() != null ? r.etat() : "—";
        Label nameLabel = new Label(r.clientName()
            + (r.clientCode() != null && !r.clientCode().isBlank() ? "  [" + r.clientCode() + "]" : ""));
        nameLabel.getStyleClass().add("detail-company-name");
        Label etatBadge = new Label(etatText);
        etatBadge.getStyleClass().addAll("etat-badge", etatCssClass(etatText));
        HBox header = new HBox(12, nameLabel, etatBadge);
        header.setAlignment(Pos.CENTER_LEFT);
        detailPanel.getChildren().add(header);

        HBox metrics = new HBox(10,
            metricCard("Encaissements CZ Phénix", String.format("%.2f €", r.encaissements())),
            metricCard("Montant à facturer TTC",   String.format("%.2f €", r.montantFacturer())),
            metricCard("Nous doit précédemment",   String.format("%.2f €", r.nousDoit())),
            metricCard("Sommes à reverser",         String.format("%.2f €", r.sommesReverser()))
        );
        detailPanel.getChildren().add(metrics);

        try {
            List<double[]> history = DatabaseManager.getInstance().getClientMonthlyHistory(r.clientName(), 6);
            if (!history.isEmpty()) {
                CategoryAxis xAxis = new CategoryAxis();
                NumberAxis   yAxis = new NumberAxis();
                yAxis.setLabel("€");
                BarChart<String, Number> chart = new BarChart<>(xAxis, yAxis);
                chart.setTitle("Montant à facturer — 6 derniers mois");
                chart.setLegendVisible(false);
                chart.setPrefHeight(180);
                chart.setAnimated(false);
                XYChart.Series<String, Number> series = new XYChart.Series<>();
                for (double[] pt : history) {
                    series.getData().add(new XYChart.Data<>(
                        String.format("%02d/%d", (int) pt[0], (int) pt[1]), pt[2]));
                }
                chart.getData().add(series);
                detailPanel.getChildren().add(chart);
            }
        } catch (Exception ignored) {}
    }

    private VBox metricCard(String label, String value) {
        Label lbl = new Label(label);
        lbl.getStyleClass().add("dp-metric-label");
        Label val = new Label(value);
        val.getStyleClass().add("dp-metric-value");
        VBox card = new VBox(4, lbl, val);
        card.getStyleClass().add("dp-metric-card");
        HBox.setHgrow(card, Priority.ALWAYS);
        return card;
    }

    private String etatCssClass(String etat) {
        if (etat == null) return "etat-debit";
        String lower = etat.toLowerCase();
        if (lower.contains("comp") && !lower.contains("non") && !lower.contains("partiel"))
            return "etat-comp";
        if (lower.contains("partiel"))
            return "etat-noncomp";
        return "etat-debit";
    }

    private String etatDot(String etat) {
        if (etat == null) return "●";
        String lower = etat.toLowerCase();
        if (lower.contains("comp") && !lower.contains("non") && !lower.contains("partiel")) return "🟢";
        if (lower.contains("partiel")) return "🟡";
        return "🔴";
    }

    // =========================================================================
    // Historique page
    // =========================================================================

    private void loadHistorique() {
        pageHistorique.getChildren().clear();

        VBox list = new VBox(10);
        list.setPadding(new Insets(16, 24, 12, 24));
        ScrollPane scroll = new ScrollPane(list);
        scroll.setFitToWidth(true);
        scroll.setStyle("-fx-background-color:transparent;-fx-border-color:transparent;");
        VBox.setVgrow(scroll, Priority.ALWAYS);
        pageHistorique.getChildren().add(scroll);

        List<TrfMonthRecord> months = monthClotureService.getAllMonths();

        LocalDate now = LocalDate.now();
        boolean currentClosed = months.stream()
            .anyMatch(m -> m.year() == now.getYear() && m.month() == now.getMonthValue()
                       && "closed".equals(m.status()));

        if (!currentClosed) {
            list.getChildren().add(buildMonthCard(now.getYear(), now.getMonthValue(),
                false, 0, 0.0, List.of()));
        }

        for (TrfMonthRecord m : months) {
            List<TrfHistoryRecord> clients = monthClotureService.getHistoryForMonth(m.id());
            list.getChildren().add(buildMonthCard(m.year(), m.month(),
                "closed".equals(m.status()), m.nbClients(), m.totalMontant(), clients));
        }
    }

    private VBox buildMonthCard(int year, int month, boolean closed,
                                int nbClients, double totalMontant,
                                List<TrfHistoryRecord> clients) {
        VBox card = new VBox(0);
        card.getStyleClass().add("month-card");

        HBox header = new HBox(10);
        header.setAlignment(Pos.CENTER_LEFT);
        header.getStyleClass().add("month-card-header");
        header.setPadding(new Insets(10, 12, 10, 12));

        String monthName = java.time.Month.of(month)
            .getDisplayName(TextStyle.FULL_STANDALONE, Locale.FRENCH);
        String displayName = capitalize(monthName) + " " + year;

        Label title = new Label("📅 " + displayName);
        title.getStyleClass().add("month-name");
        HBox.setHgrow(title, Priority.ALWAYS);
        header.getChildren().add(title);

        if (closed) {
            Label stats = new Label(nbClients + " sociétés   " + String.format("%.0f €", totalMontant));
            stats.getStyleClass().add("month-stat");
            Label badge = new Label("Clôturé");
            badge.getStyleClass().add("badge-closed");
            header.getChildren().addAll(stats, badge);
        } else {
            Label badge = new Label("En cours");
            badge.getStyleClass().add("badge-open");
            Button cloturerBtn = new Button("Clôturer le mois");
            cloturerBtn.getStyleClass().add("cloture-btn");
            final int y = year, mo = month;
            cloturerBtn.setOnAction(e -> cloturerMois(y, mo));
            header.getChildren().addAll(badge, cloturerBtn);
        }

        card.getChildren().add(header);

        VBox body = new VBox(4);
        body.setPadding(new Insets(8, 12, 8, 12));
        body.setVisible(false);
        body.setManaged(false);
        body.getStyleClass().add("month-card-body");

        if (clients.isEmpty()) {
            body.getChildren().add(new Label("Aucune donnée"));
        } else {
            for (TrfHistoryRecord r : clients) {
                Label row = new Label(etatDot(r.etat()) + "  " + r.clientName()
                    + (r.montantFacturer() != 0 ? "   " + String.format("%.2f €", r.montantFacturer()) : ""));
                row.getStyleClass().add("month-card-client-row");
                body.getChildren().add(row);
            }
        }
        card.getChildren().add(body);

        header.setOnMouseClicked(e -> {
            body.setVisible(!body.isVisible());
            body.setManaged(body.isVisible());
        });
        header.setStyle("-fx-cursor:hand;");

        return card;
    }

    private String capitalize(String s) {
        if (s == null || s.isEmpty()) return s;
        return Character.toUpperCase(s.charAt(0)) + s.substring(1);
    }

    private void cloturerMois(int year, int month) {
        String consoPath   = AppPreferences.getTrfConso();
        String listingPath = AppPreferences.getTrfListing();
        String tableauPath = AppPreferences.getTrfTableau();

        if (consoPath.isEmpty() || listingPath.isEmpty() || tableauPath.isEmpty()) {
            appendLog("ERROR: Configurez les fichiers TRF avant de clôturer.");
            showOperations();
            return;
        }

        File consoFile   = new File(consoPath);
        File listingFile = new File(listingPath);
        File tableauFile = new File(tableauPath);

        if (!consoFile.exists() || !listingFile.exists() || !tableauFile.exists()) {
            appendLog("ERROR: Fichiers TRF introuvables.");
            showOperations();
            return;
        }

        appendLog("Clôture du mois " + month + "/" + year + "…");

        executor.submit(() -> {
            try {
                monthClotureService.cloturerMois(year, month, consoFile, listingFile, tableauFile);
                Platform.runLater(() -> {
                    appendLog("Mois " + month + "/" + year + " clôturé.");
                    loadHistorique();
                });
            } catch (Exception e) {
                Platform.runLater(() -> appendLog("ERREUR clôture : " + e.getMessage()));
            }
        });
    }

    // =========================================================================
    // Configuration page
    // =========================================================================

    private void buildConfigPage() {
        configPaths = new String[]{
            AppPreferences.getMergeRoot(),
            AppPreferences.getTrfConso(),
            AppPreferences.getTrfListing(),
            AppPreferences.getTrfTableau()
        };
        String[]  labels = {"Dossier source (Dropbox)", "ConsolidationGénérale.xlsx",
                             "Listing Cabinet Phénix.xls", "Tableau de bord facturation.xlsx"};
        boolean[] isDir  = {true, false, false, false};
        configPathLabels = new Label[configPaths.length];

        configFormBox.getChildren().clear();
        for (int i = 0; i < configPaths.length; i++) {
            final int idx = i;
            HBox row = new HBox(8);
            row.setAlignment(Pos.CENTER_LEFT);
            Label lbl = new Label(labels[i] + ":");
            lbl.setMinWidth(260);
            lbl.getStyleClass().add("config-label");
            configPathLabels[i] = new Label();
            configPathLabels[i].setMaxWidth(Double.MAX_VALUE);
            HBox.setHgrow(configPathLabels[i], Priority.ALWAYS);
            updatePathLabel(configPathLabels[i], configPaths[i], isDir[i]);
            Button browseBtn = new Button("Parcourir");
            browseBtn.getStyleClass().add("browse-btn");
            browseBtn.setOnAction(ev -> {
                File chosen = isDir[idx]
                    ? dialogPickDirectory(null, labels[idx], configPaths[idx])
                    : dialogPickFile(null, labels[idx], configPaths[idx], "xlsx");
                if (chosen != null) {
                    configPaths[idx] = chosen.getAbsolutePath();
                    updatePathLabel(configPathLabels[idx], configPaths[idx], isDir[idx]);
                }
            });
            row.getChildren().addAll(lbl, configPathLabels[i], browseBtn);
            configFormBox.getChildren().add(row);
        }
    }

    @FXML
    private void saveConfig() {
        AppPreferences.setMergeRoot(configPaths[0]);
        AppPreferences.setTrfConso(configPaths[1]);
        AppPreferences.setTrfListing(configPaths[2]);
        AppPreferences.setTrfTableau(configPaths[3]);
        refreshFileBadges();
        appendLog("Configuration enregistrée.");
        showOperations();
    }

    // =========================================================================
    // Opérations — action handlers
    // =========================================================================

    private void generateTrf() {
        String consoPath   = AppPreferences.getTrfConso();
        String listingPath = AppPreferences.getTrfListing();
        String tableauPath = AppPreferences.getTrfTableau();
        String outputPath  = AppPreferences.getOutputFolder();

        if (consoPath.isEmpty() || listingPath.isEmpty() || tableauPath.isEmpty()) {
            appendLog("ERROR: Configurez les trois fichiers TRF avant de générer."); return;
        }
        File consoFile    = new File(consoPath);
        File listingFile  = new File(listingPath);
        File tableauFile  = new File(tableauPath);
        File outputFolder = new File(outputPath);

        if (!consoFile.exists())         { appendLog("ERROR: Fichier introuvable — " + consoPath);   return; }
        if (!listingFile.exists())       { appendLog("ERROR: Fichier introuvable — " + listingPath); return; }
        if (!tableauFile.exists())       { appendLog("ERROR: Fichier introuvable — " + tableauPath); return; }
        if (!outputFolder.isDirectory()) { appendLog("ERROR: Dossier sortie introuvable — " + outputPath); return; }

        setAllButtonsDisabled(true);
        statusBar.setVisible(false);
        progressBar.setProgress(0);
        logArea.clear();
        lastOutputFile = null;

        executor.submit(() -> {
            try {
                File result = trfGeneratorService.generate(consoFile, listingFile, tableauFile, outputFolder,
                    (prog, msg) -> Platform.runLater(() -> { progressBar.setProgress(prog); appendLog(msg); }));
                Platform.runLater(() -> {
                    if (result != null) {
                        lastOutputFile = result;
                        statusLabel.setText("TRF Output: " + result.getAbsolutePath());
                        openFileBtn.setVisible(true);
                        statusBar.setVisible(true);
                        refreshDashboardIfActive();
                    }
                    setAllButtonsDisabled(false);
                });
            } catch (Exception e) {
                Platform.runLater(() -> { appendLog("FATAL: " + e.getMessage()); setAllButtonsDisabled(false); });
            }
        });
    }

    private void generateEtatPublic() {
        File rootFolder = new File(AppPreferences.getMergeRoot());
        if (!rootFolder.isDirectory()) {
            appendLog("ERROR: Dossier source introuvable — " + rootFolder.getAbsolutePath()); return;
        }
        setAllButtonsDisabled(true);
        statusBar.setVisible(false);
        progressBar.setProgress(0);
        logArea.clear();
        lastOutputFile = null;

        executor.submit(() -> {
            try {
                etatPublicGenerator.generate(rootFolder,
                    (prog, msg) -> Platform.runLater(() -> { progressBar.setProgress(prog); appendLog(msg); }));
                Platform.runLater(() -> {
                    statusLabel.setText("Etat Public files written to EspacePartagé paths.");
                    openFileBtn.setVisible(false);
                    statusBar.setVisible(true);
                    setAllButtonsDisabled(false);
                });
            } catch (Exception e) {
                Platform.runLater(() -> { appendLog("FATAL: " + e.getMessage()); setAllButtonsDisabled(false); });
            }
        });
    }

    private void fixPaths() {
        File rootFolder = new File(AppPreferences.getMergeRoot());
        if (!rootFolder.isDirectory()) {
            appendLog("ERROR: Dossier source introuvable — " + rootFolder.getAbsolutePath()); return;
        }
        setAllButtonsDisabled(true);
        statusBar.setVisible(false);
        progressBar.setProgress(0);
        logArea.clear();
        lastOutputFile = null;

        executor.submit(() -> {
            try {
                File result = espacePartageFixer.fix(rootFolder,
                    (progress, msg) -> Platform.runLater(() -> { progressBar.setProgress(progress); appendLog(msg); }));
                Platform.runLater(() -> {
                    lastOutputFile = result;
                    statusLabel.setText("Saved: " + result.getAbsolutePath());
                    openFileBtn.setVisible(true);
                    statusBar.setVisible(true);
                    setAllButtonsDisabled(false);
                });
            } catch (Exception e) {
                Platform.runLater(() -> { appendLog("FATAL: " + e.getMessage()); setAllButtonsDisabled(false); });
            }
        });
    }

    private void run() {
        File rootFolder   = new File(AppPreferences.getMergeRoot());
        File outputFolder = new File(AppPreferences.getOutputFolder());

        if (!rootFolder.isDirectory()) {
            appendLog("ERROR: Dossier source introuvable — " + rootFolder.getAbsolutePath()); return;
        }
        if (!outputFolder.isDirectory()) {
            appendLog("ERROR: Dossier sortie introuvable — " + outputFolder.getAbsolutePath()); return;
        }
        setAllButtonsDisabled(true);
        statusBar.setVisible(false);
        progressBar.setProgress(0);
        logArea.clear();
        lastOutputFile = null;

        executor.submit(() -> {
            try {
                File result = mergeService.merge(rootFolder, outputFolder,
                    (progress, msg) -> Platform.runLater(() -> { progressBar.setProgress(progress); appendLog(msg); }));
                Platform.runLater(() -> {
                    if (result != null) {
                        lastOutputFile = result;
                        statusLabel.setText("Output: " + result.getAbsolutePath());
                        openFileBtn.setVisible(true);
                        statusBar.setVisible(true);
                        refreshDashboardIfActive();
                    }
                    setAllButtonsDisabled(false);
                });
            } catch (Exception e) {
                Platform.runLater(() -> { appendLog("FATAL: " + e.getMessage()); setAllButtonsDisabled(false); });
            }
        });
    }

    private void compareProcreances() {
        String procPath   = AppPreferences.getProcreancesPath();
        String consoPath  = AppPreferences.getTrfConso();
        String outputPath = AppPreferences.getOutputFolder();

        if (procPath.isEmpty() || consoPath.isEmpty()) {
            appendLog("ERROR: Configurez Export PROCREANCES et ConsolidationGénérale avant de comparer."); return;
        }
        File procFile     = new File(procPath);
        File consoFile    = new File(consoPath);
        File outputFolder = new File(outputPath);

        if (!procFile.exists())          { appendLog("ERROR: Fichier introuvable — " + procPath);  return; }
        if (!consoFile.exists())         { appendLog("ERROR: Fichier introuvable — " + consoPath); return; }
        if (!outputFolder.isDirectory()) { appendLog("ERROR: Dossier sortie introuvable — " + outputPath); return; }

        setAllButtonsDisabled(true);
        statusBar.setVisible(false);
        progressBar.setProgress(0);
        logArea.clear();
        lastOutputFile = null;

        executor.submit(() -> {
            try {
                File report = procreancesComparator.compare(procFile, consoFile, outputFolder,
                    (prog, msg) -> Platform.runLater(() -> { progressBar.setProgress(prog); appendLog(msg); }));
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
                Platform.runLater(() -> { appendLog("FATAL: " + ex.getMessage()); setAllButtonsDisabled(false); });
            }
        });
    }

    private void compareConsoControle() {
        String controlePath = AppPreferences.getControlePath();
        String consoPath    = AppPreferences.getTrfConso();
        String outputPath   = AppPreferences.getOutputFolder();

        if (controlePath.isEmpty() || consoPath.isEmpty()) {
            appendLog("ERROR: Configurez Contrôle Facturation et ConsolidationGénérale avant de comparer."); return;
        }
        File controleFile  = new File(controlePath);
        File consoFile     = new File(consoPath);
        File outputFolder  = new File(outputPath);

        if (!controleFile.exists())      { appendLog("ERROR: Fichier introuvable — " + controlePath);  return; }
        if (!consoFile.exists())         { appendLog("ERROR: Fichier introuvable — " + consoPath);      return; }
        if (!outputFolder.isDirectory()) { appendLog("ERROR: Dossier sortie introuvable — " + outputPath); return; }

        setAllButtonsDisabled(true);
        statusBar.setVisible(false);
        progressBar.setProgress(0);
        logArea.clear();
        lastOutputFile = null;

        executor.submit(() -> {
            try {
                File report = consoControleComparator.compare(controleFile, consoFile, outputFolder,
                    (prog, msg) -> Platform.runLater(() -> { progressBar.setProgress(prog); appendLog(msg); }));
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
                Platform.runLater(() -> { appendLog("FATAL: " + ex.getMessage()); setAllButtonsDisabled(false); });
            }
        });
    }

    private void recupNumFacture() {
        String recupPath = AppPreferences.getRecupFacturePath();
        String rootPath  = AppPreferences.getMergeRoot();

        if (recupPath.isEmpty()) {
            appendLog("ERROR: Configurez le fichier Récup. Num Facture avant de lancer."); return;
        }
        File recupFile  = new File(recupPath);
        File rootFolder = new File(rootPath);

        if (!recupFile.exists())       { appendLog("ERROR: Fichier introuvable — " + recupPath); return; }
        if (!rootFolder.isDirectory()) { appendLog("ERROR: Dossier source introuvable — " + rootPath); return; }

        setAllButtonsDisabled(true);
        statusBar.setVisible(false);
        progressBar.setProgress(0);
        logArea.clear();
        lastOutputFile = null;

        executor.submit(() -> {
            try {
                java.util.List<String> log = recupNumFactureService.apply(recupFile, rootFolder,
                    (prog, msg) -> Platform.runLater(() -> { progressBar.setProgress(prog); appendLog(msg); }));
                Platform.runLater(() -> {
                    log.forEach(this::appendLog);
                    statusLabel.setText("Récup. Factures terminée — " + log.size() + " dossier(s).");
                    openFileBtn.setVisible(false);
                    statusBar.setVisible(true);
                    setAllButtonsDisabled(false);
                });
            } catch (Exception ex) {
                Platform.runLater(() -> { appendLog("FATAL: " + ex.getMessage()); setAllButtonsDisabled(false); });
            }
        });
    }

    private void syncDatabase() {
        File root = new File(AppPreferences.getMergeRoot());
        if (!root.isDirectory()) {
            appendLog("ERROR: Dossier source introuvable. Configurez le chemin."); return;
        }
        syncDbBtn.setDisable(true);
        appendLog("Synchronisation DB en cours…");
        executor.submit(() -> {
            try {
                syncService.syncAll(root, (pct, msg) ->
                    Platform.runLater(() -> { progressBar.setProgress(pct); appendLog(msg); }));
            } catch (Exception e) {
                Platform.runLater(() -> appendLog("ERREUR sync : " + e.getMessage()));
            } finally {
                Platform.runLater(() -> {
                    syncDbBtn.setDisable(false);
                    refreshDashboardIfActive();
                });
            }
        });
    }

    public void shutdown() {
        executor.shutdownNow();
    }

    @FXML
    private void openFile() {
        if (lastOutputFile != null && lastOutputFile.exists()) {
            try { Desktop.getDesktop().open(lastOutputFile); }
            catch (Exception e) { appendLog("Cannot open file: " + e.getMessage()); }
        }
    }

    // =========================================================================
    // Utilities
    // =========================================================================

    private Button createActionBtn(String name, String desc, String styleClass,
                                    EventHandler<ActionEvent> handler) {
        String titleClass, subtitleClass;
        switch (styleClass) {
            case "action-card-primary" -> { titleClass = "action-card-title-primary"; subtitleClass = "action-card-subtitle-primary"; }
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

    private void updatePathLabel(Label lbl, String path, boolean isDir) {
        if (path == null || path.isEmpty()) {
            lbl.setText("(non configuré)");
            lbl.setStyle("-fx-text-fill:#791F1F;-fx-font-size:11px;");
        } else {
            boolean exists = isDir ? new File(path).isDirectory() : new File(path).exists();
            String display = path.length() > 60 ? "…" + path.substring(path.length() - 57) : path;
            lbl.setText(display);
            lbl.setStyle((exists ? "-fx-text-fill:#3B6D11;" : "-fx-text-fill:#791F1F;") + "-fx-font-size:11px;");
        }
    }

    private File dialogPickDirectory(Stage owner, String title, String lastPath) {
        DirectoryChooser dc = new DirectoryChooser();
        dc.setTitle(title);
        if (lastPath != null && !lastPath.isEmpty()) {
            File f = new File(lastPath);
            if (f.isDirectory()) dc.setInitialDirectory(f);
        }
        return dc.showDialog(owner);
    }

    private File dialogPickFile(Stage owner, String title, String lastPath, String ext) {
        FileChooser fc = new FileChooser();
        fc.setTitle(title);
        if (ext != null) {
            fc.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*." + ext));
        }
        if (lastPath != null && !lastPath.isEmpty()) {
            File parent = new File(lastPath).getParentFile();
            if (parent != null && parent.isDirectory()) fc.setInitialDirectory(parent);
        }
        return fc.showOpenDialog(owner);
    }

    private void setAllButtonsDisabled(boolean disabled) {
        if (trfBtn       != null) trfBtn.setDisable(disabled);
        if (etatBtn      != null) etatBtn.setDisable(disabled);
        if (cmpBtn       != null) cmpBtn.setDisable(disabled);
        if (fixBtn       != null) fixBtn.setDisable(disabled);
        if (controleBtn  != null) controleBtn.setDisable(disabled);
        if (recupBtn     != null) recupBtn.setDisable(disabled);
        if (syncDbBtn    != null) syncDbBtn.setDisable(disabled);
        if (runActionBtn != null) runActionBtn.setDisable(disabled);
    }

    private void appendLog(String message) {
        String ts = LocalTime.now().format(TIME_FMT);
        logArea.appendText("[" + ts + "] " + message + "\n");
    }
}
