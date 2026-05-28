package com.zeki.merger.controller;

import com.zeki.merger.db.DatabaseManager;
import com.zeki.merger.service.ConsoControleComparator;
import com.zeki.merger.service.EspacePartageFixer;
import com.zeki.merger.service.EtatCreancesSyncService;
import com.zeki.merger.service.EtatPublicGenerator;
import com.zeki.merger.service.MergeService;
import com.zeki.merger.service.MonthClotureService;
import com.zeki.merger.service.ProcreancesComparator;
import com.zeki.merger.service.RecupNumFactureService;
import com.zeki.merger.trf.TrfGeneratorService;
import javafx.fxml.FXML;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.stage.Stage;

import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.List;
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
    // Services (instantiated here, passed to sub-controllers)
    // =========================================================================

    private final TrfGeneratorService     trfGeneratorService    = new TrfGeneratorService(DatabaseManager.getInstance());
    private final MonthClotureService     monthClotureService    = new MonthClotureService(DatabaseManager.getInstance(), trfGeneratorService);
    private final MergeService            mergeService           = new MergeService(DatabaseManager.getInstance());
    private final EspacePartageFixer      espacePartageFixer     = new EspacePartageFixer();
    private final EtatPublicGenerator     etatPublicGenerator    = new EtatPublicGenerator();
    private final ProcreancesComparator   procreancesComparator  = new ProcreancesComparator();
    private final ConsoControleComparator consoControleComparator = new ConsoControleComparator();
    private final RecupNumFactureService  recupNumFactureService = new RecupNumFactureService();
    private final EtatCreancesSyncService syncService            = new EtatCreancesSyncService(DatabaseManager.getInstance());

    // =========================================================================
    // Shared executor + helpers
    // =========================================================================

    private final ExecutorService executor = Executors.newSingleThreadExecutor(r -> {
        Thread t = new Thread(r, "merge-worker");
        t.setDaemon(true);
        return t;
    });

    private static final DateTimeFormatter TIME_FMT = DateTimeFormatter.ofPattern("HH:mm:ss");

    // =========================================================================
    // Sub-controllers
    // =========================================================================

    private DashboardController   dashboardController;
    private HistoriqueController  historiqueController;
    private ConfigController      configController;
    private OperationsController  operationsController;

    private List<Button> navButtons;
    private List<VBox>   pages;

    // =========================================================================
    // JavaFX lifecycle
    // =========================================================================

    @FXML
    public void initialize() {
        navButtons = List.of(navOperations, navDashboard, navHistorique, navConfig);
        pages      = List.of(pageOperations, pageDashboard, pageHistorique, pageConfig);

        progressBar.setProgress(0);
        statusBar.setVisible(false);
        statusBar.setManaged(false);
        navOperations.getStyleClass().add("nav-active");

        dashboardController  = new DashboardController(DatabaseManager.getInstance(),
            monthClotureService, trfGeneratorService, this::appendLog, executor);

        historiqueController = new HistoriqueController(DatabaseManager.getInstance(),
            monthClotureService, trfGeneratorService, this::appendLog, executor,
            this::showOperations);

        configController = new ConfigController(configFormBox, badgesBox, missingFilesLabel,
            this::appendLog, this::showOperations);

        operationsController = new OperationsController(
            mergeService, espacePartageFixer, etatPublicGenerator, trfGeneratorService,
            procreancesComparator, consoControleComparator, recupNumFactureService, syncService,
            progressBar, statusBar, statusLabel, openFileBtn, logArea,
            this::appendLog,
            () -> { if (pageDashboard != null && pageDashboard.isVisible())
                        dashboardController.load(pageDashboard); },
            executor);

        configController.refreshBadges();
        configController.load();
        operationsController.buildButtons(actionsGrid);

        if (DatabaseManager.getInstance().getAllTrfMonths().isEmpty()) {
            appendLog("[Dashboard] Base de données vide — chargement automatique...");
            dashboardController.refresh(pageDashboard);
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
        dashboardController.load(pageDashboard);
    }

    @FXML private void showHistorique() {
        switchPage(pageHistorique, navHistorique, "Historique");
        historiqueController.load(pageHistorique);
    }

    @FXML private void showConfig() {
        switchPage(pageConfig, navConfig, "Configuration");
        configController.load();
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

    // =========================================================================
    // FXML delegates
    // =========================================================================

    @FXML private void openFileConfig() { configController.openFileConfig(); }
    @FXML private void saveConfig()     { configController.save(); }
    @FXML private void openFile()       { operationsController.openFile(); }

    // =========================================================================
    // Utilities
    // =========================================================================

    private void switchPage(VBox target, Button navBtn, String title) {
        for (VBox p : pages) {
            p.setVisible(p == target);
            p.setManaged(p == target);
        }
        for (Button b : navButtons) b.getStyleClass().remove("nav-active");
        if (navBtn != null) navBtn.getStyleClass().add("nav-active");
        pageTitle.setText(title);
    }

    private void appendLog(String message) {
        String ts = LocalTime.now().format(TIME_FMT);
        logArea.appendText("[" + ts + "] " + message + "\n");
    }

    public void shutdown() {
        executor.shutdownNow();
    }
}
