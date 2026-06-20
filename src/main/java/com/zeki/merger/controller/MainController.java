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
    @FXML private Button navConfig;
    @FXML private Button navLogFull;
    @FXML private Label  pageTitle;

    // =========================================================================
    // FXML — pages
    // =========================================================================

    @FXML private VBox pageOperations;
    @FXML private VBox pageDashboard;
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

    @FXML private VBox   sidebarPanel;
    @FXML private Button sidebarToggleBtn;
    @FXML private Label  sidebarLogoTitle;
    @FXML private Label  sidebarLogoSub;
    @FXML private Label  sidebarLabelPrincipal;
    @FXML private Label  sidebarLabelParams;
    private boolean sidebarCollapsed = false;

    private DashboardController   dashboardController;
    private ConfigController      configController;
    private OperationsController  operationsController;

    private List<Button> navButtons;
    private List<VBox>   pages;

    // =========================================================================
    // JavaFX lifecycle
    // =========================================================================

    @FXML
    public void initialize() {
        navButtons = List.of(navOperations, navDashboard, navConfig);
        pages      = List.of(pageOperations, pageDashboard, pageConfig);

        progressBar.setProgress(0);
        statusBar.setVisible(false);
        statusBar.setManaged(false);
        navOperations.getStyleClass().add("nav-active");

        dashboardController  = new DashboardController(DatabaseManager.getInstance(),
            syncService, this::appendLog, executor);

        configController = new ConfigController(configFormBox, badgesBox, missingFilesLabel,
            this::appendLog, this::showOperations, DatabaseManager.getInstance());

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

        // Dashboard loads on demand — no auto-refresh needed
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

    @FXML private void showConfig() {
        switchPage(pageConfig, navConfig, "Configuration");
        configController.load();
    }

    @FXML private void toggleSidebar() {
        sidebarCollapsed = !sidebarCollapsed;
        if (sidebarCollapsed) {
            sidebarPanel.setPrefWidth(48);
            sidebarPanel.setMinWidth(48);
            sidebarPanel.setMaxWidth(48);
            sidebarLogoTitle.setVisible(false);  sidebarLogoTitle.setManaged(false);
            sidebarLogoSub.setVisible(false);    sidebarLogoSub.setManaged(false);
            sidebarLabelPrincipal.setVisible(false); sidebarLabelPrincipal.setManaged(false);
            sidebarLabelParams.setVisible(false);    sidebarLabelParams.setManaged(false);
            navOperations.setText("⊞");
            navDashboard.setText("▦");
            navConfig.setText("⚙");
            navLogFull.setText("⌨");
            sidebarToggleBtn.setText("▶");
        } else {
            sidebarPanel.setPrefWidth(200);
            sidebarPanel.setMinWidth(200);
            sidebarPanel.setMaxWidth(200);
            sidebarLogoTitle.setVisible(true);  sidebarLogoTitle.setManaged(true);
            sidebarLogoSub.setVisible(true);    sidebarLogoSub.setManaged(true);
            sidebarLabelPrincipal.setVisible(true); sidebarLabelPrincipal.setManaged(true);
            sidebarLabelParams.setVisible(true);    sidebarLabelParams.setManaged(true);
            navOperations.setText("⊞  Opérations");
            navDashboard.setText("▦  Dashboard");
            navConfig.setText("⚙  Configuration");
            navLogFull.setText("⌨  Log complet");
            sidebarToggleBtn.setText("◀");
        }
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
    @FXML private void openModelesSignatures() {
        configController.openModelesSignatures((Stage) navConfig.getScene().getWindow());
    }

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
