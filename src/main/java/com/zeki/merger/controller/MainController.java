package com.zeki.merger.controller;

import com.zeki.merger.AppPreferences;
import com.zeki.merger.db.DatabaseManager;
import com.zeki.merger.service.EspacePartageFixer;
import com.zeki.merger.service.EtatPublicGenerator;
import com.zeki.merger.service.MergeService;
import com.zeki.merger.service.ProcreancesComparator;
import com.zeki.merger.trf.TrfGeneratorService;
import javafx.application.Platform;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Modality;
import javafx.stage.Stage;

import java.awt.Desktop;
import java.io.File;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class MainController {

    // ---- FXML injected fields ----

    @FXML private HBox        badgesBox;
    @FXML private Label       missingFilesLabel;
    @FXML private GridPane    actionsGrid;
    @FXML private ProgressBar progressBar;
    @FXML private TextArea    logArea;
    @FXML private HBox        statusBar;
    @FXML private Label       statusLabel;
    @FXML private Button      openFileBtn;
    @FXML private DashboardController dashboardController;

    // ---- Programmatic action buttons (stored for enable/disable) ----

    private Button trfBtn, etatBtn, cmpBtn, fixBtn, runActionBtn;

    // ---- Services ----

    private final MergeService          mergeService          = new MergeService(DatabaseManager.getInstance());
    private final EspacePartageFixer    espacePartageFixer    = new EspacePartageFixer();
    private final EtatPublicGenerator   etatPublicGenerator   = new EtatPublicGenerator();
    private final TrfGeneratorService   trfGeneratorService   = new TrfGeneratorService(DatabaseManager.getInstance());
    private final ProcreancesComparator procreancesComparator = new ProcreancesComparator();

    private final ExecutorService executor = Executors.newSingleThreadExecutor(r -> {
        Thread t = new Thread(r, "merge-worker");
        t.setDaemon(true);
        return t;
    });
    private static final DateTimeFormatter TIME_FMT = DateTimeFormatter.ofPattern("HH:mm:ss");
    private File lastOutputFile;

    // -------------------------------------------------------------------------
    // Lifecycle
    // -------------------------------------------------------------------------

    @FXML
    public void initialize() {
        progressBar.setProgress(0);
        statusBar.setVisible(false);
        refreshFileBadges();

        trfBtn       = createActionBtn("Générer TRF",                 "Calcul virements et compensations",      "secondary-btn", e -> generateTrf());
        etatBtn      = createActionBtn("États Publics",               "Exporter vers EspacePartagé",            "secondary-btn", e -> generateEtatPublic());
        cmpBtn       = createActionBtn("Comparer des fichiers Excel", "Détecter les écarts PROCREANCES",        "secondary-btn", e -> compareProcreances());
        fixBtn       = createActionBtn("Corriger EspacePartagé",      "Mettre à jour les chemins",              "secondary-btn", e -> fixPaths());
        runActionBtn = createActionBtn("▶  CONSOLIDER",               "Lire les états → ConsolidationGénérale", "run-btn",       e -> run());

        actionsGrid.add(trfBtn,       0, 0);
        actionsGrid.add(etatBtn,      1, 0);
        actionsGrid.add(cmpBtn,       0, 1);
        actionsGrid.add(fixBtn,       1, 1);
        GridPane.setColumnSpan(runActionBtn, 2);
        actionsGrid.add(runActionBtn, 0, 2);
    }

    // -------------------------------------------------------------------------
    // File configuration
    // -------------------------------------------------------------------------

    @FXML
    private void openFileConfig() {
        String[] paths = {
            AppPreferences.getMergeRoot(),
            AppPreferences.getOutputFolder(),
            AppPreferences.getTrfConso(),
            AppPreferences.getTrfListing(),
            AppPreferences.getTrfTableau(),
            AppPreferences.getProcreancesPath()
        };
        String[]  labels = {"Dossier source",        "Dossier de sortie",       "ConsolidationGénérale",
                             "Listing Cabinet Phénix", "Tableau de Bord",        "Export PROCREANCES"};
        boolean[] isDir  = {true,  true,  false, false, false, false};
        String[]  exts   = {null,  null,  "xlsx", "xlsx", "xlsx", "xls"};

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
            lbl.setStyle("-fx-font-weight: bold; -fx-font-family: 'Courier New', monospace;");

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

        // Footer
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
        missing += addBadge("Dossier source",       AppPreferences.getMergeRoot(),       true);
        missing += addBadge("Dossier sortie",        AppPreferences.getOutputFolder(),    true);
        missing += addBadge("ConsolidationGénérale", AppPreferences.getTrfConso(),        false);
        missing += addBadge("Listing",               AppPreferences.getTrfListing(),      false);
        missing += addBadge("Tableau de bord",       AppPreferences.getTrfTableau(),      false);
        missing += addBadge("PROCREANCES",           AppPreferences.getProcreancesPath(), false);

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

    // -------------------------------------------------------------------------
    // Action handlers
    // -------------------------------------------------------------------------

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
                        if (dashboardController != null) dashboardController.refresh();
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
                        if (dashboardController != null) dashboardController.refresh();
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
        File procFile    = new File(procPath);
        File consoFile   = new File(consoPath);
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

    @FXML
    private void openFile() {
        if (lastOutputFile != null && lastOutputFile.exists()) {
            try { Desktop.getDesktop().open(lastOutputFile); }
            catch (Exception e) { appendLog("Cannot open file: " + e.getMessage()); }
        }
    }

    // -------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------

    private Button createActionBtn(String name, String desc, String styleClass,
                                    EventHandler<ActionEvent> handler) {
        Label lName = new Label(name);
        lName.getStyleClass().add("action-btn-name");
        Label lDesc = new Label(desc);
        lDesc.getStyleClass().add("action-btn-desc");
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
        if (path.isEmpty()) {
            lbl.setText("(non configuré)");
            lbl.setStyle("-fx-text-fill: #FF4444; -fx-font-family: 'Courier New', monospace;");
        } else {
            boolean exists = isDir ? new File(path).isDirectory() : new File(path).exists();
            String display = path.length() > 60 ? "…" + path.substring(path.length() - 57) : path;
            lbl.setText(display);
            lbl.setStyle((exists ? "-fx-text-fill: #1a6b2e;" : "-fx-text-fill: #FF4444;")
                + " -fx-font-family: 'Courier New', monospace; -fx-font-size: 11px;");
        }
    }

    private File dialogPickDirectory(Stage owner, String title, String lastPath) {
        DirectoryChooser dc = new DirectoryChooser();
        dc.setTitle(title);
        if (!lastPath.isEmpty()) {
            File f = new File(lastPath);
            if (f.isDirectory()) dc.setInitialDirectory(f);
        }
        return dc.showDialog(owner);
    }

    private File dialogPickFile(Stage owner, String title, String lastPath, String ext) {
        FileChooser fc = new FileChooser();
        fc.setTitle(title);
        if (ext != null) {
            fc.getExtensionFilters().add(
                new FileChooser.ExtensionFilter("Excel Files", "*." + ext));
        }
        if (!lastPath.isEmpty()) {
            File parent = new File(lastPath).getParentFile();
            if (parent != null && parent.isDirectory()) fc.setInitialDirectory(parent);
        }
        return fc.showOpenDialog(owner);
    }

    private void setAllButtonsDisabled(boolean disabled) {
        if (trfBtn != null)       trfBtn.setDisable(disabled);
        if (etatBtn != null)      etatBtn.setDisable(disabled);
        if (cmpBtn != null)       cmpBtn.setDisable(disabled);
        if (fixBtn != null)       fixBtn.setDisable(disabled);
        if (runActionBtn != null) runActionBtn.setDisable(disabled);
    }

    private void appendLog(String message) {
        String ts = LocalTime.now().format(TIME_FMT);
        logArea.appendText("[" + ts + "] " + message + "\n");
    }
}
