package com.zeki.merger.controller;

import com.zeki.merger.AppConfig;
import com.zeki.merger.AppPreferences;
import com.zeki.merger.db.DatabaseManager;
import com.zeki.merger.service.EspacePartageFixer;
import com.zeki.merger.service.EtatPublicGenerator;
import com.zeki.merger.service.MergeService;
import com.zeki.merger.trf.TrfGeneratorService;
import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.layout.HBox;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.awt.Desktop;
import java.io.File;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class MainController {

    // ---- FXML injected fields ----

    @FXML private TextField inputFolderField;
    @FXML private TextField outputFolderField;
    @FXML private Button    browseInputBtn;
    @FXML private Button    browseOutputBtn;
    @FXML private Label     trfConsoPathLabel;
    @FXML private Label     trfListingPathLabel;
    @FXML private Label     trfTableauPathLabel;
    @FXML private Button    trfConsoBtn;
    @FXML private Button    trfListingBtn;
    @FXML private Button    trfTableauBtn;
    @FXML private Button    fixPathsBtn;
    @FXML private Button    generateEtatPublicBtn;
    @FXML private Button    generateTrfBtn;
    @FXML private Button    runBtn;
    @FXML private ProgressBar progressBar;
    @FXML private TextArea  logArea;
    @FXML private HBox      statusBar;
    @FXML private Label     statusLabel;
    @FXML private Button    openFileBtn;
    @FXML private DashboardController dashboardController;

    // ---- private state ----

    private final MergeService        mergeService        = new MergeService(DatabaseManager.getInstance());
    private final EspacePartageFixer  espacePartageFixer  = new EspacePartageFixer();
    private final EtatPublicGenerator etatPublicGenerator = new EtatPublicGenerator();
    private final TrfGeneratorService trfGeneratorService = new TrfGeneratorService(DatabaseManager.getInstance());
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
        // Restore persisted folder paths (fall back to AppConfig defaults)
        String root   = AppPreferences.getMergeRoot();
        String output = AppPreferences.getOutputFolder();
        inputFolderField.setText(root.isEmpty()   ? AppConfig.DEFAULT_ROOT_PATH   : root);
        outputFolderField.setText(output.isEmpty() ? AppConfig.DEFAULT_OUTPUT_PATH : output);

        // Restore persisted TRF file paths
        trfConsoPathLabel.setText(AppPreferences.getTrfConso());
        trfListingPathLabel.setText(AppPreferences.getTrfListing());
        trfTableauPathLabel.setText(AppPreferences.getTrfTableau());

        progressBar.setProgress(0);
        statusBar.setVisible(false);
    }

    // -------------------------------------------------------------------------
    // Folder pickers
    // -------------------------------------------------------------------------

    @FXML
    private void browseInput() {
        File chosen = pickDirectory("Select Root Folder to Scan", inputFolderField.getText());
        if (chosen != null) {
            inputFolderField.setText(chosen.getAbsolutePath());
            AppPreferences.setMergeRoot(chosen.getAbsolutePath());
        }
    }

    @FXML
    private void browseOutput() {
        File chosen = pickDirectory("Select Output Folder", outputFolderField.getText());
        if (chosen != null) {
            outputFolderField.setText(chosen.getAbsolutePath());
            AppPreferences.setOutputFolder(chosen.getAbsolutePath());
        }
    }

    // -------------------------------------------------------------------------
    // TRF file pickers
    // -------------------------------------------------------------------------

    @FXML
    private void pickTrfConso() {
        File f = pickExcelFile("Sélectionner ConsolidationGénérale", AppPreferences.getTrfConso());
        if (f != null) {
            AppPreferences.setTrfConso(f.getAbsolutePath());
            trfConsoPathLabel.setText(f.getAbsolutePath());
        }
    }

    @FXML
    private void pickTrfListing() {
        File f = pickExcelFile("Sélectionner Listing Cabinet Phénix", AppPreferences.getTrfListing());
        if (f != null) {
            AppPreferences.setTrfListing(f.getAbsolutePath());
            trfListingPathLabel.setText(f.getAbsolutePath());
        }
    }

    @FXML
    private void pickTrfTableau() {
        File f = pickExcelFile("Sélectionner Tableau de Bord", AppPreferences.getTrfTableau());
        if (f != null) {
            AppPreferences.setTrfTableau(f.getAbsolutePath());
            trfTableauPathLabel.setText(f.getAbsolutePath());
        }
    }

    // -------------------------------------------------------------------------
    // Action handlers
    // -------------------------------------------------------------------------

    @FXML
    private void generateTrf() {
        String consoPath   = AppPreferences.getTrfConso();
        String listingPath = AppPreferences.getTrfListing();
        String tableauPath = AppPreferences.getTrfTableau();
        String outputPath  = outputFolderField.getText().trim();

        if (consoPath.isEmpty() || listingPath.isEmpty() || tableauPath.isEmpty()) {
            appendLog("ERROR: Sélectionnez les trois fichiers TRF avant de générer.");
            return;
        }
        File consoFile   = new File(consoPath);
        File listingFile = new File(listingPath);
        File tableauFile = new File(tableauPath);
        File outputFolder = new File(outputPath);

        if (!consoFile.exists()) {
            appendLog("ERROR: Fichier introuvable — " + consoPath); return;
        }
        if (!listingFile.exists()) {
            appendLog("ERROR: Fichier introuvable — " + listingPath); return;
        }
        if (!tableauFile.exists()) {
            appendLog("ERROR: Fichier introuvable — " + tableauPath); return;
        }
        if (!outputFolder.isDirectory()) {
            appendLog("ERROR: Output folder does not exist — " + outputPath); return;
        }

        setAllButtonsDisabled(true);
        statusBar.setVisible(false);
        progressBar.setProgress(0);
        logArea.clear();
        lastOutputFile = null;

        executor.submit(() -> {
            try {
                File result = trfGeneratorService.generate(
                    consoFile, listingFile, tableauFile, outputFolder,
                    (prog, msg) -> Platform.runLater(() -> {
                        progressBar.setProgress(prog);
                        appendLog(msg);
                    }));

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
                Platform.runLater(() -> {
                    appendLog("FATAL: " + e.getMessage());
                    setAllButtonsDisabled(false);
                });
            }
        });
    }

    @FXML
    private void generateEtatPublic() {
        File rootFolder = new File(inputFolderField.getText().trim());
        if (!rootFolder.isDirectory()) {
            appendLog("ERROR: Root folder does not exist — " + rootFolder.getAbsolutePath());
            return;
        }

        setAllButtonsDisabled(true);
        statusBar.setVisible(false);
        progressBar.setProgress(0);
        logArea.clear();
        lastOutputFile = null;

        executor.submit(() -> {
            try {
                etatPublicGenerator.generate(rootFolder, (prog, msg) ->
                    Platform.runLater(() -> {
                        progressBar.setProgress(prog);
                        appendLog(msg);
                    })
                );

                Platform.runLater(() -> {
                    statusLabel.setText("Etat Public files written to EspacePartagé paths.");
                    openFileBtn.setVisible(false);
                    statusBar.setVisible(true);
                    setAllButtonsDisabled(false);
                });
            } catch (Exception e) {
                Platform.runLater(() -> {
                    appendLog("FATAL: " + e.getMessage());
                    setAllButtonsDisabled(false);
                });
            }
        });
    }

    @FXML
    private void fixPaths() {
        File rootFolder = new File(inputFolderField.getText().trim());
        if (!rootFolder.isDirectory()) {
            appendLog("ERROR: Root folder does not exist — " + rootFolder.getAbsolutePath());
            return;
        }

        setAllButtonsDisabled(true);
        statusBar.setVisible(false);
        progressBar.setProgress(0);
        logArea.clear();
        lastOutputFile = null;

        executor.submit(() -> {
            try {
                File result = espacePartageFixer.fix(rootFolder, (progress, msg) ->
                    Platform.runLater(() -> {
                        progressBar.setProgress(progress);
                        appendLog(msg);
                    })
                );

                Platform.runLater(() -> {
                    lastOutputFile = result;
                    statusLabel.setText("Saved: " + result.getAbsolutePath());
                    openFileBtn.setVisible(true);
                    statusBar.setVisible(true);
                    setAllButtonsDisabled(false);
                });
            } catch (Exception e) {
                Platform.runLater(() -> {
                    appendLog("FATAL: " + e.getMessage());
                    setAllButtonsDisabled(false);
                });
            }
        });
    }

    @FXML
    private void run() {
        File rootFolder   = new File(inputFolderField.getText().trim());
        File outputFolder = new File(outputFolderField.getText().trim());

        if (!rootFolder.isDirectory()) {
            appendLog("ERROR: Root folder does not exist — " + rootFolder.getAbsolutePath());
            return;
        }
        if (!outputFolder.isDirectory()) {
            appendLog("ERROR: Output folder does not exist — " + outputFolder.getAbsolutePath());
            return;
        }

        setAllButtonsDisabled(true);
        statusBar.setVisible(false);
        progressBar.setProgress(0);
        logArea.clear();
        lastOutputFile = null;

        executor.submit(() -> {
            try {
                File result = mergeService.merge(rootFolder, outputFolder, (progress, msg) ->
                    Platform.runLater(() -> {
                        progressBar.setProgress(progress);
                        appendLog(msg);
                    })
                );

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
                Platform.runLater(() -> {
                    appendLog("FATAL: " + e.getMessage());
                    setAllButtonsDisabled(false);
                });
            }
        });
    }

    @FXML
    private void openFile() {
        if (lastOutputFile != null && lastOutputFile.exists()) {
            try {
                Desktop.getDesktop().open(lastOutputFile);
            } catch (Exception e) {
                appendLog("Cannot open file: " + e.getMessage());
            }
        }
    }

    // -------------------------------------------------------------------------
    // Helpers
    // -------------------------------------------------------------------------

    private File pickDirectory(String title, String initialPath) {
        DirectoryChooser dc = new DirectoryChooser();
        dc.setTitle(title);
        File initial = new File(initialPath);
        if (initial.isDirectory()) dc.setInitialDirectory(initial);
        Stage stage = (Stage) browseInputBtn.getScene().getWindow();
        return dc.showDialog(stage);
    }

    private File pickExcelFile(String title, String lastPath) {
        FileChooser fc = new FileChooser();
        fc.setTitle(title);
        fc.getExtensionFilters().add(
            new FileChooser.ExtensionFilter("Excel Files", "*.xlsx", "*.xls"));
        if (!lastPath.isEmpty()) {
            File parent = new File(lastPath).getParentFile();
            if (parent != null && parent.isDirectory()) fc.setInitialDirectory(parent);
        }
        Stage stage = (Stage) browseInputBtn.getScene().getWindow();
        return fc.showOpenDialog(stage);
    }

    private void setAllButtonsDisabled(boolean disabled) {
        fixPathsBtn.setDisable(disabled);
        generateEtatPublicBtn.setDisable(disabled);
        generateTrfBtn.setDisable(disabled);
        runBtn.setDisable(disabled);
    }

    private void appendLog(String message) {
        String ts = LocalTime.now().format(TIME_FMT);
        logArea.appendText("[" + ts + "] " + message + "\n");
    }
}
