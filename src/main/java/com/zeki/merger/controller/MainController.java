package com.zeki.merger.controller;

import com.zeki.merger.AppConfig;
import com.zeki.merger.service.MergeService;
import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.layout.HBox;
import javafx.stage.DirectoryChooser;
import javafx.stage.Stage;

import java.awt.Desktop;
import java.io.File;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

/**
 * JavaFX controller for main.fxml.
 * Manages user interaction and delegates work to {@link MergeService} on a
 * background thread, then pushes progress updates back to the FX thread.
 */
public class MainController {

    // ---- FXML injected fields ----

    @FXML private TextField inputFolderField;
    @FXML private TextField outputFolderField;
    @FXML private Button    browseInputBtn;
    @FXML private Button    browseOutputBtn;
    @FXML private Label     configInfoLabel;
    @FXML private Button    runBtn;
    @FXML private ProgressBar progressBar;
    @FXML private TextArea  logArea;
    @FXML private HBox      statusBar;
    @FXML private Label     statusLabel;
    @FXML private Button    openFileBtn;

    // ---- private state ----

    private final MergeService mergeService = new MergeService();
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
        inputFolderField.setText(AppConfig.DEFAULT_ROOT_PATH);
        outputFolderField.setText(AppConfig.DEFAULT_OUTPUT_PATH);

        configInfoLabel.setText(
            "Target subfolder: \"" + AppConfig.TARGET_SUBFOLDER + "\""
            + "   |   File prefix: \"" + AppConfig.FILE_PREFIX + "\""
            + "   |   Filter column: " + AppConfig.FILTER_COLUMN_LABEL
            + "  (index " + AppConfig.FILTER_COLUMN_INDEX + ")"
        );

        progressBar.setProgress(0);
        statusBar.setVisible(false);
    }

    // -------------------------------------------------------------------------
    // FXML action handlers
    // -------------------------------------------------------------------------

    @FXML
    private void browseInput() {
        File chosen = pickDirectory("Select Root Folder to Scan", inputFolderField.getText());
        if (chosen != null) inputFolderField.setText(chosen.getAbsolutePath());
    }

    @FXML
    private void browseOutput() {
        File chosen = pickDirectory("Select Output Folder", outputFolderField.getText());
        if (chosen != null) outputFolderField.setText(chosen.getAbsolutePath());
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

        // Reset UI state
        runBtn.setDisable(true);
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
                    }
                    runBtn.setDisable(false);
                });
            } catch (Exception e) {
                Platform.runLater(() -> {
                    appendLog("FATAL: " + e.getMessage());
                    runBtn.setDisable(false);
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

    private void appendLog(String message) {
        String ts = LocalTime.now().format(TIME_FMT);
        logArea.appendText("[" + ts + "] " + message + "\n");
    }
}
