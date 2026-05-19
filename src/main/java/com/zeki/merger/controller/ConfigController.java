package com.zeki.merger.controller;

import com.zeki.merger.AppPreferences;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Modality;
import javafx.stage.Stage;

import java.io.File;
import java.util.function.Consumer;

public class ConfigController {

    private final VBox           configFormBox;
    private final FlowPane       badgesPane;
    private final Label          missingFilesLabel;
    private final Consumer<String> log;
    private final Runnable       onSaved;

    private String[] configPaths;
    private Label[]  configPathLabels;

    public ConfigController(VBox configFormBox, FlowPane badgesPane, Label missingFilesLabel,
                             Consumer<String> log, Runnable onSaved) {
        this.configFormBox    = configFormBox;
        this.badgesPane       = badgesPane;
        this.missingFilesLabel = missingFilesLabel;
        this.log              = log;
        this.onSaved          = onSaved;
    }

    public void load() {
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

    public void save() {
        AppPreferences.setMergeRoot(configPaths[0]);
        AppPreferences.setTrfConso(configPaths[1]);
        AppPreferences.setTrfListing(configPaths[2]);
        AppPreferences.setTrfTableau(configPaths[3]);
        refreshBadges();
        log.accept("Configuration enregistrée.");
        onSaved.run();
    }

    public void refreshBadges() {
        badgesPane.getChildren().clear();
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

    public void openFileConfig() {
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
        dialog.initOwner(badgesPane.getScene().getWindow());
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
            refreshBadges();
        });
        footer.getChildren().addAll(cancelBtn, saveBtn);
        root.getChildren().add(footer);

        Scene scene = new Scene(root);
        if (!badgesPane.getScene().getStylesheets().isEmpty()) {
            scene.getStylesheets().addAll(badgesPane.getScene().getStylesheets());
        }
        dialog.setScene(scene);
        dialog.showAndWait();
    }

    private int addBadge(String label, String path, boolean isDirectory) {
        boolean ok = !path.isEmpty()
            && (isDirectory ? new File(path).isDirectory() : new File(path).exists());
        Label badge = new Label(label + (ok ? " ✓" : " ✗"));
        badge.getStyleClass().add(ok ? "badge-ok" : "badge-missing");
        badgesPane.getChildren().add(badge);
        return ok ? 0 : 1;
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
}
