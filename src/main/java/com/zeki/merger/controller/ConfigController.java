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
                AppPreferences.getTrfTableau(),
                AppPreferences.getControlePath(),
                AppPreferences.getRecupFacturePath(),
                AppPreferences.getTableauBordPath(),
                AppPreferences.getFacturationMensuelPath(),
                AppPreferences.getEntetePdfPath(),
                AppPreferences.getTrfOutput(),
                AppPreferences.getCorrespondancePath()
        };
        String[] labels = {
                "Dossier source (Dropbox)",
                "ConsolidationGénérale.xlsx",
                "Listing Cabinet Phénix.xls",
                "Tableau de bord facturation.xlsx",
                "Contrôle Facturation.xlsx",
                "Récup Num Facture.xlsx",
                "Tableau de bord soldes.xlsx",
                "Facturation mensuel (dossier)",
                "En-tête PDF (Phénix)",
                "TRF output (classement PDF)",
                "Correspondance clients"
        };
        String[] descriptions = {
                "Dossier racine contenant tous les dossiers clients",
                "Classement de la consolidation générale",
                "Sélectionner le fichier de consolidation",
                "Listing principal de tous les clients Cabinet Phénix",
                "Utilisé pour la génération du TRF",
                "Pour comparaison avec la consolidation",
                "Sélectionner après avoir généré le Contrôle Facturation",
                "Pour le numéro de facture et le nom des factures",
                "Pour reporter les soldes clients",
                "Classer les factures chez Phénix",
                "Fichier PDF de l'en-tête Cabinet Phénix",
                "Pour le classement des factures (PDF)",
                "Fichier de correspondance client ↔ espace partagé (pour les mails)"
        };
        boolean[] isDir = {true, false, false, false, false, false, false, true, false, false, false};
        String[]  exts  = {null, "xlsx", "xls", "xlsx", "xlsx", "xlsx", "xlsx", null, "pdf", "xlsx", "xlsx"};
        configPathLabels = new Label[configPaths.length];

        configFormBox.getChildren().clear();
        for (int i = 0; i < configPaths.length; i++) {
            final int idx = i;
            HBox row = new HBox(8);
            row.setAlignment(Pos.CENTER_LEFT);
            Label lbl = new Label(labels[i] + ":");
            lbl.setStyle("-fx-font-weight: bold;");
            Label desc = new Label(descriptions[i]);
            desc.setStyle("-fx-text-fill: #888; -fx-font-size: 10px;");
            VBox labelBox = new VBox(2, lbl, desc);
            labelBox.setMinWidth(220);
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
                        : dialogPickFile(null, labels[idx], configPaths[idx], exts[idx]);
                if (chosen != null) {
                    configPaths[idx] = chosen.getAbsolutePath();
                    updatePathLabel(configPathLabels[idx], configPaths[idx], isDir[idx]);
                }
            });
            row.getChildren().addAll(labelBox, configPathLabels[i], browseBtn);
            configFormBox.getChildren().add(row);
        }
    }

    public void save() {
        AppPreferences.setMergeRoot(configPaths[0]);
        AppPreferences.setTrfConso(configPaths[1]);
        AppPreferences.setTrfListing(configPaths[2]);
        AppPreferences.setTrfTableau(configPaths[3]);
        AppPreferences.setControlePath(configPaths[4]);
        AppPreferences.setRecupFacturePath(configPaths[5]);
        if (configPaths.length > 6) AppPreferences.setTableauBordPath(configPaths[6]);
        if (configPaths.length > 7) AppPreferences.setFacturationMensuelPath(configPaths[7]);
        if (configPaths.length > 8) AppPreferences.setEntetePdfPath(configPaths[8]);
        if (configPaths.length > 9)  AppPreferences.setTrfOutput(configPaths[9]);
        if (configPaths.length > 10) AppPreferences.setCorrespondancePath(configPaths[10]);
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
        addBadge("TRF output",            AppPreferences.getTrfOutput(),        false);
        missing += addBadge("Tableau de bord",        AppPreferences.getTrfTableau(),       false);
        missing += addBadge("PROCREANCES",            AppPreferences.getProcreancesPath(),  false);
        missing += addBadge("Contrôle Fact.",         AppPreferences.getControlePath(),     false);
        missing += addBadge("Récup Factures",         AppPreferences.getRecupFacturePath(),        false);
        missing += addBadge("Fact. Mensuel",          AppPreferences.getFacturationMensuelPath(),  true);
        missing += addBadge("En-tête PDF",            AppPreferences.getEntetePdfPath(),           false);
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
                AppPreferences.getRecupFacturePath(),
                AppPreferences.getTableauBordPath(),
                AppPreferences.getFacturationMensuelPath(),
                AppPreferences.getEntetePdfPath(),
                AppPreferences.getTrfOutput()
        };
        String[]  labels = {"Dossier source", "Dossier de sortie", "ConsolidationGénérale",
                "Listing Cabinet Phénix", "Tableau de Bord", "Export PROCREANCES",
                "Contrôle Facturation", "Récup. Num Facture", "Tableau de bord soldes",
                "Facturation mensuel", "En-tête PDF (Phénix)", "TRF output (classement PDF)"};
        String[] descs = {
                "Dossier racine contenant tous les dossiers clients",
                "Classement de la consolidation générale",
                "Sélectionner le fichier de consolidation",
                "Listing principal de tous les clients Cabinet Phénix",
                "Utilisé pour la génération du TRF",
                "Pour comparaison avec la consolidation",
                "Sélectionner après avoir généré le Contrôle Facturation",
                "Pour le numéro de facture et le nom des factures",
                "Pour reporter les soldes clients",
                "Classer les factures chez Phénix",
                "Fichier PDF de l'en-tête Cabinet Phénix",
                "Pour le classement des factures (PDF)"};
        boolean[] isDir  = {true, true, false, false, false, false, false, false, false, true, false, false};
        String[]  exts   = {null, null, "xlsx", "xlsx", "xlsx", "xls", "xlsx", "xlsx", "xlsx", null, "pdf", "xlsx"};

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
            lbl.setStyle("-fx-font-weight: bold;");
            Label descLbl = new Label(i < descs.length ? descs[i] : "");
            descLbl.setStyle("-fx-text-fill: #888; -fx-font-size: 10px;");
            VBox labelBox = new VBox(2, lbl, descLbl);
            labelBox.setMinWidth(200);
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
            row.getChildren().addAll(labelBox, pathLabels[i], changeBtn);
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
            AppPreferences.setTableauBordPath(paths[8]);
            AppPreferences.setFacturationMensuelPath(paths[9]);
            AppPreferences.setEntetePdfPath(paths[10]);
            if (paths.length > 11) AppPreferences.setTrfOutput(paths[11]);
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
            if (ext.equals("xls") || ext.equals("xlsx")) {
                fc.getExtensionFilters().add(
                    new FileChooser.ExtensionFilter("Fichiers Excel", "*.xls", "*.xlsx"));
            } else {
                fc.getExtensionFilters().add(
                    new FileChooser.ExtensionFilter("Fichiers " + ext.toUpperCase(), "*." + ext));
            }
        }
        if (lastPath != null && !lastPath.isEmpty()) {
            File parent = new File(lastPath).getParentFile();
            if (parent != null && parent.isDirectory()) fc.setInitialDirectory(parent);
        }
        return fc.showOpenDialog(owner);
    }
}