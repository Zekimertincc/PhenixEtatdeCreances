package com.zeki.merger.ui;

import com.zeki.merger.db.DatabaseManager;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.stage.Modality;
import javafx.stage.Stage;

public class ModelesSignaturesDialog {

    private final Stage dialog;
    private final DatabaseManager db;

    public ModelesSignaturesDialog(Stage owner, DatabaseManager db) {
        this.db = db;
        this.dialog = new Stage();
        dialog.initModality(Modality.APPLICATION_MODAL);
        dialog.initOwner(owner);
        dialog.setTitle("Modèles & Signatures");
        dialog.setResizable(false);

        TabPane tabPane = new TabPane();
        tabPane.setTabClosingPolicy(TabPane.TabClosingPolicy.UNAVAILABLE);

        // --- Tab 1: Objets ---
        TextField objetAccuseField = new TextField(
            db.getMailConfig("objet_accuse", "Cabinet Phénix, accusé de réception de dossier(s)"));
        TextField objetFacturationField = new TextField(
            db.getMailConfig("objet_facturation", "Cabinet Phénix, votre état des créances"));

        GridPane objetsGrid = new GridPane();
        objetsGrid.setHgap(10); objetsGrid.setVgap(12);
        objetsGrid.setPadding(new Insets(20));
        objetsGrid.addRow(0, new Label("Accusé de réception :"), objetAccuseField);
        objetsGrid.addRow(1, new Label("Facturation :"),         objetFacturationField);
        ColumnConstraints col1 = new ColumnConstraints(160);
        ColumnConstraints col2 = new ColumnConstraints(350);
        objetsGrid.getColumnConstraints().addAll(col1, col2);

        Tab tabObjets = new Tab("Objets", objetsGrid);

        // --- Tab 2: Signatures ---
        TextField julienNomField    = new TextField(db.getMailConfig("signature_julien_nom",    "Julien JOUSSET"));
        TextField julienTitreField  = new TextField(db.getMailConfig("signature_julien_titre",  "Directeur Associé"));
        TextField julienTelField    = new TextField(db.getMailConfig("signature_julien_tel",    "+33 (0)6 72 86 38 78"));
        TextField gauthierNomField  = new TextField(db.getMailConfig("signature_gauthier_nom",  "Gauthier BERIS"));
        TextField gauthierTitreField= new TextField(db.getMailConfig("signature_gauthier_titre","Directeur Associé"));
        TextField gauthierTelField  = new TextField(db.getMailConfig("signature_gauthier_tel",  "+33 (0)6 22 19 61 78"));

        GridPane sigGrid = new GridPane();
        sigGrid.setHgap(10); sigGrid.setVgap(12);
        sigGrid.setPadding(new Insets(20));
        sigGrid.addRow(0, new Label("— Julien JOUSSET —"));
        sigGrid.addRow(1, new Label("Nom :"),        julienNomField);
        sigGrid.addRow(2, new Label("Titre :"),      julienTitreField);
        sigGrid.addRow(3, new Label("Tél. mobile :"),julienTelField);
        sigGrid.addRow(4, new Label(""));
        sigGrid.addRow(5, new Label("— Gauthier BERIS —"));
        sigGrid.addRow(6, new Label("Nom :"),        gauthierNomField);
        sigGrid.addRow(7, new Label("Titre :"),      gauthierTitreField);
        sigGrid.addRow(8, new Label("Tél. mobile :"),gauthierTelField);
        ColumnConstraints sc1 = new ColumnConstraints(140);
        ColumnConstraints sc2 = new ColumnConstraints(300);
        sigGrid.getColumnConstraints().addAll(sc1, sc2);

        Tab tabSignatures = new Tab("Signatures", sigGrid);

        // --- Tab 3: Modèles (info only) ---
        Label infoLabel = new Label(
            "Les modèles de corps de mail se gèrent directement\n" +
            "depuis les fenêtres Accusé et Facturation\n" +
            "(bouton « Enregistrer » / « Mes modèles »).");
        infoLabel.setStyle("-fx-text-fill: #555; -fx-font-size: 12px;");
        StackPane modelePane = new StackPane(infoLabel);
        modelePane.setPadding(new Insets(30));
        Tab tabModeles = new Tab("Modèles", modelePane);

        tabPane.getTabs().addAll(tabObjets, tabSignatures, tabModeles);

        // --- Footer ---
        Button cancelBtn = new Button("Annuler");
        cancelBtn.setOnAction(e -> dialog.close());
        Button saveBtn = new Button("Enregistrer");
        saveBtn.setDefaultButton(true);
        saveBtn.setOnAction(e -> {
            db.setMailConfig("objet_accuse",           objetAccuseField.getText().trim());
            db.setMailConfig("objet_facturation",      objetFacturationField.getText().trim());
            db.setMailConfig("signature_julien_nom",   julienNomField.getText().trim());
            db.setMailConfig("signature_julien_titre", julienTitreField.getText().trim());
            db.setMailConfig("signature_julien_tel",   julienTelField.getText().trim());
            db.setMailConfig("signature_gauthier_nom",   gauthierNomField.getText().trim());
            db.setMailConfig("signature_gauthier_titre", gauthierTitreField.getText().trim());
            db.setMailConfig("signature_gauthier_tel",   gauthierTelField.getText().trim());
            dialog.close();
        });

        HBox footer = new HBox(8, cancelBtn, saveBtn);
        footer.setAlignment(Pos.CENTER_RIGHT);
        footer.setPadding(new Insets(10, 20, 16, 20));

        VBox root = new VBox(tabPane, footer);
        Scene scene = new Scene(root);
        if (owner.getScene() != null && !owner.getScene().getStylesheets().isEmpty()) {
            scene.getStylesheets().addAll(owner.getScene().getStylesheets());
        }
        dialog.setScene(scene);
    }

    public void show() {
        dialog.showAndWait();
    }
}
