package com.zeki.merger.controller;

import com.zeki.merger.db.DatabaseManager;
import com.zeki.merger.db.TrfHistoryRecord;
import com.zeki.merger.db.TrfMonthRecord;
import com.zeki.merger.service.MonthClotureService;
import com.zeki.merger.trf.TrfGeneratorService;
import javafx.application.Platform;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.File;
import java.time.LocalDate;
import java.time.Month;
import java.time.format.TextStyle;
import java.util.List;
import java.util.Locale;
import java.util.concurrent.ExecutorService;
import java.util.function.Consumer;

public class HistoriqueController {

    private final DatabaseManager     db;
    private final MonthClotureService monthService;
    private final TrfGeneratorService trfService;
    private final Consumer<String>    log;
    private final ExecutorService     executor;
    private final Runnable            showOperations;

    private VBox currentContainer = null;

    public HistoriqueController(DatabaseManager db, MonthClotureService monthService,
                                 TrfGeneratorService trfService, Consumer<String> log,
                                 ExecutorService executor, Runnable showOperations) {
        this.db             = db;
        this.monthService   = monthService;
        this.trfService     = trfService;
        this.log            = log;
        this.executor       = executor;
        this.showOperations = showOperations;
    }

    public void load(VBox container) {
        this.currentContainer = container;
        container.getChildren().clear();

        VBox list = new VBox(10);
        list.setPadding(new Insets(16, 24, 12, 24));

        // Top bar — "Nouveau mois" button
        Button newMonthBtn = new Button("+ Nouveau mois");
        newMonthBtn.getStyleClass().add("save-btn");
        newMonthBtn.setOnAction(e -> openMonthDialog(null, container));
        HBox topBar = new HBox(newMonthBtn);
        topBar.setAlignment(Pos.CENTER_LEFT);
        topBar.setPadding(new Insets(0, 0, 8, 0));
        list.getChildren().add(topBar);

        ScrollPane scroll = new ScrollPane(list);
        scroll.setFitToWidth(true);
        scroll.setStyle("-fx-background-color:transparent;-fx-border-color:transparent;");
        VBox.setVgrow(scroll, Priority.ALWAYS);
        container.getChildren().add(scroll);

        List<TrfMonthRecord> months = monthService.getAllMonths();
        for (TrfMonthRecord m : months) {
            List<TrfHistoryRecord> clients = monthService.getHistoryForMonth(m.id());
            list.getChildren().add(buildMonthCard(m, clients, container));
        }

        if (months.isEmpty()) {
            Label empty = new Label("Aucun mois enregistré. Cliquez sur '+ Nouveau mois' pour commencer.");
            empty.setStyle("-fx-text-fill: #888; -fx-font-size: 13px;");
            list.getChildren().add(empty);
        }
    }

    // =========================================================================
    // Month card
    // =========================================================================

    private VBox buildMonthCard(TrfMonthRecord m, List<TrfHistoryRecord> clients,
                                 VBox container) {
        boolean closed = "closed".equals(m.status());

        VBox card = new VBox(0);
        card.getStyleClass().add("month-card");

        // Header
        HBox header = new HBox(10);
        header.setAlignment(Pos.CENTER_LEFT);
        header.getStyleClass().add("month-card-header");
        header.setPadding(new Insets(10, 12, 10, 12));

        String monthName = Month.of(m.month())
                .getDisplayName(TextStyle.FULL_STANDALONE, Locale.FRENCH);
        Label title = new Label("📅 " + capitalize(monthName) + " " + m.year());
        title.getStyleClass().add("month-name");
        HBox.setHgrow(title, Priority.ALWAYS);
        header.getChildren().add(title);

        Label stats = new Label(m.nbClients() + " sociétés   "
                + String.format("%.0f €", m.totalMontant()));
        stats.getStyleClass().add("month-stat");
        header.getChildren().add(stats);

        Label badge = new Label(closed ? "Clôturé" : "En cours");
        badge.getStyleClass().add(closed ? "badge-closed" : "badge-open");
        header.getChildren().add(badge);

        if (!closed) {
            Button modifierBtn = new Button("Modifier");
            modifierBtn.getStyleClass().add("action-btn");
            modifierBtn.setOnAction(e -> openMonthDialog(m, container));
            Button cloturerBtn = new Button("Clôturer");
            cloturerBtn.getStyleClass().add("cloture-btn");
            cloturerBtn.setOnAction(e -> cloturerMois(m, container));
            header.getChildren().addAll(modifierBtn, cloturerBtn);
        }

        card.getChildren().add(header);

        // Body — client list (collapsible)
        VBox body = new VBox(4);
        body.setPadding(new Insets(8, 12, 8, 12));
        body.setVisible(false);
        body.setManaged(false);
        body.getStyleClass().add("month-card-body");

        if (clients.isEmpty()) {
            body.getChildren().add(new Label("Aucune donnée chargée."));
        } else {
            HBox tHeader = clientRow("CLIENT", "ENCAISSEMENTS", "COMMISSIONS", "À REVERSER", "TYPE", true);
            body.getChildren().add(tHeader);
            for (TrfHistoryRecord r : clients) {
                body.getChildren().add(clientRow(
                        r.clientName(),
                        String.format("%.0f €", r.encaissements()),
                        String.format("%.0f €", r.montantFacturer()),
                        String.format("%.0f €", r.sommesReverser()),
                        r.nonCompensation() ? "NON COMP" : resolveEtatType(r.etat()),
                        false));
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

    // =========================================================================
    // Month dialog — create or modify
    // =========================================================================

    private void openMonthDialog(TrfMonthRecord existing, VBox container) {
        Stage dialog = new Stage();
        dialog.setTitle(existing == null ? "Nouveau mois" : "Modifier le mois");
        dialog.setWidth(520);
        dialog.setResizable(false);

        VBox root = new VBox(14);
        root.setPadding(new Insets(20));

        HBox periodRow = new HBox(10);
        periodRow.setAlignment(Pos.CENTER_LEFT);

        ComboBox<Integer> yearBox = new ComboBox<>();
        int currentYear = LocalDate.now().getYear();
        for (int y = currentYear - 3; y <= currentYear + 1; y++) yearBox.getItems().add(y);
        yearBox.setValue(existing != null ? existing.year() : currentYear);

        ComboBox<String> monthBox = new ComboBox<>();
        for (int mo = 1; mo <= 12; mo++) {
            monthBox.getItems().add(capitalize(Month.of(mo)
                    .getDisplayName(TextStyle.FULL_STANDALONE, Locale.FRENCH)));
        }
        monthBox.setValue(capitalize(Month.of(existing != null
                ? existing.month() : LocalDate.now().getMonthValue())
                .getDisplayName(TextStyle.FULL_STANDALONE, Locale.FRENCH)));

        periodRow.getChildren().addAll(
                new Label("Année :"), yearBox,
                new Label("Mois :"), monthBox);
        root.getChildren().addAll(new Label("Période :"), periodRow);

        String[] labels = {"ConsolidationGénérale :", "Listing Cabinet Phénix :", "Tableau de bord :"};
        TextField[] fields = new TextField[3];
        for (int i = 0; i < 3; i++) {
            fields[i] = new TextField();
            fields[i].setPromptText("Sélectionner un fichier...");
            fields[i].setPrefWidth(340);
            Button browseBtn = new Button("Parcourir");
            final int idx = i;
            browseBtn.setOnAction(e -> {
                FileChooser fc = new FileChooser();
                fc.getExtensionFilters().add(
                        new FileChooser.ExtensionFilter("Fichiers Excel", "*.xlsx", "*.xls"));
                File f = fc.showOpenDialog(dialog);
                if (f != null) fields[idx].setText(f.getAbsolutePath());
            });
            HBox row = new HBox(8, fields[i], browseBtn);
            row.setAlignment(Pos.CENTER_LEFT);
            root.getChildren().addAll(new Label(labels[i]), row);
        }

        HBox btnRow = new HBox(10);
        btnRow.setAlignment(Pos.CENTER_RIGHT);
        Button cancelBtn = new Button("Annuler");
        cancelBtn.setOnAction(e -> dialog.close());
        Button saveBtn = new Button("Enregistrer");
        saveBtn.getStyleClass().add("save-btn");
        saveBtn.setDefaultButton(true);
        saveBtn.setOnAction(e -> {
            String consoPath   = fields[0].getText().trim();
            String listingPath = fields[1].getText().trim();
            String tableauPath = fields[2].getText().trim();

            if (consoPath.isBlank() || listingPath.isBlank() || tableauPath.isBlank()) {
                new Alert(Alert.AlertType.WARNING,
                        "Veuillez sélectionner tous les fichiers.").showAndWait();
                return;
            }

            File consoFile   = new File(consoPath);
            File listingFile = new File(listingPath);
            File tableauFile = new File(tableauPath);

            if (!consoFile.exists() || !listingFile.exists() || !tableauFile.exists()) {
                new Alert(Alert.AlertType.WARNING,
                        "Un ou plusieurs fichiers sont introuvables.").showAndWait();
                return;
            }

            int year  = yearBox.getValue();
            int month = monthBox.getSelectionModel().getSelectedIndex() + 1;
            dialog.close();

            log.accept("Enregistrement " + month + "/" + year + "…");
            executor.submit(() -> {
                try {
                    monthService.saveOpenMonth(year, month, consoFile, listingFile, tableauFile);
                    Platform.runLater(() -> {
                        log.accept("✓ " + month + "/" + year + " enregistré.");
                        load(container);
                    });
                } catch (Exception ex) {
                    Platform.runLater(() ->
                            log.accept("ERREUR enregistrement : " + ex.getMessage()));
                }
            });
        });
        btnRow.getChildren().addAll(cancelBtn, saveBtn);
        root.getChildren().add(btnRow);

        javafx.scene.Scene scene = new javafx.scene.Scene(root);
        dialog.setScene(scene);
        dialog.show();
    }

    // =========================================================================
    // Clôturer
    // =========================================================================

    private void cloturerMois(TrfMonthRecord m, VBox container) {
        Alert confirm = new Alert(Alert.AlertType.CONFIRMATION,
                "Clôturer " + m.month() + "/" + m.year()
                        + " ? Cette action est irréversible.",
                ButtonType.YES, ButtonType.NO);
        confirm.setTitle("Confirmer la clôture");
        confirm.showAndWait().ifPresent(bt -> {
            if (bt == ButtonType.YES) {
                executor.submit(() -> {
                    try {
                        db.closeTrfMonth(m.year(), m.month());
                        Platform.runLater(() -> {
                            log.accept("Mois " + m.month() + "/" + m.year() + " clôturé.");
                            load(container);
                        });
                    } catch (Exception e) {
                        Platform.runLater(() ->
                                log.accept("ERREUR clôture : " + e.getMessage()));
                    }
                });
            }
        });
    }

    // =========================================================================
    // Helpers
    // =========================================================================

    private HBox clientRow(String c1, String c2, String c3, String c4, String c5, boolean header) {
        HBox row = new HBox();
        row.setAlignment(Pos.CENTER_LEFT);
        row.setPadding(new Insets(3, 4, 3, 4));
        if (header) row.setStyle("-fx-background-color: #2a2a2a;");

        String[] vals   = {c1, c2, c3, c4, c5};
        double[] widths = {0.35, 0.16, 0.16, 0.16, 0.14};

        for (int i = 0; i < vals.length; i++) {
            Label lbl = new Label(vals[i]);
            lbl.setStyle("-fx-font-size: " + (header ? "10" : "11") + "px;"
                    + (header ? " -fx-font-weight: bold; -fx-text-fill: #aaa;" : ""));
            lbl.setMaxWidth(Double.MAX_VALUE);
            HBox.setHgrow(lbl, Priority.ALWAYS);
            lbl.prefWidthProperty().bind(row.widthProperty().multiply(widths[i]));
            row.getChildren().add(lbl);
        }
        return row;
    }

    private String resolveEtatType(String etat) {
        if (etat == null) return "—";
        String e = etat.toLowerCase();
        if (e.contains("partiel")) return "COMP PART.";
        if (e.contains("debit") || e.contains("débit")) return "DÉBITEUR";
        return "VIREMENT";
    }

    private String capitalize(String s) {
        if (s == null || s.isEmpty()) return s;
        return Character.toUpperCase(s.charAt(0)) + s.substring(1);
    }

    private String etatDot(String etat) {
        if (etat == null) return "●";
        String lower = etat.toLowerCase();
        if (lower.contains("comp") && !lower.contains("non")
                && !lower.contains("partiel")) return "🟢";
        if (lower.contains("partiel")) return "🟡";
        return "🔴";
    }
}
