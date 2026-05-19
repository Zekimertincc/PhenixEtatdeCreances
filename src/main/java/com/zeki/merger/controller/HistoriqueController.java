package com.zeki.merger.controller;

import com.zeki.merger.AppPreferences;
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

import java.io.File;
import java.time.LocalDate;
import java.time.format.TextStyle;
import java.util.List;
import java.util.Locale;
import java.util.concurrent.ExecutorService;
import java.util.function.Consumer;

public class HistoriqueController {

    private final DatabaseManager    db;
    private final MonthClotureService monthService;
    private final TrfGeneratorService trfService;
    private final Consumer<String>   log;
    private final ExecutorService    executor;
    private final Runnable           showOperations;

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
        ScrollPane scroll = new ScrollPane(list);
        scroll.setFitToWidth(true);
        scroll.setStyle("-fx-background-color:transparent;-fx-border-color:transparent;");
        VBox.setVgrow(scroll, Priority.ALWAYS);
        container.getChildren().add(scroll);

        List<TrfMonthRecord> months = monthService.getAllMonths();

        LocalDate now = LocalDate.now();
        boolean currentClosed = months.stream()
            .anyMatch(m -> m.year() == now.getYear() && m.month() == now.getMonthValue()
                       && "closed".equals(m.status()));

        if (!currentClosed) {
            list.getChildren().add(buildMonthCard(now.getYear(), now.getMonthValue(),
                false, 0, 0.0, List.of()));
        }

        for (TrfMonthRecord m : months) {
            List<TrfHistoryRecord> clients = monthService.getHistoryForMonth(m.id());
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

    private void cloturerMois(int year, int month) {
        String consoPath   = AppPreferences.getTrfConso();
        String listingPath = AppPreferences.getTrfListing();
        String tableauPath = AppPreferences.getTrfTableau();

        if (consoPath.isEmpty() || listingPath.isEmpty() || tableauPath.isEmpty()) {
            log.accept("ERROR: Configurez les fichiers TRF avant de clôturer.");
            showOperations.run();
            return;
        }

        File consoFile   = new File(consoPath);
        File listingFile = new File(listingPath);
        File tableauFile = new File(tableauPath);

        if (!consoFile.exists() || !listingFile.exists() || !tableauFile.exists()) {
            log.accept("ERROR: Fichiers TRF introuvables.");
            showOperations.run();
            return;
        }

        log.accept("Clôture du mois " + month + "/" + year + "…");

        executor.submit(() -> {
            try {
                monthService.cloturerMois(year, month, consoFile, listingFile, tableauFile);
                Platform.runLater(() -> {
                    log.accept("Mois " + month + "/" + year + " clôturé.");
                    load(currentContainer);
                });
            } catch (Exception e) {
                Platform.runLater(() -> log.accept("ERREUR clôture : " + e.getMessage()));
            }
        });
    }

    private String capitalize(String s) {
        if (s == null || s.isEmpty()) return s;
        return Character.toUpperCase(s.charAt(0)) + s.substring(1);
    }

    private String etatDot(String etat) {
        if (etat == null) return "●";
        String lower = etat.toLowerCase();
        if (lower.contains("comp") && !lower.contains("non") && !lower.contains("partiel")) return "🟢";
        if (lower.contains("partiel")) return "🟡";
        return "🔴";
    }
}
