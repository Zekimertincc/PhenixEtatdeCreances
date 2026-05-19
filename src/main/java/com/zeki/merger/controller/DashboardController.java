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
import javafx.scene.chart.BarChart;
import javafx.scene.chart.CategoryAxis;
import javafx.scene.chart.NumberAxis;
import javafx.scene.chart.XYChart;
import javafx.scene.control.*;
import javafx.scene.layout.*;

import java.io.File;
import java.time.LocalDate;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.function.Consumer;

public class DashboardController {

    private final DatabaseManager    db;
    private final MonthClotureService monthService;
    private final TrfGeneratorService trfService;
    private final Consumer<String>   log;
    private final ExecutorService    executor;

    private Label selectedCompanyItem = null;

    public DashboardController(DatabaseManager db, MonthClotureService monthService,
                                TrfGeneratorService trfService, Consumer<String> log,
                                ExecutorService executor) {
        this.db           = db;
        this.monthService = monthService;
        this.trfService   = trfService;
        this.log          = log;
        this.executor     = executor;
    }

    public void load(VBox container) {
        container.getChildren().clear();
        container.setPadding(new Insets(16, 24, 12, 24));
        container.setSpacing(12);

        List<TrfMonthRecord> months = monthService.getAllMonths();
        TrfMonthRecord latest = months.isEmpty() ? null : months.get(0);

        long   activeSocietes = latest != null ? latest.nbClients()    : 0;
        double totalMontant   = latest != null ? latest.totalMontant() : 0;
        double totalNousDoit  = latest != null ? latest.totalNousDoit(): 0;
        String dernierTrf     = latest != null
            ? String.format("%02d/%d", latest.month(), latest.year()) : "—";

        Button refreshBtn = new Button("↻  Actualiser les données");
        refreshBtn.getStyleClass().add("save-btn");
        refreshBtn.setOnAction(e -> refresh(container));

        Region kpiSpacer = new Region();
        HBox.setHgrow(kpiSpacer, Priority.ALWAYS);
        HBox kpiHeader = new HBox(kpiSpacer, refreshBtn);
        kpiHeader.setAlignment(Pos.CENTER_RIGHT);
        container.getChildren().add(kpiHeader);

        HBox kpiRow = new HBox(12,
            kpiCard("Sociétés actives",   String.valueOf(activeSocietes)),
            kpiCard("Montant à facturer", String.format("%.2f €", totalMontant)),
            kpiCard("Nous doit",          String.format("%.2f €", totalNousDoit)),
            kpiCard("Dernier TRF",        dernierTrf)
        );
        container.getChildren().add(kpiRow);

        List<TrfHistoryRecord> clients = latest != null
            ? monthService.getHistoryForMonth(latest.id()) : List.of();

        VBox leftPanel = new VBox(4);
        leftPanel.setPrefWidth(190);
        leftPanel.setMinWidth(190);
        leftPanel.setMaxWidth(190);
        leftPanel.setPadding(new Insets(0, 8, 0, 0));

        Label societeLabel = new Label("Sociétés");
        societeLabel.getStyleClass().add("files-card-title");
        leftPanel.getChildren().add(societeLabel);

        VBox companyListBox = new VBox(2);
        ScrollPane companyScroll = new ScrollPane(companyListBox);
        companyScroll.setFitToWidth(true);
        companyScroll.setStyle("-fx-background-color:transparent;-fx-border-color:transparent;");
        VBox.setVgrow(companyScroll, Priority.ALWAYS);
        leftPanel.getChildren().add(companyScroll);

        VBox rightPanel = new VBox(10);
        HBox.setHgrow(rightPanel, Priority.ALWAYS);
        rightPanel.setPadding(new Insets(0, 0, 0, 16));

        HBox bodyRow = new HBox(0, leftPanel, rightPanel);
        VBox.setVgrow(bodyRow, Priority.ALWAYS);
        container.getChildren().add(bodyRow);

        selectedCompanyItem = null;
        for (TrfHistoryRecord r : clients) {
            Label item = new Label(etatDot(r.etat()) + "  " + r.clientName());
            item.setMaxWidth(Double.MAX_VALUE);
            item.getStyleClass().add("company-list-item");
            item.setOnMouseClicked(e -> {
                if (selectedCompanyItem != null)
                    selectedCompanyItem.getStyleClass().remove("company-list-selected");
                selectedCompanyItem = item;
                item.getStyleClass().add("company-list-selected");
                showClientDetail(r, rightPanel);
            });
            companyListBox.getChildren().add(item);
        }

        if (!clients.isEmpty()) {
            Label first = (Label) companyListBox.getChildren().get(0);
            selectedCompanyItem = first;
            first.getStyleClass().add("company-list-selected");
            showClientDetail(clients.get(0), rightPanel);
        }
    }

    public void refresh(VBox container) {
        String consoPath   = AppPreferences.getTrfConso();
        String listingPath = AppPreferences.getTrfListing();
        String tableauPath = AppPreferences.getTrfTableau();

        if (consoPath.isEmpty()) {
            log.accept("[Dashboard] ConsolidationGénérale non configurée — ouvrez Configuration.");
            return;
        }
        if (listingPath.isEmpty()) {
            log.accept("[Dashboard] Listing Cabinet Phénix non configuré — ouvrez Configuration.");
            return;
        }
        if (tableauPath.isEmpty()) {
            log.accept("[Dashboard] Tableau de bord non configuré — ouvrez Configuration.");
            return;
        }
        File consoFile   = new File(consoPath);
        File listingFile = new File(listingPath);
        File tableauFile = new File(tableauPath);

        if (!consoFile.exists()) {
            log.accept("[Dashboard] Fichier introuvable : " + consoPath);
            return;
        }
        if (!listingFile.exists()) {
            log.accept("[Dashboard] Fichier introuvable : " + listingPath);
            return;
        }
        if (!tableauFile.exists()) {
            log.accept("[Dashboard] Fichier introuvable : " + tableauPath);
            return;
        }

        log.accept("[Dashboard] Chargement des données...");
        LocalDate now = LocalDate.now();

        executor.submit(() -> {
            try {
                monthService.saveOpenMonth(now.getYear(), now.getMonthValue(),
                    consoFile, listingFile, tableauFile);
                List<TrfHistoryRecord> loaded = monthService
                    .getHistoryForMonth(db.getTrfMonthId(now.getYear(), now.getMonthValue()));
                Platform.runLater(() -> {
                    log.accept("[Dashboard] ✓ " + loaded.size() + " sociétés chargées");
                    load(container);
                });
            } catch (Exception e) {
                Platform.runLater(() -> log.accept("[Dashboard] ERREUR : " + e.getMessage()));
            }
        });
    }

    private VBox kpiCard(String label, String value) {
        Label lbl = new Label(label);
        lbl.getStyleClass().add("kpi-label");
        Label val = new Label(value);
        val.getStyleClass().add("kpi-value");
        VBox card = new VBox(4, lbl, val);
        card.getStyleClass().add("kpi-card");
        HBox.setHgrow(card, Priority.ALWAYS);
        return card;
    }

    private void showClientDetail(TrfHistoryRecord r, VBox detailPanel) {
        detailPanel.getChildren().clear();

        String etatText = r.etat() != null ? r.etat() : "—";
        Label nameLabel = new Label(r.clientName()
            + (r.clientCode() != null && !r.clientCode().isBlank() ? "  [" + r.clientCode() + "]" : ""));
        nameLabel.getStyleClass().add("detail-company-name");
        Label etatBadge = new Label(etatText);
        etatBadge.getStyleClass().addAll("etat-badge", etatCssClass(etatText));
        HBox header = new HBox(12, nameLabel, etatBadge);
        header.setAlignment(Pos.CENTER_LEFT);
        detailPanel.getChildren().add(header);

        HBox metrics = new HBox(10,
            metricCard("Encaissements CZ Phénix", String.format("%.2f €", r.encaissements())),
            metricCard("Montant à facturer TTC",   String.format("%.2f €", r.montantFacturer())),
            metricCard("Nous doit précédemment",   String.format("%.2f €", r.nousDoit())),
            metricCard("Sommes à reverser",         String.format("%.2f €", r.sommesReverser()))
        );
        detailPanel.getChildren().add(metrics);

        try {
            List<double[]> history = db.getClientMonthlyHistory(r.clientName(), 6);
            if (!history.isEmpty()) {
                CategoryAxis xAxis = new CategoryAxis();
                NumberAxis   yAxis = new NumberAxis();
                yAxis.setLabel("€");
                BarChart<String, Number> chart = new BarChart<>(xAxis, yAxis);
                chart.setTitle("Montant à facturer — 6 derniers mois");
                chart.setLegendVisible(false);
                chart.setPrefHeight(180);
                chart.setAnimated(false);
                XYChart.Series<String, Number> series = new XYChart.Series<>();
                for (double[] pt : history) {
                    series.getData().add(new XYChart.Data<>(
                        String.format("%02d/%d", (int) pt[0], (int) pt[1]), pt[2]));
                }
                chart.getData().add(series);
                detailPanel.getChildren().add(chart);
            }
        } catch (Exception ignored) {}
    }

    private VBox metricCard(String label, String value) {
        Label lbl = new Label(label);
        lbl.getStyleClass().add("dp-metric-label");
        Label val = new Label(value);
        val.getStyleClass().add("dp-metric-value");
        VBox card = new VBox(4, lbl, val);
        card.getStyleClass().add("dp-metric-card");
        HBox.setHgrow(card, Priority.ALWAYS);
        return card;
    }

    private String etatCssClass(String etat) {
        if (etat == null) return "etat-debit";
        String lower = etat.toLowerCase();
        if (lower.contains("comp") && !lower.contains("non") && !lower.contains("partiel"))
            return "etat-comp";
        if (lower.contains("partiel"))
            return "etat-noncomp";
        return "etat-debit";
    }

    private String etatDot(String etat) {
        if (etat == null) return "●";
        String lower = etat.toLowerCase();
        if (lower.contains("comp") && !lower.contains("non") && !lower.contains("partiel")) return "🟢";
        if (lower.contains("partiel")) return "🟡";
        return "🔴";
    }
}
