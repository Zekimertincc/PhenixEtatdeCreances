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
import javafx.scene.chart.*;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.scene.paint.Color;
import javafx.scene.shape.Rectangle;

import java.io.File;
import java.time.LocalDate;
import java.time.format.TextStyle;
import java.util.*;
import java.util.concurrent.ExecutorService;
import java.util.function.Consumer;
import java.util.stream.Collectors;

public class DashboardController {

    private final DatabaseManager     db;
    private final MonthClotureService monthService;
    private final TrfGeneratorService trfService;
    private final Consumer<String>    log;
    private final ExecutorService     executor;

    // Filters
    private Integer filterYear  = null;
    private Integer filterMonth = null;

    public DashboardController(DatabaseManager db, MonthClotureService monthService,
                                TrfGeneratorService trfService, Consumer<String> log,
                                ExecutorService executor) {
        this.db           = db;
        this.monthService = monthService;
        this.trfService   = trfService;
        this.log          = log;
        this.executor     = executor;
    }

    // =========================================================================
    // Load
    // =========================================================================

    public void load(VBox container) {
        container.getChildren().clear();
        container.setPadding(new Insets(16, 24, 12, 24));
        container.setSpacing(12);

        List<TrfMonthRecord> allMonths = monthService.getAllMonths();

        // Top bar — filters + refresh
        container.getChildren().add(buildTopBar(container, allMonths));

        if (allMonths.isEmpty()) {
            Label empty = new Label("Aucune donnée — cliquez sur Actualiser pour charger les données.");
            empty.setStyle("-fx-text-fill: #888; -fx-font-size: 13px;");
            container.getChildren().add(empty);
            return;
        }

        // Get filtered data
        List<TrfHistoryRecord> records = getFilteredRecords(allMonths);

        // KPI row
        container.getChildren().add(buildKpiRow(records, allMonths));

        // Charts + client table
        HBox mainRow = new HBox(16);
        VBox.setVgrow(mainRow, Priority.ALWAYS);

        VBox leftCol  = new VBox(12);
        leftCol.setPrefWidth(420);
        leftCol.setMinWidth(320);

        VBox rightCol = new VBox(12);
        HBox.setHgrow(rightCol, Priority.ALWAYS);

        leftCol.getChildren().addAll(
                buildMonthlyChart(allMonths),
                buildEtatChart(records)
        );
        rightCol.getChildren().add(buildClientTable(records));

        mainRow.getChildren().addAll(leftCol, rightCol);
        container.getChildren().add(mainRow);
    }

    // =========================================================================
    // Top bar — period filter + refresh
    // =========================================================================

    private HBox buildTopBar(VBox container, List<TrfMonthRecord> months) {
        HBox bar = new HBox(10);
        bar.setAlignment(Pos.CENTER_LEFT);

        Label periodLabel = new Label("Période :");
        periodLabel.setStyle("-fx-font-size: 12px;");

        ComboBox<String> periodBox = new ComboBox<>();
        periodBox.getItems().add("Tout l'historique");
        for (TrfMonthRecord m : months) {
            periodBox.getItems().add(monthLabel(m.month(), m.year()));
        }
        periodBox.setValue(filterYear == null ? "Tout l'historique"
                : monthLabel(filterMonth, filterYear));
        periodBox.setPrefWidth(160);
        periodBox.setOnAction(e -> {
            String selected = periodBox.getValue();
            if ("Tout l'historique".equals(selected)) {
                filterYear = null; filterMonth = null;
            } else {
                for (TrfMonthRecord m : months) {
                    if (monthLabel(m.month(), m.year()).equals(selected)) {
                        filterYear  = m.year();
                        filterMonth = m.month();
                        break;
                    }
                }
            }
            load(container);
        });

        Region spacer = new Region();
        HBox.setHgrow(spacer, Priority.ALWAYS);

        Button refreshBtn = new Button("↻  Actualiser");
        refreshBtn.getStyleClass().add("save-btn");
        refreshBtn.setOnAction(e -> refresh(container));

        bar.getChildren().addAll(periodLabel, periodBox, spacer, refreshBtn);
        return bar;
    }

    // =========================================================================
    // KPI row
    // =========================================================================

    private HBox buildKpiRow(List<TrfHistoryRecord> records, List<TrfMonthRecord> months) {
        double totalEncaissements = records.stream().mapToDouble(TrfHistoryRecord::encaissements).sum();
        double totalCommissions   = records.stream().mapToDouble(TrfHistoryRecord::montantFacturer).sum();
        double totalReverser      = records.stream().mapToDouble(TrfHistoryRecord::sommesReverser).sum();
        double totalNousDoit      = records.stream().mapToDouble(TrfHistoryRecord::nousDoit).sum();
        long   nbClients          = records.stream().map(TrfHistoryRecord::clientName).distinct().count();
        long   nbMonths           = filterYear != null ? 1 : (long) months.size();

        HBox row = new HBox(12,
                kpiCard("Clients actifs",       String.valueOf(nbClients),          "#4CAF50"),
                kpiCard("Encaissements CZ",      formatMoney(totalEncaissements),    "#2196F3"),
                kpiCard("Commissions TTC",        formatMoney(totalCommissions),      "#FF9800"),
                kpiCard("Sommes à reverser",      formatMoney(totalReverser),         "#9C27B0"),
                kpiCard("Nous doit (cumul)",      formatMoney(totalNousDoit),         "#F44336"),
                kpiCard("Mois enregistrés",       String.valueOf(nbMonths),           "#607D8B")
        );
        row.setFillHeight(true);
        return row;
    }

    // =========================================================================
    // Monthly bar chart
    // =========================================================================

    private VBox buildMonthlyChart(List<TrfMonthRecord> months) {
        Label title = new Label("Tendance mensuelle");
        title.setStyle("-fx-font-weight: bold; -fx-font-size: 12px;");

        if (months.isEmpty()) return new VBox(title);

        CategoryAxis xAxis = new CategoryAxis();
        NumberAxis   yAxis = new NumberAxis();
        yAxis.setLabel("€");
        BarChart<String, Number> chart = new BarChart<>(xAxis, yAxis);
        chart.setLegendVisible(true);
        chart.setAnimated(false);
        chart.setPrefHeight(200);
        chart.setStyle("-fx-background-color: transparent;");

        XYChart.Series<String, Number> encSeries = new XYChart.Series<>();
        encSeries.setName("Encaissements");
        XYChart.Series<String, Number> comSeries = new XYChart.Series<>();
        comSeries.setName("Commissions");

        // Show last 12 months max, oldest first
        List<TrfMonthRecord> sorted = new ArrayList<>(months);
        Collections.reverse(sorted);
        int start = Math.max(0, sorted.size() - 12);
        for (int i = start; i < sorted.size(); i++) {
            TrfMonthRecord m = sorted.get(i);
            String label = monthLabel(m.month(), m.year());
            List<TrfHistoryRecord> recs = monthService.getHistoryForMonth(m.id());
            double enc = recs.stream().mapToDouble(TrfHistoryRecord::encaissements).sum();
            double com = recs.stream().mapToDouble(TrfHistoryRecord::montantFacturer).sum();
            encSeries.getData().add(new XYChart.Data<>(label, enc));
            comSeries.getData().add(new XYChart.Data<>(label, com));
        }
        chart.getData().addAll(encSeries, comSeries);

        VBox box = new VBox(6, title, chart);
        box.getStyleClass().add("kpi-card");
        box.setPadding(new Insets(10));
        return box;
    }

    // =========================================================================
    // Etat distribution chart
    // =========================================================================

    private VBox buildEtatChart(List<TrfHistoryRecord> records) {
        Label title = new Label("Répartition par type");
        title.setStyle("-fx-font-weight: bold; -fx-font-size: 12px;");

        long comp     = records.stream().filter(r -> isComp(r.etat())).count();
        long nonComp  = records.stream().filter(r -> r.nonCompensation()).count();
        long partiel  = records.stream().filter(r -> isPartiel(r.etat())).count();
        long debiteur = records.stream().filter(r -> isDebiteur(r)).count();

        VBox bars = new VBox(8);
        long total = records.size();
        if (total > 0) {
            bars.getChildren().addAll(
                    etatBar("VIREMENTS",   comp,     total, "#4CAF50"),
                    etatBar("NON COMP",    nonComp,  total, "#F44336"),
                    etatBar("COMP PART.",  partiel,  total, "#FF9800"),
                    etatBar("DÉBITEURS",   debiteur, total, "#9C27B0")
            );
        }

        VBox box = new VBox(8, title, bars);
        box.getStyleClass().add("kpi-card");
        box.setPadding(new Insets(10));
        return box;
    }

    private HBox etatBar(String label, long count, long total, String color) {
        Label lbl = new Label(String.format("%-15s %d", label, count));
        lbl.setStyle("-fx-font-size: 11px; -fx-font-family: monospace;");
        lbl.setMinWidth(140);

        double pct = total > 0 ? (double) count / total : 0;
        Rectangle fill = new Rectangle(Math.max(2, pct * 180), 14);
        fill.setFill(Color.web(color));
        fill.setArcWidth(4); fill.setArcHeight(4);

        Label pctLbl = new Label(String.format("%.0f%%", pct * 100));
        pctLbl.setStyle("-fx-font-size: 10px; -fx-text-fill: #888;");

        HBox bar = new HBox(8, lbl, fill, pctLbl);
        bar.setAlignment(Pos.CENTER_LEFT);
        return bar;
    }

    // =========================================================================
    // Client table
    // =========================================================================

    private VBox buildClientTable(List<TrfHistoryRecord> records) {
        Label title = new Label("Détail par client");
        title.setStyle("-fx-font-weight: bold; -fx-font-size: 12px;");

        Map<String, double[]> agg     = new LinkedHashMap<>();
        Map<String, String>   etats   = new LinkedHashMap<>();
        Map<String, Boolean>  nonComps = new LinkedHashMap<>();

        for (TrfHistoryRecord r : records) {
            double[] vals = agg.computeIfAbsent(r.clientName(), k -> new double[4]);
            vals[0] += r.encaissements();
            vals[1] += r.montantFacturer();
            vals[2] += r.sommesReverser();
            vals[3] += r.nousDoit();
            etats.put(r.clientName(), r.etat());
            nonComps.put(r.clientName(), r.nonCompensation());
        }

        HBox header = tableRow("CLIENT", "ENCAISSEMENTS", "COMMISSIONS", "À REVERSER", "NOUS DOIT", "TYPE", true);
        VBox rows = new VBox(2, header);

        agg.entrySet().stream()
                .sorted((a, b) -> Double.compare(b.getValue()[0], a.getValue()[0]))
                .forEach(entry -> {
                    String name   = entry.getKey();
                    double[] vals = entry.getValue();
                    String type   = resolveType(etats.get(name), nonComps.getOrDefault(name, false));
                    rows.getChildren().add(tableRow(
                            name,
                            formatMoney(vals[0]),
                            formatMoney(vals[1]),
                            formatMoney(vals[2]),
                            formatMoney(vals[3]),
                            type,
                            false
                    ));
                });

        ScrollPane scroll = new ScrollPane(rows);
        scroll.setFitToWidth(true);
        scroll.setStyle("-fx-background-color: transparent; -fx-border-color: transparent;");
        VBox.setVgrow(scroll, Priority.ALWAYS);

        VBox box = new VBox(8, title, scroll);
        box.getStyleClass().add("kpi-card");
        box.setPadding(new Insets(10));
        VBox.setVgrow(box, Priority.ALWAYS);
        return box;
    }

    private HBox tableRow(String c1, String c2, String c3, String c4, String c5, String c6, boolean header) {
        HBox row = new HBox();
        row.setAlignment(Pos.CENTER_LEFT);
        if (header) {
            row.setStyle("-fx-background-color: #2a2a2a; -fx-padding: 4 8;");
        } else {
            row.setStyle("-fx-padding: 3 8; -fx-border-color: transparent transparent #2a2a2a transparent;");
        }

        String[] vals   = {c1, c2, c3, c4, c5, c6};
        double[] widths = {0.28, 0.15, 0.13, 0.13, 0.13, 0.10};

        for (int i = 0; i < vals.length; i++) {
            Label lbl = new Label(vals[i]);
            lbl.setStyle("-fx-font-size: " + (header ? "10" : "11") + "px;"
                    + (header ? " -fx-font-weight: bold; -fx-text-fill: #aaa;" : "")
                    + (i > 0 ? " -fx-text-alignment: right;" : ""));
            lbl.setMaxWidth(Double.MAX_VALUE);
            HBox.setHgrow(lbl, Priority.ALWAYS);
            lbl.prefWidthProperty().bind(row.widthProperty().multiply(widths[i]));
            row.getChildren().add(lbl);
        }
        return row;
    }

    // =========================================================================
    // Refresh
    // =========================================================================

    public void refresh(VBox container) {
        String consoPath   = AppPreferences.getTrfConso();
        String listingPath = AppPreferences.getTrfListing();
        String tableauPath = AppPreferences.getTrfTableau();

        if (consoPath.isBlank() || listingPath.isBlank() || tableauPath.isBlank()) {
            log.accept("[Dashboard] Fichiers non configurés — ouvrez Configuration.");
            return;
        }
        File consoFile   = new File(consoPath);
        File listingFile = new File(listingPath);
        File tableauFile = new File(tableauPath);
        if (!consoFile.exists() || !listingFile.exists() || !tableauFile.exists()) {
            log.accept("[Dashboard] Fichier(s) introuvable(s) — vérifiez Configuration.");
            return;
        }

        log.accept("[Dashboard] Chargement...");
        LocalDate now = LocalDate.now();
        executor.submit(() -> {
            try {
                monthService.saveOpenMonth(now.getYear(), now.getMonthValue(),
                        consoFile, listingFile, tableauFile);
                Platform.runLater(() -> {
                    log.accept("[Dashboard] ✓ Données actualisées.");
                    load(container);
                });
            } catch (Exception e) {
                Platform.runLater(() -> log.accept("[Dashboard] ERREUR : " + e.getMessage()));
            }
        });
    }

    // =========================================================================
    // Helpers
    // =========================================================================

    private List<TrfHistoryRecord> getFilteredRecords(List<TrfMonthRecord> months) {
        List<TrfHistoryRecord> all = new ArrayList<>();
        for (TrfMonthRecord m : months) {
            if (filterYear != null && (m.year() != filterYear || m.month() != filterMonth)) continue;
            all.addAll(monthService.getHistoryForMonth(m.id()));
        }
        return all;
    }

    private VBox kpiCard(String label, String value, String color) {
        Label lbl = new Label(label);
        lbl.getStyleClass().add("kpi-label");
        Label val = new Label(value);
        val.getStyleClass().add("kpi-value");
        val.setStyle("-fx-text-fill: " + color + ";");
        VBox card = new VBox(4, lbl, val);
        card.getStyleClass().add("kpi-card");
        HBox.setHgrow(card, Priority.ALWAYS);
        return card;
    }

    private String formatMoney(double v) {
        if (v == 0) return "—";
        return String.format("%,.0f €", v).replace(",", " ");
    }

    private String monthLabel(int month, int year) {
        return java.time.Month.of(month)
                .getDisplayName(TextStyle.SHORT, java.util.Locale.FRENCH)
                + ". " + year;
    }

    private String resolveType(String etat, boolean nonComp) {
        if (nonComp) return "NON COMP";
        if (etat == null) return "—";
        String e = etat.toLowerCase();
        if (e.contains("partiel")) return "COMP PART.";
        if (e.contains("debit") || e.contains("débit")) return "DÉBITEUR";
        return "VIREMENT";
    }

    private boolean isComp(String etat) {
        if (etat == null) return false;
        String e = etat.toLowerCase();
        return !e.contains("non") && !e.contains("partiel") && !e.contains("debit");
    }

    private boolean isPartiel(String etat) {
        return etat != null && etat.toLowerCase().contains("partiel");
    }

    private boolean isDebiteur(TrfHistoryRecord r) {
        return r.encaissements() < 0.005 && r.nousDoit() > 0.005;
    }
}
