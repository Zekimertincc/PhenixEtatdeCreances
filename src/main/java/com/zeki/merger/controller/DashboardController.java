package com.zeki.merger.controller;

import com.zeki.merger.AppPreferences;
import com.zeki.merger.db.DatabaseManager;
import com.zeki.merger.service.EtatCreancesSyncService;
import javafx.application.Platform;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.chart.*;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.scene.paint.Color;
import javafx.scene.shape.Rectangle;

import java.io.File;
import java.util.*;
import java.util.concurrent.ExecutorService;
import java.util.function.Consumer;
import java.util.stream.Collectors;

public class DashboardController {

    private final DatabaseManager         db;
    private final EtatCreancesSyncService syncService;
    private final Consumer<String>        log;
    private final ExecutorService         executor;

    private long selectedCompanyId = -1L;

    public DashboardController(DatabaseManager db,
                                EtatCreancesSyncService syncService,
                                Consumer<String> log,
                                ExecutorService executor) {
        this.db          = db;
        this.syncService = syncService;
        this.log         = log;
        this.executor    = executor;
    }

    // =========================================================================
    // Entry point
    // =========================================================================

    public void load(VBox container) {
        container.getChildren().clear();
        container.setPadding(Insets.EMPTY);
        container.setSpacing(0);

        List<Map<String, Object>> summaries = db.getAllCompanySummaries();

        HBox root = new HBox();
        VBox.setVgrow(root, Priority.ALWAYS);
        root.setPrefHeight(Double.MAX_VALUE);

        VBox sidebar = buildSidebar(summaries, root);
        sidebar.setPrefWidth(220);
        sidebar.setMinWidth(180);
        sidebar.setMaxWidth(220);

        VBox detail = new VBox();
        HBox.setHgrow(detail, Priority.ALWAYS);
        detail.setStyle("-fx-background-color: #FAFAF8;");

        if (summaries.isEmpty()) {
            detail.getChildren().add(buildEmptyState());
        } else {
            Map<String, Object> initial = summaries.stream()
                    .filter(s -> toLong(s.get("company_id")) == selectedCompanyId)
                    .findFirst()
                    .orElse(summaries.get(0));
            buildDetail(detail, initial, summaries);
            selectedCompanyId = toLong(initial.get("company_id"));
        }

        root.getChildren().addAll(sidebar, detail);
        container.getChildren().add(root);
    }

    // =========================================================================
    // Sidebar
    // =========================================================================

    private VBox buildSidebar(List<Map<String, Object>> summaries, HBox root) {
        VBox sidebar = new VBox();
        sidebar.setStyle("-fx-background-color: #F2F0EB; -fx-border-color: rgba(0,0,0,0.10); -fx-border-width: 0 1 0 0;");

        Label header = new Label("SOCIÉTÉS");
        header.setStyle("-fx-font-size: 10px; -fx-font-weight: bold; -fx-text-fill: #6B6B6B; -fx-padding: 14 12 8 14; -fx-background-color: #F2F0EB;");
        header.setMaxWidth(Double.MAX_VALUE);

        TextField search = new TextField();
        search.setPromptText("Rechercher…");
        search.getStyleClass().add("path-field");
        search.setStyle("-fx-font-size: 12px;");
        VBox.setMargin(search, new Insets(0, 10, 8, 10));

        VBox list = new VBox();
        VBox.setVgrow(list, Priority.ALWAYS);
        ScrollPane scroll = new ScrollPane(list);
        scroll.setFitToWidth(true);
        scroll.setStyle("-fx-background-color: transparent; -fx-background: transparent; -fx-border-color: transparent;");
        VBox.setVgrow(scroll, Priority.ALWAYS);

        Button syncBtn = new Button("↻  Sync sociétés");
        syncBtn.setMaxWidth(Double.MAX_VALUE);
        syncBtn.getStyleClass().add("secondary-btn");
        syncBtn.setStyle("-fx-font-size: 11px;");
        VBox.setMargin(syncBtn, new Insets(8, 10, 10, 10));
        syncBtn.setOnAction(e -> syncAll(root, syncBtn));

        sidebar.getChildren().addAll(header, search, scroll, syncBtn);

        populateCompanyList(list, summaries, root);

        search.textProperty().addListener((obs, old, val) -> {
            String q = val.trim().toLowerCase();
            List<Map<String, Object>> filtered = summaries.stream()
                    .filter(s -> q.isEmpty() || str(s, "name").toLowerCase().contains(q))
                    .collect(Collectors.toList());
            populateCompanyList(list, filtered, root);
        });

        return sidebar;
    }

    private void populateCompanyList(VBox list, List<Map<String, Object>> summaries, HBox root) {
        list.getChildren().clear();
        for (Map<String, Object> s : summaries) {
            long    id      = toLong(s.get("company_id"));
            String  name    = str(s, "name");
            int     nb      = toInt(s.get("nb_dossiers"));
            String  code    = str(s, "code_client");
            boolean synced  = nb > 0;

            Button btn = new Button();
            btn.setMaxWidth(Double.MAX_VALUE);
            btn.setAlignment(Pos.CENTER_LEFT);
            btn.getStyleClass().add("company-list-item");

            Label nameLbl = new Label(name);
            nameLbl.setStyle("-fx-font-size: 12px; -fx-font-weight: bold; -fx-text-fill: #1a1a1a; -fx-wrap-text: true;");
            nameLbl.setMaxWidth(180);
            nameLbl.setWrapText(true);

            String metaText = synced ? nb + " dossiers" + (code.isEmpty() ? "" : " · " + code) : "—  pas synchronisé";
            Label metaLbl = new Label(metaText);
            metaLbl.setStyle("-fx-font-size: 10px; -fx-text-fill: " + (synced ? "#6B6B6B" : "#9B9B9B") + ";");

            VBox content = new VBox(2, nameLbl, metaLbl);
            btn.setGraphic(content);
            btn.setPadding(new Insets(8, 12, 8, 12));

            if (id == selectedCompanyId) btn.getStyleClass().add("company-list-selected");

            btn.setOnAction(e -> {
                selectedCompanyId = id;
                VBox detail = (VBox) root.getChildren().get(1);
                buildDetail(detail, s, summaries);
                populateCompanyList(list, summaries, root);
            });

            list.getChildren().add(btn);
        }
    }

    // =========================================================================
    // Detail pane
    // =========================================================================

    private void buildDetail(VBox detail, Map<String, Object> s, List<Map<String, Object>> all) {
        detail.getChildren().clear();
        detail.setSpacing(0);

        boolean synced = toInt(s.get("nb_dossiers")) > 0;

        detail.getChildren().add(buildDetailHeader(s, synced));

        if (!synced) {
            detail.getChildren().add(buildNotSyncedPane(str(s, "name")));
            return;
        }

        detail.getChildren().add(buildKpiRow(s, all));

        HBox charts = buildChartsRow(s, all);
        VBox.setVgrow(charts, Priority.ALWAYS);
        detail.getChildren().add(charts);
    }

    private HBox buildDetailHeader(Map<String, Object> s, boolean synced) {
        HBox header = new HBox();
        header.setAlignment(Pos.CENTER_LEFT);
        header.setPadding(new Insets(14, 20, 12, 20));
        header.setStyle("-fx-background-color: #FAFAF8; -fx-border-color: rgba(0,0,0,0.10); -fx-border-width: 0 0 1 0;");

        VBox nameBox = new VBox(2);
        HBox.setHgrow(nameBox, Priority.ALWAYS);

        Label nameLbl = new Label(str(s, "name"));
        nameLbl.getStyleClass().add("detail-company-name");

        String code = str(s, "code_client");
        String resp = str(s, "responsable");
        String sync = str(s, "last_sync");
        String meta = synced
                ? (code.isEmpty() ? "" : "Code : " + code + "  ·  ")
                  + (resp.isEmpty() ? "" : resp + "  ·  ")
                  + (sync.isEmpty() ? "" : "sync " + sync.substring(0, Math.min(10, sync.length())))
                : "Pas encore synchronisé";
        Label metaLbl = new Label(meta);
        metaLbl.getStyleClass().add("detail-company-code");

        nameBox.getChildren().addAll(nameLbl, metaLbl);
        header.getChildren().add(nameBox);
        return header;
    }

    // =========================================================================
    // KPI row
    // =========================================================================

    private HBox buildKpiRow(Map<String, Object> s, List<Map<String, Object>> all) {
        double creance  = toDouble(s.get("creance_principale"));
        double recouvre = toDouble(s.get("recouvre_total"));
        double pct      = creance > 0 ? recouvre / creance * 100.0 : 0.0;
        int    nbActive = toInt(s.get("nb_dossiers")) - toInt(s.get("nb_soldes"));

        double totalCreance  = all.stream().mapToDouble(m -> toDouble(m.get("creance_principale"))).sum();
        double totalRecouvre = all.stream().mapToDouble(m -> toDouble(m.get("recouvre_total"))).sum();
        double globalPct     = totalCreance > 0 ? totalRecouvre / totalCreance * 100.0 : 0.0;

        HBox row = new HBox(10);
        row.setPadding(new Insets(14, 20, 14, 20));
        row.setStyle("-fx-background-color: #FAFAF8; -fx-border-color: rgba(0,0,0,0.10); -fx-border-width: 0 0 1 0;");

        row.getChildren().addAll(
                kpiCard("Créance principale",  fmt(creance) + " €",              "#1a1a1a"),
                kpiCard("Recouvré total",       fmt(recouvre) + " €",             "#0F6E56"),
                kpiCard("Taux de recouvrement", String.format("%.1f%%", pct),
                        pct >= 50 ? "#0F6E56" : pct >= 25 ? "#BA7517" : "#A32D2D"),
                kpiCard("Dossiers actifs",      String.valueOf(nbActive),          "#185FA5"),
                kpiCard("Taux global (toutes)", String.format("%.1f%%", globalPct), "#6B6B6B")
        );
        return row;
    }

    private VBox kpiCard(String label, String value, String color) {
        Label lbl = new Label(label);
        lbl.getStyleClass().add("kpi-label");
        Label val = new Label(value);
        val.getStyleClass().add("kpi-value");
        val.setStyle("-fx-text-fill: " + color + "; -fx-font-size: 16px;");
        VBox card = new VBox(4, lbl, val);
        card.getStyleClass().add("kpi-card");
        HBox.setHgrow(card, Priority.ALWAYS);
        return card;
    }

    // =========================================================================
    // Charts row
    // =========================================================================

    private HBox buildChartsRow(Map<String, Object> s, List<Map<String, Object>> all) {
        HBox row = new HBox(12);
        row.setPadding(new Insets(16, 20, 16, 20));
        row.setStyle("-fx-background-color: #FAFAF8;");

        VBox leftCard = buildEtatChart(s);
        leftCard.setPrefWidth(260);
        leftCard.setMinWidth(220);

        VBox rightCard = buildComparaisonChart(all);
        HBox.setHgrow(rightCard, Priority.ALWAYS);

        row.getChildren().addAll(leftCard, rightCard);
        return row;
    }

    private VBox buildEtatChart(Map<String, Object> s) {
        Label title = new Label("Répartition des dossiers");
        title.setStyle("-fx-font-size: 11px; -fx-font-weight: bold; -fx-text-fill: #6B6B6B;");

        int soldes  = toInt(s.get("nb_soldes"));
        int gestion = toInt(s.get("nb_gestion"));
        int irr     = toInt(s.get("nb_irr"));
        int arj     = toInt(s.get("nb_arj"));
        int autres  = toInt(s.get("nb_autres"));

        PieChart pie = new PieChart();
        pie.setAnimated(false);
        pie.setLegendVisible(false);
        pie.setLabelsVisible(false);
        pie.setPrefHeight(180);
        pie.setStyle("-fx-background-color: transparent;");

        if (soldes  > 0) addSlice(pie, "Soldé",   soldes,  "#1D9E75");
        if (gestion > 0) addSlice(pie, "Gestion",  gestion, "#BA7517");
        if (irr     > 0) addSlice(pie, "IRR",      irr,     "#E24B4A");
        if (arj     > 0) addSlice(pie, "ARJ",      arj,     "#378ADD");
        if (autres  > 0) addSlice(pie, "Autres",   autres,  "#888780");

        VBox legend = new VBox(4);
        legend.setPadding(new Insets(8, 0, 0, 0));
        addLegendRow(legend, "Soldé",   soldes,  "#1D9E75");
        addLegendRow(legend, "Gestion", gestion, "#BA7517");
        addLegendRow(legend, "IRR",     irr,     "#E24B4A");
        addLegendRow(legend, "ARJ",     arj,     "#378ADD");
        addLegendRow(legend, "Autres",  autres,  "#888780");

        VBox card = new VBox(8, title, pie, legend);
        card.getStyleClass().add("kpi-card");
        card.setPadding(new Insets(12));
        return card;
    }

    private void addSlice(PieChart pie, String name, int count, String hex) {
        PieChart.Data slice = new PieChart.Data(name, count);
        pie.getData().add(slice);
        slice.getNode().setStyle("-fx-pie-color: " + hex + ";");
    }

    private void addLegendRow(VBox legend, String label, int count, String hex) {
        Rectangle dot = new Rectangle(10, 10);
        dot.setFill(Color.web(hex));
        dot.setArcWidth(3); dot.setArcHeight(3);
        Label lbl = new Label(label + " — " + count);
        lbl.setStyle("-fx-font-size: 11px; -fx-text-fill: #6B6B6B;");
        HBox row = new HBox(6, dot, lbl);
        row.setAlignment(Pos.CENTER_LEFT);
        legend.getChildren().add(row);
    }

    private VBox buildComparaisonChart(List<Map<String, Object>> all) {
        Label title = new Label("Créance vs recouvré — toutes les sociétés");
        title.setStyle("-fx-font-size: 11px; -fx-font-weight: bold; -fx-text-fill: #6B6B6B;");

        List<Map<String, Object>> synced = all.stream()
                .filter(m -> toInt(m.get("nb_dossiers")) > 0)
                .sorted(Comparator.comparingDouble(m -> -toDouble(((Map<?,?>) m).get("creance_principale"))))
                .limit(8)
                .collect(Collectors.toList());

        CategoryAxis yAxis = new CategoryAxis();
        NumberAxis   xAxis = new NumberAxis();
        xAxis.setLabel("€");
        xAxis.setTickLabelFormatter(new NumberAxis.DefaultFormatter(xAxis) {
            @Override public String toString(Number v) {
                double d = v.doubleValue();
                if (d >= 1_000_000) return String.format("%.1fM", d / 1_000_000);
                if (d >= 1_000)     return String.format("%.0fk", d / 1_000);
                return String.valueOf((int) d);
            }
        });

        BarChart<Number, String> chart = new BarChart<>(xAxis, yAxis);
        chart.setAnimated(false);
        chart.setLegendVisible(false);
        chart.setPrefHeight(220);
        chart.setBarGap(2);
        chart.setCategoryGap(8);
        chart.setStyle("-fx-background-color: transparent;");

        XYChart.Series<Number, String> creanceSeries = new XYChart.Series<>();
        creanceSeries.setName("Créance");
        XYChart.Series<Number, String> recouvSeries  = new XYChart.Series<>();
        recouvSeries.setName("Recouvré");

        for (Map<String, Object> m : synced) {
            String name = abbreviate(str(m, "name"), 18);
            creanceSeries.getData().add(new XYChart.Data<>(toDouble(m.get("creance_principale")), name));
            recouvSeries.getData().add(new XYChart.Data<>(toDouble(m.get("recouvre_total")), name));
        }
        chart.getData().addAll(creanceSeries, recouvSeries);

        Platform.runLater(() -> {
            for (XYChart.Data<Number, String> d : creanceSeries.getData()) {
                if (d.getNode() != null) d.getNode().setStyle("-fx-bar-fill: #B5D4F4;");
            }
            for (XYChart.Data<Number, String> d : recouvSeries.getData()) {
                if (d.getNode() != null) d.getNode().setStyle("-fx-bar-fill: #1D9E75;");
            }
        });

        HBox legend = new HBox(12);
        legend.setPadding(new Insets(4, 0, 0, 0));
        legend.getChildren().addAll(
                legendChip("Créance",  "#B5D4F4"),
                legendChip("Recouvré", "#1D9E75")
        );

        VBox card = new VBox(8, title, chart, legend);
        card.getStyleClass().add("kpi-card");
        card.setPadding(new Insets(12));
        VBox.setVgrow(card, Priority.ALWAYS);
        return card;
    }

    private HBox legendChip(String label, String hex) {
        Rectangle dot = new Rectangle(10, 10);
        dot.setFill(Color.web(hex));
        dot.setArcWidth(3); dot.setArcHeight(3);
        Label lbl = new Label(label);
        lbl.setStyle("-fx-font-size: 11px; -fx-text-fill: #6B6B6B;");
        HBox box = new HBox(5, dot, lbl);
        box.setAlignment(Pos.CENTER_LEFT);
        return box;
    }

    // =========================================================================
    // Empty / not-synced states
    // =========================================================================

    private VBox buildEmptyState() {
        Label lbl = new Label("Aucune société synchronisée.\nCliquez sur « Sync sociétés » pour charger les données.");
        lbl.setStyle("-fx-text-fill: #9B9B9B; -fx-font-size: 13px; -fx-text-alignment: center;");
        lbl.setWrapText(true);
        VBox box = new VBox(lbl);
        box.setAlignment(Pos.CENTER);
        box.setPadding(new Insets(40));
        VBox.setVgrow(box, Priority.ALWAYS);
        return box;
    }

    private VBox buildNotSyncedPane(String name) {
        Label lbl = new Label(name + " n'a pas encore été synchronisé.\nLancez « Sync sociétés » pour charger ses données.");
        lbl.setStyle("-fx-text-fill: #9B9B9B; -fx-font-size: 13px; -fx-text-alignment: center;");
        lbl.setWrapText(true);
        VBox box = new VBox(lbl);
        box.setAlignment(Pos.CENTER);
        box.setPadding(new Insets(40));
        VBox.setVgrow(box, Priority.ALWAYS);
        return box;
    }

    // =========================================================================
    // Sync action
    // =========================================================================

    private void syncAll(HBox root, Button syncBtn) {
        String rootPath = AppPreferences.getMergeRoot();
        if (rootPath.isBlank()) {
            log.accept("[Dashboard] Dossier racine non configuré.");
            return;
        }
        File rootFolder = new File(rootPath);
        if (!rootFolder.isDirectory()) {
            log.accept("[Dashboard] Dossier introuvable : " + rootPath);
            return;
        }
        syncBtn.setDisable(true);
        syncBtn.setText("↻  Sync en cours…");
        log.accept("[Dashboard] Synchronisation…");
        executor.submit(() -> {
            try {
                syncService.syncAll(rootFolder, (pct, msg) ->
                        Platform.runLater(() -> log.accept(msg)));
                Platform.runLater(() -> {
                    syncBtn.setDisable(false);
                    syncBtn.setText("↻  Sync sociétés");
                    log.accept("[Dashboard] ✓ Synchronisation terminée.");
                    VBox container = (VBox) root.getParent();
                    if (container != null) load(container);
                });
            } catch (Exception e) {
                Platform.runLater(() -> {
                    syncBtn.setDisable(false);
                    syncBtn.setText("↻  Sync sociétés");
                    log.accept("[Dashboard] ERREUR : " + e.getMessage());
                });
            }
        });
    }

    // =========================================================================
    // Helpers
    // =========================================================================

    private String fmt(double v) {
        if (v == 0) return "—";
        return String.format("%,.0f", v).replace(",", " ");
    }

    private String str(Map<String, Object> m, String key) {
        Object v = m.get(key);
        return v == null ? "" : v.toString().trim();
    }

    private long toLong(Object v) {
        if (v == null) return -1L;
        if (v instanceof Number n) return n.longValue();
        try { return Long.parseLong(v.toString()); } catch (Exception e) { return -1L; }
    }

    private int toInt(Object v) {
        if (v == null) return 0;
        if (v instanceof Number n) return n.intValue();
        try { return Integer.parseInt(v.toString()); } catch (Exception e) { return 0; }
    }

    private double toDouble(Object v) {
        if (v == null) return 0.0;
        if (v instanceof Number n) return n.doubleValue();
        try { return Double.parseDouble(v.toString()); } catch (Exception e) { return 0.0; }
    }

    private String abbreviate(String s, int max) {
        return s.length() <= max ? s : s.substring(0, max - 1) + "…";
    }
}
