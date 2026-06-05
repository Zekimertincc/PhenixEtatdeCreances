package com.zeki.merger.controller;

import com.zeki.merger.AppPreferences;
import com.zeki.merger.db.DatabaseManager;
import com.zeki.merger.service.EtatCreancesSyncService;
import javafx.application.Platform;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.scene.paint.Color;
import javafx.scene.shape.Rectangle;

import java.io.File;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.ExecutorService;
import java.util.function.Consumer;
import java.util.stream.Collectors;

public class DashboardController {

    private final DatabaseManager         db;
    private final EtatCreancesSyncService syncService;
    private final Consumer<String>        log;
    private final ExecutorService         executor;

    private long      selectedCompanyId = -1L;
    private LocalDate filterFrom        = LocalDate.now().minusYears(1).withDayOfMonth(1);
    private LocalDate filterTo          = LocalDate.now();

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

        // Tab bar
        HBox tabBar = new HBox();
        tabBar.setStyle("-fx-background-color: #FAFAF8; -fx-border-color: rgba(0,0,0,0.10); -fx-border-width: 0 0 1 0;");
        tabBar.setPadding(new Insets(0, 20, 0, 20));

        Button tabSociete = tabBtn("Par société");
        Button tabGlobal  = tabBtn("Vue globale");
        tabBar.getChildren().addAll(tabSociete, tabGlobal);

        // Content area
        VBox content = new VBox();
        VBox.setVgrow(content, Priority.ALWAYS);
        content.setStyle("-fx-background-color: #FAFAF8;");

        final boolean[] isGlobal = {false};
        loadParSociete(content);
        styleActiveTab(tabSociete, tabGlobal, true);

        tabSociete.setOnAction(e -> {
            if (isGlobal[0]) {
                isGlobal[0] = false;
                loadParSociete(content);
                styleActiveTab(tabSociete, tabGlobal, true);
            }
        });
        tabGlobal.setOnAction(e -> {
            if (!isGlobal[0]) {
                isGlobal[0] = true;
                loadVueGlobale(content);
                styleActiveTab(tabSociete, tabGlobal, false);
            }
        });

        container.getChildren().addAll(tabBar, content);
    }

    private Button tabBtn(String label) {
        Button btn = new Button(label);
        btn.setStyle("-fx-background-color: transparent; -fx-border-color: transparent; " +
                     "-fx-font-size: 12px; -fx-padding: 10 4 10 4; -fx-cursor: hand;");
        return btn;
    }

    private void styleActiveTab(Button tabSociete, Button tabGlobal, boolean societeActive) {
        tabSociete.setStyle("-fx-background-color: transparent; -fx-font-size: 12px; -fx-padding: 10 4 10 4; -fx-cursor: hand; " +
            "-fx-border-color: " + (societeActive ? "#185FA5 transparent transparent transparent" : "transparent") + "; -fx-border-width: 0 0 2 0; " +
            "-fx-text-fill: " + (societeActive ? "#185FA5" : "#6B6B6B") + ";");
        tabGlobal.setStyle("-fx-background-color: transparent; -fx-font-size: 12px; -fx-padding: 10 4 10 4; -fx-cursor: hand; " +
            "-fx-border-color: " + (!societeActive ? "#185FA5 transparent transparent transparent" : "transparent") + "; -fx-border-width: 0 0 2 0; " +
            "-fx-text-fill: " + (!societeActive ? "#185FA5" : "#6B6B6B") + ";");
        HBox.setMargin(tabGlobal, new Insets(0, 0, 0, 20));
    }

    // =========================================================================
    // Tab: Par société
    // =========================================================================

    private void loadParSociete(VBox content) {
        content.getChildren().clear();

        List<Map<String, Object>> summaries = db.getAllCompanySummaries().stream()
                .filter(s -> toInt(s.get("nb_dossiers")) > 0)
                .collect(Collectors.toList());

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
            selectedCompanyId = toLong(initial.get("company_id"));
            buildDetail(detail, initial, summaries);
        }

        root.getChildren().addAll(sidebar, detail);
        content.getChildren().add(root);
    }

    // =========================================================================
    // Sidebar
    // =========================================================================

    private VBox buildSidebar(List<Map<String, Object>> allSummaries, HBox root) {
        VBox sidebar = new VBox();
        sidebar.setStyle("-fx-background-color: #F2F0EB; -fx-border-color: rgba(0,0,0,0.10); -fx-border-width: 0 1 0 0;");

        Label header = new Label("SOCIÉTÉS");
        header.setMaxWidth(Double.MAX_VALUE);
        header.setStyle("-fx-font-size: 10px; -fx-font-weight: bold; -fx-text-fill: #6B6B6B; -fx-padding: 14 12 8 14;");

        TextField search = new TextField();
        search.setPromptText("Rechercher…");
        search.getStyleClass().add("path-field");
        VBox.setMargin(search, new Insets(0, 10, 8, 10));

        VBox listBox = new VBox();
        ScrollPane scroll = new ScrollPane(listBox);
        scroll.setFitToWidth(true);
        scroll.setStyle("-fx-background-color: transparent; -fx-background: transparent; -fx-border-color: transparent;");
        VBox.setVgrow(scroll, Priority.ALWAYS);

        Button syncBtn = new Button("↻  Sync sociétés");
        syncBtn.setMaxWidth(Double.MAX_VALUE);
        syncBtn.getStyleClass().add("secondary-btn");
        VBox.setMargin(syncBtn, new Insets(8, 10, 10, 10));
        syncBtn.setOnAction(e -> doSync(root, syncBtn));

        sidebar.getChildren().addAll(header, search, scroll, syncBtn);

        rebuildList(listBox, allSummaries, root);

        search.textProperty().addListener((obs, old, val) -> {
            String q = val.trim().toLowerCase();
            List<Map<String, Object>> filtered = allSummaries.stream()
                    .filter(s -> q.isEmpty() || str(s, "name").toLowerCase().contains(q))
                    .collect(Collectors.toList());
            rebuildList(listBox, filtered, root);
        });

        return sidebar;
    }

    private void rebuildList(VBox listBox, List<Map<String, Object>> items, HBox root) {
        listBox.getChildren().clear();
        for (Map<String, Object> s : items) {
            long    id     = toLong(s.get("company_id"));
            String  name   = str(s, "name");
            int     nb     = toInt(s.get("nb_dossiers"));
            String  code   = str(s, "code_client");
            boolean synced = nb > 0;

            VBox cell = new VBox(2);
            cell.setPadding(new Insets(8, 12, 8, 12));
            cell.setMaxWidth(Double.MAX_VALUE);
            cell.setCursor(javafx.scene.Cursor.HAND);

            if (id == selectedCompanyId) {
                cell.setStyle("-fx-background-color: #E6F1FB; -fx-border-color: #185FA5; -fx-border-width: 0 0 0 3;");
            } else {
                cell.setStyle("-fx-background-color: transparent; -fx-border-color: rgba(0,0,0,0.07); -fx-border-width: 0 0 1 0;");
                cell.setOnMouseEntered(e -> cell.setStyle("-fx-background-color: #EAEAE6; -fx-border-color: rgba(0,0,0,0.07); -fx-border-width: 0 0 1 0;"));
                cell.setOnMouseExited(e  -> cell.setStyle("-fx-background-color: transparent; -fx-border-color: rgba(0,0,0,0.07); -fx-border-width: 0 0 1 0;"));
            }

            Label nameLbl = new Label(name);
            nameLbl.setWrapText(true);
            nameLbl.setMaxWidth(190);
            nameLbl.setStyle("-fx-font-size: 12px; -fx-font-weight: bold; -fx-text-fill: " + (id == selectedCompanyId ? "#0C447C" : "#1a1a1a") + ";");

            String meta = synced ? nb + " dossiers" + (code.isEmpty() ? "" : "  ·  " + code) : "pas synchronisé";
            Label metaLbl = new Label(meta);
            metaLbl.setStyle("-fx-font-size: 10px; -fx-text-fill: " + (synced ? "#6B6B6B" : "#9B9B9B") + ";");

            cell.getChildren().addAll(nameLbl, metaLbl);
            cell.setOnMouseClicked(e -> {
                selectedCompanyId = id;
                VBox detail  = (VBox) root.getChildren().get(1);
                VBox sidebar = (VBox) root.getChildren().get(0);
                ScrollPane sp = (ScrollPane) sidebar.getChildren().get(2);
                rebuildList((VBox) sp.getContent(), items, root);
                buildDetail(detail, s, (List<Map<String, Object>>) cell.getUserData());
            });
            cell.setUserData(items);

            listBox.getChildren().add(cell);
        }
    }

    // =========================================================================
    // Detail pane
    // =========================================================================

    private void buildDetail(VBox detail, Map<String, Object> s, List<Map<String, Object>> all) {
        detail.getChildren().clear();
        detail.setSpacing(0);

        boolean synced = toInt(s.get("nb_dossiers")) > 0;

        HBox hdr = new HBox();
        hdr.setAlignment(Pos.CENTER_LEFT);
        hdr.setPadding(new Insets(14, 20, 12, 20));
        hdr.setStyle("-fx-border-color: rgba(0,0,0,0.10); -fx-border-width: 0 0 1 0;");
        VBox nameBox = new VBox(3);
        HBox.setHgrow(nameBox, Priority.ALWAYS);
        Label nameLbl = new Label(str(s, "name"));
        nameLbl.setStyle("-fx-font-size: 14px; -fx-font-weight: bold; -fx-text-fill: #1a1a1a;");
        String code = str(s, "code_client"), resp = str(s, "responsable"), sync = str(s, "last_sync");
        String metaStr = synced
                ? List.of(code.isEmpty() ? "" : "Code : " + code,
                        resp.isEmpty() ? "" : resp,
                        sync.length() >= 10 ? "sync " + sync.substring(0, 10) : "")
                .stream().filter(x -> !x.isEmpty()).collect(Collectors.joining("  ·  "))
                : "Pas encore synchronisé — lancez Sync sociétés";
        Label metaLbl = new Label(metaStr);
        metaLbl.setStyle("-fx-font-size: 11px; -fx-text-fill: #9B9B9B; -fx-font-family: monospace;");
        nameBox.getChildren().addAll(nameLbl, metaLbl);
        hdr.getChildren().add(nameBox);
        detail.getChildren().add(hdr);

        if (!synced) {
            Label hint = new Label("Aucune donnée pour cette société.\nLancez « Sync sociétés » pour charger ses données.");
            hint.setStyle("-fx-text-fill: #9B9B9B; -fx-font-size: 13px; -fx-text-alignment: center;");
            hint.setWrapText(true);
            VBox box = new VBox(hint);
            box.setAlignment(Pos.CENTER);
            VBox.setVgrow(box, Priority.ALWAYS);
            detail.getChildren().add(box);
            return;
        }

        detail.getChildren().add(buildKpiRow(s, all));

        HBox charts = new HBox(12);
        charts.setPadding(new Insets(16, 20, 16, 20));
        VBox.setVgrow(charts, Priority.ALWAYS);

        VBox left = buildEtatBars(s);
        left.setPrefWidth(250);
        left.setMinWidth(220);

        VBox right = buildComparaisonBars(all);
        HBox.setHgrow(right, Priority.ALWAYS);

        charts.getChildren().addAll(left, right);
        detail.getChildren().add(charts);
    }

    // =========================================================================
    // KPI row
    // =========================================================================

    private HBox buildKpiRow(Map<String, Object> s, List<Map<String, Object>> all) {
        double creance  = toDouble(s.get("creance_principale"));
        double recouvre = toDouble(s.get("recouvre_total"));
        int    actifs   = Math.max(0, toInt(s.get("nb_dossiers")) - toInt(s.get("nb_soldes")));

        HBox row = new HBox(10);
        row.setPadding(new Insets(14, 20, 14, 20));
        row.setStyle("-fx-border-color: rgba(0,0,0,0.10); -fx-border-width: 0 0 1 0;");
        row.getChildren().addAll(
                kpi("Créance principale",  fmt(creance) + " €",            "#1a1a1a"),
                kpi("Recouvré total",      fmt(recouvre) + " €",           "#0F6E56"),
                kpi("Dossiers actifs",     String.valueOf(Math.max(0, actifs)), "#185FA5"),
                kpi("Dossiers soldés",     String.valueOf(toInt(s.get("nb_soldes"))),           "#0F6E56"),
                kpi("Commissions",         fmt(toDouble(s.get("commissions"))) + " €",          "#185FA5")
        );
        return row;
    }

    private VBox kpi(String label, String value, String color) {
        Label lbl = new Label(label);
        lbl.setStyle("-fx-font-size: 10px; -fx-text-fill: #9B9B9B;");
        Label val = new Label(value);
        val.setStyle("-fx-font-size: 16px; -fx-font-weight: bold; -fx-text-fill: " + color + ";");
        VBox card = new VBox(4, lbl, val);
        card.setStyle("-fx-background-color: #F2F0EB; -fx-background-radius: 8px; -fx-padding: 10 12 10 12;");
        HBox.setHgrow(card, Priority.ALWAYS);
        return card;
    }

    // =========================================================================
    // État distribution bars
    // =========================================================================

    private VBox buildEtatBars(Map<String, Object> s) {
        int total  = toInt(s.get("nb_dossiers"));
        int soldes = toInt(s.get("nb_soldes"));
        int cours  = Math.max(0, total - soldes);

        Label title = new Label("Répartition des dossiers");
        title.setStyle("-fx-font-size: 11px; -fx-font-weight: bold; -fx-text-fill: #6B6B6B;");

        VBox bars = new VBox(8);
        if (total > 0) {
            bars.getChildren().addAll(
                etatBar("Soldés",   soldes, total, "#1D9E75"),
                etatBar("En cours", cours,  total, "#BA7517")
            );
        }

        VBox card = new VBox(10, title, bars);
        card.setStyle("-fx-background-color: #F2F0EB; -fx-background-radius: 8px; -fx-padding: 12;");
        VBox.setVgrow(card, Priority.ALWAYS);
        return card;
    }

    private HBox etatBar(String label, int count, int total, String hex) {
        double pct = total > 0 ? (double) count / total : 0;

        Label nameLbl = new Label(label);
        nameLbl.setStyle("-fx-font-size: 11px; -fx-text-fill: #1a1a1a;");
        nameLbl.setMinWidth(55);

        Rectangle fill = new Rectangle(Math.max(3, pct * 130), 12);
        fill.setFill(Color.web(hex));
        fill.setArcWidth(4); fill.setArcHeight(4);

        Label cntLbl = new Label(count + "  " + String.format("%.0f%%", pct * 100));
        cntLbl.setStyle("-fx-font-size: 10px; -fx-text-fill: #9B9B9B;");

        HBox row = new HBox(8, nameLbl, fill, cntLbl);
        row.setAlignment(Pos.CENTER_LEFT);
        return row;
    }

    // =========================================================================
    // Comparaison bars
    // =========================================================================

    private VBox buildComparaisonBars(List<Map<String, Object>> all) {
        Label title = new Label("Créance vs recouvré — toutes les sociétés");
        title.setStyle("-fx-font-size: 11px; -fx-font-weight: bold; -fx-text-fill: #6B6B6B;");

        List<Map<String, Object>> synced = all.stream()
                .filter(m -> toInt(m.get("nb_dossiers")) > 0)
                .sorted(Comparator.comparingDouble(m -> -toDouble(((Map<?,?>) m).get("creance_principale"))))
                .limit(8)
                .collect(Collectors.toList());

        double maxVal = synced.stream()
                .mapToDouble(m -> toDouble(m.get("creance_principale")))
                .max().orElse(1.0);

        VBox rows = new VBox(10);
        for (Map<String, Object> m : synced) {
            String name     = abbreviate(str(m, "name"), 20);
            double creance  = toDouble(m.get("creance_principale"));
            double recouvre = toDouble(m.get("recouvre_total"));

            Label nameLbl = new Label(name);
            nameLbl.setStyle("-fx-font-size: 11px; -fx-text-fill: #1a1a1a;");
            nameLbl.setMinWidth(150);

            double barW = 200.0;
            double cW   = maxVal > 0 ? creance  / maxVal * barW : 0;
            double rW   = maxVal > 0 ? recouvre / maxVal * barW : 0;

            Rectangle cBar = new Rectangle(Math.max(2, cW), 9);
            cBar.setFill(Color.web("#B5D4F4"));
            cBar.setArcWidth(3); cBar.setArcHeight(3);

            Rectangle rBar = new Rectangle(Math.max(2, rW), 9);
            rBar.setFill(Color.web("#1D9E75"));
            rBar.setArcWidth(3); rBar.setArcHeight(3);

            Label valLbl = new Label(fmtK(creance));
            valLbl.setStyle("-fx-font-size: 10px; -fx-text-fill: #9B9B9B;");

            VBox barStack = new VBox(3, cBar, rBar);
            HBox row = new HBox(10, nameLbl, barStack, valLbl);
            row.setAlignment(Pos.CENTER_LEFT);
            rows.getChildren().add(row);
        }

        HBox legend = new HBox(12,
                legendChip("Créance",  "#B5D4F4"),
                legendChip("Recouvré", "#1D9E75"));
        legend.setPadding(new Insets(6, 0, 0, 0));

        VBox card = new VBox(10, title, rows, legend);
        card.setStyle("-fx-background-color: #F2F0EB; -fx-background-radius: 8px; -fx-padding: 12;");
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
    // Tab: Vue globale
    // =========================================================================

    private void loadVueGlobale(VBox content) {
        content.getChildren().clear();

        HBox filterRow = new HBox(12);
        filterRow.setAlignment(Pos.CENTER_LEFT);
        filterRow.setPadding(new Insets(14, 20, 14, 20));
        filterRow.setStyle("-fx-border-color: rgba(0,0,0,0.10); -fx-border-width: 0 0 1 0;");

        Label fromLbl = new Label("Du :");
        fromLbl.setStyle("-fx-font-size: 12px; -fx-text-fill: #6B6B6B;");
        DatePicker fromPicker = new DatePicker(filterFrom);
        fromPicker.setPrefWidth(140);

        Label toLbl = new Label("Au :");
        toLbl.setStyle("-fx-font-size: 12px; -fx-text-fill: #6B6B6B;");
        DatePicker toPicker = new DatePicker(filterTo);
        toPicker.setPrefWidth(140);

        Button applyBtn = new Button("Appliquer");
        applyBtn.getStyleClass().add("secondary-btn");

        filterRow.getChildren().addAll(fromLbl, fromPicker, toLbl, toPicker, applyBtn);

        VBox resultsArea = new VBox();
        VBox.setVgrow(resultsArea, Priority.ALWAYS);

        applyBtn.setOnAction(e -> {
            filterFrom = fromPicker.getValue();
            filterTo   = toPicker.getValue();
            buildGlobalResults(resultsArea);
        });

        buildGlobalResults(resultsArea);
        content.getChildren().addAll(filterRow, resultsArea);
    }

    private void buildGlobalResults(VBox area) {
        area.getChildren().clear();
        area.setPadding(new Insets(16, 20, 16, 20));
        area.setSpacing(12);

        String fmt  = "yyyy-MM-dd";
        String from = filterFrom.format(DateTimeFormatter.ofPattern(fmt));
        String to   = filterTo.format(DateTimeFormatter.ofPattern(fmt));

        List<Map<String, Object>> rows = db.getGlobalStatsByDateRange(from, to);

        if (rows.isEmpty()) {
            Label empty = new Label("Aucun dossier trouvé pour cette période.");
            empty.setStyle("-fx-text-fill: #9B9B9B; -fx-font-size: 13px;");
            area.getChildren().add(empty);
            return;
        }

        double totCreance  = rows.stream().mapToDouble(r -> toDouble(r.get("creance_principale"))).sum();
        double totRecouvre = rows.stream().mapToDouble(r -> toDouble(r.get("recouvre_total"))).sum();
        int    totDossiers = rows.stream().mapToInt(r -> toInt(r.get("nb_dossiers"))).sum();
        double globalPct   = totCreance > 0 ? totRecouvre / totCreance * 100.0 : 0.0;

        HBox kpis = new HBox(10);
        kpis.getChildren().addAll(
            kpi("Total créances",       fmt(totCreance) + " €",          "#1a1a1a"),
            kpi("Total recouvré",       fmt(totRecouvre) + " €",         "#0F6E56"),
            kpi("Dossiers sur période",  String.valueOf(totDossiers),  "#185FA5"),
            kpi("Dossiers soldés",       String.valueOf(rows.stream().mapToInt(r -> toInt(r.get("nb_soldes"))).sum()), "#0F6E56"),
            kpi("Commissions",           fmt(rows.stream().mapToDouble(r -> toDouble(r.get("commissions"))).sum()) + " €", "#9B9B9B")
        );
        area.getChildren().add(kpis);

        double maxCreance = rows.stream().mapToDouble(r -> toDouble(r.get("creance_principale"))).max().orElse(1.0);

        VBox tableCard = new VBox(6);
        tableCard.setStyle("-fx-background-color: #F2F0EB; -fx-background-radius: 8px; -fx-padding: 12;");

        HBox hdr = new HBox();
        hdr.setPadding(new Insets(0, 0, 6, 0));
        hdr.setStyle("-fx-border-color: rgba(0,0,0,0.10); -fx-border-width: 0 0 1 0;");
        hdr.getChildren().addAll(
            colHdr("Société",  200),
            colHdr("Créance",  110),
            colHdr("Recouvré", 110),
            colHdr("Taux",      70),
            colHdr("Dossiers",  70)
        );
        tableCard.getChildren().add(hdr);

        ScrollPane scroll = new ScrollPane();
        scroll.setFitToWidth(true);
        scroll.setStyle("-fx-background-color: transparent; -fx-background: transparent; -fx-border-color: transparent;");
        VBox.setVgrow(scroll, Priority.ALWAYS);

        VBox tableRows = new VBox(2);
        for (Map<String, Object> r : rows) {
            double creance  = toDouble(r.get("creance_principale"));
            double recouvre = toDouble(r.get("recouvre_total"));
            double pct      = creance > 0 ? recouvre / creance * 100.0 : 0.0;
            int    nb       = toInt(r.get("nb_dossiers"));
            String name     = str(r, "name");

            HBox row = new HBox();
            row.setAlignment(Pos.CENTER_LEFT);
            row.setPadding(new Insets(5, 0, 5, 0));
            row.setStyle("-fx-border-color: rgba(0,0,0,0.06); -fx-border-width: 0 0 1 0;");

            double barW = maxCreance > 0 ? creance / maxCreance * 180 : 0;
            StackPane nameCell = new StackPane();
            nameCell.setMinWidth(200); nameCell.setMaxWidth(200);
            nameCell.setAlignment(Pos.CENTER_LEFT);
            Rectangle bar = new Rectangle(Math.max(1, barW), 22);
            bar.setFill(Color.web("#E6F1FB"));
            bar.setArcWidth(3); bar.setArcHeight(3);
            Label nameLbl = new Label(abbreviate(name, 24));
            nameLbl.setStyle("-fx-font-size: 11px; -fx-text-fill: #1a1a1a; -fx-padding: 0 0 0 6;");
            nameCell.getChildren().addAll(bar, nameLbl);

            row.getChildren().addAll(
                nameCell,
                colVal(fmt(creance) + " €",         110, "#1a1a1a"),
                colVal(fmt(recouvre) + " €",         110, "#0F6E56"),
                colVal(String.format("%.1f%%", pct),  70,
                    pct >= 50 ? "#0F6E56" : pct >= 25 ? "#BA7517" : "#A32D2D"),
                colVal(String.valueOf(nb),             70, "#185FA5")
            );
            tableRows.getChildren().add(row);
        }
        scroll.setContent(tableRows);
        tableCard.getChildren().add(scroll);
        VBox.setVgrow(tableCard, Priority.ALWAYS);
        area.getChildren().add(tableCard);
    }

    private Label colHdr(String text, double width) {
        Label l = new Label(text);
        l.setMinWidth(width); l.setMaxWidth(width);
        l.setStyle("-fx-font-size: 10px; -fx-font-weight: bold; -fx-text-fill: #9B9B9B;");
        return l;
    }

    private Label colVal(String text, double width, String color) {
        Label l = new Label(text);
        l.setMinWidth(width); l.setMaxWidth(width);
        l.setStyle("-fx-font-size: 11px; -fx-text-fill: " + color + ";");
        return l;
    }

    // =========================================================================
    // Empty state
    // =========================================================================

    private VBox buildEmptyState() {
        Label lbl = new Label("Aucune société synchronisée.\nCliquez sur « Sync sociétés » pour charger les données.");
        lbl.setStyle("-fx-text-fill: #9B9B9B; -fx-font-size: 13px; -fx-text-alignment: center;");
        lbl.setWrapText(true);
        VBox box = new VBox(lbl);
        box.setAlignment(Pos.CENTER);
        box.setPadding(new Insets(60));
        VBox.setVgrow(box, Priority.ALWAYS);
        return box;
    }

    // =========================================================================
    // Sync
    // =========================================================================

    private void doSync(HBox root, Button syncBtn) {
        String rootPath = AppPreferences.getMergeRoot();
        if (rootPath.isBlank()) { log.accept("[Dashboard] Dossier racine non configuré."); return; }
        File rootFolder = new File(rootPath);
        if (!rootFolder.isDirectory()) { log.accept("[Dashboard] Dossier introuvable : " + rootPath); return; }

        syncBtn.setDisable(true);
        syncBtn.setText("↻  En cours…");
        log.accept("[Dashboard] Synchronisation…");

        executor.submit(() -> {
            try {
                syncService.syncAll(rootFolder, (pct, msg) -> Platform.runLater(() -> log.accept(msg)));
                Platform.runLater(() -> {
                    syncBtn.setDisable(false);
                    syncBtn.setText("↻  Sync sociétés");
                    log.accept("[Dashboard] ✓ Terminé.");
                    VBox container = (VBox) root.getParent().getParent();
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

    private String fmtK(double v) {
        if (v >= 1_000_000) return String.format("%.1fM €", v / 1_000_000);
        if (v >= 1_000)     return String.format("%.0fk €", v / 1_000);
        return String.format("%.0f €", v);
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
