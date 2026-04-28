package com.zeki.merger.controller;

import com.zeki.merger.db.CompanyRecord;
import com.zeki.merger.db.DatabaseManager;
import javafx.application.Platform;
import javafx.beans.property.ReadOnlyObjectWrapper;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.collections.transformation.FilteredList;
import javafx.fxml.FXML;
import javafx.scene.control.*;

import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Optional;

public class DashboardController {

    private static final Map<String, String> CREANCE_COL_LABELS = Map.ofEntries(
        Map.entry("id",         "ID"),
        Map.entry("company_id", "—"),
        Map.entry("row_index",  "#"),
        Map.entry("col_a",  "NBRE"),
        Map.entry("col_b",  "V/REF"),
        Map.entry("col_c",  "REMIS LE"),
        Map.entry("col_d",  "ANCIENNETÉ"),
        Map.entry("col_e",  "N/REF"),
        Map.entry("col_f",  "DÉBITEUR"),
        Map.entry("col_g",  "CRÉANCE PRINCIPALE"),
        Map.entry("col_h",  "RECOUVRÉ ET FACTURÉ"),
        Map.entry("col_i",  "ÉTAT"),
        Map.entry("col_j",  "CLÔTURE"),
        Map.entry("col_k",  "PÉNALITÉS"),
        Map.entry("col_l",  "DEPT"),
        Map.entry("col_m",  "TRANSF. L"),
        Map.entry("col_n",  "CONDITION"),
        Map.entry("col_o",  "DONT EN ATTENTE"),
        Map.entry("col_p",  "LIEU"),
        Map.entry("col_q",  "FRAIS PROCÉDURE"),
        Map.entry("col_r",  "RECOUVRÉ TOTAL"),
        Map.entry("col_s",  "DÉJÀ FACTURÉ"),
        Map.entry("col_t",  "DEPUIS LE DÉBUT"),
        Map.entry("col_u",  "COMMISSIONS"),
        Map.entry("col_v",  "PÉNALITS"),
        Map.entry("col_w",  "SOMMES CZ PHÉNIX"),
        Map.entry("col_x",  "MONTANT À FACTURER TTC"),
        Map.entry("col_y",  "SOMMES À REVERSER")
    );

    private static final Map<String, String> TRF_COL_LABELS = Map.ofEntries(
        Map.entry("id",                             "ID"),
        Map.entry("company_id",                     "—"),
        Map.entry("client_code",                    "CODE CLIENT"),
        Map.entry("iban",                           "IBAN"),
        Map.entry("bic",                            "BIC"),
        Map.entry("non_compensation",               "NON COMP."),
        Map.entry("creance_principale",             "CRÉANCE PRINCIPALE"),
        Map.entry("recouvre_et_facture",            "RECOUVRÉ ET FACTURÉ"),
        Map.entry("penalites",                      "PÉNALITÉS"),
        Map.entry("dont_en_attente",                "DONT EN ATTENTE"),
        Map.entry("frais_procedure",                "FRAIS PROCÉDURE"),
        Map.entry("recouvre_total",                 "RECOUVRÉ TOTAL"),
        Map.entry("deja_facture",                   "DÉJÀ FACTURÉ"),
        Map.entry("depuis_le_debut",                "DEPUIS LE DÉBUT"),
        Map.entry("commissions",                    "COMMISSIONS"),
        Map.entry("sommes_cz_phenix",               "SOMMES CZ PHÉNIX"),
        Map.entry("montant_a_facturer_ttc",         "MONTANT À FACTURER TTC"),
        Map.entry("sommes_a_reverser_src",          "SOMMES À REVERSER (SRC)"),
        Map.entry("nous_doit_prec",                 "NOUS DOIT PRÉC."),
        Map.entry("nous_doit_maintenant",           "NOUS DOIT MAINTENANT"),
        Map.entry("encaissements_par_compensation", "ENCAISS. PAR COMP."),
        Map.entry("sommes_a_reverser_final",        "SOMMES À REVERSER FINAL"),
        Map.entry("nous_doit_apre_facturation",     "NOUS DOIT APRÈS FACT."),
        Map.entry("etat_compensations",             "ÉTAT COMP."),
        Map.entry("virements",                      "VIREMENTS"),
        Map.entry("last_sync",                      "SYNC")
    );

    @FXML private TextField  searchField;
    @FXML private ListView<CompanyRecord> companyListView;
    @FXML private TableView<ObservableList<String>> creanceTable;
    @FXML private TableView<ObservableList<String>> trfTable;

    private final ObservableList<CompanyRecord> allCompanies = FXCollections.observableArrayList();
    private FilteredList<CompanyRecord> filteredCompanies;

    @FXML
    public void initialize() {
        filteredCompanies = new FilteredList<>(allCompanies, p -> true);
        companyListView.setItems(filteredCompanies);
        companyListView.setCellFactory(lv -> new ListCell<>() {
            @Override
            protected void updateItem(CompanyRecord item, boolean empty) {
                super.updateItem(item, empty);
                if (empty || item == null) { setText(null); return; }
                setText(item.name() + "  (" + item.rowCount() + ")");
            }
        });

        searchField.textProperty().addListener((obs, old, text) ->
            filteredCompanies.setPredicate(c ->
                text == null || text.isBlank() ||
                c.name().toLowerCase().contains(text.toLowerCase())));

        companyListView.getSelectionModel().selectedItemProperty()
            .addListener((obs, old, sel) -> { if (sel != null) loadDetails(sel); });

        refresh();
    }

    public void refresh() {
        Platform.runLater(() -> {
            DatabaseManager db = DatabaseManager.getInstance();
            if (db == null) return;
            try {
                CompanyRecord prev = companyListView.getSelectionModel().getSelectedItem();
                allCompanies.setAll(db.getAllCompanies());
                if (prev != null) {
                    allCompanies.stream()
                        .filter(c -> c.id() == prev.id())
                        .findFirst()
                        .ifPresent(c -> {
                            companyListView.getSelectionModel().select(c);
                            loadDetails(c);
                        });
                }
            } catch (SQLException ignored) {}
        });
    }

    private void loadDetails(CompanyRecord company) {
        DatabaseManager db = DatabaseManager.getInstance();
        if (db == null) return;
        try {
            populateTable(creanceTable, db.getCreanceRows(company.id()), null, CREANCE_COL_LABELS);
        } catch (SQLException ignored) {}
        try {
            Optional<Map<String, Object>> trf = db.getTrfSummary(company.id());
            populateTable(trfTable, trf.map(List::of).orElse(List.of()),
                "sommes_a_reverser_final", TRF_COL_LABELS);
        } catch (SQLException ignored) {}
    }

    private void populateTable(TableView<ObservableList<String>> table,
                                List<Map<String, Object>> data,
                                String highlightCol,
                                Map<String, String> labelMap) {
        table.getColumns().clear();
        table.getItems().clear();
        if (data.isEmpty()) return;

        // Build visible key list — skip internal DB columns
        List<String> keys = new ArrayList<>();
        for (String k : data.get(0).keySet()) {
            if (!k.equals("id") && !k.equals("company_id")) keys.add(k);
        }

        int highlightIdx = highlightCol != null ? keys.indexOf(highlightCol) : -1;

        for (int i = 0; i < keys.size(); i++) {
            final int col = i;
            String key   = keys.get(i);
            String label = labelMap != null ? labelMap.getOrDefault(key, key) : key;
            TableColumn<ObservableList<String>, String> tc = new TableColumn<>(label);
            tc.setCellValueFactory(p ->
                new ReadOnlyObjectWrapper<>(col < p.getValue().size() ? p.getValue().get(col) : ""));
            tc.setPrefWidth(110);
            table.getColumns().add(tc);
        }

        for (Map<String, Object> row : data) {
            ObservableList<String> rowData = FXCollections.observableArrayList();
            for (String key : keys) {
                Object v = row.get(key);
                rowData.add(v == null ? "" : v.toString());
            }
            table.getItems().add(rowData);
        }

        if (highlightIdx >= 0) {
            final int hi = highlightIdx;
            table.setRowFactory(tv -> new TableRow<>() {
                @Override
                protected void updateItem(ObservableList<String> item, boolean empty) {
                    super.updateItem(item, empty);
                    if (empty || item == null || hi >= item.size()) {
                        setStyle("");
                        return;
                    }
                    try {
                        double val = Double.parseDouble(item.get(hi).replace(",", "."));
                        setStyle(val > 0 ? "-fx-background-color: #FFE500;" : "");
                    } catch (NumberFormatException e) {
                        setStyle("");
                    }
                }
            });
        }
    }
}
