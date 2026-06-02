package com.zeki.merger.ui;

import com.zeki.merger.AppPreferences;
import com.zeki.merger.service.AccuseReceptionService;
import com.zeki.merger.trf.DataReader;
import com.zeki.merger.trf.model.ClientInfo;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.stage.Modality;
import javafx.stage.Stage;

import java.io.File;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.function.Consumer;

public class AccuseReceptionDialog {

    private static final DateTimeFormatter FR =
            DateTimeFormatter.ofPattern("dd/MM/yyyy");

    private final Stage stage = new Stage();
    private final Consumer<String> log;

    // Left
    private final DatePicker dateFrom = new DatePicker(LocalDate.now().withDayOfMonth(1));
    private final DatePicker dateTo   = new DatePicker(LocalDate.now());
    private final ListView<ClientRow> clientList = new ListView<>();
    private final ObservableList<ClientRow> allRows = FXCollections.observableArrayList();

    // Right
    private final TextField subjectField = new TextField();
    private final TextArea  bodyArea     = new TextArea();

    // State
    private Map<String, ClientInfo>  clientInfoMap     = new LinkedHashMap<>();
    private Map<String, String>      correspondanceMap = new LinkedHashMap<>();
    private final AccuseReceptionService service = new AccuseReceptionService();
    private File lastDraftFolder = null;

    // -------------------------------------------------------------------------

    public AccuseReceptionDialog(Stage owner, Consumer<String> log) {
        this.log = log;
        stage.initOwner(owner);
        stage.initModality(Modality.WINDOW_MODAL);
        stage.setTitle("Accusés de réception");
        stage.setWidth(900);
        stage.setHeight(620);
        stage.setResizable(true);
        stage.setScene(new Scene(buildRoot()));
        loadData();
    }

    public void show() { stage.show(); }

    // -------------------------------------------------------------------------
    // UI build
    // -------------------------------------------------------------------------

    private VBox buildRoot() {
        VBox root = new VBox(10);
        root.setPadding(new Insets(12));

        // — Top: date range row —
        HBox dateRow = new HBox(8);
        dateRow.setAlignment(Pos.CENTER_LEFT);
        dateFrom.setConverter(makeDateConverter());
        dateTo.setConverter(makeDateConverter());
        Button btnFilter = new Button("Filtrer");
        btnFilter.setOnAction(e -> applyFilter());
        dateRow.getChildren().addAll(
                new Label("Du :"), dateFrom,
                new Label("Au :"), dateTo,
                btnFilter);

        // — Center: split pane —
        SplitPane split = new SplitPane();
        split.setDividerPositions(0.42);
        VBox.setVgrow(split, Priority.ALWAYS);
        split.getItems().addAll(buildLeftPane(), buildRightPane());

        // — Bottom buttons —
        HBox bottomRow = new HBox(10);
        bottomRow.setAlignment(Pos.CENTER_RIGHT);
        Button btnCancel = new Button("Annuler");
        btnCancel.setOnAction(e -> stage.close());
        Button btnSend = new Button("Créer les drafts");
        btnSend.setDefaultButton(true);
        btnSend.setStyle("-fx-background-color: #2e7d32; -fx-text-fill: white; -fx-font-weight: bold;");
        btnSend.setOnAction(e -> createDrafts());
        bottomRow.getChildren().addAll(btnCancel, btnSend);

        root.getChildren().addAll(dateRow, split, bottomRow);
        return root;
    }

    private VBox buildLeftPane() {
        VBox pane = new VBox(8);
        pane.setPadding(new Insets(4));

        // Shortcut filter buttons
        HBox shortcuts = new HBox(6);
        Button btnAll      = new Button("Tout");
        Button btnComp     = new Button("COMP");
        Button btnNonComp  = new Button("NON COMP");
        Button btnPartiel  = new Button("COMP PART.");
        btnAll.setOnAction(e     -> selectByType(null));
        btnComp.setOnAction(e    -> selectByType("COMP"));
        btnNonComp.setOnAction(e -> selectByType("NON COMP"));
        btnPartiel.setOnAction(e -> selectByType("COMP PART."));
        shortcuts.getChildren().addAll(btnAll, btnComp, btnNonComp, btnPartiel);

        // Client list with checkboxes
        clientList.setItems(allRows);
        clientList.setCellFactory(lv -> new CheckBoxListCell());
        VBox.setVgrow(clientList, Priority.ALWAYS);

        pane.getChildren().addAll(new Label("Clients :"), shortcuts, clientList);
        return pane;
    }

    private VBox buildRightPane() {
        VBox pane = new VBox(8);
        pane.setPadding(new Insets(4));

        // Subject
        subjectField.setPromptText("Objet du mail...");
        subjectField.setText("Cabinet Phénix, votre état des créances");

        // Template shortcuts
        HBox templateRow = new HBox(6);
        Button btnTplVirement  = new Button("Virement");
        Button btnTplNonComp   = new Button("Non Comp");
        Button btnTplPartielle = new Button("Comp Partielle");
        btnTplVirement.setOnAction(e  -> bodyArea.setText(service.buildBody(AccuseReceptionService.CompType.VIREMENT)));
        btnTplNonComp.setOnAction(e   -> bodyArea.setText(service.buildBody(AccuseReceptionService.CompType.NON_COMP)));
        btnTplPartielle.setOnAction(e -> bodyArea.setText(service.buildBody(AccuseReceptionService.CompType.COMP_PARTIELLE)));
        templateRow.getChildren().addAll(
                new Label("Charger modèle :"),
                btnTplVirement, btnTplNonComp, btnTplPartielle);

        // Body
        bodyArea.setWrapText(true);
        bodyArea.setText(service.buildBody(AccuseReceptionService.CompType.VIREMENT));
        VBox.setVgrow(bodyArea, Priority.ALWAYS);

        pane.getChildren().addAll(
                new Label("Objet :"), subjectField,
                templateRow,
                new Label("Message :"), bodyArea);
        return pane;
    }

    // -------------------------------------------------------------------------
    // Data loading
    // -------------------------------------------------------------------------

    private void loadData() {
        // Load Listing
        String listingPath = AppPreferences.getTrfListing();
        if (listingPath != null && !listingPath.isBlank()) {
            try {
                clientInfoMap = new DataReader().readClientInfoMap(new File(listingPath));
            } catch (Exception e) {
                log.accept("AVERT: Listing introuvable — " + e.getMessage());
            }
        }

        // Load Correspondance
        String corrPath = AppPreferences.getCorrespondancePath();
        if (corrPath != null && !corrPath.isBlank()) {
            try {
                correspondanceMap = service.readCorrespondanceMap(new File(corrPath));
            } catch (Exception e) {
                log.accept("AVERT: Correspondance introuvable — " + e.getMessage());
            }
        }

        applyFilter();
    }

    private void applyFilter() {
        LocalDate from = dateFrom.getValue();
        LocalDate to   = dateTo.getValue();
        if (from == null || to == null) return;

        List<ClientInfo> filtered = service.filterByDateRange(clientInfoMap, from, to);
        allRows.clear();
        for (ClientInfo ci : filtered) {
            String type = resolveType(ci);
            allRows.add(new ClientRow(ci, type, true));
        }
        long nullCount = clientInfoMap.values().stream()
                .filter(ci -> ci.getDateLastDossier() == null).count();
        log.accept("Filtre appliqué : " + allRows.size() + " client(s) trouvé(s). "
                + "(dateLastDossier null: " + nullCount + "/" + clientInfoMap.size() + ")");
    }

    private String resolveType(ClientInfo ci) {
        if (ci.isNonCompensation()) return "NON COMP";
        // TODO: comp_partielle detection when TRF classification available
        return "COMP";
    }

    // -------------------------------------------------------------------------
    // Selection helpers
    // -------------------------------------------------------------------------

    private void selectByType(String type) {
        for (ClientRow row : allRows) {
            row.setSelected(type == null || row.getType().equals(type));
        }
        clientList.refresh();
    }

    // -------------------------------------------------------------------------
    // Draft creation
    // -------------------------------------------------------------------------

    private void createDrafts() {
        List<ClientRow> selected = allRows.stream()
                .filter(ClientRow::isSelected)
                .toList();

        if (selected.isEmpty()) {
            new Alert(Alert.AlertType.WARNING, "Aucun client sélectionné.").showAndWait();
            return;
        }

        String subject = subjectField.getText().trim();
        String body    = bodyArea.getText().trim();

        // Önceki klasörü temizle
        service.cleanPreviousDraftFolder(lastDraftFolder);

        List<AccuseReceptionService.DraftRequest> drafts = new ArrayList<>();

        for (ClientRow row : selected) {
            ClientInfo ci = row.getClientInfo();

            String email = ci.getEmail();
            if (email.isBlank()) {
                log.accept("AVERT: Pas d'email pour " + ci.getName() + " — ignoré");
                continue;
            }

            String rootPath = AppPreferences.getMergeRoot();
            File rootFolder = (rootPath != null && !rootPath.isBlank()) ? new File(rootPath) : null;
            File attachment = service.findEtatPublicForClient(ci.getName(), rootFolder);
            String attachPath = attachment != null ? attachment.getAbsolutePath() : "";

            log.accept("Draft → " + ci.getName() + " <" + email + ">"
                    + (attachPath.isBlank() ? " [sans PJ]" : " [" + attachment.getName() + "]"));

            drafts.add(new AccuseReceptionService.DraftRequest(
                    ci.getName(), email, subject, body, attachPath));
        }

        if (drafts.isEmpty()) {
            new Alert(Alert.AlertType.WARNING, "Aucun client avec email valide.").showAndWait();
            return;
        }

        try {
            lastDraftFolder = service.prepareDraftFolder(drafts);
            log.accept(drafts.size() + " draft(s) préparé(s) → " + lastDraftFolder.getAbsolutePath());

            Alert info = new Alert(Alert.AlertType.INFORMATION);
            info.setTitle("Drafts prêts");
            info.setHeaderText(drafts.size() + " draft(s) préparé(s)");
            info.setContentText("Le dossier s'est ouvert.\n\nDouble-cliquez sur 'lancer_tous.bat' pour envoyer tous les drafts vers Outlook.");
            info.showAndWait();

        } catch (Exception e) {
            log.accept("ERREUR: " + e.getMessage());
            new Alert(Alert.AlertType.ERROR, "Erreur: " + e.getMessage()).showAndWait();
        }
    }

    // -------------------------------------------------------------------------
    // Inner classes
    // -------------------------------------------------------------------------

    public static class ClientRow {
        private final ClientInfo ci;
        private final String type;
        private boolean selected;

        public ClientRow(ClientInfo ci, String type, boolean selected) {
            this.ci = ci; this.type = type; this.selected = selected;
        }
        public ClientInfo getClientInfo() { return ci; }
        public String     getType()       { return type; }
        public boolean    isSelected()    { return selected; }
        public void       setSelected(boolean v) { this.selected = v; }

        @Override public String toString() {
            return ci.getName() + "  —  " + type;
        }
    }

    private static class CheckBoxListCell extends ListCell<ClientRow> {
        private final CheckBox cb = new CheckBox();

        @Override
        protected void updateItem(ClientRow item, boolean empty) {
            super.updateItem(item, empty);
            if (empty || item == null) {
                setGraphic(null);
                return;
            }
            cb.setText(item.toString());
            cb.setSelected(item.isSelected());
            cb.setOnAction(e -> item.setSelected(cb.isSelected()));
            setGraphic(cb);
        }
    }

    private static javafx.util.StringConverter<LocalDate> makeDateConverter() {
        return new javafx.util.StringConverter<>() {
            @Override public String toString(LocalDate d) { return d != null ? FR.format(d) : ""; }
            @Override public LocalDate fromString(String s) {
                return (s != null && !s.isBlank()) ? LocalDate.parse(s, FR) : null;
            }
        };
    }
}
