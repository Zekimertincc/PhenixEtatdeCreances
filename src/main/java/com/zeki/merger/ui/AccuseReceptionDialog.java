package com.zeki.merger.ui;

import com.zeki.merger.AppPreferences;
import com.zeki.merger.db.DatabaseManager;
import com.zeki.merger.service.FacturationMailService;
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
    private final FacturationMailService service = new FacturationMailService();
    private FacturationMailService.Signataire selectedSignataire = FacturationMailService.Signataire.ANONYME;
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

        SplitPane split = new SplitPane();
        split.setDividerPositions(0.42);
        VBox.setVgrow(split, Priority.ALWAYS);
        split.getItems().addAll(buildLeftPane(), buildRightPane());

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

        HBox shortcuts = new HBox(6);
        Button btnAll  = new Button("Tout sélectionner");
        Button btnNone = new Button("Tout désélectionner");
        btnAll.setOnAction(e  -> { allRows.forEach(r -> r.setSelected(true));  clientList.refresh(); });
        btnNone.setOnAction(e -> { allRows.forEach(r -> r.setSelected(false)); clientList.refresh(); });
        shortcuts.getChildren().addAll(btnAll, btnNone);

        clientList.setItems(allRows);
        clientList.setCellFactory(lv -> new CheckBoxListCell());
        VBox.setVgrow(clientList, Priority.ALWAYS);

        pane.getChildren().addAll(new Label("Clients :"), shortcuts, clientList);
        return pane;
    }

    private VBox buildRightPane() {
        VBox pane = new VBox(8);
        pane.setPadding(new Insets(4));

        subjectField.setPromptText("Objet du mail...");
        subjectField.setText("Cabinet Phénix, accusé de réception de dossier(s)");

        // — Custom templates —
        HBox customTplRow = new HBox(6);
        customTplRow.setAlignment(Pos.CENTER_LEFT);
        ComboBox<String> tplCombo = new ComboBox<>();
        tplCombo.setPromptText("Mes modèles…");
        tplCombo.setPrefWidth(180);
        refreshTemplateCombo(tplCombo);

        Button btnLoadTpl = new Button("Charger");
        btnLoadTpl.setOnAction(e -> {
            String sel = tplCombo.getValue();
            if (sel == null || sel.isBlank()) return;
            DatabaseManager.getInstance().getAllMailTemplates().stream()
                    .filter(m -> sel.equals(m.get("name")))
                    .findFirst()
                    .ifPresent(m -> bodyArea.setText(m.get("body")));
        });

        Button btnSaveTpl = new Button("💾 Enregistrer");
        btnSaveTpl.setOnAction(e -> {
            String current = tplCombo.getValue();
            String defaultName = current != null && !current.isBlank() ? current : "";
            TextInputDialog dlg = new TextInputDialog(defaultName);
            dlg.setTitle("Enregistrer le modèle");
            dlg.setHeaderText(null);
            dlg.setContentText("Nom du modèle :");
            dlg.showAndWait().ifPresent(name -> {
                if (name.isBlank()) return;
                try {
                    DatabaseManager.getInstance().upsertMailTemplate(name.trim(), bodyArea.getText());
                    refreshTemplateCombo(tplCombo);
                    tplCombo.setValue(name.trim());
                } catch (Exception ex) {
                    log.accept("ERREUR sauvegarde modèle : " + ex.getMessage());
                }
            });
        });

        Button btnDeleteTpl = new Button("🗑");
        btnDeleteTpl.setStyle("-fx-text-fill: #A32D2D;");
        btnDeleteTpl.setOnAction(e -> {
            String sel = tplCombo.getValue();
            if (sel == null || sel.isBlank()) return;
            Alert confirm = new Alert(Alert.AlertType.CONFIRMATION,
                    "Supprimer le modèle « " + sel + " » ?",
                    ButtonType.YES, ButtonType.NO);
            confirm.setHeaderText(null);
            confirm.showAndWait().ifPresent(bt -> {
                if (bt == ButtonType.YES) {
                    try {
                        DatabaseManager.getInstance().deleteMailTemplate(sel);
                        refreshTemplateCombo(tplCombo);
                        tplCombo.setValue(null);
                    } catch (Exception ex) {
                        log.accept("ERREUR suppression modèle : " + ex.getMessage());
                    }
                }
            });
        });

        customTplRow.getChildren().addAll(
                new Label("Mes modèles :"), tplCombo, btnLoadTpl, btnSaveTpl, btnDeleteTpl);

        bodyArea.setWrapText(true);
        bodyArea.setText(
            "Bonjour,\n\n" +
            "Nous vous prions de bien vouloir trouver ci-joint votre (vos) état(s) de créances " +
            "à jour prenant en compte les derniers dossiers que vous nous avez confiés.\n\n" +
            "Les derniers dossiers confiés seront affichés en premier.\n\n" +
            "Notre équipe met en œuvre toutes les diligences nécessaires afin de recouvrer vos créances.\n\n" +
            "Si l'un de vos débiteurs venait à vous contacter, nous vous conseillons de l'orienter vers nous.\n\n" +
            "Par ailleurs, si un règlement vous est adressé directement, nous vous remercions de nous en tenir informés.\n\n" +
            "Pour toute remise de nouveaux dossiers, nous vous invitons à nous les adresser par e-mail aux adresses " +
            "contact@cabinetphenix.fr, info@cabinetphenix.fr ou par courrier.\n\n" +
            "Nous restons à votre disposition pour toute information."
        );
        VBox.setVgrow(bodyArea, Priority.ALWAYS);

        // Signataire selection
        HBox signataireRow = new HBox(8);
        signataireRow.setAlignment(Pos.CENTER_LEFT);
        ToggleGroup sigGroup = new ToggleGroup();
        RadioButton rbAnonyme  = new RadioButton("Anonyme");
        RadioButton rbJulien   = new RadioButton("Julien JOUSSET");
        RadioButton rbGauthier = new RadioButton("Gauthier BERIS");
        rbAnonyme.setToggleGroup(sigGroup);
        rbJulien.setToggleGroup(sigGroup);
        rbGauthier.setToggleGroup(sigGroup);

        String saved = AppPreferences.getMailSignataire();
        if ("JULIEN".equals(saved))        rbJulien.setSelected(true);
        else if ("GAUTHIER".equals(saved)) rbGauthier.setSelected(true);
        else                               rbAnonyme.setSelected(true);
        selectedSignataire = toSignataire(saved);

        sigGroup.selectedToggleProperty().addListener((obs, old, nw) -> {
            if (nw == rbJulien)        { selectedSignataire = FacturationMailService.Signataire.JULIEN;   AppPreferences.setMailSignataire("JULIEN"); }
            else if (nw == rbGauthier) { selectedSignataire = FacturationMailService.Signataire.GAUTHIER; AppPreferences.setMailSignataire("GAUTHIER"); }
            else                       { selectedSignataire = FacturationMailService.Signataire.ANONYME;  AppPreferences.setMailSignataire("ANONYME"); }
        });
        signataireRow.getChildren().addAll(new Label("Signature :"), rbAnonyme, rbJulien, rbGauthier);

        pane.getChildren().addAll(
                new Label("Objet :"), subjectField,
                customTplRow,
                new Label("Message :"), bodyArea,
                signataireRow);
        return pane;
    }

    // -------------------------------------------------------------------------
    // Data loading
    // -------------------------------------------------------------------------

    private void loadData() {
        String listingPath = AppPreferences.getTrfListing();
        if (listingPath != null && !listingPath.isBlank()) {
            try {
                clientInfoMap = new DataReader().readClientInfoMap(new File(listingPath));
            } catch (Exception e) {
                log.accept("AVERT: Listing introuvable — " + e.getMessage());
            }
        }

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
        // TODO: COMP PART. ve DÉBITEURS için TRF classification eklenecek
        return "VIREMENT";
    }

    private void refreshTemplateCombo(ComboBox<String> combo) {
        String current = combo.getValue();
        combo.getItems().clear();
        DatabaseManager.getInstance().getAllMailTemplates()
                .forEach(m -> combo.getItems().add(m.get("name")));
        if (current != null && combo.getItems().contains(current)) {
            combo.setValue(current);
        }
    }

    private FacturationMailService.Signataire toSignataire(String s) {
        if ("JULIEN".equals(s))   return FacturationMailService.Signataire.JULIEN;
        if ("GAUTHIER".equals(s)) return FacturationMailService.Signataire.GAUTHIER;
        return FacturationMailService.Signataire.ANONYME;
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
    // Draft creation — PJ = état des créances (Espace partagé/Etat des créances)
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

        service.cleanPreviousDraftFolder(lastDraftFolder);

        List<FacturationMailService.DraftRequest> drafts = new ArrayList<>();

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

            drafts.add(new FacturationMailService.DraftRequest(
                    ci.getName(), email, subject, body, attachPath, selectedSignataire));
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
            return ci.getName();
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
