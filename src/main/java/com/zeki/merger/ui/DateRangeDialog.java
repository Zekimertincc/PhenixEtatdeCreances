package com.zeki.merger.ui;

import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.stage.Modality;
import javafx.stage.Stage;
import javafx.util.StringConverter;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Optional;

public class DateRangeDialog {

    public record DateRange(LocalDate dateDebut, LocalDate dateFin) {}

    private static final DateTimeFormatter FR = DateTimeFormatter.ofPattern("dd/MM/yyyy");

    private final Stage stage = new Stage();
    private DateRange result = null;

    private final DatePicker dateFrom = new DatePicker(LocalDate.now().withDayOfMonth(1));
    private final DatePicker dateTo   = new DatePicker(LocalDate.now());

    public DateRangeDialog(Stage owner, String title) {
        stage.initOwner(owner);
        stage.initModality(Modality.WINDOW_MODAL);
        stage.setTitle(title);
        stage.setWidth(360);
        stage.setResizable(false);
        stage.setScene(new Scene(buildRoot()));
    }

    public Optional<DateRange> showAndWait() {
        stage.showAndWait();
        return Optional.ofNullable(result);
    }

    private VBox buildRoot() {
        VBox root = new VBox(14);
        root.setPadding(new Insets(18));

        StringConverter<LocalDate> cvt = new StringConverter<>() {
            public String toString(LocalDate d)   { return d == null ? "" : d.format(FR); }
            public LocalDate fromString(String s) {
                try { return s == null || s.isBlank() ? null : LocalDate.parse(s.trim(), FR); }
                catch (Exception e) { return null; }
            }
        };
        dateFrom.setConverter(cvt);
        dateTo.setConverter(cvt);
        dateFrom.setPromptText("dd/MM/yyyy");
        dateTo.setPromptText("dd/MM/yyyy");

        GridPane grid = new GridPane();
        grid.setHgap(10);
        grid.setVgap(8);
        grid.add(new Label("Remis Le — du :"), 0, 0); grid.add(dateFrom, 1, 0);
        grid.add(new Label("au :"),             0, 1); grid.add(dateTo,   1, 1);

        Button btnLancer  = new Button("Lancer");
        Button btnAnnuler = new Button("Annuler");
        btnLancer.setDefaultButton(true);
        btnLancer.setStyle("-fx-background-color:#1F4E79;-fx-text-fill:white;-fx-font-weight:bold;");
        btnLancer.setPrefWidth(100);
        btnAnnuler.setCancelButton(true);
        btnAnnuler.setPrefWidth(100);

        btnLancer.setOnAction(e -> {
            result = new DateRange(dateFrom.getValue(), dateTo.getValue());
            stage.close();
        });
        btnAnnuler.setOnAction(e -> stage.close());

        HBox btnRow = new HBox(10, btnAnnuler, btnLancer);
        btnRow.setAlignment(Pos.CENTER_RIGHT);

        root.getChildren().addAll(grid, new Separator(), btnRow);
        return root;
    }
}
