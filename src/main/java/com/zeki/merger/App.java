package com.zeki.merger;

import com.zeki.merger.db.DatabaseManager;
import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;

/**
 * JavaFX entry point.  Run with: {@code mvn javafx:run}
 */
public class App extends Application {

    @Override
    public void start(Stage primaryStage) throws Exception {
        DatabaseManager.initialize();

        FXMLLoader loader = new FXMLLoader(
            getClass().getResource("/com/zeki/merger/main.fxml"));
        Parent root = loader.load();

        Scene scene = new Scene(root, 920, 680);
        scene.getStylesheets().add(
            getClass().getResource("/com/zeki/merger/styles.css").toExternalForm());

        primaryStage.setTitle("Cabinet Phénix");
        primaryStage.setScene(scene);
        primaryStage.setMinWidth(720);
        primaryStage.setMinHeight(540);
        primaryStage.show();
    }

    public static void main(String[] args) {
        launch(args);
    }
}
