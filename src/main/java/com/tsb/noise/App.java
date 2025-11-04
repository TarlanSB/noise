package com.tsb.noise;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.net.URL;

public class App extends Application {
    private static final Logger log = LoggerFactory.getLogger(App.class);

    @Override
    public void start(Stage stage) throws Exception {
        try {
            // Загрузка FXML
            URL fxmlUrl = getClass().getResource("/com/tsb/noise/main-view.fxml");
            if (fxmlUrl == null) {
                throw new RuntimeException("FXML file not found: /com/tsb/noise/main-view.fxml");
            }

            FXMLLoader fxmlLoader = new FXMLLoader(fxmlUrl);
            Parent root = fxmlLoader.load();

            // Создание сцены
            Scene scene = new Scene(root, 1000, 800);

            // Загрузка CSS с проверкой
            URL cssUrl = getClass().getResource("/styles.css");
            if (cssUrl != null) {
                scene.getStylesheets().add(cssUrl.toExternalForm());
                log.info("CSS styles loaded successfully");
            } else {
                log.warn("CSS file not found, application will run without styles");
            }

            stage.setTitle("📊 Мероприятия по защите от шума");
            stage.setScene(scene);
            stage.setMinWidth(900);
            stage.setMinHeight(700);
            stage.show();

            log.info("Application started successfully");
        } catch (Exception e) {
            log.error("Error starting application: {}", e.getMessage(), e);
            throw e;
        }
    }

    public static void main(String[] args) {
        launch();
    }
}