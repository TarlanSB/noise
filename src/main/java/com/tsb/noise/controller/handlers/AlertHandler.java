package com.tsb.noise.controller.handlers;

import javafx.application.Platform;
import javafx.scene.control.Alert;
import javafx.scene.control.DialogPane;

/**
 * Обработчик показа alert-сообщений
 */
public class AlertHandler {

    public void showInfo(String title, String message) {
        showAlert(Alert.AlertType.INFORMATION, title, message);
    }

    public void showWarning(String title, String message) {
        showAlert(Alert.AlertType.WARNING, title, message);
    }

    public void showError(String title, String message) {
        showAlert(Alert.AlertType.ERROR, title, message);
    }

    private void showAlert(Alert.AlertType alertType, String title, String message) {
        Platform.runLater(() -> {
            Alert alert = new Alert(alertType);
            alert.setTitle(title);
            alert.setHeaderText(null);
            alert.setContentText(message);

            DialogPane dialogPane = alert.getDialogPane();
            dialogPane.getStyleClass().add("custom-alert");

            alert.showAndWait();
        });
    }
}