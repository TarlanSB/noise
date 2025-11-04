package com.tsb.noise.controller.components;

import javafx.application.Platform;
import javafx.concurrent.Task;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressBar;
import javafx.scene.layout.VBox;

/**
 * Управление прогресс-баром и статусом выполнения задач
 */
public class ProgressManager {
    private final ProgressBar progressBar;
    private final Label statusLabel;
    private final VBox progressContainer;

    public ProgressManager(ProgressBar progressBar, Label statusLabel, VBox progressContainer) {
        this.progressBar = progressBar;
        this.statusLabel = statusLabel;
        this.progressContainer = progressContainer;
        hideProgress();
    }

    public void setupTaskHandlers(Task<Void> task) {
        progressBar.progressProperty().bind(task.progressProperty());
        statusLabel.textProperty().bind(task.messageProperty());
    }

    public void showProgress() {
        Platform.runLater(() -> {
            progressContainer.setVisible(true);
            progressBar.setVisible(true);
            statusLabel.setText("🔍 Выполняется поиск файлов...");
        });
    }

    public void hideProgress() {
        Platform.runLater(() -> {
            progressContainer.setVisible(false);
            progressBar.setVisible(false);
            progressBar.progressProperty().unbind();
            statusLabel.textProperty().unbind();
            statusLabel.setText("");
        });
    }

    public void updateProgressMessage(String message) {
        Platform.runLater(() -> {
            if (statusLabel != null) {
                statusLabel.setText(message);
            }
        });
    }

    public void updateProgressValue(Double progress) {
        Platform.runLater(() -> {
            if (progressBar != null) {
                progressBar.setProgress(progress / 100.0);
            }
        });
    }
}