package com.tsb.noise.controller.managers;

import javafx.application.Platform;
import javafx.scene.control.TextArea;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

/**
 * Управление логом операций
 */
public class LogManager {

    private final TextArea logArea;
    private final DateTimeFormatter timeFormatter;

    public LogManager(TextArea logArea) {
        this.logArea = logArea;
        this.timeFormatter = DateTimeFormatter.ofPattern("HH:mm:ss");
    }

    public void logInfo(String message) {
        logMessage(message, "INFO");
    }

    public void logError(String message) {
        logMessage(message, "ERROR");
    }

    public void logWarning(String message) {
        logMessage(message, "WARNING");
    }

    private void logMessage(String message, String level) {
        Platform.runLater(() -> {
            String timestamp = LocalDateTime.now().format(timeFormatter);
            String formattedMessage = String.format("[%s] %s: %s%n", timestamp, level, message);
            logArea.appendText(formattedMessage);
            logArea.setScrollTop(Double.MAX_VALUE);
        });
    }

    public void clearLog() {
        Platform.runLater(() -> {
            logArea.clear();
            logInfo("🧹 Лог очищен");
            logInfo("⏰ " + getCurrentTimestamp());
        });
    }

    private String getCurrentTimestamp() {
        return LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd.MM.yyyy HH:mm:ss"));
    }
}