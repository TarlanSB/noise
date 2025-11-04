package com.tsb.noise.controller.managers;

import com.tsb.noise.service.utils.FileUtils;
import com.tsb.noise.model.FileType;
import javafx.scene.control.Label;

import java.io.File;
import java.util.List;

/**
 * Управление статусом и отображением состояния приложения
 */
public class StatusManager {

    private final Label statusLabel;
    private final Label processStatusLabel;

    public StatusManager(Label statusLabel, Label processStatusLabel) {
        this.statusLabel = statusLabel;
        this.processStatusLabel = processStatusLabel;
    }

    public void updateStatus(String directoryPath,
                             boolean hasSelectedFileTypes,
                             List<FileType> selectedFileTypes,
                             boolean hasOperations) {

        if (!hasValidPath(directoryPath)) {
            setStatusWarning("Выберите папку с файлами");
            return;
        }

        if (!hasSelectedFileTypes) {
            setStatusWarning("Выберите типы файлов для обработки");
            return;
        }

        int fileCount = countMatchingFiles(directoryPath, selectedFileTypes);
        updateFileStatus(fileCount, hasOperations);
    }

    private boolean hasValidPath(String directoryPath) {
        return directoryPath != null && !directoryPath.isEmpty();
    }

    private int countMatchingFiles(String directoryPath, List<FileType> selectedTypes) {
        try {
            List<File> allFiles = FileUtils.findTargetExcelFiles(directoryPath);
            return (int) allFiles.stream()
                    .filter(file -> {
                        FileType fileType = FileType.fromFileName(file.getName());
                        return fileType != null && selectedTypes.contains(fileType);
                    })
                    .count();
        } catch (Exception e) {
            return 0;
        }
    }

    private void updateFileStatus(int fileCount, boolean hasOperations) {
        if (fileCount == 0) {
            setStatusError("❌ Файлы не найдены для выбранных типов");
        } else {
            String statusText = "✅ Найдено файлов: " + fileCount;
            if (hasOperations) {
                statusText += " (операции применены)";
            }
            setStatusSuccess(statusText);
        }
    }

    public void setStatusSuccess(String message) {
        setStatus(message, "status-success", "status-error", "status-warning");
    }

    public void setStatusError(String message) {
        setStatus(message, "status-error", "status-success", "status-warning");
    }

    public void setStatusWarning(String message) {
        setStatus(message, "status-warning", "status-success", "status-error");
    }

    private void setStatus(String message, String addStyle, String... removeStyles) {
        processStatusLabel.setText(message);
        processStatusLabel.getStyleClass().removeAll(removeStyles);
        processStatusLabel.getStyleClass().add(addStyle);
    }

    public void updateProgressStatus(String message) {
        statusLabel.setText(message);
    }

    public void clearProgressStatus() {
        statusLabel.setText("");
    }
}