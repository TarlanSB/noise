package com.tsb.noise.controller.handlers;

import com.tsb.noise.service.utils.FileUtils;
import com.tsb.noise.service.utils.PreferencesService;
import javafx.application.Platform;
import javafx.scene.control.Label;
import javafx.stage.DirectoryChooser;
import javafx.stage.Window;

import java.io.File;
import java.util.List;
import java.util.function.Consumer;

/**
 * Обработчик выбора директории и управления путями
 */
public class DirectorySelectionHandler {

    private final PreferencesService preferencesService;
    private final Window parentWindow;
    private final Label pathLabel;
    private final Consumer<String> logInfoCallback;
    private final Runnable updateStatusCallback;

    private String currentSelectedPath;

    public DirectorySelectionHandler(
            PreferencesService preferencesService,
            Window parentWindow,
            Label pathLabel,
            Consumer<String> logInfoCallback,
            Runnable updateStatusCallback) {

        this.preferencesService = preferencesService;
        this.parentWindow = parentWindow;
        this.pathLabel = pathLabel;
        this.logInfoCallback = logInfoCallback;
        this.updateStatusCallback = updateStatusCallback;
    }

    public void selectDirectory() {
        try {
            DirectoryChooser directoryChooser = new DirectoryChooser();
            directoryChooser.setTitle("Выберите папку с файлами УЗД");

            if (currentSelectedPath != null && !currentSelectedPath.isEmpty()) {
                File initialDir = new File(currentSelectedPath);
                if (initialDir.exists()) {
                    directoryChooser.setInitialDirectory(initialDir);
                }
            }

            File selectedDirectory = directoryChooser.showDialog(parentWindow);
            if (selectedDirectory != null) {
                setCurrentPath(selectedDirectory.getAbsolutePath());
                preferencesService.saveLastSelectedPath(currentSelectedPath);
                logInfoCallback.accept("✅ Выбрана папка: " + currentSelectedPath);

                logFoundFilesInfo();
            }
        } catch (Exception e) {
            logInfoCallback.accept("❌ Ошибка выбора директории: " + e.getMessage());
        }
    }

    public void loadLastSelectedPath() {
        try {
            String lastPath = preferencesService.getLastSelectedPath();
            if (!lastPath.isEmpty()) {
                setCurrentPath(lastPath);
                logInfoCallback.accept("📂 Загружен последний выбранный путь: " + lastPath);
            }
        } catch (Exception e) {
            logInfoCallback.accept("❌ Ошибка загрузки сохраненного пути: " + e.getMessage());
        }
    }

    private void setCurrentPath(String path) {
        currentSelectedPath = path;
        Platform.runLater(() -> {
            pathLabel.setText(currentSelectedPath);
            updateStatusCallback.run();
        });
    }

    private void logFoundFilesInfo() {
        try {
            List<File> allFiles = FileUtils.findTargetExcelFiles(currentSelectedPath);

            if (!allFiles.isEmpty()) {
                logInfoCallback.accept("📊 Найдено файлов:");
                allFiles.forEach(file -> {
                    String fileType = FileUtils.getFileTypeDisplayName(file.getName());
                    logInfoCallback.accept("   • " + fileType + ": " + file.getName());
                });
            } else {
                logInfoCallback.accept("⚠️ Файлы для обработки не найдены");
            }
        } catch (Exception e) {
            logInfoCallback.accept("⚠️ Ошибка при поиске файлов: " + e.getMessage());
        }
    }

    public String getCurrentPath() {
        return currentSelectedPath;
    }

    public boolean hasValidPath() {
        return currentSelectedPath != null && !currentSelectedPath.isEmpty();
    }
}