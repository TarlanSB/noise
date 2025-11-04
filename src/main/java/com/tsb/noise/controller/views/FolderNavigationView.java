package com.tsb.noise.controller.views;

import com.tsb.noise.controller.handlers.AlertHandler;
import com.tsb.noise.controller.managers.LogManager;
import javafx.scene.control.Button;
import javafx.scene.control.Label;

import java.io.File;
import java.util.function.Consumer;

/**
 * Управление навигацией по папкам
 */
public class FolderNavigationView {
    private final AlertHandler alertHandler;
    private final LogManager logManager;
    private final Consumer<String> openFolderCallback;

    private String currentPath;

    public FolderNavigationView(AlertHandler alertHandler, LogManager logManager,
                                Consumer<String> openFolderCallback) {
        this.alertHandler = alertHandler;
        this.logManager = logManager;
        this.openFolderCallback = openFolderCallback;
    }

    public void setCurrentPath(String path) {
        this.currentPath = path;
    }

    public void openOutputFolder() {
        openFolder("📂 Открыта папка с результатами: ", "с результатами");
    }

    public void openSourceFolder() {
        openFolder("📁 Открыта исходная папка: ", "исходная");
    }

    public void openRtListFolder() {
        if (currentPath == null || currentPath.isEmpty()) {
            alertHandler.showWarning("Внимание", "Сначала выберите папку с файлами");
            return;
        }

        File rootDir = new File(currentPath);
        String folderName = rootDir.getName() + "_Перечень РТ";
        File rtListFolder = new File(rootDir, folderName);

        if (rtListFolder.exists()) {
            openFolderCallback.accept(rtListFolder.getAbsolutePath());
            logManager.logInfo("📊 Открыта папка с перечнем РТ: " + rtListFolder.getAbsolutePath());
        } else {
            alertHandler.showInfo("Информация",
                    "Папка с перечнем РТ еще не создана. Запустите обработку с включенной опцией 'Создать таблицу Перечень расчетных точек'");
        }
    }

    private void openFolder(String logMessage, String folderType) {
        if (currentPath == null || currentPath.isEmpty()) {
            alertHandler.showWarning("Внимание", "Сначала выберите папку с файлами");
            return;
        }

        openFolderCallback.accept(currentPath);
        logManager.logInfo(logMessage + currentPath);
    }
}