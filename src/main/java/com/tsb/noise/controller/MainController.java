package com.tsb.noise.controller;

import com.tsb.noise.controller.core.BaseController;
import com.tsb.noise.controller.core.ControllerCoordinator;
import com.tsb.noise.model.FileType;
import com.tsb.noise.service.utils.ExcelProcessor;
import com.tsb.noise.service.utils.PreferencesService;
import com.tsb.noise.service.operations.export.RtListCreator;
import com.tsb.noise.service.operations.export.SummaryTableCreator;
import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import org.controlsfx.control.ToggleSwitch;

import java.io.File;
import java.net.URL;
import java.util.ResourceBundle;

public class MainController extends BaseController {

    // UI Components
    @FXML private HBox root;
    @FXML private Label selectedPathLabel;
    @FXML private Button selectPathButton;
    @FXML private Button startProcessButton;
    @FXML private Label processStatusLabel;
    @FXML private ProgressBar progressBar;
    @FXML private TextArea logArea;
    @FXML private Label statusLabel;
    @FXML private VBox progressContainer;
    @FXML private Button openOutputButton;

    // Toggle switches для типов файлов
    @FXML private ToggleSwitch txDayToggle;
    @FXML private ToggleSwitch txNightToggle;
    @FXML private ToggleSwitch ovDayToggle;
    @FXML private ToggleSwitch ovNightToggle;
    @FXML private ToggleSwitch posDayToggle;
    @FXML private ToggleSwitch posNightToggle;

    // Общие операции
    @FXML private ToggleSwitch removeSoundIsolationToggle;
    @FXML private ToggleSwitch moveBarrierIsolationToggle;
    @FXML private ToggleSwitch correctionToggle;
    @FXML private ToggleSwitch createRtListToggle;
    @FXML private ToggleSwitch createSummaryTableToggle;
    @FXML private TextField correctionValueField;

    // Кнопки управления выбором
    @FXML private Button selectAllButton;
    @FXML private Button clearAllButton;

    // Координатор
    private ControllerCoordinator coordinator;

    @Override
    public void initialize(URL location, ResourceBundle resources) {
        try {
            initializeLogger(MainController.class);

            // Инициализация сервисов
            PreferencesService preferencesService = new PreferencesService();
            ExcelProcessor excelProcessor = new ExcelProcessor();
            RtListCreator rtListCreator = new RtListCreator();
            SummaryTableCreator summaryTableCreator = new SummaryTableCreator();

            // Создание координатора с отложенной инициализацией
            initializeCoordinator(preferencesService, excelProcessor, rtListCreator, summaryTableCreator);

            logInfo("🚀 Приложение инициализировано успешно");

        } catch (Exception e) {
            handleException("инициализации приложения", e);
            showAlert("Ошибка", "Не удалось инициализировать приложение: " + e.getMessage());
        }
    }

    /**
     * Инициализация координатора с отложенной загрузкой сцены
     */
    private void initializeCoordinator(PreferencesService preferencesService,
                                       ExcelProcessor excelProcessor,
                                       RtListCreator rtListCreator,
                                       SummaryTableCreator summaryTableCreator) {
        // Ждем когда сцена станет доступна
        root.sceneProperty().addListener((observable, oldScene, newScene) -> {
            if (newScene != null && coordinator == null) {
                Platform.runLater(() -> {
                    try {
                        createCoordinator(preferencesService, excelProcessor, rtListCreator, summaryTableCreator);
                        coordinator.loadLastSelectedPath();
                        updateUIState();
                        logInfo("✅ Сцена загружена, приложение полностью инициализировано");
                    } catch (Exception e) {
                        handleException("создания координатора", e);
                    }
                });
            }
        });

        // Если сцена уже доступна
        if (root.getScene() != null && coordinator == null) {
            Platform.runLater(() -> {
                try {
                    createCoordinator(preferencesService, excelProcessor, rtListCreator, summaryTableCreator);
                    coordinator.loadLastSelectedPath();
                    updateUIState();
                    logInfo("✅ Сцена уже доступна, приложение полностью инициализировано");
                } catch (Exception e) {
                    handleException("создания координатора", e);
                }
            });
        }
    }

    /**
     * Создание координатора
     */
    private void createCoordinator(PreferencesService preferencesService,
                                   ExcelProcessor excelProcessor,
                                   RtListCreator rtListCreator,
                                   SummaryTableCreator summaryTableCreator) {
        this.coordinator = new ControllerCoordinator(
                // UI Components
                root, selectedPathLabel, selectPathButton, startProcessButton,
                processStatusLabel, progressBar, logArea, statusLabel, progressContainer,
                openOutputButton, txDayToggle, txNightToggle, ovDayToggle, ovNightToggle,
                posDayToggle, posNightToggle, removeSoundIsolationToggle,
                moveBarrierIsolationToggle, correctionToggle, createRtListToggle,
                createSummaryTableToggle, correctionValueField, selectAllButton, clearAllButton,
                // Services
                preferencesService, excelProcessor, rtListCreator, summaryTableCreator,
                // Callbacks
                this::updateUIState, this::openFolderInExplorer);
    }

    /**
     * Обновление состояния UI
     */
    private void updateUIState() {
        if (coordinator != null) {
            coordinator.updateUIState();
            coordinator.updateProcessButtonState(startProcessButton);
        } else {
            // Базовое состояние до инициализации координатора
            startProcessButton.setDisable(true);
            processStatusLabel.setText("Инициализация...");
        }
    }

    @FXML
    private void clearLog() {
        if (coordinator != null) {
            coordinator.clearLog();
        } else {
            logArea.clear();
            logInfo("🧹 Лог очищен");
        }
    }

    @FXML
    private void openOutputFolder() {
        if (coordinator != null) {
            coordinator.openOutputFolder();
        } else {
            showAlert("Ошибка", "Система не инициализирована");
        }
    }

    @FXML
    private void openSourceFolder() {
        if (coordinator != null) {
            coordinator.openSourceFolder();
        } else {
            showAlert("Ошибка", "Система не инициализирована");
        }
    }

    @FXML
    private void openRtListFolder() {
        if (coordinator != null) {
            coordinator.openRtListFolder();
        } else {
            showAlert("Ошибка", "Система не инициализирована");
        }
    }

    @FXML
    private void handleSelectAll() {
        if (coordinator != null) {
            coordinator.handleSelectAll();
        }
    }

    @FXML
    private void handleClearAll() {
        if (coordinator != null) {
            coordinator.handleClearAll();
        }
    }

    /**
     * Открывает папку в проводнике системы
     */
    private void openFolderInExplorer(String folderPath) {
        try {
            File folder = new File(folderPath);
            if (!folder.exists()) {
                showAlert("Ошибка", "Папка не существует: " + folderPath);
                return;
            }

            String os = System.getProperty("os.name").toLowerCase();
            ProcessBuilder processBuilder;

            if (os.contains("win")) {
                // Windows
                processBuilder = new ProcessBuilder("explorer.exe", folder.getAbsolutePath());
            } else if (os.contains("mac")) {
                // macOS
                processBuilder = new ProcessBuilder("open", folder.getAbsolutePath());
            } else {
                // Linux
                processBuilder = new ProcessBuilder("xdg-open", folder.getAbsolutePath());
            }

            processBuilder.start();
            logInfo("✅ Открыта папка в проводнике: " + folderPath);

        } catch (Exception e) {
            logError("❌ Не удалось открыть папку: " + e.getMessage());
            showAlert("Ошибка", "Не удалось открыть папку в проводнике: " + e.getMessage());
        }
    }

    /**
     * Показ alert-сообщения
     */
    private void showAlert(String title, String message) {
        Platform.runLater(() -> {
            Alert alert = new Alert(Alert.AlertType.INFORMATION);
            alert.setTitle(title);
            alert.setHeaderText(null);
            alert.setContentText(message);

            DialogPane dialogPane = alert.getDialogPane();
            dialogPane.getStyleClass().add("custom-alert");

            alert.showAndWait();
        });
    }

    /**
     * Логирование информации
     */
    private void logInfo(String message) {
        Platform.runLater(() -> {
            String timestamp = java.time.LocalDateTime.now().format(java.time.format.DateTimeFormatter.ofPattern("HH:mm:ss"));
            logArea.appendText("[" + timestamp + "] " + message + "\n");
            logArea.setScrollTop(Double.MAX_VALUE);
        });
    }

    /**
     * Логирование ошибки
     */
    private void logError(String message) {
        Platform.runLater(() -> {
            String timestamp = java.time.LocalDateTime.now().format(java.time.format.DateTimeFormatter.ofPattern("HH:mm:ss"));
            logArea.appendText("[" + timestamp + "] " + message + "\n");
            logArea.setScrollTop(Double.MAX_VALUE);
        });
    }

    /**
     * Обработка исключений с логированием
     */
    @Override
    protected void handleException(String operation, Exception e) {
        if (log != null) {
            log.error("Ошибка при {}: {}", operation, e.getMessage(), e);
        }
        logError("❌ Ошибка при " + operation + ": " + e.getMessage());
    }
}