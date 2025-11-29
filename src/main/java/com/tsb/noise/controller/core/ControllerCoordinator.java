package com.tsb.noise.controller.core;

import com.tsb.noise.controller.components.FileTypeSelectionManager;
import com.tsb.noise.controller.components.OperationSettingsManager;
import com.tsb.noise.controller.components.ProgressManager;
import com.tsb.noise.controller.handlers.AlertHandler;
import com.tsb.noise.controller.handlers.DirectorySelectionHandler;
import com.tsb.noise.controller.handlers.TaskBasedProcessingHandler;
import com.tsb.noise.controller.managers.LogManager;
import com.tsb.noise.controller.managers.StatusManager;
import com.tsb.noise.controller.views.FolderNavigationView;
import com.tsb.noise.controller.views.ProcessingView;
import com.tsb.noise.model.FileType;
import com.tsb.noise.service.utils.ExcelProcessor;
import com.tsb.noise.service.utils.PreferencesService;
import com.tsb.noise.service.operations.export.RtListCreator;
import com.tsb.noise.service.operations.export.SummaryTableCreator;
import javafx.scene.control.*;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import org.controlsfx.control.ToggleSwitch;

import java.util.List;
import java.util.function.Consumer;

/**
 * Координатор всех компонентов контроллера
 */
public class ControllerCoordinator {
    private final FileTypeSelectionManager fileTypeManager;
    private final OperationSettingsManager operationManager;
    private final DirectorySelectionHandler directoryHandler;
    private final TaskBasedProcessingHandler processingHandler;
    private final StatusManager statusManager;
    private final LogManager logManager;
    private final ProgressManager progressManager;
    private final AlertHandler alertHandler;
    private final FolderNavigationView folderNavigationView;
    private final ProcessingView processingView;

    public ControllerCoordinator(
            // UI Components
            HBox root, Label selectedPathLabel, Button selectPathButton,
            Button startProcessButton, Label processStatusLabel, ProgressBar progressBar,
            TextArea logArea, Label statusLabel, VBox progressContainer, Button openOutputButton,
            ToggleSwitch txDayToggle, ToggleSwitch txNightToggle, ToggleSwitch ovDayToggle,
            ToggleSwitch ovNightToggle, ToggleSwitch posDayToggle, ToggleSwitch posNightToggle,
            ToggleSwitch removeSoundIsolationToggle, ToggleSwitch moveBarrierIsolationToggle,
            ToggleSwitch correctionToggle, ToggleSwitch createRtListToggle, ToggleSwitch createSummaryTableToggle,
            TextField correctionValueField,
            Button selectAllButton, Button clearAllButton,
            // Services
            PreferencesService preferencesService, ExcelProcessor excelProcessor,
            RtListCreator rtListCreator, SummaryTableCreator summaryTableCreator,
            // Callbacks
            Runnable updateUIStateCallback, Consumer<String> openFolderCallback) {

        // Initialize services
        this.alertHandler = new AlertHandler();
        this.logManager = new LogManager(logArea);
        this.statusManager = new StatusManager(statusLabel, processStatusLabel);
        this.progressManager = new ProgressManager(progressBar, statusLabel, progressContainer);

        // Initialize component managers
        this.fileTypeManager = initializeFileTypeManager(txDayToggle, txNightToggle, ovDayToggle,
                ovNightToggle, posDayToggle, posNightToggle, selectAllButton, clearAllButton);

        this.operationManager = new OperationSettingsManager(removeSoundIsolationToggle,
                moveBarrierIsolationToggle, correctionToggle, createRtListToggle,
                createSummaryTableToggle, correctionValueField);

        this.processingHandler = new TaskBasedProcessingHandler(excelProcessor, rtListCreator,
                summaryTableCreator, logManager::logInfo, logManager::logError);

        this.directoryHandler = new DirectorySelectionHandler(preferencesService,
                root.getScene().getWindow(), selectedPathLabel, logManager::logInfo, updateUIStateCallback);

        // Initialize views
        this.folderNavigationView = new FolderNavigationView(alertHandler, logManager, openFolderCallback);
        this.processingView = new ProcessingView(processingHandler, progressManager, statusManager,
                alertHandler, logManager, summaryTableCreator,
                directoryHandler::getCurrentPath,
                fileTypeManager::getSelectedFileTypes,
                operationManager::isRemoveSoundIsolationEnabled,
                operationManager::isMoveBarrierIsolationEnabled,
                operationManager::getCorrectionValue,
                operationManager::isCreateRtListEnabled,
                operationManager::isCreateSummaryTableEnabled,
                updateUIStateCallback);


        // Setup UI
        setupUI(selectPathButton, startProcessButton, openOutputButton, updateUIStateCallback);
    }

    private FileTypeSelectionManager initializeFileTypeManager(ToggleSwitch txDayToggle, ToggleSwitch txNightToggle,
                                                               ToggleSwitch ovDayToggle, ToggleSwitch ovNightToggle,
                                                               ToggleSwitch posDayToggle, ToggleSwitch posNightToggle,
                                                               Button selectAllButton, Button clearAllButton) {
        // Create toggle map
        java.util.Map<ToggleSwitch, FileType> toggleMap = new java.util.HashMap<>();
        toggleMap.put(txDayToggle, FileType.TX_DAY);
        toggleMap.put(txNightToggle, FileType.TX_NIGHT);
        toggleMap.put(ovDayToggle, FileType.OV_DAY);
        toggleMap.put(ovNightToggle, FileType.OV_NIGHT);
        toggleMap.put(posDayToggle, FileType.POS_DAY);
        toggleMap.put(posNightToggle, FileType.POS_NIGHT);

        return new FileTypeSelectionManager(toggleMap, selectAllButton, clearAllButton);
    }

    private void setupUI(Button selectPathButton, Button startProcessButton, Button openOutputButton,
                         Runnable updateUIStateCallback) {
        fileTypeManager.initializeFileTypeToggles(updateUIStateCallback);

        selectPathButton.setOnAction(e -> directoryHandler.selectDirectory());
        startProcessButton.setOnAction(e -> processingView.startProcessing());
        openOutputButton.setOnAction(e -> folderNavigationView.openOutputFolder());
    }

    // Делегирующие методы
    public void loadLastSelectedPath() {
        directoryHandler.loadLastSelectedPath();
    }

    public void updateUIState() {
        String directoryPath = directoryHandler.getCurrentPath();
        List<FileType> selectedTypes = fileTypeManager.getSelectedFileTypes();
        boolean hasOperations = operationManager.isRemoveSoundIsolationEnabled() ||
                operationManager.isMoveBarrierIsolationEnabled() ||
                operationManager.isCorrectionEnabled() ||
                operationManager.isCreateRtListEnabled() ||
                operationManager.isCreateSummaryTableEnabled();

        statusManager.updateStatus(directoryPath,
                fileTypeManager.hasSelectedFileTypes(), selectedTypes, hasOperations);

        folderNavigationView.setCurrentPath(directoryPath);
    }

    public void updateProcessButtonState(Button startProcessButton) {
        boolean hasPath = directoryHandler.hasValidPath();
        boolean hasSelectedFileTypes = fileTypeManager.hasSelectedFileTypes();
        boolean correctionValid = operationManager.isCorrectionValid();

        startProcessButton.setDisable(!hasPath || !hasSelectedFileTypes || !correctionValid);
    }

    // Navigation methods
    public void openOutputFolder() {
        folderNavigationView.openOutputFolder();
    }

    public void openSourceFolder() {
        folderNavigationView.openSourceFolder();
    }

    public void openRtListFolder() {
        folderNavigationView.openRtListFolder();
    }

    public void clearLog() {
        logManager.clearLog();
    }

    public void handleException(String operation, Exception e) {
        if (logManager != null) {
            logManager.logError("❌ Ошибка при " + operation + ": " + e.getMessage());
        }
        alertHandler.showError("Ошибка", "Ошибка при " + operation + ": " + e.getMessage());
    }

    // UI event handlers
    public void handleSelectPath() {
        directoryHandler.selectDirectory();
    }

    public void handleStartProcess() {
        processingView.startProcessing();
    }

    public void handleOpenOutput() {
        folderNavigationView.openOutputFolder();
    }

    public void handleSelectAll() {
        fileTypeManager.selectAllFileTypes();
    }

    public void handleClearAll() {
        fileTypeManager.clearAllFileTypes();
    }
}