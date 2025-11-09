package com.tsb.noise.controller.views;

import com.tsb.noise.controller.components.ProgressManager;
import com.tsb.noise.controller.handlers.AlertHandler;
import com.tsb.noise.controller.handlers.TaskBasedProcessingHandler;
import com.tsb.noise.controller.managers.LogManager;
import com.tsb.noise.controller.managers.StatusManager;
import com.tsb.noise.model.FileType;
import javafx.concurrent.Task;

import java.util.List;
import java.util.function.Supplier;

/**
 * Управление процессом обработки файлов
 */
public class ProcessingView {
    private final TaskBasedProcessingHandler processingHandler;
    private final ProgressManager progressManager;
    private final StatusManager statusManager;
    private final AlertHandler alertHandler;
    private final LogManager logManager;

    private final Supplier<String> directoryPathSupplier;
    private final Supplier<List<FileType>> selectedFileTypesSupplier;
    private final Supplier<Boolean> removeSoundIsolationSupplier;
    private final Supplier<Boolean> moveBarrierIsolationSupplier;
    private final Supplier<Double> correctionValueSupplier;
    private final Supplier<Boolean> createRtListSupplier;
    private final Supplier<Boolean> createSummaryTableSupplier;
    private final Runnable updateUIStateCallback;

    public ProcessingView(TaskBasedProcessingHandler processingHandler,
                          ProgressManager progressManager,
                          StatusManager statusManager,
                          AlertHandler alertHandler,
                          LogManager logManager,
                          Supplier<String> directoryPathSupplier,
                          Supplier<List<FileType>> selectedFileTypesSupplier,
                          Supplier<Boolean> removeSoundIsolationSupplier,
                          Supplier<Boolean> moveBarrierIsolationSupplier,
                          Supplier<Double> correctionValueSupplier,
                          Supplier<Boolean> createRtListSupplier,
                          Supplier<Boolean> createSummaryTableSupplier,
                          Runnable updateUIStateCallback) {
        this.processingHandler = processingHandler;
        this.progressManager = progressManager;
        this.statusManager = statusManager;
        this.alertHandler = alertHandler;
        this.logManager = logManager;
        this.directoryPathSupplier = directoryPathSupplier;
        this.selectedFileTypesSupplier = selectedFileTypesSupplier;
        this.removeSoundIsolationSupplier = removeSoundIsolationSupplier;
        this.moveBarrierIsolationSupplier = moveBarrierIsolationSupplier;
        this.correctionValueSupplier = correctionValueSupplier;
        this.createRtListSupplier = createRtListSupplier;
        this.createSummaryTableSupplier = createSummaryTableSupplier;
        this.updateUIStateCallback = updateUIStateCallback;
    }

    public void startProcessing() {
        if (!validateProcessingPreconditions()) {
            return;
        }

        String directoryPath = directoryPathSupplier.get();
        List<FileType> selectedFileTypes = selectedFileTypesSupplier.get();
        Double correctionValue = correctionValueSupplier.get();

        Task<Void> processingTask = processingHandler.createProcessingTask(
                directoryPath, selectedFileTypes,
                removeSoundIsolationSupplier.get(),
                moveBarrierIsolationSupplier.get(),
                correctionValue,
                createRtListSupplier.get(),
                createSummaryTableSupplier.get(),
                progressManager::updateProgressMessage,
                progressManager::updateProgressValue);

        setupTaskHandlers(processingTask);
        processingHandler.setCurrentTask(processingTask);

        startTaskExecution(processingTask);
    }

    private boolean validateProcessingPreconditions() {
        String directoryPath = directoryPathSupplier.get();
        if (directoryPath == null || directoryPath.isEmpty()) {
            alertHandler.showWarning("Внимание", "Сначала выберите папку с файлами УЗД");
            return false;
        }

        List<FileType> selectedFileTypes = selectedFileTypesSupplier.get();
        if (selectedFileTypes == null || selectedFileTypes.isEmpty()) {
            alertHandler.showWarning("Внимание", "Выберите хотя бы один тип файлов для обработки");
            return false;
        }

        // ИСПРАВЛЕННАЯ ЛОГИКА: проверяем поправку только если она нужна для операций
        Double correctionValue = correctionValueSupplier.get();
        boolean needsCorrection = removeSoundIsolationSupplier.get() ||
                moveBarrierIsolationSupplier.get();

        if (needsCorrection && correctionValue == null) {
            alertHandler.showError("Ошибка", "Для выбранных операций требуется значение поправки. Введите числовое значение.");
            return false;
        }

        // Дополнительная проверка: если включена поправка, но не выбраны файлы для обработки
        if (needsCorrection && selectedFileTypes.isEmpty()) {
            alertHandler.showError("Ошибка", "Для применения поправки необходимо выбрать типы файлов для обработки.");
            return false;
        }

        return true;
    }

    private void setupTaskHandlers(Task<Void> task) {
        progressManager.setupTaskHandlers(task);

        task.setOnSucceeded(e -> handleTaskCompletion("✅ Обработка завершена"));
        task.setOnFailed(e -> handleTaskFailure());
        task.setOnCancelled(e -> handleTaskCompletion("⏹️ Обработка отменена"));
    }

    private void startTaskExecution(Task<Void> task) {
        progressManager.showProgress();
        new Thread(task).start();
    }

    private void handleTaskCompletion(String message) {
        progressManager.hideProgress();
        statusManager.updateProgressStatus(message);
        updateUIStateCallback.run();
        logManager.logInfo("🏁 Все операции завершены");
    }

    private void handleTaskFailure() {
        progressManager.hideProgress();
        statusManager.updateProgressStatus("❌ Ошибка обработки");
        updateUIStateCallback.run();

        if (processingHandler.isProcessing()) {
            String errorMessage = "Неизвестная ошибка";
            logManager.logError("💥 Критическая ошибка: " + errorMessage);
            alertHandler.showError("Ошибка", "Произошла ошибка при обработке файлов: " + errorMessage);
        }
    }
}