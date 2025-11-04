package com.tsb.noise.controller.handlers;

import com.tsb.noise.service.utils.ExcelProcessor;
import com.tsb.noise.service.utils.FileUtils;
import com.tsb.noise.service.operations.export.RtListCreator;
import com.tsb.noise.service.operations.export.SummaryTableCreator;
import com.tsb.noise.model.FileType;
import javafx.concurrent.Task;
import lombok.Setter;

import java.io.File;
import java.util.List;
import java.util.function.Consumer;

/**
 * Обработчик запуска и управления процессом обработки файлов
 * Использует callback'и для полного разделения ответственности
 */
public class TaskBasedProcessingHandler {

    private final ExcelProcessor excelProcessor;
    private final RtListCreator rtListCreator;
    private final SummaryTableCreator summaryTableCreator;
    private final Consumer<String> logInfoCallback;
    private final Consumer<String> logErrorCallback;

    @Setter
    private Task<Void> currentTask;

    public TaskBasedProcessingHandler(
            ExcelProcessor excelProcessor,
            RtListCreator rtListCreator,
            SummaryTableCreator summaryTableCreator,
            Consumer<String> logInfoCallback,
            Consumer<String> logErrorCallback) {

        this.excelProcessor = excelProcessor;
        this.rtListCreator = rtListCreator;
        this.summaryTableCreator = summaryTableCreator;
        this.logInfoCallback = logInfoCallback;
        this.logErrorCallback = logErrorCallback;
    }

    /**
     * Создает задачу обработки файлов
     */
    public Task<Void> createProcessingTask(
            String directoryPath,
            List<FileType> selectedFileTypes,
            boolean removeSoundIsolation,
            boolean moveBarrierIsolation,
            Double correctionValue,
            boolean createRtList,
            boolean createSummaryTable,
            Consumer<String> progressMessageConsumer,
            Consumer<Double> progressValueConsumer) {

        return new Task<Void>() {
            @Override
            protected Void call() throws Exception {
                progressMessageConsumer.accept("Поиск файлов...");
                progressValueConsumer.accept(0.0);

                try {
                    List<File> targetFiles = findTargetFiles(directoryPath, selectedFileTypes);

                    if (targetFiles.isEmpty()) {
                        progressMessageConsumer.accept("❌ Файлы не найдены");
                        return null;
                    }

                    logInfoCallback.accept("📊 Начинается обработка выбранных файлов:");
                    targetFiles.forEach(file -> {
                        String fileType = FileUtils.getFileTypeDisplayName(file.getName());
                        logInfoCallback.accept("   • " + fileType + ": " + file.getName());
                    });

                    progressMessageConsumer.accept("Найдено файлов: " + targetFiles.size());
                    progressValueConsumer.accept(10.0);

                    // Логируем включенные операции
                    logEnabledOperations(removeSoundIsolation, moveBarrierIsolation,
                            correctionValue, createRtList, createSummaryTable);

                    processFiles(targetFiles, removeSoundIsolation, moveBarrierIsolation,
                            correctionValue, progressMessageConsumer, progressValueConsumer);

                    // Создание перечня РТ
                    if (createRtList && !isCancelled()) {
                        progressMessageConsumer.accept("Создание перечня расчетных точек...");
                        progressValueConsumer.accept(90.0);
                        createRtListTable(directoryPath);
                    }

                    // Создание сводной таблицы
                    if (createSummaryTable && !isCancelled()) {
                        progressMessageConsumer.accept("Создание сводной таблицы РТ...");
                        progressValueConsumer.accept(95.0);
                        createSummaryTable(directoryPath);
                    }

                    if (!isCancelled()) {
                        progressMessageConsumer.accept("✅ Обработка завершена");
                        progressValueConsumer.accept(100.0);
                    }

                } catch (Exception e) {
                    logErrorCallback.accept("💥 Критическая ошибка при обработке: " + e.getMessage());
                    throw e;
                }

                return null;
            }
        };
    }

    /**
     * Логирует включенные операции
     */
    private void logEnabledOperations(boolean removeSoundIsolation, boolean moveBarrierIsolation,
                                      Double correctionValue, boolean createRtList,
                                      boolean createSummaryTable) {
        if (removeSoundIsolation) {
            logInfoCallback.accept("🗑️ Режим удаления строк 'Требуемая звукоизоляция' активирован");
        }
        if (moveBarrierIsolation) {
            logInfoCallback.accept("🔄 Режим перемещения 'Звукоизоляция преградой' активирован");
        }
        if (correctionValue != null) {
            logInfoCallback.accept("📈 Режим поправки активирован: " + correctionValue);
        }
        if (createRtList) {
            logInfoCallback.accept("📋 Режим создания перечня РТ активирован");
        }
        if (createSummaryTable) {
            logInfoCallback.accept("📊 Режим создания сводной таблицы РТ активирован");
        }
    }

    /**
     * Находит файлы для обработки
     */
    private List<File> findTargetFiles(String directoryPath, List<FileType> selectedFileTypes) throws Exception {
        List<File> allFiles = FileUtils.findTargetExcelFiles(directoryPath);
        return allFiles.stream()
                .filter(file -> {
                    FileType fileType = FileType.fromFileName(file.getName());
                    return fileType != null && selectedFileTypes.contains(fileType);
                })
                .toList();
    }

    /**
     * Обрабатывает файлы
     */
    private void processFiles(List<File> targetFiles,
                              boolean removeSoundIsolation,
                              boolean moveBarrierIsolation,
                              Double correctionValue,
                              Consumer<String> progressMessageConsumer,
                              Consumer<Double> progressValueConsumer) {

        int processed = 0;
        int totalFiles = targetFiles.size();
        int successful = 0;

        for (File inputFile : targetFiles) {
            // Проверяем отмену задачи через currentTask
            if (currentTask != null && currentTask.isCancelled()) break;

            String fileType = FileUtils.getFileTypeDisplayName(inputFile.getName());
            progressMessageConsumer.accept("Обработка " + (processed + 1) + "/" + totalFiles + ": " + fileType);

            String outputFileName = FileUtils.generateOutputFileName(inputFile.getName());
            File outputFile = new File(inputFile.getParent(), outputFileName);

            boolean success = excelProcessor.processExcelFile(
                    inputFile, outputFile,
                    removeSoundIsolation, moveBarrierIsolation, correctionValue);

            if (success) {
                successful++;
                logInfoCallback.accept("✅ Успешно: " + fileType + " → " + outputFile.getName());
            } else {
                logErrorCallback.accept("❌ Ошибка: " + fileType + " → " + inputFile.getName());
            }

            processed++;
            double progress = 10 + (processed * 80.0 / totalFiles);
            progressValueConsumer.accept(progress);

            try {
                Thread.sleep(50);
            } catch (InterruptedException ie) {
                // Проверяем отмену при прерывании
                if (currentTask != null && currentTask.isCancelled()) break;
                Thread.currentThread().interrupt();
            }
        }

        String resultMessage = String.format("🎉 Обработка завершена! Успешно: %d, Ошибок: %d", successful, totalFiles - successful);
        logInfoCallback.accept(resultMessage);
    }

    /**
     * Создает перечень расчетных точек
     */
    private void createRtListTable(String directoryPath) {
        boolean rtListCreated = rtListCreator.createRtListTable(directoryPath, true);
        if (rtListCreated) {
            logInfoCallback.accept("✅ Успешно создан перечень расчетных точек");
        } else {
            logInfoCallback.accept("⚠️ Не удалось создать перечень расчетных точек");
        }
    }

    /**
     * Создает сводную таблицу РТ
     */
    private void createSummaryTable(String directoryPath) {
        boolean summaryTableCreated = summaryTableCreator.createSummaryTable(directoryPath, true);
        if (summaryTableCreated) {
            logInfoCallback.accept("✅ Успешно создана сводная таблица РТ");
        } else {
            logInfoCallback.accept("⚠️ Не удалось создать сводную таблицу РТ");
        }
    }

    public void cancelCurrentTask() {
        if (currentTask != null && currentTask.isRunning()) {
            currentTask.cancel();
        }
    }

    public boolean isProcessing() {
        return currentTask != null && currentTask.isRunning();
    }
}