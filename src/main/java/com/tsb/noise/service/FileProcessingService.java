package com.tsb.noise.service;

import java.io.File;
import java.util.List;

/**
 * Сервис для координации обработки файлов
 */
public interface FileProcessingService {

    /**
     * Обработка всех файлов в директории
     */
    ProcessingResult processDirectory(String directoryPath, ProcessingConfig config);

    /**
     * Поиск файлов для обработки
     */
    List<File> findFilesToProcess(String directoryPath);

    /**
     * Получение статуса обработки
     */
    ProcessingStatus getProcessingStatus();

    /**
     * Отмена обработки
     */
    void cancelProcessing();
}