package com.tsb.noise.service;

import java.io.File;

/**
 * Основной сервис для обработки Excel файлов
 */
public interface ExcelService {

    /**
     * Обработка Excel файла
     */
    boolean processExcelFile(File inputFile, File outputFile);

    /**
     * Проверка поддержки файла
     */
    boolean supportsFile(File file);

    /**
     * Получение информации о сервисе
     */
    String getServiceInfo();
}