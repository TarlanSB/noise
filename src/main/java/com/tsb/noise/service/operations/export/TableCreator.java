package com.tsb.noise.service.operations.export;

/**
 * Функциональный интерфейс для создания сводных таблиц
 */
@FunctionalInterface
interface TableCreator {
    boolean createTable(String rootPath, boolean createTable);
}