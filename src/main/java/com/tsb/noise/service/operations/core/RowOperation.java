package com.tsb.noise.service.operations.core;

import org.apache.poi.ss.usermodel.Sheet;

/**
 * Интерфейс для операций над строками Excel
 */
public interface RowOperation {

    /**
     * Выполняет операцию над листом
     * @param sheet лист для обработки
     * @return количество измененных строк
     */
    int execute(Sheet sheet);

    /**
     * Возвращает название операции для логирования
     */
    String getOperationName();
}