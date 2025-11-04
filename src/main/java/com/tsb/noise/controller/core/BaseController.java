package com.tsb.noise.controller.core;

import javafx.fxml.Initializable;
import org.slf4j.Logger;

/**
 * Базовый класс для всех контроллеров с общими зависимостями
 */
public abstract class BaseController implements Initializable {
    protected Logger log;

    protected void initializeLogger(Class<?> clazz) {
        this.log = org.slf4j.LoggerFactory.getLogger(clazz);
    }

    protected void handleException(String operation, Exception e) {
        log.error("Ошибка при {}: {}", operation, e.getMessage(), e);
    }
}