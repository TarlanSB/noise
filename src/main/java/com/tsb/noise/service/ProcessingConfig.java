package com.tsb.noise.service;

import lombok.Data;

/**
 * Конфигурация обработки файлов
 */
@Data
public class ProcessingConfig {
    private boolean removeSoundIsolation;
    private boolean moveBarrierIsolation;
    private boolean applyCorrection;
    private Double correctionValue;
    private boolean createRtList;

    public static ProcessingConfig defaultConfig() {
        return new ProcessingConfig();
    }
}