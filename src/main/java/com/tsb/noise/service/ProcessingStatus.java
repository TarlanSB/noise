package com.tsb.noise.service;

/**
 * Статус обработки файлов
 */
public enum ProcessingStatus {
    IDLE,
    SCANNING,
    PROCESSING,
    COMPLETED,
    FAILED,
    CANCELLED
}