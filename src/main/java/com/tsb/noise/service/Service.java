package com.tsb.noise.service;

/**
 * Базовый интерфейс для всех сервисов приложения
 */
public interface Service {

    /**
     * Инициализация сервиса
     */
    void initialize();

    /**
     * Очистка ресурсов сервиса
     */
    void shutdown();
}