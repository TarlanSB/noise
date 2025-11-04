package com.tsb.noise.controller.components;

import javafx.scene.control.TextField;
import org.controlsfx.control.ToggleSwitch;

/**
 * Управляет настройками операций обработки
 */
public class OperationSettingsManager {

    private final ToggleSwitch removeSoundIsolationToggle;
    private final ToggleSwitch moveBarrierIsolationToggle;
    private final ToggleSwitch correctionToggle;
    private final ToggleSwitch createRtListToggle;
    private final ToggleSwitch createSummaryTableToggle;
    private final TextField correctionValueField;

    public OperationSettingsManager(
            ToggleSwitch removeSoundIsolationToggle,
            ToggleSwitch moveBarrierIsolationToggle,
            ToggleSwitch correctionToggle,
            ToggleSwitch createRtListToggle,
            ToggleSwitch createSummaryTableToggle,
            TextField correctionValueField) {

        this.removeSoundIsolationToggle = removeSoundIsolationToggle;
        this.moveBarrierIsolationToggle = moveBarrierIsolationToggle;
        this.correctionToggle = correctionToggle;
        this.createRtListToggle = createRtListToggle;
        this.createSummaryTableToggle = createSummaryTableToggle;
        this.correctionValueField = correctionValueField;

        setupOperationToggles();
    }

    private void setupOperationToggles() {
        // Настройка ToggleSwitch для удаления звукоизоляции
        removeSoundIsolationToggle.setText("🗑️ Удалить строки 'Требуемая звукоизоляция'");

        // Настройка ToggleSwitch для перемещения звукоизоляции преградой
        moveBarrierIsolationToggle.setText("🔄 Переместить 'Звукоизоляция преградой' на 3 строки выше");

        // Настройка ToggleSwitch для поправки
        correctionToggle.setText("📈 Поправка на существующее/перспективное положение");
        correctionToggle.selectedProperty().addListener((obs, oldVal, newVal) -> {
            correctionValueField.setDisable(!newVal);
            if (newVal && correctionValueField.getText().isEmpty()) {
                correctionValueField.setText("0");
            }
        });

        // Настройка ToggleSwitch для создания перечня РТ
        createRtListToggle.setText("📋 Создать таблицу 'Перечень расчетных точек'");

        // Настройка ToggleSwitch для создания сводной таблицы
        createSummaryTableToggle.setText("📊 Создать сводную таблицу РТ");

        // Валидация числового значения поправки
        correctionValueField.textProperty().addListener((obs, oldVal, newVal) -> {
            if (!newVal.matches("-?\\d*([\\.,]\\d*)?")) {
                correctionValueField.setText(oldVal);
            }
        });

        correctionValueField.setDisable(true);
    }

    public boolean isRemoveSoundIsolationEnabled() {
        return removeSoundIsolationToggle.isSelected();
    }

    public boolean isMoveBarrierIsolationEnabled() {
        return moveBarrierIsolationToggle.isSelected();
    }

    public boolean isCorrectionEnabled() {
        return correctionToggle.isSelected();
    }

    public boolean isCreateRtListEnabled() {
        return createRtListToggle.isSelected();
    }

    public boolean isCreateSummaryTableEnabled() {
        return createSummaryTableToggle.isSelected();
    }

    public Double getCorrectionValue() {
        if (!isCorrectionEnabled()) {
            return null;
        }
        try {
            String valueText = correctionValueField.getText().replace(",", ".");
            return Double.parseDouble(valueText);
        } catch (NumberFormatException e) {
            return null;
        }
    }

    public boolean isCorrectionValid() {
        return !isCorrectionEnabled() || getCorrectionValue() != null;
    }
}