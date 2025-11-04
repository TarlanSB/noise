package com.tsb.noise.service.operations.row;

import com.tsb.noise.service.operations.core.RowOperation;
import com.tsb.noise.service.operations.core.StyleApplier;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.List;

@Slf4j
public class CorrectionOperation implements RowOperation {

    private static final String[] TARGET_TEXTS = {"превышение", "превышение пом."};
    private static final int TARGET_COLUMN = 1; // Колонка B
    private static final String CORRECTION_TEXT = "Поправка на существующее/перспективное положение";

    private final double correctionValue;
    private final StyleApplier styleApplier;

    public CorrectionOperation(double correctionValue, StyleApplier styleApplier) {
        this.correctionValue = correctionValue;
        this.styleApplier = styleApplier;
    }

    @Override
    public int execute(Sheet sheet) {
        log.info("🔍 Поиск строк с 'превышение' для применения поправки: {}", correctionValue);

        List<Integer> targetRows = findTargetRows(sheet);

        if (targetRows.isEmpty()) {
            log.info("❌ Строки с 'превышение' не найдены");
            return 0;
        }

        log.info("✅ Найдено строк для применения поправки: {}", targetRows.size());

        // Сортируем по убыванию для корректной вставки
        targetRows.sort((a, b) -> b - a);
        int processedCount = 0;

        for (int targetRowIndex : targetRows) {
            try {
                if (applyCorrection(sheet, targetRowIndex)) {
                    processedCount++;
                }
            } catch (Exception e) {
                log.error("❌ Ошибка при применении поправки к строке {}: {}",
                        targetRowIndex + 1, e.getMessage(), e);
            }
        }

        log.info("🎯 Применена поправка к {} строкам", processedCount);
        return processedCount;
    }

    @Override
    public String getOperationName() {
        return String.format("Поправка на существующее/перспективное положение (%.2f)", correctionValue);
    }

    /**
     * Применяет поправку к целевой строке
     */
    private boolean applyCorrection(Sheet sheet, int targetRowIndex) {
        log.debug("🔄 Применение поправки к строке {}", targetRowIndex + 1);

        try {
            Row targetRow = sheet.getRow(targetRowIndex);
            if (targetRow == null) {
                log.warn("Целевая строка {} не найдена", targetRowIndex + 1);
                return false;
            }

            // ШАГ 1: Сохраняем исходные значения ячеек D-M
            List<CellData> originalValues = saveRowValues(targetRow, 3, 12); // Колонки D-M

            // ШАГ 2: Вставляем пустую строку ПЕРЕД целевой строкой
            int correctionRowIndex = insertEmptyRowBefore(sheet, targetRowIndex);

            // ШАГ 3: Заполняем новую строку поправкой
            fillCorrectionRow(sheet, correctionRowIndex);

            // ШАГ 4: Обновляем целевую строку (применяем поправку)
            updateTargetRowWithCorrection(targetRow, originalValues);

            log.debug("✅ Успешно применена поправка: вставлена строка {}, обновлена строка {}",
                    correctionRowIndex + 1, targetRowIndex + 1);
            return true;

        } catch (Exception e) {
            log.error("❌ Критическая ошибка при применении поправки: {}", e.getMessage(), e);
            return false;
        }
    }

    /**
     * Находит строки с целевым текстом (превышение или превышение пом.)
     */
    private List<Integer> findTargetRows(Sheet sheet) {
        List<Integer> targetRows = new ArrayList<>();

        for (int rowIndex = 3; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                Cell cell = row.getCell(TARGET_COLUMN);
                if (cell != null) {
                    String cellValue = getCellStringValue(cell).trim();
                    // Проверяем на соответствие любому из целевых текстов
                    for (String targetText : TARGET_TEXTS) {
                        if (targetText.equals(cellValue)) {
                            targetRows.add(rowIndex);
                            log.debug("Найдена строка с '{}' в строке {}", targetText, rowIndex + 1);
                            break;
                        }
                    }
                }
            }
        }

        return targetRows;
    }

    /**
     * Сохраняет значения ячеек в указанном диапазоне
     */
    private List<CellData> saveRowValues(Row row, int startCol, int endCol) {
        List<CellData> values = new ArrayList<>();
        for (int colIndex = startCol; colIndex <= endCol; colIndex++) {
            Cell cell = row.getCell(colIndex);
            if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                double value = cell.getNumericCellValue();
                values.add(new CellData(colIndex, value, cell.getCellStyle()));
            } else {
                // Если ячейка пустая или не числовая, считаем значение 0
                values.add(new CellData(colIndex, 0.0,
                        cell != null ? cell.getCellStyle() : null));
            }
        }
        return values;
    }

    /**
     * Вставляет пустую строку ПЕРЕД указанной строкой
     */
    private int insertEmptyRowBefore(Sheet sheet, int targetRowIndex) {
        // Сдвигаем строки вниз начиная с целевой позиции
        if (targetRowIndex <= sheet.getLastRowNum()) {
            sheet.shiftRows(targetRowIndex, sheet.getLastRowNum(), 1, true, false);
        }

        // Создаем новую строку
        Row newRow = sheet.createRow(targetRowIndex);
        newRow.setHeightInPoints((short) (8.0 / 25.4 * 72)); // 8mm в points

        return targetRowIndex;
    }

    /**
     * Заполняет строку поправки с правильными стилями
     */
    private void fillCorrectionRow(Sheet sheet, int rowIndex) {
        Row correctionRow = sheet.getRow(rowIndex);
        Workbook workbook = sheet.getWorkbook();

        // Заполняем ячейку B
        Cell cellB = correctionRow.createCell(1);
        cellB.setCellValue(CORRECTION_TEXT);
        styleApplier.applyCellStyleWithFont(cellB); // Используем существующий стиль

        // Заполняем ячейки D-M значением поправки
        for (int colIndex = 3; colIndex <= 12; colIndex++) {
            Cell cell = correctionRow.createCell(colIndex);
            cell.setCellValue(correctionValue);
            styleApplier.applyCellStyleWithFont(cell); // Используем существующий стиль
        }

        log.debug("Заполнена строка поправки: {} со значением {}", CORRECTION_TEXT, correctionValue);
    }

    /**
     * Обновляет целевую строку с применением поправки
     */
    private void updateTargetRowWithCorrection(Row targetRow, List<CellData> originalValues) {
        for (CellData cellData : originalValues) {
            Cell cell = targetRow.getCell(cellData.columnIndex);
            if (cell == null) {
                cell = targetRow.createCell(cellData.columnIndex);
            }

            // Новое значение = исходное значение + поправка
            double newValue = cellData.numericValue + correctionValue;
            cell.setCellValue(newValue);

            // Применяем стандартный стиль
            styleApplier.applyCellStyleWithFont(cell);

            log.debug("Обновлена ячейка {}{}: {} + {} = {}",
                    (char) ('A' + cellData.columnIndex),
                    targetRow.getRowNum() + 1,
                    cellData.numericValue, correctionValue, newValue);
        }

        // Также обновляем стиль ячейки B (текст "превышение")
        Cell cellB = targetRow.getCell(1);
        if (cellB != null) {
            styleApplier.applyCellStyleWithFont(cellB);
        }
    }

    /**
     * Вспомогательный класс для хранения данных ячейки
     */
    private static class CellData {
        private final int columnIndex;
        private final double numericValue;
        private final CellStyle style;

        public CellData(int columnIndex, double numericValue, CellStyle style) {
            this.columnIndex = columnIndex;
            this.numericValue = numericValue;
            this.style = style;
        }
    }

    /**
     * Вспомогательный метод для получения строкового значения ячейки
     */
    private String getCellStringValue(Cell cell) {
        if (cell == null) return "";
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return cell.getStringCellValue();
                } catch (Exception e) {
                    return cell.getCellFormula();
                }
            default:
                return "";
        }
    }
}