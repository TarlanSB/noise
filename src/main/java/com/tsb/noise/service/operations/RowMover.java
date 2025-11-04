package com.tsb.noise.service.operations;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

@Slf4j
public class RowMover {

    /**
     * CORRECTED: Правильное перемещение строк с вставкой пустой строки
     */
    public void moveSoundIsolationBarrierRows(Sheet sheet) {
        log.info("🔍 Поиск строк с 'Звукоизоляция преградой' для перемещения...");

        List<BarrierRowInfo> barrierRows = findSoundIsolationBarrierRows(sheet);

        if (barrierRows.isEmpty()) {
            log.info("❌ Строки с 'Звукоизоляция преградой' не найдены");
            return;
        }

        log.info("✅ Найдено строк: {}", barrierRows.size());
        barrierRows.sort(Comparator.comparingInt(BarrierRowInfo::getOriginalIndex).reversed());

        int movedCount = 0;

        for (BarrierRowInfo barrierRow : barrierRows) {
            try {
                if (moveBarrierRowWithInsert(sheet, barrierRow)) {
                    movedCount++;
                }
            } catch (Exception e) {
                log.error("❌ Ошибка при перемещении строки {}: {}",
                        barrierRow.getOriginalIndex() + 1, e.getMessage(), e);
            }
        }

        log.info("🎯 Перемещено строк: {}", movedCount);
    }

    /**
     * CORRECTED: Метод с вставкой пустой строки вместо замены
     */
    private boolean moveBarrierRowWithInsert(Sheet sheet, BarrierRowInfo barrierRow) {
        int sourceIndex = barrierRow.getOriginalIndex();
        int targetIndex = barrierRow.getTargetIndex();

        log.debug("🔄 Перемещение из {} в {}", sourceIndex + 1, targetIndex + 1);

        try {
            // CORRECTED: Вставляем пустую строку в целевую позицию
            insertEmptyRow(sheet, targetIndex);

            // Копируем данные в новую пустую строку
            Row newRow = sheet.getRow(targetIndex);
            copyCompleteRowData(barrierRow.getRow(), newRow);

            // Удаляем исходную строку
            removeRow(sheet, sourceIndex + 1); // +1 потому что вставили строку выше

            log.debug("✅ Успешно перемещено с вставкой пустой строки");
            return true;

        } catch (Exception e) {
            log.error("❌ Ошибка перемещения: {}", e.getMessage(), e);
            return false;
        }
    }

    /**
     * CORRECTED: Вставка пустой строки без потери данных
     */
    private void insertEmptyRow(Sheet sheet, int targetIndex) {
        // Сдвигаем строки вниз начиная с целевой позиции
        if (targetIndex <= sheet.getLastRowNum()) {
            sheet.shiftRows(targetIndex, sheet.getLastRowNum(), 1, true, false);
        }

        // Создаем пустую строку
        Row newRow = sheet.createRow(targetIndex);
        newRow.setHeightInPoints((short) (8.0 / 25.4 * 72)); // 8mm в points

        // Заполняем пустыми ячейками
        for (int colIndex = 0; colIndex <= 13; colIndex++) {
            Cell cell = newRow.createCell(colIndex);
            cell.setCellValue("");
        }
    }

    private List<BarrierRowInfo> findSoundIsolationBarrierRows(Sheet sheet) {
        List<BarrierRowInfo> barrierRows = new ArrayList<>();

        for (int rowIndex = 3; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                Cell cellB = row.getCell(1);
                if (cellB != null) {
                    String cellValue = getCellStringValue(cellB).trim();
                    if ("Звукоизоляция преградой".equals(cellValue)) {
                        int targetIndex = rowIndex - 3;
                        if (targetIndex >= 3) {
                            barrierRows.add(new BarrierRowInfo(rowIndex, targetIndex, cellValue, row));
                        }
                    }
                }
            }
        }
        return barrierRows;
    }

    private void copyCompleteRowData(Row sourceRow, Row targetRow) {
        for (int colIndex = 0; colIndex <= 13; colIndex++) {
            Cell sourceCell = sourceRow.getCell(colIndex);
            if (sourceCell != null) {
                Cell targetCell = targetRow.createCell(colIndex);
                copyCellWithStyle(sourceCell, targetCell);
            }
        }
        targetRow.setHeight(sourceRow.getHeight());
    }

    private void copyCellWithStyle(Cell sourceCell, Cell targetCell) {
        // Копирование значения и стиля...
    }

    private void removeRow(Sheet sheet, int rowIndex) {
        if (rowIndex >= 0 && rowIndex <= sheet.getLastRowNum()) {
            if (rowIndex < sheet.getLastRowNum()) {
                sheet.shiftRows(rowIndex + 1, sheet.getLastRowNum(), -1);
            } else {
                Row row = sheet.getRow(rowIndex);
                if (row != null) {
                    sheet.removeRow(row);
                }
            }
        }
    }

    private String getCellStringValue(Cell cell) {
        if (cell == null) return "";
        // Реализация получения строкового значения...
        return "";
    }

    private static class BarrierRowInfo {
        private final int originalIndex;
        private final int targetIndex;
        private final String cellValue;
        private final Row row;

        public BarrierRowInfo(int originalIndex, int targetIndex, String cellValue, Row row) {
            this.originalIndex = originalIndex;
            this.targetIndex = targetIndex;
            this.cellValue = cellValue;
            this.row = row;
        }

        public int getOriginalIndex() { return originalIndex; }
        public int getTargetIndex() { return targetIndex; }
        public String getCellValue() { return cellValue; }
        public Row getRow() { return row; }
    }
}