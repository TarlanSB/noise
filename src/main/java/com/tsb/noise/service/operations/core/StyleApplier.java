package com.tsb.noise.service.operations.core;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;

@Slf4j
public class StyleApplier {

    private static final short FONT_HEIGHT = 10;
    private static final String FONT_NAME = "Arial Narrow";

    /**
     * Применяет тонкие границы ко всей таблице БЕЗ автопереноса
     */
    public void applyTableBorders(Sheet sheet) {
        log.info("🎨 Применение границ ко всей таблице...");

        Workbook workbook = sheet.getWorkbook();
        CellStyle borderStyle = createBorderStyle(workbook);

        int styledCells = 0;

        for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row != null) {
                for (int colIndex = 0; colIndex <= 12; colIndex++) { // A-M (после скрытия C)
                    if (colIndex == 2) continue; // Пропускаем скрытую колонку C
                    Cell cell = row.getCell(colIndex);
                    if (cell == null) {
                        // Создаем пустую ячейку с границами
                        cell = row.createCell(colIndex);
                        cell.setCellStyle(borderStyle);
                        styledCells++;
                    } else {
                        // Применяем стиль с границами к существующей ячейке
                        applyBordersToExistingCell(cell, borderStyle);
                        styledCells++;
                    }
                }
            }
        }

        log.info("✅ Применены тонкие границы к {} ячейкам БЕЗ автопереноса", styledCells);
    }

    /**
     * Создает стиль шапки с шрифтом Arial Narrow 10pt БЕЗ автопереноса
     */
    public void applyHeaderStyle(Workbook workbook, Row... headerRows) {
        log.debug("🎨 Применение стиля шапки...");

        CellStyle headerStyle = createBaseCellStyle(workbook);
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        for (Row row : headerRows) {
            for (Cell cell : row) {
                cell.setCellStyle(headerStyle);
            }
        }
    }

    /**
     * Применяет базовый стиль с шрифтом Arial Narrow 10pt к ячейке
     */
    public void applyCellStyleWithFont(Cell cell) {
        try {
            Workbook workbook = cell.getSheet().getWorkbook();
            CellStyle style = createBaseCellStyle(workbook);
            cell.setCellStyle(style);
        } catch (Exception e) {
            log.warn("⚠️ Не удалось применить стиль к ячейке: {}", e.getMessage());
        }
    }

    /**
     * Создает базовый стиль ячейки с Arial Narrow 10pt БЕЗ автопереноса
     */
    private CellStyle createBaseCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        // Выравнивание по центру
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        // УБРАН автоперенос текста
        style.setWrapText(false);

        // Шрифт Arial Narrow 10pt
        Font font = workbook.createFont();
        font.setFontName(FONT_NAME);
        font.setFontHeightInPoints(FONT_HEIGHT);
        style.setFont(font);

        return style;
    }

    /**
     * Создает стиль с тонкими границами
     */
    private CellStyle createBorderStyle(Workbook workbook) {
        CellStyle style = createBaseCellStyle(workbook);

        // Внешные и внутренние границы THIN
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        return style;
    }

    /**
     * Применяет границы к существующей ячейке сохраняя другие свойства
     */
    private void applyBordersToExistingCell(Cell cell, CellStyle borderStyle) {
        try {
            Workbook workbook = cell.getSheet().getWorkbook();
            CellStyle newStyle = workbook.createCellStyle();

            // Клонируем существующий стиль
            newStyle.cloneStyleFrom(cell.getCellStyle());

            // Применяем границы
            newStyle.setBorderTop(BorderStyle.THIN);
            newStyle.setBorderBottom(BorderStyle.THIN);
            newStyle.setBorderLeft(BorderStyle.THIN);
            newStyle.setBorderRight(BorderStyle.THIN);

            cell.setCellStyle(newStyle);
        } catch (Exception e) {
            log.warn("⚠️ Не удалось применить границы к ячейке: {}", e.getMessage());
            cell.setCellStyle(borderStyle);
        }
    }
}