package com.tsb.noise.service.processors;

import com.tsb.noise.model.ProcessConfig;
import com.tsb.noise.service.operations.*;
import com.tsb.noise.service.operations.core.StyleApplier;
import com.tsb.noise.service.operations.table.ColumnHider;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

@Slf4j
public abstract class BaseExcelProcessor {
    protected final RowMover rowMover;
    protected final RowRemover rowRemover;
    protected final ColumnHider columnHider;
    protected final StyleApplier styleApplier;

    // Константы для размеров
    protected static final double ROW_HEIGHT_MM = 8.0;
    protected static final double COLUMN_WIDTH_B_CM = 3.0;
    protected static final double COLUMN_WIDTH_OTHER_CM = 1.5;
    protected static final short FONT_HEIGHT = 10;
    protected static final String FONT_NAME = "Arial Narrow";

    public BaseExcelProcessor() {
        this.rowMover = new RowMover();
        this.rowRemover = new RowRemover();
        this.columnHider = new ColumnHider();
        this.styleApplier = new StyleApplier();
    }

    /**
     * Шаблонный метод обработки - общий для всех типов файлов
     */
    public boolean process(File inputFile, File outputFile, ProcessConfig config) {
        log.info("Обработка файла: {} (тип: {})", inputFile.getName(), config.getFileType());

        try (FileInputStream fis = new FileInputStream(inputFile);
             Workbook sourceWorkbook = WorkbookFactory.create(fis);
             Workbook outputWorkbook = new XSSFWorkbook()) {

            Sheet sourceSheet = getSourceSheet(sourceWorkbook);
            if (sourceSheet == null) {
                log.error("Исходный лист не найден в файле: {}", inputFile.getName());
                return false;
            }

            Sheet outputSheet = outputWorkbook.createSheet("Данные");

            // Общие шаги обработки
            setupSheetLayout(outputSheet);
            createTableHeader(outputSheet);
            createEmptyRowAfterHeader(outputSheet);
            copyDataFromSource(sourceSheet, outputSheet);
            processSpecificData(sourceSheet, outputSheet, config);

            // Операции по конфигурации
            applyConfigurationOperations(outputSheet, config);

            // Финальные шаги
            columnHider.hideColumnC(outputSheet);
            removeEmptyRows(outputSheet);
            styleApplier.applyTableBorders(outputSheet);

            // Сохранение
            outputFile.getParentFile().mkdirs();
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                outputWorkbook.write(fos);
            }

            log.info("Файл успешно создан: {}", outputFile.getAbsolutePath());
            return true;

        } catch (IOException e) {
            log.error("Ошибка при обработке файла {}: {}", inputFile.getName(), e.getMessage(), e);
            return false;
        }
    }

    /**
     * Абстрактные методы для реализации в конкретных процессорах
     */
    protected abstract Sheet getSourceSheet(Workbook workbook);
    protected abstract void processSpecificData(Sheet sourceSheet, Sheet targetSheet, ProcessConfig config);

    /**
     * Общие методы которые могут быть переопределены при необходимости
     */
    protected void applyConfigurationOperations(Sheet sheet, ProcessConfig config) {
        if (config.isRemoveSoundIsolation()) {
            rowRemover.removeSoundIsolationRows(sheet);
        }
        if (config.isMoveSoundIsolation()) {
            rowMover.moveSoundIsolationBarrierRows(sheet);
        }
    }

    // Общие методы для всех процессоров
    protected void setupSheetLayout(Sheet sheet) {
        short rowHeightInPoints = mmToPoints(ROW_HEIGHT_MM);
        sheet.setDefaultRowHeight(rowHeightInPoints);

        sheet.setColumnWidth(0, cmToUnits(1.5));
        sheet.setColumnWidth(1, cmToUnits(COLUMN_WIDTH_B_CM));

        for (int i = 2; i <= 13; i++) {
            sheet.setColumnWidth(i, cmToUnits(COLUMN_WIDTH_OTHER_CM));
        }
    }

    protected void createTableHeader(Sheet sheet) {
        Row headerRow1 = createRowWithFixedHeight(sheet, 0);
        Row headerRow2 = createRowWithFixedHeight(sheet, 1);

        // Реализация создания шапки...
        // (перенесена из старого ExcelProcessor)
    }

    protected void createEmptyRowAfterHeader(Sheet sheet) {
        createRowWithFixedHeight(sheet, 2);
    }

    protected void copyDataFromSource(Sheet sourceSheet, Sheet targetSheet) {
        // Реализация копирования данных...
    }

    protected void removeEmptyRows(Sheet sheet) {
        // Реализация удаления пустых строк...
    }

    // Вспомогательные методы
    protected short mmToPoints(double mm) {
        return (short) (mm / 25.4 * 72);
    }

    protected int cmToUnits(double cm) {
        return (int) (cm * 4.5 * 256);
    }

    protected Row createRowWithFixedHeight(Sheet sheet, int rowIndex) {
        Row row = sheet.createRow(rowIndex);
        row.setHeightInPoints(mmToPoints(ROW_HEIGHT_MM));
        return row;
    }
}