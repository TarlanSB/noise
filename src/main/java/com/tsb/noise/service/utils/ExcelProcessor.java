package com.tsb.noise.service.utils;

import com.tsb.noise.model.FileType;
import com.tsb.noise.service.processors.OvDataProcessor;
import com.tsb.noise.service.operations.core.RowOperation;
import com.tsb.noise.service.operations.core.SheetLayoutManager;
import com.tsb.noise.service.operations.core.StyleApplier;
import com.tsb.noise.service.operations.row.BarrierRowMover;
import com.tsb.noise.service.operations.row.CorrectionOperation;
import com.tsb.noise.service.operations.row.EmptyRowCleaner;
import com.tsb.noise.service.operations.row.SoundIsolationRemover;
import com.tsb.noise.service.operations.table.ColumnHider;
import com.tsb.noise.service.operations.table.DataCopier;
import com.tsb.noise.service.operations.table.TableHeaderCreator;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

@Slf4j
public class ExcelProcessor {
    private final RtDataProcessor rtDataProcessor;
    private final OvDataProcessor ovDataProcessor;
    private final RowOperation soundIsolationRemover;
    private final RowOperation barrierRowMover;
    private final SheetLayoutManager layoutManager;
    private final TableHeaderCreator headerCreator;
    private final DataCopier dataCopier;
    private final EmptyRowCleaner emptyRowCleaner;
    private final StyleApplier styleApplier;
    private final ColumnHider columnHider;

    // Константы
    private static final double ROW_HEIGHT_MM = 8.0;

    public ExcelProcessor() {
        this.rtDataProcessor = new RtDataProcessor();
        this.ovDataProcessor = new OvDataProcessor();
        this.soundIsolationRemover = new SoundIsolationRemover();
        this.barrierRowMover = new BarrierRowMover();
        this.styleApplier = new StyleApplier();
        this.layoutManager = new SheetLayoutManager();
        this.headerCreator = new TableHeaderCreator(styleApplier);
        this.dataCopier = new DataCopier(styleApplier);
        this.emptyRowCleaner = new EmptyRowCleaner();
        this.columnHider = new ColumnHider();
    }

    /**
     * Основной метод с поддержкой всех операций для всех типов файлов
     */
    public boolean processExcelFile(File inputFile, File outputFile,
                                    boolean removeSoundIsolation,
                                    boolean moveSoundIsolation,
                                    Double correctionValue) {
        // Определяем тип файла
        FileType fileType = FileType.fromFileName(inputFile.getName());
        if (fileType == null) {
            log.error("Неподдерживаемый тип файла: {}", inputFile.getName());
            return false;
        }

        log.info("Начало обработки файла: {} (тип: {}, удаление: {}, перемещение: {}, поправка: {})",
                inputFile.getName(), fileType.getDisplayName(), removeSoundIsolation,
                moveSoundIsolation, correctionValue != null ? correctionValue : "нет");

        try (FileInputStream fis = new FileInputStream(inputFile);
             Workbook sourceWorkbook = WorkbookFactory.create(fis);
             Workbook outputWorkbook = new XSSFWorkbook()) {

            // Используем имя листа из типа файла
            Sheet sourceSheet = sourceWorkbook.getSheet(fileType.getSheetName());
            if (sourceSheet == null) {
                log.error("Лист '{}' не найден в файле: {}", fileType.getSheetName(), inputFile.getName());
                return false;
            }

            Sheet outputSheet = outputWorkbook.createSheet("Данные");

            // Настраиваем размеры и стили
            layoutManager.setupSheetLayout(outputSheet);
            headerCreator.createTableHeader(outputSheet, ROW_HEIGHT_MM);
            createEmptyRowAfterHeader(outputSheet);

            // Копируем данные
            dataCopier.copyDataFromSource(sourceSheet, outputSheet, ROW_HEIGHT_MM);

            // Обрабатываем данные РТ (для всех типов файлов)
            log.info("Начинаем обработку данных РТ для {}...", fileType.getDisplayName());
            rtDataProcessor.processRtData(sourceSheet, outputSheet);

            // СПЕЦИАЛЬНАЯ ЛОГИКА ДЛЯ ФАЙЛОВ ОВ - добавляем поправку к ПДУ
            if (isOvFileType(fileType)) {
                log.info("🔧 Применение специальной логики для файлов ОВ...");
                ovDataProcessor.processOvData(outputSheet);
            }

            // ВЫПОЛНЯЕМ ОПЕРАЦИИ (общие для всех типов файлов)
            if (removeSoundIsolation) {
                log.info("🚀 Выполнение операции удаления для {}", fileType.getDisplayName());
                int removedCount = soundIsolationRemover.execute(outputSheet);
                log.info("✅ Удаление для {}: обработано {} строк", fileType.getDisplayName(), removedCount);
            }

            if (moveSoundIsolation) {
                log.info("🚀 Выполнение операции перемещения для {}", fileType.getDisplayName());
                int movedCount = barrierRowMover.execute(outputSheet);
                log.info("✅ Перемещение для {}: обработано {} строк", fileType.getDisplayName(), movedCount);
            }

            // Применяем поправку если указана (общая для всех типов файлов)
            if (correctionValue != null) {
                RowOperation correctionOperation = new CorrectionOperation(correctionValue, styleApplier);
                log.info("🚀 Выполнение операции поправки для {}", fileType.getDisplayName());
                int correctedCount = correctionOperation.execute(outputSheet);
                log.info("✅ Поправка для {}: обработано {} строк", fileType.getDisplayName(), correctedCount);
            }

            // Финальные операции (общие для всех типов файлов)
            columnHider.hideColumnC(outputSheet);
            emptyRowCleaner.removeEmptyRows(outputSheet);
            styleApplier.applyTableBorders(outputSheet);

            // Сохраняем файл
            outputFile.getParentFile().mkdirs();
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                outputWorkbook.write(fos);
            }

            log.info("Файл {} успешно создан: {}", fileType.getDisplayName(), outputFile.getAbsolutePath());
            return true;

        } catch (IOException e) {
            log.error("Ошибка при обработке файла {}: {}", inputFile.getName(), e.getMessage(), e);
            return false;
        }
    }

    /**
     * Проверяет, является ли тип файла ОВ (Отопление и Вентиляция)
     */
    private boolean isOvFileType(FileType fileType) {
        return fileType == FileType.OV_DAY || fileType == FileType.OV_NIGHT;
    }

    /**
     * Старый метод для обратной совместимости
     */
    public boolean processExcelFile(File inputFile, File outputFile) {
        return processExcelFile(inputFile, outputFile, false, false, null);
    }

    /**
     * Перегруженный метод для обратной совместимости
     */
    public boolean processExcelFile(File inputFile, File outputFile,
                                    boolean removeSoundIsolation,
                                    boolean moveSoundIsolation) {
        return processExcelFile(inputFile, outputFile, removeSoundIsolation, moveSoundIsolation, null);
    }

    /**
     * Создает пустую строку сразу после шапки таблицы
     */
    private void createEmptyRowAfterHeader(Sheet sheet) {
        Row row = sheet.createRow(2);
        row.setHeightInPoints((short) (ROW_HEIGHT_MM / 25.4 * 72));
        log.debug("Создана пустая строка после шапки таблицы (строка 2)");
    }
}