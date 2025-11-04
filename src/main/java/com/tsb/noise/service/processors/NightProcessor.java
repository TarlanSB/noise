package com.tsb.noise.service.processors;

import com.tsb.noise.model.ProcessConfig;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

@Slf4j
public class NightProcessor extends BaseExcelProcessor {

    @Override
    protected Sheet getSourceSheet(Workbook workbook) {
        // Для ночных файлов может быть другой лист
        Sheet sheet = workbook.getSheet("ЛИСТ2");
        if (sheet == null) {
            sheet = workbook.getSheet("ЛИСТ3"); // Резервный вариант
        }
        return sheet;
    }

    @Override
    protected void processSpecificData(Sheet sourceSheet, Sheet targetSheet, ProcessConfig config) {
        log.info("Обработка ночных данных...");
        // Специфичная логика для ночных файлов
        // Например, другой обработчик РТ данных
    }
}