package com.tsb.noise.model;

import lombok.Getter;
import lombok.RequiredArgsConstructor;

@Getter
@RequiredArgsConstructor
public enum FileType {
    TX_DAY("УЗД в РТ ТХ день_FINAL", "В записку_УЗД в РТ ТХ день_FINAL", "ТХ день", "ЛИСТ2"),
    TX_NIGHT("УЗД в РТ ТХ ночь_FINAL", "В записку_УЗД в РТ ТХ ночь_FINAL", "ТХ ночь", "ЛИСТ2"),
    OV_DAY("УЗД в РТ ОВ день_FINAL", "В записку_УЗД в РТ ОВ день_FINAL", "ОВ день", "ЛИСТ2"),
    OV_NIGHT("УЗД в РТ ОВ ночь_FINAL", "В записку_УЗД в РТ ОВ ночь_FINAL", "ОВ ночь", "ЛИСТ2"),
    POS_DAY("УЗД в РТ ПОС день_FINAL", "В записку_УЗД в РТ ПОС день_FINAL", "ПОС день", "ЛИСТ2"),
    POS_NIGHT("УЗД в РТ ПОС ночь_FINAL", "В записку_УЗД в РТ ПОС ночь_FINAL", "ПОС ночь", "ЛИСТ2");

    private final String inputPattern;
    private final String outputPattern;
    private final String displayName;
    private final String sheetName;

    public static FileType fromFileName(String fileName) {
        for (FileType type : values()) {
            if (fileName.contains(type.getInputPattern())) {
                return type;
            }
        }
        return null;
    }

    public static boolean isSupportedFile(String fileName) {
        return fromFileName(fileName) != null;
    }

    /**
     * Проверяет, является ли файл ОВ (ОВ день или ОВ ночь)
     */
    public boolean isOvFile() {
        return this == OV_DAY || this == OV_NIGHT;
    }
}