module com.tsb.noise {
    requires javafx.controls;
    requires javafx.fxml;
    requires org.kordamp.bootstrapfx.core;
    requires org.controlsfx.controls;
    requires static lombok;
    requires org.apache.poi.ooxml;
    requires org.apache.poi.poi;
    requires java.prefs;
    requires org.slf4j;

    // Открываем пакеты для JavaFX FXML
    opens com.tsb.noise to javafx.fxml;
    opens com.tsb.noise.controller to javafx.fxml;

    // Открываем model для JavaFX binding
    opens com.tsb.noise.model to javafx.base;

    // Открываем service пакеты для рефлексии
    opens com.tsb.noise.service to javafx.base;
    opens com.tsb.noise.service.utils to javafx.base;
    opens com.tsb.noise.service.processors to javafx.base;
    opens com.tsb.noise.service.operations to javafx.base;
    opens com.tsb.noise.service.operations.core to javafx.base;
    opens com.tsb.noise.service.operations.row to javafx.base;
    opens com.tsb.noise.service.operations.table to javafx.base;
    opens com.tsb.noise.service.operations.export to javafx.base;

    // Экспортируем публичные API
    exports com.tsb.noise;
    exports com.tsb.noise.controller;
    exports com.tsb.noise.model;
    exports com.tsb.noise.service;
    exports com.tsb.noise.service.utils;
    exports com.tsb.noise.service.processors;
    exports com.tsb.noise.service.operations;
    exports com.tsb.noise.service.operations.core;
    exports com.tsb.noise.service.operations.row;
    exports com.tsb.noise.service.operations.table;
    exports com.tsb.noise.service.operations.export;
}