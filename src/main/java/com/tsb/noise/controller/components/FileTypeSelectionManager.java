package com.tsb.noise.controller.components;

import com.tsb.noise.model.FileType;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import org.controlsfx.control.ToggleSwitch;

import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * Управляет выбором типов файлов через ToggleSwitch
 */
public class FileTypeSelectionManager {

    private final Map<ToggleSwitch, FileType> fileTypeToggleMap = new HashMap<>();
    private final Button selectAllButton;
    private final Button clearAllButton;

    public FileTypeSelectionManager(
            Map<ToggleSwitch, FileType> toggleMap,
            Button selectAllButton,
            Button clearAllButton) {
        this.fileTypeToggleMap.putAll(toggleMap);
        this.selectAllButton = selectAllButton;
        this.clearAllButton = clearAllButton;
        setupSelectionButtons();
    }

    public void initializeFileTypeToggles(Runnable onToggleChange) {
        for (Map.Entry<ToggleSwitch, FileType> entry : fileTypeToggleMap.entrySet()) {
            ToggleSwitch toggle = entry.getKey();
            FileType fileType = entry.getValue();

            toggle.setText(fileType.getDisplayName());
            toggle.selectedProperty().addListener((obs, oldVal, newVal) -> onToggleChange.run());
            toggle.setSelected(true); // По умолчанию включены
        }
    }

    private void setupSelectionButtons() {
        selectAllButton.setOnAction(e -> selectAllFileTypes());
        clearAllButton.setOnAction(e -> clearAllFileTypes());
    }

    public void selectAllFileTypes() {
        for (ToggleSwitch toggle : fileTypeToggleMap.keySet()) {
            toggle.setSelected(true);
        }
    }

    public void clearAllFileTypes() {
        for (ToggleSwitch toggle : fileTypeToggleMap.keySet()) {
            toggle.setSelected(false);
        }
    }

    public List<FileType> getSelectedFileTypes() {
        return fileTypeToggleMap.entrySet().stream()
                .filter(entry -> entry.getKey().isSelected())
                .map(Map.Entry::getValue)
                .collect(Collectors.toList());
    }

    public boolean hasSelectedFileTypes() {
        return fileTypeToggleMap.keySet().stream()
                .anyMatch(ToggleSwitch::isSelected);
    }
}