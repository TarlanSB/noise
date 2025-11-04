package com.tsb.noise.service.utils;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.prefs.Preferences;

public class PreferencesService {
    private static final Logger log = LoggerFactory.getLogger(PreferencesService.class);

    private static final String PREFERENCES_NODE = "com_tsb_noise_app";
    private static final String LAST_SELECTED_PATH_KEY = "last_selected_path";

    private final Preferences preferences;

    public PreferencesService() {
        try {
            this.preferences = Preferences.userRoot().node(PREFERENCES_NODE);
            log.info("Preferences service initialized successfully");
        } catch (Exception e) {
            log.error("Failed to initialize preferences service: {}", e.getMessage(), e);
            throw new RuntimeException("Cannot initialize preferences service", e);
        }
    }

    public void saveLastSelectedPath(String path) {
        try {
            if (path != null && !path.trim().isEmpty()) {
                preferences.put(LAST_SELECTED_PATH_KEY, path);
                preferences.flush();
                log.info("Path saved to preferences: {}", path);
            }
        } catch (Exception e) {
            log.error("Error saving path to preferences: {}", e.getMessage(), e);
        }
    }

    public String getLastSelectedPath() {
        try {
            String path = preferences.get(LAST_SELECTED_PATH_KEY, "");
            if (!path.isEmpty()) {
                log.info("Retrieved path from preferences: {}", path);
            }
            return path;
        } catch (Exception e) {
            log.error("Error reading path from preferences: {}", e.getMessage(), e);
            return "";
        }
    }

    public void clearPreferences() {
        try {
            preferences.remove(LAST_SELECTED_PATH_KEY);
            preferences.flush();
            log.info("Preferences cleared");
        } catch (Exception e) {
            log.error("Error clearing preferences: {}", e.getMessage(), e);
        }
    }
}