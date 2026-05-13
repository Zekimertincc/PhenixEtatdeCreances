package com.zeki.merger.controller.command;

import com.zeki.merger.service.util.ProgressNotifier;

import java.io.File;
import java.util.Map;

/**
 * Command interface for UI-triggered operations.
 *
 * Each action button in MainController maps to one implementation.
 * Commands receive a context map carrying file paths and a ProgressNotifier
 * for UI feedback — keeping the command decoupled from JavaFX.
 *
 * Context keys (all optional, validated by each command):
 *   "consoFile"    → File
 *   "listingFile"  → File
 *   "tableauFile"  → File
 *   "outputFolder" → File
 *   "rootFolder"   → File
 *   "procFile"     → File
 */
public interface ReportCommand {
    File execute(Map<String, Object> context, ProgressNotifier notifier) throws Exception;
    String getName();
}
