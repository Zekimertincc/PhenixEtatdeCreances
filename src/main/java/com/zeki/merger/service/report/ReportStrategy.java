package com.zeki.merger.service.report;

import java.io.File;
import java.util.Map;

/**
 * Strategy interface for report generation.
 * Each implementation handles one output format (XLSX, PDF, …).
 *
 * data keys:
 *   "headers" → List<String>
 *   "rows"    → List<List<Object>>
 *   "title"   → String (optional)
 */
public interface ReportStrategy {
    File generate(Map<String, Object> data, File outputPath) throws Exception;
    String getFormat();
}
