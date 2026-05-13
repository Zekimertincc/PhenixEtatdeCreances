package com.zeki.merger.service.report;

import com.zeki.merger.service.excel.ExcelSheetBuilder;

import java.io.File;
import java.io.FileOutputStream;
import java.util.List;
import java.util.Map;

/**
 * Strategy that generates a simple tabular XLSX report.
 * For domain-specific reports (TRF, EtatPublic), use their dedicated writers.
 */
public class ExcelReportStrategy implements ReportStrategy {

    @Override
    public File generate(Map<String, Object> data, File outputPath) throws Exception {
        @SuppressWarnings("unchecked")
        List<String> headers = (List<String>) data.get("headers");
        @SuppressWarnings("unchecked")
        List<List<Object>> rows = (List<List<Object>>) data.get("rows");

        String title = data.containsKey("title") ? (String) data.get("title") : "Report";

        ExcelSheetBuilder builder = new ExcelSheetBuilder(title)
            .withFrozenPane(1, 0)
            .withAutoFilter(true)
            .addHeaderRow(headers)
            .addDataRows(rows);

        try (FileOutputStream fos = new FileOutputStream(outputPath)) {
            builder.build().write(fos);
        }
        return outputPath;
    }

    @Override
    public String getFormat() {
        return "XLSX";
    }
}
