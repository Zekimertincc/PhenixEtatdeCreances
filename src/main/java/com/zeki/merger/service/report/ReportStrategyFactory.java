package com.zeki.merger.service.report;

import com.zeki.merger.core.exception.BusinessException;
import com.zeki.merger.core.exception.ErrorCode;

import java.util.HashMap;
import java.util.Map;
import java.util.Set;

/**
 * Factory that maps output format strings to ReportStrategy implementations.
 * New formats can be registered at runtime via register().
 */
public class ReportStrategyFactory {

    private final Map<String, ReportStrategy> strategies = new HashMap<>();

    public ReportStrategyFactory() {
        register("XLSX", new ExcelReportStrategy());
        register("XLS",  new ExcelReportStrategy());
        register("PDF",  new PdfReportStrategy());
    }

    public void register(String format, ReportStrategy strategy) {
        strategies.put(format.toUpperCase(), strategy);
    }

    public ReportStrategy getStrategy(String format) {
        ReportStrategy strategy = strategies.get(format.toUpperCase());
        if (strategy == null) {
            throw new BusinessException(
                ErrorCode.INVALID_FORMAT,
                "Unsupported report format: " + format,
                Map.of("requestedFormat", format, "supportedFormats", getSupportedFormats())
            );
        }
        return strategy;
    }

    public Set<String> getSupportedFormats() {
        return strategies.keySet();
    }
}
