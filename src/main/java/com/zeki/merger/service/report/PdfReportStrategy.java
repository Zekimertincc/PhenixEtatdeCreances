package com.zeki.merger.service.report;

import com.zeki.merger.core.exception.BusinessException;
import com.zeki.merger.core.exception.ErrorCode;

import java.io.File;
import java.util.Map;

/**
 * Stub PDF strategy. Domain-specific PDF output (EtatPublicGenerator.writePdf) is handled
 * by EtatPublicGenerator directly using iText7. This strategy exists to complete the
 * factory and can be wired to a generic PDF builder in a future iteration.
 */
public class PdfReportStrategy implements ReportStrategy {

    @Override
    public File generate(Map<String, Object> data, File outputPath) throws Exception {
        throw new BusinessException(
            ErrorCode.GENERATION_FAILED,
            "Generic PDF generation not implemented — use EtatPublicGenerator for PDF output"
        );
    }

    @Override
    public String getFormat() {
        return "PDF";
    }
}
