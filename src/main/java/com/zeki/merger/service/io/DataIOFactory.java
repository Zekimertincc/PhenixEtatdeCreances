package com.zeki.merger.service.io;

import com.zeki.merger.core.exception.BusinessException;
import com.zeki.merger.core.exception.ErrorCode;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

/**
 * Factory Pattern: maps file extensions to reader/writer strategies.
 *
 * Readers return Map<String,Object> — a generic property bag consumed by ReportStrategy.
 * Writers serialize that same map to a file.
 *
 * Note: The existing ExcelReader / ExcelWriter classes are domain-specific (they return
 * List<CreanceRow>). This factory is for the generic report pipeline introduced by the
 * refactoring. Domain readers are still invoked directly from their service classes.
 */
public class DataIOFactory {

    public interface FileReader {
        Map<String, Object> read(File file) throws Exception;
    }

    public interface FileWriter {
        File write(Map<String, Object> data, File outputPath) throws Exception;
    }

    private final Map<String, FileReader> readers = new HashMap<>();
    private final Map<String, FileWriter> writers = new HashMap<>();

    public DataIOFactory() {
        // Default Excel reader: returns raw row data as a property map
        FileReader excelReader = file -> {
            org.apache.poi.ss.usermodel.Workbook wb;
            try (java.io.FileInputStream fis = new java.io.FileInputStream(file)) {
                wb = file.getName().toLowerCase().endsWith(".xls")
                    ? new org.apache.poi.hssf.usermodel.HSSFWorkbook(fis)
                    : new org.apache.poi.xssf.usermodel.XSSFWorkbook(fis);
            }
            org.apache.poi.ss.usermodel.Sheet sheet = wb.getSheetAt(0);
            org.apache.poi.ss.usermodel.DataFormatter fmt = new org.apache.poi.ss.usermodel.DataFormatter();
            org.apache.poi.ss.usermodel.FormulaEvaluator eval =
                wb.getCreationHelper().createFormulaEvaluator();

            java.util.List<String> headers = new java.util.ArrayList<>();
            java.util.List<java.util.List<Object>> rows = new java.util.ArrayList<>();

            org.apache.poi.ss.usermodel.Row headerRow = sheet.getRow(0);
            if (headerRow != null) {
                for (int c = 0; c < headerRow.getLastCellNum(); c++) {
                    org.apache.poi.ss.usermodel.Cell cell = headerRow.getCell(c);
                    headers.add(cell != null ? fmt.formatCellValue(cell).trim() : "");
                }
            }
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                org.apache.poi.ss.usermodel.Row row = sheet.getRow(r);
                if (row == null) continue;
                java.util.List<Object> values = new java.util.ArrayList<>();
                for (int c = 0; c < headers.size(); c++) {
                    org.apache.poi.ss.usermodel.Cell cell =
                        row.getCell(c, org.apache.poi.ss.usermodel.Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    values.add(cell != null ? fmt.formatCellValue(cell, eval).trim() : "");
                }
                rows.add(values);
            }
            wb.close();
            return Map.of("headers", headers, "rows", rows);
        };

        readers.put("xlsx", excelReader);
        readers.put("xls",  excelReader);

        // Default Excel writer delegates to ExcelReportStrategy
        FileWriter excelWriter = (data, outputPath) -> {
            new com.zeki.merger.service.report.ExcelReportStrategy().generate(data, outputPath);
            return outputPath;
        };

        writers.put("xlsx", excelWriter);
        writers.put("xls",  excelWriter);
    }

    public void registerReader(String extension, FileReader reader) {
        readers.put(extension.toLowerCase(), reader);
    }

    public void registerWriter(String extension, FileWriter writer) {
        writers.put(extension.toLowerCase(), writer);
    }

    public FileReader getReader(String fileExtension) {
        String ext = fileExtension.toLowerCase().replace(".", "");
        FileReader reader = readers.get(ext);
        if (reader == null) {
            throw new BusinessException(
                ErrorCode.INVALID_FORMAT,
                "No reader registered for format: " + fileExtension
            );
        }
        return reader;
    }

    public FileWriter getWriter(String fileExtension) {
        String ext = fileExtension.toLowerCase().replace(".", "");
        FileWriter writer = writers.get(ext);
        if (writer == null) {
            throw new BusinessException(
                ErrorCode.INVALID_FORMAT,
                "No writer registered for format: " + fileExtension
            );
        }
        return writer;
    }

    public FileReader getReaderByFile(File file) {
        return getReader(getExtension(file));
    }

    public FileWriter getWriterByFile(File file) {
        return getWriter(getExtension(file));
    }

    private String getExtension(File file) {
        String name = file.getName();
        int dot = name.lastIndexOf('.');
        if (dot <= 0) {
            throw new BusinessException(
                ErrorCode.INVALID_FORMAT,
                "File has no extension: " + name
            );
        }
        return name.substring(dot + 1);
    }
}
