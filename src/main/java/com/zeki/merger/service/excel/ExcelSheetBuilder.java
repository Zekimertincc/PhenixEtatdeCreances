package com.zeki.merger.service.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;

/**
 * Fluent builder for simple tabular Excel sheets.
 *
 * Usage:
 *   Workbook wb = new ExcelSheetBuilder("MySheet")
 *       .withFrozenPane(1, 0)
 *       .addHeaderRow(List.of("Name", "Amount"))
 *       .addDataRow(List.of("ACME", 1234.56))
 *       .build();
 *
 * Note: This is a general-purpose builder. Domain-specific sheets (TRF, EtatPublic)
 * are handled by their own writer classes which have richer formatting logic.
 */
public class ExcelSheetBuilder {

    private final Workbook workbook;
    private final Sheet sheet;
    private final ExcelStyleFactory styleFactory;

    private int currentRow = 0;
    private int freezeRow = 0;
    private int freezeCol = 0;
    private boolean autoFilter = false;
    private int headerColCount = 0;

    public ExcelSheetBuilder(String sheetName) {
        this.workbook     = new XSSFWorkbook();
        this.sheet        = workbook.createSheet(sheetName);
        this.styleFactory = new ExcelStyleFactory();
    }

    public ExcelSheetBuilder withDefaultColumnWidth(int chars) {
        sheet.setDefaultColumnWidth(chars);
        return this;
    }

    public ExcelSheetBuilder withFrozenPane(int rows, int cols) {
        this.freezeRow = rows;
        this.freezeCol = cols;
        return this;
    }

    public ExcelSheetBuilder withAutoFilter(boolean enabled) {
        this.autoFilter = enabled;
        return this;
    }

    public ExcelSheetBuilder addHeaderRow(List<String> headers) {
        Row row = sheet.createRow(currentRow++);
        CellStyle headerStyle = styleFactory.getHeaderStyle((org.apache.poi.xssf.usermodel.XSSFWorkbook) workbook);
        for (int i = 0; i < headers.size(); i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(headers.get(i));
            cell.setCellStyle(headerStyle);
        }
        this.headerColCount = headers.size();
        return this;
    }

    public ExcelSheetBuilder addDataRow(List<Object> values) {
        Row row = sheet.createRow(currentRow++);
        CellStyle dataStyle  = styleFactory.getDataStyle((org.apache.poi.xssf.usermodel.XSSFWorkbook) workbook);
        CellStyle moneyStyle = styleFactory.getCurrencyStyle((org.apache.poi.xssf.usermodel.XSSFWorkbook) workbook);

        for (int i = 0; i < values.size(); i++) {
            Cell   cell  = row.createCell(i);
            Object value = values.get(i);

            if (value instanceof Double || value instanceof Float) {
                cell.setCellValue(((Number) value).doubleValue());
                cell.setCellStyle(moneyStyle);
            } else if (value instanceof Number) {
                cell.setCellValue(((Number) value).doubleValue());
                cell.setCellStyle(dataStyle);
            } else if (value instanceof Boolean) {
                cell.setCellValue((Boolean) value);
                cell.setCellStyle(dataStyle);
            } else {
                cell.setCellValue(value != null ? value.toString() : "");
                cell.setCellStyle(dataStyle);
            }
        }
        return this;
    }

    public ExcelSheetBuilder addDataRows(List<List<Object>> rows) {
        for (List<Object> row : rows) {
            addDataRow(row);
        }
        return this;
    }

    public Workbook build() {
        if (freezeRow > 0 || freezeCol > 0) {
            sheet.createFreezePane(freezeCol, freezeRow);
        }
        if (autoFilter && currentRow > 0 && headerColCount > 0) {
            sheet.setAutoFilter(new CellRangeAddress(0, currentRow - 1, 0, headerColCount - 1));
        }
        return workbook;
    }
}
