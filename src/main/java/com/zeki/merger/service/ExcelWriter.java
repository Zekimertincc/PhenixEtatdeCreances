package com.zeki.merger.service;

import com.zeki.merger.model.CreanceRow;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Map;

/**
 * Writes grouped company rows to a single XLSX file matching the
 * ConsolidationGenerale format:
 *
 *   [Company header row]  — company name in column B, light-blue background
 *   [Empty spacer row]
 *   [Data rows…]          — company name in column A, source data in columns B+
 *   (repeated per company)
 *
 * Companies with no matching rows are skipped entirely.
 */
public class ExcelWriter {

    /** Total number of output columns (A–AH = 34). */
    private static final int TOTAL_COLS = 34;

    /**
     * @param groupedRows  insertion-ordered map of company name → matching rows;
     *                     companies with an empty list are skipped.
     * @param outputFile   destination .xlsx file (created or overwritten).
     */
    public void write(Map<String, List<CreanceRow>> groupedRows, File outputFile)
            throws IOException {

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            XSSFSheet sheet = wb.createSheet("Créances");

            // ---- shared styles ----
            XSSFCellStyle companyHeaderStyle = buildCompanyHeaderStyle(wb);
            XSSFCellStyle dataStyle          = buildDataStyle(wb);
            XSSFCellStyle dateStyle          = buildDateStyle(wb, dataStyle);

            int rowIdx = 0;

            for (Map.Entry<String, List<CreanceRow>> entry : groupedRows.entrySet()) {
                String           company = entry.getKey();
                List<CreanceRow> rows    = entry.getValue();

                // MergeService should have filtered empty companies, but guard here too.
                if (rows == null || rows.isEmpty()) {
                    System.out.println("[" + company + "] SKIPPED - no data in column S");
                    continue;
                }

                // ---- company header row: name in column B, background across all cols ----
                XSSFRow headerRow = sheet.createRow(rowIdx++);
                for (int c = 0; c < TOTAL_COLS; c++) {
                    XSSFCell cell = headerRow.createCell(c);
                    cell.setCellStyle(companyHeaderStyle);
                }
                headerRow.getCell(1).setCellValue(company); // column B

                // ---- empty spacer row ----
                sheet.createRow(rowIdx++);

                // ---- data rows ----
                for (CreanceRow cr : rows) {
                    XSSFRow row = sheet.createRow(rowIdx++);

                    // Column A: company name
                    XSSFCell societeCell = row.createCell(0);
                    societeCell.setCellValue(company);
                    societeCell.setCellStyle(dataStyle);

                    // Columns B onwards: original source data
                    List<Object> values = cr.getCellValues();
                    for (int c = 0; c < values.size(); c++) {
                        XSSFCell cell = row.createCell(c + 1);
                        writeValue(cell, values.get(c), dataStyle, dateStyle);
                    }

                    // Fill any trailing columns up to TOTAL_COLS with the border style
                    // so borders are consistent across the full row width.
                    for (int c = values.size() + 1; c < TOTAL_COLS; c++) {
                        XSSFCell cell = row.createCell(c);
                        cell.setCellStyle(dataStyle);
                    }
                }
            }

            // ---- auto-size all columns (capped to avoid slow wide columns) ----
            for (int c = 0; c < TOTAL_COLS; c++) {
                sheet.autoSizeColumn(c);
                int w = sheet.getColumnWidth(c);
                sheet.setColumnWidth(c, Math.min(w + 512, 20_000));
            }

            // ---- freeze the very first row ----
            sheet.createFreezePane(0, 1);

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                wb.write(fos);
            }
        }
    }

    // -------------------------------------------------------------------------
    // Style builders
    // -------------------------------------------------------------------------

    /** Bold dark text, light-blue background (#BDD7EE), no borders. */
    private XSSFCellStyle buildCompanyHeaderStyle(XSSFWorkbook wb) {
        XSSFCellStyle s = wb.createCellStyle();
        XSSFFont f = wb.createFont();
        f.setBold(true);
        f.setFontHeightInPoints((short) 11);
        s.setFont(f);
        // Light blue — matches Excel's "Light Blue" theme cell fill
        s.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 0xBD, (byte) 0xD7, (byte) 0xEE}, null));
        s.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        return s;
    }

    /** No background, thin borders on all four sides. */
    private XSSFCellStyle buildDataStyle(XSSFWorkbook wb) {
        XSSFCellStyle s = wb.createCellStyle();
        s.setBorderTop(BorderStyle.THIN);
        s.setBorderBottom(BorderStyle.THIN);
        s.setBorderLeft(BorderStyle.THIN);
        s.setBorderRight(BorderStyle.THIN);
        XSSFColor borderColor = new XSSFColor(new byte[]{(byte) 0xD9, (byte) 0xD9, (byte) 0xD9}, null);
        s.setTopBorderColor(borderColor);
        s.setBottomBorderColor(borderColor);
        s.setLeftBorderColor(borderColor);
        s.setRightBorderColor(borderColor);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        return s;
    }

    /** Data style + dd/MM/yyyy date format. */
    private XSSFCellStyle buildDateStyle(XSSFWorkbook wb, XSSFCellStyle base) {
        XSSFCellStyle s = wb.createCellStyle();
        s.cloneStyleFrom(base);
        s.setDataFormat(wb.getCreationHelper().createDataFormat().getFormat("dd/MM/yyyy"));
        return s;
    }

    // -------------------------------------------------------------------------
    // Value writer
    // -------------------------------------------------------------------------

    private void writeValue(XSSFCell cell, Object val,
                            XSSFCellStyle defaultStyle, XSSFCellStyle dateStyle) {
        if (val instanceof Double d) {
            cell.setCellValue(d);
            cell.setCellStyle(defaultStyle);
        } else if (val instanceof Boolean b) {
            cell.setCellValue(b);
            cell.setCellStyle(defaultStyle);
        } else if (val instanceof LocalDateTime ldt) {
            cell.setCellValue(ldt);
            cell.setCellStyle(dateStyle);
        } else {
            cell.setCellValue(val != null ? val.toString() : "");
            cell.setCellStyle(defaultStyle);
        }
    }
}
