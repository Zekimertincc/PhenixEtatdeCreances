package com.zeki.merger.service.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Centralized factory for reusable Excel cell styles.
 * Avoids style-limit issues by caching one instance per workbook.
 *
 * TrfSheetWriter and EtatPublicGenerator each manage their own inner Styles class
 * for backwards compatibility; this factory is for new code paths.
 */
public class ExcelStyleFactory {

    /** Dark-blue header: white bold text on #1F4E79 background. */
    public XSSFCellStyle getHeaderStyle(XSSFWorkbook wb) {
        XSSFCellStyle style = wb.createCellStyle();
        XSSFFont font = wb.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 10);
        font.setColor(new XSSFColor(new byte[]{(byte) 0xFF, (byte) 0xFF, (byte) 0xFF}, null));
        style.setFont(font);
        style.setFillForegroundColor(
            new XSSFColor(new byte[]{(byte) 0x1F, (byte) 0x4E, (byte) 0x79}, null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);
        applyThinBorder(style, wb);
        return style;
    }

    /** Plain data cell: default font, thin grey borders. */
    public XSSFCellStyle getDataStyle(XSSFWorkbook wb) {
        XSSFCellStyle style = wb.createCellStyle();
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        applyThinBorder(style, wb);
        return style;
    }

    /** Monetary cell: right-aligned, #,##0.00 format. */
    public XSSFCellStyle getCurrencyStyle(XSSFWorkbook wb) {
        XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
        style.cloneStyleFrom(getDataStyle(wb));
        style.setDataFormat(wb.createDataFormat().getFormat("#,##0.00"));
        style.setAlignment(HorizontalAlignment.RIGHT);
        return style;
    }

    /** Total row: bold text on light-yellow (#FFF2CC) background. */
    public XSSFCellStyle getTotalStyle(XSSFWorkbook wb) {
        XSSFCellStyle style = wb.createCellStyle();
        XSSFFont font = wb.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 10);
        style.setFont(font);
        style.setFillForegroundColor(
            new XSSFColor(new byte[]{(byte) 0xFF, (byte) 0xF2, (byte) 0xCC}, null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        applyThinBorder(style, wb);
        return style;
    }

    /** Total-money: getTotalStyle + #,##0.00 format. */
    public XSSFCellStyle getTotalMoneyStyle(XSSFWorkbook wb) {
        XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
        style.cloneStyleFrom(getTotalStyle(wb));
        style.setDataFormat(wb.createDataFormat().getFormat("#,##0.00"));
        style.setAlignment(HorizontalAlignment.RIGHT);
        return style;
    }

    /** Date cell: dd/MM/yyyy format. */
    public XSSFCellStyle getDateStyle(XSSFWorkbook wb) {
        XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
        style.cloneStyleFrom(getDataStyle(wb));
        style.setDataFormat(wb.getCreationHelper().createDataFormat().getFormat("dd/MM/yyyy"));
        return style;
    }

    private void applyThinBorder(XSSFCellStyle style, XSSFWorkbook wb) {
        XSSFColor grey = new XSSFColor(new byte[]{(byte) 0xD9, (byte) 0xD9, (byte) 0xD9}, null);
        style.setBorderTop(BorderStyle.THIN);    style.setTopBorderColor(grey);
        style.setBorderBottom(BorderStyle.THIN); style.setBottomBorderColor(grey);
        style.setBorderLeft(BorderStyle.THIN);   style.setLeftBorderColor(grey);
        style.setBorderRight(BorderStyle.THIN);  style.setRightBorderColor(grey);
    }
}
