package com.zeki.merger.trf.model;

import java.util.List;

/**
 * Represents one row read from the ConsolidationGenerale "Consolidation" sheet.
 * Row 0 of the source is marked as the header row.
 * Rows whose column A starts with "Total " are summary (total) rows.
 */
public class ConsolidationRow {

    private final List<Object> values;
    private final boolean headerRow;
    private final boolean totalRow;
    private final String clientName;

    public ConsolidationRow(List<Object> values, boolean headerRow) {
        this.values    = List.copyOf(values);
        this.headerRow = headerRow;
        if (!headerRow && !values.isEmpty()) {
            String a = str(values.get(0));
            this.totalRow   = a.startsWith("Total ");
            this.clientName = totalRow ? a.substring(6).trim() : a;
        } else {
            this.totalRow   = false;
            this.clientName = "";
        }
    }

    // ---- Column A: 0-based index 0 ----------------------------------------

    public List<Object> getValues()  { return values; }
    public boolean isHeaderRow()     { return headerRow; }
    public boolean isTotalRow()      { return totalRow; }
    public String  getClientName()   { return clientName; }

    public String getString(int col) {
        if (col < 0 || col >= values.size()) return "";
        return str(values.get(col));
    }

    public double getDouble(int col) {
        if (col < 0 || col >= values.size()) return 0.0;
        return toDouble(values.get(col));
    }

    // ---- Static helpers (also used by DataReader) --------------------------

    private static String str(Object v) {
        return v != null ? v.toString().trim() : "";
    }

    static double toDouble(Object v) {
        if (v == null)                return 0.0;
        if (v instanceof Double d)    return d;
        if (v instanceof Number n)    return n.doubleValue();
        if (v instanceof String s)    return parseFrenchDouble(s);
        return 0.0;
    }

    /**
     * Parses French-formatted numbers: "1 680,00 €", "1.680,00", "1680.00"
     * Returns 0.0 on failure (not a number).
     */
    public static double parseFrenchDouble(String s) {
        if (s == null || s.isBlank()) return 0.0;
        // Strip currency symbols and non-breaking spaces
        String c = s.replaceAll("[€$£¥₺  ]", "").trim();
        if (c.isEmpty()) return 0.0;

        boolean hasComma = c.contains(",");
        boolean hasDot   = c.contains(".");

        if (hasComma && hasDot) {
            // Determine which is the decimal separator (the last one wins)
            int lastComma = c.lastIndexOf(',');
            int lastDot   = c.lastIndexOf('.');
            if (lastComma > lastDot) {
                // French: "1.234,56" → dots are thousands, comma is decimal
                c = c.replace(".", "").replace(",", ".");
            } else {
                // Anglo: "1,234.56" → commas are thousands, dot is decimal
                c = c.replace(",", "");
            }
        } else if (hasComma) {
            // French without dots: "1 234,56" or just "1234,56"
            c = c.replace(",", ".");
        }
        // Remove any remaining whitespace (thousands space separator)
        c = c.replaceAll("[\\s  ]", "");

        try {
            return Double.parseDouble(c);
        } catch (NumberFormatException e) {
            return 0.0;
        }
    }
}
