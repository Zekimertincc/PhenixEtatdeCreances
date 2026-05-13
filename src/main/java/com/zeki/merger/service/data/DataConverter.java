package com.zeki.merger.service.data;

import com.zeki.merger.model.CreanceRow;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Converts raw property maps from DataIOFactory into domain model objects.
 *
 * CreanceRow is immutable — values are passed to the constructor in source-column order.
 * Column order must match the source Excel file layout.
 */
public class DataConverter {

    /**
     * Converts a raw row map (col-header → value) to a List<Object> in declaration order,
     * then wraps it in a CreanceRow.
     *
     * @param companyName the société name to tag on the row
     * @param rowMap      column-name → value from DataIOFactory.FileReader output
     * @param headers     ordered list of column names (defines value ordering)
     * @param rowIndex    0-based source row index
     */
    public CreanceRow toCreanceRow(String companyName,
                                   Map<String, Object> rowMap,
                                   List<String> headers,
                                   int rowIndex) {
        List<Object> values = new ArrayList<>(headers.size());
        for (String header : headers) {
            values.add(rowMap.getOrDefault(header, ""));
        }
        return new CreanceRow(companyName, values, rowIndex);
    }

    /**
     * Converts a list of row maps (each from DataIOFactory) to CreanceRows.
     *
     * @param companyName the société name for all rows
     * @param rowMaps     list of column-name → value maps
     * @param headers     ordered header list
     * @param startIndex  row index of the first data row in the source file
     */
    public List<CreanceRow> toCreanceRows(String companyName,
                                          List<Map<String, Object>> rowMaps,
                                          List<String> headers,
                                          int startIndex) {
        List<CreanceRow> result = new ArrayList<>(rowMaps.size());
        for (int i = 0; i < rowMaps.size(); i++) {
            result.add(toCreanceRow(companyName, rowMaps.get(i), headers, startIndex + i));
        }
        return result;
    }

    private double parseDouble(String value) {
        if (value == null || value.isBlank()) return 0.0;
        try {
            return Double.parseDouble(
                value.replaceAll("[^0-9.\\-]", "").replace(",", ".")
            );
        } catch (NumberFormatException e) {
            return 0.0;
        }
    }
}
