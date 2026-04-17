package com.zeki.merger.model;

import java.util.List;


public class CreanceRow {

    private final String societe;
    private final List<Object> cellValues;   // raw values in source column order
    private final int originalRowIndex;      // 0-based row index in source sheet

    public CreanceRow(String societe, List<Object> cellValues, int originalRowIndex) {
        this.societe = societe;
        this.cellValues = List.copyOf(cellValues);
        this.originalRowIndex = originalRowIndex;
    }

    public String getSociete()              { return societe; }
    public List<Object> getCellValues()     { return cellValues; }
    public int getOriginalRowIndex()        { return originalRowIndex; }
}
