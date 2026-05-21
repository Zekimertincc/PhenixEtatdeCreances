package com.zeki.merger.trf;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.text.Normalizer;
import java.util.HashMap;
import java.util.Map;

public class TrfColumnResolver {

    public static Map<String, Integer> resolve(Row header) {

        Map<String, Integer> cols = new HashMap<>();

        for (Cell cell : header) {

            String raw = cell.getStringCellValue();

            if (raw == null) {
                continue;
            }

            String normalized = normalize(raw);

            cols.put(normalized, cell.getColumnIndex());
        }

        System.out.println("Detected columns:");
        System.out.println(cols);

        return cols;
    }

    public static String normalize(String s) {

        if (s == null) {
            return "";
        }

        String normalized = Normalizer.normalize(s, Normalizer.Form.NFD)
                .replaceAll("\\p{M}", "");

        return normalized
                .trim()
                .toUpperCase()
                .replace("_", " ")
                .replaceAll("\\s+", " ");
    }

    public static Integer get(Map<String, Integer> cols, String key) {

        String normalized = normalize(key);

        Integer value = cols.get(normalized);

        if (value == null) {
            throw new RuntimeException(
                    "Column not found: " + key
            );
        }

        return value;
    }
}