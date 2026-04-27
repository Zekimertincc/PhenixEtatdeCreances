package com.zeki.merger.trf;

import com.zeki.merger.trf.model.ClientInfo;
import com.zeki.merger.trf.model.ConsolidationRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.Normalizer;
import java.time.LocalDateTime;
import java.util.*;

/**
 * Reads all three TRF input files:
 * <ol>
 *   <li>ConsolidationGenerale.xlsx  → sheet "Consolidation"</li>
 *   <li>LISTING_CABINET_PHENIX_pour_ZEKI.xls → sheet "Feuil1" (IBAN / NonComp)</li>
 *   <li>Tableau_de_bord_facturation.xlsx → sheet "Soldes" (previous balances)</li>
 * </ol>
 */
public class DataReader {

    // Columns in ConsolidationGenerale "Consolidation" sheet that hold numeric values
    // (0-based). Strings in these columns are parsed as French-formatted numbers.
    private static final Set<Integer> NUMERIC_COLS = Set.of(
        1,              // B  NBRE
        7,              // H  CREANCE PRINCIPALE
        8,              // I  RECOUVRE ET FACTURE
        11,             // L  PENALITES
        13,             // N  Transformation colonne L
        14,             // O  CONDITION
        15,             // P  DONT EN ATTENTE
        16,             // Q  Lieu (sometimes numeric codes)
        17,             // R  Frais de procédure
        18,             // S  Recouvré total
        19,             // T  Déjà facturé
        20,             // U  Depuis le début
        21,             // V  Commissions
        22,             // W  Pénalits
        23,             // X  SOMMES CZ PHENIX
        24,             // Y  MONTANT A FACTURER TTC
        25              // Z  SOMMES A REVERSER
    );

    // -------------------------------------------------------------------------
    // Public reading methods
    // -------------------------------------------------------------------------

    /**
     * Reads every row (including the header at row 0) from the "Consolidation"
     * sheet of ConsolidationGenerale.xlsx.
     */
    public List<ConsolidationRow> readAllConsolidationRows(File file) throws IOException {
        List<ConsolidationRow> rows = new ArrayList<>();
        try (Workbook wb = openWorkbook(file)) {
            Sheet sheet = requireSheet(wb, file.getName(), "Consolidation");
            DataFormatter    fmt  = new DataFormatter();
            FormulaEvaluator eval = wb.getCreationHelper().createFormulaEvaluator();

            for (int r = 0; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) {
                    rows.add(new ConsolidationRow(List.of(), false));
                    continue;
                }
                List<Object> vals = extractConsolidationRowValues(row, fmt, eval);
                rows.add(new ConsolidationRow(vals, r == 0));
            }
        }
        return rows;
    }

    /**
     * Reads the "Feuil1" sheet of LISTING_CABINET_PHENIX_pour_ZEKI.xls and returns
     * a map keyed by <em>normalised</em> client name.
     *
     * Column indices (0-based):
     *   C=2 name | D=3 code | U=20 NonComp | V=21 IBAN | W=22 BIC
     */
    public Map<String, ClientInfo> readClientInfoMap(File file) throws IOException {
        Map<String, ClientInfo> map = new LinkedHashMap<>();
        try (Workbook wb = openWorkbook(file)) {
            Sheet sheet = findSheetFallback(wb, "Feuil1");
            DataFormatter    fmt  = new DataFormatter();
            FormulaEvaluator eval = wb.getCreationHelper().createFormulaEvaluator();

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                String name = cellStr(row, 2, fmt, eval);
                if (name.isBlank()) continue;

                String code    = cellStr(row,  3, fmt, eval);
                if (code.endsWith(".0")) code = code.substring(0, code.length() - 2);
                String nonComp = cellStr(row, 20, fmt, eval);
                String iban    = cellStr(row, 21, fmt, eval);
                String bic     = cellStr(row, 22, fmt, eval);

                map.put(normalize(name), new ClientInfo(name, code, nonComp, iban, bic));
            }
        }
        return map;
    }

    /**
     * Reads the "Soldes" sheet (index 2) of Tableau_de_bord_facturation.xlsx and returns
     * a map of normalised client name → NOUS DOIT amount.
     *
     * Column indices (0-based): A=0 name | K=10 NOUS DOIT amount
     */
    public Map<String, Double> readPreviousBalances(File file) throws IOException {
        Map<String, Double> map = new LinkedHashMap<>();
        try (Workbook wb = openWorkbook(file)) {
            Sheet sheet = findSoldesSheet(wb);
            DataFormatter    fmt  = new DataFormatter();
            FormulaEvaluator eval = wb.getCreationHelper().createFormulaEvaluator();

            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                String name = cellStr(row, 0, fmt, eval);
                if (name.isBlank()) continue;
                double amount = cellDouble(row, 10, fmt, eval);  // col K (index 10)
                map.put(normalize(name), amount);
            }
        }
        return map;
    }

    // -------------------------------------------------------------------------
    // Client lookup helpers (used by TrfCalculator)
    // -------------------------------------------------------------------------

    /** Finds a ClientInfo by name with exact-then-partial normalised matching. */
    public ClientInfo findClientInfo(String clientName, Map<String, ClientInfo> infoMap) {
        if (clientName == null || clientName.isBlank()) return null;
        String norm = normalize(clientName);
        ClientInfo ci = infoMap.get(norm);
        if (ci != null) return ci;
        // Partial match: the listing entry might be a shorter/longer version
        for (Map.Entry<String, ClientInfo> e : infoMap.entrySet()) {
            String k = e.getKey();
            if (norm.contains(k) || k.contains(norm)) return e.getValue();
        }
        return null;
    }

    /** Looks up the previous balance for a client, returning 0 if not found. */
    public double findBalance(String clientName, Map<String, Double> balanceMap) {
        if (clientName == null || clientName.isBlank()) return 0.0;
        String norm = normalize(clientName);
        Double v = balanceMap.get(norm);
        if (v != null) return v;
        for (Map.Entry<String, Double> e : balanceMap.entrySet()) {
            String k = e.getKey();
            if (norm.contains(k) || k.contains(norm)) return e.getValue();
        }
        return 0.0;
    }

    // -------------------------------------------------------------------------
    // Private helpers
    // -------------------------------------------------------------------------

    private Workbook openWorkbook(File file) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        return file.getName().toLowerCase().endsWith(".xls")
            ? new HSSFWorkbook(fis)
            : new XSSFWorkbook(fis);
    }

    private Sheet requireSheet(Workbook wb, String fileName, String sheetName) throws IOException {
        Sheet s = findSheetByName(wb, sheetName);
        if (s == null) throw new IOException(
            "Sheet '" + sheetName + "' not found in " + fileName);
        return s;
    }

    private Sheet findSheetFallback(Workbook wb, String name) {
        Sheet s = findSheetByName(wb, name);
        return s != null ? s : wb.getSheetAt(0);
    }

    private Sheet findSoldesSheet(Workbook wb) {
        Sheet s = findSheetByName(wb, "Soldes");
        if (s != null) return s;
        // Spec: Soldes is at sheet index 2; fall back to index 0 if not present
        return wb.getNumberOfSheets() > 2 ? wb.getSheetAt(2) : wb.getSheetAt(0);
    }

    private Sheet findSheetByName(Workbook wb, String name) {
        Sheet s = wb.getSheet(name);
        if (s != null) return s;
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            if (wb.getSheetName(i).equalsIgnoreCase(name)) return wb.getSheetAt(i);
        }
        return null;
    }

    private List<Object> extractConsolidationRowValues(Row row, DataFormatter fmt,
                                                        FormulaEvaluator eval) {
        int lastCell = Math.max(row.getLastCellNum(), 26); // always read at least A-Z
        List<Object> values = new ArrayList<>(lastCell);
        for (int c = 0; c < lastCell; c++) {
            Cell cell = row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            if (cell == null) { values.add(""); continue; }
            values.add(readCellValue(cell, fmt, eval, NUMERIC_COLS.contains(c)));
        }
        return values;
    }

    private Object readCellValue(Cell cell, DataFormatter fmt,
                                  FormulaEvaluator eval, boolean numericCol) {
        CellType type;
        double  numVal  = 0;
        String  strVal  = "";
        boolean boolVal = false;

        if (cell.getCellType() == CellType.FORMULA) {
            try {
                CellValue cv = eval.evaluate(cell);
                type = cv.getCellType();
                switch (type) {
                    case NUMERIC -> numVal  = cv.getNumberValue();
                    case STRING  -> strVal  = cv.getStringValue();
                    case BOOLEAN -> boolVal = cv.getBooleanValue();
                    default      -> {}
                }
            } catch (Exception ex) {
                return "";
            }
        } else {
            type = cell.getCellType();
            switch (type) {
                case NUMERIC -> numVal  = cell.getNumericCellValue();
                case STRING  -> strVal  = cell.getStringCellValue();
                case BOOLEAN -> boolVal = cell.getBooleanCellValue();
                default      -> {}
            }
        }

        return switch (type) {
            case NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    yield DateUtil.getJavaDate(numVal)
                        .toInstant().atZone(java.time.ZoneId.systemDefault()).toLocalDateTime();
                }
                yield numVal;
            }
            case BOOLEAN -> boolVal;
            case STRING  -> {
                if (!strVal.isBlank() && numericCol) {
                    double d = ConsolidationRow.parseFrenchDouble(strVal);
                    // Return numeric if parse succeeded (non-zero, or string clearly is "0")
                    if (d != 0.0 || looksLikeZero(strVal)) yield d;
                }
                yield strVal.trim();
            }
            default -> "";
        };
    }

    private static boolean looksLikeZero(String s) {
        return s.replaceAll("[0,\\.\\s€ ]", "").isEmpty();
    }

    private String cellStr(Row row, int col, DataFormatter fmt, FormulaEvaluator eval) {
        Cell cell = row.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (cell == null) return "";
        return fmt.formatCellValue(cell, eval).trim();
    }

    private double cellDouble(Row row, int col, DataFormatter fmt, FormulaEvaluator eval) {
        Cell cell = row.getCell(col, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (cell == null) return 0.0;
        CellType type = cell.getCellType() == CellType.FORMULA
            ? cell.getCachedFormulaResultType()
            : cell.getCellType();
        if (type == CellType.NUMERIC) return cell.getNumericCellValue();
        return ConsolidationRow.parseFrenchDouble(fmt.formatCellValue(cell, eval).trim());
    }

    /** Normalises a string for fuzzy matching: lowercase, no accents, single spaces. */
    public static String normalize(String s) {
        if (s == null || s.isBlank()) return "";
        return Normalizer.normalize(s.trim(), Normalizer.Form.NFD)
            .replaceAll("\\p{M}", "")
            .toLowerCase()
            .replaceAll("\\s+", " ");
    }
}
