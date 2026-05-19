package com.zeki.merger.service;

import com.itextpdf.kernel.colors.DeviceRgb;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.borders.Border;
import com.itextpdf.layout.borders.SolidBorder;
import com.itextpdf.layout.element.Cell;
import com.itextpdf.layout.element.Div;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.properties.TextAlignment;
import com.itextpdf.layout.properties.UnitValue;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.function.BiConsumer;

public class FacturePdfService {

    private final FolderScanner scanner = new FolderScanner();

    private static final DeviceRgb DARK_BLUE  = new DeviceRgb(0x1F, 0x4E, 0x79);
    private static final DeviceRgb MED_BLUE   = new DeviceRgb(0x2E, 0x75, 0xB6);
    private static final DeviceRgb LIGHT_GRAY = new DeviceRgb(0xF2, 0xF2, 0xF2);
    private static final DeviceRgb WHITE      = new DeviceRgb(0xFF, 0xFF, 0xFF);
    private static final DeviceRgb TEXT_MUTED = new DeviceRgb(0x6B, 0x6B, 0x6B);

    public List<String> apply(File rootFolder, BiConsumer<Double, String> progress) throws Exception {
        List<String> log = new ArrayList<>();
        List<FolderScanner.CompanyFile> companies = scanner.scan(rootFolder);
        if (companies.isEmpty()) {
            progress.accept(1.0, "Aucun dossier trouvé.");
            return log;
        }
        int total = companies.size();
        for (int i = 0; i < total; i++) {
            FolderScanner.CompanyFile cf = companies.get(i);
            double prog = 0.05 + 0.95 * (i + 1.0) / total;
            String result;
            try {
                result = processCompany(cf.excelFile());
            } catch (Exception e) {
                result = "ERREUR: " + e.getMessage();
            }
            log.add(cf.companyName() + " → " + result);
            progress.accept(prog, "[" + (i + 1) + "/" + total + "] " + cf.companyName() + " → " + result);
        }
        progress.accept(1.0, "Génération PDF terminée (" + total + " dossiers).");
        return log;
    }

    private String processCompany(File excelFile) throws Exception {
        try (Workbook wb = openWorkbook(excelFile)) {
            DataFormatter fmt = new DataFormatter();
            FormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator();

            Sheet creances = wb.getSheet("Créances");
            String nomClient = "";
            String codeClient = "";
            if (creances != null) {
                nomClient = cellStr(creances, 3, 7, fmt, ev);   // H4
                String a13 = cellStr(creances, 12, 0, fmt, ev); // A13
                codeClient = a13.length() >= 6 ? a13.substring(a13.length() - 6) : a13;
            }

            Sheet facture = wb.getSheet("Facture en préparation");
            if (facture == null) {
                for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                    if (wb.getSheetName(i).toLowerCase().contains("facture")) {
                        facture = wb.getSheetAt(i);
                        break;
                    }
                }
            }
            if (facture == null) return "sheet 'Facture en préparation' introuvable";

            List<String> adresseLines = new ArrayList<>();
            for (int r = 0; r <= 16; r++) {
                String d = cellStr(facture, r, 3, fmt, ev);
                String e = cellStr(facture, r, 4, fmt, ev);
                String line = (d + " " + e).trim();
                if (!line.isBlank()) adresseLines.add(line);
            }
            String numFacture = cellStr(facture, 12, 3, fmt, ev); // D13

            int ligneDuA = findMarker(facture, "A", 0, fmt, ev);
            int ligneDuD = ligneDuA >= 0 ? findMarker(facture, "D", ligneDuA + 3, fmt, ev) : -1;
            int ligneDuI = ligneDuD >= 0 ? findMarkerStartsWith(facture, "I", ligneDuD + 5, fmt, ev) : -1;
            int ligneConclusion = findMarkerContains(facture, "EN CONCLUSION", 0, fmt, ev);
            int ligneMentions = findMarkerStartsWith(facture, "Mentions", 0, fmt, ev);

            double ag             = ligneDuA >= 0 ? numVal(facture, ligneDuA,     2, fmt, ev) : 0;
            double cl             = ligneDuA >= 0 ? numVal(facture, ligneDuA + 1, 2, fmt, ev) : 0;
            double agcl           = ligneDuA >= 0 ? numVal(facture, ligneDuA + 2, 2, fmt, ev) : 0;
            double comsHt         = ligneDuD >= 0 ? numVal(facture, ligneDuD,     2, fmt, ev) : 0;
            double prodHt         = ligneDuD >= 0 ? numVal(facture, ligneDuD + 1, 2, fmt, ev) : 0;
            double totalHt        = ligneDuD >= 0 ? numVal(facture, ligneDuD + 2, 2, fmt, ev) : 0;
            double tva            = ligneDuD >= 0 ? numVal(facture, ligneDuD + 3, 2, fmt, ev) : 0;
            double ttc            = ligneDuD >= 0 ? numVal(facture, ligneDuD + 4, 2, fmt, ev) : 0;
            double solde          = ligneDuI >= 0 ? numVal(facture, ligneDuI,     2, fmt, ev) : 0;
            double retard         = ligneDuI >= 0 ? numVal(facture, ligneDuI + 1, 2, fmt, ev) : 0;
            double soldeComptable = ligneDuI >= 0 ? numVal(facture, ligneDuI + 2, 2, fmt, ev) : 0;

            String rib = "", iban = "", bic = "";
            int ribSearchStart = ligneDuI >= 0 ? ligneDuI + 3 : 0;
            for (int r = ribSearchStart; r <= facture.getLastRowNum(); r++) {
                for (int c = 0; c < 8; c++) {
                    String v = cellStr(facture, r, c, fmt, ev);
                    String vu = v.toUpperCase();
                    if (vu.contains("RIB") && rib.isBlank())
                        rib = v + " " + cellStr(facture, r, c + 1, fmt, ev);
                    if (vu.contains("IBAN") && iban.isBlank())
                        iban = v + " " + cellStr(facture, r, c + 1, fmt, ev);
                    if (vu.contains("BIC") && bic.isBlank())
                        bic = v + " " + cellStr(facture, r, c + 1, fmt, ev);
                }
            }

            String conclusionText = "";
            if (ligneConclusion >= 0) {
                StringBuilder sb = new StringBuilder();
                for (int r = ligneConclusion + 1; r <= Math.min(ligneConclusion + 4, facture.getLastRowNum()); r++) {
                    String v = cellStr(facture, r, 0, fmt, ev);
                    if (v.isBlank()) v = cellStr(facture, r, 1, fmt, ev);
                    if (!v.isBlank()) sb.append(v).append(" ");
                }
                conclusionText = sb.toString().trim();
            }

            String mentionsText = "";
            if (ligneMentions >= 0) {
                StringBuilder sb = new StringBuilder();
                for (int r = ligneMentions; r <= Math.min(ligneMentions + 5, facture.getLastRowNum()); r++) {
                    String v = cellStr(facture, r, 0, fmt, ev);
                    if (v.isBlank()) v = cellStr(facture, r, 1, fmt, ev);
                    if (!v.isBlank()) sb.append(v).append(" ");
                }
                mentionsText = sb.toString().trim();
            }

            String pdfName = excelFile.getName().replaceAll("\\.(xlsx?|xls)$", "") + "_facture.pdf";
            File pdfFile = new File(excelFile.getParent(), pdfName);

            generatePdf(pdfFile, nomClient, codeClient, numFacture, adresseLines,
                    ag, cl, agcl, comsHt, prodHt, totalHt, tva, ttc,
                    solde, retard, soldeComptable,
                    rib, iban, bic, conclusionText, mentionsText);

            return "PDF → " + pdfName;
        }
    }

    private void generatePdf(File pdfFile, String nomClient, String codeClient,
            String numFacture, List<String> adresse,
            double ag, double cl, double agcl,
            double comsHt, double prodHt, double totalHt, double tva, double ttc,
            double solde, double retard, double soldeComptable,
            String rib, String iban, String bic,
            String conclusion, String mentions) throws Exception {

        String today = LocalDate.now().format(DateTimeFormatter.ofPattern("dd/MM/yyyy"));

        try (PdfWriter writer = new PdfWriter(pdfFile);
             PdfDocument pdf = new PdfDocument(writer);
             Document doc = new Document(pdf, PageSize.A4)) {

            doc.setMargins(25, 25, 25, 25);

            // 1. Top header bar
            Table header = new Table(UnitValue.createPercentArray(new float[]{50, 50}))
                    .useAllAvailableWidth()
                    .setBackgroundColor(DARK_BLUE)
                    .setMarginBottom(12);
            header.addCell(new Cell()
                    .add(new Paragraph("Cabinet Phénix").setFontColor(WHITE).setBold().setFontSize(14))
                    .setBorder(Border.NO_BORDER).setPadding(10));
            header.addCell(new Cell()
                    .add(new Paragraph("FACTURE").setFontColor(WHITE).setBold().setFontSize(20)
                            .setTextAlignment(TextAlignment.RIGHT))
                    .setBorder(Border.NO_BORDER).setPadding(10));
            doc.add(header);

            // 2. Client address + facture details
            Table info = new Table(UnitValue.createPercentArray(new float[]{55, 45}))
                    .useAllAvailableWidth().setMarginBottom(16);

            Cell addrCell = new Cell().setBorder(Border.NO_BORDER)
                    .setBorderLeft(new SolidBorder(MED_BLUE, 3)).setPaddingLeft(8);
            for (String line : adresse) {
                addrCell.add(new Paragraph(line).setFontSize(9).setMarginBottom(1));
            }
            info.addCell(addrCell);

            Table details = new Table(UnitValue.createPercentArray(new float[]{50, 50}))
                    .useAllAvailableWidth().setBackgroundColor(LIGHT_GRAY);
            addDetailRow(details, "FACTURE N°", numFacture.isBlank() ? "—" : numFacture);
            addDetailRow(details, "Code client", codeClient.isBlank() ? "—" : codeClient);
            addDetailRow(details, "Date", today);
            info.addCell(new Cell().add(details).setBorder(Border.NO_BORDER).setPaddingLeft(8));
            doc.add(info);

            // 3. Encaissements
            doc.add(sectionHeader("LES ENCAISSEMENTS SELON LE LIEU"));
            Table encTable = new Table(UnitValue.createPercentArray(new float[]{15, 60, 25}))
                    .useAllAvailableWidth().setMarginBottom(10);
            addSectionRow(encTable, "A",     "Encaissements Phénix (AG)",  formatMoney(ag),   false);
            addSectionRow(encTable, "B",     "Encaissements Client (CL)",  formatMoney(cl),   false);
            addSectionRow(encTable, "C=A+B", "Encaissements du mois :",    formatMoney(agcl), true);
            doc.add(encTable);

            // 4. Informations facture
            doc.add(sectionHeader("LES INFORMATIONS LIÉES À LA FACTURE"));
            Table factTable = new Table(UnitValue.createPercentArray(new float[]{15, 60, 25}))
                    .useAllAvailableWidth().setMarginBottom(10);
            addSectionRow(factTable, "D",       "COMMISSIONS HT",         formatMoney(comsHt),  false);
            addSectionRow(factTable, "E",       "FRAIS DE PROCÉDURE HT",  formatMoney(prodHt),  false);
            addSectionRow(factTable, "F=D+E",   "TOTAL HT",               formatMoney(totalHt), false);
            addSectionRow(factTable, "G=F*20%", "TVA 20,00%",             formatMoney(tva),     false);
            addSectionRow(factTable, "H=F+G",   "TOTAL TTC",              formatMoney(ttc),     true);
            doc.add(factTable);

            // 5. Versement des fonds
            doc.add(sectionHeader("LES INFORMATIONS LIÉES AU VERSEMENT DES FONDS"));
            Table versTable = new Table(UnitValue.createPercentArray(new float[]{15, 60, 25}))
                    .useAllAvailableWidth().setMarginBottom(10);
            addSectionRow(versTable, "I",     "Le solde des encaissements de la période est en votre faveur de :", formatMoney(solde),         false);
            addSectionRow(versTable, "J",     "Factures en retard de paiement / avoirs en cours :",               formatMoney(retard),        false);
            addSectionRow(versTable, "K=I+J", "Solde comptable en votre faveur de :",                             formatMoney(soldeComptable), true);
            doc.add(versTable);

            // 6. EN CONCLUSION
            if (!conclusion.isBlank() || soldeComptable != 0) {
                Table concl = new Table(1).useAllAvailableWidth()
                        .setBackgroundColor(LIGHT_GRAY).setMarginBottom(10);
                Cell conclCell = new Cell().setBorder(Border.NO_BORDER).setPadding(12);
                conclCell.add(new Paragraph("EN CONCLUSION").setBold().setFontSize(10)
                        .setFontColor(DARK_BLUE).setMarginBottom(6));
                if (!conclusion.isBlank()) {
                    conclCell.add(new Paragraph(conclusion).setFontSize(9).setMarginBottom(8));
                }
                conclCell.add(new Paragraph(formatMoney(soldeComptable))
                        .setBold().setFontSize(16).setFontColor(MED_BLUE)
                        .setTextAlignment(TextAlignment.CENTER));
                concl.addCell(conclCell);
                doc.add(concl);
            }

            // 7. RIB / IBAN
            if (!rib.isBlank() || !iban.isBlank()) {
                Table ribTable = new Table(UnitValue.createPercentArray(new float[]{50, 50}))
                        .useAllAvailableWidth().setMarginBottom(10);
                Cell ribCell = new Cell().setBorder(Border.NO_BORDER)
                        .setBorderTop(new SolidBorder(MED_BLUE, 1)).setPadding(8);
                if (!rib.isBlank())  ribCell.add(new Paragraph(rib).setFontSize(8));
                if (!iban.isBlank()) ribCell.add(new Paragraph(iban).setFontSize(8));
                if (!bic.isBlank())  ribCell.add(new Paragraph(bic).setFontSize(8));
                ribTable.addCell(ribCell);
                ribTable.addCell(new Cell().setBorder(Border.NO_BORDER)
                        .setBorderTop(new SolidBorder(MED_BLUE, 1)));
                doc.add(ribTable);
            }

            // 8. Mentions obligatoires
            if (!mentions.isBlank()) {
                doc.add(new Paragraph(mentions)
                        .setFontSize(7).setFontColor(TEXT_MUTED)
                        .setItalic().setMarginTop(8)
                        .setBorderTop(new SolidBorder(LIGHT_GRAY, 1)).setPaddingTop(6));
            }
        }
    }

    // ── Helpers ──────────────────────────────────────────────────────────────

    private Div sectionHeader(String title) {
        return new Div()
                .setBackgroundColor(MED_BLUE)
                .add(new Paragraph(title).setFontColor(WHITE).setBold().setFontSize(9).setMargin(0))
                .setPadding(6).setMarginBottom(0);
    }

    private void addSectionRow(Table table, String code, String label, String value, boolean bold) {
        DeviceRgb bg = table.getNumberOfRows() % 2 == 0 ? WHITE : LIGHT_GRAY;
        float fs = bold ? 9.5f : 9f;

        Paragraph labelPara = new Paragraph(label).setFontSize(fs);
        Paragraph valuePara = new Paragraph(value).setFontSize(fs).setTextAlignment(TextAlignment.RIGHT);
        if (bold) { labelPara.setBold(); valuePara.setBold(); }

        table.addCell(new Cell()
                .add(new Paragraph(code).setFontSize(8).setFontColor(TEXT_MUTED))
                .setBackgroundColor(bg).setBorder(Border.NO_BORDER).setPadding(5));
        table.addCell(new Cell()
                .add(labelPara)
                .setBackgroundColor(bg).setBorder(Border.NO_BORDER).setPadding(5));
        table.addCell(new Cell()
                .add(valuePara)
                .setBackgroundColor(bg).setBorder(Border.NO_BORDER).setPadding(5));
    }

    private void addDetailRow(Table table, String label, String value) {
        table.addCell(new Cell()
                .add(new Paragraph(label).setFontSize(8).setFontColor(TEXT_MUTED).setBold())
                .setBorder(Border.NO_BORDER).setPadding(4));
        table.addCell(new Cell()
                .add(new Paragraph(value).setFontSize(9).setBold())
                .setBorder(Border.NO_BORDER).setPadding(4));
    }

    private int findMarker(Sheet sheet, String marker, int startRow,
                            DataFormatter fmt, FormulaEvaluator ev) {
        for (int r = startRow; r <= Math.min(sheet.getLastRowNum(), 200); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            org.apache.poi.ss.usermodel.Cell cell =
                    row.getCell(0, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            if (cell != null && marker.equals(fmt.formatCellValue(cell, ev).trim())) return r;
        }
        return -1;
    }

    private int findMarkerStartsWith(Sheet sheet, String prefix, int startRow,
                                      DataFormatter fmt, FormulaEvaluator ev) {
        for (int r = startRow; r <= Math.min(sheet.getLastRowNum(), 200); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            org.apache.poi.ss.usermodel.Cell cell =
                    row.getCell(0, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
            if (cell != null && fmt.formatCellValue(cell, ev).trim().startsWith(prefix)) return r;
        }
        return -1;
    }

    private int findMarkerContains(Sheet sheet, String text, int startRow,
                                    DataFormatter fmt, FormulaEvaluator ev) {
        for (int r = startRow; r <= Math.min(sheet.getLastRowNum(), 200); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            for (int c = 0; c < 5; c++) {
                org.apache.poi.ss.usermodel.Cell cell =
                        row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (cell != null && fmt.formatCellValue(cell, ev).trim().contains(text)) return r;
            }
        }
        return -1;
    }

    private String cellStr(Sheet sheet, int r, int c, DataFormatter fmt, FormulaEvaluator ev) {
        Row row = sheet.getRow(r);
        if (row == null) return "";
        org.apache.poi.ss.usermodel.Cell cell =
                row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (cell == null) return "";
        return fmt.formatCellValue(cell, ev).trim();
    }

    private double numVal(Sheet sheet, int rowIdx, int colIdx,
                           DataFormatter fmt, FormulaEvaluator ev) {
        Row row = sheet.getRow(rowIdx);
        if (row == null) return 0.0;
        org.apache.poi.ss.usermodel.Cell cell =
                row.getCell(colIdx, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (cell == null) return 0.0;
        CellType type = cell.getCellType() == CellType.FORMULA
                ? cell.getCachedFormulaResultType() : cell.getCellType();
        if (type == CellType.NUMERIC) return cell.getNumericCellValue();
        try {
            return Double.parseDouble(
                    fmt.formatCellValue(cell, ev).replace(",", ".").replace(" ", "").replace("€", "").trim());
        } catch (NumberFormatException e) {
            return 0.0;
        }
    }

    private String formatMoney(double val) {
        if (val == 0) return "€ -";
        return String.format("€ %,.2f", val)
                .replace(",", "X").replace(".", ",").replace("X", ".");
    }

    private Workbook openWorkbook(File file) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        return file.getName().toLowerCase().endsWith(".xls")
                ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
    }
}
