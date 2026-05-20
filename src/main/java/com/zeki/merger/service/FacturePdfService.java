package com.zeki.merger.service;

import com.itextpdf.kernel.colors.DeviceRgb;
import com.itextpdf.kernel.events.Event;
import com.itextpdf.kernel.events.IEventHandler;
import com.itextpdf.kernel.events.PdfDocumentEvent;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfPage;
import com.itextpdf.kernel.pdf.PdfReader;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.kernel.pdf.canvas.PdfCanvas;
import com.itextpdf.kernel.pdf.xobject.PdfFormXObject;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.borders.Border;
import com.itextpdf.layout.borders.SolidBorder;
import com.itextpdf.layout.element.Cell;
import com.itextpdf.layout.element.Div;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.properties.TextAlignment;
import com.itextpdf.layout.properties.UnitValue;
import com.zeki.merger.AppPreferences;
import com.zeki.merger.trf.DataReader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.text.Normalizer;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.function.BiConsumer;

public class FacturePdfService {

    private final FolderScanner scanner = new FolderScanner();

    private static final DeviceRgb MED_BLUE   = new DeviceRgb(0x2E, 0x75, 0xB6);
    private static final DeviceRgb DARK_BLUE  = new DeviceRgb(0x1F, 0x4E, 0x79);
    private static final DeviceRgb LIGHT_GRAY = new DeviceRgb(0xF2, 0xF2, 0xF2);
    private static final DeviceRgb WHITE      = new DeviceRgb(0xFF, 0xFF, 0xFF);
    private static final DeviceRgb TEXT_MUTED = new DeviceRgb(0x6B, 0x6B, 0x6B);

    // =========================================================================
    // Entry point
    // =========================================================================

    public List<String> apply(File rootFolder, File recupFile,
                               BiConsumer<Double, String> progress) throws Exception {
        List<String> log = new ArrayList<>();

        Map<String, String> factureMap = readFactureMap(recupFile); // col B = N° facture
        Map<String, String> nomMap     = readNomMap(recupFile);     // col D = NOM (for filename)

        List<FolderScanner.CompanyFile> companies = scanner.scan(rootFolder);
        if (companies.isEmpty()) {
            progress.accept(1.0, "Aucun dossier trouvé.");
            return log;
        }

        String mensuelPath = AppPreferences.getFacturationMensuelPath();
        File mensuelFolder = (mensuelPath != null && !mensuelPath.isBlank())
            ? new File(mensuelPath) : null;

        int total = companies.size();
        for (int i = 0; i < total; i++) {
            FolderScanner.CompanyFile cf = companies.get(i);
            double prog = 0.05 + 0.95 * (i + 1.0) / total;
            String result;
            try {
                result = processCompany(cf.excelFile(), cf.companyName(),
                                        factureMap, nomMap, mensuelFolder, recupFile);
            } catch (Exception e) {
                result = "ERREUR: " + e.getMessage();
            }
            log.add(cf.companyName() + " → " + result);
            progress.accept(prog, "[" + (i + 1) + "/" + total + "] "
                + cf.companyName() + " → " + result);
        }
        progress.accept(1.0, "Génération PDF terminée (" + total + " dossiers).");
        return log;
    }

    // =========================================================================
    // RecupNumFacture readers
    // =========================================================================

    /** Returns map of normalized client name → N° facture (col B). */
    Map<String, String> readFactureMap(File recupFile) throws IOException {
        Map<String, String> map = new LinkedHashMap<>();
        if (recupFile == null || !recupFile.exists()) return map;
        try (Workbook wb = openWorkbook(recupFile)) {
            Sheet sheet = wb.getSheet("Feuil1");
            if (sheet == null) sheet = wb.getSheetAt(0);
            DataFormatter fmt = new DataFormatter();
            FormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator();
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                String name = cellStr(row, 0, fmt, ev);
                if (name.isBlank()) break;
                String num = cellStr(row, 1, fmt, ev); // col B = N° facture
                if (!num.isBlank()) map.put(DataReader.normalize(name), num);
            }
        }
        return map;
    }

    /** Returns map of normalized client name → NOM (col D) for PDF filename. */
    Map<String, String> readNomMap(File recupFile) throws IOException {
        Map<String, String> map = new LinkedHashMap<>();
        if (recupFile == null || !recupFile.exists()) return map;
        try (Workbook wb = openWorkbook(recupFile)) {
            Sheet sheet = wb.getSheet("Feuil1");
            if (sheet == null) sheet = wb.getSheetAt(0);
            DataFormatter fmt = new DataFormatter();
            FormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator();
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                String name = cellStr(row, 0, fmt, ev); // col A = CLIENT
                if (name.isBlank()) break;
                String nom = cellStr(row, 3, fmt, ev);  // col D = NOM
                if (nom.isBlank()) nom = name;           // fallback to client name
                map.put(DataReader.normalize(name), nom);
            }
        }
        return map;
    }

    private String lookup(String clientName, Map<String, String> map) {
        if (clientName == null || clientName.isBlank() || map.isEmpty()) return "";
        String norm = DataReader.normalize(clientName);
        String v = map.get(norm);
        if (v != null) return v;
        for (Map.Entry<String, String> e : map.entrySet()) {
            String k = e.getKey();
            if (norm.contains(k) || k.contains(norm)) return e.getValue();
        }
        return "";
    }

    private boolean hasPartialMatch(String normName, Map<String, String> map) {
        if (map.containsKey(normName)) return true;
        for (String k : map.keySet()) {
            if (normName.contains(k) || k.contains(normName)) return true;
        }
        return false;
    }

    // =========================================================================
    // Per-company processing
    // =========================================================================

    private String processCompany(File excelFile, String companyName,
                                   Map<String, String> factureMap,
                                   Map<String, String> nomMap,
                                   File mensuelFolder,
                                   File recupFile) throws Exception {
        try (Workbook wb = openWorkbook(excelFile)) {
            DataFormatter fmt = new DataFormatter();
            FormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator();

            // Read client info from Créances sheet
            Sheet creances = wb.getSheet("Créances");
            String nomClient  = "";
            String codeClient = "";
            if (creances != null) {
                nomClient  = cellStr(creances, 3, 7, fmt, ev);   // H4
                String a13 = cellStr(creances, 12, 0, fmt, ev);  // A13
                codeClient = a13.length() >= 6 ? a13.substring(a13.length() - 6) : a13;
            }
            if (nomClient.isBlank()) nomClient = companyName;

            // Skip if not in factureMap (no facture number assigned this month)
            if (recupFile != null && !factureMap.isEmpty()) {
                String normClient = DataReader.normalize(nomClient);
                if (!hasPartialMatch(normClient, factureMap)) {
                    normClient = DataReader.normalize(companyName);
                    if (!hasPartialMatch(normClient, factureMap)) {
                        return "SKIP (pas de numéro de facture)";
                    }
                }
            }

            // Date from CI sheet row 14 col B (index 13, col 1)
            String dateFacture = "";
            Sheet ci = wb.getSheet("CI");
            if (ci != null) {
                Row ciRow13 = ci.getRow(13);
                if (ciRow13 != null) {
                    org.apache.poi.ss.usermodel.Cell dateCell =
                        ciRow13.getCell(1, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    if (dateCell != null) {
                        if (dateCell.getCellType() == CellType.NUMERIC
                                && DateUtil.isCellDateFormatted(dateCell)) {
                            java.time.LocalDate d = dateCell.getLocalDateTimeCellValue().toLocalDate();
                            dateFacture = "Paris, le " + d.format(
                                DateTimeFormatter.ofPattern("dd/MM/yyyy"));
                        } else {
                            dateFacture = fmt.formatCellValue(dateCell, ev).trim();
                        }
                    }
                }
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

            // N° facture from recupFile (col B), fallback to D13
            String numFacture = lookup(nomClient, factureMap);
            if (numFacture.isBlank()) numFacture = lookup(companyName, factureMap);
            if (numFacture.isBlank()) numFacture = cellStr(facture, 12, 3, fmt, ev);

            // NOM from recupFile (col D) for PDF filename
            String nom = lookup(nomClient, nomMap);
            if (nom.isBlank()) nom = lookup(companyName, nomMap);
            if (nom.isBlank()) nom = nomClient.isBlank() ? companyName : nomClient;
            String pdfName = sanitize(nom) + ".pdf";

            // Address lines from Facture sheet cols D+E rows 0-16
            List<String> adresseLines = new ArrayList<>();
            for (int r = 0; r <= 16; r++) {
                String d = cellStr(facture, r, 3, fmt, ev);
                String e = cellStr(facture, r, 4, fmt, ev);
                String line = (d + " " + e).trim();
                if (!line.isBlank()) adresseLines.add(line);
            }

            int ligneDuA = findMarker(facture, "A", 0, fmt, ev);
            int ligneDuD = ligneDuA >= 0 ? findMarker(facture, "D", ligneDuA + 3, fmt, ev) : -1;
            int ligneDuI = ligneDuD >= 0 ? findMarkerStartsWith(facture, "I", ligneDuD + 5, fmt, ev) : -1;
            int ligneConclusion = findMarkerContains(facture, "EN CONCLUSION", 0, fmt, ev);
            int ligneMentions   = findMarkerStartsWith(facture, "Mentions", 0, fmt, ev);

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
                    String v  = cellStr(facture, r, c, fmt, ev);
                    String vu = v.toUpperCase();
                    if (vu.contains("RIB")  && rib.isBlank())
                        rib  = v + " " + cellStr(facture, r, c + 1, fmt, ev);
                    if (vu.contains("IBAN") && iban.isBlank())
                        iban = v + " " + cellStr(facture, r, c + 1, fmt, ev);
                    if (vu.contains("BIC")  && bic.isBlank())
                        bic  = v + " " + cellStr(facture, r, c + 1, fmt, ev);
                }
            }

            String conclusionText = "";
            if (ligneConclusion >= 0) {
                StringBuilder sb = new StringBuilder();
                for (int r = ligneConclusion + 1;
                        r <= Math.min(ligneConclusion + 4, facture.getLastRowNum()); r++) {
                    String v = cellStr(facture, r, 0, fmt, ev);
                    if (v.isBlank()) v = cellStr(facture, r, 1, fmt, ev);
                    if (!v.isBlank()) sb.append(v).append(" ");
                }
                conclusionText = sb.toString().trim();
            }

            String mentionsText = "";
            if (ligneMentions >= 0) {
                StringBuilder sb = new StringBuilder();
                for (int r = ligneMentions;
                        r <= Math.min(ligneMentions + 5, facture.getLastRowNum()); r++) {
                    String v = cellStr(facture, r, 0, fmt, ev);
                    if (v.isBlank()) v = cellStr(facture, r, 1, fmt, ev);
                    if (!v.isBlank()) sb.append(v).append(" ");
                }
                mentionsText = sb.toString().trim();
            }

            // Determine save locations
            String etatSubfolder = determineEtatSubfolder(ag, ttc);
            List<File> saveTargets = new ArrayList<>();

            // Location 1+2 — facturation_mensuel/toutes/ and /toutes/{etat}/
            if (mensuelFolder != null && mensuelFolder.isDirectory()) {
                File toutesDir = new File(mensuelFolder, "toutes");
                toutesDir.mkdirs();
                saveTargets.add(new File(toutesDir, pdfName));
                File etatDir = new File(toutesDir, etatSubfolder);
                etatDir.mkdirs();
                saveTargets.add(new File(etatDir, pdfName));
            }

            // Location 3 — client espace partagé/factures/
            File companyDir    = excelFile.getParentFile();
            File espacePartage = findEspacePartage(companyDir);
            if (espacePartage != null) {
                File facturesDir = new File(espacePartage, "factures");
                facturesDir.mkdirs();
                saveTargets.add(new File(facturesDir, pdfName));
            }

            // Fallback — same folder as Excel
            if (saveTargets.isEmpty()) {
                saveTargets.add(new File(excelFile.getParent(), pdfName));
            }

            File primaryTarget = saveTargets.get(0);
            generatePdf(primaryTarget, nomClient, codeClient, numFacture, dateFacture,
                    adresseLines, ag, cl, agcl, comsHt, prodHt, totalHt, tva, ttc,
                    solde, retard, soldeComptable, rib, iban, bic, conclusionText, mentionsText);

            for (int t = 1; t < saveTargets.size(); t++) {
                try {
                    Files.copy(primaryTarget.toPath(), saveTargets.get(t).toPath(),
                        java.nio.file.StandardCopyOption.REPLACE_EXISTING);
                } catch (Exception ignored) {}
            }

            return "PDF → " + pdfName + " (" + saveTargets.size() + " emplacements)";
        }
    }

    // =========================================================================
    // Directory helpers
    // =========================================================================

    private String determineEtatSubfolder(double ag, double ttc) {
        if (ag <= 0.005 && ttc <= 0.005) return "debiteurs";
        if (ag <= 0.005 && ttc > 0.005)  return "non_comp";
        if (ag > 0.005  && ttc > ag)     return "comp_part";
        return "comp";
    }

    private File findEspacePartage(File companyDir) {
        if (companyDir == null) return null;
        // First: search inside company directory
        File[] subs = companyDir.listFiles(File::isDirectory);
        if (subs != null) {
            for (File d : subs) {
                String n = normalize(d.getName());
                if (n.contains("espace") && n.contains("partag")) return d;
            }
        }
        // Second: search in parent directory
        File parent = companyDir.getParentFile();
        if (parent != null) {
            File[] parSubs = parent.listFiles(File::isDirectory);
            if (parSubs != null) {
                for (File d : parSubs) {
                    String n = normalize(d.getName());
                    if (n.contains("espace") && n.contains("partag")) return d;
                }
            }
        }
        return null;
    }

    // =========================================================================
    // PDF generation
    // =========================================================================

    private void generatePdf(File pdfFile, String nomClient, String codeClient,
            String numFacture, String dateFacture, List<String> adresse,
            double ag, double cl, double agcl,
            double comsHt, double prodHt, double totalHt, double tva, double ttc,
            double solde, double retard, double soldeComptable,
            String rib, String iban, String bic,
            String conclusion, String mentions) throws Exception {

        String dateDisplay = dateFacture.isBlank()
            ? "Paris, le " + java.time.LocalDate.now().format(
                DateTimeFormatter.ofPattern("dd/MM/yyyy"))
            : dateFacture;

        // Load letterhead: first try user-configured path, then classpath resource
        InputStream lhStream = null;
        String entetePath = AppPreferences.getEntetePdfPath();
        if (entetePath != null && !entetePath.isBlank()) {
            File enteteFile = new File(entetePath);
            if (enteteFile.exists()) lhStream = new FileInputStream(enteteFile);
        }
        if (lhStream == null) {
            lhStream = getClass().getResourceAsStream("/com/zeki/merger/entete_phenix.pdf");
        }

        try (PdfWriter writer = new PdfWriter(pdfFile);
             PdfDocument pdf  = new PdfDocument(writer)) {

            // Letterhead as background layer on page 1
            if (lhStream != null) {
                try (PdfDocument lhPdf = new PdfDocument(new PdfReader(lhStream))) {
                    final PdfFormXObject lhXobj =
                        lhPdf.getFirstPage().copyAsFormXObject(pdf);
                    pdf.addEventHandler(PdfDocumentEvent.START_PAGE, new IEventHandler() {
                        private boolean done = false;
                        @Override
                        public void handleEvent(Event event) {
                            if (done) return;
                            done = true;
                            PdfDocumentEvent de = (PdfDocumentEvent) event;
                            PdfPage page = de.getPage();
                            try {
                                PdfCanvas canvas = new PdfCanvas(
                                    page.newContentStreamBefore(),
                                    page.getResources(), pdf);
                                canvas.addXObjectAt(lhXobj, 0, 0);
                                canvas.release();
                            } catch (Exception ignored) {}
                        }
                    });
                } catch (Exception ignored) {}
            }

            // top=140pt clears logo (~120pt), bottom=75pt clears footer (~60pt)
            try (Document doc = new Document(pdf, PageSize.A4)) {
                doc.setMargins(140, 50, 75, 50);

                // 1. Client address + facture details (no logo header — letterhead provides it)
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
                addDetailRow(details, "Date", dateDisplay);
                info.addCell(new Cell().add(details).setBorder(Border.NO_BORDER).setPaddingLeft(8));
                doc.add(info);

                // 2. Encaissements
                doc.add(sectionHeader("LES ENCAISSEMENTS SELON LE LIEU"));
                Table encTable = new Table(UnitValue.createPercentArray(new float[]{15, 60, 25}))
                        .useAllAvailableWidth().setMarginBottom(10);
                addSectionRow(encTable, "A",     "Encaissements Phénix (AG)",  formatMoney(ag),   false);
                addSectionRow(encTable, "B",     "Encaissements Client (CL)",  formatMoney(cl),   false);
                addSectionRow(encTable, "C=A+B", "Encaissements du mois :",    formatMoney(agcl), true);
                doc.add(encTable);

                // 3. Informations facture
                doc.add(sectionHeader("LES INFORMATIONS LIÉES À LA FACTURE"));
                Table factTable = new Table(UnitValue.createPercentArray(new float[]{15, 60, 25}))
                        .useAllAvailableWidth().setMarginBottom(10);
                addSectionRow(factTable, "D",       "COMMISSIONS HT",         formatMoney(comsHt),  false);
                addSectionRow(factTable, "E",       "FRAIS DE PROCÉDURE HT",  formatMoney(prodHt),  false);
                addSectionRow(factTable, "F=D+E",   "TOTAL HT",               formatMoney(totalHt), false);
                addSectionRow(factTable, "G=F*20%", "TVA 20,00%",             formatMoney(tva),     false);
                addSectionRow(factTable, "H=F+G",   "TOTAL TTC",              formatMoney(ttc),     true);
                doc.add(factTable);

                // 4. Versement des fonds
                doc.add(sectionHeader("LES INFORMATIONS LIÉES AU VERSEMENT DES FONDS"));
                Table versTable = new Table(UnitValue.createPercentArray(new float[]{15, 60, 25}))
                        .useAllAvailableWidth().setMarginBottom(10);
                addSectionRow(versTable, "I",
                    "Le solde des encaissements de la période est en votre faveur de :",
                    formatMoney(solde), false);
                addSectionRow(versTable, "J",
                    "Factures en retard de paiement / avoirs en cours :",
                    formatMoney(retard), false);
                addSectionRow(versTable, "K=I+J",
                    "Solde comptable en votre faveur de :",
                    formatMoney(soldeComptable), true);
                doc.add(versTable);

                // 5. EN CONCLUSION
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

                // 6. RIB / IBAN
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

                // 7. Mentions obligatoires
                if (!mentions.isBlank()) {
                    doc.add(new Paragraph(mentions)
                            .setFontSize(7).setFontColor(TEXT_MUTED)
                            .setItalic().setMarginTop(8)
                            .setBorderTop(new SolidBorder(LIGHT_GRAY, 1)).setPaddingTop(6));
                }
            }
        } finally {
            if (lhStream != null) {
                try { lhStream.close(); } catch (Exception ignored) {}
            }
        }
    }

    // ── Layout helpers ────────────────────────────────────────────────────────

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

    // ── Sheet search helpers ──────────────────────────────────────────────────

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

    // ── Cell readers ─────────────────────────────────────────────────────────

    private String cellStr(Sheet sheet, int r, int c, DataFormatter fmt, FormulaEvaluator ev) {
        Row row = sheet.getRow(r);
        if (row == null) return "";
        org.apache.poi.ss.usermodel.Cell cell =
                row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
        if (cell == null) return "";
        return fmt.formatCellValue(cell, ev).trim();
    }

    private String cellStr(Row row, int c, DataFormatter fmt, FormulaEvaluator ev) {
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
                    fmt.formatCellValue(cell, ev)
                       .replace(",", ".").replace(" ", "").replace("€", "").trim());
        } catch (NumberFormatException e) {
            return 0.0;
        }
    }

    // ── Formatting ───────────────────────────────────────────────────────────

    private String formatMoney(double val) {
        if (val == 0) return "€ -";
        return String.format("€ %,.2f", val)
                .replace(",", "X").replace(".", ",").replace("X", ".");
    }

    private static String normalize(String s) {
        return Normalizer.normalize(s, Normalizer.Form.NFD)
            .replaceAll("\\p{InCombiningDiacriticalMarks}+", "")
            .toLowerCase();
    }

    private static String sanitize(String name) {
        return name.replaceAll("[\\\\/:*?\"<>|]", "_");
    }

    private Workbook openWorkbook(File file) throws IOException {
        FileInputStream fis = new FileInputStream(file);
        return file.getName().toLowerCase().endsWith(".xls")
                ? new HSSFWorkbook(fis) : new XSSFWorkbook(fis);
    }
}
