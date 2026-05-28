package com.zeki.merger.service;

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
import com.itextpdf.layout.borders.SolidBorder;
import com.itextpdf.layout.element.Cell;
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

    public enum Mode { OWN, CLIENT }

    private final FolderScanner scanner = new FolderScanner();

    // =========================================================================
    // Entry point
    // =========================================================================

    public List<String> apply(File rootFolder, File recupFile,
                              BiConsumer<Double, String> progress) throws Exception {
        return apply(rootFolder, recupFile, Mode.OWN, progress);
    }

    public List<String> apply(File rootFolder, File recupFile, Mode mode,
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

        Map<String, com.zeki.merger.trf.model.ClientInfo> clientInfoMap = new java.util.LinkedHashMap<>();
        String listingPath = AppPreferences.getTrfListing();
        if (listingPath != null && !listingPath.isBlank()) {
            File listingFile = new File(listingPath);
            if (listingFile.exists()) {
                try {
                    clientInfoMap = new DataReader().readClientInfoMap(listingFile);
                } catch (Exception ignored) {}
            }
        }

        int total = companies.size();
        for (int i = 0; i < total; i++) {
            FolderScanner.CompanyFile cf = companies.get(i);
            double prog = 0.05 + 0.95 * (i + 1.0) / total;
            String result;
            try {
                result = processCompany(cf.excelFile(), cf.companyName(),
                        factureMap, nomMap, mensuelFolder, recupFile, clientInfoMap, mode);
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
                                  File recupFile,
                                  Map<String, com.zeki.merger.trf.model.ClientInfo> clientInfoMap,
                                  Mode mode) throws Exception {
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

            // Date from CI sheet row 14 col B — cell is =TODAY() formula, always format as dd/MM/yyyy
            String dateFacture = "";
            Sheet ci = wb.getSheet("CI");
            if (ci != null) {
                Row ciRow13 = ci.getRow(13);
                if (ciRow13 != null) {
                    org.apache.poi.ss.usermodel.Cell dateCell =
                            ciRow13.getCell(1, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    if (dateCell != null) {
                        CellType ct = dateCell.getCellType() == CellType.FORMULA
                                ? dateCell.getCachedFormulaResultType() : dateCell.getCellType();
                        if (ct == CellType.NUMERIC) {
                            try {
                                java.time.LocalDate d = dateCell.getLocalDateTimeCellValue().toLocalDate();
                                dateFacture = "Paris, le " + d.format(
                                        DateTimeFormatter.ofPattern("dd/MM/yyyy"));
                            } catch (Exception ignored) {}
                        }
                        if (dateFacture.isBlank()) {
                            // fallback: try to parse whatever string is there
                            String raw = fmt.formatCellValue(dateCell, ev).trim();
                            if (!raw.isBlank()) dateFacture = "Paris, le " + raw;
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
            for (int r = 4; r <= 12; r++) {
                String e = cellStr(facture, r, 4, fmt, ev);
                if (!e.isBlank()) adresseLines.add(e);
            }
            // Débiteur rows: row 17 = header (index 16), data starts row 18 (index 17)
            List<Object[]> debiteurRows = new ArrayList<>();
            for (int r = 17; r <= facture.getLastRowNum(); r++) {
                Row row = facture.getRow(r);
                if (row == null) break;
                org.apache.poi.ss.usermodel.Cell firstCell =
                        row.getCell(0, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (firstCell == null || fmt.formatCellValue(firstCell, ev).isBlank()) break;
                Object[] dr = new Object[7];
                for (int c = 0; c < 7; c++) {
                    org.apache.poi.ss.usermodel.Cell cell =
                            row.getCell(c, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                    if (cell == null) { dr[c] = ""; continue; }
                    CellType ct = cell.getCellType() == CellType.FORMULA
                            ? cell.getCachedFormulaResultType() : cell.getCellType();
                    // cols 3,4,5 = Encaissements, Commissions, Frais — always numeric amounts, never dates
                    boolean isMoneyCol = (c == 3 || c == 4 || c == 5);
                    if (ct == CellType.NUMERIC && isMoneyCol) {
                        double v = cell.getNumericCellValue();
                        dr[c] = formatMoney(v);
                    } else if (ct == CellType.NUMERIC) {
                        double v = cell.getNumericCellValue();
                        dr[c] = String.valueOf((long) v);
                    } else {
                        dr[c] = fmt.formatCellValue(cell, ev).trim();
                    }
                }
                debiteurRows.add(dr);
            }

            int ligneDuA = findMarker(facture, "A", 0, fmt, ev);
            int ligneDuD = ligneDuA >= 0 ? findMarker(facture, "D", ligneDuA + 3, fmt, ev) : -1;
            // I row exists only in COMP; in NON COMP the versement section starts at J
            int ligneDuI = ligneDuD >= 0 ? findMarkerStartsWith(facture, "I", ligneDuD + 5, fmt, ev) : -1;
            int ligneDuJ = -1;
            if (ligneDuI >= 0) {
                ligneDuJ = findMarkerStartsWith(facture, "J", ligneDuI + 1, fmt, ev);
            } else if (ligneDuD >= 0) {
                // NON COMP: no I row, J comes right after VERSEMENT header
                ligneDuJ = findMarkerStartsWith(facture, "J", ligneDuD + 5, fmt, ev);
            }
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

            // COMP: solde = I (encaissement - TTC), retard = J, soldeComptable = K
            // NON COMP: no I row, retard = J, soldeComptable = K (= TTC)
            double solde          = ligneDuI >= 0 ? numVal(facture, ligneDuI, 2, fmt, ev) : 0;
            double retard         = ligneDuJ >= 0 ? numVal(facture, ligneDuJ, 2, fmt, ev) : 0;
            int    ligneK         = ligneDuJ >= 0 ? ligneDuJ + 1 : -1;
            double soldeComptable = ligneK   >= 0 ? numVal(facture, ligneK,   2, fmt, ev) : 0;

            // NON COMP also has L row: "Aussitôt que nous aurons reçu..."
            int    ligneL         = -1;
            double montantVerse   = 0;
            String labelL         = "";
            if (ligneDuI < 0 && ligneK >= 0) {
                // try to find L row after K
                ligneL = findMarkerStartsWith(facture, "L", ligneK + 1, fmt, ev);
                if (ligneL >= 0) {
                    montantVerse = numVal(facture, ligneL, 2, fmt, ev);
                    labelL = cellStr(facture, ligneL, 1, fmt, ev);
                }
            }

            // RIB/IBAN/BIC — col C (index 2) in the sheet
            String rib = "", iban = "", bic = "";
            for (int r = 0; r <= facture.getLastRowNum(); r++) {
                String c2 = cellStr(facture, r, 2, fmt, ev);
                if (c2.toUpperCase().contains("RIB")  && rib.isBlank())  rib  = c2;
                if (c2.toUpperCase().contains("IBAN") && iban.isBlank()) iban = c2;
                if (c2.toUpperCase().contains("BIC")  && bic.isBlank())  bic  = c2;
            }

            String labelI = ligneDuI >= 0 ? cellStr(facture, ligneDuI, 1, fmt, ev) : "";
            String labelJ = ligneDuJ >= 0 ? cellStr(facture, ligneDuJ, 1, fmt, ev) : "";
            String labelK = ligneK   >= 0 ? cellStr(facture, ligneK,   1, fmt, ev) : "";
            if (labelI.isBlank()) labelI = "Le solde des encaissements de la période est en votre faveur de :";
            if (labelJ.isBlank()) labelJ = "Il n'y a pas de factures en retard de paiement ni d'avoirs en cours";
            if (labelK.isBlank()) labelK = "Solde comptable en votre faveur de :";
            if (labelL.isBlank()) labelL = "Aussitôt que nous aurons reçu votre règlement, nous vous ferons parvenir les sommes recouvrées de :";

            // Header text from row 16 (row index 15), col A
            String headerText = cellStr(facture, 15, 0, fmt, ev);

            String conclusionText = "";
            if (ligneConclusion >= 0) {
                StringBuilder sb = new StringBuilder();
                for (int r = ligneConclusion + 1;
                     r <= Math.min(ligneConclusion + 4, facture.getLastRowNum()); r++) {
                    String v = cellStr(facture, r, 0, fmt, ev);
                    if (v.isBlank()) v = cellStr(facture, r, 1, fmt, ev);
                    // Stop if we hit the RIB/Pour section
                    if (v.toLowerCase().contains("virement") || v.toLowerCase().contains("rib")
                            || v.toLowerCase().contains("iban") || v.toLowerCase().contains("mention")) break;
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

            // Determine comp/non-comp
            DataReader drReader = new DataReader();
            com.zeki.merger.trf.model.ClientInfo ciInfo = drReader.findClientInfo(nomClient, clientInfoMap);
            if (ciInfo == null) ciInfo = drReader.findClientInfo(companyName, clientInfoMap);
            // NON COMP = no I row (no virement to client), or flagged in listing
            boolean isNonComp = (ciInfo != null && ciInfo.isNonCompensation()) || (ligneDuI < 0 && ttc > 0.005);
            String etatSubfolder = isNonComp ? "non_comp" : determineEtatSubfolder(ag, ttc);
            List<File> saveTargets = new ArrayList<>();

            if (mode == Mode.OWN) {
                // Mode OWN — sadece kendi klasörlerimize: facturation_mensuel/toutes/{etat}/
                if (mensuelFolder != null && mensuelFolder.isDirectory()) {
                    File toutesDir = new File(mensuelFolder, "toutes");
                    toutesDir.mkdirs();
                    for (String folder : new String[]{"comp", "non_comp", "comp_part", "debiteurs"}) {
                        new File(toutesDir, folder).mkdirs();
                    }
                    File etatDir = new File(toutesDir, etatSubfolder);
                    saveTargets.add(new File(etatDir, pdfName));
                }
                // Fallback
                if (saveTargets.isEmpty()) {
                    saveTargets.add(new File(excelFile.getParent(), pdfName));
                }
            } else {
                // Mode CLIENT — sadece client espace partagé/factures/
                File companyDir    = excelFile.getParentFile();
                File espacePartage = findEspacePartage(companyDir);
                if (espacePartage != null) {
                    File facturesDir = new File(espacePartage, "factures");
                    facturesDir.mkdirs();
                    saveTargets.add(new File(facturesDir, pdfName));
                }
                // Fallback
                if (saveTargets.isEmpty()) {
                    saveTargets.add(new File(excelFile.getParent(), pdfName));
                }
            }

            File primaryTarget = saveTargets.get(0);
            generatePdf(primaryTarget, nomClient, codeClient, numFacture, dateFacture,
                    adresseLines, debiteurRows, ag, cl, agcl, comsHt, prodHt, totalHt, tva, ttc,
                    solde, retard, soldeComptable, montantVerse, rib, iban, bic,
                    conclusionText, mentionsText, headerText,
                    labelI, labelJ, labelK, labelL, isNonComp);

            for (int t = 1; t < saveTargets.size(); t++) {
                try {
                    Files.copy(primaryTarget.toPath(), saveTargets.get(t).toPath(),
                            java.nio.file.StandardCopyOption.REPLACE_EXISTING);
                } catch (Exception ignored) {}
            }

            return "PDF → " + pdfName + " [" + (isNonComp ? "NON COMP" : "COMP") + "] ["
                    + (mode == Mode.OWN ? "nos dossiers" : "espace partagé") + "]";
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
                             String numFacture, String dateFacture, List<String> adresse, List<Object[]> debiteurRows,
                             double ag, double cl, double agcl,
                             double comsHt, double prodHt, double totalHt, double tva, double ttc,
                             double solde, double retard, double soldeComptable, double montantVerse,
                             String rib, String iban, String bic,
                             String conclusion, String mentions, String headerText,
                             String labelI, String labelJ, String labelK, String labelL,
                             boolean isNonComp) throws Exception {

        String dateDisplay = dateFacture.isBlank()
                ? "Paris, le " + java.time.LocalDate.now().format(
                DateTimeFormatter.ofPattern("dd/MM/yyyy")): dateFacture;

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

            // top=130pt clears logo, bottom=65pt clears footer
            try (Document doc = new Document(pdf, PageSize.A4)) {
                doc.setMargins(130, 35, 65, 35);

                // 1. Date — left
                doc.add(new Paragraph(dateDisplay).setFontSize(9).setMarginBottom(2));

                // 2. Address — right-aligned
                if (!adresse.isEmpty()) {
                    Paragraph addrPara = new Paragraph()
                            .setTextAlignment(TextAlignment.RIGHT)
                            .setFontSize(9).setMarginBottom(6);
                    for (int ai = 0; ai < adresse.size(); ai++) {
                        addrPara.add(adresse.get(ai));
                        if (ai < adresse.size() - 1) addrPara.add("\n");
                    }
                    doc.add(addrPara);
                }

                // 3. FACTURE N° — bordered 2-column table
                Table factureHeader = new Table(UnitValue.createPercentArray(new float[]{40, 60}))
                        .useAllAvailableWidth().setMarginBottom(4);
                factureHeader.addCell(new Cell()
                        .add(new Paragraph("FACTURE N°").setFontSize(9).setBold())
                        .setBorder(new SolidBorder(1)).setPadding(2));
                factureHeader.addCell(new Cell()
                        .add(new Paragraph(numFacture.isBlank() ? "—" : numFacture)
                                .setFontSize(9).setBold())
                        .setBorder(new SolidBorder(1)).setPadding(2));
                doc.add(factureHeader);

                // 4. Code client — plain paragraph
                doc.add(new Paragraph("Code client : " + (codeClient.isBlank() ? "—" : codeClient))
                        .setFontSize(9).setMarginBottom(4));

                // 5. Header text + Débiteur table
                if (!headerText.isBlank()) {
                    doc.add(new Paragraph(headerText).setFontSize(8).setItalic().setMarginBottom(3));
                }
                if (!debiteurRows.isEmpty()) {
                    String[] debHeaders = {"V/REF", "N/REF", "Débiteur",
                            "Encaissements", "Commissions", "Frais de procédure", "Lieu"};
                    float[] debWidths = {10, 10, 25, 15, 15, 15, 10};
                    // cols 3,4,5 = money → right aligned; others left
                    boolean[] rightAlign = {false, false, false, true, true, true, true};
                    Table debTable = new Table(UnitValue.createPercentArray(debWidths))
                            .useAllAvailableWidth().setMarginBottom(4);
                    for (int hi = 0; hi < debHeaders.length; hi++) {
                        TextAlignment ta = rightAlign[hi] ? TextAlignment.RIGHT : TextAlignment.LEFT;
                        debTable.addHeaderCell(new Cell()
                                .add(new Paragraph(debHeaders[hi]).setFontSize(7).setBold()
                                        .setTextAlignment(ta))
                                .setTextAlignment(ta)
                                .setBorder(new SolidBorder(1)).setPadding(2));
                    }
                    for (Object[] dr : debiteurRows) {
                        for (int c = 0; c < 7; c++) {
                            String v = (c < dr.length && dr[c] != null) ? dr[c].toString() : "";
                            TextAlignment ta = rightAlign[c] ? TextAlignment.RIGHT : TextAlignment.LEFT;
                            debTable.addCell(new Cell()
                                    .add(new Paragraph(v).setFontSize(7).setTextAlignment(ta))
                                    .setTextAlignment(ta)
                                    .setBorder(new SolidBorder(1)).setPadding(2));
                        }
                    }
                    doc.add(debTable);
                }

                // 6. Encaissements
                doc.add(borderedSectionHeader("LES ENCAISSEMENTS SELON LE LIEU"));
                Table encTable = new Table(UnitValue.createPercentArray(new float[]{15, 60, 25}))
                        .useAllAvailableWidth().setMarginBottom(4);
                addBorderedRow(encTable, "A",     "Encaissements Phénix (AG)",  formatMoney(ag),   false);
                addBorderedRow(encTable, "B",     "Encaissements Client (CL)",  formatMoney(cl),   false);
                addBorderedRow(encTable, "C=A+B", "Encaissements du mois :",    formatMoney(agcl), true);
                doc.add(encTable);

                // 7. Facture
                doc.add(borderedSectionHeader("LES INFORMATIONS LIÉES À LA FACTURE"));
                Table factTable = new Table(UnitValue.createPercentArray(new float[]{15, 60, 25}))
                        .useAllAvailableWidth().setMarginBottom(4);
                addBorderedRow(factTable, "D",       "COMMISSIONS HT",         formatMoney(comsHt),  false);
                addBorderedRow(factTable, "E",       "FRAIS DE PROCÉDURE HT",  formatMoney(prodHt),  false);
                addBorderedRow(factTable, "F=D+E",   "TOTAL HT",               formatMoney(totalHt), false);
                addBorderedRow(factTable, "G=F*20%", "TVA 20,00%",             formatMoney(tva),     false);
                addBorderedRow(factTable, "H=F+G",   "TOTAL TTC",              formatMoney(ttc),     true);
                doc.add(factTable);

                // 8. Versement — different structure for COMP vs NON COMP
                doc.add(borderedSectionHeader("LES INFORMATIONS LIÉES AU VERSEMENT DES FONDS"));
                Table versTable = new Table(UnitValue.createPercentArray(new float[]{15, 60, 25}))
                        .useAllAvailableWidth().setMarginBottom(4);

                if (!isNonComp) {
                    // COMP: I = solde encaissements - TTC, J = retard, K = solde comptable
                    addBorderedRow(versTable, "I=A-H", labelI, formatMoney(solde),          false);
                    addBorderedRow(versTable, "J",     labelJ, formatMoney(retard),         false);
                    addBorderedRow(versTable, "K=I+J", labelK, formatMoney(soldeComptable), true);
                } else {
                    // NON COMP: J = retard (pas de I), K = solde (= TTC), L = montant à verser après règlement
                    addBorderedRow(versTable, "J",     labelJ, formatMoney(retard),         false);
                    addBorderedRow(versTable, "K=H+J", labelK, formatMoney(soldeComptable), true);
                    if (montantVerse > 0.005) {
                        addBorderedRow(versTable, "L=A",  labelL, formatMoney(montantVerse), false);
                    }
                }
                doc.add(versTable);

                // 9. EN CONCLUSION — only for COMP
                if (!isNonComp && (!conclusion.isBlank() || soldeComptable != 0)) {
                    Table concl = new Table(UnitValue.createPercentArray(new float[]{60, 40}))
                            .useAllAvailableWidth().setMarginBottom(4);
                    String conclText = conclusion.isBlank()
                            ? "Nous avons le plaisir de vous envoyer un règlement correspondant au solde comptable de :"
                            : conclusion;
                    Cell leftCell = new Cell().setBorder(new SolidBorder(1)).setPadding(4);
                    leftCell.add(new Paragraph("EN CONCLUSION").setFontSize(8).setBold().setMarginBottom(3));
                    leftCell.add(new Paragraph(conclText).setFontSize(8));
                    concl.addCell(leftCell);
                    concl.addCell(new Cell()
                            .add(new Paragraph(formatMoney(soldeComptable))
                                    .setBold().setFontSize(14)
                                    .setTextAlignment(TextAlignment.CENTER))
                            .setBorder(new SolidBorder(1)).setPadding(4)
                            .setVerticalAlignment(
                                    com.itextpdf.layout.properties.VerticalAlignment.MIDDLE));
                    doc.add(concl);
                }

                // 10. RIB/IBAN/BIC — 2-column bordered table
                if (!rib.isBlank() || !iban.isBlank()) {
                    Table ribTable = new Table(UnitValue.createPercentArray(new float[]{50, 50}))
                            .useAllAvailableWidth().setMarginBottom(4);
                    ribTable.addCell(new Cell()
                            .add(new Paragraph("Pour tout règlement par virement bancaire")
                                    .setFontSize(7.5f).setItalic())
                            .setBorder(new SolidBorder(1)).setPadding(4));
                    Cell ribCell = new Cell().setBorder(new SolidBorder(1)).setPadding(4);
                    if (!rib.isBlank())  ribCell.add(new Paragraph(rib).setFontSize(7.5f));
                    if (!iban.isBlank()) ribCell.add(new Paragraph(iban).setFontSize(7.5f));
                    if (!bic.isBlank())  ribCell.add(new Paragraph(bic).setFontSize(7.5f));
                    ribTable.addCell(ribCell);
                    doc.add(ribTable);
                }

                // 11. Mentions
                if (!mentions.isBlank()) {
                    doc.add(new Paragraph(mentions)
                            .setFontSize(6).setItalic().setMarginTop(4));
                }
            }
        } finally {
            if (lhStream != null) {
                try { lhStream.close(); } catch (Exception ignored) {}
            }
        }
    }

    // ── Layout helpers ────────────────────────────────────────────────────────

    private Table borderedSectionHeader(String title) {
        Table t = new Table(1).useAllAvailableWidth().setMarginTop(4).setMarginBottom(0);
        t.addCell(new Cell()
                .add(new Paragraph(title).setFontSize(8).setBold().setMargin(0))
                .setBorder(new SolidBorder(1)).setPadding(2));
        return t;
    }

    private void addBorderedRow(Table table, String code, String label, String value, boolean bold) {
        float fs = bold ? 8f : 7.5f;
        Paragraph labelPara = new Paragraph(label).setFontSize(fs);
        Paragraph valuePara = new Paragraph(value).setFontSize(fs)
                .setTextAlignment(TextAlignment.RIGHT);
        if (bold) { labelPara.setBold(); valuePara.setBold(); }
        table.addCell(new Cell()
                .add(new Paragraph(code).setFontSize(7))
                .setBorder(new SolidBorder(1)).setPadding(2));
        table.addCell(new Cell()
                .add(labelPara)
                .setBorder(new SolidBorder(1)).setPadding(2));
        table.addCell(new Cell()
                .add(valuePara)
                .setTextAlignment(TextAlignment.RIGHT)
                .setBorder(new SolidBorder(1)).setPadding(2));
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
        java.text.NumberFormat nf = java.text.NumberFormat.getNumberInstance(java.util.Locale.FRENCH);
        nf.setMinimumFractionDigits(2);
        nf.setMaximumFractionDigits(2);
        String formatted = nf.format(val).replace("\u202F", "\u00A0")
                .replace("\u0020", "\u00A0");
        return "€\u00A0" + formatted;
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