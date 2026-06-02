package com.zeki.merger.service;

import com.zeki.merger.trf.DataReader;
import com.zeki.merger.trf.model.ClientInfo;

import java.io.File;
import java.time.LocalDate;
import java.util.*;
import java.util.stream.Collectors;

public class FacturationMailService {

    public enum CompType { VIREMENT, NON_COMP, COMP_PARTIELLE }

    // -------------------------------------------------------------------------
    // Template bodies
    // -------------------------------------------------------------------------

    public String buildBody(CompType type) {
        return switch (type) {
            case VIREMENT -> """
                    Madame, Monsieur,

                    Veuillez trouver ci-joint votre état mensuel des créances.

                    Conformément à nos accords, nous vous adressons un virement correspondant au solde comptable en votre faveur.

                    Nous restons à votre disposition pour tout renseignement complémentaire.

                    Cordialement,
                    Cabinet Phénix
                    """;
            case NON_COMP -> """
                    Madame, Monsieur,

                    Veuillez trouver ci-joint votre état mensuel des créances.

                    Nous vous rappelons que votre dossier est en mode non-compensation. \
                    Le règlement de notre facture est attendu à réception de ce courrier.

                    Nous restons à votre disposition pour tout renseignement complémentaire.

                    Cordialement,
                    Cabinet Phénix
                    """;
            case COMP_PARTIELLE -> """
                    Madame, Monsieur,

                    Veuillez trouver ci-joint votre état mensuel des créances.

                    Les encaissements du mois ont été partiellement compensés avec notre facture. \
                    Un solde reste à votre charge, dont le détail figure dans le document joint.

                    Nous restons à votre disposition pour tout renseignement complémentaire.

                    Cordialement,
                    Cabinet Phénix
                    """;
        };
    }

    // -------------------------------------------------------------------------
    // Correspondance map: normalized client name → étatPublic folder path
    // Col B (index 1) = MotClé, col C (index 2) = EspacePartagé path
    // -------------------------------------------------------------------------

    public Map<String, String> readCorrespondanceMap(File file) throws Exception {
        Map<String, String> map = new LinkedHashMap<>();
        if (file == null || !file.exists()) return map;
        try (java.io.FileInputStream fis = new java.io.FileInputStream(file)) {
            byte[] bytes = fis.readAllBytes();
            java.io.ByteArrayInputStream bais = new java.io.ByteArrayInputStream(bytes);
            org.apache.poi.ss.usermodel.Workbook wb = file.getName().toLowerCase().endsWith(".xls")
                    ? new org.apache.poi.hssf.usermodel.HSSFWorkbook(bais)
                    : new org.apache.poi.xssf.usermodel.XSSFWorkbook(bais);
            org.apache.poi.ss.usermodel.Sheet sheet = wb.getSheet("Correspondance");
            if (sheet == null) sheet = wb.getSheetAt(0);
            org.apache.poi.ss.usermodel.DataFormatter fmt = new org.apache.poi.ss.usermodel.DataFormatter();
            org.apache.poi.ss.usermodel.FormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator();
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                org.apache.poi.ss.usermodel.Row row = sheet.getRow(r);
                if (row == null) continue;
                org.apache.poi.ss.usermodel.Cell motCleCell = row.getCell(1,
                        org.apache.poi.ss.usermodel.Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                org.apache.poi.ss.usermodel.Cell pathCell = row.getCell(2,
                        org.apache.poi.ss.usermodel.Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (motCleCell == null || pathCell == null) continue;
                String motCle = fmt.formatCellValue(motCleCell, ev).trim();
                String path   = fmt.formatCellValue(pathCell,   ev).trim();
                if (!motCle.isBlank() && !path.isBlank()) {
                    map.put(DataReader.normalize(motCle), path);
                }
            }
            wb.close();
        }
        return map;
    }

    // -------------------------------------------------------------------------
    // Filter clients whose dateLastDossier falls within [from, to]
    // -------------------------------------------------------------------------

    public List<ClientInfo> filterByDateRange(Map<String, ClientInfo> clientInfoMap,
                                               LocalDate from, LocalDate to) {
        List<ClientInfo> result = new ArrayList<>();
        for (ClientInfo ci : clientInfoMap.values()) {
            LocalDate d = ci.getDateLastDossier();
            if (d != null && !d.isBefore(from) && !d.isAfter(to)) {
                result.add(ci);
            }
        }
        result.sort(Comparator.comparing(ClientInfo::getName));
        return result;
    }

    // -------------------------------------------------------------------------
    // PDF finders
    // -------------------------------------------------------------------------

    public File findLatestEtatPublic(String folderPath) {
        if (folderPath == null || folderPath.isBlank()) return null;
        File dir = new File(folderPath);
        if (!dir.isDirectory()) return null;
        return findLatestPdf(dir);
    }

    /**
     * rootFolder altında clientName'e uyan şirket klasörünü bulur,
     * içindeki "Espace partagé" → "Etat des créances" klasöründeki
     * en son PDF'i döndürür.
     */
    public File findEtatPublicForClient(String clientName, File rootFolder) {
        File companyDir = findCompanyDir(clientName, rootFolder);
        if (companyDir == null) return null;
        File[] subDirs = companyDir.listFiles(File::isDirectory);
        if (subDirs == null) return null;
        for (File d : subDirs) {
            String n = DataReader.normalize(d.getName());
            if (n.contains("espace") && n.contains("partag")) {
                File[] edcDirs = d.listFiles(File::isDirectory);
                if (edcDirs != null) {
                    for (File edc : edcDirs) {
                        String en = DataReader.normalize(edc.getName());
                        if (en.contains("etat") && en.contains("cr")) {
                            return findLatestPdf(edc);
                        }
                    }
                }
            }
        }
        return null;
    }

    /**
     * rootFolder altında clientName'e uyan şirket klasörünü bulur,
     * Espace partagé → factures/ klasöründeki en son PDF'i döndürür.
     */
    public File findFacturePdfForClient(String clientName, File rootFolder) {
        File companyDir = findCompanyDir(clientName, rootFolder);
        if (companyDir == null) return null;
        File[] subDirs = companyDir.listFiles(File::isDirectory);
        if (subDirs == null) return null;
        for (File d : subDirs) {
            String n = DataReader.normalize(d.getName());
            if (n.contains("espace") && n.contains("partag")) {
                File facturesDir = new File(d, "factures");
                if (facturesDir.isDirectory()) {
                    return findLatestPdf(facturesDir);
                }
            }
        }
        return null;
    }

    private File findCompanyDir(String clientName, File rootFolder) {
        if (rootFolder == null || !rootFolder.isDirectory()) return null;
        String normClient = DataReader.normalize(clientName);
        File[] dirs = rootFolder.listFiles(File::isDirectory);
        if (dirs == null) return null;
        for (File dir : dirs) {
            String normDir = DataReader.normalize(dir.getName());
            if (normDir.contains(normClient) || normClient.contains(normDir)
                    || (normClient.length() >= 4 && normDir.startsWith(
                            normClient.substring(0, Math.min(4, normClient.length()))))) {
                return dir;
            }
        }
        return null;
    }

    private File findLatestPdf(File folder) {
        File[] pdfs = folder.listFiles(f ->
                f.isFile() && f.getName().toLowerCase().endsWith(".pdf"));
        if (pdfs == null || pdfs.length == 0) return null;
        Arrays.sort(pdfs, (a, b) -> Long.compare(b.lastModified(), a.lastModified()));
        return pdfs[0];
    }

    // -------------------------------------------------------------------------
    // Draft folder preparation
    // -------------------------------------------------------------------------

    /**
     * Tüm draft'lar için VBS dosyaları + lancer_tous.bat oluşturur,
     * klasörü Finder/Explorer'da açar.
     * @return oluşturulan klasör path'i
     */
    public File prepareDraftFolder(List<DraftRequest> drafts) throws Exception {
        String timestamp = java.time.LocalDateTime.now()
                .format(java.time.format.DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        File draftDir = new File(System.getProperty("java.io.tmpdir"), "phenix_drafts_" + timestamp);
        draftDir.mkdirs();

        StringBuilder bat = new StringBuilder("@echo off\r\n");
        bat.append("echo Envoi des drafts vers Outlook...\r\n");

        for (DraftRequest req : drafts) {
            String safeName = req.clientName.replaceAll("[^a-zA-Z0-9]", "_");
            File vbs = new File(draftDir, "draft_" + safeName + ".vbs");

            String safeBody    = req.body.replace("\"", "\"\"")
                    .replace("\n", "\" & Chr(10) & \"");
            String safeSubject = req.subject.replace("\"", "\"\"");
            String safeTo      = req.to.replace("\"", "\"\"");

            StringBuilder vbsContent = new StringBuilder();
            vbsContent.append("Set ol = CreateObject(\"Outlook.Application\")\n");
            vbsContent.append("Set mail = ol.CreateItem(0)\n");
            vbsContent.append("mail.To = \"").append(safeTo).append("\"\n");
            vbsContent.append("mail.Subject = \"").append(safeSubject).append("\"\n");
            vbsContent.append("mail.Body = \"").append(safeBody).append("\"\n");
            vbsContent.append("mail.BCC = \"info@cabinetphenix.fr\"\n");
            if (req.attachmentPath != null && !req.attachmentPath.isBlank()) {
                vbsContent.append("mail.Attachments.Add \"").append(req.attachmentPath).append("\"\n");
            }
            vbsContent.append("mail.Save\n");

            java.nio.file.Files.writeString(vbs.toPath(), vbsContent.toString(),
                    java.nio.charset.Charset.forName("windows-1252"));

            bat.append("wscript.exe \"").append(vbs.getName()).append("\"\r\n");
            bat.append("timeout /t 1 /nobreak >nul\r\n");
        }

        bat.append("echo Termine!\r\n");
        bat.append("pause\r\n");

        File batFile = new File(draftDir, "lancer_tous.bat");
        java.nio.file.Files.writeString(batFile.toPath(), bat.toString(),
                java.nio.charset.Charset.forName("windows-1252"));

        boolean isMac = System.getProperty("os.name").toLowerCase().contains("mac");
        if (isMac) {
            Runtime.getRuntime().exec(new String[]{"open", draftDir.getAbsolutePath()});
        } else {
            Runtime.getRuntime().exec(new String[]{"explorer.exe", draftDir.getAbsolutePath()});
        }

        return draftDir;
    }

    public void cleanPreviousDraftFolder(File folder) {
        if (folder == null || !folder.exists()) return;
        File[] files = folder.listFiles();
        if (files != null) for (File f : files) f.delete();
        folder.delete();
    }

    public static class DraftRequest {
        public final String clientName;
        public final String to;
        public final String subject;
        public final String body;
        public final String attachmentPath;

        public DraftRequest(String clientName, String to, String subject,
                           String body, String attachmentPath) {
            this.clientName     = clientName;
            this.to             = to;
            this.subject        = subject;
            this.body           = body;
            this.attachmentPath = attachmentPath != null ? attachmentPath : "";
        }
    }
}
