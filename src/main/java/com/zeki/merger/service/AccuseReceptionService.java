package com.zeki.merger.service;

import com.zeki.merger.trf.DataReader;
import com.zeki.merger.trf.model.ClientInfo;

import java.io.File;
import java.time.LocalDate;
import java.util.*;
import java.util.stream.Collectors;

public class AccuseReceptionService {

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
    // Reads a simple 2-column text/CSV file: "client name;/path/to/folder"
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
                // col B (index 1) = MotClé, col C (index 2) = EspacePartagé path
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
    // Clients with null dateLastDossier are always included
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
    // Find latest état public PDF in a folder
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
        if (rootFolder == null || !rootFolder.isDirectory()) return null;
        String normClient = DataReader.normalize(clientName);

        File[] dirs = rootFolder.listFiles(File::isDirectory);
        if (dirs == null) return null;

        File bestMatch = null;
        for (File dir : dirs) {
            String normDir = DataReader.normalize(dir.getName());
            if (normDir.contains(normClient) || normClient.contains(normDir)
                    || normDir.startsWith(normClient.substring(0, Math.min(4, normClient.length())))) {
                bestMatch = dir;
                break;
            }
        }
        if (bestMatch == null) return null;

        File[] subDirs = bestMatch.listFiles(File::isDirectory);
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

    private File findLatestPdf(File folder) {
        File[] pdfs = folder.listFiles(f ->
                f.isFile() && f.getName().toLowerCase().endsWith(".pdf"));
        if (pdfs == null || pdfs.length == 0) return null;
        Arrays.sort(pdfs, (a, b) -> Long.compare(b.lastModified(), a.lastModified()));
        return pdfs[0];
    }

    // -------------------------------------------------------------------------
    // Open Outlook/Mail draft — Mac (.eml) or Windows (VBScript)
    // -------------------------------------------------------------------------

    public void openOutlookDraft(String to, String subject, String body,
                                 String attachmentPath) throws Exception {
        boolean isMac = System.getProperty("os.name").toLowerCase().contains("mac");

        if (isMac) {
            System.out.println("[DRAFT LOG] To: " + to);
            System.out.println("[DRAFT LOG] Subject: " + subject);
            System.out.println("[DRAFT LOG] Attachment: " + attachmentPath);

            String safeBody    = body.replace("\"", "\\\"").replace("\n", "\\n");
            String safeTo      = to.replace("\"", "\\\"");
            String safeSubject = subject.replace("\"", "\\\"");

            String script = "tell application \"Mail\"\n"
                    + "  set newMsg to make new outgoing message with properties"
                    + " {subject:\"" + safeSubject + "\","
                    + " content:\"" + safeBody + "\","
                    + " visible:true}\n"
                    + "  tell newMsg\n"
                    + "    make new to recipient with properties {address:\"" + safeTo + "\"}\n"
                    + (attachmentPath != null && !attachmentPath.isBlank()
                        ? "    make new attachment with properties {file name:POSIX file \"" + attachmentPath + "\"}\n"
                        : "")
                    + "  end tell\n"
                    + "end tell\n";

            System.out.println("[APPLESCRIPT]\n" + script);

            File tmpScript = File.createTempFile("draft_", ".scpt");
            java.nio.file.Files.writeString(tmpScript.toPath(), script,
                    java.nio.charset.StandardCharsets.UTF_8);
            Runtime.getRuntime().exec(new String[]{"osascript", tmpScript.getAbsolutePath()});

            new Thread(() -> {
                try { Thread.sleep(5000); tmpScript.delete(); }
                catch (Exception ignored) {}
            }).start();

        } else {
            // Windows: VBScript → Outlook draft
            String safeBody = body.replace("\"", "\"\"")
                                  .replace("\n", "\" & Chr(10) & \"");
            String vbs = "Set ol = CreateObject(\"Outlook.Application\")\n"
                       + "Set mail = ol.CreateItem(0)\n"
                       + "mail.To = \"" + to + "\"\n"
                       + "mail.Subject = \"" + subject + "\"\n"
                       + "mail.Body = \"" + safeBody + "\"\n"
                       + "mail.BCC = \"info@cabinetphenix.fr\"\n"
                       + (attachmentPath != null && !attachmentPath.isBlank()
                           ? "mail.Attachments.Add \"" + attachmentPath + "\"\n"
                           : "")
                       + "mail.Display\n";

            System.out.println("[DRAFT LOG - VBS]\n" + vbs);

            File tmpVbs = File.createTempFile("draft_", ".vbs");
            java.nio.file.Files.writeString(tmpVbs.toPath(), vbs,
                    java.nio.charset.StandardCharsets.UTF_8);

            ProcessBuilder pb = new ProcessBuilder("wscript.exe", tmpVbs.getAbsolutePath());
            pb.redirectErrorStream(true);
            Process proc = pb.start();

            new Thread(() -> {
                try {
                    String out = new String(proc.getInputStream().readAllBytes());
                    if (!out.isBlank()) System.out.println("[VBS OUT] " + out);
                    proc.waitFor();
                    Thread.sleep(3000);
                    tmpVbs.delete();
                } catch (Exception ignored) {}
            }).start();
        }
    }
}
