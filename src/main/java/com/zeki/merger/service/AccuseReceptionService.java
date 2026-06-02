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
        List<String> lines = java.nio.file.Files.readAllLines(file.toPath(),
                java.nio.charset.StandardCharsets.UTF_8);
        for (String line : lines) {
            line = line.trim();
            if (line.isBlank() || line.startsWith("#")) continue;
            String[] parts = line.split(";", 2);
            if (parts.length < 2) continue;
            String name   = parts[0].trim();
            String folder = parts[1].trim();
            if (!name.isBlank() && !folder.isBlank()) {
                map.put(DataReader.normalize(name), folder);
            }
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
        File[] pdfs = dir.listFiles(f ->
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
                           ? "mail.Attachments.Add \"" + attachmentPath.replace("\\", "\\\\") + "\"\n"
                           : "")
                       + "mail.Display\n";

            System.out.println("[DRAFT LOG - VBS]\n" + vbs);

            File tmpVbs = File.createTempFile("draft_", ".vbs");
            java.nio.file.Files.writeString(tmpVbs.toPath(), vbs,
                    java.nio.charset.StandardCharsets.UTF_8);
            Runtime.getRuntime().exec(new String[]{"wscript", tmpVbs.getAbsolutePath()});

            new Thread(() -> {
                try { Thread.sleep(5000); tmpVbs.delete(); }
                catch (Exception ignored) {}
            }).start();
        }
    }
}
