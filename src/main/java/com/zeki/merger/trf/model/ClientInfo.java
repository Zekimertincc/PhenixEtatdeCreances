package com.zeki.merger.trf.model;

public class ClientInfo {
    private final String    name;
    private final String    code;
    private final String    nonCompensation;
    private final String    iban;
    private final String    bic;
    private final boolean   paiementParCheque;
    private final String    email;
    private final java.time.LocalDate dateLastDossier;

    public ClientInfo(String name, String code, String nonCompensation,
                      String iban, String bic) {
        this(name, code, nonCompensation, iban, bic, "", null);
    }

    public ClientInfo(String name, String code, String nonCompensation,
                      String iban, String bic, String email,
                      java.time.LocalDate dateLastDossier) {
        this.name            = name;
        this.code            = code;
        this.nonCompensation = nonCompensation;
        this.iban            = iban;
        this.bic             = bic;
        this.paiementParCheque = isNumeric(code);
        this.email           = email == null ? "" : email.trim();
        this.dateLastDossier = dateLastDossier;
    }

    private static boolean isNumeric(String s) {
        if (s == null || s.isBlank()) return false;
        try { Long.parseLong(s.trim()); return true; }
        catch (NumberFormatException e) { return false; }
    }

    public String              getName()            { return name; }
    public String              getCode()            { return code; }
    public String              getNonCompensation() { return nonCompensation; }
    public String              getIban()            { return iban; }
    public String              getBic()             { return bic; }
    public boolean             isPaiementParCheque(){ return paiementParCheque; }
    public boolean             isNonCompensation()  { return "OUI".equalsIgnoreCase(nonCompensation.trim()); }
    public String              getEmail()           { return email; }
    public java.time.LocalDate getDateLastDossier() { return dateLastDossier; }

    @Override
    public String toString() {
        return "ClientInfo{name='" + name + "', code='" + code
                + "', nonComp='" + nonCompensation + "', iban='" + iban
                + "', email='" + email + "', dateLastDossier=" + dateLastDossier + "}";
    }
}
