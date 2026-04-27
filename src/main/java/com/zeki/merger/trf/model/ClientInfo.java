package com.zeki.merger.trf.model;

public class ClientInfo {
    private final String name;
    private final String code;
    private final String nonCompensation;
    private final String iban;
    private final String bic;

    public ClientInfo(String name, String code, String nonCompensation, String iban, String bic) {
        this.name           = name           != null ? name.trim()           : "";
        this.code           = code           != null ? code.trim()           : "";
        this.nonCompensation= nonCompensation!= null ? nonCompensation.trim(): "";
        this.iban           = iban           != null ? iban.trim()           : "";
        this.bic            = bic            != null ? bic.trim()            : "";
    }

    public String getName()            { return name; }
    public String getCode()            { return code; }
    public String getNonCompensation() { return nonCompensation; }
    public String getIban()            { return iban; }
    public String getBic()             { return bic; }

    public boolean isNonCompensation() {
        return "OUI".equalsIgnoreCase(nonCompensation);
    }

    @Override
    public String toString() {
        return "ClientInfo{name='" + name + "', code='" + code
             + "', nonComp='" + nonCompensation + "', iban='" + iban + "'}";
    }
}
