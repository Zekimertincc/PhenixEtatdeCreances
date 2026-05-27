package com.zeki.merger;

public class Launcher {
    public static void main(String[] args) {
        // Remove XML entity size limits — required for large XLSX files (Apache POI / JAXP)
        System.setProperty("jdk.xml.maxGeneralEntitySizeLimit", "0");
        System.setProperty("jdk.xml.totalEntitySizeLimit", "0");
        System.setProperty("jdk.xml.maxXMLNameLimit", "0");
        App.main(args);
    }
}