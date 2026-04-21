package com.zeki.merger;

/**
 * Central configuration — all constants live here so they are easy to tweak
 * without hunting through multiple files.
 */
public final class AppConfig {

    /** Name of the subfolder to look for inside each company directory. */
    public static final String TARGET_SUBFOLDER = "etat de creances";

    /** Case-insensitive prefix that the target Excel file must start with. */
    public static final String FILE_PREFIX = "etat";

    /** 0-based column index used as the filter (column S = 18). */
    public static final int FILTER_COLUMN_INDEX = 18;

    /** Human-readable label for that column (shown in the UI). */
    public static final String FILTER_COLUMN_LABEL = "S";

    /** Header text for the "company name" column prepended to every output row. */
    public static final String SOCIETE_COLUMN_HEADER = "Société";

    /** Default root folder shown in the UI on first launch. */
    public static final String DEFAULT_ROOT_PATH = "/Users/zekimertinceoglu/Dropbox/ZEKI IT";

    /** Default output folder shown in the UI on first launch. */
    public static final String DEFAULT_OUTPUT_PATH = "/Users/zekimertinceoglu/Dropbox/ZEKI IT";

    /** Name of the merged output file. */
    public static final String OUTPUT_FILENAME = "etat_creances_global.xlsx";

    /** Name of the TRF export output file (timestamp appended before extension). */
    public static final String TRF_OUTPUT_FILENAME = "trf_export.xlsx";

    /** Sheet name to look for when reading a company's source Excel file. */
    public static final String CREANCES_SHEET_NAME = "Créances";

    /** Filename prefix for Etat Public output files. */
    public static final String ETAT_PUBLIC_FILENAME_PREFIX = "L_ETAT_DE_CREANCES_";

    /** Name of the correspondance file read by EspacePartageFixer. */
    public static final String ESPACE_PARTAGE_FILENAME = "CorrespondanceClient-EspacePartage.xlsx";

    /** Suffix that must appear at the end of every EspacePartagé path (canonical form). */
    public static final String ETAT_CREANCES_SUFFIX = "\\Etat des créances";

    /**
     * When true, EspacePartageFixer overwrites the source file in place.
     * When false, it writes a sibling file with a _fixed suffix.
     */
    public static final boolean FIX_OVERWRITE = true;

    private AppConfig() {}
}
