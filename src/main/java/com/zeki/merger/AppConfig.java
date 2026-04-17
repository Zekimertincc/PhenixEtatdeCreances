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

    private AppConfig() {}
}
