package com.zeki.merger;

import java.util.prefs.Preferences;

public class AppPreferences {

    private static final Preferences PREFS =
        Preferences.userNodeForPackage(AppPreferences.class);

    private static final String KEY_MERGE_ROOT      = "merge_root_folder";
    private static final String KEY_OUTPUT_FOLDER  = "output_folder";
    private static final String KEY_TRF_CONSO      = "trf_consolidation_file";
    private static final String KEY_TRF_LISTING    = "trf_listing_file";
    private static final String KEY_TRF_TABLEAU    = "trf_tableau_file";
    private static final String KEY_PROCREANCES    = "procreancesPath";
    private static final String KEY_CONSO_COMPARE  = "consoComparePath";
    private static final String KEY_CONTROLE_PATH  = "controlePath";
    private static final String KEY_RECUP_FACTURE  = "recupFacturePath";

    public static String getMergeRoot()             { return PREFS.get(KEY_MERGE_ROOT,    ""); }
    public static void   setMergeRoot(String p)     { PREFS.put(KEY_MERGE_ROOT, p);           }

    public static String getOutputFolder()          { return PREFS.get(KEY_OUTPUT_FOLDER, ""); }
    public static void   setOutputFolder(String p)  { PREFS.put(KEY_OUTPUT_FOLDER, p);         }

    public static String getTrfConso()              { return PREFS.get(KEY_TRF_CONSO,    ""); }
    public static void   setTrfConso(String p)      { PREFS.put(KEY_TRF_CONSO, p);            }

    public static String getTrfListing()            { return PREFS.get(KEY_TRF_LISTING,  ""); }
    public static void   setTrfListing(String p)    { PREFS.put(KEY_TRF_LISTING, p);          }

    public static String getTrfTableau()             { return PREFS.get(KEY_TRF_TABLEAU,   ""); }
    public static void   setTrfTableau(String p)     { PREFS.put(KEY_TRF_TABLEAU, p);           }

    public static String getProcreancesPath()        { return PREFS.get(KEY_PROCREANCES,   ""); }
    public static void   setProcreancesPath(String p){ PREFS.put(KEY_PROCREANCES, p);           }

    public static String getConsoComparePath()        { return PREFS.get(KEY_CONSO_COMPARE, ""); }
    public static void   setConsoComparePath(String p){ PREFS.put(KEY_CONSO_COMPARE, p);         }

    public static String getControlePath()            { return PREFS.get(KEY_CONTROLE_PATH, ""); }
    public static void   setControlePath(String p)    { PREFS.put(KEY_CONTROLE_PATH, p);          }

    public static String getRecupFacturePath()        { return PREFS.get(KEY_RECUP_FACTURE, ""); }
    public static void   setRecupFacturePath(String p){ PREFS.put(KEY_RECUP_FACTURE, p);          }
}
