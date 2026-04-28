package com.zeki.merger;

import java.util.prefs.Preferences;

public class AppPreferences {

    private static final Preferences PREFS =
        Preferences.userNodeForPackage(AppPreferences.class);

    private static final String KEY_MERGE_ROOT    = "merge_root_folder";
    private static final String KEY_OUTPUT_FOLDER = "output_folder";
    private static final String KEY_TRF_CONSO     = "trf_consolidation_file";
    private static final String KEY_TRF_LISTING   = "trf_listing_file";
    private static final String KEY_TRF_TABLEAU   = "trf_tableau_file";

    public static String getMergeRoot()             { return PREFS.get(KEY_MERGE_ROOT,    ""); }
    public static void   setMergeRoot(String p)     { PREFS.put(KEY_MERGE_ROOT, p);           }

    public static String getOutputFolder()          { return PREFS.get(KEY_OUTPUT_FOLDER, ""); }
    public static void   setOutputFolder(String p)  { PREFS.put(KEY_OUTPUT_FOLDER, p);         }

    public static String getTrfConso()              { return PREFS.get(KEY_TRF_CONSO,    ""); }
    public static void   setTrfConso(String p)      { PREFS.put(KEY_TRF_CONSO, p);            }

    public static String getTrfListing()            { return PREFS.get(KEY_TRF_LISTING,  ""); }
    public static void   setTrfListing(String p)    { PREFS.put(KEY_TRF_LISTING, p);          }

    public static String getTrfTableau()            { return PREFS.get(KEY_TRF_TABLEAU,  ""); }
    public static void   setTrfTableau(String p)    { PREFS.put(KEY_TRF_TABLEAU, p);          }
}
