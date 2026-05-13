package com.zeki.merger.service.data;

import java.text.Normalizer;

/**
 * String and numeric normalization utilities.
 *
 * The normalize(String) method is the canonical implementation used for fuzzy
 * client-name matching. It was duplicated in DataReader and EtatPublicGenerator —
 * those classes can delegate here in future cleanups.
 */
public class DataNormalizer {

    /**
     * Canonical name normalizer: lowercase, strip accents, collapse whitespace.
     * Used for fuzzy matching between client names across files.
     */
    public String normalize(String value) {
        if (value == null || value.isBlank()) return "";
        return Normalizer.normalize(value.trim(), Normalizer.Form.NFD)
            .replaceAll("\\p{M}", "")
            .toLowerCase()
            .replaceAll("\\s+", " ");
    }

    /** Rounds a double to 2 decimal places. */
    public double normalizeAmount(double value) {
        return Math.round(value * 100.0) / 100.0;
    }

    /**
     * Strips non-alphanumeric characters (except spaces) from a name.
     * Used to sanitize file names.
     */
    public String sanitizeFileName(String name) {
        if (name == null) return "";
        return name.replaceAll("[\\\\/:*?\"<>|]", "_");
    }

    /** Removes all non-digit and non-'+' characters from a phone string. */
    public String normalizePhoneNumber(String phone) {
        if (phone == null) return "";
        return phone.replaceAll("[^0-9+]", "");
    }

    /** Returns true if two names refer to the same entity via normalized fuzzy match. */
    public boolean fuzzyMatch(String a, String b) {
        if (a == null || b == null) return false;
        String na = normalize(a);
        String nb = normalize(b);
        return na.equals(nb) || na.contains(nb) || nb.contains(na);
    }
}
