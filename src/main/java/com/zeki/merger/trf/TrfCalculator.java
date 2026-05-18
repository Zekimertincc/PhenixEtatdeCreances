package com.zeki.merger.trf;

import com.zeki.merger.trf.model.ClientInfo;
import com.zeki.merger.trf.model.ClientSummary;
import com.zeki.merger.trf.model.ConsolidationRow;

import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * Pure business-logic layer.  Takes the raw data produced by {@link DataReader} and
 * produces a sorted list of {@link ClientSummary} objects, one per Phénix client.
 */
public class TrfCalculator {

    private static final double EPS = 0.005; // tolerance for "effectively zero"

    /**
     * Groups data rows from the ConsolidationGenerale sheet by col A (client name),
     * sums the numeric columns per group, enriches from Listing and Tableau de Bord,
     * and computes all TRF fields.
     *
     * Structure: rows where col A is blank are group-header rows (col B = label) and
     * are skipped. Rows where col A is non-blank are data rows keyed by client name.
     * Clients absent from the Listing are skipped entirely.
     *
     * @param consolidationRows all rows (including header) from the "Consolidation" sheet
     * @param clientInfoMap     map of normalised name → ClientInfo (from Listing)
     * @param balanceMap        map of normalised name → previous balance (from Tableau de Bord)
     * @return list sorted by CODE CLIENT (alphabetically), then by client name
     */
    public List<ClientSummary> buildClientSummaries(
            List<ConsolidationRow> consolidationRows,
            Map<String, ClientInfo> clientInfoMap,
            Map<String, Double>    balanceMap) {

        DataReader dr = new DataReader();

        // Accumulate per-client column sums in insertion order
        Map<String, double[]>  groupSums     = new LinkedHashMap<>();
        Map<String, String>    canonicalName = new LinkedHashMap<>();
        int[] sumCols = {7, 8, 11, 15, 17, 18, 19, 20, 21};

        for (ConsolidationRow row : consolidationRows) {
            if (row.isHeaderRow()) continue;
            String colA = row.getString(0);
            if (colA.isBlank()) continue;  // group-header row (col A null, col B = label)

            String normKey = DataReader.normalize(colA);
            canonicalName.putIfAbsent(normKey, colA);
            double[] sums = groupSums.computeIfAbsent(normKey, k -> new double[30]);
            for (int c : sumCols) {
                sums[c] += row.getDouble(c);
            }

            // Correct column mapping verified against real ConsolidationGenerale data
            double s18  = row.getDouble(18); // col S = encaissements
            double z25  = row.getDouble(25); // col Z = montant à facturer TTC (pre-computed)
            String lieu = row.getString(19).trim().toUpperCase(); // col T = AG / CL / NA

            // X (col 23) = SOMMES CZ PHENIX = SI(T="AG"; S; 0)
            if ("AG".equals(lieu)) {
                sums[23] += s18;
            }
            // Y (col 24) = MONTANT A FACTURER TTC = col Z directly
            sums[24] += z25;
        }

        List<ClientSummary> summaries = new ArrayList<>();

        for (Map.Entry<String, double[]> entry : groupSums.entrySet()) {
            String   normKey    = entry.getKey();
            String   clientName = canonicalName.getOrDefault(normKey, normKey);
            double[] sums       = entry.getValue();

            ClientSummary cs = new ClientSummary();
            cs.setClientName(clientName);
            cs.setCreancePrincipale   (sums[ 7]);
            cs.setRecouvreEtFacture   (sums[ 8]);
            cs.setPenalites           (sums[11]);
            cs.setDontEnAttente       (sums[15]);
            cs.setFraisProcedure      (sums[17]);
            cs.setRecouvreTotol       (sums[18]);
            cs.setDejaFacture         (sums[19]);
            cs.setDepuisLeDebut       (sums[20]);
            cs.setCommissions         (sums[21]);
            cs.setCommissionTtc       (sums[22]);
            cs.setPenalits            (sums[22]); // alias for compat
            cs.setSommesCzPhenix      (sums[23]);
            cs.setMontantAFacturerTtc (sums[24]);
            cs.setSommesAReverserSrc  (sums[25]);
            // SOMMES A REVERSER = max(0, X - Y)
            cs.setSommesAReverserSrc(Math.max(0, cs.getSommesCzPhenix() - cs.getMontantAFacturerTtc()));

            // Skip clients with no activity this period
            if (cs.getSommesCzPhenix() < EPS && cs.getMontantAFacturerTtc() < EPS
                    && cs.getNousDoit_Prec() < EPS) continue;

            // Enrich from Listing; skip entirely if not found
            ClientInfo ci = dr.findClientInfo(clientName, clientInfoMap);
            if (ci == null) continue;   // not in Listing → skip entirely
            cs.setClientCode(ci.getCode());
            cs.setIban(ci.getIban());
            cs.setBic(ci.getBic());
            cs.setNonCompensation(ci.isNonCompensation());
            cs.setPaiementParCheque(ci.isPaiementParCheque());

            // Previous balance from Tableau de Bord
            double prevBalance = dr.findBalance(clientName, balanceMap);
            cs.setNousDoit_Prec(prevBalance);

            // TRF calculations
            calculate(cs);

            summaries.add(cs);
        }

        // Sort by CODE CLIENT (alphabetically), fallback to client name
        summaries.sort(Comparator.comparing(
            cs -> cs.getClientCode().isBlank() ? "ZZZ" + cs.getClientName() : cs.getClientCode()
        ));

        return summaries;
    }

    // -------------------------------------------------------------------------
    // Core calculation
    // -------------------------------------------------------------------------

    /**
     * Populates all calculated fields of a ClientSummary from its input values.
     * NonComp clients skip the compensation step entirely.
     */
    public void calculate(ClientSummary cs) {
        double enc      = cs.getSommesCzPhenix();
        double montant  = cs.getMontantAFacturerTtc();
        double prevDoit = cs.getNousDoit_Prec();

        double nousDoit_Maintenant = montant + prevDoit;
        cs.setNousDoit_Maintenant(nousDoit_Maintenant);

        if (cs.isNonCompensation()) {
            // NonComp: no compensation; all encaissements must be returned to client
            cs.setEncaissementsParCompensation(0.0);
            cs.setNousDoit_ApreFacturation(nousDoit_Maintenant);   // full invoice still owed
            cs.setSommesAReverserFinal(enc);                        // all encaissements returned
            cs.setVirements(enc);
            cs.setEtatCompensations("NON COMP");
        } else {
            double compApplied       = Math.min(enc, Math.max(0, nousDoit_Maintenant));
            double reverserFinal     = Math.max(0, enc - nousDoit_Maintenant);
            double nousDoit_Apre     = Math.max(0, nousDoit_Maintenant - compApplied);

            cs.setEncaissementsParCompensation(compApplied);
            cs.setSommesAReverserFinal(reverserFinal);
            cs.setNousDoit_ApreFacturation(nousDoit_Apre);
            cs.setVirements(reverserFinal);
            cs.setEtatCompensations(determineEtat(cs, enc, compApplied, reverserFinal, nousDoit_Apre));
        }
    }

    private String determineEtat(ClientSummary cs, double enc,
                                  double compApplied, double reverserFinal, double nousDoit_Apre) {
        String compLabel = (cs.getIban() != null && !cs.getIban().isBlank()) ? "Comp VRT" : "Comp CB";
        if (reverserFinal > EPS) {
            return compLabel;
        }
        if (compApplied > EPS && nousDoit_Apre > EPS) {
            return String.format(java.util.Locale.FRENCH,
                "Comp partielle de %.2f, reste nous devoir %.2f",
                compApplied, nousDoit_Apre);
        }
        if (compApplied > EPS) {
            return compLabel;
        }
        return "";
    }
}
