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
        Map<String, double[]> groupSums = new LinkedHashMap<>();
        int[] sumCols = {7, 8, 11, 15, 17, 18, 19, 20, 21, 22, 23, 24, 25};

        for (ConsolidationRow row : consolidationRows) {
            if (row.isHeaderRow()) continue;
            String colA = row.getString(0);
            if (colA.isBlank()) continue;  // group-header row (col A null, col B = label)

            double[] sums = groupSums.computeIfAbsent(colA, k -> new double[26]);
            for (int c : sumCols) {
                sums[c] += row.getDouble(c);
            }
        }

        List<ClientSummary> summaries = new ArrayList<>();

        for (Map.Entry<String, double[]> entry : groupSums.entrySet()) {
            String   clientName = entry.getKey();
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
            cs.setPenalits            (sums[22]);
            cs.setSommesCzPhenix      (sums[23]);
            cs.setMontantAFacturerTtc (sums[24]);
            cs.setSommesAReverserSrc  (sums[25]);

            // Skip clients with no activity this period
            if (cs.getSommesCzPhenix() < EPS && cs.getMontantAFacturerTtc() < EPS) continue;

            // Enrich from Listing; skip entirely if not found
            ClientInfo ci = dr.findClientInfo(clientName, clientInfoMap);
            if (ci == null) continue;   // not in Listing → skip entirely
            cs.setClientCode(ci.getCode());
            cs.setIban(ci.getIban());
            cs.setBic(ci.getBic());
            cs.setNonCompensation(ci.isNonCompensation());

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
    void calculate(ClientSummary cs) {
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
            cs.setCheques(0.0);
            cs.setEtatCompensations("NON COMP");
        } else {
            double compApplied       = Math.min(enc, Math.max(0, nousDoit_Maintenant));
            double reverserFinal     = Math.max(0, enc - nousDoit_Maintenant);
            double nousDoit_Apre     = Math.max(0, nousDoit_Maintenant - compApplied);

            cs.setEncaissementsParCompensation(compApplied);
            cs.setSommesAReverserFinal(reverserFinal);
            cs.setNousDoit_ApreFacturation(nousDoit_Apre);
            cs.setVirements(reverserFinal);
            cs.setCheques(0.0);
            cs.setEtatCompensations(determineEtat(cs, enc, compApplied, reverserFinal, nousDoit_Apre));
        }
    }

    private String determineEtat(ClientSummary cs, double enc,
                                  double compApplied, double reverserFinal, double nousDoit_Apre) {
        if (reverserFinal > EPS) {
            // Client gets money back → compensation done, virement needed
            return "Comp Vrt";
        }
        if (compApplied > EPS && nousDoit_Apre > EPS) {
            // Only partial compensation was possible
            return String.format(java.util.Locale.FRENCH,
                "Comp partielle de %.2f, reste nous devoir %.2f",
                compApplied, nousDoit_Apre);
        }
        if (compApplied > EPS) {
            // Perfectly compensated — no transfer in either direction
            return "Comp Vrt";
        }
        if (enc < EPS && cs.getNousDoit_Maintenant() > EPS) {
            // No encaissements this period, client owes Phénix
            return "";
        }
        return "";
    }
}
