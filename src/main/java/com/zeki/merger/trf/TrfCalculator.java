package com.zeki.merger.trf;

import com.zeki.merger.trf.model.ClientInfo;
import com.zeki.merger.trf.model.ClientSummary;
import com.zeki.merger.trf.model.ConsolidationRow;

import java.util.*;

/**
 * Computes TRF summaries from raw Consolidation rows.
 *
 * Works with the 26-col ConsolidationGenerale input format:
 *   0  = CLIENT
 *   7  = CREANCE PRINCIPALE
 *   8  = RECOUVRE ET FACTURE
 *  11  = PENALITES
 *  15  = DONT EN ATTENTE DE FACTURATION  ← CZ PHENIX source
 *  16  = Lieu  ("AG" / "CL" / "NA")      ← CZ PHENIX filter
 *  17  = Frais de procédure
 *  18  = Recouvré total
 *  19  = Déjà facturé
 *  20  = dépuis le début
 *  21  = Commissions HT
 *
 * Formulas recomputed from scratch (choses-à-faire doc):
 *   CommTTC  = CommHT * 1.2
 *   CzPhenix = SUM(DONT_EN_ATTENTE) where Lieu = "AG"
 *   Montant  = SUM((CommHT + Frais) * 1.2)
 *   Reverser = MAX(0, CzPhenix - Montant)
 */
public class TrfCalculator {

    private static final double EPS = 0.005;

    // Column indices in the 26-col ConsolidationGenerale input
    private static final int COL_CREANCE  =  7;
    private static final int COL_RECOUVRE =  8;
    private static final int COL_PENALITE = 11;
    private static final int COL_DONT     = 15; // DONT EN ATTENTE DE FACTURATION
    private static final int COL_LIEU     = 16; // Lieu: "AG"/"CL"/"NA"
    private static final int COL_FRAIS    = 17; // Frais de procédure
    private static final int COL_RECTOTAT = 18; // Recouvré total
    private static final int COL_DEJAFACT = 19; // Déjà facturé
    private static final int COL_DEPUIS   = 20; // dépuis le début
    private static final int COL_COMM_HT  = 21; // Commissions HT

    public List<ClientSummary> buildClientSummaries(
            List<ConsolidationRow> consolidationRows,
            Map<String, ClientInfo> clientInfoMap,
            Map<String, Double>     balanceMap) {

        DataReader dr = new DataReader();

        Map<String, double[]> rawSums   = new LinkedHashMap<>();
        Map<String, double[]> formulas  = new LinkedHashMap<>(); // [0]=czPhenix [1]=montant [2]=commTtc
        Map<String, String>   canonical = new LinkedHashMap<>();

        for (ConsolidationRow row : consolidationRows) {
            if (row.isHeaderRow()) continue;
            String colA = row.getString(0);
            if (colA.isBlank()) continue;

            String key = DataReader.normalize(colA);
            canonical.putIfAbsent(key, colA);

            double[] s = rawSums.computeIfAbsent(key, k -> new double[25]);
            s[COL_CREANCE]  += row.getDouble(COL_CREANCE);
            s[COL_RECOUVRE] += row.getDouble(COL_RECOUVRE);
            s[COL_PENALITE] += row.getDouble(COL_PENALITE);
            s[COL_DONT]     += row.getDouble(COL_DONT);
            s[COL_FRAIS]    += row.getDouble(COL_FRAIS);
            s[COL_RECTOTAT] += row.getDouble(COL_RECTOTAT);
            s[COL_DEJAFACT] += row.getDouble(COL_DEJAFACT);
            s[COL_DEPUIS]   += row.getDouble(COL_DEPUIS);
            s[COL_COMM_HT]  += row.getDouble(COL_COMM_HT);

            double[] f      = formulas.computeIfAbsent(key, k -> new double[3]);
            String   lieu   = row.getString(COL_LIEU).trim().toUpperCase();
            double   dont   = row.getDouble(COL_DONT);
            double   frais  = row.getDouble(COL_FRAIS);
            double   commHt = row.getDouble(COL_COMM_HT);

            if ("AG".equals(lieu)) f[0] += dont;     // CzPhenix
            f[1] += (commHt + frais) * 1.2;           // Montant
            f[2] += commHt * 1.2;                      // CommTtc
        }

        List<ClientSummary> result = new ArrayList<>();

        for (String key : rawSums.keySet()) {
            String   name = canonical.getOrDefault(key, key);
            double[] s    = rawSums.get(key);
            double[] f    = formulas.getOrDefault(key, new double[3]);

            double czPhenix = f[0];
            double montant  = f[1];
            double commTtc  = f[2];
            double reverser = czPhenix >= montant ? czPhenix - montant : 0.0;

            ClientSummary cs = new ClientSummary();
            cs.setClientName         (name);
            cs.setCreancePrincipale  (s[COL_CREANCE]);
            cs.setRecouvreEtFacture  (s[COL_RECOUVRE]);
            cs.setPenalites          (s[COL_PENALITE]);
            cs.setDontEnAttente      (s[COL_DONT]);
            cs.setFraisProcedure     (s[COL_FRAIS]);
            cs.setRecouvreTotol      (s[COL_RECTOTAT]);
            cs.setDejaFacture        (s[COL_DEJAFACT]);
            cs.setDepuisLeDebut      (s[COL_DEPUIS]);
            cs.setCommissions        (s[COL_COMM_HT]);
            cs.setCommissionTtc      (commTtc);
            cs.setPenalits           (commTtc);
            cs.setSommesCzPhenix     (czPhenix);
            cs.setMontantAFacturerTtc(montant);
            cs.setSommesAReverserSrc (reverser);

            ClientInfo ci = dr.findClientInfo(name, clientInfoMap);
            if (ci == null) continue;
            cs.setClientCode        (ci.getCode());
            cs.setIban              (ci.getIban());
            cs.setBic               (ci.getBic());
            cs.setNonCompensation   (ci.isNonCompensation());
            cs.setPaiementParCheque (ci.isPaiementParCheque());

            double prev = dr.findBalance(name, balanceMap);
            cs.setNousDoit_Prec(prev);

            if (czPhenix < EPS && montant < EPS && prev < EPS) continue;

            calculate(cs);
            result.add(cs);
        }

        result.sort(Comparator.comparing(
                cs -> cs.getClientCode().isBlank() ? "ZZZ" + cs.getClientName() : cs.getClientCode()
        ));
        return result;
    }

    public void calculate(ClientSummary cs) {
        double enc   = cs.getSommesCzPhenix();
        double mont  = cs.getMontantAFacturerTtc();
        double prev  = cs.getNousDoit_Prec();
        double total = mont + prev;
        cs.setNousDoit_Maintenant(total);

        if (cs.isNonCompensation()) {
            cs.setEncaissementsParCompensation(0.0);
            cs.setNousDoit_ApreFacturation    (total);
            cs.setSommesAReverserFinal        (enc);
            cs.setVirements                   (enc);
            cs.setEtatCompensations           ("NON COMP");
        } else {
            double comp     = Math.min(enc, Math.max(0, total));
            double reverser = Math.max(0, enc - total);
            double apre     = Math.max(0, total - comp);
            cs.setEncaissementsParCompensation(comp);
            cs.setSommesAReverserFinal        (reverser);
            cs.setNousDoit_ApreFacturation    (apre);
            cs.setVirements                   (reverser);
            cs.setEtatCompensations           (etat(cs, comp, reverser, apre));
        }
    }

    private String etat(ClientSummary cs, double comp, double reverser, double apre) {
        String label = (cs.getIban() != null && !cs.getIban().isBlank()) ? "Comp VRT" : "Comp CB";
        if (reverser > EPS) return label;
        if (comp > EPS && apre > EPS)
            return String.format(java.util.Locale.FRENCH,
                    "Comp partielle de %.2f, reste nous devoir %.2f", comp, apre);
        if (comp > EPS) return label;
        return "";
    }
}