package com.zeki.merger.trf.model;

/**
 * Per-client summary combining data from ConsolidationGenerale (Total row),
 * LISTING_CABINET_PHENIX (IBAN / NonComp), and Tableau_de_bord (previous balance).
 * Calculated TRF fields are populated by TrfCalculator.
 */
public class ClientSummary {

    // ---- Identity -----------------------------------------------------------

    private String  clientName       = "";
    private String  clientCode       = "";
    private String  iban             = "";
    private String  bic              = "";
    private boolean nonCompensation  = false;

    // ---- From ConsolidationGenerale "Total [CLIENT]" row -------------------

    private double creancePrincipale;      // col H (index  7)
    private double recouvreEtFacture;      // col I (index  8)
    private double penalites;              // col L (index 11)
    private double dontEnAttente;          // col P (index 15)
    private double fraisProcedure;         // col R (index 17)
    private double recouvreTotol;          // col S (index 18)
    private double dejaFacture;            // col T (index 19)
    private double depuisLeDebut;          // col U (index 20)
    private double commissions;            // col V (index 21)
    private double penalits;               // col W (index 22)
    private double sommesCzPhenix;         // col X (index 23)
    private double montantAFacturerTtc;    // col Y (index 24)
    private double sommesAReverserSrc;     // col Z (index 25)

    // ---- From Tableau de Bord "Soldes" sheet --------------------------------

    private double nousDoit_Prec;          // col C (index 2) — previous balance owed to Phénix

    // ---- Calculated TRF fields ----------------------------------------------

    private double nousDoit_Maintenant;            // = montantAFacturerTtc + nousDoit_Prec
    private double sommesAReverserFinal;           // = max(0, sommesCzPhenix - nousDoit_Maintenant)
    private double encaissementsParCompensation;   // = min(sommesCzPhenix, nousDoit_Maintenant)  [0 if nonComp]
    private double nousDoit_ApreFacturation;       // = max(0, nousDoit_Maintenant - encaissementsParComp)
    private String etatCompensations = "";
    private double virements;
    private double cheques;

    // ---- Getters / Setters --------------------------------------------------

    public String  getClientName()       { return clientName; }
    public void    setClientName(String v){ clientName = v != null ? v : ""; }

    public String  getClientCode()       { return clientCode; }
    public void    setClientCode(String v){ clientCode = v != null ? v : ""; }

    public String  getIban()             { return iban; }
    public void    setIban(String v)     { iban = v != null ? v : ""; }

    public String  getBic()              { return bic; }
    public void    setBic(String v)      { bic = v != null ? v : ""; }

    public boolean isNonCompensation()   { return nonCompensation; }
    public void    setNonCompensation(boolean v){ nonCompensation = v; }

    public double  getCreancePrincipale()          { return creancePrincipale; }
    public void    setCreancePrincipale(double v)  { creancePrincipale = v; }

    public double  getRecouvreEtFacture()          { return recouvreEtFacture; }
    public void    setRecouvreEtFacture(double v)  { recouvreEtFacture = v; }

    public double  getPenalites()                  { return penalites; }
    public void    setPenalites(double v)          { penalites = v; }

    public double  getDontEnAttente()              { return dontEnAttente; }
    public void    setDontEnAttente(double v)      { dontEnAttente = v; }

    public double  getFraisProcedure()             { return fraisProcedure; }
    public void    setFraisProcedure(double v)     { fraisProcedure = v; }

    public double  getRecouvreTotol()              { return recouvreTotol; }
    public void    setRecouvreTotol(double v)      { recouvreTotol = v; }

    public double  getDejaFacture()                { return dejaFacture; }
    public void    setDejaFacture(double v)        { dejaFacture = v; }

    public double  getDepuisLeDebut()              { return depuisLeDebut; }
    public void    setDepuisLeDebut(double v)      { depuisLeDebut = v; }

    public double  getCommissions()                { return commissions; }
    public void    setCommissions(double v)        { commissions = v; }

    public double  getPenalits()                   { return penalits; }
    public void    setPenalits(double v)           { penalits = v; }

    public double  getSommesCzPhenix()             { return sommesCzPhenix; }
    public void    setSommesCzPhenix(double v)     { sommesCzPhenix = v; }

    public double  getMontantAFacturerTtc()        { return montantAFacturerTtc; }
    public void    setMontantAFacturerTtc(double v){ montantAFacturerTtc = v; }

    public double  getSommesAReverserSrc()         { return sommesAReverserSrc; }
    public void    setSommesAReverserSrc(double v) { sommesAReverserSrc = v; }

    public double  getNousDoit_Prec()              { return nousDoit_Prec; }
    public void    setNousDoit_Prec(double v)      { nousDoit_Prec = v; }

    public double  getNousDoit_Maintenant()        { return nousDoit_Maintenant; }
    public void    setNousDoit_Maintenant(double v){ nousDoit_Maintenant = v; }

    public double  getSommesAReverserFinal()       { return sommesAReverserFinal; }
    public void    setSommesAReverserFinal(double v){ sommesAReverserFinal = v; }

    public double  getEncaissementsParCompensation()        { return encaissementsParCompensation; }
    public void    setEncaissementsParCompensation(double v){ encaissementsParCompensation = v; }

    public double  getNousDoit_ApreFacturation()        { return nousDoit_ApreFacturation; }
    public void    setNousDoit_ApreFacturation(double v){ nousDoit_ApreFacturation = v; }

    public String  getEtatCompensations()          { return etatCompensations; }
    public void    setEtatCompensations(String v)  { etatCompensations = v != null ? v : ""; }

    public double  getVirements()                  { return virements; }
    public void    setVirements(double v)          { virements = v; }

    public double  getCheques()                    { return cheques; }
    public void    setCheques(double v)            { cheques = v; }

    // ---- Convenience helpers ------------------------------------------------

    /** True if this client has encaissements and the invoice could be fully covered. */
    public boolean isFullyCompensated() {
        return !nonCompensation && sommesAReverserFinal >= 0
            && nousDoit_ApreFacturation < 0.005;
    }

    /** True if partial compensation was applied (some encaissements covered part of the invoice). */
    public boolean isPartiallyCompensated() {
        return !nonCompensation && encaissementsParCompensation > 0.005
            && nousDoit_ApreFacturation > 0.005;
    }

    /** True if there are no encaissements this period and the invoice is still owed. */
    public boolean isDebtor() {
        return !nonCompensation && sommesCzPhenix < 0.005
            && nousDoit_ApreFacturation > 0.005;
    }

    /** True if IBAN is present and client needs a wire transfer. */
    public boolean needsVirement() {
        return !nonCompensation && sommesAReverserFinal > 0.005;
    }

    /** True if wire transfer is needed but no IBAN is available (manual handling required). */
    public boolean needsManualVirement() {
        return needsVirement() && iban.isBlank();
    }

    /** True if wire transfer can be automated (IBAN present). */
    public boolean needsAutoVirement() {
        return needsVirement() && !iban.isBlank();
    }
}
