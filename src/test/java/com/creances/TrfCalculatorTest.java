package com.creances;

import com.zeki.merger.trf.DataReader;
import com.zeki.merger.trf.TrfCalculator;
import com.zeki.merger.trf.model.ClientInfo;
import com.zeki.merger.trf.model.ClientSummary;
import com.zeki.merger.trf.model.ConsolidationRow;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import java.util.Collections;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.*;

class TrfCalculatorTest {

    static List<ClientSummary> summaries;
    static final double TOLERANCE = 0.10;

    @BeforeAll
    static void setUp() throws Exception {
        DataReader reader = new DataReader();
        TrfCalculator calculator = new TrfCalculator();

        List<ConsolidationRow> rows = reader.readAllConsolidationRows(
            TestFixtures.get("ConsolidationGenerale.xlsx"));
        Map<String, ClientInfo> clientInfo = reader.readClientInfoMap(
            TestFixtures.get("LISTING_CABINET_PHENIX.xls"));

        summaries = calculator.buildClientSummaries(rows, clientInfo, Collections.emptyMap());
    }

    private ClientSummary find(String name) {
        return summaries.stream()
            .filter(s -> s.getClientName().toUpperCase().contains(name.toUpperCase()))
            .findFirst()
            .orElseThrow(() -> new AssertionError("Client not found: " + name));
    }

    // --- Integration: structural guarantees ---

    @Test
    void buildClientSummaries_notEmpty() {
        assertThat(summaries).isNotEmpty();
    }

    @Test
    void noNullClientNames() {
        assertThat(summaries)
            .extracting(ClientSummary::getClientName)
            .doesNotContainNull()
            .doesNotContain("");
    }

    @Test
    void noDuplicateClients() {
        assertThat(summaries)
            .extracting(ClientSummary::getClientName)
            .doesNotHaveDuplicates();
    }

    @Test
    void sommesAReverserNeverNegative() {
        summaries.forEach(cs ->
            assertThat(cs.getSommesAReverserFinal())
                .as("Negatif REVERSER: " + cs.getClientName())
                .isGreaterThanOrEqualTo(-0.01)
        );
    }

    @Test
    void allClientsHaveCode() {
        summaries.forEach(cs ->
            assertThat(cs.getClientCode())
                .as("Blank code: " + cs.getClientName())
                .isNotNull()
        );
    }

    @Test
    void com2000_montantTtcPositive() {
        ClientSummary cs = find("COM 2000");
        assertThat(cs.getMontantAFacturerTtc())
            .as("COM 2000 - MONTANT A FACTURER TTC doit être positif")
            .isGreaterThan(0.0);
    }

    @Test
    void blancSas_montantTtcPositive() {
        ClientSummary cs = find("BLANC SAS");
        assertThat(cs.getMontantAFacturerTtc())
            .as("BLANC SAS - MONTANT A FACTURER TTC doit être positif")
            .isGreaterThan(0.0);
    }

    @Test
    void nousDoit_maintenantEqualsInvoicePlusPrev() {
        // nousDoit_Maintenant = montantAFacturerTtc + nousDoit_Prec (always)
        summaries.forEach(cs ->
            assertThat(cs.getNousDoit_Maintenant())
                .as("NousDoit_Maintenant incohérent: " + cs.getClientName())
                .isCloseTo(cs.getMontantAFacturerTtc() + cs.getNousDoit_Prec(), within(0.01))
        );
    }

    // --- TrfCalculator.calculate() unit tests (pure math, no file I/O) ---

    @Test
    void calculate_encGreaterThanInvoice_reverserPositive() {
        TrfCalculator calc = new TrfCalculator();
        ClientSummary cs = new ClientSummary();
        cs.setClientName("TEST");
        cs.setSommesCzPhenix(1000.0);
        cs.setMontantAFacturerTtc(600.0);
        cs.setNousDoit_Prec(0.0);

        calc.calculate(cs);

        assertThat(cs.getNousDoit_Maintenant()).isEqualTo(600.0);
        assertThat(cs.getEncaissementsParCompensation()).isEqualTo(600.0);
        assertThat(cs.getSommesAReverserFinal()).isCloseTo(400.0, within(0.01));
        assertThat(cs.getNousDoit_ApreFacturation()).isCloseTo(0.0, within(0.01));
    }

    @Test
    void calculate_encLessThanInvoice_nothingReversed() {
        TrfCalculator calc = new TrfCalculator();
        ClientSummary cs = new ClientSummary();
        cs.setClientName("TEST DEBTOR");
        cs.setSommesCzPhenix(100.0);
        cs.setMontantAFacturerTtc(300.0);
        cs.setNousDoit_Prec(0.0);

        calc.calculate(cs);

        assertThat(cs.getSommesAReverserFinal()).isCloseTo(0.0, within(0.01));
        assertThat(cs.getNousDoit_ApreFacturation()).isCloseTo(200.0, within(0.01));
    }

    @Test
    void calculate_nonCompClient_allEncaissementsReturned() {
        TrfCalculator calc = new TrfCalculator();
        ClientSummary cs = new ClientSummary();
        cs.setClientName("TEST NONCOMP");
        cs.setNonCompensation(true);
        cs.setSommesCzPhenix(500.0);
        cs.setMontantAFacturerTtc(200.0);
        cs.setNousDoit_Prec(0.0);

        calc.calculate(cs);

        assertThat(cs.getEncaissementsParCompensation()).isEqualTo(0.0);
        assertThat(cs.getSommesAReverserFinal()).isEqualTo(500.0);
        assertThat(cs.getEtatCompensations()).isEqualTo("NON COMP");
    }

    @Test
    void calculate_withPreviousBalance_addedToMaintenant() {
        TrfCalculator calc = new TrfCalculator();
        ClientSummary cs = new ClientSummary();
        cs.setClientName("TEST PREV");
        cs.setSommesCzPhenix(1000.0);
        cs.setMontantAFacturerTtc(500.0);
        cs.setNousDoit_Prec(200.0); // previous unpaid balance

        calc.calculate(cs);

        // nousDoit_Maintenant = 500 + 200 = 700
        assertThat(cs.getNousDoit_Maintenant()).isCloseTo(700.0, within(0.01));
        // enc=1000, must cover 700 → reverser = 300
        assertThat(cs.getSommesAReverserFinal()).isCloseTo(300.0, within(0.01));
    }

    @Test
    void calculate_eloPresseScenario_referenceMath() {
        // Reference values from TRF_04_2026: enc=1304.18, montant=294.75, prev=0
        TrfCalculator calc = new TrfCalculator();
        ClientSummary cs = new ClientSummary();
        cs.setClientName("ELO PRESSE");
        cs.setSommesCzPhenix(1304.18);
        cs.setMontantAFacturerTtc(294.75);
        cs.setNousDoit_Prec(0.0);

        calc.calculate(cs);

        assertThat(cs.getSommesAReverserFinal()).isCloseTo(1009.43, within(TOLERANCE));
        assertThat(cs.getNousDoit_ApreFacturation()).isCloseTo(0.0, within(TOLERANCE));
    }
}
