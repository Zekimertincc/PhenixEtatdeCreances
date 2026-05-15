package com.creances;

import com.zeki.merger.trf.DataReader;
import com.zeki.merger.trf.model.ClientInfo;
import com.zeki.merger.trf.model.ConsolidationRow;
import org.junit.jupiter.api.Test;

import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.*;

class DataReaderTest {

    DataReader reader = new DataReader();

    @Test
    void readConsolidationRows_notEmpty() throws Exception {
        List<ConsolidationRow> rows = reader.readAllConsolidationRows(
            TestFixtures.get("ConsolidationGenerale.xlsx"));
        assertThat(rows).isNotEmpty();
    }

    @Test
    void readConsolidationRows_firstRowIsHeader() throws Exception {
        List<ConsolidationRow> rows = reader.readAllConsolidationRows(
            TestFixtures.get("ConsolidationGenerale.xlsx"));
        assertThat(rows.get(0).isHeaderRow()).isTrue();
    }

    @Test
    void readConsolidationRows_hasDataRows() throws Exception {
        List<ConsolidationRow> rows = reader.readAllConsolidationRows(
            TestFixtures.get("ConsolidationGenerale.xlsx"));
        long dataRows = rows.stream().filter(r -> !r.isHeaderRow()).count();
        assertThat(dataRows).isGreaterThan(0);
    }

    @Test
    void readClientInfoMap_notEmpty() throws Exception {
        Map<String, ClientInfo> map = reader.readClientInfoMap(
            TestFixtures.get("LISTING_CABINET_PHENIX.xls"));
        assertThat(map).isNotEmpty();
    }

    @Test
    void readClientInfoMap_noBlankNames() throws Exception {
        Map<String, ClientInfo> map = reader.readClientInfoMap(
            TestFixtures.get("LISTING_CABINET_PHENIX.xls"));
        map.values().forEach(ci ->
            assertThat(ci.getName()).isNotBlank()
        );
    }

    @Test
    void readClientInfoMap_hasIbans() throws Exception {
        Map<String, ClientInfo> map = reader.readClientInfoMap(
            TestFixtures.get("LISTING_CABINET_PHENIX.xls"));
        long withIban = map.values().stream()
            .filter(ci -> !ci.getIban().isBlank())
            .count();
        assertThat(withIban).isGreaterThan(0);
    }

    @Test
    void readClientInfoMap_hasCodes() throws Exception {
        Map<String, ClientInfo> map = reader.readClientInfoMap(
            TestFixtures.get("LISTING_CABINET_PHENIX.xls"));
        long withCode = map.values().stream()
            .filter(ci -> !ci.getCode().isBlank())
            .count();
        assertThat(withCode).isGreaterThan(0);
    }

    @Test
    void normalize_removesAccents() {
        assertThat(DataReader.normalize("PHÉNIX")).isEqualTo("phenix");
        assertThat(DataReader.normalize("  Blanc SAS  ")).isEqualTo("blanc sas");
        assertThat(DataReader.normalize(null)).isEqualTo("");
    }

    @Test
    void findClientInfo_exactMatch() throws Exception {
        Map<String, ClientInfo> map = reader.readClientInfoMap(
            TestFixtures.get("LISTING_CABINET_PHENIX.xls"));
        // Just verify lookup doesn't throw; result may be null if name not in listing
        String firstName = map.values().iterator().next().getName();
        ClientInfo found = reader.findClientInfo(firstName, map);
        assertThat(found).isNotNull();
    }

    @Test
    void findBalance_returnsZeroForUnknownClient() throws Exception {
        Map<String, Double> emptyMap = Map.of();
        assertThat(reader.findBalance("UNKNOWN CLIENT XYZ", emptyMap)).isEqualTo(0.0);
    }
}
