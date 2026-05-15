package com.creances;

import com.zeki.merger.service.ConsoControleComparator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.FileInputStream;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.*;

class ConsoControleComparatorTest {

    @Test
    void comparatorInstantiates() {
        assertThatNoException().isThrownBy(ConsoControleComparator::new);
    }

    @Test
    void compareProducesOutput(@TempDir Path tempDir) throws Exception {
        if (!TestFixtures.exists("Controle_Facturation.xlsx")
                || !TestFixtures.exists("ConsolidationGenerale.xlsx")) return;

        File controleFile = TestFixtures.get("Controle_Facturation.xlsx");
        File consoFile    = TestFixtures.get("ConsolidationGenerale.xlsx");

        ConsoControleComparator comp = new ConsoControleComparator();
        File result = comp.compare(controleFile, consoFile, tempDir.toFile(), (d, s) -> {});

        assertThat(result).exists();
        assertThat(result.length()).isGreaterThan(0);
    }

    @Test
    void compareProducesOneSheet(@TempDir Path tempDir) throws Exception {
        if (!TestFixtures.exists("Controle_Facturation.xlsx")
                || !TestFixtures.exists("ConsolidationGenerale.xlsx")) return;

        File controleFile = TestFixtures.get("Controle_Facturation.xlsx");
        File consoFile    = TestFixtures.get("ConsolidationGenerale.xlsx");

        ConsoControleComparator comp = new ConsoControleComparator();
        File result = comp.compare(controleFile, consoFile, tempDir.toFile(), (d, s) -> {});

        try (FileInputStream fis = new FileInputStream(result);
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {
            assertThat(wb.getNumberOfSheets()).isEqualTo(1);
            assertThat(wb.getSheetName(0)).isEqualTo("Contrôle vs Conso");
        }
    }
}
