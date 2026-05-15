package com.creances;

import com.zeki.merger.service.ProcreancesComparator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.FileInputStream;
import java.nio.file.Path;

import static org.assertj.core.api.Assertions.*;

class ProcreancesComparatorTest {

    @Test
    void comparatorInstantiates() {
        assertThatNoException().isThrownBy(ProcreancesComparator::new);
    }

    @Test
    void compareProducesOutput(@TempDir Path tempDir) throws Exception {
        if (!TestFixtures.exists("Procreances_04_2026.xls")
                || !TestFixtures.exists("ConsolidationGenerale.xlsx")) return;

        File procFile  = TestFixtures.get("Procreances_04_2026.xls");
        File consoFile = TestFixtures.get("ConsolidationGenerale.xlsx");

        ProcreancesComparator comp = new ProcreancesComparator();
        File result = comp.compare(procFile, consoFile, tempDir.toFile(), (d, s) -> {});

        assertThat(result).exists();
        assertThat(result.length()).isGreaterThan(0);
    }

    @Test
    void compareProducesThreeSheets(@TempDir Path tempDir) throws Exception {
        if (!TestFixtures.exists("Procreances_04_2026.xls")
                || !TestFixtures.exists("ConsolidationGenerale.xlsx")) return;

        File procFile  = TestFixtures.get("Procreances_04_2026.xls");
        File consoFile = TestFixtures.get("ConsolidationGenerale.xlsx");

        ProcreancesComparator comp = new ProcreancesComparator();
        File result = comp.compare(procFile, consoFile, tempDir.toFile(), (d, s) -> {});

        try (FileInputStream fis = new FileInputStream(result);
             XSSFWorkbook wb = new XSSFWorkbook(fis)) {
            assertThat(wb.getNumberOfSheets()).isEqualTo(3);
            assertThat(wb.getSheetName(0)).isEqualTo("Récapitulatif");
            assertThat(wb.getSheetName(1)).isEqualTo("Écarts");
            assertThat(wb.getSheetName(2)).isEqualTo("Non appariés");
        }
    }
}
