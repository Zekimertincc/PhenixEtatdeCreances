package com.creances;

import com.zeki.merger.model.CreanceRow;
import com.zeki.merger.service.ExcelReader;
import com.zeki.merger.service.ExcelWriter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static org.assertj.core.api.Assertions.*;

class ExcelWriterConsoTest {

    @Test
    void ftTrading_lieuIsInColumnQ() throws Exception {
        ExcelReader reader = new ExcelReader();
        ExcelWriter writer = new ExcelWriter();
        File etatFile = TestFixtures.get("ETAT DE CREANCES FT TRADING.xlsx");
        File outputFile = File.createTempFile("conso_test_", ".xlsx");
        outputFile.deleteOnExit();

        List<CreanceRow> rows = reader.readFiltered("FT TRADING", etatFile);
        assertThat(rows).as("No rows read from FT TRADING fixture").isNotEmpty();

        Map<String, List<CreanceRow>> grouped = new LinkedHashMap<>();
        grouped.put("FT TRADING", rows);
        writer.write(grouped, outputFile);

        try (Workbook wb = new XSSFWorkbook(outputFile)) {
            Sheet sheet = wb.getSheet("Consolidation");
            assertThat(sheet).isNotNull();

            // Find first data row (col A = "FT TRADING")
            Row dataRow = null;
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                Cell cellA = row.getCell(0);
                if (cellA != null
                        && cellA.getCellType() == CellType.STRING
                        && "FT TRADING".equals(cellA.getStringCellValue())) {
                    dataRow = row;
                    break;
                }
            }
            assertThat(dataRow).as("No FT TRADING data row found in Consolidation sheet").isNotNull();

            // col Q = index 16 = Lieu, must be "AG" (string)
            Cell cellQ = dataRow.getCell(16);
            assertThat(cellQ).as("col Q (index 16) must exist").isNotNull();
            assertThat(cellQ.getCellType())
                .as("col Q (Lieu) must be STRING")
                .isEqualTo(CellType.STRING);
            assertThat(cellQ.getStringCellValue())
                .as("col Q (Lieu) must be AG")
                .isEqualTo("AG");

            // col G = index 6 = DEBITEUR, must be a string (debtor name, not a number)
            Cell cellG = dataRow.getCell(6);
            assertThat(cellG).as("col G (index 6) must exist").isNotNull();
            assertThat(cellG.getCellType())
                .as("col G (DEBITEUR) must be STRING")
                .isEqualTo(CellType.STRING);
        }
    }
}
