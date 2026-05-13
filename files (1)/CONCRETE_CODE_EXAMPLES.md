# REFACTORING BAŞLANGICI - CONCRETE CODE EXAMPLES

## ADIM 1: Yeni Package Yapısı Oluştur

```
src/main/java/com/zeki/merger/
├── core/
│   ├── config/
│   │   └── ApplicationConfig.java
│   ├── exception/
│   │   ├── BusinessException.java
│   │   └── ErrorCode.java
│   └── pattern/
│       ├── strategy/
│       ├── factory/
│       ├── builder/
│       ├── command/
│       └── observer/
├── service/
│   ├── excel/
│   │   ├── ExcelStyleFactory.java      [YENİ]
│   │   ├── ExcelSheetBuilder.java      [YENİ]
│   │   └── ExcelFormatterService.java  [YENİ]
│   ├── report/
│   │   ├── ReportGenerator.java        [NEW FACADE]
│   │   ├── PdfReportBuilder.java       [YENİ]
│   │   └── ReportValidator.java        [YENİ]
│   ├── data/
│   │   ├── DataExtractor.java          [YENİ]
│   │   ├── DataNormalizer.java         [YENİ]
│   │   └── DataConverter.java          [YENİ]
│   ├── io/
│   │   ├── DataIOFactory.java          [YENİ]
│   │   ├── FileTransferService.java    [YENİ]
│   │   └── FileValidator.java          [YENİ]
│   └── util/
│       └── ProgressNotifier.java       [YENİ]
└── controller/
    └── command/
        └── ReportCommandFactory.java   [YENİ]
```

---

## ADIM 2: Core Exception Handling

### File: `core/exception/ErrorCode.java`

```java
package com.zeki.merger.core.exception;

public enum ErrorCode {
    FILE_NOT_FOUND(1001, "Dosya bulunamadı"),
    INVALID_FORMAT(1002, "Geçersiz dosya formatı"),
    DATABASE_ERROR(1003, "Veritabanı hatası"),
    GENERATION_FAILED(1004, "Rapor oluşturma başarısız"),
    INVALID_DATA(1005, "Geçersiz veri"),
    IO_ERROR(1006, "Giriş/çıkış hatası"),
    NORMALIZATION_ERROR(1007, "Veri normalleştirme hatası");

    private final int code;
    private final String defaultMessage;

    ErrorCode(int code, String defaultMessage) {
        this.code = code;
        this.defaultMessage = defaultMessage;
    }

    public int getCode() { return code; }
    public String getDefaultMessage() { return defaultMessage; }
}
```

### File: `core/exception/BusinessException.java`

```java
package com.zeki.merger.core.exception;

import java.util.Collections;
import java.util.HashMap;
import java.util.Map;

public class BusinessException extends RuntimeException {
    private final ErrorCode errorCode;
    private final Map<String, Object> context;

    public BusinessException(ErrorCode errorCode, String message) {
        super(message);
        this.errorCode = errorCode;
        this.context = new HashMap<>();
    }

    public BusinessException(ErrorCode errorCode, String message, Throwable cause) {
        super(message, cause);
        this.errorCode = errorCode;
        this.context = new HashMap<>();
    }

    public BusinessException(ErrorCode errorCode, String message, Map<String, Object> context) {
        super(message);
        this.errorCode = errorCode;
        this.context = new HashMap<>(context);
    }

    public ErrorCode getErrorCode() { return errorCode; }
    public Map<String, Object> getContext() { 
        return Collections.unmodifiableMap(context); 
    }

    public String getDetailedMessage() {
        StringBuilder sb = new StringBuilder()
            .append("[").append(errorCode.getCode()).append("] ")
            .append(getMessage());
        
        if (!context.isEmpty()) {
            sb.append(" | Context: ").append(context);
        }
        return sb.toString();
    }
}
```

---

## ADIM 3: Observer Pattern (Progress Tracking)

### File: `service/util/ProgressObserver.java`

```java
package com.zeki.merger.service.util;

public interface ProgressObserver {
    void onProgressUpdate(double progress, String message);
    void onCompleted();
    void onFailed(Exception exception);
}
```

### File: `service/util/ProgressNotifier.java`

```java
package com.zeki.merger.service.util;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

public class ProgressNotifier {
    private final List<ProgressObserver> observers = new ArrayList<>();

    public void subscribe(ProgressObserver observer) {
        Objects.requireNonNull(observer, "Observer cannot be null");
        observers.add(observer);
    }

    public void unsubscribe(ProgressObserver observer) {
        observers.remove(observer);
    }

    public void notifyProgress(double progress, String message) {
        for (ProgressObserver observer : new ArrayList<>(observers)) {
            try {
                observer.onProgressUpdate(progress, message);
            } catch (Exception e) {
                System.err.println("Observer notification failed: " + e.getMessage());
            }
        }
    }

    public void notifyCompleted() {
        for (ProgressObserver observer : new ArrayList<>(observers)) {
            try {
                observer.onCompleted();
            } catch (Exception e) {
                System.err.println("Observer notification failed: " + e.getMessage());
            }
        }
    }

    public void notifyFailed(Exception exception) {
        for (ProgressObserver observer : new ArrayList<>(observers)) {
            try {
                observer.onFailed(exception);
            } catch (Exception e) {
                System.err.println("Observer notification failed: " + e.getMessage());
            }
        }
    }
}
```

---

## ADIM 4: Builder Pattern (Excel Sheet)

### File: `service/excel/ExcelSheetBuilder.java`

```java
package com.zeki.merger.service.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.List;
import java.util.Map;
import java.util.Objects;

public class ExcelSheetBuilder {
    private final Workbook workbook;
    private final Sheet sheet;
    private final ExcelStyleFactory styleFactory;
    
    private int currentRow = 0;
    private int freezeRowCount = 0;
    private int freezeColCount = 0;
    private boolean hasAutoFilter = false;

    public ExcelSheetBuilder(String sheetName) {
        this.workbook = new XSSFWorkbook();
        this.sheet = workbook.createSheet(sheetName);
        this.styleFactory = new ExcelStyleFactory();
    }

    public ExcelSheetBuilder withDimensions(int rows, int cols) {
        sheet.setDefaultColumnWidth(20);
        for (int i = 0; i < cols; i++) {
            sheet.setColumnWidth(i, 20 * 256);
        }
        return this;
    }

    public ExcelSheetBuilder withFrozenPane(int freezeRow, int freezeCol) {
        this.freezeRowCount = freezeRow;
        this.freezeColCount = freezeCol;
        return this;
    }

    public ExcelSheetBuilder addHeaderRow(List<String> headers) {
        Row headerRow = sheet.createRow(currentRow);
        CellStyle headerStyle = styleFactory.getHeaderStyle(workbook);
        
        for (int i = 0; i < headers.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers.get(i));
            cell.setCellStyle(headerStyle);
        }
        
        currentRow++;
        this.hasAutoFilter = true;
        return this;
    }

    public ExcelSheetBuilder addDataRow(List<Object> values) {
        Row dataRow = sheet.createRow(currentRow);
        CellStyle dataStyle = styleFactory.getDataStyle(workbook);
        
        for (int i = 0; i < values.size(); i++) {
            Cell cell = dataRow.createCell(i);
            Object value = values.get(i);
            
            if (value instanceof Number) {
                cell.setCellValue(((Number) value).doubleValue());
            } else if (value instanceof Boolean) {
                cell.setCellValue((Boolean) value);
            } else {
                cell.setCellValue(value != null ? value.toString() : "");
            }
            cell.setCellStyle(dataStyle);
        }
        
        currentRow++;
        return this;
    }

    public ExcelSheetBuilder addDataRows(List<List<Object>> rows) {
        for (List<Object> row : rows) {
            addDataRow(row);
        }
        return this;
    }

    public ExcelSheetBuilder addFormulaCell(int row, int col, String formula) {
        Row xlsRow = sheet.getRow(row);
        if (xlsRow == null) {
            xlsRow = sheet.createRow(row);
        }
        Cell cell = xlsRow.createCell(col);
        cell.setCellFormula(formula);
        return this;
    }

    public ExcelSheetBuilder withBorders(BorderStyle style) {
        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    Cell cell = row.getCell(j);
                    if (cell != null) {
                        CellStyle cellStyle = workbook.createCellStyle();
                        cellStyle.cloneStyleFrom(cell.getCellStyle());
                        cellStyle.setBorderLeft(style);
                        cellStyle.setBorderRight(style);
                        cellStyle.setBorderTop(style);
                        cellStyle.setBorderBottom(style);
                        cell.setCellStyle(cellStyle);
                    }
                }
            }
        }
        return this;
    }

    public Workbook build() {
        if (freezeRowCount > 0 || freezeColCount > 0) {
            sheet.createFreezePane(freezeColCount, freezeRowCount);
        }
        
        if (hasAutoFilter && sheet.getLastRowNum() >= 0) {
            sheet.setAutoFilter(new org.apache.poi.ss.util.CellRangeAddress(
                0, sheet.getLastRowNum(), 0, sheet.getRow(0).getLastCellNum() - 1
            ));
        }
        
        return workbook;
    }
}
```

### File: `service/excel/ExcelStyleFactory.java`

```java
package com.zeki.merger.service.excel;

import org.apache.poi.ss.usermodel.*;

public class ExcelStyleFactory {
    
    public CellStyle getHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontName("Calibri");
        font.setFontHeightInPoints((short) 12);
        font.setBold(true);
        font.setColor(IndexedColors.WHITE.getIndex());
        
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.DARK_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        
        return style;
    }

    public CellStyle getDataStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontName("Calibri");
        font.setFontHeightInPoints((short) 11);
        
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        
        return style;
    }

    public CellStyle getCurrencyStyle(Workbook workbook) {
        CellStyle style = getDataStyle(workbook);
        style.setDataFormat(workbook.createDataFormat().getFormat("\"€\" #,##0.00"));
        style.setAlignment(HorizontalAlignment.RIGHT);
        return style;
    }

    public CellStyle getPercentageStyle(Workbook workbook) {
        CellStyle style = getDataStyle(workbook);
        style.setDataFormat(workbook.createDataFormat().getFormat("0.00%"));
        style.setAlignment(HorizontalAlignment.RIGHT);
        return style;
    }

    public CellStyle getTotalStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontName("Calibri");
        font.setFontHeightInPoints((short) 11);
        font.setBold(true);
        
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.LIGHT_GRAY.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.RIGHT);
        
        return style;
    }
}
```

---

## ADIM 5: Strategy Pattern (Output Formats)

### File: `service/report/ReportStrategy.java`

```java
package com.zeki.merger.service.report;

import java.io.File;
import java.util.Map;

public interface ReportStrategy {
    File generate(Map<String, Object> data, File outputPath) throws Exception;
    String getFormat();
}
```

### File: `service/report/ExcelReportStrategy.java`

```java
package com.zeki.merger.service.report;

import com.zeki.merger.service.excel.ExcelSheetBuilder;
import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

public class ExcelReportStrategy implements ReportStrategy {
    
    @Override
    public File generate(Map<String, Object> data, File outputPath) throws Exception {
        ExcelSheetBuilder builder = new ExcelSheetBuilder("Report");
        
        @SuppressWarnings("unchecked")
        List<String> headers = (List<String>) data.get("headers");
        @SuppressWarnings("unchecked")
        List<List<Object>> rows = (List<List<Object>>) data.get("rows");
        
        builder.withDimensions(rows.size() + 1, headers.size())
               .addHeaderRow(headers)
               .addDataRows(rows)
               .withBorders(org.apache.poi.ss.usermodel.BorderStyle.THIN);
        
        try (FileOutputStream fos = new FileOutputStream(outputPath)) {
            builder.build().write(fos);
        }
        
        return outputPath;
    }

    @Override
    public String getFormat() {
        return "XLSX";
    }
}
```

### File: `service/report/PdfReportStrategy.java`

```java
package com.zeki.merger.service.report;

import java.io.File;
import java.util.List;
import java.util.Map;

public class PdfReportStrategy implements ReportStrategy {
    
    @Override
    public File generate(Map<String, Object> data, File outputPath) throws Exception {
        // PDF generation logic
        // (Using iText or Apache PDFBox)
        throw new UnsupportedOperationException("PDF generation not yet implemented");
    }

    @Override
    public String getFormat() {
        return "PDF";
    }
}
```

### File: `service/report/ReportStrategyFactory.java`

```java
package com.zeki.merger.service.report;

import com.zeki.merger.core.exception.BusinessException;
import com.zeki.merger.core.exception.ErrorCode;
import java.util.HashMap;
import java.util.Map;

public class ReportStrategyFactory {
    private final Map<String, ReportStrategy> strategies = new HashMap<>();

    public ReportStrategyFactory() {
        register("XLSX", new ExcelReportStrategy());
        register("PDF", new PdfReportStrategy());
        // Add more strategies as needed
    }

    public void register(String format, ReportStrategy strategy) {
        strategies.put(format.toUpperCase(), strategy);
    }

    public ReportStrategy getStrategy(String format) {
        String key = format.toUpperCase();
        ReportStrategy strategy = strategies.get(key);
        
        if (strategy == null) {
            throw new BusinessException(
                ErrorCode.INVALID_FORMAT,
                "Unsupported report format: " + format,
                Map.of("supportedFormats", strategies.keySet())
            );
        }
        
        return strategy;
    }
}
```

---

## ADIM 6: Factory Pattern (Data I/O)

### File: `service/io/DataIOFactory.java`

```java
package com.zeki.merger.service.io;

import com.zeki.merger.core.exception.BusinessException;
import com.zeki.merger.core.exception.ErrorCode;
import com.zeki.merger.service.ExcelReader;
import com.zeki.merger.service.ExcelWriter;
import java.io.File;
import java.util.HashMap;
import java.util.Map;

public class DataIOFactory {
    private final Map<String, FileReader> readers = new HashMap<>();
    private final Map<String, FileWriter> writers = new HashMap<>();

    public interface FileReader {
        Map<String, Object> read(File file) throws Exception;
    }

    public interface FileWriter {
        File write(Map<String, Object> data, File outputPath) throws Exception;
    }

    public DataIOFactory() {
        // Register Excel handlers
        readers.put("xlsx", new ExcelReader());
        readers.put("xls", new ExcelReader());
        
        writers.put("xlsx", new ExcelWriter());
        writers.put("xls", new ExcelWriter());
    }

    public FileReader getReader(String fileExtension) {
        String ext = fileExtension.toLowerCase().replace(".", "");
        FileReader reader = readers.get(ext);
        
        if (reader == null) {
            throw new BusinessException(
                ErrorCode.INVALID_FORMAT,
                "No reader available for format: " + fileExtension
            );
        }
        
        return reader;
    }

    public FileWriter getWriter(String fileExtension) {
        String ext = fileExtension.toLowerCase().replace(".", "");
        FileWriter writer = writers.get(ext);
        
        if (writer == null) {
            throw new BusinessException(
                ErrorCode.INVALID_FORMAT,
                "No writer available for format: " + fileExtension
            );
        }
        
        return writer;
    }

    public FileReader getReaderByFile(File file) {
        String ext = getFileExtension(file);
        return getReader(ext);
    }

    public FileWriter getWriterByFile(File file) {
        String ext = getFileExtension(file);
        return getWriter(ext);
    }

    private String getFileExtension(File file) {
        String name = file.getName();
        int lastIndexOfDot = name.lastIndexOf(".");
        
        if (lastIndexOfDot <= 0) {
            throw new BusinessException(
                ErrorCode.INVALID_FORMAT,
                "File has no extension: " + name
            );
        }
        
        return name.substring(lastIndexOfDot + 1);
    }
}
```

---

## ADIM 7: Command Pattern (UI Actions)

### File: `controller/command/ReportCommand.java`

```java
package com.zeki.merger.controller.command;

import com.zeki.merger.service.util.ProgressNotifier;

public interface ReportCommand {
    void execute(ProgressNotifier progressNotifier) throws Exception;
    String getName();
}
```

### File: `controller/command/GenerateTrfCommand.java`

```java
package com.zeki.merger.controller.command;

import com.zeki.merger.service.util.ProgressNotifier;
import com.zeki.merger.trf.TrfGeneratorService;
import com.zeki.merger.db.DatabaseManager;

public class GenerateTrfCommand implements ReportCommand {
    private final TrfGeneratorService trfGeneratorService;

    public GenerateTrfCommand(DatabaseManager dbManager) {
        this.trfGeneratorService = new TrfGeneratorService(dbManager);
    }

    @Override
    public void execute(ProgressNotifier progressNotifier) throws Exception {
        progressNotifier.notifyProgress(0.0, "Generating TRF report...");
        
        try {
            // TRF generation logic
            progressNotifier.notifyProgress(0.5, "Processing data...");
            // ...
            progressNotifier.notifyProgress(1.0, "TRF generation completed!");
            progressNotifier.notifyCompleted();
        } catch (Exception e) {
            progressNotifier.notifyFailed(e);
            throw e;
        }
    }

    @Override
    public String getName() {
        return "GENERATE_TRF";
    }
}
```

### File: `controller/command/ReportCommandFactory.java`

```java
package com.zeki.merger.controller.command;

import com.zeki.merger.core.exception.BusinessException;
import com.zeki.merger.core.exception.ErrorCode;
import com.zeki.merger.db.DatabaseManager;
import java.util.HashMap;
import java.util.Map;

public class ReportCommandFactory {
    private final Map<String, ReportCommand> commandMap = new HashMap<>();
    private final DatabaseManager dbManager;

    public ReportCommandFactory(DatabaseManager dbManager) {
        this.dbManager = dbManager;
        registerCommands();
    }

    private void registerCommands() {
        commandMap.put("GENERATE_TRF", new GenerateTrfCommand(dbManager));
        // commandMap.put("EXPORT_PUBLIC", new ExportPublicCommand(dbManager));
        // commandMap.put("COMPARE_FILES", new CompareFilesCommand());
        // commandMap.put("FIX_PATHS", new FixPathsCommand());
    }

    public ReportCommand getCommand(String commandName) {
        ReportCommand command = commandMap.get(commandName.toUpperCase());
        
        if (command == null) {
            throw new BusinessException(
                ErrorCode.GENERATION_FAILED,
                "Unknown command: " + commandName
            );
        }
        
        return command;
    }
}
```

---

## ADIM 8: Refactored Data Layer

### File: `service/data/DataExtractor.java`

```java
package com.zeki.merger.service.data;

import org.apache.poi.ss.usermodel.*;
import java.util.*;

public class DataExtractor {
    private final DataFormatter formatter = new DataFormatter();

    public String extractString(Row row, int columnIndex, FormulaEvaluator evaluator) {
        Cell cell = row.getCell(columnIndex);
        if (cell == null) return "";
        
        try {
            return formatter.formatCellValue(cell, evaluator).trim();
        } catch (Exception e) {
            return "";
        }
    }

    public double extractDouble(Row row, int columnIndex, FormulaEvaluator evaluator) {
        Cell cell = row.getCell(columnIndex);
        if (cell == null) return 0.0;
        
        try {
            CellValue cellValue = evaluator.evaluate(cell);
            if (cellValue.getNumberValue() == 0) return 0.0;
            return cellValue.getNumberValue();
        } catch (Exception e) {
            return 0.0;
        }
    }

    public List<String> extractRow(Row row, int... columnIndices) {
        List<String> result = new ArrayList<>();
        FormulaEvaluator evaluator = null;
        
        for (int colIndex : columnIndices) {
            result.add(extractString(row, colIndex, evaluator));
        }
        
        return result;
    }
}
```

### File: `service/data/DataNormalizer.java`

```java
package com.zeki.merger.service.data;

public class DataNormalizer {
    
    public String normalize(String value) {
        if (value == null || value.trim().isEmpty()) {
            return "";
        }
        
        return value
            .trim()
            .replaceAll("\\s+", " ")
            .toLowerCase();
    }

    public double normalizeAmount(double value) {
        // Round to 2 decimal places
        return Math.round(value * 100.0) / 100.0;
    }

    public String normalizeClientName(String name) {
        return normalize(name)
            .replaceAll("[^a-z0-9\\s]", "");
    }

    public String normalizePhoneNumber(String phone) {
        return phone
            .replaceAll("[^0-9+]", "");
    }
}
```

### File: `service/data/DataConverter.java`

```java
package com.zeki.merger.service.data;

import com.zeki.merger.model.CreanceRow;
import java.util.*;

public class DataConverter {
    
    public CreanceRow mapToCreanceRow(Map<String, String> dataMap) {
        CreanceRow row = new CreanceRow();
        row.setClientCode(dataMap.getOrDefault("clientCode", ""));
        row.setClientName(dataMap.getOrDefault("clientName", ""));
        row.setAmount(parseDouble(dataMap.getOrDefault("amount", "0")));
        // Map other fields...
        return row;
    }

    public List<CreanceRow> mapToCreanceRows(List<Map<String, String>> dataList) {
        List<CreanceRow> rows = new ArrayList<>();
        for (Map<String, String> dataMap : dataList) {
            rows.add(mapToCreanceRow(dataMap));
        }
        return rows;
    }

    private double parseDouble(String value) {
        try {
            return Double.parseDouble(value.replaceAll("[^0-9.-]", ""));
        } catch (NumberFormatException e) {
            return 0.0;
        }
    }
}
```

---

## ADIM 9: Updated MainController

### File: `controller/MainController.java` (Refactored)

```java
package com.zeki.merger.controller;

import com.zeki.merger.controller.command.*;
import com.zeki.merger.db.DatabaseManager;
import com.zeki.merger.service.util.ProgressObserver;
import com.zeki.merger.service.util.ProgressNotifier;
import javafx.application.Platform;
import javafx.fxml.FXML;
import javafx.scene.control.*;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class MainController implements ProgressObserver {
    
    @FXML private ProgressBar progressBar;
    @FXML private TextArea logArea;
    @FXML private Label statusLabel;
    @FXML private GridPane actionsGrid;
    
    private final ReportCommandFactory commandFactory;
    private final ProgressNotifier progressNotifier = new ProgressNotifier();
    private final ExecutorService executor = Executors.newSingleThreadExecutor();

    public MainController() {
        this.commandFactory = new ReportCommandFactory(DatabaseManager.getInstance());
    }

    @FXML
    public void initialize() {
        progressNotifier.subscribe(this);
        createActionButtons();
    }

    private void createActionButtons() {
        Button trfBtn = createButton("Generate TRF", () -> executeCommand("GENERATE_TRF"));
        Button cmpBtn = createButton("Compare Files", () -> executeCommand("COMPARE_FILES"));
        
        actionsGrid.add(trfBtn, 0, 0);
        actionsGrid.add(cmpBtn, 1, 0);
    }

    private Button createButton(String text, Runnable action) {
        Button btn = new Button(text);
        btn.setOnAction(e -> action.run());
        return btn;
    }

    private void executeCommand(String commandName) {
        executor.execute(() -> {
            try {
                ReportCommand command = commandFactory.getCommand(commandName);
                command.execute(progressNotifier);
            } catch (Exception e) {
                progressNotifier.notifyFailed(e);
            }
        });
    }

    @Override
    public void onProgressUpdate(double progress, String message) {
        Platform.runLater(() -> {
            progressBar.setProgress(progress);
            statusLabel.setText(message);
            logArea.appendText(message + "\n");
        });
    }

    @Override
    public void onCompleted() {
        Platform.runLater(() -> statusLabel.setText("✓ Completed successfully"));
    }

    @Override
    public void onFailed(Exception exception) {
        Platform.runLater(() -> {
            statusLabel.setText("✗ Error: " + exception.getMessage());
            logArea.appendText("\n[ERROR] " + exception.getMessage() + "\n");
        });
    }
}
```

---

## ADIM 10: Configuration Management

### File: `core/config/ApplicationConfig.java`

```java
package com.zeki.merger.core.config;

import com.zeki.merger.controller.command.ReportCommandFactory;
import com.zeki.merger.db.DatabaseManager;
import com.zeki.merger.service.report.ReportStrategyFactory;
import com.zeki.merger.service.io.DataIOFactory;

public class ApplicationConfig {
    private static ApplicationConfig instance;
    private final DatabaseManager databaseManager;
    private final ReportStrategyFactory reportStrategyFactory;
    private final DataIOFactory dataIOFactory;
    private final ReportCommandFactory reportCommandFactory;

    private ApplicationConfig() {
        this.databaseManager = DatabaseManager.getInstance();
        this.reportStrategyFactory = new ReportStrategyFactory();
        this.dataIOFactory = new DataIOFactory();
        this.reportCommandFactory = new ReportCommandFactory(databaseManager);
    }

    public static synchronized ApplicationConfig getInstance() {
        if (instance == null) {
            instance = new ApplicationConfig();
        }
        return instance;
    }

    public DatabaseManager getDatabaseManager() { return databaseManager; }
    public ReportStrategyFactory getReportStrategyFactory() { return reportStrategyFactory; }
    public DataIOFactory getDataIOFactory() { return dataIOFactory; }
    public ReportCommandFactory getReportCommandFactory() { return reportCommandFactory; }
}
```

---

## ÖZET: Uygulanacak Kod Değişiklikleri

1. ✅ **core/exception/** → Custom exception handling
2. ✅ **service/util/** → Observer pattern for progress
3. ✅ **service/excel/** → Builder pattern for sheets + Style factory
4. ✅ **service/report/** → Strategy pattern for multiple formats
5. ✅ **service/io/** → Factory pattern for readers/writers
6. ✅ **service/data/** → Extract, normalize, convert logic
7. ✅ **controller/command/** → Command pattern for actions
8. ✅ **core/config/** → Centralized configuration
9. ✅ **controller/MainController** → Simplified, delegating

## Sonraki Adımlar

1. Bu dosyaları projeye ekle (copy-paste)
2. Imports'ları düzelt
3. Mevcut sınıfları refactor et (TrfSheetWriter, EtatPublicGenerator vs)
4. Unit testleri yaz
5. Integration testleri yaz
6. Maven build et ve test et

---

**Başlama Zamanı:** 09:00  
**Tahmini Tamamlama:** 17:00  
**Kod Yazma Süresi:** ~4-5 saat (copy-paste + düzeltmeler)  
**Test Yazma Süresi:** ~2 saat

Hazır mısın? 🚀
