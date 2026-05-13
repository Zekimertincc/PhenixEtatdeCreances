# ⚡ QUICK REFERENCE - REFACTORING CHECKLIST

## 🎯 GÜNÜN PLANLAMASI

```
09:00-09:30  Başlangıç        - Bu dosyayı oku, setup yap
09:30-12:00  PART 1          - TrfSheetWriter refactor
12:00-13:00  LUNCH & BREAK   - 💪 Dinlenme
13:00-15:30  PART 2          - Data layer refactor
15:30-16:30  PART 3          - Controller refactor
16:30-17:00  Testing & Wrap  - Build + Test
```

---

## 🚀 ÖNCESİ vs SONRASI

### ÖNCESİ: Spaghetti Code
```java
// TrfSheetWriter.java (717 satır) - PROBLEM!
public class TrfSheetWriter {
    public File generate(List<Data> data) {
        Workbook wb = new XSSFWorkbook();  // ← Style logic burada
        Sheet sheet = wb.createSheet();     // ← Format logic burada
        Row headerRow = sheet.createRow(0); // ← Color logic burada
        Cell cell = headerRow.createCell(0); // ← Font logic burada
        // ... 700 satır daha karışık kod ...
    }
}
```

### SONRASI: Clean Architecture
```java
// Refactored - TEMIZ!
public class TrfSheetWriter {
    private final ExcelStyleFactory styleFactory;
    private final ExcelSheetBuilder builder;
    
    public File generate(List<Data> data) {
        builder.withStyle(styleFactory.getHeaderStyle())
               .addHeaderRow(getHeaders())
               .addDataRows(data)
               .build()
               .write(outputFile);
    }
}
```

---

## 📋 REFACTORING CHECKLIST

### PART 1: Excel Layer (TrfSheetWriter)

**Tahmini Zaman: 2 saat**

- [ ] `ExcelStyleFactory.java` oluştur
  - [ ] Header style
  - [ ] Data style
  - [ ] Currency style
  - [ ] Total style

- [ ] `ExcelFormatterService.java` oluştur
  - [ ] Format currency
  - [ ] Format percentage
  - [ ] Format date
  - [ ] Apply alignment

- [ ] `ExcelSheetBuilder.java` oluştur (Builder Pattern)
  - [ ] Constructor
  - [ ] withStyle()
  - [ ] addHeaderRow()
  - [ ] addDataRow()
  - [ ] withBorders()
  - [ ] build()

- [ ] `TrfSheetWriter.java` refactor (100 satırdan az)
  - [ ] Sadece generate() metodu
  - [ ] Sadece orchestration
  - [ ] StyleFactory kullan
  - [ ] SheetBuilder kullan

- [ ] Test Et
  - [ ] Excel file oluşturulur mu?
  - [ ] Styles doğru mu?
  - [ ] Formulas çalışır mı?

---

### PART 2: Data Layer

**Tahmini Zaman: 2 saat**

- [ ] Exception Handling oluştur
  - [ ] `ErrorCode.java` enum
  - [ ] `BusinessException.java` class

- [ ] Observer Pattern
  - [ ] `ProgressObserver.java` interface
  - [ ] `ProgressNotifier.java` class

- [ ] Data Processing Classes
  - [ ] `DataExtractor.java` - Excel okuma
  - [ ] `DataNormalizer.java` - String normalize
  - [ ] `DataConverter.java` - Model mapping

- [ ] Factory Pattern
  - [ ] `DataIOFactory.java` - Reader/Writer selection
  - [ ] Register Excel handlers

- [ ] Strategy Pattern
  - [ ] `ReportStrategy.java` interface
  - [ ] `ExcelReportStrategy.java`
  - [ ] `PdfReportStrategy.java` (stub)
  - [ ] `ReportStrategyFactory.java`

- [ ] Test Et
  - [ ] Data extract doğru mu?
  - [ ] Normalize çalışıyor mu?
  - [ ] Factory correct strategy döndürüyor mu?

---

### PART 3: Controller Layer

**Tahmini Zaman: 1.5 saat**

- [ ] Command Pattern
  - [ ] `ReportCommand.java` interface
  - [ ] `GenerateTrfCommand.java`
  - [ ] `CompareFilesCommand.java`
  - [ ] `FixPathsCommand.java`
  - [ ] `ExportPublicCommand.java`
  - [ ] `ReportCommandFactory.java`

- [ ] Configuration
  - [ ] `ApplicationConfig.java` - Centralized setup

- [ ] MainController Refactor
  - [ ] Progress observers subscribe et
  - [ ] Executor management
  - [ ] Command execution

- [ ] Test Et
  - [ ] Command'ler çalışıyor mu?
  - [ ] Progress updates geliyor mu?
  - [ ] Error handling çalışıyor mu?

---

## 🔧 COPY-PASTE CODE LOCATIONS

Heryerde bu kodu bulabilirsin (CONCRETE_CODE_EXAMPLES.md içinde):

| Sınıf | Satır | Kopyala |
|-------|-------|--------|
| ExcelStyleFactory | ~70 | 1️⃣ |
| ExcelSheetBuilder | ~150 | 2️⃣ |
| BusinessException | ~40 | 3️⃣ |
| ProgressNotifier | ~50 | 4️⃣ |
| ReportStrategy | ~20 | 5️⃣ |
| DataIOFactory | ~60 | 6️⃣ |
| ReportCommand | ~10 | 7️⃣ |
| ApplicationConfig | ~30 | 8️⃣ |

---

## 🎓 PATTERN QUICK REFERENCE

### Builder Pattern
```java
// ✅ Kullan:
builder.withStyle(style)
       .addHeader(headers)
       .addRows(data)
       .build();

// ❌ Değil:
builder.setStyle(style);
builder.setHeaders(headers);
builder.setRows(data);
builder.build(); // çirkin!
```

### Strategy Pattern
```java
// ✅ Kullan:
ReportStrategy strategy = factory.getStrategy("XLSX");
File result = strategy.generate(data, output);

// ❌ Değil:
if (format.equals("XLSX")) {
    // Excel logic
} else if (format.equals("PDF")) {
    // PDF logic
} // 100 satır if-else!
```

### Factory Pattern
```java
// ✅ Kullan:
DataReader reader = factory.getReader("xlsx");
Map<String, Object> data = reader.read(file);

// ❌ Değil:
if (file.endsWith(".xlsx")) {
    reader = new XlsxReader();
} else if (file.endsWith(".xls")) {
    reader = new XlsReader();
} // Sertleştirilmiş!
```

### Observer Pattern
```java
// ✅ Kullan:
notifier.subscribe(uiController);
notifier.subscribe(logService);
notifier.notifyProgress(0.5, "Processing...");

// ❌ Değil:
service.setUiController(uiController);
service.setLogService(logService);
service.notifyUI(0.5);
service.notifyLog(0.5); // Tight coupling!
```

### Command Pattern
```java
// ✅ Kullan:
ReportCommand cmd = factory.getCommand("GENERATE_TRF");
cmd.execute(progressNotifier);

// ❌ Değil:
if (action.equals("GENERATE_TRF")) {
    generateTrf();
} else if (action.equals("COMPARE")) {
    compare();
} // Switch-case spaghetti!
```

---

## ⚠️ YAYGIN HATALAR & ÇÖZÜMLERI

### Hata 1: Null Pointer Exceptions
```java
// ❌ YANLIŞ:
public void doSomething(File file) {
    file.getName(); // file null olabilir!
}

// ✅ DOĞRU:
public void doSomething(File file) {
    Objects.requireNonNull(file, "File cannot be null");
    file.getName();
}
```

### Hata 2: Unchecked Exceptions
```java
// ❌ YANLIŞ:
try {
    data = read(file);
} catch (Exception e) {
    System.out.println("Error"); // Context yok!
}

// ✅ DOĞRU:
try {
    data = read(file);
} catch (IOException e) {
    throw new BusinessException(
        ErrorCode.FILE_NOT_FOUND,
        "Cannot read file: " + file.getName(),
        Map.of("file", file.getPath(), "cause", e.getMessage())
    );
}
```

### Hata 3: Tight Coupling
```java
// ❌ YANLIŞ:
public class ReportService {
    private ExcelWriter writer = new ExcelWriter();
    private PdfWriter writer2 = new PdfWriter();
}

// ✅ DOĞRU:
public class ReportService {
    private final ReportStrategy strategy;
    
    public ReportService(ReportStrategy strategy) {
        this.strategy = strategy;
    }
}
```

### Hata 4: Forgotten Progress Updates
```java
// ❌ YANLIŞ:
public void generate(File input, File output) {
    data = readFile(input);
    // ... 10 satır işlem ...
    // UI hiç update olmadı!
    writeFile(data, output);
}

// ✅ DOĞRU:
public void generate(File input, File output, ProgressNotifier notifier) {
    notifier.notifyProgress(0.0, "Reading file...");
    data = readFile(input);
    
    notifier.notifyProgress(0.5, "Processing...");
    // ... işlem ...
    
    notifier.notifyProgress(0.9, "Writing output...");
    writeFile(data, output);
    
    notifier.notifyProgress(1.0, "Done!");
    notifier.notifyCompleted();
}
```

---

## 📦 MAVEN BUILD KOMUTU

```bash
# Clean & Build
mvn clean package

# With Tests
mvn clean package -DskipTests=false

# Skip Tests (fast)
mvn clean package -DskipTests

# Run Single Test
mvn test -Dtest=TrfSheetWriterTest

# Generate JAR
mvn clean package -DskipTests
# Output: target/etat-creances-merger-1.0.0.jar
```

---

## 🧪 TESTING STRATEGY

### Unit Tests - Yazılacaklar
```java
public class ExcelSheetBuilderTest {
    @Test
    public void testHeadersAdded() { }
    
    @Test
    public void testDataRowsAdded() { }
    
    @Test
    public void testFrozenPaneApplied() { }
}

public class DataNormalizerTest {
    @Test
    public void testNormalizeString() { }
    
    @Test
    public void testNormalizeAmount() { }
}

public class ReportCommandFactoryTest {
    @Test
    public void testGetValidCommand() { }
    
    @Test
    public void testGetInvalidCommand() { }
}
```

### Integration Tests
```java
public class ReportGenerationIntegrationTest {
    @Test
    public void testEndToEndTrfGeneration() { }
    
    @Test
    public void testProgressNotification() { }
    
    @Test
    public void testErrorHandling() { }
}
```

---

## 📊 ÖLÇÜM METRIKLERI

Başlangıçta vs Sonunda karşılaştır:

### Kod Kalitesi
```
Başında:          Sonunda:
─────────────────────────
Avg Class Size:   175 satır    →  80 satır ✓
Max Class Size:   717 satır    →  150 satır ✓
Responsibilities: 3.2/class    →  1.1/class ✓
Duplication:      12%          →  2% ✓
Test Coverage:    5%           →  65% ✓
```

### Performance
- Derleme: 30 saniye (no change)
- JAR boyutu: ~28MB (no change)
- Runtime: Same (refactoring yapısal)

---

## 🎯 GÜN SONU BEKLENTILER

### Tamamlanması Gereken:
- [ ] 8+ yeni sınıf oluşturulmuş
- [ ] 5+ design pattern uygulanmış
- [ ] 10+ sınıf refactored
- [ ] 20+ test yazılmış
- [ ] 0 compilation error
- [ ] Proje successfully builds

### Beklenen Kod Kalitesi:
- [ ] Tüm sınıflar < 200 satır
- [ ] SRP uyumlu (1 sorumluluk/sınıf)
- [ ] No duplication
- [ ] All exceptions handled
- [ ] All UI updates threadsafe
- [ ] All resources properly closed

---

## 🆘 SORUN ÇÖZÜM REHBERI

| Sorun | Çözüm |
|-------|-------|
| **Compilation Error** | Imports kontrol et, package adları eşleş |
| **NullPointerException** | Objects.requireNonNull() ekle |
| **Progress update yok** | ProgressNotifier.subscribe() çağrıldı mı? |
| **Memory Leak** | try-with-resources kullan (.close()) |
| **Thread deadlock** | Platform.runLater() kullan UI updates için |
| **Slow build** | `mvn clean` temizle, cache clear et |

---

## 📚 KAYNAKLAR

- Gang of Four Design Patterns: https://en.wikipedia.org/wiki/Design_Patterns
- SOLID Principles: https://en.wikipedia.org/wiki/SOLID
- JavaFX Threading: https://docs.oracle.com/javase/8/javafx/
- Apache POI: https://poi.apache.org/

---

## 💬 NOT

Bu refactoring:
- ✅ Kod yazma mahareti geliştir
- ✅ Pattern bilgini kat kat arttır
- ✅ Professional code yapısını öğren
- ✅ Projeyi kontrol altına al
- ✅ Future features kolay eklenebilir hale getir

**Sen yapabilirsin! 💪**

---

**Başlangıç:** 09:00  
**Target Bitiş:** 17:00  
**Break:** 12:00-13:00 (1 saat)  
**Aktif Çalışma:** ~7 saat

Let's GO! 🚀🚀🚀
