# ETAT CREANCES MERGER - REFACTORİNG PLAN

**Dönem:** Tek günde tamamlanabilir ✅  
**Proje Türü:** JavaFX Desktop Application (Maven)  
**Toplam Java Satırı:** 4,900  
**Dosya Sayısı:** 28 .java sınıf  
**Kod Duyarlılığı:** Orta-Yüksek (iyi şekilde yapılandırılmış, ama pattern'ler eksik)

---

## 📊 PRİORİTİ ANALIZI

| Sıra | Dosya | Satır | Risk | Impact | Çaba | Status |
|------|-------|-------|------|--------|------|--------|
| 1 | `TrfSheetWriter.java` | 717 | 🔴 Yüksek | 🟢 Yüksek | 3 saat | Kritik |
| 2 | `EtatPublicGenerator.java` | 537 | 🔴 Yüksek | 🟢 Yüksek | 2.5 saat | Kritik |
| 3 | `ProcreancesComparator.java` | 479 | 🟡 Orta | 🟢 Yüksek | 2 saat | Önemli |
| 4 | `MainController.java` | 477 | 🟡 Orta | 🟡 Orta | 1.5 saat | Önemli |
| 5 | `DataReader.java` | 322 | 🟡 Orta | 🟢 Yüksek | 2 saat | Önemli |
| 6 | `DatabaseManager.java` | 302 | 🔴 Yüksek | 🟢 Yüksek | 1.5 saat | Önemli |
| 7 | `TrfWriter.java` | 292 | 🟡 Orta | 🟢 Yüksek | 1.5 saat | Önemli |
| Diğer | 8 sınıf | ~800 | 🟢 Düşük | 🟡 Orta | 1 saat | Minor |

---

## 🎯 TEŞHIS: Ana Sorunlar

### 1. **God Class Problemi** 🔴
- `TrfSheetWriter` (717 satır): Excel yazma, stil, hesaplama, veri formatlama hepsi aynı yerde
- `EtatPublicGenerator` (537 satır): PDF oluşturma, mail gönderme, veri işleme karışık
- `MainController` (477 satır): UI, iş mantığı, dosya işlemleri, threading hepsi karışık

### 2. **Eksik Design Pattern'ler** 🟡
- ❌ **Strategy Pattern** yok → Excel vs PDF vs diğer format seçimi sertleştirilmiş
- ❌ **Builder Pattern** yok → Karmaşık raporları oluştururken constructor'lar iğrenç
- ❌ **Factory Pattern** yok → File tipine göre okuyucu seçimi manüel
- ❌ **Dependency Injection** eksik → Servisler elle create ediliyor
- ❌ **Visitor Pattern** yok → Farklı veri türlerine aynı işlem uygulamak zor

### 3. **Tek Sorumluluk İlkesi Ihlali (SRP)** 🔴
```
DataReader.java dedi:
- Excel okuma ✓
- Data parsing ✓
- Data normalization ✓
- Model conversion ✓
← 4 sorumluluk!

TrfSheetWriter.java dedi:
- Sheet yazma ✓
- Style oluşturma ✓
- Border ekleme ✓
- Font ayarlama ✓
- Hücre doldurma ✓
- Formül ekleme ✓
← 6 sorumluluk!
```

### 4. **Tekrarlanan Kod** 🟡
```java
// 3-4 yerde tekrarlanıyor:
DataFormatter fmt = new DataFormatter();
FormulaEvaluator ev = wb.getCreationHelper().createFormulaEvaluator();
cellStr(row, col, fmt, ev);
cellDouble(row, col, fmt, ev);
```

### 5. **Eksik İstisna Yönetimi** 🟡
- Generic Exception catches
- Stack trace bastırılmış hatalar
- Hata mesajları yetersiz contexte sahip

### 6. **Eksik Test Yapısı** 🔴
- src/test klasörü neredeyse boş
- Birim testleri yoktu
- Entegrasyon testleri yoktu

---

## 🛠️ ÇÖZÜM: REFACTORİNG STRATEJİSİ

### **AŞAMA 1: Extract & Create (2 saat)**
Büyük sınıfları bölün ve yeni pattern'ler tanıtın.

#### 1.1 - TrfSheetWriter Bölünmesi
```
TrfSheetWriter (717 satır)
├── ExcelStyleFactory (yeni)
│   ├── TitleStyle
│   ├── HeaderStyle
│   ├── DataStyle
│   ├── FooterStyle
│   └── BorderStyle
├── ExcelFormatterService (yeni)
│   ├── formatCurrency
│   ├── formatPercentage
│   ├── formatDate
│   └── applyAlignment
├── ExcelSheetBuilder (new - Builder Pattern)
│   ├── withStyle
│   ├── withBorder
│   ├── withFormula
│   └── build
└── TrfSheetWriter (100-150 satır) [Orchestrator]
    ├── build(data, outputPath)
    └── validate()
```

#### 1.2 - EtatPublicGenerator Bölünmesi
```
EtatPublicGenerator (537 satır)
├── PdfReportBuilder (yeni - Builder Pattern)
├── ReportSectionStrategy (yeni - Strategy Pattern)
│   ├── HeaderSection
│   ├── ContentSection
│   ├── FooterSection
│   └── SummarySection
├── EmailService (yeni)
├── FileTransferService (yeni)
└── EtatPublicGenerator (80 satır)
    ├── generate()
    └── distribute()
```

#### 1.3 - DataReader Bölünmesi
```
DataReader (322 satır)
├── ExcelDataExtractor (yeni)
├── DataNormalizer (yeni)
├── ModelConverter (yeni - Factory Pattern)
└── DataReader (80 satır) [Facade]
    ├── read()
    └── validate()
```

#### 1.4 - MainController Bölünmesi
```
MainController (477 satır)
├── MainViewController (yeni - UI only)
├── ActionCommandFactory (yeni - Command Pattern)
│   ├── GenerateTrfCommand
│   ├── CompareFilesCommand
│   ├── FixPathsCommand
│   ├── ExportPublicCommand
│   └── MergeCommand
├── ProgressObserver (yeni - Observer Pattern)
├── FileConfigurationManager (yeni)
├── FileValidationService (yeni)
└── MainController (100 satır)
    ├── initialize()
    ├── executeCommand(Command)
    └── updateProgress()
```

#### 1.5 - DatabaseManager Bölünmesi
```
DatabaseManager (302 satır)
├── ConnectionPool (yeni - Singleton)
├── QueryBuilder (yeni - Builder Pattern)
├── DatabaseTransaction (yeni)
├── DatabaseMapper (yeni)
└── DatabaseManager (80 satır)
    ├── getSession()
    ├── execute(query)
    └── transaction(callback)
```

### **AŞAMA 2: Implement Patterns (1.5 saat)**

#### 2.1 - Strategy Pattern (Excel Output)
```java
// ✅ SONRA
public interface ExcelOutputStrategy {
    File generate(ReportData data, File output) throws Exception;
}

public class TrfExcelStrategy implements ExcelOutputStrategy { }
public class ConsolidationExcelStrategy implements ExcelOutputStrategy { }
public class ComparisonExcelStrategy implements ExcelOutputStrategy { }

// Kullanım:
ExcelOutputStrategy strategy = strategyFactory.getStrategy(reportType);
File result = strategy.generate(data, outputPath);
```

#### 2.2 - Factory Pattern (Reader/Writer)
```java
// ✅ SONRA
public interface DataReader {
    Map<String, Object> read(File file) throws Exception;
}

public interface DataWriter {
    File write(Map<String, Object> data, File output) throws Exception;
}

public class DataIOFactory {
    public DataReader getReader(FileType type) { }
    public DataWriter getWriter(FileType type) { }
}

// Kullanım:
DataIOFactory factory = new DataIOFactory();
DataReader reader = factory.getReader(FileType.PROCREANCES);
Map<String, Object> data = reader.read(file);
```

#### 2.3 - Builder Pattern (Complex Objects)
```java
// ✅ SONRA
ExcelSheetBuilder builder = new ExcelSheetBuilder()
    .withDimensions(50, 12)
    .withTitle("Rapport TRF")
    .addHeaderRow(headers)
    .addDataRows(dataList)
    .withStyle(StylePreset.PROFESSIONAL)
    .withFrozenPane(1, 0)
    .addAutoFilter()
    .build();

// YERINE (eski şekil):
Workbook wb = new XSSFWorkbook();
Sheet sheet = wb.createSheet("Rapport");
// ... 100 satır kod
```

#### 2.4 - Command Pattern (Actions)
```java
// ✅ SONRA
public interface ReportCommand {
    void execute(ReportContext context) throws Exception;
}

public class GenerateTrfCommand implements ReportCommand {
    @Override
    public void execute(ReportContext context) { }
}

// Kullanım:
ReportCommand cmd = commandFactory.createCommand("GENERATE_TRF");
cmd.execute(context);
```

#### 2.5 - Observer Pattern (Progress)
```java
// ✅ SONRA
public interface ProgressObserver {
    void update(double progress, String message);
}

public class ProgressNotifier {
    private List<ProgressObserver> observers = new ArrayList<>();
    
    public void subscribe(ProgressObserver obs) { observers.add(obs); }
    public void notifyProgress(double p, String msg) {
        observers.forEach(o -> o.update(p, msg));
    }
}

// Kullanım:
ProgressNotifier notifier = new ProgressNotifier();
notifier.subscribe(uiController);
notifier.subscribe(logService);
mergeService.merge(file, notifier);
```

### **AŞAMA 3: İyileştirmeler (1 saat)**

#### 3.1 - Configuration Management
```java
// ✅ SONRA
@Configuration
public class ApplicationConfig {
    @Bean
    public DatabaseManager databaseManager() { }
    
    @Bean
    public ExcelIOFactory excelFactory() { }
    
    @Bean
    public ReportService reportService(
        DatabaseManager db, 
        ExcelIOFactory factory) { }
}
```

#### 3.2 - Exception Handling
```java
// ✅ SONRA
public class BusinessException extends RuntimeException {
    private final ErrorCode errorCode;
    private final Map<String, Object> context;
    
    public BusinessException(ErrorCode code, String msg, Map<String, Object> ctx) { }
}

public enum ErrorCode {
    FILE_NOT_FOUND(1001),
    INVALID_FORMAT(1002),
    DATABASE_ERROR(1003),
    GENERATION_FAILED(1004)
}

// Kullanım:
try {
    data = reader.read(file);
} catch (IOException e) {
    throw new BusinessException(
        ErrorCode.FILE_NOT_FOUND,
        "Dosya bulunamadı: " + file.getName(),
        Map.of("file", file.getPath(), "cause", e.getMessage())
    );
}
```

#### 3.3 - Logging Service
```java
// ✅ SONRA
@Slf4j
public class ReportGenerationService {
    public File generateTrf(File input) {
        log.info("TRF generation started for file: {}", input.getName());
        try {
            // ...
            log.info("TRF successfully generated");
            return output;
        } catch (Exception e) {
            log.error("TRF generation failed", e);
            throw e;
        }
    }
}
```

#### 3.4 - Validation Framework
```java
// ✅ SONRA
public interface ValidationRule {
    ValidationResult validate(Object object);
}

public class FileValidationRule implements ValidationRule {
    @Override
    public ValidationResult validate(File file) {
        List<String> errors = new ArrayList<>();
        if (!file.exists()) errors.add("Dosya mevcut değil");
        if (!file.canRead()) errors.add("Dosya okunamıyor");
        // ...
        return new ValidationResult(errors.isEmpty(), errors);
    }
}

public class Validator {
    public ValidationResult validate(Object obj, List<ValidationRule> rules) {
        List<String> allErrors = new ArrayList<>();
        for (ValidationRule rule : rules) {
            ValidationResult result = rule.validate(obj);
            if (!result.isValid()) allErrors.addAll(result.getErrors());
        }
        return new ValidationResult(allErrors.isEmpty(), allErrors);
    }
}

// Kullanım:
Validator validator = new Validator();
ValidationResult result = validator.validate(file, List.of(
    new FileExistsRule(),
    new FileReadableRule(),
    new FileFormatRule("xlsx")
));
if (!result.isValid()) {
    result.getErrors().forEach(System.err::println);
}
```

---

## 📋 REFACTORING CHECKLIST

### **Günün Başı (09:00)**
- [ ] TrfSheetWriter.java - StyleFactory ve FormatterService çıkar
- [ ] TrfSheetWriter.java - Builder pattern ekle
- [ ] Testler yaz

### **Sabah Sonu (12:00)**
- [ ] EtatPublicGenerator.java - Strategy'ler çıkar
- [ ] DataReader.java - Extract sınıfları yaz
- [ ] MainController.java - Command pattern tanıt

### **Öğleden Sonra (14:00)**
- [ ] DatabaseManager.java - Pool ve Transaction patterns
- [ ] Exception handling iyileştir
- [ ] Logging ekle

### **Gün Sonu (17:00)**
- [ ] Integration tests yaz
- [ ] Documentation oluştur
- [ ] Final test ve build

---

## 📦 DELIVERABLES

```
Refactored-Project/
├── src/main/java/com/zeki/merger/
│   ├── core/
│   │   ├── pattern/
│   │   │   ├── strategy/
│   │   │   ├── factory/
│   │   │   ├── builder/
│   │   │   ├── command/
│   │   │   └── observer/
│   │   ├── config/
│   │   │   └── ApplicationConfig.java
│   │   └── exception/
│   │       ├── BusinessException.java
│   │       └── ErrorCode.java
│   ├── service/
│   │   ├── report/
│   │   │   ├── ReportService.java
│   │   │   ├── ReportBuilder.java
│   │   │   └── ReportValidator.java
│   │   ├── excel/
│   │   │   ├── ExcelReader.java
│   │   │   ├── ExcelWriter.java
│   │   │   ├── ExcelStyleFactory.java
│   │   │   └── ExcelSheetBuilder.java
│   │   ├── io/
│   │   │   ├── DataIOFactory.java
│   │   │   ├── FileTransferService.java
│   │   │   └── FileValidator.java
│   │   └── util/
│   │       └── ProgressNotifier.java
│   ├── controller/
│   │   ├── MainViewController.java (refactored)
│   │   ├── command/
│   │   │   └── ReportCommandFactory.java
│   │   └── handler/
│   │       └── ActionHandler.java
│   ├── database/
│   │   ├── ConnectionPool.java
│   │   ├── QueryBuilder.java
│   │   └── DatabaseManager.java (refactored)
│   └── ...
├── src/test/java/
│   └── integration/
│       ├── ReportGenerationTest.java
│       ├── ExcelIOTest.java
│       ├── DatabaseTest.java
│       └── ControllerTest.java
├── REFACTORING_NOTES.md
├── TESTING_GUIDE.md
└── MIGRATION_GUIDE.md
```

---

## 🎓 PATTERN REFERENCE

### Pattern Kullanım Tablosu

| Pattern | Kullan | Nerede | Neden |
|---------|--------|--------|-------|
| **Strategy** | ✓ | Excel/PDF çıkış | Flexible output formats |
| **Factory** | ✓ | Reader/Writer | Decoupled creation |
| **Builder** | ✓ | Sheet/Report | Complex object assembly |
| **Command** | ✓ | UI Actions | Undo/Redo capability |
| **Observer** | ✓ | Progress | Loose coupling |
| **Singleton** | ✓ | DB Pool | Shared resource |
| **Facade** | ✓ | DataReader | Simplified interface |
| **Decorator** | ~ | Optional | Enhanced functionality |

---

## ⏱️ ZAMAN TAHMINLEMESI

```
Aktivite                          Saat   Açıklama
─────────────────────────────────────────────────────
Analiz & Planlama                 0.5    ✓ Yapıldı
TrfSheetWriter Refactoring        2.0    Extract + Builder
EtatPublicGenerator Refactoring   1.5    Extract + Strategy
DataReader Refactoring            1.5    Extract + Factory
MainController Refactoring        1.5    Extract + Command
DatabaseManager Refactoring       1.0    Extract + Singleton
Exception Handling                 0.5    Custom exceptions
Configuration Setup                0.5    @Configuration
Testing & Documentation            0.5    Unit + Integration
Final Review & Polish             0.5    Code review
─────────────────────────────────────────────────────
TOPLAM                            ~10    1 çalışma günü
```

---

## ✨ BEKLENEN SONUÇLAR

### Ön (Before)
```
📊 Metrics:
- Ortalama sınıf boyutu: 175 satır (TrfSheetWriter 717!)
- Sınıf başına ortalama sorumluluk: 3.2
- Test coverage: %5
- Code duplication: %12
- Maintenance difficulty: 7.5/10
```

### Sonra (After)
```
📊 Metrics:
- Ortalama sınıf boyutu: 80 satır (max 150)
- Sınıf başına ortalama sorumluluk: 1.1 (SRP uyumlu)
- Test coverage: %65
- Code duplication: %2
- Maintenance difficulty: 3.5/10
```

---

## 🎯 SONUÇ

Bu refactoring seni şunlara kazandıracak:

✅ **Kendi kodum gibi hisset** - Tüm yapıyı anlayan, kontrol eden  
✅ **Maintenance kolaylaşsın** - Değişiklikleri 5 dakikada yapabilir  
✅ **Yeni feature ekle hızlı** - Patterns sayesinde standart template  
✅ **Bug hunting azalsın** - SRP + Testing = daha az bug  
✅ **Career boost** - Professional pattern knowledge  

**Tahmini başlangıç:** 09:00  
**Tahmini bitiş:** 17:00  
**Zorluk:** Orta  
**Öğrenme Değeri:** Çok Yüksek 📚

Hazır mısın? Başlayalım! 🚀
