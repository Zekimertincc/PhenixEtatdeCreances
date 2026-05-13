# PhenixEtatDeCreances — Code Overview

**Proje Tipi:** JavaFX 17 desktop app (Maven, Java 17+)  
**Ana Kullanım:** Cabinet Phénix tarafından müşteri alacak takibi. Excel ve PDF çıktıları üretir, PROCREANCES sistemi ile karşılaştırma yapar, TRF (Transfert) belgesi oluşturur.  
**Build:** `mvn clean package` → `target/app-shaded.jar`  
**Çalıştırma:** `java -jar target/app-shaded.jar`

---

## Paket Mimarisi

```
com.zeki.merger/
├── App.java / Launcher.java        # JavaFX entry point
├── AppConfig.java                  # Sabit konfigürasyon değerleri
├── AppPreferences.java             # Kullanıcı tercihlerini java.util.prefs ile saklar
│
├── core/
│   ├── config/ApplicationConfig   # Singleton DI hub (tüm servisler buradan alınır)
│   └── exception/
│       ├── ErrorCode              # Hata kodları enum (1001–1008)
│       └── BusinessException      # RuntimeException + ErrorCode + context map
│
├── controller/
│   ├── MainController             # JavaFX ana ekran (FXML bağlı)
│   ├── DashboardController        # Dashboard tab (SQLite özet)
│   └── command/
│       ├── ReportCommand          # Interface: execute(context, notifier) → File
│       ├── GenerateTrfCommand     # TRF üretimi komutu
│       └── ReportCommandFactory   # Komut adı → ReportCommand factory
│
├── service/
│   ├── MergeService               # Ana konsolidasyon (Créances → global xlsx)
│   ├── EtatPublicGenerator        # Şirket başına xlsx + PDF üretimi
│   ├── EspacePartageFixer         # EspacePartagé yollarını düzeltir
│   ├── ProcreancesComparator      # PROCREANCES vs ConsolidationGenerale karşılaştırması
│   ├── FolderScanner              # Şirket klasörlerini tarar
│   ├── ExcelReader                # Domain-specific okuyucu (Créances formatı)
│   ├── ExcelWriter                # Domain-specific yazıcı (global output)
│   ├── TrfWriter                  # Eski TRF yazıcı (TrfSheetWriter ile birlikte)
│   ├── ComparisonResult           # PROCREANCES karşılaştırma sonucu veri modeli
│   ├── DiffRow / UnmatchedProcRow / UnmatchedConsoRow  # Karşılaştırma satır modelleri
│   │
│   ├── excel/                     # [YENİ] Genel amaçlı Excel yardımcıları
│   │   ├── ExcelStyleFactory      # XSSFCellStyle fabrika (header/data/money/total)
│   │   ├── ExcelSheetBuilder      # Fluent builder — basit tablo sheets için
│   │   └── ExcelFormatterService  # Hücre okuma/yazma + colLetter() yardımcıları
│   │
│   ├── report/                    # [YENİ] Strategy Pattern
│   │   ├── ReportStrategy         # Interface: generate(data, outputPath) → File
│   │   ├── ExcelReportStrategy    # XLSX implementasyonu (ExcelSheetBuilder kullanır)
│   │   ├── PdfReportStrategy      # Stub — EtatPublicGenerator.writePdf() domain PDF'i işler
│   │   └── ReportStrategyFactory  # Format string → ReportStrategy factory
│   │
│   ├── io/                        # [YENİ] Factory Pattern
│   │   └── DataIOFactory          # Uzantı → FileReader/FileWriter factory
│   │
│   ├── data/                      # [YENİ] Veri dönüşüm yardımcıları
│   │   ├── DataExtractor          # Apache POI hücre okuma (tekrar eden kodu ortadan kaldırır)
│   │   ├── DataNormalizer         # İsim normalizasyonu, fuzzy match, sayısal yuvarlama
│   │   └── DataConverter          # Ham satır maplerini CreanceRow modeline çevirir
│   │
│   └── util/                      # [YENİ] Observer Pattern
│       ├── ProgressObserver       # Interface: onProgressUpdate / onCompleted / onFailed
│       └── ProgressNotifier       # Observer listesi yönetir + BiConsumer adapter
│
├── trf/                           # TRF (Transfert) alt sistemi
│   ├── DataReader                 # 3 giriş dosyasını okur (ConsoGénérale, Listing, Tableau)
│   ├── TrfCalculator              # Hesaplama motoru: ConsolidationRow → ClientSummary
│   ├── TrfGeneratorService        # Orkestratör: okuma → hesap → yazma
│   ├── TrfSheetWriter             # 4 sayfalı TRF workbook'u yazar (717 satır)
│   └── model/
│       ├── ClientInfo             # Listing'den gelen IBAN / NonComp / BIC
│       ├── ClientSummary          # Hesaplanmış müşteri özeti (tüm TRF sütunları)
│       └── ConsolidationRow       # Ham ConsolidationGenerale satırı + parseFrenchDouble()
│
├── db/
│   ├── DatabaseManager            # SQLite singleton (JDBC, ~/.cabinet_phenix/data.db)
│   └── CompanyRecord              # DB'den okunan şirket kaydı
│
└── model/
    └── CreanceRow                 # Tek alacak satırı (societe + cellValues + rowIndex)
```

---

## Nasıl Çalışır — Ana Akışlar

### 1. Konsolidasyon (`▶ CONSOLIDER` butonu)

```
MainController.run()
  └─ MergeService.merge(rootFolder, outputFolder, progress)
       ├─ FolderScanner.scan(rootFolder)     → List<CompanyFile>
       ├─ ExcelReader.readFiltered(company, file) → List<CreanceRow>  (her şirket için)
       └─ ExcelWriter.write(allRows, outputFile)  → global xlsx
```

Her şirket için `AppConfig.FILTER_COLUMN_INDEX` (kolon S=18) dolu olan satırlar alınır. Boş formül hücreleri atlanır (`hasRealData()`). Çıktı: `etat_creances_global.xlsx`.

### 2. TRF Üretimi (`Générer TRF` butonu)

```
MainController.generateTrf()
  └─ TrfGeneratorService.generate(consoFile, listingFile, tableauFile, outputFolder)
       ├─ DataReader.readAllConsolidationRows(consoFile)    → List<ConsolidationRow>
       ├─ DataReader.readClientInfoMap(listingFile)         → Map<normName, ClientInfo>
       ├─ DataReader.readPreviousBalances(tableauFile)      → Map<normName, Double>
       ├─ TrfCalculator.buildClientSummaries(...)          → List<ClientSummary>
       ├─ DatabaseManager.upsertCompany / replaceTrfSummary  (kalıcı kayıt)
       └─ TrfSheetWriter.write(allRows, summaries, outFile) → TRF_MM_YYYY.xlsx
```

**TRF workbook 4 sayfa içerir:**
- `Consolidation` — kaynak verinin aynısı + müşteri bazında SUBTOTAL satırları
- `Feuil1` — müşteri başına özet satır (26 sütun)
- `TRF` — ana transfer belgesi + 5 alt bölüm (Virements, Manuelles, NonComp, CompPartielle, Debiteurs)
- `Feuil3` — boş, referans format gereksinimi

### 3. Etat Public Üretimi (`États Publics` butonu)

```
MainController.generateEtatPublic()
  └─ EtatPublicGenerator.generate(rootFolder, progress)
       ├─ FolderScanner.scan(rootFolder) → List<CompanyFile>
       └─ (her şirket için):
            ├─ resolveDestDir()  → "Espace partagé/Etat des créances" veya fallback
            ├─ Eski L_ETAT_ dosyalarını sil
            ├─ generateForClient()
            │    ├─ Kaynak "Créances" sheet'inden 16. satırdan veri oku
            │    ├─ writeOutput() → L_ETAT_DE_CREANCES_[şirket].xlsx
            │    └─ writePdf()   → L_ETAT_DE_CREANCES_[şirket].pdf
```

`SOURCE_COL_MAP = {0,1,2,3,5,6,7,8,17,9,10}` — kaynak kolon indekslerini 11 çıktı kolonuna eşler. PDF iText7 ile landscape A4 üretilir.

### 4. PROCREANCES Karşılaştırması (`Comparer des fichiers Excel`)

```
ProcreancesComparator.compare(procFile, consoFile, outputFolder, progress)
  ├─ Her iki dosyayı oku, müşteri adına göre topla
  ├─ normalize() ile fuzzy match yap
  ├─ 3 değeri karşılaştır: Hono.TTC / Disponible / Reversement
  ├─ Tolerance = 0.05 (küçük yuvarlama farklarını görmezden gel)
  └─ comparison_PROCREANCES_vs_CONSO_[timestamp].xlsx yaz
       ├─ sheet "Récapitulatif" — tüm eşleşmeler
       ├─ sheet "Écarts"        — sadece fark olanlar (kırmızı/yeşil)
       └─ sheet "Non appariés"  — eşleşemeyen satırlar
```

### 5. EspacePartagé Yol Düzeltmesi (`Corriger EspacePartagé`)

```
EspacePartageFixer.fix(rootFolder, progress)
  └─ CorrespondanceClient-EspacePartage.xlsx'i okur
     Şirket adı → EspacePartagé yolu eşlemesini günceller / yazar
```

---

## Yeni Katmanlar (Refactoring ile eklenenler)

### Exception Handling (`core/exception/`)

Mevcut kodda `catch (Exception e)` yaygındı, hata mesajları bağlamsızdı. Yeni yaklaşım:

```java
// Önce (eski):
} catch (IOException e) {
    System.err.println("error");
}

// Sonra (yeni):
} catch (IOException e) {
    throw new BusinessException(
        ErrorCode.FILE_NOT_FOUND,
        "ConsolidationGenerale okunamadı: " + file.getName(),
        Map.of("file", file.getPath(), "cause", e.getMessage())
    );
}
```

`ErrorCode` enum'u hem hata kodu (int) hem de varsayılan mesaj taşır. `BusinessException.getDetailedMessage()` tüm bağlamı tek stringde verir.

### Observer Pattern (`service/util/`)

Mevcut `BiConsumer<Double, String>` progress callback'leri JavaFX'e sıkı bağlıydı. `ProgressNotifier`, bu callback'i saran bir Observer hub'ı:

```java
ProgressNotifier notifier = new ProgressNotifier();
notifier.subscribe(uiController);   // JavaFX güncelleme
notifier.subscribe(logService);     // opsiyonel loglama

// Geriye dönük uyum: eski BiConsumer API'leriyle çalışır
service.generate(consoFile, listingFile, tableauFile, outputFolder,
    notifier.asBiConsumer());
```

`ProgressObserver` interface üç metot sunar: `onProgressUpdate`, `onCompleted`, `onFailed`.

### Builder Pattern (`service/excel/`)

`ExcelSheetBuilder` yeni feature'lar için temiz bir API sağlar. Mevcut `TrfSheetWriter` ve `EtatPublicGenerator` kendi `Styles` inner class'larını kullanmaya devam eder — bu sınıflar çok domain-specific.

```java
Workbook wb = new ExcelSheetBuilder("Rapport")
    .withFrozenPane(1, 0)
    .withAutoFilter(true)
    .addHeaderRow(List.of("Client", "Montant"))
    .addDataRow(List.of("ACME SAS", 12345.67))
    .build();
```

`ExcelStyleFactory` → `ExcelFormatterService` → `ExcelSheetBuilder` şeklinde katmanlı:
- `ExcelStyleFactory`: XSSFCellStyle nesneleri üretir (header/data/money/total/date)
- `ExcelFormatterService`: hücre değeri okuma/yazma + `columnLetter()` util
- `ExcelSheetBuilder`: fluent sheet builder

### Strategy Pattern (`service/report/`)

Farklı çıktı formatları için açık/kapalı prensip:

```java
ReportStrategyFactory factory = new ReportStrategyFactory();
ReportStrategy strategy = factory.getStrategy("XLSX");  // veya "PDF"
File output = strategy.generate(data, outputFile);

// Yeni format eklemek için:
factory.register("CSV", new CsvReportStrategy());  // MainController değişmez
```

`data` map'i: `{"headers": List<String>, "rows": List<List<Object>>, "title": String}`

### Factory Pattern (`service/io/`)

`DataIOFactory` dosya uzantısını okuyucu/yazıcıya eşler:

```java
DataIOFactory factory = new DataIOFactory();
Map<String, Object> data = factory.getReaderByFile(excelFile).read(excelFile);
factory.getWriterByFile(outputFile).write(data, outputFile);
```

Not: Mevcut `ExcelReader` (domain-specific, `List<CreanceRow>` döner) ile `DataIOFactory.FileReader` farklı abstraction'lardır. Domain okuyucuları servis sınıflarından doğrudan çağrılmaya devam eder.

### Data Layer (`service/data/`)

Üç sınıf `DataReader`, `ProcreancesComparator`, `EtatPublicGenerator`'da tekrarlanan kodu merkezileştirir:

| Sınıf | Sorumluluk |
|-------|-----------|
| `DataExtractor` | POI `Row` → String/double okuma (null-safe) |
| `DataNormalizer` | normalize(), fuzzyMatch(), sanitizeFileName(), normalizeAmount() |
| `DataConverter` | Ham satır map'ini `CreanceRow` modeline çevirir |

`DataNormalizer.normalize()` canonical implementasyondur — `DataReader.normalize()` ve `EtatPublicGenerator.normalize()` ile özdeştir.

### Command Pattern (`controller/command/`)

Her UI aksiyonu bir `ReportCommand` olur. Faydası: MainController'ın `executeCommand("GENERATE_TRF")` demesi yeterli, işin detayı komut içinde.

```java
// Mevcut MainController içinde:
Map<String, Object> ctx = Map.of(
    "consoFile",    new File(AppPreferences.getTrfConso()),
    "listingFile",  new File(AppPreferences.getTrfListing()),
    "tableauFile",  new File(AppPreferences.getTrfTableau()),
    "outputFolder", new File(AppPreferences.getOutputFolder())
);
ReportCommand cmd = commandFactory.getCommand("GENERATE_TRF");
cmd.execute(ctx, progressNotifier);
```

Şu an sadece `GenerateTrfCommand` implement edildi. Diğer 4 aksiyon (EtatPublic, Compare, Fix, Consolidate) `ReportCommandFactory`'de comment olarak bırakıldı — sıradaki adım.

### ApplicationConfig (`core/config/`)

Double-checked locking Singleton. Tüm bağımlılıkları tek yerden verir:

```java
ApplicationConfig cfg = ApplicationConfig.getInstance();
cfg.getDatabaseManager()
cfg.getReportStrategyFactory()
cfg.getDataIOFactory()
cfg.getReportCommandFactory()
```

---

## Veri Modelleri

### CreanceRow
Tek alacak satırını temsil eder. Immutable.
- `societe`: şirket adı
- `cellValues`: List<Object> — kaynak Excel'deki ham değerler (aynı sütun sırası)
- `originalRowIndex`: kaynak dosyadaki 0-based satır indeksi

### ConsolidationRow
ConsolidationGenerale.xlsx'den okunan bir satır.
- `values`: List<Object> — 26 sütun
- `headerRow`: boolean (ilk satır mı?)
- `parseFrenchDouble(String)`: "1 234,56 €" → 1234.56 (static, proje genelinde kullanılır)

### ClientSummary
TRF hesaplamalarının çıktısı. Bir müşteriye ait tüm finansal özeti içerir.
- `clientName`, `clientCode`, `iban`
- Finansal alanlar: `creancePrincipale`, `recouvreEtFacture`, `commissions`, `penalites`, `sommesCzPhenix`, `montantAFacturerTtc`, `sommesAReverserSrc/Final`, `nousDoit_Prec/ApreFacturation`, `encaissementsParCompensation`
- Durum bayrakları: `isNonCompensation()`, `isPartiallyCompensated()`, `isDebtor()`, `needsManualVirement()`

### ClientInfo
Listing dosyasından (LISTING_CABINET_PHENIX) okunan müşteri meta verisi.
- `name`, `code`, `nonComp`, `iban`, `bic`

---

## Kritik İş Mantığı

### TRF Hesaplama Formülleri (TrfSheetWriter)

TRF sayfası, müşteri başına 12 sütun üretir (B-L):

```
E = C + D                         (Nous Doit Maintenant = Montant Facturer + Nous Doit Précédemment)

Compensasyon durumuna göre:
  NonComp: F = B,  G = 0,  H = E - G
  Normal:  F = IF(B=0,0, IF(B<E, 0, B-E))   (Sommes à Reverser)
           G = IF(B=0,0, IF(B>E, E, B))     (Encaissements par Compensation)
           H = E - G                         (Nous Doit Après Facturation)
```

Excel formülleri hücrelere yazılır (evaluate edilmez), böylece dosyayı açan Excel kullanıcısı formülleri görebilir.

### Fuzzy Client Name Matching (DataReader / DataNormalizer)

İsim eşleştirme iki adımlı:
1. `normalize()` → küçük harf, aksan kaldır, fazla boşluk sil
2. Exact match → bulunamazsa `contains` partial match

Bu yaklaşım "ACME SAS" ile "ACME" gibi isimleri eşleştirir. `TOLERANCE = 0.05` sayısal karşılaştırmalarda yuvarlama farklarını absorbe eder.

### EtatPublicGenerator — Hedef Klasör Çözümleme

3 seviyeli fallback:
1. `[şirket]/[Espace partagé]/[Etat des créances]` → varsa kullan
2. `[şirket]/Etat des créances` → doğrudan alt klasör
3. `[şirket]/Etat des créances` → oluştur

Normalize edilmiş isim karşılaştırması ("espace" + "partage" içeriyorsa).

---

## Veritabanı (SQLite)

Konum: `~/.cabinet_phenix/data.db`  
Bağlantı: `DatabaseManager` (Singleton, synchronized metotlar)

**Tablolar:**
- `companies`: `id, name, code, last_updated`
- `trf_summaries`: TRF sonuçları, `company_id` FK'sı ile ilişkili
- `creance_rows`: Konsolidasyon satırları (MergeService tarafından yazılır)

`DatabaseManager.getInstance()` null dönebilir — TRF oluşturma DB'siz de çalışır (DB hatası işlemi durdurmaz, sadece loglar).

---

## Konfigürasyon

### AppConfig (sabitler)
```java
AppConfig.CREANCES_SHEET_NAME      = "Créances"         // kaynak sheet adı
AppConfig.FILTER_COLUMN_INDEX      = 18                  // kolon S
AppConfig.ETAT_PUBLIC_FILENAME_PREFIX = "L_ETAT_DE_CREANCES_"
AppConfig.ESPACE_PARTAGE_FILENAME  = "CorrespondanceClient-EspacePartage.xlsx"
```

### AppPreferences (kullanıcı tercihleri)
`java.util.prefs.Preferences` ile kalıcı. UI "Configuration des fichiers" diyalogu ile değiştirilir.
- `getMergeRoot()` / `getOutputFolder()`
- `getTrfConso()` / `getTrfListing()` / `getTrfTableau()`
- `getProcreancesPath()`

---

## Yeni Feature Eklemek

### Yeni bir rapor formatı eklemek (örn. CSV)

1. `service/report/CsvReportStrategy.java` yaz (`ReportStrategy` implement et)
2. `ReportStrategyFactory` constructor'ına `register("CSV", new CsvReportStrategy())` ekle
3. Bitişi budur. Mevcut hiçbir sınıf değişmez.

### Yeni bir UI aksiyonu eklemek (örn. "Export to Dropbox")

1. `controller/command/ExportDropboxCommand.java` yaz (`ReportCommand` implement et)
2. `ReportCommandFactory` constructor'ında uncomment satırına benzer şekilde `register(new ExportDropboxCommand())` ekle
3. `MainController.initialize()`'da butona `executeCommand("EXPORT_DROPBOX")` bağla

### Yeni bir giriş formatı desteklemek (örn. CSV okuma)

1. `DataIOFactory`'de `registerReader("csv", file -> { ... })` lambda ekle
2. `registerWriter("csv", (data, out) -> { ... })` ekle

### Yeni bir müşteri alanı eklemek

1. `ClientSummary.java`'ya alan ekle
2. `TrfCalculator.buildClientSummaries()`'da hesaplamayı yaz
3. `TrfSheetWriter.writeFeuil1Sheet()` ve `writeTrfSheet()`'e sütun ekle
4. `DatabaseManager` şemasını güncelle (gerekirse)

---

## Threading Modeli

- Tüm servis çağrıları `ExecutorService` (single daemon thread `merge-worker`) üzerinde çalışır
- JavaFX güncellemeleri `Platform.runLater(() -> ...)` ile UI thread'e gönderilir
- `ProgressNotifier.asBiConsumer()` → mevcut `BiConsumer<Double, String>` callback'leriyle uyumlu köprü
- `DatabaseManager` metotları `synchronized` — background thread ile JavaFX thread aynı anda DB'ye erişebilir

---

## Bağımlılıklar (pom.xml)

| Kütüphane | Kullanım |
|-----------|---------|
| `org.apache.poi:poi-ooxml` | Excel okuma/yazma (xlsx) |
| `org.apache.poi:poi` | Excel okuma/yazma (xls) |
| `com.itextpdf:kernel`, `layout` | PDF üretimi (EtatPublicGenerator) |
| `org.xerial:sqlite-jdbc` | Embedded SQLite |
| `org.openjfx:javafx-*` | UI framework |

---

## Dosya Sayısı

| Kategori | Dosya Sayısı |
|----------|-------------|
| Önceki | 28 Java sınıf |
| Sonrası | 47 Java sınıf |
| Yeni | 19 sınıf (tümü `mvn compile` ile doğrulandı) |

Yeni sınıfların hiçbiri mevcut kodu değiştirmemiştir — tamamen additive. Mevcut testler (varsa) kırılmaz.
