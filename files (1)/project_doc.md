# Documentation Complète du Projet — Cabinet Phénix : Etat des Créances Merger

> **Version :** 1.0.0  
> **Technologie principale :** Java 17+, JavaFX 21, Apache POI 5.2.5, iText7, SQLite  
> **Build :** Maven  
> **Date de rédaction :** 2026-05-12

---

## Table des Matières

1. [Vue d'ensemble du projet](#1-vue-densemble-du-projet)
2. [Architecture générale](#2-architecture-générale)
3. [Structure des packages](#3-structure-des-packages)
4. [Modèle de données](#4-modèle-de-données)
5. [Couche de persistance (Base de données SQLite)](#5-couche-de-persistance-base-de-données-sqlite)
6. [Couche service — Pipeline de consolidation](#6-couche-service--pipeline-de-consolidation)
7. [Couche service — Génération TRF](#7-couche-service--génération-trf)
8. [Couche service — États Publics](#8-couche-service--états-publics)
9. [Couche service — Comparaison PROCREANCES](#9-couche-service--comparaison-procreances)
10. [Couche service — Nouveaux services refactorisés](#10-couche-service--nouveaux-services-refactorisés)
11. [Couche contrôleur (JavaFX)](#11-couche-contrôleur-javafx)
12. [Patrons de conception appliqués](#12-patrons-de-conception-appliqués)
13. [Interface utilisateur (JavaFX/FXML)](#13-interface-utilisateur-javafxfxml)
14. [Configuration et préférences](#14-configuration-et-préférences)
15. [Dépendances Maven](#15-dépendances-maven)
16. [Construction et exécution](#16-construction-et-exécution)
17. [Flux de données complets (Diagrammes)](#17-flux-de-données-complets-diagrammes)
18. [Glossaire métier](#18-glossaire-métier)
19. [Points d'attention et limitations connues](#19-points-dattention-et-limitations-connues)

---

## 1. Vue d'ensemble du projet

**Cabinet Phénix — Etat des Créances Merger** est une application de bureau JavaFX destinée aux équipes du cabinet de recouvrement Phénix. Elle automatise les traitements mensuels sur des fichiers Excel liés aux créances des clients.

### Fonctionnalités principales

| Fonctionnalité | Description |
|---|---|
| **CONSOLIDER** | Scanne l'arborescence des dossiers clients, lit chaque fichier "Etat des créances", filtre les lignes actives (colonne S non vide) et fusionne tout dans un fichier `etat_creances_global_[timestamp].xlsx` |
| **Générer TRF** | Lit 3 fichiers sources (ConsolidationGénérale, Listing Cabinet Phénix, Tableau de Bord), calcule les montants de compensation/reversement et produit le classeur TRF 4 feuilles |
| **États Publics** | Pour chaque client, génère un fichier `L_ETAT_DE_CREANCES_[client].xlsx` + `.pdf` dans son dossier "Espace Partagé" |
| **Comparer des fichiers Excel** | Compare l'export PROCREANCES vs la ConsolidationGénérale, identifie les écarts et produit un rapport Excel à 3 feuilles |
| **Corriger EspacePartagé** | Met à jour les chemins dans le fichier de correspondance `CorrespondanceClient-EspacePartage.xlsx` |

### Domaine métier (Cabinet Phénix)

Cabinet Phénix est une société de recouvrement. Pour chaque client qu'elle représente :
- Elle collecte des créances (montants dus par des débiteurs)
- Elle prélève des commissions sur les montants recouvrés
- Chaque mois, elle calcule ce qu'elle doit reverser au client (encaissements − montant à facturer)
- Le **TRF** (Transfert et Reversement Financier) est le document récapitulatif mensuel des virements et compensations

---

## 2. Architecture générale

L'application suit une **architecture en couches** avec séparation claire des responsabilités :

```
┌─────────────────────────────────────────────────────────┐
│                    INTERFACE UTILISATEUR                 │
│          JavaFX / FXML (App.java, MainController,        │
│                  DashboardController)                    │
└──────────────────────┬──────────────────────────────────┘
                       │
┌──────────────────────▼──────────────────────────────────┐
│                  COUCHE COMMANDE                         │
│     controller/command/ (Command Pattern)                │
│  ReportCommand ← GenerateTrfCommand, ...                 │
│          ReportCommandFactory                            │
└──────────────────────┬──────────────────────────────────┘
                       │
┌──────────────────────▼──────────────────────────────────┐
│                  COUCHE SERVICE (Domaine)                │
│  MergeService │ TrfGeneratorService │ EtatPublicGenerator│
│  ProcreancesComparator │ EspacePartageFixer              │
└──────────┬─────────────────────────────┬────────────────┘
           │                             │
┌──────────▼──────────┐    ┌────────────▼────────────────┐
│   COUCHE TRF        │    │  COUCHE SERVICE — UTILITAIRES│
│  trf/               │    │  service/excel/  (styles)    │
│  DataReader         │    │  service/data/   (extraction)│
│  TrfCalculator      │    │  service/io/     (factory)   │
│  TrfSheetWriter     │    │  service/report/ (strategy)  │
│  TrfGeneratorService│    │  service/util/   (observer)  │
└──────────┬──────────┘    └────────────┬────────────────┘
           │                            │
┌──────────▼────────────────────────────▼────────────────┐
│               COUCHE PERSISTANCE (SQLite)               │
│                   db/DatabaseManager                    │
│       companies │ creance_rows │ trf_summaries          │
└─────────────────────────────────────────────────────────┘
```

### Principes architecturaux

1. **Séparation MVC** : L'interface JavaFX (controllers) ne contient pas de logique métier. Elle délègue aux services.
2. **Thread Safety** : Toutes les opérations longues s'exécutent dans un `ExecutorService` (thread de fond). Les mises à jour UI se font via `Platform.runLater()`.
3. **Couplage lâche** : Les patterns Strategy, Command et Observer permettent d'ajouter de nouvelles fonctionnalités sans modifier le code existant.
4. **Singleton contrôlé** : `DatabaseManager` et `ApplicationConfig` sont des singletons thread-safe (double-checked locking).

---

## 3. Structure des packages

```
com.zeki.merger/
│
├── App.java                          # Point d'entrée JavaFX
├── Launcher.java                     # Wrapper pour le JAR exécutable
├── AppConfig.java                    # Constantes de configuration
├── AppPreferences.java               # Préférences utilisateur persistées (Java Preferences API)
│
├── controller/
│   ├── MainController.java           # Contrôleur FXML principal (5 actions)
│   ├── DashboardController.java      # Contrôleur FXML du tableau de bord (DB viewer)
│   └── command/                      # Patron Commande
│       ├── ReportCommand.java        # Interface commande
│       ├── GenerateTrfCommand.java   # Commande : génération TRF
│       └── ReportCommandFactory.java # Registre de commandes
│
├── core/
│   ├── config/
│   │   └── ApplicationConfig.java   # Singleton DI — assemble les services
│   └── exception/
│       ├── BusinessException.java   # Exception métier structurée
│       └── ErrorCode.java           # Enum des codes d'erreur (1001–1008)
│
├── db/
│   ├── DatabaseManager.java         # Singleton SQLite — 3 tables
│   └── CompanyRecord.java           # DTO pour les requêtes dashboard
│
├── model/
│   └── CreanceRow.java              # Objet valeur : une ligne de créance
│
├── service/
│   ├── MergeService.java            # Pipeline : scan→lire→filtrer→écrire
│   ├── EtatPublicGenerator.java     # Génère L_ETAT_DE_CREANCES_*.xlsx + .pdf
│   ├── ProcreancesComparator.java   # Compare PROCREANCES vs ConsolidationGénérale
│   ├── EspacePartageFixer.java      # Corrige les chemins EspacePartagé
│   ├── FolderScanner.java           # Scan de l'arborescence des dossiers clients
│   ├── ExcelReader.java             # Lecture filtrée (colonne S) des fichiers clients
│   ├── ExcelWriter.java             # Écriture du fichier de consolidation global
│   ├── TrfWriter.java               # Écriture de l'export TRF depuis MergeService
│   ├── ComparisonResult.java        # Résultat de comparaison
│   ├── DiffRow.java                 # Record : une ligne de différence
│   ├── UnmatchedProcRow.java        # Record : client non apparié (PROCREANCES)
│   ├── UnmatchedConsoRow.java       # Record : client non apparié (Conso)
│   │
│   ├── data/                        # Extraction et normalisation de données
│   │   ├── DataExtractor.java       # Lecture bas niveau Apache POI
│   │   ├── DataNormalizer.java      # Normalisation strings/montants
│   │   └── DataConverter.java       # Conversion Map→CreanceRow
│   │
│   ├── excel/                       # Couche Excel réutilisable
│   │   ├── ExcelStyleFactory.java   # Fabrique de styles POI
│   │   ├── ExcelSheetBuilder.java   # Constructeur fluide de feuilles Excel
│   │   └── ExcelFormatterService.java # Lecture/écriture de valeurs POI
│   │
│   ├── io/                          # Entrées/sorties génériques
│   │   └── DataIOFactory.java       # Fabrique lecteurs/écrivains par extension
│   │
│   ├── report/                      # Stratégies de génération de rapports
│   │   ├── ReportStrategy.java      # Interface stratégie
│   │   ├── ExcelReportStrategy.java # Stratégie Excel (.xlsx)
│   │   ├── PdfReportStrategy.java   # Stratégie PDF (stub — à implémenter)
│   │   └── ReportStrategyFactory.java # Registre des stratégies
│   │
│   └── util/                        # Utilitaires transverses
│       ├── ProgressObserver.java    # Interface observateur de progression
│       └── ProgressNotifier.java    # Notificateur observable
│
└── trf/                             # Module TRF complet
    ├── DataReader.java              # Lecture des 3 fichiers sources TRF
    ├── TrfCalculator.java           # Logique pure de calcul TRF
    ├── TrfGeneratorService.java     # Orchestrateur : lire→calculer→écrire
    ├── TrfSheetWriter.java          # Écriture du classeur Excel TRF (4 feuilles)
    └── model/
        ├── ClientInfo.java          # DTO : données du Listing (IBAN, NonComp...)
        ├── ClientSummary.java       # Objet riche : tous les champs TRF d'un client
        └── ConsolidationRow.java    # Wrapper d'une ligne de la feuille Consolidation
```

---

## 4. Modèle de données

### 4.1 `CreanceRow` — Ligne de créance brute

Objet **valeur immuable** représentant une ligne filtrée depuis le fichier Excel d'un client.

```java
CreanceRow {
    String       societe          // Nom de la société (dossier client)
    List<Object> cellValues       // Valeurs brutes dans l'ordre des colonnes source
    int          originalRowIndex // Index 0-based dans la feuille source
}
```

Utilisé dans le pipeline `ExcelReader → MergeService → ExcelWriter → DatabaseManager`.

---

### 4.2 `ConsolidationRow` — Ligne de la ConsolidationGénérale

Wrapper pour une ligne de la feuille "Consolidation" de `ConsolidationGenerale.xlsx`.

```java
ConsolidationRow {
    List<Object> values      // Valeurs des colonnes A–Z (26 colonnes minimum)
    boolean      headerRow   // True si row index == 0 (ligne d'en-tête)
    boolean      totalRow    // True si colonne A commence par "Total "
    String       clientName  // Nom du client (extrait de col A, sans "Total ")
}
```

Méthode clé : `parseFrenchDouble(String s)` — parse les nombres français (`"1 234,56 €"`, `"1.234,56"`, etc.)

---

### 4.3 `ClientInfo` — Données du Listing Cabinet Phénix

DTO lu depuis `LISTING_CABINET_PHENIX_pour_ZEKI.xls`, feuille "Feuil1".

```java
ClientInfo {
    String  name              // Nom du client (col C, index 2)
    String  code              // Code client (col D, index 3)
    String  nonCompensation   // "OUI" ou "" (col U, index 20)
    String  iban              // IBAN bancaire (col V, index 21)
    String  bic               // Code BIC (col W, index 22)
    boolean paiementParCheque // Déduit : code purement numérique → paiement par chèque
}
```

`isNonCompensation()` retourne `true` si le champ `nonCompensation` vaut `"OUI"` (insensible à la casse).

---

### 4.4 `ClientSummary` — Résumé TRF complet d'un client

Objet riche qui accumule toutes les données nécessaires au calcul TRF. Enrichi progressivement par `TrfCalculator`.

#### Champs sources (depuis ConsolidationGénérale — feuille "Consolidation")

| Champ Java | Colonne Excel | Description |
|---|---|---|
| `creancePrincipale` | H (index 7) | Montant total des créances confiées |
| `recouvreEtFacture` | I (index 8) | Montant recouvré et déjà facturé |
| `penalites` | L (index 11) | Pénalités de retard |
| `dontEnAttente` | P (index 15) | Dont en attente de facturation |
| `fraisProcedure` | R (index 17) | Frais de procédure judiciaire |
| `recouvreTotol` | S (index 18) | Recouvré total (encaissements bruts) |
| `dejaFacture` | T (index 19) | Déjà facturé aux clients |
| `depuisLeDebut` | U (index 20) | Total depuis le début du mandat |
| `commissions` | V (index 21) | Commissions Phénix |
| `penalits` | W (index 22) | Pénalités (variante de calcul) |
| `sommesCzPhenix` | X (index 23) | **Encaissements CZ Phénix** — total des fonds collectés |
| `montantAFacturerTtc` | Y (index 24) | **Montant à facturer TTC** — facture du mois |
| `sommesAReverserSrc` | Z (index 25) | Sommes à reverser (source) |

#### Champs enrichis (depuis Tableau de Bord — feuille "Soldes")

| Champ Java | Description |
|---|---|
| `nousDoit_Prec` | Solde précédent — montant que le client nous devait avant ce mois |

#### Champs calculés (par `TrfCalculator`)

| Champ Java | Formule |
|---|---|
| `nousDoit_Maintenant` | `montantAFacturerTtc + nousDoit_Prec` |
| `encaissementsParCompensation` | `min(sommesCzPhenix, max(0, nousDoit_Maintenant))` |
| `sommesAReverserFinal` | `max(0, sommesCzPhenix - nousDoit_Maintenant)` |
| `nousDoit_ApreFacturation` | `max(0, nousDoit_Maintenant - encaissementsParCompensation)` |
| `etatCompensations` | `"Comp VRT"`, `"Comp CB"`, `"NON COMP"`, ou description partielle |
| `virements` | `= sommesAReverserFinal` (montant à virer au client) |

---

## 5. Couche de persistance (Base de données SQLite)

### Localisation

```
~/.cabinet_phenix/data.db     (SQLite, créé automatiquement au 1er démarrage)
```

### Schéma de la base de données

#### Table `companies`

```sql
CREATE TABLE companies (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    name        TEXT    NOT NULL UNIQUE,   -- Nom du client (dossier)
    source_path TEXT,                      -- Chemin du fichier Excel source
    last_sync   TEXT                       -- Horodatage ISO de la dernière sync
);
```

#### Table `creance_rows`

```sql
CREATE TABLE creance_rows (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    company_id INTEGER NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
    row_index  INTEGER,      -- Index 0-based dans la feuille source
    col_a TEXT, col_b TEXT, col_c TEXT, col_d TEXT, col_e TEXT,
    col_f TEXT, col_g REAL, col_h REAL, col_i TEXT, col_j TEXT,
    col_k REAL, col_l TEXT, col_m TEXT, col_n TEXT, col_o REAL,
    col_p TEXT, col_q REAL, col_r REAL, col_s REAL, col_t REAL,
    col_u REAL, col_v REAL, col_w REAL, col_x REAL, col_y REAL
);
```

#### Table `trf_summaries`

```sql
CREATE TABLE trf_summaries (
    id                             INTEGER PRIMARY KEY AUTOINCREMENT,
    company_id                     INTEGER NOT NULL REFERENCES companies(id),
    client_code                    TEXT,
    iban                           TEXT,
    bic                            TEXT,
    non_compensation               INTEGER,  -- 0 ou 1 (booléen SQLite)
    creance_principale             REAL,
    recouvre_et_facture            REAL,
    penalites                      REAL,
    dont_en_attente                REAL,
    frais_procedure                REAL,
    recouvre_total                 REAL,
    deja_facture                   REAL,
    depuis_le_debut                REAL,
    commissions                    REAL,
    sommes_cz_phenix               REAL,
    montant_a_facturer_ttc         REAL,
    sommes_a_reverser_src          REAL,
    nous_doit_prec                 REAL,
    nous_doit_maintenant           REAL,
    encaissements_par_compensation REAL,
    sommes_a_reverser_final        REAL,
    nous_doit_apre_facturation     REAL,
    etat_compensations             TEXT,
    virements                      REAL,
    last_sync                      TEXT
);
```

### Classe `DatabaseManager`

- **Singleton** initialisé dans `App.start()` via `DatabaseManager.initialize()`
- Toutes les méthodes publiques sont `synchronized` pour la sécurité multi-thread
- Utilise `PRAGMA foreign_keys = ON` pour les cascades de suppression
- Connexion JDBC directe sur `jdbc:sqlite:~/.cabinet_phenix/data.db`

#### Méthodes principales

| Méthode | Description |
|---|---|
| `upsertCompany(name, sourcePath)` | INSERT OR UPDATE sur companies, retourne l'ID |
| `replaceCreanceRows(companyId, rows)` | DELETE puis INSERT batch dans une transaction |
| `replaceTrfSummary(companyId, cs)` | DELETE puis INSERT du résumé TRF |
| `getAllCompanies()` | SELECT avec COUNT des lignes, pour le dashboard |
| `getCreanceRows(companyId)` | SELECT toutes les lignes d'une société |
| `getTrfSummary(companyId)` | SELECT le résumé TRF d'une société |

---

## 6. Couche service — Pipeline de consolidation

### Vue d'ensemble du flux

```
rootFolder/
  ├── CLIENT_A/
  │   └── Etat des créances/
  │       └── etat_creances_CLIENT_A.xlsx
  ├── CLIENT_B/
  │   └── etat de creances/
  │       └── etat_2024.xlsx
  └── ...
        │
        ▼
  FolderScanner.scan(rootFolder)
        │
        ▼  List<CompanyFile>
  ExcelReader.readFiltered(companyName, excelFile)
        │     (filtre : colonne S non vide)
        ▼  List<CreanceRow>
  groupedRows : Map<String, List<CreanceRow>>
        │
        ▼
  ExcelWriter.write(groupedRows, outputFile)
        │
        ▼  etat_creances_global_YYYY-MM-DD_HH-mm-ss.xlsx
```

### `FolderScanner`

Parcourt `rootFolder`, trouve les sous-dossiers ayant un dossier "Etat des créances" (insensible aux accents et à la casse — cherche "etat" + "cr") contenant un fichier `.xlsx` ou `.xls` commençant par "etat".

- Tri alphabétique des sociétés pour ordre reproductible
- Si plusieurs fichiers correspondent → prend le plus récemment modifié

### `ExcelReader`

Lit la première feuille du fichier Excel d'un client :
- Ligne 0 = en-tête (ignorée)
- Lignes 1+ : incluses si colonne S (`AppConfig.FILTER_COLUMN_INDEX = 18`) contient une valeur réelle
- Une "valeur réelle" : texte non vide, nombre ≠ 0, ou booléen
- Les formules sont évaluées via `FormulaEvaluator`
- Colonnes numériques spéciales (indices 8,9,12,13,14,18,20,21,22,23,24) : chaînes parsées en double

### `ExcelWriter`

Écrit le fichier de consolidation global :
- Une **ligne d'en-tête** (fond bleu clair `#BDD7EE`) avec le nom de la société en colonne B
- Une **ligne vide** de séparation
- Les **lignes de données** : nom société en colonne A, données source en colonnes B+
- 34 colonnes au total (A–AH)
- Auto-taille des colonnes, figement de la première ligne

---

## 7. Couche service — Génération TRF

### Les 3 fichiers d'entrée

| Fichier | Feuille lue | Données extraites |
|---|---|---|
| `ConsolidationGenerale.xlsx` | "Consolidation" (ou "Créances" ou Sheet 0) | Toutes les lignes par client (26 colonnes A–Z) |
| `LISTING_CABINET_PHENIX_pour_ZEKI.xls` | "Feuil1" | Code, IBAN, BIC, NonComp par client (colonnes C,D,U,V,W) |
| `Tableau_de_bord_facturation.xlsx` | "Soldes" (index 2) | Solde précédent par client (colonne K, index 10) |

### `DataReader` — Lecture des fichiers sources

#### `readAllConsolidationRows(File)`

Lit toutes les lignes de la feuille "Consolidation" (y compris la ligne d'en-tête). Colonnes numériques (indices 1,7,8,11,13,14,15,16,17,18,19,20,21,22,23,24,25) : si une cellule string ressemble à un nombre français, elle est convertie en `Double`.

**Matching de feuille** (par ordre de priorité) :
1. Feuille nommée "Consolidation"
2. Feuille nommée "Créances"
3. Première feuille disponible

#### `readClientInfoMap(File, Consumer<String>)`

Lit la feuille "Feuil1" du Listing. Mapping normalisé `normalize(name)` → `ClientInfo`. Le debug Consumer (optionnel) dump les 3 premières lignes dans le journal UI pour diagnostic.

#### `readPreviousBalances(File)`

Lit la feuille "Soldes" du Tableau de Bord. Colonne A = nom client, colonne K (index 10) = montant "Nous Doit".

#### `normalize(String)` — Clé de fuzzy matching

```java
// NFD décompose les caractères accentués en lettre + combinant
// La regex supprime tous les combinants (catégorie Unicode \p{M})
return Normalizer.normalize(s.trim(), Normalizer.Form.NFD)
    .replaceAll("\\p{M}", "")
    .toLowerCase()
    .replaceAll("\\s+", " ");
```

Exemples : `"SOCIÉTÉ GÉNÉRALE"` → `"societe generale"`, `"Phénix"` → `"phenix"`

#### `findClientInfo(String, Map)` — Matching client

1. Lookup exact sur la clé normalisée
2. Fallback : matching par sous-chaîne (si la clé du listing est contenue dans le nom normalisé ou vice-versa)

### `TrfCalculator` — Logique de calcul

#### `buildClientSummaries()`

1. **Regroupement** : parcourt les `ConsolidationRow`, agrège par nom de client (colonne A), somme les 13 colonnes monétaires (indices 7,8,11,15,17,18,19,20,21,22,23,24,25)
2. **Filtrage** : ignore les clients avec `sommesCzPhenix < 0.005` ET `montantAFacturerTtc < 0.005` (pas d'activité ce mois)
3. **Enrichissement** : cherche le `ClientInfo` correspondant via `DataReader.findClientInfo()`. Si non trouvé → client exclu
4. **Calcul** : appelle `calculate(cs)` pour chaque `ClientSummary`
5. **Tri** : par code client alphabétique (les clients sans code sont placés en fin)

#### `calculate(ClientSummary)` — Règles de compensation

```
nous_doit_maintenant = montant_a_facturer_ttc + nous_doit_prec

Si NON COMP :
    encaissements_par_compensation = 0
    nous_doit_apre_facturation = nous_doit_maintenant   (facture toujours due)
    sommes_a_reverser_final = sommes_cz_phenix           (tout reversé)
    etat = "NON COMP"

Sinon :
    comp_appliquee = min(enc, max(0, nous_doit_maintenant))
    reverser_final = max(0, enc - nous_doit_maintenant)
    nous_doit_apre = max(0, nous_doit_maintenant - comp_appliquee)
    etat = "Comp VRT" | "Comp CB" | "Comp partielle de X, reste nous devoir Y" | ""
```

`determineEtat()` : 
- Si `reverserFinal > 0.005` → `"Comp VRT"` (virement) ou `"Comp CB"` (chèque)
- Si compensation partielle (comp > 0 et reste_doit > 0) → description textuelle
- Si compensation totale → `"Comp VRT"` ou `"Comp CB"`
- Sinon → chaîne vide

### `TrfSheetWriter` — Les 4 feuilles du classeur TRF

#### Feuille 1 : "Consolidation"

- **En-tête** : 26 colonnes fixes `CONSO_HEADERS[]` (style bleu foncé)
- **Données** : toutes les `ConsolidationRow` source avec leur style selon le type (monétaire = `#,##0.00`, date = `dd/MM/yyyy`)
- **Sous-totaux** : à chaque changement de client (`col A`), insertion d'une ligne `"Total [NomClient]"` avec `SUBTOTAL(9, range)` sur les colonnes monétaires (`MONEY_COLS = {7,8,11,15,17,18,19,20,21,22,23,24,25}`)
- Colonnes monétaires = indices : 7,8,11,15,17,18,19,20,21,22,23,24,25

#### Feuille 2 : "Feuil1"

- Un résumé par client sur la structure des 26 colonnes
- Colonne 0 (A) = nom du client, colonne 1 (B) = code client
- Colonnes monétaires mappées sur leurs indices de Consolidation
- Ligne finale `"TOTAUX"` avec `SUM()` sur toutes les colonnes monétaires

#### Feuille 3 : "TRF" — Document principal de transfert

**En-tête de colonne (12 colonnes A–L) :**
```
A : CLIENTS EN FACTURATION MM/YY    F : SOMMES A REVERSER AU FINAL
B : ENCAISSEMENTS CZ PHENIX          G : ENCAISSEMENTS PAR COMPENSATION
C : MONTANT A FACTURER TTC           H : NOUS DOIT APRES FACTURATION
D : NOUS DOIT précédemment           I : ETAT DE COMPENSATIONS
E : NOUS DOIT MAINTENANT             J : VIREMENTS
                                     K : CHEQUES
                                     L : CODE CLIENT
```

**Formules Excel par ligne client :**
- `E = C + D` (nous doit maintenant = facturer + précédent)
- Si NON COMP : `F = B`, `G = 0`, `H = E - G`
- Sinon : `F = IF(B=0,0,IF(B<E,0,B-E))`, `G = IF(B=0,0,IF(B>E,E,B))`, `H = E - G`

**Sections du bas (après TOTAUX) :**
1. `VIREMENTS CLIENTS` — clients avec `sommesAReverserFinal > 0` (colonnes CLIENT, IBAN, MONTANT)
2. `VIREMENTS MANUELLES` — clients sans IBAN mais avec reversement
3. `NON COMP` — clients en mode non-compensation
4. `COMP PARTIELLE` — clients avec compensation partielle appliquée
5. `DEBITEURS` — clients sans encaissements mais encore débiteurs

#### Feuille 4 : "Feuil3"

Onglet vide requis par le format de référence.

---

## 8. Couche service — États Publics

### `EtatPublicGenerator`

#### Flux de traitement

```
rootFolder/
  └── CLIENT_X/
      └── Espace partagé [ou variante] /
          └── Etat des créances [ou variante] /
              └── L_ETAT_DE_CREANCES_CLIENT_X.xlsx  ← généré
              └── L_ETAT_DE_CREANCES_CLIENT_X.pdf   ← généré
```

#### Algorithme `generate()`

1. `FolderScanner.scan()` trouve tous les fichiers `Etat des créances` clients
2. Pour chaque client :
   a. **Résolution du dossier destination** (3 niveaux de fallback) :
      - Cherche "Espace partagé" → cherche "Etat des créances" dedans
      - Cherche directement "Etat des créances" dans le dossier client
      - Crée "Etat des créances" si rien trouvé
   b. Supprime les anciens fichiers `L_ETAT_*`
   c. Génère `.xlsx` + `.pdf`

#### Lecture de la feuille "Créances" source

- Lignes de métadonnées : `(3,7)=société`, `(4,7)=adresse1`, `(5,7)=adresse2`, `(7,7)=contact`, `(12,0)=codeClient`
- Données : lignes 16+ jusqu'à colonne 0 (NBRE) vide

**Mapping de colonnes source → sortie (11 colonnes de sortie) :**

| Colonne sortie | Index sortie | Colonne source (index) |
|---|---|---|
| NOMBRE | 0 | col 0 |
| V/REF | 1 | col 1 |
| REMIS LE | 2 | col 2 |
| ANCIENNETE | 3 | col 3 |
| N/REF | 4 | col 5 |
| DEBITEUR | 5 | col 6 |
| CREANCE PRINCIPALE | 6 | col 7 |
| RECOUVRE | 7 | col 8 |
| DONT EN ATTENTE DE FACTURATION | 8 | col 17 |
| ETAT | 9 | col 9 |
| CLOTURE | 10 | col 10 |

#### Sortie Excel (`.xlsx`)

- 8 lignes d'en-tête client (société, adresse, contact, code)
- En-tête tableau (fond bleu foncé `#1F4E79`, texte blanc)
- Lignes de données avec bordures grises
- Ligne `TOTAUX` avec formules `SUM()` sur CREANCE PRINCIPALE, RECOUVRE, DONT EN ATTENTE
- Figement à la ligne 9 (après les métadonnées client)

#### Sortie PDF (iText7, A4 paysage)

- Marges 20pt de chaque côté
- En-tête client en texte
- Tableau : largeurs proportionnelles `[38,58,58,58,58,130,78,78,92,58,58]` points
- En-tête tableau : fond bleu foncé `#1F4E79`, texte blanc gras
- Données : alternance blanc/gris clair `#F2F2F2`
- Ligne TOTAUX : fond jaune `#FFF2CC`, texte gras
- Calcul des totaux côté Java (pas de formules Excel)

---

## 9. Couche service — Comparaison PROCREANCES

### `ProcreancesComparator`

#### Colonnes PROCREANCES (première feuille)

| Constante | Index | Description |
|---|---|---|
| `PC_CODE` | 1 | N° Client |
| `PC_NOM` | 2 | Nom du client |
| `PC_HONO` | 5 | Honoraires TTC |
| `PC_DISPO` | 6 | Disponible |
| `PC_REV` | 8 | Reversement |

#### Colonnes ConsolidationGénérale

| Constante | Index | Contexte |
|---|---|---|
| `CS_NAME` | 0 | Nom client (col A) |
| `CS_CODE_COL` | 1 | Code (col B) |
| `CS_COMM_FEUIL1` | 2 | Commissions TTC (Feuil1, col C) |
| `CS_COMM_CONSO` | 21 | Commissions (Consolidation, col V) |
| `CS_CZ` | 23 | Sommes CZ Phénix (col X) |
| `CS_SOMMES_REV` | 25 | Sommes à reverser (col Z) |

#### Algorithme de comparaison

1. Lecture PROCREANCES → `Map<normalizedName, double[3]>` (hono, dispo, rev) + `Map<normalizedName, String[2]>` (name, code)
2. Lecture ConsolidationGénérale → même structure (préférence Feuil1 si disponible)
3. **Matching** : lookup exact sur nom normalisé → fallback sous-chaîne
4. **Calcul des écarts** : `diff = round2(proc) - round2(conso)`, seuil de tolérance 0,05€
5. **Résultat** : `ComparisonResult(allRows, discrepancies, unmatchedProc, unmatchedConso)`

#### Rapport Excel de sortie

Fichier `comparison_PROCREANCES_vs_CONSO_YYYY-MM-DD_HH-mm.xlsx` à 3 feuilles :

**Feuille "Récapitulatif"** : tous les clients appariés (11 colonnes : CLIENT, N°CLIENT, PROC Hono, CONSO Comm, DIFF, PROC Dispo, CONSO CZ, DIFF, PROC Rev, CONSO Rev, DIFF)

**Feuille "Écarts"** : uniquement les clients avec `|diff| > 0,05€` (mêmes 11 colonnes + ligne TOTAUX)

**Feuille "Non appariés"** : 2 tableaux
- Clients dans PROCREANCES mais absents de Conso
- Clients dans Conso mais absents de PROCREANCES

**Style des écarts** :
- Diff > 0 → vert `#C6EFCE` (texte `#276221`) 
- Diff < 0 → rouge `#FFC7CE` (texte `#9C0006`)
- Diff ≈ 0 → style monétaire standard

---

## 10. Couche service — Nouveaux services refactorisés

### 10.1 `service/excel/` — Couche Excel réutilisable

#### `ExcelStyleFactory` (Patron Fabrique)

Produit des styles Apache POI cohérents pour tous les classeurs. Évite la limite de 64000 styles Excel en centralisant la création.

| Méthode | Style produit |
|---|---|
| `getHeaderStyle(wb)` | Bleu foncé `#1F4E79`, texte blanc gras, bordures grises fines |
| `getDataStyle(wb)` | Transparent, bordures grises fines, centrage vertical |
| `getCurrencyStyle(wb)` | DataStyle + format `#,##0.00`, alignement droit |
| `getTotalStyle(wb)` | Jaune clair `#FFF2CC`, texte gras, bordures grises |
| `getTotalMoneyStyle(wb)` | TotalStyle + format `#,##0.00` |
| `getDateStyle(wb)` | DataStyle + format `dd/MM/yyyy` |

#### `ExcelSheetBuilder` (Patron Constructeur)

API fluide pour créer des feuilles tabulaires simples.

```java
// Exemple d'utilisation
Workbook wb = new ExcelSheetBuilder("Rapport")
    .withDefaultColumnWidth(20)
    .withFrozenPane(1, 0)
    .withAutoFilter(true)
    .addHeaderRow(List.of("Client", "Montant", "Date"))
    .addDataRow(List.of("ACME Corp", 1500.00, "01/05/2026"))
    .addDataRows(otherRows)
    .build();
```

#### `ExcelFormatterService` (Service sans état)

Utilitaires de lecture/écriture de cellules. Extrait le code dupliqué entre `TrfSheetWriter`, `EtatPublicGenerator` et `ProcreancesComparator`.

- `writeValue(cell, val, defaultStyle, dateStyle)` : écrit `Double`, `Number`, `Boolean`, `LocalDateTime`, `String` (avec parsing de nombres français)
- `formatForDisplay(val)` : formatage texte pour PDF et journaux
- `readString(row, col, fmt, eval)` : lecture chaîne null-safe
- `readDouble(row, col, fmt, eval)` : lecture numérique avec fallback parsing français
- `columnLetter(int)` : conversion index colonne → lettre Excel (`0→"A"`, `25→"Z"`, `26→"AA"`)

---

### 10.2 `service/data/` — Extraction et normalisation

#### `DataExtractor` — Lecture bas niveau Apache POI

Centralise les patterns `cellStr`/`cellDouble` répétés dans `DataReader`, `ProcreancesComparator` et `EtatPublicGenerator`.

- `extractString(row, col, evaluator)` : chaîne trimée, `""` si null
- `extractDouble(row, col, evaluator)` : double avec fallback parsing, `0.0` si null
- `extractStrings(row, evaluator, int... columns)` : lecture multi-colonnes en un appel
- `extractSheetCell(sheet, rowIndex, colIndex, evaluator)` : lecture par position de feuille

#### `DataNormalizer` — Normalisation

- `normalize(String)` : NFD + suppression diacritiques + toLowerCase + collapse espaces → clé de fuzzy matching
- `normalizeAmount(double)` : arrondi 2 décimales
- `sanitizeFileName(String)` : remplace `\/:*?"<>|` par `_`
- `fuzzyMatch(a, b)` : `equals || contains (dans les 2 sens)` sur noms normalisés

#### `DataConverter` — Conversion modèles

- `toCreanceRow(companyName, rowMap, headers, rowIndex)` : `Map<String,Object>` → `CreanceRow`
- `toCreanceRows(companyName, rowMaps, headers, startIndex)` : conversion batch

---

### 10.3 `service/io/DataIOFactory` (Patron Fabrique)

Mappe les extensions de fichiers vers des `FileReader`/`FileWriter` (interfaces internes).

```java
// Enregistrement par défaut
readers.put("xlsx", excelReader);   // → Map<"headers",List<String>> + Map<"rows",List<List<Object>>>
readers.put("xls",  excelReader);
writers.put("xlsx", excelWriter);   // → ExcelReportStrategy.generate()
writers.put("xls",  excelWriter);

// Usage
FileReader reader = factory.getReaderByFile(new File("data.xlsx"));
Map<String,Object> data = reader.read(file);
```

---

### 10.4 `service/util/` — Patron Observateur

#### `ProgressObserver` (interface)
```java
interface ProgressObserver {
    void onProgressUpdate(double progress, String message);  // 0.0..1.0
    void onCompleted();
    void onFailed(Exception exception);
}
```

#### `ProgressNotifier` (Observable)
- Maintient une `List<ProgressObserver>` (copie défensive pour thread safety)
- `subscribe(observer)` / `unsubscribe(observer)`
- `notifyProgress(prog, msg)` / `notifyCompleted()` / `notifyFailed(ex)`
- `asBiConsumer()` : adaptateur vers `BiConsumer<Double,String>` pour la compatibilité avec les services legacy

---

### 10.5 `service/report/` — Patron Stratégie

```java
// Interface
interface ReportStrategy {
    File generate(Map<String,Object> data, File outputPath) throws Exception;
    String getFormat();  // "XLSX", "PDF", etc.
}

// Factory
ReportStrategyFactory factory = new ReportStrategyFactory();
ReportStrategy strategy = factory.getStrategy("XLSX");
File result = strategy.generate(data, outputPath);
```

- `ExcelReportStrategy` : utilise `ExcelSheetBuilder` → écrit `.xlsx`
- `PdfReportStrategy` : stub (non implémenté — lance `UnsupportedOperationException`)
- `ReportStrategyFactory` : registre `HashMap<String, ReportStrategy>`, `register()` public pour extensions

---

## 11. Couche contrôleur (JavaFX)

### `MainController`

Contrôleur FXML principal chargé depuis `main.fxml`.

#### Champs FXML injectés

| Champ | Type | Rôle |
|---|---|---|
| `badgesBox` | `HBox` | Conteneur des badges d'état des fichiers |
| `missingFilesLabel` | `Label` | Avertissement "N fichier(s) manquant(s)" |
| `actionsGrid` | `GridPane` | Grille des 5 boutons d'action |
| `progressBar` | `ProgressBar` | Barre de progression (0.0..1.0) |
| `logArea` | `TextArea` | Journal des opérations (horodaté HH:mm:ss) |
| `statusBar` | `HBox` | Barre d'état (masquée quand inactif) |
| `statusLabel` | `Label` | Message de résultat |
| `openFileBtn` | `Button` | "Ouvrir" le fichier de sortie |
| `dashboardController` | `DashboardController` | Contrôleur imbriqué (rafraîchi après opération) |

#### Services instanciés

```java
MergeService          mergeService          = new MergeService(DatabaseManager.getInstance());
EspacePartageFixer    espacePartageFixer    = new EspacePartageFixer();
EtatPublicGenerator   etatPublicGenerator   = new EtatPublicGenerator();
TrfGeneratorService   trfGeneratorService   = new TrfGeneratorService(DatabaseManager.getInstance());
ProcreancesComparator procreancesComparator = new ProcreancesComparator();
```

#### Threading — Pattern standard pour chaque action

```java
setAllButtonsDisabled(true);
progressBar.setProgress(0);
logArea.clear();

executor.submit(() -> {
    try {
        File result = service.doWork(files,
            (prog, msg) -> Platform.runLater(() -> {
                progressBar.setProgress(prog);
                appendLog(msg);
            }));
        Platform.runLater(() -> {
            // Mise à jour UI après succès
            setAllButtonsDisabled(false);
        });
    } catch (Exception e) {
        Platform.runLater(() -> {
            appendLog("FATAL: " + e.getMessage());
            setAllButtonsDisabled(false);
        });
    }
});
```

#### Badges de statut

Chaque badge indique si le fichier/dossier configuré existe :
- ✓ vert (`badge-ok`) : configuré et accessible
- ✗ rouge (`badge-missing`) : manquant ou non configuré

#### Dialogue de configuration

`openFileConfig()` crée une fenêtre modale (`Modality.APPLICATION_MODAL`) avec 6 lignes :

| Label | Type | Clé Préférence |
|---|---|---|
| Dossier source | Répertoire | `merge_root_folder` |
| Dossier de sortie | Répertoire | `output_folder` |
| ConsolidationGénérale | Fichier `.xlsx` | `trf_consolidation_file` |
| Listing Cabinet Phénix | Fichier `.xlsx` | `trf_listing_file` |
| Tableau de Bord | Fichier `.xlsx` | `trf_tableau_file` |
| Export PROCREANCES | Fichier `.xls` | `procreancesPath` |

---

### `ApplicationConfig` (Singleton DI)

Assemble les dépendances au démarrage de l'application. Évite l'injection de dépendances manuelle dans `MainController`.

```java
ApplicationConfig cfg = ApplicationConfig.getInstance();
cfg.getDatabaseManager();        // DatabaseManager.getInstance()
cfg.getReportStrategyFactory();  // new ReportStrategyFactory()
cfg.getDataIOFactory();          // new DataIOFactory()
cfg.getReportCommandFactory();   // new ReportCommandFactory(dbManager)
```

### `ReportCommandFactory` + commandes

```java
// Actuellement enregistrée :
commands.put("GENERATE_TRF", new GenerateTrfCommand(dbManager));

// Prévues (à implémenter) :
// new GenerateEtatPublicCommand()
// new CompareFilesCommand()
// new FixPathsCommand()
// new RunConsolidationCommand(dbManager)
```

`GenerateTrfCommand.execute(context, notifier)` :
1. Extrait `consoFile`, `listingFile`, `tableauFile`, `outputFolder` du `context`
2. Valide existence (lève `BusinessException(FILE_NOT_FOUND)` si absent)
3. Délègue à `TrfGeneratorService.generate()` avec `notifier.asBiConsumer()`

---

## 12. Patrons de conception appliqués

| Patron | Classe(s) | Problème résolu |
|---|---|---|
| **Singleton** | `DatabaseManager`, `ApplicationConfig` | Ressource partagée unique — pas de double initialisation |
| **Builder** | `ExcelSheetBuilder` | Construction fluide d'objets Excel complexes |
| **Factory** | `ExcelStyleFactory`, `DataIOFactory`, `ReportStrategyFactory`, `ReportCommandFactory` | Création découplée d'objets selon le type |
| **Strategy** | `ReportStrategy`, `ExcelReportStrategy`, `PdfReportStrategy` | Algorithme de génération interchangeable au runtime |
| **Command** | `ReportCommand`, `GenerateTrfCommand` | Encapsulation des actions UI — extensible sans modifier le contrôleur |
| **Observer** | `ProgressObserver`, `ProgressNotifier` | Découplage entre services et UI — les services ne dépendent pas de JavaFX |
| **Facade** | `TrfGeneratorService`, `ApplicationConfig` | Interface simplifiée sur des sous-systèmes complexes |
| **Template Method** | Flux dans `MergeService`, `TrfGeneratorService` | Squelette d'algorithme fixe avec étapes interchangeables |

---

## 13. Interface utilisateur (JavaFX/FXML)

### Fichiers de ressources

```
src/main/resources/com/zeki/merger/
    main.fxml        — fenêtre principale
    dashboard.fxml   — vue tableau de bord (incluse dans main.fxml)
    styles.css       — feuille de style JavaFX
```

### Styles CSS principaux

| Classe CSS | Usage |
|---|---|
| `.secondary-btn` | Boutons d'action secondaires (TRF, États, Comparer, Corriger) |
| `.run-btn` | Bouton principal "▶ CONSOLIDER" |
| `.badge-ok` | Badge vert ✓ |
| `.badge-missing` | Badge rouge ✗ |
| `.action-btn-name` | Label nom de l'action dans le bouton |
| `.action-btn-desc` | Label description de l'action dans le bouton |

### Dimensions

Fenêtre : 920×680 px, minimum 720×540 px.

---

## 14. Configuration et préférences

### `AppConfig` — Constantes globales

```java
TARGET_SUBFOLDER         = "etat de creances"          // Dossier cible cherché
FILE_PREFIX              = "etat"                      // Préfixe du fichier Excel
FILTER_COLUMN_INDEX      = 18                          // Colonne S (0-based)
FILTER_COLUMN_LABEL      = "S"
SOCIETE_COLUMN_HEADER    = "Société"
DEFAULT_ROOT_PATH        = "/Users/zekimertinceoglu/Dropbox/ZEKI IT"
DEFAULT_OUTPUT_PATH      = "/Users/zekimertinceoglu/Dropbox/ZEKI IT"
OUTPUT_FILENAME          = "etat_creances_global.xlsx"
TRF_OUTPUT_FILENAME      = "trf_export.xlsx"
CREANCES_SHEET_NAME      = "Créances"
ETAT_PUBLIC_FILENAME_PREFIX = "L_ETAT_DE_CREANCES_"
ESPACE_PARTAGE_FILENAME  = "CorrespondanceClient-EspacePartage.xlsx"
ETAT_CREANCES_SUFFIX     = "\\Etat des créances"
FIX_OVERWRITE            = true                        // Écraser en place
```

### `AppPreferences` — Préférences utilisateur

Stockées via `java.util.prefs.Preferences.userNodeForPackage(AppPreferences.class)`.

```
Clé                    Valeur
merge_root_folder      Chemin du dossier racine des clients
output_folder          Chemin du dossier de sortie
trf_consolidation_file Chemin de ConsolidationGenerale.xlsx
trf_listing_file       Chemin du Listing Cabinet Phénix
trf_tableau_file       Chemin du Tableau de Bord
procreancesPath        Chemin de l'export PROCREANCES (.xls)
consoComparePath       (Réservé — non encore utilisé dans l'UI)
```

---

## 15. Dépendances Maven

```xml
<groupId>com.zeki</groupId>
<artifactId>etat-creances-merger</artifactId>
<version>1.0.0</version>
```

| Dépendance | Version | Usage |
|---|---|---|
| `javafx-controls` | 21.0.4 | Composants UI JavaFX |
| `javafx-fxml` | 21.0.4 | Chargement FXML |
| `poi-ooxml` | 5.2.5 | Lecture/écriture Excel `.xlsx` et `.xls` |
| `slf4j-simple` | 2.0.13 | Logging Apache POI |
| `sqlite-jdbc` | 3.45.3.0 | Base de données embarquée SQLite |
| `itext7-core` | 7.2.5 | Génération PDF (A4 paysage, tableaux) |

**Plugins Maven :**
- `maven-compiler-plugin 3.13.0` — Java 17+, `--enable-preview`
- `javafx-maven-plugin 0.0.8` — `mvn javafx:run`
- `maven-shade-plugin 3.5.1` — JAR fat avec `Launcher` comme Main-Class
- `jpackage-maven-plugin 1.6.5` — Packaging natif macOS/Windows

---

## 16. Construction et exécution

### Prérequis

- Java 17+ (JDK)
- Maven 3.8+
- JavaFX 21 (inclus via Maven)

### Commandes Maven

```bash
# Compilation uniquement
mvn clean compile

# Compilation + packaging (JAR fat)
mvn clean package -DskipTests

# Exécution directe (mode développement)
mvn javafx:run

# Tests unitaires
mvn test

# Test spécifique
mvn test -Dtest=NomDuTest

# Génération JAR exécutable
mvn clean package -DskipTests
# → target/etat-creances-merger-1.0.0.jar
```

### Exécution du JAR

```bash
java -jar target/etat-creances-merger-1.0.0.jar
```

### Structure des fichiers générés

```
~/.cabinet_phenix/
    data.db                          ← Base de données SQLite

[outputFolder]/
    etat_creances_global_YYYY-MM-DD_HH-mm-ss.xlsx   ← Consolidation
    TRF_MM_YYYY.xlsx                                  ← Classeur TRF
    comparison_PROCREANCES_vs_CONSO_YYYY-MM-DD_HH-mm.xlsx ← Rapport comparaison

[rootFolder]/
  CLIENT_X/
    Espace partagé .../
      Etat des créances/
        L_ETAT_DE_CREANCES_CLIENT_X.xlsx             ← Etat public
        L_ETAT_DE_CREANCES_CLIENT_X.pdf              ← Etat public PDF
```

---

## 17. Flux de données complets (Diagrammes)

### Flux 1 — Consolidation (bouton CONSOLIDER)

```
Utilisateur clique "CONSOLIDER"
    │
    ├─ Validation : rootFolder existe ? outputFolder existe ?
    │
    ├─ ExecutorService.submit(task)
    │   │
    │   ├─ FolderScanner.scan(rootFolder)
    │   │   └─ returns List<CompanyFile>
    │   │
    │   ├─ Pour chaque CompanyFile :
    │   │   ├─ ExcelReader.readFiltered(companyName, excelFile)
    │   │   │   └─ returns List<CreanceRow> (col S non vide)
    │   │   ├─ groupedRows.put(company, rows)
    │   │   └─ DatabaseManager.upsertCompany() + replaceCreanceRows()
    │   │
    │   └─ ExcelWriter.write(groupedRows, outputFile)
    │       └─ etat_creances_global_[timestamp].xlsx
    │
    └─ Platform.runLater() → UI update (statusLabel, openFileBtn)
```

### Flux 2 — Génération TRF

```
Utilisateur clique "Générer TRF"
    │
    ├─ Validation : 3 fichiers + outputFolder existent ?
    │
    ├─ ExecutorService.submit(task)
    │   │
    │   ├─ DataReader.readAllConsolidationRows(consoFile)
    │   │   └─ List<ConsolidationRow>
    │   │
    │   ├─ DataReader.readClientInfoMap(listingFile, debugCallback)
    │   │   └─ Map<normalizedName, ClientInfo>
    │   │
    │   ├─ DataReader.readPreviousBalances(tableauFile)
    │   │   └─ Map<normalizedName, Double>
    │   │
    │   ├─ TrfCalculator.buildClientSummaries(allRows, clientInfoMap, balanceMap)
    │   │   ├─ Regroupement par col A, sommes des 13 colonnes monétaires
    │   │   ├─ Filtrage (pas d'activité → skip)
    │   │   ├─ Enrichissement ClientInfo (IBAN, NonComp)
    │   │   ├─ calculate() → tous les champs TRF
    │   │   └─ Tri par code client
    │   │
    │   ├─ DatabaseManager.replaceTrfSummary() pour chaque client
    │   │
    │   └─ TrfSheetWriter.write(allRows, summaries, outputFile)
    │       ├─ Feuille "Consolidation" avec sous-totaux
    │       ├─ Feuille "Feuil1" avec résumés par client
    │       ├─ Feuille "TRF" avec formules et sections virements
    │       └─ Feuille "Feuil3" (vide)
    │
    └─ Platform.runLater() → DashboardController.refresh()
```

### Flux 3 — États Publics

```
Utilisateur clique "États Publics"
    │
    ├─ FolderScanner.scan(rootFolder)
    │
    ├─ Pour chaque CompanyFile :
    │   ├─ Résolution destDir (3 niveaux de fallback)
    │   ├─ Suppression anciens L_ETAT_*
    │   ├─ generateForClient()
    │   │   ├─ Ouverture workbook source
    │   │   ├─ Extraction métadonnées (société, adresse, contact, code)
    │   │   ├─ Extraction lignes de données (16+, stop si col 0 vide)
    │   │   ├─ writeOutput() → L_ETAT_DE_CREANCES_[client].xlsx
    │   │   └─ writePdf() → L_ETAT_DE_CREANCES_[client].pdf
    │   └─ progress.accept(prog, message)
    │
    └─ Done
```

---

## 18. Glossaire métier

| Terme | Définition |
|---|---|
| **TRF** | Transfert et Reversement Financier — document récapitulatif mensuel des mouvements financiers |
| **Créance** | Somme d'argent due à un créancier (client de Phénix) par un débiteur |
| **Recouvrement** | Processus de collecte des créances impayées |
| **Encaissement** | Réception effective d'un paiement d'un débiteur |
| **Compensation** | Mécanisme par lequel les encaissements collectés sont déduits du montant facturé avant reversement |
| **NON COMP** | Client en mode "Non Compensation" — les encaissements ne peuvent pas être compensés avec la facture (réglementaire) |
| **Reversement** | Montant que Phénix doit rendre au client après compensation |
| **Nous Doit** | Montant que le client doit à Phénix (facture non encore payée) |
| **IBAN** | Numéro de compte bancaire international — pour les virements automatiques |
| **Comp VRT** | Compensation par virement bancaire |
| **Comp CB** | Compensation par chèque bancaire |
| **Comp Partielle** | Compensation partielle : les encaissements couvrent seulement une partie de la facture |
| **ConsolidationGénérale** | Fichier Excel central agrégeant les données de tous les clients Phénix |
| **PROCREANCES** | Système externe de gestion de créances — l'export permet la comparaison avec Phénix |
| **Espace Partagé** | Dossier partagé avec chaque client, dans lequel sont déposés les états publics |
| **Etat Public** | Document mensuel remis au client résumant l'état de ses créances |
| **Listing Cabinet Phénix** | Fichier Excel interne listant tous les clients avec leurs données bancaires |
| **Tableau de Bord** | Fichier Excel interne avec les soldes précédents par client |

---

## 19. Points d'attention et limitations connues

### Limitations techniques

1. **Limite de 64000 styles Excel** : Apache POI impose un maximum de styles par classeur. Les classes `TrfSheetWriter` et `EtatPublicGenerator` créent leurs styles dans des objets `Styles` internes qui n'utilisent pas encore `ExcelStyleFactory`. Un refactoring supplémentaire serait nécessaire pour éliminer cette duplication.

2. **PDF non implémenté pour la stratégie** : `PdfReportStrategy` dans `service/report/` lève `UnsupportedOperationException`. La génération PDF fonctionne uniquement via `EtatPublicGenerator.writePdf()` (code propriétaire iText7).

3. **Commandes partiellement enregistrées** : `ReportCommandFactory` n'enregistre que `GenerateTrfCommand`. Les autres actions (États Publics, Comparer, Corriger, Consolider) sont encore directement invoquées dans `MainController` sans passer par le pattern Command.

4. **`DatabaseManager` sans pool de connexions** : Une seule connexion SQLite partagée, protégée par `synchronized`. Convenable pour un usage mono-utilisateur mais pas scalable.

5. **Parsing des nombres français** : `ConsolidationRow.parseFrenchDouble()` gère les formats courants mais pourrait échouer sur des formats inhabituels (ex : espaces insécables ` ` comme séparateur de milliers dans certains locales).

6. **Matching fuzzy client** : L'algorithme de fuzzy matching (sous-chaîne normalisée) peut produire des faux positifs si deux noms de clients sont des sous-chaînes l'un de l'autre.

### Bonnes pratiques à maintenir

- **Toujours utiliser `Platform.runLater()`** pour toute modification de composant JavaFX depuis un thread de fond
- **Ne jamais appeler `DatabaseManager.getInstance()`** avant `DatabaseManager.initialize()`
- **Les `BiConsumer<Double, String>`** (callbacks de progression) doivent être appelés depuis les threads de fond uniquement — c'est eux qui déclenchent le `Platform.runLater()` côté contrôleur
- **La normalisation des noms** (`DataReader.normalize()` ou `DataNormalizer.normalize()`) doit toujours être utilisée avant de comparer des noms de clients entre fichiers différents

### Évolutions planifiées

- Implémenter les commandes manquantes (`GenerateEtatPublicCommand`, `CompareFilesCommand`, `FixPathsCommand`, `RunConsolidationCommand`)
- Refactoriser `TrfSheetWriter` pour utiliser `ExcelStyleFactory` et `ExcelFormatterService`
- Refactoriser `EtatPublicGenerator` pour utiliser `ExcelStyleFactory`
- Écrire des tests unitaires pour `TrfCalculator`, `DataNormalizer`, `ConsolidationRow.parseFrenchDouble()`
- Intégrer SLF4J/Logback pour le logging structuré (remplacer les `System.out.println`)

---

*Document généré le 2026-05-12. Ce document reflète l'état du projet après refactoring selon le plan de `/files (1)/REFACTORING_PLAN.md`.*
