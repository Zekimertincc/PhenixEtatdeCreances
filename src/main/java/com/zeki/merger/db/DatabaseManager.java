package com.zeki.merger.db;

import com.zeki.merger.model.CreanceRow;
import com.zeki.merger.trf.model.ClientSummary;

import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.Collections;

/**
 * Singleton embedded SQLite database.
 * Call {@link #initialize()} once in App.start() before loading FXML.
 * Thereafter any code that needs the DB calls {@link #getInstance()}.
 *
 * All public methods are {@code synchronized} to guard against the background
 * merge thread writing concurrently with the FX thread reading for the dashboard.
 */
public class DatabaseManager {

    private static DatabaseManager instance;

    private final Connection conn;
    private static final DateTimeFormatter ISO = DateTimeFormatter.ISO_LOCAL_DATE_TIME;

    // -------------------------------------------------------------------------
    // Lifecycle
    // -------------------------------------------------------------------------

    public static void initialize() throws Exception {
        if (instance == null) {
            instance = new DatabaseManager();
        }
    }

    public static DatabaseManager getInstance() {
        return instance;
    }

    private DatabaseManager() throws Exception {
        Path dbPath = Path.of(System.getProperty("user.home"), ".cabinet_phenix", "data.db");
        Files.createDirectories(dbPath.getParent());
        conn = DriverManager.getConnection("jdbc:sqlite:" + dbPath);
        conn.createStatement().execute("PRAGMA foreign_keys = ON");
        createSchema();
    }

    private void createSchema() throws SQLException {
        try (Statement st = conn.createStatement()) {
            st.executeUpdate("""
                CREATE TABLE IF NOT EXISTS companies (
                    id          INTEGER PRIMARY KEY AUTOINCREMENT,
                    name        TEXT    NOT NULL UNIQUE,
                    source_path TEXT,
                    last_sync   TEXT
                )""");

            try {
                st.executeUpdate("""
                    DELETE FROM companies
                    WHERE id NOT IN (
                        SELECT MIN(id) FROM companies GROUP BY LOWER(TRIM(name))
                    )
                """);
            } catch (Exception ignored) {}
            try {
                st.executeUpdate("CREATE UNIQUE INDEX IF NOT EXISTS idx_companies_name_ci ON companies(name COLLATE NOCASE)");
            } catch (Exception ignored) {}

            st.executeUpdate("""
                CREATE TABLE IF NOT EXISTS creance_rows (
                    id         INTEGER PRIMARY KEY AUTOINCREMENT,
                    company_id INTEGER NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
                    row_index  INTEGER,
                    col_a TEXT, col_b TEXT, col_c TEXT, col_d TEXT, col_e TEXT,
                    col_f TEXT, col_g REAL, col_h REAL, col_i TEXT, col_j TEXT,
                    col_k REAL, col_l TEXT, col_m TEXT, col_n TEXT, col_o REAL,
                    col_p TEXT, col_q REAL, col_r REAL, col_s REAL, col_t REAL,
                    col_u REAL, col_v REAL, col_w REAL, col_x REAL, col_y REAL
                )""");

            st.executeUpdate("""
                CREATE TABLE IF NOT EXISTS trf_months (
                    id             INTEGER PRIMARY KEY AUTOINCREMENT,
                    year           INTEGER NOT NULL,
                    month          INTEGER NOT NULL,
                    status         TEXT    NOT NULL DEFAULT 'open',
                    nb_clients     INTEGER,
                    total_montant  REAL,
                    total_nous_doit REAL,
                    closed_at      TEXT,
                    created_at     TEXT DEFAULT (datetime('now')),
                    UNIQUE(year, month)
                )""");

            st.executeUpdate("""
                CREATE TABLE IF NOT EXISTS trf_history (
                    id               INTEGER PRIMARY KEY AUTOINCREMENT,
                    month_id         INTEGER NOT NULL REFERENCES trf_months(id),
                    client_name      TEXT NOT NULL,
                    client_code      TEXT,
                    encaissements    REAL,
                    montant_facturer REAL,
                    nous_doit_prec   REAL,
                    sommes_reverser  REAL,
                    etat             TEXT,
                    iban             TEXT,
                    non_compensation INTEGER DEFAULT 0
                )""");

            st.executeUpdate("""
                CREATE TABLE IF NOT EXISTS company_summaries (
                    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
                    company_id          INTEGER NOT NULL UNIQUE REFERENCES companies(id) ON DELETE CASCADE,
                    code_client         TEXT,
                    responsable         TEXT,
                    nb_dossiers         INTEGER,
                    nb_soldes           INTEGER,
                    nb_gestion          INTEGER,
                    nb_irr              INTEGER,
                    nb_arj              INTEGER,
                    nb_autres           INTEGER,
                    creance_principale  REAL,
                    recouvre_total      REAL,
                    commissions         REAL,
                    dernier_dossier     TEXT,
                    last_sync           TEXT
                )""");

            st.executeUpdate("""
                CREATE TABLE IF NOT EXISTS mail_templates (
                    id         INTEGER PRIMARY KEY AUTOINCREMENT,
                    name       TEXT NOT NULL UNIQUE COLLATE NOCASE,
                    body       TEXT NOT NULL,
                    created_at TEXT,
                    updated_at TEXT
                )""");

            st.executeUpdate("""
                CREATE TABLE IF NOT EXISTS trf_summaries (
                    id                             INTEGER PRIMARY KEY AUTOINCREMENT,
                    company_id                     INTEGER NOT NULL REFERENCES companies(id) ON DELETE CASCADE,
                    client_code                    TEXT,
                    iban                           TEXT,
                    bic                            TEXT,
                    non_compensation               INTEGER,
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
                )""");
        }
    }

    public void close() {
        try { if (conn != null && !conn.isClosed()) conn.close(); }
        catch (SQLException ignored) {}
    }

    // -------------------------------------------------------------------------
    // Write methods
    // -------------------------------------------------------------------------

    /** Upsert company by name and return its row id. */
    public synchronized long upsertCompany(String name, String sourcePath) throws SQLException {
        name = name.trim();
        String now = LocalDateTime.now().format(ISO);
        try (PreparedStatement ps = conn.prepareStatement("""
                INSERT INTO companies (name, source_path, last_sync) VALUES (?,?,?)
                ON CONFLICT(name) DO UPDATE
                  SET source_path = excluded.source_path,
                      last_sync   = excluded.last_sync
                """)) {
            ps.setString(1, name);
            ps.setString(2, sourcePath);
            ps.setString(3, now);
            ps.executeUpdate();
        }
        try (PreparedStatement ps = conn.prepareStatement(
                "SELECT id FROM companies WHERE name = ?")) {
            ps.setString(1, name);
            try (ResultSet rs = ps.executeQuery()) {
                return rs.next() ? rs.getLong(1) : -1L;
            }
        }
    }

    /** Delete all creance rows for companyId and insert the new list in a single transaction. */
    public synchronized void replaceCreanceRows(long companyId, List<CreanceRow> rows)
            throws SQLException {
        try (PreparedStatement del = conn.prepareStatement(
                "DELETE FROM creance_rows WHERE company_id = ?")) {
            del.setLong(1, companyId);
            del.executeUpdate();
        }
        if (rows.isEmpty()) return;

        String sql = """
                INSERT INTO creance_rows
                  (company_id, row_index,
                   col_a,col_b,col_c,col_d,col_e,col_f,
                   col_g,col_h,col_i,col_j,col_k,col_l,col_m,col_n,col_o,col_p,
                   col_q,col_r,col_s,col_t,col_u,col_v,col_w,col_x,col_y)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """;

        conn.setAutoCommit(false);
        try (PreparedStatement ps = conn.prepareStatement(sql)) {
            for (CreanceRow cr : rows) {
                List<Object> vals = cr.getCellValues();
                ps.setLong(1, companyId);
                ps.setInt(2, cr.getOriginalRowIndex());
                for (int i = 0; i < 25; i++) {
                    Object v = i < vals.size() ? vals.get(i) : null;
                    bindValue(ps, 3 + i, v);
                }
                ps.addBatch();
            }
            ps.executeBatch();
            conn.commit();
        } catch (SQLException e) {
            conn.rollback();
            throw e;
        } finally {
            conn.setAutoCommit(true);
        }
    }

    /** Delete existing TRF summary for companyId and insert the new one. */
    public synchronized void replaceTrfSummary(long companyId, ClientSummary cs)
            throws SQLException {
        try (PreparedStatement del = conn.prepareStatement(
                "DELETE FROM trf_summaries WHERE company_id = ?")) {
            del.setLong(1, companyId);
            del.executeUpdate();
        }
        try (PreparedStatement ps = conn.prepareStatement("""
                INSERT INTO trf_summaries
                  (company_id, client_code, iban, bic, non_compensation,
                   creance_principale, recouvre_et_facture, penalites, dont_en_attente,
                   frais_procedure, recouvre_total, deja_facture, depuis_le_debut, commissions,
                   sommes_cz_phenix, montant_a_facturer_ttc, sommes_a_reverser_src,
                   nous_doit_prec, nous_doit_maintenant, encaissements_par_compensation,
                   sommes_a_reverser_final, nous_doit_apre_facturation,
                   etat_compensations, virements, last_sync)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """)) {
            ps.setLong  (1,  companyId);
            ps.setString(2,  cs.getClientCode());
            ps.setString(3,  cs.getIban());
            ps.setString(4,  cs.getBic());
            ps.setInt   (5,  cs.isNonCompensation() ? 1 : 0);
            ps.setDouble(6,  cs.getCreancePrincipale());
            ps.setDouble(7,  cs.getRecouvreEtFacture());
            ps.setDouble(8,  cs.getPenalites());
            ps.setDouble(9,  cs.getDontEnAttente());
            ps.setDouble(10, cs.getFraisProcedure());
            ps.setDouble(11, cs.getRecouvreTotol());
            ps.setDouble(12, cs.getDejaFacture());
            ps.setDouble(13, cs.getDepuisLeDebut());
            ps.setDouble(14, cs.getCommissions());
            ps.setDouble(15, cs.getSommesCzPhenix());
            ps.setDouble(16, cs.getMontantAFacturerTtc());
            ps.setDouble(17, cs.getSommesAReverserSrc());
            ps.setDouble(18, cs.getNousDoit_Prec());
            ps.setDouble(19, cs.getNousDoit_Maintenant());
            ps.setDouble(20, cs.getEncaissementsParCompensation());
            ps.setDouble(21, cs.getSommesAReverserFinal());
            ps.setDouble(22, cs.getNousDoit_ApreFacturation());
            ps.setString(23, cs.getEtatCompensations());
            ps.setDouble(24, cs.getVirements());
            ps.setString(25, LocalDateTime.now().format(ISO));
            ps.executeUpdate();
        }
    }

    // -------------------------------------------------------------------------
    // Read methods (dashboard queries)
    // -------------------------------------------------------------------------

    public synchronized List<CompanyRecord> getAllCompanies() throws SQLException {
        String sql = """
                SELECT c.id, c.name, COUNT(cr.id) AS row_count, c.last_sync
                FROM companies c
                LEFT JOIN creance_rows cr ON cr.company_id = c.id
                GROUP BY c.id
                ORDER BY c.name COLLATE NOCASE
                """;
        List<CompanyRecord> out = new ArrayList<>();
        try (Statement st = conn.createStatement();
             ResultSet rs = st.executeQuery(sql)) {
            while (rs.next()) {
                out.add(new CompanyRecord(
                    rs.getLong("id"),
                    rs.getString("name"),
                    rs.getInt("row_count"),
                    rs.getString("last_sync")));
            }
        }
        return out;
    }

    public synchronized List<Map<String, Object>> getCreanceRows(long companyId)
            throws SQLException {
        List<Map<String, Object>> out = new ArrayList<>();
        try (PreparedStatement ps = conn.prepareStatement(
                "SELECT * FROM creance_rows WHERE company_id = ? ORDER BY row_index")) {
            ps.setLong(1, companyId);
            try (ResultSet rs = ps.executeQuery()) {
                ResultSetMetaData meta = rs.getMetaData();
                int cols = meta.getColumnCount();
                while (rs.next()) {
                    Map<String, Object> row = new LinkedHashMap<>();
                    for (int i = 1; i <= cols; i++) {
                        row.put(meta.getColumnName(i), rs.getObject(i));
                    }
                    out.add(row);
                }
            }
        }
        return out;
    }

    public synchronized Optional<Map<String, Object>> getTrfSummary(long companyId)
            throws SQLException {
        try (PreparedStatement ps = conn.prepareStatement(
                "SELECT * FROM trf_summaries WHERE company_id = ?")) {
            ps.setLong(1, companyId);
            try (ResultSet rs = ps.executeQuery()) {
                if (!rs.next()) return Optional.empty();
                ResultSetMetaData meta = rs.getMetaData();
                Map<String, Object> row = new LinkedHashMap<>();
                for (int i = 1; i <= meta.getColumnCount(); i++) {
                    row.put(meta.getColumnName(i), rs.getObject(i));
                }
                return Optional.of(row);
            }
        }
    }

    // -------------------------------------------------------------------------
    // trf_months / trf_history DAO
    // -------------------------------------------------------------------------

    public synchronized void insertOrUpdateTrfMonth(int year, int month, String status,
            int nbClients, double totalMontant, double totalNousDoit) throws SQLException {
        String now = LocalDateTime.now().format(ISO);
        try (PreparedStatement ps = conn.prepareStatement("""
                INSERT INTO trf_months (year, month, status, nb_clients, total_montant, total_nous_doit, closed_at)
                VALUES (?,?,?,?,?,?,?)
                ON CONFLICT(year,month) DO UPDATE
                  SET status=excluded.status, nb_clients=excluded.nb_clients,
                      total_montant=excluded.total_montant, total_nous_doit=excluded.total_nous_doit,
                      closed_at=excluded.closed_at
                """)) {
            ps.setInt(1, year);
            ps.setInt(2, month);
            ps.setString(3, status);
            ps.setInt(4, nbClients);
            ps.setDouble(5, totalMontant);
            ps.setDouble(6, totalNousDoit);
            ps.setString(7, "closed".equals(status) ? now : null);
            ps.executeUpdate();
        }
    }

    public synchronized long getTrfMonthId(int year, int month) throws SQLException {
        try (PreparedStatement ps = conn.prepareStatement(
                "SELECT id FROM trf_months WHERE year=? AND month=?")) {
            ps.setInt(1, year);
            ps.setInt(2, month);
            try (ResultSet rs = ps.executeQuery()) {
                return rs.next() ? rs.getLong(1) : -1L;
            }
        }
    }

    public synchronized void deleteTrfHistory(long monthId) throws SQLException {
        try (PreparedStatement ps = conn.prepareStatement(
                "DELETE FROM trf_history WHERE month_id=?")) {
            ps.setLong(1, monthId);
            ps.executeUpdate();
        }
    }

    public synchronized void insertTrfHistory(long monthId, ClientSummary cs) throws SQLException {
        try (PreparedStatement ps = conn.prepareStatement("""
                INSERT INTO trf_history
                  (month_id, client_name, client_code, encaissements, montant_facturer,
                   nous_doit_prec, sommes_reverser, etat, iban, non_compensation)
                VALUES (?,?,?,?,?,?,?,?,?,?)
                """)) {
            ps.setLong(1, monthId);
            ps.setString(2, cs.getClientName());
            ps.setString(3, cs.getClientCode());
            ps.setDouble(4, cs.getSommesCzPhenix());
            ps.setDouble(5, cs.getMontantAFacturerTtc());
            ps.setDouble(6, cs.getNousDoit_Prec());
            ps.setDouble(7, cs.getSommesAReverserFinal());
            ps.setString(8, cs.getEtatCompensations());
            ps.setString(9, cs.getIban());
            ps.setInt(10, cs.isNonCompensation() ? 1 : 0);
            ps.executeUpdate();
        }
    }

    public synchronized List<TrfMonthRecord> getAllTrfMonths() {
        List<TrfMonthRecord> out = new ArrayList<>();
        try (Statement st = conn.createStatement();
             ResultSet rs = st.executeQuery("""
                 SELECT id, year, month, status, nb_clients, total_montant, total_nous_doit, closed_at
                 FROM trf_months ORDER BY year DESC, month DESC
                 """)) {
            while (rs.next()) {
                out.add(new TrfMonthRecord(
                    rs.getLong("id"), rs.getInt("year"), rs.getInt("month"),
                    rs.getString("status"), rs.getInt("nb_clients"),
                    rs.getDouble("total_montant"), rs.getDouble("total_nous_doit"),
                    rs.getString("closed_at")));
            }
        } catch (SQLException ignored) {}
        return out;
    }

    public synchronized List<TrfHistoryRecord> getTrfHistoryForMonth(long monthId) {
        List<TrfHistoryRecord> out = new ArrayList<>();
        try (PreparedStatement ps = conn.prepareStatement("""
                SELECT id, month_id, client_name, client_code, encaissements, montant_facturer,
                       nous_doit_prec, sommes_reverser, etat, iban, non_compensation
                FROM trf_history WHERE month_id=? ORDER BY client_name
                """)) {
            ps.setLong(1, monthId);
            try (ResultSet rs = ps.executeQuery()) {
                while (rs.next()) {
                    out.add(new TrfHistoryRecord(
                        rs.getLong("id"), rs.getLong("month_id"),
                        rs.getString("client_name"), rs.getString("client_code"),
                        rs.getDouble("encaissements"), rs.getDouble("montant_facturer"),
                        rs.getDouble("nous_doit_prec"), rs.getDouble("sommes_reverser"),
                        rs.getString("etat"), rs.getString("iban"),
                        rs.getInt("non_compensation") == 1));
                }
            }
        } catch (SQLException ignored) {}
        return out;
    }

    public synchronized List<double[]> getClientMonthlyHistory(String clientName, int limit)
            throws SQLException {
        List<double[]> out = new ArrayList<>();
        try (PreparedStatement ps = conn.prepareStatement("""
                SELECT tm.month, tm.year, th.montant_facturer
                FROM trf_history th
                JOIN trf_months tm ON th.month_id = tm.id
                WHERE th.client_name = ?
                ORDER BY tm.year DESC, tm.month DESC
                LIMIT ?
                """)) {
            ps.setString(1, clientName);
            ps.setInt(2, limit);
            try (ResultSet rs = ps.executeQuery()) {
                while (rs.next()) {
                    out.add(new double[]{rs.getInt("month"), rs.getInt("year"),
                                        rs.getDouble("montant_facturer")});
                }
            }
        }
        Collections.reverse(out);
        return out;
    }

    // -------------------------------------------------------------------------
    // Internal helper
    // -------------------------------------------------------------------------

    private static void bindValue(PreparedStatement ps, int idx, Object v) throws SQLException {
        if (v instanceof Number n) {
            ps.setDouble(idx, n.doubleValue());
        } else if (v != null) {
            ps.setString(idx, v.toString());
        } else {
            ps.setNull(idx, Types.NULL);
        }
    }

    /** Upsert the computed summary for a company (from Érat de Créances sheet). */
    public synchronized void upsertCompanySummary(long companyId, String codeClient,
            String responsable, int nbDossiers, int nbSoldes, int nbGestion,
            int nbIrr, int nbArj, int nbAutres,
            double creancePrincipale, double recouvreTotal, double commissions,
            String dernierDossier) throws SQLException {
        String now = LocalDateTime.now().format(ISO);
        try (PreparedStatement ps = conn.prepareStatement("""
                INSERT INTO company_summaries
                  (company_id, code_client, responsable, nb_dossiers,
                   nb_soldes, nb_gestion, nb_irr, nb_arj, nb_autres,
                   creance_principale, recouvre_total, commissions,
                   dernier_dossier, last_sync)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                ON CONFLICT(company_id) DO UPDATE
                  SET code_client=excluded.code_client, responsable=excluded.responsable,
                      nb_dossiers=excluded.nb_dossiers, nb_soldes=excluded.nb_soldes,
                      nb_gestion=excluded.nb_gestion, nb_irr=excluded.nb_irr,
                      nb_arj=excluded.nb_arj, nb_autres=excluded.nb_autres,
                      creance_principale=excluded.creance_principale,
                      recouvre_total=excluded.recouvre_total,
                      commissions=excluded.commissions,
                      dernier_dossier=excluded.dernier_dossier,
                      last_sync=excluded.last_sync
                """)) {
            ps.setLong  (1,  companyId);
            ps.setString(2,  codeClient);
            ps.setString(3,  responsable);
            ps.setInt   (4,  nbDossiers);
            ps.setInt   (5,  nbSoldes);
            ps.setInt   (6,  nbGestion);
            ps.setInt   (7,  nbIrr);
            ps.setInt   (8,  nbArj);
            ps.setInt   (9,  nbAutres);
            ps.setDouble(10, creancePrincipale);
            ps.setDouble(11, recouvreTotal);
            ps.setDouble(12, commissions);
            ps.setString(13, dernierDossier);
            ps.setString(14, now);
            ps.executeUpdate();
        }
    }

    public synchronized List<Map<String, Object>> getAllCompanySummaries() {
        List<Map<String, Object>> out = new ArrayList<>();
        String sql = """
                SELECT c.id AS company_id, c.name, c.source_path,
                       cs.code_client, cs.responsable, cs.nb_dossiers,
                       cs.nb_soldes, cs.nb_gestion, cs.nb_irr, cs.nb_arj, cs.nb_autres,
                       cs.creance_principale, cs.recouvre_total, cs.commissions,
                       cs.dernier_dossier, cs.last_sync
                FROM companies c
                LEFT JOIN company_summaries cs ON cs.company_id = c.id
                ORDER BY c.name COLLATE NOCASE
                """;
        try (Statement st = conn.createStatement();
             ResultSet rs = st.executeQuery(sql)) {
            ResultSetMetaData meta = rs.getMetaData();
            int cols = meta.getColumnCount();
            while (rs.next()) {
                Map<String, Object> row = new LinkedHashMap<>();
                for (int i = 1; i <= cols; i++) {
                    row.put(meta.getColumnName(i), rs.getObject(i));
                }
                out.add(row);
            }
        } catch (SQLException e) {
            System.err.println("[DB] getAllCompanySummaries error: " + e.getMessage());
        }
        return out;
    }

    public synchronized void closeTrfMonth(int year, int month) throws SQLException {
        try (PreparedStatement ps = conn.prepareStatement(
                "UPDATE trf_months SET status='closed', closed_at=datetime('now') WHERE year=? AND month=?")) {
            ps.setInt(1, year);
            ps.setInt(2, month);
            ps.executeUpdate();
        }
    }

    /**
     * Returns per-company aggregates from creance_rows filtered by REMIS LE date range.
     * dateFrom and dateTo are inclusive, format "YYYY-MM-DD".
     * Returns list of maps with keys: company_id, name, creance_principale, recouvre_total,
     * nb_dossiers, nb_soldes.
     */
    public synchronized List<Map<String, Object>> getGlobalStatsByDateRange(
            String dateFrom, String dateTo) {
        List<Map<String, Object>> out = new ArrayList<>();

        boolean filter = dateFrom != null && dateTo != null
                && !dateFrom.isBlank() && !dateTo.isBlank();

        String sql = """
                SELECT c.id AS company_id, c.name,
                       COALESCE(cs.nb_dossiers, 0)         AS nb_dossiers,
                       COALESCE(cs.nb_soldes,   0)         AS nb_soldes,
                       COALESCE(cs.creance_principale, 0)  AS creance_principale,
                       COALESCE(cs.recouvre_total,     0)  AS recouvre_total,
                       COALESCE(cs.commissions,        0)  AS commissions
                FROM companies c
                LEFT JOIN company_summaries cs ON cs.company_id = c.id
                WHERE COALESCE(cs.nb_dossiers, 0) > 0
                """
                + (filter ? "  AND cs.dernier_dossier >= ? AND cs.dernier_dossier <= ?\n" : "")
                + "ORDER BY cs.creance_principale DESC NULLS LAST";

        try (PreparedStatement ps = conn.prepareStatement(sql)) {
            if (filter) {
                ps.setString(1, dateFrom);
                ps.setString(2, dateTo);
            }
            try (ResultSet rs = ps.executeQuery()) {
                ResultSetMetaData meta = rs.getMetaData();
                int cols = meta.getColumnCount();
                while (rs.next()) {
                    Map<String, Object> row = new LinkedHashMap<>();
                    for (int i = 1; i <= cols; i++) {
                        row.put(meta.getColumnName(i), rs.getObject(i));
                    }
                    out.add(row);
                }
            }
        } catch (SQLException e) {
            System.err.println("[DB] getGlobalStatsByDateRange error: " + e.getMessage());
        }
        return out;
    }

    public synchronized List<Map<String, String>> getAllMailTemplates() {
        List<Map<String, String>> out = new ArrayList<>();
        try (Statement st = conn.createStatement();
             ResultSet rs = st.executeQuery(
                     "SELECT id, name, body FROM mail_templates ORDER BY name COLLATE NOCASE")) {
            while (rs.next()) {
                Map<String, String> row = new LinkedHashMap<>();
                row.put("id",   rs.getString("id"));
                row.put("name", rs.getString("name"));
                row.put("body", rs.getString("body"));
                out.add(row);
            }
        } catch (SQLException e) {
            System.err.println("[DB] getAllMailTemplates: " + e.getMessage());
        }
        return out;
    }

    public synchronized void upsertMailTemplate(String name, String body) throws SQLException {
        String now = LocalDateTime.now().format(ISO);
        try (PreparedStatement ps = conn.prepareStatement("""
                INSERT INTO mail_templates (name, body, created_at, updated_at)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(name) DO UPDATE
                  SET body=excluded.body, updated_at=excluded.updated_at
                """)) {
            ps.setString(1, name.trim());
            ps.setString(2, body);
            ps.setString(3, now);
            ps.setString(4, now);
            ps.executeUpdate();
        }
    }

    public synchronized void deleteMailTemplate(String name) throws SQLException {
        try (PreparedStatement ps = conn.prepareStatement(
                "DELETE FROM mail_templates WHERE name = ?")) {
            ps.setString(1, name.trim());
            ps.executeUpdate();
        }
    }
}
