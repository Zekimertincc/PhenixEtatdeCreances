package com.zeki.merger.db;

import com.zeki.merger.model.CreanceRow;
import com.zeki.merger.trf.model.ClientSummary;

import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

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
}
