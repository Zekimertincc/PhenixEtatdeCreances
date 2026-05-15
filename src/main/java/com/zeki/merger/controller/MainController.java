package com.zeki.merger.controller;

import com.zeki.merger.AppPreferences;
import com.zeki.merger.db.DatabaseManager;
import com.zeki.merger.service.ClientInfoService;
import com.zeki.merger.service.ConsoControleComparator;
import com.zeki.merger.service.EspacePartageFixer;
import com.zeki.merger.service.EtatCreancesSyncService;
import com.zeki.merger.service.EtatPublicGenerator;
import com.zeki.merger.service.FolderWatchService;
import com.zeki.merger.service.MergeService;
import com.zeki.merger.service.ProcreancesComparator;
import com.zeki.merger.service.RecupNumFactureService;
import com.zeki.merger.trf.TrfGeneratorService;
import javafx.application.Platform;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.*;
import javafx.scene.layout.*;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import javafx.stage.Modality;
import javafx.stage.Stage;

import java.awt.Desktop;
import java.io.File;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

/**
 * Contrôleur principal de l'interface graphique Cabinet Phénix.
 *
 * <p>Ce contrôleur est lié au fichier FXML {@code main.fxml} et constitue
 * le point central d'orchestration de toutes les actions utilisateur.
 * Il suit le patron <em>Controller</em> de l'architecture MVC/FXML de JavaFX :
 * les composants graphiques sont injectés par le {@code FXMLLoader} via
 * les annotations {@code @FXML}, et les actions métier sont déléguées
 * aux services spécialisés.
 *
 * <h2>Responsabilités</h2>
 * <ul>
 *   <li>Initialiser la grille des boutons d'action au démarrage ({@link #initialize()}).</li>
 *   <li>Afficher et mettre à jour les badges d'état des fichiers configurés.</li>
 *   <li>Ouvrir la boîte de dialogue de configuration des chemins de fichiers.</li>
 *   <li>Lancer chaque tâche longue (TRF, États Publics, Comparaison, Correction,
 *       Consolidation) dans un <em>fil de fond</em> dédié via {@link ExecutorService}.</li>
 *   <li>Répercuter toute mise à jour de l'interface (barre de progression, journal,
 *       barre d'état) sur le fil JavaFX via {@code Platform.runLater()}.</li>
 * </ul>
 *
 * <h2>Architecture des fils d'exécution</h2>
 * <p>Un unique {@code ExecutorService} à fil unique ({@code newSingleThreadExecutor})
 * est utilisé pour toutes les tâches d'arrière-plan. Cela garantit qu'une seule
 * opération de lecture/écriture Excel s'exécute à la fois, évitant ainsi les
 * conflits de fichiers. Le fil est marqué <em>daemon</em> pour ne pas bloquer
 * la fermeture de l'application.
 */
public class MainController {

    // =========================================================================
    // Champs FXML — injectés automatiquement par le FXMLLoader
    // =========================================================================

    /** Conteneur horizontal affichant les badges d'état (un badge par fichier configuré). */
    @FXML private HBox        badgesBox;

    /**
     * Étiquette affichant le nombre de fichiers manquants (ex. "2 fichier(s) manquant(s)").
     * Masquée lorsque tous les fichiers sont correctement configurés.
     */
    @FXML private Label       missingFilesLabel;

    /**
     * Grille d'actions (2 colonnes × 3 rangées) dans laquelle les boutons
     * d'action sont ajoutés de manière programmatique dans {@link #initialize()}.
     */
    @FXML private GridPane    actionsGrid;

    /**
     * Barre de progression affichant l'avancement de la tâche en cours.
     * La valeur varie entre 0.0 (début) et 1.0 (terminé).
     */
    @FXML private ProgressBar progressBar;

    /**
     * Zone de journal (log) affichant les messages horodatés émis par
     * les services pendant leur exécution.
     */
    @FXML private TextArea    logArea;

    /** Conteneur horizontal de la barre d'état, masqué tant qu'aucune tâche n'est terminée. */
    @FXML private HBox        statusBar;

    /**
     * Étiquette de la barre d'état affichant le chemin du dernier fichier produit
     * (ex. "TRF Output: /chemin/vers/trf_export.xlsx").
     */
    @FXML private Label       statusLabel;

    /**
     * Bouton "Ouvrir le fichier" affiché dans la barre d'état après la
     * génération d'un fichier de sortie. Déclenche {@link #openFile()}.
     */
    @FXML private Button      openFileBtn;

    /**
     * Référence vers le contrôleur du tableau de bord (panneau secondaire).
     * Permet de rafraîchir le tableau de bord après une consolidation ou
     * une génération TRF, lorsque les données SQLite ont été mises à jour.
     * Peut être {@code null} si le panneau tableau de bord n'est pas présent
     * dans la scène courante.
     */
    @FXML private DashboardController dashboardController;

    // =========================================================================
    // Boutons d'action — créés programmatiquement, stockés pour enable/disable
    // =========================================================================

    /**
     * Bouton "Générer TRF" — lance le calcul des virements et compensations.
     * Stocké en champ pour pouvoir être désactivé pendant l'exécution d'une tâche.
     */
    private Button trfBtn;

    /**
     * Bouton "États Publics" — exporte les états de créances vers l'EspacePartagé.
     * Stocké en champ pour pouvoir être désactivé pendant l'exécution d'une tâche.
     */
    private Button etatBtn;

    /**
     * Bouton "Comparer des fichiers Excel" — détecte les écarts entre
     * PROCREANCES et ConsolidationGénérale.
     * Stocké en champ pour pouvoir être désactivé pendant l'exécution d'une tâche.
     */
    private Button cmpBtn;

    /**
     * Bouton "Corriger EspacePartagé" — met à jour les chemins dans le fichier
     * de correspondance clients/EspacePartagé.
     * Stocké en champ pour pouvoir être désactivé pendant l'exécution d'une tâche.
     */
    private Button fixBtn;

    /**
     * Bouton principal "CONSOLIDER" — lit tous les états de créances par société
     * et produit le fichier de consolidation générale.
     * Stocké en champ pour pouvoir être désactivé pendant l'exécution d'une tâche.
     */
    private Button runActionBtn;

    private Button controleBtn;
    private Button recupBtn;
    private Button infoBtn;
    private Button syncDbBtn;
    private Button watchToggleBtn;

    // =========================================================================
    // Services métier — instanciés une seule fois à la création du contrôleur
    // =========================================================================

    /**
     * Service de consolidation : scanne le dossier source, lit les fichiers Excel
     * de chaque société et produit le fichier {@code etat_creances_global.xlsx}.
     * Reçoit l'instance singleton de {@link DatabaseManager} pour persister
     * les lignes de créances en base SQLite après chaque consolidation.
     */
    private final MergeService          mergeService          = new MergeService(DatabaseManager.getInstance());

    /**
     * Service de correction des chemins EspacePartagé : lit le fichier de
     * correspondance {@code CorrespondanceClient-EspacePartage.xlsx} et met
     * à jour les chemins obsolètes ou mal formés.
     */
    private final EspacePartageFixer    espacePartageFixer    = new EspacePartageFixer();

    /**
     * Service de génération des états publics : produit un fichier Excel
     * au format "L_ETAT_DE_CREANCES_[SOCIETE].xlsx" pour chaque société,
     * destiné à être déposé dans l'EspacePartagé.
     */
    private final EtatPublicGenerator   etatPublicGenerator   = new EtatPublicGenerator();

    /**
     * Service de génération TRF (Tableau de Remboursement et de Facturation) :
     * lit les trois fichiers sources (ConsolidationGénérale, Listing Cabinet Phénix,
     * Tableau de Bord), calcule virements et compensations par client, et produit
     * le fichier {@code trf_export.xlsx}. Reçoit l'instance singleton de
     * {@link DatabaseManager} pour persister les résumés TRF en base SQLite.
     */
    private final TrfGeneratorService   trfGeneratorService   = new TrfGeneratorService(DatabaseManager.getInstance());

    /**
     * Service de comparaison PROCREANCES : met en regard les données du fichier
     * PROCREANCES (.xls) et du fichier ConsolidationGénérale (.xlsx), produit
     * un rapport Excel détaillant les écarts supérieurs à 0,05 €.
     */
    private final ProcreancesComparator   procreancesComparator   = new ProcreancesComparator();
    private final ConsoControleComparator  consoControleComparator = new ConsoControleComparator();
    private final RecupNumFactureService   recupNumFactureService  = new RecupNumFactureService();
    private final ClientInfoService        clientInfoService       = new ClientInfoService();
    private final EtatCreancesSyncService  syncService             = new EtatCreancesSyncService(DatabaseManager.getInstance());
    private final FolderWatchService       watchService            = new FolderWatchService(syncService, this::onWatchEvent);

    // =========================================================================
    // Fil de fond et utilitaires
    // =========================================================================

    /**
     * Exécuteur à fil unique utilisé pour toutes les tâches d'arrière-plan.
     *
     * <p>Le fil unique garantit la sérialisation des opérations (une seule
     * tâche à la fois) et évite les conflits d'accès concurrentiel aux fichiers.
     * Le fil est nommé {@code "merge-worker"} pour faciliter le débogage dans
     * les traces de pile. Il est marqué <em>daemon</em> afin que la JVM puisse
     * se terminer normalement sans attendre la fin d'une éventuelle tâche pendante.
     */
    private final ExecutorService executor = Executors.newSingleThreadExecutor(r -> {
        Thread t = new Thread(r, "merge-worker");
        t.setDaemon(true);
        return t;
    });

    /**
     * Formateur d'heure utilisé pour horodater les messages du journal.
     * Format : {@code HH:mm:ss} (ex. "14:32:07").
     */
    private static final DateTimeFormatter TIME_FMT = DateTimeFormatter.ofPattern("HH:mm:ss");

    /**
     * Référence vers le dernier fichier produit par une tâche.
     * Utilisée par {@link #openFile()} pour ouvrir ce fichier
     * avec l'application système associée à son extension.
     * Réinitialisée à {@code null} au début de chaque nouvelle tâche.
     */
    private File lastOutputFile;

    // =========================================================================
    // Cycle de vie JavaFX
    // =========================================================================

    /**
     * Méthode d'initialisation appelée automatiquement par le {@code FXMLLoader}
     * après l'injection de tous les champs {@code @FXML}.
     *
     * <p>Effectue dans l'ordre :
     * <ol>
     *   <li>Réinitialisation de la barre de progression à 0.</li>
     *   <li>Masquage de la barre d'état (aucun résultat disponible au démarrage).</li>
     *   <li>Rafraîchissement des badges de fichiers (vérifie l'existence de chaque
     *       chemin sauvegardé dans {@link AppPreferences}).</li>
     *   <li>Création programmatique des 5 boutons d'action et ajout dans la grille.</li>
     * </ol>
     *
     * <p>Les boutons sont créés ici et non dans le FXML pour permettre un graphisme
     * riche (nom + description sur deux lignes) et un stockage en champ pour
     * l'activation/désactivation groupée.
     */
    @FXML
    public void initialize() {
        // Barre de progression : aucun travail en cours au démarrage
        progressBar.setProgress(0);

        // Barre d'état masquée tant qu'aucune tâche n'a produit de fichier
        statusBar.setVisible(false);

        // Vérification et affichage de l'état de chaque fichier configuré
        refreshFileBadges();

        // Création des boutons d'action avec leur nom, description et style CSS
        trfBtn       = createActionBtn("Générer TRF",                 "Calcul virements et compensations",      "secondary-btn", e -> generateTrf());
        etatBtn      = createActionBtn("États Publics",               "Exporter vers EspacePartagé",            "secondary-btn", e -> generateEtatPublic());
        cmpBtn       = createActionBtn("Comparer des fichiers Excel", "Détecter les écarts PROCREANCES",        "secondary-btn", e -> compareProcreances());
        fixBtn       = createActionBtn("Corriger EspacePartagé",      "Mettre à jour les chemins",              "secondary-btn", e -> fixPaths());
        controleBtn  = createActionBtn("Contrôle Facturation",        "Comparer Contrôle vs Consolidation",     "secondary-btn", e -> compareConsoControle());
        recupBtn     = createActionBtn("Récup. Factures",             "Écrire N° facture → D13",               "secondary-btn", e -> recupNumFacture());
        infoBtn      = createActionBtn("Info Clients",                "Créer feuille Infos depuis Listing",     "secondary-btn", e -> clientInfo());
        runActionBtn = createActionBtn("▶  CONSOLIDER",               "Lire les états → ConsolidationGénérale", "run-btn",       e -> run());

        // Placement des boutons dans la grille
        actionsGrid.add(trfBtn,       0, 0);
        actionsGrid.add(etatBtn,      1, 0);
        actionsGrid.add(cmpBtn,       0, 1);
        actionsGrid.add(fixBtn,       1, 1);

        GridPane.setColumnSpan(controleBtn, 2);
        actionsGrid.add(controleBtn,  0, 2);

        actionsGrid.add(recupBtn,     0, 3);
        actionsGrid.add(infoBtn,      1, 3);

        syncDbBtn      = createActionBtn("DB Güncelle",   "Synchroniser toutes les sociétés",  "secondary-btn", e -> syncDatabase());
        watchToggleBtn = createActionBtn("▶ Surveiller",  "Surveillance automatique",           "secondary-btn", e -> toggleWatch());
        actionsGrid.add(syncDbBtn,      0, 4);
        actionsGrid.add(watchToggleBtn, 1, 4);

        GridPane.setColumnSpan(runActionBtn, 2);
        actionsGrid.add(runActionBtn, 0, 5);
        controleBtn.setVisible(false);
        controleBtn.setManaged(false);
        recupBtn.setVisible(false);
        recupBtn.setManaged(false);
        infoBtn.setVisible(false);
        infoBtn.setManaged(false);

        if (AppPreferences.isWatchEnabled()) {
            File root = new File(AppPreferences.getMergeRoot());
            if (root.isDirectory()) {
                watchService.start(root);
                updateWatchToggleLabel(true);
                appendLog("Surveillance automatique activée.");
            }
        }
    }

    // =========================================================================
    // Configuration des fichiers
    // =========================================================================

    /**
     * Ouvre la boîte de dialogue modale de configuration des fichiers.
     *
     * <p>Cette méthode est référencée dans {@code main.fxml} via l'attribut
     * {@code onAction} du bouton "⚙ Configurer", d'où la nécessité de
     * l'annotation {@code @FXML}.
     *
     * <p>La boîte de dialogue affiche 6 lignes de configuration, chacune composée
     * d'une étiquette, d'un libellé de chemin (coloré en vert si valide, rouge sinon)
     * et d'un bouton "Choisir"/"Changer". Les chemins sont stockés temporairement
     * dans le tableau local {@code paths[]} et ne sont persistés dans
     * {@link AppPreferences} qu'après un clic sur "Enregistrer".
     *
     * <p>Les 6 fichiers configurables sont :
     * <ol>
     *   <li>Dossier source (dossier racine des sociétés)</li>
     *   <li>Dossier de sortie (destination des fichiers générés)</li>
     *   <li>ConsolidationGénérale.xlsx (données principales TRF)</li>
     *   <li>Listing Cabinet Phénix.xlsx (métadonnées clients : IBAN, BIC, code)</li>
     *   <li>Tableau de Bord.xlsx (soldes précédents par client)</li>
     *   <li>Export PROCREANCES.xls (données PROCREANCES pour comparaison)</li>
     * </ol>
     */
    @FXML
    private void openFileConfig() {
        // Lecture des chemins actuellement sauvegardés dans les préférences utilisateur
        String[] paths = {
            AppPreferences.getMergeRoot(),
            AppPreferences.getOutputFolder(),
            AppPreferences.getTrfConso(),
            AppPreferences.getTrfListing(),
            AppPreferences.getTrfTableau(),
            AppPreferences.getProcreancesPath(),
            AppPreferences.getControlePath(),
            AppPreferences.getRecupFacturePath()
        };

        // Libellés affichés devant chaque champ de sélection
        String[]  labels = {"Dossier source",        "Dossier de sortie",        "ConsolidationGénérale",
                             "Listing Cabinet Phénix", "Tableau de Bord",         "Export PROCREANCES",
                             "Contrôle Facturation",  "Récup. Num Facture"};

        // true = sélecteur de dossier, false = sélecteur de fichier
        boolean[] isDir  = {true,  true,  false, false, false, false, false, false};

        // Extensions autorisées pour les sélecteurs de fichiers (null = pas de filtre)
        String[]  exts   = {null,  null,  "xlsx", "xlsx", "xlsx", "xls", "xlsx", "xlsx"};

        // Création de la fenêtre modale de configuration
        Stage dialog = new Stage();
        dialog.initModality(Modality.APPLICATION_MODAL);          // bloque la fenêtre principale
        dialog.initOwner(badgesBox.getScene().getWindow());        // rattaché à la fenêtre parente
        dialog.setTitle("Configuration des fichiers");
        dialog.setResizable(false);

        VBox root = new VBox(10);
        root.setPadding(new Insets(20));
        root.setPrefWidth(700);

        // Tableau de références vers les étiquettes de chemin, pour mise à jour dynamique
        Label[] pathLabels = new Label[paths.length];

        // Construction des 6 lignes de sélection (label + chemin + bouton Changer)
        for (int i = 0; i < paths.length; i++) {
            final int idx = i; // copie finale pour utilisation dans le lambda

            HBox row = new HBox(8);
            row.setAlignment(Pos.CENTER_LEFT);

            // Étiquette descriptive du champ (ex. "Dossier source :")
            Label lbl = new Label(labels[i] + ":");
            lbl.setMinWidth(200);
            lbl.setStyle("-fx-font-weight: bold; -fx-font-family: 'Courier New', monospace;");

            // Étiquette affichant le chemin courant (coloré selon sa validité)
            pathLabels[i] = new Label();
            pathLabels[i].setMaxWidth(Double.MAX_VALUE);
            HBox.setHgrow(pathLabels[i], Priority.ALWAYS);
            updatePathLabel(pathLabels[i], paths[i], isDir[i]);

            // Bouton "Choisir" (chemin vide) ou "Changer" (chemin déjà défini)
            Button changeBtn = new Button(paths[i].isEmpty() ? "Choisir" : "Changer");
            changeBtn.getStyleClass().add("secondary-btn");
            changeBtn.setOnAction(ev -> {
                // Ouvrir le sélecteur approprié (dossier ou fichier)
                File chosen = isDir[idx]
                    ? dialogPickDirectory(dialog, labels[idx], paths[idx])
                    : dialogPickFile(dialog, labels[idx], paths[idx], exts[idx]);
                if (chosen != null) {
                    // Mise à jour du tableau temporaire et du libellé affiché
                    paths[idx] = chosen.getAbsolutePath();
                    updatePathLabel(pathLabels[idx], paths[idx], isDir[idx]);
                    changeBtn.setText("Changer"); // un chemin est maintenant défini
                }
            });

            row.getChildren().addAll(lbl, pathLabels[i], changeBtn);
            root.getChildren().add(row);
        }

        // ---- Pied de page : boutons Annuler et Enregistrer ----
        HBox footer = new HBox(8);
        footer.setAlignment(Pos.CENTER_RIGHT);
        footer.setPadding(new Insets(10, 0, 0, 0));

        // Bouton Annuler : ferme la boîte de dialogue sans sauvegarder les modifications
        Button cancelBtn = new Button("Annuler");
        cancelBtn.getStyleClass().add("secondary-btn");
        cancelBtn.setOnAction(ev -> dialog.close());

        // Bouton Enregistrer : persiste tous les chemins dans AppPreferences
        // puis ferme la boîte de dialogue et rafraîchit les badges
        Button saveBtn = new Button("Enregistrer");
        saveBtn.getStyleClass().add("run-btn");
        saveBtn.setOnAction(ev -> {
            // Persistence de chaque chemin dans les préférences Java (registre/plist)
            AppPreferences.setMergeRoot(paths[0]);
            AppPreferences.setOutputFolder(paths[1]);
            AppPreferences.setTrfConso(paths[2]);
            AppPreferences.setTrfListing(paths[3]);
            AppPreferences.setTrfTableau(paths[4]);
            AppPreferences.setProcreancesPath(paths[5]);
            AppPreferences.setControlePath(paths[6]);
            AppPreferences.setRecupFacturePath(paths[7]);
            dialog.close();
            // Mise à jour des badges pour refléter les nouveaux chemins
            refreshFileBadges();
        });

        footer.getChildren().addAll(cancelBtn, saveBtn);
        root.getChildren().add(footer);

        // Application de la feuille de style CSS de la fenêtre principale à la boîte de dialogue
        Scene scene = new Scene(root);
        if (!badgesBox.getScene().getStylesheets().isEmpty()) {
            scene.getStylesheets().addAll(badgesBox.getScene().getStylesheets());
        }
        dialog.setScene(scene);
        dialog.showAndWait(); // attente bloquante jusqu'à la fermeture de la boîte de dialogue
    }

    /**
     * Rafraîchit la zone des badges d'état ({@code badgesBox}) en reconstruisant
     * un badge pour chacun des 6 fichiers/dossiers configurés.
     *
     * <p>Chaque badge affiche le nom abrégé du fichier suivi de "✓" (vert, classe CSS
     * {@code badge-ok}) si le chemin est valide et accessible, ou de "✗" (rouge,
     * classe CSS {@code badge-missing}) dans le cas contraire.
     *
     * <p>Le libellé {@code missingFilesLabel} est affiché ou masqué selon le nombre
     * de fichiers manquants détectés.
     */
    private void refreshFileBadges() {
        // Nettoyage des anciens badges
        badgesBox.getChildren().clear();
        int missing = 0;

        // Création d'un badge pour chacun des 6 éléments configurables
        missing += addBadge("Dossier source",       AppPreferences.getMergeRoot(),       true);
        missing += addBadge("Dossier sortie",        AppPreferences.getOutputFolder(),    true);
        missing += addBadge("ConsolidationGénérale", AppPreferences.getTrfConso(),        false);
        missing += addBadge("Listing",               AppPreferences.getTrfListing(),      false);
        missing += addBadge("Tableau de bord",       AppPreferences.getTrfTableau(),      false);
        missing += addBadge("PROCREANCES",           AppPreferences.getProcreancesPath(), false);
        missing += addBadge("Contrôle Fact.",        AppPreferences.getControlePath(),    false);
        missing += addBadge("Récup Factures",        AppPreferences.getRecupFacturePath(), false);

        // Affichage ou masquage du compteur de fichiers manquants
        if (missing > 0) {
            missingFilesLabel.setText(missing + " fichier(s) manquant(s)");
            missingFilesLabel.setVisible(true);
            missingFilesLabel.setManaged(true);
        } else {
            // Tous les fichiers sont configurés et accessibles : masquer le libellé
            missingFilesLabel.setVisible(false);
            missingFilesLabel.setManaged(false);
        }
    }

    /**
     * Crée et ajoute un badge d'état pour un fichier ou dossier donné dans
     * la zone {@code badgesBox}.
     *
     * @param label       Libellé court affiché dans le badge (ex. "Dossier source").
     * @param path        Chemin absolu du fichier ou dossier à vérifier.
     * @param isDirectory {@code true} si le chemin désigne un dossier, {@code false} pour un fichier.
     * @return {@code 0} si le fichier/dossier est valide et accessible, {@code 1} sinon.
     */
    private int addBadge(String label, String path, boolean isDirectory) {
        // Vérification de l'existence : isDirectory() pour un dossier, exists() pour un fichier
        boolean ok = !path.isEmpty()
            && (isDirectory ? new File(path).isDirectory() : new File(path).exists());

        // Le badge affiche "✓" (OK) ou "✗" (manquant) avec la classe CSS correspondante
        Label badge = new Label(label + (ok ? " ✓" : " ✗"));
        badge.getStyleClass().add(ok ? "badge-ok" : "badge-missing");
        badgesBox.getChildren().add(badge);

        // Retourne 0 si OK, 1 si manquant (permet au caller d'additionner les manques)
        return ok ? 0 : 1;
    }

    // =========================================================================
    // Gestionnaires d'actions (tâches de fond)
    // =========================================================================

    /**
     * Lance la génération du TRF dans un fil de fond.
     *
     * <p>Valide au préalable la présence des trois fichiers sources requis
     * (ConsolidationGénérale, Listing Cabinet Phénix, Tableau de Bord) ainsi que
     * l'existence du dossier de sortie. En cas d'absence, un message d'erreur
     * est immédiatement inscrit dans le journal et la méthode retourne sans lancer
     * aucune tâche.
     *
     * <p>Si la validation réussit :
     * <ol>
     *   <li>Tous les boutons d'action sont désactivés.</li>
     *   <li>La barre de progression, le journal et la barre d'état sont réinitialisés.</li>
     *   <li>La tâche est soumise à l'{@link #executor} : elle appelle
     *       {@link TrfGeneratorService#generate} en transmettant un callback
     *       de progression qui met à jour barre et journal via {@code Platform.runLater()}.</li>
     *   <li>À la fin de la tâche, les boutons sont réactivés, le fichier résultat
     *       est affiché dans la barre d'état et le tableau de bord est rafraîchi.</li>
     * </ol>
     *
     * <p>Toutes les modifications de l'interface dans le fil de fond passent
     * obligatoirement par {@code Platform.runLater()} pour respecter le modèle
     * de thread unique de JavaFX.
     */
    private void generateTrf() {
        // Lecture des chemins depuis les préférences utilisateur
        String consoPath   = AppPreferences.getTrfConso();
        String listingPath = AppPreferences.getTrfListing();
        String tableauPath = AppPreferences.getTrfTableau();
        String outputPath  = AppPreferences.getOutputFolder();

        // Validation : les trois fichiers TRF doivent être configurés
        if (consoPath.isEmpty() || listingPath.isEmpty() || tableauPath.isEmpty()) {
            appendLog("ERROR: Configurez les trois fichiers TRF avant de générer."); return;
        }

        File consoFile    = new File(consoPath);
        File listingFile  = new File(listingPath);
        File tableauFile  = new File(tableauPath);
        File outputFolder = new File(outputPath);

        // Vérification de l'existence physique de chaque fichier et du dossier de sortie
        if (!consoFile.exists())         { appendLog("ERROR: Fichier introuvable — " + consoPath);   return; }
        if (!listingFile.exists())       { appendLog("ERROR: Fichier introuvable — " + listingPath); return; }
        if (!tableauFile.exists())       { appendLog("ERROR: Fichier introuvable — " + tableauPath); return; }
        if (!outputFolder.isDirectory()) { appendLog("ERROR: Dossier sortie introuvable — " + outputPath); return; }

        // Préparation de l'interface avant le démarrage de la tâche
        setAllButtonsDisabled(true);   // désactivation des boutons pendant le traitement
        statusBar.setVisible(false);   // masquage de l'ancienne barre d'état
        progressBar.setProgress(0);    // réinitialisation de la barre de progression
        logArea.clear();               // effacement du journal précédent
        lastOutputFile = null;         // réinitialisation de la référence au fichier de sortie

        // Soumission de la tâche TRF au fil de fond
        executor.submit(() -> {
            try {
                // Appel du service TRF avec un callback de progression double (valeur + message)
                // Le callback est exécuté depuis le fil de fond → Platform.runLater() est obligatoire
                File result = trfGeneratorService.generate(consoFile, listingFile, tableauFile, outputFolder,
                    (prog, msg) -> Platform.runLater(() -> { progressBar.setProgress(prog); appendLog(msg); }));

                // Mise à jour de l'interface dans le fil JavaFX après la fin du traitement
                Platform.runLater(() -> {
                    if (result != null) {
                        lastOutputFile = result;
                        statusLabel.setText("TRF Output: " + result.getAbsolutePath());
                        openFileBtn.setVisible(true);   // affichage du bouton "Ouvrir"
                        statusBar.setVisible(true);     // affichage de la barre d'état
                        // Rafraîchissement du tableau de bord si disponible (données SQLite mises à jour)
                        if (dashboardController != null) dashboardController.refresh();
                    }
                    setAllButtonsDisabled(false); // réactivation des boutons
                });
            } catch (Exception e) {
                // En cas d'erreur non récupérée, affichage dans le journal et réactivation des boutons
                Platform.runLater(() -> { appendLog("FATAL: " + e.getMessage()); setAllButtonsDisabled(false); });
            }
        });
    }

    /**
     * Lance la génération des États Publics dans un fil de fond.
     *
     * <p>Parcourt le dossier source (racine des sociétés) et produit pour chaque
     * société un fichier Excel au format {@code L_ETAT_DE_CREANCES_[SOCIETE].xlsx},
     * destiné à être déposé dans l'EspacePartagé correspondant.
     *
     * <p>Valide d'abord l'existence du dossier source. Si celui-ci est absent,
     * un message d'erreur est écrit dans le journal et la méthode retourne.
     *
     * <p>Toutes les mises à jour de l'interface sont effectuées via
     * {@code Platform.runLater()} pour respecter le fil JavaFX.
     */
    private void generateEtatPublic() {
        // Vérification du dossier source configuré
        File rootFolder = new File(AppPreferences.getMergeRoot());
        if (!rootFolder.isDirectory()) {
            appendLog("ERROR: Dossier source introuvable — " + rootFolder.getAbsolutePath()); return;
        }

        // Préparation de l'interface
        setAllButtonsDisabled(true);
        statusBar.setVisible(false);
        progressBar.setProgress(0);
        logArea.clear();
        lastOutputFile = null;

        // Soumission de la tâche de génération des états publics au fil de fond
        executor.submit(() -> {
            try {
                // Le callback de progression met à jour barre et journal sur le fil JavaFX
                etatPublicGenerator.generate(rootFolder,
                    (prog, msg) -> Platform.runLater(() -> { progressBar.setProgress(prog); appendLog(msg); }));

                // Mise à jour de la barre d'état après succès
                Platform.runLater(() -> {
                    statusLabel.setText("Etat Public files written to EspacePartagé paths.");
                    openFileBtn.setVisible(false); // pas de fichier unique à ouvrir
                    statusBar.setVisible(true);
                    setAllButtonsDisabled(false);
                });
            } catch (Exception e) {
                Platform.runLater(() -> { appendLog("FATAL: " + e.getMessage()); setAllButtonsDisabled(false); });
            }
        });
    }

    /**
     * Lance la correction des chemins EspacePartagé dans un fil de fond.
     *
     * <p>Lit le fichier de correspondance {@code CorrespondanceClient-EspacePartage.xlsx}
     * situé dans le dossier source, vérifie chaque chemin EspacePartagé et corrige
     * ceux qui sont obsolètes ou mal formés. Le fichier corrigé est soit sauvegardé
     * sur place (si {@code AppConfig.FIX_OVERWRITE == true}), soit écrit avec un
     * suffixe {@code _fixed}.
     *
     * <p>Valide d'abord l'existence du dossier source. Toutes les mises à jour
     * de l'interface sont effectuées via {@code Platform.runLater()}.
     */
    private void fixPaths() {
        // Vérification du dossier source
        File rootFolder = new File(AppPreferences.getMergeRoot());
        if (!rootFolder.isDirectory()) {
            appendLog("ERROR: Dossier source introuvable — " + rootFolder.getAbsolutePath()); return;
        }

        // Préparation de l'interface
        setAllButtonsDisabled(true);
        statusBar.setVisible(false);
        progressBar.setProgress(0);
        logArea.clear();
        lastOutputFile = null;

        // Soumission de la tâche de correction au fil de fond
        executor.submit(() -> {
            try {
                // Le callback de progression met à jour barre et journal sur le fil JavaFX
                File result = espacePartageFixer.fix(rootFolder,
                    (progress, msg) -> Platform.runLater(() -> { progressBar.setProgress(progress); appendLog(msg); }));

                // Affichage du chemin du fichier corrigé dans la barre d'état
                Platform.runLater(() -> {
                    lastOutputFile = result;
                    statusLabel.setText("Saved: " + result.getAbsolutePath());
                    openFileBtn.setVisible(true);
                    statusBar.setVisible(true);
                    setAllButtonsDisabled(false);
                });
            } catch (Exception e) {
                Platform.runLater(() -> { appendLog("FATAL: " + e.getMessage()); setAllButtonsDisabled(false); });
            }
        });
    }

    /**
     * Lance la consolidation complète des états de créances dans un fil de fond.
     *
     * <p>C'est l'action principale de l'application. Le service {@link MergeService}
     * scanne le dossier source, identifie les sous-dossiers "etat de creances" de
     * chaque société, lit le fichier {@code etat_*.xlsx} correspondant, filtre les
     * lignes pertinentes (colonne S non vide) et produit le fichier consolidé
     * {@code etat_creances_global.xlsx} dans le dossier de sortie.
     *
     * <p>Valide d'abord l'existence du dossier source et du dossier de sortie.
     * En cas de succès, le tableau de bord est rafraîchi (nouvelles données SQLite).
     * Toutes les mises à jour de l'interface passent par {@code Platform.runLater()}.
     */
    private void run() {
        // Lecture des dossiers configurés
        File rootFolder   = new File(AppPreferences.getMergeRoot());
        File outputFolder = new File(AppPreferences.getOutputFolder());

        // Validation des deux dossiers requis
        if (!rootFolder.isDirectory()) {
            appendLog("ERROR: Dossier source introuvable — " + rootFolder.getAbsolutePath()); return;
        }
        if (!outputFolder.isDirectory()) {
            appendLog("ERROR: Dossier sortie introuvable — " + outputFolder.getAbsolutePath()); return;
        }

        // Préparation de l'interface
        setAllButtonsDisabled(true);
        statusBar.setVisible(false);
        progressBar.setProgress(0);
        logArea.clear();
        lastOutputFile = null;

        // Soumission de la tâche de consolidation au fil de fond
        executor.submit(() -> {
            try {
                // Fusion de tous les fichiers Excel — le callback met à jour l'interface sur le fil JavaFX
                File result = mergeService.merge(rootFolder, outputFolder,
                    (progress, msg) -> Platform.runLater(() -> { progressBar.setProgress(progress); appendLog(msg); }));

                // Affichage du résultat et rafraîchissement du tableau de bord
                Platform.runLater(() -> {
                    if (result != null) {
                        lastOutputFile = result;
                        statusLabel.setText("Output: " + result.getAbsolutePath());
                        openFileBtn.setVisible(true);
                        statusBar.setVisible(true);
                        // Les données SQLite ont été mises à jour : rafraîchissement du tableau de bord
                        if (dashboardController != null) dashboardController.refresh();
                    }
                    setAllButtonsDisabled(false);
                });
            } catch (Exception e) {
                Platform.runLater(() -> { appendLog("FATAL: " + e.getMessage()); setAllButtonsDisabled(false); });
            }
        });
    }

    /**
     * Lance la comparaison entre le fichier PROCREANCES et le fichier
     * ConsolidationGénérale dans un fil de fond.
     *
     * <p>Le service {@link ProcreancesComparator} met en correspondance les clients
     * des deux fichiers, calcule les écarts sur trois colonnes (Hono. TTC,
     * Disponible, Reversement) et produit un rapport Excel horodaté avec trois
     * feuilles : "Tous les clients", "Écarts uniquement" et "Non appariés".
     *
     * <p>Valide d'abord la présence du fichier PROCREANCES (.xls), du fichier
     * ConsolidationGénérale (.xlsx) et du dossier de sortie. Une fois le rapport
     * généré, il est automatiquement ouvert avec l'application système par défaut
     * (ex. Microsoft Excel) via {@code Desktop.getDesktop().open()}.
     *
     * <p>Toutes les mises à jour de l'interface passent par {@code Platform.runLater()}.
     */
    private void compareProcreances() {
        // Lecture des chemins depuis les préférences utilisateur
        String procPath   = AppPreferences.getProcreancesPath();
        String consoPath  = AppPreferences.getTrfConso();   // même fichier que pour le TRF
        String outputPath = AppPreferences.getOutputFolder();

        // Validation : les deux fichiers sources doivent être configurés
        if (procPath.isEmpty() || consoPath.isEmpty()) {
            appendLog("ERROR: Configurez Export PROCREANCES et ConsolidationGénérale avant de comparer."); return;
        }

        File procFile    = new File(procPath);
        File consoFile   = new File(consoPath);
        File outputFolder = new File(outputPath);

        // Vérification de l'existence physique des fichiers et du dossier de sortie
        if (!procFile.exists())          { appendLog("ERROR: Fichier introuvable — " + procPath);  return; }
        if (!consoFile.exists())         { appendLog("ERROR: Fichier introuvable — " + consoPath); return; }
        if (!outputFolder.isDirectory()) { appendLog("ERROR: Dossier sortie introuvable — " + outputPath); return; }

        // Préparation de l'interface
        setAllButtonsDisabled(true);
        statusBar.setVisible(false);
        progressBar.setProgress(0);
        logArea.clear();
        lastOutputFile = null;

        // Soumission de la tâche de comparaison au fil de fond
        executor.submit(() -> {
            try {
                // Exécution de la comparaison — le callback met à jour barre et journal sur le fil JavaFX
                File report = procreancesComparator.compare(procFile, consoFile, outputFolder,
                    (prog, msg) -> Platform.runLater(() -> { progressBar.setProgress(prog); appendLog(msg); }));

                // Affichage du rapport et ouverture automatique avec l'application système
                Platform.runLater(() -> {
                    setAllButtonsDisabled(false);
                    if (report != null) {
                        lastOutputFile = report;
                        statusLabel.setText("Rapport: " + report.getAbsolutePath());
                        openFileBtn.setVisible(true);
                        statusBar.setVisible(true);
                        // Ouverture automatique du rapport avec l'application associée (ex. Excel)
                        try { Desktop.getDesktop().open(report); } catch (Exception ignored) {}
                    }
                });
            } catch (Exception ex) {
                Platform.runLater(() -> { appendLog("FATAL: " + ex.getMessage()); setAllButtonsDisabled(false); });
            }
        });
    }

    /**
     * Lance la comparaison Contrôle_Facturation vs ConsolidationGénérale dans un fil de fond.
     */
    private void compareConsoControle() {
        String controlePath = AppPreferences.getControlePath();
        String consoPath    = AppPreferences.getTrfConso();
        String outputPath   = AppPreferences.getOutputFolder();

        if (controlePath.isEmpty() || consoPath.isEmpty()) {
            appendLog("ERROR: Configurez Contrôle Facturation et ConsolidationGénérale avant de comparer."); return;
        }

        File controleFile  = new File(controlePath);
        File consoFile     = new File(consoPath);
        File outputFolder  = new File(outputPath);

        if (!controleFile.exists())      { appendLog("ERROR: Fichier introuvable — " + controlePath);  return; }
        if (!consoFile.exists())         { appendLog("ERROR: Fichier introuvable — " + consoPath);      return; }
        if (!outputFolder.isDirectory()) { appendLog("ERROR: Dossier sortie introuvable — " + outputPath); return; }

        setAllButtonsDisabled(true);
        statusBar.setVisible(false);
        progressBar.setProgress(0);
        logArea.clear();
        lastOutputFile = null;

        executor.submit(() -> {
            try {
                File report = consoControleComparator.compare(controleFile, consoFile, outputFolder,
                    (prog, msg) -> Platform.runLater(() -> { progressBar.setProgress(prog); appendLog(msg); }));

                Platform.runLater(() -> {
                    setAllButtonsDisabled(false);
                    if (report != null) {
                        lastOutputFile = report;
                        statusLabel.setText("Rapport: " + report.getAbsolutePath());
                        openFileBtn.setVisible(true);
                        statusBar.setVisible(true);
                        try { Desktop.getDesktop().open(report); } catch (Exception ignored) {}
                    }
                });
            } catch (Exception ex) {
                Platform.runLater(() -> { appendLog("FATAL: " + ex.getMessage()); setAllButtonsDisabled(false); });
            }
        });
    }

    /**
     * Lance la récupération des numéros de facture dans un fil de fond.
     * Lit RecupNumFacture.xlsx et écrit chaque N° en D13 de la feuille
     * "Facture en préparation" de chaque dossier société.
     */
    private void recupNumFacture() {
        String recupPath  = AppPreferences.getRecupFacturePath();
        String rootPath   = AppPreferences.getMergeRoot();

        if (recupPath.isEmpty()) {
            appendLog("ERROR: Configurez le fichier Récup. Num Facture avant de lancer."); return;
        }
        File recupFile  = new File(recupPath);
        File rootFolder = new File(rootPath);

        if (!recupFile.exists())          { appendLog("ERROR: Fichier introuvable — " + recupPath); return; }
        if (!rootFolder.isDirectory())    { appendLog("ERROR: Dossier source introuvable — " + rootPath); return; }

        setAllButtonsDisabled(true);
        statusBar.setVisible(false);
        progressBar.setProgress(0);
        logArea.clear();
        lastOutputFile = null;

        executor.submit(() -> {
            try {
                java.util.List<String> log = recupNumFactureService.apply(recupFile, rootFolder,
                    (prog, msg) -> Platform.runLater(() -> { progressBar.setProgress(prog); appendLog(msg); }));
                Platform.runLater(() -> {
                    log.forEach(this::appendLog);
                    statusLabel.setText("Récup. Factures terminée — " + log.size() + " dossier(s).");
                    openFileBtn.setVisible(false);
                    statusBar.setVisible(true);
                    setAllButtonsDisabled(false);
                });
            } catch (Exception ex) {
                Platform.runLater(() -> { appendLog("FATAL: " + ex.getMessage()); setAllButtonsDisabled(false); });
            }
        });
    }

    /**
     * Lance la création des feuilles Info Clients dans un fil de fond.
     * Lit le Listing et écrit une feuille "Infos" dans chaque dossier société.
     */
    private void clientInfo() {
        String listingPath = AppPreferences.getTrfListing();
        String rootPath    = AppPreferences.getMergeRoot();

        if (listingPath.isEmpty()) {
            appendLog("ERROR: Configurez le fichier Listing avant de lancer."); return;
        }
        File listingFile = new File(listingPath);
        File rootFolder  = new File(rootPath);

        if (!listingFile.exists())        { appendLog("ERROR: Fichier introuvable — " + listingPath); return; }
        if (!rootFolder.isDirectory())    { appendLog("ERROR: Dossier source introuvable — " + rootPath); return; }

        setAllButtonsDisabled(true);
        statusBar.setVisible(false);
        progressBar.setProgress(0);
        logArea.clear();
        lastOutputFile = null;

        executor.submit(() -> {
            try {
                java.util.List<String> log = clientInfoService.apply(listingFile, rootFolder,
                    (prog, msg) -> Platform.runLater(() -> { progressBar.setProgress(prog); appendLog(msg); }));
                Platform.runLater(() -> {
                    log.forEach(this::appendLog);
                    statusLabel.setText("Info Clients terminée — " + log.size() + " dossier(s).");
                    openFileBtn.setVisible(false);
                    statusBar.setVisible(true);
                    setAllButtonsDisabled(false);
                });
            } catch (Exception ex) {
                Platform.runLater(() -> { appendLog("FATAL: " + ex.getMessage()); setAllButtonsDisabled(false); });
            }
        });
    }

    private void syncDatabase() {
        File root = new File(AppPreferences.getMergeRoot());
        if (!root.isDirectory()) {
            appendLog("ERROR: Dossier source introuvable. Configurez le chemin.");
            return;
        }
        syncDbBtn.setDisable(true);
        appendLog("Synchronisation DB en cours...");
        executor.submit(() -> {
            try {
                syncService.syncAll(root, (pct, msg) ->
                    Platform.runLater(() -> { progressBar.setProgress(pct); appendLog(msg); }));
            } catch (Exception e) {
                Platform.runLater(() -> appendLog("ERREUR sync : " + e.getMessage()));
            } finally {
                Platform.runLater(() -> {
                    syncDbBtn.setDisable(false);
                    if (dashboardController != null) dashboardController.refresh();
                });
            }
        });
    }

    private void toggleWatch() {
        if (watchService.isRunning()) {
            watchService.stop();
            updateWatchToggleLabel(false);
            appendLog("Surveillance arrêtée.");
            AppPreferences.setWatchEnabled(false);
        } else {
            File root = new File(AppPreferences.getMergeRoot());
            if (!root.isDirectory()) {
                appendLog("ERROR: Dossier source introuvable."); return;
            }
            watchService.start(root);
            updateWatchToggleLabel(true);
            appendLog("Surveillance démarrée : " + root.getAbsolutePath());
            AppPreferences.setWatchEnabled(true);
        }
    }

    private void updateWatchToggleLabel(boolean active) {
        if (watchToggleBtn == null) return;
        Label lName = (Label) ((javafx.scene.layout.VBox) watchToggleBtn.getGraphic()).getChildren().get(0);
        lName.setText(active ? "⏹ Arrêter" : "▶ Surveiller");
    }

    private void onWatchEvent(String companyName, String message) {
        Platform.runLater(() -> {
            appendLog("[WATCH] " + companyName + " — " + message);
            if (message.startsWith("✓") && dashboardController != null) {
                dashboardController.refresh();
            }
        });
    }

    public void shutdown() {
        watchService.stop();
        executor.shutdownNow();
    }

    @FXML
    private void openFile() {
        if (lastOutputFile != null && lastOutputFile.exists()) {
            try { Desktop.getDesktop().open(lastOutputFile); }
            catch (Exception e) { appendLog("Cannot open file: " + e.getMessage()); }
        }
    }

    // =========================================================================
    // Méthodes utilitaires privées
    // =========================================================================

    /**
     * Crée un bouton d'action stylisé avec un nom et une description sur deux lignes.
     *
     * <p>Le bouton utilise un {@link VBox} comme graphique, contenant deux étiquettes :
     * <ul>
     *   <li>{@code lName} — nom principal de l'action, classe CSS {@code action-btn-name}.</li>
     *   <li>{@code lDesc} — description courte, classe CSS {@code action-btn-desc}.</li>
     * </ul>
     *
     * <p>Le bouton est configuré pour s'étirer à la largeur et hauteur maximales disponibles
     * dans la cellule de la grille ({@code setMaxWidth/setMaxHeight(Double.MAX_VALUE)}).
     *
     * @param name       Nom principal affiché en gras (ex. "Générer TRF").
     * @param desc       Description secondaire en italique (ex. "Calcul virements et compensations").
     * @param styleClass Classe CSS appliquée au bouton (ex. {@code "secondary-btn"} ou {@code "run-btn"}).
     * @param handler    Gestionnaire d'événement déclenché au clic.
     * @return Le bouton JavaFX configuré et prêt à être ajouté dans la grille.
     */
    private Button createActionBtn(String name, String desc, String styleClass,
                                    EventHandler<ActionEvent> handler) {
        // Étiquette du nom de l'action
        Label lName = new Label(name);
        lName.getStyleClass().add("action-btn-name");

        // Étiquette de la description de l'action
        Label lDesc = new Label(desc);
        lDesc.getStyleClass().add("action-btn-desc");

        // Conteneur vertical regroupant nom et description
        VBox vb = new VBox(2, lName, lDesc);

        Button btn = new Button();
        btn.setGraphic(vb);
        btn.getStyleClass().add(styleClass);
        btn.setMaxWidth(Double.MAX_VALUE);   // le bouton s'étire en largeur
        btn.setMaxHeight(Double.MAX_VALUE);  // le bouton s'étire en hauteur
        btn.setOnAction(handler);
        return btn;
    }

    /**
     * Met à jour une étiquette de chemin dans la boîte de dialogue de configuration.
     *
     * <p>Si le chemin est vide, affiche "(non configuré)" en rouge.
     * Sinon, affiche le chemin en vert s'il est valide (fichier/dossier existant)
     * ou en rouge s'il est invalide. Les chemins longs (> 60 caractères) sont
     * tronqués par la gauche avec une ellipse ("…").
     *
     * @param lbl    Étiquette JavaFX à mettre à jour.
     * @param path   Chemin absolu du fichier ou dossier.
     * @param isDir  {@code true} si le chemin désigne un dossier, {@code false} pour un fichier.
     */
    private void updatePathLabel(Label lbl, String path, boolean isDir) {
        if (path.isEmpty()) {
            // Chemin non configuré : texte rouge et placeholder
            lbl.setText("(non configuré)");
            lbl.setStyle("-fx-text-fill: #FF4444; -fx-font-family: 'Courier New', monospace;");
        } else {
            // Vérification de l'existence du chemin sur le disque
            boolean exists = isDir ? new File(path).isDirectory() : new File(path).exists();
            // Troncature des chemins trop longs (affichage de la fin, plus reconnaissable)
            String display = path.length() > 60 ? "…" + path.substring(path.length() - 57) : path;
            lbl.setText(display);
            // Couleur verte si valide, rouge si invalide
            lbl.setStyle((exists ? "-fx-text-fill: #1a6b2e;" : "-fx-text-fill: #FF4444;")
                + " -fx-font-family: 'Courier New', monospace; -fx-font-size: 11px;");
        }
    }

    /**
     * Affiche un sélecteur de dossier natif et retourne le dossier choisi.
     *
     * <p>Si {@code lastPath} est un dossier existant, il est utilisé comme
     * répertoire initial du sélecteur pour faciliter la navigation.
     *
     * @param owner    Fenêtre propriétaire du sélecteur (pour le mode modal).
     * @param title    Titre affiché dans la fenêtre du sélecteur.
     * @param lastPath Dernier chemin connu, utilisé comme répertoire initial (peut être vide).
     * @return Le dossier choisi par l'utilisateur, ou {@code null} si annulé.
     */
    private File dialogPickDirectory(Stage owner, String title, String lastPath) {
        DirectoryChooser dc = new DirectoryChooser();
        dc.setTitle(title);
        // Pré-navigation vers le dernier dossier connu si valide
        if (!lastPath.isEmpty()) {
            File f = new File(lastPath);
            if (f.isDirectory()) dc.setInitialDirectory(f);
        }
        return dc.showDialog(owner);
    }

    /**
     * Affiche un sélecteur de fichier natif avec un filtre d'extension et retourne
     * le fichier choisi.
     *
     * <p>Si {@code lastPath} est un fichier existant, le répertoire parent est utilisé
     * comme répertoire initial du sélecteur pour faciliter la navigation.
     *
     * @param owner    Fenêtre propriétaire du sélecteur (pour le mode modal).
     * @param title    Titre affiché dans la fenêtre du sélecteur.
     * @param lastPath Dernier chemin de fichier connu (peut être vide).
     * @param ext      Extension à filtrer (ex. {@code "xlsx"} ou {@code "xls"}),
     *                 ou {@code null} pour n'appliquer aucun filtre.
     * @return Le fichier choisi par l'utilisateur, ou {@code null} si annulé.
     */
    private File dialogPickFile(Stage owner, String title, String lastPath, String ext) {
        FileChooser fc = new FileChooser();
        fc.setTitle(title);
        // Ajout du filtre d'extension si spécifié (ex. "Excel Files — *.xlsx")
        if (ext != null) {
            fc.getExtensionFilters().add(
                new FileChooser.ExtensionFilter("Excel Files", "*." + ext));
        }
        // Pré-navigation vers le dossier parent du dernier fichier connu si valide
        if (!lastPath.isEmpty()) {
            File parent = new File(lastPath).getParentFile();
            if (parent != null && parent.isDirectory()) fc.setInitialDirectory(parent);
        }
        return fc.showOpenDialog(owner);
    }

    /**
     * Active ou désactive simultanément tous les boutons d'action de la grille.
     *
     * <p>Appelée systématiquement en début de tâche ({@code disabled = true}) pour
     * éviter les lancements concurrents, et en fin de tâche ({@code disabled = false})
     * pour rendre l'interface de nouveau opérationnelle. Les vérifications
     * {@code != null} protègent contre un appel prématuré avant {@link #initialize()}.
     *
     * @param disabled {@code true} pour désactiver tous les boutons, {@code false} pour les activer.
     */
    private void setAllButtonsDisabled(boolean disabled) {
        if (trfBtn != null)       trfBtn.setDisable(disabled);
        if (etatBtn != null)      etatBtn.setDisable(disabled);
        if (cmpBtn != null)       cmpBtn.setDisable(disabled);
        if (fixBtn != null)       fixBtn.setDisable(disabled);
        if (controleBtn != null)  controleBtn.setDisable(disabled);
        if (recupBtn != null)     recupBtn.setDisable(disabled);
        if (infoBtn != null)      infoBtn.setDisable(disabled);
        if (syncDbBtn != null)    syncDbBtn.setDisable(disabled);
        if (runActionBtn != null) runActionBtn.setDisable(disabled);
    }

    /**
     * Ajoute un message horodaté dans le journal ({@link #logArea}).
     *
     * <p>L'horodatage est au format {@code HH:mm:ss} (ex. "[14:32:07] Traitement terminé.").
     * Cette méthode doit être appelée depuis le fil JavaFX ; si elle est appelée
     * depuis un fil de fond, il faut l'encapsuler dans {@code Platform.runLater()}.
     *
     * @param message Message à inscrire dans le journal.
     */
    private void appendLog(String message) {
        // Formatage de l'heure courante pour l'horodatage du message
        String ts = LocalTime.now().format(TIME_FMT);
        logArea.appendText("[" + ts + "] " + message + "\n");
    }
}
