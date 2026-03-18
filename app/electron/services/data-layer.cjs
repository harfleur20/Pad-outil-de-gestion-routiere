const fs = require("fs");
const path = require("path");
const Database = require("better-sqlite3");
const XLSX = require("xlsx");

const SHEET_DEFINITIONS = [
  {
    name: "Feuil1",
    title: "Mesures de deflexion",
    description: "Lecture comparateur et deflexion",
    table: "sheet_feuil1",
    columns: ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]
  },
  {
    name: "Feuil2",
    title: "Listing des sections",
    description: "Sections par SAP",
    table: "sheet_feuil2",
    columns: ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"]
  },
  {
    name: "Feuil3",
    title: "Etat chaussee et intervention",
    description: "Observation metier",
    table: "sheet_feuil3",
    columns: ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L"]
  },
  {
    name: "Feuil4",
    title: "Programme d'evaluation",
    description: "Donnees d'entree et regles",
    table: "sheet_feuil4",
    columns: ["A", "B", "C", "D", "E", "F"]
  },
  {
    name: "Feuil5",
    title: "Sections avec assainissement",
    description: "Profil complet de section",
    table: "sheet_feuil5",
    columns: ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"]
  },
  {
    name: "Feuil6",
    title: "Repertoire codifie des voies",
    description: "Noms proposes et bornes",
    table: "sheet_feuil6",
    columns: ["A", "B", "C", "D", "E", "F", "G"]
  },
  {
    name: "Feuil7",
    title: "Degradations et causes",
    description: "Catalogue des degradations",
    table: "sheet_feuil7",
    columns: ["A", "B", "C", "D", "E", "F", "G"]
  }
];

const SHEET_COLUMN_LABELS = {
  Feuil1: {
    A: "PK (lecture)",
    B: "Lecture comparateur - Gauche",
    C: "Lecture comparateur - Axe",
    D: "Lecture comparateur - Droit",
    E: "PK (deflexion)",
    F: "Deflexion - Gauche",
    G: "Deflexion - Axe",
    H: "Deflexion - Droit",
    I: "Deflexion brute moy.",
    J: "Ecart type",
    K: "Deflexion caracteristique Dc"
  },
  Feuil2: {
    A: "N° troncon",
    B: "N° sections",
    C: "Voies",
    D: "Designation",
    E: "Debut",
    F: "Fin",
    G: "Longueur (m)",
    H: "SAP (groupe)",
    I: "SAP",
    J: "Designation SAP",
    K: "Debut SAP",
    L: "Fin SAP"
  },
  Feuil3: {
    A: "Voies",
    B: "Designation",
    C: "Debut",
    D: "Fin",
    E: "Longueur (m)",
    F: "Largeur min. facade (m)",
    G: "Nature du revetement",
    H: "Etat de la chaussee",
    I: "Type caniveaux",
    J: "Description assainissement",
    K: "Largeur min. trottoirs (m)",
    L: "Nature de l'intervention"
  },
  Feuil4: {
    A: "Libelle",
    B: "Valeur entree",
    C: "Evaluation etat",
    D: "Intervention recommandee",
    E: "Zone controle",
    F: "Resultat formule"
  },
  Feuil5: {
    A: "N° troncon",
    B: "N° sections",
    C: "Voies",
    D: "Designation",
    E: "Debut",
    F: "Fin",
    G: "Longueur (m)",
    H: "Largeur min. facade (m)",
    I: "Nature du revetement",
    J: "Etat de la chaussee",
    K: "Type assainissement",
    L: "Etat assainissement",
    M: "Largeur min. trottoirs (m)",
    N: "Stationnement gauche",
    O: "Stationnement droit",
    P: "Stationnement autre"
  },
  Feuil6: {
    A: "N°",
    B: "Type de voie",
    C: "Code voie",
    D: "Lineaire (ml)",
    E: "Nom propose",
    F: "Debut / fin",
    G: "Justification"
  },
  Feuil7: {
    A: "Categorie",
    B: "Reference",
    C: "Degradation",
    D: "Famille",
    E: "Sous-famille",
    F: "Notes",
    G: "Cause probable"
  }
};

const COLUMN_KEYS = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P"];
const DB_COLUMN_KEYS = COLUMN_KEYS.map((key) => `col_${key.toLowerCase()}`);
const DB_COLUMN_MAP = Object.fromEntries(COLUMN_KEYS.map((key, index) => [key, DB_COLUMN_KEYS[index]]));
const DEFAULT_INTERVENTION_TEXT = "a determiner (A D)";
const SOLUTION_TEMPLATES_SEED = [
  {
    templateKey: "REPRISE_SURFACE_CHAUSSEE",
    title: "Reprise de la surface de chaussee",
    description:
      "Delimitation de la zone, scarification eventuelle, application d'un liant d'accrochage, rechargement en enrobe adapte, repandage du 0/4 ou 0/6 puis compactage."
  },
  {
    templateKey: "NIDS_DE_POULE",
    title: "Traitement des nids de poule",
    description:
      "Delimitation, excavation jusqu'au support sec et sain, rebouchage en enrobe adapte, couche d'accrochage, compactage puis impermeabilisation de la zone reparee."
  },
  {
    templateKey: "ORNIERAGE",
    title: "Traitement de l'ornierage",
    description:
      "Reprofilage (ornierage inferieur a 5 cm) ou rechargement (ornierage superieur a 5 cm), avec verification des pentes, compactage intense et drainage final."
  },
  {
    templateKey: "PELADE",
    title: "Traitement de la pelade",
    description:
      "Bouchage local avec enrobe adapte et couche d'accrochage; si degradation generalisee, reprofilage en enrobe a chaud ou enduit superficiel avec compactage."
  },
  {
    templateKey: "RESSUAGE",
    title: "Traitement du ressuage",
    description:
      "Application d'un enduit superficiel ou sablage avec sable grossier, etalement uniforme puis cylindrage leger pour stabiliser la surface."
  }
];

const DEGRADATION_SOLUTION_RULES_SEED = [
  { degradationKey: "REPRISE_DE_LA_SURFACE_DE_CHAUSSEE", templateKey: "REPRISE_SURFACE_CHAUSSEE" },
  { degradationKey: "NIDS_DE_POULE", templateKey: "NIDS_DE_POULE" },
  { degradationKey: "ORNIERAGE", templateKey: "ORNIERAGE" },
  { degradationKey: "ORNIERAGES", templateKey: "ORNIERAGE" },
  { degradationKey: "PELADE", templateKey: "PELADE" },
  { degradationKey: "RESSUAGE", templateKey: "RESSUAGE" }
];

const DEFLECTION_RULES_SEED = [
  { ruleOrder: 1, minValue: null, maxValue: 60, severity: "FAIBLE", recommendation: "PAS D'ENTRETIEN" },
  { ruleOrder: 2, minValue: 60, maxValue: 80, severity: "MOYEN", recommendation: "RENFORCEMENT LEGER" },
  { ruleOrder: 3, minValue: 80, maxValue: 90, severity: "FORT", recommendation: "RENFORCEMENT LOURD" },
  {
    ruleOrder: 4,
    minValue: 90,
    maxValue: null,
    severity: "TRES FORT",
    recommendation: "REHABILITATION COUCHE DE ROULEMENT ET DE BASE"
  }
];

const DEGRADATION_CAUSE_FALLBACKS = {
  ARRACHEMENTS_DU_REVETEMENT:
    "Perte d'adherence ou vieillissement du revetement entrainant le depart progressif des granulats.",
  EPAUFFURES:
    "Degradation des bords ou des joints liee aux chocs, a une mise en oeuvre insuffisante ou a l'infiltration d'eau.",
  CANIVEAUX_OBSTRUES:
    "Accumulation de dechets, sables et boues dans les caniveaux, empechant l'ecoulement normal des eaux."
};

const DRAINAGE_RULES_SEED = [
  {
    ruleOrder: 10,
    matchOperator: "CONTAINS",
    pattern: "OBSTR",
    askRequired: 0,
    needsAttention: 1,
    recommendation: "Curage et nettoyage des caniveaux, puis verification du fonctionnement hydraulique."
  },
  {
    ruleOrder: 20,
    matchOperator: "CONTAINS",
    pattern: "MAUV",
    askRequired: 0,
    needsAttention: 1,
    recommendation: "Curage et nettoyage des caniveaux, puis verification du fonctionnement hydraulique."
  },
  {
    ruleOrder: 30,
    matchOperator: "CONTAINS",
    pattern: "NON FONCTION",
    askRequired: 0,
    needsAttention: 1,
    recommendation: "Curage et nettoyage des caniveaux, puis verification du fonctionnement hydraulique."
  },
  {
    ruleOrder: 40,
    matchOperator: "CONTAINS",
    pattern: "DEFICIENT",
    askRequired: 0,
    needsAttention: 1,
    recommendation: "Curage et nettoyage des caniveaux, puis verification du fonctionnement hydraulique."
  },
  {
    ruleOrder: 50,
    matchOperator: "CONTAINS",
    pattern: "CURAGE",
    askRequired: 0,
    needsAttention: 1,
    recommendation: "Curage et nettoyage des caniveaux, puis verification du fonctionnement hydraulique."
  },
  {
    ruleOrder: 900,
    matchOperator: "ALWAYS",
    pattern: "",
    askRequired: 1,
    needsAttention: 0,
    recommendation: "Verifier l'etat des caniveaux et programmer un entretien preventif si necessaire."
  },
  {
    ruleOrder: 999,
    matchOperator: "ALWAYS",
    pattern: "",
    askRequired: 0,
    needsAttention: 0,
    recommendation: "Aucune alerte assainissement immediate."
  }
];

const BACKUP_TABLES = [
  "app_meta",
  ...SHEET_DEFINITIONS.map((sheet) => sheet.table),
  "sap_sector",
  "road",
  "road_alias",
  "degradation",
  "degradation_cause",
  "road_degradation",
  "road_section",
  "measurement_campaign",
  "road_measurement",
  "decision_profile_input",
  "maintenance_solution_template",
  "degradation_solution_assignment",
  "degradation_solution_override_rel",
  "degradation_definition",
  "deflection_rule",
  "drainage_rule",
  "decision_history",
  "maintenance_intervention"
];

function setupDataLayer({ app }) {
  const dataDir = path.join(app.getPath("userData"), "data");
  fs.mkdirSync(dataDir, { recursive: true });

  const dbPath = path.join(dataDir, "pad-maintenance.db");
  const db = new Database(dbPath);
  db.pragma("journal_mode = WAL");
  db.pragma("foreign_keys = ON");

  preEnsureLegacyColumns(db);
  migrateLegacySchema(db);
  ensureSchema(db);
  repairLegacyRelationalSchema(db);
  seedSolutionCatalog(db);
  seedDeflectionRules(db);
  seedDrainageRules(db);

  if (isDatabaseEmpty(db)) {
    const autoPath = resolveDefaultExcelPath();
    if (autoPath) {
      try {
        importFromExcelInternal(db, autoPath);
      } catch (error) {
        console.error("[PAD] Import initial impossible:", error.message);
      }
    }
  }

  rebuildNormalizedCatalogs(db);
  migrateLegacySolutionMapping(db);

  return {
    dbPath,
    listSheetDefinitions: () => listSheetDefinitions(),
    getDataStatus: () => getDataStatus(db),
    getDataIntegrityReport: () => getDataIntegrityReport(db),
    getDashboardSummary: () => getDashboardSummary(db),
    importFromExcel: (excelPath) => importFromExcelInternal(db, excelPath || resolveDefaultExcelPath()),
    previewExcelImport: (excelPath) => previewExcelImport(excelPath || resolveDefaultExcelPath()),
    exportBackup: (filePath) => exportBackupSnapshot(db, filePath),
    restoreBackup: (filePath) => restoreBackupSnapshot(db, filePath),
    listSheetRows: (sheetName, filters) => listSheetRows(db, sheetName, filters),
    createSheetRow: (sheetName, payload) => createSheetRow(db, sheetName, payload),
    updateSheetRow: (sheetName, rowId, payload) => updateSheetRow(db, sheetName, rowId, payload),
    deleteSheetRow: (sheetName, rowId) => deleteSheetRow(db, sheetName, rowId),
    listRoadCatalog: (filters) => listRoadCatalog(db, filters),
    listRoadSections: (filters) => listRoadSections(db, filters),
    listMeasurementCampaigns: (filters) => listMeasurementCampaigns(db, filters),
    listRoadMeasurements: (filters) => listRoadMeasurements(db, filters),
    upsertMeasurementCampaign: (payload) => upsertMeasurementCampaign(db, payload),
    deleteMeasurementCampaign: (campaignId) => deleteMeasurementCampaign(db, campaignId),
    upsertRoadMeasurement: (payload) => upsertRoadMeasurement(db, payload),
    deleteRoadMeasurement: (measurementId) => deleteRoadMeasurement(db, measurementId),
    listSapSectors: () => listSapSectors(db),
    listDegradationCatalog: () => listDegradationCatalog(db),
    listDrainageRules: () => listDrainageRules(db),
    upsertDrainageRule: (payload) => upsertDrainageRule(db, payload),
    deleteDrainageRule: (ruleId) => deleteDrainageRule(db, ruleId),
    listSolutionTemplates: () => listSolutionTemplates(db),
    upsertSolutionTemplate: (payload) => upsertSolutionTemplate(db, payload),
    assignTemplateToDegradation: (degradationCode, templateKey) =>
      assignTemplateToDegradation(db, degradationCode, templateKey),
    setDegradationSolutionOverride: (degradationCode, solutionText) =>
      setDegradationSolutionOverride(db, degradationCode, solutionText),
    clearDegradationSolutionOverride: (degradationCode) => clearDegradationSolutionOverride(db, degradationCode),
    listMaintenanceInterventions: (filters) => listMaintenanceInterventions(db, filters),
    upsertMaintenanceIntervention: (payload) => upsertMaintenanceIntervention(db, payload),
    deleteMaintenanceIntervention: (interventionId) => deleteMaintenanceIntervention(db, interventionId),
    evaluateDecision: (payload) => evaluateDecision(db, payload),
    listDecisionHistory: (filters) => listDecisionHistory(db, filters),
    clearDecisionHistory: () => clearDecisionHistory(db),
    exportReportWorkbook: (reportType, filePath) => exportReportWorkbook(db, reportType, filePath)
  };
}

function preEnsureLegacyColumns(db) {
  ensureColumnIfMissing(db, "road", "sap_code", "sap_code TEXT");
  ensureColumnIfMissing(db, "road", "road_code", "road_code TEXT NOT NULL DEFAULT ''");
  ensureColumnIfMissing(db, "decision_history", "sap_code", "sap_code TEXT NOT NULL DEFAULT ''");
  ensureColumnIfMissing(db, "degradation", "degradation_code", "degradation_code TEXT NOT NULL DEFAULT ''");
  ensureColumnIfMissing(db, "degradation", "name", "name TEXT NOT NULL DEFAULT ''");
  ensureColumnIfMissing(db, "degradation_cause", "degradation_code", "degradation_code TEXT NOT NULL DEFAULT ''");
  ensureColumnIfMissing(
    db,
    "maintenance_intervention",
    "responsible_name",
    "responsible_name TEXT NOT NULL DEFAULT ''"
  );
  ensureColumnIfMissing(
    db,
    "maintenance_intervention",
    "attachment_path",
    "attachment_path TEXT NOT NULL DEFAULT ''"
  );
  ensureColumnIfMissing(db, "road_degradation", "road_id", "road_id INTEGER");
  ensureColumnIfMissing(db, "deflection_rule", "rule_order", "rule_order INTEGER");
  ensureColumnIfMissing(db, "drainage_rule", "rule_order", "rule_order INTEGER");
  ensureColumnIfMissing(db, "drainage_rule", "is_active", "is_active INTEGER NOT NULL DEFAULT 1");
}

function migrateLegacySchema(db) {
  const columnDefsByTable = {
    sap_sector: [
      "name TEXT NOT NULL DEFAULT ''",
      "sort_order INTEGER NOT NULL DEFAULT 999"
    ],
    road: [
      "road_key TEXT NOT NULL DEFAULT ''",
      "road_code TEXT NOT NULL DEFAULT ''",
      "designation TEXT NOT NULL DEFAULT ''",
      "sap_code TEXT",
      "start_label TEXT NOT NULL DEFAULT ''",
      "end_label TEXT NOT NULL DEFAULT ''",
      "length_m REAL",
      "width_m REAL",
      "surface_type TEXT NOT NULL DEFAULT ''",
      "pavement_state TEXT NOT NULL DEFAULT ''",
      "drainage_type TEXT NOT NULL DEFAULT ''",
      "drainage_state TEXT NOT NULL DEFAULT ''",
      "sidewalk_min_m REAL",
      "parking_left TEXT NOT NULL DEFAULT ''",
      "parking_right TEXT NOT NULL DEFAULT ''",
      "parking_other TEXT NOT NULL DEFAULT ''",
      "itinerary TEXT NOT NULL DEFAULT ''",
      "justification TEXT NOT NULL DEFAULT ''",
      "intervention_hint TEXT NOT NULL DEFAULT ''"
    ],
    degradation: [
      "degradation_code TEXT NOT NULL DEFAULT ''",
      "name TEXT NOT NULL DEFAULT ''"
    ],
    degradation_cause: [
      "degradation_code TEXT NOT NULL DEFAULT ''",
      "cause_text TEXT NOT NULL DEFAULT ''"
    ],
    road_degradation: [
      "is_active INTEGER NOT NULL DEFAULT 1",
      "created_at TEXT NOT NULL DEFAULT ''",
      "updated_at TEXT NOT NULL DEFAULT ''"
    ],
    road_section: [
      "section_key TEXT NOT NULL DEFAULT ''",
      "source_sheet TEXT NOT NULL DEFAULT ''",
      "source_row_no INTEGER NOT NULL DEFAULT 0",
      "troncon_no TEXT NOT NULL DEFAULT ''",
      "section_no TEXT NOT NULL DEFAULT ''",
      "road_key TEXT NOT NULL DEFAULT ''",
      "road_id INTEGER",
      "sap_code TEXT NOT NULL DEFAULT ''",
      "road_code TEXT NOT NULL DEFAULT ''",
      "designation TEXT NOT NULL DEFAULT ''",
      "start_label TEXT NOT NULL DEFAULT ''",
      "end_label TEXT NOT NULL DEFAULT ''",
      "length_m REAL",
      "width_m REAL",
      "surface_type TEXT NOT NULL DEFAULT ''",
      "pavement_state TEXT NOT NULL DEFAULT ''",
      "drainage_type TEXT NOT NULL DEFAULT ''",
      "drainage_state TEXT NOT NULL DEFAULT ''",
      "sidewalk_min_m REAL",
      "intervention_hint TEXT NOT NULL DEFAULT ''",
      "source_payload TEXT NOT NULL DEFAULT ''",
      "is_active INTEGER NOT NULL DEFAULT 1"
    ],
    road_measurement: [
      "measurement_key TEXT NOT NULL DEFAULT ''",
      "source_sheet TEXT NOT NULL DEFAULT ''",
      "source_row_no INTEGER NOT NULL DEFAULT 0",
      "campaign_key TEXT NOT NULL DEFAULT ''",
      "measurement_date TEXT NOT NULL DEFAULT ''",
      "road_id INTEGER",
      "road_key TEXT NOT NULL DEFAULT ''",
      "road_code TEXT NOT NULL DEFAULT ''",
      "designation TEXT NOT NULL DEFAULT ''",
      "start_label TEXT NOT NULL DEFAULT ''",
      "end_label TEXT NOT NULL DEFAULT ''",
      "pk_label TEXT NOT NULL DEFAULT ''",
      "pk_m REAL",
      "lecture_left REAL",
      "lecture_axis REAL",
      "lecture_right REAL",
      "deflection_left REAL",
      "deflection_axis REAL",
      "deflection_right REAL",
      "deflection_avg REAL",
      "std_dev REAL",
      "deflection_dc REAL",
      "source_payload TEXT NOT NULL DEFAULT ''"
    ],
    decision_profile_input: [
      "profile_key TEXT NOT NULL DEFAULT ''",
      "source_sheet TEXT NOT NULL DEFAULT ''",
      "source_row_no INTEGER NOT NULL DEFAULT 0",
      "param_label TEXT NOT NULL DEFAULT ''",
      "param_value TEXT NOT NULL DEFAULT ''",
      "aux_value_1 TEXT NOT NULL DEFAULT ''",
      "aux_value_2 TEXT NOT NULL DEFAULT ''",
      "aux_value_3 TEXT NOT NULL DEFAULT ''",
      "aux_value_4 TEXT NOT NULL DEFAULT ''",
      "source_payload TEXT NOT NULL DEFAULT ''"
    ],
    degradation_definition: [
      "degradation_code TEXT NOT NULL DEFAULT ''",
      "category TEXT NOT NULL DEFAULT ''",
      "reference TEXT NOT NULL DEFAULT ''",
      "family TEXT NOT NULL DEFAULT ''",
      "subfamily TEXT NOT NULL DEFAULT ''",
      "notes TEXT NOT NULL DEFAULT ''"
    ],
    decision_history: [
      "road_id INTEGER",
      "road_code TEXT NOT NULL DEFAULT ''",
      "road_designation TEXT NOT NULL DEFAULT ''",
      "sap_code TEXT NOT NULL DEFAULT ''",
      "start_label TEXT NOT NULL DEFAULT ''",
      "end_label TEXT NOT NULL DEFAULT ''",
      "degradation_name TEXT NOT NULL DEFAULT ''",
      "probable_cause TEXT NOT NULL DEFAULT ''",
      "maintenance_solution TEXT NOT NULL DEFAULT ''",
      "contextual_intervention TEXT NOT NULL DEFAULT ''",
      "deflection_value REAL",
      "deflection_severity TEXT NOT NULL DEFAULT ''",
      "deflection_recommendation TEXT NOT NULL DEFAULT ''",
      "drainage_needs_attention INTEGER NOT NULL DEFAULT 0",
      "drainage_recommendation TEXT NOT NULL DEFAULT ''"
    ],
    drainage_rule: [
      "match_operator TEXT NOT NULL DEFAULT 'CONTAINS'",
      "pattern TEXT NOT NULL DEFAULT ''",
      "ask_required INTEGER NOT NULL DEFAULT 0",
      "needs_attention INTEGER NOT NULL DEFAULT 0",
      "recommendation TEXT NOT NULL DEFAULT ''",
      "is_active INTEGER NOT NULL DEFAULT 1"
    ]
  };

  for (const [tableName, columnDefs] of Object.entries(columnDefsByTable)) {
    for (const columnDef of columnDefs) {
      const columnName = columnDef.split(/\s+/)[0];
      ensureColumnIfMissing(db, tableName, columnName, columnDef);
    }
  }
}

function repairLegacyRelationalSchema(db) {
  if (!tableExists(db, "road_degradation")) {
    return;
  }

  let mismatchDetected = false;
  try {
    db.prepare("INSERT INTO road_degradation (road_id, degradation_code, is_active) VALUES (?, ?, 1)");
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    if (/foreign key mismatch/i.test(message)) {
      mismatchDetected = true;
    } else {
      throw error;
    }
  }

  if (!mismatchDetected) {
    return;
  }

  db.pragma("foreign_keys = OFF");
  db.exec(`
    DROP TABLE IF EXISTS road_degradation;
    DROP TABLE IF EXISTS degradation_cause;
    DROP TABLE IF EXISTS degradation_solution_override_rel;
    DROP TABLE IF EXISTS degradation_solution_assignment;
    DROP TABLE IF EXISTS degradation;
    DROP TABLE IF EXISTS road;
    DROP TABLE IF EXISTS sap_sector;
  `);
  db.pragma("foreign_keys = ON");

  ensureSchema(db);
  migrateLegacySchema(db);
}

function ensureSchema(db) {
  db.exec(`
    CREATE TABLE IF NOT EXISTS app_meta (
      meta_key TEXT PRIMARY KEY,
      meta_value TEXT
    );
  `);

  db.exec(`
    CREATE TABLE IF NOT EXISTS decision_history (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      road_id INTEGER,
      road_code TEXT NOT NULL DEFAULT '',
      road_designation TEXT NOT NULL DEFAULT '',
      sap_code TEXT NOT NULL DEFAULT '',
      start_label TEXT NOT NULL DEFAULT '',
      end_label TEXT NOT NULL DEFAULT '',
      degradation_name TEXT NOT NULL DEFAULT '',
      probable_cause TEXT NOT NULL DEFAULT '',
      maintenance_solution TEXT NOT NULL DEFAULT '',
      contextual_intervention TEXT NOT NULL DEFAULT '',
      deflection_value REAL,
      deflection_severity TEXT NOT NULL DEFAULT '',
      deflection_recommendation TEXT NOT NULL DEFAULT '',
      drainage_needs_attention INTEGER NOT NULL DEFAULT 0,
      drainage_recommendation TEXT NOT NULL DEFAULT ''
    );
    CREATE INDEX IF NOT EXISTS idx_decision_history_created_at ON decision_history(created_at DESC);
    CREATE INDEX IF NOT EXISTS idx_decision_history_sap ON decision_history(sap_code);

    CREATE TABLE IF NOT EXISTS maintenance_intervention (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      road_id INTEGER,
      road_key TEXT NOT NULL DEFAULT '',
      road_code TEXT NOT NULL DEFAULT '',
      road_designation TEXT NOT NULL DEFAULT '',
      sap_code TEXT,
      degradation_code TEXT,
      degradation_name TEXT NOT NULL DEFAULT '',
      intervention_type TEXT NOT NULL DEFAULT '',
      status TEXT NOT NULL DEFAULT 'PREVU',
      intervention_date TEXT NOT NULL DEFAULT '',
      completion_date TEXT NOT NULL DEFAULT '',
      state_before TEXT NOT NULL DEFAULT '',
      state_after TEXT NOT NULL DEFAULT '',
      deflection_before REAL,
      deflection_after REAL,
      solution_applied TEXT NOT NULL DEFAULT '',
      contractor_name TEXT NOT NULL DEFAULT '',
      responsible_name TEXT NOT NULL DEFAULT '',
      attachment_path TEXT NOT NULL DEFAULT '',
      observation TEXT NOT NULL DEFAULT '',
      cost_amount REAL,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (road_id) REFERENCES road(id) ON DELETE SET NULL,
      FOREIGN KEY (degradation_code) REFERENCES degradation(degradation_code) ON DELETE SET NULL,
      FOREIGN KEY (sap_code) REFERENCES sap_sector(code)
    );
    CREATE INDEX IF NOT EXISTS idx_maintenance_intervention_road ON maintenance_intervention(road_id);
    CREATE INDEX IF NOT EXISTS idx_maintenance_intervention_sap ON maintenance_intervention(sap_code);
    CREATE INDEX IF NOT EXISTS idx_maintenance_intervention_status ON maintenance_intervention(status);
    CREATE INDEX IF NOT EXISTS idx_maintenance_intervention_date ON maintenance_intervention(intervention_date DESC);
  `);

  db.exec(`
    CREATE TABLE IF NOT EXISTS sap_sector (
      code TEXT PRIMARY KEY,
      name TEXT NOT NULL DEFAULT '',
      sort_order INTEGER NOT NULL DEFAULT 999
    );

    CREATE TABLE IF NOT EXISTS road (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      road_key TEXT NOT NULL UNIQUE,
      road_code TEXT NOT NULL DEFAULT '',
      designation TEXT NOT NULL DEFAULT '',
      sap_code TEXT,
      start_label TEXT NOT NULL DEFAULT '',
      end_label TEXT NOT NULL DEFAULT '',
      length_m REAL,
      width_m REAL,
      surface_type TEXT NOT NULL DEFAULT '',
      pavement_state TEXT NOT NULL DEFAULT '',
      drainage_type TEXT NOT NULL DEFAULT '',
      drainage_state TEXT NOT NULL DEFAULT '',
      sidewalk_min_m REAL,
      parking_left TEXT NOT NULL DEFAULT '',
      parking_right TEXT NOT NULL DEFAULT '',
      parking_other TEXT NOT NULL DEFAULT '',
      itinerary TEXT NOT NULL DEFAULT '',
      justification TEXT NOT NULL DEFAULT '',
      intervention_hint TEXT NOT NULL DEFAULT '',
      FOREIGN KEY (sap_code) REFERENCES sap_sector(code)
    );
    CREATE INDEX IF NOT EXISTS idx_road_sap ON road(sap_code);
    CREATE INDEX IF NOT EXISTS idx_road_code ON road(road_code);

    CREATE TABLE IF NOT EXISTS road_alias (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      road_id INTEGER,
      road_key TEXT NOT NULL DEFAULT '',
      alias_key TEXT NOT NULL DEFAULT '',
      alias_text TEXT NOT NULL DEFAULT '',
      alias_type TEXT NOT NULL DEFAULT '',
      source_sheet TEXT NOT NULL DEFAULT '',
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (road_id) REFERENCES road(id) ON DELETE CASCADE,
      UNIQUE(road_key, alias_key)
    );
    CREATE INDEX IF NOT EXISTS idx_road_alias_road_id ON road_alias(road_id);
    CREATE INDEX IF NOT EXISTS idx_road_alias_key ON road_alias(alias_key);

    CREATE TABLE IF NOT EXISTS degradation (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      degradation_code TEXT NOT NULL UNIQUE,
      name TEXT NOT NULL DEFAULT ''
    );
    CREATE INDEX IF NOT EXISTS idx_degradation_name ON degradation(name);

    CREATE TABLE IF NOT EXISTS degradation_cause (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      degradation_code TEXT NOT NULL,
      cause_text TEXT NOT NULL DEFAULT '',
      FOREIGN KEY (degradation_code) REFERENCES degradation(degradation_code) ON DELETE CASCADE,
      UNIQUE(degradation_code, cause_text)
    );
    CREATE INDEX IF NOT EXISTS idx_degradation_cause_code ON degradation_cause(degradation_code);

    CREATE TABLE IF NOT EXISTS road_degradation (
      road_id INTEGER NOT NULL,
      degradation_code TEXT NOT NULL,
      is_active INTEGER NOT NULL DEFAULT 1,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      PRIMARY KEY (road_id, degradation_code),
      FOREIGN KEY (road_id) REFERENCES road(id) ON DELETE CASCADE,
      FOREIGN KEY (degradation_code) REFERENCES degradation(degradation_code) ON DELETE CASCADE
    );
    CREATE INDEX IF NOT EXISTS idx_road_degradation_road ON road_degradation(road_id);

    CREATE TABLE IF NOT EXISTS road_section (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      section_key TEXT NOT NULL UNIQUE,
      source_sheet TEXT NOT NULL DEFAULT '',
      source_row_no INTEGER NOT NULL DEFAULT 0,
      troncon_no TEXT NOT NULL DEFAULT '',
      section_no TEXT NOT NULL DEFAULT '',
      road_key TEXT NOT NULL DEFAULT '',
      road_id INTEGER,
      sap_code TEXT NOT NULL DEFAULT '',
      road_code TEXT NOT NULL DEFAULT '',
      designation TEXT NOT NULL DEFAULT '',
      start_label TEXT NOT NULL DEFAULT '',
      end_label TEXT NOT NULL DEFAULT '',
      length_m REAL,
      width_m REAL,
      surface_type TEXT NOT NULL DEFAULT '',
      pavement_state TEXT NOT NULL DEFAULT '',
      drainage_type TEXT NOT NULL DEFAULT '',
      drainage_state TEXT NOT NULL DEFAULT '',
      sidewalk_min_m REAL,
      intervention_hint TEXT NOT NULL DEFAULT '',
      source_payload TEXT NOT NULL DEFAULT '',
      is_active INTEGER NOT NULL DEFAULT 1,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (road_id) REFERENCES road(id) ON DELETE SET NULL,
      FOREIGN KEY (sap_code) REFERENCES sap_sector(code)
    );
    CREATE INDEX IF NOT EXISTS idx_road_section_road_id ON road_section(road_id);
    CREATE INDEX IF NOT EXISTS idx_road_section_sap ON road_section(sap_code);
    CREATE INDEX IF NOT EXISTS idx_road_section_source ON road_section(source_sheet, source_row_no);

    CREATE TABLE IF NOT EXISTS measurement_campaign (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      campaign_key TEXT NOT NULL UNIQUE,
      source_sheet TEXT NOT NULL DEFAULT 'Feuil1',
      source_row_no INTEGER NOT NULL DEFAULT 0,
      road_id INTEGER,
      road_key TEXT NOT NULL DEFAULT '',
      road_code TEXT NOT NULL DEFAULT '',
      designation TEXT NOT NULL DEFAULT '',
      section_label TEXT NOT NULL DEFAULT '',
      start_label TEXT NOT NULL DEFAULT '',
      end_label TEXT NOT NULL DEFAULT '',
      measurement_date TEXT NOT NULL DEFAULT '',
      pk_start_m REAL,
      pk_end_m REAL,
      source_payload TEXT NOT NULL DEFAULT '',
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (road_id) REFERENCES road(id) ON DELETE SET NULL
    );
    CREATE INDEX IF NOT EXISTS idx_measurement_campaign_road_id ON measurement_campaign(road_id);
    CREATE INDEX IF NOT EXISTS idx_measurement_campaign_date ON measurement_campaign(measurement_date);

    CREATE TABLE IF NOT EXISTS road_measurement (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      measurement_key TEXT NOT NULL UNIQUE,
      source_sheet TEXT NOT NULL DEFAULT 'Feuil1',
      source_row_no INTEGER NOT NULL DEFAULT 0,
      campaign_key TEXT NOT NULL DEFAULT '',
      measurement_date TEXT NOT NULL DEFAULT '',
      road_id INTEGER,
      road_key TEXT NOT NULL DEFAULT '',
      road_code TEXT NOT NULL DEFAULT '',
      designation TEXT NOT NULL DEFAULT '',
      start_label TEXT NOT NULL DEFAULT '',
      end_label TEXT NOT NULL DEFAULT '',
      pk_label TEXT NOT NULL DEFAULT '',
      pk_m REAL,
      lecture_left REAL,
      lecture_axis REAL,
      lecture_right REAL,
      deflection_left REAL,
      deflection_axis REAL,
      deflection_right REAL,
      deflection_avg REAL,
      std_dev REAL,
      deflection_dc REAL,
      source_payload TEXT NOT NULL DEFAULT '',
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (road_id) REFERENCES road(id) ON DELETE SET NULL
    );
    CREATE INDEX IF NOT EXISTS idx_road_measurement_road_id ON road_measurement(road_id);
    CREATE INDEX IF NOT EXISTS idx_road_measurement_pk ON road_measurement(pk_m);
    CREATE INDEX IF NOT EXISTS idx_road_measurement_source ON road_measurement(source_sheet, source_row_no);
    CREATE INDEX IF NOT EXISTS idx_road_measurement_campaign ON road_measurement(campaign_key);

    CREATE TABLE IF NOT EXISTS decision_profile_input (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      profile_key TEXT NOT NULL UNIQUE,
      source_sheet TEXT NOT NULL DEFAULT 'Feuil4',
      source_row_no INTEGER NOT NULL DEFAULT 0,
      param_label TEXT NOT NULL DEFAULT '',
      param_value TEXT NOT NULL DEFAULT '',
      aux_value_1 TEXT NOT NULL DEFAULT '',
      aux_value_2 TEXT NOT NULL DEFAULT '',
      aux_value_3 TEXT NOT NULL DEFAULT '',
      aux_value_4 TEXT NOT NULL DEFAULT '',
      source_payload TEXT NOT NULL DEFAULT '',
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    );
    CREATE INDEX IF NOT EXISTS idx_decision_profile_input_source ON decision_profile_input(source_sheet, source_row_no);

    CREATE TABLE IF NOT EXISTS maintenance_solution_template (
      template_key TEXT PRIMARY KEY,
      title TEXT NOT NULL,
      description TEXT NOT NULL,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    );

    CREATE TABLE IF NOT EXISTS degradation_solution_assignment (
      degradation_code TEXT PRIMARY KEY,
      template_key TEXT NOT NULL,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (degradation_code) REFERENCES degradation(degradation_code) ON DELETE CASCADE,
      FOREIGN KEY (template_key) REFERENCES maintenance_solution_template(template_key)
    );

    CREATE TABLE IF NOT EXISTS degradation_solution_override_rel (
      degradation_code TEXT PRIMARY KEY,
      solution_text TEXT NOT NULL,
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (degradation_code) REFERENCES degradation(degradation_code) ON DELETE CASCADE
    );

    CREATE TABLE IF NOT EXISTS degradation_definition (
      degradation_code TEXT PRIMARY KEY,
      category TEXT NOT NULL DEFAULT '',
      reference TEXT NOT NULL DEFAULT '',
      family TEXT NOT NULL DEFAULT '',
      subfamily TEXT NOT NULL DEFAULT '',
      notes TEXT NOT NULL DEFAULT '',
      updated_at TEXT NOT NULL DEFAULT (datetime('now')),
      FOREIGN KEY (degradation_code) REFERENCES degradation(degradation_code) ON DELETE CASCADE
    );
    CREATE INDEX IF NOT EXISTS idx_degradation_definition_category ON degradation_definition(category);

    CREATE TABLE IF NOT EXISTS deflection_rule (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      rule_order INTEGER NOT NULL,
      min_value REAL,
      max_value REAL,
      severity TEXT NOT NULL,
      recommendation TEXT NOT NULL
    );
    CREATE UNIQUE INDEX IF NOT EXISTS idx_deflection_rule_order ON deflection_rule(rule_order);

    CREATE TABLE IF NOT EXISTS drainage_rule (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      rule_order INTEGER NOT NULL,
      match_operator TEXT NOT NULL DEFAULT 'CONTAINS',
      pattern TEXT NOT NULL DEFAULT '',
      ask_required INTEGER NOT NULL DEFAULT 0,
      needs_attention INTEGER NOT NULL DEFAULT 0,
      recommendation TEXT NOT NULL DEFAULT '',
      is_active INTEGER NOT NULL DEFAULT 1
    );
    CREATE UNIQUE INDEX IF NOT EXISTS idx_drainage_rule_order ON drainage_rule(rule_order);
    CREATE INDEX IF NOT EXISTS idx_drainage_rule_active ON drainage_rule(is_active);
  `);

  for (const sheet of SHEET_DEFINITIONS) {
    const columnsSql = DB_COLUMN_KEYS.map((column) => `${column} TEXT NOT NULL DEFAULT ''`).join(",\n      ");
    db.exec(`
      CREATE TABLE IF NOT EXISTS ${sheet.table} (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        row_no INTEGER NOT NULL,
        ${columnsSql},
        created_at TEXT NOT NULL DEFAULT (datetime('now')),
        updated_at TEXT NOT NULL DEFAULT (datetime('now'))
      );
      CREATE INDEX IF NOT EXISTS idx_${sheet.table}_row_no ON ${sheet.table}(row_no);
    `);
  }
}

function seedSolutionCatalog(db) {
  const insertTemplate = db.prepare(`
    INSERT INTO maintenance_solution_template (template_key, title, description)
    VALUES (?, ?, ?)
    ON CONFLICT(template_key) DO NOTHING
  `);

  const tx = db.transaction(() => {
    for (const item of SOLUTION_TEMPLATES_SEED) {
      insertTemplate.run(item.templateKey, item.title, item.description);
    }
  });

  tx();
}

function seedDeflectionRules(db) {
  const insert = db.prepare(`
    INSERT INTO deflection_rule (rule_order, min_value, max_value, severity, recommendation)
    VALUES (?, ?, ?, ?, ?)
    ON CONFLICT(rule_order) DO NOTHING
  `);

  const tx = db.transaction(() => {
    for (const rule of DEFLECTION_RULES_SEED) {
      insert.run(rule.ruleOrder, rule.minValue, rule.maxValue, rule.severity, rule.recommendation);
    }
  });

  tx();
}

function seedDrainageRules(db) {
  const insert = db.prepare(`
    INSERT INTO drainage_rule (
      rule_order, match_operator, pattern, ask_required, needs_attention, recommendation, is_active
    ) VALUES (?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(rule_order) DO NOTHING
  `);

  const tx = db.transaction(() => {
    for (const rule of DRAINAGE_RULES_SEED) {
      insert.run(
        rule.ruleOrder,
        toText(rule.matchOperator) || "CONTAINS",
        toText(rule.pattern),
        Number(rule.askRequired) ? 1 : 0,
        Number(rule.needsAttention) ? 1 : 0,
        toText(rule.recommendation),
        1
      );
    }
  });

  tx();
}

function listDrainageRules(db) {
  return db
    .prepare(
      `
      SELECT
        id,
        rule_order AS ruleOrder,
        match_operator AS matchOperator,
        pattern,
        ask_required AS askRequired,
        needs_attention AS needsAttention,
        recommendation,
        is_active AS isActive
      FROM drainage_rule
      ORDER BY rule_order, id
    `
    )
    .all()
    .map((row) => ({
      ...row,
      askRequired: Boolean(row.askRequired),
      needsAttention: Boolean(row.needsAttention),
      isActive: Boolean(row.isActive)
    }));
}

function upsertDrainageRule(db, payload = {}) {
  const id = Number(payload.id);
  const ruleOrder = Number(payload.ruleOrder);
  const matchOperator = normalizeDrainageMatchOperator(payload.matchOperator);
  const pattern = toText(payload.pattern);
  const askRequired = Number(payload.askRequired) ? 1 : 0;
  const needsAttention = Number(payload.needsAttention) ? 1 : 0;
  const recommendation = toText(payload.recommendation);
  const isActive = payload.isActive === false || payload.isActive === 0 ? 0 : 1;

  if (!Number.isFinite(ruleOrder) || ruleOrder <= 0) {
    throw new Error("ruleOrder invalide.");
  }
  if (!recommendation) {
    throw new Error("La recommandation assainissement est obligatoire.");
  }
  if (matchOperator !== "ALWAYS" && !pattern) {
    throw new Error("Le pattern est obligatoire sauf pour matchOperator=ALWAYS.");
  }

  if (Number.isFinite(id) && id > 0) {
    db.prepare(
      `
      UPDATE drainage_rule
      SET
        rule_order = ?,
        match_operator = ?,
        pattern = ?,
        ask_required = ?,
        needs_attention = ?,
        recommendation = ?,
        is_active = ?
      WHERE id = ?
    `
    ).run(ruleOrder, matchOperator, pattern, askRequired, needsAttention, recommendation, isActive, id);

    const updated = db
      .prepare(
        `
        SELECT
          id,
          rule_order AS ruleOrder,
          match_operator AS matchOperator,
          pattern,
          ask_required AS askRequired,
          needs_attention AS needsAttention,
          recommendation,
          is_active AS isActive
        FROM drainage_rule
        WHERE id = ?
      `
      )
      .get(id);
    if (!updated) {
      throw new Error("Regle assainissement introuvable.");
    }
    return {
      ...updated,
      askRequired: Boolean(updated.askRequired),
      needsAttention: Boolean(updated.needsAttention),
      isActive: Boolean(updated.isActive)
    };
  }

  const result = db
    .prepare(
      `
      INSERT INTO drainage_rule (
        rule_order, match_operator, pattern, ask_required, needs_attention, recommendation, is_active
      ) VALUES (?, ?, ?, ?, ?, ?, ?)
    `
    )
    .run(ruleOrder, matchOperator, pattern, askRequired, needsAttention, recommendation, isActive);

  return (
    listDrainageRules(db).find((item) => item.id === Number(result.lastInsertRowid)) || null
  );
}

function deleteDrainageRule(db, ruleId) {
  const id = Number(ruleId);
  if (!Number.isFinite(id) || id <= 0) {
    throw new Error("ID regle assainissement invalide.");
  }
  db.prepare("DELETE FROM drainage_rule WHERE id = ?").run(id);
  return { deleted: true };
}

function listSolutionTemplates(db) {
  return db
    .prepare(
      `
      SELECT
        template_key AS templateKey,
        title,
        description
      FROM maintenance_solution_template
      ORDER BY title
    `
    )
    .all();
}

function upsertSolutionTemplate(db, payload = {}) {
  const templateKey = normalizeDegradationKey(payload.templateKey);
  const title = toText(payload.title);
  const description = toText(payload.description);

  if (!templateKey) {
    throw new Error("Template de solution invalide.");
  }
  if (!title || !description) {
    throw new Error("Titre et description de solution sont obligatoires.");
  }

  db.prepare(
    `
    INSERT INTO maintenance_solution_template (template_key, title, description)
    VALUES (?, ?, ?)
    ON CONFLICT(template_key) DO UPDATE SET
      title = excluded.title,
      description = excluded.description,
      updated_at = datetime('now')
  `
  ).run(templateKey, title, description);

  return (
    db
      .prepare(
        `
        SELECT
          template_key AS templateKey,
          title,
          description
        FROM maintenance_solution_template
        WHERE template_key = ?
      `
      )
      .get(templateKey) || null
  );
}

function assignTemplateToDegradation(db, degradationCode, templateKeyInput) {
  const degradationKey = normalizeDegradationKey(degradationCode);
  const templateKey = normalizeDegradationKey(templateKeyInput);

  if (!degradationKey || !templateKey) {
    throw new Error("Parametres de liaison degradation/solution invalides.");
  }

  const templateExists = db
    .prepare("SELECT 1 AS ok FROM maintenance_solution_template WHERE template_key = ?")
    .get(templateKey);
  if (!templateExists) {
    throw new Error("Template de solution introuvable.");
  }

  const degradationExists = db
    .prepare("SELECT 1 AS ok FROM degradation WHERE degradation_code = ?")
    .get(degradationKey);
  if (!degradationExists) {
    throw new Error("Degradation introuvable dans le referentiel.");
  }

  db.prepare(
    `
    INSERT INTO degradation_solution_assignment (degradation_code, template_key)
    VALUES (?, ?)
    ON CONFLICT(degradation_code) DO UPDATE SET
      template_key = excluded.template_key,
      updated_at = datetime('now')
  `
  ).run(degradationKey, templateKey);

  return { degradationCode: degradationKey, templateKey };
}

function setDegradationSolutionOverride(db, degradationCode, solutionTextInput) {
  const degradationKey = normalizeDegradationKey(degradationCode);
  const solutionText = toText(solutionTextInput);

  if (!degradationKey) {
    throw new Error("Degradation invalide.");
  }
  if (!solutionText) {
    throw new Error("La solution personnalisee est vide.");
  }

  const degradationExists = db
    .prepare("SELECT 1 AS ok FROM degradation WHERE degradation_code = ?")
    .get(degradationKey);
  if (!degradationExists) {
    throw new Error("Degradation introuvable dans le referentiel.");
  }

  db.prepare(
    `
    INSERT INTO degradation_solution_override_rel (degradation_code, solution_text)
    VALUES (?, ?)
    ON CONFLICT(degradation_code) DO UPDATE SET
      solution_text = excluded.solution_text,
      updated_at = datetime('now')
  `
  ).run(degradationKey, solutionText);

  return { degradationCode: degradationKey, solutionText };
}

function clearDegradationSolutionOverride(db, degradationCode) {
  const degradationKey = normalizeDegradationKey(degradationCode);
  if (!degradationKey) {
    throw new Error("Degradation invalide.");
  }
  db.prepare("DELETE FROM degradation_solution_override_rel WHERE degradation_code = ?").run(degradationKey);
  return { degradationCode: degradationKey, cleared: true };
}

function listSheetDefinitions() {
  return SHEET_DEFINITIONS.map((sheet) => ({
    name: sheet.name,
    title: sheet.title,
    description: sheet.description,
    columns: sheet.columns,
    columnLabels: resolveSheetColumnLabels(sheet)
  }));
}

function resolveSheetColumnLabels(sheet) {
  const configured = SHEET_COLUMN_LABELS[sheet.name] || {};
  const labels = {};
  for (const column of sheet.columns) {
    labels[column] = configured[column] || column;
  }
  return labels;
}

function resolveSheet(sheetName) {
  const found = SHEET_DEFINITIONS.find((sheet) => sheet.name === sheetName);
  if (!found) {
    throw new Error(`Feuille inconnue: ${sheetName}`);
  }
  return found;
}

function rebuildNormalizedCatalogs(db) {
  const tx = db.transaction(() => {
    repairRoadCodesInSourceSheets(db);
    repairSapAssignmentsFromFeuil6(db);
    const roads = buildRoadCatalog(db);
    const roadAliases = buildRoadAliasCatalog(db, roads);
    const degradationItems = buildDegradationCatalog(db);
    const roadSections = buildRoadSectionCatalog(db);
    const measurementCatalog = buildRoadMeasurementCatalog(db);
    const roadMeasurements = measurementCatalog.items;
    const measurementCampaigns = measurementCatalog.campaigns;
    const decisionInputs = buildDecisionProfileInputs(db);
    const degradationDefinitions = buildDegradationDefinitions(db);

    db.prepare("DELETE FROM road_degradation").run();

    const sapCodes = [...new Set(roads.map((road) => toText(road.sapCode)).filter(Boolean))].sort(
      (a, b) => Number((a.match(/[0-9]+/) || ["99"])[0]) - Number((b.match(/[0-9]+/) || ["99"])[0])
    );

    const upsertSap = db.prepare(`
      INSERT INTO sap_sector (code, name, sort_order)
      VALUES (?, ?, ?)
      ON CONFLICT(code) DO UPDATE SET
        name = excluded.name,
        sort_order = excluded.sort_order
    `);
    for (const code of sapCodes) {
      upsertSap.run(code, code, Number((code.match(/[0-9]+/) || ["99"])[0]));
    }

    const upsertRoad = db.prepare(`
      INSERT INTO road (
        road_key, road_code, designation, sap_code, start_label, end_label,
        length_m, width_m, surface_type, pavement_state, drainage_type, drainage_state,
        sidewalk_min_m, parking_left, parking_right, parking_other, itinerary, justification, intervention_hint
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      ON CONFLICT(road_key) DO UPDATE SET
        road_code = excluded.road_code,
        designation = excluded.designation,
        sap_code = excluded.sap_code,
        start_label = excluded.start_label,
        end_label = excluded.end_label,
        length_m = excluded.length_m,
        width_m = excluded.width_m,
        surface_type = excluded.surface_type,
        pavement_state = excluded.pavement_state,
        drainage_type = excluded.drainage_type,
        drainage_state = excluded.drainage_state,
        sidewalk_min_m = excluded.sidewalk_min_m,
        parking_left = excluded.parking_left,
        parking_right = excluded.parking_right,
        parking_other = excluded.parking_other,
        itinerary = excluded.itinerary,
        justification = excluded.justification,
        intervention_hint = excluded.intervention_hint
    `);

    const keepRoadKeys = [];
    for (const road of roads) {
      const roadKey = toText(road.roadKey);
      keepRoadKeys.push(roadKey);
      upsertRoad.run(
        toText(road.roadKey),
        toText(road.roadCode),
        toText(road.designation),
        toText(road.sapCode) || null,
        toText(road.startLabel),
        toText(road.endLabel),
        toNumber(road.lengthM),
        toNumber(road.widthM),
        toText(road.surfaceType),
        toText(road.pavementState),
        toText(road.drainageType),
        toText(road.drainageState),
        toNumber(road.sidewalkMinM),
        toText(road.parkingLeft),
        toText(road.parkingRight),
        toText(road.parkingOther),
        toText(road.itinerary),
        toText(road.justification),
        toText(road.interventionHint)
      );
    }

    if (keepRoadKeys.length > 0) {
      db.prepare(
        `
        DELETE FROM road
        WHERE road_key NOT IN (${keepRoadKeys.map(() => "?").join(", ")})
      `
      ).run(...keepRoadKeys);
    } else {
      db.prepare("DELETE FROM road").run();
    }

    syncRoadAliases(db, roadAliases);

    for (const item of degradationItems) {
      db.prepare(
        `
        INSERT INTO degradation (degradation_code, name)
        VALUES (?, ?)
        ON CONFLICT(degradation_code) DO UPDATE SET
          name = excluded.name
      `
      ).run(item.code, item.name);
    }

    const keepCodes = degradationItems.map((item) => item.code);
    if (keepCodes.length > 0) {
      db.prepare(
        `
        DELETE FROM degradation
        WHERE degradation_code NOT IN (${keepCodes.map(() => "?").join(", ")})
      `
      ).run(...keepCodes);
    } else {
      db.prepare("DELETE FROM degradation").run();
    }

    syncDegradationDefinitions(db, degradationDefinitions, new Set(keepCodes));

    db.prepare("DELETE FROM degradation_cause").run();
    const insertCause = db.prepare(
      "INSERT OR IGNORE INTO degradation_cause (degradation_code, cause_text) VALUES (?, ?)"
    );
    for (const item of degradationItems) {
      for (const cause of item.causes) {
        if (toText(cause)) {
          insertCause.run(item.code, toText(cause));
        }
      }
    }

    seedDefaultDegradationAssignments(db, new Set(keepCodes));

    const roadLookup = buildRoadLookupMaps(db);
    syncRoadSections(db, roadSections, roadLookup);
    syncMeasurementCampaigns(db, measurementCampaigns, roadLookup);
    syncRoadMeasurements(db, roadMeasurements, roadLookup);
    syncDecisionProfileInputs(db, decisionInputs);

    const roadIds = db.prepare("SELECT id FROM road ORDER BY id").all().map((row) => Number(row.id));
    const insertRoadDegradation = db.prepare(
      "INSERT INTO road_degradation (road_id, degradation_code, is_active) VALUES (?, ?, 1)"
    );
    for (const roadId of roadIds) {
      for (const code of keepCodes) {
        insertRoadDegradation.run(roadId, code);
      }
    }

    db.prepare(
      `
      DELETE FROM sap_sector
      WHERE code NOT IN (
        SELECT DISTINCT sap_code
        FROM road
        WHERE sap_code IS NOT NULL AND sap_code <> ''
      )
    `
      ).run();
  });

  tx();
}

function syncRoadAliases(db, aliases) {
  const roadByKey = new Map(
    db
      .prepare(
        `
        SELECT
          id,
          road_key AS roadKey
        FROM road
      `
      )
      .all()
      .map((row) => [toText(row.roadKey), Number(row.id)])
  );

  const upsert = db.prepare(`
    INSERT INTO road_alias (
      road_id,
      road_key,
      alias_key,
      alias_text,
      alias_type,
      source_sheet
    ) VALUES (?, ?, ?, ?, ?, ?)
    ON CONFLICT(road_key, alias_key) DO UPDATE SET
      road_id = excluded.road_id,
      alias_text = excluded.alias_text,
      alias_type = excluded.alias_type,
      source_sheet = excluded.source_sheet,
      updated_at = datetime('now')
  `);

  const keepPairs = [];
  for (const alias of aliases) {
    const roadKey = toText(alias.roadKey);
    const aliasKey = toText(alias.aliasKey);
    if (!roadKey || !aliasKey) {
      continue;
    }

    keepPairs.push({ roadKey, aliasKey });
    const roadId = roadByKey.get(roadKey);

    upsert.run(
      Number.isFinite(roadId) ? roadId : null,
      roadKey,
      aliasKey,
      toText(alias.aliasText),
      toText(alias.aliasType),
      toText(alias.sourceSheet)
    );
  }

  if (keepPairs.length > 0) {
    const conditions = keepPairs.map(() => "(road_key = ? AND alias_key = ?)").join(" OR ");
    const params = keepPairs.flatMap((pair) => [pair.roadKey, pair.aliasKey]);
    db.prepare(`DELETE FROM road_alias WHERE NOT (${conditions})`).run(...params);
  } else {
    db.prepare("DELETE FROM road_alias").run();
  }
}

function buildRoadLookupMaps(db) {
  const rows = db
    .prepare(
      `
      SELECT
        id,
        road_key AS roadKey,
        road_code AS roadCode,
        designation,
        COALESCE(sap_code, '') AS sapCode
      FROM road
    `
    )
    .all();

  const aliasRows = db
    .prepare(
      `
      SELECT
        a.alias_key AS aliasKey,
        r.id,
        r.road_key AS roadKey,
        r.road_code AS roadCode,
        r.designation,
        COALESCE(r.sap_code, '') AS sapCode,
        r.start_label AS startLabel,
        r.end_label AS endLabel
      FROM road_alias a
      INNER JOIN road r ON r.road_key = a.road_key
    `
    )
    .all();

  const byKey = new Map();
  const byCode = new Map();
  const byDesignation = new Map();
  const byAlias = new Map();
  const byBounds = new Map();

  for (const row of rows) {
    const key = toText(row.roadKey);
    const code = normalizeRoadCode(row.roadCode);
    const designation = normalizeText(row.designation);
    const simplifiedDesignation = normalizeText(simplifyRoadDesignation(row.designation));
    const boundsKey = makeRoadBoundsLookupKey(row.designation, row.startLabel, row.endLabel);

    if (key && !byKey.has(key)) {
      byKey.set(key, row);
    }
    setUniqueRoadLookup(byCode, code, row);
    setUniqueRoadLookup(byDesignation, designation, row);
    setUniqueRoadLookup(byDesignation, simplifiedDesignation, row);
    setUniqueRoadLookup(byBounds, boundsKey, row);
  }

  for (const row of aliasRows) {
    const aliasKey = toText(row.aliasKey);
    if (aliasKey.startsWith("BOUNDS:")) {
      setUniqueRoadLookup(byBounds, aliasKey, row);
      continue;
    }
    setUniqueRoadLookup(byAlias, aliasKey, row);
  }

  return { byKey, byCode, byDesignation, byAlias, byBounds };
}

function resolveRoadRef(entry, roadLookup) {
  const roadKey = toText(entry.roadKey);
  if (roadKey && roadLookup.byKey.has(roadKey)) {
    return roadLookup.byKey.get(roadKey);
  }

  const roadCode = normalizeRoadCode(entry.roadCode);
  if (roadCode && roadLookup.byCode.has(roadCode)) {
    return roadLookup.byCode.get(roadCode);
  }

  const boundsKey = makeRoadBoundsLookupKey(entry.designation || entry.sectionLabel, entry.startLabel, entry.endLabel);
  if (boundsKey && roadLookup.byBounds.has(boundsKey)) {
    const road = roadLookup.byBounds.get(boundsKey);
    if (road) {
      return road;
    }
  }

  const aliasCandidates = [entry.sectionLabel, entry.designation, simplifyRoadDesignation(entry.designation)];
  for (const candidate of aliasCandidates) {
    const aliasKey = makeRoadTextAliasKey(candidate);
    if (aliasKey && roadLookup.byAlias.has(aliasKey)) {
      const road = roadLookup.byAlias.get(aliasKey);
      if (road) {
        return road;
      }
    }
  }

  const designation = normalizeText(entry.designation);
  if (designation && roadLookup.byDesignation.has(designation)) {
    return roadLookup.byDesignation.get(designation);
  }

  return null;
}

function syncRoadSections(db, sections, roadLookup) {
  const upsert = db.prepare(`
    INSERT INTO road_section (
      section_key,
      source_sheet,
      source_row_no,
      troncon_no,
      section_no,
      road_key,
      road_id,
      sap_code,
      road_code,
      designation,
      start_label,
      end_label,
      length_m,
      width_m,
      surface_type,
      pavement_state,
      drainage_type,
      drainage_state,
      sidewalk_min_m,
      intervention_hint,
      source_payload,
      is_active
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
    ON CONFLICT(section_key) DO UPDATE SET
      source_sheet = excluded.source_sheet,
      source_row_no = excluded.source_row_no,
      troncon_no = excluded.troncon_no,
      section_no = excluded.section_no,
      road_key = excluded.road_key,
      road_id = excluded.road_id,
      sap_code = excluded.sap_code,
      road_code = excluded.road_code,
      designation = excluded.designation,
      start_label = excluded.start_label,
      end_label = excluded.end_label,
      length_m = excluded.length_m,
      width_m = excluded.width_m,
      surface_type = excluded.surface_type,
      pavement_state = excluded.pavement_state,
      drainage_type = excluded.drainage_type,
      drainage_state = excluded.drainage_state,
      sidewalk_min_m = excluded.sidewalk_min_m,
      intervention_hint = excluded.intervention_hint,
      source_payload = excluded.source_payload,
      is_active = excluded.is_active,
      updated_at = datetime('now')
  `);

  const keepKeys = [];
  for (const section of sections) {
    const roadRef = resolveRoadRef(section, roadLookup);
    const roadId = roadRef ? Number(roadRef.id) : null;
    const sapCode = toText(roadRef?.sapCode) || toText(section.sapCode);
    const sectionNo = alignSectionNoWithSap(section.sectionNo, sapCode);
    const sectionKey = toText(section.sectionKey);
    keepKeys.push(sectionKey);

    upsert.run(
      sectionKey,
      toText(section.sourceSheet),
      Number(section.sourceRowNo) || 0,
      toText(section.tronconNo),
      sectionNo,
      toText(section.roadKey),
      Number.isFinite(roadId) ? roadId : null,
      sapCode || null,
      toText(section.roadCode),
      toText(section.designation),
      toText(section.startLabel),
      toText(section.endLabel),
      toNumber(section.lengthM),
      toNumber(section.widthM),
      toText(section.surfaceType),
      toText(section.pavementState),
      toText(section.drainageType),
      toText(section.drainageState),
      toNumber(section.sidewalkMinM),
      toText(section.interventionHint),
      toText(section.sourcePayload)
    );
  }

  if (keepKeys.length > 0) {
    db.prepare(
      `
      DELETE FROM road_section
      WHERE section_key NOT IN (${keepKeys.map(() => "?").join(", ")})
    `
    ).run(...keepKeys);
  } else {
    db.prepare("DELETE FROM road_section").run();
  }
}

function syncMeasurementCampaigns(db, campaigns, roadLookup) {
  const upsert = db.prepare(`
    INSERT INTO measurement_campaign (
      campaign_key,
      source_sheet,
      source_row_no,
      road_id,
      road_key,
      road_code,
      designation,
      section_label,
      start_label,
      end_label,
      measurement_date,
      pk_start_m,
      pk_end_m,
      source_payload
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(campaign_key) DO UPDATE SET
      source_sheet = excluded.source_sheet,
      source_row_no = excluded.source_row_no,
      road_id = excluded.road_id,
      road_key = excluded.road_key,
      road_code = excluded.road_code,
      designation = excluded.designation,
      section_label = excluded.section_label,
      start_label = excluded.start_label,
      end_label = excluded.end_label,
      measurement_date = excluded.measurement_date,
      pk_start_m = excluded.pk_start_m,
      pk_end_m = excluded.pk_end_m,
      source_payload = excluded.source_payload,
      updated_at = datetime('now')
  `);

  const keepKeys = [];
  for (const campaign of campaigns) {
    const roadRef = resolveRoadRef(campaign, roadLookup);
    const roadId = roadRef ? Number(roadRef.id) : null;
    const campaignKey = toText(campaign.campaignKey);
    keepKeys.push(campaignKey);

    upsert.run(
      campaignKey,
      toText(campaign.sourceSheet) || "Feuil1",
      Number(campaign.sourceRowNo) || 0,
      Number.isFinite(roadId) ? roadId : null,
      toText(campaign.roadKey) || toText(roadRef?.roadKey),
      toText(campaign.roadCode) || toText(roadRef?.roadCode),
      toText(campaign.designation) || toText(roadRef?.designation),
      toText(campaign.sectionLabel),
      toText(campaign.startLabel) || toText(roadRef?.startLabel),
      toText(campaign.endLabel) || toText(roadRef?.endLabel),
      toText(campaign.measurementDate),
      toNumber(campaign.pkStartM),
      toNumber(campaign.pkEndM),
      toText(campaign.sourcePayload)
    );
  }

  if (keepKeys.length > 0) {
    db.prepare(
      `
      DELETE FROM measurement_campaign
      WHERE source_sheet = 'Feuil1'
        AND campaign_key NOT IN (${keepKeys.map(() => "?").join(", ")})
    `
    ).run(...keepKeys);
  } else {
    db.prepare("DELETE FROM measurement_campaign WHERE source_sheet = 'Feuil1'").run();
  }
}

function syncRoadMeasurements(db, measurements, roadLookup) {
  const upsert = db.prepare(`
    INSERT INTO road_measurement (
      measurement_key,
      source_sheet,
      source_row_no,
      campaign_key,
      measurement_date,
      road_id,
      road_key,
      road_code,
      designation,
      start_label,
      end_label,
      pk_label,
      pk_m,
      lecture_left,
      lecture_axis,
      lecture_right,
      deflection_left,
      deflection_axis,
      deflection_right,
      deflection_avg,
      std_dev,
      deflection_dc,
      source_payload
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(measurement_key) DO UPDATE SET
      source_sheet = excluded.source_sheet,
      source_row_no = excluded.source_row_no,
      campaign_key = excluded.campaign_key,
      measurement_date = excluded.measurement_date,
      road_id = excluded.road_id,
      road_key = excluded.road_key,
      road_code = excluded.road_code,
      designation = excluded.designation,
      start_label = excluded.start_label,
      end_label = excluded.end_label,
      pk_label = excluded.pk_label,
      pk_m = excluded.pk_m,
      lecture_left = excluded.lecture_left,
      lecture_axis = excluded.lecture_axis,
      lecture_right = excluded.lecture_right,
      deflection_left = excluded.deflection_left,
      deflection_axis = excluded.deflection_axis,
      deflection_right = excluded.deflection_right,
      deflection_avg = excluded.deflection_avg,
      std_dev = excluded.std_dev,
      deflection_dc = excluded.deflection_dc,
      source_payload = excluded.source_payload,
      updated_at = datetime('now')
  `);

  const keepKeys = [];
  for (const item of measurements) {
    const roadRef = resolveRoadRef(item, roadLookup);
    const roadId = roadRef ? Number(roadRef.id) : null;
    const measurementKey = toText(item.measurementKey);
    keepKeys.push(measurementKey);

    upsert.run(
      measurementKey,
      toText(item.sourceSheet) || "Feuil1",
      Number(item.sourceRowNo) || 0,
      toText(item.campaignKey),
      toText(item.measurementDate),
      Number.isFinite(roadId) ? roadId : null,
      toText(item.roadKey) || toText(roadRef?.roadKey),
      toText(item.roadCode) || toText(roadRef?.roadCode),
      toText(item.designation) || toText(roadRef?.designation),
      toText(item.startLabel) || toText(roadRef?.startLabel),
      toText(item.endLabel) || toText(roadRef?.endLabel),
      toText(item.pkLabel),
      toNumber(item.pkM),
      toNumber(item.lectureLeft),
      toNumber(item.lectureAxis),
      toNumber(item.lectureRight),
      toNumber(item.deflectionLeft),
      toNumber(item.deflectionAxis),
      toNumber(item.deflectionRight),
      toNumber(item.deflectionAvg),
      toNumber(item.stdDev),
      toNumber(item.deflectionDc),
      toText(item.sourcePayload)
    );
  }

  if (keepKeys.length > 0) {
    db.prepare(
      `
      DELETE FROM road_measurement
      WHERE source_sheet = 'Feuil1'
        AND measurement_key NOT IN (${keepKeys.map(() => "?").join(", ")})
    `
    ).run(...keepKeys);
  } else {
    db.prepare("DELETE FROM road_measurement WHERE source_sheet = 'Feuil1'").run();
  }
}

function syncDecisionProfileInputs(db, inputs) {
  db.prepare("DELETE FROM decision_profile_input").run();
  const insert = db.prepare(`
    INSERT INTO decision_profile_input (
      profile_key,
      source_sheet,
      source_row_no,
      param_label,
      param_value,
      aux_value_1,
      aux_value_2,
      aux_value_3,
      aux_value_4,
      source_payload,
      updated_at
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now'))
  `);

  for (const item of inputs) {
    insert.run(
      toText(item.profileKey),
      "Feuil4",
      Number(item.sourceRowNo) || 0,
      toText(item.paramLabel),
      toText(item.paramValue),
      toText(item.auxValue1),
      toText(item.auxValue2),
      toText(item.auxValue3),
      toText(item.auxValue4),
      toText(item.sourcePayload)
    );
  }
}

function syncDegradationDefinitions(db, definitions, allowedCodes) {
  const upsert = db.prepare(`
    INSERT INTO degradation_definition (
      degradation_code,
      category,
      reference,
      family,
      subfamily,
      notes,
      updated_at
    ) VALUES (?, ?, ?, ?, ?, ?, datetime('now'))
    ON CONFLICT(degradation_code) DO UPDATE SET
      category = excluded.category,
      reference = excluded.reference,
      family = excluded.family,
      subfamily = excluded.subfamily,
      notes = excluded.notes,
      updated_at = datetime('now')
  `);

  const keepCodes = [];
  for (const item of definitions) {
    const code = normalizeDegradationKey(item.degradationCode);
    if (!allowedCodes.has(code)) {
      continue;
    }
    keepCodes.push(code);
    upsert.run(
      code,
      toText(item.category),
      toText(item.reference),
      toText(item.family),
      toText(item.subfamily),
      toText(item.notes)
    );
  }

  if (keepCodes.length > 0) {
    db.prepare(
      `
      DELETE FROM degradation_definition
      WHERE degradation_code NOT IN (${keepCodes.map(() => "?").join(", ")})
    `
    ).run(...keepCodes);
  } else {
    db.prepare("DELETE FROM degradation_definition").run();
  }
}

function buildRoadSectionCatalog(db) {
  const rowsFeuil2 = listSheetRows(db, "Feuil2", { limit: 5000 });
  const rowsFeuil3 = listSheetRows(db, "Feuil3", { limit: 5000 });
  const rowsFeuil5 = listSheetRows(db, "Feuil5", { limit: 5000 });

  const sectionsByKey = new Map();
  const sectionNoIndex = new Map();
  const boundsIndex = new Map();

  function buildCanonicalSectionKey(entry) {
    const roadKey = toText(entry.roadKey);
    const sectionNo = normalizeText(entry.sectionNo).replace(/[^A-Z0-9]+/g, "_");
    const boundsKey = makeRoadBoundsLookupKey(entry.designation, entry.startLabel, entry.endLabel);
    const sectionIndexKey = roadKey && sectionNo ? `${roadKey}::SECTION:${sectionNo}` : "";
    const boundsIndexKey = roadKey && boundsKey ? `${roadKey}::${boundsKey}` : "";

    if (sectionIndexKey && sectionNoIndex.has(sectionIndexKey)) {
      return sectionNoIndex.get(sectionIndexKey);
    }
    if (boundsIndexKey && boundsIndex.has(boundsIndexKey)) {
      return boundsIndex.get(boundsIndexKey);
    }
    if (sectionIndexKey) {
      return sectionIndexKey;
    }
    if (roadKey && boundsKey) {
      return `${roadKey}::${boundsKey}`;
    }
    if (roadKey) {
      return `${roadKey}::DERIVED:${toText(entry.sourceSheet)}:${Number(entry.sourceRowNo) || 0}`;
    }
    return `SECTION::${toText(entry.sourceSheet)}:${Number(entry.sourceRowNo) || 0}`;
  }

  function indexSection(entry, canonicalKey) {
    const roadKey = toText(entry.roadKey);
    const sectionNo = normalizeText(entry.sectionNo).replace(/[^A-Z0-9]+/g, "_");
    const boundsKey = makeRoadBoundsLookupKey(entry.designation, entry.startLabel, entry.endLabel);

    if (roadKey && sectionNo) {
      sectionNoIndex.set(`${roadKey}::SECTION:${sectionNo}`, canonicalKey);
    }
    if (roadKey && boundsKey) {
      boundsIndex.set(`${roadKey}::${boundsKey}`, canonicalKey);
    }
  }

  function setIfPresent(target, key, value) {
    if (value === null || value === undefined) {
      return;
    }
    if (typeof value === "number") {
      if (Number.isFinite(value)) {
        target[key] = value;
      }
      return;
    }
    const text = toText(value);
    if (text) {
      target[key] = text;
    }
  }

  function setIfMissing(target, key, value) {
    const current = target[key];
    if (current !== null && current !== undefined && current !== "") {
      return;
    }
    setIfPresent(target, key, value);
  }

  function mergeSectionEntry(entry) {
    const canonicalKey = buildCanonicalSectionKey(entry);
    let current = sectionsByKey.get(canonicalKey);

    if (!current) {
      current = {
        sectionKey: canonicalKey,
        sourceSheet: toText(entry.sourceSheet),
        sourceRowNo: Number(entry.sourceRowNo) || 0,
        tronconNo: toText(entry.tronconNo),
        sectionNo: toText(entry.sectionNo),
        roadKey: toText(entry.roadKey),
        sapCode: toText(entry.sapCode),
        roadCode: toText(entry.roadCode),
        designation: toText(entry.designation),
        startLabel: toText(entry.startLabel),
        endLabel: toText(entry.endLabel),
        lengthM: toNumber(entry.lengthM),
        widthM: toNumber(entry.widthM),
        surfaceType: toText(entry.surfaceType),
        pavementState: toText(entry.pavementState),
        drainageType: toText(entry.drainageType),
        drainageState: toText(entry.drainageState),
        sidewalkMinM: toNumber(entry.sidewalkMinM),
        interventionHint: toText(entry.interventionHint),
        sourcePayload: safeJson({
          primarySource: toText(entry.sourceSheet),
          sources: {
            [toText(entry.sourceSheet)]: {
              rowNo: Number(entry.sourceRowNo) || 0,
              payload: entry.sourcePayload ? JSON.parse(entry.sourcePayload) : {}
            }
          }
        })
      };
      sectionsByKey.set(canonicalKey, current);
      indexSection(entry, canonicalKey);
      return;
    }

    setIfMissing(current, "tronconNo", entry.tronconNo);
    setIfMissing(current, "sectionNo", entry.sectionNo);
    setIfMissing(current, "sapCode", entry.sapCode);
    setIfMissing(current, "roadCode", entry.roadCode);
    setIfMissing(current, "designation", entry.designation);
    setIfMissing(current, "startLabel", entry.startLabel);
    setIfMissing(current, "endLabel", entry.endLabel);
    if (current.lengthM === null) {
      current.lengthM = toNumber(entry.lengthM);
    }

    if (toText(entry.sourceSheet) === "Feuil5") {
      if (current.widthM === null) {
        current.widthM = toNumber(entry.widthM);
      }
      setIfMissing(current, "surfaceType", entry.surfaceType);
      setIfMissing(current, "pavementState", entry.pavementState);
      setIfMissing(current, "drainageType", entry.drainageType);
      setIfMissing(current, "drainageState", entry.drainageState);
      if (current.sidewalkMinM === null) {
        current.sidewalkMinM = toNumber(entry.sidewalkMinM);
      }
    }

    if (toText(entry.sourceSheet) === "Feuil3") {
      const widthM = toNumber(entry.widthM);
      const sidewalkMinM = toNumber(entry.sidewalkMinM);
      if (widthM !== null) {
        current.widthM = widthM;
      }
      setIfPresent(current, "surfaceType", entry.surfaceType);
      setIfPresent(current, "pavementState", entry.pavementState);
      setIfPresent(current, "drainageType", entry.drainageType);
      setIfPresent(current, "drainageState", entry.drainageState);
      if (sidewalkMinM !== null) {
        current.sidewalkMinM = sidewalkMinM;
      }
      setIfPresent(current, "interventionHint", entry.interventionHint);
    }

    try {
      const currentPayload = current.sourcePayload ? JSON.parse(current.sourcePayload) : {};
      const nextSources = {
        ...(currentPayload.sources || {}),
        [toText(entry.sourceSheet)]: {
          rowNo: Number(entry.sourceRowNo) || 0,
          payload: entry.sourcePayload ? JSON.parse(entry.sourcePayload) : {}
        }
      };
      current.sourcePayload = safeJson({
        primarySource: currentPayload.primarySource || current.sourceSheet,
        sources: nextSources
      });
    } catch {
      current.sourcePayload = safeJson({
        primarySource: current.sourceSheet,
        sources: {
          [current.sourceSheet]: {
            rowNo: Number(current.sourceRowNo) || 0
          },
          [toText(entry.sourceSheet)]: {
            rowNo: Number(entry.sourceRowNo) || 0
          }
        }
      });
    }

    indexSection(current, canonicalKey);
  }

  for (const row of rowsFeuil2) {
    const roadCode = canonicalizeRoadCode(row.C);
    const designation = toText(row.D);
    if (!isRoadLabel(roadCode, designation)) {
      continue;
    }
    const roadKey = makeRoadKey(roadCode, designation);
    mergeSectionEntry({
      sectionKey: "",
      sourceSheet: "Feuil2",
      sourceRowNo: Number(row.rowNo) || 0,
      tronconNo: toText(row.A),
      sectionNo: toText(row.B),
      roadKey,
      sapCode: parseSapCode(row.B, row.A),
      roadCode,
      designation,
      startLabel: toText(row.E),
      endLabel: toText(row.F),
      lengthM: toNumber(row.G),
      widthM: null,
      surfaceType: "",
      pavementState: "",
      drainageType: "",
      drainageState: "",
      sidewalkMinM: null,
      interventionHint: "",
      sourcePayload: safeJson(row)
    });
  }

  for (const row of rowsFeuil5) {
    const roadCode = canonicalizeRoadCode(row.C);
    const designation = toText(row.D);
    if (!isRoadLabel(roadCode, designation)) {
      continue;
    }
    const roadKey = makeRoadKey(roadCode, designation);
    mergeSectionEntry({
      sectionKey: "",
      sourceSheet: "Feuil5",
      sourceRowNo: Number(row.rowNo) || 0,
      tronconNo: toText(row.A),
      sectionNo: toText(row.B),
      roadKey,
      sapCode: parseSapCode(row.B, row.A),
      roadCode,
      designation,
      startLabel: toText(row.E),
      endLabel: toText(row.F),
      lengthM: toNumber(row.G),
      widthM: toNumber(row.H),
      surfaceType: toText(row.I),
      pavementState: toText(row.J),
      drainageType: toText(row.K),
      drainageState: toText(row.L),
      sidewalkMinM: toNumber(row.M),
      interventionHint: "",
      sourcePayload: safeJson(row)
    });
  }

  for (const row of rowsFeuil3) {
    const roadCode = canonicalizeRoadCode(row.A);
    const designation = toText(row.B);
    if (!isRoadLabel(roadCode, designation)) {
      continue;
    }
    const roadKey = makeRoadKey(roadCode, designation);
    mergeSectionEntry({
      sectionKey: "",
      sourceSheet: "Feuil3",
      sourceRowNo: Number(row.rowNo) || 0,
      tronconNo: "",
      sectionNo: "",
      roadKey,
      sapCode: "",
      roadCode,
      designation,
      startLabel: toText(row.C),
      endLabel: toText(row.D),
      lengthM: toNumber(row.E),
      widthM: toNumber(row.F),
      surfaceType: toText(row.G),
      pavementState: toText(row.H),
      drainageType: toText(row.I),
      drainageState: toText(row.J),
      sidewalkMinM: toNumber(row.K),
      interventionHint: toText(row.L),
      sourcePayload: safeJson(row)
    });
  }

  return [...sectionsByKey.values()];
}

function buildRoadMeasurementCatalog(db) {
  const rows = listSheetRows(db, "Feuil1", { limit: 10000 });
  const items = [];
  const campaignsByKey = new Map();
  let reportTitle = "";
  let currentContext = {
    sourceRowNo: 0,
    sectionLabel: "",
    designation: "",
    startLabel: "",
    endLabel: "",
    measurementDate: "",
    pkStartM: null,
    pkEndM: null
  };

  for (const row of rows) {
    const titleCandidate = resolveMeasurementReportTitle(row);
    if (titleCandidate) {
      reportTitle = titleCandidate;
    }

    const roadContext = resolveMeasurementRoadContext(row);
    if (roadContext.sectionLabel) {
      currentContext = {
        ...currentContext,
        sourceRowNo: Number(row.rowNo) || 0,
        sectionLabel: roadContext.sectionLabel,
        designation: roadContext.designation,
        startLabel: roadContext.startLabel,
        endLabel: roadContext.endLabel,
        pkStartM: roadContext.pkStartM,
        pkEndM: roadContext.pkEndM
      };
    }

    const measurementDate = extractMeasurementDate(row);
    if (measurementDate) {
      currentContext = {
        ...currentContext,
        measurementDate
      };
    }

    const pkLecture = toText(row.A);
    const pkDeflection = toText(row.E);
    const hasMetric = [row.B, row.C, row.D, row.F, row.G, row.H, row.I, row.J, row.K].some(
      (value) => toNumber(value) !== null
    );
    const hasPk = isLikelyPkValue(pkLecture) || isLikelyPkValue(pkDeflection);

    if (!hasMetric && !hasPk) {
      continue;
    }

    const designation = simplifyRoadDesignation(currentContext.designation);
    const roadKey = designation ? makeRoadKey("", designation) : "";
    const pkM = toNumber(pkDeflection) ?? toNumber(pkLecture);
    const campaignKey = buildMeasurementCampaignKey(currentContext);

    if (campaignKey) {
      if (!campaignsByKey.has(campaignKey)) {
        campaignsByKey.set(campaignKey, {
          campaignKey,
          sourceSheet: "Feuil1",
          sourceRowNo: Number(currentContext.sourceRowNo) || Number(row.rowNo) || 0,
          roadKey,
          roadCode: "",
          designation,
          sectionLabel: toText(currentContext.sectionLabel),
          startLabel: toText(currentContext.startLabel),
          endLabel: toText(currentContext.endLabel),
          measurementDate: toText(currentContext.measurementDate),
          pkStartM: currentContext.pkStartM,
          pkEndM: currentContext.pkEndM,
          sourcePayload: safeJson({
            reportTitle,
            sectionLabel: currentContext.sectionLabel,
            measurementDate: currentContext.measurementDate
          })
        });
      }

      const campaign = campaignsByKey.get(campaignKey);
      if (pkM !== null) {
        if (campaign.pkStartM === null || pkM < campaign.pkStartM) {
          campaign.pkStartM = pkM;
        }
        if (campaign.pkEndM === null || pkM > campaign.pkEndM) {
          campaign.pkEndM = pkM;
        }
      }
    }

    items.push({
      measurementKey: `FEUIL1:${row.rowNo}:${campaignKey || "NO_CAMPAIGN"}`,
      sourceSheet: "Feuil1",
      sourceRowNo: Number(row.rowNo) || 0,
      campaignKey,
      measurementDate: toText(currentContext.measurementDate),
      roadKey,
      roadCode: "",
      designation,
      startLabel: toText(currentContext.startLabel),
      endLabel: toText(currentContext.endLabel),
      sectionLabel: toText(currentContext.sectionLabel),
      pkLabel: pkDeflection || pkLecture,
      pkM,
      lectureLeft: toNumber(row.B),
      lectureAxis: toNumber(row.C),
      lectureRight: toNumber(row.D),
      deflectionLeft: toNumber(row.F),
      deflectionAxis: toNumber(row.G),
      deflectionRight: toNumber(row.H),
      deflectionAvg: toNumber(row.I),
      stdDev: toNumber(row.J),
      deflectionDc: toNumber(row.K),
      sourcePayload: safeJson(row)
    });
  }

  return {
    campaigns: [...campaignsByKey.values()],
    items
  };
}

function resolveMeasurementReportTitle(row) {
  const candidates = [row.A, row.B];
  for (const candidate of candidates) {
    const text = toText(candidate);
    if (/MESURE\s+DES\s+DEFLEXIONS/i.test(text)) {
      return text;
    }
  }
  return "";
}

function resolveMeasurementRoadContext(row) {
  const candidates = [row.B, row.A, row.E];
  for (const candidate of candidates) {
    const text = toText(candidate);
    const designation = simplifyRoadDesignation(text);
    if (!isRoadDesignationText(designation)) {
      continue;
    }

    const bounds = extractMeasurementBounds(text);
    return {
      sectionLabel: text,
      designation,
      startLabel: bounds.startLabel,
      endLabel: bounds.endLabel,
      pkStartM: bounds.pkStartM,
      pkEndM: bounds.pkEndM
    };
  }

  return {
    sectionLabel: "",
    designation: "",
    startLabel: "",
    endLabel: "",
    pkStartM: null,
    pkEndM: null
  };
}

function extractMeasurementBounds(value) {
  const text = toText(value);
  const tronconMatch = text.match(/\bTRON[CÇ]ON\b\s+(.+?)\s*-\s*(.+?)(?:\s+DU\s+PK|\)|$)/i);
  const startLabel = tronconMatch ? toText(tronconMatch[1]) : "";
  const endLabel = tronconMatch ? toText(tronconMatch[2]) : "";
  const pkMatch = text.match(/\bDU\s+PK\b\s*([0-9+.,]+)\s+\bAU\s+PK\b\s*([0-9+.,]+)/i);

  return {
    startLabel,
    endLabel,
    pkStartM: pkMatch ? parsePkDistance(pkMatch[1]) : null,
    pkEndM: pkMatch ? parsePkDistance(pkMatch[2]) : null
  };
}

function extractMeasurementDate(row) {
  const values = Object.values(row).map((value) => toText(value)).filter(Boolean);
  for (const value of values) {
    const match = value.match(/(\d{2})[\/.-](\d{2})[\/.-](\d{4})/);
    if (match) {
      return `${match[3]}-${match[2]}-${match[1]}`;
    }
  }
  return "";
}

function parsePkDistance(value) {
  const text = toText(value).replace(/\s+/g, "");
  const chainageMatch = text.match(/^(\d+)\+(\d+(?:[.,]\d+)?)$/);
  if (chainageMatch) {
    return Number(chainageMatch[1]) * 1000 + Number(String(chainageMatch[2]).replace(",", "."));
  }
  return toNumber(text);
}

function buildMeasurementCampaignKey(context) {
  const designationKey = normalizeRoadAliasValue(context.designation);
  const dateKey = toText(context.measurementDate) || "NO_DATE";
  const startKey = normalizeRoadAliasValue(context.startLabel) || "NO_START";
  const endKey = normalizeRoadAliasValue(context.endLabel) || "NO_END";
  const rowKey = Number(context.sourceRowNo) || 0;

  if (!designationKey) {
    return rowKey ? `FEUIL1:${rowKey}:NO_ROAD:${dateKey}` : "";
  }

  return `FEUIL1:${rowKey}:${designationKey}:${startKey}:${endKey}:${dateKey}`;
}

function buildDecisionProfileInputs(db) {
  const rows = listSheetRows(db, "Feuil4", { limit: 1000 });
  return rows
    .filter((row) => [row.A, row.B, row.C, row.D, row.E, row.F].some((value) => toText(value)))
    .map((row) => ({
      profileKey: `FEUIL4:${row.rowNo}`,
      sourceRowNo: Number(row.rowNo) || 0,
      paramLabel: toText(row.A),
      paramValue: toText(row.B),
      auxValue1: toText(row.C),
      auxValue2: toText(row.D),
      auxValue3: toText(row.E),
      auxValue4: toText(row.F),
      sourcePayload: safeJson(row)
    }));
}

function buildDegradationDefinitions(db) {
  const rows = listSheetRows(db, "Feuil7", { limit: 5000 });
  const map = new Map();

  for (const row of rows) {
    const name = toText(row.C);
    if (!isDegradationLabel(name)) {
      continue;
    }

    const code = normalizeDegradationKey(name);
    if (!map.has(code)) {
      map.set(code, {
        degradationCode: code,
        category: toText(row.A),
        reference: toText(row.B),
        family: toText(row.D),
        subfamily: toText(row.E),
        notes: toText(row.F)
      });
      continue;
    }

    const current = map.get(code);
    current.category = current.category || toText(row.A);
    current.reference = current.reference || toText(row.B);
    current.family = current.family || toText(row.D);
    current.subfamily = current.subfamily || toText(row.E);
    current.notes = current.notes || toText(row.F);
  }

  return [...map.values()];
}

function seedDefaultDegradationAssignments(db, degradationCodes) {
  const insert = db.prepare(`
    INSERT INTO degradation_solution_assignment (degradation_code, template_key)
    VALUES (?, ?)
    ON CONFLICT(degradation_code) DO NOTHING
  `);

  for (const rule of DEGRADATION_SOLUTION_RULES_SEED) {
    if (degradationCodes.has(rule.degradationKey)) {
      insert.run(rule.degradationKey, rule.templateKey);
    }
  }
}

function migrateLegacySolutionMapping(db) {
  if (tableExists(db, "degradation_solution_rule")) {
    db.prepare(
      `
      INSERT OR IGNORE INTO degradation_solution_assignment (degradation_code, template_key)
      SELECT legacy.degradation_key, legacy.template_key
      FROM degradation_solution_rule legacy
      INNER JOIN degradation d ON d.degradation_code = legacy.degradation_key
      INNER JOIN maintenance_solution_template t ON t.template_key = legacy.template_key
    `
    ).run();
  }

  if (tableExists(db, "degradation_solution_override")) {
    db.prepare(
      `
      INSERT OR IGNORE INTO degradation_solution_override_rel (degradation_code, solution_text)
      SELECT legacy.degradation_key, legacy.solution_text
      FROM degradation_solution_override legacy
      INNER JOIN degradation d ON d.degradation_code = legacy.degradation_key
    `
    ).run();
  }
}

function tableExists(db, tableName) {
  return Boolean(
    db.prepare("SELECT 1 AS ok FROM sqlite_master WHERE type = 'table' AND name = ? LIMIT 1").get(tableName)
  );
}

function tableHasColumn(db, tableName, columnName) {
  if (!tableExists(db, tableName)) {
    return false;
  }

  const infoRows = db.prepare(`PRAGMA table_info(${tableName})`).all();
  return infoRows.some((row) => String(row.name || "").toLowerCase() === String(columnName).toLowerCase());
}

function ensureColumnIfMissing(db, tableName, columnName, columnDefinition) {
  if (!tableExists(db, tableName)) {
    return;
  }
  if (tableHasColumn(db, tableName, columnName)) {
    return;
  }
  db.prepare(`ALTER TABLE ${tableName} ADD COLUMN ${columnDefinition}`).run();
}

function ensureMeasurementRuntimeSchema(db) {
  ensureColumnIfMissing(db, "road_measurement", "campaign_key", "campaign_key TEXT NOT NULL DEFAULT ''");
  ensureColumnIfMissing(db, "road_measurement", "measurement_date", "measurement_date TEXT NOT NULL DEFAULT ''");
  ensureColumnIfMissing(db, "road_measurement", "start_label", "start_label TEXT NOT NULL DEFAULT ''");
  ensureColumnIfMissing(db, "road_measurement", "end_label", "end_label TEXT NOT NULL DEFAULT ''");
}

function isDatabaseEmpty(db) {
  for (const sheet of SHEET_DEFINITIONS) {
    const count = db.prepare(`SELECT COUNT(*) AS count FROM ${sheet.table}`).get().count;
    if (count > 0) {
      return false;
    }
  }
  return true;
}

function resolveDefaultExcelPath() {
  const candidates = [
    process.env.PAD_EXCEL_PATH,
    "C:\\Users\\harfl\\OneDrive\\Desktop\\pad\\programme ayissi.xlsx",
    "C:\\Users\\harfl\\OneDrive\\Desktop\\programme ayissi.xlsx"
  ].filter(Boolean);

  for (const candidate of candidates) {
    if (fs.existsSync(candidate)) {
      return candidate;
    }
  }
  return null;
}

function importFromExcelInternal(db, excelPath) {
  if (!excelPath || !fs.existsSync(excelPath)) {
    throw new Error("Fichier Excel introuvable.");
  }

  const workbook = XLSX.readFile(excelPath);
  const missingSheets = SHEET_DEFINITIONS.filter((sheet) => !workbook.Sheets[sheet.name]).map((sheet) => sheet.name);
  if (missingSheets.length > 0) {
    throw new Error(`Fichier Excel invalide: feuilles manquantes (${missingSheets.join(", ")}).`);
  }

  const tx = db.transaction(() => {
    for (const sheet of SHEET_DEFINITIONS) {
      const rows = getSheetRows(workbook, sheet);
      db.prepare(`DELETE FROM ${sheet.table}`).run();
      resetAutoincrementSequence(db, sheet.table);
      const insert = db.prepare(`
        INSERT INTO ${sheet.table} (
          row_no,
          ${DB_COLUMN_KEYS.join(", ")}
        ) VALUES (
          ?,
          ${DB_COLUMN_KEYS.map(() => "?").join(", ")}
        )
      `);

      rows.forEach((row, index) => {
        const values = COLUMN_KEYS.map((key) => toText(row[key]));
        insert.run(index + 1, ...values);
      });
    }

    rebuildNormalizedCatalogs(db);
    setMeta(db, "last_import_path", excelPath);
    setMeta(db, "last_import_at", new Date().toISOString());
  });

  tx();
  return getDataStatus(db);
}

function listSheetRows(db, sheetName, filters = {}) {
  const sheet = resolveSheet(sheetName);
  const where = [];
  const params = {};

  if (filters.search) {
    where.push(`(${DB_COLUMN_KEYS.map((key) => `${key} LIKE @search`).join(" OR ")})`);
    params.search = `%${String(filters.search).trim()}%`;
  }

  const limit = Math.min(Math.max(Number(filters.limit) || 2000, 1), 5000);
  params.limit = limit;

  const selectedColumns = COLUMN_KEYS.map((key, index) => `${DB_COLUMN_KEYS[index]} AS ${key}`).join(", ");
  const sql = `
    SELECT
      id,
      row_no AS rowNo,
      ${selectedColumns}
    FROM ${sheet.table}
    ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
    ORDER BY row_no, id
    LIMIT @limit
  `;

  return db.prepare(sql).all(params);
}

function getRequiredSheetColumns(sheetName) {
  if (sheetName === "Feuil2") {
    return ["A", "B", "C"];
  }
  if (sheetName === "Feuil3") {
    return ["A", "F", "G", "H", "I", "J", "L"];
  }
  if (sheetName === "Feuil5") {
    return ["C", "H", "K", "L", "M"];
  }
  if (sheetName === "Feuil6") {
    return ["B", "C", "D", "E", "F", "G"];
  }
  if (sheetName === "Feuil7") {
    return ["A", "B", "C", "G"];
  }
  if (sheetName === "Feuil4") {
    return ["A"];
  }
  return [];
}

function getSheetFieldRequiredMessage(sheetName, column) {
  if (sheetName === "Feuil2") {
    if (column === "A") return "Veuillez renseigner le numéro du tronçon.";
    if (column === "B") return "Veuillez renseigner le numéro de section.";
    if (column === "C") return "Veuillez choisir la voie concernée.";
  }
  if (sheetName === "Feuil3") {
    if (column === "A") return "Veuillez choisir la voie à diagnostiquer.";
    if (column === "F") return "Veuillez renseigner la largeur minimale côté façade.";
    if (column === "G") return "Veuillez renseigner le type de revêtement.";
    if (column === "H") return "Veuillez renseigner l'état de la chaussée.";
    if (column === "I") return "Veuillez renseigner le type de caniveaux.";
    if (column === "J") return "Veuillez renseigner l'état de l'assainissement.";
    if (column === "L") return "Veuillez renseigner l'intervention à prévoir.";
  }
  if (sheetName === "Feuil5") {
    if (column === "C") return "Veuillez choisir la voie à compléter.";
    if (column === "H") return "Veuillez renseigner la largeur minimale côté façade.";
    if (column === "K") return "Veuillez renseigner le type d'assainissement.";
    if (column === "L") return "Veuillez renseigner l'état de l'assainissement.";
    if (column === "M") return "Veuillez renseigner la largeur minimale des trottoirs.";
  }
  if (sheetName === "Feuil6") {
    if (column === "B") return "Veuillez renseigner le type de voie.";
    if (column === "C") return "Veuillez renseigner le code de la voie.";
    if (column === "D") return "Veuillez renseigner le linéaire en mètres.";
    if (column === "E") return "Veuillez renseigner le nom proposé.";
    if (column === "F") return "Veuillez renseigner le début et la fin de la voie.";
    if (column === "G") return "Veuillez renseigner la justification.";
  }
  if (sheetName === "Feuil7") {
    if (column === "A") return "Veuillez renseigner la catégorie.";
    if (column === "B") return "Veuillez renseigner la référence.";
    if (column === "C") return "Veuillez renseigner le nom de la dégradation.";
    if (column === "G") return "Veuillez renseigner la cause probable.";
  }
  if (sheetName === "Feuil4" && column === "A") {
    return "Veuillez renseigner le libellé.";
  }
  const label = SHEET_COLUMN_LABELS?.[sheetName]?.[column] || column;
  return `Veuillez renseigner ${label}.`;
}

function parseRoadSectionSourcePayload(sourcePayload) {
  try {
    const parsed = JSON.parse(sourcePayload || "{}");
    return parsed && typeof parsed === "object" ? parsed : {};
  } catch {
    return {};
  }
}

function getRoadSectionSourceRowNo(section, sheetName) {
  const parsed = parseRoadSectionSourcePayload(section.sourcePayload);
  const rowNo = Number(parsed?.sources?.[sheetName]?.rowNo);
  if (Number.isFinite(rowNo) && rowNo > 0) {
    return rowNo;
  }
  if (section.sourceSheet === sheetName && Number(section.sourceRowNo) > 0) {
    return Number(section.sourceRowNo);
  }
  return 0;
}

function normalizeRoadCompareKey(value) {
  return toText(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/BOULEVARD/g, "BVD")
    .replace(/AVENUE/g, "AV")
    .replace(/[().,;:_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function buildMergedSheetCells(current = null, payload = {}) {
  const cells = {};
  for (const key of COLUMN_KEYS) {
    if (Object.prototype.hasOwnProperty.call(payload, key)) {
      cells[key] = toText(payload[key]);
    } else {
      cells[key] = current ? toText(current[key]) : "";
    }
  }
  return cells;
}

function resolveRoadFromSheetCells(roads, sheetName, cells) {
  if (!Array.isArray(roads) || roads.length === 0) {
    return null;
  }

  if (sheetName === "Feuil2" || sheetName === "Feuil5") {
    const codeKey = normalizeRoadCompareKey(cells.C);
    const designationKey = normalizeRoadCompareKey(cells.D);
    const startKey = normalizeRoadCompareKey(cells.E);
    const endKey = normalizeRoadCompareKey(cells.F);
    const sapKey = normalizeRoadCompareKey(parseSapCode(cells.B, cells.A));

    return (
      roads.find((road) => {
        const roadCode = normalizeRoadCompareKey(road.roadCode);
        const roadDesignation = normalizeRoadCompareKey(road.designation);
        const roadStart = normalizeRoadCompareKey(road.startLabel);
        const roadEnd = normalizeRoadCompareKey(road.endLabel);
        const roadSap = normalizeRoadCompareKey(road.sapCode);

        if (codeKey && roadCode === codeKey) {
          return true;
        }
        if (designationKey && roadDesignation === designationKey && (!sapKey || roadSap === sapKey)) {
          return true;
        }
        return Boolean(
          designationKey &&
            startKey &&
            endKey &&
            roadDesignation === designationKey &&
            roadStart === startKey &&
            roadEnd === endKey
        );
      }) || null
    );
  }

  if (sheetName === "Feuil3") {
    const codeKey = normalizeRoadCompareKey(cells.A);
    const designationKey = normalizeRoadCompareKey(cells.B);
    const startKey = normalizeRoadCompareKey(cells.C);
    const endKey = normalizeRoadCompareKey(cells.D);

    return (
      roads.find((road) => {
        const roadCode = normalizeRoadCompareKey(road.roadCode);
        const roadDesignation = normalizeRoadCompareKey(road.designation);
        const roadStart = normalizeRoadCompareKey(road.startLabel);
        const roadEnd = normalizeRoadCompareKey(road.endLabel);

        if (codeKey && roadCode === codeKey) {
          return true;
        }
        if (designationKey && roadDesignation === designationKey && startKey && endKey) {
          return roadStart === startKey && roadEnd === endKey;
        }
        return Boolean(designationKey && roadDesignation === designationKey);
      }) || null
    );
  }

  if (sheetName === "Feuil6") {
    const codeKey = normalizeRoadCompareKey(cells.C);
    const designationKey = normalizeRoadCompareKey(cells.E);
    const bounds = splitItinerary(cells.F);
    const startKey = normalizeRoadCompareKey(bounds.startLabel);
    const endKey = normalizeRoadCompareKey(bounds.endLabel);

    return (
      roads.find((road) => {
        const roadCode = normalizeRoadCompareKey(road.roadCode);
        const roadDesignation = normalizeRoadCompareKey(road.designation);
        const roadStart = normalizeRoadCompareKey(road.startLabel);
        const roadEnd = normalizeRoadCompareKey(road.endLabel);

        if (codeKey && roadCode === codeKey) {
          return true;
        }
        if (designationKey && roadDesignation === designationKey) {
          return true;
        }
        return Boolean(startKey && endKey && roadStart === startKey && roadEnd === endKey);
      }) || null
    );
  }

  return null;
}

function validateSheetRowPayload(db, sheetName, cells, options = {}) {
  const currentRowId = Number(options.rowId);
  const currentRowNo = Number(options.rowNo);
  const currentRow = options.currentRow || null;

  for (const column of getRequiredSheetColumns(sheetName)) {
    if (!toText(cells[column])) {
      throw new Error(getSheetFieldRequiredMessage(sheetName, column));
    }
  }

  if (sheetName === "Feuil2") {
    if (toText(cells.A) && !/^[1-9][0-9]*$/.test(toText(cells.A))) {
      throw new Error("Le numéro du tronçon doit être un nombre entier positif.");
    }
    if (toText(cells.B) && !/^[1-9][0-9]*_[1-9][0-9]*$/.test(toText(cells.B))) {
      throw new Error("Écrivez par exemple 1_1, 2_3 ou 6_1. Le nombre avant _ crée le groupe SAP.");
    }
    const lengthM = toNumber(cells.G);
    if (toText(cells.G) && (!Number.isFinite(lengthM) || Number(lengthM) <= 0)) {
      throw new Error("La longueur doit être un nombre supérieur à 0.");
    }
  }

  if (sheetName === "Feuil3") {
    const lengthM = toNumber(cells.E);
    const facadeWidthM = toNumber(cells.F);
    const sidewalkWidthM = toNumber(cells.K);
    if (toText(cells.E) && (!Number.isFinite(lengthM) || Number(lengthM) <= 0)) {
      throw new Error("La longueur doit être un nombre supérieur à 0.");
    }
    if (toText(cells.F) && (!Number.isFinite(facadeWidthM) || Number(facadeWidthM) <= 0)) {
      throw new Error("La largeur côté façade doit être un nombre supérieur à 0.");
    }
    if (toText(cells.K) && (!Number.isFinite(sidewalkWidthM) || Number(sidewalkWidthM) < 0)) {
      throw new Error("La largeur des trottoirs doit être un nombre positif ou nul.");
    }
  }

  if (sheetName === "Feuil5") {
    const numericColumns = [
      ["H", "La largeur côté façade doit être un nombre supérieur à 0.", true],
      ["M", "La largeur des trottoirs doit être un nombre positif ou nul.", false],
      ["N", "La valeur du stationnement à gauche doit être un nombre positif ou nul.", false],
      ["O", "La valeur du stationnement à droite doit être un nombre positif ou nul.", false]
    ];

    for (const [column, message, strictPositive] of numericColumns) {
      const rawValue = toText(cells[column]);
      const numericValue = toNumber(cells[column]);
      if (!rawValue) {
        continue;
      }
      if (!Number.isFinite(numericValue) || (strictPositive ? Number(numericValue) <= 0 : Number(numericValue) < 0)) {
        throw new Error(message);
      }
    }
  }

  if (sheetName === "Feuil6") {
    const linearM = toNumber(cells.D);
    if (toText(cells.D) && (!Number.isFinite(linearM) || Number(linearM) <= 0)) {
      throw new Error("Le linéaire doit être un nombre supérieur à 0.");
    }
    const bounds = splitItinerary(cells.F);
    if (!bounds.startLabel || !bounds.endLabel) {
      throw new Error("Veuillez renseigner le début et la fin de la voie.");
    }
  }

  if (sheetName === "Feuil4") {
    const hasValue = ["B", "C", "D", "E", "F"].some((column) => toText(cells[column]));
    if (!hasValue) {
      throw new Error("Veuillez remplir au moins une valeur utile sur cette ligne.");
    }
  }

  const roads = ["Feuil2", "Feuil3", "Feuil5", "Feuil6"].includes(sheetName) ? listRoadCatalog(db, {}) : [];
  const sections = ["Feuil2", "Feuil3", "Feuil5"].includes(sheetName) ? listRoadSections(db, {}) : [];
  const draftRoad = resolveRoadFromSheetCells(roads, sheetName, cells);
  const currentRoad = currentRow ? resolveRoadFromSheetCells(roads, sheetName, currentRow) : null;
  const editingRoadId = currentRoad?.id || 0;

  if ((sheetName === "Feuil2" || sheetName === "Feuil3" || sheetName === "Feuil5") && !draftRoad) {
    throw new Error("La voie choisie doit déjà exister dans le référentiel central des voies.");
  }

  if (sheetName === "Feuil2" || sheetName === "Feuil5") {
    const sectionNo = normalizeText(cells.B);
    const startKey = normalizeText(cells.E);
    const endKey = normalizeText(cells.F);
    const roadCodeKey = normalizeRoadCode(cells.C);

    const duplicateSection = sections.find((section) => {
      const sheetRowNo = getRoadSectionSourceRowNo(section, sheetName);
      if (!sheetRowNo || (currentRowNo && sheetRowNo === currentRowNo)) {
        return false;
      }

      const sameRoad = draftRoad
        ? (section.roadId ? Number(section.roadId) === Number(draftRoad.id) : normalizeText(section.roadKey) === normalizeText(draftRoad.roadKey))
        : normalizeRoadCode(section.roadCode) === roadCodeKey;

      if (!sameRoad) {
        return false;
      }
      if (sectionNo && normalizeText(section.sectionNo) === sectionNo) {
        return true;
      }
      return Boolean(startKey && endKey && normalizeText(section.startLabel) === startKey && normalizeText(section.endLabel) === endKey);
    });

    if (duplicateSection) {
      throw new Error(`Cette section existe déjà pour cette voie (${toText(duplicateSection.sectionNo) || "-"}).`);
    }
  }

  if (sheetName === "Feuil3") {
    const duplicateSection = sections.find((section) => {
      const sheetRowNo = getRoadSectionSourceRowNo(section, "Feuil3");
      if (!sheetRowNo || (currentRowNo && sheetRowNo === currentRowNo)) {
        return false;
      }
      if (draftRoad) {
        return section.roadId ? Number(section.roadId) === Number(draftRoad.id) : normalizeText(section.roadKey) === normalizeText(draftRoad.roadKey);
      }
      return (
        normalizeRoadCode(section.roadCode) === normalizeRoadCode(cells.A) ||
        normalizeRoadCompareKey(section.designation) === normalizeRoadCompareKey(cells.B)
      );
    });

    if (duplicateSection) {
      throw new Error(`Cette voie possède déjà un profil dans cette feuille (${toText(duplicateSection.roadCode) || "-"}).`);
    }
  }

  if (sheetName === "Feuil6") {
    const codeKey = normalizeRoadCode(cells.C);
    const designationKey = normalizeRoadCompareKey(cells.E);
    const bounds = splitItinerary(cells.F);
    const startKey = normalizeRoadCompareKey(bounds.startLabel);
    const endKey = normalizeRoadCompareKey(bounds.endLabel);

    const duplicateByCode = roads.find(
      (road) => codeKey && normalizeRoadCode(road.roadCode) === codeKey && (!editingRoadId || Number(road.id) !== Number(editingRoadId))
    );
    if (duplicateByCode) {
      throw new Error(`Ce code de voie existe déjà dans le répertoire (${toText(duplicateByCode.roadCode) || "-"}).`);
    }

    const duplicateByDesignation = roads.find(
      (road) =>
        designationKey &&
        normalizeRoadCompareKey(road.designation) === designationKey &&
        (!editingRoadId || Number(road.id) !== Number(editingRoadId))
    );
    if (duplicateByDesignation) {
      throw new Error(`Cette voie existe déjà dans le répertoire (${toText(duplicateByDesignation.designation) || "-"}).`);
    }

    const duplicateByBounds = roads.find(
      (road) =>
        startKey &&
        endKey &&
        normalizeRoadCompareKey(road.startLabel) === startKey &&
        normalizeRoadCompareKey(road.endLabel) === endKey &&
        (!editingRoadId || Number(road.id) !== Number(editingRoadId))
    );
    if (duplicateByBounds) {
      throw new Error(`Cette combinaison début / fin existe déjà dans le répertoire (${toText(duplicateByBounds.designation) || "-"}).`);
    }
  }

  if (sheetName === "Feuil7" || sheetName === "Feuil4") {
    const comparableRows = listSheetRows(db, sheetName, { limit: 5000 }).filter((row) => !currentRowId || Number(row.id) !== currentRowId);

    if (sheetName === "Feuil7") {
      const referenceKey = normalizeText(cells.B);
      const degradationKey = normalizeRoadCompareKey(cells.C);
      const duplicate = comparableRows.find(
        (row) =>
          (referenceKey && normalizeText(row.B) === referenceKey) ||
          (degradationKey && normalizeRoadCompareKey(row.C) === degradationKey)
      );
      if (duplicate) {
        if (referenceKey && normalizeText(duplicate.B) === referenceKey) {
          throw new Error(`Cette référence existe déjà (${toText(duplicate.B) || "-"}).`);
        }
        throw new Error(`Cette dégradation existe déjà (${toText(duplicate.C) || "-"}).`);
      }
    }

    if (sheetName === "Feuil4") {
      const labelKey = normalizeRoadCompareKey(cells.A);
      const duplicate = comparableRows.find((row) => labelKey && normalizeRoadCompareKey(row.A) === labelKey);
      if (duplicate) {
        throw new Error(`Cette ligne existe déjà dans le programme d'évaluation (${toText(duplicate.A) || "-"}).`);
      }
    }
  }
}

function createSheetRow(db, sheetName, payload = {}) {
  const sheet = resolveSheet(sheetName);
  const explicitRowNo = Number(payload.rowNo);
  const maxRowNo = db.prepare(`SELECT COALESCE(MAX(row_no), 0) AS maxRowNo FROM ${sheet.table}`).get().maxRowNo;
  const rowNo = Number.isFinite(explicitRowNo) && explicitRowNo > 0 ? explicitRowNo : maxRowNo + 1;
  const cells = buildMergedSheetCells(null, payload);
  validateSheetRowPayload(db, sheet.name, cells, { rowNo });
  const values = COLUMN_KEYS.map((key) => cells[key]);

  const insert = db.prepare(`
    INSERT INTO ${sheet.table} (
      row_no,
      ${DB_COLUMN_KEYS.join(", ")}
    ) VALUES (
      ?,
      ${DB_COLUMN_KEYS.map(() => "?").join(", ")}
    )
  `);

  const tx = db.transaction(() => {
    const insertedId = insert.run(rowNo, ...values).lastInsertRowid;
    rebuildNormalizedCatalogs(db);
    return getSheetRowById(db, sheet, insertedId);
  });

  return tx();
}

function updateSheetRow(db, sheetName, rowId, payload = {}) {
  const sheet = resolveSheet(sheetName);
  const id = Number(rowId);
  if (!Number.isFinite(id) || id <= 0) {
    throw new Error("ID de ligne invalide.");
  }

  const current = getSheetRowById(db, sheet, id);
  if (!current) {
    throw new Error("Ligne introuvable.");
  }

  const explicitRowNo = Number(payload.rowNo);
  const rowNo = Number.isFinite(explicitRowNo) && explicitRowNo > 0 ? explicitRowNo : current.rowNo;
  const cells = buildMergedSheetCells(current, payload);
  validateSheetRowPayload(db, sheet.name, cells, {
    rowId: id,
    rowNo,
    currentRow: current
  });
  const values = COLUMN_KEYS.map((key) => cells[key]);

  const update = db.prepare(`
    UPDATE ${sheet.table}
    SET
      row_no = ?,
      ${DB_COLUMN_KEYS.map((key) => `${key} = ?`).join(",\n      ")},
      updated_at = datetime('now')
    WHERE id = ?
  `);

  const tx = db.transaction(() => {
    update.run(rowNo, ...values, id);
    rebuildNormalizedCatalogs(db);
    return getSheetRowById(db, sheet, id);
  });

  return tx();
}

function deleteSheetRow(db, sheetName, rowId) {
  const sheet = resolveSheet(sheetName);
  const id = Number(rowId);
  if (!Number.isFinite(id) || id <= 0) {
    throw new Error("ID de ligne invalide.");
  }
  db.prepare(`DELETE FROM ${sheet.table} WHERE id = ?`).run(id);
  rebuildNormalizedCatalogs(db);
  return true;
}

function getSheetRowById(db, sheet, id) {
  const selectedColumns = COLUMN_KEYS.map((key, index) => `${DB_COLUMN_KEYS[index]} AS ${key}`).join(", ");
  return (
    db
      .prepare(`
        SELECT
          id,
          row_no AS rowNo,
          ${selectedColumns}
        FROM ${sheet.table}
        WHERE id = ?
      `)
      .get(id) || null
  );
}

function getDataStatus(db) {
  const sheetCounts = {};
  let totalRows = 0;

  for (const sheet of SHEET_DEFINITIONS) {
    const count = db.prepare(`SELECT COUNT(*) AS count FROM ${sheet.table}`).get().count;
    sheetCounts[sheet.name] = count;
    totalRows += count;
  }

  const lastImportPath = db.prepare("SELECT meta_value FROM app_meta WHERE meta_key = 'last_import_path'").get()?.meta_value || null;
  const lastImportAt = db.prepare("SELECT meta_value FROM app_meta WHERE meta_key = 'last_import_at'").get()?.meta_value || null;
  const decisionHistoryCount = db.prepare("SELECT COUNT(*) AS count FROM decision_history").get().count;

  return {
    sheetCounts,
    totalRows,
    decisionHistoryCount,
    lastImportPath,
    lastImportAt
  };
}

function getDataIntegrityReport(db) {
  const scalar = (sql, ...params) => {
    const row = db.prepare(sql).get(...params);
    const value = row ? Object.values(row)[0] : 0;
    return Number(value) || 0;
  };

  const totals = {
    roads: scalar("SELECT COUNT(*) AS count FROM road"),
    degradations: scalar("SELECT COUNT(*) AS count FROM degradation"),
    roadSections: scalar("SELECT COUNT(*) AS count FROM road_section"),
    roadMeasurements: scalar("SELECT COUNT(*) AS count FROM road_measurement"),
    profileInputs: scalar("SELECT COUNT(*) AS count FROM decision_profile_input"),
    decisionHistory: scalar("SELECT COUNT(*) AS count FROM decision_history")
  };

  const issues = [];
  const addIssue = (code, level, count, message) => {
    if (count > 0) {
      issues.push({ code, level, count, message });
    }
  };

  addIssue(
    "ROAD_MISSING_IDENTITY",
    "ERROR",
    scalar("SELECT COUNT(*) AS count FROM road WHERE TRIM(road_code) = '' OR TRIM(designation) = ''"),
    "Voies sans code ou designation."
  );

  addIssue(
    "ROAD_DUPLICATE_BY_SAP",
    "WARNING",
    scalar(`
      SELECT COUNT(*) AS count FROM (
        SELECT COALESCE(sap_code, '') AS sap, road_code, COUNT(*) AS c
        FROM road
        WHERE TRIM(road_code) <> ''
        GROUP BY COALESCE(sap_code, ''), road_code
        HAVING COUNT(*) > 1
      ) x
    `),
    "Doublons de code voie dans un meme SAP."
  );

  addIssue(
    "ROAD_SECTION_ORPHAN",
    "WARNING",
    scalar("SELECT COUNT(*) AS count FROM road_section WHERE road_id IS NULL"),
    "Sections non rattachees a une voie."
  );

  addIssue(
    "ROAD_MEASUREMENT_ORPHAN",
    "WARNING",
    scalar("SELECT COUNT(*) AS count FROM road_measurement WHERE road_id IS NULL"),
    "Mesures non rattachees a une voie."
  );

  addIssue(
    "DEGRADATION_NO_CAUSE",
    "WARNING",
    scalar(`
      SELECT COUNT(*) AS count
      FROM degradation d
      LEFT JOIN degradation_cause c ON c.degradation_code = d.degradation_code
      WHERE c.id IS NULL
    `),
    "Degradations sans cause probable."
  );

  addIssue(
    "DEGRADATION_NO_DEFINITION",
    "WARNING",
    scalar(`
      SELECT COUNT(*) AS count
      FROM degradation d
      LEFT JOIN degradation_definition dd ON dd.degradation_code = d.degradation_code
      WHERE dd.degradation_code IS NULL
    `),
    "Degradations sans metadonnees definies (categorie/famille)."
  );

  const activeDrainageRuleCount = scalar("SELECT COUNT(*) AS count FROM drainage_rule WHERE is_active = 1");
  addIssue(
    "DRAINAGE_RULE_INACTIVE",
    "ERROR",
    activeDrainageRuleCount === 0 ? 1 : 0,
    "Aucune regle assainissement active."
  );

  addIssue(
    "PROFILE_INPUT_EMPTY",
    "WARNING",
    totals.profileInputs === 0 ? 1 : 0,
    "Feuil4 non alimentee en base technique."
  );

  const status = issues.length === 0 ? "OK" : "WARNING";

  return {
    generatedAt: new Date().toISOString(),
    status,
    totals,
    issues
  };
}

function getDashboardSummary(db) {
  const scalar = (sql, ...params) => {
    const row = db.prepare(sql).get(...params);
    const value = row ? Object.values(row)[0] : 0;
    return Number(value) || 0;
  };

  const totals = {
    roads: scalar("SELECT COUNT(*) AS count FROM road"),
    degradations: scalar("SELECT COUNT(*) AS count FROM degradation"),
    decisionHistory: scalar("SELECT COUNT(*) AS count FROM decision_history"),
    maintenance: scalar("SELECT COUNT(*) AS count FROM maintenance_intervention"),
    pendingMaintenance: scalar(
      "SELECT COUNT(*) AS count FROM maintenance_intervention WHERE status IN ('PREVU', 'EN_COURS')"
    ),
    completedMaintenance: scalar(
      "SELECT COUNT(*) AS count FROM maintenance_intervention WHERE status = 'TERMINE'"
    ),
    estimatedBudget: scalar("SELECT COALESCE(SUM(cost_amount), 0) AS total FROM maintenance_intervention"),
    urgentDrainage: scalar(`
      SELECT COUNT(*) AS count
      FROM road
      WHERE UPPER(COALESCE(drainage_state, '')) LIKE '%OBSTR%'
         OR UPPER(COALESCE(drainage_state, '')) LIKE '%MAUV%'
         OR UPPER(COALESCE(drainage_state, '')) LIKE '%NON FONCTION%'
    `)
  };

  const roadsBySap = db
    .prepare(
      `
      SELECT COALESCE(sap_code, 'NON RENSEIGNE') AS label, COUNT(*) AS count
      FROM road
      GROUP BY COALESCE(sap_code, 'NON RENSEIGNE')
      ORDER BY count DESC, label
    `
    )
    .all()
    .map((row) => ({ label: toText(row.label), count: Number(row.count) || 0 }));

  const roadsByState = db
    .prepare(
      `
      SELECT COALESCE(NULLIF(TRIM(pavement_state), ''), 'NON RENSEIGNE') AS label, COUNT(*) AS count
      FROM road
      GROUP BY COALESCE(NULLIF(TRIM(pavement_state), ''), 'NON RENSEIGNE')
      ORDER BY count DESC, label
      LIMIT 8
    `
    )
    .all()
    .map((row) => ({ label: toText(row.label), count: Number(row.count) || 0 }));

  const maintenanceByStatus = db
    .prepare(
      `
      SELECT status AS label, COUNT(*) AS count
      FROM maintenance_intervention
      GROUP BY status
      ORDER BY count DESC, status
    `
    )
    .all()
    .map((row) => ({ label: normalizeMaintenanceStatus(row.label) || "PREVU", count: Number(row.count) || 0 }));

  const topDegradations = db
    .prepare(
      `
      SELECT degradation_name AS label, COUNT(*) AS count
      FROM (
        SELECT degradation_name FROM decision_history
        UNION ALL
        SELECT degradation_name FROM maintenance_intervention
      ) x
      WHERE TRIM(COALESCE(degradation_name, '')) <> ''
      GROUP BY degradation_name
      ORDER BY count DESC, degradation_name
      LIMIT 8
    `
    )
    .all()
    .map((row) => ({ label: toText(row.label), count: Number(row.count) || 0 }));

  return {
    generatedAt: new Date().toISOString(),
    totals,
    roadsBySap,
    roadsByState,
    maintenanceByStatus,
    topDegradations,
    integrity: getDataIntegrityReport(db),
    recentMaintenance: listMaintenanceInterventions(db, { limit: 5 })
  };
}

function previewExcelImport(excelPath) {
  if (!excelPath || !fs.existsSync(excelPath)) {
    throw new Error("Fichier Excel introuvable.");
  }

  const workbook = XLSX.readFile(excelPath);
  const workbookSheetNames = Array.isArray(workbook.SheetNames) ? workbook.SheetNames.map((name) => toText(name)) : [];
  const sheetPreviews = SHEET_DEFINITIONS.map((sheet) => {
    const rows = workbook.Sheets[sheet.name] ? getSheetRows(workbook, sheet) : [];
    return {
      name: sheet.name,
      title: sheet.title,
      present: Boolean(workbook.Sheets[sheet.name]),
      rowCount: rows.length,
      expectedColumns: sheet.columns.length
    };
  });

  const missingSheets = sheetPreviews.filter((item) => !item.present).map((item) => item.name);
  const warnings = [
    ...sheetPreviews
      .filter((item) => item.present && item.rowCount === 0)
      .map((item) => `${item.name} est presente mais vide.`),
    ...missingSheets.map((sheetName) => `${sheetName} est absente du fichier.`)
  ];

  const estimates = estimateWorkbookEntities(workbook);

  return {
    filePath: excelPath,
    workbookSheetNames,
    missingSheets,
    warnings,
    ready: missingSheets.length === 0,
    totals: {
      rows: sheetPreviews.reduce((sum, item) => sum + item.rowCount, 0),
      roads: estimates.roads,
      degradations: estimates.degradations,
      sections: estimates.sections
    },
    sheetPreviews
  };
}

function estimateWorkbookEntities(workbook) {
  const roadKeys = new Set();
  for (const row of getSheetRows(workbook, "Feuil6")) {
    const roadKey = makeRoadKey(row.C, row.E);
    if (roadKey && roadKey !== "NAME_") {
      roadKeys.add(roadKey);
    }
  }
  if (roadKeys.size === 0) {
    for (const row of getSheetRows(workbook, "Feuil2")) {
      if (isRoadLabel(row.C, row.D)) {
        roadKeys.add(makeRoadKey(row.C, row.D));
      }
    }
  }

  const degradationKeys = new Set();
  for (const row of getSheetRows(workbook, "Feuil7")) {
    const code = normalizeDegradationKey(row.C || row.B);
    if (code) {
      degradationKeys.add(code);
    }
  }

  return {
    roads: roadKeys.size,
    degradations: degradationKeys.size,
    sections: getSheetRows(workbook, "Feuil2").length
  };
}

function exportBackupSnapshot(db, filePath) {
  const targetPath = toText(filePath);
  if (!targetPath) {
    throw new Error("Chemin de sauvegarde invalide.");
  }

  const payload = {
    version: 1,
    exportedAt: new Date().toISOString(),
    tables: {}
  };

  for (const tableName of BACKUP_TABLES) {
    if (!tableExists(db, tableName)) {
      continue;
    }
    payload.tables[tableName] = db.prepare(`SELECT * FROM ${tableName}`).all();
  }

  fs.mkdirSync(path.dirname(targetPath), { recursive: true });
  fs.writeFileSync(targetPath, JSON.stringify(payload, null, 2), "utf8");
  const stats = fs.statSync(targetPath);
  return {
    filePath: targetPath,
    size: Number(stats.size) || 0,
    exportedAt: payload.exportedAt
  };
}

function restoreBackupSnapshot(db, filePath) {
  const sourcePath = toText(filePath);
  if (!sourcePath || !fs.existsSync(sourcePath)) {
    throw new Error("Fichier de sauvegarde introuvable.");
  }

  const payload = JSON.parse(fs.readFileSync(sourcePath, "utf8"));
  if (!payload || typeof payload !== "object" || !payload.tables || typeof payload.tables !== "object") {
    throw new Error("Format de sauvegarde invalide.");
  }

  const tx = db.transaction(() => {
    db.pragma("foreign_keys = OFF");

    for (const tableName of [...BACKUP_TABLES].reverse()) {
      if (!tableExists(db, tableName)) {
        continue;
      }
      db.prepare(`DELETE FROM ${tableName}`).run();
      resetAutoincrementSequence(db, tableName);
    }

    for (const tableName of BACKUP_TABLES) {
      if (!tableExists(db, tableName)) {
        continue;
      }
      const rows = Array.isArray(payload.tables[tableName]) ? payload.tables[tableName] : [];
      if (rows.length === 0) {
        continue;
      }

      const validColumns = new Set(
        db.prepare(`PRAGMA table_info(${tableName})`).all().map((column) => toText(column.name))
      );

      for (const row of rows) {
        const columns = Object.keys(row).filter((column) => validColumns.has(column));
        if (columns.length === 0) {
          continue;
        }
        const placeholders = columns.map(() => "?").join(", ");
        const values = columns.map((column) => row[column]);
        db.prepare(`INSERT INTO ${tableName} (${columns.join(", ")}) VALUES (${placeholders})`).run(...values);
      }
    }

    db.pragma("foreign_keys = ON");
  });

  tx();

  seedSolutionCatalog(db);
  seedDeflectionRules(db);
  seedDrainageRules(db);
  if (isDatabaseEmpty(db) === false) {
    rebuildNormalizedCatalogs(db);
    migrateLegacySolutionMapping(db);
  }
  setMeta(db, "backup_restored_at", new Date().toISOString());
  setMeta(db, "backup_restored_from", sourcePath);
  return getDataStatus(db);
}

function exportReportWorkbook(db, reportType, filePath) {
  const targetPath = toText(filePath);
  if (!targetPath) {
    throw new Error("Chemin d'export invalide.");
  }

  const workbook = XLSX.utils.book_new();
  let rows = [];
  let sheetName = "Rapport";

  if (reportType === "maintenance") {
    sheetName = "Entretiens";
    rows = listMaintenanceInterventions(db, { limit: 5000 }).map((item) => ({
      Date: item.interventionDate,
      Statut: item.status,
      SAP: item.sapCode,
      CodeVoie: item.roadCode,
      Voie: item.roadDesignation,
      TypeEntretien: item.interventionType,
      Degradation: item.degradationName,
      EtatAvant: item.stateBefore,
      EtatApres: item.stateAfter,
      SolutionAppliquee: item.solutionApplied,
      Prestataire: item.contractorName,
      Responsable: item.responsibleName,
      Cout: item.costAmount,
      PieceJointe: item.attachmentPath,
      Observation: item.observation
    }));
  } else {
    sheetName = "Decisions";
    rows = listDecisionHistory(db, { limit: 5000 }).map((item) => ({
      Date: item.createdAt,
      SAP: item.sapCode,
      CodeVoie: item.roadCode,
      Voie: item.roadDesignation,
      Degradation: item.degradationName,
      CauseProbable: item.probableCause,
      Deflexion: item.deflectionValue,
      Severite: item.deflectionSeverity,
      Recommendation: item.deflectionRecommendation,
      Assainissement: item.drainageRecommendation
    }));
  }

  const worksheet = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
  fs.mkdirSync(path.dirname(targetPath), { recursive: true });
  XLSX.writeFile(workbook, targetPath);
  return {
    filePath: targetPath,
    reportType: reportType === "maintenance" ? "maintenance" : "history",
    rowCount: rows.length
  };
}

function listSapSectors(db) {
  return db
    .prepare(
      `
      SELECT code, name
      FROM sap_sector
      ORDER BY sort_order, code
    `
    )
    .all()
    .map((row) => ({
      code: toText(row.code),
      name: toText(row.name) || toText(row.code)
    }));
}

function listRoadCatalog(db, filters = {}) {
  const where = [];
  const params = {};

  const sapCode = toText(filters.sapCode);
  if (sapCode) {
    where.push("sap_code = @sapCode");
    params.sapCode = sapCode;
  }

  const search = toText(filters.search);
  if (search) {
    where.push("(road_code LIKE @search OR designation LIKE @search OR start_label LIKE @search OR end_label LIKE @search)");
    params.search = `%${search}%`;
  }

  const rows = db
    .prepare(
      `
      SELECT
        id,
        road_key AS roadKey,
        road_code AS roadCode,
        designation,
        COALESCE(sap_code, '') AS sapCode,
        start_label AS startLabel,
        end_label AS endLabel,
        length_m AS lengthM,
        width_m AS widthM,
        surface_type AS surfaceType,
        pavement_state AS pavementState,
        drainage_type AS drainageType,
        drainage_state AS drainageState,
        sidewalk_min_m AS sidewalkMinM,
        parking_left AS parkingLeft,
        parking_right AS parkingRight,
        parking_other AS parkingOther,
        itinerary,
        justification,
        intervention_hint AS interventionHint
      FROM road
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY sap_code, road_code, designation
      LIMIT 5000
    `
    )
    .all(params);

  return rows.map((row) => ({
    ...row,
    lengthM: toNumber(row.lengthM),
    widthM: toNumber(row.widthM),
    sidewalkMinM: toNumber(row.sidewalkMinM)
  }));
}

function listRoadSections(db, filters = {}) {
  const where = [];
  const params = {};

  const sapCode = toText(filters.sapCode);
  if (sapCode) {
    where.push("sap_code = @sapCode");
    params.sapCode = sapCode;
  }

  const search = toText(filters.search);
  if (search) {
    where.push(
      `(
        troncon_no LIKE @search OR
        section_no LIKE @search OR
        road_code LIKE @search OR
        designation LIKE @search OR
        start_label LIKE @search OR
        end_label LIKE @search
      )`
    );
    params.search = `%${search}%`;
  }

  const rows = db
    .prepare(
      `
      SELECT
        id,
        section_key AS sectionKey,
        source_sheet AS sourceSheet,
        source_row_no AS sourceRowNo,
        troncon_no AS tronconNo,
        section_no AS sectionNo,
        road_key AS roadKey,
        road_id AS roadId,
        COALESCE(sap_code, '') AS sapCode,
        road_code AS roadCode,
        designation,
        start_label AS startLabel,
        end_label AS endLabel,
        length_m AS lengthM,
        width_m AS widthM,
        surface_type AS surfaceType,
        pavement_state AS pavementState,
        drainage_type AS drainageType,
        drainage_state AS drainageState,
        sidewalk_min_m AS sidewalkMinM,
        intervention_hint AS interventionHint,
        source_payload AS sourcePayload
      FROM road_section
      ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
      ORDER BY sap_code, troncon_no, section_no, designation
      LIMIT 5000
    `
    )
    .all(params);

  return rows.map((row) => ({
    ...row,
    roadId: Number.isFinite(Number(row.roadId)) ? Number(row.roadId) : null,
    sourceRowNo: Number(row.sourceRowNo) || 0,
    lengthM: toNumber(row.lengthM),
    widthM: toNumber(row.widthM),
    sidewalkMinM: toNumber(row.sidewalkMinM)
  }));
}

function listMeasurementCampaigns(db, filters = {}) {
  ensureMeasurementRuntimeSchema(db);

  const where = [];
  const params = {};

  const roadId = Number(filters.roadId);
  if (Number.isFinite(roadId) && roadId > 0) {
    where.push("c.road_id = @roadId");
    params.roadId = roadId;
  }

  const search = toText(filters.search);
  if (search) {
    where.push(
      `(
        c.road_code LIKE @search OR
        c.designation LIKE @search OR
        c.section_label LIKE @search OR
        COALESCE(r.sap_code, '') LIKE @search OR
        c.measurement_date LIKE @search
      )`
    );
    params.search = `%${search}%`;
  }

  const limit = Math.min(Math.max(Number(filters.limit) || 500, 1), 5000);
  params.limit = limit;

  const sql = `
    SELECT
      c.id,
      c.campaign_key AS campaignKey,
      c.road_id AS roadId,
      c.road_key AS roadKey,
      c.road_code AS roadCode,
      c.designation,
      COALESCE(r.sap_code, '') AS sapCode,
      c.section_label AS sectionLabel,
      c.start_label AS startLabel,
      c.end_label AS endLabel,
      c.measurement_date AS measurementDate,
      c.pk_start_m AS pkStartM,
      c.pk_end_m AS pkEndM,
      COUNT(m.id) AS measurementCount,
      MAX(m.deflection_dc) AS maxDeflectionDc,
      AVG(m.deflection_dc) AS avgDeflectionDc
    FROM measurement_campaign c
    LEFT JOIN road r ON r.id = c.road_id
    LEFT JOIN road_measurement m ON m.campaign_key = c.campaign_key
    ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
    GROUP BY
      c.id,
      c.campaign_key,
      c.road_id,
      c.road_key,
      c.road_code,
      c.designation,
      r.sap_code,
      c.section_label,
      c.start_label,
      c.end_label,
      c.measurement_date,
      c.pk_start_m,
      c.pk_end_m
    ORDER BY c.measurement_date DESC, c.id DESC
    LIMIT @limit
  `;

  return db.prepare(sql).all(params).map((row) => ({
    ...row,
    roadId: Number.isFinite(Number(row.roadId)) ? Number(row.roadId) : null,
    pkStartM: toNumber(row.pkStartM),
    pkEndM: toNumber(row.pkEndM),
    measurementCount: Number(row.measurementCount) || 0,
    maxDeflectionDc: toNumber(row.maxDeflectionDc),
    avgDeflectionDc: toNumber(row.avgDeflectionDc)
  }));
}

function listRoadMeasurements(db, filters = {}) {
  ensureMeasurementRuntimeSchema(db);

  const where = [];
  const params = {};

  const campaignKey = toText(filters.campaignKey);
  if (campaignKey) {
    where.push("campaign_key = @campaignKey");
    params.campaignKey = campaignKey;
  }

  const roadId = Number(filters.roadId);
  if (Number.isFinite(roadId) && roadId > 0) {
    where.push("road_id = @roadId");
    params.roadId = roadId;
  }

  const limit = Math.min(Math.max(Number(filters.limit) || 2000, 1), 10000);
  params.limit = limit;

  const sql = `
    SELECT
      id,
      campaign_key AS campaignKey,
      measurement_date AS measurementDate,
      road_id AS roadId,
      road_key AS roadKey,
      road_code AS roadCode,
      designation,
      start_label AS startLabel,
      end_label AS endLabel,
      pk_label AS pkLabel,
      pk_m AS pkM,
      lecture_left AS lectureLeft,
      lecture_axis AS lectureAxis,
      lecture_right AS lectureRight,
      deflection_left AS deflectionLeft,
      deflection_axis AS deflectionAxis,
      deflection_right AS deflectionRight,
      deflection_avg AS deflectionAvg,
      std_dev AS stdDev,
      deflection_dc AS deflectionDc
    FROM road_measurement
    ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
    ORDER BY COALESCE(pk_m, 999999999), source_row_no, id
    LIMIT @limit
  `;

  return db.prepare(sql).all(params).map((row) => ({
    ...row,
    roadId: Number.isFinite(Number(row.roadId)) ? Number(row.roadId) : null,
    pkM: toNumber(row.pkM),
    lectureLeft: toNumber(row.lectureLeft),
    lectureAxis: toNumber(row.lectureAxis),
    lectureRight: toNumber(row.lectureRight),
    deflectionLeft: toNumber(row.deflectionLeft),
    deflectionAxis: toNumber(row.deflectionAxis),
    deflectionRight: toNumber(row.deflectionRight),
    deflectionAvg: toNumber(row.deflectionAvg),
    stdDev: toNumber(row.stdDev),
    deflectionDc: toNumber(row.deflectionDc)
  }));
}

function normalizeMeasurementDateInput(value) {
  const text = toText(value);
  if (!text) {
    return "";
  }
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) {
    return text;
  }
  const match = text.match(/^(\d{2})[\/.-](\d{2})[\/.-](\d{4})$/);
  if (match) {
    return `${match[3]}-${match[2]}-${match[1]}`;
  }
  return "";
}

function normalizeMeasurementPkKey(value) {
  return toText(value).replace(/\s+/g, "").replace(",", ".").toUpperCase();
}

function formatMeasurementPkLabel(value) {
  const numeric = toNumber(value);
  if (numeric === null) {
    return "";
  }
  return numeric.toFixed(3);
}

function makeManualMeasurementCampaignKey(road, measurementDate) {
  const roadKey = normalizeRoadAliasValue(road?.designation || road?.roadCode || String(road?.id || ""));
  const dateKey = toText(measurementDate) || "NO_DATE";
  const entropy = `${Date.now()}${Math.random().toString(36).slice(2, 8)}`;
  return `MANUAL:${roadKey || "NO_ROAD"}:${dateKey}:${entropy}`;
}

function makeManualMeasurementKey(campaignKey, pkLabel) {
  const pkKey = normalizeMeasurementPkKey(pkLabel) || "NO_PK";
  const entropy = `${Date.now()}${Math.random().toString(36).slice(2, 8)}`;
  return `MANUAL:${campaignKey}:${pkKey}:${entropy}`;
}

function getRoadForMeasurement(db, roadId) {
  const normalizedRoadId = Number(roadId);
  if (!Number.isFinite(normalizedRoadId) || normalizedRoadId <= 0) {
    return null;
  }
  return (
    db
      .prepare(
        `
        SELECT
          id,
          road_key AS roadKey,
          road_code AS roadCode,
          designation,
          COALESCE(sap_code, '') AS sapCode,
          COALESCE(start_label, '') AS startLabel,
          COALESCE(end_label, '') AS endLabel
        FROM road
        WHERE id = ?
      `
      )
      .get(normalizedRoadId) || null
  );
}

function getMeasurementCampaignRecordById(db, campaignId) {
  const normalizedCampaignId = Number(campaignId);
  if (!Number.isFinite(normalizedCampaignId) || normalizedCampaignId <= 0) {
    return null;
  }
  return db.prepare("SELECT * FROM measurement_campaign WHERE id = ?").get(normalizedCampaignId) || null;
}

function getMeasurementCampaignRecordByKey(db, campaignKey) {
  const normalizedKey = toText(campaignKey);
  if (!normalizedKey) {
    return null;
  }
  return db.prepare("SELECT * FROM measurement_campaign WHERE campaign_key = ?").get(normalizedKey) || null;
}

function getMeasurementCampaignItemById(db, campaignId) {
  const normalizedCampaignId = Number(campaignId);
  if (!Number.isFinite(normalizedCampaignId) || normalizedCampaignId <= 0) {
    return null;
  }

  const row = db
    .prepare(
      `
      SELECT
        c.id,
        c.campaign_key AS campaignKey,
        c.road_id AS roadId,
        c.road_key AS roadKey,
        c.road_code AS roadCode,
        c.designation,
        COALESCE(r.sap_code, '') AS sapCode,
        c.section_label AS sectionLabel,
        c.start_label AS startLabel,
        c.end_label AS endLabel,
        c.measurement_date AS measurementDate,
        c.pk_start_m AS pkStartM,
        c.pk_end_m AS pkEndM,
        COUNT(m.id) AS measurementCount,
        MAX(m.deflection_dc) AS maxDeflectionDc,
        AVG(m.deflection_dc) AS avgDeflectionDc
      FROM measurement_campaign c
      LEFT JOIN road r ON r.id = c.road_id
      LEFT JOIN road_measurement m ON m.campaign_key = c.campaign_key
      WHERE c.id = @campaignId
      GROUP BY
        c.id,
        c.campaign_key,
        c.road_id,
        c.road_key,
        c.road_code,
        c.designation,
        r.sap_code,
        c.section_label,
        c.start_label,
        c.end_label,
        c.measurement_date,
        c.pk_start_m,
        c.pk_end_m
    `
    )
    .get({ campaignId: normalizedCampaignId });

  if (!row) {
    return null;
  }

  return {
    ...row,
    roadId: Number.isFinite(Number(row.roadId)) ? Number(row.roadId) : null,
    pkStartM: toNumber(row.pkStartM),
    pkEndM: toNumber(row.pkEndM),
    measurementCount: Number(row.measurementCount) || 0,
    maxDeflectionDc: toNumber(row.maxDeflectionDc),
    avgDeflectionDc: toNumber(row.avgDeflectionDc)
  };
}

function getRoadMeasurementRecordById(db, measurementId) {
  const normalizedMeasurementId = Number(measurementId);
  if (!Number.isFinite(normalizedMeasurementId) || normalizedMeasurementId <= 0) {
    return null;
  }
  return db.prepare("SELECT * FROM road_measurement WHERE id = ?").get(normalizedMeasurementId) || null;
}

function getRoadMeasurementItemById(db, measurementId) {
  const normalizedMeasurementId = Number(measurementId);
  if (!Number.isFinite(normalizedMeasurementId) || normalizedMeasurementId <= 0) {
    return null;
  }

  const row = db
    .prepare(
      `
      SELECT
        id,
        campaign_key AS campaignKey,
        measurement_date AS measurementDate,
        road_id AS roadId,
        road_key AS roadKey,
        road_code AS roadCode,
        designation,
        start_label AS startLabel,
        end_label AS endLabel,
        pk_label AS pkLabel,
        pk_m AS pkM,
        lecture_left AS lectureLeft,
        lecture_axis AS lectureAxis,
        lecture_right AS lectureRight,
        deflection_left AS deflectionLeft,
        deflection_axis AS deflectionAxis,
        deflection_right AS deflectionRight,
        deflection_avg AS deflectionAvg,
        std_dev AS stdDev,
        deflection_dc AS deflectionDc
      FROM road_measurement
      WHERE id = ?
    `
    )
    .get(normalizedMeasurementId);

  if (!row) {
    return null;
  }

  return {
    ...row,
    roadId: Number.isFinite(Number(row.roadId)) ? Number(row.roadId) : null,
    pkM: toNumber(row.pkM),
    lectureLeft: toNumber(row.lectureLeft),
    lectureAxis: toNumber(row.lectureAxis),
    lectureRight: toNumber(row.lectureRight),
    deflectionLeft: toNumber(row.deflectionLeft),
    deflectionAxis: toNumber(row.deflectionAxis),
    deflectionRight: toNumber(row.deflectionRight),
    deflectionAvg: toNumber(row.deflectionAvg),
    stdDev: toNumber(row.stdDev),
    deflectionDc: toNumber(row.deflectionDc)
  };
}

function assertMeasurementCampaignPayload(payload) {
  const sectionLabel = toText(payload.sectionLabel);
  const startLabel = toText(payload.startLabel);
  const endLabel = toText(payload.endLabel);
  const measurementDate = normalizeMeasurementDateInput(payload.measurementDate);
  const pkStartM = payload.pkStartM === "" || payload.pkStartM == null ? null : toNumber(payload.pkStartM);
  const pkEndM = payload.pkEndM === "" || payload.pkEndM == null ? null : toNumber(payload.pkEndM);

  if (!sectionLabel) {
    throw new Error("Nom du tronçon obligatoire.");
  }
  if (!startLabel) {
    throw new Error("Nom du point de départ obligatoire.");
  }
  if (!endLabel) {
    throw new Error("Nom du point d'arrivée obligatoire.");
  }
  if (!measurementDate) {
    throw new Error("Date de mesure obligatoire.");
  }
  if ((pkStartM === null) !== (pkEndM === null)) {
    throw new Error("Renseigne à la fois le PK début et le PK fin.");
  }
  if (pkStartM !== null && pkEndM !== null && pkStartM > pkEndM) {
    throw new Error("Le PK début doit être inférieur ou égal au PK fin.");
  }

  return {
    sectionLabel,
    startLabel,
    endLabel,
    measurementDate,
    pkStartM,
    pkEndM
  };
}

function upsertMeasurementCampaign(db, payload = {}) {
  ensureMeasurementRuntimeSchema(db);

  const road = getRoadForMeasurement(db, payload.roadId);
  if (!road) {
    throw new Error("Choisis la voie concernée par cette campagne.");
  }

  const campaignValues = assertMeasurementCampaignPayload(payload);
  const campaignId = Number(payload.id);

  if (Number.isFinite(campaignId) && campaignId > 0) {
    const current = getMeasurementCampaignRecordById(db, campaignId);
    if (!current) {
      throw new Error("Campagne introuvable.");
    }

    db.prepare(
      `
      UPDATE measurement_campaign
      SET
        road_id = ?,
        road_key = ?,
        road_code = ?,
        designation = ?,
        section_label = ?,
        start_label = ?,
        end_label = ?,
        measurement_date = ?,
        pk_start_m = ?,
        pk_end_m = ?,
        source_payload = ?,
        updated_at = datetime('now')
      WHERE id = ?
    `
    ).run(
      Number(road.id),
      toText(road.roadKey),
      toText(road.roadCode),
      toText(road.designation),
      campaignValues.sectionLabel,
      campaignValues.startLabel,
      campaignValues.endLabel,
      campaignValues.measurementDate,
      campaignValues.pkStartM,
      campaignValues.pkEndM,
      safeJson({ mode: "manual", source: current.source_sheet || "MANUAL" }),
      campaignId
    );

    db.prepare(
      `
      UPDATE road_measurement
      SET
        measurement_date = ?,
        road_id = ?,
        road_key = ?,
        road_code = ?,
        designation = ?,
        start_label = ?,
        end_label = ?,
        updated_at = datetime('now')
      WHERE campaign_key = ?
    `
    ).run(
      campaignValues.measurementDate,
      Number(road.id),
      toText(road.roadKey),
      toText(road.roadCode),
      toText(road.designation),
      campaignValues.startLabel,
      campaignValues.endLabel,
      toText(current.campaign_key)
    );

    return getMeasurementCampaignItemById(db, campaignId);
  }

  const campaignKey = makeManualMeasurementCampaignKey(road, campaignValues.measurementDate);
  const insertedId = db
    .prepare(
      `
      INSERT INTO measurement_campaign (
        campaign_key,
        source_sheet,
        source_row_no,
        road_id,
        road_key,
        road_code,
        designation,
        section_label,
        start_label,
        end_label,
        measurement_date,
        pk_start_m,
        pk_end_m,
        source_payload
      ) VALUES (?, 'MANUAL', 0, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `
    )
    .run(
      campaignKey,
      Number(road.id),
      toText(road.roadKey),
      toText(road.roadCode),
      toText(road.designation),
      campaignValues.sectionLabel,
      campaignValues.startLabel,
      campaignValues.endLabel,
      campaignValues.measurementDate,
      campaignValues.pkStartM,
      campaignValues.pkEndM,
      safeJson({ mode: "manual" })
    ).lastInsertRowid;

  return getMeasurementCampaignItemById(db, insertedId);
}

function deleteMeasurementCampaign(db, campaignId) {
  ensureMeasurementRuntimeSchema(db);

  const current = getMeasurementCampaignRecordById(db, campaignId);
  if (!current) {
    throw new Error("Campagne introuvable.");
  }

  db.prepare("DELETE FROM road_measurement WHERE campaign_key = ?").run(toText(current.campaign_key));
  db.prepare("DELETE FROM measurement_campaign WHERE id = ?").run(Number(campaignId));
  return { deleted: true };
}

function assertMeasurementPayload(payload) {
  const pkLabelDraft = toText(payload.pkLabel);
  const numericPk = payload.pkM === "" || payload.pkM == null ? null : toNumber(payload.pkM);
  const pkLabel = pkLabelDraft || formatMeasurementPkLabel(numericPk);
  const pkM = numericPk !== null ? numericPk : pkLabel ? parsePkDistance(pkLabel) : null;

  if (!pkLabel && pkM === null) {
    throw new Error("PK obligatoire.");
  }

  const values = {
    lectureLeft: payload.lectureLeft === "" || payload.lectureLeft == null ? null : toNumber(payload.lectureLeft),
    lectureAxis: payload.lectureAxis === "" || payload.lectureAxis == null ? null : toNumber(payload.lectureAxis),
    lectureRight: payload.lectureRight === "" || payload.lectureRight == null ? null : toNumber(payload.lectureRight),
    deflectionLeft: payload.deflectionLeft === "" || payload.deflectionLeft == null ? null : toNumber(payload.deflectionLeft),
    deflectionAxis: payload.deflectionAxis === "" || payload.deflectionAxis == null ? null : toNumber(payload.deflectionAxis),
    deflectionRight: payload.deflectionRight === "" || payload.deflectionRight == null ? null : toNumber(payload.deflectionRight),
    deflectionAvg: payload.deflectionAvg === "" || payload.deflectionAvg == null ? null : toNumber(payload.deflectionAvg),
    stdDev: payload.stdDev === "" || payload.stdDev == null ? null : toNumber(payload.stdDev),
    deflectionDc: payload.deflectionDc === "" || payload.deflectionDc == null ? null : toNumber(payload.deflectionDc)
  };

  if (Object.values(values).every((value) => value === null)) {
    throw new Error("Renseigne au moins une valeur de mesure.");
  }

  return {
    pkLabel,
    pkM,
    ...values
  };
}

function assertMeasurementPkUniqueness(db, campaignKey, pkLabel, pkM, excludedId) {
  const rows = db
    .prepare(
      `
      SELECT id, pk_label AS pkLabel, pk_m AS pkM
      FROM road_measurement
      WHERE campaign_key = ?
        AND (? IS NULL OR id <> ?)
    `
    )
    .all(toText(campaignKey), Number.isFinite(Number(excludedId)) ? Number(excludedId) : null, Number(excludedId) || null);

  const nextPkKey = normalizeMeasurementPkKey(pkLabel);
  for (const row of rows) {
    const sameNumericPk = pkM !== null && toNumber(row.pkM) !== null && toNumber(row.pkM) === pkM;
    const sameLabelPk = nextPkKey && normalizeMeasurementPkKey(row.pkLabel) === nextPkKey;
    if (sameNumericPk || sameLabelPk) {
      throw new Error("Une ligne avec ce PK existe déjà dans cette campagne.");
    }
  }
}

function upsertRoadMeasurement(db, payload = {}) {
  ensureMeasurementRuntimeSchema(db);

  const measurementId = Number(payload.id);
  const current = Number.isFinite(measurementId) && measurementId > 0 ? getRoadMeasurementRecordById(db, measurementId) : null;
  if (Number.isFinite(measurementId) && measurementId > 0 && !current) {
    throw new Error("Ligne de mesure introuvable.");
  }

  const campaignKey = toText(payload.campaignKey) || toText(current?.campaign_key);
  const campaign = getMeasurementCampaignRecordByKey(db, campaignKey);
  if (!campaign) {
    throw new Error("Choisis d'abord une campagne de mesure.");
  }

  const measurementValues = assertMeasurementPayload(payload);
  assertMeasurementPkUniqueness(db, campaignKey, measurementValues.pkLabel, measurementValues.pkM, current?.id);

  if (current) {
    db.prepare(
      `
      UPDATE road_measurement
      SET
        campaign_key = ?,
        measurement_date = ?,
        road_id = ?,
        road_key = ?,
        road_code = ?,
        designation = ?,
        start_label = ?,
        end_label = ?,
        pk_label = ?,
        pk_m = ?,
        lecture_left = ?,
        lecture_axis = ?,
        lecture_right = ?,
        deflection_left = ?,
        deflection_axis = ?,
        deflection_right = ?,
        deflection_avg = ?,
        std_dev = ?,
        deflection_dc = ?,
        source_payload = ?,
        updated_at = datetime('now')
      WHERE id = ?
    `
    ).run(
      campaignKey,
      toText(campaign.measurement_date),
      Number(campaign.road_id) || null,
      toText(campaign.road_key),
      toText(campaign.road_code),
      toText(campaign.designation),
      toText(campaign.start_label),
      toText(campaign.end_label),
      measurementValues.pkLabel,
      measurementValues.pkM,
      measurementValues.lectureLeft,
      measurementValues.lectureAxis,
      measurementValues.lectureRight,
      measurementValues.deflectionLeft,
      measurementValues.deflectionAxis,
      measurementValues.deflectionRight,
      measurementValues.deflectionAvg,
      measurementValues.stdDev,
      measurementValues.deflectionDc,
      safeJson({ mode: "manual", campaignKey }),
      Number(current.id)
    );

    return getRoadMeasurementItemById(db, current.id);
  }

  const measurementKey = makeManualMeasurementKey(campaignKey, measurementValues.pkLabel);
  const insertedId = db
    .prepare(
      `
      INSERT INTO road_measurement (
        measurement_key,
        source_sheet,
        source_row_no,
        campaign_key,
        measurement_date,
        road_id,
        road_key,
        road_code,
        designation,
        start_label,
        end_label,
        pk_label,
        pk_m,
        lecture_left,
        lecture_axis,
        lecture_right,
        deflection_left,
        deflection_axis,
        deflection_right,
        deflection_avg,
        std_dev,
        deflection_dc,
        source_payload
      ) VALUES (?, 'MANUAL', 0, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `
    )
    .run(
      measurementKey,
      campaignKey,
      toText(campaign.measurement_date),
      Number(campaign.road_id) || null,
      toText(campaign.road_key),
      toText(campaign.road_code),
      toText(campaign.designation),
      toText(campaign.start_label),
      toText(campaign.end_label),
      measurementValues.pkLabel,
      measurementValues.pkM,
      measurementValues.lectureLeft,
      measurementValues.lectureAxis,
      measurementValues.lectureRight,
      measurementValues.deflectionLeft,
      measurementValues.deflectionAxis,
      measurementValues.deflectionRight,
      measurementValues.deflectionAvg,
      measurementValues.stdDev,
      measurementValues.deflectionDc,
      safeJson({ mode: "manual", campaignKey })
    ).lastInsertRowid;

  return getRoadMeasurementItemById(db, insertedId);
}

function deleteRoadMeasurement(db, measurementId) {
  ensureMeasurementRuntimeSchema(db);

  const current = getRoadMeasurementRecordById(db, measurementId);
  if (!current) {
    throw new Error("Ligne de mesure introuvable.");
  }

  db.prepare("DELETE FROM road_measurement WHERE id = ?").run(Number(measurementId));
  return { deleted: true };
}

function listDegradationCatalog(db) {
  const items = db
    .prepare(
      `
      SELECT
        id,
        degradation_code AS code,
        name
      FROM degradation
      ORDER BY name
    `
    )
    .all();

  if (items.length === 0) {
    return [];
  }

  const causeRows = db
    .prepare(
      `
      SELECT
        degradation_code AS code,
        cause_text AS cause
      FROM degradation_cause
      ORDER BY id
    `
    )
    .all();

  const causesByCode = new Map();
  for (const row of causeRows) {
    const code = toText(row.code);
    if (!causesByCode.has(code)) {
      causesByCode.set(code, []);
    }
    causesByCode.get(code).push(toText(row.cause));
  }

  const templates = getSolutionTemplateMap(db);
  const rules = getDegradationSolutionRuleMap(db);
  const overrides = getDegradationSolutionOverrideMap(db);

  return items.map((item) => {
    const code = toText(item.code);
    const resolved = resolveSolutionForDegradation(code, templates, rules, overrides);
    return {
      id: Number(item.id),
      code,
      name: toText(item.name),
      causes: causesByCode.get(code) || [],
      solution: resolved.solution,
      solutionSource: resolved.source,
      templateKey: resolved.templateKey
    };
  });
}

function evaluateDecision(db, payload = {}) {
  const roads = listRoadCatalog(db, {});
  const degradations = listDegradationCatalog(db);

  if (roads.length === 0) {
    throw new Error("Aucune voie disponible. Importe d'abord le fichier Excel.");
  }
  if (degradations.length === 0) {
    throw new Error("Aucune degradation disponible. Verifie la feuille Feuil7.");
  }

  const road = resolveRoad(roads, payload);
  if (!road) {
    throw new Error("Voie introuvable.");
  }

  const degradation = resolveDegradation(degradations, payload);
  if (!degradation) {
    throw new Error("Degradation introuvable.");
  }

  const roadHasDegradation = db
    .prepare(
      `
      SELECT is_active AS isActive
      FROM road_degradation
      WHERE road_id = ? AND degradation_code = ?
    `
    )
    .get(road.id, degradation.code);

  if (!roadHasDegradation || Number(roadHasDegradation.isActive) !== 1) {
    throw new Error("Cette degradation n'est pas active pour la voie selectionnee.");
  }

  const deflectionValue = toNumber(payload.deflectionValue);
  const deflection = classifyDeflection(db, deflectionValue);
  const drainage = buildDrainageAssessment(db, road, Boolean(payload.askDrainage));
  const probableCause = degradation.causes[0] || "Cause probable a confirmer sur le terrain.";

  const result = {
    road,
    degradation,
    probableCause,
    maintenanceSolution: degradation.solution,
    contextualIntervention: toText(road.interventionHint) || DEFAULT_INTERVENTION_TEXT,
    deflection,
    drainage
  };

  saveDecisionHistory(db, result);
  return result;
}

function listDecisionHistory(db, filters = {}) {
  const where = [];
  const params = {};

  const sapCode = toText(filters.sapCode);
  if (sapCode) {
    where.push("sap_code = @sapCode");
    params.sapCode = sapCode;
  }

  const search = toText(filters.search);
  if (search) {
    where.push(
      "(road_code LIKE @search OR road_designation LIKE @search OR degradation_name LIKE @search OR probable_cause LIKE @search)"
    );
    params.search = `%${search}%`;
  }

  const limit = Math.min(Math.max(Number(filters.limit) || 250, 1), 5000);
  params.limit = limit;

  const sql = `
    SELECT
      id,
      created_at AS createdAt,
      road_id AS roadId,
      road_code AS roadCode,
      road_designation AS roadDesignation,
      sap_code AS sapCode,
      start_label AS startLabel,
      end_label AS endLabel,
      degradation_name AS degradationName,
      probable_cause AS probableCause,
      maintenance_solution AS maintenanceSolution,
      contextual_intervention AS contextualIntervention,
      deflection_value AS deflectionValue,
      deflection_severity AS deflectionSeverity,
      deflection_recommendation AS deflectionRecommendation,
      drainage_needs_attention AS drainageNeedsAttention,
      drainage_recommendation AS drainageRecommendation
    FROM decision_history
    ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
    ORDER BY datetime(created_at) DESC, id DESC
    LIMIT @limit
  `;

  return db
    .prepare(sql)
    .all(params)
    .map((row) => ({
      ...row,
      drainageNeedsAttention: Boolean(row.drainageNeedsAttention)
    }));
}

function clearDecisionHistory(db) {
  db.prepare("DELETE FROM decision_history").run();
  return { deleted: true };
}

function listMaintenanceInterventions(db, filters = {}) {
  const where = [];
  const params = {};

  const sapCode = toText(filters.sapCode);
  if (sapCode) {
    where.push("sap_code = @sapCode");
    params.sapCode = sapCode;
  }

  const roadId = Number(filters.roadId);
  if (Number.isFinite(roadId) && roadId > 0) {
    where.push("road_id = @roadId");
    params.roadId = roadId;
  }

  const status = normalizeMaintenanceStatus(filters.status);
  if (status) {
    where.push("status = @status");
    params.status = status;
  }

  const search = toText(filters.search);
  if (search) {
    where.push(
      `(
        road_code LIKE @search OR
        road_designation LIKE @search OR
        degradation_name LIKE @search OR
        intervention_type LIKE @search OR
        solution_applied LIKE @search OR
        contractor_name LIKE @search OR
        responsible_name LIKE @search OR
        attachment_path LIKE @search OR
        observation LIKE @search
      )`
    );
    params.search = `%${search}%`;
  }

  const limit = Math.min(Math.max(Number(filters.limit) || 250, 1), 5000);
  params.limit = limit;

  const sql = `
    SELECT
      id,
      created_at AS createdAt,
      updated_at AS updatedAt,
      road_id AS roadId,
      road_key AS roadKey,
      road_code AS roadCode,
      road_designation AS roadDesignation,
      sap_code AS sapCode,
      degradation_code AS degradationCode,
      degradation_name AS degradationName,
      intervention_type AS interventionType,
      status,
      intervention_date AS interventionDate,
      completion_date AS completionDate,
      state_before AS stateBefore,
      state_after AS stateAfter,
      deflection_before AS deflectionBefore,
      deflection_after AS deflectionAfter,
      solution_applied AS solutionApplied,
      contractor_name AS contractorName,
      responsible_name AS responsibleName,
      attachment_path AS attachmentPath,
      observation,
      cost_amount AS costAmount
    FROM maintenance_intervention
    ${where.length ? `WHERE ${where.join(" AND ")}` : ""}
    ORDER BY date(intervention_date) DESC, id DESC
    LIMIT @limit
  `;

  return db
    .prepare(sql)
    .all(params)
    .map((row) => ({
      ...row,
      roadId: Number.isFinite(Number(row.roadId)) ? Number(row.roadId) : null,
      deflectionBefore: toNumber(row.deflectionBefore),
      deflectionAfter: toNumber(row.deflectionAfter),
      costAmount: toNumber(row.costAmount)
    }));
}

function upsertMaintenanceIntervention(db, payload = {}) {
  const interventionId = Number(payload.id);
  const roadId = Number(payload.roadId);
  if (!Number.isFinite(roadId) || roadId <= 0) {
    throw new Error("Voie d'entretien invalide.");
  }

  const road = db
    .prepare(
      `
      SELECT
        id,
        road_key AS roadKey,
        road_code AS roadCode,
        designation AS roadDesignation,
        COALESCE(sap_code, '') AS sapCode
      FROM road
      WHERE id = ?
    `
    )
    .get(roadId);
  if (!road) {
    throw new Error("Voie introuvable pour l'entretien.");
  }

  const degradationCode = normalizeDegradationKey(payload.degradationCode);
  let degradationName = "";
  if (degradationCode) {
    const degradation = db
      .prepare("SELECT name FROM degradation WHERE degradation_code = ?")
      .get(degradationCode);
    if (!degradation) {
      throw new Error("Degradation introuvable pour l'entretien.");
    }
    degradationName = toText(degradation.name);
  }

  const interventionType = toText(payload.interventionType);
  if (!interventionType) {
    throw new Error("Type d'entretien obligatoire.");
  }

  const interventionDate = toText(payload.interventionDate);
  if (!interventionDate) {
    throw new Error("Date d'intervention obligatoire.");
  }

  const stateBefore = toText(payload.stateBefore);
  if (!stateBefore) {
    throw new Error("Etat avant obligatoire.");
  }

  const responsibleName = toText(payload.responsibleName);
  if (!responsibleName) {
    throw new Error("Responsable PAD obligatoire.");
  }

  const costAmount = toNumber(payload.costAmount);
  if (costAmount === null || !Number.isFinite(costAmount) || costAmount < 0) {
    throw new Error("Cout estime obligatoire et valide.");
  }

  const status = normalizeMaintenanceStatus(payload.status) || "PREVU";
  const completionDate =
    toText(payload.completionDate) || (status === "TERMINE" ? interventionDate : "");

  const values = [
    Number(road.id),
    toText(road.roadKey),
    toText(road.roadCode),
    toText(road.roadDesignation),
    toText(road.sapCode) || null,
    degradationCode || null,
    degradationName,
    interventionType,
    status,
    interventionDate,
    completionDate,
    stateBefore,
    toText(payload.stateAfter),
    toNumber(payload.deflectionBefore),
    toNumber(payload.deflectionAfter),
    toText(payload.solutionApplied),
    toText(payload.contractorName),
    responsibleName,
    toText(payload.attachmentPath),
    toText(payload.observation),
    costAmount
  ];

  let targetId = null;
  if (Number.isFinite(interventionId) && interventionId > 0) {
    const existing = db.prepare("SELECT id FROM maintenance_intervention WHERE id = ?").get(interventionId);
    if (!existing) {
      throw new Error("Intervention d'entretien introuvable.");
    }

    db.prepare(
      `
      UPDATE maintenance_intervention
      SET
        road_id = ?,
        road_key = ?,
        road_code = ?,
        road_designation = ?,
        sap_code = ?,
        degradation_code = ?,
        degradation_name = ?,
        intervention_type = ?,
        status = ?,
        intervention_date = ?,
        completion_date = ?,
        state_before = ?,
        state_after = ?,
        deflection_before = ?,
        deflection_after = ?,
        solution_applied = ?,
        contractor_name = ?,
        responsible_name = ?,
        attachment_path = ?,
        observation = ?,
        cost_amount = ?,
        updated_at = datetime('now')
      WHERE id = ?
    `
    ).run(...values, interventionId);
    targetId = interventionId;
  } else {
    targetId = Number(
      db.prepare(
        `
        INSERT INTO maintenance_intervention (
          road_id,
          road_key,
          road_code,
          road_designation,
          sap_code,
          degradation_code,
          degradation_name,
          intervention_type,
          status,
          intervention_date,
          completion_date,
          state_before,
          state_after,
          deflection_before,
          deflection_after,
          solution_applied,
          contractor_name,
          responsible_name,
          attachment_path,
          observation,
          cost_amount
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `
      ).run(...values).lastInsertRowid
    );
  }

  return listMaintenanceInterventions(db, { roadId, limit: 5000 }).find((item) => item.id === targetId) || null;
}

function deleteMaintenanceIntervention(db, interventionId) {
  const id = Number(interventionId);
  if (!Number.isFinite(id) || id <= 0) {
    throw new Error("Intervention d'entretien invalide.");
  }
  db.prepare("DELETE FROM maintenance_intervention WHERE id = ?").run(id);
  return { deleted: true };
}

function saveDecisionHistory(db, result) {
  const insert = db.prepare(`
    INSERT INTO decision_history (
      road_id,
      road_code,
      road_designation,
      sap_code,
      start_label,
      end_label,
      degradation_name,
      probable_cause,
      maintenance_solution,
      contextual_intervention,
      deflection_value,
      deflection_severity,
      deflection_recommendation,
      drainage_needs_attention,
      drainage_recommendation
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);

  insert.run(
    result.road.id || null,
    toText(result.road.roadCode),
    toText(result.road.designation),
    toText(result.road.sapCode),
    toText(result.road.startLabel),
    toText(result.road.endLabel),
    toText(result.degradation.name),
    toText(result.probableCause),
    toText(result.maintenanceSolution),
    toText(result.contextualIntervention),
    Number.isFinite(result.deflection.value) ? result.deflection.value : null,
    toText(result.deflection.severity),
    toText(result.deflection.recommendation),
    result.drainage.needsAttention ? 1 : 0,
    toText(result.drainage.recommendation)
  );
}

function buildRoadCatalog(db) {
  const rowsFeuil5 = listSheetRows(db, "Feuil5", { limit: 5000 });
  const rowsFeuil2 = listSheetRows(db, "Feuil2", { limit: 5000 });
  const rowsFeuil3 = listSheetRows(db, "Feuil3", { limit: 5000 });
  const rowsFeuil6 = listSheetRows(db, "Feuil6", { limit: 5000 });
  const sapAssignmentsFromFeuil6 = buildFeuil6SapAssignments(rowsFeuil6);

  const roadsByKey = new Map();

  for (const row of rowsFeuil5) {
    const roadCode = canonicalizeRoadCode(row.C);
    const designation = toText(row.D);
    if (!isRoadLabel(roadCode, designation)) {
      continue;
    }
    const key = makeRoadKey(roadCode, designation);
    const road = ensureRoad(roadsByKey, key, roadCode, designation);
    road.sapCode = toText(sapAssignmentsFromFeuil6.get(key)) || road.sapCode || parseSapCode(row.B, row.A);
    road.startLabel = road.startLabel || toText(row.E);
    road.endLabel = road.endLabel || toText(row.F);
    road.lengthM = road.lengthM ?? toNumber(row.G);
    road.widthM = road.widthM ?? toNumber(row.H);
    road.surfaceType = road.surfaceType || toText(row.I);
    road.pavementState = road.pavementState || toText(row.J);
    road.drainageType = road.drainageType || toText(row.K);
    road.drainageState = road.drainageState || toText(row.L);
    road.sidewalkMinM = road.sidewalkMinM ?? toNumber(row.M);
    road.parkingLeft = road.parkingLeft || toText(row.N);
    road.parkingRight = road.parkingRight || toText(row.O);
    road.parkingOther = road.parkingOther || toText(row.P);
  }

  for (const row of rowsFeuil2) {
    const roadCode = canonicalizeRoadCode(row.C);
    const designation = toText(row.D);
    if (!isRoadLabel(roadCode, designation)) {
      continue;
    }
    const key = makeRoadKey(roadCode, designation);
    const road = ensureRoad(roadsByKey, key, roadCode, designation);
    road.sapCode = toText(sapAssignmentsFromFeuil6.get(key)) || road.sapCode || parseSapCode(row.B, row.A);
    road.startLabel = road.startLabel || toText(row.E);
    road.endLabel = road.endLabel || toText(row.F);
    road.lengthM = road.lengthM ?? toNumber(row.G);
  }

  // Feuil3 porte les champs metier "CHAUSSEE" et "ASSAINISSEMENT".
  // On l'applique sur le catalogue pour que l'evaluation utilise ces valeurs.
  for (const row of rowsFeuil3) {
    const roadCode = canonicalizeRoadCode(row.A);
    const designation = toText(row.B);
    if (!isRoadLabel(roadCode, designation)) {
      continue;
    }

    const key = makeRoadKey(roadCode, designation);
    const road = ensureRoad(roadsByKey, key, roadCode, designation);

    road.startLabel = road.startLabel || toText(row.C);
    road.endLabel = road.endLabel || toText(row.D);

    const lengthM = toNumber(row.E);
    if (lengthM !== null) {
      road.lengthM = road.lengthM ?? lengthM;
    }

    const widthM = toNumber(row.F);
    if (widthM !== null) {
      road.widthM = widthM;
    }

    const surfaceType = toText(row.G);
    if (surfaceType) {
      road.surfaceType = surfaceType;
    }

    const pavementState = toText(row.H);
    if (pavementState) {
      road.pavementState = pavementState;
    }

    const drainageType = toText(row.I);
    if (drainageType) {
      road.drainageType = drainageType;
    }

    const drainageState = toText(row.J);
    if (drainageState) {
      road.drainageState = drainageState;
    }

    const sidewalkMinM = toNumber(row.K);
    if (sidewalkMinM !== null) {
      road.sidewalkMinM = sidewalkMinM;
    }

    const interventionHint = toText(row.L);
    if (interventionHint) {
      road.interventionHint = interventionHint;
    }
  }

  let currentSapFromFeuil6 = "";
  for (const row of rowsFeuil6) {
    const sapMarker = parseSapCode(row.A, row.A);
    if (sapMarker) {
      currentSapFromFeuil6 = sapMarker;
    }

    const roadCode = canonicalizeRoadCode(row.C);
    const designation = toText(row.E);
    if (!isRoadLabel(roadCode, designation)) {
      continue;
    }
    const key = makeRoadKey(roadCode, designation);
    const road = ensureRoad(roadsByKey, key, roadCode, designation);
    road.sapCode = toText(currentSapFromFeuil6) || road.sapCode;
    road.designation = road.designation || designation;
    road.lengthM = road.lengthM ?? toNumber(row.D);
    road.itinerary = road.itinerary || toText(row.F);
    road.justification = road.justification || toText(row.G);

    const itineraryBounds = splitItinerary(road.itinerary);
    road.startLabel = road.startLabel || itineraryBounds.startLabel;
    road.endLabel = road.endLabel || itineraryBounds.endLabel;
  }

  const interventionsByDesignation = new Map();
  for (const row of rowsFeuil3) {
    const designation = toText(row.B);
    const intervention = toText(row.L);
    if (!designation || !intervention || /DESIGNATION/i.test(designation)) {
      continue;
    }
    interventionsByDesignation.set(normalizeText(designation), intervention);
  }

  const roads = [...roadsByKey.values()].map((road) => {
    const interventionHint =
      interventionsByDesignation.get(normalizeText(road.designation)) ||
      toText(road.interventionHint) ||
      DEFAULT_INTERVENTION_TEXT;
    return {
      ...road,
      interventionHint
    };
  });

  roads.sort((a, b) => {
    const sapA = Number((a.sapCode.match(/[0-9]+/) || ["99"])[0]);
    const sapB = Number((b.sapCode.match(/[0-9]+/) || ["99"])[0]);
    if (sapA !== sapB) {
      return sapA - sapB;
    }
    return normalizeText(a.roadCode).localeCompare(normalizeText(b.roadCode));
  });

  return roads.map((road, index) => ({
    id: index + 1,
    ...road
  }));
}

function buildRoadAliasCatalog(db, roads) {
  const aliases = new Map();
  const knownRoadKeys = new Set(roads.map((road) => toText(road.roadKey)).filter(Boolean));

  function registerAlias(roadKey, aliasText, aliasType, sourceSheet) {
    const key = toText(roadKey);
    if (!key || !knownRoadKeys.has(key)) {
      return;
    }

    const aliasKey = makeRoadTextAliasKey(aliasText);
    if (!aliasKey) {
      return;
    }

    const uniqueKey = `${key}::${aliasKey}`;
    if (!aliases.has(uniqueKey)) {
      aliases.set(uniqueKey, {
        roadKey: key,
        aliasKey,
        aliasText: toText(aliasText),
        aliasType: toText(aliasType) || "TEXT",
        sourceSheet: toText(sourceSheet)
      });
    }
  }

  function registerBoundsAlias(roadKey, designation, startLabel, endLabel, sourceSheet) {
    const key = toText(roadKey);
    if (!key || !knownRoadKeys.has(key)) {
      return;
    }

    const boundsKey = makeRoadBoundsAliasKey(designation, startLabel, endLabel);
    if (!boundsKey) {
      return;
    }

    const uniqueKey = `${key}::${boundsKey}`;
    if (!aliases.has(uniqueKey)) {
      aliases.set(uniqueKey, {
        roadKey: key,
        aliasKey: boundsKey,
        aliasText: [toText(designation), toText(startLabel), toText(endLabel)].filter(Boolean).join(" | "),
        aliasType: "SECTION_BOUNDS",
        sourceSheet: toText(sourceSheet)
      });
    }
  }

  function processSourceRow(roadCode, designation, startLabel, endLabel, sourceSheet) {
    const normalizedRoadCode = canonicalizeRoadCode(roadCode);
    if (!isRoadLabel(normalizedRoadCode, designation)) {
      return;
    }

    const roadKey = makeRoadKey(normalizedRoadCode, designation);
    registerAlias(roadKey, normalizedRoadCode, "ROAD_CODE", sourceSheet);
    registerAlias(roadKey, designation, "DESIGNATION", sourceSheet);
    registerAlias(roadKey, simplifyRoadDesignation(designation), "SIMPLIFIED_DESIGNATION", sourceSheet);
    registerBoundsAlias(roadKey, designation, startLabel, endLabel, sourceSheet);
  }

  for (const road of roads) {
    registerAlias(road.roadKey, road.roadCode, "ROAD_CODE", "ROAD");
    registerAlias(road.roadKey, road.designation, "DESIGNATION", "ROAD");
    registerAlias(road.roadKey, simplifyRoadDesignation(road.designation), "SIMPLIFIED_DESIGNATION", "ROAD");
    registerBoundsAlias(road.roadKey, road.designation, road.startLabel, road.endLabel, "ROAD");
  }

  for (const row of listSheetRows(db, "Feuil2", { limit: 5000 })) {
    processSourceRow(row.C, row.D, row.E, row.F, "Feuil2");
  }

  for (const row of listSheetRows(db, "Feuil3", { limit: 5000 })) {
    processSourceRow(row.A, row.B, row.C, row.D, "Feuil3");
  }

  for (const row of listSheetRows(db, "Feuil5", { limit: 5000 })) {
    processSourceRow(row.C, row.D, row.E, row.F, "Feuil5");
  }

  for (const row of listSheetRows(db, "Feuil6", { limit: 5000 })) {
    const itineraryBounds = splitItinerary(row.F);
    processSourceRow(row.C, row.E, itineraryBounds.startLabel, itineraryBounds.endLabel, "Feuil6");
  }

  return [...aliases.values()];
}

function buildDegradationCatalog(db) {
  const rows = listSheetRows(db, "Feuil7", { limit: 5000 });
  const map = new Map();
  const templates = getSolutionTemplateMap(db);
  const rules = getDegradationSolutionRuleMap(db);
  const overrides = getDegradationSolutionOverrideMap(db);

  for (let index = 0; index < rows.length; index += 1) {
    const row = rows[index];
    const name = toText(row.C);
    if (!isDegradationLabel(name)) {
      continue;
    }

    const key = normalizeText(name);
    if (!map.has(key)) {
      map.set(key, {
        name,
        causes: []
      });
    }
    const entry = map.get(key);

    const firstCause = toText(row.G);
    if (isCauseLabel(firstCause)) {
      entry.causes.push(firstCause);
    }

    let cursor = index + 1;
    while (cursor < rows.length) {
      const nextRow = rows[cursor];
      if (isDegradationLabel(nextRow.C)) {
        break;
      }
      const causeText = toText(nextRow.G);
      if (isCauseLabel(causeText)) {
        entry.causes.push(causeText);
      }
      cursor += 1;
    }
    index = cursor - 1;
  }

  const items = [...map.values()].map((entry, index) => {
    const degradationCode = normalizeDegradationKey(entry.name);
    const fallbackCause = DEGRADATION_CAUSE_FALLBACKS[degradationCode] || "";
    const uniqueCauses = [...new Set(entry.causes.map((cause) => toText(cause)).filter(Boolean))];
    if (uniqueCauses.length === 0 && fallbackCause) {
      uniqueCauses.push(fallbackCause);
    }
    const resolved = resolveSolutionForDegradation(degradationCode, templates, rules, overrides);

    return {
      id: index + 1,
      code: degradationCode,
      name: entry.name,
      causes: uniqueCauses,
      solution: resolved.solution,
      solutionSource: resolved.source,
      templateKey: resolved.templateKey
    };
  });

  items.sort((a, b) => normalizeText(a.name).localeCompare(normalizeText(b.name)));
  return items;
}

function getSolutionTemplateMap(db) {
  const rows = db
    .prepare(
      `
      SELECT template_key, title, description
      FROM maintenance_solution_template
    `
    )
    .all();

  const map = new Map();
  for (const row of rows) {
    map.set(row.template_key, {
      title: toText(row.title),
      description: toText(row.description)
    });
  }
  return map;
}

function getDegradationSolutionRuleMap(db) {
  const rows = db
    .prepare(
      `
      SELECT degradation_code, template_key
      FROM degradation_solution_assignment
    `
    )
    .all();

  const map = new Map();
  for (const row of rows) {
    map.set(row.degradation_code, row.template_key);
  }
  return map;
}

function getDegradationSolutionOverrideMap(db) {
  const rows = db
    .prepare(
      `
      SELECT degradation_code, solution_text
      FROM degradation_solution_override_rel
    `
    )
    .all();

  const map = new Map();
  for (const row of rows) {
    map.set(row.degradation_code, toText(row.solution_text));
  }
  return map;
}

function resolveSolutionForDegradation(degradationCode, templates, rules, overrides) {
  const overrideText = overrides.get(degradationCode);
  if (overrideText) {
    return {
      solution: overrideText,
      source: "OVERRIDE",
      templateKey: rules.get(degradationCode) || null
    };
  }

  const templateKey = rules.get(degradationCode) || null;
  const template = templateKey ? templates.get(templateKey) : null;
  if (template) {
    return {
      solution: `${template.title}: ${template.description}`,
      source: "TEMPLATE",
      templateKey
    };
  }

  return {
    solution: "Solution a parametrer dans le catalogue de maintenance.",
    source: "MISSING",
    templateKey: null
  };
}

function resolveRoad(roads, payload) {
  const byId = Number(payload.roadId);
  if (Number.isFinite(byId) && byId > 0) {
    return roads.find((road) => road.id === byId) || null;
  }

  const roadKey = normalizeText(payload.roadKey);
  if (roadKey) {
    return roads.find((road) => normalizeText(road.roadKey) === roadKey) || null;
  }

  const roadCode = normalizeRoadCode(payload.roadCode);
  if (roadCode) {
    return roads.find((road) => normalizeRoadCode(road.roadCode) === roadCode) || null;
  }

  const designation = normalizeText(payload.designation);
  if (designation) {
    return roads.find((road) => normalizeText(road.designation) === designation) || null;
  }

  return null;
}

function resolveDegradation(degradations, payload) {
  const id = Number(payload.degradationId);
  if (Number.isFinite(id) && id > 0) {
    return degradations.find((item) => item.id === id) || null;
  }

  const name = normalizeText(payload.degradationName);
  if (!name) {
    return null;
  }

  return degradations.find((item) => normalizeText(item.name) === name) || null;
}

function classifyDeflection(db, value) {
  if (!Number.isFinite(value)) {
    return {
      value: null,
      severity: "NON RENSEIGNE",
      recommendation: "Renseigner une valeur de deflexion (D) pour obtenir le niveau d'intervention."
    };
  }

  const rules = db
    .prepare(
      `
      SELECT
        min_value AS minValue,
        max_value AS maxValue,
        severity,
        recommendation
      FROM deflection_rule
      ORDER BY rule_order
    `
    )
    .all();

  for (const rule of rules) {
    const minOk = rule.minValue === null || value >= Number(rule.minValue);
    const maxOk = rule.maxValue === null || value < Number(rule.maxValue);
    if (minOk && maxOk) {
      return {
        value,
        severity: toText(rule.severity),
        recommendation: toText(rule.recommendation)
      };
    }
  }

  return {
    value,
    severity: "TRES FORT",
    recommendation: "REHABILITATION COUCHE DE ROULEMENT ET DE BASE"
  };
}

function buildDrainageAssessment(db, road, askDrainage) {
  const normalized = normalizeText(`${road.drainageType} ${road.drainageState}`);
  const rules = db
    .prepare(
      `
      SELECT
        id,
        rule_order AS ruleOrder,
        match_operator AS matchOperator,
        pattern,
        ask_required AS askRequired,
        needs_attention AS needsAttention,
        recommendation
      FROM drainage_rule
      WHERE is_active = 1
      ORDER BY rule_order, id
    `
    )
    .all();

  for (const rule of rules) {
    if (Number(rule.askRequired) === 1 && !askDrainage) {
      continue;
    }

    const operator = normalizeDrainageMatchOperator(rule.matchOperator);
    const pattern = toText(rule.pattern);
    let matched = false;

    if (operator === "ALWAYS") {
      matched = true;
    } else if (operator === "CONTAINS") {
      matched = Boolean(pattern) && normalized.includes(normalizeText(pattern));
    } else if (operator === "EQUALS") {
      matched = Boolean(pattern) && normalized === normalizeText(pattern);
    } else if (operator === "REGEX") {
      try {
        const regex = new RegExp(pattern, "i");
        matched = regex.test(normalized);
      } catch {
        matched = false;
      }
    }

    if (matched) {
      return {
        needsAttention: Boolean(rule.needsAttention),
        recommendation: toText(rule.recommendation),
        ruleId: Number(rule.id) || null
      };
    }
  }

  return {
    needsAttention: false,
    recommendation: askDrainage
      ? "Verifier l'etat des caniveaux et programmer un entretien preventif si necessaire."
      : "Aucune alerte assainissement immediate.",
    ruleId: null
  };
}

function ensureRoad(map, key, roadCode, designation) {
  if (!map.has(key)) {
    map.set(key, {
      roadKey: key,
      roadCode,
      designation,
      sapCode: "",
      startLabel: "",
      endLabel: "",
      lengthM: null,
      widthM: null,
      surfaceType: "",
      pavementState: "",
      drainageType: "",
      drainageState: "",
      sidewalkMinM: null,
      parkingLeft: "",
      parkingRight: "",
      parkingOther: "",
      itinerary: "",
      justification: "",
      interventionHint: ""
    });
  }
  return map.get(key);
}

function setUniqueRoadLookup(map, key, row) {
  const normalizedKey = toText(key);
  if (!normalizedKey) {
    return;
  }

  if (!map.has(normalizedKey)) {
    map.set(normalizedKey, row);
    return;
  }

  const existing = map.get(normalizedKey);
  if (existing && Number(existing.id) !== Number(row.id)) {
    map.set(normalizedKey, null);
  }
}

function makeRoadKey(roadCode, designation) {
  const codeKey = normalizeRoadCode(roadCode);
  if (codeKey) {
    return codeKey;
  }
  return `NAME_${normalizeText(designation).replace(/[^A-Z0-9]+/g, "_")}`;
}

function parseSapCode(sectionNo, rowLabel) {
  const fromSection = toText(sectionNo).match(/^([1-9][0-9]?)_/);
  if (fromSection) {
    return `SAP${fromSection[1]}`;
  }
  const fromLabel = normalizeText(rowLabel).match(/SAP\s*([1-9][0-9]?)/);
  if (fromLabel) {
    return `SAP${fromLabel[1]}`;
  }
  return "";
}

function canonicalizeRoadCode(value) {
  const raw = toText(value);
  if (!raw) {
    return "";
  }

  const compact = normalizeText(raw)
    .replace(/BOULEVARD/g, "BVD")
    .replace(/AVENUE/g, "AV")
    .replace(/[^A-Z0-9]/g, "");

  const match = compact.match(/^(RUE|BVD|AV)([0-9]{1,3})$/);
  if (!match) {
    return raw;
  }

  const prefix = match[1] === "BVD" ? "Bvd" : match[1] === "AV" ? "Av" : "Rue";
  const digits = match[2].padStart(2, "0");
  return `${prefix}.${digits}`;
}

function buildFeuil6SapAssignments(rowsFeuil6) {
  const sapByRoadKey = new Map();
  let currentSap = "";

  for (const row of rowsFeuil6) {
    const sapMarker = parseSapCode(row.A, row.A);
    if (sapMarker) {
      currentSap = sapMarker;
    }

    const roadCode = toText(row.C);
    const designation = toText(row.E);
    if (!currentSap || !isRoadLabel(roadCode, designation)) {
      continue;
    }

    sapByRoadKey.set(makeRoadKey(roadCode, designation), currentSap);
  }

  return sapByRoadKey;
}

function alignSectionNoWithSap(sectionNo, sapCode) {
  const text = toText(sectionNo);
  const sectionMatch = text.match(/^([1-9][0-9]?)_(.+)$/);
  const sapMatch = toText(sapCode).match(/^SAP([1-9][0-9]?)$/i);
  if (!sectionMatch || !sapMatch) {
    return text;
  }

  const targetPrefix = sapMatch[1];
  if (sectionMatch[1] === targetPrefix) {
    return text;
  }

  return `${targetPrefix}_${sectionMatch[2]}`;
}

function repairSapAssignmentsFromFeuil6(db) {
  if (!tableExists(db, "sheet_feuil2") || !tableExists(db, "sheet_feuil5") || !tableExists(db, "sheet_feuil6")) {
    return;
  }

  const sapAssignmentsFromFeuil6 = buildFeuil6SapAssignments(listSheetRows(db, "Feuil6", { limit: 5000 }));
  if (sapAssignmentsFromFeuil6.size === 0) {
    return;
  }

  const migratedRoads = new Map();

  for (const sheetName of ["Feuil2", "Feuil5"]) {
    const sheet = resolveSheet(sheetName);
    const rows = listSheetRows(db, sheetName, { limit: 5000 });
    const updateSectionNo = db.prepare(`
      UPDATE ${sheet.table}
      SET col_b = ?, updated_at = datetime('now')
      WHERE id = ?
    `);

    for (const row of rows) {
      const roadCode = toText(row.C);
      const designation = toText(row.D);
      if (!isRoadLabel(roadCode, designation)) {
        continue;
      }

      const roadKey = makeRoadKey(roadCode, designation);
      const authoritativeSap = toText(sapAssignmentsFromFeuil6.get(roadKey));
      if (!authoritativeSap) {
        continue;
      }

      const nextSectionNo = alignSectionNoWithSap(row.B, authoritativeSap);
      if (!nextSectionNo || nextSectionNo === toText(row.B)) {
        continue;
      }

      updateSectionNo.run(nextSectionNo, Number(row.id));
      migratedRoads.set(roadKey, {
        roadCode,
        designation,
        sapCode: authoritativeSap
      });
    }
  }

  if (migratedRoads.size === 0) {
    return;
  }

  if (tableExists(db, "decision_history")) {
    const updateDecisionHistoryByCode = db.prepare(`
      UPDATE decision_history
      SET sap_code = ?
      WHERE road_code = ? AND COALESCE(sap_code, '') <> ?
    `);
    const updateDecisionHistoryByDesignation = db.prepare(`
      UPDATE decision_history
      SET sap_code = ?
      WHERE road_code = '' AND road_designation = ? AND COALESCE(sap_code, '') <> ?
    `);

    for (const road of migratedRoads.values()) {
      if (toText(road.roadCode)) {
        updateDecisionHistoryByCode.run(road.sapCode, road.roadCode, road.sapCode);
      } else if (toText(road.designation)) {
        updateDecisionHistoryByDesignation.run(road.sapCode, road.designation, road.sapCode);
      }
    }
  }

  if (tableExists(db, "maintenance_intervention")) {
    const updateMaintenanceByCode = db.prepare(`
      UPDATE maintenance_intervention
      SET sap_code = ?, updated_at = datetime('now')
      WHERE road_code = ? AND COALESCE(sap_code, '') <> ?
    `);
    const updateMaintenanceByDesignation = db.prepare(`
      UPDATE maintenance_intervention
      SET sap_code = ?, updated_at = datetime('now')
      WHERE road_code = '' AND road_designation = ? AND COALESCE(sap_code, '') <> ?
    `);

    for (const road of migratedRoads.values()) {
      if (toText(road.roadCode)) {
        updateMaintenanceByCode.run(road.sapCode, road.roadCode, road.sapCode);
      } else if (toText(road.designation)) {
        updateMaintenanceByDesignation.run(road.sapCode, road.designation, road.sapCode);
      }
    }
  }
}

function repairRoadCodesInSourceSheets(db) {
  const sheetSpecs = [
    { name: "Feuil2", codeColumn: "C", designationColumn: "D" },
    { name: "Feuil3", codeColumn: "A", designationColumn: "B" },
    { name: "Feuil5", codeColumn: "C", designationColumn: "D" },
    { name: "Feuil6", codeColumn: "C", designationColumn: "E" }
  ];

  const updatedCodes = new Map();

  for (const spec of sheetSpecs) {
    const sheet = resolveSheet(spec.name);
    const codeDbColumn = DB_COLUMN_MAP[spec.codeColumn];
    const rows = listSheetRows(db, spec.name, { limit: 5000 });
    const updateCode = db.prepare(`
      UPDATE ${sheet.table}
      SET ${codeDbColumn} = ?, updated_at = datetime('now')
      WHERE id = ?
    `);

    for (const row of rows) {
      const currentCode = toText(row[spec.codeColumn]);
      const designation = toText(row[spec.designationColumn]);
      const nextCode = canonicalizeRoadCode(currentCode);
      if (!nextCode || nextCode === currentCode || !isRoadLabel(nextCode, designation)) {
        continue;
      }

      updateCode.run(nextCode, Number(row.id));
      updatedCodes.set(`${spec.name}:${Number(row.id)}`, { currentCode, nextCode });
    }
  }

  if (updatedCodes.size === 0) {
    return;
  }

  const repairTableRoadCode = (tableName) => {
    if (!tableExists(db, tableName)) {
      return;
    }

    const rows = db.prepare(`SELECT id, road_code AS roadCode FROM ${tableName}`).all();
    const setClause = tableHasColumn(db, tableName, "updated_at")
      ? "road_code = ?, updated_at = datetime('now')"
      : "road_code = ?";
    const update = db.prepare(`UPDATE ${tableName} SET ${setClause} WHERE id = ?`);

    for (const row of rows) {
      const currentCode = toText(row.roadCode);
      const nextCode = canonicalizeRoadCode(currentCode);
      if (!nextCode || nextCode === currentCode) {
        continue;
      }
      update.run(nextCode, Number(row.id));
    }
  };

  for (const tableName of [
    "road",
    "road_section",
    "measurement_campaign",
    "road_measurement",
    "decision_history",
    "maintenance_intervention"
  ]) {
    repairTableRoadCode(tableName);
  }
}

function normalizeRoadAliasValue(value) {
  return toText(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/BOULEVARD/g, "BVD")
    .replace(/AVENUE/g, "AV")
    .replace(/[().,;:_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function makeRoadTextAliasKey(value) {
  const normalized = normalizeRoadAliasValue(simplifyRoadDesignation(value) || value);
  return normalized ? `TEXT:${normalized}` : "";
}

function makeRoadBoundsLookupKey(designation, startLabel, endLabel) {
  const designationKey = normalizeRoadAliasValue(simplifyRoadDesignation(designation));
  const startKey = normalizeRoadAliasValue(startLabel);
  const endKey = normalizeRoadAliasValue(endLabel);
  if (!designationKey || !startKey || !endKey) {
    return "";
  }
  return `BOUNDS:${designationKey}|${startKey}|${endKey}`;
}

function makeRoadBoundsAliasKey(designation, startLabel, endLabel) {
  return makeRoadBoundsLookupKey(designation, startLabel, endLabel);
}

function splitItinerary(itinerary) {
  const text = toText(itinerary);
  if (!text) {
    return { startLabel: "", endLabel: "" };
  }
  const itineraryMatch = text.match(/^(?:DE\s+)?(.+?)\s+[AÀ]\s+(.+)$/i);
  if (itineraryMatch) {
    return {
      startLabel: toText(itineraryMatch[1]),
      endLabel: toText(itineraryMatch[2])
    };
  }
  const slashMatch = text.match(/^(.+?)\s*\/\s*(.+)$/);
  if (slashMatch) {
    return {
      startLabel: toText(slashMatch[1]),
      endLabel: toText(slashMatch[2])
    };
  }
  return { startLabel: text, endLabel: "" };
}

function isRoadLabel(roadCode, designation) {
  const code = normalizeText(roadCode);
  const label = normalizeText(designation);
  if (!code || !label) {
    return false;
  }
  if (["VOIES", "DESIGNATION", "TYPE DE VOIE", "N", "N TRONCON", "N SECTIONS"].includes(code)) {
    return false;
  }
  return /^(RUE|BVD|AV|BOULEVARD|AVENUE)/.test(code);
}

function simplifyRoadDesignation(value) {
  let text = toText(value);
  if (!text) {
    return "";
  }

  text = text.replace(/\s+/g, " ").trim();
  text = text.replace(/\([^)]*\)/g, " ").replace(/\s+/g, " ").trim();
  text = text.replace(/\bTRON[CÇ]ON\b.*$/i, "").trim();
  text = text.replace(/\bDU\s+PK\b.*$/i, "").trim();

  const roadMatch = text.match(/((?:RUE|BVD|BOULEVARD|AV|AVENUE)[^,;]*)/i);
  if (roadMatch) {
    text = roadMatch[1].trim();
  }

  return text;
}

function isRoadDesignationText(value) {
  const text = simplifyRoadDesignation(value);
  if (!text || isLikelyPkValue(text)) {
    return false;
  }
  return /(?:^|\s)(RUE|BVD|BOULEVARD|AV|AVENUE)(?:\s|$)/i.test(text);
}

function isLikelyPkValue(value) {
  const text = toText(value).replace(/\s+/g, "");
  if (!text) {
    return false;
  }
  return /^-?\d+(?:[.,]\d+)?$/.test(text);
}

function isDegradationLabel(value) {
  const text = toText(value);
  const normalized = normalizeText(text);
  if (!text || normalized.length < 4) {
    return false;
  }
  if (/[.;:]/.test(text) || text.length > 80) {
    return false;
  }
  if (["DEGRADATIONS", "CAUSES", "OBSERVATIONS"].includes(normalized)) {
    return false;
  }

  const lettersOnly = text.replace(/[^A-Za-z]/g, "");
  if (!lettersOnly) {
    return false;
  }
  const uppercaseOnly = lettersOnly.replace(/[^A-Z]/g, "");
  const ratio = uppercaseOnly.length / lettersOnly.length;
  return ratio >= 0.5;
}

function isCauseLabel(value) {
  const text = toText(value);
  if (!text) {
    return false;
  }
  if (text.length < 8) {
    return false;
  }
  const normalized = normalizeText(text);
  return !["CAUSE", "CAUSES"].includes(normalized);
}

function normalizeText(value) {
  return toText(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase();
}

function normalizeDegradationKey(value) {
  return normalizeText(value).replace(/[^A-Z0-9]+/g, "_");
}

function normalizeDrainageMatchOperator(value) {
  const normalized = normalizeText(value);
  if (["CONTAINS", "EQUALS", "REGEX", "ALWAYS"].includes(normalized)) {
    return normalized;
  }
  return "CONTAINS";
}

function normalizeMaintenanceStatus(value) {
  const normalized = normalizeText(value).replace(/[\s-]+/g, "_");
  if (["PREVU", "EN_COURS", "TERMINE"].includes(normalized)) {
    return normalized;
  }
  return "";
}

function normalizeRoadCode(value) {
  return normalizeText(canonicalizeRoadCode(value))
    .replace(/BOULEVARD/g, "BVD")
    .replace(/AVENUE/g, "AV")
    .replace(/[^A-Z0-9]/g, "");
}

function toNumber(value) {
  const text = toText(value);
  if (!text) {
    return null;
  }
  const normalized = text.replace(",", ".");
  const num = Number(normalized);
  return Number.isFinite(num) ? num : null;
}

function setMeta(db, key, value) {
  db.prepare(
    "INSERT INTO app_meta (meta_key, meta_value) VALUES (?, ?) ON CONFLICT(meta_key) DO UPDATE SET meta_value = excluded.meta_value"
  ).run(key, value);
}

function resetAutoincrementSequence(db, tableName) {
  if (!tableExists(db, tableName)) {
    return;
  }
  db.prepare("DELETE FROM sqlite_sequence WHERE name = ?").run(tableName);
}

function getSheetRows(workbook, sheetRef) {
  const sheetName = typeof sheetRef === "string" ? sheetRef : sheetRef?.name;
  const sheetDef =
    typeof sheetRef === "string" ? SHEET_DEFINITIONS.find((item) => item.name === sheetRef) || null : sheetRef;
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    return [];
  }

  const expectedColumns = sheetDef?.columns || COLUMN_KEYS;
  const rawRows = XLSX.utils.sheet_to_json(sheet, {
    header: "A",
    defval: "",
    raw: false,
    blankrows: false
  });

  const parsedRows = [];
  for (const rawRow of rawRows) {
    const row = {};
    let hasValue = false;

    for (const column of expectedColumns) {
      const value = toText(rawRow[column]);
      row[column] = value;
      if (value) {
        hasValue = true;
      }
    }

    if (!hasValue) {
      continue;
    }
    parsedRows.push(row);
  }

  return parsedRows;
}

function toText(value) {
  return String(value ?? "").trim();
}

function safeJson(value) {
  try {
    return JSON.stringify(value ?? {});
  } catch {
    return "";
  }
}

module.exports = {
  setupDataLayer
};
