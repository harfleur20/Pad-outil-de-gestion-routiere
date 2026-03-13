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
    description: "Noms proposes et itineraires",
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
    F: "Itineraire",
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

function setupDataLayer({ app }) {
  const dataDir = path.join(app.getPath("userData"), "data");
  fs.mkdirSync(dataDir, { recursive: true });

  const dbPath = path.join(dataDir, "pad-maintenance.db");
  const db = new Database(dbPath);
  db.pragma("journal_mode = WAL");
  db.pragma("foreign_keys = ON");

  preEnsureLegacyColumns(db);
  ensureSchema(db);
  migrateLegacySchema(db);
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
    importFromExcel: (excelPath) => importFromExcelInternal(db, excelPath || resolveDefaultExcelPath()),
    listSheetRows: (sheetName, filters) => listSheetRows(db, sheetName, filters),
    createSheetRow: (sheetName, payload) => createSheetRow(db, sheetName, payload),
    updateSheetRow: (sheetName, rowId, payload) => updateSheetRow(db, sheetName, rowId, payload),
    deleteSheetRow: (sheetName, rowId) => deleteSheetRow(db, sheetName, rowId),
    listRoadCatalog: (filters) => listRoadCatalog(db, filters),
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
    evaluateDecision: (payload) => evaluateDecision(db, payload),
    listDecisionHistory: (filters) => listDecisionHistory(db, filters),
    clearDecisionHistory: () => clearDecisionHistory(db)
  };
}

function preEnsureLegacyColumns(db) {
  ensureColumnIfMissing(db, "road", "sap_code", "sap_code TEXT");
  ensureColumnIfMissing(db, "road", "road_code", "road_code TEXT NOT NULL DEFAULT ''");
  ensureColumnIfMissing(db, "decision_history", "sap_code", "sap_code TEXT NOT NULL DEFAULT ''");
  ensureColumnIfMissing(db, "degradation", "degradation_code", "degradation_code TEXT NOT NULL DEFAULT ''");
  ensureColumnIfMissing(db, "degradation", "name", "name TEXT NOT NULL DEFAULT ''");
  ensureColumnIfMissing(db, "degradation_cause", "degradation_code", "degradation_code TEXT NOT NULL DEFAULT ''");
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
      "road_id INTEGER",
      "road_key TEXT NOT NULL DEFAULT ''",
      "road_code TEXT NOT NULL DEFAULT ''",
      "designation TEXT NOT NULL DEFAULT ''",
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

    CREATE TABLE IF NOT EXISTS road_measurement (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      measurement_key TEXT NOT NULL UNIQUE,
      source_sheet TEXT NOT NULL DEFAULT 'Feuil1',
      source_row_no INTEGER NOT NULL DEFAULT 0,
      road_id INTEGER,
      road_key TEXT NOT NULL DEFAULT '',
      road_code TEXT NOT NULL DEFAULT '',
      designation TEXT NOT NULL DEFAULT '',
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
  const roads = buildRoadCatalog(db);
  const degradationItems = buildDegradationCatalog(db);
  const roadSections = buildRoadSectionCatalog(db);
  const roadMeasurements = buildRoadMeasurementCatalog(db);
  const decisionInputs = buildDecisionProfileInputs(db);
  const degradationDefinitions = buildDegradationDefinitions(db);

  const tx = db.transaction(() => {
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

    const roadLookup = buildRoadLookupMaps(db);
    syncRoadSections(db, roadSections, roadLookup);
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
  });

  tx();
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

  const byKey = new Map();
  const byCode = new Map();
  const byDesignation = new Map();

  for (const row of rows) {
    const key = toText(row.roadKey);
    const code = normalizeRoadCode(row.roadCode);
    const designation = normalizeText(row.designation);

    if (key && !byKey.has(key)) {
      byKey.set(key, row);
    }
    if (code && !byCode.has(code)) {
      byCode.set(code, row);
    }
    if (designation && !byDesignation.has(designation)) {
      byDesignation.set(designation, row);
    }
  }

  return { byKey, byCode, byDesignation };
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
    const sapCode = toText(section.sapCode) || toText(roadRef?.sapCode);
    const sectionKey = toText(section.sectionKey);
    keepKeys.push(sectionKey);

    upsert.run(
      sectionKey,
      toText(section.sourceSheet),
      Number(section.sourceRowNo) || 0,
      toText(section.tronconNo),
      toText(section.sectionNo),
      toText(section.roadKey),
      Number.isFinite(roadId) ? roadId : null,
      sapCode,
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

function syncRoadMeasurements(db, measurements, roadLookup) {
  const upsert = db.prepare(`
    INSERT INTO road_measurement (
      measurement_key,
      source_sheet,
      source_row_no,
      road_id,
      road_key,
      road_code,
      designation,
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
    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    ON CONFLICT(measurement_key) DO UPDATE SET
      source_sheet = excluded.source_sheet,
      source_row_no = excluded.source_row_no,
      road_id = excluded.road_id,
      road_key = excluded.road_key,
      road_code = excluded.road_code,
      designation = excluded.designation,
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
      Number.isFinite(roadId) ? roadId : null,
      toText(item.roadKey),
      toText(item.roadCode),
      toText(item.designation),
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
      WHERE measurement_key NOT IN (${keepKeys.map(() => "?").join(", ")})
    `
    ).run(...keepKeys);
  } else {
    db.prepare("DELETE FROM road_measurement").run();
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

  const sections = [];

  for (const row of rowsFeuil2) {
    const roadCode = toText(row.C);
    const designation = toText(row.D);
    if (!isRoadLabel(roadCode, designation)) {
      continue;
    }
    const roadKey = makeRoadKey(roadCode, designation);
    sections.push({
      sectionKey: `FEUIL2:${row.rowNo}:${roadKey}`,
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
    const roadCode = toText(row.C);
    const designation = toText(row.D);
    if (!isRoadLabel(roadCode, designation)) {
      continue;
    }
    const roadKey = makeRoadKey(roadCode, designation);
    sections.push({
      sectionKey: `FEUIL5:${row.rowNo}:${roadKey}`,
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
    const roadCode = toText(row.A);
    const designation = toText(row.B);
    if (!isRoadLabel(roadCode, designation)) {
      continue;
    }
    const roadKey = makeRoadKey(roadCode, designation);
    sections.push({
      sectionKey: `FEUIL3:${row.rowNo}:${roadKey}`,
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

  return sections;
}

function buildRoadMeasurementCatalog(db) {
  const rows = listSheetRows(db, "Feuil1", { limit: 10000 });
  const items = [];
  let currentDesignation = "";

  for (const row of rows) {
    const headerCandidate = toText(row.A) || toText(row.E);
    if (isRoadDesignationText(headerCandidate)) {
      currentDesignation = headerCandidate;
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

    const designation = currentDesignation;
    const roadKey = designation ? makeRoadKey("", designation) : "";
    const pkM = toNumber(pkDeflection) ?? toNumber(pkLecture);

    items.push({
      measurementKey: `FEUIL1:${row.rowNo}`,
      sourceSheet: "Feuil1",
      sourceRowNo: Number(row.rowNo) || 0,
      roadKey,
      roadCode: "",
      designation,
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

  return items;
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

function createSheetRow(db, sheetName, payload = {}) {
  const sheet = resolveSheet(sheetName);
  const explicitRowNo = Number(payload.rowNo);
  const maxRowNo = db.prepare(`SELECT COALESCE(MAX(row_no), 0) AS maxRowNo FROM ${sheet.table}`).get().maxRowNo;
  const rowNo = Number.isFinite(explicitRowNo) && explicitRowNo > 0 ? explicitRowNo : maxRowNo + 1;
  const values = COLUMN_KEYS.map((key) => toText(payload[key]));

  const insert = db.prepare(`
    INSERT INTO ${sheet.table} (
      row_no,
      ${DB_COLUMN_KEYS.join(", ")}
    ) VALUES (
      ?,
      ${DB_COLUMN_KEYS.map(() => "?").join(", ")}
    )
  `);

  const insertedId = insert.run(rowNo, ...values).lastInsertRowid;
  rebuildNormalizedCatalogs(db);
  return getSheetRowById(db, sheet, insertedId);
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

  const values = COLUMN_KEYS.map((key) => {
    if (Object.prototype.hasOwnProperty.call(payload, key)) {
      return toText(payload[key]);
    }
    return toText(current[key]);
  });

  const update = db.prepare(`
    UPDATE ${sheet.table}
    SET
      row_no = ?,
      ${DB_COLUMN_KEYS.map((key) => `${key} = ?`).join(",\n      ")},
      updated_at = datetime('now')
    WHERE id = ?
  `);

  update.run(rowNo, ...values, id);
  rebuildNormalizedCatalogs(db);
  return getSheetRowById(db, sheet, id);
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

  const roadsByKey = new Map();

  for (const row of rowsFeuil5) {
    const roadCode = toText(row.C);
    const designation = toText(row.D);
    if (!isRoadLabel(roadCode, designation)) {
      continue;
    }
    const key = makeRoadKey(roadCode, designation);
    const road = ensureRoad(roadsByKey, key, roadCode, designation);
    road.sapCode = road.sapCode || parseSapCode(row.B, row.A);
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
    const roadCode = toText(row.C);
    const designation = toText(row.D);
    if (!isRoadLabel(roadCode, designation)) {
      continue;
    }
    const key = makeRoadKey(roadCode, designation);
    const road = ensureRoad(roadsByKey, key, roadCode, designation);
    road.sapCode = road.sapCode || parseSapCode(row.B, row.A);
    road.startLabel = road.startLabel || toText(row.E);
    road.endLabel = road.endLabel || toText(row.F);
    road.lengthM = road.lengthM ?? toNumber(row.G);
  }

  // Feuil3 porte les champs metier "CHAUSSEE" et "ASSAINISSEMENT".
  // On l'applique sur le catalogue pour que l'evaluation utilise ces valeurs.
  for (const row of rowsFeuil3) {
    const roadCode = toText(row.A);
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

    const roadCode = toText(row.C);
    const designation = toText(row.E);
    if (!isRoadLabel(roadCode, designation)) {
      continue;
    }
    const key = makeRoadKey(roadCode, designation);
    const road = ensureRoad(roadsByKey, key, roadCode, designation);
    road.sapCode = road.sapCode || currentSapFromFeuil6;
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
    const uniqueCauses = [...new Set(entry.causes.map((cause) => toText(cause)).filter(Boolean))];
    const degradationCode = normalizeDegradationKey(entry.name);
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

function splitItinerary(itinerary) {
  const text = toText(itinerary);
  const match = text.match(/^(?:DE\s+)?(.+?)\s+[AÀ]\s+(.+)$/i);
  if (!match) {
    return { startLabel: "", endLabel: "" };
  }
  return {
    startLabel: toText(match[1]),
    endLabel: toText(match[2])
  };
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

function isRoadDesignationText(value) {
  const text = toText(value);
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

function normalizeRoadCode(value) {
  return normalizeText(value)
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
