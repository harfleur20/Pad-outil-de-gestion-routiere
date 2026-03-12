PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS sap_sector (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  code TEXT NOT NULL UNIQUE,
  name TEXT NOT NULL,
  description TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS road (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  sap_sector_id INTEGER NOT NULL,
  road_code TEXT NOT NULL,
  road_type TEXT,
  name TEXT NOT NULL,
  start_label TEXT,
  end_label TEXT,
  length_m REAL,
  width_m REAL,
  surface_type TEXT,
  pavement_state TEXT,
  drainage_type TEXT,
  drainage_state TEXT,
  sidewalk_left_m REAL,
  sidewalk_right_m REAL,
  parking_left INTEGER NOT NULL DEFAULT 0,
  parking_right INTEGER NOT NULL DEFAULT 0,
  notes TEXT,
  is_active INTEGER NOT NULL DEFAULT 1,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (sap_sector_id) REFERENCES sap_sector(id),
  UNIQUE (sap_sector_id, road_code)
);

CREATE TABLE IF NOT EXISTS degradation (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  code TEXT NOT NULL UNIQUE,
  name TEXT NOT NULL,
  description TEXT,
  is_active INTEGER NOT NULL DEFAULT 1,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS degradation_cause (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  degradation_id INTEGER NOT NULL,
  cause_text TEXT NOT NULL,
  weight INTEGER,
  is_active INTEGER NOT NULL DEFAULT 1,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (degradation_id) REFERENCES degradation(id)
);

CREATE TABLE IF NOT EXISTS maintenance_solution (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  degradation_id INTEGER NOT NULL,
  title TEXT NOT NULL,
  description TEXT NOT NULL,
  intervention_level TEXT,
  estimated_duration_days INTEGER,
  is_active INTEGER NOT NULL DEFAULT 1,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (degradation_id) REFERENCES degradation(id)
);

CREATE TABLE IF NOT EXISTS road_observation (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  road_id INTEGER NOT NULL,
  degradation_id INTEGER NOT NULL,
  observed_at TEXT NOT NULL,
  observed_by TEXT,
  severity_level TEXT,
  deflection_value REAL,
  pavement_note TEXT,
  drainage_note TEXT,
  attachments_json TEXT,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (road_id) REFERENCES road(id),
  FOREIGN KEY (degradation_id) REFERENCES degradation(id)
);

CREATE TABLE IF NOT EXISTS recommendation (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  observation_id INTEGER NOT NULL,
  primary_cause_id INTEGER,
  primary_solution_id INTEGER,
  recommended_action TEXT NOT NULL,
  assainissement_action TEXT,
  confidence_score REAL,
  rule_version TEXT NOT NULL DEFAULT 'v1',
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (observation_id) REFERENCES road_observation(id),
  FOREIGN KEY (primary_cause_id) REFERENCES degradation_cause(id),
  FOREIGN KEY (primary_solution_id) REFERENCES maintenance_solution(id)
);

CREATE INDEX IF NOT EXISTS idx_road_sap ON road (sap_sector_id);
CREATE INDEX IF NOT EXISTS idx_obs_road ON road_observation (road_id);
CREATE INDEX IF NOT EXISTS idx_obs_degradation ON road_observation (degradation_id);
CREATE INDEX IF NOT EXISTS idx_cause_degradation ON degradation_cause (degradation_id);
CREATE INDEX IF NOT EXISTS idx_solution_degradation ON maintenance_solution (degradation_id);
