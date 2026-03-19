const fs = require("fs");
const os = require("os");
const path = require("path");
const Database = require("better-sqlite3");
const { setupDataLayer } = require("../app/electron/services/data-layer.cjs");

const DEFAULT_OUTPUT_PATH = path.join(__dirname, "..", "app", "electron", "assets", "pad-maintenance.seed.db");

function getArgValue(flagName) {
  const index = process.argv.indexOf(flagName);
  if (index >= 0 && process.argv[index + 1]) {
    return process.argv[index + 1];
  }
  return "";
}

function resolveSourceExcelPath() {
  const cliValue = getArgValue("--source");
  if (cliValue && fs.existsSync(cliValue)) {
    return cliValue;
  }

  const envValue = process.env.PAD_EXCEL_PATH;
  if (envValue && fs.existsSync(envValue)) {
    return envValue;
  }

  const currentDbPath = path.join(
    process.env.APPDATA || path.join(os.homedir(), "AppData", "Roaming"),
    "projet-pad",
    "data",
    "pad-maintenance.db"
  );
  if (fs.existsSync(currentDbPath)) {
    try {
      const db = new Database(currentDbPath, { readonly: true });
      const row = db
        .prepare("SELECT meta_value AS value FROM app_meta WHERE meta_key = 'last_import_path' LIMIT 1")
        .get();
      db.close();
      const candidate = String(row?.value || "").trim();
      if (candidate && fs.existsSync(candidate)) {
        return candidate;
      }
    } catch {}
  }

  const candidates = [
    "C:\\Users\\harfl\\OneDrive\\Desktop\\pad\\programme ayissi.xlsx",
    "C:\\Users\\harfl\\OneDrive\\Desktop\\programme ayissi.xlsx"
  ];
  for (const candidate of candidates) {
    if (fs.existsSync(candidate)) {
      return candidate;
    }
  }
  return "";
}

function resolveOutputPath() {
  const cliValue = getArgValue("--output");
  return cliValue ? path.resolve(cliValue) : DEFAULT_OUTPUT_PATH;
}

function upsertMeta(db, key, value) {
  db.prepare(
    `
    INSERT INTO app_meta (meta_key, meta_value)
    VALUES (?, ?)
    ON CONFLICT(meta_key) DO UPDATE SET
      meta_value = excluded.meta_value
  `
  ).run(key, value);
}

function cleanSeedDatabase(db) {
  db.prepare("DELETE FROM decision_history").run();
  db.prepare("DELETE FROM maintenance_intervention").run();
  db.prepare("DELETE FROM sqlite_sequence WHERE name IN ('decision_history', 'maintenance_intervention')").run();
  upsertMeta(db, "last_import_path", "Base métier embarquée");
  upsertMeta(db, "last_import_at", new Date().toISOString());
}

function escapeSqliteLiteral(value) {
  return String(value).replace(/'/g, "''");
}

function safeRemoveDirectory(targetPath) {
  try {
    fs.rmSync(targetPath, { recursive: true, force: true });
  } catch (error) {
    console.warn(`[PAD] Nettoyage temporaire ignoré: ${error.message}`);
  }
}

Promise.resolve().then(() => {
  const sourceExcelPath = resolveSourceExcelPath();
  if (!sourceExcelPath) {
    throw new Error("Aucun fichier Excel source n'a été trouvé pour construire la base métier.");
  }

  const outputPath = resolveOutputPath();
  const buildRoot = path.join(__dirname, "..", ".tmp-seed-build");
  fs.rmSync(buildRoot, { recursive: true, force: true });
  fs.mkdirSync(buildRoot, { recursive: true });

  const fakeApp = {
    getPath(name) {
      if (name === "userData") {
        return buildRoot;
      }
      return buildRoot;
    }
  };

  const dataLayer = setupDataLayer({ app: fakeApp });
  dataLayer.importFromExcel(sourceExcelPath);

  const tempDbPath = path.join(buildRoot, "data", "pad-maintenance.db");
  const tempDb = new Database(tempDbPath);
  cleanSeedDatabase(tempDb);

  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  if (fs.existsSync(outputPath)) {
    fs.unlinkSync(outputPath);
  }
  tempDb.exec(`VACUUM INTO '${escapeSqliteLiteral(outputPath)}'`);
  tempDb.close();

  console.log(`[PAD] Base métier générée: ${outputPath}`);
  console.log(`[PAD] Source Excel: ${sourceExcelPath}`);

  safeRemoveDirectory(buildRoot);
  process.exit(0);
}).catch((error) => {
  console.error("[PAD] Échec génération base métier:", error);
  safeRemoveDirectory(path.join(__dirname, "..", ".tmp-seed-build"));
  process.exitCode = 1;
});
