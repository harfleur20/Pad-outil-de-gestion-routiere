const fs = require("fs");
const path = require("path");
const { setupDataLayer } = require("../app/electron/services/data-layer.cjs");

const buildRoot = path.join(__dirname, "..", ".tmp-seed-smoke");
const seedPath = path.join(__dirname, "..", "app", "electron", "assets", "pad-maintenance.seed.db");

process.env.PAD_SEED_DB_PATH = seedPath;

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

try {
  const dataLayer = setupDataLayer({ app: fakeApp });
  const status = dataLayer.getDataStatus();
  const roads = dataLayer.listRoadCatalog({});
  const sections = dataLayer.listRoadSections({});

  console.log(
    JSON.stringify(
      {
        totalRows: status.totalRows,
        decisionHistoryCount: status.decisionHistoryCount,
        lastImportPath: status.lastImportPath,
        roads: roads.length,
        sections: sections.length,
        hasRue34: roads.some((item) => item.roadCode === "Rue.34"),
        hasSap5: roads.some((item) => item.sapCode === "SAP5")
      },
      null,
      2
    )
  );
  process.exit(0);
} catch (error) {
  console.error("[PAD] Smoke seed bootstrap failed:", error);
  process.exit(1);
}
