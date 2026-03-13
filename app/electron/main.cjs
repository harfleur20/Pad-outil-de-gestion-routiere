const { app, BrowserWindow, ipcMain, dialog } = require("electron");
const path = require("path");

const isDev = !app.isPackaged;
let dataLayer = null;
const appIconPath = path.join(__dirname, "assets", "icon.ico");

if (process.platform === "win32") {
  app.setAppUserModelId("cm.pad.maintenance");
}

function createMainWindow() {
  const win = new BrowserWindow({
    width: 1320,
    height: 860,
    minWidth: 1100,
    minHeight: 720,
    backgroundColor: "#08214f",
    icon: appIconPath,
    webPreferences: {
      preload: path.join(__dirname, "preload.cjs"),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  if (isDev) {
    win.loadURL("http://localhost:5173");
    win.webContents.openDevTools({ mode: "detach" });
  } else {
    win.loadFile(path.join(__dirname, "../web/dist/index.html"));
  }
}

function registerIpcHandlers() {
  ipcMain.handle("data:status", () => dataLayer.getDataStatus());
  ipcMain.handle("data:importFromExcel", (_event, excelPath) => dataLayer.importFromExcel(excelPath));
  ipcMain.handle("data:pickExcelFile", async () => {
    const result = await dialog.showOpenDialog({
      title: "Selectionner un fichier Excel",
      properties: ["openFile"],
      filters: [
        { name: "Excel", extensions: ["xlsx", "xlsm", "xls"] },
        { name: "Tous les fichiers", extensions: ["*"] }
      ]
    });
    if (result.canceled || result.filePaths.length === 0) {
      return null;
    }
    return result.filePaths[0];
  });
  ipcMain.handle("audit:integrity", () => dataLayer.getDataIntegrityReport());

  ipcMain.handle("sheet:definitions", () => dataLayer.listSheetDefinitions());
  ipcMain.handle("sheet:list", (_event, sheetName, filters) => dataLayer.listSheetRows(sheetName, filters));
  ipcMain.handle("sheet:create", (_event, sheetName, payload) => dataLayer.createSheetRow(sheetName, payload));
  ipcMain.handle("sheet:update", (_event, sheetName, rowId, payload) =>
    dataLayer.updateSheetRow(sheetName, rowId, payload)
  );
  ipcMain.handle("sheet:delete", (_event, sheetName, rowId) => dataLayer.deleteSheetRow(sheetName, rowId));

  ipcMain.handle("sap:list", () => dataLayer.listSapSectors());
  ipcMain.handle("roads:list", (_event, filters) => dataLayer.listRoadCatalog(filters));
  ipcMain.handle("degradations:list", () => dataLayer.listDegradationCatalog());
  ipcMain.handle("drainageRules:list", () => dataLayer.listDrainageRules());
  ipcMain.handle("drainageRules:upsert", (_event, payload) => dataLayer.upsertDrainageRule(payload));
  ipcMain.handle("drainageRules:delete", (_event, ruleId) => dataLayer.deleteDrainageRule(ruleId));
  ipcMain.handle("solutions:listTemplates", () => dataLayer.listSolutionTemplates());
  ipcMain.handle("solutions:upsertTemplate", (_event, payload) => dataLayer.upsertSolutionTemplate(payload));
  ipcMain.handle("solutions:assignTemplate", (_event, degradationCode, templateKey) =>
    dataLayer.assignTemplateToDegradation(degradationCode, templateKey)
  );
  ipcMain.handle("solutions:setOverride", (_event, degradationCode, solutionText) =>
    dataLayer.setDegradationSolutionOverride(degradationCode, solutionText)
  );
  ipcMain.handle("solutions:clearOverride", (_event, degradationCode) =>
    dataLayer.clearDegradationSolutionOverride(degradationCode)
  );
  ipcMain.handle("decision:evaluate", (_event, payload) => dataLayer.evaluateDecision(payload));
  ipcMain.handle("reporting:listHistory", (_event, filters) => dataLayer.listDecisionHistory(filters));
  ipcMain.handle("reporting:clearHistory", () => dataLayer.clearDecisionHistory());
}

function getStartupErrorMessage(error) {
  const details = error instanceof Error ? error.message : String(error);
  const nodeModuleMismatch = details.includes("NODE_MODULE_VERSION") || details.includes("better_sqlite3.node");

  if (nodeModuleMismatch) {
    return [
      "Le module SQLite n'est pas compatible avec la version Electron actuelle.",
      "",
      "Execute ces commandes dans le projet:",
      "1) npm install",
      "2) npm run rebuild:electron",
      "3) npm run dev",
      "",
      `Detail technique: ${details}`
    ].join("\n");
  }

  return `Erreur de demarrage:\n${details}`;
}

process.on("unhandledRejection", (reason) => {
  console.error("[PAD] Unhandled rejection:", reason);
});

async function bootstrap() {
  try {
    const { setupDataLayer } = require("./services/data-layer.cjs");
    dataLayer = setupDataLayer({ app });
    registerIpcHandlers();
    createMainWindow();

    app.on("activate", () => {
      if (BrowserWindow.getAllWindows().length === 0) {
        createMainWindow();
      }
    });
  } catch (error) {
    console.error("[PAD] Echec au demarrage:", error);
    dialog.showErrorBox("PAD - Erreur de demarrage", getStartupErrorMessage(error));
    app.quit();
  }
}

app.whenReady().then(bootstrap);

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});
