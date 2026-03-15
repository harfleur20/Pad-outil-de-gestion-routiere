const { app, BrowserWindow, ipcMain, dialog, shell } = require("electron");
const path = require("path");
const fs = require("fs");
const { getAppMetadata } = require("./app-metadata.cjs");

const isDev = !app.isPackaged;
let dataLayer = null;
const appIconPath = path.join(__dirname, "assets", "icon.ico");
const { appName, appVersion } = getAppMetadata();
const MAX_ATTACHMENT_SIZE_BYTES = 2 * 1024 * 1024;
const MAINTENANCE_ATTACHMENT_EXTENSIONS = [
  "png",
  "jpg",
  "jpeg",
  "webp",
  "pdf",
  "doc",
  "docx",
  "xls",
  "xlsx",
  "txt"
];

if (process.platform === "win32") {
  app.setAppUserModelId("cm.pad.maintenance");
}

function createMainWindow() {
  const win = new BrowserWindow({
    width: 1320,
    height: 860,
    minWidth: 1100,
    minHeight: 720,
    backgroundColor: "#ffffff",
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
  ipcMain.on("app:getMetadataSync", (event) => {
    event.returnValue = { appName, appVersion };
  });
  ipcMain.handle("data:status", () => dataLayer.getDataStatus());
  ipcMain.handle("data:importFromExcel", (_event, excelPath) => dataLayer.importFromExcel(excelPath));
  ipcMain.handle("data:previewExcelImport", (_event, excelPath) => dataLayer.previewExcelImport(excelPath));
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
  ipcMain.handle("dashboard:summary", () => dataLayer.getDashboardSummary());
  ipcMain.handle("backup:export", async () => {
    const result = await dialog.showSaveDialog({
      title: "Sauvegarder les donnees PAD",
      defaultPath: "pad-maintenance-backup.json",
      filters: [{ name: "Sauvegarde JSON", extensions: ["json"] }]
    });
    if (result.canceled || !result.filePath) {
      return null;
    }
    return dataLayer.exportBackup(result.filePath);
  });
  ipcMain.handle("backup:restore", async () => {
    const result = await dialog.showOpenDialog({
      title: "Restaurer une sauvegarde PAD",
      properties: ["openFile"],
      filters: [{ name: "Sauvegarde JSON", extensions: ["json"] }]
    });
    if (result.canceled || result.filePaths.length === 0) {
      return null;
    }
    return dataLayer.restoreBackup(result.filePaths[0]);
  });

  ipcMain.handle("sheet:definitions", () => dataLayer.listSheetDefinitions());
  ipcMain.handle("sheet:list", (_event, sheetName, filters) => dataLayer.listSheetRows(sheetName, filters));
  ipcMain.handle("sheet:create", (_event, sheetName, payload) => dataLayer.createSheetRow(sheetName, payload));
  ipcMain.handle("sheet:update", (_event, sheetName, rowId, payload) =>
    dataLayer.updateSheetRow(sheetName, rowId, payload)
  );
  ipcMain.handle("sheet:delete", (_event, sheetName, rowId) => dataLayer.deleteSheetRow(sheetName, rowId));

  ipcMain.handle("sap:list", () => dataLayer.listSapSectors());
  ipcMain.handle("roads:list", (_event, filters) => dataLayer.listRoadCatalog(filters));
  ipcMain.handle("roadSections:list", (_event, filters) => dataLayer.listRoadSections(filters));
  ipcMain.handle("measurement:listCampaigns", (_event, filters) => dataLayer.listMeasurementCampaigns(filters));
  ipcMain.handle("measurement:listRows", (_event, filters) => dataLayer.listRoadMeasurements(filters));
  ipcMain.handle("measurement:upsertCampaign", (_event, payload) => dataLayer.upsertMeasurementCampaign(payload));
  ipcMain.handle("measurement:deleteCampaign", (_event, campaignId) => dataLayer.deleteMeasurementCampaign(campaignId));
  ipcMain.handle("measurement:upsertRow", (_event, payload) => dataLayer.upsertRoadMeasurement(payload));
  ipcMain.handle("measurement:deleteRow", (_event, measurementId) => dataLayer.deleteRoadMeasurement(measurementId));
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
  ipcMain.handle("maintenance:list", (_event, filters) => dataLayer.listMaintenanceInterventions(filters));
  ipcMain.handle("maintenance:upsert", (_event, payload) => dataLayer.upsertMaintenanceIntervention(payload));
  ipcMain.handle("maintenance:delete", (_event, interventionId) =>
    dataLayer.deleteMaintenanceIntervention(interventionId)
  );
  ipcMain.handle("maintenance:pickAttachment", async () => {
    const result = await dialog.showOpenDialog({
      title: "SÃ©lectionner une piÃ¨ce jointe d'entretien",
      properties: ["openFile"],
      filters: [
        { name: "PiÃ¨ces jointes", extensions: MAINTENANCE_ATTACHMENT_EXTENSIONS },
        { name: "Images", extensions: ["png", "jpg", "jpeg", "webp"] },
        { name: "Documents", extensions: ["pdf", "doc", "docx", "xls", "xlsx", "txt"] }
      ]
    });
    if (result.canceled || result.filePaths.length === 0) {
      return null;
    }
    return copyMaintenanceAttachment(result.filePaths[0]);
  });
  ipcMain.handle("maintenance:openAttachment", async (_event, attachmentPath) => {
    const targetPath = String(attachmentPath || "").trim();
    if (!targetPath) {
      return { opened: false };
    }
    if (!fs.existsSync(targetPath)) {
      throw new Error("PiÃ¨ce jointe introuvable sur le poste.");
    }
    const openResult = await shell.openPath(targetPath);
    if (openResult) {
      throw new Error(openResult);
    }
    return { opened: true };
  });
  ipcMain.handle("decision:evaluate", (_event, payload) => dataLayer.evaluateDecision(payload));
  ipcMain.handle("reporting:listHistory", (_event, filters) => dataLayer.listDecisionHistory(filters));
  ipcMain.handle("reporting:clearHistory", () => dataLayer.clearDecisionHistory());
  ipcMain.handle("reporting:exportHistoryXlsx", async () => {
    const result = await dialog.showSaveDialog({
      title: "Exporter l'historique des decisions",
      defaultPath: "pad-historique-decisions.xlsx",
      filters: [{ name: "Classeur Excel", extensions: ["xlsx"] }]
    });
    if (result.canceled || !result.filePath) {
      return null;
    }
    return dataLayer.exportReportWorkbook("history", result.filePath);
  });
  ipcMain.handle("reporting:exportMaintenanceXlsx", async () => {
    const result = await dialog.showSaveDialog({
      title: "Exporter l'historique des entretiens",
      defaultPath: "pad-historique-entretiens.xlsx",
      filters: [{ name: "Classeur Excel", extensions: ["xlsx"] }]
    });
    if (result.canceled || !result.filePath) {
      return null;
    }
    return dataLayer.exportReportWorkbook("maintenance", result.filePath);
  });
  ipcMain.handle("printing:exportCurrentViewPdf", async (event, suggestedName) => {
    const win = BrowserWindow.fromWebContents(event.sender);
    if (!win) {
      throw new Error("FenÃªtre d'impression introuvable.");
    }

    const normalizedBaseName =
      String(suggestedName || "pad-impression")
        .trim()
        .replace(/[<>:\"/\\\\|?*]+/g, "-")
        .replace(/\s+/g, "-")
        .replace(/-+/g, "-")
        .replace(/^-|-$/g, "")
        .slice(0, 80) || "pad-impression";

    const result = await dialog.showSaveDialog({
      title: "Exporter en PDF",
      defaultPath: `${normalizedBaseName}.pdf`,
      filters: [{ name: "PDF", extensions: ["pdf"] }]
    });
    if (result.canceled || !result.filePath) {
      return null;
    }

    await win.webContents.executeJavaScript(
      "document.fonts && document.fonts.ready ? document.fonts.ready.then(() => true) : true",
      true
    );

    const footerTemplate = `
      <div style="width:100%;font-size:8px;color:#526580;background:#ffffff;padding:0 10px;border-top:1px solid #d9d9d9;display:flex;align-items:center;justify-content:space-between;">
        <span style="flex:1;text-align:left;">${appName}</span>
        <span style="flex:1;text-align:center;">Page <span class="pageNumber"></span> / <span class="totalPages"></span></span>
        <span style="flex:1;text-align:right;">Version ${appVersion}</span>
      </div>
    `;

    const pdfBuffer = await win.webContents.printToPDF({
      printBackground: true,
      landscape: true,
      pageSize: "A4",
      preferCSSPageSize: true,
      marginsType: 1,
      displayHeaderFooter: true,
      headerTemplate: "<div></div>",
      footerTemplate
    });

    fs.writeFileSync(result.filePath, pdfBuffer);
    return { filePath: result.filePath };
  });
}

function copyMaintenanceAttachment(sourcePath) {
  const normalizedSource = String(sourcePath || "").trim();
  if (!normalizedSource || !fs.existsSync(normalizedSource)) {
    throw new Error("Fichier de piÃ¨ce jointe introuvable.");
  }

  const stats = fs.statSync(normalizedSource);
  if (!stats.isFile()) {
    throw new Error("La piÃ¨ce jointe sÃ©lectionnÃ©e est invalide.");
  }
  if (stats.size > MAX_ATTACHMENT_SIZE_BYTES) {
    throw new Error("La piÃ¨ce jointe dÃ©passe 2 Mo.");
  }

  const extension = path.extname(normalizedSource).replace(".", "").toLowerCase();
  if (!MAINTENANCE_ATTACHMENT_EXTENSIONS.includes(extension)) {
    throw new Error(
      `Extension non autorisÃ©e. Formats acceptÃ©s: ${MAINTENANCE_ATTACHMENT_EXTENSIONS.join(", ")}.`
    );
  }

  const originalName = path.basename(normalizedSource);
  const safeBaseName = path
    .basename(originalName, path.extname(originalName))
    .replace(/[^a-zA-Z0-9-_]+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 80) || "piece-jointe";

  const attachmentsDir = path.join(app.getPath("userData"), "attachments");
  fs.mkdirSync(attachmentsDir, { recursive: true });

  const targetName = `${Date.now()}-${safeBaseName}.${extension}`;
  const targetPath = path.join(attachmentsDir, targetName);
  fs.copyFileSync(normalizedSource, targetPath);

  return {
    storedPath: targetPath,
    fileName: originalName,
    size: stats.size
  };
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

