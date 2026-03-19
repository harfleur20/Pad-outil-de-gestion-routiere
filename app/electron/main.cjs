const { app, BrowserWindow, ipcMain, dialog, shell } = require("electron");
const path = require("path");
const fs = require("fs");
const { getAppMetadata } = require("./app-metadata.cjs");

const isDev = !app.isPackaged;
let dataLayer = null;
let mainWindow = null;
let splashWindow = null;
let startupFallbackTimer = null;
let startupFinalizeTimer = null;
let splashStageState = {
  status: "Initialisation du logiciel",
  caption: "Préparation de l'environnement",
  progress: 18
};
const appIconPath = path.join(__dirname, "assets", "icon.ico");
const { appName, appVersion } = getAppMetadata();
const appTagline = "Pilotez la maintenance routière du PAD avec des décisions rapides et fiables.";
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

const hasSingleInstanceLock = app.requestSingleInstanceLock();
if (!hasSingleInstanceLock) {
  app.quit();
}

function createMainWindow() {
  const win = new BrowserWindow({
    show: false,
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

  win.on("closed", () => {
    if (mainWindow === win) {
      mainWindow = null;
    }
  });

  mainWindow = win;
  return win;
}

function resolveSplashLogoPath() {
  const candidates = [
    path.join(__dirname, "../web/public/logo-pad.png"),
    path.join(__dirname, "../web/dist/logo-pad.png")
  ];

  for (const candidate of candidates) {
    if (fs.existsSync(candidate)) {
      return candidate;
    }
  }

  return "";
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function buildSplashHtml() {
  const splashLogoPath = resolveSplashLogoPath();
  const splashLogoDataUrl = splashLogoPath
    ? `data:image/${path.extname(splashLogoPath).replace(".", "") || "png"};base64,${fs
        .readFileSync(splashLogoPath)
        .toString("base64")}`
    : "";

  return `<!doctype html>
  <html lang="fr">
    <head>
      <meta charset="UTF-8" />
      <meta name="viewport" content="width=device-width, initial-scale=1.0" />
      <title>${escapeHtml(appName)}</title>
      <style>
        :root {
          --pad-blue-dark: #0a2f8f;
          --pad-blue: #115eb8;
          --pad-cyan: #25afe7;
          --pad-white: #ffffff;
          --pad-muted: #546788;
          --pad-surface: #f4f9ff;
          --pad-line: rgba(17, 94, 184, 0.12);
        }

        * { box-sizing: border-box; }
        body {
          margin: 0;
          min-height: 100vh;
          display: grid;
          place-items: center;
          overflow: hidden;
          font-family: "Montserrat", "Trebuchet MS", sans-serif;
          background: transparent;
          color: #07224b;
        }

        .startup-shell {
          width: 100%;
          height: 100vh;
          padding: 18px;
          display: grid;
          place-items: center;
        }

        .startup-card {
          width: min(100%, 420px);
          min-height: 420px;
          padding: 28px 28px 24px;
          border-radius: 24px;
          background: #ffffff;
          border: 1px solid var(--pad-line);
          box-shadow: 0 24px 60px rgba(10, 47, 143, 0.12);
          position: relative;
          overflow: hidden;
          text-align: center;
          display: grid;
          align-content: center;
        }

        .startup-card::before {
          content: "";
          position: absolute;
          inset: 0 auto auto 0;
          width: 100%;
          height: 4px;
          background: linear-gradient(90deg, var(--pad-blue-dark), var(--pad-blue), var(--pad-cyan));
          opacity: 0.9;
        }

        .startup-logo {
          width: 68px;
          height: 68px;
          margin: 0 auto 16px;
          object-fit: contain;
          display: block;
        }

        .startup-title {
          margin: 0;
          color: var(--pad-blue-dark);
          font-size: 1.58rem;
          line-height: 1.08;
          font-weight: 800;
          letter-spacing: -0.02em;
        }

        .startup-tagline {
          margin: 10px auto 0;
          max-width: 320px;
          color: var(--pad-muted);
          font-size: 0.88rem;
          line-height: 1.5;
        }

        .startup-status {
          margin-top: 22px;
          color: var(--pad-muted);
          font-size: 0.86rem;
          line-height: 1.45;
          min-height: 1.4em;
        }

        .startup-progress {
          margin-top: 12px;
          width: 100%;
          height: 6px;
          border-radius: 999px;
          background: #e8f0fb;
          overflow: hidden;
          position: relative;
        }

        .startup-progress__bar {
          width: 18%;
          height: 100%;
          border-radius: inherit;
          background: linear-gradient(90deg, var(--pad-blue-dark), var(--pad-blue), var(--pad-cyan));
          transition: width 320ms ease;
        }

        .startup-caption {
          margin-top: 10px;
          font-size: 0.72rem;
          color: var(--pad-blue);
          font-weight: 600;
          letter-spacing: 0.04em;
          text-transform: uppercase;
          min-height: 1.2em;
        }

        .startup-version {
          margin-top: 18px;
          font-size: 0.78rem;
          color: var(--pad-muted);
        }
      </style>
    </head>
    <body>
      <div class="startup-shell">
        <section class="startup-card" aria-label="Écran de lancement PAD">
          ${splashLogoDataUrl ? `<img class="startup-logo" src="${splashLogoDataUrl}" alt="Logo Port Autonome de Douala" />` : ""}
          <h1 class="startup-title">${escapeHtml(appName)}</h1>
          <p class="startup-tagline">${escapeHtml(appTagline)}</p>
          <div class="startup-status" id="startup-status">${escapeHtml(splashStageState.status)}</div>
          <div class="startup-progress" aria-hidden="true">
            <div class="startup-progress__bar" id="startup-progress-bar" style="width:${Math.max(6, Math.min(100, splashStageState.progress))}%"></div>
          </div>
          <div class="startup-caption" id="startup-caption">${escapeHtml(splashStageState.caption)}</div>
          <div class="startup-version">Version ${escapeHtml(appVersion)}</div>
        </section>
      </div>
    </body>
  </html>`;
}

function clearStartupFallbackTimer() {
  if (startupFallbackTimer) {
    clearTimeout(startupFallbackTimer);
    startupFallbackTimer = null;
  }
}

function clearStartupFinalizeTimer() {
  if (startupFinalizeTimer) {
    clearTimeout(startupFinalizeTimer);
    startupFinalizeTimer = null;
  }
}

function applySplashStage() {
  if (!splashWindow || splashWindow.isDestroyed()) {
    return;
  }

  const script = `
    (() => {
      const status = document.getElementById("startup-status");
      const caption = document.getElementById("startup-caption");
      const bar = document.getElementById("startup-progress-bar");
      if (status) status.textContent = ${JSON.stringify(splashStageState.status)};
      if (caption) caption.textContent = ${JSON.stringify(splashStageState.caption)};
      if (bar) bar.style.width = ${JSON.stringify(`${Math.max(6, Math.min(100, splashStageState.progress))}%`)};
    })();
  `;

  splashWindow.webContents.executeJavaScript(script, true).catch(() => {});
}

function setSplashStage(status, caption, progress) {
  splashStageState = {
    status,
    caption,
    progress
  };
  applySplashStage();
}

function finalizeStartup() {
  clearStartupFallbackTimer();
  clearStartupFinalizeTimer();

  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.show();
    mainWindow.focus();
  }

  if (splashWindow && !splashWindow.isDestroyed()) {
    splashWindow.close();
    splashWindow = null;
  }
}

function createSplashWindow() {
  if (splashWindow && !splashWindow.isDestroyed()) {
    return splashWindow;
  }

  const win = new BrowserWindow({
    width: 520,
    height: 520,
    resizable: false,
    movable: true,
    minimizable: false,
    maximizable: false,
    fullscreenable: false,
    frame: false,
    transparent: true,
    autoHideMenuBar: true,
    skipTaskbar: true,
    alwaysOnTop: true,
    backgroundColor: "#00000000",
    show: true
  });

  win.loadURL(`data:text/html;charset=UTF-8,${encodeURIComponent(buildSplashHtml())}`);
  win.webContents.on("did-finish-load", () => {
    applySplashStage();
  });
  win.on("closed", () => {
    if (splashWindow === win) {
      splashWindow = null;
    }
  });

  splashWindow = win;
  return win;
}

function registerIpcHandlers() {
  ipcMain.on("app:getMetadataSync", (event) => {
    event.returnValue = { appName, appVersion };
  });
  ipcMain.on("app:ready", () => {
    setSplashStage("Ouverture du logiciel", "Préparation de l'interface", 100);
    clearStartupFinalizeTimer();
    startupFinalizeTimer = setTimeout(finalizeStartup, 220);
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
  const databaseLocked = /database is locked|SQLITE_BUSY/i.test(details);

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

  if (databaseLocked) {
    return [
      "La base PAD est déjà utilisée par une autre instance ou par une opération encore en cours.",
      "",
      "Fermez les autres fenêtres PAD puis relancez le logiciel.",
      "",
      `Detail technique: ${details}`
    ].join("\n");
  }

  return `Erreur de demarrage:\n${details}`;
}

function isDatabaseLockError(error) {
  const details = error instanceof Error ? error.message : String(error);
  return /database is locked|SQLITE_BUSY/i.test(details);
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function initializeDataLayerWithRetry() {
  const { setupDataLayer } = require("./services/data-layer.cjs");
  const attempts = 6;
  let lastError = null;

  for (let attempt = 1; attempt <= attempts; attempt += 1) {
    try {
      return setupDataLayer({ app });
    } catch (error) {
      lastError = error;
      if (!isDatabaseLockError(error) || attempt === attempts) {
        throw error;
      }
      await delay(500);
    }
  }

  throw lastError;
}

process.on("unhandledRejection", (reason) => {
  console.error("[PAD] Unhandled rejection:", reason);
});

async function bootstrap() {
  try {
    createSplashWindow();
    setSplashStage("Initialisation du logiciel", "Préparation de l'environnement", 18);
    dataLayer = await initializeDataLayerWithRetry();
    setSplashStage("Chargement du référentiel", "Ouverture des catalogues PAD", 68);
    registerIpcHandlers();
    createMainWindow();
    setSplashStage("Ouverture du logiciel", "Préparation de l'interface", 88);
    clearStartupFallbackTimer();
    startupFallbackTimer = setTimeout(finalizeStartup, 15000);

    app.on("activate", () => {
      if (BrowserWindow.getAllWindows().length === 0) {
        createSplashWindow();
        setSplashStage("Initialisation du logiciel", "Préparation de l'environnement", 18);
        createMainWindow();
        setSplashStage("Ouverture du logiciel", "Préparation de l'interface", 88);
        clearStartupFallbackTimer();
        startupFallbackTimer = setTimeout(finalizeStartup, 15000);
      }
    });
  } catch (error) {
    clearStartupFallbackTimer();
    clearStartupFinalizeTimer();
    if (splashWindow && !splashWindow.isDestroyed()) {
      splashWindow.close();
      splashWindow = null;
    }
    console.error("[PAD] Echec au demarrage:", error);
    dialog.showErrorBox("PAD - Erreur de demarrage", getStartupErrorMessage(error));
    app.quit();
  }
}

app.on("second-instance", () => {
  if (mainWindow && !mainWindow.isDestroyed()) {
    if (mainWindow.isMinimized()) {
      mainWindow.restore();
    }
    mainWindow.focus();
    return;
  }
  if (splashWindow && !splashWindow.isDestroyed()) {
    splashWindow.focus();
  }
});

app.whenReady().then(bootstrap);

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});

