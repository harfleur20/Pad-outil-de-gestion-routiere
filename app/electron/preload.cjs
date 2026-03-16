const { contextBridge, ipcRenderer } = require("electron");

const metadata = ipcRenderer.sendSync("app:getMetadataSync") || {};
const appName = typeof metadata.appName === "string" ? metadata.appName : "PAD Maintenance Routière";
const appVersion = typeof metadata.appVersion === "string" ? metadata.appVersion : "0.0.0";
contextBridge.exposeInMainWorld("padApp", {
  appName,
  appVersion,
  data: {
    getStatus: () => ipcRenderer.invoke("data:status"),
    importFromExcel: (excelPath) => ipcRenderer.invoke("data:importFromExcel", excelPath),
    previewExcelImport: (excelPath) => ipcRenderer.invoke("data:previewExcelImport", excelPath),
    pickExcelFile: () => ipcRenderer.invoke("data:pickExcelFile")
  },
  audit: {
    integrity: () => ipcRenderer.invoke("audit:integrity")
  },
  dashboard: {
    summary: () => ipcRenderer.invoke("dashboard:summary")
  },
  backup: {
    export: () => ipcRenderer.invoke("backup:export"),
    restore: () => ipcRenderer.invoke("backup:restore")
  },
  sheet: {
    definitions: () => ipcRenderer.invoke("sheet:definitions"),
    list: (sheetName, filters) => ipcRenderer.invoke("sheet:list", sheetName, filters),
    create: (sheetName, payload) => ipcRenderer.invoke("sheet:create", sheetName, payload),
    update: (sheetName, rowId, payload) => ipcRenderer.invoke("sheet:update", sheetName, rowId, payload),
    delete: (sheetName, rowId) => ipcRenderer.invoke("sheet:delete", sheetName, rowId)
  },
  sap: {
    list: () => ipcRenderer.invoke("sap:list")
  },
  roads: {
    list: (filters) => ipcRenderer.invoke("roads:list", filters)
  },
  roadSections: {
    list: (filters) => ipcRenderer.invoke("roadSections:list", filters)
  },
  measurement: {
    listCampaigns: (filters) => ipcRenderer.invoke("measurement:listCampaigns", filters),
    listRows: (filters) => ipcRenderer.invoke("measurement:listRows", filters),
    upsertCampaign: (payload) => ipcRenderer.invoke("measurement:upsertCampaign", payload),
    deleteCampaign: (campaignId) => ipcRenderer.invoke("measurement:deleteCampaign", campaignId),
    upsertRow: (payload) => ipcRenderer.invoke("measurement:upsertRow", payload),
    deleteRow: (measurementId) => ipcRenderer.invoke("measurement:deleteRow", measurementId)
  },
  degradations: {
    list: () => ipcRenderer.invoke("degradations:list")
  },
  drainageRules: {
    list: () => ipcRenderer.invoke("drainageRules:list"),
    upsert: (payload) => ipcRenderer.invoke("drainageRules:upsert", payload),
    delete: (ruleId) => ipcRenderer.invoke("drainageRules:delete", ruleId)
  },
  solutions: {
    listTemplates: () => ipcRenderer.invoke("solutions:listTemplates"),
    upsertTemplate: (payload) => ipcRenderer.invoke("solutions:upsertTemplate", payload),
    assignTemplate: (degradationCode, templateKey) =>
      ipcRenderer.invoke("solutions:assignTemplate", degradationCode, templateKey),
    setOverride: (degradationCode, solutionText) =>
      ipcRenderer.invoke("solutions:setOverride", degradationCode, solutionText),
    clearOverride: (degradationCode) => ipcRenderer.invoke("solutions:clearOverride", degradationCode)
  },
  maintenance: {
    list: (filters) => ipcRenderer.invoke("maintenance:list", filters),
    upsert: (payload) => ipcRenderer.invoke("maintenance:upsert", payload),
    delete: (interventionId) => ipcRenderer.invoke("maintenance:delete", interventionId),
    pickAttachment: () => ipcRenderer.invoke("maintenance:pickAttachment"),
    openAttachment: (attachmentPath) => ipcRenderer.invoke("maintenance:openAttachment", attachmentPath)
  },
  decision: {
    evaluate: (payload) => ipcRenderer.invoke("decision:evaluate", payload)
  },
  reporting: {
    listHistory: (filters) => ipcRenderer.invoke("reporting:listHistory", filters),
    clearHistory: () => ipcRenderer.invoke("reporting:clearHistory"),
    exportHistoryXlsx: () => ipcRenderer.invoke("reporting:exportHistoryXlsx"),
    exportMaintenanceXlsx: () => ipcRenderer.invoke("reporting:exportMaintenanceXlsx")
  },
  printing: {
    exportCurrentViewPdf: (suggestedName) => ipcRenderer.invoke("printing:exportCurrentViewPdf", suggestedName)
  },
  lifecycle: {
    notifyReady: () => ipcRenderer.send("app:ready")
  },
  ping: () => true
});



