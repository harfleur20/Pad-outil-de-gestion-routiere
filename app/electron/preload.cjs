const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("padApp", {
  appName: "PAD Maintenance Routière",
  appVersion: "0.5.0",
  data: {
    getStatus: () => ipcRenderer.invoke("data:status"),
    importFromExcel: (excelPath) => ipcRenderer.invoke("data:importFromExcel", excelPath),
    pickExcelFile: () => ipcRenderer.invoke("data:pickExcelFile")
  },
  audit: {
    integrity: () => ipcRenderer.invoke("audit:integrity")
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
    delete: (interventionId) => ipcRenderer.invoke("maintenance:delete", interventionId)
  },
  decision: {
    evaluate: (payload) => ipcRenderer.invoke("decision:evaluate", payload)
  },
  reporting: {
    listHistory: (filters) => ipcRenderer.invoke("reporting:listHistory", filters),
    clearHistory: () => ipcRenderer.invoke("reporting:clearHistory")
  },
  ping: () => true
});

