const { contextBridge } = require("electron");

contextBridge.exposeInMainWorld("padApp", {
  appName: "PAD Maintenance Routiere",
  appVersion: "0.1.0"
});
