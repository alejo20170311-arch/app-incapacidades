const path = require("path");
const { app, BrowserWindow, dialog, ipcMain, shell } = require("electron");
const { createStore } = require("./db");
const { importBukWorkbook } = require("./importer");
const { exportCasesReport, exportImportTemplate } = require("./reports");

let mainWindow;
let store;

const isDev = !app.isPackaged;

async function createWindow() {
  store = await createStore(app.getPath("userData"));

  mainWindow = new BrowserWindow({
    width: 1440,
    height: 920,
    minWidth: 1120,
    minHeight: 720,
    title: "Recobro Incapacidades",
    backgroundColor: "#0f172a",
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  if (isDev) {
    await mainWindow.loadURL("http://127.0.0.1:5173");
  } else {
    await mainWindow.loadFile(path.join(__dirname, "..", "dist", "index.html"));
  }
}

app.whenReady().then(createWindow);

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

app.on("activate", () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});

ipcMain.handle("cases:list", (_event, filters) => store.listCases(filters || {}));
ipcMain.handle("cases:get", (_event, id) => store.getCase(id));
ipcMain.handle("cases:update", (_event, id, patch) => store.updateCase(id, patch));
ipcMain.handle("cases:status", (_event, id, status) => store.changeCaseStatus(id, status));
ipcMain.handle("dashboard:get", () => store.getDashboard());
ipcMain.handle("management:add", (_event, caseId, entry) => store.addManagement(caseId, entry));
ipcMain.handle("catalog:list", (_event, type) => store.listCatalog(type));
ipcMain.handle("catalog:save", (_event, type, item) => store.saveCatalog(type, item));
ipcMain.handle("catalog:delete", (_event, type, id) => store.deleteCatalog(type, id));

ipcMain.handle("support:open", async (_event, url) => {
  if (!url) return { ok: false };
  await shell.openExternal(url);
  return { ok: true };
});

ipcMain.handle("buk:import", async (_event, epsName) => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: "Importar Excel de BUK",
    filters: [{ name: "Excel", extensions: ["xlsx", "xls", "csv"] }],
    properties: ["openFile"]
  });
  if (result.canceled || result.filePaths.length === 0) return { canceled: true };
  return importBukWorkbook(store, result.filePaths[0], epsName);
});

ipcMain.handle("buk:template", async () => {
  const result = await dialog.showSaveDialog(mainWindow, {
    title: "Descargar modelo de importacion",
    defaultPath: "modelo-importacion-incapacidades-buk.xlsx",
    filters: [{ name: "Excel", extensions: ["xlsx"] }]
  });
  if (result.canceled || !result.filePath) return { canceled: true };
  exportImportTemplate(result.filePath, store.listCatalog("eps"));
  return { ok: true, path: result.filePath };
});

ipcMain.handle("reports:export", async (_event, filters) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    title: "Exportar reporte",
    defaultPath: `reporte-incapacidades-${new Date().toISOString().slice(0, 10)}.xlsx`,
    filters: [{ name: "Excel", extensions: ["xlsx"] }]
  });
  if (result.canceled || !result.filePath) return { canceled: true };
  const cases = store.listCases(filters || {});
  exportCasesReport(cases, result.filePath);
  return { ok: true, path: result.filePath, total: cases.length };
});
