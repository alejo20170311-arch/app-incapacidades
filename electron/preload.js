const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("recobro", {
  listCases: (filters) => ipcRenderer.invoke("cases:list", filters),
  getCase: (id) => ipcRenderer.invoke("cases:get", id),
  updateCase: (id, patch) => ipcRenderer.invoke("cases:update", id, patch),
  changeStatus: (id, status) => ipcRenderer.invoke("cases:status", id, status),
  importBuk: (eps) => ipcRenderer.invoke("buk:import", eps),
  downloadImportTemplate: () => ipcRenderer.invoke("buk:template"),
  dashboard: () => ipcRenderer.invoke("dashboard:get"),
  addManagement: (caseId, entry) => ipcRenderer.invoke("management:add", caseId, entry),
  listCatalog: (type) => ipcRenderer.invoke("catalog:list", type),
  saveCatalog: (type, item) => ipcRenderer.invoke("catalog:save", type, item),
  deleteCatalog: (type, id) => ipcRenderer.invoke("catalog:delete", type, id),
  exportReport: (filters) => ipcRenderer.invoke("reports:export", filters),
  openSupport: (url) => ipcRenderer.invoke("support:open", url)
});
