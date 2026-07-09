function electronApi() {
  return window.recobro || null;
}

async function jsonFetch(url, options = {}) {
  const response = await fetch(url, {
    headers: { "Content-Type": "application/json", ...(options.headers || {}) },
    ...options
  });
  if (!response.ok) throw new Error(await response.text());
  return response.json();
}

function query(filters = {}) {
  const params = new URLSearchParams();
  Object.entries(filters).forEach(([key, value]) => {
    if (value !== undefined && value !== null && value !== "") params.set(key, value);
  });
  const text = params.toString();
  return text ? `?${text}` : "";
}

function chooseFile(accept) {
  return new Promise((resolve) => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = accept;
    input.onchange = () => resolve(input.files?.[0] || null);
    input.click();
  });
}

async function downloadBlob(url, filename, options = {}) {
  const response = await fetch(url, options);
  if (!response.ok) throw new Error(await response.text());
  const blob = await response.blob();
  const objectUrl = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = objectUrl;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(objectUrl);
}

export const api = {
  listCases(filters) {
    const native = electronApi();
    return native ? native.listCases(filters) : jsonFetch(`/api/cases${query(filters)}`);
  },
  getCase(id) {
    const native = electronApi();
    return native ? native.getCase(id) : jsonFetch(`/api/cases/${id}`);
  },
  updateCase(id, patch) {
    const native = electronApi();
    return native
      ? native.updateCase(id, patch)
      : jsonFetch(`/api/cases/${id}`, { method: "PATCH", body: JSON.stringify(patch) });
  },
  changeStatus(id, status) {
    const native = electronApi();
    return native
      ? native.changeStatus(id, status)
      : jsonFetch(`/api/cases/${id}/status`, { method: "PATCH", body: JSON.stringify({ status }) });
  },
  async importBuk() {
    const native = electronApi();
    if (native) return native.importBuk("");
    const file = await chooseFile(".xlsx,.xls,.csv");
    if (!file) return { canceled: true };
    const form = new FormData();
    form.append("file", file);
    const response = await fetch("/api/buk/import", { method: "POST", body: form });
    if (!response.ok) throw new Error(await response.text());
    return response.json();
  },
  dashboard() {
    const native = electronApi();
    return native ? native.dashboard() : jsonFetch("/api/dashboard");
  },
  addManagement(caseId, entry) {
    const native = electronApi();
    return native
      ? native.addManagement(caseId, entry)
      : jsonFetch(`/api/cases/${caseId}/management`, { method: "POST", body: JSON.stringify(entry) });
  },
  listCatalog(type) {
    const native = electronApi();
    return native ? native.listCatalog(type) : jsonFetch(`/api/catalogs/${type}`);
  },
  saveCatalog(type, item) {
    const native = electronApi();
    return native
      ? native.saveCatalog(type, item)
      : jsonFetch(`/api/catalogs/${type}`, { method: "POST", body: JSON.stringify(item) });
  },
  deleteCatalog(type, id) {
    const native = electronApi();
    return native ? native.deleteCatalog(type, id) : jsonFetch(`/api/catalogs/${type}/${id}`, { method: "DELETE" });
  },
  async exportReport(filters) {
    const native = electronApi();
    if (native) return native.exportReport(filters);
    await downloadBlob("/api/reports/export", `reporte-incapacidades-${new Date().toISOString().slice(0, 10)}.xlsx`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(filters || {})
    });
    return { ok: true };
  },
  async downloadImportTemplate() {
    const native = electronApi();
    if (native) return native.downloadImportTemplate();
    await downloadBlob("/api/buk/template", "modelo-importacion-incapacidades-buk.xlsx");
    return { ok: true };
  },
  async downloadPaymentTemplate(filters) {
    await downloadBlob("/api/payments/template", `plantilla-aplicacion-pagos-${new Date().toISOString().slice(0, 10)}.xlsx`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(filters || {})
    });
    return { ok: true };
  },
  async importPayments() {
    const file = await chooseFile(".xlsx,.xls,.csv");
    if (!file) return { canceled: true };
    const form = new FormData();
    form.append("file", file);
    const response = await fetch("/api/payments/import", { method: "POST", body: form });
    if (!response.ok) throw new Error(await response.text());
    return response.json();
  },
  applyPayments(payments) {
    return jsonFetch("/api/payments/apply", {
      method: "POST",
      body: JSON.stringify({ payments })
    });
  },
  openSupport(url) {
    const native = electronApi();
    if (native) return native.openSupport(url);
    if (url) window.open(url, "_blank", "noopener,noreferrer");
    return Promise.resolve({ ok: Boolean(url) });
  }
};
