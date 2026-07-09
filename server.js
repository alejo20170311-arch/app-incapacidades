const fs = require("fs");
const os = require("os");
const path = require("path");
const express = require("express");
const multer = require("multer");
const { createStore } = require("./electron/db");
const { importBukWorkbook } = require("./electron/importer");
const { importPaymentsWorkbook } = require("./electron/payments");
const { exportCasesReport, exportImportTemplate, exportPaymentTemplate } = require("./electron/reports");

const PORT = Number(process.env.PORT || 5174);
const HOST = process.env.HOST || "0.0.0.0";
const DATA_DIR = process.env.RECOBRO_DATA_DIR || path.join(__dirname, "data");
const upload = multer({ dest: path.join(os.tmpdir(), "recobro-uploads") });

async function main() {
  const store = await createStore(DATA_DIR);
  const app = express();

  app.use(express.json({ limit: "10mb" }));

  app.get("/api/health", (_req, res) => res.json({ ok: true }));

  app.get("/api/cases", (req, res) => {
    res.json(store.listCases(req.query || {}));
  });

  app.get("/api/cases/:id", (req, res) => {
    const record = store.getCase(Number(req.params.id));
    if (!record) return res.status(404).json({ error: "No encontrado" });
    return res.json(record);
  });

  app.patch("/api/cases/:id", (req, res) => {
    res.json(store.updateCase(Number(req.params.id), req.body || {}));
  });

  app.patch("/api/cases/:id/status", (req, res) => {
    res.json(store.changeCaseStatus(Number(req.params.id), req.body.status));
  });

  app.post("/api/cases/:id/management", (req, res) => {
    res.json(store.addManagement(Number(req.params.id), req.body || {}));
  });

  app.get("/api/catalogs/:type", (req, res) => {
    res.json(store.listCatalog(req.params.type));
  });

  app.post("/api/catalogs/:type", (req, res) => {
    res.json(store.saveCatalog(req.params.type, req.body || {}));
  });

  app.delete("/api/catalogs/:type/:id", (req, res) => {
    res.json(store.deleteCatalog(req.params.type, Number(req.params.id)));
  });

  app.get("/api/dashboard", (_req, res) => {
    res.json(store.getDashboard());
  });

  app.post("/api/buk/import", upload.single("file"), (req, res) => {
    if (!req.file) return res.status(400).json({ error: "Archivo requerido" });
    try {
      const result = importBukWorkbook(store, req.file.path);
      return res.json(result);
    } finally {
      fs.rmSync(req.file.path, { force: true });
    }
  });

  app.get("/api/buk/template", (_req, res) => {
    const filePath = path.join(os.tmpdir(), `modelo-importacion-incapacidades-buk-${Date.now()}.xlsx`);
    exportImportTemplate(filePath, store.listCatalog("eps"));
    res.download(filePath, "modelo-importacion-incapacidades-buk.xlsx", () => {
      fs.rmSync(filePath, { force: true });
    });
  });

  app.post("/api/payments/template", (req, res) => {
    const filePath = path.join(os.tmpdir(), `plantilla-aplicacion-pagos-${Date.now()}.xlsx`);
    const cases = store.listCases(req.body || {});
    exportPaymentTemplate(cases, filePath);
    res.download(filePath, `plantilla-aplicacion-pagos-${new Date().toISOString().slice(0, 10)}.xlsx`, () => {
      fs.rmSync(filePath, { force: true });
    });
  });

  app.post("/api/payments/import", upload.single("file"), (req, res) => {
    if (!req.file) return res.status(400).json({ error: "Archivo requerido" });
    try {
      const result = importPaymentsWorkbook(store, req.file.path);
      return res.json(result);
    } finally {
      fs.rmSync(req.file.path, { force: true });
    }
  });

  app.post("/api/payments/apply", (req, res) => {
    const payments = Array.isArray(req.body?.payments) ? req.body.payments : [];
    let applied = 0;
    let updated = 0;
    let skipped = 0;
    const errors = [];
    payments.forEach((payment, index) => {
      const result = store.applyPayment(Number(payment.case_id), {
        amount: Number(payment.amount || 0),
        payment_date: payment.payment_date,
        observation: payment.observation || "",
        responsible: "Registro manual de pago"
      });
      if (result.applied && result.updated) updated += 1;
      else if (result.applied) applied += 1;
      else {
        skipped += 1;
        errors.push({ row: index + 1, message: result.reason || "No aplicado" });
      }
    });
    res.json({ ok: true, applied, updated, skipped, errors });
  });

  app.post("/api/payments/clear", (_req, res) => {
    res.json(store.clearPayments());
  });

  app.post("/api/reports/export", (req, res) => {
    const filePath = path.join(os.tmpdir(), `reporte-incapacidades-${Date.now()}.xlsx`);
    const cases = store.listCases(req.body || {});
    exportCasesReport(cases, filePath);
    res.download(filePath, `reporte-incapacidades-${new Date().toISOString().slice(0, 10)}.xlsx`, () => {
      fs.rmSync(filePath, { force: true });
    });
  });

  const distPath = path.join(__dirname, "dist");
  if (fs.existsSync(distPath)) {
    app.use(express.static(distPath));
    app.use((req, res, next) => {
      if (req.path.startsWith("/api")) return next();
      return res.sendFile(path.join(distPath, "index.html"));
    });
  }

  const httpServer = app.listen(PORT, HOST, () => {
    console.log(`Recobro Incapacidades web en http://127.0.0.1:${PORT}`);
  });
  globalThis.recobroHttpServer = httpServer;
  globalThis.recobroKeepAlive = setInterval(() => {}, 2147483647);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
