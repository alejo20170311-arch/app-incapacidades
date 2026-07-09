const fs = require("fs");
const path = require("path");
const initSqlJs = require("sql.js");

const DEFAULT_EPS = [
  "COOSALUD EPS-S",
  "NUEVA EPS",
  "MUTUAL SER",
  "SALUD MIA",
  "ALIANSALUD EPS",
  "SALUD TOTAL EPS S.A.",
  "EPS SANITAS",
  "EPS SURA",
  "FAMISANAR",
  "SERVICIO OCCIDENTAL DE SALUD EPS SOS",
  "COMFENALCO VALLE",
  "COMPENSAR EPS",
  "EPM - EMPRESAS PUBLICAS DE MEDELLIN",
  "FONDO DE PASIVO SOCIAL DE FERROCARRILES NACIONALES DE COLOMBIA",
  "CAJACOPI ATLANTICO",
  "CAPRESOCA",
  "COMFACHOCO",
  "COMFAORIENTE",
  "EPS FAMILIAR DE COLOMBIA",
  "ASMET SALUD",
  "EMSSANAR E.S.S.",
  "CAPITAL SALUD EPS-S",
  "SAVIA SALUD EPS",
  "DUSAKAWI EPSI",
  "ASOCIACION INDIGENA DEL CAUCA EPSI",
  "ANAS WAYUU EPSI",
  "MALLAMAS EPSI",
  "PIJAOS SALUD EPSI",
  "ARL SURA (Accidentes de trabajo)"
];

const DEFAULT_STATES = [
  "Pendiente",
  "Lista para radicar",
  "Radicada",
  "En revision",
  "Devuelta",
  "Corregida",
  "Pagada",
  "Rechazada",
  "Archivada"
];

const DEFAULT_RESPONSIBLES = ["Alejandro"];

async function createStore(userDataPath) {
  const wasmPath = resolveWasmPath();
  const SQL = await initSqlJs({ locateFile: () => wasmPath });
  const dbPath = path.join(userDataPath, "recobro.sqlite");
  fs.mkdirSync(userDataPath, { recursive: true });

  const db = fs.existsSync(dbPath)
    ? new SQL.Database(fs.readFileSync(dbPath))
    : new SQL.Database();

  const store = new Store(db, dbPath);
  store.migrate();
  store.seed();
  store.normalizeAutomaticStatuses();
  store.save();
  return store;
}

function resolveWasmPath() {
  const devPath = path.join(__dirname, "..", "node_modules", "sql.js", "dist", "sql-wasm.wasm");
  const resourcePath = path.join(process.resourcesPath || "", "sql-wasm.wasm");
  return fs.existsSync(resourcePath) ? resourcePath : devPath;
}

class Store {
  constructor(db, dbPath) {
    this.db = db;
    this.dbPath = dbPath;
  }

  migrate() {
    this.exec(`
      PRAGMA foreign_keys = ON;

      CREATE TABLE IF NOT EXISTS catalogs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        type TEXT NOT NULL,
        name TEXT NOT NULL,
        active INTEGER NOT NULL DEFAULT 1,
        UNIQUE(type, name)
      );

      CREATE TABLE IF NOT EXISTS cases (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        employee_name TEXT NOT NULL,
        document TEXT NOT NULL,
        employee_code TEXT,
        salary REAL DEFAULT 0,
        eps TEXT NOT NULL,
        area TEXT,
        position TEXT,
        incapacity_type TEXT NOT NULL,
        diagnosis TEXT,
        start_date TEXT NOT NULL,
        end_date TEXT NOT NULL,
        days INTEGER DEFAULT 0,
        chargeable_days INTEGER DEFAULT 0,
        buk_status TEXT,
        app_status TEXT NOT NULL DEFAULT 'Pendiente',
        incapacity_number TEXT,
        support_url TEXT,
        filing_date TEXT,
        filing_number TEXT,
        responsible TEXT,
        expected_value REAL DEFAULT 0,
        recognized_value REAL DEFAULT 0,
        recovered_value REAL DEFAULT 0,
        pending_value REAL DEFAULT 0,
        observations TEXT,
        next_action_date TEXT,
        pending_documents INTEGER NOT NULL DEFAULT 0,
        created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(document, start_date, end_date, incapacity_type)
      );

      CREATE TABLE IF NOT EXISTS management_entries (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        case_id INTEGER NOT NULL,
        entry_date TEXT NOT NULL,
        note TEXT NOT NULL,
        next_action_date TEXT,
        responsible TEXT,
        created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(case_id) REFERENCES cases(id) ON DELETE CASCADE
      );

      CREATE TABLE IF NOT EXISTS payments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        case_id INTEGER NOT NULL,
        payment_date TEXT NOT NULL,
        amount REAL NOT NULL DEFAULT 0,
        observation TEXT,
        created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(case_id, payment_date),
        FOREIGN KEY(case_id) REFERENCES cases(id) ON DELETE CASCADE
      );
    `);
    this.ensureColumn("cases", "salary", "REAL DEFAULT 0");
    this.ensureColumn("cases", "chargeable_days", "INTEGER DEFAULT 0");
  }

  seed() {
    DEFAULT_EPS.forEach((name) => this.ensureCatalog("eps", name));
    DEFAULT_STATES.forEach((name) => this.ensureCatalog("state", name));
    DEFAULT_RESPONSIBLES.forEach((name) => this.ensureCatalog("responsible", name));
  }

  save() {
    fs.writeFileSync(this.dbPath, Buffer.from(this.db.export()));
  }

  exec(sql) {
    this.db.exec(sql);
  }

  run(sql, params = []) {
    const stmt = this.db.prepare(sql);
    stmt.bind(params);
    stmt.step();
    stmt.free();
    this.save();
  }

  all(sql, params = []) {
    const stmt = this.db.prepare(sql);
    stmt.bind(params);
    const rows = [];
    while (stmt.step()) rows.push(stmt.getAsObject());
    stmt.free();
    return rows;
  }

  one(sql, params = []) {
    return this.all(sql, params)[0] || null;
  }

  ensureColumn(table, column, definition) {
    const exists = this.all(`PRAGMA table_info(${table})`).some((item) => item.name === column);
    if (!exists) this.exec(`ALTER TABLE ${table} ADD COLUMN ${column} ${definition}`);
  }

  ensureCatalog(type, name) {
    if (!name) return;
    const stmt = this.db.prepare("INSERT OR IGNORE INTO catalogs (type, name) VALUES (?, ?)");
    stmt.run([type, String(name).trim()]);
    stmt.free();
  }

  listCatalog(type) {
    return this.all(
      "SELECT id, type, name, active FROM catalogs WHERE type = ? ORDER BY name COLLATE NOCASE",
      [type]
    );
  }

  saveCatalog(type, item) {
    const name = String(item.name || "").trim();
    if (!name) throw new Error("El nombre es obligatorio");
    if (item.id) {
      this.run("UPDATE catalogs SET name = ?, active = ? WHERE id = ? AND type = ?", [
        name,
        item.active === false ? 0 : 1,
        item.id,
        type
      ]);
    } else {
      this.ensureCatalog(type, name);
      this.save();
    }
    return this.listCatalog(type);
  }

  deleteCatalog(type, id) {
    this.run("DELETE FROM catalogs WHERE id = ? AND type = ?", [id, type]);
    return this.listCatalog(type);
  }

  insertCase(data) {
    this.ensureCatalog("eps", data.eps);
    if (data.responsible) this.ensureCatalog("responsible", data.responsible);
    const existing = this.one(
      "SELECT * FROM cases WHERE document = ? AND start_date = ? AND end_date = ? AND incapacity_type = ?",
      [data.document, data.start_date, data.end_date, data.incapacity_type]
    );
    if (existing) return this.updateImportedCase(existing, data);

    const pending = Math.max(Number(data.expected_value || 0) - Number(data.recovered_value || 0), 0);
    const appStatus = resolveCaseStatus(
      {
        ...data,
        pending_value: pending
      },
      data.app_status || "Pendiente"
    );
    this.run(
      `INSERT INTO cases (
        employee_name, document, employee_code, salary, eps, area, position, incapacity_type,
        diagnosis, start_date, end_date, days, chargeable_days, buk_status, app_status, incapacity_number,
        support_url, filing_date, filing_number, responsible, expected_value,
        recognized_value, recovered_value, pending_value, observations, next_action_date,
        pending_documents
      ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
      [
        data.employee_name,
        data.document,
        data.employee_code || "",
        Number(data.salary || 0),
        data.eps,
        data.area || "",
        data.position || "",
        data.incapacity_type,
        data.diagnosis || "",
        data.start_date,
        data.end_date,
        Number(data.days || 0),
        Number(data.chargeable_days || 0),
        data.buk_status || "",
        appStatus,
        data.incapacity_number || "",
        data.support_url || "",
        data.filing_date || "",
        data.filing_number || "",
        data.responsible || "",
        Number(data.expected_value || 0),
        Number(data.recognized_value || 0),
        Number(data.recovered_value || 0),
        pending,
        data.observations || "",
        data.next_action_date || "",
        data.pending_documents ? 1 : 0
      ]
    );
    const row = this.one(
      "SELECT id FROM cases WHERE document = ? AND start_date = ? AND end_date = ? AND incapacity_type = ?",
      [data.document, data.start_date, data.end_date, data.incapacity_type]
    );
    return { inserted: true, id: row.id };
  }

  updateImportedCase(existing, data) {
    const hasSalary = Number(data.salary || 0) > 0;
    const salary = hasSalary ? Number(data.salary || 0) : Number(existing.salary || 0);
    const chargeableDays = hasSalary ? Number(data.chargeable_days || 0) : Number(existing.chargeable_days || 0);
    const expected = hasSalary ? Number(data.expected_value || 0) : Number(existing.expected_value || 0);
    const recovered = Number(existing.recovered_value || 0);
    const pending = Math.max(expected - recovered, 0);
    const recognized = Math.max(Number(existing.recognized_value || 0), recovered);
    const nextStatus = resolveCaseStatus(
      {
        days: Number(data.days || existing.days || 0),
        chargeable_days: chargeableDays,
        expected_value: expected,
        recovered_value: recovered,
        pending_value: pending
      },
      existing.app_status || "Pendiente"
    );

    this.run(
      `UPDATE cases
       SET employee_name = ?,
           employee_code = ?,
           salary = ?,
           eps = ?,
           area = ?,
           position = ?,
           diagnosis = ?,
           days = ?,
           chargeable_days = ?,
           buk_status = ?,
           incapacity_number = ?,
           support_url = ?,
           expected_value = ?,
           recognized_value = ?,
           pending_value = ?,
           app_status = ?,
           updated_at = CURRENT_TIMESTAMP
       WHERE id = ?`,
      [
        data.employee_name || existing.employee_name,
        data.employee_code || existing.employee_code || "",
        salary,
        data.eps || existing.eps,
        data.area || existing.area || "",
        data.position || existing.position || "",
        data.diagnosis || existing.diagnosis || "",
        Number(data.days || existing.days || 0),
        chargeableDays,
        data.buk_status || existing.buk_status || "",
        data.incapacity_number || existing.incapacity_number || "",
        data.support_url || existing.support_url || "",
        expected,
        recognized,
        pending,
        nextStatus,
        existing.id
      ]
    );
    return { inserted: false, updated: true, id: existing.id };
  }

  listCases(filters = {}) {
    const clauses = [];
    const params = [];
    const q = String(filters.q || "").trim();
    if (q) {
      clauses.push(`(
        document LIKE ? OR employee_name LIKE ? OR eps LIKE ? OR app_status LIKE ? OR
        incapacity_type LIKE ? OR responsible LIKE ?
      )`);
      const like = `%${q}%`;
      params.push(like, like, like, like, like, like);
    }
    ["eps", "app_status", "incapacity_type", "responsible"].forEach((field) => {
      if (filters[field]) {
        clauses.push(`${field} = ?`);
        params.push(filters[field]);
      }
    });
    if (filters.from) {
      clauses.push("start_date >= ?");
      params.push(filters.from);
    }
    if (filters.to) {
      clauses.push("start_date <= ?");
      params.push(filters.to);
    }
    const where = clauses.length ? `WHERE ${clauses.join(" AND ")}` : "";
    return this.all(`SELECT * FROM cases ${where} ORDER BY start_date DESC, id DESC`, params).map(enrichCase);
  }

  getCase(id) {
    const record = this.one("SELECT * FROM cases WHERE id = ?", [id]);
    if (!record) return null;
    return {
      ...enrichCase(record),
      history: this.all(
        "SELECT * FROM management_entries WHERE case_id = ? ORDER BY entry_date DESC, id DESC",
        [id]
      )
    };
  }

  updateCase(id, patch) {
    const allowed = [
      "employee_name",
      "document",
      "employee_code",
      "salary",
      "eps",
      "area",
      "position",
      "incapacity_type",
      "diagnosis",
      "start_date",
      "end_date",
      "days",
      "chargeable_days",
      "buk_status",
      "app_status",
      "incapacity_number",
      "support_url",
      "filing_date",
      "filing_number",
      "responsible",
      "expected_value",
      "recognized_value",
      "recovered_value",
      "pending_value",
      "observations",
      "next_action_date",
      "pending_documents"
    ];
    const keys = allowed.filter((key) => Object.prototype.hasOwnProperty.call(patch, key));
    if (!keys.length) return this.getCase(id);
    if (patch.eps) this.ensureCatalog("eps", patch.eps);
    if (patch.responsible) this.ensureCatalog("responsible", patch.responsible);

    const assignments = keys.map((key) => `${key} = ?`).join(", ");
    const values = keys.map((key) => {
      if (["salary", "chargeable_days", "expected_value", "recognized_value", "recovered_value", "pending_value"].includes(key)) {
        return Number(patch[key] || 0);
      }
      if (key === "pending_documents") return patch[key] ? 1 : 0;
      return patch[key] ?? "";
    });
    this.run(`UPDATE cases SET ${assignments}, updated_at = CURRENT_TIMESTAMP WHERE id = ?`, [...values, id]);
    this.recalculatePending(id);
    this.applyAutomaticStatus(id);
    return this.getCase(id);
  }

  changeCaseStatus(id, status) {
    this.ensureCatalog("state", status);
    this.run("UPDATE cases SET app_status = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?", [status, id]);
    return this.getCase(id);
  }

  recalculatePending(id) {
    this.run(
      `UPDATE cases
       SET pending_value = CASE
         WHEN expected_value - recovered_value > 0 THEN expected_value - recovered_value
         ELSE 0
       END
       WHERE id = ?`,
      [id]
    );
  }

  applyAutomaticStatus(id) {
    const record = this.one("SELECT * FROM cases WHERE id = ?", [id]);
    if (!record) return;
    const nextStatus = resolveCaseStatus(record, record.app_status);
    if (nextStatus !== record.app_status) {
      this.run("UPDATE cases SET app_status = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?", [nextStatus, id]);
    }
  }

  getContinuousIncapacityDays(document, diagnosis, startDate) {
    if (!document || !diagnosis || !startDate) return 0;
    let total = 0;
    let cursor = addDays(startDate, -1);
    const normalizedDiagnosis = normalizeText(diagnosis);
    while (cursor) {
      const previous = this.one(
        `SELECT id, start_date, end_date, days, diagnosis
         FROM cases
         WHERE document = ? AND end_date = ?
         ORDER BY start_date DESC, id DESC
         LIMIT 1`,
        [document, cursor]
      );
      if (!previous || normalizeText(previous.diagnosis) !== normalizedDiagnosis) break;
      total += Number(previous.days || daysBetweenInclusive(previous.start_date, previous.end_date));
      cursor = addDays(previous.start_date, -1);
    }
    return total;
  }

  addManagement(caseId, entry) {
    this.run(
      `INSERT INTO management_entries (case_id, entry_date, note, next_action_date, responsible)
       VALUES (?, ?, ?, ?, ?)`,
      [
        caseId,
        entry.entry_date || today(),
        String(entry.note || "").trim(),
        entry.next_action_date || "",
        entry.responsible || ""
      ]
    );
    const patch = {};
    if (entry.next_action_date) patch.next_action_date = entry.next_action_date;
    if (entry.responsible) patch.responsible = entry.responsible;
    if (Object.keys(patch).length) this.updateCase(caseId, patch);
    return this.getCase(caseId);
  }

  applyPayment(caseId, payment) {
    const record = this.getCase(caseId);
    if (!record) return { applied: false, reason: "No encontrado" };
    const amount = Number(payment.amount || 0);
    if (!amount || amount <= 0) return { applied: false, reason: "Valor girado vacio" };
    const paymentDate = payment.payment_date || today();
    const existingPayment = this.one(
      "SELECT * FROM payments WHERE case_id = ? AND payment_date = ?",
      [caseId, paymentDate]
    );
    const previousAmount = Number(existingPayment?.amount || 0);
    const delta = amount - previousAmount;
    if (delta === 0 && existingPayment) {
      const nextStatus = resolveCaseStatus(record, record.app_status);
      this.run(
        `UPDATE payments
         SET observation = ?, updated_at = CURRENT_TIMESTAMP
         WHERE id = ?`,
        [payment.observation || existingPayment.observation || "", existingPayment.id]
      );
      if (nextStatus !== record.app_status) {
        this.run("UPDATE cases SET app_status = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?", [nextStatus, caseId]);
      }
      return { applied: true, updated: true, id: caseId, amount, delta: 0, pending: record.pending_value, status: nextStatus };
    }

    const nextRecovered = Math.max(Number(record.recovered_value || 0) + delta, 0);
    const nextRecognized = Math.max(Number(record.recognized_value || 0), nextRecovered);
    const expected = Number(record.expected_value || 0);
    const nextPending = Math.max(expected - nextRecovered, 0);
    const nextStatus = resolveCaseStatus(
      {
        ...record,
        recovered_value: nextRecovered,
        recognized_value: nextRecognized,
        pending_value: nextPending
      },
      record.app_status
    );
    const note = [
      existingPayment ? `Pago actualizado de ${formatCurrency(previousAmount)} a ${formatCurrency(amount)}` : `Pago aplicado por ${formatCurrency(amount)}`,
      `Fecha pago: ${paymentDate}`,
      payment.observation ? `Observacion: ${payment.observation}` : ""
    ].filter(Boolean).join(". ");

    if (existingPayment) {
      this.run(
        "UPDATE payments SET amount = ?, observation = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?",
        [amount, payment.observation || existingPayment.observation || "", existingPayment.id]
      );
    } else {
      this.run(
        "INSERT INTO payments (case_id, payment_date, amount, observation) VALUES (?, ?, ?, ?)",
        [caseId, paymentDate, amount, payment.observation || ""]
      );
    }

    this.run(
      `UPDATE cases
       SET recovered_value = ?, recognized_value = ?, pending_value = ?, app_status = ?,
           observations = CASE
             WHEN ? = '' THEN observations
             WHEN observations = '' THEN ?
             ELSE observations || char(10) || ?
           END,
           updated_at = CURRENT_TIMESTAMP
       WHERE id = ?`,
      [
        nextRecovered,
        nextRecognized,
        nextPending,
        nextStatus,
        payment.observation || "",
        payment.observation || "",
        payment.observation || "",
        caseId
      ]
    );
    this.addManagement(caseId, {
      entry_date: paymentDate,
      note,
      responsible: payment.responsible || "Importacion pagos"
    });
    return { applied: true, updated: Boolean(existingPayment), id: caseId, amount, delta, pending: nextPending, status: nextStatus };
  }

  clearPayments() {
    this.run("DELETE FROM payments");
    this.run("DELETE FROM management_entries WHERE note LIKE 'Pago aplicado por%' OR note LIKE 'Pago actualizado de%'");
    this.run(
      `UPDATE cases
       SET recovered_value = 0,
           recognized_value = 0,
           pending_value = expected_value,
           app_status = CASE WHEN app_status = 'Pagada' THEN 'Radicada' ELSE app_status END,
           updated_at = CURRENT_TIMESTAMP`
    );
    return { ok: true };
  }

  normalizeAutomaticStatuses() {
    this.exec(`
      UPDATE cases
      SET pending_value = CASE
            WHEN expected_value - recovered_value > 0 THEN expected_value - recovered_value
            ELSE 0
          END,
          updated_at = CURRENT_TIMESTAMP;

      UPDATE cases
      SET app_status = 'Pagada',
          updated_at = CURRENT_TIMESTAMP
      WHERE recovered_value > 0
        AND pending_value <= 0;

      UPDATE cases
      SET app_status = 'Archivada',
          updated_at = CURRENT_TIMESTAMP
      WHERE recovered_value <= 0
        AND pending_value <= 0
        AND expected_value <= 0
        AND chargeable_days <= 0
        AND days BETWEEN 1 AND 2
        AND app_status = 'Pendiente';
    `);
    this.save();
    return { ok: true };
  }

  getDashboard() {
    const cases = this.listCases({});
    const totals = cases.reduce(
      (acc, item) => {
        acc.active += item.app_status !== "Archivada" ? 1 : 0;
        acc.pending += item.app_status === "Pendiente" ? 1 : 0;
        acc.filed += item.app_status === "Radicada" ? 1 : 0;
        acc.paid += item.app_status === "Pagada" ? 1 : 0;
        acc.returned += item.app_status === "Devuelta" ? 1 : 0;
        acc.rejected += item.app_status === "Rechazada" ? 1 : 0;
        acc.pendingValue += Number(item.pending_value || 0);
        acc.recoveredValue += Number(item.recovered_value || 0);
        return acc;
      },
      { active: 0, pending: 0, filed: 0, paid: 0, returned: 0, rejected: 0, pendingValue: 0, recoveredValue: 0 }
    );

    return {
      totals,
      byEps: groupCount(cases, "eps"),
      pendingByEps: groupSum(cases, "eps", "pending_value"),
      byStatus: groupCount(cases, "app_status"),
      monthlyRecovery: monthlySum(cases, "recovered_value"),
      aging: agingBuckets(cases),
      epsRanking: groupSum(cases, "eps", "recovered_value").sort((a, b) => b.value - a.value),
      alerts: buildAlerts(cases)
    };
  }
}

function groupCount(items, key) {
  const map = new Map();
  items.forEach((item) => map.set(item[key] || "Sin dato", (map.get(item[key] || "Sin dato") || 0) + 1));
  return Array.from(map, ([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
}

function groupSum(items, key, valueKey) {
  const map = new Map();
  items.forEach((item) => {
    const name = item[key] || "Sin dato";
    map.set(name, (map.get(name) || 0) + Number(item[valueKey] || 0));
  });
  return Array.from(map, ([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
}

function monthlySum(items, valueKey) {
  const map = new Map();
  items.forEach((item) => {
    const month = String(item.filing_date || item.start_date || "").slice(0, 7) || "Sin fecha";
    map.set(month, (map.get(month) || 0) + Number(item[valueKey] || 0));
  });
  return Array.from(map, ([name, value]) => ({ name, value })).sort((a, b) => a.name.localeCompare(b.name));
}

function agingBuckets(items) {
  const buckets = [
    { name: "0-30 dias", value: 0 },
    { name: "31-60 dias", value: 0 },
    { name: "61-90 dias", value: 0 },
    { name: "90+ dias", value: 0 }
  ];
  items.forEach((item) => {
    const age = daysBetween(item.start_date, today());
    if (age <= 30) buckets[0].value += Number(item.pending_value || 0);
    else if (age <= 60) buckets[1].value += Number(item.pending_value || 0);
    else if (age <= 90) buckets[2].value += Number(item.pending_value || 0);
    else buckets[3].value += Number(item.pending_value || 0);
  });
  return buckets;
}

function buildAlerts(items) {
  const now = today();
  return items.flatMap((item) => {
    const alerts = [];
    if (!item.filing_date && daysBetween(item.start_date, now) > 30) {
      alerts.push({ type: "Sin radicar", message: `${item.employee_name} lleva mas de 30 dias sin radicar`, caseId: item.id });
    }
    if (item.filing_date && !["Pagada", "Rechazada", "Archivada"].includes(item.app_status) && daysBetween(item.filing_date, now) > 60) {
      alerts.push({ type: "Sin respuesta", message: `${item.employee_name} lleva mas de 60 dias sin respuesta`, caseId: item.id });
    }
    if (item.pending_documents) {
      alerts.push({ type: "Documentos", message: `${item.employee_name} tiene documentos pendientes`, caseId: item.id });
    }
    if (item.next_action_date === now) {
      alerts.push({ type: "Gestion hoy", message: `Gestion programada hoy para ${item.employee_name}`, caseId: item.id });
    }
    if (Number(item.recovered_value || 0) > 0 && Number(item.pending_value || 0) > 0) {
      alerts.push({ type: "Pago parcial", message: `${item.employee_name} tiene pago parcial`, caseId: item.id });
    }
    return alerts;
  });
}

function enrichCase(item) {
  return {
    ...item,
    pending_documents: Boolean(item.pending_documents),
    alert_count: buildAlerts([item]).length
  };
}

function daysBetween(from, to) {
  if (!from || !to) return 0;
  const start = new Date(`${from}T00:00:00`);
  const end = new Date(`${to}T00:00:00`);
  return Math.floor((end - start) / 86400000);
}

function daysBetweenInclusive(from, to) {
  return daysBetween(from, to) + 1;
}

function addDays(dateText, amount) {
  if (!dateText) return "";
  const date = new Date(`${dateText}T00:00:00`);
  date.setDate(date.getDate() + amount);
  return date.toISOString().slice(0, 10);
}

function normalizeText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
}

function today() {
  return new Date().toISOString().slice(0, 10);
}

function formatCurrency(value) {
  return new Intl.NumberFormat("es-CO", {
    style: "currency",
    currency: "COP",
    maximumFractionDigits: 0
  }).format(Number(value || 0));
}

function resolveCaseStatus(record, fallbackStatus = "Pendiente") {
  const recovered = Number(record.recovered_value || 0);
  const pending = Number(record.pending_value || 0);
  if (recovered > 0 && pending <= 0) return "Pagada";
  if (isNoChargeIncapacity(record) && recovered <= 0) return "Archivada";
  return fallbackStatus || "Pendiente";
}

function isNoChargeIncapacity(record) {
  const days = Number(record.days || 0);
  return (
    days > 0 &&
    days <= 2 &&
    Number(record.chargeable_days || 0) <= 0 &&
    Number(record.expected_value || 0) <= 0
  );
}

module.exports = { createStore };
