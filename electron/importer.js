const XLSX = require("xlsx");

const FIELD_MAP = {
  document: ["colaborador - numero de documento", "numero de documento", "número de documento", "documento", "cedula", "cédula"],
  employee_name: ["colaborador - nombre completo", "nombre del colaborador", "nombre", "colaborador", "empleado"],
  employee_code: ["colaborador - codigo", "colaborador - código", "codigo del colaborador", "código del colaborador", "codigo empleado", "código empleado"],
  start_date: ["novedades - fecha inicio", "fecha inicio", "inicio", "fecha de inicio"],
  end_date: ["novedades - fecha de termino", "novedades - fecha de término", "fecha fin", "fin", "fecha de fin"],
  days: ["novedades - dias tomados", "novedades - días tomados", "cantidad de dias", "cantidad de días", "dias", "días"],
  buk_status: ["novedades - estado licencia", "estado", "estado buk"],
  incapacity_type: ["novedades - descripcion tipo", "novedades - descripción tipo", "novedades - nombre tipo", "tipo de licencia", "tipo de incapacidad", "licencia", "incapacidad"],
  support_url: ["campos personalizados de licencia - soporte licencia / incapacidad", "enlace del soporte", "soporte", "url", "link", "archivo"],
  eps: ["eps", "entidad", "entidad recobro", "aseguradora"],
  salary: ["salario", "sueldo", "salario base", "salario mensual"],
  area: ["area", "área"],
  position: ["cargo", "puesto"],
  diagnosis: ["novedades - codigo", "diagnostico", "diagnóstico"],
  incapacity_number: ["numero de incapacidad", "número de incapacidad", "radicado incapacidad"]
};

const MINIMUM_WAGES = {
  2023: 1160000,
  2024: 1300000,
  2025: 1423500,
  2026: 1750905
};

const INCAPACITY_RATE = 0.6667;

function importBukWorkbook(store, filePath) {
  const workbook = XLSX.readFile(filePath, { cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  validateRequiredHeaders(sheet, ["eps"]);

  let inserted = 0;
  let updated = 0;
  let skipped = 0;
  const errors = [];

  const normalizedRows = rows
    .map((raw, index) => ({ index, normalized: normalizeRow(raw) }))
    .sort((a, b) => String(a.normalized.start_date).localeCompare(String(b.normalized.start_date)));
  const importChains = new Map();

  normalizedRows.forEach(({ index, normalized }) => {
    try {
      if (!normalized.document || !normalized.start_date || !normalized.end_date || !normalized.incapacity_type) {
        skipped += 1;
        return;
      }
      const eps = normalized.eps;
      if (!eps) {
        errors.push({ row: index + 2, message: "La fila no tiene EPS" });
        skipped += 1;
        return;
      }
      const previousDays = getPreviousDays(store, importChains, normalized);
      const chargeableDays = calculateChargeableDays(eps, previousDays, normalized.days);
      const expectedValue = calculateExpectedValue(normalized.salary, chargeableDays, normalized.start_date, eps);
      const appStatus = chargeableDays <= 0 && Number(normalized.days || 0) <= 2 ? "Archivada" : "Pendiente";
      const result = store.insertCase({
        ...normalized,
        eps,
        chargeable_days: chargeableDays,
        expected_value: expectedValue,
        pending_value: expectedValue,
        app_status: appStatus
      });
      if (result.inserted) inserted += 1;
      else if (result.updated) updated += 1;
      else skipped += 1;
      updateImportChain(importChains, normalized, previousDays);
    } catch (error) {
      errors.push({ row: index + 2, message: error.message });
    }
  });

  return { ok: true, filePath, total: rows.length, inserted, updated, skipped, errors };
}

function validateRequiredHeaders(sheet, requiredFields) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  const headers = (rows[0] || []).map(normalizeKey);
  const missing = requiredFields.filter((field) => {
    const labels = FIELD_MAP[field].map(normalizeKey);
    return !labels.some((label) => headers.includes(label));
  });
  if (missing.length) {
    throw new Error("El Excel debe traer la columna EPS. Descarga el modelo de importacion y diligencia esa columna.");
  }
}

function normalizeRow(raw) {
  const source = {};
  Object.entries(raw).forEach(([key, value]) => {
    source[normalizeKey(key)] = value;
  });

  const item = {};
  Object.entries(FIELD_MAP).forEach(([field, labels]) => {
    const value = labels.map(normalizeKey).map((key) => source[key]).find((entry) => entry !== undefined && entry !== "");
    item[field] = value ?? "";
  });

  return {
    employee_name: clean(item.employee_name) || "Sin nombre",
    document: clean(item.document),
    employee_code: clean(item.employee_code),
    eps: clean(item.eps),
    salary: parseMoney(item.salary),
    area: clean(item.area),
    position: clean(item.position),
    incapacity_type: clean(item.incapacity_type) || "Sin tipo",
    diagnosis: clean(item.diagnosis),
    start_date: parseDate(item.start_date),
    end_date: parseDate(item.end_date),
    days: Number(item.days || 0),
    buk_status: clean(item.buk_status),
    incapacity_number: clean(item.incapacity_number),
    support_url: clean(item.support_url)
  };
}

function getPreviousDays(store, chains, item) {
  const key = chainKey(item);
  const chain = chains.get(key);
  if (chain && addDays(chain.lastEnd, 1) === item.start_date) return chain.days;
  return store.getContinuousIncapacityDays(item.document, item.diagnosis, item.start_date);
}

function updateImportChain(chains, item, previousDays) {
  chains.set(chainKey(item), {
    lastEnd: item.end_date,
    days: previousDays + Number(item.days || 0)
  });
}

function calculateChargeableDays(eps, previousDays, days) {
  const totalDays = Number(days || 0);
  if (isArl(eps)) return totalDays;
  const before = Math.max(Number(previousDays || 0) - 2, 0);
  const after = Math.max(Number(previousDays || 0) + totalDays - 2, 0);
  return Math.max(after - before, 0);
}

function calculateExpectedValue(salary, chargeableDays, startDate, eps = "") {
  const monthlySalary = Number(salary || 0);
  const days = Number(chargeableDays || 0);
  if (monthlySalary <= 0 || days <= 0) return 0;
  if (isArl(eps)) return Math.round((monthlySalary / 30) * days);

  const minimumWage = minimumWageForDate(startDate);
  const monthlyBase =
    minimumWage && monthlySalary > minimumWage
      ? Math.max(monthlySalary * INCAPACITY_RATE, minimumWage)
      : monthlySalary;

  return Math.round((monthlyBase / 30) * days);
}

function minimumWageForDate(dateText) {
  const year = Number(String(dateText || "").slice(0, 4));
  return MINIMUM_WAGES[year] || 0;
}

function isArl(eps) {
  return normalizeKey(eps).includes("arl");
}

function chainKey(item) {
  return `${item.document}|${normalizeKey(item.diagnosis)}`;
}

function normalizeKey(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
}

function clean(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function parseMoney(value) {
  if (typeof value === "number") return value;
  const text = clean(value).replace(/\$/g, "").replace(/\s/g, "");
  if (!text) return 0;
  if (text.includes(",") && text.includes(".")) return Number(text.replace(/\./g, "").replace(",", "."));
  if (text.includes(",")) return Number(text.replace(",", "."));
  if ((text.match(/\./g) || []).length > 1) return Number(text.replace(/\./g, ""));
  return Number(text) || 0;
}

function addDays(dateText, amount) {
  if (!dateText) return "";
  const date = new Date(`${dateText}T00:00:00`);
  date.setDate(date.getDate() + amount);
  return date.toISOString().slice(0, 10);
}

function parseDate(value) {
  if (!value) return "";
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return "";
    return `${parsed.y}-${pad(parsed.m)}-${pad(parsed.d)}`;
  }
  const text = clean(value);
  const iso = text.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (iso) return `${iso[1]}-${pad(iso[2])}-${pad(iso[3])}`;
  const local = text.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  if (local) return `${local[3]}-${pad(local[2])}-${pad(local[1])}`;
  const date = new Date(text);
  return Number.isNaN(date.getTime()) ? "" : date.toISOString().slice(0, 10);
}

function pad(value) {
  return String(value).padStart(2, "0");
}

module.exports = { importBukWorkbook, calculateExpectedValue };
