const XLSX = require("xlsx");

const FIELD_MAP = {
  id: ["id incapacidad", "id", "caso", "case id"],
  amount: ["valor girado", "valor pagado", "valor pago", "monto", "pago"],
  payment_date: ["fecha pago", "fecha de pago", "fecha giro", "fecha transaccion", "fecha transacción"],
  observation: ["observacion pago", "observación pago", "observacion", "observación", "nota"]
};

function importPaymentsWorkbook(store, filePath) {
  const workbook = XLSX.readFile(filePath, { cellDates: true });
  const sheetName = workbook.SheetNames.find((name) => normalizeKey(name).includes("pago")) || workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  let applied = 0;
  let updated = 0;
  let skipped = 0;
  const errors = [];

  rows.forEach((raw, index) => {
    const item = normalizeRow(raw);
    if (!item.id && !item.amount) return;
    if (!item.id || !item.amount) {
      skipped += 1;
      errors.push({ row: index + 2, message: "Debe tener ID incapacidad y Valor girado" });
      return;
    }
    const result = store.applyPayment(item.id, item);
    if (result.applied && result.updated) updated += 1;
    else if (result.applied) applied += 1;
    else {
      skipped += 1;
      errors.push({ row: index + 2, message: result.reason || "No aplicado" });
    }
  });

  return { ok: true, total: rows.length, applied, updated, skipped, errors };
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
    id: Number(clean(item.id)),
    amount: parseMoney(item.amount),
    payment_date: parseDate(item.payment_date),
    observation: clean(item.observation)
  };
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
  return Number(text);
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

module.exports = { importPaymentsWorkbook };
