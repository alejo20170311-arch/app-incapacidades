import { createClient } from "@supabase/supabase-js";
import * as XLSX from "xlsx";

const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL || "";
const SUPABASE_ANON_KEY = import.meta.env.VITE_SUPABASE_ANON_KEY || "";
const supabase = SUPABASE_URL && SUPABASE_ANON_KEY ? createClient(SUPABASE_URL, SUPABASE_ANON_KEY) : null;

const IMPORT_TEMPLATE_HEADERS = [
  "Colaborador - Numero de Documento",
  "Colaborador - Nombre Completo",
  "Novedades - Codigo",
  "Campos Personalizados de Licencia - Soporte Licencia / Incapacidad",
  "Novedades - Fecha Inicio",
  "Novedades - Fecha de Termino",
  "Novedades - Dias Tomados",
  "Novedades - Estado Licencia",
  "Novedades - Descripcion Tipo",
  "Novedades - Justificacion",
  "Novedades - Nombre Tipo",
  "Novedades - Tipo de",
  "EPS",
  "Salario"
];

const BUK_FIELD_MAP = {
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

const PAYMENT_FIELD_MAP = {
  id: ["id incapacidad", "id", "caso", "case id"],
  amount: ["valor girado", "valor pagado", "valor pago", "monto", "pago"],
  payment_date: ["fecha pago", "fecha de pago", "fecha giro", "fecha transaccion", "fecha transacción"],
  observation: ["observacion pago", "observación pago", "observacion", "observación", "nota"]
};

const MINIMUM_WAGES = {
  2023: 1160000,
  2024: 1300000,
  2025: 1423500,
  2026: 1750905
};

const INCAPACITY_RATE = 0.6667;

function electronApi() {
  return window.recobro || null;
}

function useSupabase() {
  return Boolean(supabase) && !electronApi();
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

async function readWorkbook(file) {
  const data = await file.arrayBuffer();
  return XLSX.read(data, { cellDates: true });
}

async function downloadBlob(url, filename, options = {}) {
  const response = await fetch(url, options);
  if (!response.ok) throw new Error(await response.text());
  const blob = await response.blob();
  triggerDownload(blob, filename);
}

function triggerWorkbook(workbook, filename) {
  const data = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  triggerDownload(
    new Blob([data], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    filename
  );
}

function triggerDownload(blob, filename) {
  const objectUrl = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = objectUrl;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(objectUrl);
}

async function supabaseAll(table, filters = {}) {
  let request = supabase.from(table).select("*");
  Object.entries(filters).forEach(([key, value]) => {
    if (value !== undefined && value !== null && value !== "") request = request.eq(key, value);
  });
  const { data, error } = await request;
  if (error) throw error;
  return data || [];
}

async function supabaseListCases(filters = {}) {
  let request = supabase.from("cases").select("*");
  const q = String(filters.q || "").trim();
  if (q) {
    const escaped = q.replace(/[%_,]/g, "\\$&");
    request = request.or([
      `document.ilike.%${escaped}%`,
      `employee_name.ilike.%${escaped}%`,
      `eps.ilike.%${escaped}%`,
      `app_status.ilike.%${escaped}%`,
      `incapacity_type.ilike.%${escaped}%`,
      `responsible.ilike.%${escaped}%`
    ].join(","));
  }
  ["eps", "app_status", "incapacity_type", "responsible"].forEach((field) => {
    if (filters[field]) request = request.eq(field, filters[field]);
  });
  if (filters.from) request = request.gte("start_date", filters.from);
  if (filters.to) request = request.lte("start_date", filters.to);

  const { data, error } = await request.order("start_date", { ascending: false }).order("id", { ascending: false });
  if (error) throw error;
  return (data || []).map(enrichCase);
}

async function supabaseGetCase(id) {
  const { data, error } = await supabase.from("cases").select("*").eq("id", id).single();
  if (error) throw error;
  const { data: history, error: historyError } = await supabase
    .from("management_entries")
    .select("*")
    .eq("case_id", id)
    .order("entry_date", { ascending: false })
    .order("id", { ascending: false });
  if (historyError) throw historyError;
  return { ...enrichCase(data), history: history || [] };
}

async function supabaseUpdateCase(id, patch) {
  const cleanPatch = normalizePatch(patch);
  cleanPatch.updated_at = new Date().toISOString();
  const recalculated = recalculatePatch(cleanPatch);
  const { error } = await supabase.from("cases").update(recalculated).eq("id", id);
  if (error) throw error;
  return supabaseGetCase(id);
}

async function supabaseChangeStatus(id, status) {
  const { error } = await supabase
    .from("cases")
    .update({ app_status: status, updated_at: new Date().toISOString() })
    .eq("id", id);
  if (error) throw error;
  return supabaseGetCase(id);
}

async function supabaseAddManagement(caseId, entry) {
  const record = {
    case_id: Number(caseId),
    entry_date: emptyToNull(entry.entry_date) || today(),
    note: String(entry.note || "").trim(),
    next_action_date: emptyToNull(entry.next_action_date),
    responsible: entry.responsible || ""
  };
  const { error } = await supabase.from("management_entries").insert(record);
  if (error) throw error;

  const patch = {};
  if (record.next_action_date) patch.next_action_date = record.next_action_date;
  if (record.responsible) patch.responsible = record.responsible;
  if (Object.keys(patch).length) await supabaseUpdateCase(caseId, patch);
  return supabaseGetCase(caseId);
}

async function supabaseSaveCatalog(type, item) {
  const record = { type, name: String(item.name || "").trim(), active: item.active === false ? false : true };
  if (!record.name) throw new Error("El nombre es obligatorio");
  if (item.id) {
    const { error } = await supabase.from("catalogs").update(record).eq("id", item.id).eq("type", type);
    if (error) throw error;
  } else {
    const { error } = await supabase.from("catalogs").upsert(record, { onConflict: "type,name" });
    if (error) throw error;
  }
  return api.listCatalog(type);
}

async function supabaseDeleteCatalog(type, id) {
  const { error } = await supabase.from("catalogs").delete().eq("id", id).eq("type", type);
  if (error) throw error;
  return api.listCatalog(type);
}

async function supabaseApplyPayments(payments) {
  let applied = 0;
  let updated = 0;
  let skipped = 0;
  const errors = [];

  for (const [index, payment] of payments.entries()) {
    try {
      const result = await applySinglePayment(payment);
      if (result.updated) updated += 1;
      else applied += 1;
    } catch (error) {
      skipped += 1;
      errors.push({ row: index + 1, message: error.message || "No aplicado" });
    }
  }

  return { ok: true, applied, updated, skipped, errors };
}

async function applySinglePayment(payment) {
  const caseId = Number(payment.case_id || payment.id || 0);
  const amount = Number(payment.amount || 0);
  if (!caseId) throw new Error("ID incapacidad requerido");
  if (!amount || amount <= 0) throw new Error("Valor girado vacio");

  const record = await supabaseGetCase(caseId);
  const paymentDate = emptyToNull(payment.payment_date) || today();
  const { data: existingPayment, error: existingError } = await supabase
    .from("payments")
    .select("*")
    .eq("case_id", caseId)
    .eq("payment_date", paymentDate)
    .maybeSingle();
  if (existingError) throw existingError;

  const previousAmount = Number(existingPayment?.amount || 0);
  const delta = amount - previousAmount;
  const observation = payment.observation || existingPayment?.observation || "";

  if (existingPayment) {
    const { error } = await supabase
      .from("payments")
      .update({ amount, observation, updated_at: new Date().toISOString() })
      .eq("id", existingPayment.id);
    if (error) throw error;
  } else {
    const { error } = await supabase.from("payments").insert({
      case_id: caseId,
      payment_date: paymentDate,
      amount,
      observation
    });
    if (error) throw error;
  }

  const nextRecovered = Math.max(Number(record.recovered_value || 0) + delta, 0);
  const nextRecognized = Math.max(Number(record.recognized_value || 0), nextRecovered);
  const expected = Number(record.expected_value || 0);
  const nextPending = Math.max(expected - nextRecovered, 0);
  const nextStatus = resolveCaseStatus(
    { ...record, recovered_value: nextRecovered, recognized_value: nextRecognized, pending_value: nextPending },
    record.app_status
  );

  const { error: updateError } = await supabase.from("cases").update({
    recovered_value: nextRecovered,
    recognized_value: nextRecognized,
    pending_value: nextPending,
    app_status: nextStatus,
    observations: appendObservation(record.observations, payment.observation),
    updated_at: new Date().toISOString()
  }).eq("id", caseId);
  if (updateError) throw updateError;

  const note = [
    existingPayment ? `Pago actualizado de ${formatCurrency(previousAmount)} a ${formatCurrency(amount)}` : `Pago aplicado por ${formatCurrency(amount)}`,
    `Fecha pago: ${paymentDate}`,
    payment.observation ? `Observacion: ${payment.observation}` : ""
  ].filter(Boolean).join(". ");
  await supabaseAddManagement(caseId, {
    entry_date: paymentDate,
    note,
    responsible: payment.responsible || "Registro manual de pago"
  });

  return { updated: Boolean(existingPayment) };
}

async function supabaseDashboard() {
  const cases = await supabaseListCases({});
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

async function supabaseImportPayments() {
  const file = await chooseFile(".xlsx,.xls,.csv");
  if (!file) return { canceled: true };
  const workbook = await readWorkbook(file);
  const sheetName = workbook.SheetNames.find((name) => normalizeKey(name).includes("pago")) || workbook.SheetNames[0];
  const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
  const payments = rows.map(normalizePaymentRow).filter((item) => item.id || item.amount);
  return supabaseApplyPayments(payments.map((item) => ({ ...item, case_id: item.id, responsible: "Importacion pagos" })));
}

async function supabaseImportBuk() {
  const file = await chooseFile(".xlsx,.xls,.csv");
  if (!file) return { canceled: true };
  const workbook = await readWorkbook(file);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  const normalizedRows = rows
    .map((raw, index) => ({ index, normalized: normalizeBukRow(raw) }))
    .sort((a, b) => String(a.normalized.start_date).localeCompare(String(b.normalized.start_date)));

  let inserted = 0;
  let updated = 0;
  let skipped = 0;
  const errors = [];
  const importChains = new Map();

  for (const { index, normalized } of normalizedRows) {
    try {
      if (!normalized.document || !normalized.start_date || !normalized.end_date || !normalized.incapacity_type) {
        skipped += 1;
        continue;
      }
      if (!normalized.eps) {
        skipped += 1;
        errors.push({ row: index + 2, message: "La fila no tiene EPS" });
        continue;
      }
      const previousDays = await getPreviousDays(importChains, normalized);
      const chargeableDays = calculateChargeableDays(normalized.eps, previousDays, normalized.days);
      const expectedValue = calculateExpectedValue(normalized.salary, chargeableDays, normalized.start_date, normalized.eps);
      const recovered = 0;
      const pending = Math.max(expectedValue - recovered, 0);
      const appStatus = resolveCaseStatus(
        { ...normalized, chargeable_days: chargeableDays, expected_value: expectedValue, recovered_value: recovered, pending_value: pending },
        chargeableDays <= 0 && Number(normalized.days || 0) <= 2 ? "Archivada" : "Pendiente"
      );
      const result = await upsertImportedCase({
        ...normalized,
        chargeable_days: chargeableDays,
        expected_value: expectedValue,
        pending_value: pending,
        app_status: appStatus
      });
      if (result.inserted) inserted += 1;
      else updated += 1;
      updateImportChain(importChains, normalized, previousDays);
    } catch (error) {
      errors.push({ row: index + 2, message: error.message });
    }
  }

  return { ok: true, filePath: file.name, total: rows.length, inserted, updated, skipped, errors };
}

async function upsertImportedCase(item) {
  const { data: existing, error: existingError } = await supabase
    .from("cases")
    .select("*")
    .eq("document", item.document)
    .eq("start_date", item.start_date)
    .eq("end_date", item.end_date)
    .eq("incapacity_type", item.incapacity_type)
    .maybeSingle();
  if (existingError) throw existingError;

  await ensureCatalog("eps", item.eps);
  const payload = normalizeCasePayload(item);
  if (existing) {
    const recovered = Number(existing.recovered_value || 0);
    const pending = Math.max(Number(payload.expected_value || 0) - recovered, 0);
    const patch = {
      ...payload,
      recovered_value: recovered,
      recognized_value: Math.max(Number(existing.recognized_value || 0), recovered),
      pending_value: pending,
      app_status: resolveCaseStatus({ ...payload, recovered_value: recovered, pending_value: pending }, existing.app_status),
      updated_at: new Date().toISOString()
    };
    const { error } = await supabase.from("cases").update(patch).eq("id", existing.id);
    if (error) throw error;
    return { inserted: false, id: existing.id };
  }

  const { data, error } = await supabase.from("cases").insert(payload).select("id").single();
  if (error) throw error;
  return { inserted: true, id: data.id };
}

async function ensureCatalog(type, name) {
  if (!name) return;
  const { error } = await supabase.from("catalogs").upsert(
    { type, name: String(name).trim(), active: true },
    { onConflict: "type,name" }
  );
  if (error) throw error;
}

function exportCasesWorkbook(cases, filename) {
  const rows = cases.map((item) => ({
    ID: item.id,
    Colaborador: item.employee_name,
    Documento: item.document,
    EPS: item.eps,
    "Fecha inicio": item.start_date,
    "Fecha fin": item.end_date,
    Dias: item.days,
    "Dias cobrables": item.chargeable_days,
    Estado: item.app_status,
    Salario: item.salary,
    "Valor esperado": item.expected_value,
    "Valor reconocido": item.recognized_value,
    "Valor recuperado": item.recovered_value,
    "Valor pendiente": item.pending_value,
    Responsable: item.responsible,
    Observaciones: item.observations
  }));
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, sheet, "Incapacidades");
  triggerWorkbook(workbook, filename);
}

function exportPaymentTemplate(cases, filename) {
  const rows = cases
    .filter((item) => Number(item.pending_value || 0) > 0)
    .map((item) => ({
      "ID Incapacidad": item.id,
      Colaborador: item.employee_name,
      Documento: item.document,
      EPS: item.eps,
      Inicio: item.start_date,
      Fin: item.end_date,
      "Valor pendiente": item.pending_value,
      "Valor girado": "",
      "Fecha pago": today(),
      "Observacion pago": ""
    }));
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, sheet, "Pagos");
  triggerWorkbook(workbook, filename);
}

function exportImportTemplate(filename) {
  const rows = [
    {
      "Colaborador - Numero de Documento": "1049629472",
      "Colaborador - Nombre Completo": "Ejemplo Colaborador",
      "Novedades - Codigo": "J00X",
      "Campos Personalizados de Licencia - Soporte Licencia / Incapacidad": "https://...",
      "Novedades - Fecha Inicio": "2026-01-13",
      "Novedades - Fecha de Termino": "2026-01-14",
      "Novedades - Dias Tomados": 2,
      "Novedades - Estado Licencia": "Aprobado",
      "Novedades - Descripcion Tipo": "Incapacidad Medica",
      "Novedades - Justificacion": "",
      "Novedades - Nombre Tipo": "Incapacidad",
      "Novedades - Tipo de": "Incapacidad",
      EPS: "EPS SURA",
      Salario: 1800000
    }
  ];
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(rows, { header: IMPORT_TEMPLATE_HEADERS });
  XLSX.utils.book_append_sheet(workbook, sheet, "Modelo BUK");
  triggerWorkbook(workbook, filename);
}

function normalizePatch(patch) {
  const dateFields = ["start_date", "end_date", "filing_date", "next_action_date"];
  const numericFields = ["salary", "days", "chargeable_days", "expected_value", "recognized_value", "recovered_value", "pending_value"];
  return Object.fromEntries(Object.entries(patch).map(([key, value]) => {
    if (dateFields.includes(key)) return [key, emptyToNull(value)];
    if (numericFields.includes(key)) return [key, Number(value || 0)];
    if (key === "pending_documents") return [key, Boolean(value)];
    return [key, value ?? ""];
  }));
}

function recalculatePatch(patch) {
  if ("expected_value" in patch || "recovered_value" in patch) {
    const expected = Number(patch.expected_value || 0);
    const recovered = Number(patch.recovered_value || 0);
    patch.pending_value = Math.max(expected - recovered, 0);
    patch.app_status = resolveCaseStatus(patch, patch.app_status);
  }
  return patch;
}

function normalizeCasePayload(item) {
  return {
    employee_name: item.employee_name || "Sin nombre",
    document: String(item.document || "").trim(),
    employee_code: item.employee_code || "",
    salary: Number(item.salary || 0),
    eps: item.eps,
    area: item.area || "",
    position: item.position || "",
    incapacity_type: item.incapacity_type || "Sin tipo",
    diagnosis: item.diagnosis || "",
    start_date: emptyToNull(item.start_date),
    end_date: emptyToNull(item.end_date),
    days: Number(item.days || 0),
    chargeable_days: Number(item.chargeable_days || 0),
    buk_status: item.buk_status || "",
    app_status: item.app_status || "Pendiente",
    incapacity_number: item.incapacity_number || "",
    support_url: item.support_url || "",
    filing_date: emptyToNull(item.filing_date),
    filing_number: item.filing_number || "",
    responsible: item.responsible || "",
    expected_value: Number(item.expected_value || 0),
    recognized_value: Number(item.recognized_value || 0),
    recovered_value: Number(item.recovered_value || 0),
    pending_value: Number(item.pending_value || 0),
    observations: item.observations || "",
    next_action_date: emptyToNull(item.next_action_date),
    pending_documents: Boolean(item.pending_documents)
  };
}

function normalizeBukRow(raw) {
  const source = normalizeSource(raw);
  const item = mapFields(source, BUK_FIELD_MAP);
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

function normalizePaymentRow(raw) {
  const source = normalizeSource(raw);
  const item = mapFields(source, PAYMENT_FIELD_MAP);
  return {
    id: Number(clean(item.id)),
    amount: parseMoney(item.amount),
    payment_date: parseDate(item.payment_date),
    observation: clean(item.observation)
  };
}

function normalizeSource(raw) {
  const source = {};
  Object.entries(raw).forEach(([key, value]) => {
    source[normalizeKey(key)] = value;
  });
  return source;
}

function mapFields(source, map) {
  const item = {};
  Object.entries(map).forEach(([field, labels]) => {
    const value = labels.map(normalizeKey).map((key) => source[key]).find((entry) => entry !== undefined && entry !== "");
    item[field] = value ?? "";
  });
  return item;
}

async function getPreviousDays(chains, item) {
  const key = chainKey(item);
  const chain = chains.get(key);
  if (chain && addDays(chain.lastEnd, 1) === item.start_date) return chain.days;

  let total = 0;
  let cursor = addDays(item.start_date, -1);
  const normalizedDiagnosis = normalizeKey(item.diagnosis);
  while (cursor) {
    const { data, error } = await supabase
      .from("cases")
      .select("id,start_date,end_date,days,diagnosis")
      .eq("document", item.document)
      .eq("end_date", cursor)
      .order("start_date", { ascending: false })
      .order("id", { ascending: false })
      .limit(1)
      .maybeSingle();
    if (error) throw error;
    if (!data || normalizeKey(data.diagnosis) !== normalizedDiagnosis) break;
    total += Number(data.days || daysBetweenInclusive(data.start_date, data.end_date));
    cursor = addDays(data.start_date, -1);
  }
  return total;
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

  const minimumWage = MINIMUM_WAGES[Number(String(startDate || "").slice(0, 4))] || 0;
  const monthlyBase =
    minimumWage && monthlySalary > minimumWage
      ? Math.max(monthlySalary * INCAPACITY_RATE, minimumWage)
      : monthlySalary;
  return Math.round((monthlyBase / 30) * days);
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
  return days > 0 && days <= 2 && Number(record.chargeable_days || 0) <= 0 && Number(record.expected_value || 0) <= 0;
}

function appendObservation(current, next) {
  if (!next) return current || "";
  if (!current) return next;
  return `${current}\n${next}`;
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

function emptyToNull(value) {
  return value === "" || value === undefined ? null : value;
}

function isArl(eps) {
  return normalizeKey(eps).includes("arl");
}

function chainKey(item) {
  return `${item.document}|${normalizeKey(item.diagnosis)}`;
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

export const api = {
  authRequired() {
    return useSupabase();
  },
  async getSession() {
    if (!useSupabase()) return null;
    const { data, error } = await supabase.auth.getSession();
    if (error) throw error;
    return data.session;
  },
  onAuthStateChange(callback) {
    if (!useSupabase()) return () => {};
    const { data } = supabase.auth.onAuthStateChange((_event, session) => callback(session));
    return () => data.subscription.unsubscribe();
  },
  async signIn(email, password) {
    if (!useSupabase()) return null;
    const { data, error } = await supabase.auth.signInWithPassword({ email, password });
    if (error) throw error;
    return data.session;
  },
  async signOut() {
    if (!useSupabase()) return;
    const { error } = await supabase.auth.signOut();
    if (error) throw error;
  },
  async updatePassword(password) {
    if (!useSupabase()) return;
    const { error } = await supabase.auth.updateUser({ password });
    if (error) throw error;
  },
  listCases(filters) {
    const native = electronApi();
    if (native) return native.listCases(filters);
    if (useSupabase()) return supabaseListCases(filters);
    return jsonFetch(`/api/cases${query(filters)}`);
  },
  getCase(id) {
    const native = electronApi();
    if (native) return native.getCase(id);
    if (useSupabase()) return supabaseGetCase(id);
    return jsonFetch(`/api/cases/${id}`);
  },
  updateCase(id, patch) {
    const native = electronApi();
    if (native) return native.updateCase(id, patch);
    if (useSupabase()) return supabaseUpdateCase(id, patch);
    return jsonFetch(`/api/cases/${id}`, { method: "PATCH", body: JSON.stringify(patch) });
  },
  changeStatus(id, status) {
    const native = electronApi();
    if (native) return native.changeStatus(id, status);
    if (useSupabase()) return supabaseChangeStatus(id, status);
    return jsonFetch(`/api/cases/${id}/status`, { method: "PATCH", body: JSON.stringify({ status }) });
  },
  async importBuk() {
    const native = electronApi();
    if (native) return native.importBuk("");
    if (useSupabase()) return supabaseImportBuk();
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
    if (native) return native.dashboard();
    if (useSupabase()) return supabaseDashboard();
    return jsonFetch("/api/dashboard");
  },
  addManagement(caseId, entry) {
    const native = electronApi();
    if (native) return native.addManagement(caseId, entry);
    if (useSupabase()) return supabaseAddManagement(caseId, entry);
    return jsonFetch(`/api/cases/${caseId}/management`, { method: "POST", body: JSON.stringify(entry) });
  },
  listCatalog(type) {
    const native = electronApi();
    if (native) return native.listCatalog(type);
    if (useSupabase()) return supabaseAll("catalogs", { type }).then((rows) => rows.sort((a, b) => a.name.localeCompare(b.name)));
    return jsonFetch(`/api/catalogs/${type}`);
  },
  saveCatalog(type, item) {
    const native = electronApi();
    if (native) return native.saveCatalog(type, item);
    if (useSupabase()) return supabaseSaveCatalog(type, item);
    return jsonFetch(`/api/catalogs/${type}`, { method: "POST", body: JSON.stringify(item) });
  },
  deleteCatalog(type, id) {
    const native = electronApi();
    if (native) return native.deleteCatalog(type, id);
    if (useSupabase()) return supabaseDeleteCatalog(type, id);
    return jsonFetch(`/api/catalogs/${type}/${id}`, { method: "DELETE" });
  },
  async exportReport(filters) {
    const native = electronApi();
    if (native) return native.exportReport(filters);
    if (useSupabase()) {
      exportCasesWorkbook(await supabaseListCases(filters || {}), `reporte-incapacidades-${today()}.xlsx`);
      return { ok: true };
    }
    await downloadBlob("/api/reports/export", `reporte-incapacidades-${today()}.xlsx`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(filters || {})
    });
    return { ok: true };
  },
  async downloadImportTemplate() {
    const native = electronApi();
    if (native) return native.downloadImportTemplate();
    if (useSupabase()) {
      exportImportTemplate("modelo-importacion-incapacidades-buk.xlsx");
      return { ok: true };
    }
    await downloadBlob("/api/buk/template", "modelo-importacion-incapacidades-buk.xlsx");
    return { ok: true };
  },
  async downloadPaymentTemplate(filters) {
    if (useSupabase()) {
      exportPaymentTemplate(await supabaseListCases(filters || {}), `plantilla-aplicacion-pagos-${today()}.xlsx`);
      return { ok: true };
    }
    await downloadBlob("/api/payments/template", `plantilla-aplicacion-pagos-${today()}.xlsx`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(filters || {})
    });
    return { ok: true };
  },
  async importPayments() {
    if (useSupabase()) return supabaseImportPayments();
    const file = await chooseFile(".xlsx,.xls,.csv");
    if (!file) return { canceled: true };
    const form = new FormData();
    form.append("file", file);
    const response = await fetch("/api/payments/import", { method: "POST", body: form });
    if (!response.ok) throw new Error(await response.text());
    return response.json();
  },
  applyPayments(payments) {
    if (useSupabase()) return supabaseApplyPayments(payments);
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
