const XLSX = require("xlsx");

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

const PAYMENT_TEMPLATE_HEADERS = [
  "ID incapacidad",
  "Documento",
  "Colaborador",
  "EPS",
  "Tipo incapacidad",
  "Fecha inicio",
  "Fecha fin",
  "Dias",
  "Dias cobrables",
  "Estado recobro",
  "Numero radicado",
  "Valor esperado",
  "Valor recuperado actual",
  "Valor pendiente",
  "Valor girado",
  "Fecha pago",
  "Observacion pago"
];

function exportCasesReport(cases, filePath) {
  const rows = cases.map((item) => ({
    Documento: item.document,
    Colaborador: item.employee_name,
    "Codigo empleado": item.employee_code,
    Salario: item.salary,
    EPS: item.eps,
    Area: item.area,
    Cargo: item.position,
    "Tipo incapacidad": item.incapacity_type,
    Diagnostico: item.diagnosis,
    Inicio: item.start_date,
    Fin: item.end_date,
    Dias: item.days,
    "Dias cobrables": item.chargeable_days,
    "Estado BUK": item.buk_status,
    "Estado recobro": item.app_status,
    "Numero incapacidad": item.incapacity_number,
    "URL soporte": item.support_url,
    "Fecha radicacion": item.filing_date,
    "Numero radicado": item.filing_number,
    Responsable: item.responsible,
    "Valor esperado": item.expected_value,
    "Valor reconocido": item.recognized_value,
    "Valor recuperado": item.recovered_value,
    "Valor pendiente": item.pending_value,
    "Proxima gestion": item.next_action_date,
    Observaciones: item.observations
  }));
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, sheet, "Incapacidades");
  XLSX.writeFile(workbook, filePath);
}

function exportImportTemplate(filePath, entities = []) {
  const rows = [
    {
      "Colaborador - Numero de Documento": "1049629472",
      "Colaborador - Nombre Completo": "Ejemplo Colaborador",
      "Novedades - Codigo": "J00X",
      "Campos Personalizados de Licencia - Soporte Licencia / Incapacidad": "https://bukwebapp-enterprise-colombia.s3.us-east-2.amazonaws.com/...",
      "Novedades - Fecha Inicio": "2026-01-13",
      "Novedades - Fecha de Termino": "2026-01-14",
      "Novedades - Dias Tomados": 2,
      "Novedades - Estado Licencia": "Aprobado",
      "Novedades - Descripcion Tipo": "Incapacidad Medica",
      "Novedades - Justificacion": "",
      "Novedades - Nombre Tipo": "Incapacidad",
      "Novedades - Tipo de": "Incapacidad",
      EPS: "SURA",
      Salario: 1800000
    }
  ];
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(rows, { header: IMPORT_TEMPLATE_HEADERS });
  sheet["!autofilter"] = { ref: XLSX.utils.encode_range(XLSX.utils.decode_range(sheet["!ref"])) };
  sheet["!cols"] = IMPORT_TEMPLATE_HEADERS.map((header) => ({ wch: Math.max(16, Math.min(58, header.length + 3)) }));
  XLSX.utils.book_append_sheet(workbook, sheet, "Modelo BUK");
  appendEntitiesSheet(workbook, entities);
  XLSX.writeFile(workbook, filePath);
}

function exportPaymentTemplate(cases, filePath) {
  const rows = cases.map((item) => ({
    "ID incapacidad": item.id,
    Documento: item.document,
    Colaborador: item.employee_name,
    EPS: item.eps,
    "Tipo incapacidad": item.incapacity_type,
    "Fecha inicio": item.start_date,
    "Fecha fin": item.end_date,
    Dias: item.days,
    "Dias cobrables": item.chargeable_days,
    "Estado recobro": item.app_status,
    "Numero radicado": item.filing_number,
    "Valor esperado": item.expected_value,
    "Valor recuperado actual": item.recovered_value,
    "Valor pendiente": item.pending_value,
    "Valor girado": "",
    "Fecha pago": "",
    "Observacion pago": ""
  }));
  const workbook = XLSX.utils.book_new();
  const sheet = XLSX.utils.json_to_sheet(rows, { header: PAYMENT_TEMPLATE_HEADERS });
  sheet["!autofilter"] = { ref: XLSX.utils.encode_range(XLSX.utils.decode_range(sheet["!ref"])) };
  sheet["!cols"] = PAYMENT_TEMPLATE_HEADERS.map((header) => ({ wch: Math.max(14, Math.min(32, header.length + 4)) }));
  XLSX.utils.book_append_sheet(workbook, sheet, "Aplicar pagos");
  const notes = [
    { Campo: "ID incapacidad", Uso: "No modificar. La app lo usa para saber a que incapacidad aplicar el pago." },
    { Campo: "Valor girado", Uso: "Escribir el valor que corresponde a esa incapacidad dentro de la transaccion de la EPS/ARL/AFP." },
    { Campo: "Fecha pago", Uso: "Fecha real del giro o de la transaccion. Formato recomendado: AAAA-MM-DD." },
    { Campo: "Observacion pago", Uso: "Opcional. Ejemplo: Transaccion Famisanar 5 millones." }
  ];
  const instructions = XLSX.utils.json_to_sheet(notes);
  instructions["!cols"] = [{ wch: 24 }, { wch: 96 }];
  XLSX.utils.book_append_sheet(workbook, instructions, "Instrucciones");
  XLSX.writeFile(workbook, filePath);
}

function appendEntitiesSheet(workbook, entities) {
  const rows = entities.length
    ? entities.map((item) => ({ Entidad: item.name || item }))
    : [{ Entidad: "SURA" }, { Entidad: "ARL SURA" }];
  const sheet = XLSX.utils.json_to_sheet(rows, { header: ["Entidad"] });
  sheet["!autofilter"] = { ref: XLSX.utils.encode_range(XLSX.utils.decode_range(sheet["!ref"])) };
  sheet["!cols"] = [{ wch: 58 }];
  XLSX.utils.book_append_sheet(workbook, sheet, "Entidades");
}

module.exports = { exportCasesReport, exportImportTemplate, exportPaymentTemplate };
