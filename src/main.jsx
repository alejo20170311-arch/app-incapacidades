import React, { useEffect, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";
import {
  AlertTriangle,
  BarChart3,
  CheckCircle2,
  ChevronDown,
  ChevronRight,
  CircleDollarSign,
  ClipboardList,
  Download,
  ExternalLink,
  FileSpreadsheet,
  KeyRound,
  Plus,
  RefreshCw,
  Search,
  Settings,
  LogOut,
  UserRound,
  WalletCards
} from "lucide-react";
import {
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Line,
  LineChart,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis
} from "recharts";
import * as XLSX from "xlsx";
import { api } from "./api";
import "./styles.css";
import mslLogo from "./assets/msl-logo-white.png";

const COLORS = ["#2563eb", "#16a34a", "#f59e0b", "#dc2626", "#7c3aed", "#0891b2", "#475569"];
const STATUS_COLORS = {
  Pendiente: "#f59e0b",
  "Lista para radicar": "#2563eb",
  Radicada: "#0ea5e9",
  "En revision": "#7c3aed",
  Devuelta: "#f97316",
  Corregida: "#14b8a6",
  Pagada: "#16a34a",
  Rechazada: "#dc2626",
  Archivada: "#64748b"
};

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

function App() {
  const [view, setView] = useState("dashboard");
  const [session, setSession] = useState(null);
  const [authReady, setAuthReady] = useState(!api.authRequired());
  const [cases, setCases] = useState([]);
  const [dashboard, setDashboard] = useState(null);
  const [catalogs, setCatalogs] = useState({ eps: [], state: [], responsible: [] });
  const [filters, setFilters] = useState({ q: "" });
  const [selectedId, setSelectedId] = useState(null);
  const [selected, setSelected] = useState(null);
  const [toast, setToast] = useState("");
  const [passwordModalOpen, setPasswordModalOpen] = useState(false);
  const [importModal, setImportModal] = useState(null);
  const [entitiesNormalized, setEntitiesNormalized] = useState(false);

  useEffect(() => {
    document.documentElement.dataset.theme = "light";
    localStorage.setItem("theme", "light");
  }, []);

  useEffect(() => {
    if (!api.authRequired()) return;
    let mounted = true;
    api
      .getSession()
      .then((currentSession) => {
        if (!mounted) return;
        setSession(currentSession);
        setAuthReady(true);
      })
      .catch((error) => {
        if (!mounted) return;
        setToast(cleanError(error));
        setAuthReady(true);
      });
    const unsubscribe = api.onAuthStateChange((nextSession) => {
      setSession(nextSession);
      setAuthReady(true);
    });
    return () => {
      mounted = false;
      unsubscribe();
    };
  }, []);

  useEffect(() => {
    if (!authReady || (api.authRequired() && !session)) return;
    if (entitiesNormalized) {
      refreshAll();
      return;
    }
    api
      .normalizeEntities()
      .catch((error) => setToast(cleanError(error)))
      .finally(() => setEntitiesNormalized(true));
  }, [authReady, session, entitiesNormalized]);

  useEffect(() => {
    if (!authReady || (api.authRequired() && !session)) return;
    loadCases();
  }, [filters, authReady, session]);

  useEffect(() => {
    if (selectedId) loadSelected(selectedId);
  }, [selectedId]);

  async function refreshAll() {
    await Promise.all([loadCases(), loadDashboard(), loadCatalogs()]);
  }

  async function handleLogin(email, password) {
    try {
      const nextSession = await api.signIn(email, password);
      setSession(nextSession);
      setToast("");
    } catch (error) {
      setToast(cleanError(error));
    }
  }

  async function handleLogout() {
    await api.signOut();
    setSession(null);
    setCases([]);
    setDashboard(null);
    setSelected(null);
    setSelectedId(null);
  }

  async function handleChangePassword(password) {
    await api.updatePassword(password);
    setPasswordModalOpen(false);
    setToast("Clave actualizada.");
  }

  async function loadCases() {
    setCases(await api.listCases(filters));
  }

  async function loadDashboard() {
    setDashboard(await api.dashboard());
  }

  async function loadCatalogs() {
    const [eps, state, responsible] = await Promise.all([
      api.listCatalog("eps"),
      api.listCatalog("state"),
      api.listCatalog("responsible")
    ]);
    setCatalogs({ eps, state, responsible });
  }

  async function loadSelected(id) {
    setSelected(await api.getCase(id));
  }

  async function handleImport() {
    const result = await runImportWithProgress("Importar incapacidades BUK", () => api.importBuk(), formatBukImportSummary);
    if (result?.ok) {
      await refreshAll();
    }
  }

  async function downloadTemplate() {
    try {
      const result = await api.downloadImportTemplate();
      if (result?.ok) setToast("Modelo de importacion descargado.");
      return;
    } catch (_error) {
      downloadImportTemplateInBrowser();
      setToast("Modelo de importacion descargado.");
    }
  }

  async function changeStatus(id, status) {
    await api.changeStatus(id, status);
    await refreshAll();
    if (selectedId === id) await loadSelected(id);
  }

  async function saveCase(patch) {
    try {
      await api.updateCase(selected.id, patch);
      setToast("Incapacidad actualizada.");
      await refreshAll();
      await loadSelected(selected.id);
    } catch (error) {
      setToast(cleanError(error));
    }
  }

  async function exportReport(extra = {}) {
    const result = await api.exportReport({ ...filters, ...extra });
    if (result?.ok) setToast(`Reporte exportado con ${result.total} registros.`);
  }

  async function downloadPaymentTemplate(extra = {}) {
    const result = await api.downloadPaymentTemplate({ ...filters, ...extra });
    if (result?.ok) setToast("Plantilla de pagos descargada.");
  }

  async function importPayments() {
    const result = await runImportWithProgress("Importar pagos aplicados", () => api.importPayments(), formatPaymentImportSummary);
    if (result?.ok) {
      await refreshAll();
      if (selectedId) await loadSelected(selectedId);
    }
  }

  async function applyManualPayments(payments) {
    const result = await api.applyPayments(payments);
    setToast(`Pago registrado: ${result.applied} nuevos, ${result.updated || 0} actualizados, ${result.skipped} omitidos.`);
    await refreshAll();
    if (selectedId) await loadSelected(selectedId);
    return result;
  }

  async function runImportWithProgress(title, action, formatSummary) {
    const timers = [];
    setImportModal({
      title,
      progress: 0,
      status: "running",
      message: "Selecciona el archivo de Excel.",
      summary: "",
      errors: [],
      canClose: false
    });
    timers.push(setTimeout(() => setImportModal((current) => advanceImportModal(current, 20, "Leyendo archivo...")), 350));
    timers.push(setTimeout(() => setImportModal((current) => advanceImportModal(current, 55, "Validando datos...")), 1100));
    timers.push(setTimeout(() => setImportModal((current) => advanceImportModal(current, 82, "Guardando informacion...")), 2200));

    try {
      const result = await action();
      timers.forEach(clearTimeout);
      if (result?.canceled) {
        setImportModal({
          title,
          progress: 100,
          status: "warning",
          message: "Importacion cancelada.",
          summary: "No se realizaron cambios.",
          errors: [],
          canClose: true
        });
        return result;
      }
      const errors = normalizeImportErrors(result?.errors || []);
      setImportModal({
        title,
        progress: 100,
        status: errors.length ? "warning" : "success",
        message: errors.length ? "Importacion terminada con observaciones." : "Importacion completada.",
        summary: formatSummary(result),
        errors,
        canClose: true
      });
      return result;
    } catch (error) {
      timers.forEach(clearTimeout);
      setImportModal({
        title,
        progress: 100,
        status: "error",
        message: "No se pudo completar la importacion.",
        summary: "",
        errors: normalizeImportErrors(error),
        canClose: true
      });
      return null;
    }
  }

  if (!authReady) return <LoadingScreen />;

  if (api.authRequired() && !session) {
    return (
      <>
        <LoginView onLogin={handleLogin} />
        {toast && (
          <button className="toast" onClick={() => setToast("")}>
            {toast}
          </button>
        )}
      </>
    );
  }

  return (
    <div className="app-shell">
      <Sidebar view={view} setView={setView} />
      <main className="workspace">
        <header className="topbar">
          <div>
            <p className="eyebrow">Recobro de incapacidades</p>
            <h1>{viewTitle(view)}</h1>
            <span className="subtitle">Mercadeo Sin Limites - Gestion Humana</span>
          </div>
          <div className="top-actions">
            <button className="icon-button" title="Actualizar" onClick={refreshAll}>
              <RefreshCw size={18} />
            </button>
            <button className="secondary" onClick={downloadTemplate}>
              <Download size={18} />
              Modelo importacion
            </button>
            <button className="primary" onClick={handleImport}>
              <FileSpreadsheet size={18} />
              Importar BUK
            </button>
            {api.authRequired() && (
              <AccountMenu
                email={session?.user?.email || "Usuario"}
                onChangePassword={() => setPasswordModalOpen(true)}
                onLogout={handleLogout}
              />
            )}
          </div>
        </header>

        {passwordModalOpen && (
          <PasswordModal
            onCancel={() => setPasswordModalOpen(false)}
            onSave={handleChangePassword}
          />
        )}

        {importModal && (
          <ImportProgressModal
            state={importModal}
            onClose={() => setImportModal(null)}
          />
        )}

        {toast && (
          <button className="toast" onClick={() => setToast("")}>
            {toast}
          </button>
        )}

        {view === "dashboard" && dashboard && (
          <Dashboard dashboard={dashboard} setFilters={setFilters} setView={setView} openCase={setSelectedId} />
        )}

        {view === "cases" && (
          <CasesView
            cases={cases}
            filters={filters}
            setFilters={setFilters}
            catalogs={catalogs}
            openCase={(id) => {
              setSelectedId(id);
              setView("detail");
            }}
            changeStatus={changeStatus}
            exportReport={exportReport}
          />
        )}

        {view === "detail" && (
          <DetailView
            selected={selected}
            catalogs={catalogs}
            onSave={saveCase}
            onStatus={changeStatus}
            onBack={() => setView("cases")}
            onRefresh={async () => {
              await refreshAll();
              if (selectedId) await loadSelected(selectedId);
            }}
          />
        )}

        {view === "reports" && (
          <ReportsView
            catalogs={catalogs}
            exportReport={exportReport}
            downloadPaymentTemplate={downloadPaymentTemplate}
            importPayments={importPayments}
            applyManualPayments={applyManualPayments}
          />
        )}

        {view === "settings" && (
          <SettingsView catalogs={catalogs} refresh={loadCatalogs} />
        )}
      </main>
    </div>
  );
}

function AccountMenu({ email, onChangePassword, onLogout }) {
  const [open, setOpen] = useState(false);

  function choose(action) {
    setOpen(false);
    action();
  }

  return (
    <div className="account-menu">
      <button className="account-trigger" onClick={() => setOpen((current) => !current)}>
        <span className="account-avatar">
          <UserRound size={20} />
        </span>
        <span className="account-copy">
          <strong>{emailName(email)}</strong>
          <small>{email}</small>
        </span>
        <ChevronDown size={16} />
      </button>
      {open && (
        <div className="account-popover">
          <button onClick={() => choose(onChangePassword)}>
            <KeyRound size={17} />
            Cambiar clave
          </button>
          <button onClick={() => choose(onLogout)}>
            <LogOut size={17} />
            Cerrar sesion
          </button>
        </div>
      )}
    </div>
  );
}

function PasswordModal({ onCancel, onSave }) {
  const [password, setPassword] = useState("");
  const [confirm, setConfirm] = useState("");
  const [error, setError] = useState("");
  const [saving, setSaving] = useState(false);

  async function submit(event) {
    event.preventDefault();
    if (password.length < 8) {
      setError("La clave debe tener minimo 8 caracteres.");
      return;
    }
    if (password !== confirm) {
      setError("Las claves no coinciden.");
      return;
    }
    setSaving(true);
    try {
      await onSave(password);
    } catch (saveError) {
      setError(cleanError(saveError));
      setSaving(false);
    }
  }

  return (
    <div className="modal-backdrop">
      <form className="password-modal" onSubmit={submit}>
        <h2>Cambiar clave</h2>
        <label>
          Nueva clave
          <input
            type="password"
            autoComplete="new-password"
            value={password}
            onChange={(event) => setPassword(event.target.value)}
          />
        </label>
        <label>
          Confirmar clave
          <input
            type="password"
            autoComplete="new-password"
            value={confirm}
            onChange={(event) => setConfirm(event.target.value)}
          />
        </label>
        {error && <p className="form-error">{error}</p>}
        <div className="modal-actions">
          <button className="secondary" type="button" onClick={onCancel}>
            Cancelar
          </button>
          <button className="primary" type="submit" disabled={saving}>
            {saving ? "Guardando..." : "Guardar clave"}
          </button>
        </div>
      </form>
    </div>
  );
}

function emailName(email) {
  return String(email || "Usuario").split("@")[0];
}

function LoadingScreen() {
  return (
    <main className="login-shell">
      <div className="login-panel">
        <img src={mslLogo} alt="MSL Group" />
        <p>Cargando...</p>
      </div>
    </main>
  );
}

function ImportProgressModal({ state, onClose }) {
  const statusLabel = {
    running: "Procesando",
    success: "Completado",
    warning: "Con observaciones",
    error: "Error"
  }[state.status] || "Procesando";

  return (
    <div className="modal-backdrop">
      <div className={`import-modal ${state.status}`}>
        <div className="import-modal-header">
          <div>
            <h2>{state.title}</h2>
            <span>{statusLabel}</span>
          </div>
          <strong>{state.progress}%</strong>
        </div>
        <div className="progress-track">
          <div className="progress-fill" style={{ width: `${state.progress}%` }} />
        </div>
        <p className="import-message">{state.message}</p>
        {state.summary && <p className="import-summary">{state.summary}</p>}
        {state.errors?.length > 0 && (
          <div className="import-errors">
            <strong>Errores y observaciones</strong>
            <ul>
              {state.errors.map((error, index) => (
                <li key={`${error.row || "general"}-${index}`}>
                  {error.row ? `Fila ${error.row}: ` : ""}
                  {error.message}
                </li>
              ))}
            </ul>
          </div>
        )}
        <div className="modal-actions">
          <button className="primary" onClick={onClose} disabled={!state.canClose}>
            Cerrar
          </button>
        </div>
      </div>
    </div>
  );
}

function LoginView({ onLogin }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);

  async function submit(event) {
    event.preventDefault();
    if (!email.trim() || !password) return;
    setLoading(true);
    try {
      await onLogin(email.trim(), password);
    } finally {
      setLoading(false);
    }
  }

  return (
    <main className="login-shell">
      <form className="login-panel" onSubmit={submit}>
        <img src={mslLogo} alt="MSL Group" />
        <div>
          <p className="eyebrow">Recobro de incapacidades</p>
          <span className="subtitle">Gestion Humana</span>
        </div>
        <label>
          Usuario
          <input
            type="email"
            autoComplete="email"
            value={email}
            onChange={(event) => setEmail(event.target.value)}
          />
        </label>
        <label>
          Clave
          <input
            type="password"
            autoComplete="current-password"
            value={password}
            onChange={(event) => setPassword(event.target.value)}
          />
        </label>
        <button className="primary" type="submit" disabled={loading}>
          <UserRound size={18} />
          {loading ? "Ingresando..." : "Ingresar"}
        </button>
      </form>
    </main>
  );
}

function Sidebar({ view, setView }) {
  const items = [
    ["dashboard", BarChart3, "Dashboard"],
    ["cases", ClipboardList, "Incapacidades"],
    ["reports", Download, "Reportes"],
    ["settings", Settings, "Configuracion"]
  ];
  return (
    <aside className="sidebar">
      <div className="brand">
        <img className="brand-logo" src={mslLogo} alt="MSL Group" />
        <div>
          <strong>MSL Group</strong>
          <span>Gestion Humana</span>
        </div>
      </div>
      <nav>
        {items.map(([id, Icon, label]) => (
          <button key={id} className={view === id ? "active" : ""} onClick={() => setView(id)}>
            <Icon size={19} />
            {label}
          </button>
        ))}
      </nav>
    </aside>
  );
}

function Dashboard({ dashboard, setFilters, setView, openCase }) {
  const { totals } = dashboard;
  const cards = [
    ["Activas", totals.active, ClipboardList],
    ["Pendientes", totals.pending, AlertTriangle],
    ["Radicadas", totals.filed, FileSpreadsheet],
    ["Pagadas", totals.paid, CheckCircle2],
    ["Devueltas", totals.returned, RefreshCw],
    ["Rechazadas", totals.rejected, AlertTriangle],
    ["Valor pendiente", money(totals.pendingValue), WalletCards],
    ["Valor recuperado", money(totals.recoveredValue), CircleDollarSign]
  ];

  return (
    <section className="page-stack">
      <div className="metric-grid">
        {cards.map(([label, value, Icon]) => (
          <div className="metric" key={label}>
            <Icon size={21} />
            <span>{label}</span>
            <strong>{value}</strong>
          </div>
        ))}
      </div>

      <div className="dashboard-grid">
        <ChartPanel title="Incapacidades por EPS">
          <ResponsiveContainer width="100%" height={260}>
            <BarChart data={dashboard.byEps.slice(0, 8)}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} />
              <XAxis dataKey="name" tick={{ fontSize: 11 }} />
              <YAxis allowDecimals={false} />
              <Tooltip />
              <Bar dataKey="value" radius={[6, 6, 0, 0]}>
                {dashboard.byEps.map((_, index) => <Cell key={index} fill={COLORS[index % COLORS.length]} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </ChartPanel>
        <ChartPanel title="Valor pendiente por EPS">
          <ResponsiveContainer width="100%" height={260}>
            <BarChart data={dashboard.pendingByEps.slice(0, 8)}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} />
              <XAxis dataKey="name" tick={{ fontSize: 11 }} />
              <YAxis tickFormatter={shortMoney} />
              <Tooltip formatter={(value) => money(value)} />
              <Bar dataKey="value" fill="#0ea5e9" radius={[6, 6, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </ChartPanel>
        <ChartPanel title="Incapacidades por estado">
          <ResponsiveContainer width="100%" height={260}>
            <PieChart>
              <Pie data={dashboard.byStatus} dataKey="value" nameKey="name" innerRadius={52} outerRadius={92} paddingAngle={3}>
                {dashboard.byStatus.map((entry, index) => (
                  <Cell key={entry.name} fill={STATUS_COLORS[entry.name] || COLORS[index % COLORS.length]} />
                ))}
              </Pie>
              <Tooltip />
            </PieChart>
          </ResponsiveContainer>
        </ChartPanel>
        <ChartPanel title="Recuperacion mensual">
          <ResponsiveContainer width="100%" height={260}>
            <LineChart data={dashboard.monthlyRecovery}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} />
              <XAxis dataKey="name" />
              <YAxis tickFormatter={shortMoney} />
              <Tooltip formatter={(value) => money(value)} />
              <Line type="monotone" dataKey="value" stroke="#16a34a" strokeWidth={3} dot={false} />
            </LineChart>
          </ResponsiveContainer>
        </ChartPanel>
      </div>

      <div className="split-grid">
        <ChartPanel title="Cartera por antiguedad">
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={dashboard.aging}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} />
              <XAxis dataKey="name" />
              <YAxis tickFormatter={shortMoney} />
              <Tooltip formatter={(value) => money(value)} />
              <Bar dataKey="value" fill="#f59e0b" radius={[6, 6, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </ChartPanel>
        <div className="panel">
          <div className="panel-header">
            <h2>Alertas</h2>
            <span>{dashboard.alerts.length}</span>
          </div>
          <div className="alert-list">
            {dashboard.alerts.slice(0, 8).map((alert, index) => (
              <button
                key={`${alert.caseId}-${index}`}
                onClick={() => {
                  openCase(alert.caseId);
                  setView("detail");
                }}
              >
                <AlertTriangle size={17} />
                <div>
                  <strong>{alert.type}</strong>
                  <span>{alert.message}</span>
                </div>
                <ChevronRight size={16} />
              </button>
            ))}
            {!dashboard.alerts.length && <p className="empty">Sin alertas por ahora.</p>}
          </div>
        </div>
      </div>
    </section>
  );
}

function CasesView({ cases, filters, setFilters, catalogs, openCase, changeStatus, exportReport }) {
  return (
    <section className="page-stack">
      <FilterBar filters={filters} setFilters={setFilters} catalogs={catalogs} />
      <div className="table-actions">
        <button className="secondary" onClick={() => exportReport({})}>
          <Download size={17} />
          Exportar vista
        </button>
      </div>
      <div className="data-table">
        <table>
          <thead>
            <tr>
              <th>Colaborador</th>
              <th>Documento</th>
              <th>EPS</th>
              <th>Inicio</th>
              <th>Dias</th>
              <th>Estado</th>
              <th>Valor pendiente</th>
              <th>Proxima gestion</th>
              <th></th>
            </tr>
          </thead>
          <tbody>
            {cases.map((item) => (
              <tr key={item.id}>
                <td>
                  <strong>{item.employee_name}</strong>
                  <span>{item.incapacity_type}</span>
                </td>
                <td>{item.document}</td>
                <td>{item.eps}</td>
                <td>{item.start_date}</td>
                <td>{item.days}</td>
                <td>
                  <select value={item.app_status} onChange={(event) => changeStatus(item.id, event.target.value)}>
                    {catalogs.state.map((state) => (
                      <option key={state.id}>{state.name}</option>
                    ))}
                  </select>
                </td>
                <td>{money(item.pending_value)}</td>
                <td>{item.next_action_date || "-"}</td>
                <td>
                  <button className="row-button" onClick={() => openCase(item.id)}>
                    Abrir
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        {!cases.length && <p className="empty">No hay incapacidades con los filtros actuales.</p>}
      </div>
    </section>
  );
}

function FilterBar({ filters, setFilters, catalogs }) {
  return (
    <div className="filters">
      <label className="search-box">
        <Search size={18} />
        <input value={filters.q || ""} onChange={(event) => setFilters({ ...filters, q: event.target.value })} placeholder="Buscar documento, nombre, EPS, estado, tipo o responsable" />
      </label>
      <select value={filters.eps || ""} onChange={(event) => setFilters({ ...filters, eps: event.target.value })}>
        <option value="">Todas las EPS</option>
        {catalogs.eps.map((item) => (
          <option key={item.id}>{item.name}</option>
        ))}
      </select>
      <select value={filters.app_status || ""} onChange={(event) => setFilters({ ...filters, app_status: event.target.value })}>
        <option value="">Todos los estados</option>
        {catalogs.state.map((item) => (
          <option key={item.id}>{item.name}</option>
        ))}
      </select>
      <input type="date" value={filters.from || ""} onChange={(event) => setFilters({ ...filters, from: event.target.value })} />
      <input type="date" value={filters.to || ""} onChange={(event) => setFilters({ ...filters, to: event.target.value })} />
    </div>
  );
}

function DetailView({ selected, catalogs, onSave, onStatus, onBack, onRefresh }) {
  const [draft, setDraft] = useState(null);
  const [entry, setEntry] = useState({ entry_date: today(), note: "", next_action_date: "", responsible: "" });

  useEffect(() => setDraft(selected), [selected]);

  if (!selected || !draft) return <p className="empty">Selecciona una incapacidad.</p>;

  async function addEntry() {
    if (!entry.note.trim()) return;
    await api.addManagement(selected.id, entry);
    setEntry({ entry_date: today(), note: "", next_action_date: "", responsible: entry.responsible });
    await onRefresh();
  }

  return (
    <section className="detail-layout">
      <div className="detail-main">
        <div className="panel">
          <div className="detail-title">
            <button className="secondary" onClick={onBack}>Volver</button>
            <div>
              <h2>{selected.employee_name}</h2>
              <p>{selected.document} · {selected.eps}</p>
            </div>
            <button className="primary" onClick={() => api.openSupport(selected.support_url)} disabled={!selected.support_url}>
              <ExternalLink size={17} />
              Abrir soporte
            </button>
          </div>
          <div className="status-strip">
            {catalogs.state.map((state) => (
              <button
                key={state.id}
                className={selected.app_status === state.name ? "active" : ""}
                onClick={() => onStatus(selected.id, state.name)}
              >
                {state.name}
              </button>
            ))}
          </div>
          <EditableForm draft={draft} setDraft={setDraft} catalogs={catalogs} onSave={() => onSave(draft)} />
        </div>
      </div>
      <aside className="detail-side">
        <div className="panel">
          <h2>Registrar gestion</h2>
          <div className="form-grid one">
            <label>Fecha<input type="date" value={entry.entry_date} onChange={(event) => setEntry({ ...entry, entry_date: event.target.value })} /></label>
            <label>Responsable<input list="responsibles" value={entry.responsible} onChange={(event) => setEntry({ ...entry, responsible: event.target.value })} /></label>
            <label>Proxima gestion<input type="date" value={entry.next_action_date} onChange={(event) => setEntry({ ...entry, next_action_date: event.target.value })} /></label>
            <label>Observacion<textarea value={entry.note} onChange={(event) => setEntry({ ...entry, note: event.target.value })} /></label>
            <button className="primary" onClick={addEntry}><Plus size={17} /> Guardar gestion</button>
          </div>
        </div>
        <div className="panel">
          <h2>Historial</h2>
          <div className="timeline">
            {selected.history.map((item) => (
              <article key={item.id}>
                <time>{item.entry_date}</time>
                <p>{item.note}</p>
                <span>{item.responsible || "Sin responsable"} {item.next_action_date ? `· Proxima: ${item.next_action_date}` : ""}</span>
              </article>
            ))}
            {!selected.history.length && <p className="empty">Sin gestiones registradas.</p>}
          </div>
        </div>
      </aside>
      <datalist id="responsibles">
        {catalogs.responsible.map((item) => <option key={item.id} value={item.name} />)}
      </datalist>
    </section>
  );
}

function EditableForm({ draft, setDraft, catalogs, onSave }) {
  const set = (key, value) => setDraft({ ...draft, [key]: value });
  const fields = useMemo(() => [
    ["employee_name", "Colaborador"],
    ["document", "Documento"],
    ["employee_code", "Codigo empleado"],
    ["salary", "Salario", "number"],
    ["eps", "EPS"],
    ["area", "Area"],
    ["position", "Cargo"],
    ["incapacity_type", "Tipo de incapacidad"],
    ["diagnosis", "Diagnostico"],
    ["start_date", "Fecha inicio", "date"],
    ["end_date", "Fecha fin", "date"],
    ["days", "Dias", "number"],
    ["chargeable_days", "Dias cobrables", "number"],
    ["buk_status", "Estado BUK"],
    ["incapacity_number", "Numero incapacidad"],
    ["filing_date", "Fecha radicacion", "date"],
    ["filing_number", "Numero radicado"],
    ["responsible", "Responsable"],
    ["expected_value", "Valor esperado", "number"],
    ["recognized_value", "Valor reconocido", "number"],
    ["recovered_value", "Valor recuperado", "number"],
    ["pending_value", "Valor pendiente", "number"],
    ["next_action_date", "Proxima gestion", "date"],
    ["support_url", "URL soporte"]
  ], []);

  return (
    <>
      <div className="form-grid">
        {fields.map(([key, label, type]) => (
          <label key={key}>
            {label}
            <input
              list={key === "eps" ? "eps-list" : key === "responsible" ? "responsibles" : undefined}
              type={type || "text"}
              value={draft[key] ?? ""}
              onChange={(event) => set(key, type === "number" ? Number(event.target.value) : event.target.value)}
            />
          </label>
        ))}
        <label className="checkbox-line">
          <input type="checkbox" checked={Boolean(draft.pending_documents)} onChange={(event) => set("pending_documents", event.target.checked)} />
          Documentos pendientes
        </label>
        <label className="wide">Observaciones<textarea value={draft.observations || ""} onChange={(event) => set("observations", event.target.value)} /></label>
      </div>
      <button className="primary save-button" onClick={onSave}>Guardar cambios</button>
      <datalist id="eps-list">
        {catalogs.eps.map((item) => <option key={item.id} value={item.name} />)}
      </datalist>
    </>
  );
}

function ReportsView({ catalogs, exportReport, downloadPaymentTemplate, importPayments, applyManualPayments }) {
  const reportButtons = [
    ["Pendientes", { app_status: "Pendiente" }],
    ["Pagadas", { app_status: "Pagada" }],
    ["Radicadas", { app_status: "Radicada" }],
    ["Devueltas", { app_status: "Devuelta" }],
    ["Rechazadas", { app_status: "Rechazada" }]
  ];
  return (
    <section className="page-stack">
      <PaymentRegistration catalogs={catalogs} applyManualPayments={applyManualPayments} />
      <div className="panel report-panel">
        <h2>Aplicacion de pagos</h2>
        <div className="report-grid compact">
          <button className="report-card featured" onClick={() => downloadPaymentTemplate({})}>
            <Download size={20} />
            Descargar plantilla pagos
          </button>
          <button className="report-card featured" onClick={importPayments}>
            <FileSpreadsheet size={20} />
            Importar pagos aplicados
          </button>
        </div>
      </div>
      <div className="panel report-panel">
        <h2>Exportar a Excel</h2>
        <div className="report-grid">
          {reportButtons.map(([label, filter]) => (
            <button key={label} className="report-card" onClick={() => exportReport(filter)}>
              <Download size={20} />
              {label}
            </button>
          ))}
          {catalogs.eps.map((eps) => (
            <button key={eps.id} className="report-card" onClick={() => exportReport({ eps: eps.name })}>
              <FileSpreadsheet size={20} />
              {eps.name}
            </button>
          ))}
        </div>
      </div>
    </section>
  );
}

function PaymentRegistration({ catalogs, applyManualPayments }) {
  const [entity, setEntity] = useState("");
  const [paymentDate, setPaymentDate] = useState(today());
  const [transactionValue, setTransactionValue] = useState("");
  const [observation, setObservation] = useState("");
  const [pendingCases, setPendingCases] = useState([]);
  const [selected, setSelected] = useState({});
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    if (!entity) {
      setPendingCases([]);
      setSelected({});
      return;
    }
    let cancelled = false;
    setLoading(true);
    api
      .listCases({ eps: entity })
      .then((rows) => {
        if (cancelled) return;
        const pending = rows.filter((item) => Number(item.pending_value || 0) > 0);
        setPendingCases(pending);
        setSelected({});
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [entity]);

  const selectedRows = pendingCases.filter((item) => selected[item.id]?.checked);
  const selectedTotal = selectedRows.reduce((sum, item) => sum + Number(selected[item.id]?.amount || 0), 0);
  const difference = Number(transactionValue || 0) - selectedTotal;

  function toggleCase(item, checked) {
    setSelected((current) => ({
      ...current,
      [item.id]: {
        checked,
        amount: current[item.id]?.amount ?? Number(item.pending_value || 0)
      }
    }));
  }

  function setAmount(item, amount) {
    setSelected((current) => ({
      ...current,
      [item.id]: {
        checked: current[item.id]?.checked ?? true,
        amount: Number(amount || 0)
      }
    }));
  }

  async function submitPayment() {
    const payments = selectedRows
      .map((item) => ({
        case_id: item.id,
        amount: Number(selected[item.id]?.amount || 0),
        payment_date: paymentDate,
        observation
      }))
      .filter((item) => item.amount > 0);
    if (!entity || !paymentDate || !payments.length) return;
    await applyManualPayments(payments);
    setSelected({});
    setTransactionValue("");
    setObservation("");
    setPendingCases((await api.listCases({ eps: entity })).filter((item) => Number(item.pending_value || 0) > 0));
  }

  return (
    <div className="panel payment-panel">
      <h2>Registrar pago</h2>
      <div className="payment-form">
        <label>
          Entidad
          <select value={entity} onChange={(event) => setEntity(event.target.value)}>
            <option value="">Selecciona entidad</option>
            {catalogs.eps.map((item) => (
              <option key={item.id}>{item.name}</option>
            ))}
          </select>
        </label>
        <label>
          Fecha pago
          <input type="date" value={paymentDate} onChange={(event) => setPaymentDate(event.target.value)} />
        </label>
        <label>
          Valor transaccion
          <input type="number" value={transactionValue} onChange={(event) => setTransactionValue(event.target.value)} placeholder="0" />
        </label>
        <label>
          Observacion
          <input value={observation} onChange={(event) => setObservation(event.target.value)} placeholder="Ej. Giro Famisanar julio" />
        </label>
      </div>

      <div className="payment-summary">
        <div><span>Seleccionado</span><strong>{money(selectedTotal)}</strong></div>
        <div><span>Valor transaccion</span><strong>{money(transactionValue)}</strong></div>
        <div className={difference === 0 ? "balanced" : "unbalanced"}><span>Diferencia</span><strong>{money(difference)}</strong></div>
      </div>

      <div className="data-table payment-table">
        <table>
          <thead>
            <tr>
              <th></th>
              <th>Colaborador</th>
              <th>Documento</th>
              <th>Inicio</th>
              <th>Fin</th>
              <th>Pendiente</th>
              <th>Valor pagado</th>
            </tr>
          </thead>
          <tbody>
            {pendingCases.map((item) => (
              <tr key={item.id}>
                <td>
                  <input
                    type="checkbox"
                    checked={Boolean(selected[item.id]?.checked)}
                    onChange={(event) => toggleCase(item, event.target.checked)}
                  />
                </td>
                <td>
                  <strong>{item.employee_name}</strong>
                  <span>{item.incapacity_type}</span>
                </td>
                <td>{item.document}</td>
                <td>{item.start_date}</td>
                <td>{item.end_date}</td>
                <td>{money(item.pending_value)}</td>
                <td>
                  <input
                    type="number"
                    value={selected[item.id]?.amount ?? ""}
                    placeholder={Number(item.pending_value || 0)}
                    onChange={(event) => setAmount(item, event.target.value)}
                  />
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        {!entity && <p className="empty">Selecciona una entidad para ver incapacidades pendientes de pago.</p>}
        {entity && loading && <p className="empty">Cargando incapacidades pendientes...</p>}
        {entity && !loading && !pendingCases.length && <p className="empty">No hay incapacidades pendientes para esta entidad.</p>}
      </div>
      <button className="primary save-button" onClick={submitPayment} disabled={!selectedRows.length || !paymentDate}>
        Registrar pago seleccionado
      </button>
    </div>
  );
}

function SettingsView({ catalogs, refresh }) {
  return (
    <section className="settings-grid">
      <CatalogEditor title="EPS" type="eps" items={catalogs.eps} refresh={refresh} />
      <CatalogEditor title="Responsables" type="responsible" items={catalogs.responsible} refresh={refresh} />
      <CatalogEditor title="Estados" type="state" items={catalogs.state} refresh={refresh} />
    </section>
  );
}

function CatalogEditor({ title, type, items, refresh }) {
  const [name, setName] = useState("");
  async function add() {
    if (!name.trim()) return;
    await api.saveCatalog(type, { name });
    setName("");
    await refresh();
  }
  async function remove(id) {
    await api.deleteCatalog(type, id);
    await refresh();
  }
  async function edit(item) {
    const next = prompt(`Editar ${title}`, item.name);
    if (!next || next.trim() === item.name) return;
    await api.saveCatalog(type, { ...item, name: next.trim() });
    await refresh();
  }
  return (
    <div className="panel">
      <h2>{title}</h2>
      <div className="catalog-add">
        <input value={name} onChange={(event) => setName(event.target.value)} placeholder="Nuevo valor" />
        <button className="primary" onClick={add}><Plus size={17} /></button>
      </div>
      <div className="catalog-list">
        {items.map((item) => (
          <div key={item.id}>
            <span>{item.name}</span>
            <div className="catalog-actions">
              <button onClick={() => edit(item)}>Editar</button>
              <button onClick={() => remove(item.id)}>Eliminar</button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function ChartPanel({ title, children }) {
  return (
    <div className="panel chart-panel">
      <h2>{title}</h2>
      {children}
    </div>
  );
}

function viewTitle(view) {
  return {
    dashboard: "Dashboard",
    cases: "Seguimiento",
    detail: "Detalle de incapacidad",
    reports: "Reportes",
    settings: "Configuracion"
  }[view];
}

function money(value) {
  return new Intl.NumberFormat("es-CO", { style: "currency", currency: "COP", maximumFractionDigits: 0 }).format(Number(value || 0));
}

function shortMoney(value) {
  const amount = Number(value || 0);
  if (amount >= 1000000) return `$${Math.round(amount / 1000000)}M`;
  if (amount >= 1000) return `$${Math.round(amount / 1000)}K`;
  return `$${amount}`;
}

function downloadImportTemplateInBrowser() {
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
  const data = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  const blob = new Blob([data], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "modelo-importacion-incapacidades-buk.xlsx";
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function today() {
  return new Date().toISOString().slice(0, 10);
}

function cleanError(error) {
  const text = String(error?.message || error || "No se pudo completar la accion");
  try {
    const parsed = JSON.parse(text);
    return parsed.error || text;
  } catch (_parseError) {
    return text.replace(/^Error:\s*/i, "");
  }
}

function advanceImportModal(current, progress, message) {
  if (!current || current.status !== "running") return current;
  return {
    ...current,
    progress: Math.max(current.progress, progress),
    message
  };
}

function formatBukImportSummary(result = {}) {
  return `Total: ${result.total || 0}. Nuevos: ${result.inserted || 0}. Actualizados: ${result.updated || 0}. Omitidos: ${result.skipped || 0}.`;
}

function formatPaymentImportSummary(result = {}) {
  return `Nuevos: ${result.applied || 0}. Actualizados: ${result.updated || 0}. Omitidos: ${result.skipped || 0}.`;
}

function normalizeImportErrors(input) {
  if (Array.isArray(input)) {
    return input.map((item) => ({
      row: item.row,
      message: String(item.message || item.error || item)
    }));
  }

  const text = String(input?.message || input || "No se pudo completar la importacion");
  try {
    const parsed = JSON.parse(text);
    if (Array.isArray(parsed.errors)) return normalizeImportErrors(parsed.errors);
    return [{ message: parsed.error || parsed.message || text }];
  } catch (_parseError) {
    return [{ message: text.replace(/^Error:\s*/i, "") }];
  }
}

createRoot(document.getElementById("root")).render(<App />);
