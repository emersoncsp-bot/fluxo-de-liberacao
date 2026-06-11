import { useState, useCallback } from "react";

// ─── MOCK USERS ───────────────────────────────────────────────────────────────
const USERS = [
  { email: "qualidade.tecnica@empresa.com",      password: "123456", name: "Ana Paula Silva",  dept: "Controle da Qualidade [Área técnica]" },
  { email: "qualidade.lib@empresa.com",          password: "123456", name: "Carlos Mendes",    dept: "Controle da Qualidade [Liberação Intermediária]" },
  { email: "planejamento.uap@empresa.com",       password: "123456", name: "Fernanda Rocha",   dept: "Planejamento UAP" },
  { email: "planejamento.central@empresa.com",   password: "123456", name: "Ricardo Alves",    dept: "Planejamento Central" },
  { email: "admin@empresa.com",                  password: "admin",  name: "Administrador",    dept: "Admin" },
];

// ─── STAGES ───────────────────────────────────────────────────────────────────
const STAGES = [
  { id: 1, label: "Análise do Bloqueio",     short: "Bloqueio",     dept: "Controle da Qualidade [Área técnica]" },
  { id: 2, label: "Definição do Recurso",    short: "Recurso",      dept: "Planejamento UAP" },
  { id: 3, label: "Criação de Ordem",        short: "Ordem",        dept: "Planejamento Central" },
  { id: 4, label: "Instrução da Qualidade",  short: "Instrução",    dept: "Controle da Qualidade [Área técnica]" },
  { id: 5, label: "Liberação para Vínculo",  short: "Lib. Vínculo", dept: "Controle da Qualidade [Liberação Intermediária]" },
  { id: 6, label: "Vínculo dos Lotes",       short: "Vínculo",      dept: "Planejamento UAP" },
  { id: 7, label: "Ativação de Flag",        short: "Flag",         dept: "Controle da Qualidade [Liberação Intermediária]" },
  { id: 8, label: "Concluído",              short: "Concluído",     dept: null },
];

const STAGE_COLORS = [
  "#007AFF","#34C759","#FF9500","#AF52DE",
  "#FF2D55","#5AC8FA","#FF6B35","#30B0C7"
];

// ─── TABLE COLUMNS ────────────────────────────────────────────────────────────
const CSV_COLUMNS = [
  "ultima_ordem","lote","ippn","qualidade_qts","deposito_sap",
  "motivo_bloqueio","motivo_bloqueio_texto","razao_bloq","num_cassete",
];
const COL_LABELS = {
  ultima_ordem:          "Última Ordem",
  lote:                  "Lote",
  ippn:                  "IPPN",
  qualidade_qts:         "Qualidade QTS",
  deposito_sap:          "Depósito SAP",
  motivo_bloqueio:       "Motivo Bloqueio",
  motivo_bloqueio_texto: "Motivo Bloqueio Texto",
  razao_bloq:            "Razão Bloq.",
  num_cassete:           "Nº Cassete",
};

// ─── HELPERS ──────────────────────────────────────────────────────────────────
function parseCSV(text) {
  const lines = text.trim().split(/\r?\n/);
  if (lines.length < 2) return [];
  const rawHeaders = lines[0].split(/[;,\t]/);
  const headers = rawHeaders.map(h =>
    h.trim().toLowerCase()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, "_")
      .replace(/-/g, "_")
      .replace(/[^a-z0-9_]/g, "")
  );
  return lines.slice(1).filter(l => l.trim()).map((line, i) => {
    const vals = line.split(/[;,\t]/);
    const obj = {
      _id: `row_${Date.now()}_${i}`,
      tratativa: "",
      checked: false,
      history: [],
    };
    CSV_COLUMNS.forEach(col => {
      const idx = headers.indexOf(col);
      obj[col] = idx >= 0 ? (vals[idx] || "").trim() : "";
    });
    return obj;
  });
}

function initStageData() {
  const d = {};
  STAGES.forEach(s => { d[s.id] = []; });
  return d;
}

// ─── STYLES ───────────────────────────────────────────────────────────────────
const S = {
  app: {
    fontFamily: "-apple-system, BlinkMacSystemFont, 'SF Pro Display', 'Segoe UI', sans-serif",
    minHeight: "100vh",
    background: "#F2F2F7",
    color: "#1C1C1E",
  },

  // ── Header
  header: {
    background: "linear-gradient(135deg, #0A2240 0%, #1A3A5C 100%)",
    padding: "0 24px",
    height: 56,
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    boxShadow: "0 1px 0 rgba(255,255,255,0.08), 0 4px 16px rgba(0,0,0,0.35)",
    position: "sticky",
    top: 0,
    zIndex: 100,
  },
  headerLeft: { display: "flex", alignItems: "center", gap: 12 },
  headerIcon: {
    width: 30, height: 30, borderRadius: 8,
    background: "rgba(255,255,255,0.12)",
    display: "flex", alignItems: "center", justifyContent: "center",
    fontSize: 17,
  },
  headerTitle: { color: "#fff", fontSize: 17, fontWeight: 600, letterSpacing: "-0.3px" },
  headerUser: {
    display: "flex", alignItems: "center", gap: 8,
    color: "rgba(255,255,255,0.9)", fontSize: 14, cursor: "pointer",
  },
  avatar: {
    width: 30, height: 30, borderRadius: "50%",
    background: "linear-gradient(135deg, #007AFF, #5AC8FA)",
    display: "flex", alignItems: "center", justifyContent: "center",
    fontSize: 13, fontWeight: 700, color: "#fff", flexShrink: 0,
  },
  userName: { display: "flex", flexDirection: "column", alignItems: "flex-end" },
  userNameText: { fontWeight: 600, fontSize: 13, lineHeight: 1.2 },
  userDeptText: { fontSize: 10, color: "rgba(255,255,255,0.5)", lineHeight: 1.2 },

  // ── Login
  loginWrap: {
    minHeight: "100vh",
    display: "flex", alignItems: "center", justifyContent: "center",
    background: "linear-gradient(160deg, #0A2240 0%, #1A3A5C 55%, #0D3B6B 100%)",
  },
  loginCard: {
    background: "rgba(255,255,255,0.97)",
    borderRadius: 20, padding: "40px 36px", width: 360,
    boxShadow: "0 24px 60px rgba(0,0,0,0.4)",
  },
  loginLogoBox: {
    width: 60, height: 60, borderRadius: 16, margin: "0 auto 16px",
    background: "linear-gradient(135deg, #007AFF, #0051D4)",
    display: "flex", alignItems: "center", justifyContent: "center",
    fontSize: 28, boxShadow: "0 6px 20px rgba(0,122,255,0.35)",
  },
  loginTitle: {
    fontSize: 24, fontWeight: 800, color: "#0A2240",
    letterSpacing: "-0.5px", textAlign: "center", marginBottom: 4,
  },
  loginSub: { fontSize: 13, color: "#8E8E93", textAlign: "center", marginBottom: 28 },
  loginHint: {
    marginTop: 20, fontSize: 11, color: "#C7C7CC",
    textAlign: "center", lineHeight: 1.7,
    borderTop: "1px solid #F2F2F7", paddingTop: 14,
  },

  // ── Form
  fieldWrap: { marginBottom: 14 },
  label: { fontSize: 13, fontWeight: 600, color: "#3A3A3C", marginBottom: 6, display: "block" },
  input: {
    width: "100%", border: "1.5px solid #E5E5EA", borderRadius: 10,
    padding: "11px 14px", fontSize: 15, outline: "none",
    background: "#fff", transition: "border-color 0.15s, box-shadow 0.15s",
  },
  btnPrimary: {
    width: "100%", background: "linear-gradient(135deg, #007AFF, #0051D4)",
    color: "#fff", border: "none", borderRadius: 12, padding: "13px",
    fontSize: 16, fontWeight: 600, cursor: "pointer",
    letterSpacing: "-0.2px", marginTop: 6,
  },
  errorBox: {
    color: "#FF3B30", fontSize: 13, textAlign: "center",
    background: "#FFF2F0", borderRadius: 8, padding: "8px 12px", marginBottom: 12,
  },

  // ── Pipeline
  pipelineWrap: {
    background: "#fff",
    borderBottom: "1px solid #E5E5EA",
    padding: "14px 20px 0",
    overflowX: "auto",
  },
  pipelineInner: {
    display: "flex", alignItems: "flex-start",
    minWidth: 720, paddingBottom: 0,
  },
  stageItem: {
    display: "flex", flexDirection: "column", alignItems: "center",
    flex: 1, cursor: "pointer", paddingBottom: 12, position: "relative",
  },
  stageDot: (active, done) => ({
    width: 34, height: 34, borderRadius: "50%",
    display: "flex", alignItems: "center", justifyContent: "center",
    fontSize: 13, fontWeight: 700, marginBottom: 6,
    background: done ? "#34C759" : active ? "#007AFF" : "#E5E5EA",
    color: done || active ? "#fff" : "#8E8E93",
    boxShadow: active ? "0 0 0 5px rgba(0,122,255,0.15)" : "none",
    transition: "all 0.2s",
  }),
  stageLabel: (active, done) => ({
    fontSize: 10, fontWeight: active ? 700 : 500,
    color: active ? "#007AFF" : done ? "#34C759" : "#8E8E93",
    textAlign: "center", lineHeight: 1.3, maxWidth: 72,
  }),
  connector: (done) => ({
    flex: 1, height: 2, marginTop: 16,
    background: done ? "#34C759" : "#E5E5EA",
    transition: "background 0.3s",
    marginLeft: -1, marginRight: -1,
  }),
  stageBadge: (count) => ({
    position: "absolute", top: -5, right: "calc(50% - 24px)",
    background: count > 0 ? "#FF3B30" : "#E5E5EA",
    color: count > 0 ? "#fff" : "#8E8E93",
    borderRadius: 10, fontSize: 10, fontWeight: 700,
    padding: "1px 6px", minWidth: 18, textAlign: "center",
    lineHeight: "16px", height: 16,
  }),
  activeUnderline: {
    position: "absolute", bottom: 0, left: "50%",
    transform: "translateX(-50%)",
    width: 28, height: 3,
    background: "#007AFF", borderRadius: 2,
  },

  // ── Nav tabs
  tabsWrap: { padding: "14px 24px 0", maxWidth: 1200, margin: "0 auto" },
  tabsInner: {
    display: "flex", gap: 2, background: "#E5E5EA",
    borderRadius: 10, padding: 3, overflowX: "auto",
    alignSelf: "flex-start", flexWrap: "nowrap",
  },
  tab: (active) => ({
    padding: "7px 14px", borderRadius: 8, fontSize: 12,
    fontWeight: active ? 700 : 500,
    color: active ? "#1C1C1E" : "#8E8E93",
    background: active ? "#fff" : "transparent",
    border: "none", cursor: "pointer", whiteSpace: "nowrap",
    boxShadow: active ? "0 1px 4px rgba(0,0,0,0.12)" : "none",
    transition: "all 0.15s",
  }),
  tabBadge: (color) => ({
    marginLeft: 5, background: color, color: "#fff",
    borderRadius: 8, fontSize: 9, fontWeight: 700,
    padding: "1px 5px",
  }),

  // ── Content
  content: { padding: "20px 24px 40px", maxWidth: 1200, margin: "0 auto" },
  pageTitle: { fontSize: 22, fontWeight: 800, color: "#1C1C1E", letterSpacing: "-0.4px", marginBottom: 4 },
  pageSub: { fontSize: 13, color: "#8E8E93", marginBottom: 20 },

  // ── Dashboard
  kpiRow: { display: "flex", gap: 12, marginBottom: 24, flexWrap: "wrap" },
  kpiCard: (color) => ({
    background: "#fff", borderRadius: 14, padding: "16px 20px",
    boxShadow: "0 1px 4px rgba(0,0,0,0.07)",
    display: "flex", alignItems: "center", gap: 14,
    flex: 1, minWidth: 150,
    borderLeft: `4px solid ${color}`,
  }),
  kpiNum: (color) => ({ fontSize: 30, fontWeight: 900, color, letterSpacing: "-1px" }),
  kpiLabel: { fontSize: 12, color: "#8E8E93", fontWeight: 500, marginTop: 1 },
  dashGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(155px,1fr))", gap: 12 },
  dashCard: (color) => ({
    background: "#fff", borderRadius: 14, padding: "16px",
    boxShadow: "0 1px 4px rgba(0,0,0,0.07)",
    borderTop: `3px solid ${color}`,
    cursor: "pointer", transition: "transform 0.15s, box-shadow 0.15s",
  }),
  dashNum: { fontSize: 34, fontWeight: 900, letterSpacing: "-1px", color: "#1C1C1E" },
  dashEtapa: { fontSize: 10, fontWeight: 700, marginTop: 2 },
  dashStageName: { fontSize: 11, color: "#8E8E93", marginTop: 2, lineHeight: 1.3 },
  dashLink: { fontSize: 10, color: "#007AFF", fontWeight: 700, marginTop: 8 },

  // ── Import / Dropzone
  dropzone: (drag) => ({
    border: `2px dashed ${drag ? "#007AFF" : "#C7C7CC"}`,
    borderRadius: 16, padding: "44px 24px", textAlign: "center",
    background: drag ? "rgba(0,122,255,0.04)" : "#F9F9FB",
    cursor: "pointer", transition: "all 0.2s", marginBottom: 20,
  }),
  dropIcon: { fontSize: 40, marginBottom: 10 },
  dropTitle: { fontSize: 15, fontWeight: 700, color: "#3A3A3C", marginBottom: 4 },
  dropSub: { fontSize: 12, color: "#8E8E93" },

  // ── Table
  tableWrap: {
    overflowX: "auto", borderRadius: 14,
    boxShadow: "0 1px 6px rgba(0,0,0,0.08)", background: "#fff",
  },
  table: { width: "100%", borderCollapse: "collapse", fontSize: 13 },
  th: {
    background: "#F2F2F7", padding: "10px 12px",
    textAlign: "left", fontWeight: 700, color: "#3A3A3C",
    borderBottom: "1px solid #E5E5EA", whiteSpace: "nowrap", fontSize: 11,
  },
  thTratativa: {
    background: "#EDF7EE", padding: "10px 12px",
    textAlign: "left", fontWeight: 700, color: "#1A7A3A",
    borderBottom: "1px solid #E5E5EA", whiteSpace: "nowrap", fontSize: 11,
  },
  td: {
    padding: "10px 12px", borderBottom: "1px solid #F2F2F7",
    verticalAlign: "middle", color: "#1C1C1E",
  },
  tdAlt: {
    padding: "10px 12px", borderBottom: "1px solid #F2F2F7",
    verticalAlign: "middle", color: "#1C1C1E",
    background: "#FAFAFA",
  },
  checkbox: { width: 18, height: 18, accentColor: "#007AFF", cursor: "pointer" },
  tratativaInput: {
    border: "1.5px solid #E5E5EA", borderRadius: 8, padding: "6px 10px",
    fontSize: 12, width: 180, outline: "none", background: "#fff",
    transition: "border-color 0.15s, box-shadow 0.15s",
  },

  // ── Buttons
  btnRow: { display: "flex", gap: 10, alignItems: "center", marginTop: 20, flexWrap: "wrap" },
  btnPrimaryAct: {
    background: "linear-gradient(135deg, #007AFF, #0051D4)",
    color: "#fff", border: "none", borderRadius: 12,
    padding: "11px 22px", fontSize: 14, fontWeight: 600,
    cursor: "pointer", display: "flex", alignItems: "center", gap: 8,
    transition: "opacity 0.15s",
  },
  btnSecondary: {
    background: "#F2F2F7", color: "#007AFF", border: "none",
    borderRadius: 12, padding: "11px 22px", fontSize: 14,
    fontWeight: 600, cursor: "pointer",
  },
  btnDanger: {
    background: "linear-gradient(135deg, #FF3B30, #C0392B)",
    color: "#fff", border: "none", borderRadius: 12,
    padding: "11px 22px", fontSize: 14, fontWeight: 600, cursor: "pointer",
  },

  // ── Dept tag
  deptTag: (dept) => {
    const map = {
      "Controle da Qualidade [Área técnica]":           { bg: "#E8F4FD", fg: "#1A6FA8" },
      "Controle da Qualidade [Liberação Intermediária]":{ bg: "#EDF7EE", fg: "#1A7A3A" },
      "Planejamento UAP":                               { bg: "#FFF3E0", fg: "#B45309" },
      "Planejamento Central":                           { bg: "#F3E8FF", fg: "#7C3AED" },
    };
    const c = map[dept] || { bg: "#F2F2F7", fg: "#3A3A3C" };
    return {
      background: c.bg, color: c.fg, borderRadius: 6,
      padding: "2px 9px", fontSize: 11, fontWeight: 600,
      display: "inline-block",
    };
  },

  // ── Empty state
  emptyWrap: { textAlign: "center", padding: "56px 24px", color: "#8E8E93" },
  emptyIcon: { fontSize: 44, marginBottom: 12 },
  emptyTitle: { fontSize: 16, fontWeight: 700, color: "#3A3A3C", marginBottom: 4 },
  emptySub: { fontSize: 13, color: "#8E8E93" },

  // ── Modal
  overlay: {
    position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)",
    zIndex: 200, display: "flex", alignItems: "center",
    justifyContent: "center", padding: 20,
  },
  modal: {
    background: "#fff", borderRadius: 20, padding: "28px 28px 24px",
    maxWidth: 420, width: "100%",
    boxShadow: "0 28px 70px rgba(0,0,0,0.3)",
  },
  modalTitle: { fontSize: 18, fontWeight: 800, color: "#1C1C1E", marginBottom: 8 },
  modalBody: { fontSize: 14, color: "#3A3A3C", marginBottom: 22, lineHeight: 1.65 },
  modalActions: { display: "flex", gap: 10, justifyContent: "flex-end" },

  // ── Toast
  toast: {
    position: "fixed", bottom: 28, left: "50%",
    transform: "translateX(-50%)",
    background: "#1C1C1E", color: "#fff",
    borderRadius: 12, padding: "12px 22px",
    fontSize: 14, fontWeight: 500, zIndex: 300,
    boxShadow: "0 8px 28px rgba(0,0,0,0.28)",
    display: "flex", alignItems: "center", gap: 8,
    whiteSpace: "nowrap",
  },

  // ── History
  historyItem: {
    fontSize: 11, color: "#3A3A3C", marginBottom: 6,
    borderLeft: "2px solid #E5E5EA", paddingLeft: 8, lineHeight: 1.6,
  },
};

// ─── COMPONENTS ───────────────────────────────────────────────────────────────

function Toast({ msg }) {
  if (!msg) return null;
  return (
    <div style={S.toast}>
      <span style={{ color: "#34C759", fontWeight: 800 }}>✓</span>
      {msg}
    </div>
  );
}

function Modal({ title, body, onConfirm, onCancel, confirmLabel = "Confirmar", danger = false }) {
  return (
    <div style={S.overlay} onClick={onCancel}>
      <div style={S.modal} onClick={e => e.stopPropagation()}>
        <div style={S.modalTitle}>{title}</div>
        <div style={S.modalBody}>{body}</div>
        <div style={S.modalActions}>
          <button style={S.btnSecondary} onClick={onCancel}>Cancelar</button>
          <button style={danger ? S.btnDanger : S.btnPrimaryAct} onClick={onConfirm}>
            {confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Login ─────────────────────────────────────────────────────────────────────
function LoginPage({ onLogin }) {
  const [email, setEmail]       = useState("");
  const [password, setPassword] = useState("");
  const [error, setError]       = useState("");
  const [loading, setLoading]   = useState(false);

  function handleLogin() {
    if (!email || !password) { setError("Preencha e-mail e senha."); return; }
    setLoading(true);
    setTimeout(() => {
      const user = USERS.find(u => u.email === email.trim().toLowerCase() && u.password === password);
      if (user) { setError(""); onLogin(user); }
      else setError("E-mail ou senha inválidos.");
      setLoading(false);
    }, 350);
  }

  return (
    <div style={S.loginWrap}>
      <div style={S.loginCard}>
        <div style={{ textAlign: "center" }}>
          <div style={S.loginLogoBox}>⚗️</div>
          <div style={S.loginTitle}>Fluxo de Liberação</div>
          <div style={S.loginSub}>Gestão de Produtos — Controle da Qualidade</div>
        </div>

        {error && <div style={S.errorBox}>{error}</div>}

        <div style={S.fieldWrap}>
          <label style={S.label}>E-mail</label>
          <input
            style={S.input}
            type="email"
            placeholder="seu@empresa.com"
            value={email}
            onChange={e => setEmail(e.target.value)}
            onKeyDown={e => e.key === "Enter" && handleLogin()}
            autoFocus
            autoComplete="username"
          />
        </div>

        <div style={S.fieldWrap}>
          <label style={S.label}>Senha</label>
          <input
            style={S.input}
            type="password"
            placeholder="••••••••"
            value={password}
            onChange={e => setPassword(e.target.value)}
            onKeyDown={e => e.key === "Enter" && handleLogin()}
            autoComplete="current-password"
          />
        </div>

        <button
          style={{ ...S.btnPrimary, opacity: loading ? 0.7 : 1 }}
          onClick={handleLogin}
          disabled={loading}
        >
          {loading ? "Entrando…" : "Entrar"}
        </button>

        <div style={S.loginHint}>
          Contas demo:<br />
          qualidade.tecnica@empresa.com / 123456<br />
          admin@empresa.com / admin
        </div>
      </div>
    </div>
  );
}

// ── Pipeline bar ──────────────────────────────────────────────────────────────
function Pipeline({ stageData, activeStage, onSelectStage }) {
  return (
    <div style={S.pipelineWrap}>
      <div style={S.pipelineInner}>
        {STAGES.map((stage, idx) => {
          const done   = activeStage > stage.id;
          const active = activeStage === stage.id;
          const count  = stageData[stage.id]?.length || 0;
          return (
            <div key={stage.id} style={{ display: "flex", alignItems: "flex-start", flex: 1 }}>
              <div style={S.stageItem} onClick={() => onSelectStage(stage.id)}>
                {count > 0 && stage.id !== 8 && (
                  <div style={S.stageBadge(count)}>{count}</div>
                )}
                <div style={S.stageDot(active, done)}>
                  {done ? "✓" : stage.id}
                </div>
                <div style={S.stageLabel(active, done)}>{stage.short}</div>
                {active && <div style={S.activeUnderline} />}
              </div>
              {idx < STAGES.length - 1 && (
                <div style={{ ...S.connector(done), marginTop: 17 }} />
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ── Import step (Stage 1 empty state) ────────────────────────────────────────
function ImportStep({ onImport }) {
  const [drag, setDrag]       = useState(false);
  const [preview, setPreview] = useState(null);
  const [error, setError]     = useState("");

  function processFile(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const rows = parseCSV(e.target.result);
        if (rows.length === 0) { setError("Nenhuma linha válida encontrada no arquivo."); return; }
        setPreview(rows);
        setError("");
      } catch { setError("Erro ao processar arquivo. Verifique o formato."); }
    };
    reader.readAsText(file, "utf-8");
  }

  const handleDrop = useCallback(e => {
    e.preventDefault();
    setDrag(false);
    processFile(e.dataTransfer.files[0]);
  }, []);

  return (
    <div>
      <div style={S.pageTitle}>Importar Tubos Bloqueados</div>
      <div style={S.pageSub}>Carregue um arquivo CSV ou TXT para iniciar o fluxo de liberação na Etapa 1.</div>

      <div
        style={S.dropzone(drag)}
        onDragOver={e => { e.preventDefault(); setDrag(true); }}
        onDragLeave={() => setDrag(false)}
        onDrop={handleDrop}
        onClick={() => document.getElementById("fileInput").click()}
      >
        <div style={S.dropIcon}>📂</div>
        <div style={S.dropTitle}>Arraste o arquivo aqui ou clique para selecionar</div>
        <div style={S.dropSub}>CSV ou TXT — separadores aceitos: ponto-e-vírgula, vírgula ou tab</div>
        <input
          id="fileInput" type="file" accept=".csv,.txt"
          style={{ display: "none" }}
          onChange={e => processFile(e.target.files[0])}
        />
      </div>

      {error && <div style={{ ...S.errorBox, marginBottom: 16 }}>{error}</div>}

      {preview ? (
        <div>
          <div style={{ fontSize: 13, color: "#34C759", fontWeight: 700, marginBottom: 12 }}>
            ✓ {preview.length} tubo{preview.length !== 1 ? "s" : ""} encontrado{preview.length !== 1 ? "s" : ""} — confira abaixo
          </div>
          <div style={{ ...S.tableWrap, marginBottom: 16, maxHeight: 300, overflowY: "auto" }}>
            <table style={S.table}>
              <thead>
                <tr>{CSV_COLUMNS.map(c => <th key={c} style={S.th}>{COL_LABELS[c]}</th>)}</tr>
              </thead>
              <tbody>
                {preview.slice(0, 12).map((row, i) => (
                  <tr key={i}>
                    {CSV_COLUMNS.map(c => (
                      <td key={c} style={i % 2 === 0 ? S.td : S.tdAlt}>
                        {row[c] || <span style={{ color: "#C7C7CC" }}>—</span>}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {preview.length > 12 && (
            <div style={{ fontSize: 12, color: "#8E8E93", marginBottom: 12 }}>
              Exibindo 12 de {preview.length} linhas
            </div>
          )}
          <div style={S.btnRow}>
            <button style={S.btnPrimaryAct} onClick={() => onImport(preview)}>
              <span>📥</span> Importar {preview.length} Tubos para Etapa 1
            </button>
            <button style={S.btnSecondary} onClick={() => setPreview(null)}>Cancelar</button>
          </div>
        </div>
      ) : (
        <div style={{ background: "#fff", borderRadius: 14, padding: "16px 18px", boxShadow: "0 1px 4px rgba(0,0,0,0.07)" }}>
          <div style={{ fontSize: 12, fontWeight: 700, color: "#3A3A3C", marginBottom: 8 }}>
            Colunas esperadas no arquivo:
          </div>
          <div style={{ fontSize: 12, color: "#8E8E93", lineHeight: 2 }}>
            {Object.values(COL_LABELS).join("  ·  ")}
          </div>
        </div>
      )}
    </div>
  );
}

// ── Stage view (table + advance) ─────────────────────────────────────────────
function StageView({ stage, rows, user, onAdvance }) {
  const [selected, setSelected]     = useState(new Set());
  const [tratativas, setTratativas] = useState({});
  const [showConfirm, setShowConfirm] = useState(false);

  const isLast    = stage.id === 7;
  const nextStage = STAGES.find(s => s.id === stage.id + 1);

  function toggleAll(checked) {
    setSelected(checked ? new Set(rows.map(r => r._id)) : new Set());
  }
  function toggleRow(id) {
    const next = new Set(selected);
    next.has(id) ? next.delete(id) : next.add(id);
    setSelected(next);
  }

  function handleAdvance() {
    const toMove = rows
      .filter(r => selected.has(r._id))
      .map(r => ({
        ...r,
        tratativa: tratativas[r._id] ?? r.tratativa,
        history: [
          ...(r.history || []),
          {
            stage:      stage.id,
            stageLabel: stage.label,
            user:       user.name,
            dept:       user.dept,
            tratativa:  tratativas[r._id] ?? r.tratativa,
            date:       new Date().toLocaleString("pt-BR"),
          },
        ],
      }));
    const remaining = rows.filter(r => !selected.has(r._id));
    onAdvance(toMove, remaining);
    setSelected(new Set());
    setTratativas({});
    setShowConfirm(false);
  }

  if (rows.length === 0) {
    return (
      <div>
        <div style={S.pageTitle}>{stage.label}</div>
        <div style={{ marginBottom: 16 }}>
          <span style={S.deptTag(stage.dept)}>{stage.dept}</span>
        </div>
        <div style={{ ...S.emptyWrap, background: "#fff", borderRadius: 16 }}>
          <div style={S.emptyIcon}>📭</div>
          <div style={S.emptyTitle}>Nenhum tubo nesta etapa</div>
          <div style={S.emptySub}>Os tubos aparecerão aqui quando forem enviados para esta etapa.</div>
        </div>
      </div>
    );
  }

  return (
    <div>
      <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: 14, flexWrap: "wrap", gap: 10 }}>
        <div>
          <div style={S.pageTitle}>{stage.label}</div>
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginTop: 6 }}>
            <span style={S.deptTag(stage.dept)}>{stage.dept}</span>
            <span style={{ fontSize: 12, color: "#8E8E93" }}>
              {rows.length} tubo{rows.length !== 1 ? "s" : ""} pendente{rows.length !== 1 ? "s" : ""}
            </span>
          </div>
        </div>
        {selected.size > 0 && (
          <div style={{ fontSize: 13, fontWeight: 700, color: "#007AFF", background: "rgba(0,122,255,0.08)", borderRadius: 8, padding: "6px 14px" }}>
            {selected.size} selecionado{selected.size !== 1 ? "s" : ""}
          </div>
        )}
      </div>

      <div style={S.tableWrap}>
        <table style={S.table}>
          <thead>
            <tr>
              <th style={{ ...S.th, width: 38 }}>
                <input
                  type="checkbox" style={S.checkbox}
                  checked={selected.size === rows.length && rows.length > 0}
                  onChange={e => toggleAll(e.target.checked)}
                />
              </th>
              {CSV_COLUMNS.map(c => <th key={c} style={S.th}>{COL_LABELS[c]}</th>)}
              <th style={S.thTratativa}>Tratativa</th>
              {stage.id > 1 && <th style={S.th}>Histórico</th>}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, i) => {
              const tdStyle = i % 2 === 0 ? S.td : S.tdAlt;
              return (
                <tr key={row._id}>
                  <td style={tdStyle}>
                    <input
                      type="checkbox" style={S.checkbox}
                      checked={selected.has(row._id)}
                      onChange={() => toggleRow(row._id)}
                    />
                  </td>
                  {CSV_COLUMNS.map(c => (
                    <td key={c} style={tdStyle}>
                      {row[c] || <span style={{ color: "#C7C7CC" }}>—</span>}
                    </td>
                  ))}
                  <td style={tdStyle}>
                    <input
                      type="text"
                      style={S.tratativaInput}
                      placeholder="Registrar tratativa…"
                      value={tratativas[row._id] ?? row.tratativa ?? ""}
                      onChange={e => setTratativas(t => ({ ...t, [row._id]: e.target.value }))}
                    />
                  </td>
                  {stage.id > 1 && (
                    <td style={tdStyle}>
                      {row.history?.length > 0 ? (
                        <details>
                          <summary style={{ fontSize: 11, color: "#007AFF", cursor: "pointer", fontWeight: 600 }}>
                            {row.history.length} etapa{row.history.length !== 1 ? "s" : ""}
                          </summary>
                          <div style={{ marginTop: 6 }}>
                            {row.history.map((h, hi) => (
                              <div key={hi} style={S.historyItem}>
                                <strong>{h.stageLabel}</strong><br />
                                {h.user}
                                {h.dept && <span style={{ color: "#8E8E93" }}> · {h.dept}</span>}<br />
                                {h.tratativa && <em style={{ color: "#555" }}>"{h.tratativa}"</em>}
                                {h.tratativa && <br />}
                                <span style={{ color: "#C7C7CC" }}>{h.date}</span>
                              </div>
                            ))}
                          </div>
                        </details>
                      ) : (
                        <span style={{ color: "#C7C7CC", fontSize: 12 }}>—</span>
                      )}
                    </td>
                  )}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      <div style={S.btnRow}>
        <button
          style={{ ...S.btnPrimaryAct, opacity: selected.size === 0 ? 0.4 : 1 }}
          disabled={selected.size === 0}
          onClick={() => setShowConfirm(true)}
        >
          <span>{isLast ? "✅" : "→"}</span>
          {isLast
            ? `Finalizar ${selected.size} tubo${selected.size !== 1 ? "s" : ""}`
            : `Enviar para: ${nextStage?.short} (${selected.size})`
          }
        </button>
        {selected.size === 0 && (
          <span style={{ fontSize: 12, color: "#8E8E93" }}>
            Selecione ao menos um tubo para avançar
          </span>
        )}
      </div>

      {showConfirm && (
        <Modal
          title={isLast ? "Finalizar tubos?" : `Enviar para "${nextStage?.label}"?`}
          body={`${selected.size} tubo(s) serão ${isLast ? "marcados como concluídos" : `enviados para "${nextStage?.label}"`}. Esta ação será registrada em nome de ${user.name}.`}
          onConfirm={handleAdvance}
          onCancel={() => setShowConfirm(false)}
          confirmLabel={isLast ? "Finalizar" : "Confirmar envio"}
        />
      )}
    </div>
  );
}

// ── Dashboard ─────────────────────────────────────────────────────────────────
function Dashboard({ stageData, onSelectStage }) {
  const allRows   = Object.values(stageData).flat();
  const total     = allRows.length;
  const completed = stageData[8]?.length || 0;
  const inFlow    = total - completed;

  return (
    <div>
      <div style={S.pageTitle}>Dashboard</div>
      <div style={S.pageSub}>Visão geral do fluxo de liberação</div>

      <div style={S.kpiRow}>
        {[
          { label: "Em fluxo",      value: inFlow,    color: "#007AFF", icon: "🔄" },
          { label: "Concluídos",    value: completed,  color: "#34C759", icon: "✅" },
          { label: "Total no sistema", value: total,  color: "#FF9500", icon: "📦" },
        ].map(k => (
          <div key={k.label} style={S.kpiCard(k.color)}>
            <span style={{ fontSize: 30 }}>{k.icon}</span>
            <div>
              <div style={S.kpiNum(k.color)}>{k.value}</div>
              <div style={S.kpiLabel}>{k.label}</div>
            </div>
          </div>
        ))}
      </div>

      <div style={S.dashGrid}>
        {STAGES.filter(s => s.id !== 8).map((stage, idx) => {
          const count = stageData[stage.id]?.length || 0;
          const color = STAGE_COLORS[idx];
          return (
            <div
              key={stage.id}
              style={S.dashCard(color)}
              onClick={() => onSelectStage(stage.id)}
              onMouseEnter={e => {
                e.currentTarget.style.transform = "translateY(-3px)";
                e.currentTarget.style.boxShadow = "0 6px 18px rgba(0,0,0,0.13)";
              }}
              onMouseLeave={e => {
                e.currentTarget.style.transform = "";
                e.currentTarget.style.boxShadow = "0 1px 4px rgba(0,0,0,0.07)";
              }}
            >
              <div style={S.dashNum}>{count}</div>
              <div style={{ ...S.dashEtapa, color }}>{`ETAPA ${stage.id}`}</div>
              <div style={S.dashStageName}>{stage.short}</div>
              {count > 0 && <div style={S.dashLink}>Ver tubos →</div>}
            </div>
          );
        })}
      </div>

      {total === 0 && (
        <div style={{ ...S.emptyWrap, background: "#fff", borderRadius: 16, marginTop: 20 }}>
          <div style={S.emptyIcon}>📋</div>
          <div style={S.emptyTitle}>Nenhum tubo no sistema</div>
          <div style={S.emptySub}>Acesse a Etapa 1 para importar o arquivo de tubos bloqueados.</div>
        </div>
      )}
    </div>
  );
}

// ── Completed view ────────────────────────────────────────────────────────────
function CompletedView({ rows }) {
  if (rows.length === 0) {
    return (
      <div>
        <div style={S.pageTitle}>Tubos Concluídos</div>
        <div style={{ ...S.emptyWrap, background: "#fff", borderRadius: 16, marginTop: 12 }}>
          <div style={S.emptyIcon}>🏁</div>
          <div style={S.emptyTitle}>Nenhum tubo concluído ainda</div>
          <div style={S.emptySub}>Os tubos aparecerão aqui após percorrerem todas as etapas do fluxo.</div>
        </div>
      </div>
    );
  }

  return (
    <div>
      <div style={S.pageTitle}>Tubos Concluídos</div>
      <div style={S.pageSub}>{rows.length} tubo{rows.length !== 1 ? "s" : ""} completaram o fluxo de liberação.</div>
      <div style={S.tableWrap}>
        <table style={S.table}>
          <thead>
            <tr>
              {CSV_COLUMNS.map(c => <th key={c} style={S.th}>{COL_LABELS[c]}</th>)}
              <th style={S.th}>Histórico Completo</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row, i) => {
              const tdStyle = i % 2 === 0 ? S.td : S.tdAlt;
              return (
                <tr key={row._id}>
                  {CSV_COLUMNS.map(c => (
                    <td key={c} style={tdStyle}>{row[c] || "—"}</td>
                  ))}
                  <td style={tdStyle}>
                    <details>
                      <summary style={{ fontSize: 11, color: "#34C759", cursor: "pointer", fontWeight: 700 }}>
                        ✓ {row.history?.length || 0} etapa{(row.history?.length || 0) !== 1 ? "s" : ""}
                      </summary>
                      <div style={{ marginTop: 6 }}>
                        {row.history?.map((h, hi) => (
                          <div key={hi} style={{ ...S.historyItem, borderLeftColor: "#34C759" }}>
                            <strong>{h.stageLabel}</strong><br />
                            {h.user}
                            {h.dept && <span style={{ color: "#8E8E93" }}> · {h.dept}</span>}<br />
                            {h.tratativa && <><em style={{ color: "#555" }}>"{h.tratativa}"</em><br /></>}
                            <span style={{ color: "#C7C7CC" }}>{h.date}</span>
                          </div>
                        ))}
                      </div>
                    </details>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ─── MAIN APP ─────────────────────────────────────────────────────────────────
export default function App() {
  const [user, setUser]           = useState(null);
  const [stageData, setStageData] = useState(initStageData);
  const [activeStage, setActive]  = useState(0);   // 0 = dashboard
  const [toast, setToast]         = useState("");
  const [showLogout, setShowLogout] = useState(false);

  function showToast(msg) {
    setToast(msg);
    setTimeout(() => setToast(""), 3200);
  }

  function handleImport(rows) {
    setStageData(d => ({ ...d, 1: [...d[1], ...rows] }));
    setActive(1);
    showToast(`${rows.length} tubos importados para a Etapa 1`);
  }

  function handleAdvance(toMove, remaining) {
    const nextId    = activeStage + 1;
    const nextStage = STAGES.find(s => s.id === nextId);
    setStageData(d => ({
      ...d,
      [activeStage]: remaining,
      [nextId]: [...(d[nextId] || []), ...toMove],
    }));
    showToast(`${toMove.length} tubo(s) enviado(s) para "${nextStage?.label}"`);
  }

  if (!user) return <LoginPage onLogin={setUser} />;

  const currentStage = STAGES.find(s => s.id === activeStage);

  return (
    <div style={S.app}>
      {/* ── Header */}
      <header style={S.header}>
        <div style={S.headerLeft}>
          <div style={S.headerIcon}>⚗️</div>
          <span style={S.headerTitle}>Fluxo de Liberação</span>
        </div>
        <div style={S.headerUser} onClick={() => setShowLogout(true)} title="Clique para sair">
          <div style={S.userName}>
            <span style={S.userNameText}>{user.name}</span>
            <span style={S.userDeptText}>{user.dept}</span>
          </div>
          <div style={S.avatar}>{user.name.charAt(0).toUpperCase()}</div>
        </div>
      </header>

      {/* ── Pipeline */}
      <Pipeline stageData={stageData} activeStage={activeStage} onSelectStage={setActive} />

      {/* ── Nav tabs */}
      <div style={S.tabsWrap}>
        <div style={S.tabsInner}>
          <button style={S.tab(activeStage === 0)} onClick={() => setActive(0)}>
            Dashboard
          </button>
          {STAGES.filter(s => s.id !== 8).map(s => {
            const cnt = stageData[s.id]?.length || 0;
            return (
              <button key={s.id} style={S.tab(activeStage === s.id)} onClick={() => setActive(s.id)}>
                E{s.id}: {s.short}
                {cnt > 0 && <span style={S.tabBadge("#FF3B30")}>{cnt}</span>}
              </button>
            );
          })}
          <button style={S.tab(activeStage === 8)} onClick={() => setActive(8)}>
            Concluídos
            {(stageData[8]?.length || 0) > 0 && (
              <span style={S.tabBadge("#34C759")}>{stageData[8].length}</span>
            )}
          </button>
        </div>
      </div>

      {/* ── Page content */}
      <div style={S.content}>
        {activeStage === 0 && (
          <Dashboard stageData={stageData} onSelectStage={setActive} />
        )}

        {activeStage === 1 && (
          stageData[1].length === 0 ? (
            <ImportStep onImport={handleImport} />
          ) : (
            <div>
              <StageView
                stage={STAGES[0]}
                rows={stageData[1]}
                user={user}
                onAdvance={handleAdvance}
              />
              <div style={{ marginTop: 16 }}>
                <button
                  style={{ ...S.btnSecondary, fontSize: 12 }}
                  onClick={() => {
                    if (window.confirm("Isso removerá todos os tubos da Etapa 1. Continuar?")) {
                      setStageData(d => ({ ...d, 1: [] }));
                    }
                  }}
                >
                  ↩ Importar novo arquivo
                </button>
              </div>
            </div>
          )
        )}

        {activeStage > 1 && activeStage <= 7 && (
          <StageView
            stage={STAGES[activeStage - 1]}
            rows={stageData[activeStage]}
            user={user}
            onAdvance={handleAdvance}
          />
        )}

        {activeStage === 8 && (
          <CompletedView rows={stageData[8]} />
        )}
      </div>

      {/* ── Toast */}
      <Toast msg={toast} />

      {/* ── Logout confirm */}
      {showLogout && (
        <Modal
          title="Sair da conta"
          body={`Deseja encerrar a sessão de ${user.name}?`}
          onConfirm={() => { setUser(null); setShowLogout(false); setStageData(initStageData()); setActive(0); }}
          onCancel={() => setShowLogout(false)}
          confirmLabel="Sair"
          danger
        />
      )}
    </div>
  );
}
