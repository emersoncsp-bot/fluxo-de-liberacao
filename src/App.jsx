import { useState, useCallback, useRef } from "react";

// ─── USERS ────────────────────────────────────────────────────────────────────
const USERS = [
  { email: "qualidade.tecnica@empresa.com",    password: "123456", name: "Ana Paula Silva",  dept: "Controle da Qualidade [Área técnica]" },
  { email: "qualidade.lib@empresa.com",        password: "123456", name: "Carlos Mendes",    dept: "Controle da Qualidade [Liberação Intermediária]" },
  { email: "planejamento.uap@empresa.com",     password: "123456", name: "Fernanda Rocha",   dept: "Planejamento UAP" },
  { email: "planejamento.central@empresa.com", password: "123456", name: "Ricardo Alves",    dept: "Planejamento Central" },
  { email: "admin@empresa.com",                password: "admin",  name: "Administrador",    dept: "Admin" },
];

// ─── STAGES ───────────────────────────────────────────────────────────────────
const STAGES = [
  { id: 1, label: "Análise do Bloqueio",    short: "Análise de bloqueio",  dept: "CQ Área Técnica" },
  { id: 2, label: "Definição do Recurso",   short: "Análise de recurso",   dept: "Planejamento UAP" },
  { id: 3, label: "Criação de Ordem",       short: "Definição de ordem",   dept: "Planejamento Central" },
  { id: 4, label: "Instrução da Qualidade", short: "Instrução",            dept: "CQ Área Técnica" },
  { id: 5, label: "Liberação para Vínculo", short: "Lib. Vínculo",         dept: "CQ Lib. Intermediária" },
  { id: 6, label: "Vínculo dos Lotes",      short: "Vínculo",              dept: "Planejamento UAP" },
  { id: 7, label: "Ativação de Flag",       short: "Flag",                 dept: "CQ Lib. Intermediária" },
  { id: 8, label: "Pendente Execução",      short: "Pendente Execução",    dept: null },
];

const STAGE_COLORS = ["#007AFF","#34C759","#FF9500","#AF52DE","#FF2D55","#5AC8FA","#FF6B35","#8E8E93"];

const DEPT_FULL = {
  "CQ Área Técnica":       "Controle da Qualidade [Área técnica]",
  "Planejamento UAP":      "Planejamento UAP",
  "Planejamento Central":  "Planejamento Central",
  "CQ Lib. Intermediária": "Controle da Qualidade [Liberação Intermediária]",
};

const CSV_COLUMNS = [
  "ultima_ordem","lote","ippn","qualidade_qts","deposito_sap",
  "motivo_bloqueio","motivo_bloqueio_texto","razao_bloq","num_cassete",
];
const COL_LABELS = {
  ultima_ordem: "Última Ordem", lote: "Lote", ippn: "IPPN",
  qualidade_qts: "Qualidade QTS", deposito_sap: "Depósito SAP",
  motivo_bloqueio: "Motivo Bloqueio", motivo_bloqueio_texto: "Motivo Bloqueio Texto",
  razao_bloq: "Razão Bloq.", num_cassete: "Nº Cassete",
};

// ─── HELPERS ──────────────────────────────────────────────────────────────────
function parseCSV(text) {
  const lines = text.trim().split(/\r?\n/);
  if (lines.length < 2) return [];
  const headers = lines[0].split(/[;,\t]/).map(h =>
    h.trim().toLowerCase()
      .normalize("NFD").replace(/[\u0300-\u036f]/g,"")
      .replace(/[\s-]+/g,"_").replace(/[^a-z0-9_]/g,"")
  );
  return lines.slice(1).filter(l => l.trim()).map((line, i) => {
    const vals = line.split(/[;,\t]/);
    const obj = { _id: `row_${Date.now()}_${i}`, tratativa: "", history: [] };
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

// Group rows by lote, each group gets a tratativa
function groupByLote(rows) {
  const map = new Map();
  rows.forEach(row => {
    const key = row.lote || row._id;
    if (!map.has(key)) map.set(key, { lote: key, rows: [], tratativa: row.tratativa || "" });
    map.get(key).rows.push(row);
  });
  return Array.from(map.values());
}

// ─── DEPT TAG COLORS ─────────────────────────────────────────────────────────
const DEPT_COLORS = {
  "CQ Área Técnica":       { bg: "#E8F4FD", fg: "#1A6FA8" },
  "Planejamento UAP":      { bg: "#FFF3E0", fg: "#B45309" },
  "Planejamento Central":  { bg: "#F3E8FF", fg: "#7C3AED" },
  "CQ Lib. Intermediária": { bg: "#EDF7EE", fg: "#1A7A3A" },
};

function DeptTag({ dept }) {
  const c = DEPT_COLORS[dept] || { bg: "#F2F2F7", fg: "#3A3A3C" };
  return (
    <span style={{ background: c.bg, color: c.fg, borderRadius: 6, padding: "2px 9px", fontSize: 11, fontWeight: 600 }}>
      {DEPT_FULL[dept] || dept}
    </span>
  );
}

// ─── TOAST ────────────────────────────────────────────────────────────────────
function Toast({ msg }) {
  if (!msg) return null;
  return (
    <div style={{ position:"fixed", bottom:28, left:"50%", transform:"translateX(-50%)", background:"#1C1C1E", color:"#fff", borderRadius:12, padding:"12px 22px", fontSize:14, fontWeight:500, zIndex:300, boxShadow:"0 8px 28px rgba(0,0,0,0.28)", display:"flex", alignItems:"center", gap:8, whiteSpace:"nowrap" }}>
      <span style={{ color:"#34C759", fontWeight:800 }}>✓</span>{msg}
    </div>
  );
}

// ─── MODAL ────────────────────────────────────────────────────────────────────
function Modal({ title, body, onConfirm, onCancel, confirmLabel = "Confirmar", danger = false }) {
  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.5)", zIndex:200, display:"flex", alignItems:"center", justifyContent:"center", padding:20 }} onClick={onCancel}>
      <div style={{ background:"#fff", borderRadius:20, padding:"28px 28px 24px", maxWidth:420, width:"100%", boxShadow:"0 28px 70px rgba(0,0,0,0.3)" }} onClick={e => e.stopPropagation()}>
        <div style={{ fontSize:18, fontWeight:800, color:"#1C1C1E", marginBottom:8 }}>{title}</div>
        <div style={{ fontSize:14, color:"#3A3A3C", marginBottom:22, lineHeight:1.65 }}>{body}</div>
        <div style={{ display:"flex", gap:10, justifyContent:"flex-end" }}>
          <button style={{ background:"#F2F2F7", color:"#007AFF", border:"none", borderRadius:12, padding:"11px 22px", fontSize:14, fontWeight:600, cursor:"pointer" }} onClick={onCancel}>Cancelar</button>
          <button style={{ background: danger ? "linear-gradient(135deg,#FF3B30,#C0392B)" : "linear-gradient(135deg,#007AFF,#0051D4)", color:"#fff", border:"none", borderRadius:12, padding:"11px 22px", fontSize:14, fontWeight:600, cursor:"pointer" }} onClick={onConfirm}>{confirmLabel}</button>
        </div>
      </div>
    </div>
  );
}

// ─── LOGIN ────────────────────────────────────────────────────────────────────
function LoginPage({ onLogin }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

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

  const inp = { width:"100%", border:"1.5px solid #E5E5EA", borderRadius:10, padding:"11px 14px", fontSize:15, outline:"none", background:"#fff", boxSizing:"border-box" };

  return (
    <div style={{ minHeight:"100vh", display:"flex", alignItems:"center", justifyContent:"center", background:"linear-gradient(160deg,#0A2240 0%,#1A3A5C 55%,#0D3B6B 100%)" }}>
      <div style={{ background:"rgba(255,255,255,0.97)", borderRadius:20, padding:"40px 36px", width:360, boxShadow:"0 24px 60px rgba(0,0,0,0.4)" }}>
        <div style={{ textAlign:"center", marginBottom:28 }}>
          <div style={{ width:60, height:60, borderRadius:16, margin:"0 auto 16px", background:"linear-gradient(135deg,#007AFF,#0051D4)", display:"flex", alignItems:"center", justifyContent:"center", fontSize:28, boxShadow:"0 6px 20px rgba(0,122,255,0.35)" }}>⚗️</div>
          <div style={{ fontSize:24, fontWeight:800, color:"#0A2240", letterSpacing:"-0.5px" }}>Fluxo de Liberação</div>
          <div style={{ fontSize:13, color:"#8E8E93", marginTop:4 }}>Gestão de Produtos — Controle da Qualidade</div>
        </div>
        {error && <div style={{ color:"#FF3B30", fontSize:13, textAlign:"center", background:"#FFF2F0", borderRadius:8, padding:"8px 12px", marginBottom:12 }}>{error}</div>}
        <div style={{ marginBottom:14 }}>
          <label style={{ fontSize:13, fontWeight:600, color:"#3A3A3C", marginBottom:6, display:"block" }}>E-mail</label>
          <input style={inp} type="email" placeholder="seu@empresa.com" value={email} onChange={e=>setEmail(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handleLogin()} autoFocus />
        </div>
        <div style={{ marginBottom:20 }}>
          <label style={{ fontSize:13, fontWeight:600, color:"#3A3A3C", marginBottom:6, display:"block" }}>Senha</label>
          <input style={inp} type="password" placeholder="••••••••" value={password} onChange={e=>setPassword(e.target.value)} onKeyDown={e=>e.key==="Enter"&&handleLogin()} />
        </div>
        <button style={{ width:"100%", background:"linear-gradient(135deg,#007AFF,#0051D4)", color:"#fff", border:"none", borderRadius:12, padding:"13px", fontSize:16, fontWeight:600, cursor:"pointer", opacity:loading?0.7:1 }} onClick={handleLogin} disabled={loading}>
          {loading ? "Entrando…" : "Entrar"}
        </button>
        <div style={{ marginTop:20, fontSize:11, color:"#C7C7CC", textAlign:"center", lineHeight:1.7, borderTop:"1px solid #F2F2F7", paddingTop:14 }}>
          Demo: qualidade.tecnica@empresa.com / 123456<br/>admin@empresa.com / admin
        </div>
      </div>
    </div>
  );
}

// ─── PIPELINE ─────────────────────────────────────────────────────────────────
function Pipeline({ stageData, activeStage, onSelectStage }) {
  return (
    <div style={{ background:"#fff", borderBottom:"1px solid #E5E5EA", overflowX:"auto" }}>
      <div style={{ display:"flex", alignItems:"stretch", minWidth:900, padding:"0 12px" }}>
        {STAGES.map((stage, idx) => {
          const active = activeStage === stage.id;
          const count  = stageData[stage.id]?.length || 0;
          return (
            <div key={stage.id} style={{ display:"flex", alignItems:"stretch", flex:1 }}>
              {/* Stage cell */}
              <div
                onClick={() => onSelectStage(stage.id)}
                style={{
                  flex:1, display:"flex", flexDirection:"column", alignItems:"center",
                  justifyContent:"flex-start", cursor:"pointer", padding:"14px 4px 10px",
                  position:"relative",
                  borderBottom: active ? "3px solid #007AFF" : "3px solid transparent",
                  background: active ? "rgba(0,122,255,0.04)" : "transparent",
                  transition:"background 0.15s",
                  minWidth: 90,
                }}
              >
                {/* Badge */}
                {count > 0 && stage.id !== 8 && (
                  <div style={{ position:"absolute", top:8, right:"calc(50% - 20px)", background:"#FF3B30", color:"#fff", borderRadius:10, fontSize:9, fontWeight:700, padding:"1px 5px", lineHeight:"14px", height:14, minWidth:16, textAlign:"center" }}>{count}</div>
                )}
                {/* Dot */}
                <div style={{
                  width:30, height:30, borderRadius:"50%", display:"flex", alignItems:"center", justifyContent:"center",
                  fontSize:12, fontWeight:700, marginBottom:6,
                  background: active ? "#007AFF" : "#E5E5EA",
                  color: active ? "#fff" : "#8E8E93",
                  boxShadow: active ? "0 0 0 4px rgba(0,122,255,0.15)" : "none",
                  transition:"all 0.2s", flexShrink:0,
                }}>
                  {stage.id}
                </div>
                {/* Stage name */}
                <div style={{ fontSize:10, fontWeight: active ? 700 : 500, color: active ? "#007AFF" : "#6B6B6B", textAlign:"center", lineHeight:1.3, marginBottom:4 }}>
                  {stage.short}
                </div>
                {/* Dept name */}
                {stage.dept && (
                  <div style={{ fontSize:9, color: active ? "rgba(0,122,255,0.7)" : "#ABABAB", textAlign:"center", lineHeight:1.3, fontWeight:500, maxWidth:88 }}>
                    {DEPT_FULL[stage.dept] || stage.dept}
                  </div>
                )}
              </div>
              {/* Connector */}
              {idx < STAGES.length - 1 && (
                <div style={{ width:8, display:"flex", alignItems:"center", marginBottom:3 }}>
                  <div style={{ width:"100%", height:2, background:"#E5E5EA" }} />
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

// ─── IMPORT STEP ──────────────────────────────────────────────────────────────
function ImportStep({ onImport }) {
  const [drag, setDrag] = useState(false);
  const [preview, setPreview] = useState(null);
  const [error, setError] = useState("");

  function processFile(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const rows = parseCSV(e.target.result);
        if (rows.length === 0) { setError("Nenhuma linha válida no arquivo."); return; }
        setPreview(rows); setError("");
      } catch { setError("Erro ao processar arquivo."); }
    };
    reader.readAsText(file, "utf-8");
  }

  const handleDrop = useCallback(e => { e.preventDefault(); setDrag(false); processFile(e.dataTransfer.files[0]); }, []);

  return (
    <div>
      <div style={{ fontSize:22, fontWeight:800, color:"#1C1C1E", letterSpacing:"-0.4px", marginBottom:4 }}>Importar Tubos Bloqueados</div>
      <div style={{ fontSize:13, color:"#8E8E93", marginBottom:20 }}>Carregue um arquivo CSV ou TXT para iniciar o fluxo na Etapa 1.</div>
      <div
        style={{ border:`2px dashed ${drag?"#007AFF":"#C7C7CC"}`, borderRadius:16, padding:"44px 24px", textAlign:"center", background:drag?"rgba(0,122,255,0.04)":"#F9F9FB", cursor:"pointer", transition:"all 0.2s", marginBottom:20 }}
        onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)} onDrop={handleDrop}
        onClick={()=>document.getElementById("fileInput").click()}
      >
        <div style={{ fontSize:40, marginBottom:10 }}>📂</div>
        <div style={{ fontSize:15, fontWeight:700, color:"#3A3A3C", marginBottom:4 }}>Arraste o arquivo ou clique para selecionar</div>
        <div style={{ fontSize:12, color:"#8E8E93" }}>CSV ou TXT — separadores: ponto-e-vírgula, vírgula ou tab</div>
        <input id="fileInput" type="file" accept=".csv,.txt" style={{ display:"none" }} onChange={e=>processFile(e.target.files[0])} />
      </div>
      {error && <div style={{ color:"#FF3B30", fontSize:13, background:"#FFF2F0", borderRadius:8, padding:"8px 12px", marginBottom:12 }}>{error}</div>}
      {preview ? (
        <div>
          <div style={{ fontSize:13, color:"#34C759", fontWeight:700, marginBottom:12 }}>✓ {preview.length} tubos encontrados</div>
          <div style={{ overflowX:"auto", borderRadius:14, boxShadow:"0 1px 6px rgba(0,0,0,0.08)", background:"#fff", maxHeight:280, overflowY:"auto", marginBottom:16 }}>
            <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12 }}>
              <thead><tr>{CSV_COLUMNS.map(c=><th key={c} style={{ background:"#F2F2F7", padding:"8px 10px", textAlign:"left", fontWeight:700, color:"#3A3A3C", borderBottom:"1px solid #E5E5EA", whiteSpace:"nowrap", position:"sticky", top:0 }}>{COL_LABELS[c]}</th>)}</tr></thead>
              <tbody>{preview.slice(0,10).map((row,i)=>(<tr key={i}>{CSV_COLUMNS.map(c=><td key={c} style={{ padding:"8px 10px", borderBottom:"1px solid #F2F2F7", color:"#1C1C1E" }}>{row[c]||"—"}</td>)}</tr>))}</tbody>
            </table>
          </div>
          {preview.length>10&&<div style={{ fontSize:12, color:"#8E8E93", marginBottom:12 }}>Exibindo 10 de {preview.length} linhas</div>}
          <div style={{ display:"flex", gap:10 }}>
            <button style={{ background:"linear-gradient(135deg,#007AFF,#0051D4)", color:"#fff", border:"none", borderRadius:12, padding:"11px 22px", fontSize:14, fontWeight:600, cursor:"pointer", display:"flex", alignItems:"center", gap:8 }} onClick={()=>onImport(preview)}>📥 Importar {preview.length} Tubos</button>
            <button style={{ background:"#F2F2F7", color:"#007AFF", border:"none", borderRadius:12, padding:"11px 22px", fontSize:14, fontWeight:600, cursor:"pointer" }} onClick={()=>setPreview(null)}>Cancelar</button>
          </div>
        </div>
      ) : (
        <div style={{ background:"#fff", borderRadius:14, padding:"16px 18px", boxShadow:"0 1px 4px rgba(0,0,0,0.07)" }}>
          <div style={{ fontSize:12, fontWeight:700, color:"#3A3A3C", marginBottom:8 }}>Colunas esperadas:</div>
          <div style={{ fontSize:12, color:"#8E8E93", lineHeight:2 }}>{Object.values(COL_LABELS).join("  ·  ")}</div>
        </div>
      )}
    </div>
  );
}

// ─── STAGE VIEW (grouped by lote) ────────────────────────────────────────────
function StageView({ stage, rows, user, onAdvance, onReturn }) {
  const [selectedLotes, setSelectedLotes]   = useState(new Set());
  const [loteTratativas, setLoteTratativas] = useState({});
  const [expandedLotes, setExpandedLotes]   = useState(new Set());
  const [showConfirm, setShowConfirm]       = useState(false);
  const [showReturnConfirm, setShowReturnConfirm] = useState(false);
  const tableRef = useRef(null);

  const isLast    = stage.id === 7;
  const nextStage = STAGES.find(s => s.id === stage.id + 1);
  const prevStage = STAGES.find(s => s.id === stage.id - 1);
  const groups    = groupByLote(rows);

  function toggleLote(lote) {
    const n = new Set(selectedLotes);
    n.has(lote) ? n.delete(lote) : n.add(lote);
    setSelectedLotes(n);
  }
  function toggleAll(checked) {
    setSelectedLotes(checked ? new Set(groups.map(g => g.lote)) : new Set());
  }
  function toggleExpand(lote) {
    const n = new Set(expandedLotes);
    n.has(lote) ? n.delete(lote) : n.add(lote);
    setExpandedLotes(n);
  }

  function handleAdvance() {
    const toMoveRows = rows.filter(r => selectedLotes.has(r.lote || r._id)).map(r => ({
      ...r,
      tratativa: loteTratativas[r.lote || r._id] ?? r.tratativa,
      history: [...(r.history||[]), {
        stage: stage.id, stageLabel: stage.label,
        user: user.name, dept: user.dept,
        tratativa: loteTratativas[r.lote || r._id] ?? r.tratativa,
        date: new Date().toLocaleString("pt-BR"),
      }],
    }));
    const remaining = rows.filter(r => !selectedLotes.has(r.lote || r._id));
    onAdvance(toMoveRows, remaining);
    setSelectedLotes(new Set());
    setLoteTratativas({});
    setShowConfirm(false);
  }

  function handleReturn() {
    const toReturnRows = rows.filter(r => selectedLotes.has(r.lote || r._id)).map(r => ({
      ...r,
      tratativa: loteTratativas[r.lote || r._id] ?? r.tratativa,
      history: [...(r.history||[]), {
        stage: stage.id, stageLabel: `${stage.label} → retorno`,
        user: user.name, dept: user.dept,
        tratativa: loteTratativas[r.lote || r._id] ?? r.tratativa,
        date: new Date().toLocaleString("pt-BR"),
      }],
    }));
    const remaining = rows.filter(r => !selectedLotes.has(r.lote || r._id));
    onReturn(toReturnRows, remaining);
    setSelectedLotes(new Set());
    setLoteTratativas({});
    setShowReturnConfirm(false);
  }

  const selectedCount = rows.filter(r => selectedLotes.has(r.lote || r._id)).length;
  const selectedLotesCount = selectedLotes.size;

  if (rows.length === 0) {
    return (
      <div>
        <div style={{ display:"flex", alignItems:"center", gap:12, marginBottom:16 }}>
          <div style={{ fontSize:22, fontWeight:800, color:"#1C1C1E", letterSpacing:"-0.4px" }}>{stage.label}</div>
          <DeptTag dept={stage.dept} />
        </div>
        <div style={{ textAlign:"center", padding:"56px 24px", background:"#fff", borderRadius:16, boxShadow:"0 1px 4px rgba(0,0,0,0.07)" }}>
          <div style={{ fontSize:44, marginBottom:12 }}>📭</div>
          <div style={{ fontSize:16, fontWeight:700, color:"#3A3A3C", marginBottom:4 }}>Nenhum tubo nesta etapa</div>
          <div style={{ fontSize:13, color:"#8E8E93" }}>Os tubos aparecerão aqui quando forem enviados para esta etapa.</div>
        </div>
      </div>
    );
  }

  return (
    <div style={{ display:"flex", flexDirection:"column", gap:0 }}>
      {/* Top bar: title + action buttons */}
      <div style={{ display:"flex", alignItems:"flex-start", justifyContent:"space-between", marginBottom:14, flexWrap:"wrap", gap:10 }}>
        <div>
          <div style={{ fontSize:22, fontWeight:800, color:"#1C1C1E", letterSpacing:"-0.4px", marginBottom:6 }}>{stage.label}</div>
          <div style={{ display:"flex", alignItems:"center", gap:10, flexWrap:"wrap" }}>
            <DeptTag dept={stage.dept} />
            <span style={{ fontSize:12, color:"#8E8E93" }}>{groups.length} lote{groups.length!==1?"s":""} · {rows.length} tubo{rows.length!==1?"s":""}</span>
            {selectedLotes.size > 0 && (
              <span style={{ fontSize:12, fontWeight:700, color:"#007AFF", background:"rgba(0,122,255,0.08)", borderRadius:8, padding:"3px 10px" }}>
                {selectedLotesCount} lote{selectedLotesCount!==1?"s":""} selecionado{selectedLotesCount!==1?"s":""} ({selectedCount} tubo{selectedCount!==1?"s":""})
              </span>
            )}
          </div>
        </div>
        {/* Action buttons top right */}
        <div style={{ display:"flex", gap:8, flexWrap:"wrap" }}>
          {stage.id > 1 && onReturn && (
            <button
              style={{ background: selectedLotes.size===0?"#F2F2F7":"linear-gradient(135deg,#FF9500,#FF6B00)", color: selectedLotes.size===0?"#8E8E93":"#fff", border:"none", borderRadius:12, padding:"10px 18px", fontSize:13, fontWeight:600, cursor: selectedLotes.size===0?"not-allowed":"pointer", display:"flex", alignItems:"center", gap:6, opacity: selectedLotes.size===0?0.5:1 }}
              disabled={selectedLotes.size === 0}
              onClick={() => setShowReturnConfirm(true)}
            >
              ← Retornar para {prevStage?.short}
            </button>
          )}
          <button
            style={{ background: selectedLotes.size===0?"#E5E5EA":"linear-gradient(135deg,#007AFF,#0051D4)", color: selectedLotes.size===0?"#8E8E93":"#fff", border:"none", borderRadius:12, padding:"10px 18px", fontSize:13, fontWeight:600, cursor: selectedLotes.size===0?"not-allowed":"pointer", display:"flex", alignItems:"center", gap:6, opacity: selectedLotes.size===0?0.5:1 }}
            disabled={selectedLotes.size === 0}
            onClick={() => setShowConfirm(true)}
          >
            {isLast ? `✅ Finalizar (${selectedLotesCount} lote${selectedLotesCount!==1?"s":""})` : `→ ${nextStage?.short} (${selectedLotesCount} lote${selectedLotesCount!==1?"s":""})`}
          </button>
        </div>
      </div>

      {/* Table with sticky header, dual scrollbars */}
      <div style={{ position:"relative" }}>
        {/* Scrollable table container */}
        <div
          ref={tableRef}
          style={{ borderRadius:14, boxShadow:"0 1px 6px rgba(0,0,0,0.08)", background:"#fff", overflowX:"auto", overflowY:"auto", maxHeight:"calc(100vh - 340px)", minHeight:200 }}
        >
          <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13, minWidth:900 }}>
            <thead>
              <tr>
                <th style={{ background:"#0A2240", padding:"10px 8px", textAlign:"center", fontWeight:700, color:"#fff", borderBottom:"2px solid #1A3A5C", whiteSpace:"nowrap", fontSize:11, position:"sticky", top:0, zIndex:10, width:36 }}>
                  <input type="checkbox" style={{ accentColor:"#5AC8FA", cursor:"pointer", width:16, height:16 }}
                    checked={selectedLotes.size===groups.length&&groups.length>0}
                    onChange={e=>toggleAll(e.target.checked)} />
                </th>
                <th style={{ background:"#0A2240", padding:"10px 8px", textAlign:"center", fontWeight:700, color:"#fff", borderBottom:"2px solid #1A3A5C", whiteSpace:"nowrap", fontSize:11, position:"sticky", top:0, zIndex:10, width:36 }}>▶</th>
                <th style={{ background:"#0A2240", padding:"10px 12px", textAlign:"left", fontWeight:700, color:"#fff", borderBottom:"2px solid #1A3A5C", whiteSpace:"nowrap", fontSize:11, position:"sticky", top:0, zIndex:10 }}>Lote</th>
                <th style={{ background:"#0A2240", padding:"10px 12px", textAlign:"left", fontWeight:700, color:"#fff", borderBottom:"2px solid #1A3A5C", whiteSpace:"nowrap", fontSize:11, position:"sticky", top:0, zIndex:10 }}>Qtd Tubos</th>
                <th style={{ background:"#0A2240", padding:"10px 12px", textAlign:"left", fontWeight:700, color:"#fff", borderBottom:"2px solid #1A3A5C", whiteSpace:"nowrap", fontSize:11, position:"sticky", top:0, zIndex:10 }}>IPPNs</th>
                {["ultima_ordem","qualidade_qts","deposito_sap","motivo_bloqueio","motivo_bloqueio_texto","razao_bloq"].map(c=>(
                  <th key={c} style={{ background:"#0A2240", padding:"10px 12px", textAlign:"left", fontWeight:700, color:"#fff", borderBottom:"2px solid #1A3A5C", whiteSpace:"nowrap", fontSize:11, position:"sticky", top:0, zIndex:10 }}>{COL_LABELS[c]}</th>
                ))}
                <th style={{ background:"#1A5C2A", padding:"10px 12px", textAlign:"left", fontWeight:700, color:"#fff", borderBottom:"2px solid #1A7A3A", whiteSpace:"nowrap", fontSize:11, position:"sticky", top:0, zIndex:10, minWidth:200 }}>Tratativa do Lote</th>
                {stage.id > 1 && <th style={{ background:"#0A2240", padding:"10px 12px", textAlign:"left", fontWeight:700, color:"#fff", borderBottom:"2px solid #1A3A5C", whiteSpace:"nowrap", fontSize:11, position:"sticky", top:0, zIndex:10 }}>Histórico</th>}
              </tr>
            </thead>
            <tbody>
              {groups.map((group, gi) => {
                const isExpanded = expandedLotes.has(group.lote);
                const isSelected = selectedLotes.has(group.lote);
                const firstRow = group.rows[0];
                const ippnList = group.rows.map(r => r.ippn).filter(Boolean).join(", ");
                const rowBg = gi % 2 === 0 ? "#fff" : "#F8F9FB";
                const selBg = isSelected ? "rgba(0,122,255,0.06)" : rowBg;

                return [
                  // ── Group summary row
                  <tr key={`group-${group.lote}`} style={{ background: selBg }}>
                    <td style={{ padding:"10px 8px", borderBottom:"1px solid #F0F0F5", textAlign:"center", verticalAlign:"middle" }}>
                      <input type="checkbox" style={{ accentColor:"#007AFF", cursor:"pointer", width:16, height:16 }}
                        checked={isSelected} onChange={()=>toggleLote(group.lote)} />
                    </td>
                    <td style={{ padding:"10px 8px", borderBottom:"1px solid #F0F0F5", textAlign:"center", verticalAlign:"middle" }}>
                      <button
                        onClick={()=>toggleExpand(group.lote)}
                        style={{ background:"none", border:"none", cursor:"pointer", fontSize:14, color:"#007AFF", fontWeight:700, padding:"2px 4px", lineHeight:1 }}
                        title={isExpanded?"Recolher":"Expandir IPPNs"}
                      >
                        {isExpanded ? "▼" : "▶"}
                      </button>
                    </td>
                    <td style={{ padding:"10px 12px", borderBottom:"1px solid #F0F0F5", verticalAlign:"middle" }}>
                      <span style={{ fontWeight:700, color:"#1C1C1E", fontSize:13 }}>{group.lote || "—"}</span>
                    </td>
                    <td style={{ padding:"10px 12px", borderBottom:"1px solid #F0F0F5", verticalAlign:"middle" }}>
                      <span style={{ background:"#E8F4FD", color:"#1A6FA8", borderRadius:6, padding:"2px 8px", fontSize:12, fontWeight:700 }}>{group.rows.length}</span>
                    </td>
                    <td style={{ padding:"10px 12px", borderBottom:"1px solid #F0F0F5", verticalAlign:"middle", maxWidth:200, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap", fontSize:12, color:"#555" }}>
                      {ippnList || "—"}
                    </td>
                    {["ultima_ordem","qualidade_qts","deposito_sap","motivo_bloqueio","motivo_bloqueio_texto","razao_bloq"].map(c=>(
                      <td key={c} style={{ padding:"10px 12px", borderBottom:"1px solid #F0F0F5", verticalAlign:"middle", fontSize:12, color:"#3A3A3C" }}>
                        {firstRow[c]||<span style={{ color:"#C7C7CC" }}>—</span>}
                      </td>
                    ))}
                    <td style={{ padding:"10px 12px", borderBottom:"1px solid #F0F0F5", verticalAlign:"middle", background: gi%2===0?"rgba(26,90,42,0.03)":"rgba(26,90,42,0.06)" }}>
                      <input
                        type="text"
                        style={{ border:"1.5px solid #E5E5EA", borderRadius:8, padding:"6px 10px", fontSize:12, width:"100%", outline:"none", background:"#fff", minWidth:180 }}
                        placeholder="Registrar tratativa do lote…"
                        value={loteTratativas[group.lote] ?? group.tratativa ?? ""}
                        onChange={e => setLoteTratativas(t=>({...t,[group.lote]:e.target.value}))}
                      />
                    </td>
                    {stage.id > 1 && (
                      <td style={{ padding:"10px 12px", borderBottom:"1px solid #F0F0F5", verticalAlign:"middle" }}>
                        {firstRow.history?.length > 0 ? (
                          <details>
                            <summary style={{ fontSize:11, color:"#007AFF", cursor:"pointer", fontWeight:600 }}>{firstRow.history.length} etapa{firstRow.history.length!==1?"s":""}</summary>
                            <div style={{ marginTop:6 }}>
                              {firstRow.history.map((h,hi)=>(
                                <div key={hi} style={{ fontSize:11, color:"#3A3A3C", marginBottom:6, borderLeft:"2px solid #E5E5EA", paddingLeft:8, lineHeight:1.6 }}>
                                  <strong>{h.stageLabel}</strong><br/>{h.user}<br/>
                                  {h.tratativa&&<><em style={{ color:"#555" }}>"{h.tratativa}"</em><br/></>}
                                  <span style={{ color:"#C7C7CC" }}>{h.date}</span>
                                </div>
                              ))}
                            </div>
                          </details>
                        ) : <span style={{ color:"#C7C7CC", fontSize:12 }}>—</span>}
                      </td>
                    )}
                  </tr>,
                  // ── Expanded IPPN rows
                  ...(isExpanded ? group.rows.map((row, ri) => (
                    <tr key={`expanded-${row._id}`} style={{ background:"#F0F7FF" }}>
                      <td style={{ padding:"8px 8px", borderBottom:"1px solid #E8F4FD" }}></td>
                      <td style={{ padding:"8px 8px", borderBottom:"1px solid #E8F4FD", textAlign:"center" }}>
                        <span style={{ fontSize:9, color:"#8E8E93" }}>└</span>
                      </td>
                      <td style={{ padding:"8px 12px", borderBottom:"1px solid #E8F4FD", fontSize:11, color:"#8E8E93" }}>
                        <span style={{ fontWeight:600, color:"#1A6FA8" }}>IPPN {ri+1}/{group.rows.length}</span>
                      </td>
                      <td style={{ padding:"8px 12px", borderBottom:"1px solid #E8F4FD" }}></td>
                      <td style={{ padding:"8px 12px", borderBottom:"1px solid #E8F4FD", fontSize:12, fontWeight:700, color:"#1A6FA8" }}>{row.ippn||"—"}</td>
                      {["ultima_ordem","qualidade_qts","deposito_sap","motivo_bloqueio","motivo_bloqueio_texto","razao_bloq"].map(c=>(
                        <td key={c} style={{ padding:"8px 12px", borderBottom:"1px solid #E8F4FD", fontSize:11, color:"#555" }}>{row[c]||"—"}</td>
                      ))}
                      <td style={{ padding:"8px 12px", borderBottom:"1px solid #E8F4FD" }}>
                        <span style={{ fontSize:11, color:"#8E8E93" }}>Nº Cassete: <strong style={{ color:"#3A3A3C" }}>{row.num_cassete||"—"}</strong></span>
                      </td>
                      {stage.id > 1 && <td style={{ borderBottom:"1px solid #E8F4FD" }}></td>}
                    </tr>
                  )) : [])
                ];
              })}
            </tbody>
          </table>
        </div>
      </div>

      {/* Confirm modals */}
      {showConfirm && (
        <Modal
          title={isLast ? "Finalizar lotes?" : `Enviar para "${nextStage?.label}"?`}
          body={`${selectedLotesCount} lote(s) / ${selectedCount} tubo(s) serão ${isLast ? "marcados como Pendente Execução" : `enviados para "${nextStage?.label}"`}. Registrado em nome de ${user.name}.`}
          onConfirm={handleAdvance}
          onCancel={()=>setShowConfirm(false)}
          confirmLabel={isLast ? "Finalizar" : "Confirmar envio"}
        />
      )}
      {showReturnConfirm && (
        <Modal
          title={`Retornar para "${prevStage?.label}"?`}
          body={`${selectedLotesCount} lote(s) / ${selectedCount} tubo(s) serão retornados para a etapa "${prevStage?.label}". Esta ação será registrada em nome de ${user.name}.`}
          onConfirm={handleReturn}
          onCancel={()=>setShowReturnConfirm(false)}
          confirmLabel="Confirmar retorno"
          danger
        />
      )}
    </div>
  );
}

// ─── DASHBOARD ────────────────────────────────────────────────────────────────
function Dashboard({ stageData, onSelectStage }) {
  const allRows    = Object.values(stageData).flat();
  const total      = allRows.length;
  const etapa1     = stageData[1]?.length || 0;
  const inFlow     = [2,3,4,5,6,7].reduce((s,id)=>s+(stageData[id]?.length||0),0);
  const pendExec   = stageData[8]?.length || 0;

  // Por depósito — all stages
  const byDeposito = {};
  allRows.forEach(r => {
    const dep = r.deposito_sap || "Sem Depósito";
    byDeposito[dep] = (byDeposito[dep]||0)+1;
  });

  // Por departamento
  const byDept = {};
  STAGES.filter(s=>s.id!==8&&s.dept).forEach(s => {
    const cnt = stageData[s.id]?.length || 0;
    if (!byDept[s.dept]) byDept[s.dept] = 0;
    byDept[s.dept] += cnt;
  });

  const kpis = [
    { label:"Total Bloqueado",    value:total,    color:"#FF3B30", icon:"🔴" },
    { label:"Análise Pendente",   value:etapa1,   color:"#007AFF", icon:"🔵", stageId:1 },
    { label:"Em Fluxo",          value:inFlow,   color:"#FF9500", icon:"🟠" },
    { label:"Pendente Execução",  value:pendExec, color:"#34C759", icon:"🟢", stageId:8 },
  ];

  return (
    <div>
      <div style={{ fontSize:22, fontWeight:800, color:"#1C1C1E", letterSpacing:"-0.4px", marginBottom:4 }}>Dashboard</div>
      <div style={{ fontSize:13, color:"#8E8E93", marginBottom:20 }}>Visão geral do fluxo de liberação</div>

      {/* KPI row */}
      <div style={{ display:"flex", gap:12, marginBottom:24, flexWrap:"wrap" }}>
        {kpis.map(k=>(
          <div key={k.label}
            style={{ background:"#fff", borderRadius:16, padding:"18px 20px", flex:1, minWidth:140, boxShadow:"0 1px 6px rgba(0,0,0,0.08)", borderLeft:`4px solid ${k.color}`, cursor:k.stageId?"pointer":"default" }}
            onClick={()=>k.stageId&&onSelectStage(k.stageId)}
          >
            <div style={{ fontSize:11, color:"#8E8E93", fontWeight:600, marginBottom:4, textTransform:"uppercase", letterSpacing:"0.5px" }}>{k.label}</div>
            <div style={{ fontSize:36, fontWeight:900, color:k.color, letterSpacing:"-1.5px", lineHeight:1 }}>{k.value}</div>
          </div>
        ))}
      </div>

      {/* Stage cards */}
      <div style={{ fontSize:14, fontWeight:700, color:"#3A3A3C", marginBottom:12 }}>Tubos por etapa</div>
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(148px,1fr))", gap:10, marginBottom:28 }}>
        {STAGES.filter(s=>s.id!==8).map((stage,idx)=>{
          const count = stageData[stage.id]?.length||0;
          const color = STAGE_COLORS[idx];
          return (
            <div key={stage.id}
              style={{ background:"#fff", borderRadius:14, padding:"14px", boxShadow:"0 1px 4px rgba(0,0,0,0.07)", borderTop:`3px solid ${color}`, cursor:"pointer", transition:"transform 0.15s,box-shadow 0.15s" }}
              onClick={()=>onSelectStage(stage.id)}
              onMouseEnter={e=>{e.currentTarget.style.transform="translateY(-2px)";e.currentTarget.style.boxShadow="0 5px 14px rgba(0,0,0,0.12)";}}
              onMouseLeave={e=>{e.currentTarget.style.transform="";e.currentTarget.style.boxShadow="0 1px 4px rgba(0,0,0,0.07)";}}
            >
              <div style={{ fontSize:30, fontWeight:900, color:"#1C1C1E", letterSpacing:"-1px" }}>{count}</div>
              <div style={{ fontSize:9, fontWeight:800, color, marginTop:2, textTransform:"uppercase", letterSpacing:"0.4px" }}>ETAPA {stage.id}</div>
              <div style={{ fontSize:11, color:"#6B6B6B", marginTop:3, lineHeight:1.3 }}>{stage.short}</div>
              {count>0&&<div style={{ fontSize:10, color:"#007AFF", fontWeight:700, marginTop:8 }}>Ver →</div>}
            </div>
          );
        })}
      </div>

      {/* By department */}
      <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:16, marginBottom:24 }}>
        <div style={{ background:"#fff", borderRadius:16, padding:"18px", boxShadow:"0 1px 6px rgba(0,0,0,0.08)" }}>
          <div style={{ fontSize:14, fontWeight:700, color:"#3A3A3C", marginBottom:14 }}>Pendentes por Departamento</div>
          {Object.entries(byDept).length === 0
            ? <div style={{ fontSize:13, color:"#8E8E93" }}>Nenhum tubo em fluxo</div>
            : Object.entries(byDept).map(([dept, cnt]) => {
                const c = DEPT_COLORS[dept]||{bg:"#F2F2F7",fg:"#3A3A3C"};
                const pct = total > 0 ? Math.round(cnt/total*100) : 0;
                return (
                  <div key={dept} style={{ marginBottom:12 }}>
                    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:5 }}>
                      <span style={{ fontSize:12, fontWeight:600, color:"#3A3A3C" }}>{DEPT_FULL[dept]||dept}</span>
                      <span style={{ fontSize:13, fontWeight:800, color:c.fg }}>{cnt}</span>
                    </div>
                    <div style={{ height:6, background:"#F2F2F7", borderRadius:3, overflow:"hidden" }}>
                      <div style={{ width:`${pct}%`, height:"100%", background:c.fg, borderRadius:3, transition:"width 0.5s" }}/>
                    </div>
                  </div>
                );
              })
          }
        </div>

        {/* By deposito */}
        <div style={{ background:"#fff", borderRadius:16, padding:"18px", boxShadow:"0 1px 6px rgba(0,0,0,0.08)" }}>
          <div style={{ fontSize:14, fontWeight:700, color:"#3A3A3C", marginBottom:14 }}>Bloqueio por Depósito</div>
          {Object.entries(byDeposito).length === 0
            ? <div style={{ fontSize:13, color:"#8E8E93" }}>Nenhum tubo no sistema</div>
            : Object.entries(byDeposito).sort((a,b)=>b[1]-a[1]).map(([dep, cnt],i) => {
                const colors = ["#007AFF","#34C759","#FF9500","#AF52DE","#FF2D55","#5AC8FA"];
                const color = colors[i % colors.length];
                const pct = total > 0 ? Math.round(cnt/total*100) : 0;
                return (
                  <div key={dep} style={{ marginBottom:12 }}>
                    <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:5 }}>
                      <span style={{ fontSize:12, fontWeight:600, color:"#3A3A3C" }}>{dep}</span>
                      <span style={{ fontSize:13, fontWeight:800, color }}>{cnt}</span>
                    </div>
                    <div style={{ height:6, background:"#F2F2F7", borderRadius:3, overflow:"hidden" }}>
                      <div style={{ width:`${pct}%`, height:"100%", background:color, borderRadius:3, transition:"width 0.5s" }}/>
                    </div>
                  </div>
                );
              })
          }
        </div>
      </div>

      {total === 0 && (
        <div style={{ textAlign:"center", padding:"56px 24px", background:"#fff", borderRadius:16, boxShadow:"0 1px 4px rgba(0,0,0,0.07)" }}>
          <div style={{ fontSize:44, marginBottom:12 }}>📋</div>
          <div style={{ fontSize:16, fontWeight:700, color:"#3A3A3C", marginBottom:4 }}>Nenhum tubo no sistema</div>
          <div style={{ fontSize:13, color:"#8E8E93" }}>Acesse a Etapa 1 para importar o arquivo de tubos bloqueados.</div>
        </div>
      )}
    </div>
  );
}

// ─── COMPLETED VIEW ───────────────────────────────────────────────────────────
function PendingExecView({ rows }) {
  const groups = groupByLote(rows);
  const [expandedLotes, setExpandedLotes] = useState(new Set());
  function toggleExpand(lote) {
    const n = new Set(expandedLotes);
    n.has(lote) ? n.delete(lote) : n.add(lote);
    setExpandedLotes(n);
  }
  if (rows.length === 0) return (
    <div>
      <div style={{ fontSize:22, fontWeight:800, color:"#1C1C1E", marginBottom:4 }}>Pendente Execução</div>
      <div style={{ textAlign:"center", padding:"56px", background:"#fff", borderRadius:16, marginTop:12 }}>
        <div style={{ fontSize:44, marginBottom:12 }}>🏁</div>
        <div style={{ fontSize:16, fontWeight:700, color:"#3A3A3C" }}>Nenhum tubo concluído ainda</div>
      </div>
    </div>
  );
  return (
    <div>
      <div style={{ fontSize:22, fontWeight:800, color:"#1C1C1E", marginBottom:4 }}>Pendente Execução</div>
      <div style={{ fontSize:13, color:"#8E8E93", marginBottom:16 }}>{groups.length} lote{groups.length!==1?"s":""} · {rows.length} tubo{rows.length!==1?"s":""} completaram o fluxo.</div>
      <div style={{ borderRadius:14, boxShadow:"0 1px 6px rgba(0,0,0,0.08)", background:"#fff", overflowX:"auto", overflowY:"auto", maxHeight:"calc(100vh - 300px)" }}>
        <table style={{ width:"100%", borderCollapse:"collapse", fontSize:13, minWidth:800 }}>
          <thead>
            <tr>
              {["","Lote","Qtd","IPPNs",...Object.values(COL_LABELS).slice(0,6),"Histórico"].map((h,i)=>(
                <th key={i} style={{ background:"#0A2240", padding:"10px 12px", textAlign:"left", fontWeight:700, color:"#fff", borderBottom:"2px solid #1A3A5C", whiteSpace:"nowrap", fontSize:11, position:"sticky", top:0, zIndex:10 }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {groups.map((group,gi)=>{
              const isExpanded = expandedLotes.has(group.lote);
              const firstRow = group.rows[0];
              const ippnList = group.rows.map(r=>r.ippn).filter(Boolean).join(", ");
              return [
                <tr key={`g-${group.lote}`} style={{ background: gi%2===0?"#fff":"#F8F9FB" }}>
                  <td style={{ padding:"10px 8px", borderBottom:"1px solid #F0F0F5", textAlign:"center" }}>
                    <button onClick={()=>toggleExpand(group.lote)} style={{ background:"none", border:"none", cursor:"pointer", fontSize:12, color:"#007AFF", fontWeight:700 }}>
                      {isExpanded?"▼":"▶"}
                    </button>
                  </td>
                  <td style={{ padding:"10px 12px", borderBottom:"1px solid #F0F0F5", fontWeight:700, color:"#1C1C1E" }}>{group.lote||"—"}</td>
                  <td style={{ padding:"10px 12px", borderBottom:"1px solid #F0F0F5" }}><span style={{ background:"#EDF7EE", color:"#1A7A3A", borderRadius:6, padding:"2px 8px", fontSize:12, fontWeight:700 }}>{group.rows.length}</span></td>
                  <td style={{ padding:"10px 12px", borderBottom:"1px solid #F0F0F5", fontSize:12, color:"#555", maxWidth:180, overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{ippnList||"—"}</td>
                  {["ultima_ordem","qualidade_qts","deposito_sap","motivo_bloqueio","motivo_bloqueio_texto","razao_bloq"].map(c=>(
                    <td key={c} style={{ padding:"10px 12px", borderBottom:"1px solid #F0F0F5", fontSize:12, color:"#3A3A3C" }}>{firstRow[c]||"—"}</td>
                  ))}
                  <td style={{ padding:"10px 12px", borderBottom:"1px solid #F0F0F5" }}>
                    <details>
                      <summary style={{ fontSize:11, color:"#34C759", cursor:"pointer", fontWeight:700 }}>✓ {firstRow.history?.length||0} etapas</summary>
                      <div style={{ marginTop:6 }}>
                        {firstRow.history?.map((h,hi)=>(
                          <div key={hi} style={{ fontSize:11, color:"#3A3A3C", marginBottom:6, borderLeft:"2px solid #34C759", paddingLeft:8, lineHeight:1.6 }}>
                            <strong>{h.stageLabel}</strong><br/>{h.user}<br/>
                            {h.tratativa&&<><em>"{h.tratativa}"</em><br/></>}
                            <span style={{ color:"#C7C7CC" }}>{h.date}</span>
                          </div>
                        ))}
                      </div>
                    </details>
                  </td>
                </tr>,
                ...(isExpanded ? group.rows.map((row,ri)=>(
                  <tr key={`exp-${row._id}`} style={{ background:"#F0F7FF" }}>
                    <td style={{ padding:"8px 8px", borderBottom:"1px solid #E8F4FD", textAlign:"center" }}><span style={{ fontSize:9, color:"#8E8E93" }}>└</span></td>
                    <td style={{ padding:"8px 12px", borderBottom:"1px solid #E8F4FD", fontSize:11, color:"#8E8E93" }}>IPPN {ri+1}</td>
                    <td style={{ borderBottom:"1px solid #E8F4FD" }}></td>
                    <td style={{ padding:"8px 12px", borderBottom:"1px solid #E8F4FD", fontSize:12, fontWeight:700, color:"#1A6FA8" }}>{row.ippn||"—"}</td>
                    {["ultima_ordem","qualidade_qts","deposito_sap","motivo_bloqueio","motivo_bloqueio_texto","razao_bloq"].map(c=>(
                      <td key={c} style={{ padding:"8px 12px", borderBottom:"1px solid #E8F4FD", fontSize:11, color:"#555" }}>{row[c]||"—"}</td>
                    ))}
                    <td style={{ borderBottom:"1px solid #E8F4FD" }}></td>
                  </tr>
                )) : [])
              ];
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
  const [activeStage, setActive]  = useState(0);
  const [toast, setToast]         = useState("");
  const [showLogout, setShowLogout] = useState(false);

  function showToast(msg) {
    setToast(msg);
    setTimeout(()=>setToast(""), 3200);
  }

  function handleImport(rows) {
    setStageData(d=>({...d, 1:[...d[1],...rows]}));
    setActive(1);
    showToast(`${rows.length} tubos importados para a Etapa 1`);
  }

  function handleAdvance(toMove, remaining) {
    const nextId = activeStage + 1;
    const nextStage = STAGES.find(s=>s.id===nextId);
    setStageData(d=>({ ...d, [activeStage]:remaining, [nextId]:[...(d[nextId]||[]),...toMove] }));
    showToast(`${toMove.length} tubo(s) enviado(s) para "${nextStage?.label}"`);
  }

  function handleReturn(toReturn, remaining) {
    const prevId = activeStage - 1;
    const prevStage = STAGES.find(s=>s.id===prevId);
    setStageData(d=>({ ...d, [activeStage]:remaining, [prevId]:[...(d[prevId]||[]),...toReturn] }));
    showToast(`${toReturn.length} tubo(s) retornado(s) para "${prevStage?.label}"`);
  }

  if (!user) return <LoginPage onLogin={setUser} />;

  return (
    <div style={{ fontFamily:"-apple-system, BlinkMacSystemFont, 'SF Pro Display', 'Segoe UI', sans-serif", minHeight:"100vh", background:"#F2F2F7", color:"#1C1C1E" }}>
      {/* Header */}
      <header style={{ background:"linear-gradient(135deg,#0A2240 0%,#1A3A5C 100%)", padding:"0 24px", height:52, display:"flex", alignItems:"center", justifyContent:"space-between", boxShadow:"0 1px 0 rgba(255,255,255,0.08),0 4px 16px rgba(0,0,0,0.35)", position:"sticky", top:0, zIndex:100 }}>
        <div style={{ display:"flex", alignItems:"center", gap:12 }}>
          <div style={{ width:28, height:28, borderRadius:8, background:"rgba(255,255,255,0.12)", display:"flex", alignItems:"center", justifyContent:"center", fontSize:16 }}>⚗️</div>
          <span style={{ color:"#fff", fontSize:16, fontWeight:700, letterSpacing:"-0.3px" }}>Fluxo de Liberação</span>
        </div>
        <div style={{ display:"flex", alignItems:"center", gap:8, cursor:"pointer" }} onClick={()=>setShowLogout(true)}>
          <div style={{ textAlign:"right" }}>
            <div style={{ color:"#fff", fontSize:13, fontWeight:600 }}>{user.name}</div>
            <div style={{ color:"rgba(255,255,255,0.5)", fontSize:10 }}>{user.dept}</div>
          </div>
          <div style={{ width:30, height:30, borderRadius:"50%", background:"linear-gradient(135deg,#007AFF,#5AC8FA)", display:"flex", alignItems:"center", justifyContent:"center", fontSize:13, fontWeight:700, color:"#fff" }}>{user.name.charAt(0)}</div>
        </div>
      </header>

      {/* Pipeline */}
      <Pipeline stageData={stageData} activeStage={activeStage} onSelectStage={setActive} />

      {/* Nav tabs */}
      <div style={{ padding:"12px 24px 0", maxWidth:1200, margin:"0 auto" }}>
        <div style={{ display:"flex", gap:2, background:"#E5E5EA", borderRadius:10, padding:3, overflowX:"auto", flexWrap:"nowrap" }}>
          <button style={{ padding:"6px 14px", borderRadius:8, fontSize:12, fontWeight:activeStage===0?700:500, color:activeStage===0?"#1C1C1E":"#8E8E93", background:activeStage===0?"#fff":"transparent", border:"none", cursor:"pointer", whiteSpace:"nowrap", boxShadow:activeStage===0?"0 1px 4px rgba(0,0,0,0.12)":"none" }} onClick={()=>setActive(0)}>Dashboard</button>
          {STAGES.filter(s=>s.id!==8).map(s=>{
            const cnt = stageData[s.id]?.length||0;
            const isA = activeStage===s.id;
            return (
              <button key={s.id} style={{ padding:"6px 12px", borderRadius:8, fontSize:12, fontWeight:isA?700:500, color:isA?"#1C1C1E":"#8E8E93", background:isA?"#fff":"transparent", border:"none", cursor:"pointer", whiteSpace:"nowrap", boxShadow:isA?"0 1px 4px rgba(0,0,0,0.12)":"none" }} onClick={()=>setActive(s.id)}>
                E{s.id}
                {cnt>0&&<span style={{ marginLeft:4, background:"#FF3B30", color:"#fff", borderRadius:8, fontSize:9, fontWeight:700, padding:"1px 5px" }}>{cnt}</span>}
              </button>
            );
          })}
          <button style={{ padding:"6px 12px", borderRadius:8, fontSize:12, fontWeight:activeStage===8?700:500, color:activeStage===8?"#1C1C1E":"#8E8E93", background:activeStage===8?"#fff":"transparent", border:"none", cursor:"pointer", whiteSpace:"nowrap", boxShadow:activeStage===8?"0 1px 4px rgba(0,0,0,0.12)":"none" }} onClick={()=>setActive(8)}>
            Pend. Exec.
            {(stageData[8]?.length||0)>0&&<span style={{ marginLeft:4, background:"#34C759", color:"#fff", borderRadius:8, fontSize:9, fontWeight:700, padding:"1px 5px" }}>{stageData[8].length}</span>}
          </button>
        </div>
      </div>

      {/* Content */}
      <div style={{ padding:"18px 24px 40px", maxWidth:1200, margin:"0 auto" }}>
        {activeStage===0 && <Dashboard stageData={stageData} onSelectStage={setActive} />}

        {activeStage===1 && (
          stageData[1].length===0
            ? <ImportStep onImport={handleImport} />
            : <div>
                <StageView stage={STAGES[0]} rows={stageData[1]} user={user} onAdvance={handleAdvance} onReturn={null} />
                <div style={{ marginTop:14 }}>
                  <button style={{ background:"#F2F2F7", color:"#007AFF", border:"none", borderRadius:10, padding:"8px 16px", fontSize:12, fontWeight:600, cursor:"pointer" }}
                    onClick={()=>{ if(window.confirm("Isso removerá todos os tubos da Etapa 1. Continuar?")) setStageData(d=>({...d,1:[]})); }}>
                    ↩ Importar novo arquivo
                  </button>
                </div>
              </div>
        )}

        {activeStage>1 && activeStage<=7 && (
          <StageView
            stage={STAGES[activeStage-1]}
            rows={stageData[activeStage]}
            user={user}
            onAdvance={handleAdvance}
            onReturn={handleReturn}
          />
        )}

        {activeStage===8 && <PendingExecView rows={stageData[8]} />}
      </div>

      <Toast msg={toast} />

      {showLogout && (
        <Modal
          title="Sair da conta"
          body={`Deseja encerrar a sessão de ${user.name}?`}
          onConfirm={()=>{ setUser(null); setShowLogout(false); setStageData(initStageData()); setActive(0); }}
          onCancel={()=>setShowLogout(false)}
          confirmLabel="Sair"
          danger
        />
      )}
    </div>
  );
}
