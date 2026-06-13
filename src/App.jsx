import { useState, useCallback, useRef, useEffect } from "react";
import { createPortal } from "react-dom";
import * as XLSX from "xlsx";
import {
  LayoutGrid, GitBranch, History as HistoryIcon, Settings, Search,
  ChevronRight, ChevronDown, ArrowRight, ArrowLeft, CheckCircle2,
  Upload, LogOut, Plus, Pencil, Trash2, Shield, Users, Inbox,
  Clock, AlertTriangle, RotateCcw, FlaskConical, Loader2, Layers,
  BarChart3, Package, Building2, Receipt, PieChart, X,
} from "lucide-react";

// ─────────────────────────────────────────────────────────────────────────────
// CONSTANTS
// ─────────────────────────────────────────────────────────────────────────────
const STAGES = [
  { id:1, label:"Análise do Bloqueio",    short:"Análise de Bloqueio",  dept:"CQ Área Técnica" },
  { id:2, label:"Definição do Recurso",   short:"Avaliar Recurso",      dept:"Planejamento UAP" },
  { id:3, label:"Criação de Ordem",       short:"Criar Ordem",          dept:"Planejamento Central" },
  { id:4, label:"Instrução da Qualidade", short:"Atualizar IC",         dept:"CQ Área Técnica" },
  { id:5, label:"Liberação para Vínculo", short:"Desbloqueio",          dept:"CQ Lib. Intermediária" },
  { id:6, label:"Vínculo dos Lotes",      short:"Vínculo",              dept:"Planejamento UAP" },
  { id:7, label:"Ativação de Flag",       short:"Flag",                 dept:"CQ Lib. Intermediária" },
  { id:8, label:"Pendente Execução",      short:"Execução",             dept:"CQ Lib. Intermediária" },
];

const STAGE_COLORS = ["#0A84FF","#34C759","#FF9F0A","#AF52DE","#FF2D55","#5AC8FA","#FF6B35","#8E8E93"];

const DEPT_FULL = {
  "CQ Área Técnica":       "Controle da Qualidade",
  "Planejamento UAP":      "Planejamento UAP",
  "Planejamento Central":  "Planejamento Central",
  "CQ Lib. Intermediária": "Liberação de produtos",
  "Admin":                 "Administrador",
};

const DEPT_COLORS = {
  "CQ Área Técnica":       { bg:"#E8F4FD", fg:"#1A6FA8" },
  "Planejamento UAP":      { bg:"#FFF3E0", fg:"#B45309" },
  "Planejamento Central":  { bg:"#F3E8FF", fg:"#7C3AED" },
  "CQ Lib. Intermediária": { bg:"#EDF7EE", fg:"#1A7A3A" },
  "Admin":                 { bg:"#F2F2F7", fg:"#3A3A3C" },
};

// Default stage permissions per dept key
const DEPT_DEFAULT_STAGES = {
  "CQ Área Técnica":       [1,4],
  "Planejamento UAP":      [2,6],
  "Planejamento Central":  [3],
  "CQ Lib. Intermediária": [5,7,8],
  "Admin":                 [1,2,3,4,5,6,7,8],
};

const DEFAULT_USERS = [
  { id:"u1", email:"qualidade.tecnica@empresa.com", password:"123456", name:"Ana Paula Silva",  dept:"CQ Área Técnica",       allowedStages:[1,4],     active:true },
  { id:"u2", email:"qualidade.lib@empresa.com",     password:"123456", name:"Carlos Mendes",    dept:"CQ Lib. Intermediária", allowedStages:[5,7,8],   active:true },
  { id:"u3", email:"planejamento.uap@empresa.com",  password:"123456", name:"Fernanda Rocha",   dept:"Planejamento UAP",      allowedStages:[2,6],     active:true },
  { id:"u4", email:"planejamento.central@empresa.com",password:"123456",name:"Ricardo Alves",   dept:"Planejamento Central",  allowedStages:[3],       active:true },
  { id:"u5", email:"admin@empresa.com",             password:"admin",  name:"Administrador",    dept:"Admin",                 allowedStages:[1,2,3,4,5,6,7,8], active:true },
];

// All importable columns
const ALL_COLS = ["pedido","item","material","descricao","data_bloqueio","ultima_ordem","lote","ippn","qualidade_qts","deposito_sap","motivo_bloqueio","motivo_bloqueio_texto","razao_bloq","descricao_motivo","num_cassete"];
const COL_LABELS = {
  pedido:"Pedido", item:"Item", material:"Material", descricao:"Descrição", data_bloqueio:"Data Bloqueio",
  ultima_ordem:"Última Ordem", lote:"Lote", ippn:"IPPN", qualidade_qts:"Qualidade QTS",
  deposito_sap:"Depósito SAP", motivo_bloqueio:"Motivo Bloqueio", motivo_bloqueio_texto:"Motivo Bloqueio Texto",
  razao_bloq:"Razão Bloq.", descricao_motivo:"Descrição Motivo", num_cassete:"Nº Cassete",
};

// Options for "Definição" na Etapa 1 (recursos necessários)
const RECURSOS_OPTIONS = ["Inspeção VD","Inspeção Ultrassom","Serra","Inspeção EMI","Inspeção de Drift","Faceamento"];

// ─────────────────────────────────────────────────────────────────────────────
// DESIGN TOKENS — Apple HIG, adapted for desktop (NavigationSplitView pattern)
// ─────────────────────────────────────────────────────────────────────────────
const FONT = "-apple-system,BlinkMacSystemFont,'SF Pro Display','SF Pro Text','Segoe UI',Helvetica,Arial,sans-serif";
const SPRING = "cubic-bezier(0.34, 1.45, 0.55, 1)";
const EASE = "cubic-bezier(0.25, 0.1, 0.25, 1)";
const ACCENT = "#0A84FF";          // iOS systemBlue
// Display/large-title typeface — rounded SF Pro for a softer, "Apple report" feel
const TITLE_FONT = "'Manrope', 'SF Pro Rounded', ui-rounded, -apple-system, BlinkMacSystemFont, 'SF Pro Display', 'Segoe UI', Helvetica, Arial, sans-serif";
const SIDEBAR_W = 232;
const APP_BG = "#F2F2F4";          // systemGroupedBackground
const SURFACE = "#FFFFFF";
const SEPARATOR = "rgba(60,60,67,0.13)";
const HEADER_GLASS = "rgba(244,244,247,0.78)";
const SIDEBAR_GLASS = "rgba(255,255,255,0.65)";

// Shared layout container — desktop: use full available width
const CONTAINER = { padding:"0 28px" };
const WIDE = { maxWidth:1440, margin:"0 auto", padding:"0 28px 40px" };

// ─────────────────────────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────────────────────────
function normalizeHdr(h) {
  return String(h).trim().toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g,"")
    .replace(/[\s\-\/\\]+/g,"_").replace(/[^a-z0-9_]/g,"");
}

// Take 10 leftmost chars and format as DD/MM/AAAA
function formatDataBloqueio(raw) {
  if (raw==null || raw==="") return "";
  if (raw instanceof Date) {
    const d=String(raw.getDate()).padStart(2,"0");
    const m=String(raw.getMonth()+1).padStart(2,"0");
    const y=raw.getFullYear();
    return `${d}/${m}/${y}`;
  }
  const s = String(raw).trim().slice(0,10);
  let m = s.match(/^(\d{4})[-\/](\d{2})[-\/](\d{2})$/); // YYYY-MM-DD or YYYY/MM/DD
  if (m) return `${m[3]}/${m[2]}/${m[1]}`;
  m = s.match(/^(\d{2})[-\/](\d{2})[-\/](\d{4})$/); // DD-MM-YYYY or DD/MM/YYYY
  if (m) return `${m[1]}/${m[2]}/${m[3]}`;
  return s;
}

function rowsFromMatrix(matrix) {
  if (matrix.length < 2) return [];
  const headers = matrix[0].map(normalizeHdr);
  return matrix.slice(1).filter(v=>v.some(x=>String(x??"").trim())).map((vals,i)=>{
    const obj = { _id:`row_${Date.now()}_${i}`, tratativa:"", history:[], stageEnteredMs: Date.now() };
    ALL_COLS.forEach(col=>{
      const idx = headers.indexOf(col);
      const raw = idx>=0 ? vals[idx] : "";
      if (col==="data_bloqueio") {
        obj[col] = formatDataBloqueio(raw);
      } else {
        obj[col] = raw==null ? "" : String(raw).trim();
      }
    });
    return obj;
  });
}

async function parseFile(file) {
  return new Promise((resolve,reject)=>{
    const reader = new FileReader();
    const isXLS = /\.(xlsx?|xls)$/i.test(file.name);
    reader.onload = e=>{
      try {
        if (isXLS) {
          const wb = XLSX.read(e.target.result,{type:"array",cellDates:true});
          const ws = wb.Sheets[wb.SheetNames[0]];
          const matrix = XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
          resolve(rowsFromMatrix(matrix));
        } else {
          const text = new TextDecoder("utf-8").decode(e.target.result);
          const lines = text.trim().split(/\r?\n/);
          if (lines.length<2){resolve([]);return;}
          const sep = lines[0].includes(";")?";":(lines[0].includes("\t")?"\t":",");
          const matrix = lines.map(l=>l.split(sep).map(v=>v.trim().replace(/^"|"$/g,"")));
          resolve(rowsFromMatrix(matrix));
        }
      } catch(err){reject(err);}
    };
    reader.onerror=reject;
    reader.readAsArrayBuffer(file);
  });
}

function groupByLote(rows) {
  const map = new Map();
  rows.forEach(row=>{
    const key = row.lote||row._id;
    if(!map.has(key)) map.set(key,{lote:key,rows:[],tratativa:row.tratativa||"",
      pedido:row.pedido||"",item:row.item||"",material:row.material||"",
      descricao:row.descricao||"",data_bloqueio:row.data_bloqueio||"",
      ultima_ordem:row.ultima_ordem||"",qualidade_qts:row.qualidade_qts||"",
      deposito_sap:row.deposito_sap||"",motivo_bloqueio:row.motivo_bloqueio||"",
      motivo_bloqueio_texto:row.motivo_bloqueio_texto||"",razao_bloq:row.razao_bloq||"",
      descricao_motivo:row.descricao_motivo||"",num_cassete:row.num_cassete||"",
    });
    map.get(key).rows.push(row);
  });
  return Array.from(map.values());
}

function parseTokens(str) {
  if(!str||!str.trim()) return [];
  return str.split(/[\s,;]+/).map(s=>s.trim().toLowerCase()).filter(Boolean);
}

function rowMatchesSearch(row, filters) {
  const tPedidoItem = parseTokens(filters.pedido_item);
  const tLotes      = parseTokens(filters.lotes);
  const tIppns      = parseTokens(filters.ippns);
  const tDeposito   = parseTokens(filters.deposito);
  const match = (tokens, fields) => !tokens.length || tokens.some(t=>fields.some(f=>String(row[f]||"").toLowerCase().includes(t)));
  return match(tPedidoItem,["pedido","item"]) && match(tLotes,["lote"]) && match(tIppns,["ippn"]) && match(tDeposito,["deposito_sap"]);
}

function initStageData() { return {1:[],2:[],3:[],4:[],5:[],6:[],7:[],8:[]}; }

// Format a duration in days (unit = "dias")
function formatDuration(ms) {
  if(!ms||ms<=0) return "—";
  const days = ms/86400000;
  const val = days<10 ? days.toFixed(1) : Math.round(days);
  const num = Number(val);
  return `${val} ${num===1?"dia":"dias"}`;
}

// Compute average time spent at each stage from history
function computeTimingStats(stageData, historyRows) {
  const allRows = [...Object.values(stageData).flat(), ...historyRows];
  const stageMs = {}; // stageId -> [ms]
  const deptMs  = {}; // dept    -> [ms]

  allRows.forEach(row=>{
    const hist = row.history||[];
    let prevMs = row.stageEnteredMs||null;
    hist.forEach(h=>{
      if(!h.dateMs||!prevMs) { if(h.dateMs) prevMs=h.dateMs; return; }
      const dur = h.dateMs - prevMs;
      if(dur>0 && dur < 90*24*3600000) {
        const sid = h.stage;
        if(!stageMs[sid]) stageMs[sid]=[];
        stageMs[sid].push(dur);
        const dept = STAGES.find(s=>s.id===sid)?.dept;
        if(dept){
          if(!deptMs[dept]) deptMs[dept]=[];
          deptMs[dept].push(dur);
        }
      }
      prevMs = h.dateMs;
    });
  });

  const avg = arr => arr.length ? arr.reduce((a,b)=>a+b,0)/arr.length : null;
  const stageAvg = {};
  Object.entries(stageMs).forEach(([id,arr])=>{ stageAvg[Number(id)]=avg(arr); });
  const deptAvg = {};
  Object.entries(deptMs).forEach(([d,arr])=>{ deptAvg[d]=avg(arr); });
  return { stageAvg, deptAvg };
}

// Render a recurso/tratativa ("Definição") value — handles array (Etapa 1) or string
function renderTratativaValue(v){
  if(Array.isArray(v)) return v.length?v.join(", "):"—";
  return v||"—";
}

// Convert "#RRGGBB" to "rgba(r,g,b,a)" — used for soft tinted icon badges
function hexToRgba(hex, alpha){
  const h = hex.replace("#","");
  const r = parseInt(h.substring(0,2),16);
  const g = parseInt(h.substring(2,4),16);
  const b = parseInt(h.substring(4,6),16);
  return `rgba(${r},${g},${b},${alpha})`;
}

// ─────────────────────────────────────────────────────────────────────────────
// GLOBAL STYLES — spring press feedback, hover lift, focus rings
// ─────────────────────────────────────────────────────────────────────────────
function GlobalStyles(){
  return (
    <style>{`
      * { -webkit-tap-highlight-color: transparent; }
      .spring-btn { transition: transform 0.22s ${SPRING}, opacity 0.15s ${EASE}, background 0.15s ${EASE}, box-shadow 0.15s ${EASE}; }
      .spring-btn:active { transform: scale(0.96); }
      .hover-lift { transition: transform 0.2s ${EASE}, box-shadow 0.2s ${EASE}; }
      .hover-lift:hover { transform: translateY(-2px); box-shadow: 0 8px 22px rgba(0,0,0,0.08); }
      input:focus, select:focus, textarea:focus { border-color: ${ACCENT} !important; box-shadow: 0 0 0 3px rgba(10,132,255,0.15) !important; }
      ::-webkit-scrollbar { width: 7px; height: 7px; }
      ::-webkit-scrollbar-track { background: transparent; }
      ::-webkit-scrollbar-thumb { background: rgba(120,120,128,0.28); border-radius: 4px; }
      ::-webkit-scrollbar-thumb:hover { background: rgba(120,120,128,0.45); }
      details summary::-webkit-details-marker { display: none; }
      details summary { list-style: none; cursor: pointer; }
    `}</style>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// SMALL UI ATOMS
// ─────────────────────────────────────────────────────────────────────────────
function Toast({msg}){ if(!msg)return null; return(<div style={{position:"fixed",bottom:28,left:`calc(50% + ${SIDEBAR_W/2}px)`,transform:"translateX(-50%)",background:"rgba(28,28,30,0.92)",backdropFilter:"blur(12px)",color:"#fff",borderRadius:14,padding:"12px 22px",fontSize:14,fontWeight:500,zIndex:500,boxShadow:"0 10px 30px rgba(0,0,0,0.25)",display:"flex",alignItems:"center",gap:8,whiteSpace:"nowrap",pointerEvents:"none",fontFamily:FONT}}><CheckCircle2 size={16} color="#34C759"/>{msg}</div>); }

function Modal({title,body,onConfirm,onCancel,confirmLabel="Confirmar",danger=false,children}){
  return(<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.4)",backdropFilter:"blur(2px)",zIndex:400,display:"flex",alignItems:"center",justifyContent:"center",padding:20}} onClick={onCancel}><div style={{background:SURFACE,borderRadius:18,padding:"28px 28px 24px",maxWidth:480,width:"100%",boxShadow:"0 28px 70px rgba(0,0,0,0.25)",fontFamily:FONT}} onClick={e=>e.stopPropagation()}><div style={{fontSize:19,fontWeight:800,color:"#1C1C1E",marginBottom:8,fontFamily:TITLE_FONT,letterSpacing:"-0.3px"}}>{title}</div>{body&&<div style={{fontSize:14,color:"#3A3A3C",marginBottom:children?12:22,lineHeight:1.65}}>{body}</div>}{children&&<div style={{marginBottom:20}}>{children}</div>}<div style={{display:"flex",gap:10,justifyContent:"flex-end"}}><button className="spring-btn" style={{background:"#F2F2F7",color:ACCENT,border:"none",borderRadius:12,padding:"11px 22px",fontSize:14,fontWeight:600,cursor:"pointer"}} onClick={onCancel}>Cancelar</button><button className="spring-btn" style={{background:danger?"linear-gradient(135deg,#FF453A,#C0392B)":`linear-gradient(135deg,${ACCENT},#0051D4)`,color:"#fff",border:"none",borderRadius:12,padding:"11px 22px",fontSize:14,fontWeight:600,cursor:"pointer"}} onClick={onConfirm}>{confirmLabel}</button></div></div></div>);
}

function DeptTag({dept}){
  const c=DEPT_COLORS[dept]||{bg:"#F2F2F7",fg:"#3A3A3C"};
  return <span style={{background:c.bg,color:c.fg,borderRadius:6,padding:"2px 9px",fontSize:11,fontWeight:600,display:"inline-block"}}>{DEPT_FULL[dept]||dept}</span>;
}

function Btn({children,variant="primary",disabled,onClick,small,icon:Icon,style={}}){
  const variants={
    primary:{background:`linear-gradient(135deg,${ACCENT},#0060DF)`,color:"#fff"},
    secondary:{background:"#F2F2F7",color:ACCENT},
    danger:{background:"linear-gradient(135deg,#FF453A,#C0392B)",color:"#fff"},
    warning:{background:"linear-gradient(135deg,#FF9F0A,#E8890A)",color:"#fff"},
    ghost:{background:"transparent",color:ACCENT},
  };
  return(<button className="spring-btn" disabled={disabled} onClick={onClick} style={{border:"none",borderRadius:10,padding:small?"7px 14px":"10px 18px",fontSize:small?12:13,fontWeight:600,cursor:disabled?"not-allowed":"pointer",opacity:disabled?0.4:1,display:"inline-flex",alignItems:"center",gap:6,fontFamily:FONT,...variants[variant],...style}}>{Icon&&<Icon size={small?13:15} strokeWidth={2.4}/>}{children}</button>);
}

const INP = {width:"100%",border:"1.5px solid #E5E5EA",borderRadius:9,padding:"9px 12px",fontSize:13,outline:"none",background:"#fff",boxSizing:"border-box",fontFamily:FONT};
const TH = {background:"#F6F6F8",padding:"10px 11px",textAlign:"left",fontWeight:700,color:"#6B6B70",borderBottom:`1px solid ${SEPARATOR}`,whiteSpace:"nowrap",fontSize:10.5,textTransform:"uppercase",letterSpacing:"0.4px",position:"sticky",top:0,zIndex:10};
const TH_ACCENT = {...TH, background:"rgba(10,132,255,0.08)", color:ACCENT};
const TD = (alt)=>({padding:"10px 11px",borderBottom:`1px solid ${SEPARATOR}`,verticalAlign:"middle",fontSize:13,color:"#1C1C1E",background:alt?"#FAFAFB":"#fff"});

// Inline history record list — stages separated by a vertical bar "|", clean look (no cards)
function HistoryInline({history}){
  if(!history || history.length===0) return <span style={{color:"#C7C7CC",fontSize:11}}>—</span>;
  return(
    <div style={{display:"flex",alignItems:"center",flexWrap:"wrap",gap:0,fontSize:11,lineHeight:1.6,fontFamily:FONT}}>
      {history.map((h,hi)=>(
        <span key={hi} style={{display:"inline-flex",alignItems:"baseline",whiteSpace:"nowrap"}}>
          {hi>0&&<span style={{color:"#D1D1D6",margin:"0 10px",fontWeight:300}}>|</span>}
          <span style={{fontWeight:700,color:"#1C1C1E"}}>{h.stage}. {h.stageLabel}:</span>
          <span style={{color:ACCENT,fontWeight:600,marginLeft:5}}>{renderTratativaValue(h.tratativa)}</span>
          <span style={{color:"#8E8E93",marginLeft:6}}>— {h.user}</span>
          <span style={{color:"#C7C7CC",marginLeft:5}}>({h.date})</span>
        </span>
      ))}
    </div>
  );
}

// Multi-select "dropdown" for Etapa 1 — uses a portal so it isn't clipped by the scroll container
function RecursosSelect({value,onChange,readOnly}){
  const [open,setOpen]=useState(false);
  const [pos,setPos]=useState({top:0,left:0,width:200});
  const btnRef=useRef(null);
  const panelRef=useRef(null);
  const arr=Array.isArray(value)?value:(value?[value]:[]);

  useEffect(()=>{
    function h(e){
      if(btnRef.current?.contains(e.target))return;
      if(panelRef.current?.contains(e.target))return;
      setOpen(false);
    }
    document.addEventListener("mousedown",h);
    return ()=>document.removeEventListener("mousedown",h);
  },[]);

  if(readOnly) return <span style={{fontSize:11,color:"#555"}}>{arr.length?arr.join(", "):"—"}</span>;

  function openPanel(){
    const r=btnRef.current.getBoundingClientRect();
    setPos({top:r.bottom+4,left:r.left,width:Math.max(r.width,210)});
    setOpen(true);
  }
  function toggle(opt){
    onChange(arr.includes(opt)?arr.filter(o=>o!==opt):[...arr,opt]);
  }

  return(
    <>
      <button ref={btnRef} type="button" className="spring-btn" onClick={()=>open?setOpen(false):openPanel()} style={{width:"100%",minWidth:170,textAlign:"left",border:"1.5px solid #E5E5EA",borderRadius:8,padding:"5px 8px",fontSize:11,background:"#fff",cursor:"pointer",color:arr.length?"#1C1C1E":"#8E8E93",display:"flex",justifyContent:"space-between",gap:4,alignItems:"center",fontFamily:FONT}}>
        <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{arr.length?arr.join(", "):"Selecionar…"}</span>
        <ChevronDown size={12} color="#8E8E93"/>
      </button>
      {open&&createPortal(
        <div ref={panelRef} style={{position:"fixed",top:pos.top,left:pos.left,minWidth:pos.width,background:"#fff",border:`1px solid ${SEPARATOR}`,borderRadius:12,boxShadow:"0 10px 30px rgba(0,0,0,0.18)",zIndex:1000,padding:5,fontFamily:FONT}}>
          {RECURSOS_OPTIONS.map(opt=>{
            const sel=arr.includes(opt);
            return(
              <label key={opt} className="spring-btn" style={{display:"flex",alignItems:"center",gap:8,padding:"7px 9px",fontSize:12,cursor:"pointer",borderRadius:8,background:sel?"rgba(10,132,255,0.08)":"transparent",fontWeight:sel?600:400,color:"#1C1C1E"}}>
                <input type="checkbox" checked={sel} onChange={()=>toggle(opt)} style={{accentColor:ACCENT,width:14,height:14,cursor:"pointer"}}/>
                {opt}
              </label>
            );
          })}
        </div>,
        document.body
      )}
    </>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// LOGIN
// ─────────────────────────────────────────────────────────────────────────────
function LoginPage({onLogin,users}){
  const [email,setEmail]=useState("");
  const [pw,setPw]=useState("");
  const [err,setErr]=useState("");
  const [loading,setLoading]=useState(false);
  function go(){
    if(!email||!pw){setErr("Preencha e-mail e senha.");return;}
    setLoading(true);
    setTimeout(()=>{
      const u=users.find(u=>u.email===email.trim().toLowerCase()&&u.password===pw&&u.active!==false);
      if(u){setErr("");onLogin(u);}else setErr("E-mail ou senha inválidos.");
      setLoading(false);
    },350);
  }
  return(
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"linear-gradient(160deg,#0A2240 0%,#1A3A5C 55%,#0A84FF 130%)",fontFamily:FONT}}>
      <GlobalStyles/>
      <div style={{background:"rgba(255,255,255,0.96)",backdropFilter:"blur(20px)",borderRadius:24,padding:"42px 38px",width:380,boxShadow:"0 30px 70px rgba(0,0,0,0.35)"}}>
        <div style={{textAlign:"center",marginBottom:30}}>
          <div style={{width:60,height:60,borderRadius:18,margin:"0 auto 16px",background:`linear-gradient(135deg,${ACCENT},#0051D4)`,display:"flex",alignItems:"center",justifyContent:"center",boxShadow:"0 8px 22px rgba(10,132,255,0.35)"}}><FlaskConical size={28} color="#fff" strokeWidth={2}/></div>
          <div style={{fontSize:25,fontWeight:800,color:"#0A2240",letterSpacing:"-0.5px"}}>Fluxo de Liberação</div>
          <div style={{fontSize:13,color:"#8E8E93",marginTop:4}}>Gestão de Produtos — Controle da Qualidade</div>
        </div>
        {err&&<div style={{color:"#FF453A",fontSize:13,textAlign:"center",background:"#FFF2F0",borderRadius:10,padding:"8px 12px",marginBottom:12}}>{err}</div>}
        <div style={{marginBottom:14}}><label style={{fontSize:13,fontWeight:600,color:"#3A3A3C",marginBottom:6,display:"block"}}>E-mail</label><input style={INP} type="email" placeholder="seu@empresa.com" value={email} onChange={e=>setEmail(e.target.value)} onKeyDown={e=>e.key==="Enter"&&go()} autoFocus/></div>
        <div style={{marginBottom:22}}><label style={{fontSize:13,fontWeight:600,color:"#3A3A3C",marginBottom:6,display:"block"}}>Senha</label><input style={INP} type="password" placeholder="••••••••" value={pw} onChange={e=>setPw(e.target.value)} onKeyDown={e=>e.key==="Enter"&&go()}/></div>
        <button className="spring-btn" style={{width:"100%",background:`linear-gradient(135deg,${ACCENT},#0051D4)`,color:"#fff",border:"none",borderRadius:14,padding:"14px",fontSize:16,fontWeight:600,cursor:"pointer",opacity:loading?0.7:1,fontFamily:FONT}} onClick={go} disabled={loading}>{loading?"Entrando…":"Entrar"}</button>
        <div style={{marginTop:20,fontSize:11,color:"#C7C7CC",textAlign:"center",lineHeight:1.7,borderTop:`1px solid ${SEPARATOR}`,paddingTop:14}}>Demo: qualidade.tecnica@empresa.com / 123456<br/>admin@empresa.com / admin</div>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// SIDEBAR (desktop NavigationSplitView pattern — appears on hover, hides otherwise)
// ─────────────────────────────────────────────────────────────────────────────
function Sidebar({page,setPage}){
  const [open,setOpen]=useState(false);
  const items=[
    {key:"dashboard",icon:LayoutGrid,label:"Dashboard"},
    {key:"pipeline",icon:GitBranch,label:"Pipeline"},
    {key:"historico",icon:HistoryIcon,label:"Histórico"},
    {key:"configuracoes",icon:Settings,label:"Configurações"},
  ];
  return(
    <>
      {/* Edge hover zone — hovering near the left edge reveals the sidebar */}
      <div onMouseEnter={()=>setOpen(true)} style={{position:"fixed",left:0,top:0,width:14,height:"100vh",zIndex:99}}/>

      {/* Persistent handle — signals the hidden sidebar is there, even when collapsed */}
      <div
        onMouseEnter={()=>setOpen(true)}
        style={{
          position:"fixed", left:0, top:"50%", zIndex:98,
          width:22, height:68, borderRadius:"0 16px 16px 0",
          background:SIDEBAR_GLASS, backdropFilter:"blur(16px) saturate(180%)", WebkitBackdropFilter:"blur(16px) saturate(180%)",
          border:`1px solid ${SEPARATOR}`, borderLeft:"none",
          display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", gap:4,
          boxShadow:"2px 0 12px rgba(0,0,0,0.07)", cursor:"pointer",
          transform:open?"translate(-28px,-50%)":"translate(0,-50%)",
          opacity:open?0:1,
          transition:`transform 0.32s ${SPRING}, opacity 0.25s ${EASE}`,
        }}
      >
        <ChevronRight size={13} color={ACCENT} strokeWidth={2.6}/>
        <div style={{display:"flex",flexDirection:"column",gap:3}}>
          <span style={{width:3,height:3,borderRadius:"50%",background:"#C7C7CC"}}/>
          <span style={{width:3,height:3,borderRadius:"50%",background:"#C7C7CC"}}/>
          <span style={{width:3,height:3,borderRadius:"50%",background:"#C7C7CC"}}/>
        </div>
      </div>

      <div
        onMouseEnter={()=>setOpen(true)}
        onMouseLeave={()=>setOpen(false)}
        style={{
          position:"fixed", left:0, top:0, width:SIDEBAR_W, height:"100vh", zIndex:100,
          background:SIDEBAR_GLASS, backdropFilter:"blur(24px) saturate(180%)", WebkitBackdropFilter:"blur(24px) saturate(180%)",
          borderRight:`1px solid ${SEPARATOR}`, display:"flex", flexDirection:"column",
          padding:"18px 12px", boxSizing:"border-box", fontFamily:FONT,
          transform:open?"translateX(0)":"translateX(-100%)",
          boxShadow:open?"6px 0 30px rgba(0,0,0,0.12)":"none",
          transition:`transform 0.32s ${SPRING}, box-shadow 0.32s ${EASE}`,
        }}
      >
        <div style={{display:"flex",alignItems:"center",gap:10,padding:"4px 8px",marginBottom:24}}>
          <div style={{width:34,height:34,borderRadius:9,background:`linear-gradient(135deg,${ACCENT},#0051D4)`,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,boxShadow:"0 4px 12px rgba(10,132,255,0.3)"}}>
            <FlaskConical size={18} color="#fff" strokeWidth={2.2}/>
          </div>
          <div style={{fontSize:15,fontWeight:800,color:"#1C1C1E",letterSpacing:"-0.3px",lineHeight:1.25,fontFamily:TITLE_FONT}}>Fluxo de<br/>Liberação</div>
        </div>

        <nav style={{display:"flex",flexDirection:"column",gap:2}}>
          {items.map(it=>{
            const active=page===it.key;
            const Icon=it.icon;
            return(
              <button key={it.key} className="spring-btn" onClick={()=>setPage(it.key)} style={{
                display:"flex",alignItems:"center",gap:10,padding:"9px 12px",borderRadius:10,border:"none",
                background:active?"rgba(10,132,255,0.12)":"transparent",
                color:active?ACCENT:"#3A3A3C", fontSize:13.5,fontWeight:active?700:500,
                cursor:"pointer",textAlign:"left",width:"100%",fontFamily:FONT,
              }}>
                <Icon size={17} strokeWidth={2.2}/>
                {it.label}
              </button>
            );
          })}
        </nav>
      </div>
    </>
  );
}

// Glass top bar — large title for the active page + logged-in user (top-right)
function TopBar({title,subtitle,user,onLogoutClick}){
  return(
    <div style={{
      position:"sticky", top:0, zIndex:40,
      background:HEADER_GLASS, backdropFilter:"blur(20px) saturate(180%)", WebkitBackdropFilter:"blur(20px) saturate(180%)",
      borderBottom:`1px solid ${SEPARATOR}`, padding:"16px 28px", fontFamily:FONT,
      display:"flex", alignItems:"center", justifyContent:"space-between", gap:16,
    }}>
      <div>
        <div style={{fontSize:26,fontWeight:800,color:"#1C1C1E",letterSpacing:"-0.5px",fontFamily:TITLE_FONT}}>{title}</div>
        {subtitle&&<div style={{fontSize:12.5,color:"#8E8E93",marginTop:2}}>{subtitle}</div>}
      </div>
      <div style={{display:"flex",alignItems:"center",gap:10,flexShrink:0}}>
        <div style={{textAlign:"right"}}>
          <div style={{fontSize:13,fontWeight:700,color:"#1C1C1E",lineHeight:1.3}}>{user.name}</div>
          <div style={{fontSize:11,color:"#8E8E93",lineHeight:1.3}}>{DEPT_FULL[user.dept]||user.dept}</div>
        </div>
        <div style={{width:34,height:34,borderRadius:"50%",background:`linear-gradient(135deg,${ACCENT},#5AC8FA)`,display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontSize:13,fontWeight:700,flexShrink:0}}>{user.name.charAt(0)}</div>
        <button className="spring-btn" onClick={onLogoutClick} title="Sair" style={{background:"#F2F2F7",border:"none",cursor:"pointer",color:"#8E8E93",padding:8,borderRadius:9,display:"flex"}}>
          <LogOut size={16} strokeWidth={2.2}/>
        </button>
      </div>
    </div>

  );
}

// ─────────────────────────────────────────────────────────────────────────────
// PIPELINE BAR
// ─────────────────────────────────────────────────────────────────────────────
function PipelineBar({stageData,activeStage,onSelectStage,searchResult}){
  return(
    <div style={{background:SURFACE,borderBottom:`1px solid ${SEPARATOR}`}}>
      <div style={{...CONTAINER,overflowX:"auto"}}>
        <div style={{display:"flex",alignItems:"stretch",minWidth:850}}>
          {STAGES.map((stage,idx)=>{
            const active=activeStage===stage.id;
            const count=searchResult ? (searchResult[stage.id]||0) : (stageData[stage.id]?.length||0);
            const showBadge=count>0;
            return(
              <div key={stage.id} style={{display:"flex",alignItems:"stretch",flex:1}}>
                <button className="spring-btn" onClick={()=>onSelectStage(stage.id)} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"flex-start",cursor:"pointer",padding:"13px 4px 10px",position:"relative",border:"none",borderBottom:active?`3px solid ${ACCENT}`:"3px solid transparent",background:active?"rgba(10,132,255,0.05)":"transparent",minWidth:88,fontFamily:FONT}}>
                  {showBadge&&<div style={{position:"absolute",top:6,right:"calc(50% - 22px)",background:searchResult?"#FF9F0A":"#FF453A",color:"#fff",borderRadius:10,fontSize:9,fontWeight:700,padding:"1px 5px",lineHeight:"14px",height:14,minWidth:16,textAlign:"center"}}>{count}</div>}
                  <div style={{width:28,height:28,borderRadius:"50%",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700,marginBottom:5,background:active?ACCENT:"#EDEDF0",color:active?"#fff":"#8E8E93",boxShadow:active?"0 0 0 4px rgba(10,132,255,0.14)":"none",transition:`all 0.3s ${SPRING}`,flexShrink:0}}>{stage.id}</div>
                  <div style={{fontSize:9,fontWeight:active?700:500,color:active?ACCENT:"#6B6B6B",textAlign:"center",lineHeight:1.3,marginBottom:3}}>{stage.short}</div>
                  {stage.dept&&<div style={{fontSize:8,color:active?"rgba(10,132,255,0.7)":"#ABABAB",textAlign:"center",lineHeight:1.3,fontWeight:500,maxWidth:84}}>{DEPT_FULL[stage.dept]||stage.dept}</div>}
                </button>
                {idx<STAGES.length-1&&<div style={{width:6,display:"flex",alignItems:"center",marginBottom:3}}><div style={{width:"100%",height:2,background:SEPARATOR}}/></div>}
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// SEARCH BAR
// ─────────────────────────────────────────────────────────────────────────────
function SearchBar({filters,onChange,onClear}){
  const fields=[
    {key:"pedido_item",label:"Pedido / Item",ph:"ex: 4500001 10"},
    {key:"lotes",label:"Lotes",ph:"ex: LOTE-001; LOTE-002"},
    {key:"ippns",label:"IPPNs",ph:"ex: IPN001 IPN002"},
    {key:"deposito",label:"Depósito SAP",ph:"ex: DEP-01"},
  ];
  const hasAny=fields.some(f=>filters[f.key]);
  return(
    <div style={{background:SURFACE,borderRadius:14,padding:"14px 16px",boxShadow:"0 1px 4px rgba(0,0,0,0.04)",marginBottom:16}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10}}>
        <span style={{fontSize:12.5,fontWeight:700,color:"#3A3A3C",display:"flex",alignItems:"center",gap:6}}><Search size={14} color={ACCENT}/> Busca</span>
        {hasAny&&<button className="spring-btn" onClick={onClear} style={{background:"none",border:"none",color:"#FF453A",fontSize:12,fontWeight:600,cursor:"pointer"}}>Limpar</button>}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(210px,1fr))",gap:10}}>
        {fields.map(f=>(
          <div key={f.key}>
            <div style={{fontSize:10,fontWeight:600,color:"#8E8E93",marginBottom:3}}>{f.label}</div>
            <input value={filters[f.key]||""} onChange={e=>onChange({...filters,[f.key]:e.target.value})} placeholder={f.ph} style={{width:"100%",border:"1.5px solid #E5E5EA",borderRadius:9,padding:"7px 10px",fontSize:12.5,outline:"none",background:"#FAFAFB",boxSizing:"border-box",fontFamily:FONT}}/>
          </div>
        ))}
      </div>
      {hasAny&&<div style={{marginTop:8,fontSize:10,color:"#8E8E93"}}>Separe múltiplos valores com espaço, vírgula ou ponto-e-vírgula</div>}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// IMPORT STEP
// ─────────────────────────────────────────────────────────────────────────────
function ImportStep({onImport}){
  const [drag,setDrag]=useState(false);
  const [preview,setPreview]=useState(null);
  const [error,setError]=useState("");
  const [loading,setLoading]=useState(false);

  async function processFile(file){
    if(!file)return;
    setLoading(true);
    try{
      const rows=await parseFile(file);
      if(!rows.length){setError("Nenhuma linha válida encontrada.");setLoading(false);return;}
      setPreview(rows);setError("");
    }catch(e){setError("Erro ao processar arquivo: "+e.message);}
    setLoading(false);
  }
  const handleDrop=useCallback(e=>{e.preventDefault();setDrag(false);processFile(e.dataTransfer.files[0]);},[]);

  const previewCols=["pedido_item","material","lote","ippn","deposito_sap","motivo_bloqueio","data_bloqueio"];
  function previewCell(row,col){
    if(col==="pedido_item") return (row.pedido||row.item)?`${row.pedido||"—"}/${row.item||"—"}`:"—";
    return row[col]||"—";
  }
  function previewLabel(col){
    return col==="pedido_item" ? "Pedido/Item" : (COL_LABELS[col]||col);
  }
  return(
    <div>
      <div style={{fontSize:13,color:"#8E8E93",marginBottom:20}}>Carregue um arquivo CSV, TXT ou XLS/XLSX para iniciar o fluxo na Etapa 1.</div>
      <div style={{border:`2px dashed ${drag?ACCENT:"#D1D1D6"}`,borderRadius:18,padding:"44px 24px",textAlign:"center",background:drag?"rgba(10,132,255,0.04)":"#FAFAFB",cursor:"pointer",transition:`all 0.25s ${EASE}`,marginBottom:18}} onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)} onDrop={handleDrop} onClick={()=>!loading&&document.getElementById("fileInputImp").click()}>
        <div style={{display:"flex",justifyContent:"center",marginBottom:12}}>
          {loading?<Loader2 size={34} color={ACCENT} className="spin" style={{animation:"spin 1s linear infinite"}}/>:<Upload size={34} color={ACCENT} strokeWidth={1.8}/>}
        </div>
        <div style={{fontSize:14,fontWeight:700,color:"#3A3A3C",marginBottom:3}}>{loading?"Processando…":"Arraste o arquivo ou clique para selecionar"}</div>
        <div style={{fontSize:11,color:"#8E8E93"}}>CSV · TXT · XLS · XLSX</div>
        <input id="fileInputImp" type="file" accept=".csv,.txt,.xls,.xlsx" style={{display:"none"}} onChange={e=>processFile(e.target.files[0])}/>
      </div>
      {error&&<div style={{color:"#FF453A",fontSize:13,background:"#FFF2F0",borderRadius:10,padding:"8px 12px",marginBottom:12}}>{error}</div>}
      {preview?(
        <div>
          <div style={{fontSize:13,color:"#34C759",fontWeight:700,marginBottom:10,display:"flex",alignItems:"center",gap:6}}><CheckCircle2 size={15}/> {preview.length} tubos encontrados</div>
          <div style={{overflowX:"auto",borderRadius:14,boxShadow:"0 1px 5px rgba(0,0,0,0.06)",background:"#fff",maxHeight:260,overflowY:"auto",marginBottom:14}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:600}}>
              <thead><tr>{previewCols.map(c=><th key={c} style={TH}>{previewLabel(c)}</th>)}</tr></thead>
              <tbody>{preview.slice(0,10).map((row,i)=><tr key={i}>{previewCols.map(c=><td key={c} style={TD(i%2===1)}>{previewCell(row,c)}</td>)}</tr>)}</tbody>
            </table>
          </div>
          {preview.length>10&&<div style={{fontSize:11,color:"#8E8E93",marginBottom:10}}>Exibindo 10 de {preview.length} linhas</div>}
          <div style={{display:"flex",gap:10}}>
            <Btn onClick={()=>onImport(preview)} icon={Upload}>Importar {preview.length} Tubos</Btn>
            <Btn variant="secondary" onClick={()=>setPreview(null)}>Cancelar</Btn>
          </div>
        </div>
      ):(
        <div style={{background:"#fff",borderRadius:14,padding:"14px 16px",boxShadow:"0 1px 4px rgba(0,0,0,0.04)"}}>
          <div style={{fontSize:12,fontWeight:700,color:"#3A3A3C",marginBottom:6}}>Colunas esperadas (mapeamento automático por nome):</div>
          <div style={{fontSize:11,color:"#8E8E93",lineHeight:2}}>{Object.values(COL_LABELS).join(" · ")}</div>
          <div style={{fontSize:11,color:"#8E8E93",marginTop:6}}>A coluna <strong>Data Bloqueio</strong> é automaticamente ajustada para o formato DD/MM/AAAA. As colunas <strong>Pedido</strong> e <strong>Item</strong> são exibidas em conjunto como <strong>Pedido/Item</strong>.</div>
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// STAGE VIEW
// ─────────────────────────────────────────────────────────────────────────────
function StageView({stage,rows,user,onAdvance,onReturn,onComplete,filters,faturamento}){
  const [selLotes,setSelLotes]=useState(new Set());
  const [tratativas,setTratativas]=useState({});
  const [expanded,setExpanded]=useState(new Set());
  const [confirmAdv,setConfirmAdv]=useState(false);
  const [confirmRet,setConfirmRet]=useState(false);
  const [confirmDone,setConfirmDone]=useState(false);

  const isStage8 = stage.id===8;
  const isStage1 = stage.id===1;
  const nextStage = STAGES.find(s=>s.id===stage.id+1);
  const prevStage = STAGES.find(s=>s.id===stage.id-1);

  const hasSearch = filters && Object.values(filters).some(Boolean);
  const allGroups = groupByLote(rows);
  // Search also filters the table itself (not just the pipeline badge counts)
  const groups = hasSearch ? allGroups.filter(g=>g.rows.some(r=>rowMatchesSearch(r,filters))) : allGroups;

  // Pedido/Item combos registered in Faturamento — highlighted in the table
  const fatSet = new Set((faturamento||[]).map(f=>`${(f.pedido||"").trim()}|${(f.item||"").trim()}`.toLowerCase()));
  const isFaturado = (g)=>fatSet.has(`${(g.pedido||"").trim()}|${(g.item||"").trim()}`.toLowerCase());

  const canInteract = user.dept==="Admin" || (user.allowedStages||[]).includes(stage.id);

  function toggleSel(lote){const n=new Set(selLotes);n.has(lote)?n.delete(lote):n.add(lote);setSelLotes(n);}
  function toggleAll(c){setSelLotes(c?new Set(groups.map(g=>g.lote)):new Set());}
  function toggleExp(lote){const n=new Set(expanded);n.has(lote)?n.delete(lote):n.add(lote);setExpanded(n);}

  function buildMoved(loteSet){
    return rows.filter(r=>loteSet.has(r.lote||r._id)).map(r=>{
      const entered = tratativas[r.lote||r._id] ?? r.tratativa;
      return {
        ...r,
        tratativa:"", // cleared on advance/return — history keeps the record
        history:[...(r.history||[]),{
          stage:stage.id,stageLabel:stage.label,
          user:user.name,dept:user.dept,
          tratativa:entered,
          date:new Date().toLocaleString("pt-BR"),
          dateMs:Date.now(),
        }],
        stageEnteredMs:Date.now(),
      };
    });
  }

  function doAdvance(){const moved=buildMoved(selLotes);const rem=rows.filter(r=>!selLotes.has(r.lote||r._id));onAdvance(moved,rem);setSelLotes(new Set());setTratativas({});setConfirmAdv(false);}
  function doReturn(){const moved=buildMoved(selLotes);const rem=rows.filter(r=>!selLotes.has(r.lote||r._id));onReturn(moved,rem);setSelLotes(new Set());setTratativas({});setConfirmRet(false);}
  function doComplete(){const moved=buildMoved(selLotes);const rem=rows.filter(r=>!selLotes.has(r.lote||r._id));onComplete(moved,rem);setSelLotes(new Set());setTratativas({});setConfirmDone(false);}

  const selCount=rows.filter(r=>selLotes.has(r.lote||r._id)).length;
  const selLotesCount=selLotes.size;

  if(rows.length===0) return(
    <div>
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14}}><DeptTag dept={stage.dept}/></div>
      <div style={{textAlign:"center",padding:"60px 24px",background:"#fff",borderRadius:18,boxShadow:"0 1px 4px rgba(0,0,0,0.04)"}}>
        <div style={{display:"flex",justifyContent:"center",marginBottom:12}}><Inbox size={40} color="#C7C7CC" strokeWidth={1.5}/></div>
        <div style={{fontSize:15,fontWeight:700,color:"#3A3A3C",marginBottom:4}}>Nenhum tubo nesta etapa</div>
        <div style={{fontSize:13,color:"#8E8E93"}}>Os tubos aparecerão aqui quando avançarem para esta etapa.</div>
      </div>
    </div>
  );

  // Sticky checkbox column (left-fixed)
  const stickyTh = {...TH, width:34, position:"sticky", left:0, zIndex:25};
  const stickyTd = (bg)=>({...TD(false), background:bg, textAlign:"center", position:"sticky", left:0, zIndex:5});
  const histTh = {...TH, minWidth:380};
  const histTd = (alt)=>({...TD(alt), minWidth:380, padding:"10px 14px"});
  const descTh = {...TH, minWidth:220};
  const descTd = (bg)=>({...TD(false), background:bg, minWidth:220});

  // total column count (for IPPN expand row colSpan)
  const totalCols = (canInteract?1:0) /*checkbox*/ + 1 /*expand*/ + 3 /*Lote,Tubos,IPPNs*/ + 2 /*DataBloqueio,Cassete*/ + 3 /*Pedido/Item,Material,Descricao*/ + 4 /*UltimaOrdem,QTS,Deposito,Motivo*/ + 3 /*MotivoTexto,RazaoBloq,DescricaoMotivo*/ + 1 /*Definição*/ + (stage.id>1?1:0) /*Historico*/;

  return(
    <div>
      {/* Top bar */}
      <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:14,flexWrap:"wrap",gap:10}}>
        <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
          <DeptTag dept={stage.dept}/>
          <span style={{fontSize:12,color:"#8E8E93"}}>{groups.length} lote{groups.length!==1?"s":""} · {rows.length} tubo{rows.length!==1?"s":""}</span>
          {selLotes.size>0&&<span style={{fontSize:12,fontWeight:700,color:ACCENT,background:"rgba(10,132,255,0.08)",borderRadius:8,padding:"3px 10px"}}>{selLotesCount} lote{selLotesCount!==1?"s":""} ({selCount} tubo{selCount!==1?"s":""})</span>}
          {!canInteract&&<span style={{fontSize:11,color:"#FF453A",background:"#FFF2F0",borderRadius:6,padding:"2px 8px",fontWeight:600,display:"flex",alignItems:"center",gap:4}}><AlertTriangle size={11}/> Sem permissão nesta etapa</span>}
          {faturamento&&faturamento.length>0&&(
            <span style={{display:"inline-flex",alignItems:"center",gap:6,background:"rgba(255,69,58,0.10)",borderRadius:6,padding:"2px 9px",fontSize:11,fontWeight:600,color:"#1C1C1E"}}>
              <span style={{width:9,height:9,borderRadius:3,background:"rgba(255,69,58,0.35)",display:"inline-block"}}/>
              Faturamento
            </span>
          )}
        </div>
        {/* Action buttons */}
        {canInteract&&(
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
            {stage.id>1&&onReturn&&<Btn variant="warning" icon={ArrowLeft} disabled={selLotes.size===0} onClick={()=>setConfirmRet(true)}>Retornar: {prevStage?.short}</Btn>}
            {!isStage8&&<Btn icon={ArrowRight} disabled={selLotes.size===0} onClick={()=>setConfirmAdv(true)}>{nextStage?.short} ({selLotesCount})</Btn>}
            {isStage8&&<Btn variant="danger" icon={CheckCircle2} disabled={selLotes.size===0} onClick={()=>setConfirmDone(true)}>Concluir ({selLotesCount})</Btn>}
          </div>
        )}
      </div>

      {/* Table */}
      <div style={{borderRadius:14,boxShadow:"0 1px 6px rgba(0,0,0,0.05)",background:"#fff",overflowX:"auto",overflowY:"auto",maxHeight:"calc(100vh - 280px)",minHeight:180}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:1320}}>
          <thead>
            <tr>
              {canInteract&&<th style={stickyTh}><input type="checkbox" style={{accentColor:ACCENT,cursor:"pointer",width:15,height:15}} checked={selLotes.size===groups.length&&groups.length>0} onChange={e=>toggleAll(e.target.checked)}/></th>}
              <th style={{...TH,width:30}}><ChevronRight size={12}/></th>
              <th style={TH}>Lote</th>
              <th style={TH}>Tubos</th>
              <th style={TH}>IPPNs</th>
              <th style={TH}>Data Bloqueio</th>
              <th style={TH}>Nº Cassete</th>
              <th style={TH}>Pedido/Item</th>
              <th style={TH}>Material</th>
              <th style={descTh}>Descrição</th>
              <th style={TH}>Última Ordem</th>
              <th style={TH}>Qualidade QTS</th>
              <th style={TH}>Depósito SAP</th>
              <th style={TH}>Motivo Bloqueio</th>
              <th style={TH}>Motivo Bloqueio Texto</th>
              <th style={TH}>Razão Bloq.</th>
              <th style={TH}>Descrição Motivo</th>
              <th style={TH_ACCENT}>Definição</th>
              {stage.id>1&&<th style={histTh}>Histórico</th>}
            </tr>
          </thead>
          <tbody>
            {groups.map((group,gi)=>{
              const isExp=expanded.has(group.lote);
              const isSel=selLotes.has(group.lote);
              const fr=group.rows[0];
              const ippnList=group.rows.map(r=>r.ippn).filter(Boolean).join(", ");
              const alt=gi%2===1;
              const fat=isFaturado(group);
              const bg=isSel?"rgba(10,132,255,0.06)":(fat?"rgba(255,69,58,0.07)":(alt?"#FAFAFB":"#fff"));
              const pedidoItem = (group.pedido||group.item) ? `${group.pedido||"—"}/${group.item||"—"}` : "—";
              const currentTratValue = tratativas[group.lote] ?? group.tratativa;

              return[
                <tr key={`g-${group.lote}`} style={{background:bg}}>
                  {canInteract&&<td style={stickyTd(bg)}><input type="checkbox" style={{accentColor:ACCENT,cursor:"pointer",width:15,height:15}} checked={isSel} onChange={()=>toggleSel(group.lote)}/></td>}
                  {/* Expand IPPNs */}
                  <td style={{...TD(false),background:bg,textAlign:"center"}}><button className="spring-btn" onClick={()=>toggleExp(group.lote)} style={{background:"none",border:"none",cursor:"pointer",color:ACCENT,padding:"1px 3px",display:"flex"}}>{isExp?<ChevronDown size={14}/>:<ChevronRight size={14}/>}</button></td>
                  <td style={{...TD(false),background:bg,fontWeight:700,color:"#1C1C1E"}}>{group.lote||"—"}</td>
                  <td style={{...TD(false),background:bg}}><span style={{background:"#E8F4FD",color:"#1A6FA8",borderRadius:6,padding:"2px 7px",fontSize:11,fontWeight:700}}>{group.rows.length}</span></td>
                  <td style={{...TD(false),background:bg,maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:"#555"}}>{ippnList||"—"}</td>
                  <td style={{...TD(false),background:bg}}>{group.data_bloqueio||"—"}</td>
                  <td style={{...TD(false),background:bg}}>{group.num_cassete||"—"}</td>
                  <td style={{...TD(false),background:bg}}>{pedidoItem}</td>
                  <td style={{...TD(false),background:bg}}>{group.material||"—"}</td>
                  <td style={descTd(bg)}>{group.descricao||"—"}</td>
                  <td style={{...TD(false),background:bg}}>{group.ultima_ordem||"—"}</td>
                  <td style={{...TD(false),background:bg}}>{group.qualidade_qts||"—"}</td>
                  <td style={{...TD(false),background:bg}}>{group.deposito_sap||"—"}</td>
                  <td style={{...TD(false),background:bg}}>{group.motivo_bloqueio||"—"}</td>
                  <td style={{...TD(false),background:bg}}>{group.motivo_bloqueio_texto||"—"}</td>
                  <td style={{...TD(false),background:bg}}>{group.razao_bloq||"—"}</td>
                  <td style={{...TD(false),background:bg}}>{group.descricao_motivo||"—"}</td>
                  <td style={{...TD(false),background:gi%2===0?"rgba(10,132,255,0.03)":"rgba(10,132,255,0.06)",minWidth:190}}>
                    {isStage1?(
                      <RecursosSelect value={currentTratValue} onChange={v=>setTratativas(t=>({...t,[group.lote]:v}))} readOnly={!canInteract}/>
                    ):(
                      canInteract?(<input type="text" style={{border:"1.5px solid #E5E5EA",borderRadius:8,padding:"5px 9px",fontSize:11,width:"100%",outline:"none",background:"#fff",minWidth:160,fontFamily:FONT}} placeholder="Definição…" value={currentTratValue||""} onChange={e=>setTratativas(t=>({...t,[group.lote]:e.target.value}))}/>):(<span style={{fontSize:11,color:"#555"}}>{currentTratValue||"—"}</span>)
                    )}
                  </td>
                  {stage.id>1&&<td style={histTd(alt)}><HistoryInline history={fr.history}/></td>}
                </tr>,
                // IPPN expanded info row — only the IPPN itself (Cassete/Data Bloqueio are now table columns)
                ...(isExp?group.rows.map((row,ri)=>(
                  <tr key={`exp-${row._id}`} style={{background:"rgba(10,132,255,0.04)"}}>
                    <td colSpan={totalCols} style={{padding:"8px 14px",borderBottom:`1px solid ${SEPARATOR}`,fontSize:11,color:"#555"}}>
                      <span style={{color:ACCENT,fontWeight:700}}>IPPN {ri+1}:</span> {row.ippn||"—"}
                    </td>
                  </tr>
                )):[]),
              ];
            })}
          </tbody>
        </table>
      </div>

      {/* Confirm modals */}
      {confirmAdv&&<Modal title={`Enviar para "${nextStage?.label}"?`} body={`${selLotesCount} lote(s) / ${selCount} tubo(s) → "${nextStage?.label}". Registrado em nome de ${user.name}.`} onConfirm={doAdvance} onCancel={()=>setConfirmAdv(false)} confirmLabel="Confirmar envio"/>}
      {confirmRet&&<Modal title={`Retornar para "${prevStage?.label}"?`} body={`${selLotesCount} lote(s) / ${selCount} tubo(s) retornam para "${prevStage?.label}". Registrado em nome de ${user.name}.`} onConfirm={doReturn} onCancel={()=>setConfirmRet(false)} confirmLabel="Confirmar retorno" danger/>}
      {confirmDone&&<Modal title="Concluir lotes?" body={`${selLotesCount} lote(s) / ${selCount} tubo(s) serão marcados como concluídos e movidos para o Histórico. Registrado em nome de ${user.name}.`} onConfirm={doComplete} onCancel={()=>setConfirmDone(false)} confirmLabel="Concluir" danger/>}
    </div>
  );
}

// Section title with a small tinted icon badge — used across the Dashboard
function SectionHeader({icon:Icon, title, color=ACCENT}){
  return(
    <div style={{display:"flex",alignItems:"center",gap:9,marginBottom:14}}>
      <div style={{width:28,height:28,borderRadius:9,background:hexToRgba(color,0.12),display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
        <Icon size={14} color={color} strokeWidth={2.4}/>
      </div>
      <div style={{fontSize:14.5,fontWeight:800,color:"#1C1C1E",letterSpacing:"-0.2px",fontFamily:TITLE_FONT}}>{title}</div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// DASHBOARD
// ─────────────────────────────────────────────────────────────────────────────
function Dashboard({stageData,historyRows,onSelectStage}){
  const [selectedStage,setSelectedStage]=useState(null);
  function toggleStageFilter(id){setSelectedStage(s=>s===id?null:id);}

  const allActive=Object.values(stageData).flat();
  const total=allActive.length;
  const etapa1=stageData[1]?.length||0;
  const inFlow=[2,3,4,5,6,7].reduce((s,id)=>s+(stageData[id]?.length||0),0);
  const pendExec=stageData[8]?.length||0;
  const concluded=historyRows.length;

  // Rows used by the cross-filterable charts (Bloqueio por Depósito + Motivos de Bloqueio)
  const filteredRows = selectedStage ? (stageData[selectedStage]||[]) : allActive;
  const filteredStageInfo = selectedStage!=null ? STAGES.find(s=>s.id===selectedStage) : null;

  const byDep={};
  filteredRows.forEach(r=>{const d=r.deposito_sap||"Sem Depósito";byDep[d]=(byDep[d]||0)+1;});

  const byDept={};
  STAGES.filter(s=>s.id!==8&&s.dept).forEach(s=>{const cnt=stageData[s.id]?.length||0;if(!byDept[s.dept])byDept[s.dept]=0;byDept[s.dept]+=cnt;});

  // ─ Pareto: principais Motivos de Bloqueio (até 80% acumulado) ─
  const motivoCounts={};
  filteredRows.forEach(r=>{const m=(r.motivo_bloqueio_texto||"").trim();if(m)motivoCounts[m]=(motivoCounts[m]||0)+1;});
  const motivoTotal=Object.values(motivoCounts).reduce((a,b)=>a+b,0);
  const motivoSorted=Object.entries(motivoCounts).sort((a,b)=>b[1]-a[1]);
  const paretoItems=[];
  let cum=0;
  for(const [label,cnt] of motivoSorted){
    cum+=cnt;
    paretoItems.push({label,cnt,cumPct:motivoTotal?Math.round(cum/motivoTotal*100):0});
    if(motivoTotal && cum/motivoTotal>=0.8) break;
  }

  // ─ Timing stats ─
  const {stageAvg,deptAvg}=computeTimingStats(stageData,historyRows);

  const kpis=[
    {label:"Total Bloqueado",  value:total,     color:"#FF453A", icon:AlertTriangle},
    {label:"Análise Pendente", value:etapa1,    color:ACCENT,    icon:Clock, sid:1},
    {label:"Em Fluxo",         value:inFlow,    color:"#FF9F0A", icon:GitBranch},
    {label:"Pend. Execução",   value:pendExec,  color:"#8E8E93", icon:CheckCircle2, sid:8},
  ];

  const hasTimingData=Object.keys(stageAvg).length>0||Object.keys(deptAvg).length>0;
  const card = {background:SURFACE, borderRadius:16, padding:"16px", boxShadow:"0 1px 6px rgba(0,0,0,0.04)"};
  const now = new Date();

  return(
    <div style={{...WIDE,paddingTop:20}}>
      {/* Report header */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"baseline",marginBottom:20,flexWrap:"wrap",gap:8}}>
        <div style={{fontSize:22,fontWeight:900,color:"#1C1C1E",letterSpacing:"-0.6px",fontFamily:TITLE_FONT}}>Visão Geral</div>
        <div style={{fontSize:11,color:"#8E8E93"}}>Atualizado em {now.toLocaleDateString("pt-BR")} às {now.toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"})}</div>
      </div>

      {/* KPIs */}
      <div style={{display:"flex",gap:14,marginBottom:24,flexWrap:"wrap"}}>
        {kpis.map(k=>{
          const Icon=k.icon;
          return(
            <div key={k.label} className="hover-lift spring-btn" style={{...card,borderRadius:18,flex:1,minWidth:170,minHeight:100,border:`1px solid ${SEPARATOR}`,cursor:k.sid?"pointer":"default",display:"flex",flexDirection:"column",justifyContent:"space-between"}} onClick={()=>k.sid&&onSelectStage(k.sid)}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                <div style={{fontSize:11,color:"#8E8E93",fontWeight:600,textTransform:"uppercase",letterSpacing:"0.5px",maxWidth:"70%"}}>{k.label}</div>
                <div style={{width:32,height:32,borderRadius:10,background:hexToRgba(k.color,0.12),display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>
                  <Icon size={16} color={k.color} strokeWidth={2.4}/>
                </div>
              </div>
              <div style={{fontSize:36,fontWeight:900,color:"#1C1C1E",letterSpacing:"-1.5px",lineHeight:1,fontFamily:TITLE_FONT}}>{k.value}</div>
            </div>
          );
        })}
      </div>

      {/* Funil do Processo — full-width executive funnel chart */}
      <div style={{...card,marginBottom:20}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8,marginBottom:14}}>
          <SectionHeader icon={BarChart3} title="Funil do Processo" color={ACCENT}/>
          {selectedStage!=null&&(
            <button className="spring-btn" onClick={()=>setSelectedStage(null)} style={{display:"flex",alignItems:"center",gap:6,background:"rgba(10,132,255,0.1)",color:ACCENT,border:"none",borderRadius:8,padding:"5px 11px",fontSize:11,fontWeight:700,cursor:"pointer"}}>
              Filtro: Etapa {selectedStage} · {filteredStageInfo?.short}
              <X size={12} strokeWidth={2.6}/>
            </button>
          )}
        </div>
        <div style={{display:"flex",alignItems:"stretch",gap:8}}>
          {STAGES.map((stage,idx)=>{
            const count=stageData[stage.id]?.length||0;
            const maxCount=Math.max(...STAGES.map(s=>stageData[s.id]?.length||0),1);
            const barH=count===0?4:Math.max(6,Math.round(count/maxCount*84));
            const color=STAGE_COLORS[idx];
            const isSelected=selectedStage===stage.id;
            return(
              <div key={stage.id} className="spring-btn" onClick={()=>toggleStageFilter(stage.id)} title="Filtrar dashboard por esta etapa" style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",cursor:"pointer",borderRadius:10,padding:"6px 2px",background:isSelected?hexToRgba(color,0.10):"transparent",boxShadow:isSelected?`0 0 0 2px ${hexToRgba(color,0.5)}`:"none",transition:`all 0.25s ${EASE}`}}>
                <div style={{fontSize:17,fontWeight:900,color:"#1C1C1E",marginBottom:6,fontFamily:TITLE_FONT,letterSpacing:"-0.5px"}}>{count}</div>
                <div style={{height:84,display:"flex",alignItems:"flex-end",width:"100%"}}>
                  <div style={{width:"100%",height:barH,background:hexToRgba(color,0.88),borderRadius:6,transition:`height 0.6s ${EASE}`}}/>
                </div>
                <div style={{fontSize:10,fontWeight:800,color,marginTop:8,textAlign:"center"}}>E{stage.id}</div>
                <div style={{fontSize:9,color:"#8E8E93",textAlign:"center",lineHeight:1.3,marginTop:1}}>{stage.short}</div>
              </div>
            );
          })}
        </div>
      </div>

      {/* Pendências por Departamento (left) + Bloqueio por Depósito (right) */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:20,alignItems:"start"}}>
        <div style={card}>
          <SectionHeader icon={Building2} title="Pendências por Departamento" color="#5E5CE6"/>
          {!Object.keys(byDept).length?<div style={{fontSize:12,color:"#8E8E93"}}>Nenhum tubo em fluxo</div>:Object.entries(byDept).map(([dept,cnt])=>{
            const c=DEPT_COLORS[dept]||{bg:"#F2F2F7",fg:"#3A3A3C"};
            const pct=total>0?Math.round(cnt/total*100):0;
            return(<div key={dept} style={{marginBottom:11}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}><span style={{fontSize:11,fontWeight:600,color:"#3A3A3C"}}>{DEPT_FULL[dept]||dept}</span><span style={{fontSize:12,fontWeight:800,color:c.fg}}>{cnt}</span></div><div style={{height:6,background:"#F2F2F7",borderRadius:3,overflow:"hidden"}}><div style={{width:`${pct}%`,height:"100%",background:c.fg,borderRadius:3,transition:`width 0.6s ${EASE}`}}/></div></div>);
          })}
        </div>

        <div style={card}>
          <SectionHeader icon={Package} title={selectedStage!=null?`Bloqueio por Depósito · E${selectedStage}`:"Bloqueio por Depósito"} color="#FF9F0A"/>
          {!Object.keys(byDep).length?<div style={{fontSize:12,color:"#8E8E93"}}>Nenhum tubo</div>:Object.entries(byDep).sort((a,b)=>b[1]-a[1]).map(([dep,cnt],i)=>{
            const clrs=STAGE_COLORS;
            const col=clrs[i%clrs.length];
            const fTotal=filteredRows.length||1;
            const pct=Math.round(cnt/fTotal*100);
            return(<div key={dep} style={{marginBottom:11}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}><span style={{fontSize:11,fontWeight:600,color:"#3A3A3C"}}>{dep}</span><span style={{fontSize:12,fontWeight:800,color:col}}>{cnt}</span></div><div style={{height:6,background:"#F2F2F7",borderRadius:3,overflow:"hidden"}}><div style={{width:`${pct}%`,height:"100%",background:col,borderRadius:3,transition:`width 0.6s ${EASE}`}}/></div></div>);
          })}
        </div>
      </div>

      {/* Tempo médio por etapa — full width, split per stage / per department */}
      <div style={{...card,marginBottom:20}}>
        <SectionHeader icon={Clock} title="Tempo médio por etapa" color="#34C759"/>
        {!hasTimingData?(
          <div style={{fontSize:13,color:"#8E8E93",textAlign:"center",padding:"12px 0"}}>
            Dados de tempo disponíveis após os primeiros tubos avançarem entre etapas
          </div>
        ):(
          <div style={{display:"grid",gridTemplateColumns:Object.keys(deptAvg).length>0?"1fr 1fr":"1fr",gap:24}}>
            {/* Per-stage bar chart */}
            <div>
              <div style={{fontSize:12,fontWeight:700,color:"#3A3A3C",marginBottom:10}}>Por etapa</div>
              {STAGES.filter(s=>stageAvg[s.id]).map((stage,idx)=>{
                const ms=stageAvg[stage.id];
                const color=STAGE_COLORS[idx];
                const allMs=Object.values(stageAvg).filter(Boolean);
                const maxMs=Math.max(...allMs,1);
                const pct=Math.round(ms/maxMs*100);
                return(
                  <div key={stage.id} style={{marginBottom:10}}>
                    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
                      <div style={{fontSize:11,fontWeight:600,color:"#3A3A3C",display:"flex",alignItems:"center",gap:6}}>
                        <span style={{width:8,height:8,borderRadius:"50%",background:color,display:"inline-block",flexShrink:0}}></span>
                        E{stage.id} · {stage.short}
                      </div>
                      <span style={{fontSize:12,fontWeight:800,color}}>{formatDuration(ms)}</span>
                    </div>
                    <div style={{height:7,background:"#F2F2F7",borderRadius:4,overflow:"hidden"}}>
                      <div style={{width:`${pct}%`,height:"100%",background:color,borderRadius:4,transition:`width 0.6s ${EASE}`}}/>
                    </div>
                  </div>
                );
              })}
            </div>
            {/* Per-dept */}
            {Object.keys(deptAvg).length>0&&(
              <div>
                <div style={{fontSize:12,fontWeight:700,color:"#3A3A3C",marginBottom:10}}>Por departamento</div>
                {Object.entries(deptAvg).map(([dept,ms])=>{
                  const c=DEPT_COLORS[dept]||{bg:"#F2F2F7",fg:"#3A3A3C"};
                  const allMs=Object.values(deptAvg).filter(Boolean);
                  const maxMs=Math.max(...allMs,1);
                  const pct=Math.round(ms/maxMs*100);
                  return(
                    <div key={dept} style={{marginBottom:10}}>
                      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
                        <span style={{fontSize:11,fontWeight:600,color:"#3A3A3C"}}>{DEPT_FULL[dept]||dept}</span>
                        <span style={{fontSize:12,fontWeight:800,color:c.fg}}>{formatDuration(ms)}</span>
                      </div>
                      <div style={{height:7,background:"#F2F2F7",borderRadius:4,overflow:"hidden"}}>
                        <div style={{width:`${pct}%`,height:"100%",background:c.fg,borderRadius:4,transition:`width 0.6s ${EASE}`}}/>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        )}
      </div>

      {/* Motivos de Bloqueio — Pareto (80% acumulado) */}
      <div style={{...card,marginBottom:20}}>
        <SectionHeader icon={PieChart} title={selectedStage!=null?`Motivos de Bloqueio · E${selectedStage}`:"Motivos de Bloqueio"} color="#FF453A"/>
        {paretoItems.length===0?(
          <div style={{fontSize:13,color:"#8E8E93",textAlign:"center",padding:"12px 0"}}>
            Nenhum motivo de bloqueio registrado para os tubos selecionados
          </div>
        ):(
          <div>
            {paretoItems.map((it,i)=>{
              const maxCnt=paretoItems[0].cnt;
              const pct=Math.round(it.cnt/maxCnt*100);
              const col=STAGE_COLORS[i%STAGE_COLORS.length];
              return(
                <div key={it.label} style={{marginBottom:10}}>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4,gap:10}}>
                    <span style={{fontSize:11,fontWeight:600,color:"#3A3A3C",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{it.label}</span>
                    <span style={{display:"flex",alignItems:"center",gap:6,flexShrink:0}}>
                      <span style={{fontSize:12,fontWeight:800,color:col}}>{it.cnt}</span>
                      <span style={{fontSize:10,fontWeight:700,color:"#8E8E93",background:"#F2F2F7",borderRadius:6,padding:"1px 7px"}}>{it.cumPct}% acum.</span>
                    </span>
                  </div>
                  <div style={{height:7,background:"#F2F2F7",borderRadius:4,overflow:"hidden"}}>
                    <div style={{width:`${pct}%`,height:"100%",background:col,borderRadius:4,transition:`width 0.6s ${EASE}`}}/>
                  </div>
                </div>
              );
            })}
            <div style={{fontSize:10,color:"#8E8E93",marginTop:8}}>Principais motivos que somam até 80% do total de tubos bloqueados{selectedStage!=null?` na Etapa ${selectedStage}`:""}.</div>
          </div>
        )}
      </div>

      {total===0&&concluded===0&&(
        <div style={{textAlign:"center",padding:"60px 24px",background:"#fff",borderRadius:18,boxShadow:"0 1px 4px rgba(0,0,0,0.04)"}}>
          <div style={{display:"flex",justifyContent:"center",marginBottom:12}}><Layers size={40} color="#C7C7CC" strokeWidth={1.5}/></div>
          <div style={{fontSize:15,fontWeight:700,color:"#3A3A3C",marginBottom:4}}>Nenhum tubo no sistema</div>
          <div style={{fontSize:13,color:"#8E8E93"}}>Acesse a Pipeline → Etapa 1 para importar tubos bloqueados.</div>
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// HISTÓRICO
// ─────────────────────────────────────────────────────────────────────────────
function HistoricoPage({historyRows}){
  const [filters,setFilters]=useState({pedido:"",item:"",material:"",lotes:"",ippns:""});

  function rowMatches(row){
    const f=(tokens,fields)=>!tokens.length||tokens.some(t=>fields.some(f=>String(row[f]||"").toLowerCase().includes(t)));
    return f(parseTokens(filters.pedido),["pedido"])&&f(parseTokens(filters.item),["item"])&&f(parseTokens(filters.material),["material"])&&f(parseTokens(filters.lotes),["lote"])&&f(parseTokens(filters.ippns),["ippn"]);
  }

  const filtered=historyRows.filter(rowMatches);
  const groups=groupByLote(filtered);
  const hasFilters=Object.values(filters).some(Boolean);

  const searchFields=[
    {key:"pedido",label:"Pedido",ph:"ex: 4500001"},
    {key:"item",label:"Item",ph:"ex: 0010"},
    {key:"material",label:"Material",ph:"ex: MAT-001"},
    {key:"lotes",label:"Lotes",ph:"ex: LOTE-001"},
    {key:"ippns",label:"IPPNs",ph:"ex: IPN001"},
  ];

  const histTh = {...TH, minWidth:380};
  const histTd = (alt)=>({...TD(alt), minWidth:380, padding:"10px 14px"});

  return(
    <div style={{...WIDE,paddingTop:20}}>
      {/* Search */}
      <div style={{background:SURFACE,borderRadius:14,padding:"14px 16px",boxShadow:"0 1px 4px rgba(0,0,0,0.04)",marginBottom:16}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10}}>
          <span style={{fontSize:12.5,fontWeight:700,color:"#3A3A3C",display:"flex",alignItems:"center",gap:6}}><Search size={14} color={ACCENT}/> Pesquisar Histórico</span>
          {hasFilters&&<button className="spring-btn" onClick={()=>setFilters({pedido:"",item:"",material:"",lotes:"",ippns:""})} style={{background:"none",border:"none",color:"#FF453A",fontSize:12,fontWeight:600,cursor:"pointer"}}>Limpar</button>}
        </div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(190px,1fr))",gap:10}}>
          {searchFields.map(f=><div key={f.key}><div style={{fontSize:10,fontWeight:600,color:"#8E8E93",marginBottom:3}}>{f.label}</div><input value={filters[f.key]} onChange={e=>setFilters(v=>({...v,[f.key]:e.target.value}))} placeholder={f.ph} style={{width:"100%",border:"1.5px solid #E5E5EA",borderRadius:9,padding:"7px 10px",fontSize:12.5,outline:"none",background:"#FAFAFB",boxSizing:"border-box",fontFamily:FONT}}/></div>)}
        </div>
        {hasFilters&&<div style={{marginTop:8,fontSize:10,color:"#8E8E93"}}>Separe múltiplos valores com espaço, vírgula ou ponto-e-vírgula</div>}
      </div>

      {historyRows.length===0?(
        <div style={{textAlign:"center",padding:"60px 24px",background:"#fff",borderRadius:18,boxShadow:"0 1px 4px rgba(0,0,0,0.04)"}}>
          <div style={{display:"flex",justifyContent:"center",marginBottom:12}}><Inbox size={40} color="#C7C7CC" strokeWidth={1.5}/></div>
          <div style={{fontSize:15,fontWeight:700,color:"#3A3A3C",marginBottom:4}}>Nenhum registro ainda</div>
          <div style={{fontSize:13,color:"#8E8E93"}}>Os tubos concluídos na Etapa 8 aparecerão aqui.</div>
        </div>
      ):(
        <div>
          <div style={{fontSize:12,color:"#8E8E93",marginBottom:10}}>{groups.length} lote{groups.length!==1?"s":""} · {filtered.length} tubo{filtered.length!==1?"s":""}{hasFilters&&` (filtrado de ${historyRows.length})`}</div>
          <div style={{borderRadius:14,boxShadow:"0 1px 6px rgba(0,0,0,0.05)",background:"#fff",overflowX:"auto",overflowY:"auto",maxHeight:"calc(100vh - 320px)"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:900}}>
              <thead>
                <tr>
                  <th style={{...TH,width:30}}></th>
                  {["Lote","Tubos","IPPNs","Pedido/Item","Material","Depósito","Motivo Bloqueio","Data Bloqueio","Nº Cassete"].map(h=><th key={h} style={TH}>{h}</th>)}
                  <th style={histTh}>Histórico</th>
                </tr>
              </thead>
              <tbody>
                {groups.map((group,gi)=>{
                  const fr=group.rows[0];
                  const ippnList=group.rows.map(r=>r.ippn).filter(Boolean).join(", ");
                  const alt=gi%2===1;
                  return(
                    <tr key={`h-${group.lote}`} style={{background:alt?"#FAFAFB":"#fff"}}>
                      <td style={{...TD(alt),textAlign:"center"}}></td>
                      <td style={{...TD(alt),fontWeight:700,color:"#1C1C1E"}}>{group.lote||"—"}</td>
                      <td style={TD(alt)}><span style={{background:"#EDF7EE",color:"#1A7A3A",borderRadius:6,padding:"2px 7px",fontSize:11,fontWeight:700}}>{group.rows.length}</span></td>
                      <td style={{...TD(alt),maxWidth:150,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{ippnList||"—"}</td>
                      <td style={TD(alt)}>{(group.pedido||group.item)?`${group.pedido||"—"}/${group.item||"—"}`:"—"}</td>
                      <td style={TD(alt)}>{group.material||"—"}</td>
                      <td style={TD(alt)}>{group.deposito_sap||"—"}</td>
                      <td style={TD(alt)}>{group.motivo_bloqueio||"—"}</td>
                      <td style={TD(alt)}>{group.data_bloqueio||"—"}</td>
                      <td style={TD(alt)}>{group.num_cassete||"—"}</td>
                      <td style={histTd(alt)}><HistoryInline history={fr.history}/></td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// CONFIGURAÇÕES
// ─────────────────────────────────────────────────────────────────────────────
function ConfigPage({users,setUsers,faturamento,setFaturamento,stageData}){
  const [tab,setTab]=useState("users");
  const [editing,setEditing]=useState(null); // null or user obj
  const [form,setForm]=useState({});
  const [showDel,setShowDel]=useState(null);
  const [saveMsg,setSaveMsg]=useState("");
  const [fatForm,setFatForm]=useState({pedido:"",item:"",descricao:""});
  const [fatLoading,setFatLoading]=useState(false);
  const [fatError,setFatError]=useState("");

  const deptOptions=Object.keys(DEPT_FULL).filter(d=>d!=="Admin");

  function openNew(){
    setForm({id:`u_${Date.now()}`,email:"",password:"",name:"",dept:deptOptions[0],allowedStages:DEPT_DEFAULT_STAGES[deptOptions[0]]||[],active:true});
    setEditing("new");
  }
  function openEdit(u){setForm({...u,allowedStages:[...(u.allowedStages||[])]});setEditing(u.id);}
  function saveUser(){
    if(!form.email||!form.name||!form.password){setSaveMsg("Preencha todos os campos obrigatórios.");return;}
    if(editing==="new"){
      setUsers(us=>[...us,{...form}]);
    } else {
      setUsers(us=>us.map(u=>u.id===editing?{...form}:u));
    }
    setEditing(null);setSaveMsg("Usuário salvo com sucesso.");setTimeout(()=>setSaveMsg(""),3000);
  }
  function deleteUser(){setUsers(us=>us.filter(u=>u.id!==showDel));setShowDel(null);}
  function toggleStage(sid){
    const arr=form.allowedStages||[];
    setForm(f=>({...f,allowedStages:arr.includes(sid)?arr.filter(s=>s!==sid):[...arr,sid].sort((a,b)=>a-b)}));
  }
  function onDeptChange(dept){
    setForm(f=>({...f,dept,allowedStages:DEPT_DEFAULT_STAGES[dept]||[]}));
  }

  // ─ Faturamento: count how many tubes across the whole pipeline match a Pedido/Item ─
  function countTubos(pedido,item){
    const p=(pedido||"").trim().toLowerCase(), it=(item||"").trim().toLowerCase();
    let count=0;
    Object.values(stageData||{}).forEach(arr=>{
      (arr||[]).forEach(r=>{
        if((r.pedido||"").trim().toLowerCase()===p && (r.item||"").trim().toLowerCase()===it) count++;
      });
    });
    return count;
  }
  function addFaturamento(){
    if(!fatForm.pedido||!fatForm.item){setFatError("Informe Pedido e Item.");return;}
    setFaturamento(f=>{
      const exists=f.some(x=>x.pedido.trim().toLowerCase()===fatForm.pedido.trim().toLowerCase()&&x.item.trim().toLowerCase()===fatForm.item.trim().toLowerCase());
      if(exists) return f;
      return [...f,{id:`fat_${Date.now()}`,pedido:fatForm.pedido.trim(),item:fatForm.item.trim(),descricao:fatForm.descricao.trim()}];
    });
    setFatForm({pedido:"",item:"",descricao:""});setFatError("");
  }
  function removeFaturamento(id){setFaturamento(f=>f.filter(x=>x.id!==id));}
  async function importFaturamento(file){
    if(!file)return;
    setFatLoading(true);setFatError("");
    try{
      const rows=await parseFile(file);
      if(!rows.length){setFatError("Nenhuma linha válida encontrada.");setFatLoading(false);return;}
      setFaturamento(f=>{
        const seen=new Set(f.map(x=>`${x.pedido.trim().toLowerCase()}|${x.item.trim().toLowerCase()}`));
        const novos=[];
        rows.forEach((r,i)=>{
          if(!r.pedido||!r.item) return;
          const key=`${r.pedido.trim().toLowerCase()}|${r.item.trim().toLowerCase()}`;
          if(seen.has(key)) return;
          seen.add(key);
          novos.push({id:`fat_${Date.now()}_${i}`,pedido:r.pedido.trim(),item:r.item.trim(),descricao:(r.descricao||"").trim()});
        });
        return [...f,...novos];
      });
    }catch(e){setFatError("Erro ao processar arquivo: "+e.message);}
    setFatLoading(false);
  }

  const tabStyle=(active)=>({padding:"8px 18px",borderRadius:9,border:"none",fontSize:13,fontWeight:active?700:500,color:active?"#1C1C1E":"#8E8E93",background:active?"#fff":"transparent",cursor:"pointer",boxShadow:active?"0 1px 4px rgba(0,0,0,0.08)":"none",display:"flex",alignItems:"center",gap:6,fontFamily:FONT});

  return(
    <div style={{...WIDE,maxWidth:1000,paddingTop:20}}>
      <div style={{display:"flex",gap:2,background:"#E9E9EC",borderRadius:11,padding:3,marginBottom:20,width:"fit-content"}}>
        <button className="spring-btn" style={tabStyle(tab==="users")} onClick={()=>setTab("users")}><Users size={14}/> Usuários</button>
        <button className="spring-btn" style={tabStyle(tab==="perms")} onClick={()=>setTab("perms")}><Shield size={14}/> Permissões</button>
        <button className="spring-btn" style={tabStyle(tab==="faturamento")} onClick={()=>setTab("faturamento")}><Receipt size={14}/> Faturamento</button>
      </div>

      {saveMsg&&<div style={{background:"#EDF7EE",color:"#1A7A3A",borderRadius:10,padding:"8px 14px",marginBottom:14,fontSize:13,fontWeight:600,display:"flex",alignItems:"center",gap:6}}><CheckCircle2 size={14}/>{saveMsg}</div>}

      {tab==="users"&&(
        <div>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
            <div style={{fontSize:14,fontWeight:700,color:"#3A3A3C"}}>Usuários cadastrados ({users.length})</div>
            <Btn onClick={openNew} icon={Plus}>Novo Usuário</Btn>
          </div>
          <div style={{background:"#fff",borderRadius:14,boxShadow:"0 1px 6px rgba(0,0,0,0.05)",overflow:"hidden"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
              <thead><tr>{["Nome","E-mail","Departamento","Etapas","Status","Ações"].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead>
              <tbody>
                {users.map((u,i)=>(
                  <tr key={u.id} style={{background:i%2?"#FAFAFB":"#fff"}}>
                    <td style={{...TD(i%2===1),fontWeight:600}}>{u.name}</td>
                    <td style={{...TD(i%2===1),color:"#555"}}>{u.email}</td>
                    <td style={TD(i%2===1)}><DeptTag dept={u.dept}/></td>
                    <td style={{...TD(i%2===1),fontSize:11}}>{(u.allowedStages||[]).map(s=>`E${s}`).join(", ")||"—"}</td>
                    <td style={TD(i%2===1)}><span style={{background:u.active!==false?"#EDF7EE":"#F2F2F7",color:u.active!==false?"#1A7A3A":"#8E8E93",borderRadius:6,padding:"2px 8px",fontSize:11,fontWeight:600}}>{u.active!==false?"Ativo":"Inativo"}</span></td>
                    <td style={TD(i%2===1)}>
                      <div style={{display:"flex",gap:6}}>
                        <Btn small variant="secondary" icon={Pencil} onClick={()=>openEdit(u)}>Editar</Btn>
                        <Btn small variant="danger" icon={Trash2} onClick={()=>setShowDel(u.id)}>Excluir</Btn>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {tab==="perms"&&(
        <div>
          <div style={{fontSize:14,fontWeight:700,color:"#3A3A3C",marginBottom:14}}>Permissões por Departamento (padrão)</div>
          <div style={{background:"#fff",borderRadius:14,padding:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.05)"}}>
            {Object.entries(DEPT_DEFAULT_STAGES).filter(([d])=>d!=="Admin").map(([dept,stages])=>{
              const c=DEPT_COLORS[dept]||{bg:"#F2F2F7",fg:"#3A3A3C"};
              return(
                <div key={dept} style={{marginBottom:16,paddingBottom:16,borderBottom:`1px solid ${SEPARATOR}`}}>
                  <div style={{marginBottom:8}}><DeptTag dept={dept}/></div>
                  <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                    {STAGES.map(s=>{
                      const allowed=stages.includes(s.id);
                      return(<div key={s.id} style={{background:allowed?c.bg:"#F2F2F7",color:allowed?c.fg:"#C7C7CC",borderRadius:8,padding:"4px 10px",fontSize:11,fontWeight:600}}>{allowed?"✓":""} E{s.id}: {s.short}</div>);
                    })}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {tab==="faturamento"&&(
        <div>
          <div style={{fontSize:14,fontWeight:700,color:"#3A3A3C",marginBottom:14}}>Pedidos de Faturamento (mês corrente)</div>

          {/* Manual add form */}
          <div style={{background:"#fff",borderRadius:14,padding:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.05)",marginBottom:14}}>
            <div style={{fontSize:12,fontWeight:700,color:"#3A3A3C",marginBottom:10}}>Adicionar manualmente</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 2fr auto",gap:10,alignItems:"end"}}>
              <div><label style={{fontSize:11,fontWeight:600,color:"#8E8E93",marginBottom:4,display:"block"}}>Pedido *</label><input style={INP} value={fatForm.pedido} onChange={e=>setFatForm(f=>({...f,pedido:e.target.value}))} placeholder="ex: 12000010"/></div>
              <div><label style={{fontSize:11,fontWeight:600,color:"#8E8E93",marginBottom:4,display:"block"}}>Item *</label><input style={INP} value={fatForm.item} onChange={e=>setFatForm(f=>({...f,item:e.target.value}))} placeholder="ex: 10"/></div>
              <div><label style={{fontSize:11,fontWeight:600,color:"#8E8E93",marginBottom:4,display:"block"}}>Descrição</label><input style={INP} value={fatForm.descricao} onChange={e=>setFatForm(f=>({...f,descricao:e.target.value}))} placeholder="Descrição do produto"/></div>
              <Btn icon={Plus} onClick={addFaturamento}>Adicionar</Btn>
            </div>

            <div style={{display:"flex",alignItems:"center",gap:10,marginTop:14,paddingTop:14,borderTop:`1px solid ${SEPARATOR}`}}>
              <span style={{fontSize:11,fontWeight:600,color:"#8E8E93"}}>ou importar arquivo:</span>
              <Btn variant="secondary" small icon={fatLoading?Loader2:Upload} onClick={()=>document.getElementById("fatFileInput").click()} disabled={fatLoading}>{fatLoading?"Processando…":"CSV / XLS / XLSX"}</Btn>
              <input id="fatFileInput" type="file" accept=".csv,.txt,.xls,.xlsx" style={{display:"none"}} onChange={e=>{importFaturamento(e.target.files[0]);e.target.value="";}}/>
            </div>
            {fatError&&<div style={{color:"#FF453A",fontSize:12,marginTop:10}}>{fatError}</div>}
          </div>

          {/* Registered table */}
          <div style={{background:"#fff",borderRadius:14,boxShadow:"0 1px 6px rgba(0,0,0,0.05)",overflow:"hidden"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
              <thead><tr>{["Pedido/Item","Descrição","Tubos",""].map(h=><th key={h} style={TH}>{h}</th>)}</tr></thead>
              <tbody>
                {faturamento.length===0?(
                  <tr><td colSpan={4} style={{...TD(false),textAlign:"center",color:"#8E8E93",padding:"24px"}}>Nenhum pedido de faturamento cadastrado.</td></tr>
                ):faturamento.map((f,i)=>(
                  <tr key={f.id} style={{background:i%2?"#FAFAFB":"#fff"}}>
                    <td style={{...TD(i%2===1),fontWeight:700,color:"#1C1C1E"}}>{f.pedido}/{f.item}</td>
                    <td style={{...TD(i%2===1),color:"#555"}}>{f.descricao||"—"}</td>
                    <td style={TD(i%2===1)}><span style={{background:"#E8F4FD",color:"#1A6FA8",borderRadius:6,padding:"2px 8px",fontSize:11,fontWeight:700}}>{countTubos(f.pedido,f.item)}</span></td>
                    <td style={TD(i%2===1)}><Btn small variant="danger" icon={Trash2} onClick={()=>removeFaturamento(f.id)}>Excluir</Btn></td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          <div style={{fontSize:11,color:"#8E8E93",marginTop:8}}>Pedidos cadastrados aqui são destacados em vermelho claro na tabela da Pipeline.</div>
        </div>
      )}


      {editing&&(
        <Modal title={editing==="new"?"Novo Usuário":"Editar Usuário"} onConfirm={saveUser} onCancel={()=>setEditing(null)} confirmLabel="Salvar">
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
            <div><label style={{fontSize:12,fontWeight:600,color:"#3A3A3C",marginBottom:4,display:"block"}}>Nome *</label><input style={INP} value={form.name||""} onChange={e=>setForm(f=>({...f,name:e.target.value}))} placeholder="Nome completo"/></div>
            <div><label style={{fontSize:12,fontWeight:600,color:"#3A3A3C",marginBottom:4,display:"block"}}>E-mail *</label><input style={INP} type="email" value={form.email||""} onChange={e=>setForm(f=>({...f,email:e.target.value}))} placeholder="email@empresa.com"/></div>
            <div><label style={{fontSize:12,fontWeight:600,color:"#3A3A3C",marginBottom:4,display:"block"}}>Senha *</label><input style={INP} type="password" value={form.password||""} onChange={e=>setForm(f=>({...f,password:e.target.value}))} placeholder="Senha"/></div>
            <div><label style={{fontSize:12,fontWeight:600,color:"#3A3A3C",marginBottom:4,display:"block"}}>Departamento</label><select style={{...INP,cursor:"pointer"}} value={form.dept||""} onChange={e=>onDeptChange(e.target.value)}>{deptOptions.map(d=><option key={d} value={d}>{DEPT_FULL[d]||d}</option>)}</select></div>
          </div>
          <div style={{marginBottom:10}}>
            <label style={{fontSize:12,fontWeight:600,color:"#3A3A3C",marginBottom:6,display:"block"}}>Etapas autorizadas</label>
            <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
              {STAGES.map(s=>{
                const allowed=(form.allowedStages||[]).includes(s.id);
                return(<button key={s.id} type="button" className="spring-btn" onClick={()=>toggleStage(s.id)} style={{background:allowed?ACCENT:"#F2F2F7",color:allowed?"#fff":"#555",border:"none",borderRadius:8,padding:"5px 11px",fontSize:11,fontWeight:600,cursor:"pointer"}}>E{s.id}: {s.short}</button>);
              })}
            </div>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <input type="checkbox" checked={form.active!==false} onChange={e=>setForm(f=>({...f,active:e.target.checked}))} style={{accentColor:ACCENT,width:15,height:15}}/>
            <label style={{fontSize:12,fontWeight:600,color:"#3A3A3C"}}>Usuário ativo</label>
          </div>
        </Modal>
      )}
      {showDel&&<Modal title="Excluir usuário?" body="Esta ação não pode ser desfeita." onConfirm={deleteUser} onCancel={()=>setShowDel(null)} confirmLabel="Excluir" danger/>}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// PIPELINE PAGE (main flow view)
// ─────────────────────────────────────────────────────────────────────────────
function PipelinePage({stageData,setStageData,user,historyRows,setHistoryRows,showToast,initialStage,faturamento}){
  const [activeStage,setActiveStage]=useState(initialStage||1);
  const [filters,setFilters]=useState({pedido_item:"",lotes:"",ippns:"",deposito:""});

  const hasSearch=Object.values(filters).some(Boolean);

  // Compute search result per stage
  const searchResult = hasSearch ? (() => {
    const res={};
    STAGES.forEach(s=>{
      const matching=(stageData[s.id]||[]).filter(r=>rowMatchesSearch(r,filters));
      res[s.id]=matching.length;
    });
    return res;
  })() : null;

  function handleImport(rows){
    const stamped=rows.map(r=>({...r,stageEnteredMs:Date.now()}));
    setStageData(d=>({...d,1:[...d[1],...stamped]}));
    setActiveStage(1);
    showToast(`${rows.length} tubos importados para a Etapa 1`);
  }
  function handleAdvance(toMove,remaining){
    const nextId=activeStage+1;
    setStageData(d=>({...d,[activeStage]:remaining,[nextId]:[...(d[nextId]||[]),...toMove]}));
    showToast(`${toMove.length} tubo(s) → Etapa ${nextId}`);
  }
  function handleReturn(toReturn,remaining){
    const prevId=activeStage-1;
    setStageData(d=>({...d,[activeStage]:remaining,[prevId]:[...(d[prevId]||[]),...toReturn]}));
    showToast(`${toReturn.length} tubo(s) retornado(s) para Etapa ${prevId}`);
  }
  function handleComplete(toComplete,remaining){
    setStageData(d=>({...d,8:remaining}));
    setHistoryRows(h=>[...h,...toComplete]);
    showToast(`${toComplete.length} tubo(s) movido(s) para o Histórico`);
  }

  const stage=STAGES.find(s=>s.id===activeStage);

  return(
    <div>
      <PipelineBar stageData={stageData} activeStage={activeStage} onSelectStage={setActiveStage} searchResult={searchResult}/>
      <div style={{...CONTAINER,paddingTop:16,paddingBottom:40}}>
        <SearchBar filters={filters} onChange={setFilters} onClear={()=>setFilters({pedido_item:"",lotes:"",ippns:"",deposito:""})}/>

        {/* Search results summary */}
        {hasSearch&&searchResult&&(
          <div style={{background:"#fff",borderRadius:12,padding:"10px 14px",marginBottom:16,boxShadow:"0 1px 4px rgba(0,0,0,0.04)",fontSize:12,color:"#3A3A3C"}}>
            <span style={{fontWeight:700}}>Resultado da busca: </span>
            {STAGES.map(s=>(searchResult[s.id]>0)&&<span key={s.id} style={{marginRight:10,background:"rgba(255,159,10,0.1)",color:"#B45309",borderRadius:6,padding:"2px 8px",fontWeight:600}}>E{s.id}: {searchResult[s.id]}</span>)}
            {Object.values(searchResult).every(v=>v===0)&&<span style={{color:"#8E8E93"}}>Nenhum tubo encontrado</span>}
          </div>
        )}

        {activeStage===1&&stageData[1].length===0?(
          <ImportStep onImport={handleImport}/>
        ):(
          <div>
            <StageView
              stage={stage}
              rows={stageData[activeStage]||[]}
              user={user}
              onAdvance={handleAdvance}
              onReturn={activeStage>1?handleReturn:null}
              onComplete={activeStage===8?handleComplete:null}
              filters={filters}
              faturamento={faturamento}
            />
            {activeStage===1&&stageData[1].length>0&&(
              <div style={{marginTop:14}}>
                <Btn variant="secondary" small icon={RotateCcw} onClick={()=>{if(window.confirm("Remove todos os tubos da Etapa 1. Continuar?"))setStageData(d=>({...d,1:[]}));}}>Novo arquivo</Btn>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// MAIN APP
// ─────────────────────────────────────────────────────────────────────────────
const PAGE_META = {
  dashboard:     { title:"Dashboard",     subtitle:"Visão geral do fluxo de liberação" },
  pipeline:      { title:"Pipeline",      subtitle:"Etapas do fluxo de liberação" },
  historico:     { title:"Histórico",     subtitle:"Tubos que completaram o fluxo de liberação" },
  configuracoes: { title:"Configurações", subtitle:"Usuários e permissões" },
};

export default function App(){
  const [users,setUsers]=useState(DEFAULT_USERS);
  const [user,setUser]=useState(null);
  const [stageData,setStageData]=useState(initStageData);
  const [historyRows,setHistoryRows]=useState([]);
  const [page,setPage]=useState("dashboard");
  const [pipelineTarget,setPipelineTarget]=useState(1);
  const [toast,setToast]=useState("");
  const [showLogout,setShowLogout]=useState(false);
  const [faturamento,setFaturamento]=useState([]);

  function showToast(msg){setToast(msg);setTimeout(()=>setToast(""),3200);}

  function goToPipelineStage(stageId){
    setPipelineTarget(stageId);
    setPage("pipeline");
  }

  if(!user) return <LoginPage onLogin={setUser} users={users}/>;

  const meta = PAGE_META[page];

  return(
    <div style={{height:"100vh",background:APP_BG,color:"#1C1C1E",fontFamily:FONT,overflow:"hidden"}}>
      <GlobalStyles/>
      <Sidebar page={page} setPage={setPage}/>

      <div style={{display:"flex",flexDirection:"column",height:"100vh",overflow:"hidden"}}>
        <TopBar title={meta.title} subtitle={meta.subtitle} user={user} onLogoutClick={()=>setShowLogout(true)}/>
        <div style={{flex:1,overflowY:"auto"}}>
          {page==="dashboard"&&<Dashboard stageData={stageData} historyRows={historyRows} onSelectStage={goToPipelineStage}/>}
          {page==="pipeline"&&<PipelinePage key={pipelineTarget} initialStage={pipelineTarget} stageData={stageData} setStageData={setStageData} user={user} historyRows={historyRows} setHistoryRows={setHistoryRows} showToast={showToast} faturamento={faturamento}/>}
          {page==="historico"&&<HistoricoPage historyRows={historyRows}/>}
          {page==="configuracoes"&&<ConfigPage users={users} setUsers={setUsers} faturamento={faturamento} setFaturamento={setFaturamento} stageData={stageData}/>}
        </div>
      </div>

      <Toast msg={toast}/>
      {showLogout&&<Modal title="Sair da conta" body={`Encerrar sessão de ${user.name}?`} onConfirm={()=>{setUser(null);setShowLogout(false);setStageData(initStageData());setHistoryRows([]);setPage("dashboard");}} onCancel={()=>setShowLogout(false)} confirmLabel="Sair" danger/>}
    </div>
  );
}
