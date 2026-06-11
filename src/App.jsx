import { useState, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ─────────────────────────────────────────────────────────────────────────────
// CONSTANTS
// ─────────────────────────────────────────────────────────────────────────────
const STAGES = [
  { id:1, label:"Análise do Bloqueio",    short:"Análise de Bloqueio",  dept:"CQ Área Técnica" },
  { id:2, label:"Definição do Recurso",   short:"Análise de Recurso",   dept:"Planejamento UAP" },
  { id:3, label:"Criação de Ordem",       short:"Definição de Ordem",   dept:"Planejamento Central" },
  { id:4, label:"Instrução da Qualidade", short:"Instrução",            dept:"CQ Área Técnica" },
  { id:5, label:"Liberação para Vínculo", short:"Lib. Vínculo",         dept:"CQ Lib. Intermediária" },
  { id:6, label:"Vínculo dos Lotes",      short:"Vínculo",              dept:"Planejamento UAP" },
  { id:7, label:"Ativação de Flag",       short:"Flag",                 dept:"CQ Lib. Intermediária" },
  { id:8, label:"Pendente Execução",      short:"Pend. Execução",       dept:"CQ Lib. Intermediária" },
];

const STAGE_COLORS = ["#007AFF","#34C759","#FF9500","#AF52DE","#FF2D55","#5AC8FA","#FF6B35","#8E8E93"];

const DEPT_FULL = {
  "CQ Área Técnica":       "Controle da Qualidade [Área técnica]",
  "Planejamento UAP":      "Planejamento UAP",
  "Planejamento Central":  "Planejamento Central",
  "CQ Lib. Intermediária": "Controle da Qualidade [Liberação Intermediária]",
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
const ALL_COLS = ["pedido","item","material","descricao","data_bloqueio","ultima_ordem","lote","ippn","qualidade_qts","deposito_sap","motivo_bloqueio","motivo_bloqueio_texto","razao_bloq","num_cassete"];
const COL_LABELS = {
  pedido:"Pedido", item:"Item", material:"Material", descricao:"Descrição", data_bloqueio:"Data Bloqueio",
  ultima_ordem:"Última Ordem", lote:"Lote", ippn:"IPPN", qualidade_qts:"Qualidade QTS",
  deposito_sap:"Depósito SAP", motivo_bloqueio:"Motivo Bloqueio", motivo_bloqueio_texto:"Motivo Bloqueio Texto",
  razao_bloq:"Razão Bloq.", num_cassete:"Nº Cassete",
};

// ─────────────────────────────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────────────────────────────
function normalizeHdr(h) {
  return String(h).trim().toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g,"")
    .replace(/[\s\-\/\\]+/g,"_").replace(/[^a-z0-9_]/g,"");
}

function rowsFromMatrix(matrix) {
  if (matrix.length < 2) return [];
  const headers = matrix[0].map(normalizeHdr);
  return matrix.slice(1).filter(v=>v.some(x=>String(x||"").trim())).map((vals,i)=>{
    const obj = { _id:`row_${Date.now()}_${i}`, tratativa:"", history:[], stageEnteredMs: Date.now() };
    ALL_COLS.forEach(col=>{
      const idx = headers.indexOf(col);
      obj[col] = idx>=0 ? String(vals[idx]??"").trim() : "";
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
          const wb = XLSX.read(e.target.result,{type:"array"});
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
    isXLS ? reader.readAsArrayBuffer(file) : reader.readAsArrayBuffer(file);
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

function formatDuration(ms) {
  if(!ms||ms<=0) return "—";
  const h = Math.floor(ms/3600000);
  const d = Math.floor(ms/86400000);
  if(d>=1) return `${d}d ${Math.floor((ms%86400000)/3600000)}h`;
  if(h>=1) return `${h}h ${Math.floor((ms%3600000)/60000)}min`;
  return `${Math.floor(ms/60000)}min`;
}

// Compute average time spent at each stage from history
function computeTimingStats(stageData, historyRows) {
  // All rows everywhere (including history)
  const allRows = [...Object.values(stageData).flat(), ...historyRows];
  // Per stage: collect durations from history entries that have dateMs
  const stageMs = {}; // stageId -> [ms]
  const deptMs  = {}; // dept    -> [ms]

  allRows.forEach(row=>{
    const hist = row.history||[];
    // build timeline: entry into each stage + how long until next event
    // stageEnteredMs is when it entered current stage; each history entry records when it LEFT a stage
    // We stored dateMs when advancing
    let prevMs = row.stageEnteredMs||null;
    hist.forEach(h=>{
      if(!h.dateMs||!prevMs) { if(h.dateMs) prevMs=h.dateMs; return; }
      const dur = h.dateMs - prevMs;
      if(dur>0 && dur < 90*24*3600000) { // ignore >90d (data errors)
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

// ─────────────────────────────────────────────────────────────────────────────
// SMALL UI ATOMS
// ─────────────────────────────────────────────────────────────────────────────
function Toast({msg}){ if(!msg)return null; return(<div style={{position:"fixed",bottom:28,left:"50%",transform:"translateX(-50%)",background:"#1C1C1E",color:"#fff",borderRadius:12,padding:"12px 22px",fontSize:14,fontWeight:500,zIndex:500,boxShadow:"0 8px 28px rgba(0,0,0,0.3)",display:"flex",alignItems:"center",gap:8,whiteSpace:"nowrap",pointerEvents:"none"}}><span style={{color:"#34C759",fontWeight:800}}>✓</span>{msg}</div>); }

function Modal({title,body,onConfirm,onCancel,confirmLabel="Confirmar",danger=false,children}){
  return(<div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:400,display:"flex",alignItems:"center",justifyContent:"center",padding:20}} onClick={onCancel}><div style={{background:"#fff",borderRadius:20,padding:"28px 28px 24px",maxWidth:480,width:"100%",boxShadow:"0 28px 70px rgba(0,0,0,0.3)"}} onClick={e=>e.stopPropagation()}><div style={{fontSize:18,fontWeight:800,color:"#1C1C1E",marginBottom:8}}>{title}</div>{body&&<div style={{fontSize:14,color:"#3A3A3C",marginBottom:children?12:22,lineHeight:1.65}}>{body}</div>}{children&&<div style={{marginBottom:20}}>{children}</div>}<div style={{display:"flex",gap:10,justifyContent:"flex-end"}}><button style={{background:"#F2F2F7",color:"#007AFF",border:"none",borderRadius:12,padding:"11px 22px",fontSize:14,fontWeight:600,cursor:"pointer"}} onClick={onCancel}>Cancelar</button><button style={{background:danger?"linear-gradient(135deg,#FF3B30,#C0392B)":"linear-gradient(135deg,#007AFF,#0051D4)",color:"#fff",border:"none",borderRadius:12,padding:"11px 22px",fontSize:14,fontWeight:600,cursor:"pointer"}} onClick={onConfirm}>{confirmLabel}</button></div></div></div>);
}

function DeptTag({dept}){
  const c=DEPT_COLORS[dept]||{bg:"#F2F2F7",fg:"#3A3A3C"};
  return <span style={{background:c.bg,color:c.fg,borderRadius:6,padding:"2px 9px",fontSize:11,fontWeight:600,display:"inline-block"}}>{DEPT_FULL[dept]||dept}</span>;
}

function Btn({children,variant="primary",disabled,onClick,small,style={}}){
  const variants={primary:{background:"linear-gradient(135deg,#007AFF,#0051D4)",color:"#fff"},secondary:{background:"#F2F2F7",color:"#007AFF"},danger:{background:"linear-gradient(135deg,#FF3B30,#C0392B)",color:"#fff"},warning:{background:"linear-gradient(135deg,#FF9500,#E8890A)",color:"#fff"},ghost:{background:"transparent",color:"#007AFF"}};
  return(<button disabled={disabled} onClick={onClick} style={{border:"none",borderRadius:10,padding:small?"7px 14px":"10px 18px",fontSize:small?12:13,fontWeight:600,cursor:disabled?"not-allowed":"pointer",opacity:disabled?0.42:1,display:"inline-flex",alignItems:"center",gap:6,transition:"opacity 0.15s",...variants[variant],...style}}>{children}</button>);
}

const INP = {width:"100%",border:"1.5px solid #E5E5EA",borderRadius:9,padding:"9px 12px",fontSize:13,outline:"none",background:"#fff",boxSizing:"border-box"};
const TH_DARK = {background:"#0A2240",padding:"9px 11px",textAlign:"left",fontWeight:700,color:"#fff",borderBottom:"2px solid #1A3A5C",whiteSpace:"nowrap",fontSize:11,position:"sticky",top:0,zIndex:10};
const TH_GREEN = {...{background:"#0A2240",padding:"9px 11px",textAlign:"left",fontWeight:700,color:"#fff",borderBottom:"2px solid #1A3A5C",whiteSpace:"nowrap",fontSize:11,position:"sticky",top:0,zIndex:10}, background:"#1A5C2A",borderBottom:"2px solid #1A7A3A"};
const TD = (alt)=>({padding:"9px 11px",borderBottom:"1px solid #F0F0F5",verticalAlign:"middle",fontSize:12,color:"#1C1C1E",background:alt?"#F8F9FB":"#fff"});

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
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"linear-gradient(160deg,#0A2240 0%,#1A3A5C 55%,#0D3B6B 100%)"}}>
      <div style={{background:"rgba(255,255,255,0.97)",borderRadius:20,padding:"40px 36px",width:360,boxShadow:"0 24px 60px rgba(0,0,0,0.4)"}}>
        <div style={{textAlign:"center",marginBottom:28}}>
          <div style={{width:60,height:60,borderRadius:16,margin:"0 auto 16px",background:"linear-gradient(135deg,#007AFF,#0051D4)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:28,boxShadow:"0 6px 20px rgba(0,122,255,0.35)"}}>⚗️</div>
          <div style={{fontSize:24,fontWeight:800,color:"#0A2240",letterSpacing:"-0.5px"}}>Fluxo de Liberação</div>
          <div style={{fontSize:13,color:"#8E8E93",marginTop:4}}>Gestão de Produtos — Controle da Qualidade</div>
        </div>
        {err&&<div style={{color:"#FF3B30",fontSize:13,textAlign:"center",background:"#FFF2F0",borderRadius:8,padding:"8px 12px",marginBottom:12}}>{err}</div>}
        <div style={{marginBottom:14}}><label style={{fontSize:13,fontWeight:600,color:"#3A3A3C",marginBottom:6,display:"block"}}>E-mail</label><input style={INP} type="email" placeholder="seu@empresa.com" value={email} onChange={e=>setEmail(e.target.value)} onKeyDown={e=>e.key==="Enter"&&go()} autoFocus/></div>
        <div style={{marginBottom:20}}><label style={{fontSize:13,fontWeight:600,color:"#3A3A3C",marginBottom:6,display:"block"}}>Senha</label><input style={INP} type="password" placeholder="••••••••" value={pw} onChange={e=>setPw(e.target.value)} onKeyDown={e=>e.key==="Enter"&&go()}/></div>
        <button style={{width:"100%",background:"linear-gradient(135deg,#007AFF,#0051D4)",color:"#fff",border:"none",borderRadius:12,padding:"13px",fontSize:16,fontWeight:600,cursor:"pointer",opacity:loading?0.7:1}} onClick={go} disabled={loading}>{loading?"Entrando…":"Entrar"}</button>
        <div style={{marginTop:20,fontSize:11,color:"#C7C7CC",textAlign:"center",lineHeight:1.7,borderTop:"1px solid #F2F2F7",paddingTop:14}}>Demo: qualidade.tecnica@empresa.com / 123456<br/>admin@empresa.com / admin</div>
      </div>
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// NAV MENU (hamburger)
// ─────────────────────────────────────────────────────────────────────────────
function NavMenu({page,setPage}){
  const [open,setOpen]=useState(false);
  const ref=useRef(null);
  useEffect(()=>{
    function h(e){if(ref.current&&!ref.current.contains(e.target))setOpen(false);}
    document.addEventListener("mousedown",h);return()=>document.removeEventListener("mousedown",h);
  },[]);
  const items=[
    {key:"pipeline",icon:"🔄",label:"Pipeline"},
    {key:"dashboard",icon:"📊",label:"Dashboard"},
    {key:"historico",icon:"📋",label:"Histórico"},
    {key:"configuracoes",icon:"⚙️",label:"Configurações"},
  ];
  return(
    <div ref={ref} style={{position:"relative"}}>
      <button onClick={()=>setOpen(o=>!o)} style={{background:"rgba(255,255,255,0.12)",border:"none",borderRadius:8,width:34,height:34,display:"flex",alignItems:"center",justifyContent:"center",cursor:"pointer",color:"#fff",fontSize:18,transition:"background 0.15s"}} title="Menu">
        {open?"✕":"☰"}
      </button>
      {open&&(
        <div style={{position:"absolute",top:42,left:0,background:"#fff",borderRadius:14,boxShadow:"0 8px 32px rgba(0,0,0,0.2)",padding:"6px",minWidth:190,zIndex:300}}>
          {items.map(it=>(
            <button key={it.key} onClick={()=>{setPage(it.key);setOpen(false);}} style={{display:"flex",alignItems:"center",gap:10,width:"100%",background:page===it.key?"rgba(0,122,255,0.08)":"transparent",border:"none",borderRadius:9,padding:"10px 12px",fontSize:14,fontWeight:page===it.key?700:500,color:page===it.key?"#007AFF":"#1C1C1E",cursor:"pointer",textAlign:"left"}}>
              <span style={{fontSize:18}}>{it.icon}</span>{it.label}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// PIPELINE BAR
// ─────────────────────────────────────────────────────────────────────────────
function PipelineBar({stageData,activeStage,onSelectStage,searchResult}){
  // searchResult: if set, show filtered counts per stage
  return(
    <div style={{background:"#fff",borderBottom:"1px solid #E5E5EA",overflowX:"auto"}}>
      <div style={{display:"flex",alignItems:"stretch",minWidth:900,padding:"0 12px"}}>
        {STAGES.map((stage,idx)=>{
          const active=activeStage===stage.id;
          const count=searchResult ? (searchResult[stage.id]||0) : (stageData[stage.id]?.length||0);
          const showBadge=count>0;
          return(
            <div key={stage.id} style={{display:"flex",alignItems:"stretch",flex:1}}>
              <div onClick={()=>onSelectStage(stage.id)} style={{flex:1,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"flex-start",cursor:"pointer",padding:"12px 4px 8px",position:"relative",borderBottom:active?"3px solid #007AFF":"3px solid transparent",background:active?"rgba(0,122,255,0.04)":"transparent",transition:"background 0.15s",minWidth:88}}>
                {showBadge&&<div style={{position:"absolute",top:6,right:"calc(50% - 22px)",background:searchResult?"#FF9500":"#FF3B30",color:"#fff",borderRadius:10,fontSize:9,fontWeight:700,padding:"1px 5px",lineHeight:"14px",height:14,minWidth:16,textAlign:"center"}}>{count}</div>}
                <div style={{width:28,height:28,borderRadius:"50%",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700,marginBottom:5,background:active?"#007AFF":"#E5E5EA",color:active?"#fff":"#8E8E93",boxShadow:active?"0 0 0 4px rgba(0,122,255,0.15)":"none",transition:"all 0.2s",flexShrink:0}}>{stage.id}</div>
                <div style={{fontSize:9,fontWeight:active?700:500,color:active?"#007AFF":"#6B6B6B",textAlign:"center",lineHeight:1.3,marginBottom:3}}>{stage.short}</div>
                {stage.dept&&<div style={{fontSize:8,color:active?"rgba(0,122,255,0.7)":"#ABABAB",textAlign:"center",lineHeight:1.3,fontWeight:500,maxWidth:84}}>{DEPT_FULL[stage.dept]||stage.dept}</div>}
              </div>
              {idx<STAGES.length-1&&<div style={{width:6,display:"flex",alignItems:"center",marginBottom:3}}><div style={{width:"100%",height:2,background:"#E5E5EA"}}/></div>}
            </div>
          );
        })}
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
    <div style={{background:"#fff",borderRadius:12,padding:"12px 14px",boxShadow:"0 1px 5px rgba(0,0,0,0.07)",marginBottom:14}}>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8}}>
        <span style={{fontSize:12,fontWeight:700,color:"#3A3A3C"}}>🔍 Busca</span>
        {hasAny&&<button onClick={onClear} style={{background:"none",border:"none",color:"#FF3B30",fontSize:12,fontWeight:600,cursor:"pointer"}}>Limpar</button>}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(190px,1fr))",gap:8}}>
        {fields.map(f=>(
          <div key={f.key}>
            <div style={{fontSize:10,fontWeight:600,color:"#8E8E93",marginBottom:3}}>{f.label}</div>
            <input value={filters[f.key]||""} onChange={e=>onChange({...filters,[f.key]:e.target.value})} placeholder={f.ph} style={{width:"100%",border:"1.5px solid #E5E5EA",borderRadius:7,padding:"6px 9px",fontSize:12,outline:"none",background:"#F9F9FB",boxSizing:"border-box"}}/>
          </div>
        ))}
      </div>
      {hasAny&&<div style={{marginTop:6,fontSize:10,color:"#8E8E93"}}>Separe múltiplos valores com espaço, vírgula ou ponto-e-vírgula</div>}
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

  const previewCols=["pedido","item","material","lote","ippn","deposito_sap","motivo_bloqueio"];
  return(
    <div>
      <div style={{fontSize:22,fontWeight:800,color:"#1C1C1E",letterSpacing:"-0.4px",marginBottom:4}}>Importar Tubos Bloqueados</div>
      <div style={{fontSize:13,color:"#8E8E93",marginBottom:20}}>Carregue um arquivo CSV, TXT ou XLS/XLSX para iniciar o fluxo na Etapa 1.</div>
      <div style={{border:`2px dashed ${drag?"#007AFF":"#C7C7CC"}`,borderRadius:16,padding:"40px 24px",textAlign:"center",background:drag?"rgba(0,122,255,0.04)":"#F9F9FB",cursor:"pointer",transition:"all 0.2s",marginBottom:18}} onDragOver={e=>{e.preventDefault();setDrag(true);}} onDragLeave={()=>setDrag(false)} onDrop={handleDrop} onClick={()=>!loading&&document.getElementById("fileInputImp").click()}>
        <div style={{fontSize:38,marginBottom:8}}>{loading?"⏳":"📂"}</div>
        <div style={{fontSize:14,fontWeight:700,color:"#3A3A3C",marginBottom:3}}>{loading?"Processando…":"Arraste o arquivo ou clique para selecionar"}</div>
        <div style={{fontSize:11,color:"#8E8E93"}}>CSV · TXT · XLS · XLSX</div>
        <input id="fileInputImp" type="file" accept=".csv,.txt,.xls,.xlsx" style={{display:"none"}} onChange={e=>processFile(e.target.files[0])}/>
      </div>
      {error&&<div style={{color:"#FF3B30",fontSize:13,background:"#FFF2F0",borderRadius:8,padding:"8px 12px",marginBottom:12}}>{error}</div>}
      {preview?(
        <div>
          <div style={{fontSize:13,color:"#34C759",fontWeight:700,marginBottom:10}}>✓ {preview.length} tubos encontrados</div>
          <div style={{overflowX:"auto",borderRadius:12,boxShadow:"0 1px 5px rgba(0,0,0,0.08)",background:"#fff",maxHeight:260,overflowY:"auto",marginBottom:14}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:600}}>
              <thead><tr>{previewCols.map(c=><th key={c} style={{background:"#0A2240",padding:"8px 10px",textAlign:"left",fontWeight:700,color:"#fff",borderBottom:"1px solid #1A3A5C",whiteSpace:"nowrap",position:"sticky",top:0}}>{COL_LABELS[c]||c}</th>)}</tr></thead>
              <tbody>{preview.slice(0,10).map((row,i)=><tr key={i}>{previewCols.map(c=><td key={c} style={{padding:"7px 10px",borderBottom:"1px solid #F2F2F7"}}>{row[c]||"—"}</td>)}</tr>)}</tbody>
            </table>
          </div>
          {preview.length>10&&<div style={{fontSize:11,color:"#8E8E93",marginBottom:10}}>Exibindo 10 de {preview.length} linhas</div>}
          <div style={{display:"flex",gap:10}}>
            <Btn onClick={()=>onImport(preview)}>📥 Importar {preview.length} Tubos</Btn>
            <Btn variant="secondary" onClick={()=>setPreview(null)}>Cancelar</Btn>
          </div>
        </div>
      ):(
        <div style={{background:"#fff",borderRadius:12,padding:"14px 16px",boxShadow:"0 1px 4px rgba(0,0,0,0.06)"}}>
          <div style={{fontSize:12,fontWeight:700,color:"#3A3A3C",marginBottom:6}}>Colunas esperadas (mapeamento automático por nome):</div>
          <div style={{fontSize:11,color:"#8E8E93",lineHeight:2}}>{Object.values(COL_LABELS).join(" · ")}</div>
        </div>
      )}
    </div>
  );
}

// ─────────────────────────────────────────────────────────────────────────────
// STAGE VIEW
// ─────────────────────────────────────────────────────────────────────────────
function StageView({stage,rows,user,onAdvance,onReturn,onComplete}){
  const [selLotes,setSelLotes]=useState(new Set());
  const [tratativas,setTratativas]=useState({});
  const [expanded,setExpanded]=useState(new Set());
  const [confirmAdv,setConfirmAdv]=useState(false);
  const [confirmRet,setConfirmRet]=useState(false);
  const [confirmDone,setConfirmDone]=useState(false);
  // Group 1 expand (Pedido/Item/Material/Desc)
  const [g1Expanded,setG1Expanded]=useState(new Set());
  // Group 2 expand (Última Ordem / QTS / Depósito / Motivo)
  const [g2Expanded,setG2Expanded]=useState(new Set());

  const isStage8 = stage.id===8;
  const nextStage = STAGES.find(s=>s.id===stage.id+1);
  const prevStage = STAGES.find(s=>s.id===stage.id-1);
  const groups = groupByLote(rows);

  const canInteract = user.dept==="Admin" || (user.allowedStages||[]).includes(stage.id);

  function toggleSel(lote){const n=new Set(selLotes);n.has(lote)?n.delete(lote):n.add(lote);setSelLotes(n);}
  function toggleAll(c){setSelLotes(c?new Set(groups.map(g=>g.lote)):new Set());}
  function toggleExp(lote){const n=new Set(expanded);n.has(lote)?n.delete(lote):n.add(lote);setExpanded(n);}
  function toggleG1(lote){const n=new Set(g1Expanded);n.has(lote)?n.delete(lote):n.add(lote);setG1Expanded(n);}
  function toggleG2(lote){const n=new Set(g2Expanded);n.has(lote)?n.delete(lote):n.add(lote);setG2Expanded(n);}

  function buildMoved(loteSet){
    return rows.filter(r=>loteSet.has(r.lote||r._id)).map(r=>({
      ...r,
      tratativa:tratativas[r.lote||r._id]??r.tratativa,
      history:[...(r.history||[]),{
        stage:stage.id,stageLabel:stage.label,
        user:user.name,dept:user.dept,
        tratativa:tratativas[r.lote||r._id]??r.tratativa,
        date:new Date().toLocaleString("pt-BR"),
        dateMs:Date.now(),
      }],
      stageEnteredMs:Date.now(),
    }));
  }

  function doAdvance(){const moved=buildMoved(selLotes);const rem=rows.filter(r=>!selLotes.has(r.lote||r._id));onAdvance(moved,rem);setSelLotes(new Set());setTratativas({});setConfirmAdv(false);}
  function doReturn(){const moved=buildMoved(selLotes);const rem=rows.filter(r=>!selLotes.has(r.lote||r._id));onReturn(moved,rem);setSelLotes(new Set());setTratativas({});setConfirmRet(false);}
  function doComplete(){const moved=buildMoved(selLotes);const rem=rows.filter(r=>!selLotes.has(r.lote||r._id));onComplete(moved,rem);setSelLotes(new Set());setTratativas({});setConfirmDone(false);}

  const selCount=rows.filter(r=>selLotes.has(r.lote||r._id)).length;
  const selLotesCount=selLotes.size;

  if(rows.length===0) return(
    <div>
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14}}><div style={{fontSize:20,fontWeight:800,color:"#1C1C1E"}}>{stage.label}</div><DeptTag dept={stage.dept}/></div>
      <div style={{textAlign:"center",padding:"52px 24px",background:"#fff",borderRadius:16,boxShadow:"0 1px 4px rgba(0,0,0,0.07)"}}><div style={{fontSize:40,marginBottom:10}}>📭</div><div style={{fontSize:15,fontWeight:700,color:"#3A3A3C",marginBottom:4}}>Nenhum tubo nesta etapa</div><div style={{fontSize:13,color:"#8E8E93"}}>Os tubos aparecerão aqui quando avançarem para esta etapa.</div></div>
    </div>
  );

  return(
    <div>
      {/* Top bar */}
      <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",marginBottom:12,flexWrap:"wrap",gap:10}}>
        <div>
          <div style={{fontSize:20,fontWeight:800,color:"#1C1C1E",letterSpacing:"-0.3px",marginBottom:5}}>{stage.label}</div>
          <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
            <DeptTag dept={stage.dept}/>
            <span style={{fontSize:12,color:"#8E8E93"}}>{groups.length} lote{groups.length!==1?"s":""} · {rows.length} tubo{rows.length!==1?"s":""}</span>
            {selLotes.size>0&&<span style={{fontSize:12,fontWeight:700,color:"#007AFF",background:"rgba(0,122,255,0.08)",borderRadius:8,padding:"3px 10px"}}>{selLotesCount} lote{selLotesCount!==1?"s":""} ({selCount} tubo{selCount!==1?"s":""})</span>}
            {!canInteract&&<span style={{fontSize:11,color:"#FF3B30",background:"#FFF2F0",borderRadius:6,padding:"2px 8px",fontWeight:600}}>⚠ Sem permissão nesta etapa</span>}
          </div>
        </div>
        {/* Action buttons */}
        {canInteract&&(
          <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
            {stage.id>1&&onReturn&&<Btn variant="warning" disabled={selLotes.size===0} onClick={()=>setConfirmRet(true)}>← Retornar: {prevStage?.short}</Btn>}
            {!isStage8&&<Btn disabled={selLotes.size===0} onClick={()=>setConfirmAdv(true)}>→ {nextStage?.short} ({selLotesCount})</Btn>}
            {isStage8&&<Btn variant="danger" disabled={selLotes.size===0} onClick={()=>setConfirmDone(true)}>✅ Concluir ({selLotesCount})</Btn>}
          </div>
        )}
      </div>

      {/* Table */}
      <div style={{borderRadius:12,boxShadow:"0 1px 6px rgba(0,0,0,0.08)",background:"#fff",overflowX:"auto",overflowY:"auto",maxHeight:"calc(100vh - 320px)",minHeight:180}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:1000}}>
          <thead>
            <tr>
              {canInteract&&<th style={{...TH_DARK,width:34}}><input type="checkbox" style={{accentColor:"#5AC8FA",cursor:"pointer",width:15,height:15}} checked={selLotes.size===groups.length&&groups.length>0} onChange={e=>toggleAll(e.target.checked)}/></th>}
              <th style={{...TH_DARK,width:30}} title="Expandir IPPNs">▶</th>
              <th style={{...TH_DARK,width:38}} title="Grupo 1: Pedido/Item">G1</th>
              <th style={{...TH_DARK,width:38}} title="Grupo 2: Ordem/QTS">G2</th>
              <th style={TH_DARK}>Lote</th>
              <th style={TH_DARK}>Tubos</th>
              <th style={TH_DARK}>IPPNs</th>
              <th style={TH_DARK}>Pedido</th>
              <th style={TH_DARK}>Item</th>
              <th style={TH_DARK}>Material</th>
              <th style={TH_DARK}>Última Ordem</th>
              <th style={TH_DARK}>Depósito SAP</th>
              <th style={TH_DARK}>Motivo Bloqueio</th>
              <th style={TH_GREEN}>Tratativa do Lote</th>
              {stage.id>1&&<th style={TH_DARK}>Histórico</th>}
            </tr>
          </thead>
          <tbody>
            {groups.map((group,gi)=>{
              const isExp=expanded.has(group.lote);
              const isSel=selLotes.has(group.lote);
              const isG1=g1Expanded.has(group.lote);
              const isG2=g2Expanded.has(group.lote);
              const fr=group.rows[0];
              const ippnList=group.rows.map(r=>r.ippn).filter(Boolean).join(", ");
              const alt=gi%2===1;
              const bg=isSel?"rgba(0,122,255,0.06)":(alt?"#F8F9FB":"#fff");

              return[
                <tr key={`g-${group.lote}`} style={{background:bg}}>
                  {canInteract&&<td style={{...TD(false),background:bg,textAlign:"center"}}><input type="checkbox" style={{accentColor:"#007AFF",cursor:"pointer",width:15,height:15}} checked={isSel} onChange={()=>toggleSel(group.lote)}/></td>}
                  {/* Expand IPPNs */}
                  <td style={{...TD(false),background:bg,textAlign:"center"}}><button onClick={()=>toggleExp(group.lote)} style={{background:"none",border:"none",cursor:"pointer",fontSize:12,color:"#007AFF",fontWeight:700,padding:"1px 3px"}}>{isExp?"▼":"▶"}</button></td>
                  {/* G1 toggle */}
                  <td style={{...TD(false),background:bg,textAlign:"center"}} title="Pedido · Item · Material · Descrição"><button onClick={()=>toggleG1(group.lote)} style={{background:isG1?"#007AFF":"#E5E5EA",border:"none",borderRadius:4,cursor:"pointer",fontSize:10,color:isG1?"#fff":"#555",padding:"2px 5px",fontWeight:700}}>G1</button></td>
                  {/* G2 toggle */}
                  <td style={{...TD(false),background:bg,textAlign:"center"}} title="Ordem · QTS · Depósito · Motivo"><button onClick={()=>toggleG2(group.lote)} style={{background:isG2?"#FF9500":"#E5E5EA",border:"none",borderRadius:4,cursor:"pointer",fontSize:10,color:isG2?"#fff":"#555",padding:"2px 5px",fontWeight:700}}>G2</button></td>
                  <td style={{...TD(false),background:bg,fontWeight:700,color:"#1C1C1E"}}>{group.lote||"—"}</td>
                  <td style={{...TD(false),background:bg}}><span style={{background:"#E8F4FD",color:"#1A6FA8",borderRadius:5,padding:"2px 7px",fontSize:11,fontWeight:700}}>{group.rows.length}</span></td>
                  <td style={{...TD(false),background:bg,maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:"#555"}}>{ippnList||"—"}</td>
                  <td style={{...TD(false),background:bg}}>{group.pedido||"—"}</td>
                  <td style={{...TD(false),background:bg}}>{group.item||"—"}</td>
                  <td style={{...TD(false),background:bg}}>{group.material||"—"}</td>
                  <td style={{...TD(false),background:bg}}>{group.ultima_ordem||"—"}</td>
                  <td style={{...TD(false),background:bg}}>{group.deposito_sap||"—"}</td>
                  <td style={{...TD(false),background:bg}}>{group.motivo_bloqueio||"—"}</td>
                  <td style={{...TD(false),background:gi%2===0?"rgba(26,90,42,0.03)":"rgba(26,90,42,0.06)"}}>
                    {canInteract?(<input type="text" style={{border:"1.5px solid #E5E5EA",borderRadius:7,padding:"5px 9px",fontSize:11,width:"100%",outline:"none",background:"#fff",minWidth:160}} placeholder="Tratativa do lote…" value={tratativas[group.lote]??group.tratativa??""} onChange={e=>setTratativas(t=>({...t,[group.lote]:e.target.value}))}/>):(<span style={{fontSize:11,color:"#555"}}>{group.tratativa||"—"}</span>)}
                  </td>
                  {stage.id>1&&<td style={{...TD(false),background:bg}}>
                    {fr.history?.length>0?(<details><summary style={{fontSize:10,color:"#007AFF",cursor:"pointer",fontWeight:600}}>{fr.history.length} etapa{fr.history.length!==1?"s":""}</summary><div style={{marginTop:5}}>{fr.history.map((h,hi)=><div key={hi} style={{fontSize:10,color:"#3A3A3C",marginBottom:5,borderLeft:"2px solid #E5E5EA",paddingLeft:7,lineHeight:1.55}}><strong>{h.stageLabel}</strong><br/>{h.user}{h.tratativa&&<><br/><em style={{color:"#555"}}>"{h.tratativa}"</em></>}<br/><span style={{color:"#C7C7CC"}}>{h.date}</span></div>)}</div></details>):(<span style={{color:"#C7C7CC",fontSize:11}}>—</span>)}
                  </td>}
                </tr>,
                // G1 row: Pedido / Item / Material / Descrição
                ...(isG1?[<tr key={`g1-${group.lote}`} style={{background:"#FFF9F0"}}><td colSpan={canInteract?4:3}></td><td colSpan={2} style={{padding:"8px 11px",fontSize:11,color:"#B45309",fontWeight:700}}>📦 Grupo 1</td><td style={{padding:"8px 11px",fontSize:11,color:"#555"}}><span style={{fontWeight:600}}>Ped:</span> {group.pedido||"—"}</td><td style={{padding:"8px 11px",fontSize:11,color:"#555"}}><span style={{fontWeight:600}}>Item:</span> {group.item||"—"}</td><td style={{padding:"8px 11px",fontSize:11,color:"#555"}}><span style={{fontWeight:600}}>Mat:</span> {group.material||"—"}</td><td colSpan={3} style={{padding:"8px 11px",fontSize:11,color:"#555"}}><span style={{fontWeight:600}}>Desc:</span> {fr.descricao||"—"}</td><td colSpan={2}></td>{stage.id>1&&<td></td>}</tr>]:[]),
                // G2 row: Última Ordem / QTS / Depósito / Motivo
                ...(isG2?[<tr key={`g2-${group.lote}`} style={{background:"#FFF3E0"}}><td colSpan={canInteract?4:3}></td><td colSpan={2} style={{padding:"8px 11px",fontSize:11,color:"#E8890A",fontWeight:700}}>🗂 Grupo 2</td><td colSpan={2} style={{padding:"8px 11px",fontSize:11,color:"#555"}}><span style={{fontWeight:600}}>Última Ord:</span> {group.ultima_ordem||"—"} &nbsp;|&nbsp; <span style={{fontWeight:600}}>QTS:</span> {group.qualidade_qts||fr.qualidade_qts||"—"}</td><td style={{padding:"8px 11px",fontSize:11,color:"#555"}}><span style={{fontWeight:600}}>Depósito:</span> {group.deposito_sap||"—"}</td><td colSpan={3} style={{padding:"8px 11px",fontSize:11,color:"#555"}}><span style={{fontWeight:600}}>Motivo:</span> {group.motivo_bloqueio||"—"} — {fr.motivo_bloqueio_texto||"—"}</td><td></td>{stage.id>1&&<td></td>}</tr>]:[]),
                // IPPN expanded rows
                ...(isExp?group.rows.map((row,ri)=><tr key={`exp-${row._id}`} style={{background:"#F0F7FF"}}>{canInteract&&<td></td>}<td style={{padding:"7px 8px",borderBottom:"1px solid #E8F4FD",textAlign:"center"}}><span style={{fontSize:9,color:"#8E8E93"}}>└</span></td><td colSpan={3} style={{padding:"7px 11px",borderBottom:"1px solid #E8F4FD",fontSize:11,color:"#1A6FA8",fontWeight:600}}>IPPN {ri+1}: {row.ippn||"—"}</td><td style={{padding:"7px 11px",borderBottom:"1px solid #E8F4FD",fontSize:11,color:"#555"}}><span style={{fontWeight:600}}>Cassete:</span> {row.num_cassete||"—"}</td><td colSpan={2} style={{padding:"7px 11px",borderBottom:"1px solid #E8F4FD",fontSize:11,color:"#555"}}><span style={{fontWeight:600}}>Data Bloqueio:</span> {row.data_bloqueio||"—"}</td><td colSpan={2} style={{padding:"7px 11px",borderBottom:"1px solid #E8F4FD",fontSize:11,color:"#555"}}><span style={{fontWeight:600}}>Razão:</span> {row.razao_bloq||"—"}</td><td colSpan={3} style={{borderBottom:"1px solid #E8F4FD"}}></td>{stage.id>1&&<td style={{borderBottom:"1px solid #E8F4FD"}}></td>}</tr>):[]),
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

// ─────────────────────────────────────────────────────────────────────────────
// DASHBOARD
// ─────────────────────────────────────────────────────────────────────────────
function Dashboard({stageData,historyRows,onSelectStage}){
  const allActive=Object.values(stageData).flat();
  const total=allActive.length;
  const etapa1=stageData[1]?.length||0;
  const inFlow=[2,3,4,5,6,7].reduce((s,id)=>s+(stageData[id]?.length||0),0);
  const pendExec=stageData[8]?.length||0;
  const concluded=historyRows.length;

  const byDep={};
  allActive.forEach(r=>{const d=r.deposito_sap||"Sem Depósito";byDep[d]=(byDep[d]||0)+1;});

  const byDept={};
  STAGES.filter(s=>s.id!==8&&s.dept).forEach(s=>{const cnt=stageData[s.id]?.length||0;if(!byDept[s.dept])byDept[s.dept]=0;byDept[s.dept]+=cnt;});

  // ─ Timing stats ─
  const {stageAvg,deptAvg}=computeTimingStats(stageData,historyRows);

  const kpis=[
    {label:"Total Bloqueado",  value:total,     color:"#FF3B30"},
    {label:"Análise Pendente", value:etapa1,    color:"#007AFF",sid:1},
    {label:"Em Fluxo",         value:inFlow,    color:"#FF9500"},
    {label:"Pend. Execução",   value:pendExec,  color:"#8E8E93",sid:8},
    {label:"Concluídos",       value:concluded, color:"#34C759",sid:"hist"},
  ];

  const hasTimingData=Object.keys(stageAvg).length>0||Object.keys(deptAvg).length>0;

  return(
    <div style={{maxWidth:1200,margin:"0 auto",padding:"18px 24px 40px"}}>
      <div style={{fontSize:22,fontWeight:800,color:"#1C1C1E",letterSpacing:"-0.4px",marginBottom:4}}>Dashboard</div>
      <div style={{fontSize:13,color:"#8E8E93",marginBottom:20}}>Visão geral do fluxo de liberação</div>

      {/* KPIs */}
      <div style={{display:"flex",gap:12,marginBottom:24,flexWrap:"wrap"}}>
        {kpis.map(k=>(
          <div key={k.label} style={{background:"#fff",borderRadius:16,padding:"16px 18px",flex:1,minWidth:130,boxShadow:"0 1px 6px rgba(0,0,0,0.08)",borderLeft:`4px solid ${k.color}`,cursor:k.sid?"pointer":"default"}} onClick={()=>k.sid&&onSelectStage(k.sid)}>
            <div style={{fontSize:10,color:"#8E8E93",fontWeight:600,marginBottom:4,textTransform:"uppercase",letterSpacing:"0.5px"}}>{k.label}</div>
            <div style={{fontSize:34,fontWeight:900,color:k.color,letterSpacing:"-1.5px",lineHeight:1}}>{k.value}</div>
          </div>
        ))}
      </div>

      {/* Stage cards */}
      <div style={{fontSize:13,fontWeight:700,color:"#3A3A3C",marginBottom:10}}>Tubos por etapa</div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(138px,1fr))",gap:10,marginBottom:24}}>
        {STAGES.map((stage,idx)=>{
          const count=stageData[stage.id]?.length||0;
          const color=STAGE_COLORS[idx];
          const avgMs=stageAvg[stage.id];
          return(
            <div key={stage.id} style={{background:"#fff",borderRadius:12,padding:"13px",boxShadow:"0 1px 4px rgba(0,0,0,0.07)",borderTop:`3px solid ${color}`,cursor:"pointer",transition:"transform 0.15s,box-shadow 0.15s"}} onClick={()=>onSelectStage(stage.id)} onMouseEnter={e=>{e.currentTarget.style.transform="translateY(-2px)";e.currentTarget.style.boxShadow="0 5px 14px rgba(0,0,0,0.12)";}} onMouseLeave={e=>{e.currentTarget.style.transform="";e.currentTarget.style.boxShadow="0 1px 4px rgba(0,0,0,0.07)";}}>
              <div style={{fontSize:28,fontWeight:900,color:"#1C1C1E",letterSpacing:"-1px"}}>{count}</div>
              <div style={{fontSize:9,fontWeight:800,color,marginTop:2,textTransform:"uppercase",letterSpacing:"0.4px"}}>ETAPA {stage.id}</div>
              <div style={{fontSize:10,color:"#6B6B6B",marginTop:2,lineHeight:1.3}}>{stage.short}</div>
              {avgMs&&<div style={{fontSize:9,color:"#8E8E93",marginTop:4}}>⏱ Média: {formatDuration(avgMs)}</div>}
              {count>0&&<div style={{fontSize:9,color:"#007AFF",fontWeight:700,marginTop:5}}>Ver →</div>}
            </div>
          );
        })}
      </div>

      {/* Timing section */}
      <div style={{fontSize:13,fontWeight:700,color:"#3A3A3C",marginBottom:10}}>⏱ Tempo médio por etapa</div>
      {!hasTimingData?(
        <div style={{background:"#fff",borderRadius:12,padding:"20px",boxShadow:"0 1px 4px rgba(0,0,0,0.07)",marginBottom:24,fontSize:13,color:"#8E8E93",textAlign:"center"}}>
          Dados de tempo disponíveis após os primeiros tubos avançarem entre etapas
        </div>
      ):(
        <div style={{background:"#fff",borderRadius:14,padding:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.08)",marginBottom:24}}>
          {/* Per-stage bar chart */}
          <div style={{marginBottom:16}}>
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
                    <div style={{width:`${pct}%`,height:"100%",background:color,borderRadius:4,transition:"width 0.6s"}}/>
                  </div>
                </div>
              );
            })}
          </div>
          {/* Per-dept */}
          {Object.keys(deptAvg).length>0&&(
            <div style={{borderTop:"1px solid #F2F2F7",paddingTop:14}}>
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
                      <div style={{width:`${pct}%`,height:"100%",background:c.fg,borderRadius:4,transition:"width 0.6s"}}/>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
        </div>
      )}

      {/* By dept pending + by deposit */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:20}}>
        <div style={{background:"#fff",borderRadius:14,padding:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.08)"}}>
          <div style={{fontSize:13,fontWeight:700,color:"#3A3A3C",marginBottom:12}}>Pendentes por Departamento</div>
          {!Object.keys(byDept).length?<div style={{fontSize:12,color:"#8E8E93"}}>Nenhum tubo em fluxo</div>:Object.entries(byDept).map(([dept,cnt])=>{
            const c=DEPT_COLORS[dept]||{bg:"#F2F2F7",fg:"#3A3A3C"};
            const pct=total>0?Math.round(cnt/total*100):0;
            return(<div key={dept} style={{marginBottom:11}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}><span style={{fontSize:11,fontWeight:600,color:"#3A3A3C"}}>{DEPT_FULL[dept]||dept}</span><span style={{fontSize:12,fontWeight:800,color:c.fg}}>{cnt}</span></div><div style={{height:6,background:"#F2F2F7",borderRadius:3,overflow:"hidden"}}><div style={{width:`${pct}%`,height:"100%",background:c.fg,borderRadius:3,transition:"width 0.5s"}}/></div></div>);
          })}
        </div>
        <div style={{background:"#fff",borderRadius:14,padding:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.08)"}}>
          <div style={{fontSize:13,fontWeight:700,color:"#3A3A3C",marginBottom:12}}>Bloqueio por Depósito</div>
          {!Object.keys(byDep).length?<div style={{fontSize:12,color:"#8E8E93"}}>Nenhum tubo</div>:Object.entries(byDep).sort((a,b)=>b[1]-a[1]).map(([dep,cnt],i)=>{
            const clrs=["#007AFF","#34C759","#FF9500","#AF52DE","#FF2D55","#5AC8FA"];
            const col=clrs[i%clrs.length];
            const pct=total>0?Math.round(cnt/total*100):0;
            return(<div key={dep} style={{marginBottom:11}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}><span style={{fontSize:11,fontWeight:600,color:"#3A3A3C"}}>{dep}</span><span style={{fontSize:12,fontWeight:800,color:col}}>{cnt}</span></div><div style={{height:6,background:"#F2F2F7",borderRadius:3,overflow:"hidden"}}><div style={{width:`${pct}%`,height:"100%",background:col,borderRadius:3,transition:"width 0.5s"}}/></div></div>);
          })}
        </div>
      </div>

      {total===0&&concluded===0&&(
        <div style={{textAlign:"center",padding:"52px 24px",background:"#fff",borderRadius:16,boxShadow:"0 1px 4px rgba(0,0,0,0.07)"}}>
          <div style={{fontSize:40,marginBottom:10}}>📋</div>
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
  const [expanded,setExpanded]=useState(new Set());
  function toggleExp(lote){const n=new Set(expanded);n.has(lote)?n.delete(lote):n.add(lote);setExpanded(n);}

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

  return(
    <div style={{maxWidth:1200,margin:"0 auto",padding:"18px 24px 40px"}}>
      <div style={{fontSize:22,fontWeight:800,color:"#1C1C1E",letterSpacing:"-0.4px",marginBottom:4}}>Histórico</div>
      <div style={{fontSize:13,color:"#8E8E93",marginBottom:16}}>Tubos que completaram o fluxo de liberação.</div>

      {/* Search */}
      <div style={{background:"#fff",borderRadius:12,padding:"12px 14px",boxShadow:"0 1px 5px rgba(0,0,0,0.07)",marginBottom:14}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8}}><span style={{fontSize:12,fontWeight:700,color:"#3A3A3C"}}>🔍 Pesquisar Histórico</span>{hasFilters&&<button onClick={()=>setFilters({pedido:"",item:"",material:"",lotes:"",ippns:""})} style={{background:"none",border:"none",color:"#FF3B30",fontSize:12,fontWeight:600,cursor:"pointer"}}>Limpar</button>}</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(180px,1fr))",gap:8}}>
          {searchFields.map(f=><div key={f.key}><div style={{fontSize:10,fontWeight:600,color:"#8E8E93",marginBottom:3}}>{f.label}</div><input value={filters[f.key]} onChange={e=>setFilters(v=>({...v,[f.key]:e.target.value}))} placeholder={f.ph} style={{width:"100%",border:"1.5px solid #E5E5EA",borderRadius:7,padding:"6px 9px",fontSize:12,outline:"none",background:"#F9F9FB",boxSizing:"border-box"}}/></div>)}
        </div>
        {hasFilters&&<div style={{marginTop:6,fontSize:10,color:"#8E8E93"}}>Separe múltiplos valores com espaço, vírgula ou ponto-e-vírgula</div>}
      </div>

      {historyRows.length===0?(
        <div style={{textAlign:"center",padding:"52px 24px",background:"#fff",borderRadius:16,boxShadow:"0 1px 4px rgba(0,0,0,0.07)"}}>
          <div style={{fontSize:40,marginBottom:10}}>📂</div>
          <div style={{fontSize:15,fontWeight:700,color:"#3A3A3C",marginBottom:4}}>Nenhum registro ainda</div>
          <div style={{fontSize:13,color:"#8E8E93"}}>Os tubos concluídos na Etapa 8 aparecerão aqui.</div>
        </div>
      ):(
        <div>
          <div style={{fontSize:12,color:"#8E8E93",marginBottom:10}}>{groups.length} lote{groups.length!==1?"s":""} · {filtered.length} tubo{filtered.length!==1?"s":""}{hasFilters&&` (filtrado de ${historyRows.length})`}</div>
          <div style={{borderRadius:12,boxShadow:"0 1px 6px rgba(0,0,0,0.08)",background:"#fff",overflowX:"auto",overflowY:"auto",maxHeight:"calc(100vh - 340px)"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:800}}>
              <thead>
                <tr>
                  <th style={{...TH_DARK,width:30}}></th>
                  {["Lote","Tubos","IPPNs","Pedido","Item","Material","Depósito","Motivo Bloqueio","Data Bloqueio","Histórico"].map(h=><th key={h} style={TH_DARK}>{h}</th>)}
                </tr>
              </thead>
              <tbody>
                {groups.map((group,gi)=>{
                  const isExp=expanded.has(group.lote);
                  const fr=group.rows[0];
                  const ippnList=group.rows.map(r=>r.ippn).filter(Boolean).join(", ");
                  const alt=gi%2===1;
                  return[
                    <tr key={`h-${group.lote}`} style={{background:alt?"#F8F9FB":"#fff"}}>
                      <td style={{...TD(alt),textAlign:"center"}}><button onClick={()=>toggleExp(group.lote)} style={{background:"none",border:"none",cursor:"pointer",fontSize:12,color:"#007AFF",fontWeight:700}}>{isExp?"▼":"▶"}</button></td>
                      <td style={{...TD(alt),fontWeight:700,color:"#1C1C1E"}}>{group.lote||"—"}</td>
                      <td style={TD(alt)}><span style={{background:"#EDF7EE",color:"#1A7A3A",borderRadius:5,padding:"2px 7px",fontSize:11,fontWeight:700}}>{group.rows.length}</span></td>
                      <td style={{...TD(alt),maxWidth:150,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{ippnList||"—"}</td>
                      <td style={TD(alt)}>{group.pedido||"—"}</td>
                      <td style={TD(alt)}>{group.item||"—"}</td>
                      <td style={TD(alt)}>{group.material||"—"}</td>
                      <td style={TD(alt)}>{group.deposito_sap||"—"}</td>
                      <td style={TD(alt)}>{group.motivo_bloqueio||"—"}</td>
                      <td style={TD(alt)}>{fr.data_bloqueio||"—"}</td>
                      <td style={TD(alt)}>
                        {fr.history?.length>0?<details><summary style={{fontSize:10,color:"#34C759",cursor:"pointer",fontWeight:700}}>✓ {fr.history.length} etapa{fr.history.length!==1?"s":""}</summary><div style={{marginTop:5}}>{fr.history.map((h,hi)=><div key={hi} style={{fontSize:10,color:"#3A3A3C",marginBottom:5,borderLeft:"2px solid #34C759",paddingLeft:7,lineHeight:1.55}}><strong>{h.stageLabel}</strong><br/>{h.user}{h.tratativa&&<><br/><em>"{h.tratativa}"</em></>}<br/><span style={{color:"#C7C7CC"}}>{h.date}</span></div>)}</div></details>:<span style={{color:"#C7C7CC",fontSize:11}}>—</span>}
                      </td>
                    </tr>,
                    ...(isExp?group.rows.map((row,ri)=><tr key={`he-${row._id}`} style={{background:"#F0F7FF"}}><td style={{padding:"6px 8px",borderBottom:"1px solid #E8F4FD",textAlign:"center"}}><span style={{fontSize:9,color:"#8E8E93"}}>└</span></td><td colSpan={3} style={{padding:"6px 11px",borderBottom:"1px solid #E8F4FD",fontSize:11,color:"#1A6FA8",fontWeight:600}}>IPPN {ri+1}: {row.ippn||"—"}</td><td style={{padding:"6px 11px",borderBottom:"1px solid #E8F4FD",fontSize:11}}>{row.pedido||"—"}</td><td style={{padding:"6px 11px",borderBottom:"1px solid #E8F4FD",fontSize:11}}>{row.item||"—"}</td><td style={{padding:"6px 11px",borderBottom:"1px solid #E8F4FD",fontSize:11}}>{row.material||"—"}</td><td colSpan={4} style={{borderBottom:"1px solid #E8F4FD"}}></td></tr>):[]),
                  ];
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
function ConfigPage({users,setUsers}){
  const [tab,setTab]=useState("users");
  const [editing,setEditing]=useState(null); // null or user obj
  const [form,setForm]=useState({});
  const [showDel,setShowDel]=useState(null);
  const [saveMsg,setSaveMsg]=useState("");

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

  const tabStyle=(active)=>({padding:"8px 18px",borderRadius:8,border:"none",fontSize:13,fontWeight:active?700:500,color:active?"#1C1C1E":"#8E8E93",background:active?"#fff":"transparent",cursor:"pointer",boxShadow:active?"0 1px 4px rgba(0,0,0,0.1)":"none"});

  return(
    <div style={{maxWidth:1000,margin:"0 auto",padding:"18px 24px 40px"}}>
      <div style={{fontSize:22,fontWeight:800,color:"#1C1C1E",letterSpacing:"-0.4px",marginBottom:16}}>Configurações</div>
      <div style={{display:"flex",gap:2,background:"#E5E5EA",borderRadius:10,padding:3,marginBottom:20,width:"fit-content"}}>
        <button style={tabStyle(tab==="users")} onClick={()=>setTab("users")}>👤 Usuários</button>
        <button style={tabStyle(tab==="perms")} onClick={()=>setTab("perms")}>🔐 Permissões</button>
      </div>

      {saveMsg&&<div style={{background:"#EDF7EE",color:"#1A7A3A",borderRadius:8,padding:"8px 14px",marginBottom:14,fontSize:13,fontWeight:600}}>{saveMsg}</div>}

      {tab==="users"&&(
        <div>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
            <div style={{fontSize:14,fontWeight:700,color:"#3A3A3C"}}>Usuários cadastrados ({users.length})</div>
            <Btn onClick={openNew}>+ Novo Usuário</Btn>
          </div>
          <div style={{background:"#fff",borderRadius:14,boxShadow:"0 1px 6px rgba(0,0,0,0.08)",overflow:"hidden"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
              <thead><tr>{["Nome","E-mail","Departamento","Etapas","Status","Ações"].map(h=><th key={h} style={{background:"#0A2240",padding:"10px 12px",textAlign:"left",fontWeight:700,color:"#fff",fontSize:11,borderBottom:"2px solid #1A3A5C"}}>{h}</th>)}</tr></thead>
              <tbody>
                {users.map((u,i)=>(
                  <tr key={u.id} style={{background:i%2?"#F8F9FB":"#fff"}}>
                    <td style={{padding:"10px 12px",borderBottom:"1px solid #F0F0F5",fontWeight:600}}>{u.name}</td>
                    <td style={{padding:"10px 12px",borderBottom:"1px solid #F0F0F5",color:"#555"}}>{u.email}</td>
                    <td style={{padding:"10px 12px",borderBottom:"1px solid #F0F0F5"}}><DeptTag dept={u.dept}/></td>
                    <td style={{padding:"10px 12px",borderBottom:"1px solid #F0F0F5",fontSize:11}}>{(u.allowedStages||[]).map(s=>`E${s}`).join(", ")||"—"}</td>
                    <td style={{padding:"10px 12px",borderBottom:"1px solid #F0F0F5"}}><span style={{background:u.active!==false?"#EDF7EE":"#F2F2F7",color:u.active!==false?"#1A7A3A":"#8E8E93",borderRadius:6,padding:"2px 8px",fontSize:11,fontWeight:600}}>{u.active!==false?"Ativo":"Inativo"}</span></td>
                    <td style={{padding:"10px 12px",borderBottom:"1px solid #F0F0F5"}}>
                      <div style={{display:"flex",gap:6}}>
                        <Btn small variant="secondary" onClick={()=>openEdit(u)}>Editar</Btn>
                        <Btn small variant="danger" onClick={()=>setShowDel(u.id)}>Excluir</Btn>
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
          <div style={{background:"#fff",borderRadius:14,padding:"16px",boxShadow:"0 1px 6px rgba(0,0,0,0.08)"}}>
            {Object.entries(DEPT_DEFAULT_STAGES).filter(([d])=>d!=="Admin").map(([dept,stages])=>{
              const c=DEPT_COLORS[dept]||{bg:"#F2F2F7",fg:"#3A3A3C"};
              return(
                <div key={dept} style={{marginBottom:16,paddingBottom:16,borderBottom:"1px solid #F2F2F7"}}>
                  <div style={{marginBottom:8}}><DeptTag dept={dept}/></div>
                  <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
                    {STAGES.map(s=>{
                      const allowed=stages.includes(s.id);
                      return(<div key={s.id} style={{background:allowed?c.bg:"#F2F2F7",color:allowed?c.fg:"#C7C7CC",borderRadius:7,padding:"4px 10px",fontSize:11,fontWeight:600}}>{allowed?"✓":""} E{s.id}: {s.short}</div>);
                    })}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* Edit / New user modal */}
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
                return(<button key={s.id} type="button" onClick={()=>toggleStage(s.id)} style={{background:allowed?"#007AFF":"#F2F2F7",color:allowed?"#fff":"#555",border:"none",borderRadius:7,padding:"5px 11px",fontSize:11,fontWeight:600,cursor:"pointer"}}>E{s.id}: {s.short}</button>);
              })}
            </div>
          </div>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <input type="checkbox" checked={form.active!==false} onChange={e=>setForm(f=>({...f,active:e.target.checked}))} style={{accentColor:"#007AFF",width:15,height:15}}/>
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
function PipelinePage({stageData,setStageData,user,historyRows,setHistoryRows,showToast}){
  const [activeStage,setActiveStage]=useState(1);
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
      <div style={{maxWidth:1200,margin:"0 auto",padding:"14px 24px 40px"}}>
        <SearchBar filters={filters} onChange={setFilters} onClear={()=>setFilters({pedido_item:"",lotes:"",ippns:"",deposito:""})}/>

        {/* Search results summary */}
        {hasSearch&&searchResult&&(
          <div style={{background:"#fff",borderRadius:10,padding:"10px 14px",marginBottom:14,boxShadow:"0 1px 4px rgba(0,0,0,0.07)",fontSize:12,color:"#3A3A3C"}}>
            <span style={{fontWeight:700}}>Resultado da busca: </span>
            {STAGES.map(s=>(searchResult[s.id]>0)&&<span key={s.id} style={{marginRight:10,background:"rgba(255,150,0,0.1)",color:"#B45309",borderRadius:5,padding:"2px 8px",fontWeight:600}}>E{s.id}: {searchResult[s.id]}</span>)}
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
            />
            {activeStage===1&&stageData[1].length>0&&(
              <div style={{marginTop:12}}>
                <Btn variant="secondary" small onClick={()=>{if(window.confirm("Remove todos os tubos da Etapa 1. Continuar?"))setStageData(d=>({...d,1:[]}));}}>↩ Novo arquivo</Btn>
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
export default function App(){
  const [users,setUsers]=useState(DEFAULT_USERS);
  const [user,setUser]=useState(null);
  const [stageData,setStageData]=useState(initStageData);
  const [historyRows,setHistoryRows]=useState([]);
  const [page,setPage]=useState("pipeline");
  const [toast,setToast]=useState("");
  const [showLogout,setShowLogout]=useState(false);

  function showToast(msg){setToast(msg);setTimeout(()=>setToast(""),3200);}

  if(!user) return <LoginPage onLogin={setUser} users={users}/>;

  return(
    <div style={{fontFamily:"-apple-system,BlinkMacSystemFont,'SF Pro Display','Segoe UI',sans-serif",minHeight:"100vh",background:"#F2F2F7",color:"#1C1C1E"}}>
      {/* Header */}
      <header style={{background:"linear-gradient(135deg,#0A2240 0%,#1A3A5C 100%)",padding:"0 20px",height:50,display:"flex",alignItems:"center",justifyContent:"space-between",boxShadow:"0 1px 0 rgba(255,255,255,0.08),0 4px 16px rgba(0,0,0,0.35)",position:"sticky",top:0,zIndex:200}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <NavMenu page={page} setPage={setPage}/>
          <div style={{width:26,height:26,borderRadius:7,background:"rgba(255,255,255,0.12)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:15}}>⚗️</div>
          <span style={{color:"#fff",fontSize:15,fontWeight:700,letterSpacing:"-0.3px"}}>Fluxo de Liberação</span>
          {/* Breadcrumb page indicator */}
          <span style={{color:"rgba(255,255,255,0.4)",fontSize:12}}>›</span>
          <span style={{color:"rgba(255,255,255,0.8)",fontSize:12,fontWeight:500,textTransform:"capitalize"}}>{page==="pipeline"?"Pipeline":page==="dashboard"?"Dashboard":page==="historico"?"Histórico":"Configurações"}</span>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8,cursor:"pointer"}} onClick={()=>setShowLogout(true)}>
          <div style={{textAlign:"right"}}>
            <div style={{color:"#fff",fontSize:13,fontWeight:600}}>{user.name}</div>
            <div style={{color:"rgba(255,255,255,0.5)",fontSize:10}}>{DEPT_FULL[user.dept]||user.dept}</div>
          </div>
          <div style={{width:30,height:30,borderRadius:"50%",background:"linear-gradient(135deg,#007AFF,#5AC8FA)",display:"flex",alignItems:"center",justifyContent:"center",fontSize:13,fontWeight:700,color:"#fff"}}>{user.name.charAt(0)}</div>
        </div>
      </header>

      {/* Page content */}
      {page==="pipeline"&&<PipelinePage stageData={stageData} setStageData={setStageData} user={user} historyRows={historyRows} setHistoryRows={setHistoryRows} showToast={showToast}/>}
      {page==="dashboard"&&<Dashboard stageData={stageData} historyRows={historyRows} onSelectStage={sid=>{setPage("pipeline");}}/>}
      {page==="historico"&&<HistoricoPage historyRows={historyRows}/>}
      {page==="configuracoes"&&<ConfigPage users={users} setUsers={setUsers}/>}

      <Toast msg={toast}/>
      {showLogout&&<Modal title="Sair da conta" body={`Encerrar sessão de ${user.name}?`} onConfirm={()=>{setUser(null);setShowLogout(false);setStageData(initStageData());setHistoryRows([]);setPage("pipeline");}} onCancel={()=>setShowLogout(false)} confirmLabel="Sair" danger/>}
    </div>
  );
}
