import { useState, useMemo, useRef, useEffect } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import { signInWithEmailAndPassword, signOut, onAuthStateChanged } from "firebase/auth";
import { auth } from "./firebase";

const BR_HOL = new Set(["2025-01-01","2025-04-18","2025-04-21","2025-05-01","2025-06-19","2025-09-07","2025-10-12","2025-11-02","2025-11-15","2025-12-25","2026-01-01","2026-04-03","2026-04-21","2026-05-01","2026-06-04","2026-09-07","2026-10-12","2026-11-02","2026-11-15","2026-12-25","2027-01-01","2027-04-02","2027-04-21","2027-05-01","2027-05-27","2027-09-07","2027-10-12","2027-11-02","2027-11-15","2027-12-25"]);
const T={bg:"#121212",surface:"#181818",card:"#282828",hover:"#3E3E3E",accent:"#1DB954",white:"#FFFFFF",gray:"#B3B3B3",muted:"#535353",danger:"#E22134",warning:"#FF9800",purple:"#9c27b0",sidebar:"#000000",border:"#282828",font:"'Segoe UI',system-ui,sans-serif"};
const TODAY=new Date().toISOString().slice(0,10);

const normStr=s=>(s||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().trim();
const getCol=(row,...keys)=>{const rk=Object.keys(row);for(const k of keys){if(row[k]!==undefined&&row[k]!==null&&String(row[k]).trim()!=="")return String(row[k]).trim();const f=rk.find(r=>normStr(r)===normStr(k));if(f&&row[f]!==undefined&&row[f]!==null&&String(row[f]).trim()!=="")return String(row[f]).trim();}return "";};
const getIdx=(row,i)=>{const v=Object.values(row)[i];return v!==undefined&&v!==null?String(v).trim():"";};
const parseD=s=>{if(!s)return null;s=String(s).trim();if(/^\d{2}\/\d{2}\/\d{4}/.test(s)){const[d,m,y]=s.split("/");return`${y}-${m.padStart(2,"0")}-${d.padStart(2,"0")}`;}if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.slice(0,10);return null;};
const parseAny=v=>{if(!v&&v!==0)return null;const s=String(v).trim();const d=parseD(s);if(d)return d;const n=parseFloat(s);if(!isNaN(n)&&n>40000&&n<60000){const dt=new Date(Math.round((n-25569)*86400*1000));return dt.toISOString().slice(0,10);}try{const dt=new Date(s);if(!isNaN(dt)&&String(s).length>6)return dt.toISOString().slice(0,10);}catch(e){}return null;};
const normRef=s=>{if(!s&&s!==0)return"";const n=parseFloat(String(s));if(!isNaN(n)&&n>0)return String(Math.round(n));return String(s).trim();};
const addBiz=(ds,n)=>{if(!ds)return null;const dt=new Date(ds+"T12:00:00Z");let c=0;while(c<n){dt.setUTCDate(dt.getUTCDate()+1);const k=dt.toISOString().slice(0,10),w=dt.getUTCDay();if(w!==0&&w!==6&&!BR_HOL.has(k))c++;}return dt.toISOString().slice(0,10);};
const fD=s=>{if(!s)return"—";const[y,m,d]=s.split("-");return`${d}/${m}/${y}`;};
const pV=s=>{if(!s)return 0;return parseFloat(String(s).replace(/\./g,"").replace(",","."))||0;};
const fV=v=>v.toLocaleString("pt-BR",{style:"currency",currency:"BRL"});

const loadFile=(file,enc,cb)=>{const ext=file.name.split(".").pop().toLowerCase();if(["xlsx","xlsb","xls"].includes(ext)){const fr=new FileReader();fr.onload=e=>{const wb=XLSX.read(e.target.result,{type:"array",cellDates:true});const ws=wb.Sheets[wb.SheetNames[0]];cb(XLSX.utils.sheet_to_json(ws,{defval:"",raw:false}));};fr.readAsArrayBuffer(file);}else{const fr=new FileReader();fr.onload=e=>cb(Papa.parse(e.target.result,{header:true,delimiter:";",skipEmptyLines:true}).data);fr.readAsText(file,enc);}};

function analyze5125(gaRows,ctrlRows){
  const canByProt={},bckByRef={},dupTrack={};
  gaRows.forEach(r=>{
    const dep=getCol(r,"Departamento","DEPARTAMENTO").toUpperCase();
    const ec=getCol(r,"EC","ec");
    if(dep==="CAN"){
      const prot=normRef(getCol(r,"Protocolo Cancelamento","PROTOCOLO CANCELAMENTO")||getIdx(r,24));
      const auth=(getCol(r,"Autorizacao","Autorizacao","Autoriza\u00e7\u00e3o","Autorizacao")||getIdx(r,28)).toUpperCase();
      const sd=parseAny(getCol(r,"Data da venda","DATA DA VENDA","Data da Venda")||getIdx(r,27));
      const cd=parseAny(getCol(r,"Data de criacao","Data de Cria\u00e7\u00e3o","Data de criacao")||getIdx(r,11));
      if(prot)canByProt[prot]={...r,_canDate:cd};
      if(auth&&sd){const k=`${ec}|${auth}|${sd}`;if(!dupTrack[k])dupTrack[k]=[];dupTrack[k].push(prot||"?");}
    }
    if(dep==="BCK"){
      const obs=getCol(r,"Observacoes","Observa\u00e7\u00f5es","Observações")||getIdx(r,12);
      const cd=parseAny(getCol(r,"Data de criacao","Data de Cria\u00e7\u00e3o","Data de criacao")||getIdx(r,11));
      const m=obs.match(/REF\.\s*(\d+)/i);
      if(m&&cd&&(!bckByRef[m[1]]||bckByRef[m[1]]>cd))bckByRef[m[1]]=cd;
    }
  });
  const dupProts=new Set();
  Object.values(dupTrack).forEach(ps=>{if(ps.length>1)ps.forEach(p=>p&&dupProts.add(p));});
  return ctrlRows.map(c=>{
    const ref=normRef(getCol(c,"REFERENCIA","Referencia","REFER\u00caNCIA","Refer\u00eancia"));
    const ec=getCol(c,"ESTABELECIMENTO","Estabelecimento");
    const auth=getCol(c,"AUTORIZACAO","Autoriza\u00e7\u00e3o","Autorizacao","AUTORIZACAO","AUTORIZA\u00c7\u00c3O");
    const sd=parseAny(getCol(c,"DATA DA VENDA","Data da Venda"));
    const od=parseAny(getCol(c,"DATA","DATA ABERTURA","Data Abertura","Data de Abertura"));
    const bdCtrl=parseAny(getCol(c,"DATA DO AJUSTE A CREDITO","Data do Ajuste a Credito","DATA DO AJUSTE A CR\u00c9DITO"));
    const bd=bdCtrl||bckByRef[ref]||null;
    const valor=pV(getCol(c,"VALOR DA TRANSA\u00c7\u00c3O","VALOR DA TRANSACAO","Valor da Transa\u00e7\u00e3o"));
    const cval=pV(getCol(c,"VALOR DO CANCELAMENTO","Valor do Cancelamento"));
    const gaRec=canByProt[ref]||null;
    const canDate=gaRec?._canDate||null;
    const canDl=addBiz(od,2);
    const bckDl=canDate?addBiz(canDate,2):null;
    const canOk=canDate&&canDl?canDate<=canDl:null;
    const bckOk=bd&&bckDl?bd<=bckDl:(!bd&&bckDl&&TODAY>bckDl?false:null);
    const isDup=dupProts.has(ref);
    const issues=[];
    if(isDup)issues.push("DUP");
    if(!gaRec)issues.push("SEM_CAN");
    else if(canOk===false)issues.push("SLA_CAN");
    if(bckOk===false)issues.push("SLA_BCK");
    return{ref,ec,auth,sd,od,bd,valor,cval,analista:getCol(c,"ANALISTA","Analista"),ajuste:getCol(c,"AJUSTE EFETUADO?","Ajuste Efetuado?"),trans3943:getCol(c,"TRANSFERIDO PARA 3943","Transferido para 3943"),canDate,canDl,canOk,bckDl,bckOk,isDup,issues,ok:issues.length===0,_ga:gaRec,_c:c};
  });
}

const MODULES=[
  {id:"5125",name:"Evento 5125",group:"Eventos",icon:"⚡",desc:"Cancelamento sem saldo · Boleto / PIX",slots:[{key:"ctrl",label:"Planilha Controle (analistas)",enc:"latin1"},{key:"ga",label:"Relatório G.A — Gestor de Ajustes",enc:"ISO-8859-1"}],canRun:s=>s.ctrl?.length>0&&s.ga?.length>0,run:s=>analyze5125(s.ga,s.ctrl),is5125:true},
  {id:"7922",name:"Evento 7922",group:"Eventos",icon:"📋",desc:"Análise em desenvolvimento",slots:[{key:"file",label:"Planilha do Evento",enc:"UTF-8"}],canRun:s=>s.file?.length>0,run:s=>s.file},
  {id:"9066",name:"Evento 9066",group:"Eventos",icon:"📋",desc:"Análise em desenvolvimento",slots:[{key:"file",label:"Planilha do Evento",enc:"UTF-8"}],canRun:s=>s.file?.length>0,run:s=>s.file},
  {id:"reg-fin",name:"Regularizações Financeiras",group:"Caixas de E-mail",icon:"💼",desc:"Análise em desenvolvimento",slots:[{key:"file",label:"Planilha de Regularizações",enc:"UTF-8"}],canRun:s=>s.file?.length>0,run:s=>s.file},
  {id:"saldo-aud",name:"Saldo Auditoria",group:"Caixas de E-mail",icon:"🔍",desc:"Análise em desenvolvimento",slots:[{key:"file",label:"Planilha de Auditoria",enc:"UTF-8"}],canRun:s=>s.file?.length>0,run:s=>s.file},
  {id:"gest-alug",name:"Gestão Aluguel",group:"Caixas de E-mail",icon:"🏢",desc:"Análise em desenvolvimento",slots:[{key:"file",label:"Planilha de Gestão",enc:"UTF-8"}],canRun:s=>s.file?.length>0,run:s=>s.file},
  {id:"est-alug",name:"Estorno Gestão Aluguel",group:"Caixas de E-mail",icon:"↩️",desc:"Análise em desenvolvimento",slots:[{key:"file",label:"Planilha de Estornos",enc:"UTF-8"}],canRun:s=>s.file?.length>0,run:s=>s.file},
  {id:"incentivo",name:"Incentivo",group:"Caixas de E-mail",icon:"🎯",desc:"Análise em desenvolvimento",slots:[{key:"file",label:"Planilha de Incentivos",enc:"UTF-8"}],canRun:s=>s.file?.length>0,run:s=>s.file},
  {id:"desfaz",name:"Desfazimento",group:"Caixas de E-mail",icon:"🔄",desc:"Análise em desenvolvimento",slots:[{key:"file",label:"Planilha de Desfazimento",enc:"UTF-8"}],canRun:s=>s.file?.length>0,run:s=>s.file},
  {id:"std-aerea",name:"STD — Cia Aérea",group:"Caixas de E-mail",icon:"✈️",desc:"Análise em desenvolvimento",slots:[{key:"file",label:"Planilha STD Aérea",enc:"UTF-8"}],canRun:s=>s.file?.length>0,run:s=>s.file},
  {id:"std-corp",name:"STD-Corporate",group:"Caixas de E-mail",icon:"🏛️",desc:"Análise em desenvolvimento",slots:[{key:"file",label:"Planilha STD Corporate",enc:"UTF-8"}],canRun:s=>s.file?.length>0,run:s=>s.file},
];
const MODULE_BY_ID=Object.fromEntries(MODULES.map(m=>[m.id,m]));
const GROUPS=[...new Set(MODULES.map(m=>m.group))];

const BADGES={OK:{bg:"#1a3a1a",fg:T.accent,txt:"✓ OK"},ONTIME:{bg:"#1a3a1a",fg:T.accent,txt:"NO PRAZO"},LATE:{bg:"#3a1a1a",fg:"#ff4d4d",txt:"ATRASADO"},DUP:{bg:"#3a2a00",fg:T.warning,txt:"DUPLICATA"},SEM_CAN:{bg:"#3a1a1a",fg:"#ff4d4d",txt:"SEM CAN"},SLA_CAN:{bg:"#3a1a1a",fg:"#ff4d4d",txt:"⏰ CAN"},SLA_BCK:{bg:"#3a1a1a",fg:"#ff4d4d",txt:"⏰ BCK"},PEND:{bg:T.card,fg:T.muted,txt:"—"}};
const Badge=({type})=>{const s=BADGES[type]||BADGES.PEND;return<span style={{display:"inline-block",padding:"3px 8px",borderRadius:20,fontSize:10,fontWeight:700,letterSpacing:.5,background:s.bg,color:s.fg,marginRight:3,whiteSpace:"nowrap"}}>{s.txt}</span>;};

const Login=()=>{
  const[email,setEmail]=useState("");const[pass,setPass]=useState("");const[err,setErr]=useState("");const[busy,setBusy]=useState(false);
  const go=async()=>{if(!email||!pass){setErr("Preencha e-mail e senha.");return;}setBusy(true);setErr("");try{await signInWithEmailAndPassword(auth,email,pass);}catch(e){const m={"auth/invalid-credential":"Credenciais inválidas.","auth/user-not-found":"Usuário não encontrado.","auth/wrong-password":"Senha incorreta.","auth/too-many-requests":"Muitas tentativas. Aguarde."};setErr(m[e.code]||"Erro ao autenticar.");setBusy(false);}};
  return(<div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:`radial-gradient(ellipse at 60% 40%,#1a3a2a 0%,${T.bg} 60%)`,fontFamily:T.font}}><div style={{width:380,padding:"48px 40px",background:T.surface,borderRadius:16,boxShadow:"0 24px 80px rgba(0,0,0,.8)"}}><div style={{textAlign:"center",marginBottom:36}}><div style={{width:60,height:60,borderRadius:"50%",background:T.accent,display:"inline-flex",alignItems:"center",justifyContent:"center",fontSize:26,marginBottom:20,boxShadow:`0 8px 32px ${T.accent}55`}}>◈</div><h1 style={{margin:0,fontSize:24,fontWeight:900,color:T.white,letterSpacing:-.5}}>Painel de Ajustes</h1><p style={{margin:"8px 0 0",fontSize:12,color:T.gray,letterSpacing:.5}}>CONTROLE · AUDITORIA · ANÁLISE</p></div>{[{l:"E-MAIL",t:"email",v:email,s:setEmail,ph:"usuario@empresa.com.br"},{l:"SENHA",t:"password",v:pass,s:setPass,ph:"••••••••",ip:true}].map(({l,t,v,s,ph,ip})=>(<div key={l} style={{marginBottom:16}}><label style={{display:"block",fontSize:11,fontWeight:700,color:T.gray,marginBottom:8,letterSpacing:1}}>{l}</label><input type={t} value={v} onChange={e=>{s(e.target.value);if(ip)setErr("");}} onKeyDown={e=>e.key==="Enter"&&go()} placeholder={ph} style={{width:"100%",padding:"14px 16px",background:T.card,border:`1px solid ${ip&&err?"#ff4d4d":T.border}`,borderRadius:8,color:T.white,fontSize:14,outline:"none",boxSizing:"border-box"}}/></div>))}{err&&<p style={{color:"#ff4d4d",fontSize:12,margin:"-8px 0 14px"}}>{err}</p>}<button onClick={go} disabled={busy} style={{width:"100%",padding:"14px",background:busy?T.muted:T.accent,color:busy?T.gray:"#000",border:"none",borderRadius:50,fontSize:15,fontWeight:900,cursor:busy?"not-allowed":"pointer"}}>{busy?"Verificando…":"ENTRAR"}</button><p style={{textAlign:"center",fontSize:11,color:T.muted,margin:"24px 0 0"}}>Uso Interno · Acesso Restrito</p></div></div>);
};

const Sidebar=({activeId,onSelect})=>(<div style={{width:240,background:T.sidebar,flexShrink:0,overflowY:"auto",paddingTop:8}}>{GROUPS.map(g=>(<div key={g}><div style={{fontSize:10,fontWeight:700,color:T.muted,letterSpacing:1.5,padding:"16px 24px 6px",textTransform:"uppercase"}}>{g}</div>{MODULES.filter(m=>m.group===g).map(m=>(<button key={m.id} onClick={()=>onSelect(m.id)} style={{display:"flex",alignItems:"center",gap:12,width:"100%",padding:"10px 24px",border:"none",textAlign:"left",cursor:"pointer",background:activeId===m.id?"#282828":"transparent",color:activeId===m.id?T.white:T.gray,fontSize:13,fontWeight:activeId===m.id?700:400,borderLeft:activeId===m.id?`3px solid ${T.accent}`:"3px solid transparent"}}><span style={{fontSize:16}}>{m.icon}</span><span>{m.name}</span></button>))}</div>))}</div>);

const UploadZone=({label,count,onFile,enc})=>{const ref=useRef();return(<div onClick={()=>ref.current?.click()} style={{background:T.card,borderRadius:10,border:`1px dashed ${count?T.accent:T.muted}`,padding:"16px 20px",cursor:"pointer"}}><input ref={ref} type="file" accept=".csv,.xlsx,.xlsb,.xls" style={{display:"none"}} onChange={e=>e.target.files[0]&&loadFile(e.target.files[0],enc,onFile)}/><div style={{fontSize:10,fontWeight:700,color:count?T.accent:T.gray,letterSpacing:.8,marginBottom:4}}>{label.toUpperCase()}</div><div style={{fontSize:12,color:count?T.accent:T.muted}}>{count?`✅ ${count} registros carregados`:"📎 CSV · XLSX · XLSB"}</div></div>);};

const Stat=({label,value,color,icon})=>(<div style={{background:T.card,borderRadius:12,padding:"18px 16px",boxShadow:"0 4px 16px rgba(0,0,0,.4)",position:"relative",overflow:"hidden"}}><div style={{position:"absolute",top:-10,right:-10,fontSize:48,opacity:.06}}>{icon}</div><div style={{fontSize:30,fontWeight:900,color,lineHeight:1}}>{value}</div><div style={{fontSize:11,color:T.gray,marginTop:6}}>{label}</div></div>);

const GenericTable=({data,moduleId})=>{const[search,setSearch]=useState("");const cols=data.length>0?Object.keys(data[0]).filter(k=>k!==""):[];const rows=useMemo(()=>{if(!search.trim())return data;const s=search.toLowerCase();return data.filter(r=>Object.values(r).some(v=>String(v).toLowerCase().includes(s)));},[data,search]);const doExport=()=>{const ws=XLSX.utils.json_to_sheet(rows);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Dados");XLSX.writeFile(wb,`export_${moduleId}_${TODAY}.xlsx`);};return(<div><div style={{display:"flex",gap:10,alignItems:"center",marginBottom:14,flexWrap:"wrap"}}><input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar…" style={{flex:1,minWidth:200,padding:"10px 14px",background:T.card,border:`1px solid ${T.border}`,borderRadius:8,color:T.white,fontSize:13,outline:"none"}}/><span style={{fontSize:12,color:T.gray}}>{rows.length}/{data.length}</span><button onClick={doExport} style={{padding:"10px 20px",background:T.accent,color:"#000",border:"none",borderRadius:50,fontSize:12,fontWeight:700,cursor:"pointer"}}>⬇ Exportar</button></div><div style={{background:T.card,borderRadius:12,overflow:"hidden"}}><div style={{overflowX:"auto",maxHeight:"55vh",overflowY:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}><thead style={{position:"sticky",top:0,zIndex:2}}><tr style={{background:"#1a1a1a"}}>{cols.map(h=><th key={h} style={{padding:"11px 12px",textAlign:"left",fontWeight:700,color:T.gray,fontSize:10,letterSpacing:.8,whiteSpace:"nowrap",borderBottom:`1px solid ${T.border}`}}>{h}</th>)}</tr></thead><tbody>{rows.map((row,i)=>(<tr key={i} style={{background:i%2===0?T.card:"#1f1f1f",borderBottom:`1px solid ${T.border}`}}>{cols.map(k=><td key={k} style={{padding:"9px 12px",color:T.white,whiteSpace:"nowrap",maxWidth:200,overflow:"hidden",textOverflow:"ellipsis"}}>{String(row[k]||"")}</td>)}</tr>))}</tbody></table>{rows.length===0&&<div style={{textAlign:"center",padding:40,color:T.muted}}>Nenhum registro encontrado.</div>}</div></div></div>);};

const View5125=({results,onExport})=>{
  const[search,setSearch]=useState("");const[onlyIssues,setOnlyIssues]=useState(false);const[expanded,setExpanded]=useState(null);
  const stats=useMemo(()=>({total:results.length,ok:results.filter(r=>r.ok).length,issues:results.filter(r=>!r.ok).length,dup:results.filter(r=>r.isDup).length,slaCan:results.filter(r=>r.canOk===false).length,slaBck:results.filter(r=>r.bckOk===false).length,semCan:results.filter(r=>r.issues.includes("SEM_CAN")).length}),[results]);
  const shown=useMemo(()=>{let r=results;if(onlyIssues)r=r.filter(x=>!x.ok);if(search.trim()){const s=search.toLowerCase();r=r.filter(x=>x.ref.includes(s)||x.ec.includes(s)||x.auth.toLowerCase().includes(s)||x.analista.toLowerCase().includes(s));}return r;},[results,search,onlyIssues]);
  const TH=({c})=><th style={{padding:"11px 12px",textAlign:"left",fontWeight:700,color:T.gray,fontSize:10,letterSpacing:.8,whiteSpace:"nowrap",borderBottom:`1px solid ${T.border}`}}>{c}</th>;
  return(<div>
    <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:12,marginBottom:24}}>
      {[[stats.total,"Total","#60a5fa","📊"],[stats.ok,"OK",T.accent,"✅"],[stats.issues,"Pendência","#ff4d4d","⚠️"],[stats.dup,"Duplicata",T.warning,"🔁"],[stats.slaCan,"SLA CAN","#ff4d4d","⏰"],[stats.slaBck,"SLA BCK","#ff4d4d","⏰"],[stats.semCan,"Sem CAN",T.purple,"❌"]].map(([v,l,c,ic])=><Stat key={l} label={l} value={v} color={c} icon={ic}/>)}
    </div>
    <div style={{display:"flex",gap:10,alignItems:"center",marginBottom:14,flexWrap:"wrap"}}>
      <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar por referência, EC, autorização, analista…" style={{flex:1,minWidth:200,padding:"10px 14px",background:T.card,border:`1px solid ${T.border}`,borderRadius:8,color:T.white,fontSize:13,outline:"none"}}/>
      <label style={{display:"flex",alignItems:"center",gap:8,fontSize:13,cursor:"pointer",color:T.gray}}><input type="checkbox" checked={onlyIssues} onChange={e=>setOnlyIssues(e.target.checked)} style={{accentColor:T.accent}}/>Apenas pendências</label>
      <span style={{fontSize:12,color:T.muted}}>{shown.length}/{results.length}</span>
      <button onClick={()=>onExport(shown)} style={{padding:"10px 20px",background:T.accent,color:"#000",border:"none",borderRadius:50,fontSize:12,fontWeight:700,cursor:"pointer"}}>⬇ Exportar XLSX</button>
    </div>
    <div style={{background:T.card,borderRadius:12,overflow:"hidden",boxShadow:"0 4px 16px rgba(0,0,0,.4)"}}>
      <div style={{overflowX:"auto",maxHeight:"50vh",overflowY:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
          <thead style={{position:"sticky",top:0,zIndex:2}}><tr style={{background:"#1a1a1a"}}>{["Referência","EC","Autorização","Data Venda","Valor","Data Abertura","Analista","Data CAN","Prazo CAN","SLA CAN","Data BCK","Prazo BCK","SLA BCK","Situação"].map(c=><TH key={c} c={c}/>)}</tr></thead>
          <tbody>
            {shown.map((r,i)=>(<>
              <tr key={`r${i}`} onClick={()=>setExpanded(expanded===i?null:i)} style={{background:r.isDup?"#2a1f00":!r.ok?"#2a0000":i%2===0?T.card:"#1f1f1f",borderBottom:`1px solid ${T.border}`,cursor:"pointer"}}>
                <td style={{padding:"9px 12px",fontWeight:700,color:T.accent}}>{r.ref||"—"}</td>
                <td style={{padding:"9px 12px",fontFamily:"monospace",fontSize:10,color:T.gray}}>{r.ec}</td>
                <td style={{padding:"9px 12px",fontFamily:"monospace",color:T.white}}>{r.auth||"—"}</td>
                <td style={{padding:"9px 12px",color:T.white}}>{fD(r.sd)}</td>
                <td style={{padding:"9px 12px",fontWeight:700,color:T.white}}>{fV(r.valor)}</td>
                <td style={{padding:"9px 12px",color:T.white}}>{fD(r.od)}</td>
                <td style={{padding:"9px 12px",color:T.gray}}>{r.analista||"—"}</td>
                <td style={{padding:"9px 12px",color:r.canOk===false?"#ff4d4d":r.canOk?T.accent:T.muted}}>{fD(r.canDate)}</td>
                <td style={{padding:"9px 12px",color:T.muted}}>{fD(r.canDl)}</td>
                <td style={{padding:"9px 12px"}}><Badge type={r.canOk===true?"ONTIME":r.canOk===false?"LATE":"PEND"}/></td>
                <td style={{padding:"9px 12px",color:r.bckOk===false?"#ff4d4d":r.bckOk?T.accent:T.muted}}>{fD(r.bd)}</td>
                <td style={{padding:"9px 12px",color:T.muted}}>{fD(r.bckDl)}</td>
                <td style={{padding:"9px 12px"}}><Badge type={r.bckOk===true?"ONTIME":r.bckOk===false?"LATE":"PEND"}/></td>
                <td style={{padding:"9px 12px"}}>{r.ok?<Badge type="OK"/>:r.issues.map(t=><Badge key={t} type={t}/>)}</td>
              </tr>
              {expanded===i&&(<tr key={`e${i}`} style={{background:"#1a1a2a"}}><td colSpan={14} style={{padding:"14px 18px"}}>
                <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:10}}>
                  {[["Valor Cancelamento",fV(r.cval)],["Ajuste Efetuado",r.ajuste||"—"],["Transf. 3943",r.trans3943||"—"],["CAN no G.A",r._ga?"✅ Localizado":"❌ Não localizado"],["Dias abert.→CAN",r.canDate&&r.od?`${Math.round((new Date(r.canDate)-new Date(r.od))/86400000)} dias`:"—"],["Dias CAN→BCK",r.canDate&&r.bd?`${Math.round((new Date(r.bd)-new Date(r.canDate))/86400000)} dias`:"—"],["Prazo CAN",fD(r.canDl)],["Prazo BCK",fD(r.bckDl)]].map(([l,v])=>(
                    <div key={l} style={{background:T.card,padding:"10px 12px",borderRadius:8,border:`1px solid ${T.border}`}}><div style={{fontSize:10,color:T.muted,marginBottom:3}}>{l}</div><div style={{fontWeight:700,color:T.white,fontSize:13}}>{v}</div></div>
                  ))}
                </div>
              </td></tr>)}
            </>))}
          </tbody>
        </table>
        {shown.length===0&&<div style={{textAlign:"center",padding:40,color:T.muted}}>Nenhum registro encontrado.</div>}
      </div>
    </div>
    <div style={{marginTop:12,display:"flex",gap:12,flexWrap:"wrap",alignItems:"center"}}>
      <span style={{fontSize:11,color:T.gray,fontWeight:700}}>Legenda:</span>
      {[["Duplicata","DUP"],["Sem CAN","SEM_CAN"],["SLA CAN","SLA_CAN"],["SLA BCK","SLA_BCK"],["No prazo","ONTIME"],["OK","OK"]].map(([l,t])=>(<span key={t} style={{display:"flex",alignItems:"center",gap:5,fontSize:11,color:T.gray}}><Badge type={t}/>{l}</span>))}
      <span style={{fontSize:10,color:T.muted,marginLeft:"auto"}}>Clique na linha para detalhar · D+2 considera feriados nacionais</span>
    </div>
  </div>);
};

const ModuleContent=({moduleId,files,setFiles,results,setResults})=>{
  const mod=MODULE_BY_ID[moduleId];const slotData=files[moduleId]||{};const moduleResults=results[moduleId]||null;
  const setSlot=(key,data)=>setFiles(f=>({...f,[moduleId]:{...f[moduleId],[key]:data}}));
  const run=()=>{if(mod.canRun(slotData))setResults(r=>({...r,[moduleId]:mod.run(slotData)}));};
  const export5125=rows=>{const out=rows.map(r=>({"Referência":r.ref,"EC":r.ec,"Autorização":r.auth,"Data Venda":fD(r.sd),"Valor":fV(r.valor),"Data Abertura":fD(r.od),"Analista":r.analista,"Data CAN":fD(r.canDate),"Prazo CAN":fD(r.canDl),"SLA CAN":r.canOk===true?"NO PRAZO":r.canOk===false?"ATRASADO":"—","Data BCK":fD(r.bd),"Prazo BCK":fD(r.bckDl),"SLA BCK":r.bckOk===true?"NO PRAZO":r.bckOk===false?"ATRASADO":"—","Pendências":r.issues.join(", ")||"OK"}));const ws=XLSX.utils.json_to_sheet(out);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Análise");XLSX.writeFile(wb,`analise_5125_${TODAY}.xlsx`);};
  return(<div style={{flex:1,padding:"24px 28px",overflowY:"auto",fontFamily:T.font}}>
    <div style={{marginBottom:24,paddingBottom:16,borderBottom:`1px solid ${T.border}`,display:"flex",alignItems:"center",gap:14}}>
      <span style={{fontSize:28}}>{mod.icon}</span>
      <div><h2 style={{margin:0,fontSize:22,fontWeight:900,color:T.white}}>{mod.name}</h2><p style={{margin:0,fontSize:13,color:T.gray}}>{mod.desc}</p></div>
      {!mod.is5125&&<span style={{marginLeft:"auto",padding:"4px 14px",background:T.card,borderRadius:20,fontSize:11,color:T.gray,border:`1px solid ${T.border}`}}>Em desenvolvimento</span>}
    </div>
    <div style={{display:"grid",gridTemplateColumns:`repeat(${mod.slots.length},1fr) auto`,gap:12,alignItems:"end",marginBottom:24}}>
      {mod.slots.map(s=><UploadZone key={s.key} label={s.label} count={slotData[s.key]?.length||0} onFile={d=>setSlot(s.key,d)} enc={s.enc}/>)}
      <button onClick={run} disabled={!mod.canRun(slotData)} style={{padding:"0 28px",height:60,border:"none",borderRadius:50,fontSize:13,fontWeight:900,letterSpacing:.5,whiteSpace:"nowrap",background:mod.canRun(slotData)?T.accent:T.muted,color:mod.canRun(slotData)?"#000":T.card,cursor:mod.canRun(slotData)?"pointer":"not-allowed",boxShadow:mod.canRun(slotData)?`0 4px 20px ${T.accent}55`:"none"}}>▶ ANALISAR</button>
    </div>
    {moduleResults&&mod.is5125&&<View5125 results={moduleResults} onExport={export5125}/>}
    {moduleResults&&!mod.is5125&&<GenericTable data={moduleResults} moduleId={moduleId}/>}
    {!moduleResults&&(<div style={{textAlign:"center",padding:"72px 24px",color:T.muted}}><div style={{fontSize:56,marginBottom:20}}>{mod.icon}</div>{mod.is5125?(<><p style={{fontSize:15,fontWeight:700,color:T.gray,margin:"0 0 12px"}}>Carregue as planilhas e clique em Analisar</p><p style={{fontSize:12,margin:0,lineHeight:2,color:T.muted}}>✔ Cancelamentos duplicados · ✔ SLA CAN D+2 · ✔ SLA BCK D+2 · ✔ Feriados 2025–2027</p></>):(<><p style={{fontSize:15,fontWeight:700,color:T.gray,margin:"0 0 8px"}}>Carregue o arquivo para visualizar os dados</p><p style={{fontSize:12,color:T.muted}}>Análise personalizada em breve</p></>)}</div>)}
  </div>);
};

const Footer=()=>(<div style={{background:T.sidebar,borderTop:`1px solid ${T.border}`,padding:"10px 28px",display:"flex",alignItems:"center",gap:12,flexShrink:0}}>
  <span style={{fontSize:11,color:T.muted}}>Desenvolvido por</span>
  <a href="https://github.com/SentinelaQA" target="_blank" rel="noreferrer" style={{display:"flex",alignItems:"center",gap:7,fontSize:12,color:T.gray,textDecoration:"none",fontWeight:600}}>
    <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M12 0C5.37 0 0 5.37 0 12c0 5.3 3.438 9.8 8.205 11.385.6.113.82-.258.82-.577 0-.285-.01-1.04-.015-2.04-3.338.724-4.042-1.61-4.042-1.61C4.422 18.07 3.633 17.7 3.633 17.7c-1.087-.744.084-.729.084-.729 1.205.084 1.838 1.236 1.838 1.236 1.07 1.835 2.809 1.305 3.495.998.108-.776.417-1.305.76-1.605-2.665-.3-5.466-1.332-5.466-5.93 0-1.31.465-2.38 1.235-3.22-.135-.303-.54-1.523.105-3.176 0 0 1.005-.322 3.3 1.23.96-.267 1.98-.399 3-.405 1.02.006 2.04.138 3 .405 2.28-1.552 3.285-1.23 3.285-1.23.645 1.653.24 2.873.12 3.176.765.84 1.23 1.91 1.23 3.22 0 4.61-2.805 5.625-5.475 5.92.42.36.81 1.096.81 2.22 0 1.606-.015 2.896-.015 3.286 0 .315.21.69.825.57C20.565 21.795 24 17.295 24 12c0-6.63-5.37-12-12-12"/></svg>
    Yasmin de Melo Campos — Analista de Qualidade Jr
  </a>
</div>);

export default function App(){
  const[user,setUser]=useState(null);const[loading,setLoading]=useState(true);
  const[activeModule,setActiveModule]=useState("5125");
  const[files,setFiles]=useState({});const[results,setResults]=useState({});
  useEffect(()=>{const u=onAuthStateChanged(auth,u=>{setUser(u);setLoading(false);});return u;},[]);
  if(loading)return<div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:T.bg}}><div style={{color:T.accent,fontSize:40,fontWeight:900}}>◈</div></div>;
  if(!user)return<Login/>;
  return(<div style={{display:"flex",flexDirection:"column",height:"100vh",fontFamily:T.font,background:T.bg,overflow:"hidden",color:T.white}}>
    <div style={{background:T.sidebar,height:56,display:"flex",alignItems:"center",padding:"0 24px",borderBottom:`1px solid ${T.border}`,flexShrink:0,zIndex:10}}>
      <div style={{display:"flex",alignItems:"center",gap:10,fontWeight:900,fontSize:15}}><span style={{color:T.accent,fontSize:20}}>◈</span><span>Painel de Ajustes</span></div>
      <div style={{margin:"0 16px",color:T.border}}>|</div>
      <div style={{fontSize:12,color:T.muted}}>{MODULE_BY_ID[activeModule]?.name}</div>
      <div style={{flex:1}}/>
      <div style={{fontSize:11,color:T.muted,marginRight:16}}>{user.email}</div>
      <button onClick={()=>{signOut(auth);setFiles({});setResults({});}} style={{background:"transparent",border:`1px solid ${T.border}`,color:T.gray,padding:"6px 16px",borderRadius:50,fontSize:11,cursor:"pointer",fontWeight:700}}>Sair</button>
    </div>
    <div style={{display:"flex",flex:1,overflow:"hidden"}}>
      <Sidebar activeId={activeModule} onSelect={setActiveModule}/>
      <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden"}}>
        <ModuleContent moduleId={activeModule} files={files} setFiles={setFiles} results={results} setResults={setResults}/>
        <Footer/>
      </div>
    </div>
  </div>);
}
