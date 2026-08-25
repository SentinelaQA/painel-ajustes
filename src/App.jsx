import { useState, useMemo, useRef, useEffect } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import { signInWithEmailAndPassword, signOut, onAuthStateChanged } from "firebase/auth";
import { auth } from "./firebase";

const BR_HOL = new Set(["2025-01-01","2025-04-18","2025-04-21","2025-05-01","2025-06-19","2025-09-07","2025-10-12","2025-11-02","2025-11-15","2025-12-25","2026-01-01","2026-04-03","2026-04-21","2026-05-01","2026-06-04","2026-09-07","2026-10-12","2026-11-02","2026-11-15","2026-12-25","2027-01-01","2027-04-02","2027-04-21","2027-05-01","2027-05-27","2027-09-07","2027-10-12","2027-11-02","2027-11-15","2027-12-25"]);
const T={
  bg:      "hsl(228,30%,13%)",
  surface: "hsl(228,28%,16%)",
  card:    "hsl(228,26%,18%)",
  hover:   "hsl(228,24%,22%)",
  accent:  "#00b4d8",
  accentHover:"#00d4f8",
  white:   "#f0f4ff",
  gray:    "#8892a4",
  muted:   "#5a6478",
  danger:  "#ff5252",
  dangerDark:"hsl(0,62.8%,18%)",
  warning: "#ffab40",
  success: "#00e676",
  purple:  "#b388ff",
  sidebar: "hsl(228,35%,11%)",
  sidebarAccent:"hsl(228,30%,16%)",
  border:  "hsl(228,24%,22%)",
  font:    "'Segoe UI',system-ui,sans-serif",
};
const TODAY=new Date().toISOString().slice(0,10);

const normStr=s=>(s||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().trim();
const getCol=(row,...keys)=>{const rk=Object.keys(row);for(const k of keys){if(row[k]!==undefined&&row[k]!==null&&String(row[k]).trim()!=="")return String(row[k]).trim();const f=rk.find(r=>normStr(r)===normStr(k));if(f&&row[f]!==undefined&&row[f]!==null&&String(row[f]).trim()!=="")return String(row[f]).trim();}return "";};
const getIdx=(row,i)=>{const v=Object.values(row)[i];return v!==undefined&&v!==null?String(v).trim():"";};
const parseD=s=>{if(!s)return null;s=String(s).trim();if(/^\d{2}\/\d{2}\/\d{4}/.test(s)){const[d,m,y]=s.split("/");return`${y}-${m.padStart(2,"0")}-${d.padStart(2,"0")}`;}if(/^\d{4}-\d{2}-\d{2}/.test(s))return s.slice(0,10);return null;};
const parseAny=v=>{if(!v&&v!==0)return null;const s=String(v).trim();if(!s)return null;const d=parseD(s);if(d)return d;const n=parseFloat(s);if(!isNaN(n)&&n>40000&&n<60000){try{const dt=new Date(Math.floor((n-25569)*86400*1000));if(!isNaN(dt.getTime()))return dt.toISOString().slice(0,10);}catch(e){}}const dm=s.match(/(\d{2}\/\d{2}\/\d{4})/);if(dm)return parseD(dm[1]);try{const dt=new Date(s);if(!isNaN(dt.getTime())){const iso=dt.toISOString().slice(0,10);if(iso>="2000-01-01"&&iso<="2035-12-31")return iso;}}catch(e){}return null;};
const isScheduled=v=>{if(!v)return false;const s=String(v).trim();return !parseD(s)&&/(\d{2}\/\d{2}\/\d{4})/.test(s);};

const normRef=s=>{if(!s&&s!==0)return"";const n=parseFloat(String(s));if(!isNaN(n)&&n>0)return String(Math.round(n));return String(s).trim();};
const addBiz=(ds,n)=>{if(!ds)return null;try{const dt=new Date(ds+"T12:00:00Z");if(isNaN(dt.getTime()))return null;let c=0;while(c<n){dt.setUTCDate(dt.getUTCDate()+1);const k=dt.toISOString().slice(0,10),w=dt.getUTCDay();if(w!==0&&w!==6&&!BR_HOL.has(k))c++;}return dt.toISOString().slice(0,10);}catch(e){return null;}};
const fD=s=>{if(!s)return"—";const[y,m,d]=s.split("-");return`${d}/${m}/${y}`;};
const pV=s=>{
  if(!s)return 0;
  s=String(s).replace(/R\$\s*/g,"").trim();
  if(!s)return 0;
  if(s.includes(","))return parseFloat(s.replace(/\./g,"").replace(",","."))||0;
  return parseFloat(s)||0;
};
const fV=v=>v.toLocaleString("pt-BR",{style:"currency",currency:"BRL"});
// Formata taxa % pra exibição, arredondando erro de ponto flutuante (0.8500000000000001 -> 0.85)
const fP=v=>v==null?"—":`${(Math.round(v*100)/100).toLocaleString("pt-BR",{minimumFractionDigits:2,maximumFractionDigits:2})}%`;

// Em planilhas com múltiplas abas (ex: "Resumo" + "Faturamento"), pega a aba com mais linhas —
// a aba de detalhe transacional é sempre a maior; um resumo/pivot tem poucas linhas.
const loadFile=(file,enc,cb)=>{const ext=file.name.split(".").pop().toLowerCase();if(["xlsx","xlsb","xls"].includes(ext)){const fr=new FileReader();fr.onload=e=>{const wb=XLSX.read(e.target.result,{type:"array"});let best=[];wb.SheetNames.forEach(name=>{const rows=XLSX.utils.sheet_to_json(wb.Sheets[name],{defval:"",raw:true});if(rows.length>best.length)best=rows;});cb(best);};fr.readAsArrayBuffer(file);}else{const fr=new FileReader();fr.onload=e=>cb(Papa.parse(e.target.result,{header:true,delimiter:";",skipEmptyLines:true}).data);fr.readAsText(file,enc);}};

function analyze5125(gaRows,ctrlRows){
  const canByProt={},bckByRef={},dupTrack={};
  gaRows.forEach(r=>{
    const ec=getCol(r,"EC","ec");
    const cod=getIdx(r,8);
    const cd=parseAny(getIdx(r,11));
    const obs=getIdx(r,12);
    const is97=/^\s*97[\s\-]/i.test(cod)||cod.toUpperCase().includes("CANCELAMENTO DE VENDA");
    const is962=/^\s*(C?\s*962|962)[\s\-]/i.test(cod)||cod.toUpperCase().includes("ACERTO")||cod.toUpperCase().includes("C962");
    if(is97){
      const prot=normRef(getCol(r,"Protocolo Cancelamento","PROTOCOLO CANCELAMENTO")||getIdx(r,24));
      const auth=(getCol(r,"Autorizacao","Autorizacao","Autoriza\u00e7\u00e3o","Autorizacao")||getIdx(r,28)).toUpperCase();
      const sd=parseAny(getIdx(r,27));
      if(prot) canByProt[prot]={...r,_canDate:cd};
      if(auth&&sd){const k=`${ec}|${auth}|${sd}`;if(!dupTrack[k])dupTrack[k]=[];dupTrack[k].push(prot||"?");}
    }
    if(is962){
      const dep=getCol(r,"Departamento","DEPARTAMENTO").toUpperCase()||"CAN";
      const m=obs.match(/REF\.\s*(\d+)/i);
      if(m&&cd&&(!bckByRef[m[1]]||bckByRef[m[1]].date>cd))
        bckByRef[m[1]]={date:cd,row:r,obs,dept:dep,cod};
    }
    if(!is97&&!is962&&getCol(r,"Departamento").toUpperCase()==="BCK"){
      const m=obs.match(/REF\.\s*(\d+)/i);
      if(m&&cd&&(!bckByRef[m[1]]||bckByRef[m[1]].date>cd))
        bckByRef[m[1]]={date:cd,row:r,obs,dept:"BCK",cod};
    }
  });

  // Duplicate detection: same EC+AUTH+DATA_VENDA+VALOR_TRANS+CARTAO in control
  const ctrlKeyMap={};
  ctrlRows.forEach(cr=>{
    const ec=getCol(cr,"ESTABELECIMENTO","Estabelecimento");
    const sd=parseAny(getCol(cr,"DATA DA VENDA","Data da Venda","DATA DA VENDA"));
    const vt=String(Math.round(pV(getCol(cr,"VALOR DA TRANSAÇÃO","VALOR DA TRANSACAO","Valor da Transação"))*100));
    const auth=getCol(cr,"AUTORIZAÇÃO","AUTORIZACAO","Autorização","Autorizacao").toUpperCase().trim();
    const cartao=getCol(cr,"CARTÃO","CARTAO","Cartão").trim();
    const ref=normRef(getCol(cr,"REFERÊNCIA","REFERENCIA","Referencia"));
    if(ec&&auth&&sd){
      const k=`${ec}|${auth}|${sd}|${vt}|${cartao}`;
      if(!ctrlKeyMap[k])ctrlKeyMap[k]=[];
      ctrlKeyMap[k].push(ref);
    }
  });
  const ctrlDupRefs=new Set();
  const ctrlDupGroups={};
  Object.entries(ctrlKeyMap).forEach(([,refs])=>{
    if(refs.length>1){
      refs.forEach(ref=>{ctrlDupRefs.add(ref);ctrlDupGroups[ref]=refs;});
    }
  });

  // Build a lookup of ALL control rows by ref for duplicate group display
  const ctrlByRef={};
  ctrlRows.forEach(cr=>{
    const ref=normRef(getCol(cr,"REFERÊNCIA","REFERENCIA","Referencia"));
    if(ref) ctrlByRef[ref]=cr;
  });

  return ctrlRows.map(cr=>{
    const ref=normRef(getCol(cr,"REFERÊNCIA","REFERENCIA","Referencia","Referência"));
    const ec=getCol(cr,"ESTABELECIMENTO","Estabelecimento");
    const auth=getCol(cr,"AUTORIZAÇÃO","AUTORIZACAO","Autorização","Autorizacao");
    const od=parseAny(getCol(cr,"DATA ABERTURA","DATA","Data Abertura","Data de Abertura"));
    const sd=parseAny(getCol(cr,"DATA DA VENDA","Data da Venda","DATA DA VENDA"));
    const bdRaw=getCol(cr,"DATA DO AJUSTE A CREDITO","Data do Ajuste a Credito");
    const bdCtrl=parseAny(bdRaw);
    const bdScheduled=isScheduled(bdRaw);
    const bckRec=bckByRef[ref]||null;
    const bd=bdCtrl||bckRec?.date||null;
    const valor=pV(getCol(cr,"VALOR DA TRANSAÇÃO","VALOR DA TRANSACAO","Valor da Transação"));
    const cval=pV(getCol(cr,"VALOR DO CANCELAMENTO","Valor do Cancelamento"));
    const vBoleto=pV(getCol(cr,"VALOR DO BOLETO","Valor do Boleto"));
    const obs=getCol(cr,"OBSERVAÇÃO","OBSERVACAO","Observação","Observacao");
    const bckValor=bckRec?pV(getIdx(bckRec.row,16)):null;
    const gaRec=canByProt[ref]||null;
    const sdFinal=sd||(gaRec?parseAny(getIdx(gaRec,27)):null);
    const canDate=gaRec?._canDate||null;
    const canDl=addBiz(od,2);
    const bckDl=canDate?addBiz(canDate,2):null;
    const canOk=canDate&&canDl?canDate<=canDl:null;
    const bckOk=bd&&bckDl?bd<=bckDl:(!bd&&bckDl&&TODAY>bckDl?false:null);
    const bckStatus=bdScheduled?"SCHED":bckOk===true?"OK":bckOk===false?"LATE":"PEND";
    const isDup=ctrlDupRefs.has(ref);
    // Build duplicate group detail for display
    const dupGroupRefs=ctrlDupGroups[ref]||[];
    const dupGroupDetail=dupGroupRefs.map(r=>({
      ref:r,
      cval:pV(getCol(ctrlByRef[r]||{},"VALOR DO CANCELAMENTO","Valor do Cancelamento")),
      hasCAN:!!canByProt[r],
      canDate:canByProt[r]?._canDate||null,
      hasBCK:!!bckByRef[r],
      bckDate:bckByRef[r]?.date||null,
      bckValor:bckByRef[r]?pV(getIdx(bckByRef[r].row,16)):null,
    }));
    const issues=[];
    if(isDup)issues.push("DUP");
    if(!gaRec)issues.push("SEM_CAN");
    if(bckOk===false&&!bdScheduled)issues.push("SLA_BCK");
    if(bckRec&&bckValor!==null&&cval>0&&Math.abs(bckValor-cval)>0.05)issues.push("VALOR_DIFF");
    return{ref,ec,auth,sd:sdFinal,od,bd,valor,cval,vBoleto,obs,bckValor,
      analista:getCol(cr,"ANALISTA","Analista"),
      ajuste:getCol(cr,"AJUSTE EFETUADO?","Ajuste Efetuado?"),
      trans3943:getCol(cr,"TRANSFERIDO PARA 3943","Transferido para 3943"),
      boleto:getCol(cr,"NÚMERO BOLETO","NUMERO BOLETO","Numero Boleto"),
      tipoPag:getCol(cr,"TIPO DE PAGAMENTO","Tipo de Pagamento"),
      cartao:getCol(cr,"CARTÃO","CARTAO","Cartão"),
      canDate,canDl,canOk,bckDl,bckOk,bdScheduled,bckStatus,
      isDup,dupGroupRefs,dupGroupDetail,issues,ok:issues.length===0,
      _ga:gaRec,_bck:bckRec,_c:cr};
  });
}

function analyze9066(gaRows,ctrlRows){
  const gaByEcLogico={},gaByEc={};
  gaRows.forEach(r=>{
    const ec=String(r['EC']||'').trim();
    const vals=Object.values(r);
    const logico=normRef(String(vals[13]||''));
    const cod=String(vals[8]||'');
    const bandeira=String(vals[15]||'');
    const valorGA=pV(String(vals[16]||''));
    const dv=vals[11];
    const dataCriacao=parseAny(dv instanceof Date?dv.toISOString().slice(0,10):String(dv||''));
    const status=String(vals[1]||'');
    const obs=String(vals[12]||'');
    const rec={...r,_ec:ec,_logico:logico,_cod:cod,_bandeira:bandeira,_valor:valorGA,_date:dataCriacao,_status:status,_obs:obs};
    const key=`${ec}|${logico}`;
    if(!gaByEcLogico[key])gaByEcLogico[key]=[];
    gaByEcLogico[key].push(rec);
    if(!gaByEc[ec])gaByEc[ec]=[];
    gaByEc[ec].push(rec);
  });
  return ctrlRows
    .filter(cr=>getCol(cr,'EC','ec').trim()&&getCol(cr,'Logico','LOGICO','logico').trim())
    .map(cr=>{
    const ec=getCol(cr,'EC','ec').trim();
    const logico=normRef(getCol(cr,'Logico','LOGICO','logico'));
    const ref=getCol(cr,'Ref/ Chamado','Ref/Chamado','REF/ CHAMADO');
    const valor=pV(getCol(cr,'Valor','VALOR'));
    const analista=getCol(cr,'Analista','ANALISTA');
    const dataEntrada=parseAny(getCol(cr,'Data Entrada','DATA ENTRADA'));
    const dataFin=parseAny(getCol(cr,'Data','DATA'));
    const finalizado=getCol(cr,'Finalizado','FINALIZADO');
    const key=`${ec}|${logico}`;
    const gaList=gaByEcLogico[key]||[];
    const gaRec=gaList.find(r=>r._logico===logico)||gaByEc[ec]?.find(r=>r._logico===logico)||gaByEc[ec]?.[0]||null;
    const slaDl=addBiz(dataEntrada,4);
    const ajusteDate=gaRec?._date||null;
    const slaOk=ajusteDate&&slaDl?ajusteDate<=slaDl:(!ajusteDate&&slaDl&&TODAY>slaDl?false:null);
    const cod984=gaRec?/984/.test(gaRec._cod):null;
    const logicoMatch=gaRec?gaRec._logico===logico:null;
    const bandVisa=gaRec?/VISA/i.test(gaRec._bandeira):null;
    const valorMatch=gaRec?Math.abs(gaRec._valor-valor)<0.05:null;
    const issues=[];
    if(!gaRec)issues.push('SEM_984');
    else{
      if(!cod984)issues.push('COD_ERRADO');
      if(logicoMatch===false)issues.push('LOGICO_DIFF');
      if(bandVisa===false)issues.push('BAND_ERRADA');
      if(valorMatch===false)issues.push('VALOR_DIFF');
    }
    if(slaOk===false)issues.push('SLA_BCK');
    return{ec,logico,ref,valor,analista,dataEntrada,dataFin,finalizado,
      slaDl,ajusteDate,slaOk,cod984,logicoMatch,bandVisa,valorMatch,
      gaValor:gaRec?._valor||null,gaBandeira:gaRec?._bandeira||'',
      gaStatus:gaRec?._status||'',gaCod:gaRec?._cod||'',gaObs:gaRec?._obs||'',
      gaLogico:gaRec?._logico||'',
      issues,ok:issues.length===0,_ga:gaRec,_c:cr};
  });
}

const TAXA_TABLE={"VISA|D":0.85,"VISA|0":3.08,"VISA|2":4.25,"VISA|3":4.86,"VISA|4":5.45,"VISA|5":6.00,"VISA|6":6.57,"VISA|7":7.63,"VISA|8":8.14,"VISA|9":8.70,"VISA|10":9.24,"VISA|11":9.76,"VISA|12":10.26,"MASTERCARD|D":0.85,"MASTERCARD|0":3.08,"MASTERCARD|2":4.25,"MASTERCARD|3":4.86,"MASTERCARD|4":5.45,"MASTERCARD|5":6.00,"MASTERCARD|6":6.57,"MASTERCARD|7":7.63,"MASTERCARD|8":8.14,"MASTERCARD|9":8.70,"MASTERCARD|10":9.24,"MASTERCARD|11":9.76,"MASTERCARD|12":10.26,"ELO|D":1.48,"ELO|0":3.48,"ELO|2":4.93,"ELO|3":5.54,"ELO|4":6.13,"ELO|5":6.68,"ELO|6":7.25,"ELO|7":8.13,"ELO|8":8.64,"ELO|9":9.20,"ELO|10":9.74,"ELO|11":10.26,"ELO|12":10.76,"AMEX|0":3.48,"AMEX|2":4.66,"AMEX|3":5.27,"AMEX|4":5.86,"AMEX|5":6.41,"AMEX|6":6.98,"AMEX|7":8.13,"AMEX|8":8.64,"AMEX|9":9.20,"AMEX|10":9.74,"AMEX|11":10.26,"AMEX|12":10.76,"DINERS|0":3.19,"DINERS|2":4.12,"DINERS|3":4.73,"DINERS|4":5.32,"DINERS|5":5.87,"DINERS|6":6.44,"DINERS|7":7.13,"DINERS|8":7.64,"DINERS|9":8.20,"DINERS|10":8.74,"DINERS|11":9.26,"DINERS|12":9.76,"HIPERCARD|0":5.72,"HIPERCARD|2":6.99,"HIPERCARD|3":7.60,"HIPERCARD|4":8.19,"HIPERCARD|5":8.74,"HIPERCARD|6":9.31,"HIPERCARD|7":8.68,"HIPERCARD|8":9.19,"HIPERCARD|9":9.75,"HIPERCARD|10":10.29,"HIPERCARD|11":10.81,"HIPERCARD|12":11.31,"AGIPLAN|0":3.65,"AGIPLAN|2":6.40,"AGIPLAN|3":7.01,"AGIPLAN|4":7.85,"AGIPLAN|5":8.40,"AGIPLAN|6":8.97,"AGIPLAN|7":9.84,"AGIPLAN|8":10.35,"AGIPLAN|9":10.91,"AGIPLAN|10":11.45,"AGIPLAN|11":11.97,"AGIPLAN|12":12.47,"BANESCARD|D":2.45,"BANESCARD|0":4.88,"BANESCARD|2":6.40,"BANESCARD|3":7.01,"BANESCARD|4":7.85,"BANESCARD|5":8.40,"BANESCARD|6":8.97,"BANESCARD|7":9.84,"BANESCARD|8":10.35,"BANESCARD|9":10.91,"BANESCARD|10":11.45,"BANESCARD|11":11.97,"BANESCARD|12":12.47,"SOROCRED|0":4.88,"SOROCRED|2":6.40,"SOROCRED|3":7.01,"SOROCRED|4":7.85,"SOROCRED|5":8.40,"SOROCRED|6":8.97,"SOROCRED|7":9.84,"SOROCRED|8":10.35,"SOROCRED|9":10.91,"SOROCRED|10":11.45,"SOROCRED|11":11.97,"SOROCRED|12":12.47,"CABAL|D":2.45,"CABAL|0":4.88,"CABAL|2":6.40,"CABAL|3":7.01,"CABAL|4":7.85,"CABAL|5":8.40,"CABAL|6":8.97,"CABAL|7":9.84,"CABAL|8":10.35,"CABAL|9":10.91,"CABAL|10":11.45,"CABAL|11":11.97,"CABAL|12":12.47};

function detectBandeira(produto){
  const p=String(produto||'').toUpperCase().trim();
  let b='VISA';
  if(p.includes('MASTERCARD')||p.includes('MAESTRO')) b='MASTERCARD';
  else if(p.includes('ELO')) b='ELO';
  else if(p.includes('AMEX')) b='AMEX';
  else if(p.includes('DINERS')) b='DINERS';
  else if(p.includes('HIPERCARD')||p.includes('HIPER')) b='HIPERCARD';
  else if(p.includes('AGIPLAN')) b='AGIPLAN';
  else if(p.includes('BANESCARD')) b='BANESCARD';
  else if(p.includes('SOROCRED')) b='SOROCRED';
  else if(p.includes('CABAL')) b='CABAL';
  else if(p.includes('VISA')) b='VISA';
  const isDebito=p.includes('DEBITO')||p.includes('DÉBITO')||p.includes('MAESTRO');
  // Voucher/Multibenefícios (VA/VR) roda na bandeira do cartão mas com tarifa própria, diferente
  // da tarifa de crédito à vista da mesma bandeira — precisa de token de parcela dedicado (ver
  // parcTokenNum/parcTokenRaw), senão as duas taxas colidem na mesma chave EC|bandeira|0.
  const isVoucher=p.includes('VOUCHER')||p.includes('MULTIBENEFICIO');
  return{bandeira:b,isDebito,isVoucher};
}

// Token de parcela usado como sufixo da chave: 'V' = voucher/multibenefícios, 'D' = débito,
// '0' = à vista, ou o nº de parcelas (2..12). Voucher checado antes de débito/parcelas porque é
// informação mais específica (ex "Visa Voucher Multibenefícios" não é nem débito nem parcelado).
const parcTokenNum=(parc,isDebito,isVoucher)=>isVoucher?'V':isDebito?'D':parc>1?String(parc):'0';
const parcTokenRaw=(raw,isDebito,isVoucher)=>{
  const s=String(raw||'').toUpperCase();
  if(isVoucher||s.includes('VOUCHER')||s.includes('MULTIBENEFICIO')) return 'V';
  // "Débito" pode vir só na coluna Produto/Tipo de Venda (isDebito) OU só na própria célula de
  // Parcelas ("Débito") — checa as duas, senão a tarifa de débito da planilha de taxas cai no
  // mesmo balde de "à vista" e é sobrescrita.
  if(isDebito||s.includes('DEBITO')||s.includes('DÉBITO')) return 'D';
  if(s.includes('VISTA')) return '0';
  const n=parseInt(s.replace(/\D+/g,''))||0;
  return n>1?String(n):'0';
};

// Monta o mapa de taxas específicas por EC a partir da 3ª planilha (EC | Produto Cielo | Parcelas | Taxa %)
function buildTaxaMap(taxasRaw){
  const map={};
  (taxasRaw||[]).forEach(r=>{
    const ec=String(getCol(r,'EC','Número do EC','Numero do EC','EC Cielo','Estabelecimento')||'').trim();
    const produto=String(getCol(r,'Produto Cielo','Produto cielo','Bandeira','Produto')||'').trim();
    if(!ec||!produto) return;
    const tipoVenda=String(getCol(r,'Tipo da Venda','Tipo de Venda','Tipo Venda')||'').trim();
    const parcRaw=getCol(r,'Parcelas','Quantidade de Parcelas','Qtd Parcelas','Nº Parcelas','Parcela');
    const taxaRaw=String(getCol(r,'Taxa %','Taxa%','Taxa','Taxa Atual','Taxa atual')||'').replace('%','').trim();
    let taxa=pV(taxaRaw);
    // Quando a coluna vem formatada como % no Excel, o valor "cru" da célula é a fração (0,0085
    // pra 0,85%), não 0,85 — normaliza pra escala percentual (0.85, 4.86...) igual o resto do app
    // usa. Taxas reais da Cielo vão de ~0,8% a ~15%, então threshold 0.5 separa bem os dois casos.
    if(taxa>0&&taxa<0.5) taxa=Math.round(taxa*100*10000)/10000;
    if(!taxa) return;
    const{bandeira,isDebito,isVoucher}=detectBandeira(`${produto} ${tipoVenda}`);
    const tk=parcTokenRaw(parcRaw,isDebito,isVoucher);
    map[`${ec}|${bandeira}|${tk}`]=taxa;
  });
  return map;
}

// Resolve a taxa correta: 1º tenta a tarifa específica do EC (3ª planilha), senão cai na tabela padrão Cielo
function getTaxaCorreta(produto,parcelas,ec,taxaMap){
  const{bandeira,isDebito,isVoucher}=detectBandeira(produto);
  const parc=parseInt(parcelas)||0;
  const tk=parcTokenNum(parc,isDebito,isVoucher);
  const ecKey=ec?`${String(ec).trim()}|${bandeira}|${tk}`:null;
  if(ecKey&&taxaMap&&taxaMap[ecKey]!==undefined) return{taxa:taxaMap[ecKey],origem:'EC'};
  const padrao=TAXA_TABLE[`${bandeira}|${tk}`];
  if(padrao!==undefined) return{taxa:padrao,origem:'PADRAO'};
  return{taxa:null,origem:null};
}

// Reconstrói o grid original (título + linha em branco + cabeçalho real + dados) a partir do
// JSON que o SheetJS já gerou usando a linha 1 como header, e re-ancora no cabeçalho de verdade
// onde quer que ele esteja (linha 1, 2, 3...). Robusto tanto pra arquivos "limpos" (header já na
// linha 1) quanto pra exports do SIE que têm título + linha em branco antes do cabeçalho real.
function parseSIEFile(rawRows){
  if(!rawRows||!rawRows.length) return [];
  const keyRow=Object.keys(rawRows[0]);
  const grid=[keyRow,...rawRows.map(r=>Object.values(r))];
  const hIdx=grid.findIndex(row=>row.some(v=>String(v||'').includes('Produto')||String(v||'').includes('Flag MDR')));
  if(hIdx===-1) return rawRows;
  const hVals=grid[hIdx];
  return grid.slice(hIdx+1).map(row=>Object.fromEntries(hVals.map((h,i)=>[h||`_c${i}`,row[i]])))
    .filter(r=>Object.values(r).some(v=>v!==''&&v!=null));
}

function analyzeComissaoMinima(sieRaw,analRaw,taxasRaw){
  const taxaMap=buildTaxaMap(taxasRaw);
  const sieRows=parseSIEFile(sieRaw).filter(r=>{
    const flag=String(getCol(r,'Flag MDR Mínimo','Flag MDR','FLAG MDR MÍNIMO')||'').trim().toLowerCase();
    return flag==='sim';
  });
  const analByAuth={};
  analRaw.forEach(r=>{
    const auth=String(r['Autorização']||r['Autorizacao']||'').trim();
    if(auth) analByAuth[auth]=r;
  });
  const rows=sieRows.map(r=>{
    const produto=String(getCol(r,'Produto cielo','Produto Cielo')||'').trim();
    const parcelas=parseInt(getCol(r,'Quantidade de Parcelas','Parcelas')||0)||0;
    const vt=pV(String(getCol(r,'Valor da Transação','Valor da Transacao')||0));
    const vcb=pV(String(getCol(r,'Valor Comissão Bruta','Valor Comissao Bruta')||0));
    const pctDesc=pV(String(getCol(r,'Percentual de Desconto')||0));
    const auth=String(getCol(r,'Código Autorização (Transação)','Codigo Autorizacao')||'').trim();
    const dtTrans=getCol(r,'Data da Transação','Data da Transacao');
    const ec=String(getCol(r,'Número do EC','EC')||'').trim();
    const ro=String(getCol(r,'Número RO')||'').trim();
    const{taxa:taxaCorreta,origem:taxaOrigem}=getTaxaCorreta(produto,parcelas,ec,taxaMap);
    const comissaoCorreta=taxaCorreta!==null?Math.round(vt*(taxaCorreta/100)*10000)/10000:null;
    const difSistema=comissaoCorreta!==null?Math.round((vcb-comissaoCorreta)*10000)/10000:null;
    const anal=analByAuth[auth]||null;
    const taxaAnal=anal?pV(String(anal['Taxa %']||anal['Taxa%']||0)):null;
    const comissaoAnal=anal?pV(String(anal['Comissão']||anal['Comissao']||0)):null;
    const difAnal=anal?pV(String(anal['Diferença']||anal['Diferenca']||0)):null;
    const taxaMatch=taxaCorreta!==null&&taxaAnal!==null?Math.abs(taxaCorreta-taxaAnal)<0.001:null;
    const difMatch=difSistema!==null&&difAnal!==null?Math.abs(difSistema-difAnal)<0.005:null;
    const issues=[];
    if(taxaCorreta===null) issues.push('TAXA_NAO_ENCONTRADA');
    // "Sem tarifa do EC" é só informativo (mostrado no badge Origem + no card do topo) — não conta
    // como divergência sozinho, porque a tabela padrão pode dar o valor certo mesmo sem tarifa
    // específica do EC cadastrada. Só vira problema de verdade se o cálculo não bater com o analista.
    if(!anal) issues.push('SEM_ANALISTA');
    else if(taxaMatch===false) issues.push('TAXA_DIVERGE');
    else if(difMatch===false) issues.push('CALCULO_DIVERGE');
    return{ec,auth,ro,produto,parcelas,dtTrans,vt,vcb,pctDesc,taxaCorreta,taxaOrigem,comissaoCorreta,difSistema,taxaAnal,comissaoAnal,difAnal,taxaMatch,difMatch,hasAnal:!!anal,issues,ok:issues.length===0};
  });
  const byProd={};
  rows.forEach(r=>{
    const k=r.produto;
    if(!byProd[k]) byProd[k]={produto:k,qtd:0,totalVCB:0,totalCSistema:0,totalDifSistema:0,totalCAnal:0,totalDifAnal:0,divergencias:0};
    byProd[k].qtd++;
    byProd[k].totalVCB+=r.vcb;
    if(r.comissaoCorreta!==null) byProd[k].totalCSistema+=r.comissaoCorreta;
    if(r.difSistema!==null) byProd[k].totalDifSistema+=r.difSistema;
    if(r.comissaoAnal!==null) byProd[k].totalCAnal+=r.comissaoAnal;
    if(r.difAnal!==null) byProd[k].totalDifAnal+=r.difAnal;
    if(!r.ok) byProd[k].divergencias++;
  });
  // Reconciliação com a planilha do analista: linhas que existem lá mas não têm par no SIE
  // filtrado (Flag=Sim) — é isso que explica o total do sistema não bater com o total do
  // analista mesmo quando toda linha comparada está OK.
  const matchedAuths=new Set(rows.map(r=>r.auth));
  const analTotalGeral=Math.round(analRaw.reduce((s,r)=>s+pV(String(r['Diferença']||r['Diferenca']||0)),0)*10000)/10000;
  const analOnly=analRaw.map(r=>{
    const auth=String(r['Autorização']||r['Autorizacao']||'').trim();
    return{auth,r};
  }).filter(({auth})=>auth&&!matchedAuths.has(auth)).map(({auth,r})=>({
    auth,
    ec:String(r['EC']||r['Número do EC']||'').trim(),
    produto:String(r['Produto cielo']||r['Produto Cielo']||'').trim(),
    dtTrans:r['Data da Transação']||r['Data']||'',
    flag:String(r['Flag MDR Mínimo']||'').trim(),
    difAnal:pV(String(r['Diferença']||r['Diferenca']||0)),
  }));

  return{rows,summary:Object.values(byProd),analOnly,analTotalGeral};
}

// ===== Evento 7922 — Ressarcimento/Estorno de Cobrança de Ativos (Terminais não recuperados, D297) =====

// Conta dias úteis (desconsiderando sáb/dom/feriados) estritamente entre duas datas ISO
// "YYYY-MM-DD". Positivo quando b é depois de a, negativo quando b é antes de a. 0 = mesmo dia.
const diffBiz=(a,b)=>{
  if(!a||!b) return null;
  try{
    let da=new Date(a+"T12:00:00Z"),db=new Date(b+"T12:00:00Z");
    if(isNaN(da.getTime())||isNaN(db.getTime())) return null;
    let sign=1;
    if(da>db){const t=da;da=db;db=t;sign=-1;}
    const cur=new Date(da);let c=0;
    while(cur<db){
      cur.setUTCDate(cur.getUTCDate()+1);
      const k=cur.toISOString().slice(0,10),w=cur.getUTCDay();
      if(w!==0&&w!==6&&!BR_HOL.has(k)) c++;
    }
    return c*sign;
  }catch(e){return null;}
};

// Carrega o workbook inteiro (não converte nenhuma aba ainda) — usado no Evento 7922, onde a aba
// certa varia de nome mês a mês (ex: "7922 Agosto", "Planilha9") e precisa ser achada pelo
// cabeçalho, não pela posição/tamanho. CSV não tem conceito de aba, então vira uma "tabela única".
const loadWorkbookRaw=(file,enc,cb)=>{
  const ext=file.name.split(".").pop().toLowerCase();
  if(["xlsx","xlsb","xls"].includes(ext)){
    const fr=new FileReader();
    fr.onload=e=>{const wb=XLSX.read(e.target.result,{type:"array"});cb({type:"wb",wb});};
    fr.readAsArrayBuffer(file);
  }else{
    const fr=new FileReader();
    fr.onload=e=>cb({type:"rows",rows:Papa.parse(e.target.result,{header:true,delimiter:";",skipEmptyLines:true}).data});
    fr.readAsText(file,enc);
  }
};

// Exports de sistema legado costumam vir com o "!ref" da aba inflado até a última linha do Excel
// (1.048.576) mesmo tendo só algumas centenas de linhas reais — converter isso com sheet_to_json
// direto processaria mais de 1 milhão de linhas à toa (dezenas de segundos a minutos, travando o
// navegador). sheetLastRow acha a última linha realmente preenchida escaneando só as chaves do
// objeto esparso da planilha (rápido, independe do "!ref"), e sheetToJsonFast restringe a
// conversão a esse intervalo real antes de chamar o parser.
const sheetLastRow=ws=>{
  let maxR=-1;
  Object.keys(ws||{}).forEach(addr=>{
    if(addr[0]==="!") return;
    const v=ws[addr]?.v;
    if(v===undefined||v==="") return;
    const c=XLSX.utils.decode_cell(addr);
    if(c.r>maxR) maxR=c.r;
  });
  return maxR;
};
const sheetHeaderRow=ws=>{
  if(!ws||!ws["!ref"]) return[];
  const range=XLSX.utils.decode_range(ws["!ref"]);
  const headers=[];
  for(let C=range.s.c;C<=range.e.c;C++){
    const cell=ws[XLSX.utils.encode_cell({r:range.s.r,c:C})];
    headers.push(cell?.v!==undefined?String(cell.v).trim():"");
  }
  return headers;
};
const sheetToJsonFast=ws=>{
  if(!ws||!ws["!ref"]) return[];
  const range=XLSX.utils.decode_range(ws["!ref"]);
  const lastRow=sheetLastRow(ws);
  if(lastRow<0) return[];
  const restricted={...range,e:{...range.e,r:lastRow}};
  return XLSX.utils.sheet_to_json(ws,{defval:"",raw:true,range:restricted});
};

const headerMatches=(keys,requiredHeaders)=>requiredHeaders.every(h=>{
  const nh=normStr(h);
  return keys.some(k=>k===nh||(nh.length>=6&&k.includes(nh)));
});

// Acha, num workbook, a MELHOR aba cujo cabeçalho contenha as colunas pedidas — só lê o cabeçalho
// (barato) de cada aba antes de decidir, e só converte de verdade a aba escolhida. Quando mais de
// uma aba bate (planilhas de controle costumam ter uma aba "mestre"/histórico com o mesmo layout
// da aba do mês/evento específico), prioriza pelo NOME (preferName, ex "7922") e, sem nenhuma
// batendo por nome, pela de menos linhas reais — a mestre/histórico é sempre a maior das duas.
const findBestSheet=(wb,requiredHeaders,preferName)=>{
  const cands=[];
  (wb?.SheetNames||[]).forEach(name=>{
    const ws=wb.Sheets[name];
    if(!ws||!ws["!ref"]) return;
    const keys=sheetHeaderRow(ws).map(normStr);
    if(headerMatches(keys,requiredHeaders)) cands.push({name,ws});
  });
  if(!cands.length) return null;
  if(preferName){
    const byName=cands.find(c=>normStr(c.name).includes(normStr(preferName)));
    if(byName) return sheetToJsonFast(byName.ws);
  }
  let best=null,bestRows=Infinity;
  cands.forEach(c=>{const r=sheetLastRow(c.ws);if(r>=0&&r<bestRows){bestRows=r;best=c;}});
  return sheetToJsonFast((best||cands[0]).ws);
};

// Procura as linhas certas em qualquer um dos arquivos carregados (workbook ou CSV já em linhas) —
// assim não importa em qual dos dois slots o usuário soltou qual arquivo.
const findRowsAcross=(dataList,requiredHeaders,preferName)=>{
  for(const data of dataList){
    if(!data) continue;
    if(data.type==="rows"){
      const keys=Object.keys(data.rows?.[0]||{}).map(normStr);
      if(headerMatches(keys,requiredHeaders)) return data.rows;
    }else if(data.type==="wb"){
      const rows=findBestSheet(data.wb,requiredHeaders,preferName);
      if(rows) return rows;
    }
  }
  return[];
};

// Distância de edição (Levenshtein) — usada só pra tolerar erro de digitação de 1-2 letras numa
// palavra-chave (ex "conta" em vez de "consta"), não pra parear frases inteiras.
function levenshtein(a,b){
  const m=a.length,n=b.length;
  if(!m) return n; if(!n) return m;
  let prev=Array.from({length:n+1},(_,j)=>j);
  for(let i=1;i<=m;i++){
    const cur=[i];
    for(let j=1;j<=n;j++) cur[j]=a[i-1]===b[j-1]?prev[j-1]:1+Math.min(prev[j-1],prev[j],cur[j-1]);
    prev=cur;
  }
  return prev[n];
}
// true se ALGUMA palavra do texto está a 1-2 letras de "target" (erro de digitação pequeno tipo
// letra faltando/trocada/a mais) — palavras de até 5 letras toleram 1 erro, mais longas toleram 2.
const fuzzyHasWord=(s,target)=>{
  const maxDist=target.length<=5?1:2;
  return s.split(/\s+/).some(w=>w.length>=3&&Math.abs(w.length-target.length)<=maxDist&&levenshtein(w,target)<=maxDist);
};

// Agrupa "Motivo se Improcedente" (coluna W do controle) nos 4 baldes pedidos, tolerando
// variações reais de digitação da planilha: maiúsculas, "D97" em vez de "D297", e erro de
// digitação pequeno numa palavra-chave (ex "conta" em vez de "consta") — só cai em "Outros"
// quando o texto é bem diferente dos 4 motivos esperados, não por causa de 1 letra errada.
function normMotivoImprocedente(m){
  const s=normStr(m);
  if(!s) return"Outros";
  if(s.includes("acionamento")||fuzzyHasWord(s,"acionamento")) return"Acionamento indevido";
  if(s.includes("duplicidade")||fuzzyHasWord(s,"duplicidade")) return"Evento aberto em duplicidade";
  if(s.includes("consta estorno")||(fuzzyHasWord(s,"consta")&&fuzzyHasWord(s,"estorno"))) return"Já consta estorno";
  const temConsta=s.includes("consta")||fuzzyHasWord(s,"consta");
  const temRef=s.includes("d297")||s.includes("d97")||s.includes("cobranca")||s.includes("debito")||fuzzyHasWord(s,"cobranca")||fuzzyHasWord(s,"debito");
  if(temConsta&&temRef) return"Não consta cobrança D297";
  return"Outros";
}

function analyze7922(ajustesData,controleData){
  const sources=[ajustesData,controleData];
  const ajusteRows=findRowsAcross(sources,["EC","Código + Motivo de ajuste","Número RO"]);
  const controleRows=findRowsAcross(sources,["Protocolo","Procedente/ Improcedente","Estabelecimento"],"7922");

  // --- Planilha de Ajustes (crédito D297) ---
  const ajustes=ajusteRows.map(r=>{
    const ec=String(getCol(r,"EC")||"").trim();
    const codigoMotivo=String(getCol(r,"Código + Motivo de ajuste","Codigo + Motivo de ajuste")||"").trim();
    const is297=/297/.test(codigoMotivo);
    const bandeira=String(getCol(r,"Bandeira")||"").trim();
    const isVisa=/VISA/i.test(bandeira);
    const valor=pV(String(getCol(r,"Valor total do ajuste")||0));
    const ro=String(getCol(r,"Número RO","Numero RO")||"").trim();
    const obs=String(getCol(r,"Observações","Observacoes")||"").trim();
    // A Observação começa com um número de referência do estorno (ex "00889227 - Estorno
    // Cobranca de Ativo - OUTUBRO DE 2022", às vezes separado por "|" em vez de "-") — não é o
    // número lógico da máquina, é o identificador que diferencia dois ajustes que por coincidência
    // têm o mesmo RO + Valor + EC (mesmo cliente, mesmo valor padrão, meses diferentes de cobrança).
    const obsRef=(obs.match(/^(\d+)/)||[])[1]||"";
    const dtCriacao=parseAny(getCol(r,"Data de criação","Data de criacao"));
    const solicitacao=String(getCol(r,"Solicitação","Solicitacao")||"").trim();
    const tipoAjuste=String(getCol(r,"Tipo de ajuste")||"").trim();
    return{solicitacao,ec,codigoMotivo,is297,bandeira,isVisa,valor,ro,obs,obsRef,dtCriacao,tipoAjuste};
  }).filter(a=>a.is297);

  // Duplicidade: mesma combinação de Número RO (X) + Valor (Q) + EC (E) *e* mesma referência da
  // Observação (M) aparecendo mais de uma vez. RO+Valor+EC sozinhos dão falso positivo quando o
  // mesmo cliente tem duas cobranças D297 de meses diferentes com o mesmo RO/valor — só é
  // duplicidade de verdade quando a referência da Observação também é igual.
  const dupKeyCount={};
  ajustes.forEach(a=>{const k=`${a.ro}|${a.valor}|${a.ec}|${a.obsRef}`;dupKeyCount[k]=(dupKeyCount[k]||0)+1;});
  ajustes.forEach(a=>{a.isDup=dupKeyCount[`${a.ro}|${a.valor}|${a.ec}|${a.obsRef}`]>1;});

  const ajustesByEc={};
  ajustes.forEach(a=>{(ajustesByEc[a.ec]=ajustesByEc[a.ec]||[]).push(a);});

  // --- Planilha de Controle dos Analistas (Tabulador Consultoria Financeira) ---
  const controle=controleRows.map(r=>{
    const protocolo=String(getCol(r,"Protocolo")||"").trim();
    const ec=String(getCol(r,"Estabelecimento")||"").trim();
    const analista=String(getCol(r,"Analista")||"").trim();
    const procImprocRaw=String(getCol(r,"Procedente/ Improcedente","Procedente/Improcedente")||"").trim();
    const procedente=/^procedente$/i.test(procImprocRaw);
    const improcedente=/improcedente/i.test(procImprocRaw);
    const motivoW=String(getCol(r,"Motivo se Improcedente")||"").trim();
    const obsX=String(getCol(r,"Observação","Observacao")||"").trim();
    const dtFinalizada=parseAny(getCol(r,"Data Finalizada"));
    return{protocolo,ec,analista,procImprocRaw,procedente,improcedente,motivoW,obsX,dtFinalizada};
  }).filter(c=>c.procedente||c.improcedente);

  // --- PROCEDENTE: casa por EC com o ajuste D297 (o mais próximo, em data de criação, da Data
  // Finalizada) e valida bandeira VISA, D+2 dias úteis (Data Finalizada -> Data de criação do
  // ajuste, desconsiderando feriados) e duplicidade ---
  const procedentes=controle.filter(c=>c.procedente).map(c=>{
    const cands=ajustesByEc[c.ec]||[];
    let match=null;
    if(cands.length===1) match=cands[0];
    else if(cands.length>1){
      match=c.dtFinalizada
        ?cands.reduce((best,cur)=>{
          if(!cur.dtCriacao) return best;
          if(!best) return cur;
          return Math.abs(new Date(cur.dtCriacao)-new Date(c.dtFinalizada))<Math.abs(new Date(best.dtCriacao)-new Date(c.dtFinalizada))?cur:best;
        },null)
        :cands[0];
    }
    const bd=match&&c.dtFinalizada&&match.dtCriacao?diffBiz(c.dtFinalizada,match.dtCriacao):null;
    const d2Ok=bd===null?null:(bd>=0&&bd<=2);
    const bandeiraOk=match?match.isVisa:null;
    const issues=[];
    if(!match) issues.push("SEM_AJUSTE_D297");
    else{
      if(bandeiraOk===false) issues.push("BANDEIRA_ERRADA");
      if(d2Ok===false) issues.push("FORA_D2");
      if(match.isDup) issues.push("DUPLICIDADE");
    }
    return{...c,match,outrosAjustes:cands.length>1?cands:[],bd,d2Ok,bandeiraOk,issues,ok:issues.length===0};
  });

  // --- IMPROCEDENTE: agrupa pela coluna W (Motivo se Improcedente) nos 4 baldes + Outros,
  // trazendo a observação da coluna X pra cada um ---
  const MOTIVOS=["Acionamento indevido","Evento aberto em duplicidade","Já consta estorno","Não consta cobrança D297","Outros"];
  const improcedentes=controle.filter(c=>c.improcedente).map(c=>({...c,motivoNorm:normMotivoImprocedente(c.motivoW)}));
  const porMotivo=MOTIVOS.map(m=>({motivo:m,itens:improcedentes.filter(i=>i.motivoNorm===m)}));

  return{procedentes,improcedentes,porMotivo,ajustes,ajustesDup:ajustes.filter(a=>a.isDup)};
}

const MODULES=[
  {id:"5125",name:"Evento 5125",group:"Eventos",icon:"⚡",desc:"Cancelamento sem saldo · Boleto / PIX",slots:[{key:"ctrl",label:"Planilha Controle (analistas)",enc:"latin1"},{key:"ga",label:"Relatório G.A — Gestor de Ajustes",enc:"ISO-8859-1"}],canRun:s=>s.ctrl?.length>0&&s.ga?.length>0,run:s=>analyze5125(s.ga,s.ctrl),is5125:true},
  {id:"7922",name:"Evento 7922",group:"Eventos",icon:"📋",desc:"Ressarcimento de Ativos (terminais não recuperados) — D297 · Crédito VISA · SLA D+2",
    slots:[
      {key:"ajustes",label:"Planilha de Ajustes 7922 (D297)",enc:"UTF-8",allSheets:true},
      {key:"controle",label:"Controle dos Analistas (Tabulador Consultoria Financeira)",enc:"UTF-8",allSheets:true},
    ],
    canRun:s=>!!s.ajustes&&!!s.controle,
    run:s=>analyze7922(s.ajustes,s.controle),
    is7922:true},
  {id:"9066",name:"Evento 9066",group:"Eventos",icon:"🔧",desc:"Sinistro · Perda ou Roubo de Maquininha · D+4",slots:[{key:"ctrl",label:"Controle Sinistro (analistas)",enc:"latin1"},{key:"ga",label:"Ajustes G.A",enc:"UTF-8"}],canRun:s=>s.ctrl?.length>0&&s.ga?.length>0,run:s=>analyze9066(s.ga,s.ctrl),is9066:true},
  {id:"reg-fin",name:"Regularizações Financeiras",group:"Caixas de E-mail",icon:"💼",desc:"Comissão Mínima — MDR incorreto cobrado ao cliente",
    slots:[
      {key:"sie",label:"Faturamento Contábil EC — SIE",enc:"UTF-8"},
      {key:"anal",label:"Planilha do Analista (com cálculo)",enc:"UTF-8"},
      {key:"taxas",label:"Tabela de Taxas por EC (EC | Produto Cielo | Parcelas | Taxa %)",enc:"UTF-8"},
    ],
    canRun:s=>s.sie?.length>0&&s.anal?.length>0&&s.taxas?.length>0,
    run:s=>analyzeComissaoMinima(s.sie,s.anal,s.taxas),
    isComissao:true},
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

const BADGES={OK:{bg:"rgba(0,230,118,.15)",fg:"#00e676",txt:"✓ OK"},ONTIME:{bg:"rgba(0,230,118,.15)",fg:"#00e676",txt:"NO PRAZO"},LATE:{bg:"rgba(255,82,82,.15)",fg:"#ff5252",txt:"ATRASADO"},DUP:{bg:"rgba(255,171,64,.15)",fg:"#ffab40",txt:"DUPLICATA"},SEM_CAN:{bg:"rgba(255,82,82,.15)",fg:"#ff5252",txt:"SEM CAN"},SLA_CAN:{bg:"rgba(255,171,64,.15)",fg:"#ffab40",txt:"ℹ️ CAN"},SLA_BCK:{bg:"rgba(255,82,82,.15)",fg:"#ff5252",txt:"⏰ BCK"},SCHED:{bg:"rgba(0,180,216,.15)",fg:"#00b4d8",txt:"AGENDADO"},VALOR_DIFF:{bg:"rgba(255,171,64,.15)",fg:"#ffab40",txt:"⚠️ VALOR"},PEND:{bg:T.hover,fg:T.gray,txt:"—"}};
const Badge=({type})=>{const s=BADGES[type]||BADGES.PEND;return<span style={{display:"inline-block",padding:"3px 8px",borderRadius:20,fontSize:10,fontWeight:700,letterSpacing:.5,background:s.bg,color:s.fg,marginRight:3,whiteSpace:"nowrap"}}>{s.txt}</span>;};

const Login=()=>{
  const[email,setEmail]=useState("");const[pass,setPass]=useState("");const[err,setErr]=useState("");const[busy,setBusy]=useState(false);
  const go=async()=>{if(!email||!pass){setErr("Preencha e-mail e senha.");return;}setBusy(true);setErr("");try{await signInWithEmailAndPassword(auth,email,pass);}catch(e){const m={"auth/invalid-credential":"Credenciais inválidas.","auth/user-not-found":"Usuário não encontrado.","auth/wrong-password":"Senha incorreta.","auth/too-many-requests":"Muitas tentativas. Aguarde."};setErr(m[e.code]||"Erro ao autenticar.");setBusy(false);}};
  return(<div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:`linear-gradient(135deg,hsl(228,35%,8%) 0%,hsl(228,30%,13%) 100%)`,fontFamily:T.font}}>
  <div style={{width:420,padding:"48px 40px",background:T.surface,borderRadius:20,boxShadow:"0 24px 80px rgba(0,0,0,.6)",border:`1px solid ${T.border}`}}>
    <div style={{textAlign:"center",marginBottom:36}}>
      <div style={{width:72,height:72,borderRadius:18,background:"linear-gradient(135deg,#00b4d8,#00e676)",display:"inline-flex",alignItems:"center",justifyContent:"center",fontSize:32,marginBottom:18,boxShadow:"0 8px 24px rgba(0,180,216,.4)"}}>◈</div>
      <h1 style={{margin:0,fontSize:26,fontWeight:800,color:T.accent,letterSpacing:-.5}}>Painel de Ajustes</h1>
      <p style={{margin:"6px 0 0",fontSize:12,color:T.gray,letterSpacing:.5}}>CONTROLE · AUDITORIA · ANÁLISE</p>
    </div>
    {[{l:"E-mail",t:"email",v:email,s:setEmail,ph:"seu.email@empresa.com.br"},{l:"Senha",t:"password",v:pass,s:setPass,ph:"Digite sua senha",ip:true}].map(({l,t,v,s,ph,ip})=>(
      <div key={l} style={{marginBottom:16}}>
        <label style={{display:"block",fontSize:13,fontWeight:600,color:T.gray,marginBottom:8}}>{l}</label>
        <input type={t} value={v} onChange={e=>{s(e.target.value);if(ip)setErr("");}} onKeyDown={e=>e.key==="Enter"&&go()} placeholder={ph}
          style={{width:"100%",padding:"14px 16px",background:T.card,border:`1.5px solid ${ip&&err?T.danger:T.border}`,borderRadius:10,color:T.white,fontSize:14,outline:"none",boxSizing:"border-box",WebkitTextFillColor:T.white}}/>
      </div>
    ))}
    {err&&<p style={{color:T.danger,fontSize:12,margin:"-8px 0 14px"}}>{err}</p>}
    <button onClick={go} disabled={busy} style={{width:"100%",padding:"15px",background:busy?T.muted:"linear-gradient(to right,#00b4d8,#00e676)",color:busy?T.gray:"#0a1628",border:"none",borderRadius:50,fontSize:16,fontWeight:800,cursor:busy?"not-allowed":"pointer",boxShadow:busy?"none":"0 4px 20px rgba(0,180,216,.4)",letterSpacing:.5,marginTop:4}}>
      {busy?"Verificando…":"Entrar"}
    </button>
    <p style={{textAlign:"center",fontSize:11,color:T.muted,margin:"20px 0 0"}}>Uso Interno · Acesso Restrito</p>
  </div>
</div>);
};

const Sidebar=({activeId,onSelect})=>(<div style={{width:240,background:T.sidebar,flexShrink:0,overflowY:"auto",paddingTop:8}}>{GROUPS.map(g=>(<div key={g}><div style={{fontSize:10,fontWeight:700,color:T.muted,letterSpacing:1.5,padding:"16px 24px 6px",textTransform:"uppercase"}}>{g}</div>{MODULES.filter(m=>m.group===g).map(m=>(<button key={m.id} onClick={()=>onSelect(m.id)} style={{display:"flex",alignItems:"center",gap:12,width:"100%",padding:"10px 24px",border:"none",textAlign:"left",cursor:"pointer",background:activeId===m.id?T.sidebarAccent+"":"transparent",color:activeId===m.id?T.white:T.gray,fontSize:13,fontWeight:activeId===m.id?700:400,borderLeft:activeId===m.id?`3px solid ${T.accent}`:"3px solid transparent"}}><span style={{fontSize:16}}>{m.icon}</span><span>{m.name}</span></button>))}</div>))}</div>);

const UploadZone=({label,count,countLabel,onFile,enc,allSheets})=>{const ref=useRef();return(<div onClick={()=>ref.current?.click()} style={{background:T.card,borderRadius:10,border:`1px dashed ${count?T.accent:T.muted}`,padding:"16px 20px",cursor:"pointer"}}><input ref={ref} type="file" accept=".csv,.xlsx,.xlsb,.xls" style={{display:"none"}} onChange={e=>e.target.files[0]&&(allSheets?loadWorkbookRaw(e.target.files[0],enc,onFile):loadFile(e.target.files[0],enc,onFile))}/><div style={{fontSize:10,fontWeight:700,color:count?T.accent:T.gray,letterSpacing:.8,marginBottom:4}}>{label.toUpperCase()}</div><div style={{fontSize:12,color:count?T.accent:T.muted}}>{count?`✅ ${count} ${countLabel||"registros carregados"}`:"📎 CSV · XLSX · XLSB"}</div></div>);};

const Stat=({label,value,color,icon,active,onClick})=>(<div onClick={onClick} style={{background:active?`${color}22`:T.card,borderRadius:12,padding:"18px 16px",boxShadow:active?`0 0 0 2px ${color},0 4px 16px rgba(0,0,0,.4)`:`0 4px 16px rgba(0,0,0,.4)`,position:"relative",overflow:"hidden",cursor:"pointer",transition:"all .2s"}}><div style={{position:"absolute",top:-10,right:-10,fontSize:48,opacity:.06}}>{icon}</div><div style={{fontSize:30,fontWeight:900,color,lineHeight:1}}>{value}</div><div style={{fontSize:11,color:T.gray,marginTop:6,letterSpacing:.2}}>{label}</div>{active&&<div style={{position:"absolute",bottom:0,left:0,right:0,height:3,background:color,borderRadius:"0 0 12px 12px"}}/>}</div>);

const GenericTable=({data,moduleId})=>{const[search,setSearch]=useState("");const cols=data.length>0?Object.keys(data[0]).filter(k=>k!==""):[];const rows=useMemo(()=>{if(!search.trim())return data;const s=search.toLowerCase();return data.filter(r=>Object.values(r).some(v=>String(v).toLowerCase().includes(s)));},[data,search]);const doExport=()=>{const ws=XLSX.utils.json_to_sheet(rows);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Dados");XLSX.writeFile(wb,`export_${moduleId}_${TODAY}.xlsx`);};return(<div><div style={{display:"flex",gap:10,alignItems:"center",marginBottom:14,flexWrap:"wrap"}}><input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar…" style={{flex:1,minWidth:200,padding:"10px 14px",background:T.card,border:`1px solid ${T.border}`,borderRadius:8,color:T.white,fontSize:13,outline:"none"}}/><span style={{fontSize:12,color:T.gray}}>{rows.length}/{data.length}</span><button onClick={doExport} style={{padding:"10px 20px",background:T.accent,color:"#000",border:"none",borderRadius:50,fontSize:12,fontWeight:700,cursor:"pointer"}}>⬇ Exportar</button></div><div style={{background:T.card,borderRadius:12,overflow:"hidden"}}><div style={{overflowX:"auto",maxHeight:"55vh",overflowY:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}><thead style={{position:"sticky",top:0,zIndex:2}}><tr style={{background:T.surface}}>{cols.map(h=><th key={h} style={{padding:"11px 12px",textAlign:"left",fontWeight:700,color:T.gray,fontSize:10,letterSpacing:.8,whiteSpace:"nowrap",borderBottom:`1px solid ${T.border}`}}>{h}</th>)}</tr></thead><tbody>{rows.map((row,i)=>(<tr key={i} style={{background:i%2===0?T.card:T.card,borderBottom:`1px solid ${T.border}`}}>{cols.map(k=><td key={k} style={{padding:"9px 12px",color:T.white,whiteSpace:"nowrap",maxWidth:200,overflow:"hidden",textOverflow:"ellipsis"}}>{String(row[k]||"")}</td>)}</tr>))}</tbody></table>{rows.length===0&&<div style={{textAlign:"center",padding:40,color:T.muted}}>Nenhum registro encontrado.</div>}</div></div></div>);};

const View5125=({results,onExport})=>{
  const[search,setSearch]=useState("");const[onlyIssues,setOnlyIssues]=useState(false);const[expanded,setExpanded]=useState(null);
  const stats=useMemo(()=>({total:results.length,ok:results.filter(r=>r.ok).length,issues:results.filter(r=>!r.ok).length,dup:results.filter(r=>r.isDup).length,slaCan:results.filter(r=>r.canOk===false).length,slaBck:results.filter(r=>r.bckOk===false&&!r.bdScheduled).length,semCan:results.filter(r=>r.issues.includes("SEM_CAN")).length,agendado:results.filter(r=>r.bdScheduled).length}),[results]);
  const[activeFilter,setActiveFilter]=useState(null);
  const toggleFilter=f=>setActiveFilter(af=>af===f?null:f);
  const shown=useMemo(()=>{
    let r=results;
    if(activeFilter==="ok")r=r.filter(x=>x.ok);
    else if(activeFilter==="issues")r=r.filter(x=>!x.ok);
    else if(activeFilter==="dup")r=r.filter(x=>x.isDup);
    else if(activeFilter==="canTardio")r=r.filter(x=>x.canOk===false);
    else if(activeFilter==="slaBck")r=r.filter(x=>x.bckOk===false&&!x.bdScheduled);
    else if(activeFilter==="semCan")r=r.filter(x=>x.issues.includes("SEM_CAN"));
    else if(activeFilter==="agendado")r=r.filter(x=>x.bdScheduled);
    if(onlyIssues)r=r.filter(x=>!x.ok);
    if(search.trim()){const s=search.toLowerCase();r=r.filter(x=>x.ref.includes(s)||x.ec.includes(s)||x.auth.toLowerCase().includes(s)||x.analista.toLowerCase().includes(s));}
    return r;
  },[results,search,onlyIssues,activeFilter]);
  const TH=({c})=><th style={{padding:"11px 12px",textAlign:"left",fontWeight:700,color:T.gray,fontSize:10,letterSpacing:.8,whiteSpace:"nowrap",borderBottom:`1px solid ${T.border}`}}>{c}</th>;
  return(<div>
    <div style={{display:"grid",gridTemplateColumns:"repeat(8,1fr)",gap:10,marginBottom:24}}>
      {[
          [stats.total,"Total",T.accent,"📊",null],
          [stats.ok,"OK",T.success,"✅","ok"],
          [stats.issues,"Pendência",T.danger,"⚠️","issues"],
          [stats.dup,"Duplicata",T.warning,"🔁","dup"],
          [stats.slaCan,"CAN Tardio",T.muted,"ℹ️","canTardio"],
          [stats.slaBck,"SLA BCK",T.danger,"⏰","slaBck"],
          [stats.semCan,"Sem CAN",T.purple,"❌","semCan"],
          [stats.agendado,"Agendado",T.accent,"📅","agendado"],
        ].map(([v,l,clr,ic,fk])=><Stat key={l} label={l} value={v} color={clr} icon={ic}
          active={activeFilter===fk}
          onClick={()=>fk?toggleFilter(fk):setActiveFilter(null)}
        />)}
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
          <thead style={{position:"sticky",top:0,zIndex:2}}><tr style={{background:T.surface}}>{["Referência","EC","Autorização","Data Venda","Valor","Data Abertura","Analista","Data CAN","Prazo CAN","CAN Info","Data BCK","Prazo BCK","SLA BCK","Situação"].map(c=><TH key={c} c={c}/>)}</tr></thead>
          <tbody>
            {shown.map((r,i)=>(<>
              <tr key={`r${i}`} onClick={()=>setExpanded(expanded===i?null:i)} style={{background:r.isDup?"hsl(38,92%,8%)":!r.ok?"hsl(0,62.8%,8%)":i%2===0?T.card:T.hover,borderBottom:`1px solid ${T.border}`,cursor:"pointer"}}>
                <td style={{padding:"9px 12px",fontWeight:700,color:T.accent}}>{r.ref||"—"}</td>
                <td style={{padding:"9px 12px",fontFamily:"monospace",fontSize:10,color:T.gray}}>{r.ec}</td>
                <td style={{padding:"9px 12px",fontFamily:"monospace",color:T.white}}>{r.auth||"—"}</td>
                <td style={{padding:"9px 12px",color:T.white}}>{fD(r.sd)}</td>
                <td style={{padding:"9px 12px",fontWeight:700,color:T.white}}>{fV(r.valor)}</td>
                <td style={{padding:"9px 12px",color:T.white}}>{fD(r.od)}</td>
                <td style={{padding:"9px 12px",color:T.gray}}>{r.analista||"—"}</td>
                <td style={{padding:"9px 12px",color:r.canOk===false?T.warning:r.canOk?T.success:T.muted}}>{fD(r.canDate)}</td>
                <td style={{padding:"9px 12px",color:T.muted}}>{fD(r.canDl)}</td>
                <td style={{padding:"9px 12px"}}><Badge type={r.canOk===true?"ONTIME":r.canOk===false?"SLA_CAN":"PEND"}/></td>
                <td style={{padding:"9px 12px",color:r.bckStatus==="SCHED"?T.accent:r.bckOk===false?T.danger:r.bckOk?T.success:T.muted}}>{fD(r.bd)}</td>
                <td style={{padding:"9px 12px",color:T.muted}}>{fD(r.bckDl)}</td>
                <td style={{padding:"9px 12px"}}><Badge type={r.bckStatus==="SCHED"?"SCHED":r.bckOk===true?"ONTIME":r.bckOk===false?"LATE":"PEND"}/></td>
                <td style={{padding:"9px 12px"}}>{r.ok?<Badge type="OK"/>:r.issues.map(t=><Badge key={t} type={t}/>)}</td>
              </tr>
              {expanded===i&&(<tr key={`e${i}`} style={{background:T.bg}}><td colSpan={14} style={{padding:"16px 20px"}}>
                {/* Duplicate group detail */}
                {r.isDup&&r.dupGroupDetail?.length>0&&(
                  <div style={{marginBottom:14,border:`1px solid ${T.warning}44`,borderRadius:10,overflow:"hidden"}}>
                    <div style={{background:`${T.warning}22`,padding:"8px 14px",display:"flex",alignItems:"center",gap:8}}>
                      <span style={{color:T.warning,fontWeight:700,fontSize:12}}>⚠️ DUPLICATA DETECTADA</span>
                      <span style={{color:T.gray,fontSize:11}}>Mesmos campos: {fD(r.sd)} · {fV(r.valor)} · {r.auth} · Cartão {r.cartao}</span>
                    </div>
                    <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                      <thead><tr style={{background:`${T.warning}11`}}>
                        {["Referência","Val. Cancelamento","CAN (97)","Data CAN","BCK (962)","Data BCK","Val. Ajuste GA","✓ Match"].map(h=>(
                          <th key={h} style={{padding:"7px 10px",textAlign:"left",color:T.gray,fontSize:10,letterSpacing:.5,fontWeight:700,borderBottom:`1px solid ${T.border}`}}>{h}</th>
                        ))}
                      </tr></thead>
                      <tbody>
                        {r.dupGroupDetail.map((g,gi)=>{
                          const match=g.bckValor!==null&&g.cval>0?Math.abs(g.bckValor-g.cval)<0.05:null;
                          return(
                            <tr key={gi} style={{background:gi%2===0?T.card:T.hover,borderBottom:`1px solid ${T.border}`}}>
                              <td style={{padding:"7px 10px",fontWeight:700,color:g.ref===r.ref?T.accent:T.white}}>{g.ref}{g.ref===r.ref?" ◀":""}</td>
                              <td style={{padding:"7px 10px",color:T.white}}>{fV(g.cval)}</td>
                              <td style={{padding:"7px 10px",color:g.hasCAN?T.success:T.danger}}>{g.hasCAN?"✅ Sim":"❌ Não"}</td>
                              <td style={{padding:"7px 10px",color:T.gray}}>{fD(g.canDate)}</td>
                              <td style={{padding:"7px 10px",color:g.hasBCK?T.success:T.danger}}>{g.hasBCK?"✅ Sim":"❌ Não"}</td>
                              <td style={{padding:"7px 10px",color:T.gray}}>{fD(g.bckDate)}</td>
                              <td style={{padding:"7px 10px",color:T.white}}>{g.bckValor!==null?fV(g.bckValor):"—"}</td>
                              <td style={{padding:"7px 10px"}}>{match===true?<span style={{color:T.success}}>✅</span>:match===false?<span style={{color:T.danger}}>⚠️ {fV(g.cval)}</span>:<span style={{color:T.muted}}>—</span>}</td>
                            </tr>
                          );
                        })}
                        <tr style={{background:`${T.warning}11`,borderTop:`2px solid ${T.warning}44`}}>
                          <td colSpan={2} style={{padding:"7px 10px",color:T.warning,fontWeight:700,fontSize:11}}>
                            Total cancelamentos: {fV(r.dupGroupDetail.reduce((s,g)=>s+g.cval,0))}
                            {r.dupGroupDetail.reduce((s,g)=>s+g.cval,0)>r.valor+0.05?
                              <span style={{color:T.danger,marginLeft:8}}>⚠️ Excede valor original {fV(r.valor)}</span>:
                              <span style={{color:T.success,marginLeft:8}}>✅ Dentro do valor original {fV(r.valor)}</span>
                            }
                          </td>
                          <td colSpan={6} style={{padding:"7px 10px",color:T.muted,fontSize:11}}>
                            Valor original da transação: <strong style={{color:T.white}}>{fV(r.valor)}</strong>
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                )}
                {/* Summary row */}
                <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8,marginBottom:14}}>
                  {[["Valor Cancelamento",fV(r.cval)],["Tipo Pagamento",r.tipoPag||"—"],["Número Boleto",r.boleto||"—"],["Cartão (4 dígitos)",r.cartao||"—"],["Ajuste Efetuado",r.ajuste||"—"],["Transf. 3943",r.trans3943||"—"],["Dias abert.→CAN",r.canDate&&r.od?`${Math.round((new Date(r.canDate)-new Date(r.od))/86400000)} dias`:"—"],["Dias CAN→BCK",r.canDate&&r.bd?`${Math.round((new Date(r.bd)-new Date(r.canDate))/86400000)} dias`:"—"]].map(([l,v])=>(
                    <div key={l} style={{background:T.card,padding:"10px 12px",borderRadius:8,border:`1px solid ${T.border}`}}><div style={{fontSize:9,color:T.muted,marginBottom:3,letterSpacing:.5}}>{l.toUpperCase()}</div><div style={{fontWeight:700,color:T.white,fontSize:12}}>{v}</div></div>
                  ))}
                </div>
                {/* Observação do controle */}
                {r.obs&&<div style={{background:T.card,padding:"10px 14px",borderRadius:8,border:`1px solid ${T.border}`,marginBottom:10}}><span style={{fontSize:10,color:T.muted,letterSpacing:.5}}>OBSERVAÇÃO (CONTROLE) </span><span style={{color:"#facc15",fontWeight:600,fontSize:12}}>{r.obs}</span></div>}
                {/* CAN details from GA */}
                {r._ga&&<div style={{marginBottom:10}}><div style={{fontSize:10,fontWeight:700,color:T.accent,letterSpacing:.8,marginBottom:6}}>◈ DETALHES CAN — G.A</div>
                  <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:6}}>
                    {[["Departamento",getCol(r._ga,"Departamento")],["Solicitante",getCol(r._ga,"Solicitante")],["Status Solicitação",getCol(r._ga,"Status da solicitacao","Status da solicita\u00e7\u00e3o")],["Tipo Exceção",getCol(r._ga,"Tipo de excecao","Tipo de exce\u00e7\u00e3o")||getIdx(r._ga,25)],["Motivo Cancelamento",getCol(r._ga,"Motivo do cancelamento")||getIdx(r._ga,26)],["Bandeira",getCol(r._ga,"Bandeira")],["Produto",getCol(r._ga,"Produto")],["Valor da Venda",fV(pV(getCol(r._ga,"Valor da venda")||getIdx(r._ga,29)))],["Data Prev. Liq.",fD(parseAny(getCol(r._ga,"Data prevista de liquidacao","Data prevista de liquida\u00e7\u00e3o")||getIdx(r._ga,20)))],["N° Lógico",getCol(r._ga,"Numero logico","N\u00famero l\u00f3gico")||getIdx(r._ga,13)]].map(([l,v])=>(
                      <div key={l} style={{background:T.bg,padding:"8px 10px",borderRadius:6,border:`1px solid #2a2a3a`}}><div style={{fontSize:9,color:T.muted,marginBottom:2,letterSpacing:.4}}>{l.toUpperCase()}</div><div style={{fontWeight:600,color:T.white,fontSize:11,wordBreak:"break-word"}}>{v||"—"}</div></div>
                    ))}
                  </div>
                </div>}
                {/* BCK details */}
                {r._bck&&<div><div style={{fontSize:10,fontWeight:700,color:T.accent,letterSpacing:.8,marginBottom:6}}>◈ DETALHES CRÉDITO ({r._bck.dept||"BCK"}) — G.A</div>
                  <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:6}}>
                    {[["Data BCK",fD(r._bck.date)],["Observações GA",r._bck.obs||"—"],["Departamento",getCol(r._bck.row,"Departamento")],["Solicitante",getCol(r._bck.row,"Solicitante")],["Valor Ajuste GA",`${fV(r.bckValor||0)} ${r.bckValor!==null&&r.cval>0?(Math.abs(r.bckValor-r.cval)<0.05?"✅ OK":"⚠️ ≠ "+fV(r.cval)):""}`.trim()],["Status",getCol(r._bck.row,"Status do ajuste","Status da solicita\u00e7\u00e3o")],["Bandeira",getCol(r._bck.row,"Bandeira")],["Data Prev. Liq.",fD(parseAny(getCol(r._bck.row,"Data prevista de liquidacao")||getIdx(r._bck.row,20)))],["N° RO",getCol(r._bck.row,"Numero RO","N\u00famero RO")||getIdx(r._bck.row,23)],["Produto",getCol(r._bck.row,"Produto")]].map(([l,v])=>(
                      <div key={l} style={{background:T.surface,padding:"8px 10px",borderRadius:6,border:`1px solid #1a2a3a`}}><div style={{fontSize:9,color:T.muted,marginBottom:2,letterSpacing:.4}}>{l.toUpperCase()}</div><div style={{fontWeight:600,color:T.white,fontSize:11,wordBreak:"break-word"}}>{v||"—"}</div></div>
                    ))}
                  </div>
                </div>}
              </td></tr>)}
            </>))}
          </tbody>
        </table>
        {shown.length===0&&<div style={{textAlign:"center",padding:40,color:T.muted}}>Nenhum registro encontrado.</div>}
      </div>
    </div>
    <div style={{marginTop:12,display:"flex",gap:12,flexWrap:"wrap",alignItems:"center"}}>
      <span style={{fontSize:11,color:T.gray,fontWeight:700}}>Legenda:</span>
      {[["Duplicata","DUP"],["Sem CAN","SEM_CAN"],["CAN tardio (info)","SLA_CAN"],["SLA BCK","SLA_BCK"],["Agendado","SCHED"],["Valor divergente","VALOR_DIFF"],["No prazo","ONTIME"],["OK","OK"]].map(([l,t])=>(<span key={t} style={{display:"flex",alignItems:"center",gap:5,fontSize:11,color:T.gray}}><Badge type={t}/>{l}</span>))}
      <span style={{fontSize:10,color:T.muted,marginLeft:"auto"}}>Clique na linha para detalhar · D+2 considera feriados nacionais</span>
    </div>
  </div>);
};

const BADGES9066={
  SEM_984:{bg:"rgba(255,82,82,.15)",fg:"#ff5252",txt:"SEM 984"},
  COD_ERRADO:{bg:"rgba(255,82,82,.15)",fg:"#ff5252",txt:"CÓD ERRADO"},
  LOGICO_DIFF:{bg:"rgba(255,171,64,.15)",fg:"#ffab40",txt:"LÓGICO ≠"},
  BAND_ERRADA:{bg:"rgba(179,136,255,.15)",fg:"#b388ff",txt:"BAND ≠"},
  VALOR_DIFF:{bg:"rgba(255,171,64,.15)",fg:"#ffab40",txt:"VALOR ≠"},
  SLA_BCK:{bg:"rgba(255,82,82,.15)",fg:"#ff5252",txt:"⏰ SLA"},
  OK9:{bg:"rgba(0,230,118,.15)",fg:"#00e676",txt:"✓ OK"},
  ONTIME9:{bg:"rgba(0,230,118,.15)",fg:"#00e676",txt:"NO PRAZO"},
  LATE9:{bg:"rgba(255,82,82,.15)",fg:"#ff5252",txt:"ATRASADO"},
};
const Badge9=({type})=>{const s=BADGES9066[type]||{bg:T.hover,fg:T.gray,txt:"—"};return<span style={{display:"inline-block",padding:"3px 8px",borderRadius:20,fontSize:10,fontWeight:700,letterSpacing:.5,background:s.bg,color:s.fg,marginRight:3,whiteSpace:"nowrap"}}>{s.txt}</span>;};
const Check=({ok,yes,no})=>ok===true?<span style={{color:"#00e676",fontWeight:700}}>{yes||"✅"}</span>:ok===false?<span style={{color:"#ff5252",fontWeight:700}}>{no||"❌"}</span>:<span style={{color:T.muted}}>—</span>;

const View9066=({results})=>{
  const[search,setSearch]=useState("");const[onlyIssues,setOnlyIssues]=useState(false);const[expanded,setExpanded]=useState(null);const[activeFilter,setActiveFilter]=useState(null);
  const stats=useMemo(()=>({
    total:results.length,ok:results.filter(r=>r.ok).length,issues:results.filter(r=>!r.ok).length,
    sem984:results.filter(r=>r.issues.includes('SEM_984')).length,
    sla:results.filter(r=>r.slaOk===false).length,
    logicoDiff:results.filter(r=>r.issues.includes('LOGICO_DIFF')).length,
    valorDiff:results.filter(r=>r.issues.includes('VALOR_DIFF')).length,
    bandErrada:results.filter(r=>r.issues.includes('BAND_ERRADA')).length,
  }),[results]);
  const shown=useMemo(()=>{
    let r=results;
    if(activeFilter==="ok")r=r.filter(x=>x.ok);
    else if(activeFilter==="issues")r=r.filter(x=>!x.ok);
    else if(activeFilter==="sem984")r=r.filter(x=>x.issues.includes('SEM_984'));
    else if(activeFilter==="sla")r=r.filter(x=>x.slaOk===false);
    else if(activeFilter==="logico")r=r.filter(x=>x.issues.includes('LOGICO_DIFF'));
    else if(activeFilter==="valor")r=r.filter(x=>x.issues.includes('VALOR_DIFF'));
    if(onlyIssues)r=r.filter(x=>!x.ok);
    if(search.trim()){const s=search.toLowerCase();r=r.filter(x=>x.ec.includes(s)||x.logico.includes(s)||x.ref.toLowerCase().includes(s)||x.analista.toLowerCase().includes(s));}
    return r;
  },[results,search,onlyIssues,activeFilter]);
  const TH=({c})=><th style={{padding:"9px 10px",textAlign:"left",fontWeight:700,color:T.gray,fontSize:10,letterSpacing:.8,whiteSpace:"nowrap",borderBottom:`1px solid ${T.border}`}}>{c}</th>;
  const doExport=()=>{
    const out=shown.map(r=>({'EC':r.ec,'Lógico':r.logico,'Ref/Chamado':r.ref,'Valor Controle':fV(r.valor),'Valor GA':r.gaValor!==null?fV(r.gaValor):'—','Data Entrada':fD(r.dataEntrada),'Prazo D+4':fD(r.slaDl),'Data Ajuste':fD(r.ajusteDate),'SLA':r.slaOk===true?'NO PRAZO':r.slaOk===false?'ATRASADO':'—','Cód 984':r.cod984===true?'✓':r.cod984===false?'✗':'—','Lógico OK':r.logicoMatch===true?'✓':r.logicoMatch===false?'✗':'—','Bandeira VISA':r.bandVisa===true?'✓':r.bandVisa===false?'✗':'—','Valor OK':r.valorMatch===true?'✓':r.valorMatch===false?'✗':'—','Finalizado':r.finalizado,'Analista':r.analista,'Pendências':r.issues.join(', ')||'OK'}));
    const ws=XLSX.utils.json_to_sheet(out);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'9066');XLSX.writeFile(wb,`analise_9066_${TODAY}.xlsx`);
  };
  return(<div>
    <div style={{display:"grid",gridTemplateColumns:"repeat(8,1fr)",gap:10,marginBottom:16}}>
      {[[stats.total,"Total",T.accent,"📊",null],[stats.ok,"OK",T.success,"✅","ok"],[stats.issues,"Pendência",T.danger,"⚠️","issues"],[stats.sem984,"Sem 984",T.danger,"🚫","sem984"],[stats.sla,"SLA",T.danger,"⏰","sla"],[stats.logicoDiff,"Lógico ≠",T.warning,"🔢","logico"],[stats.valorDiff,"Valor ≠",T.warning,"💰","valor"],[stats.bandErrada,"Bandeira ≠",T.purple,"💳",null]].map(([v,l,clr,ic,fk])=>(
        <Stat key={l} label={l} value={v} color={clr} icon={ic} active={activeFilter===fk} onClick={()=>fk?setActiveFilter(af=>af===fk?null:fk):null}/>
      ))}
    </div>
    {activeFilter&&<div style={{textAlign:"center",fontSize:11,color:T.muted,marginBottom:8}}>
      Filtro ativo: <strong style={{color:T.accent}}>{activeFilter}</strong> · <span style={{cursor:"pointer",color:T.danger}} onClick={()=>setActiveFilter(null)}>✕ Limpar</span>
    </div>}
    <div style={{display:"flex",gap:10,alignItems:"center",marginBottom:12,flexWrap:"wrap"}}>
      <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar EC, lógico, referência, analista…" style={{flex:1,minWidth:200,padding:"10px 14px",background:T.card,border:`1px solid ${T.border}`,borderRadius:8,color:T.white,fontSize:13,outline:"none"}}/>
      <label style={{display:"flex",alignItems:"center",gap:8,fontSize:13,cursor:"pointer",color:T.gray}}><input type="checkbox" checked={onlyIssues} onChange={e=>setOnlyIssues(e.target.checked)} style={{accentColor:T.accent}}/>Apenas pendências</label>
      <span style={{fontSize:12,color:T.muted}}>{shown.length}/{results.length}</span>
      <button onClick={doExport} style={{padding:"10px 20px",background:T.accent,color:"#0a1628",border:"none",borderRadius:50,fontSize:12,fontWeight:700,cursor:"pointer"}}>⬇ Exportar XLSX</button>
    </div>
    <div style={{background:T.card,borderRadius:12,overflow:"hidden",boxShadow:"0 4px 16px rgba(0,0,0,.4)"}}>
      <div style={{overflowX:"auto",maxHeight:"52vh",overflowY:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
          <thead style={{position:"sticky",top:0,zIndex:2}}><tr style={{background:T.surface}}>
            {["EC","Lógico (Ctrl)","Lógico (GA)","Ref/Chamado","Valor Ctrl","Valor GA","Data Entrada","Analista","Prazo D+4","Data Ajuste","SLA","Cód 984","Lógico ✓","Bandeira","Valor ✓","Situação"].map(col=><TH key={col} c={col}/>)}
          </tr></thead>
          <tbody>
            {shown.map((r,i)=>(<>
              <tr key={`r${i}`} onClick={()=>setExpanded(expanded===i?null:i)} style={{background:!r.ok?"hsl(0,62.8%,8%)":i%2===0?T.card:T.hover,borderBottom:`1px solid ${T.border}`,cursor:"pointer"}}>
                <td style={{padding:"8px 10px",fontFamily:"monospace",fontSize:10,color:T.gray}}>{r.ec}</td>
                <td style={{padding:"8px 10px",fontWeight:700,color:T.accent}}>{r.logico}</td>
                <td style={{padding:"8px 10px",color:r.logicoMatch===false?T.danger:r.logicoMatch?T.success:T.muted,fontFamily:"monospace"}}>{r.gaLogico||"—"}</td>
                <td style={{padding:"8px 10px",color:T.white}}>{r.ref||"—"}</td>
                <td style={{padding:"8px 10px",fontWeight:700,color:T.white}}>{fV(r.valor)}</td>
                <td style={{padding:"8px 10px",color:r.valorMatch===false?T.danger:r.valorMatch?T.success:T.muted}}>{r.gaValor!==null?fV(r.gaValor):"—"}</td>
                <td style={{padding:"8px 10px",color:T.white}}>{fD(r.dataEntrada)}</td>
                <td style={{padding:"8px 10px",color:T.gray}}>{r.analista||"—"}</td>
                <td style={{padding:"8px 10px",color:T.muted}}>{fD(r.slaDl)}</td>
                <td style={{padding:"8px 10px",color:r.slaOk===false?T.danger:r.slaOk?T.success:T.muted}}>{fD(r.ajusteDate)}</td>
                <td style={{padding:"8px 10px"}}><Badge9 type={r.slaOk===true?"ONTIME9":r.slaOk===false?"LATE9":"—"}/></td>
                <td style={{padding:"8px 10px"}}><Check ok={r.cod984}/></td>
                <td style={{padding:"8px 10px"}}><Check ok={r.logicoMatch}/></td>
                <td style={{padding:"8px 10px",color:r.bandVisa===false?T.danger:r.bandVisa?T.success:T.muted,fontSize:10}}>{r.gaBandeira||"—"}</td>
                <td style={{padding:"8px 10px"}}><Check ok={r.valorMatch}/></td>
                <td style={{padding:"8px 10px"}}>{r.ok?<Badge9 type="OK9"/>:r.issues.map(t=><Badge9 key={t} type={t}/>)}</td>
              </tr>
              {expanded===i&&(<tr key={`e${i}`} style={{background:T.bg}}><td colSpan={16} style={{padding:"14px 18px"}}>
                <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8,marginBottom:10}}>
                  {[["Ref/Chamado",r.ref||"—"],["Finalizado",r.finalizado||"—"],["Data Finalização",fD(r.dataFin)],["Analista",r.analista||"—"],["Cód GA",r.gaCod||"—"],["Status GA",r.gaStatus||"—"],["Observação GA",r.gaObs||"—"],["Prazo SLA D+4",fD(r.slaDl)]].map(([l,v])=>(
                    <div key={l} style={{background:T.card,padding:"9px 12px",borderRadius:8,border:`1px solid ${T.border}`}}>
                      <div style={{fontSize:9,color:T.muted,marginBottom:3,letterSpacing:.5}}>{l.toUpperCase()}</div>
                      <div style={{fontWeight:600,color:T.white,fontSize:12,wordBreak:"break-word"}}>{v}</div>
                    </div>
                  ))}
                </div>
                <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:6}}>
                  {[["Cód 984",r.cod984===true?"✅ Correto":r.cod984===false?"❌ Errado":"—"],["Lógico Ctrl",r.logico],["Lógico GA",r.gaLogico||"—"],["Match Lógico",r.logicoMatch===true?"✅ Bate":r.logicoMatch===false?"❌ Diverge":"—"],["Bandeira",r.gaBandeira||"—"],["É VISA?",r.bandVisa===true?"✅ Sim":r.bandVisa===false?"❌ Não":"—"],["Valor Controle",fV(r.valor)],["Valor GA",r.gaValor!==null?fV(r.gaValor):"—"],["Match Valor",r.valorMatch===true?"✅ Bate":r.valorMatch===false?"❌ Diverge":"—"],["Dias p/ SLA",r.dataEntrada&&r.ajusteDate?`${Math.round((new Date(r.ajusteDate)-new Date(r.dataEntrada))/86400000)} dias corridos`:"—"]].map(([l,v])=>(
                    <div key={l} style={{background:"#0d1a2a",padding:"8px 10px",borderRadius:6,border:`1px solid #1a2a3a`}}>
                      <div style={{fontSize:9,color:T.muted,marginBottom:2,letterSpacing:.4}}>{l.toUpperCase()}</div>
                      <div style={{fontWeight:600,color:T.white,fontSize:11,wordBreak:"break-word"}}>{v}</div>
                    </div>
                  ))}
                </div>
              </td></tr>)}
            </>))}
          </tbody>
        </table>
        {shown.length===0&&<div style={{textAlign:"center",padding:40,color:T.muted}}>Nenhum registro encontrado.</div>}
      </div>
    </div>
    <div style={{marginTop:10,display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
      <span style={{fontSize:11,color:T.gray,fontWeight:700}}>Legenda:</span>
      {[["Sem 984","SEM_984"],["Cód Errado","COD_ERRADO"],["Lógico ≠","LOGICO_DIFF"],["Valor ≠","VALOR_DIFF"],["SLA","SLA_BCK"],["OK","OK9"]].map(([l,t])=>(
        <span key={t} style={{display:"flex",alignItems:"center",gap:4,fontSize:11,color:T.gray}}><Badge9 type={t}/>{l}</span>
      ))}
      <span style={{fontSize:10,color:T.muted,marginLeft:"auto"}}>SLA D+4 úteis a partir da Data Entrada · Feriados nacionais excluídos</span>
    </div>
  </div>);
};

const ViewComissao=({results})=>{
  const{rows,summary,analOnly=[],analTotalGeral=null}=results;
  const[search,setSearch]=useState("");const[tab,setTab]=useState("summary");const[expanded,setExpanded]=useState(null);
  const[onlyDiv,setOnlyDiv]=useState(false);
  const totalDifSistema=summary.reduce((s,r)=>s+r.totalDifSistema,0);
  const totalDifAnal=summary.reduce((s,r)=>s+r.totalDifAnal,0);
  const totalDiv=rows.filter(r=>!r.ok).length;
  const totalSemTaxaEc=rows.filter(r=>r.taxaOrigem==='PADRAO').length;
  // Gap entre o total do sistema (só as transações que casaram com o SIE filtrado) e o total real
  // da planilha do analista (arquivo inteiro) — normalmente vem de linhas que estão no analista
  // mas não apareceram no SIE que você subiu (período diferente, EC diferente, etc).
  const gapAnalista=analTotalGeral!==null?Math.round((analTotalGeral-totalDifSistema)*10000)/10000:null;

  const shown=useMemo(()=>{
    let r=rows;
    if(onlyDiv) r=r.filter(x=>!x.ok);
    if(search.trim()){const s=search.toLowerCase();r=r.filter(x=>x.produto.toLowerCase().includes(s)||x.auth.includes(s)||x.ec.includes(s));}
    return r;
  },[rows,search,onlyDiv]);

  const doExport=()=>{
    const out=rows.map(r=>({'EC':r.ec,'Autorização':r.auth,'Produto':r.produto,'Parcelas':r.parcelas||0,'Data':r.dtTrans,'Valor Transação':r.vt,'VCB':r.vcb,'Taxa Sistema%':r.taxaCorreta,'Origem Taxa':r.taxaOrigem==='EC'?'Tarifa do EC':r.taxaOrigem==='PADRAO'?'Tabela padrão':'—','Comissão Sistema':r.comissaoCorreta,'Diferença Sistema':r.difSistema,'Taxa Analista%':r.taxaAnal,'Comissão Analista':r.comissaoAnal,'Diferença Analista':r.difAnal,'Taxa Match':r.taxaMatch?'✓':'✗','Pendência':r.issues.join(', ')||'OK'}));
    const ws=XLSX.utils.json_to_sheet(out);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'ComissaoMinima');XLSX.writeFile(wb,`comissao_minima_${TODAY}.xlsx`);
  };

  const TH=({c})=><th style={{padding:"9px 10px",textAlign:"left",fontWeight:700,color:T.gray,fontSize:10,letterSpacing:.8,whiteSpace:"nowrap",borderBottom:`1px solid ${T.border}`}}>{c}</th>;

  return(<div>
    {/* Summary cards */}
    <div style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:12,marginBottom:20}}>
      {[[rows.length,"Transações (Flag=Sim)",T.accent,"📊",false],[summary.length,"Produtos únicos",T.success,"📦",false],[totalDiv,"Divergências",T.danger,"⚠️",false],[totalSemTaxaEc,"Sem tarifa do EC (padrão)",T.purple,"❔",false],[totalDifSistema,"Total Estorno (Sistema)",T.warning,"💰",true]].map(([v,l,clr,ic,isMoney])=>(
        <Stat key={l} label={l} value={isMoney?fV(v):v} color={clr} icon={ic}/>
      ))}
    </div>

    {/* Reconciliação com o total real do analista */}
    {analTotalGeral!==null&&Math.abs(gapAnalista)>=0.01&&(
      <div style={{marginBottom:16,padding:"12px 16px",background:"rgba(179,136,255,.08)",border:`1px solid ${T.purple}`,borderRadius:10,fontSize:12}}>
        <strong style={{color:T.purple}}>⚠ Total do sistema ({fV(totalDifSistema)}) ≠ total da planilha do analista ({fV(analTotalGeral)})</strong>
        <span style={{color:T.gray,marginLeft:8}}>Diferença de {fV(Math.abs(gapAnalista))}. </span>
        <span style={{color:T.gray}}>{analOnly.length>0?`${analOnly.length} transação(ões) da planilha do analista não têm par no SIE que você subiu — veja a aba "Só no Analista".`:"Confira se o período/EC dos dois arquivos é o mesmo."}</span>
      </div>
    )}

    {/* Tabs */}
    <div style={{display:"flex",gap:2,marginBottom:16,borderBottom:`1px solid ${T.border}`}}>
      {[["summary","📊 Resumo por Produto"],["detail","📋 Detalhe por Transação"],["diff","⚠️ Divergências"],...(analOnly.length>0?[["analonly",`🔎 Só no Analista (${analOnly.length})`]]:[])].map(([id,label])=>(
        <button key={id} onClick={()=>setTab(id)} style={{padding:"8px 16px",background:"transparent",border:"none",cursor:"pointer",fontSize:12,fontWeight:tab===id?700:400,color:tab===id?T.accent:T.gray,borderBottom:tab===id?`2px solid ${T.accent}`:"2px solid transparent",marginBottom:-1}}>
          {label}
        </button>
      ))}
      <div style={{flex:1}}/>
      <button onClick={doExport} style={{padding:"8px 18px",background:T.accent,color:"#0a1628",border:"none",borderRadius:50,fontSize:11,fontWeight:700,cursor:"pointer",marginBottom:4}}>⬇ Exportar</button>
    </div>

    {/* RESUMO TAB */}
    {tab==="summary"&&(
      <div>
        <div style={{background:T.card,borderRadius:12,overflow:"hidden",boxShadow:"0 4px 16px rgba(0,0,0,.4)"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr style={{background:T.surface}}>
              {["Produto Cielo","Qtd","Total VCB","Comissão Sistema","Estorno Sistema","Comissão Analista","Estorno Analista","Δ Analista vs Sistema","Status"].map(h=><TH key={h} c={h}/>)}
            </tr></thead>
            <tbody>
              {summary.map((s,i)=>{
                const delta=Math.abs(s.totalDifSistema-s.totalDifAnal);
                const ok=delta<0.01;
                return(<tr key={i} style={{background:!ok?"hsl(0,62.8%,8%)":i%2===0?T.card:T.hover,borderBottom:`1px solid ${T.border}`}}>
                  <td style={{padding:"9px 12px",fontWeight:600,color:T.white}}>{s.produto}</td>
                  <td style={{padding:"9px 12px",color:T.gray}}>{s.qtd}</td>
                  <td style={{padding:"9px 12px",color:T.white}}>{fV(s.totalVCB)}</td>
                  <td style={{padding:"9px 12px",color:T.accent}}>{fV(s.totalCSistema)}</td>
                  <td style={{padding:"9px 12px",fontWeight:700,color:T.warning}}>{fV(s.totalDifSistema)}</td>
                  <td style={{padding:"9px 12px",color:T.gray}}>{fV(s.totalCAnal)}</td>
                  <td style={{padding:"9px 12px",color:T.gray}}>{fV(s.totalDifAnal)}</td>
                  <td style={{padding:"9px 12px",color:ok?T.success:T.danger,fontWeight:700}}>{ok?"✅ OK":`⚠️ ${fV(delta)}`}</td>
                  <td style={{padding:"9px 12px"}}>{s.divergencias>0?<span style={{background:"rgba(255,82,82,.15)",color:"#ff5252",padding:"2px 8px",borderRadius:12,fontSize:10,fontWeight:700}}>{s.divergencias} diverg.</span>:<span style={{color:T.success,fontSize:11}}>✓ OK</span>}</td>
                </tr>);
              })}
              <tr style={{background:T.surface,borderTop:`2px solid ${T.border}`}}>
                <td colSpan={2} style={{padding:"10px 12px",fontWeight:700,color:T.white}}>TOTAL GERAL</td>
                <td style={{padding:"10px 12px",fontWeight:700,color:T.white}}>{fV(summary.reduce((s,r)=>s+r.totalVCB,0))}</td>
                <td style={{padding:"10px 12px",fontWeight:700,color:T.accent}}>{fV(summary.reduce((s,r)=>s+r.totalCSistema,0))}</td>
                <td style={{padding:"10px 12px",fontWeight:900,color:T.warning,fontSize:14}}>{fV(totalDifSistema)}</td>
                <td style={{padding:"10px 12px",fontWeight:700,color:T.gray}}>{fV(summary.reduce((s,r)=>s+r.totalCAnal,0))}</td>
                <td style={{padding:"10px 12px",fontWeight:700,color:T.gray}}>{fV(totalDifAnal)}</td>
                <td style={{padding:"10px 12px",fontWeight:700,color:Math.abs(totalDifSistema-totalDifAnal)<0.01?T.success:T.danger}}>{Math.abs(totalDifSistema-totalDifAnal)<0.01?"✅ IGUAL":`⚠️ ${fV(Math.abs(totalDifSistema-totalDifAnal))}`}</td>
                <td/>
              </tr>
            </tbody>
          </table>
        </div>
        <div style={{marginTop:10,padding:"12px 16px",background:T.card,borderRadius:10,border:`1px solid ${T.border}`,fontSize:12,color:T.gray}}>
          <strong style={{color:T.white}}>Como interpretar:</strong> "Estorno" = VCB - Comissão Correta. Positivo = cliente foi cobrado a mais (MDR mínimo aplicado incorretamente). É o valor a ser creditado de volta ao cliente via ajuste C005/D005.
          A taxa correta é buscada primeiro na tabela de taxas por EC (3º upload); quando o EC não tem tarifa cadastrada pra aquele produto/parcela, o sistema usa a tabela padrão da Cielo e sinaliza a transação como "SEM TAXA EC" pra revisão manual.
        </div>
      </div>
    )}

    {/* DETALHE TAB */}
    {/* SÓ NO ANALISTA TAB — transações da planilha do analista sem par no SIE filtrado */}
    {tab==="analonly"&&(
      <div>
        <div style={{marginBottom:10,padding:"10px 14px",background:T.card,borderRadius:8,border:`1px solid ${T.border}`,fontSize:12,color:T.gray}}>
          Essas autorizações aparecem na planilha do analista mas não foram encontradas entre as transações com Flag MDR Mínimo = Sim do SIE que você subiu agora. Confira se são de outro período/extração — não são divergência de cálculo, é ausência de dado pra comparar.
        </div>
        <div style={{background:T.card,borderRadius:12,overflow:"hidden",boxShadow:"0 4px 16px rgba(0,0,0,.4)"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr style={{background:T.surface}}>
              {["EC","Autorização","Produto","Data","Flag MDR Mínimo","Diferença (Analista)"].map(h=><TH key={h} c={h}/>)}
            </tr></thead>
            <tbody>
              {analOnly.map((r,i)=>(
                <tr key={r.auth+i} style={{background:i%2===0?T.card:T.hover,borderBottom:`1px solid ${T.border}`}}>
                  <td style={{padding:"9px 12px",color:T.gray}}>{r.ec||"—"}</td>
                  <td style={{padding:"9px 12px",fontFamily:"monospace",color:T.accent}}>{r.auth}</td>
                  <td style={{padding:"9px 12px",color:T.white}}>{r.produto||"—"}</td>
                  <td style={{padding:"9px 12px",color:T.gray}}>{String(r.dtTrans||"—")}</td>
                  <td style={{padding:"9px 12px",color:T.gray}}>{r.flag||"—"}</td>
                  <td style={{padding:"9px 12px",fontWeight:700,color:T.purple}}>{fV(r.difAnal)}</td>
                </tr>
              ))}
              <tr style={{background:T.surface,borderTop:`2px solid ${T.border}`}}>
                <td colSpan={5} style={{padding:"10px 12px",fontWeight:700,color:T.white}}>TOTAL SÓ NO ANALISTA</td>
                <td style={{padding:"10px 12px",fontWeight:900,color:T.purple}}>{fV(analOnly.reduce((s,r)=>s+r.difAnal,0))}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    )}

    {(tab==="detail"||tab==="diff")&&(
      <div>
        <div style={{display:"flex",gap:10,alignItems:"center",marginBottom:12}}>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar produto, autorização, EC…" style={{flex:1,padding:"10px 14px",background:T.card,border:`1px solid ${T.border}`,borderRadius:8,color:T.white,fontSize:13,outline:"none"}}/>
          <label style={{display:"flex",alignItems:"center",gap:8,fontSize:13,cursor:"pointer",color:T.gray}}>
            <input type="checkbox" checked={onlyDiv} onChange={e=>setOnlyDiv(e.target.checked)} style={{accentColor:T.accent}}/>Apenas divergências
          </label>
          <span style={{fontSize:12,color:T.muted}}>{shown.filter(r=>tab==="diff"?!r.ok:true).length}/{rows.length}</span>
        </div>
        <div style={{background:T.card,borderRadius:12,overflow:"hidden",boxShadow:"0 4px 16px rgba(0,0,0,.4)"}}>
          <div style={{overflowX:"auto",maxHeight:"50vh",overflowY:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
              <thead style={{position:"sticky",top:0,zIndex:2}}><tr style={{background:T.surface}}>
                {["Autorização","Produto","Parcelas","Data","Valor Trans.","VCB","% Sistema","Origem","Comiss. Sistema","Est. Sistema","% Analista","Comiss. Analista","Est. Analista","Taxa OK?","Situação"].map(h=><TH key={h} c={h}/>)}
              </tr></thead>
              <tbody>
                {shown.filter(r=>tab==="diff"?!r.ok:true).map((r,i)=>(
                  <>
                    <tr key={`r${i}`} onClick={()=>setExpanded(expanded===i?null:i)} style={{background:!r.ok?"hsl(0,62.8%,8%)":i%2===0?T.card:T.hover,borderBottom:`1px solid ${T.border}`,cursor:"pointer"}}>
                      <td style={{padding:"8px 10px",fontFamily:"monospace",color:T.accent,fontSize:10}}>{r.auth}</td>
                      <td style={{padding:"8px 10px",color:T.white,fontSize:10}}>{r.produto}</td>
                      <td style={{padding:"8px 10px",color:T.gray,textAlign:"center"}}>{r.parcelas||"—"}</td>
                      <td style={{padding:"8px 10px",color:T.gray}}>{r.dtTrans||"—"}</td>
                      <td style={{padding:"8px 10px",fontWeight:600,color:T.white}}>{fV(r.vt)}</td>
                      <td style={{padding:"8px 10px",color:T.white}}>{fV(r.vcb)}</td>
                      <td style={{padding:"8px 10px",color:T.accent,fontWeight:700}}>{fP(r.taxaCorreta)}</td>
                      <td style={{padding:"8px 10px"}}>{r.taxaOrigem==='EC'?<span style={{background:"rgba(0,230,118,.15)",color:T.success,padding:"2px 7px",borderRadius:10,fontSize:9,fontWeight:700}}>EC</span>:r.taxaOrigem==='PADRAO'?<span style={{background:"rgba(179,136,255,.15)",color:T.purple,padding:"2px 7px",borderRadius:10,fontSize:9,fontWeight:700}}>PADRÃO</span>:<span style={{color:T.muted}}>—</span>}</td>
                      <td style={{padding:"8px 10px",color:T.accent}}>{r.comissaoCorreta!==null?fV(r.comissaoCorreta):"—"}</td>
                      <td style={{padding:"8px 10px",fontWeight:700,color:r.difSistema>0?T.warning:T.success}}>{r.difSistema!==null?fV(r.difSistema):"—"}</td>
                      <td style={{padding:"8px 10px",color:r.taxaMatch===false?T.danger:T.gray}}>{fP(r.taxaAnal)}</td>
                      <td style={{padding:"8px 10px",color:T.gray}}>{r.comissaoAnal!==null?fV(r.comissaoAnal):"—"}</td>
                      <td style={{padding:"8px 10px",color:r.difMatch===false?T.danger:T.gray}}>{r.difAnal!==null?fV(r.difAnal):"—"}</td>
                      <td style={{padding:"8px 10px"}}>{r.taxaMatch===true?<span style={{color:T.success}}>✅</span>:r.taxaMatch===false?<span style={{color:T.danger}}>❌</span>:<span style={{color:T.muted}}>—</span>}</td>
                      <td style={{padding:"8px 10px"}}>{r.ok?<span style={{color:T.success,fontSize:10,fontWeight:700}}>✓ OK</span>:r.issues.map(iss=><span key={iss} style={{background:"rgba(255,82,82,.15)",color:"#ff5252",padding:"2px 6px",borderRadius:10,fontSize:9,fontWeight:700,marginRight:2}}>{iss.replace(/_/g,' ')}</span>)}</td>
                    </tr>
                    {expanded===i&&(<tr key={`e${i}`} style={{background:T.bg}}><td colSpan={15} style={{padding:"12px 16px"}}>
                      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8,fontSize:11}}>
                        {[["EC",r.ec],["Número RO",r.ro||"—"],["Percentual Desconto Cobrado",`${r.pctDesc}%`],["Parcelas",r.parcelas||"À vista"],
                          ["Taxa Sistema",r.taxaCorreta!==null?fP(r.taxaCorreta):"Não encontrada"],["Origem da Taxa",r.taxaOrigem==='EC'?"✅ Tarifa específica do EC":r.taxaOrigem==='PADRAO'?"❔ Tabela padrão (EC não cadastrado)":"—"],["Comissão Sistema",r.comissaoCorreta!==null?fV(r.comissaoCorreta):"—"],["Estorno Sistema",r.difSistema!==null?fV(r.difSistema):"—"],
                          ["Taxa Analista",r.taxaAnal!==null?fP(r.taxaAnal):"—"],["Comissão Analista",r.comissaoAnal!==null?fV(r.comissaoAnal):"—"],["Estorno Analista",r.difAnal!==null?fV(r.difAnal):"—"],
                          ["Taxa correta?",r.taxaMatch===true?"✅ Sim":r.taxaMatch===false?"❌ Não":"—"],["Cálculo correto?",r.difMatch===true?"✅ Sim":r.difMatch===false?"❌ Não":"—"]
                        ].map(([l,v])=>(
                          <div key={l} style={{background:T.card,padding:"9px 12px",borderRadius:8,border:`1px solid ${T.border}`}}>
                            <div style={{fontSize:9,color:T.muted,marginBottom:3,letterSpacing:.5}}>{l.toUpperCase()}</div>
                            <div style={{fontWeight:600,color:T.white,fontSize:12}}>{v}</div>
                          </div>
                        ))}
                      </div>
                      {!r.taxaMatch&&r.taxaCorreta!==null&&r.taxaAnal!==null&&(
                        <div style={{marginTop:10,padding:"10px 14px",background:"rgba(255,82,82,.1)",borderRadius:8,border:"1px solid rgba(255,82,82,.3)",fontSize:12}}>
                          <strong style={{color:"#ff5252"}}>⚠️ Taxa incorreta:</strong>
                          <span style={{color:T.gray,marginLeft:8}}>Analista usou <strong style={{color:T.white}}>{fP(r.taxaAnal)}</strong> mas a taxa correta é <strong style={{color:T.accent}}>{fP(r.taxaCorreta)}</strong> para <strong style={{color:T.white}}>{r.produto}</strong>{r.parcelas>1?` em ${r.parcelas}x`:''}.</span>
                          <span style={{color:T.warning,marginLeft:8}}>Diferença no estorno: {fV(Math.abs(r.difSistema-r.difAnal))}</span>
                        </div>
                      )}
                    </td></tr>)}
                  </>
                ))}
              </tbody>
            </table>
            {shown.filter(r=>tab==="diff"?!r.ok:true).length===0&&<div style={{textAlign:"center",padding:40,color:T.muted}}>{tab==="diff"?"Nenhuma divergência encontrada ✅":"Nenhum registro."}</div>}
          </div>
        </div>
      </div>
    )}
  </div>);
};

const View7922=({results})=>{
  const{procedentes,porMotivo,ajustesDup}=results;
  const[tab,setTab]=useState("procedente");
  const[search,setSearch]=useState("");
  const[onlyIssues,setOnlyIssues]=useState(false);
  const[expanded,setExpanded]=useState(null);
  const[openMotivo,setOpenMotivo]=useState(null);
  const[activeFilter,setActiveFilter]=useState(null);

  const stats=useMemo(()=>({
    totalProc:procedentes.length,
    ok:procedentes.filter(r=>r.ok).length,
    pend:procedentes.filter(r=>!r.ok).length,
    semAjuste:procedentes.filter(r=>r.issues.includes("SEM_AJUSTE_D297")).length,
    bandeiraErrada:procedentes.filter(r=>r.issues.includes("BANDEIRA_ERRADA")).length,
    foraD2:procedentes.filter(r=>r.issues.includes("FORA_D2")).length,
    totalImproc:porMotivo.reduce((s,m)=>s+m.itens.length,0),
  }),[procedentes,porMotivo]);

  // Cards do topo são clicáveis: cada um leva pra aba certa (Procedente/Improcedente) já filtrada
  // pro recorte que ele representa — clicar de novo no mesmo card volta a mostrar tudo.
  const goToStat=key=>{
    if(key==="totalImproc"){setTab("improcedente");setActiveFilter(null);return;}
    setTab("procedente");
    setActiveFilter(af=>af===key?null:key);
  };

  const shownProc=useMemo(()=>{
    let r=procedentes;
    if(activeFilter==="ok") r=r.filter(x=>x.ok);
    else if(activeFilter==="pend") r=r.filter(x=>!x.ok);
    else if(activeFilter==="semAjuste") r=r.filter(x=>x.issues.includes("SEM_AJUSTE_D297"));
    else if(activeFilter==="foraD2") r=r.filter(x=>x.issues.includes("FORA_D2"));
    if(onlyIssues) r=r.filter(x=>!x.ok);
    if(search.trim()){const s=search.toLowerCase();r=r.filter(x=>x.protocolo.includes(s)||x.ec.includes(s)||x.analista.toLowerCase().includes(s));}
    return r;
  },[procedentes,search,onlyIssues,activeFilter]);

  const TH=({c})=><th style={{padding:"9px 10px",textAlign:"left",fontWeight:700,color:T.gray,fontSize:10,letterSpacing:.8,whiteSpace:"nowrap",borderBottom:`1px solid ${T.border}`}}>{c}</th>;

  const exportProc=()=>{
    const out=procedentes.map(r=>({"Protocolo":r.protocolo,"EC":r.ec,"Analista":r.analista,"Data Finalizada":fD(r.dtFinalizada),"Bandeira":r.match?.bandeira||"—","D+2 (dias úteis)":r.bd??"—","Número RO":r.match?.ro||"—","Valor":r.match?.valor??"—","Observação (ajuste)":r.match?.obs||"—","Pendências":r.issues.join(", ")||"OK"}));
    const ws=XLSX.utils.json_to_sheet(out);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Procedente");XLSX.writeFile(wb,`evento7922_procedente_${TODAY}.xlsx`);
  };
  const exportImproc=()=>{
    const out=porMotivo.flatMap(m=>m.itens.map(i=>({"Motivo":m.motivo,"Protocolo":i.protocolo,"EC":i.ec,"Analista":i.analista,"Motivo (original)":i.motivoW,"Observação":i.obsX})));
    const ws=XLSX.utils.json_to_sheet(out);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Improcedente");XLSX.writeFile(wb,`evento7922_improcedente_${TODAY}.xlsx`);
  };

  return(<div>
    <div style={{display:"grid",gridTemplateColumns:"repeat(6,1fr)",gap:10,marginBottom:20}}>
      {[[stats.totalProc,"Procedentes",T.accent,"📊","totalProc"],[stats.ok,"OK (bandeira + D+2)",T.success,"✅","ok"],[stats.pend,"Com pendência",T.danger,"⚠️","pend"],[stats.semAjuste,"Sem ajuste D297",T.purple,"❔","semAjuste"],[stats.foraD2,"Fora do D+2",T.warning,"⏰","foraD2"],[stats.totalImproc,"Improcedentes",T.gray,"🚫","totalImproc"]].map(([v,l,clr,ic,key])=>(
        <Stat key={l} label={l} value={v} color={clr} icon={ic}
          active={key==="totalImproc"?tab==="improcedente":tab==="procedente"&&activeFilter===key}
          onClick={()=>goToStat(key)}
        />
      ))}
    </div>

    {ajustesDup.length>0&&(
      <div style={{marginBottom:16,padding:"12px 16px",background:"rgba(255,171,64,.08)",border:`1px solid ${T.warning}`,borderRadius:10,fontSize:12}}>
        <strong style={{color:T.warning}}>⚠ {ajustesDup.length} ajuste(s) D297 com Número RO + Valor + EC duplicados</strong>
        <span style={{color:T.gray,marginLeft:8}}>Risco de crédito em duplicidade — as linhas afetadas estão marcadas na aba Procedente.</span>
      </div>
    )}

    <div style={{display:"flex",gap:2,marginBottom:16,borderBottom:`1px solid ${T.border}`}}>
      {[["procedente",`✅ Procedente (${procedentes.length})`],["improcedente",`🚫 Improcedente (${stats.totalImproc})`]].map(([id,label])=>(
        <button key={id} onClick={()=>{setTab(id);setActiveFilter(null);}} style={{padding:"8px 16px",background:"transparent",border:"none",cursor:"pointer",fontSize:12,fontWeight:tab===id?700:400,color:tab===id?T.accent:T.gray,borderBottom:tab===id?`2px solid ${T.accent}`:"2px solid transparent",marginBottom:-1}}>
          {label}
        </button>
      ))}
      <div style={{flex:1}}/>
      <button onClick={tab==="procedente"?exportProc:exportImproc} style={{padding:"8px 18px",background:T.accent,color:"#0a1628",border:"none",borderRadius:50,fontSize:11,fontWeight:700,cursor:"pointer",marginBottom:4}}>⬇ Exportar</button>
    </div>

    {tab==="procedente"&&(
      <div>
        <div style={{marginBottom:10,padding:"10px 14px",background:T.card,borderRadius:8,border:`1px solid ${T.border}`,fontSize:12,color:T.gray}}>
          Pra cada protocolo Procedente, cruza pelo EC com a planilha de ajustes D297 e valida: crédito na bandeira <strong style={{color:T.white}}>VISA</strong>, ajuste criado em até <strong style={{color:T.white}}>D+2 dias úteis</strong> após a Data Finalizada (feriados desconsiderados), e ausência de duplicidade (Nº RO + Valor + EC).
        </div>
        <div style={{display:"flex",gap:10,alignItems:"center",marginBottom:12}}>
          <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Buscar protocolo, EC, analista…" style={{flex:1,padding:"10px 14px",background:T.card,border:`1px solid ${T.border}`,borderRadius:8,color:T.white,fontSize:13,outline:"none"}}/>
          <label style={{display:"flex",alignItems:"center",gap:8,fontSize:13,cursor:"pointer",color:T.gray}}>
            <input type="checkbox" checked={onlyIssues} onChange={e=>setOnlyIssues(e.target.checked)} style={{accentColor:T.accent}}/>Apenas pendências
          </label>
          <span style={{fontSize:12,color:T.muted}}>{shownProc.length}/{procedentes.length}</span>
        </div>
        <div style={{background:T.card,borderRadius:12,overflow:"hidden",boxShadow:"0 4px 16px rgba(0,0,0,.4)"}}>
          <div style={{overflowX:"auto",maxHeight:"55vh",overflowY:"auto"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
              <thead style={{position:"sticky",top:0,zIndex:2}}><tr style={{background:T.surface}}>
                {["Protocolo","EC","Analista","Data Finalizada","Bandeira","D+2 (dias úteis)","Nº RO","Valor","Observação","Situação"].map(h=><TH key={h} c={h}/>)}
              </tr></thead>
              <tbody>
                {shownProc.map((r,i)=>(
                  <>
                    <tr key={`r${i}`} onClick={()=>setExpanded(expanded===i?null:i)} style={{background:!r.ok?"hsl(0,62.8%,8%)":i%2===0?T.card:T.hover,borderBottom:`1px solid ${T.border}`,cursor:"pointer"}}>
                      <td style={{padding:"8px 10px",fontFamily:"monospace",color:T.accent}}>{r.protocolo}</td>
                      <td style={{padding:"8px 10px",color:T.white}}>{r.ec}</td>
                      <td style={{padding:"8px 10px",color:T.gray}}>{r.analista}</td>
                      <td style={{padding:"8px 10px",color:T.gray}}>{fD(r.dtFinalizada)}</td>
                      <td style={{padding:"8px 10px",color:r.bandeiraOk===false?T.danger:T.white}}>{r.match?.bandeira||"—"}</td>
                      <td style={{padding:"8px 10px",fontWeight:700,color:r.d2Ok===false?T.danger:r.d2Ok===true?T.success:T.muted}}>{r.bd!==null?`D+${r.bd}`:"—"}</td>
                      <td style={{padding:"8px 10px",color:T.gray}}>{r.match?.ro||"—"}</td>
                      <td style={{padding:"8px 10px",color:T.white}}>{r.match?fV(r.match.valor):"—"}</td>
                      <td style={{padding:"8px 10px",color:T.gray,maxWidth:220,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.match?.obs||"—"}</td>
                      <td style={{padding:"8px 10px"}}>{r.ok?<span style={{color:T.success,fontSize:10,fontWeight:700}}>✓ OK</span>:r.issues.map(iss=><span key={iss} style={{background:"rgba(255,82,82,.15)",color:"#ff5252",padding:"2px 6px",borderRadius:10,fontSize:9,fontWeight:700,marginRight:2}}>{iss.replace(/_/g," ")}</span>)}</td>
                    </tr>
                    {expanded===i&&(<tr key={`e${i}`} style={{background:T.bg}}><td colSpan={10} style={{padding:"12px 16px"}}>
                      <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:8,fontSize:11}}>
                        {[["Solicitação",r.match?.solicitacao||"—"],["Tipo de Ajuste",r.match?.tipoAjuste||"—"],["Código + Motivo",r.match?.codigoMotivo||"—"],["Data de Criação (ajuste)",fD(r.match?.dtCriacao)],
                          ["Bandeira correta?",r.bandeiraOk===true?"✅ Sim (VISA)":r.bandeiraOk===false?`❌ Não (${r.match?.bandeira})`:"—"],["Dentro do D+2?",r.d2Ok===true?"✅ Sim":r.d2Ok===false?"❌ Não":"—"],["Duplicidade?",r.match?.isDup?"⚠️ Sim":"Não"],["Outros ajustes p/ mesmo EC",r.outrosAjustes.length||"—"],
                        ].map(([l,v])=>(
                          <div key={l} style={{background:T.card,padding:"9px 12px",borderRadius:8,border:`1px solid ${T.border}`}}>
                            <div style={{fontSize:9,color:T.muted,marginBottom:3,letterSpacing:.5}}>{l.toUpperCase()}</div>
                            <div style={{fontWeight:600,color:T.white,fontSize:12}}>{v}</div>
                          </div>
                        ))}
                      </div>
                      {r.issues.includes("SEM_AJUSTE_D297")&&(
                        <div style={{marginTop:10,padding:"10px 14px",background:"rgba(179,136,255,.1)",borderRadius:8,border:`1px solid ${T.purple}`,fontSize:12}}>
                          <strong style={{color:T.purple}}>❔ Nenhum ajuste D297 encontrado</strong>
                          <span style={{color:T.gray,marginLeft:8}}>Não existe, na planilha de ajustes, nenhuma linha D297 pra o EC {r.ec} — confira se o crédito ainda não foi lançado ou se é de outro período/extração.</span>
                        </div>
                      )}
                    </td></tr>)}
                  </>
                ))}
              </tbody>
            </table>
            {shownProc.length===0&&<div style={{textAlign:"center",padding:40,color:T.muted}}>Nenhum registro encontrado.</div>}
          </div>
        </div>
      </div>
    )}

    {tab==="improcedente"&&(
      <div>
        <div style={{marginBottom:10,padding:"10px 14px",background:T.card,borderRadius:8,border:`1px solid ${T.border}`,fontSize:12,color:T.gray}}>
          Agrupado pelo motivo (coluna W do controle) — clique num motivo pra ver o protocolo, EC, analista e a observação (coluna X) de cada caso.
        </div>
        {porMotivo.map(m=>(
          <div key={m.motivo} style={{marginBottom:10,background:T.card,borderRadius:12,overflow:"hidden",boxShadow:"0 4px 16px rgba(0,0,0,.4)"}}>
            <div onClick={()=>setOpenMotivo(openMotivo===m.motivo?null:m.motivo)} style={{display:"flex",alignItems:"center",gap:10,padding:"14px 16px",cursor:"pointer",background:T.surface}}>
              <span style={{fontSize:13,fontWeight:700,color:T.white,flex:1}}>{m.motivo}</span>
              <span style={{background:m.itens.length?"rgba(255,82,82,.15)":T.hover,color:m.itens.length?"#ff5252":T.muted,padding:"3px 10px",borderRadius:12,fontSize:11,fontWeight:700}}>{m.itens.length}</span>
              <span style={{color:T.muted,fontSize:12}}>{openMotivo===m.motivo?"▲":"▼"}</span>
            </div>
            {openMotivo===m.motivo&&(
              <div style={{overflowX:"auto",maxHeight:"40vh",overflowY:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}>
                  <thead style={{position:"sticky",top:0,zIndex:2}}><tr style={{background:T.surface}}>
                    {["Protocolo","EC","Analista","Motivo (original)","Observação"].map(h=><TH key={h} c={h}/>)}
                  </tr></thead>
                  <tbody>
                    {m.itens.map((it,i)=>(
                      <tr key={i} style={{background:i%2===0?T.card:T.hover,borderBottom:`1px solid ${T.border}`}}>
                        <td style={{padding:"8px 10px",fontFamily:"monospace",color:T.accent}}>{it.protocolo}</td>
                        <td style={{padding:"8px 10px",color:T.white}}>{it.ec}</td>
                        <td style={{padding:"8px 10px",color:T.gray}}>{it.analista}</td>
                        <td style={{padding:"8px 10px",color:T.gray}}>{it.motivoW||"—"}</td>
                        <td style={{padding:"8px 10px",color:T.gray}}>{it.obsX||"—"}</td>
                      </tr>
                    ))}
                    {m.itens.length===0&&<tr><td colSpan={5} style={{textAlign:"center",padding:20,color:T.muted}}>Nenhum caso nesse motivo.</td></tr>}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        ))}
      </div>
    )}
  </div>);
};

// Conta registros de um slot carregado — array normal (loadFile) ou {type:"wb"|"rows",...}
// (loadWorkbookRaw, usado no Evento 7922, onde o slot guarda o workbook bruto e a aba certa só é
// escolhida/convertida na hora de Analisar). Pra "wb" mostra nº de abas em vez de linhas, já que
// contar linhas reais exigiria escanear todas as abas — e é só um indicador visual de "carregou".
const slotCount=d=>{
  if(Array.isArray(d)) return d.length;
  if(d&&d.type==="rows") return d.rows?.length||0;
  if(d&&d.type==="wb") return d.wb?.SheetNames?.length||0;
  return 0;
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
      {!mod.is5125&&!mod.is9066&&!mod.isComissao&&!mod.is7922&&<span style={{marginLeft:"auto",padding:"4px 14px",background:T.card,borderRadius:20,fontSize:11,color:T.gray,border:`1px solid ${T.border}`}}>Em desenvolvimento</span>}
    </div>
    <div style={{display:"grid",gridTemplateColumns:`repeat(${mod.slots.length},1fr) auto`,gap:12,alignItems:"end",marginBottom:24}}>
      {mod.slots.map(s=><UploadZone key={s.key} label={s.label} count={slotCount(slotData[s.key])} countLabel={s.allSheets?"abas carregadas":undefined} onFile={d=>setSlot(s.key,d)} enc={s.enc} allSheets={s.allSheets}/>)}
      <button onClick={run} disabled={!mod.canRun(slotData)} style={{padding:"0 28px",height:60,border:"none",borderRadius:50,fontSize:13,fontWeight:900,letterSpacing:.5,whiteSpace:"nowrap",background:mod.canRun(slotData)?T.accent:T.muted,color:mod.canRun(slotData)?"#000":T.card,cursor:mod.canRun(slotData)?"pointer":"not-allowed",boxShadow:mod.canRun(slotData)?`0 4px 20px ${T.accent}55`:"none"}}>▶ ANALISAR</button>
    </div>
    {moduleResults&&mod.is5125&&<View5125 results={moduleResults} onExport={export5125}/>}
    {moduleResults&&mod.is9066&&<View9066 results={moduleResults}/>}{moduleResults&&mod.isComissao&&<ViewComissao results={moduleResults}/>}{moduleResults&&mod.is7922&&<View7922 results={moduleResults}/>}
    {moduleResults&&!mod.is5125&&!mod.is9066&&!mod.isComissao&&!mod.is7922&&<GenericTable data={moduleResults} moduleId={moduleId}/>}
    {!moduleResults&&(<div style={{textAlign:"center",padding:"72px 24px",color:T.muted}}><div style={{fontSize:56,marginBottom:20}}>{mod.icon}</div>{mod.is5125?(<><p style={{fontSize:15,fontWeight:700,color:T.gray,margin:"0 0 12px"}}>Carregue as planilhas e clique em Analisar</p><p style={{fontSize:12,margin:0,lineHeight:2,color:T.muted}}>✔ Cancelamentos duplicados · ✔ SLA BCK: D+2 após CAN · ✔ CAN Tardio: informativo · ✔ Feriados 2025–2027</p></>):mod.is9066?(<><p style={{fontSize:15,fontWeight:700,color:T.gray,margin:"0 0 12px"}}>Carregue o Controle Sinistro e os Ajustes G.A</p><p style={{fontSize:12,margin:0,lineHeight:2,color:T.muted}}>✔ Código 984 · ✔ Número lógico · ✔ Bandeira VISA · ✔ Valor · ✔ SLA D+4 úteis · ✔ Feriados excluídos</p></>):mod.is7922?(<><p style={{fontSize:15,fontWeight:700,color:T.gray,margin:"0 0 12px"}}>Carregue a Planilha de Ajustes 7922 e o Controle dos Analistas</p><p style={{fontSize:12,margin:0,lineHeight:2,color:T.muted}}>✔ Filtra D297 · ✔ Bandeira VISA · ✔ SLA D+2 úteis · ✔ Duplicidade RO+Valor+EC · ✔ Motivos de improcedência</p></>):(<><p style={{fontSize:15,fontWeight:700,color:T.gray,margin:"0 0 8px"}}>Carregue o arquivo para visualizar os dados</p><p style={{fontSize:12,color:T.muted}}>Análise personalizada em breve</p></>)}</div>)}
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
