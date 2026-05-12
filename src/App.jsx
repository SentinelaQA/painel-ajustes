import { useState, useMemo, useRef, useEffect } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import { signInWithEmailAndPassword, signOut, onAuthStateChanged } from "firebase/auth";
import { auth } from "./firebase";

/* ─── Brazilian holidays 2025–2027 ──────────────────────────────────────────── */
const BR_HOL = new Set([
  "2025-01-01","2025-04-18","2025-04-21","2025-05-01","2025-06-19",
  "2025-09-07","2025-10-12","2025-11-02","2025-11-15","2025-12-25",
  "2026-01-01","2026-04-03","2026-04-21","2026-05-01","2026-06-04",
  "2026-09-07","2026-10-12","2026-11-02","2026-11-15","2026-12-25",
  "2027-01-01","2027-04-02","2027-04-21","2027-05-01","2027-05-27",
  "2027-09-07","2027-10-12","2027-11-02","2027-11-15","2027-12-25",
]);

/* ─── Shared utilities ───────────────────────────────────────────────────────── */
const parseD = (s) => {
  if (!s) return null;
  s = String(s).trim();
  if (/^\d{2}\/\d{2}\/\d{4}/.test(s)) {
    const [d, m, y] = s.split("/");
    return `${y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
  }
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
  return null;
};

const addBiz = (ds, n) => {
  if (!ds) return null;
  const dt = new Date(ds + "T12:00:00Z");
  let c = 0;
  while (c < n) {
    dt.setUTCDate(dt.getUTCDate() + 1);
    const k = dt.toISOString().slice(0, 10);
    const w = dt.getUTCDay();
    if (w !== 0 && w !== 6 && !BR_HOL.has(k)) c++;
  }
  return dt.toISOString().slice(0, 10);
};

const fD = (s) => {
  if (!s) return "—";
  const [y, m, d] = s.split("-");
  return `${d}/${m}/${y}`;
};
const pV = (s) => {
  if (!s) return 0;
  return parseFloat(String(s).replace(/\./g, "").replace(",", ".")) || 0;
};
const fV = (v) => v.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
const getCol = (row, ...keys) => {
  for (const k of keys) {
    const v = row[k];
    if (v !== undefined && v !== null && String(v).trim() !== "") return String(v).trim();
  }
  return "";
};
const TODAY = new Date().toISOString().slice(0, 10);

/* ─── File loader ────────────────────────────────────────────────────────────── */
const loadFile = (file, enc, cb) => {
  const ext = file.name.split(".").pop().toLowerCase();
  if (["xlsx", "xlsb", "xls"].includes(ext)) {
    const fr = new FileReader();
    fr.onload = (e) => {
      const wb = XLSX.read(e.target.result, { type: "array", raw: false });
      const ws = wb.Sheets[wb.SheetNames[0]];
      cb(XLSX.utils.sheet_to_json(ws, { defval: "", raw: false }));
    };
    fr.readAsArrayBuffer(file);
  } else {
    const fr = new FileReader();
    fr.onload = (e) =>
      cb(Papa.parse(e.target.result, { header: true, delimiter: ";", skipEmptyLines: true }).data);
    fr.readAsText(file, enc);
  }
};

/* ─── Evento 5125 analysis ───────────────────────────────────────────────────── */
function analyze5125(gaRows, ctrlRows) {
  const canByProt = {};
  const dupTrack = {};

  gaRows.forEach((r) => {
    const dep = getCol(r, "Departamento", "DEPARTAMENTO", "departamento").toUpperCase();
    if (dep !== "CAN") return;
    const prot = getCol(r, "Protocolo Cancelamento", "PROTOCOLO CANCELAMENTO");
    const ec = getCol(r, "EC", "ec");
    const auth = getCol(r, "Autorização", "Autorizacao", "AUTORIZAÇÃO", "AUTORIZACAO").toUpperCase();
    const sd = parseD(getCol(r, "Data da venda", "DATA DA VENDA", "Data da Venda"));
    const rawCd = getCol(r, "Data de criação", "Data de Criacao", "DATA DE CRIAÇÃO");
    const cd = parseD(rawCd.slice(0, 10));
    if (prot) canByProt[prot] = { ...r, _canDate: cd };
    if (auth && sd) {
      const k = `${ec}|${auth}|${sd}`;
      if (!dupTrack[k]) dupTrack[k] = [];
      dupTrack[k].push(prot || "?");
    }
  });

  const dupProts = new Set();
  Object.values(dupTrack).forEach((ps) => {
    if (ps.length > 1) ps.forEach((p) => p && dupProts.add(p));
  });

  return ctrlRows.map((c) => {
    const ref = getCol(c, "REFERÊNCIA", "REFERENCIA", "Referência", "Referencia");
    const ec = getCol(c, "ESTABELECIMENTO", "Estabelecimento");
    const auth = getCol(c, "AUTORIZAÇÃO", "AUTORIZACAO", "Autorização", "Autorizacao");
    const sd = parseD(getCol(c, "DATA DA VENDA", "Data da Venda"));
    const od = parseD(getCol(c, "DATA ABERTURA", "Data Abertura"));
    const bd = parseD(getCol(c, "DATA DO AJUSTE A CREDITO", "Data do Ajuste a Credito"));
    const valor = pV(getCol(c, "VALOR DA TRANSAÇÃO", "VALOR DA TRANSACAO", "Valor da Transação"));
    const cval = pV(getCol(c, "VALOR DO CANCELAMENTO", "Valor do Cancelamento"));

    const gaRec = canByProt[ref] || null;
    const canDate = gaRec?._canDate || null;
    const canDl = addBiz(od, 2);
    const bckDl = canDate ? addBiz(canDate, 2) : null;
    const canOk = canDate && canDl ? canDate <= canDl : null;
    const bckOk =
      bd && bckDl ? bd <= bckDl : !bd && bckDl && TODAY > bckDl ? false : null;
    const isDup = dupProts.has(ref);

    const issues = [];
    if (isDup) issues.push("DUP");
    if (!gaRec) issues.push("SEM_CAN");
    else if (canOk === false) issues.push("SLA_CAN");
    if (bckOk === false) issues.push("SLA_BCK");

    return {
      ref, ec, auth, sd, od, bd, valor, cval,
      analista: getCol(c, "ANALISTA", "Analista"),
      ajuste: getCol(c, "AJUSTE EFETUADO?", "Ajuste Efetuado?"),
      trans3943: getCol(c, "TRANSFERIDO PARA 3943", "Transferido para 3943"),
      canDate, canDl, canOk, bckDl, bckOk, isDup, issues,
      ok: issues.length === 0,
      _ga: gaRec, _c: c,
    };
  });
}

/* ─── Module registry ────────────────────────────────────────────────────────── */
const MODULES = [
  {
    id: "5125", name: "Evento 5125", group: "Eventos", icon: "⚡",
    desc: "Cancelamento sem saldo · Boleto / PIX",
    slots: [
      { key: "ctrl", label: "Planilha Controle (analistas)", enc: "UTF-8" },
      { key: "ga", label: "Relatório G.A — Gestor de Ajustes", enc: "ISO-8859-1" },
    ],
    canRun: (s) => s.ctrl?.length > 0 && s.ga?.length > 0,
    run: (s) => analyze5125(s.ga, s.ctrl),
    has5125: true,
  },
  { id: "7922", name: "Evento 7922", group: "Eventos", icon: "📋", desc: "Análise em desenvolvimento", slots: [{ key: "file", label: "Planilha do Evento", enc: "UTF-8" }], canRun: (s) => s.file?.length > 0, run: (s) => s.file },
  { id: "9066", name: "Evento 9066", group: "Eventos", icon: "📋", desc: "Análise em desenvolvimento", slots: [{ key: "file", label: "Planilha do Evento", enc: "UTF-8" }], canRun: (s) => s.file?.length > 0, run: (s) => s.file },
  { id: "reg-fin", name: "Regularizações Financeiras", group: "Caixas de E-mail", icon: "💼", desc: "Análise em desenvolvimento", slots: [{ key: "file", label: "Planilha de Regularizações", enc: "UTF-8" }], canRun: (s) => s.file?.length > 0, run: (s) => s.file },
  { id: "saldo-aud", name: "Saldo Auditoria", group: "Caixas de E-mail", icon: "🔍", desc: "Análise em desenvolvimento", slots: [{ key: "file", label: "Planilha de Auditoria", enc: "UTF-8" }], canRun: (s) => s.file?.length > 0, run: (s) => s.file },
  { id: "gest-alug", name: "Gestão Aluguel", group: "Caixas de E-mail", icon: "🏢", desc: "Análise em desenvolvimento", slots: [{ key: "file", label: "Planilha de Gestão", enc: "UTF-8" }], canRun: (s) => s.file?.length > 0, run: (s) => s.file },
  { id: "est-alug", name: "Estorno Gestão Aluguel", group: "Caixas de E-mail", icon: "↩️", desc: "Análise em desenvolvimento", slots: [{ key: "file", label: "Planilha de Estornos", enc: "UTF-8" }], canRun: (s) => s.file?.length > 0, run: (s) => s.file },
  { id: "incentivo", name: "Incentivo", group: "Caixas de E-mail", icon: "🎯", desc: "Análise em desenvolvimento", slots: [{ key: "file", label: "Planilha de Incentivos", enc: "UTF-8" }], canRun: (s) => s.file?.length > 0, run: (s) => s.file },
  { id: "desfaz", name: "Desfazimento", group: "Caixas de E-mail", icon: "🔄", desc: "Análise em desenvolvimento", slots: [{ key: "file", label: "Planilha de Desfazimento", enc: "UTF-8" }], canRun: (s) => s.file?.length > 0, run: (s) => s.file },
  { id: "std-aerea", name: "STD — Cia Aérea", group: "Caixas de E-mail", icon: "✈️", desc: "Análise em desenvolvimento", slots: [{ key: "file", label: "Planilha STD Aérea", enc: "UTF-8" }], canRun: (s) => s.file?.length > 0, run: (s) => s.file },
  { id: "std-corp", name: "STD-Corporate", group: "Caixas de E-mail", icon: "🏛️", desc: "Análise em desenvolvimento", slots: [{ key: "file", label: "Planilha STD Corporate", enc: "UTF-8" }], canRun: (s) => s.file?.length > 0, run: (s) => s.file },
];

const MODULE_BY_ID = Object.fromEntries(MODULES.map((m) => [m.id, m]));
const GROUPS = [...new Set(MODULES.map((m) => m.group))];

/* ─── Badge ──────────────────────────────────────────────────────────────────── */
const BADGE_MAP = {
  OK:      { bg: "#e8f5e9", fg: "#2e7d32", txt: "✓ OK" },
  ONTIME:  { bg: "#e8f5e9", fg: "#2e7d32", txt: "NO PRAZO" },
  LATE:    { bg: "#fce4e4", fg: "#c62828", txt: "ATRASADO" },
  DUP:     { bg: "#fff3e0", fg: "#e65100", txt: "DUPLICATA" },
  SEM_CAN: { bg: "#fce4e4", fg: "#c62828", txt: "SEM CAN" },
  SLA_CAN: { bg: "#fce4e4", fg: "#c62828", txt: "⏰ CAN" },
  SLA_BCK: { bg: "#fce4e4", fg: "#c62828", txt: "⏰ BCK" },
  PEND:    { bg: "#f5f5f5", fg: "#aaa",    txt: "—" },
};
const Badge = ({ type }) => {
  const s = BADGE_MAP[type] || BADGE_MAP.PEND;
  return (
    <span style={{ display: "inline-block", padding: "2px 7px", borderRadius: 4, fontSize: 10, fontWeight: 700, letterSpacing: 0.4, background: s.bg, color: s.fg, marginRight: 3, whiteSpace: "nowrap" }}>
      {s.txt}
    </span>
  );
};

/* ─── Login ──────────────────────────────────────────────────────────────────── */
const Login = () => {
  const [email, setEmail] = useState("");
  const [pass, setPass] = useState("");
  const [err, setErr] = useState("");
  const [busy, setBusy] = useState(false);

  const go = async () => {
    if (!email || !pass) { setErr("Preencha e-mail e senha."); return; }
    setBusy(true);
    setErr("");
    try {
      await signInWithEmailAndPassword(auth, email, pass);
      // onAuthStateChanged in App handles the redirect
    } catch (e) {
      const msgs = {
        "auth/user-not-found": "Usuário não encontrado.",
        "auth/wrong-password": "Senha incorreta.",
        "auth/invalid-email": "E-mail inválido.",
        "auth/too-many-requests": "Muitas tentativas. Aguarde e tente novamente.",
        "auth/invalid-credential": "Credenciais inválidas.",
      };
      setErr(msgs[e.code] || "Erro ao autenticar. Verifique suas credenciais.");
      setBusy(false);
    }
  };

  const inp = { width: "100%", padding: "10px 12px", fontSize: 13, borderRadius: 8, outline: "none", boxSizing: "border-box" };

  return (
    <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "linear-gradient(135deg,#0f172a 0%,#1e3a5f 55%,#1e40af 100%)", fontFamily: "'Segoe UI',system-ui,sans-serif" }}>
      <div style={{ position: "fixed", top: -100, right: -100, width: 400, height: 400, borderRadius: "50%", background: "rgba(255,255,255,.03)", pointerEvents: "none" }} />
      <div style={{ background: "white", borderRadius: 18, padding: "44px 40px", width: 360, boxShadow: "0 32px 80px rgba(0,0,0,.4)" }}>
        <div style={{ textAlign: "center", marginBottom: 32 }}>
          <div style={{ width: 54, height: 54, borderRadius: 14, background: "linear-gradient(135deg,#1e3a5f,#1e40af)", display: "inline-flex", alignItems: "center", justifyContent: "center", fontSize: 24, color: "white", marginBottom: 14, boxShadow: "0 8px 24px rgba(30,64,175,.4)" }}>◈</div>
          <h1 style={{ margin: 0, fontSize: 21, fontWeight: 800, color: "#0f172a", letterSpacing: -0.5 }}>Painel de Ajustes</h1>
          <p style={{ margin: "6px 0 0", fontSize: 11, color: "#94a3b8", letterSpacing: 0.3 }}>CONTROLE · AUDITORIA · ANÁLISE</p>
        </div>
        <div style={{ marginBottom: 14 }}>
          <label style={{ display: "block", fontSize: 10, fontWeight: 700, color: "#64748b", marginBottom: 5, letterSpacing: 0.7 }}>E-MAIL</label>
          <input type="email" value={email} onChange={(e) => setEmail(e.target.value)} onKeyDown={(e) => e.key === "Enter" && go()} placeholder="usuario@empresa.com.br"
            style={{ ...inp, border: "1.5px solid #e2e8f0" }} />
        </div>
        <div style={{ marginBottom: err ? 4 : 20 }}>
          <label style={{ display: "block", fontSize: 10, fontWeight: 700, color: "#64748b", marginBottom: 5, letterSpacing: 0.7 }}>SENHA</label>
          <input type="password" value={pass} onChange={(e) => { setPass(e.target.value); setErr(""); }} onKeyDown={(e) => e.key === "Enter" && go()} placeholder="••••••••"
            style={{ ...inp, border: `1.5px solid ${err ? "#e53935" : "#e2e8f0"}` }} />
        </div>
        {err && <p style={{ color: "#e53935", fontSize: 11, margin: "0 0 14px" }}>{err}</p>}
        <button onClick={go} disabled={busy} style={{ width: "100%", padding: "12px", background: busy ? "#94a3b8" : "linear-gradient(135deg,#1e3a5f,#1e40af)", color: "white", border: "none", borderRadius: 8, fontSize: 14, fontWeight: 700, cursor: busy ? "not-allowed" : "pointer", boxShadow: "0 4px 14px rgba(30,64,175,.35)" }}>
          {busy ? "Verificando…" : "Entrar"}
        </button>
        <p style={{ textAlign: "center", fontSize: 10, color: "#cbd5e1", margin: "18px 0 0" }}>Uso Interno · Acesso Restrito</p>
      </div>
    </div>
  );
};

/* ─── Sidebar ────────────────────────────────────────────────────────────────── */
const Sidebar = ({ activeId, onSelect }) => (
  <div style={{ width: 220, background: "#0f172a", color: "white", minHeight: "100%", padding: "16px 0", flexShrink: 0, overflowY: "auto" }}>
    {GROUPS.map((group) => (
      <div key={group} style={{ marginBottom: 8 }}>
        <div style={{ fontSize: 9, fontWeight: 700, color: "#475569", letterSpacing: 1.2, padding: "8px 16px 4px", textTransform: "uppercase" }}>{group}</div>
        {MODULES.filter((m) => m.group === group).map((m) => (
          <button key={m.id} onClick={() => onSelect(m.id)} style={{ display: "flex", alignItems: "center", gap: 10, width: "100%", padding: "9px 16px", border: "none", textAlign: "left", cursor: "pointer", background: activeId === m.id ? "rgba(30,64,175,.35)" : "transparent", color: activeId === m.id ? "#93c5fd" : "#94a3b8", fontSize: 12, fontWeight: activeId === m.id ? 700 : 400, borderLeft: activeId === m.id ? "3px solid #3b82f6" : "3px solid transparent", transition: "all .15s" }}>
            <span style={{ fontSize: 14 }}>{m.icon}</span>
            <span style={{ lineHeight: 1.3 }}>{m.name}</span>
          </button>
        ))}
      </div>
    ))}
  </div>
);

/* ─── Upload zone (self-contained ref) ───────────────────────────────────────── */
const UploadZone = ({ label, count, onFile, enc }) => {
  const ref = useRef();
  return (
    <div onClick={() => ref.current?.click()} style={{ background: "white", borderRadius: 10, border: `2px dashed ${count ? "#1e40af" : "#d1d5db"}`, padding: "14px 18px", cursor: "pointer", boxShadow: "0 1px 6px rgba(0,0,0,.06)" }}>
      <input ref={ref} type="file" accept=".csv,.xlsx,.xlsb,.xls" style={{ display: "none" }} onChange={(e) => e.target.files[0] && loadFile(e.target.files[0], enc, onFile)} />
      <div style={{ fontSize: 10, fontWeight: 700, color: "#1e3a5f", letterSpacing: 0.5, marginBottom: 3 }}>{label.toUpperCase()}</div>
      <div style={{ fontSize: 12, color: count ? "#16a34a" : "#9ca3af" }}>{count ? `✅ ${count} registros carregados` : "📎 CSV · XLSX · XLSB"}</div>
    </div>
  );
};

/* ─── Stat card ──────────────────────────────────────────────────────────────── */
const Stat = ({ label, value, color }) => (
  <div style={{ background: "white", borderRadius: 10, padding: "14px 10px", textAlign: "center", boxShadow: "0 1px 6px rgba(0,0,0,.06)", borderTop: `3px solid ${color}` }}>
    <div style={{ fontSize: 24, fontWeight: 800, color, lineHeight: 1 }}>{value}</div>
    <div style={{ fontSize: 10, color: "#6b7280", marginTop: 5, lineHeight: 1.3 }}>{label}</div>
  </div>
);

/* ─── Generic table ──────────────────────────────────────────────────────────── */
const GenericTable = ({ data, moduleId }) => {
  const [search, setSearch] = useState("");
  const cols = data.length > 0 ? Object.keys(data[0]).filter((k) => k !== "") : [];
  const rows = useMemo(() => {
    if (!search.trim()) return data;
    const s = search.toLowerCase();
    return data.filter((r) => Object.values(r).some((v) => String(v).toLowerCase().includes(s)));
  }, [data, search]);

  const doExport = () => {
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Dados");
    XLSX.writeFile(wb, `export_${moduleId}_${TODAY}.xlsx`);
  };

  return (
    <div>
      <div style={{ display: "flex", gap: 10, alignItems: "center", marginBottom: 12, flexWrap: "wrap" }}>
        <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Buscar em todos os campos…" style={{ flex: 1, minWidth: 200, padding: "7px 12px", border: "1.5px solid #e2e8f0", borderRadius: 6, fontSize: 12, outline: "none" }} />
        <span style={{ fontSize: 12, color: "#6b7280", whiteSpace: "nowrap" }}>{rows.length} / {data.length} registros</span>
        <button onClick={doExport} style={{ padding: "7px 16px", background: "#1e3a5f", color: "white", border: "none", borderRadius: 6, fontSize: 12, fontWeight: 600, cursor: "pointer" }}>⬇ Exportar XLSX</button>
      </div>
      <div style={{ background: "white", borderRadius: 12, boxShadow: "0 2px 8px rgba(0,0,0,.07)", overflow: "hidden" }}>
        <div style={{ overflowX: "auto", maxHeight: "55vh", overflowY: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
            <thead style={{ position: "sticky", top: 0, zIndex: 2 }}>
              <tr style={{ background: "#f1f5f9", borderBottom: "2px solid #e2e8f0" }}>
                {cols.map((h) => (<th key={h} style={{ padding: "9px 10px", textAlign: "left", fontWeight: 700, color: "#374151", fontSize: 10, letterSpacing: 0.4, whiteSpace: "nowrap" }}>{h}</th>))}
              </tr>
            </thead>
            <tbody>
              {rows.map((row, i) => (
                <tr key={i} style={{ background: i % 2 === 0 ? "white" : "#f9fafb", borderBottom: "1px solid #f0f0f2" }}>
                  {cols.map((k) => (<td key={k} style={{ padding: "7px 10px", color: "#374151", whiteSpace: "nowrap", maxWidth: 200, overflow: "hidden", textOverflow: "ellipsis" }}>{String(row[k] || "")}</td>))}
                </tr>
              ))}
            </tbody>
          </table>
          {rows.length === 0 && <div style={{ textAlign: "center", padding: 32, color: "#9ca3af", fontSize: 13 }}>Nenhum registro encontrado.</div>}
        </div>
      </div>
    </div>
  );
};

/* ─── Evento 5125 view ───────────────────────────────────────────────────────── */
const View5125 = ({ results, onExport }) => {
  const [search, setSearch] = useState("");
  const [onlyIssues, setOnlyIssues] = useState(false);
  const [expanded, setExpanded] = useState(null);

  const stats = useMemo(() => ({
    total: results.length,
    ok: results.filter((r) => r.ok).length,
    issues: results.filter((r) => !r.ok).length,
    dup: results.filter((r) => r.isDup).length,
    slaCan: results.filter((r) => r.canOk === false).length,
    slaBck: results.filter((r) => r.bckOk === false).length,
    semCan: results.filter((r) => r.issues.includes("SEM_CAN")).length,
  }), [results]);

  const shown = useMemo(() => {
    let r = results;
    if (onlyIssues) r = r.filter((x) => !x.ok);
    if (search.trim()) {
      const s = search.toLowerCase();
      r = r.filter((x) => x.ref.includes(s) || x.ec.includes(s) || x.auth.toLowerCase().includes(s) || x.analista.toLowerCase().includes(s));
    }
    return r;
  }, [results, search, onlyIssues]);

  const TH = ({ c }) => <th style={{ padding: "9px 10px", textAlign: "left", fontWeight: 700, color: "#374151", fontSize: 10, letterSpacing: 0.4, whiteSpace: "nowrap" }}>{c}</th>;

  return (
    <div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(7,1fr)", gap: 10, marginBottom: 16 }}>
        {[["Total", stats.total, "#1e3a5f"], ["✅ OK", stats.ok, "#16a34a"], ["⚠️ Pendência", stats.issues, "#dc2626"], ["🔁 Duplicata", stats.dup, "#d97706"], ["⏰ SLA CAN", stats.slaCan, "#c62828"], ["⏰ SLA BCK", stats.slaBck, "#c62828"], ["❌ Sem CAN", stats.semCan, "#7c3aed"]].map(([l, v, c]) => (
          <Stat key={l} label={l} value={v} color={c} />
        ))}
      </div>

      <div style={{ display: "flex", gap: 10, alignItems: "center", marginBottom: 10, flexWrap: "wrap" }}>
        <input value={search} onChange={(e) => setSearch(e.target.value)} placeholder="Buscar por referência, EC, autorização, analista…" style={{ flex: 1, minWidth: 200, padding: "7px 12px", border: "1.5px solid #e2e8f0", borderRadius: 6, fontSize: 12, outline: "none" }} />
        <label style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, cursor: "pointer", whiteSpace: "nowrap" }}>
          <input type="checkbox" checked={onlyIssues} onChange={(e) => setOnlyIssues(e.target.checked)} style={{ accentColor: "#1e40af" }} />
          Apenas pendências
        </label>
        <span style={{ fontSize: 12, color: "#6b7280", whiteSpace: "nowrap" }}>{shown.length} / {results.length}</span>
        <button onClick={() => onExport(shown)} style={{ padding: "7px 16px", background: "#166534", color: "white", border: "none", borderRadius: 6, fontSize: 12, fontWeight: 600, cursor: "pointer" }}>⬇ Exportar XLSX</button>
      </div>

      <div style={{ background: "white", borderRadius: 12, boxShadow: "0 2px 8px rgba(0,0,0,.07)", overflow: "hidden" }}>
        <div style={{ overflowX: "auto", maxHeight: "52vh", overflowY: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
            <thead style={{ position: "sticky", top: 0, zIndex: 2 }}>
              <tr style={{ background: "#f1f5f9", borderBottom: "2px solid #e2e8f0" }}>
                {["Referência","EC","Autorização","Data Venda","Valor","Data Abertura","Analista","Data CAN","Prazo CAN","SLA CAN","Data BCK","Prazo BCK","SLA BCK","Situação"].map((c) => <TH key={c} c={c} />)}
              </tr>
            </thead>
            <tbody>
              {shown.map((r, i) => (
                <>
                  <tr key={`r${i}`} onClick={() => setExpanded(expanded === i ? null : i)} style={{ background: r.isDup ? "#fffbeb" : !r.ok ? "#fef2f2" : i % 2 === 0 ? "white" : "#f9fafb", borderBottom: "1px solid #f0f2f5", cursor: "pointer" }}>
                    <td style={{ padding: "7px 10px", fontWeight: 600, color: "#1e3a5f" }}>{r.ref || "—"}</td>
                    <td style={{ padding: "7px 10px", fontFamily: "monospace", fontSize: 10, color: "#555" }}>{r.ec}</td>
                    <td style={{ padding: "7px 10px", fontFamily: "monospace" }}>{r.auth || "—"}</td>
                    <td style={{ padding: "7px 10px" }}>{fD(r.sd)}</td>
                    <td style={{ padding: "7px 10px", fontWeight: 600 }}>{fV(r.valor)}</td>
                    <td style={{ padding: "7px 10px" }}>{fD(r.od)}</td>
                    <td style={{ padding: "7px 10px" }}>{r.analista || "—"}</td>
                    <td style={{ padding: "7px 10px", color: r.canOk === false ? "#c62828" : r.canOk ? "#16a34a" : "#9ca3af" }}>{fD(r.canDate)}</td>
                    <td style={{ padding: "7px 10px", color: "#6b7280" }}>{fD(r.canDl)}</td>
                    <td style={{ padding: "7px 10px" }}><Badge type={r.canOk === true ? "ONTIME" : r.canOk === false ? "LATE" : "PEND"} /></td>
                    <td style={{ padding: "7px 10px", color: r.bckOk === false ? "#c62828" : r.bckOk ? "#16a34a" : "#9ca3af" }}>{fD(r.bd)}</td>
                    <td style={{ padding: "7px 10px", color: "#6b7280" }}>{fD(r.bckDl)}</td>
                    <td style={{ padding: "7px 10px" }}><Badge type={r.bckOk === true ? "ONTIME" : r.bckOk === false ? "LATE" : "PEND"} /></td>
                    <td style={{ padding: "7px 10px" }}>{r.ok ? <Badge type="OK" /> : r.issues.map((t) => <Badge key={t} type={t} />)}</td>
                  </tr>
                  {expanded === i && (
                    <tr key={`e${i}`} style={{ background: "#f8faff" }}>
                      <td colSpan={14} style={{ padding: "12px 16px" }}>
                        <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 8, fontSize: 11 }}>
                          {[
                            ["Valor Cancelamento", fV(r.cval)],
                            ["Ajuste Efetuado", r.ajuste || "—"],
                            ["Transf. 3943", r.trans3943 || "—"],
                            ["CAN no G.A", r._ga ? "✅ Localizado" : "❌ Não localizado"],
                            ["Dias abert. → CAN", r.canDate && r.od ? `${Math.round((new Date(r.canDate) - new Date(r.od)) / 86400000)} dias corridos` : "—"],
                            ["Dias CAN → BCK", r.canDate && r.bd ? `${Math.round((new Date(r.bd) - new Date(r.canDate)) / 86400000)} dias corridos` : "—"],
                            ["Prazo CAN", fD(r.canDl)],
                            ["Prazo BCK", fD(r.bckDl)],
                          ].map(([l, v]) => (
                            <div key={l} style={{ background: "white", padding: "8px 10px", borderRadius: 6, border: "1px solid #e2e8f0" }}>
                              <div style={{ fontSize: 10, color: "#9ca3af", marginBottom: 2 }}>{l}</div>
                              <div style={{ fontWeight: 600, color: "#1f2937" }}>{v}</div>
                            </div>
                          ))}
                        </div>
                      </td>
                    </tr>
                  )}
                </>
              ))}
            </tbody>
          </table>
          {shown.length === 0 && <div style={{ textAlign: "center", padding: 32, color: "#9ca3af", fontSize: 13 }}>Nenhum registro encontrado.</div>}
        </div>
      </div>

      <div style={{ marginTop: 10, display: "flex", gap: 12, flexWrap: "wrap", alignItems: "center" }}>
        <span style={{ fontSize: 11, color: "#6b7280", fontWeight: 600 }}>Legenda:</span>
        {[["Duplicata","DUP"],["Sem CAN","SEM_CAN"],["SLA CAN","SLA_CAN"],["SLA BCK","SLA_BCK"],["No prazo","ONTIME"],["OK","OK"]].map(([l, t]) => (
          <span key={t} style={{ display: "flex", alignItems: "center", gap: 4, fontSize: 11, color: "#6b7280" }}><Badge type={t} />{l}</span>
        ))}
        <span style={{ fontSize: 10, color: "#9ca3af", marginLeft: "auto" }}>Clique na linha para detalhar · D+2 considera feriados nacionais</span>
      </div>
    </div>
  );
};

/* ─── Module content ─────────────────────────────────────────────────────────── */
const ModuleContent = ({ moduleId, files, setFiles, results, setResults }) => {
  const mod = MODULE_BY_ID[moduleId];
  const slotData = files[moduleId] || {};
  const moduleResults = results[moduleId] || null;

  const setSlot = (key, data) =>
    setFiles((f) => ({ ...f, [moduleId]: { ...f[moduleId], [key]: data } }));

  const runAnalysis = () => {
    if (!mod.canRun(slotData)) return;
    setResults((r) => ({ ...r, [moduleId]: mod.run(slotData) }));
  };

  const export5125 = (rows) => {
    const out = rows.map((r) => ({
      Referência: r.ref, EC: r.ec, Autorização: r.auth,
      "Data Venda": fD(r.sd), Valor: fV(r.valor), "Data Abertura": fD(r.od),
      Analista: r.analista, "Data CAN": fD(r.canDate), "Prazo CAN": fD(r.canDl),
      "SLA CAN": r.canOk === true ? "NO PRAZO" : r.canOk === false ? "ATRASADO" : "—",
      "Data BCK": fD(r.bd), "Prazo BCK": fD(r.bckDl),
      "SLA BCK": r.bckOk === true ? "NO PRAZO" : r.bckOk === false ? "ATRASADO" : "—",
      Pendências: r.issues.join(", ") || "OK",
    }));
    const ws = XLSX.utils.json_to_sheet(out);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Análise");
    XLSX.writeFile(wb, `analise_5125_${TODAY}.xlsx`);
  };

  return (
    <div style={{ flex: 1, padding: "20px", overflowY: "auto" }}>
      <div style={{ marginBottom: 20, paddingBottom: 14, borderBottom: "1px solid #e2e8f0" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{ fontSize: 22 }}>{mod.icon}</span>
          <div>
            <h2 style={{ margin: 0, fontSize: 18, fontWeight: 800, color: "#0f172a" }}>{mod.name}</h2>
            <p style={{ margin: 0, fontSize: 12, color: "#6b7280" }}>{mod.desc}</p>
          </div>
          {!mod.has5125 && (
            <span style={{ marginLeft: "auto", padding: "3px 10px", background: "#f1f5f9", borderRadius: 20, fontSize: 11, color: "#64748b", fontWeight: 600 }}>Em desenvolvimento</span>
          )}
        </div>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: `repeat(${mod.slots.length},1fr) auto`, gap: 12, alignItems: "end", marginBottom: 20 }}>
        {mod.slots.map((slot) => (
          <UploadZone key={slot.key} label={slot.label} count={slotData[slot.key]?.length || 0} onFile={(data) => setSlot(slot.key, data)} enc={slot.enc} />
        ))}
        <button onClick={runAnalysis} disabled={!mod.canRun(slotData)} style={{ padding: "0 20px", height: 48, whiteSpace: "nowrap", border: "none", borderRadius: 8, fontSize: 13, fontWeight: 700, background: mod.canRun(slotData) ? "linear-gradient(135deg,#1e3a5f,#1e40af)" : "#e2e8f0", color: mod.canRun(slotData) ? "white" : "#9ca3af", cursor: mod.canRun(slotData) ? "pointer" : "not-allowed", boxShadow: mod.canRun(slotData) ? "0 4px 12px rgba(30,64,175,.3)" : "none" }}>
          ▶ Analisar
        </button>
      </div>

      {moduleResults && mod.has5125 && <View5125 results={moduleResults} onExport={export5125} />}
      {moduleResults && !mod.has5125 && <GenericTable data={moduleResults} moduleId={moduleId} />}

      {!moduleResults && (
        <div style={{ textAlign: "center", padding: "56px 24px", color: "#9ca3af" }}>
          <div style={{ fontSize: 48, marginBottom: 16 }}>{mod.icon}</div>
          {mod.has5125 ? (
            <>
              <p style={{ fontSize: 14, fontWeight: 600, color: "#475569", margin: "0 0 8px" }}>Carregue as planilhas e clique em Analisar</p>
              <p style={{ fontSize: 12, margin: 0, lineHeight: 1.8, color: "#94a3b8" }}>
                ✔ Cancelamentos duplicados (EC + autorização + data da venda)<br />
                ✔ SLA CAN — Data Abertura + D+2 dias úteis<br />
                ✔ SLA BCK — Data CAN + D+2 dias úteis<br />
                ✔ Feriados nacionais 2025–2027 incluídos
              </p>
            </>
          ) : (
            <>
              <p style={{ fontSize: 14, fontWeight: 600, color: "#475569", margin: "0 0 8px" }}>Carregue o arquivo para visualizar os dados</p>
              <p style={{ fontSize: 12, margin: 0, color: "#94a3b8" }}>Análise personalizada disponível em breve.<br />Os dados serão exibidos em tabela com busca e exportação.</p>
            </>
          )}
        </div>
      )}
    </div>
  );
};

/* ─── App root ───────────────────────────────────────────────────────────────── */
export default function App() {
  const [user, setUser] = useState(null);       // Firebase user object
  const [loading, setLoading] = useState(true); // while Firebase checks auth state
  const [activeModule, setActiveModule] = useState("5125");
  const [files, setFiles] = useState({});
  const [results, setResults] = useState({});

  // Firebase persistent auth — runs once on mount
  useEffect(() => {
    const unsub = onAuthStateChanged(auth, (u) => {
      setUser(u);
      setLoading(false);
    });
    return unsub; // cleanup on unmount
  }, []);

  const handleLogout = async () => {
    await signOut(auth);
    setFiles({});
    setResults({});
  };

  // Splash while Firebase resolves
  if (loading) {
    return (
      <div style={{ minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#0f172a" }}>
        <div style={{ color: "white", fontSize: 28, fontWeight: 800, opacity: 0.6 }}>◈</div>
      </div>
    );
  }

  if (!user) return <Login />;

  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100vh", fontFamily: "'Segoe UI',system-ui,sans-serif", overflow: "hidden" }}>
      {/* Topbar */}
      <div style={{ background: "#0f172a", color: "white", height: 54, display: "flex", alignItems: "center", padding: "0 20px", boxShadow: "0 2px 10px rgba(0,0,0,.3)", flexShrink: 0, zIndex: 10 }}>
        <div style={{ fontSize: 16, fontWeight: 800, letterSpacing: -0.3, display: "flex", alignItems: "center", gap: 8 }}>
          <span style={{ fontSize: 20 }}>◈</span>
          <span>Painel de Ajustes</span>
        </div>
        <div style={{ margin: "0 14px", opacity: 0.25 }}>|</div>
        <div style={{ fontSize: 12, opacity: 0.6 }}>{MODULE_BY_ID[activeModule]?.name}</div>
        <div style={{ flex: 1 }} />
        <div style={{ fontSize: 11, opacity: 0.55, marginRight: 14 }}>{user.email}</div>
        <button onClick={handleLogout} style={{ background: "rgba(255,255,255,.1)", border: "none", color: "white", padding: "5px 12px", borderRadius: 6, fontSize: 11, cursor: "pointer" }}>Sair</button>
      </div>

      {/* Body */}
      <div style={{ display: "flex", flex: 1, overflow: "hidden" }}>
        <Sidebar activeId={activeModule} onSelect={setActiveModule} />
        <div style={{ flex: 1, overflow: "auto", background: "#f1f5f9" }}>
          <ModuleContent moduleId={activeModule} files={files} setFiles={setFiles} results={results} setResults={setResults} />
        </div>
      </div>
    </div>
  );
}
