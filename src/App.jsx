import { useState } from "react";
import * as XLSX from "xlsx";

const ISO = [
  { id: "old",  label: "Slechte isolatie",   sub: "Voor 1980",           w: 120 },
  { id: "avg",  label: "Gemiddelde isolatie", sub: "1980–2000",           w: 100 },
  { id: "good", label: "Goede isolatie",      sub: "Na 2000 / renovatie", w: 70  },
  { id: "new",  label: "Nieuwbouw",           sub: "Hoog E-peil",         w: 50  },
];

const IU = [
  { id: "MIDMSCB1BU09HRFN8ME", lbl: "Wandunit 2.7 kW",      kw: 2.7, pipeLiq: '1/4"', pipeGas: '3/8"', cable: "5×2,5 mm²" },
  { id: "MIDMSCB1BU12HRFN8ME", lbl: "Wandunit 3.5 kW",      kw: 3.5, pipeLiq: '1/4"', pipeGas: '3/8"', cable: "5×2,5 mm²" },
  { id: "MIDMSCB1CU18HRFN8ME", lbl: "Wandunit 5.2 kW Twin", kw: 5.2, pipeLiq: '1/4"', pipeGas: '1/2"', cable: "5×2,5 mm²" },
  { id: "MIDMSCB1DU24HRF8ME",  lbl: "Wandunit 7.0 kW Twin", kw: 7.0, pipeLiq: '3/8"', pipeGas: '5/8"', cable: "5×2,5 mm²" },
];

const OM = [
  { id: "MIDMOX10309HFNXME", lbl: "Monosplit 2.7 kW R32", kw: 2.7,  mz: 1, mn: 2.0, mx: 3.2, breaker: "10A traag", ouCable: "3×2,5 mm²" },
  { id: "MIDMOX10312HFNXME", lbl: "Monosplit 3.5 kW R32", kw: 3.5,  mz: 1, mn: 3.0, mx: 4.0, breaker: "10A traag", ouCable: "3×2,5 mm²" },
  { id: "MIDMOX30318HFNXME", lbl: "Monosplit 5.2 kW R32", kw: 5.2,  mz: 1, mn: 4.5, mx: 5.5, breaker: "16A traag", ouCable: "3×2,5 mm²" },
  { id: "MIDMOX40124HFN8ME", lbl: "Monosplit 7.0 kW R32", kw: 7.0,  mz: 1, mn: 6.0, mx: 8.0, breaker: "20A traag", ouCable: "3×2,5 mm²" },
];

const OMU = [
  { id: "MIDM20D18HFN8Q", lbl: "Multisplit 5.27 kW R32",        kw: 5.27,  mz: 2, mn: 3.5, mx: 5.8,  breaker: "16A traag", ouCable: "3×2,5 mm²" },
  { id: "MIDM4O27N8HRU",  lbl: "Multisplit CIRQ HP 8.2 kW R32", kw: 8.20,  mz: 4, mn: 5.0, mx: 9.5,  breaker: "20A traag", ouCable: "3×2,5 mm²" },
  { id: "MIDM4036FN8Q",   lbl: "Multisplit 10.55 kW R32",       kw: 10.55, mz: 4, mn: 6.0, mx: 11.5, breaker: "25A traag", ouCable: "3×4,0 mm²" },
];

const C = {
  accent: "#C5E22A", accentDark: "#9BB81A", accentTxt: "#1a1e0a",
  bg: "#0f1117", surface: "#1a1e26", surfaceHigh: "#222733",
  border: "#2e3340", borderHigh: "#3d4455",
  white: "#ffffff", textPrimary: "#f0f2f5", textSecondary: "#8b93a8", textMuted: "#555e72",
  green: "#1a2e1a", greenBrd: "#2d5a2d", greenTxt: "#7ed87e",
  red: "#2e1a1a",   redBrd: "#5a2d2d",   redTxt: "#f08080",
  amber: "#2e2a1a", amberBrd: "#EF9F27", amberTxt: "#f0c860",
  blueLt: "#1e2a0d",
};

function ckw(sqm, w) { return Math.round(sqm * w) / 1000; }

function sidx(sqm, w, overdim = 15) {
  const k = ckw(sqm, w) * (1 - overdim / 100);
  for (let i = 0; i < 4; i++) if (k <= IU[i].kw) return i;
  return 3;
}

function sou(grs, ouOverdim = 15) {
  const tk = grs.reduce((s, r) => s + IU[r.ii].kw, 0), n = grs.length;
  const ouF = 1 + ouOverdim / 100;
  if (n === 1) {
    const k = IU[grs[0].ii].kw;
    const minOUSingle = k * (1 - ouOverdim / 100);
    return OM.find(u => u.kw >= minOUSingle && k >= u.mn && k <= u.mx * ouF) || null;
  }
  const minOU = tk * (1 - ouOverdim / 100);
  return OMU.filter(u => n <= u.mz && u.kw >= minOU && tk >= u.mn && tk <= u.mx * ouF)
    .sort((a, b) => a.kw - b.kw)[0] || null;
}

function vg(grs, overdim = 15, ouOverdim = 15) {
  const n = grs.length;
  const tk = grs.reduce((s, r) => s + IU[r.ii].kw, 0);
  const ou = sou(grs, ouOverdim);
  const errs = [];
  if (n > 4) errs.push("Maximum 4 binnenunits per buitenunit.");
  else if (!ou) {
    if (n > 1) {
      const mk = OMU.filter(u => n <= u.mz).reduce((m, u) => Math.max(m, u.mx * (1 + ouOverdim / 100)), 0);
      errs.push(tk > mk
        ? `Totaalvermogen (${tk.toFixed(1)} kW) te hoog (max ${mk.toFixed(1)} kW). Verdeel over meerdere buitenunits.`
        : "Geen compatibele multisplit buitenunit gevonden.");
    } else errs.push("Geen compatibele monosplit buitenunit gevonden.");
  }
  return { ou, errs, tk };
}

let nextId = 1;
function makeRoom(name, w, od = 15) {
  const idx = sidx(20, w, od);
  return { id: nextId++, name, sqm: 20, sqmRaw: "20", ii: idx, di: idx, pipeDist: "", pipeDistRaw: "", condensePump: false };
}
function makeGroup() { return { mountType: "wall", fuseDistRaw: "", fuseDist: "" }; }

function Btn({ active, onClick, children, small }) {
  return (
    <button onClick={onClick} style={{
      padding: small ? "3px 10px" : "3px 12px", borderRadius: 6, cursor: "pointer", fontSize: 15,
      border: active ? `1.5px solid ${C.accent}` : `1px solid ${C.border}`,
      background: active ? C.blueLt : C.surfaceHigh,
      color: active ? C.accent : C.textSecondary, fontWeight: active ? 600 : 400,
    }}>{children}</button>
  );
}

function SecLabel({ children }) {
  return <div style={{ fontSize: 12, fontWeight: 700, color: C.accent, letterSpacing: "0.12em", textTransform: "uppercase", marginBottom: 8 }}>{children}</div>;
}

function SectionHead({ children }) {
  return <div style={{ fontSize: 13, fontWeight: 700, color: C.accent, textTransform: "uppercase", letterSpacing: "0.12em", marginBottom: 12 }}>{children}</div>;
}

function MatTable({ rows }) {
  return (
    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 15 }}>
      <thead>
        <tr>{["Omschrijving", "Spec", "Detail", "Aantal"].map((h, i) => (
          <th key={h} style={{ textAlign: i === 3 ? "right" : "left", padding: "4px 8px", fontSize: 12, fontWeight: 700, color: C.textMuted, borderBottom: `1px solid ${C.border}`, background: C.surfaceHigh, letterSpacing: "0.06em", textTransform: "uppercase" }}>{h}</th>
        ))}</tr>
      </thead>
      <tbody>
        {rows.map((row, i) => (
          <tr key={i} style={{ borderBottom: `1px solid ${C.border}` }}>
            <td style={{ padding: "6px 8px", color: C.textPrimary, fontWeight: 500 }}>{row[0]}</td>
            <td style={{ padding: "6px 8px", color: C.accent, fontFamily: "monospace", fontSize: 13 }}>{row[1]}</td>
            <td style={{ padding: "6px 8px", color: C.textMuted, fontSize: 13 }}>{row[2]}</td>
            <td style={{ padding: "6px 8px", color: C.textPrimary, textAlign: "right", fontWeight: 700 }}>{row[3]}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );
}

export default function App() {
  const [iso, setIso] = useState(ISO[1]);
  const [rooms, setRooms] = useState(() => [makeRoom("Ruimte 1", ISO[1].w)]);
  const [groups, setGroups] = useState(() => [[1]]);
  const [groupMeta, setGroupMeta] = useState(() => [makeGroup()]);
  const [dragSrc, setDragSrc] = useState(null);
  const [dragOv, setDragOv] = useState(null);
  const [overdim, setOverdim] = useState(15);
  const [overdimRaw, setOverdimRaw] = useState("15");
  const [overdimOU, setOverdimOU] = useState(15);
  const [overdimOURaw, setOverdimOURaw] = useState("15");

  function setIsoLive(newIso) {
    setIso(newIso);
    setRooms(prev => prev.map(r => { const ni = sidx(r.sqm, newIso.w, overdim); return { ...r, di: ni, ii: r.ii === r.di ? ni : r.ii }; }));
  }

  function addRoom() {
    const r = makeRoom(`Ruimte ${rooms.length + 1}`, iso.w, overdim);
    setRooms(prev => [...prev, r]);
    setGroups(prev => { const ng = [...prev]; ng[0] = [...ng[0], r.id]; return ng; });
  }

  function removeRoom(id) {
    const nr = rooms.filter(r => r.id !== id);
    setRooms(nr);
    const ids = new Set(nr.map(r => r.id));
    setGroups(prev => prev.map(g => g.filter(x => ids.has(x))).filter(g => g.length > 0));
  }

  function updateRoom(id, patch) {
    setRooms(prev => prev.map(r => {
      if (r.id !== id) return r;
      const next = { ...r, ...patch };
      if (patch.sqm !== undefined) { const ni = sidx(patch.sqm, iso.w, overdim); next.di = ni; if (r.ii === r.di) next.ii = ni; }
      return next;
    }));
  }

  function handleSqm(id, raw) {
    const num = parseInt(raw);
    setRooms(prev => prev.map(r => {
      if (r.id !== id) return r;
      if (Number.isFinite(num) && num > 0) { const ni = sidx(num, iso.w, overdim); return { ...r, sqmRaw: raw, sqm: num, di: ni, ii: r.ii === r.di ? ni : r.ii }; }
      return { ...r, sqmRaw: raw };
    }));
  }

  function handlePipe(id, raw) {
    const num = parseInt(raw);
    setRooms(prev => prev.map(r => r.id !== id ? r : { ...r, pipeDistRaw: raw, pipeDist: Number.isFinite(num) && num >= 0 ? num : r.pipeDist }));
  }

  function updateMeta(gi, patch) { setGroupMeta(prev => prev.map((g, i) => i === gi ? { ...g, ...patch } : g)); }

  function handleFuse(gi, raw) {
    const num = parseInt(raw);
    setGroupMeta(prev => prev.map((g, i) => i === gi ? { ...g, fuseDistRaw: raw, fuseDist: Number.isFinite(num) && num >= 0 ? num : g.fuseDist } : g));
  }

  function moveRoom(rid, fgi, tgi) {
    if (fgi === tgi) return;
    const ng = groups.map(g => [...g]);
    ng[fgi] = ng[fgi].filter(x => x !== rid);
    let tg = tgi;
    if (ng[fgi].length === 0 && ng.length > 1) { ng.splice(fgi, 1); if (tgi > fgi) tg = tgi - 1; }
    if (!ng[tg]) ng[tg] = [];
    ng[tg].push(rid);
    const cl = ng.filter(g => g.length > 0);
    setGroups(cl);
    setGroupMeta(prev => { const next = [...prev]; while (next.length < cl.length) next.push(makeGroup()); return next.slice(0, cl.length); });
  }

  const activeGroups = groups.filter(g => g.length > 0);
  const gmeta = groupMeta.length >= activeGroups.length ? groupMeta : [...groupMeta, ...Array(activeGroups.length - groupMeta.length).fill(null).map(makeGroup)];


  function exportToExcel() {
    const wb = XLSX.utils.book_new();
    const allRows = [["Buitenunit", "Omschrijving", "Spec", "Detail", "Aantal"]];

    activeGroups.forEach((grp, gi) => {
      const grs = grp.map(id => rooms.find(r => r.id === id)).filter(Boolean);
      if (!grs.length) return;
      const { ou } = vg(grs, overdim, overdimOU);
      const meta = gmeta[gi] || makeGroup();
      const fuseDistNum = Number(meta.fuseDist) || 0;
      const label = `Buitenunit ${gi + 1}`;

      allRows.push([label, "Buitenunit", ou ? ou.lbl : "—", ou ? ou.id : "—", "1 st"]);
      allRows.push([label, "Automaat zekeringkast", ou ? ou.breaker : "—", "", "1 st"]);
      allRows.push([label, "Voedingskabel zekeringkast → buitenunit", ou ? ou.ouCable : "—", "", fuseDistNum > 0 ? `${fuseDistNum} m` : "—"]);
      allRows.push([label, "Montage", meta.mountType === "wall" ? "Muurbeugel" : "Voetmodel", "", "1 st"]);

      grs.forEach(r => {
        const u = IU[r.ii];
        const dist = Number(r.pipeDist) || 0;
        const rolls = dist > 0 ? Math.ceil(dist / 20) : 1;
        const ruimte = `${label} — ${r.name}`;
        allRows.push([ruimte, "Binnenunit", u.lbl, u.id, "1 st"]);
        allRows.push([ruimte, `Koelleiding ${u.pipeLiq} + ${u.pipeGas}`, "Rol 20 m", dist > 0 ? `${dist} m → ${rolls} rol${rolls > 1 ? "len" : ""}` : "afstand?", `${rolls} rol${rolls > 1 ? "len" : ""}`]);
        allRows.push([ruimte, "Verbindingskabel buiten → binnen", u.cable, "", dist > 0 ? `${dist} m` : "—"]);
        if (r.condensePump) allRows.push([ruimte, "Condenspomp", "1 st", "", ""]);
      });
    });

    const ws = XLSX.utils.aoa_to_sheet(allRows);
    ws["!cols"] = [{ wch: 28 }, { wch: 42 }, { wch: 22 }, { wch: 28 }, { wch: 14 }];
    XLSX.utils.book_append_sheet(wb, ws, "Materiaallijst");
    XLSX.writeFile(wb, "Materiaallijst_Hivolta.xlsx");
  }

  const inp = (extra = {}) => ({ border: `1px solid ${C.border}`, borderRadius: 6, padding: "4px 8px", fontSize: 18, fontWeight: 500, color: C.textPrimary, background: C.surfaceHigh, outline: "none", textAlign: "right", ...extra });

  return (
    <div style={{ fontFamily: "system-ui, sans-serif", fontSize: 18, color: C.textPrimary, background: C.bg, minHeight: "100vh" }}>

      {/* Header */}
      <div style={{ background: C.surface, borderBottom: `1px solid ${C.border}`, padding: "13px 20px", display: "flex", alignItems: "center", gap: 10 }}>
        <div style={{ width: 30, height: 30, borderRadius: 7, background: C.accent, color: C.accentTxt, fontSize: 19, fontWeight: 800, display: "flex", alignItems: "center", justifyContent: "center" }}>M</div>
        <span style={{ fontSize: 18, fontWeight: 700, color: C.white, letterSpacing: "-0.01em" }}>Midea warmtepomp configurator</span>
        <div style={{ marginLeft: "auto", fontSize: 13, color: C.textMuted }}>by hivolta</div>
      </div>

      <div style={{ maxWidth: 1100, margin: "0 auto", padding: "24px 16px 64px" }}>

        {/* Isolatie */}
        <SectionHead>Isolatiegraad gebouw</SectionHead>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 8, marginBottom: 28 }}>
          {ISO.map(l => {
            const sel = iso.id === l.id;
            return (
              <button key={l.id} onClick={() => setIsoLive(l)} style={{ background: sel ? C.blueLt : C.surface, border: sel ? `1.5px solid ${C.accent}` : `1px solid ${C.border}`, borderRadius: 10, padding: "10px 12px", cursor: "pointer", textAlign: "left", width: "100%" }}>
                <div style={{ fontSize: 16, fontWeight: 600, color: sel ? C.accent : C.textPrimary }}>{l.label}</div>
                <div style={{ fontSize: 13, color: C.textSecondary, marginTop: 2 }}>{l.sub}</div>
                <span style={{ display: "inline-block", fontSize: 13, padding: "2px 6px", borderRadius: 999, marginTop: 5, background: sel ? C.accent : C.surfaceHigh, color: sel ? C.accentTxt : C.textSecondary }}>{l.w} W/m²</span>
              </button>
            );
          })}
        </div>

        {/* Overdimensionering */}
        <SectionHead>Overdimensionering</SectionHead>
        <div style={{ display: "flex", gap: 10, marginBottom: 28, flexWrap: "wrap" }}>
          {[
            { title: "BINNENUNIT", label: "Binnenunit mag groter zijn dan buitenunit", val: overdimRaw, setVal: setOverdimRaw, commit: v => { setOverdim(v); setRooms(prev => prev.map(r => { const ni = sidx(r.sqm, iso.w, v); return { ...r, di: ni, ii: r.ii === r.di ? ni : r.ii }; })); }, cur: overdim },
            { title: "BUITENUNIT", label: "Buitenunit mag groter zijn dan binnenunit", val: overdimOURaw, setVal: setOverdimOURaw, commit: v => setOverdimOU(v), cur: overdimOU },
          ].map(({ title, label, val, setVal, commit, cur }) => (
            <div key={title} style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 10, padding: "12px 14px", flex: 1, minWidth: 220 }}>
              <div style={{ fontSize: 12, fontWeight: 700, color: C.accent, letterSpacing: "0.1em", marginBottom: 8 }}>{title}</div>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 5 }}>
                <span style={{ fontSize: 15, color: C.textPrimary, flex: 1 }}>{label}</span>
                <input type="number" min={0} max={50} value={val}
                  onChange={e => { setVal(e.target.value); const v = parseInt(e.target.value); if (Number.isFinite(v) && v >= 0 && v <= 50) commit(v); }}
                  onBlur={e => { const v = parseInt(e.target.value); const s = Number.isFinite(v) && v >= 0 && v <= 50 ? v : cur; commit(s); setVal(String(s)); }}
                  style={{ ...inp({ width: 50, fontSize: 19 }) }} />
                <span style={{ fontSize: 15, color: C.textSecondary }}>%</span>
              </div>
              <div style={{ fontSize: 13, color: C.textMuted }}>Bij {cur}% en 4 kW: min. <strong style={{ color: C.textPrimary }}>{(4 * (1 - cur / 100)).toFixed(1)} kW</strong> buitenunit</div>
            </div>
          ))}
        </div>

        {/* Ruimtes */}
        <SectionHead>Ruimtes</SectionHead>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(230px, 1fr))", gap: 10, marginBottom: 10 }}>
          {rooms.map(r => {
            const gi = groups.findIndex(g => g.includes(r.id));
            const kw = ckw(r.sqm, iso.w);
            const u = IU[r.ii];
            const isDefault = r.ii === r.di;
            const dov = dragOv?.rid === r.id && dragSrc?.rid !== r.id;
            return (
              <div key={r.id} draggable
                onDragStart={() => setDragSrc({ rid: r.id, gi })} onDragEnd={() => setDragSrc(null)}
                onDragOver={e => { e.preventDefault(); setDragOv({ rid: r.id, gi }); }} onDragLeave={() => setDragOv(null)}
                onDrop={e => { e.preventDefault(); if (dragSrc && dragSrc.rid !== r.id) moveRoom(dragSrc.rid, dragSrc.gi, gi); setDragSrc(null); setDragOv(null); }}
                style={{ background: C.surface, border: dov ? `1.5px solid ${C.accent}` : `1px solid ${C.border}`, borderRadius: 12, padding: 13, opacity: dragSrc?.rid === r.id ? 0.4 : 1, cursor: "grab", userSelect: "none" }}>

                <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 11 }}>
                  <div style={{ display: "flex", flexDirection: "column", gap: 2.5 }}>
                    {[0,1,2].map(i => <span key={i} style={{ display: "block", width: 10, height: 1.5, background: C.borderHigh, borderRadius: 1 }} />)}
                  </div>
                  <input value={r.name} onChange={e => updateRoom(r.id, { name: e.target.value })}
                    style={{ flex: 1, border: "none", background: "transparent", fontSize: 18, fontWeight: 600, color: C.textPrimary, outline: "none", minWidth: 0 }} />
                  {rooms.length > 1 && <button onClick={() => removeRoom(r.id)} style={{ background: "none", border: "none", cursor: "pointer", color: C.textMuted, fontSize: 21, lineHeight: 1 }}>×</button>}
                </div>

                {/* m² */}
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
                  <span style={{ fontSize: 15, color: C.textSecondary, minWidth: 76, flexShrink: 0 }}>Oppervlakte</span>
                  <input type="number" min={1} max={500} value={r.sqmRaw}
                    onChange={e => handleSqm(r.id, e.target.value)}
                    onBlur={() => { const v = parseInt(r.sqmRaw); const s = Number.isFinite(v) && v > 0 ? v : r.sqm; updateRoom(r.id, { sqm: s, sqmRaw: String(s) }); }}
                    style={{ ...inp({ width: 60 }) }} />
                  <span style={{ fontSize: 15, color: C.textSecondary }}>m²</span>
                </div>

                <div style={{ fontSize: 13, color: C.textMuted, marginBottom: 10, paddingLeft: 84 }}>
                  {kw.toFixed(1)} kW nodig · met {overdim}% marge: {(kw * (1 - overdim / 100)).toFixed(1)} kW
                </div>

                {/* Afstand buitenunit */}
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 9 }}>
                  <span style={{ fontSize: 15, color: C.textSecondary, minWidth: 76, flexShrink: 0 }}>Afstand buiten</span>
                  <input type="number" min={0} value={r.pipeDistRaw}
                    onChange={e => handlePipe(r.id, e.target.value)}
                    onBlur={() => { const v = parseInt(r.pipeDistRaw); if (!Number.isFinite(v) || v < 0) updateRoom(r.id, { pipeDistRaw: r.pipeDist !== "" ? String(r.pipeDist) : "" }); }}
                    style={{ ...inp({ width: 60 }) }} />
                  <span style={{ fontSize: 15, color: C.textSecondary }}>m</span>
                </div>

                {/* Condenspomp */}
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 11 }}>
                  <span style={{ fontSize: 15, color: C.textSecondary, minWidth: 76, flexShrink: 0 }}>Condenspomp</span>
                  <div style={{ display: "flex", gap: 5 }}>
                    <Btn active={r.condensePump === true}  onClick={() => updateRoom(r.id, { condensePump: true  })}>Ja</Btn>
                    <Btn active={r.condensePump === false} onClick={() => updateRoom(r.id, { condensePump: false })}>Nee</Btn>
                  </div>
                </div>

                {/* IU suggestion */}
                <div style={{ background: C.surfaceHigh, borderRadius: 8, padding: "7px 9px", display: "flex", alignItems: "center", gap: 8 }}>
                  <div style={{ width: 32, height: 32, borderRadius: 5, background: C.accent, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, fontSize: 12, fontWeight: 800, color: C.accentTxt, textAlign: "center", lineHeight: 1.3 }}>
                    {u.kw}<br />kW
                  </div>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontSize: 15, fontWeight: 600, color: C.textPrimary, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{u.lbl}</div>
                    <div style={{ fontSize: 12, color: C.textMuted, marginTop: 1, fontFamily: "monospace" }}>{u.id}</div>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 3 }}>
                    {[["↑", () => setRooms(p => p.map(x => x.id === r.id && x.ii < IU.length-1 ? {...x, ii: x.ii+1} : x)), r.ii >= IU.length-1],
                      ["↓", () => setRooms(p => p.map(x => x.id === r.id && x.ii > 0 ? {...x, ii: x.ii-1} : x)), r.ii <= 0]
                    ].map(([lbl, fn, dis], i) => (
                      <button key={i} onClick={fn} disabled={dis} style={{ width: 22, height: 22, borderRadius: 4, border: `1px solid ${C.border}`, background: C.surface, cursor: dis ? "not-allowed" : "pointer", fontSize: 15, color: C.textSecondary, display: "flex", alignItems: "center", justifyContent: "center", opacity: dis ? 0.25 : 1 }}>{lbl}</button>
                    ))}
                    <button onClick={() => setRooms(p => p.map(x => x.id === r.id ? {...x, ii: x.di} : x))}
                      style={{ visibility: isDefault ? "hidden" : "visible", fontSize: 12, padding: "2px 5px", borderRadius: 4, border: `1px solid ${C.border}`, background: C.surface, cursor: "pointer", color: C.textSecondary }}>reset</button>
                  </div>
                </div>

                {activeGroups.length > 1 && gi >= 0 && (
                  <div style={{ marginTop: 8, textAlign: "right" }}>
                    <span style={{ fontSize: 13, padding: "2px 7px", borderRadius: 999, background: C.blueLt, color: C.accent, border: `1px solid ${C.accentDark}` }}>Buitenunit {gi + 1}</span>
                  </div>
                )}
              </div>
            );
          })}

          <button onClick={addRoom} style={{ background: "transparent", border: `1px dashed ${C.borderHigh}`, borderRadius: 12, padding: 14, cursor: "pointer", color: C.textMuted, fontSize: 16, minHeight: 80, display: "flex", alignItems: "center", justifyContent: "center", gap: 6 }}>
            <span style={{ fontSize: 26, lineHeight: 1 }}>+</span> Ruimte toevoegen
          </button>
        </div>

        {rooms.length > 1 && (
          <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 28, flexWrap: "wrap" }}>
            <span style={{ fontSize: 15, color: C.textMuted }}>Sleep ruimtes samen voor 1 buitenunit.</span>
            {activeGroups.length < rooms.length && (
              <button onClick={() => { setGroups(g => [...g, []]); setGroupMeta(g => [...g, makeGroup()]); }}
                style={{ fontSize: 15, padding: "4px 10px", borderRadius: 6, border: `1px dashed ${C.borderHigh}`, background: "transparent", cursor: "pointer", color: C.textSecondary }}>
                + Extra buitenunit
              </button>
            )}
          </div>
        )}

        {/* Resultaat */}
        <div style={{ borderTop: `1px solid ${C.border}`, paddingTop: 24 }}>
          <SectionHead>Resultaat installatie</SectionHead>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(270px, 1fr))", gap: 14 }}>
            {activeGroups.map((grp, gi) => {
              const grs = grp.map(id => rooms.find(r => r.id === id)).filter(Boolean);
              if (!grs.length) return null;
              const { ou, errs, tk } = vg(grs, overdim, overdimOU);
              const hasErr = errs.length > 0;
              const pct = ou ? Math.min(100, Math.round((tk / ou.kw) * 100)) : 0;
              const meta = gmeta[gi] || makeGroup();

              return (
                <div key={gi} style={{ background: C.surface, border: `1px solid ${hasErr ? C.redBrd : C.border}`, borderRadius: 12, overflow: "hidden" }}>
                  <div style={{ height: 3, background: hasErr ? "#f08080" : C.accent }} />
                  <div style={{ padding: 14 }}>
                    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
                      <span style={{ fontSize: 19, fontWeight: 700 }}>Buitenunit {gi + 1}</span>
                      <span style={{ fontSize: 13, padding: "2px 8px", borderRadius: 999, background: hasErr ? C.red : C.green, border: `1px solid ${hasErr ? C.redBrd : C.greenBrd}`, color: hasErr ? C.redTxt : C.greenTxt, fontWeight: 700 }}>{hasErr ? "Fout" : "Compatibel"}</span>
                    </div>

                    {ou && !hasErr ? (
                      <>
                        <div style={{ background: C.surfaceHigh, borderRadius: 8, padding: "9px 11px", marginBottom: 10 }}>
                          <div style={{ fontSize: 18, fontWeight: 700, color: C.textPrimary, marginBottom: 2 }}>{ou.lbl}</div>
                          <div style={{ fontSize: 13, color: C.textMuted, fontFamily: "monospace" }}>{ou.id}</div>
                        </div>
                        <div style={{ marginBottom: 12 }}>
                          <div style={{ height: 4, background: C.border, borderRadius: 2, overflow: "hidden", marginBottom: 3 }}>
                            <div style={{ height: "100%", width: `${pct}%`, borderRadius: 2, background: pct > 100 ? "#f08080" : C.accent, transition: "width .3s" }} />
                          </div>
                          <div style={{ display: "flex", justifyContent: "space-between", fontSize: 13, color: C.textMuted }}>
                            <span>Vermogensbezetting</span><span>{tk.toFixed(1)} / {ou.kw} kW</span>
                          </div>
                        </div>
                      </>
                    ) : (
                      <div style={{ background: C.red, border: `1px solid ${C.redBrd}`, borderRadius: 8, padding: "9px 11px", marginBottom: 12 }}>
                        {errs.map((e, i) => <div key={i} style={{ display: "flex", gap: 6, fontSize: 15, color: C.redTxt, ...(i > 0 ? { marginTop: 4 } : {}) }}><span>✕</span><span>{e}</span></div>)}
                      </div>
                    )}

                    <div style={{ background: C.surfaceHigh, border: `1px solid ${C.border}`, borderRadius: 8, padding: "10px 11px", marginBottom: 12 }}>
                      <SecLabel>Installatie buitenunit</SecLabel>
                      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                        <span style={{ fontSize: 15, color: C.textSecondary, minWidth: 76, flexShrink: 0 }}>Montage</span>
                        <div style={{ display: "flex", gap: 6 }}>
                          <Btn active={meta.mountType === "wall"} onClick={() => updateMeta(gi, { mountType: "wall" })}>Muurbeugel</Btn>
                          <Btn active={meta.mountType === "floor"} onClick={() => updateMeta(gi, { mountType: "floor" })}>Voetmodel</Btn>
                        </div>
                      </div>
                      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        <span style={{ fontSize: 15, color: C.textSecondary, minWidth: 76, flexShrink: 0 }}>Zekeringkast</span>
                        <input type="number" min={0} max={999} value={meta.fuseDistRaw}
                          onChange={e => handleFuse(gi, e.target.value)}
                          onBlur={() => { const v = parseInt(meta.fuseDistRaw); if (!Number.isFinite(v) || v < 0) updateMeta(gi, { fuseDistRaw: meta.fuseDist !== "" ? String(meta.fuseDist) : "" }); }}
                          placeholder="afstand" style={{ ...inp({ width: 65 }) }} />
                        <span style={{ fontSize: 15, color: C.textSecondary }}>m</span>
                      </div>
                    </div>

                    <div style={{ borderTop: `1px solid ${C.border}`, paddingTop: 10 }}>
                      <SecLabel>Binnenunits ({grs.length})</SecLabel>
                      {grs.map(r => {
                        const u = IU[r.ii];
                        return (
                          <div key={r.id} style={{ marginBottom: 8, padding: "8px 10px", background: C.surfaceHigh, borderRadius: 8, border: `1px solid ${C.border}` }}>
                            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 5 }}>
                              <div style={{ width: 26, height: 26, borderRadius: 4, background: C.accent, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, fontSize: 11, fontWeight: 800, color: C.accentTxt, textAlign: "center", lineHeight: 1.3 }}>{u.kw}<br />kW</div>
                              <div style={{ flex: 1, minWidth: 0 }}>
                                <div style={{ fontSize: 13, fontWeight: 600, color: C.textPrimary, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{u.lbl}</div>
                                <div style={{ fontSize: 12, color: C.textMuted, fontFamily: "monospace" }}>{u.id}</div>
                              </div>
                              <div style={{ fontSize: 13, fontWeight: 700, color: C.accent }}>{u.kw} kW</div>
                            </div>
                            <div style={{ fontSize: 13, color: C.textSecondary, marginBottom: 5 }}>{r.name} · {r.sqm} m²</div>
                            <div style={{ display: "flex", gap: 5, flexWrap: "wrap" }}>
                              {r.pipeDist !== "" && <span style={{ fontSize: 13, padding: "2px 7px", borderRadius: 999, background: C.surface, border: `1px solid ${C.border}`, color: C.textPrimary }}>Leiding: {r.pipeDist} m</span>}
                              <span style={{ fontSize: 13, padding: "2px 7px", borderRadius: 999, background: r.condensePump ? C.amber : C.surface, border: `1px solid ${r.condensePump ? C.amberBrd : C.border}`, color: r.condensePump ? C.amberTxt : C.textMuted }}>
                                Condenspomp: {r.condensePump ? "Ja" : "Nee"}
                              </span>
                            </div>
                          </div>
                        );
                      })}
                    </div>

                    {(meta.fuseDist !== "") && (
                      <div style={{ borderTop: `1px solid ${C.border}`, paddingTop: 8, marginTop: 4, display: "flex", gap: 5, flexWrap: "wrap" }}>
                        <span style={{ fontSize: 13, padding: "2px 8px", borderRadius: 999, background: C.blueLt, color: C.accent, border: `1px solid ${C.accentDark}` }}>Kabel zekeringkast: {meta.fuseDist} m</span>
                        <span style={{ fontSize: 13, padding: "2px 8px", borderRadius: 999, background: C.surfaceHigh, color: C.textSecondary, border: `1px solid ${C.border}` }}>{meta.mountType === "wall" ? "Muurbeugel" : "Voetmodel"}</span>
                      </div>
                    )}
                  </div>
                </div>
              );
            })}
          </div>
        </div>

        {/* Materiaallijst */}
        <div style={{ borderTop: `1px solid ${C.border}`, paddingTop: 24, marginTop: 24 }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
            <SectionHead>Materiaallijst</SectionHead>
            <button onClick={exportToExcel} style={{ display: "flex", alignItems: "center", gap: 6, padding: "6px 14px", borderRadius: 7, border: `1.5px solid ${C.accent}`, background: C.blueLt, color: C.accent, cursor: "pointer", fontSize: 15, fontWeight: 700, letterSpacing: "0.04em" }}>
              ⬇ Download Excel
            </button>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            {activeGroups.map((grp, gi) => {
              const grs = grp.map(id => rooms.find(r => r.id === id)).filter(Boolean);
              if (!grs.length) return null;
              const { ou } = vg(grs, overdim, overdimOU);
              const meta = gmeta[gi] || makeGroup();
              const fuseDistNum = Number(meta.fuseDist) || 0;

              return (
                <div key={gi} style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 12, overflow: "hidden" }}>
                  <div style={{ background: C.accent, padding: "9px 14px" }}>
                    <span style={{ fontSize: 16, fontWeight: 800, color: C.accentTxt }}>Buitenunit {gi + 1} — {ou ? ou.lbl : "geen unit"}</span>
                  </div>
                  <div style={{ padding: "12px 14px" }}>
                    <SecLabel>Buitenunit</SecLabel>
                    <MatTable rows={[
                      ["Buitenunit", ou ? ou.lbl : "—", ou ? ou.id : "—", "1 st"],
                      ["Automaat zekeringkast", ou ? ou.breaker : "—", "", "1 st"],
                      ["Voedingskabel zekeringkast → buitenunit", ou ? ou.ouCable : "—", "", fuseDistNum > 0 ? `${fuseDistNum} m` : "—"],
                      ["Montage", meta.mountType === "wall" ? "Muurbeugel" : "Voetmodel", "", "1 st"],
                    ]} />

                    {grs.map((r, ri) => {
                      const u = IU[r.ii];
                      const dist = Number(r.pipeDist) || 0;
                      const rolls = dist > 0 ? Math.ceil(dist / 20) : 1;
                      return (
                        <div key={r.id} style={{ marginTop: 14 }}>
                          <SecLabel>{r.name} — {u.lbl}</SecLabel>
                          <MatTable rows={[
                            ["Binnenunit", u.lbl, u.id, "1 st"],
                            [`Koelleiding ${u.pipeLiq} + ${u.pipeGas}`, "Rol 20 m", dist > 0 ? `${dist} m → ${rolls} rol${rolls > 1 ? "len" : ""}` : "afstand?", `${rolls} rol${rolls > 1 ? "len" : ""}`],
                            ["Verbindingskabel buiten → binnen", u.cable, "", dist > 0 ? `${dist} m` : "—"],
                            r.condensePump ? ["Condenspomp", "1 st", "", ""] : null,
                          ].filter(Boolean)} />
                        </div>
                      );
                    })}
                  </div>
                </div>
              );
            })}
          </div>
        </div>

      </div>
    </div>
  );
}
