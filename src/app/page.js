'use client';

import { useState, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import html2canvas from "html2canvas";

const BASELINE_KEY = "fibernc_baselines";
const ARCHIV_KEY = "fibernc_archiv";
const DEFAULT_BASELINES = {
  gesamt: { cc_rate: 97.6, termintreue: 97.7, loesungsquote: 96.0, nps: 69.9 },
  fs5335: { cc_rate: 99.6, termintreue: 99.1, loesungsquote: 96.9, nps: 74.4 },
  fs5336: { cc_rate: 95.7, termintreue: 96.7, loesungsquote: 97.2, nps: 66.7 },
};
const OT_BASELINE = { a_ges: 95.0, a1: 60.0 };
const STORAGE_KEY = "fibernc_kpi_v2";
const KONTAKTE_KEY = "fibernc_kontakte";

const KATEGORIEN = [
  { id: "alle", label: "Alle" },
  { id: "smsfeedback", label: "SMS-Feedback" },
  { id: "smsfeedbackschalten", label: "Schalten" },
  { id: "nftq", label: "NFTQ" },
  { id: "onetouch", label: "OneTouch" },
  { id: "standard", label: "Manuell" },
];

const SYSTEM_PROMPT_FN = (bl) => `Du bist ein operativer KPI-Analyseagent für ein Telekommunikations-Subunternehmen (Telekom-Subunternehmer, Kupfer & FTTH, Bergheim NRW).
Baseline KW13-19: CC=${bl.gesamt.cc_rate}% | Termintreue=${bl.gesamt.termintreue}% | Lösungsquote=${bl.gesamt.loesungsquote}%
FS5335: CC=${bl.fs5335.cc_rate}% | Termintreue=${bl.fs5335.termintreue}% | Lösungsquote=${bl.fs5335.loesungsquote}%
FS5336: CC=${bl.fs5336.cc_rate}% | Termintreue=${bl.fs5336.termintreue}% | Lösungsquote=${bl.fs5336.loesungsquote}%
KW20 schlechteste Woche (NPS 26, Termintreue 85,7%). KW23-24 FS5336 kritisch: CC 70%, SearchCall 37%.
OneTouch: A1=erster Besuch erledigt (Ziel >=60%), AX=Abbruch, A0=nicht erledigt (kritisch >10%).
Aufgabe: Techniker-KPIs bewerten, Frühwarnungen bei >=7% Abweichung, Leitstellen-Empfehlungen.
Antworte auf Deutsch, direkt und operativ.

WICHTIG: Gib am Ende der Analyse einen JSON-Block aus:
<MASSNAHMEN>
{
  "massnahmen": [
    {"name": "Vollständiger Name", "status": "kritisch|warnung|gut", "massnahme": "Konkrete Maßnahme in einem Satz", "betreff": "Email-Betreff"}
  ]
}
</MASSNAHMEN>

## KPI-Übersicht
## Frühwarnungen
## Baseline-Vergleich
## Empfehlungen Leitstelle`;

function parsePercent(val) {
  if (val === undefined || val === null || val === "") return null;
  const s = String(val).replace(",", ".").replace("%", "").trim();
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function cleanHeader(h) {
  return String(h).replace(/^\uFEFF/, "").replace(/^"|"$/g, "").trim();
}

function isJunkRow(row) {
  const vals = Object.values(row).map(v => String(v || "").trim());
  const joined = vals.join(" ").toLowerCase();
  return joined.includes("diese datei") || vals.every(v => v === "" || v === "nan" || v === "NaN");
}

function detectFormat(headers) {
  const h = headers.map(s => cleanHeader(s).toLowerCase());
  if (h.some(x => x === "a1" || x === "a ges." || x === "a ges" || x === "a0")) return "onetouch";
  if (h.some(x => x.includes("nftq b") || x.includes("nftq s"))) return "nftq";
  if (h.some(x => x.includes("courtesy call") || x.includes("abschluss call"))) return "smsfeedbackschalten";
  if (h.some(x => x.includes("cc anzahl") || x.includes("nps bs") || x.includes("nps pb"))) return "smsfeedback";
  if (h.some(x => x === "cc_rate" || x === "loesungsquote")) return "standard";
  return "standard";
}

function aggregateOneTouch(rawRows) {
  const byName = {};
  rawRows.forEach(row => {
    const rawHeaders = Object.keys(row);
    const get = (...keys) => {
      for (const key of keys) {
        const rh = rawHeaders.find(h => cleanHeader(h).toLowerCase() === key.toLowerCase());
        if (rh !== undefined && row[rh] !== undefined && row[rh] !== "") return row[rh];
      }
      return null;
    };
    const name = String(get("techniker", "name") || "").trim();
    if (!name || name.length < 2) return;
    if (!byName[name]) byName[name] = [];
    byName[name].push({
      gesamt: parseFloat(get("gesamt") || 0) || 0,
      a_ges: parsePercent(get("a ges.", "a ges")),
      a1: parsePercent(get("a1")),
      a2: parsePercent(get("a2")),
      a2plus: parsePercent(get("a2+")),
      ax: parsePercent(get("ax")),
      a0: parsePercent(get("a0")),
    });
  });
  return Object.entries(byName).map(([name, days]) => {
    const total = days.reduce((s, d) => s + d.gesamt, 0);
    const wavg = (key) => {
      const num = days.reduce((s, d) => d[key] !== null ? s + d[key] * d.gesamt : s, 0);
      const den = days.reduce((s, d) => d[key] !== null ? s + d.gesamt : s, 0);
      return den > 0 ? Math.round(num / den * 10) / 10 : null;
    };
    return {
      name, standort: "5335", auftraege: total,
      a_ges: wavg("a_ges"), a1: wavg("a1"), a2: wavg("a2"),
      a2plus: wavg("a2plus"), ax: wavg("ax"), a0: wavg("a0"),
      tage: days.length, quelle: "onetouch",
      cc_rate: null, termintreue: null, loesungsquote: null, nps: null,
    };
  });
}

function normalizeRows(rawRows) {
  if (!rawRows || !rawRows.length) return [];
  const filtered = rawRows.filter(row => !isJunkRow(row));
  if (!filtered.length) return [];
  const rawHeaders = Object.keys(filtered[0]);
  const headers = rawHeaders.map(cleanHeader);
  const fmt = detectFormat(headers);
  if (fmt === "onetouch") return aggregateOneTouch(filtered);
  const get = (row, ...keys) => {
    for (const key of keys) {
      const raw = rawHeaders.find(h => cleanHeader(h).toLowerCase() === key.toLowerCase());
      if (raw !== undefined && row[raw] !== undefined && row[raw] !== "") return row[raw];
    }
    return null;
  };
  return filtered
    .filter(row => { const name = get(row, "name"); return name && String(name).trim().length > 2; })
    .map(row => {
      const name = String(get(row, "name") || "").trim();
      const rawStandort = String(get(row, "standort") || "").trim();
      const standortKlar = rawStandort === "5335" || rawStandort === "5336";
      const standort = standortKlar ? rawStandort : "5335";
      if (fmt === "smsfeedback") return { name, standort: String(get(row, "od") || "5335"), cc_rate: parsePercent(get(row, "cc")), termintreue: parsePercent(get(row, "termintreue")), loesungsquote: null, nps: parsePercent(get(row, "nps pb", "nps bs")), auftraege: get(row, "anzahl") || "—", quelle: "smsfeedback", standortUnbekannt: false };
      if (fmt === "smsfeedbackschalten") return { name, standort: "5335", cc_rate: parsePercent(get(row, "courtesy call")), termintreue: parsePercent(get(row, "termintreue mit st vo", "termintreue ohne st vo")), loesungsquote: null, nps: parsePercent(get(row, "nps")), auftraege: get(row, "anzahl") || "—", quelle: "smsfeedbackschalten", standortUnbekannt: false };
      if (fmt === "nftq") return { name, standort: "5335", cc_rate: null, termintreue: null, loesungsquote: null, nftq_b: parsePercent(get(row, "nftq b")), nftq_s: parsePercent(get(row, "nftq s")), nftq_m: parsePercent(get(row, "nftq m")), nftq_p: parsePercent(get(row, "nftq p")), auftraege: get(row, "anzahl") || "—", quelle: "nftq", standortUnbekannt: false };
      return { name, standort, cc_rate: parsePercent(get(row, "cc_rate")), termintreue: parsePercent(get(row, "termintreue")), loesungsquote: parsePercent(get(row, "loesungsquote")), nps: parsePercent(get(row, "nps")), auftraege: get(row, "auftraege") || "—", quelle: "standard", standortUnbekannt: !standortKlar };
    });
}

function parseCSV(text) {
  const lines = text.trim().split("\n");
  if (lines.length < 2) return [];
  const sep = lines[0].includes(";") ? ";" : ",";
  const headers = lines[0].split(sep).map(h => h.trim().replace(/^"|"$/g, ""));
  const rows = lines.slice(1).filter(l => l.trim()).map(line => {
    const values = line.split(sep).map(v => v.trim().replace(/^"|"$/g, ""));
    const obj = {};
    headers.forEach((h, i) => (obj[h] = values[i] || ""));
    return obj;
  });
  return normalizeRows(rows);
}

function getNPSStatus(nps) {
  if (nps === null || nps === undefined || isNaN(nps)) return null;
  if (nps < 0) return "kritisch";
  if (nps < 30) return "warnung";
  return "gut";
}

function getStatus(value, baseline) {
  if (value === null || value === undefined || isNaN(value)) return "unbekannt";
  const ratio = value / baseline;
  if (ratio < 0.85) return "kritisch";
  if (ratio < 0.93) return "warnung";
  return "gut";
}

function getOTStatus(tech) {
  if (tech.a0 !== null && tech.a0 > 10) return "kritisch";
  if (tech.a_ges !== null && tech.a_ges < 85) return "kritisch";
  if (tech.a1 !== null && tech.a1 < 45) return "kritisch";
  if (tech.a1 !== null && tech.a1 < 60) return "warnung";
  if (tech.ax !== null && tech.ax > 20) return "warnung";
  return "gut";
}

function parseMassnahmen(text) {
  try {
    const match = text.match(/<MASSNAHMEN>([\s\S]*?)<\/MASSNAHMEN>/);
    if (!match) return { massnahmen: [], fehler: "Kein <MASSNAHMEN>-Block gefunden." };
    const clean = match[1].trim().replace(/```json|```/g, "").trim();
    const json = JSON.parse(clean);
    if (!Array.isArray(json.massnahmen)) return { massnahmen: [], fehler: "JSON-Format ungültig." };
    return { massnahmen: json.massnahmen, fehler: null };
  } catch(e) {
    return { massnahmen: [], fehler: `JSON-Parsing fehlgeschlagen: ${e.message}` };
  }
}

function getKW(date = new Date()) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return { kw: Math.ceil((((d - yearStart) / 86400000) + 1) / 7), jahr: d.getUTCFullYear() };
}

function formatArchivLabel(date = new Date()) {
  const { kw, jahr } = getKW(date);
  const datum = date.toLocaleDateString("de-DE", { day: "2-digit", month: "2-digit", year: "numeric" });
  return `KW${String(kw).padStart(2, "0")} ${jahr} · ${datum}`;
}

const STATUS_STYLE = {
  gut:       { bg: "#0f2e1a", color: "#4ade80", label: "GUT" },
  warnung:   { bg: "#2e1f00", color: "#fbbf24", label: "WARNUNG" },
  kritisch:  { bg: "#2e0f0f", color: "#f87171", label: "KRITISCH" },
  unbekannt: { bg: "#1a1a2e", color: "#6b7280", label: "—" },
};

function StatusBadge({ status }) {
  const s = STATUS_STYLE[status] || STATUS_STYLE.gut;
  return <span style={{ background: s.bg, color: s.color, padding: "2px 8px", borderRadius: 3, fontSize: 11, fontFamily: "monospace", fontWeight: 700 }}>{s.label}</span>;
}

function KPIBar({ value, baseline, label }) {
  if (value === null || value === undefined || isNaN(value)) return null;
  const color = value / baseline < 0.85 ? "#f87171" : value / baseline < 0.93 ? "#fbbf24" : "#4ade80";
  return (
    <div style={{ marginBottom: 6 }}>
      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11, color: "#9ca3af", marginBottom: 2 }}>
        <span>{label}</span><span style={{ color }}>{value.toFixed(1)}% / {baseline}%</span>
      </div>
      <div style={{ background: "#1f2937", borderRadius: 2, height: 6, position: "relative" }}>
        <div style={{ width: `${Math.min(100, value)}%`, background: color, height: "100%", borderRadius: 2 }} />
        <div style={{ position: "absolute", left: `${Math.min(100, baseline)}%`, top: -3, width: 2, height: 12, background: "#6b7280" }} />
      </div>
    </div>
  );
}

function NFTQBar({ value, label }) {
  if (value === null || isNaN(value)) return null;
  const color = value > 10 ? "#f87171" : value > 5 ? "#fbbf24" : "#4ade80";
  return (
    <div style={{ marginBottom: 6 }}>
      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11, color: "#9ca3af", marginBottom: 2 }}>
        <span>{label}</span><span style={{ color }}>{value.toFixed(2)}%</span>
      </div>
      <div style={{ background: "#1f2937", borderRadius: 2, height: 6 }}>
        <div style={{ width: `${Math.min(100, value * 4)}%`, background: color, height: "100%", borderRadius: 2 }} />
      </div>
    </div>
  );
}

function OTStackedBar({ tech }) {
  const a1 = tech.a1 || 0, a2 = tech.a2 || 0, a2plus = tech.a2plus || 0, ax = tech.ax || 0, a0 = tech.a0 || 0;
  const segments = [
    { key: "A1", val: a1, color: "#4ade80" }, { key: "A2", val: a2, color: "#60a5fa" },
    { key: "A2+", val: a2plus, color: "#818cf8" }, { key: "AX", val: ax, color: "#fbbf24" },
    { key: "A0", val: a0, color: "#f87171" },
  ];
  return (
    <div style={{ marginBottom: 10 }}>
      <div style={{ fontSize: 11, color: "#9ca3af", marginBottom: 4 }}>Auftragsverteilung</div>
      <div style={{ display: "flex", height: 10, borderRadius: 3, overflow: "hidden", background: "#1f2937" }}>
        {segments.map(s => s.val > 0 ? <div key={s.key} style={{ width: `${s.val}%`, background: s.color }} title={`${s.key}: ${s.val.toFixed(1)}%`} /> : null)}
      </div>
      <div style={{ display: "flex", gap: 10, marginTop: 5, flexWrap: "wrap" }}>
        {segments.map(s => <span key={s.key} style={{ fontSize: 10, color: s.val > 0 ? s.color : "#374151" }}>{s.key} {s.val.toFixed(0)}%</span>)}
      </div>
    </div>
  );
}

function TechCard({ tech, baselines }) {
  const isNFTQ = tech.quelle === "nftq";
  const isOT = tech.quelle === "onetouch";
  const bl = String(tech.standort) === "5336" ? baselines.fs5336 : baselines.fs5335;
  let worst = "gut";
  if (isOT) worst = getOTStatus(tech);
  else if (isNFTQ) {
    const vals = [tech.nftq_b, tech.nftq_s, tech.nftq_m, tech.nftq_p].filter(v => v !== null);
    worst = vals.some(v => v > 10) ? "kritisch" : vals.some(v => v > 5) ? "warnung" : "gut";
  } else {
    const statuses = [
      tech.cc_rate !== null ? getStatus(tech.cc_rate, bl.cc_rate) : null,
      tech.termintreue !== null ? getStatus(tech.termintreue, bl.termintreue) : null,
      tech.loesungsquote !== null ? getStatus(tech.loesungsquote, bl.loesungsquote) : null,
      tech.nps !== null ? getNPSStatus(tech.nps) : null,
    ].filter(Boolean);
    worst = statuses.includes("kritisch") ? "kritisch" : statuses.includes("warnung") ? "warnung" : "gut";
  }
  const borderColor = worst === "kritisch" ? "#7f1d1d" : worst === "warnung" ? "#78350f" : "#14532d";
  const quelleLabel = { smsfeedback: "SMS-Feedback", smsfeedbackschalten: "Schalten", nftq: "NFTQ", standard: "Manuell", onetouch: "OneTouch" }[tech.quelle] || "";
  const npsStatus = tech.nps !== null ? getNPSStatus(tech.nps) : null;
  const npsColor = npsStatus === "kritisch" ? "#f87171" : npsStatus === "warnung" ? "#fbbf24" : "#4ade80";
  return (
    <div style={{ background: "#111827", border: `1px solid ${borderColor}`, borderRadius: 8, padding: "16px 18px", marginBottom: 12 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
        <div>
          <div style={{ fontWeight: 700, fontSize: 15, color: "#f9fafb" }}>{tech.name}</div>
          <div style={{ fontSize: 11, color: "#6b7280", marginTop: 2 }}>
            {tech.standortUnbekannt ? <span style={{ color: "#fbbf24" }}>⚠ Standort unbekannt → FS5335</span> : `FS${tech.standort}`}
            {" · "}{tech.auftraege} Aufträge
            {isOT && tech.tage ? <span style={{ marginLeft: 6 }}>· {tech.tage} Tage</span> : null}
            <span style={{ marginLeft: 8, color: "#374151", background: "#1f2937", padding: "1px 6px", borderRadius: 3 }}>{quelleLabel}</span>
          </div>
        </div>
        <StatusBadge status={worst} />
      </div>
      {isOT && (<>
        <OTStackedBar tech={tech} />
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
          <KPIBar value={tech.a_ges} baseline={OT_BASELINE.a_ges} label="Gesamterfolg" />
          <KPIBar value={tech.a1} baseline={OT_BASELINE.a1} label="Erstlösung (A1)" />
        </div>
        {tech.a0 > 0 ? <div style={{ marginTop: 6, fontSize: 11, color: "#f87171" }}>⚠ A0: {tech.a0.toFixed(1)}%</div> : null}
      </>)}
      {isNFTQ && (<>
        <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 8 }}>NFTQ Fehlerquoten</div>
        <NFTQBar value={tech.nftq_b} label="Bereitstellung" />
        <NFTQBar value={tech.nftq_s} label="Schalten" />
        <NFTQBar value={tech.nftq_m} label="Montage" />
        <NFTQBar value={tech.nftq_p} label="Problembehebung" />
      </>)}
      {!isOT && !isNFTQ && (<>
        <KPIBar value={tech.cc_rate} baseline={bl.cc_rate} label="CC-Rate" />
        <KPIBar value={tech.termintreue} baseline={bl.termintreue} label="Termintreue" />
        <KPIBar value={tech.loesungsquote} baseline={bl.loesungsquote} label="Lösungsquote" />
        {tech.nps !== null ? (
          <div style={{ marginTop: 8, display: "flex", alignItems: "center", gap: 8 }}>
            <span style={{ fontSize: 11, color: "#9ca3af" }}>NPS:</span>
            <span style={{ color: npsColor, fontWeight: 700, fontSize: 13, fontFamily: "monospace" }}>{tech.nps.toFixed(0)}</span>
            <span style={{ background: STATUS_STYLE[npsStatus]?.bg, color: npsColor, padding: "1px 6px", borderRadius: 3, fontSize: 10, fontFamily: "monospace", fontWeight: 700 }}>
              {npsStatus === "kritisch" ? "KRITISCH" : npsStatus === "warnung" ? "WARNUNG" : "GUT"}
            </span>
            <span style={{ fontSize: 10, color: "#4b5563" }}>Basis: {String(tech.standort) === "5336" ? baselines.fs5336.nps : baselines.fs5335.nps}</span>
          </div>
        ) : null}
      </>)}
    </div>
  );
}

function renderMarkdown(text) {
  return text
    .replace(/<MASSNAHMEN>[\s\S]*?<\/MASSNAHMEN>/g, "")
    .replace(/## (.*)/g, '<h3 style="color:#f9fafb;margin:20px 0 8px;font-size:14px">$1</h3>')
    .replace(/\*\*(.*?)\*\*/g, '<strong style="color:#e5e7eb">$1</strong>')
    .replace(/\n/g, "<br/>");
}

function MassnahmenPanel({ massnahmen, parseError, kontakte }) {
  if (parseError) {
    return (
      <div style={{ marginTop: 16, background: "#2e0f0f", border: "1px solid #7f1d1d", borderRadius: 8, padding: "12px 16px" }}>
        <div style={{ fontSize: 12, color: "#f87171", fontWeight: 700 }}>⚠ Maßnahmen konnten nicht geladen werden</div>
        <div style={{ fontSize: 11, color: "#9ca3af", marginTop: 4 }}>{parseError}</div>
      </div>
    );
  }
  if (!massnahmen.length) return null;
  const statusBg = { kritisch: "#2e0f0f", warnung: "#2e1f00", gut: "#0f2e1a" };
  const statusColor = { kritisch: "#f87171", warnung: "#fbbf24", gut: "#4ade80" };
  return (
    <div style={{ marginTop: 24 }}>
      <div style={{ fontSize: 13, fontWeight: 700, color: "#f9fafb", marginBottom: 12, borderBottom: "1px solid #1f2937", paddingBottom: 8 }}>📋 Maßnahmen pro Techniker</div>
      {massnahmen.map((m, i) => {
        const k = kontakte[m.name] || {};
        const body = `Hallo ${m.name.split(" ")[0]},\n\nfolgende Maßnahme wurde für Sie festgelegt:\n\n${m.massnahme}\n\nBitte bestätigen Sie die Umsetzung.\n\nMit freundlichen Grüßen\nFiberNC Leitstelle`;
        const mailto = `mailto:${k.email || ""}?subject=${encodeURIComponent(m.betreff || "KPI Maßnahme")}&body=${encodeURIComponent(body)}`;
        const waLink = k.mobil ? `https://wa.me/${k.mobil.replace(/[^0-9]/g, "")}?text=${encodeURIComponent(m.massnahme)}` : null;
        return (
          <div key={i} style={{ background: statusBg[m.status] || "#111827", border: `1px solid ${statusColor[m.status] || "#1f2937"}`, borderRadius: 8, padding: "12px 16px", marginBottom: 10 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8 }}>
              <div style={{ flex: 1 }}>
                <div style={{ fontWeight: 700, fontSize: 13, color: "#f9fafb", marginBottom: 4 }}>{m.name}</div>
                <div style={{ fontSize: 12, color: "#d1d5db", lineHeight: 1.5 }}>{m.massnahme}</div>
              </div>
              <div style={{ display: "flex", gap: 6, flexShrink: 0 }}>
                <a href={mailto} style={{ background: "#1d4ed8", color: "#fff", padding: "5px 10px", borderRadius: 5, fontSize: 11, textDecoration: "none", fontWeight: 600 }}>📧 Email</a>
                {waLink ? <a href={waLink} target="_blank" rel="noreferrer" style={{ background: "#15803d", color: "#fff", padding: "5px 10px", borderRadius: 5, fontSize: 11, textDecoration: "none", fontWeight: 600 }}>💬 WA</a> : null}
              </div>
            </div>
            {(!k.email && !k.mobil) && <div style={{ marginTop: 6, fontSize: 10, color: "#6b7280" }}>⚠ Keine Kontaktdaten — unter "👥 Kontakte" eintragen</div>}
          </div>
        );
      })}
    </div>
  );
}

function AddKPIRow({ onAdd }) {
  const [name, setName] = useState("");
  const [val, setVal] = useState("");
  const inputStyle = { background: "#1f2937", border: "1px solid #374151", borderRadius: 5, padding: "4px 7px", color: "#e5e7eb", fontSize: 11 };
  return (
    <div style={{ display: "flex", gap: 6, marginTop: 10, alignItems: "center" }}>
      <input value={name} onChange={e => setName(e.target.value)} placeholder="KPI-Name z.B. cc_rate" style={{ ...inputStyle, flex: 1 }} />
      <input value={val} onChange={e => setVal(e.target.value)} placeholder="Wert" type="number" style={{ ...inputStyle, width: 70 }} />
      <button onClick={() => { if (name && val) { onAdd(name.trim(), parseFloat(val)); setName(""); setVal(""); } }}
        style={{ background: "#1d4ed8", color: "#fff", border: "none", borderRadius: 5, cursor: "pointer", fontSize: 13, padding: "4px 10px", fontWeight: 700 }}>+</button>
    </div>
  );
}

function BaselineEditor({ baselines, onSave, onClose }) {
  const [local, setLocal] = useState(JSON.parse(JSON.stringify(baselines)));
  const kpiLabels = { cc_rate: "CC-Rate %", termintreue: "Termintreue %", loesungsquote: "Lösungsquote %", nps: "NPS" };
  const standortLabels = { gesamt: "Gesamt (KW13-19)", fs5335: "FS5335", fs5336: "FS5336" };
  const inputStyle = { background: "#1f2937", border: "1px solid #374151", borderRadius: 5, padding: "5px 8px", color: "#e5e7eb", fontSize: 12, width: "80px", textAlign: "right" };
  return (
    <div style={{ position: "fixed", top: 0, left: 0, right: 0, bottom: 0, background: "rgba(0,0,0,0.85)", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ background: "#111827", border: "1px solid #1f2937", borderRadius: 12, padding: 24, width: 620, maxHeight: "85vh", overflowY: "auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
          <span style={{ fontWeight: 700, fontSize: 15, color: "#f9fafb" }}>📊 Baseline-Werte verwalten</span>
          <button onClick={onClose} style={{ background: "none", border: "none", color: "#6b7280", cursor: "pointer", fontSize: 18 }}>✕</button>
        </div>
        {Object.entries(standortLabels).map(([standort, standortName]) => (
          <div key={standort} style={{ marginBottom: 16, background: "#0f172a", border: "1px solid #1f2937", borderRadius: 8, padding: "14px 16px" }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: "#60a5fa", marginBottom: 12, textTransform: "uppercase", letterSpacing: 1 }}>{standortName}</div>
            {Object.entries(local[standort] || {}).map(([kpi, wert]) => (
              <div key={kpi} style={{ display: "grid", gridTemplateColumns: "1fr auto auto", gap: 8, alignItems: "center", marginBottom: 8 }}>
                <div style={{ fontSize: 12, color: "#9ca3af" }}>{kpiLabels[kpi] || kpi}</div>
                <input type="number" step="0.1" value={wert}
                  onChange={e => setLocal(prev => ({ ...prev, [standort]: { ...prev[standort], [kpi]: parseFloat(e.target.value) || 0 } }))}
                  style={inputStyle} />
                <button onClick={() => setLocal(prev => { const n = { ...prev, [standort]: { ...prev[standort] } }; delete n[standort][kpi]; return n; })}
                  style={{ background: "#2e0f0f", color: "#f87171", border: "1px solid #7f1d1d", borderRadius: 4, cursor: "pointer", fontSize: 11, padding: "3px 8px" }}>✕</button>
              </div>
            ))}
            <AddKPIRow onAdd={(kpi, val) => setLocal(prev => ({ ...prev, [standort]: { ...prev[standort], [kpi]: val } }))} />
          </div>
        ))}
        <div style={{ display: "flex", gap: 10, marginTop: 4 }}>
          <button onClick={() => setLocal(JSON.parse(JSON.stringify(DEFAULT_BASELINES)))}
            style={{ flex: 1, background: "#1f2937", color: "#9ca3af", border: "1px solid #374151", borderRadius: 8, padding: "10px", fontSize: 12, cursor: "pointer" }}>🔄 Standard</button>
          <button onClick={() => { onSave(local); onClose(); }}
            style={{ flex: 2, background: "#1d4ed8", color: "#fff", border: "none", borderRadius: 8, padding: "10px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>💾 Speichern</button>
        </div>
      </div>
    </div>
  );
}

function TechnikerVerwaltung({ gespeichert, onUpdate, onClose }) {
  const [local, setLocal] = useState(JSON.parse(JSON.stringify(gespeichert)));
  const kategorieLabels = { smsfeedback: "SMS-Feedback", smsfeedbackschalten: "Schalten", nftq: "NFTQ", standard: "Manuell", onetouch: "OneTouch" };
  return (
    <div style={{ position: "fixed", top: 0, left: 0, right: 0, bottom: 0, background: "rgba(0,0,0,0.85)", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ background: "#111827", border: "1px solid #1f2937", borderRadius: 12, padding: 24, width: 660, maxHeight: "88vh", overflowY: "auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
          <span style={{ fontWeight: 700, fontSize: 15, color: "#f9fafb" }}>🧑‍🔧 Techniker-Einträge verwalten</span>
          <button onClick={onClose} style={{ background: "none", border: "none", color: "#6b7280", cursor: "pointer", fontSize: 18 }}>✕</button>
        </div>
        {Object.keys(local).length === 0 && <div style={{ textAlign: "center", padding: "30px", color: "#6b7280" }}>Keine Daten geladen.</div>}
        {Object.entries(local).map(([kat, rows]) => (
          <div key={kat} style={{ marginBottom: 20 }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: "#60a5fa", textTransform: "uppercase", letterSpacing: 1, marginBottom: 8 }}>{kategorieLabels[kat] || kat} — {rows.length} Einträge</div>
            {rows.map((t, i) => (
              <div key={i} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: "#0f172a", border: "1px solid #1f2937", borderRadius: 6, padding: "8px 12px", marginBottom: 6 }}>
                <div>
                  <span style={{ fontSize: 13, fontWeight: 600, color: "#f9fafb" }}>{t.name}</span>
                  <span style={{ fontSize: 11, color: "#6b7280", marginLeft: 10 }}>FS{t.standort}</span>
                  {t.cc_rate !== null && <span style={{ fontSize: 11, color: "#9ca3af", marginLeft: 8 }}>CC {t.cc_rate?.toFixed(1)}%</span>}
                  {t.termintreue !== null && <span style={{ fontSize: 11, color: "#9ca3af", marginLeft: 8 }}>TT {t.termintreue?.toFixed(1)}%</span>}
                  {t.nps !== null && <span style={{ fontSize: 11, color: "#9ca3af", marginLeft: 8 }}>NPS {t.nps?.toFixed(0)}</span>}
                  {t.a1 !== null && <span style={{ fontSize: 11, color: "#9ca3af", marginLeft: 8 }}>A1 {t.a1?.toFixed(1)}%</span>}
                </div>
                <button onClick={() => setLocal(prev => { const n = { ...prev, [kat]: prev[kat].filter((_, j) => j !== i) }; if (!n[kat].length) delete n[kat]; return n; })}
                  style={{ background: "#2e0f0f", color: "#f87171", border: "1px solid #7f1d1d", borderRadius: 4, cursor: "pointer", fontSize: 11, padding: "3px 10px" }}>✕ Löschen</button>
              </div>
            ))}
          </div>
        ))}
        <button onClick={() => { onUpdate(local); onClose(); }}
          style={{ width: "100%", marginTop: 8, background: "#1d4ed8", color: "#fff", border: "none", borderRadius: 8, padding: "12px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>💾 Änderungen speichern</button>
      </div>
    </div>
  );
}

function KontakteEditor({ kontakte, onSave, onClose }) {
  const [local, setLocal] = useState({ ...kontakte });
  const [neuerName, setNeuerName] = useState("");
  const [neuerEmail, setNeuerEmail] = useState("");
  const [neuerMobil, setNeuerMobil] = useState("");
  const hinzufuegen = () => {
    if (!neuerName.trim()) return;
    setLocal(prev => ({ ...prev, [neuerName.trim()]: { email: neuerEmail.trim(), mobil: neuerMobil.trim() } }));
    setNeuerName(""); setNeuerEmail(""); setNeuerMobil("");
  };
  const inputStyle = { background: "#1f2937", border: "1px solid #374151", borderRadius: 5, padding: "5px 8px", color: "#e5e7eb", fontSize: 12, width: "100%" };
  return (
    <div style={{ position: "fixed", top: 0, left: 0, right: 0, bottom: 0, background: "rgba(0,0,0,0.85)", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ background: "#111827", border: "1px solid #1f2937", borderRadius: 12, padding: 24, width: 580, maxHeight: "85vh", overflowY: "auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
          <span style={{ fontWeight: 700, fontSize: 15, color: "#f9fafb" }}>👥 Techniker Stammdaten</span>
          <button onClick={onClose} style={{ background: "none", border: "none", color: "#6b7280", cursor: "pointer", fontSize: 18 }}>✕</button>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "150px 1fr 130px 28px", gap: 8, marginBottom: 8, padding: "0 0 8px", borderBottom: "1px solid #1f2937" }}>
          <div style={{ fontSize: 10, color: "#6b7280", fontWeight: 700, textTransform: "uppercase" }}>Name</div>
          <div style={{ fontSize: 10, color: "#6b7280", fontWeight: 700, textTransform: "uppercase" }}>Email</div>
          <div style={{ fontSize: 10, color: "#6b7280", fontWeight: 700, textTransform: "uppercase" }}>Mobil</div>
          <div />
        </div>
        {Object.entries(local).map(([name, k]) => (
          <div key={name} style={{ display: "grid", gridTemplateColumns: "150px 1fr 130px 28px", gap: 8, marginBottom: 8, alignItems: "center" }}>
            <div style={{ fontSize: 12, color: "#f9fafb", fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }} title={name}>{name}</div>
            <input value={k.email || ""} onChange={e => setLocal(prev => ({ ...prev, [name]: { ...prev[name], email: e.target.value } }))} placeholder="email@beispiel.de" style={inputStyle} />
            <input value={k.mobil || ""} onChange={e => setLocal(prev => ({ ...prev, [name]: { ...prev[name], mobil: e.target.value } }))} placeholder="+4915..." style={inputStyle} />
            <button onClick={() => { const n = { ...local }; delete n[name]; setLocal(n); }} style={{ background: "none", border: "none", color: "#4b5563", cursor: "pointer", fontSize: 16, padding: 0 }}>✕</button>
          </div>
        ))}
        <div style={{ borderTop: "1px solid #1f2937", paddingTop: 16, marginTop: 8 }}>
          <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 8 }}>➕ Neuen Techniker hinzufügen:</div>
          <div style={{ display: "grid", gridTemplateColumns: "150px 1fr 130px 40px", gap: 8 }}>
            <input value={neuerName} onChange={e => setNeuerName(e.target.value)} placeholder="Vor- Nachname" style={inputStyle} onKeyDown={e => e.key === "Enter" && hinzufuegen()} />
            <input value={neuerEmail} onChange={e => setNeuerEmail(e.target.value)} placeholder="email@beispiel.de" style={inputStyle} onKeyDown={e => e.key === "Enter" && hinzufuegen()} />
            <input value={neuerMobil} onChange={e => setNeuerMobil(e.target.value)} placeholder="+4915..." style={inputStyle} onKeyDown={e => e.key === "Enter" && hinzufuegen()} />
            <button onClick={hinzufuegen} style={{ background: "#1d4ed8", color: "#fff", border: "none", borderRadius: 5, cursor: "pointer", fontSize: 16, fontWeight: 700 }}>+</button>
          </div>
        </div>
        <button onClick={() => { onSave(local); onClose(); }}
          style={{ width: "100%", marginTop: 20, background: "#1d4ed8", color: "#fff", border: "none", borderRadius: 8, padding: "12px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>💾 Speichern</button>
      </div>
    </div>
  );
}

function ArchivPanel({ archiv, onDelete, onClose }) {
  const [aufgeklappt, setAufgeklappt] = useState(null);
  const kategorieLabels = { smsfeedback: "SMS-Feedback", smsfeedbackschalten: "Schalten", nftq: "NFTQ", standard: "Manuell", onetouch: "OneTouch" };
  const exportCSV = (eintrag) => {
    const rows = Object.entries(eintrag.daten).flatMap(([kat, techs]) =>
      techs.map(t => ({ Kategorie: kategorieLabels[kat] || kat, Name: t.name, Standort: `FS${t.standort}`, CC: t.cc_rate ?? "", Termintreue: t.termintreue ?? "", Loesungsquote: t.loesungsquote ?? "", NPS: t.nps ?? "", A1: t.a1 ?? "", A0: t.a0 ?? "" }))
    );
    const header = Object.keys(rows[0]).join(";");
    const csv = [header, ...rows.map(r => Object.values(r).join(";"))].join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href = url; a.download = `KPI-Archiv-${eintrag.label.replace(/[^a-zA-Z0-9]/g, "_")}.csv`; a.click();
    URL.revokeObjectURL(url);
  };
  return (
    <div style={{ position: "fixed", top: 0, left: 0, right: 0, bottom: 0, background: "rgba(0,0,0,0.88)", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ background: "#111827", border: "1px solid #1f2937", borderRadius: 12, padding: 24, width: 680, maxHeight: "88vh", overflowY: "auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
          <span style={{ fontWeight: 700, fontSize: 15, color: "#f9fafb" }}>🗂 KPI-Archiv</span>
          <button onClick={onClose} style={{ background: "none", border: "none", color: "#6b7280", cursor: "pointer", fontSize: 18 }}>✕</button>
        </div>
        {archiv.length === 0 && <div style={{ textAlign: "center", padding: "40px", color: "#6b7280" }}>Noch keine archivierten Einträge.</div>}
        {[...archiv].reverse().map((eintrag, i) => {
          const idx = archiv.length - 1 - i;
          const totalTechs = Object.values(eintrag.daten).flat().length;
          const kritisch = Object.values(eintrag.daten).flat().filter(t => t._status === "kritisch").length;
          const isOpen = aufgeklappt === idx;
          return (
            <div key={idx} style={{ marginBottom: 10, background: "#0f172a", border: "1px solid #1f2937", borderRadius: 8, overflow: "hidden" }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "12px 16px", cursor: "pointer" }}
                onClick={() => setAufgeklappt(isOpen ? null : idx)}>
                <div>
                  <div style={{ fontSize: 13, fontWeight: 700, color: "#f9fafb" }}>{eintrag.label}</div>
                  <div style={{ fontSize: 11, color: "#6b7280", marginTop: 2 }}>
                    {totalTechs} Techniker
                    {kritisch > 0 ? <span style={{ color: "#f87171", marginLeft: 8 }}>· {kritisch} kritisch</span> : null}
                    {eintrag.analyse ? <span style={{ color: "#4ade80", marginLeft: 8 }}>· KI ✓</span> : null}
                  </div>
                </div>
                <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                  <button onClick={e => { e.stopPropagation(); exportCSV(eintrag); }}
                    style={{ background: "#1f2937", color: "#9ca3af", border: "1px solid #374151", borderRadius: 4, cursor: "pointer", fontSize: 10, padding: "3px 8px" }}>📥 CSV</button>
                  <button onClick={e => { e.stopPropagation(); if (window.confirm("Eintrag löschen?")) onDelete(idx); }}
                    style={{ background: "#2e0f0f", color: "#f87171", border: "1px solid #7f1d1d", borderRadius: 4, cursor: "pointer", fontSize: 10, padding: "3px 8px" }}>✕</button>
                  <span style={{ color: "#6b7280", fontSize: 14 }}>{isOpen ? "▲" : "▼"}</span>
                </div>
              </div>
              {isOpen && (
                <div style={{ borderTop: "1px solid #1f2937", padding: "12px 16px" }}>
                  {Object.entries(eintrag.daten).map(([kat, techs]) => (
                    <div key={kat} style={{ marginBottom: 12 }}>
                      <div style={{ fontSize: 10, color: "#60a5fa", fontWeight: 700, textTransform: "uppercase", marginBottom: 6 }}>{kategorieLabels[kat] || kat} ({techs.length})</div>
                      {techs.map((t, ti) => (
                        <div key={ti} style={{ display: "flex", justifyContent: "space-between", fontSize: 11, color: "#9ca3af", padding: "4px 0", borderBottom: "1px solid #1f2937" }}>
                          <span style={{ color: "#f9fafb", fontWeight: 600 }}>{t.name}</span>
                          <span style={{ display: "flex", gap: 12 }}>
                            {t.cc_rate != null && <span>CC {t.cc_rate.toFixed(1)}%</span>}
                            {t.termintreue != null && <span>TT {t.termintreue.toFixed(1)}%</span>}
                            {t.nps != null && <span>NPS {t.nps.toFixed(0)}</span>}
                            {t.a1 != null && <span>A1 {t.a1.toFixed(1)}%</span>}
                            {t._status && <span style={{ color: t._status === "kritisch" ? "#f87171" : t._status === "warnung" ? "#fbbf24" : "#4ade80", fontWeight: 700 }}>{t._status.toUpperCase()}</span>}
                          </span>
                        </div>
                      ))}
                    </div>
                  ))}
                </div>
              )}
            </div>
          );
        })}
      </div>
    </div>
  );
}

export default function KPIAgent() {
  const [gespeichert, setGespeichert] = useState({});
  const [kontakte, setKontakte] = useState({});
  const [baselines, setBaselines] = useState(() => {
    try { const s = localStorage.getItem(BASELINE_KEY); return s ? JSON.parse(s) : DEFAULT_BASELINES; } catch(e) { return DEFAULT_BASELINES; }
  });
  const [archiv, setArchiv] = useState(() => {
    try { const s = localStorage.getItem(ARCHIV_KEY); return s ? JSON.parse(s) : []; } catch(e) { return []; }
  });
  const [aktiveKategorie, setAktiveKategorie] = useState("alle");
  const [aiAnalysis, setAiAnalysis] = useState("");
  const [massnahmen, setMassnahmen] = useState([]);
  const [massnahmenFehler, setMassnahmenFehler] = useState(null);
  const [techBewertungen, setTechBewertungen] = useState({});
  const [bewertungLoading, setBewertungLoading] = useState({});
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [error, setError] = useState("");
  const [activeTab, setActiveTab] = useState("dashboard");
  const [pending, setPending] = useState(null);
  const [showKontakte, setShowKontakte] = useState(false);
  const [showBaseline, setShowBaseline] = useState(false);
  const [showTechVerwaltung, setShowTechVerwaltung] = useState(false);
  const [showArchiv, setShowArchiv] = useState(false);
  const dashboardRef = useRef(null);

  useEffect(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) setGespeichert(JSON.parse(saved));
      const savedK = localStorage.getItem(KONTAKTE_KEY);
      if (savedK) setKontakte(JSON.parse(savedK));
    } catch(e) {}
  }, []);

  useEffect(() => { try { localStorage.setItem(STORAGE_KEY, JSON.stringify(gespeichert)); } catch(e) {} }, [gespeichert]);
  useEffect(() => { try { localStorage.setItem(KONTAKTE_KEY, JSON.stringify(kontakte)); } catch(e) {} }, [kontakte]);
  useEffect(() => { try { localStorage.setItem(BASELINE_KEY, JSON.stringify(baselines)); } catch(e) {} }, [baselines]);
  useEffect(() => { try { localStorage.setItem(ARCHIV_KEY, JSON.stringify(archiv)); } catch(e) {} }, [archiv]);

  useEffect(() => {
    if (!loading && pending) {
      setGespeichert(prev => ({ ...prev, [pending.quelle]: pending.rows }));
      setAktiveKategorie(pending.quelle);
      setAiAnalysis(""); setMassnahmen([]); setMassnahmenFehler(null); setTechBewertungen({});
      setActiveTab("dashboard"); setPending(null); setError("");
    }
  }, [loading, pending]);

  useEffect(() => {
    const alleNamen = Object.values(gespeichert).flat().map(t => t.name);
    const unique = [...new Set(alleNamen)];
    setKontakte(prev => {
      const neu = { ...prev };
      unique.forEach(name => { if (!neu[name]) neu[name] = { email: "", mobil: "" }; });
      return neu;
    });
  }, [gespeichert]);

  const handleRows = useCallback((rows) => {
    if (!rows.length) { setError("Keine Daten gefunden."); return; }
    const quellen = [...new Set(rows.map(r => r.quelle))];
    const quelle = quellen.length === 1 ? quellen[0] : "standard";
    if (loading) {
      setPending({ rows, quelle });
      setError("✓ Gespeichert — wird nach Analyse geladen");
    } else {
      setGespeichert(prev => ({ ...prev, [quelle]: rows }));
      setAktiveKategorie(quelle);
      setAiAnalysis(""); setMassnahmen([]); setMassnahmenFehler(null); setTechBewertungen({});
      setActiveTab("dashboard"); setError("");
    }
  }, [loading]);

  const processXLSX = useCallback(async (file) => {
    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      handleRows(normalizeRows(XLSX.utils.sheet_to_json(ws, { defval: "" })));
    } catch (e) { setError("Fehler: " + e.message); }
  }, [handleRows]);

  const handleFile = useCallback((e) => {
    const file = e.target.files[0];
    if (!file) return;
    e.target.value = "";
    if (file.name.match(/\.xlsx?$/i)) { processXLSX(file); }
    else {
      const reader = new FileReader();
      reader.onload = (ev) => handleRows(parseCSV(ev.target.result));
      reader.readAsText(file, "utf-8");
    }
  }, [processXLSX, handleRows]);

  const angezeigt = aktiveKategorie === "alle"
    ? Object.values(gespeichert).flat().filter((t, idx, arr) => arr.findLastIndex(x => x.name === t.name) === idx)
    : (gespeichert[aktiveKategorie] || []);

  const hatDaten = Object.keys(gespeichert).length > 0;

  const berechneTechScore = useCallback((tech) => {
    const bl = String(tech.standort) === "5336" ? baselines.fs5336 : baselines.fs5335;
    const scores = [];
    if (tech.cc_rate !== null) scores.push(Math.min(10, (tech.cc_rate / bl.cc_rate) * 10));
    if (tech.termintreue !== null) scores.push(Math.min(10, (tech.termintreue / bl.termintreue) * 10));
    if (tech.loesungsquote !== null) scores.push(Math.min(10, (tech.loesungsquote / bl.loesungsquote) * 10));
    if (tech.nps !== null) scores.push(Math.min(10, Math.max(0, (tech.nps + 100) / 20)));
    if (tech.a1 !== null) scores.push(Math.min(10, (tech.a1 / 60) * 10));
    if (tech.a_ges !== null) scores.push(Math.min(10, (tech.a_ges / 95) * 10));
    if (!scores.length) return null;
    return Math.round((scores.reduce((s, v) => s + v, 0) / scores.length) * 10) / 10;
  }, [baselines]);

  const scoreColor = (s) => !s ? "#6b7280" : s >= 8.5 ? "#4ade80" : s >= 7 ? "#fbbf24" : "#f87171";
  const scoreLabel = (s) => !s ? "—" : s >= 9 ? "Ausgezeichnet" : s >= 8 ? "Gut" : s >= 7 ? "Befriedigend" : s >= 5 ? "Verbesserungsbedarf" : "Kritisch";

  const bewerteEinzelTechniker = useCallback(async (tech) => {
    const bl = String(tech.standort) === "5336" ? baselines.fs5336 : baselines.fs5335;
    const score = berechneTechScore(tech);
    const kpiText = tech.quelle === "onetouch"
      ? `A-Gesamt=${tech.a_ges?.toFixed(1) ?? "—"}%, A1=${tech.a1?.toFixed(1) ?? "—"}%, AX=${tech.ax?.toFixed(1) ?? "—"}%, A0=${tech.a0?.toFixed(1) ?? "—"}%`
      : `CC=${tech.cc_rate?.toFixed(1) ?? "—"}% (Basis ${bl.cc_rate}%), Termintreue=${tech.termintreue?.toFixed(1) ?? "—"}% (Basis ${bl.termintreue}%), Lösungsquote=${tech.loesungsquote?.toFixed(1) ?? "—"}%, NPS=${tech.nps?.toFixed(0) ?? "—"}`;
    setBewertungLoading(prev => ({ ...prev, [tech.name]: true }));
    try {
      const res = await fetch("/api/analyse", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-6", max_tokens: 500,
          system: `Du bist KPI-Bewerter für Telekom-Techniker. Antworte NUR mit JSON ohne Backticks:
{"kommentar": "1-2 Sätze persönliche Bewertung mit Namen", "staerken": ["max 2 Stärken"], "schwaechen": ["max 2 Schwächen"], "massnahme": "Eine konkrete Maßnahme"}`,
          messages: [{ role: "user", content: `Techniker: ${tech.name}, Score: ${score}/10, KPIs: ${kpiText}, Aufträge: ${tech.auftraege}` }]
        }),
      });
      const data = await res.json();
      const text = data.content?.map(b => b.text || "").join("") || "";
      const clean = text.replace(/```json|```/g, "").trim();
      const parsed = JSON.parse(clean);
      setTechBewertungen(prev => ({ ...prev, [tech.name]: { ...parsed, score } }));
    } catch(e) {
      setTechBewertungen(prev => ({ ...prev, [tech.name]: { kommentar: "Bewertung fehlgeschlagen.", score } }));
    } finally {
      setBewertungLoading(prev => ({ ...prev, [tech.name]: false }));
    }
  }, [baselines, berechneTechScore]);

  const bewerteAlle = useCallback(async () => {
    for (const tech of angezeigt) {
      await bewerteEinzelTechniker(tech);
    }
    setActiveTab("firmendashboard");
  }, [angezeigt, bewerteEinzelTechniker]);

  const runAnalysis = async () => {
    if (!angezeigt.length) return;
    setLoading(true); setError(""); setAiAnalysis(""); setMassnahmen([]); setMassnahmenFehler(null);
    const dataStr = angezeigt.map(t => {
      if (t.quelle === "onetouch") return `${t.name}: A-Ges=${t.a_ges?.toFixed(1) ?? "—"}%, A1=${t.a1?.toFixed(1) ?? "—"}%, AX=${t.ax?.toFixed(1) ?? "—"}%, A0=${t.a0?.toFixed(1) ?? "—"}%, Aufträge=${t.auftraege}`;
      if (t.quelle === "nftq") return `${t.name}: NFTQ-B=${t.nftq_b?.toFixed(2) ?? "—"}%, NFTQ-S=${t.nftq_s?.toFixed(2) ?? "—"}%, NFTQ-M=${t.nftq_m?.toFixed(2) ?? "—"}%, NFTQ-P=${t.nftq_p?.toFixed(2) ?? "—"}%, Aufträge=${t.auftraege}`;
      return `${t.name} (FS${t.standort}): CC=${t.cc_rate?.toFixed(1) ?? "—"}%, Termintreue=${t.termintreue?.toFixed(1) ?? "—"}%, Lösungsquote=${t.loesungsquote?.toFixed(1) ?? "—"}%, NPS=${t.nps?.toFixed(0) ?? "—"}, Aufträge=${t.auftraege}`;
    }).join("\n");
    try {
      const res = await fetch("/api/analyse", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ model: "claude-sonnet-4-6", max_tokens: 2000, system: SYSTEM_PROMPT_FN(baselines), messages: [{ role: "user", content: `Analysiere diese Techniker-KPIs:\n\n${dataStr}` }] }),
      });
      const data = await res.json();
      const text = data.content?.map(b => b.text || "").join("") || "";
      setAiAnalysis(text);
      const { massnahmen: parsed, fehler } = parseMassnahmen(text);
      setMassnahmen(parsed); setMassnahmenFehler(fehler);
      setActiveTab("analyse");
    } catch (e) { setError("Fehler bei der KI-Analyse."); }
    finally { setLoading(false); }
  };

  const archivieren = useCallback(() => {
    if (!hatDaten) return;
    const now = new Date();
    const datatenMitStatus = {};
    Object.entries(gespeichert).forEach(([kat, techs]) => {
      datatenMitStatus[kat] = techs.map(t => {
        const bl = String(t.standort) === "5336" ? baselines.fs5336 : baselines.fs5335;
        let status = "gut";
        if (t.quelle === "onetouch") status = getOTStatus(t);
        else if (t.quelle === "nftq") {
          const vals = [t.nftq_b, t.nftq_s, t.nftq_m, t.nftq_p].filter(Boolean);
          status = vals.some(v => v > 10) ? "kritisch" : vals.some(v => v > 5) ? "warnung" : "gut";
        } else {
          const statuses = [
            t.cc_rate !== null ? getStatus(t.cc_rate, bl.cc_rate) : null,
            t.termintreue !== null ? getStatus(t.termintreue, bl.termintreue) : null,
            t.nps !== null ? getNPSStatus(t.nps) : null,
          ].filter(Boolean);
          status = statuses.includes("kritisch") ? "kritisch" : statuses.includes("warnung") ? "warnung" : "gut";
        }
        return { ...t, _status: status, _score: berechneTechScore(t) };
      });
    });
    setArchiv(prev => [...prev, {
      label: formatArchivLabel(now), datum: now.toISOString(),
      daten: datatenMitStatus, analyse: aiAnalysis || "", bewertungen: techBewertungen,
    }]);
    setGespeichert({}); setAktiveKategorie("alle");
    setAiAnalysis(""); setMassnahmen([]); setMassnahmenFehler(null); setTechBewertungen({});
  }, [gespeichert, baselines, aiAnalysis, hatDaten, techBewertungen, berechneTechScore]);

  const exportPDF = async () => {
    setExporting(true);
    try {
      const canvas = await html2canvas(dashboardRef.current, { backgroundColor: "#0a0e1a", scale: 2 });
      const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
      const w = pdf.internal.pageSize.getWidth();
      pdf.addImage(canvas.toDataURL("image/png"), "PNG", 0, 12, w, (canvas.height * w) / canvas.width);
      pdf.save(`KPI-${new Date().toLocaleDateString("de-DE").replace(/\./g, "-")}.pdf`);
    } catch (e) { setError("PDF fehlgeschlagen."); }
    finally { setExporting(false); }
  };

  const loescheKategorie = (quelle) => {
    setGespeichert(prev => { const n = { ...prev }; delete n[quelle]; return n; });
    if (aktiveKategorie === quelle) setAktiveKategorie("alle");
  };

  const criticalCount = angezeigt.filter(t => {
    if (t.quelle === "onetouch") return getOTStatus(t) === "kritisch";
    if (t.quelle === "nftq") return [t.nftq_b, t.nftq_s, t.nftq_m, t.nftq_p].filter(Boolean).some(v => v > 10);
    const bl = String(t.standort) === "5336" ? baselines.fs5336 : baselines.fs5335;
    return [
      t.cc_rate !== null ? getStatus(t.cc_rate, bl.cc_rate) : null,
      t.termintreue !== null ? getStatus(t.termintreue, bl.termintreue) : null,
      t.nps !== null ? getNPSStatus(t.nps) : null,
    ].includes("kritisch");
  }).length;

  const avg = (key) => {
    const vals = angezeigt.map(t => t[key]).filter(v => v !== null && !isNaN(v));
    return vals.length ? (vals.reduce((s, v) => s + v, 0) / vals.length).toFixed(1) : "—";
  };

  const teamAvgScore = () => {
    const scores = angezeigt.map(t => berechneTechScore(t)).filter(s => s !== null);
    return scores.length ? (scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(1) : "—";
  };

  const isOTView = aktiveKategorie === "onetouch";

  const FirmendashboardTab = () => {
    const sorted = [...angezeigt].sort((a, b) => (berechneTechScore(b) || 0) - (berechneTechScore(a) || 0));
    return (
      <div>
        <div style={{ background: "#0f172a", border: "1px solid #1f2937", borderRadius: 8, padding: "16px", marginBottom: 20 }}>
          <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: 1 }}>Team-Übersicht</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12 }}>
            {[
              { label: "Ø Team-Score", value: `${teamAvgScore()}/10`, color: scoreColor(parseFloat(teamAvgScore())) },
              { label: "Kritisch", value: criticalCount, color: criticalCount > 0 ? "#f87171" : "#4ade80" },
              { label: "Bewertet", value: `${Object.keys(techBewertungen).length}/${angezeigt.length}`, color: "#60a5fa" },
            ].map(s => (
              <div key={s.label} style={{ background: "#111827", borderRadius: 6, padding: "10px 12px" }}>
                <div style={{ fontSize: 10, color: "#6b7280", marginBottom: 4 }}>{s.label}</div>
                <div style={{ fontSize: 18, fontWeight: 700, color: s.color, fontFamily: "monospace" }}>{s.value}</div>
              </div>
            ))}
          </div>
        </div>
        <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 10, fontWeight: 700, textTransform: "uppercase", letterSpacing: 1 }}>Rangliste</div>
        {sorted.map((tech, i) => {
          const score = berechneTechScore(tech);
          const bew = techBewertungen[tech.name];
          const isLoadingThis = bewertungLoading[tech.name];
          const k = kontakte[tech.name] || {};
          const mailBody = bew
            ? `Hallo ${tech.name.split(" ")[0]},\n\nhier ist Ihre persönliche KPI-Bewertung:\n\nScore: ${score}/10 — ${scoreLabel(score)}\n\n${bew.kommentar}\n\n${bew.staerken?.length ? `Stärken:\n${bew.staerken.map(s => `• ${s}`).join("\n")}\n\n` : ""}${bew.schwaechen?.length ? `Verbesserungsbedarf:\n${bew.schwaechen.map(s => `• ${s}`).join("\n")}\n\n` : ""}Maßnahme: ${bew.massnahme || ""}\n\nMit freundlichen Grüßen\nFiberNC Leitstelle`
            : "";
          const mailto = `mailto:${k.email || ""}?subject=${encodeURIComponent(`KPI-Bewertung ${tech.name}`)}&body=${encodeURIComponent(mailBody)}`;
          return (
            <div key={tech.name} style={{ background: "#111827", border: `1px solid ${score >= 8 ? "#14532d" : score >= 6 ? "#78350f" : "#7f1d1d"}`, borderRadius: 8, padding: "14px 16px", marginBottom: 10 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <div style={{ fontSize: 16, fontWeight: 700, color: scoreColor(score), fontFamily: "monospace", minWidth: 28 }}>#{i + 1}</div>
                  <div>
                    <div style={{ fontWeight: 700, fontSize: 14, color: "#f9fafb" }}>{tech.name}</div>
                    <div style={{ fontSize: 11, color: "#6b7280" }}>FS{tech.standort} · {tech.auftraege} Aufträge</div>
                  </div>
                </div>
                {score !== null && (
                  <div style={{ textAlign: "center" }}>
                    <div style={{ fontSize: 22, fontWeight: 700, color: scoreColor(score), fontFamily: "monospace" }}>{score}</div>
                    <div style={{ fontSize: 9, color: scoreColor(score) }}>/10 · {scoreLabel(score)}</div>
                  </div>
                )}
              </div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 10 }}>
                {tech.cc_rate !== null && <span style={{ fontSize: 10, background: "#1f2937", color: "#9ca3af", padding: "2px 8px", borderRadius: 3 }}>CC {tech.cc_rate.toFixed(1)}%</span>}
                {tech.termintreue !== null && <span style={{ fontSize: 10, background: "#1f2937", color: "#9ca3af", padding: "2px 8px", borderRadius: 3 }}>TT {tech.termintreue.toFixed(1)}%</span>}
                {tech.loesungsquote !== null && <span style={{ fontSize: 10, background: "#1f2937", color: "#9ca3af", padding: "2px 8px", borderRadius: 3 }}>LQ {tech.loesungsquote.toFixed(1)}%</span>}
                {tech.nps !== null && <span style={{ fontSize: 10, background: "#1f2937", color: "#9ca3af", padding: "2px 8px", borderRadius: 3 }}>NPS {tech.nps.toFixed(0)}</span>}
                {tech.a1 !== null && <span style={{ fontSize: 10, background: "#1f2937", color: "#9ca3af", padding: "2px 8px", borderRadius: 3 }}>A1 {tech.a1.toFixed(1)}%</span>}
                {tech.a0 !== null && tech.a0 > 0 && <span style={{ fontSize: 10, background: "#2e0f0f", color: "#f87171", padding: "2px 8px", borderRadius: 3 }}>A0 {tech.a0.toFixed(1)}%</span>}
              </div>
              {isLoadingThis && <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 8 }}>⏳ KI bewertet...</div>}
              {bew && !isLoadingThis && (
                <div style={{ background: "#0f172a", borderRadius: 6, padding: "10px 12px", marginBottom: 10 }}>
                  <div style={{ fontSize: 12, color: "#d1d5db", lineHeight: 1.6, marginBottom: 8 }}>{bew.kommentar}</div>
                  <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
                    {bew.staerken?.map((s, si) => <span key={si} style={{ fontSize: 10, color: "#4ade80" }}>✓ {s}</span>)}
                    {bew.schwaechen?.map((s, si) => <span key={si} style={{ fontSize: 10, color: "#f87171" }}>✗ {s}</span>)}
                  </div>
                  {bew.massnahme && <div style={{ fontSize: 11, color: "#fbbf24", marginTop: 6 }}>→ {bew.massnahme}</div>}
                </div>
              )}
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                {!bew && !isLoadingThis && (
                  <button onClick={() => bewerteEinzelTechniker(tech)}
                    style={{ background: "#1d4ed8", color: "#fff", border: "none", borderRadius: 5, cursor: "pointer", fontSize: 11, padding: "5px 12px", fontWeight: 600 }}>🤖 Bewerten</button>
                )}
                {bew && k.email && (
                  <a href={mailto} style={{ background: "#1d4ed8", color: "#fff", padding: "5px 12px", borderRadius: 5, fontSize: 11, textDecoration: "none", fontWeight: 600 }}>📧 Mail senden</a>
                )}
                {bew && !k.email && <span style={{ fontSize: 10, color: "#6b7280" }}>⚠ Keine Email — unter 👥 eintragen</span>}
                {k.mobil && bew && (
                  <a href={`https://wa.me/${k.mobil.replace(/[^0-9]/g, "")}?text=${encodeURIComponent(`Hallo ${tech.name.split(" ")[0]}, Ihr KPI-Score: ${score}/10 — ${scoreLabel(score)}. ${bew?.massnahme || ""}`)}`}
                    target="_blank" rel="noreferrer"
                    style={{ background: "#15803d", color: "#fff", padding: "5px 12px", borderRadius: 5, fontSize: 11, textDecoration: "none", fontWeight: 600 }}>💬 WhatsApp</a>
                )}
              </div>
            </div>
          );
        })}
      </div>
    );
  };

  return (
    <div style={{ background: "#0a0e1a", minHeight: "100vh", fontFamily: "system-ui, sans-serif", color: "#e5e7eb" }}>
      {showKontakte && <KontakteEditor kontakte={kontakte} onSave={setKontakte} onClose={() => setShowKontakte(false)} />}
      {showBaseline && <BaselineEditor baselines={baselines} onSave={setBaselines} onClose={() => setShowBaseline(false)} />}
      {showTechVerwaltung && <TechnikerVerwaltung gespeichert={gespeichert} onUpdate={setGespeichert} onClose={() => setShowTechVerwaltung(false)} />}
      {showArchiv && <ArchivPanel archiv={archiv} onDelete={(idx) => setArchiv(prev => prev.filter((_, i) => i !== idx))} onClose={() => setShowArchiv(false)} />}

      <div style={{ borderBottom: "1px solid #1f2937", padding: "12px 24px", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 8 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
          <div style={{ width: 8, height: 8, borderRadius: "50%", background: "#4ade80", boxShadow: "0 0 6px #4ade80", flexShrink: 0 }} />
          <button onClick={() => window.location.reload()} style={{ background: "none", border: "none", cursor: "pointer", padding: 0 }}>
            <span style={{ fontWeight: 700, fontSize: 13, letterSpacing: 2, color: "#9ca3af", textTransform: "uppercase" }}>KPI Agent ↻</span>
          </button>
          <span style={{ color: "#374151" }}>·</span>
          <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
            {KATEGORIEN.map(k => {
              const anzahl = k.id === "alle"
                ? Object.values(gespeichert).flat().filter((t, idx, arr) => arr.findLastIndex(x => x.name === t.name) === idx).length
                : (gespeichert[k.id] || []).length;
              const aktiv = aktiveKategorie === k.id;
              const hatDatenInKat = k.id === "alle" ? hatDaten : anzahl > 0;
              return (
                <button key={k.id} onClick={() => setAktiveKategorie(k.id)} style={{
                  background: aktiv ? "#1d4ed8" : hatDatenInKat ? "#111827" : "transparent",
                  color: aktiv ? "#fff" : hatDatenInKat ? "#60a5fa" : "#374151",
                  border: `1px solid ${aktiv ? "#1d4ed8" : hatDatenInKat ? "#1e3a5f" : "#1f2937"}`,
                  padding: "3px 10px", borderRadius: 5, cursor: "pointer", fontSize: 11, fontWeight: aktiv ? 700 : 400,
                }}>
                  {k.label}{hatDatenInKat && anzahl > 0 ? ` (${anzahl})` : ""}
                </button>
              );
            })}
          </div>
        </div>
        <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" }}>
          <button onClick={() => setShowKontakte(true)} style={{ background: "#111827", color: "#9ca3af", border: "1px solid #374151", padding: "6px 12px", borderRadius: 6, cursor: "pointer", fontSize: 11 }}>👥 Kontakte</button>
          <button onClick={() => setShowBaseline(true)} style={{ background: "#111827", color: "#9ca3af", border: "1px solid #374151", padding: "6px 12px", borderRadius: 6, cursor: "pointer", fontSize: 11 }}>📊 Baselines</button>
          <button onClick={() => setShowTechVerwaltung(true)} style={{ background: "#111827", color: "#9ca3af", border: "1px solid #374151", padding: "6px 12px", borderRadius: 6, cursor: "pointer", fontSize: 11 }}>🧑‍🔧 Techniker</button>
          <button onClick={() => setShowArchiv(true)} style={{ background: "#111827", color: "#9ca3af", border: "1px solid #374151", padding: "6px 12px", borderRadius: 6, cursor: "pointer", fontSize: 11 }}>🗂 Archiv{archiv.length > 0 ? ` (${archiv.length})` : ""}</button>
          <label style={{ background: loading ? "#1a2e1a" : "#1f2937", color: loading ? "#4ade80" : "#9ca3af", border: `1px solid ${loading ? "#14532d" : "#374151"}`, padding: "6px 12px", borderRadius: 6, cursor: "pointer", fontSize: 11 }}>
            {loading ? "⏳ Nächste" : "📂 Upload"}
            <input type="file" accept=".csv,.xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
          </label>
          {angezeigt.length > 0 && <button onClick={exportPDF} disabled={exporting} style={{ background: "#1f2937", color: "#9ca3af", border: "1px solid #374151", padding: "6px 12px", borderRadius: 6, cursor: "pointer", fontSize: 11 }}>📄 PDF</button>}
          {aktiveKategorie !== "alle" && gespeichert[aktiveKategorie] && <button onClick={() => loescheKategorie(aktiveKategorie)} style={{ background: "#2e0f0f", color: "#f87171", border: "1px solid #7f1d1d", padding: "6px 10px", borderRadius: 6, cursor: "pointer", fontSize: 11 }}>✕</button>}
          <button onClick={async () => { await fetch("/api/logout", { method: "POST" }); window.location.href = "/login"; }}
            style={{ background: "#2e0f0f", color: "#f87171", border: "1px solid #7f1d1d", padding: "6px 12px", borderRadius: 6, cursor: "pointer", fontSize: 11 }}>🔒 Logout</button>
        </div>
      </div>

      <div ref={dashboardRef} style={{ maxWidth: 720, margin: "0 auto", padding: "24px 20px" }}>
        {!hatDaten && (
          <div style={{ textAlign: "center", padding: "60px 20px" }}>
            <div style={{ fontSize: 36, marginBottom: 12 }}>📁</div>
            <div style={{ fontSize: 18, fontWeight: 700, color: "#f9fafb", marginBottom: 8 }}>Telekom-Export hochladen</div>
            <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 24 }}>Wird automatisch der richtigen Kategorie zugeordnet</div>
            <label style={{ display: "inline-block", background: "#1d4ed8", color: "#fff", padding: "10px 24px", borderRadius: 6, cursor: "pointer", fontSize: 13, fontWeight: 600 }}>
              📂 Datei wählen (.csv / .xlsx)
              <input type="file" accept=".csv,.xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
            </label>
            {error ? <div style={{ marginTop: 16, color: "#4ade80", fontSize: 13 }}>{error}</div> : null}
          </div>
        )}

        {hatDaten && angezeigt.length === 0 && (
          <div style={{ textAlign: "center", padding: "40px 20px" }}>
            <div style={{ fontSize: 14, color: "#6b7280", marginBottom: 16 }}>Noch keine Daten für "{KATEGORIEN.find(k => k.id === aktiveKategorie)?.label}"</div>
            <label style={{ display: "inline-block", background: "#1d4ed8", color: "#fff", padding: "10px 24px", borderRadius: 6, cursor: "pointer", fontSize: 13, fontWeight: 600 }}>
              📂 Export hochladen
              <input type="file" accept=".csv,.xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
            </label>
          </div>
        )}

        {angezeigt.length > 0 && (
          <>
            {pending && <div style={{ fontSize: 11, color: "#4ade80", marginBottom: 8 }}>⏳ Nächste Datei bereit</div>}
            {error && <div style={{ fontSize: 11, color: "#4ade80", marginBottom: 8 }}>{error}</div>}

            <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 10, marginBottom: 20 }}>
              {(isOTView ? [
                { label: "Techniker", value: angezeigt.length, color: "#60a5fa" },
                { label: "Kritisch", value: criticalCount, color: criticalCount > 0 ? "#f87171" : "#4ade80" },
                { label: "Ø A1-Rate", value: avg("a1") !== "—" ? `${avg("a1")}%` : "—", color: "#4ade80" },
                { label: "Ø Score", value: `${teamAvgScore()}/10`, color: scoreColor(parseFloat(teamAvgScore())) },
              ] : [
                { label: "Techniker", value: angezeigt.length, color: "#60a5fa" },
                { label: "Kritisch", value: criticalCount, color: criticalCount > 0 ? "#f87171" : "#4ade80" },
                { label: "Ø CC-Rate", value: avg("cc_rate") !== "—" ? `${avg("cc_rate")}%` : "—", color: "#fbbf24" },
                { label: "Ø Score", value: `${teamAvgScore()}/10`, color: scoreColor(parseFloat(teamAvgScore())) },
              ]).map(s => (
                <div key={s.label} style={{ background: "#111827", border: "1px solid #1f2937", borderRadius: 8, padding: "12px 14px" }}>
                  <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 4 }}>{s.label}</div>
                  <div style={{ fontSize: 20, fontWeight: 700, color: s.color, fontFamily: "monospace" }}>{s.value}</div>
                </div>
              ))}
            </div>

            <div style={{ display: "flex", marginBottom: 16, borderBottom: "1px solid #1f2937" }}>
              {[
                { id: "dashboard", label: "Dashboard" },
                { id: "firmendashboard", label: `🏢 Firmendashboard${Object.keys(techBewertungen).length > 0 ? ` (${Object.keys(techBewertungen).length})` : ""}` },
                { id: "analyse", label: "KI-Analyse" + (aiAnalysis ? " ✓" : "") },
              ].map(tab => (
                <button key={tab.id} onClick={() => setActiveTab(tab.id)}
                  style={{ background: "none", border: "none", borderBottom: activeTab === tab.id ? "2px solid #3b82f6" : "2px solid transparent", color: activeTab === tab.id ? "#f9fafb" : "#6b7280", padding: "8px 14px", cursor: "pointer", fontSize: 12, fontWeight: activeTab === tab.id ? 600 : 400, marginBottom: -1, whiteSpace: "nowrap" }}>
                  {tab.label}
                </button>
              ))}
            </div>

            {activeTab === "dashboard" && (
              <>
                <div style={{ marginBottom: 16 }}>{angezeigt.map((t, i) => <TechCard key={i} tech={t} baselines={baselines} />)}</div>
                <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                  <button onClick={runAnalysis} disabled={loading}
                    style={{ width: "100%", background: loading ? "#1f2937" : "#1d4ed8", color: loading ? "#6b7280" : "#fff", border: "none", borderRadius: 8, padding: "14px", fontSize: 14, fontWeight: 700, cursor: loading ? "not-allowed" : "pointer" }}>
                    {loading ? "⏳ KI analysiert..." : "🤖 Team-Analyse starten"}
                  </button>
                  <button onClick={bewerteAlle}
                    style={{ width: "100%", background: "#0f172a", color: "#60a5fa", border: "1px solid #1e3a5f", borderRadius: 8, padding: "12px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                    🏢 Alle bewerten → Firmendashboard
                  </button>
                  <button onClick={archivieren}
                    style={{ width: "100%", background: "#0f172a", color: "#6b7280", border: "1px solid #1f2937", borderRadius: 8, padding: "10px", fontSize: 12, cursor: "pointer" }}>
                    🗂 Archivieren & Dashboard leeren
                  </button>
                </div>
              </>
            )}

            {activeTab === "firmendashboard" && <FirmendashboardTab />}

            {activeTab === "analyse" && (
              <div>
                {!aiAnalysis && !loading && <div style={{ textAlign: "center", padding: "40px", color: "#6b7280" }}>Noch keine Analyse. Dashboard öffnen und starten.</div>}
                {loading && <div style={{ textAlign: "center", padding: "40px", color: "#6b7280" }}>⏳ KI analysiert...</div>}
                {aiAnalysis && (
                  <>
                    <div style={{ background: "#111827", border: "1px solid #1f2937", borderRadius: 8, padding: "20px", fontSize: 13, lineHeight: 1.8, color: "#d1d5db" }}
                      dangerouslySetInnerHTML={{ __html: renderMarkdown(aiAnalysis) }} />
                    <MassnahmenPanel massnahmen={massnahmen} parseError={massnahmenFehler} kontakte={kontakte} />
                  </>
                )}
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}
