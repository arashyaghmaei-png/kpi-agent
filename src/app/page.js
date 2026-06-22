'use client';

import { useState, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import html2canvas from "html2canvas";

const BASELINE = { cc_rate: 97.6, termintreue: 97.7, loesungsquote: 96.0, nps: 69.9 };
const BASELINE_FS5335 = { cc_rate: 99.6, termintreue: 99.1, loesungsquote: 96.9, nps: 74.4 };
const BASELINE_FS5336 = { cc_rate: 95.7, termintreue: 96.7, loesungsquote: 97.2, nps: 66.7 };
const OT_BASELINE = { a_ges: 95.0, a1: 60.0 };
const STORAGE_KEY = "fibernc_kpi_sessions";

const SYSTEM_PROMPT = `Du bist ein operativer KPI-Analyseagent für ein Telekommunikations-Subunternehmen (Telekom-Subunternehmer, Kupfer & FTTH, Bergheim NRW).
Baseline KW13-19: CC=${BASELINE.cc_rate}% | Termintreue=${BASELINE.termintreue}% | Lösungsquote=${BASELINE.loesungsquote}%
FS5335: CC=${BASELINE_FS5335.cc_rate}% | Termintreue=${BASELINE_FS5335.termintreue}% | Lösungsquote=${BASELINE_FS5335.loesungsquote}%
FS5336: CC=${BASELINE_FS5336.cc_rate}% | Termintreue=${BASELINE_FS5336.termintreue}% | Lösungsquote=${BASELINE_FS5336.loesungsquote}%
KW20 schlechteste Woche (NPS 26, Termintreue 85,7%). KW23-24 FS5336 kritisch: CC 70%, SearchCall 37%.
OneTouch: A1=erster Besuch erledigt (Ziel >=60%), AX=Abbruch, A0=nicht erledigt (kritisch >10%).
Aufgabe: Techniker-KPIs bewerten, Frühwarnungen bei >=7% Abweichung, Leitstellen-Empfehlungen.
Antworte auf Deutsch, direkt und operativ.

## KPI-Übersicht
[Techniker, Wert, Baseline-Delta, Status]

## Frühwarnungen
[Nur kritische Fälle mit Name und Problem]

## Baseline-Vergleich
[Teamdurchschnitt vs KW13-19]

## Empfehlungen Leitstelle
[3-5 konkrete Maßnahmen für heute]`;

const EXAMPLE_CSV = `name,standort,cc_rate,termintreue,loesungsquote,auftraege
Mehmet K.,5335,91,94,88,5
Tobias R.,5335,78,82,75,4
Sven L.,5336,65,70,68,3
Igor P.,5336,88,90,85,5
Kai B.,5335,82,85,79,4`;

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
  return null;
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
      if (fmt === "smsfeedback") return { name, standort: String(get(row, "od") || "5335"), cc_rate: parsePercent(get(row, "cc")), termintreue: parsePercent(get(row, "termintreue")), loesungsquote: null, nps: parsePercent(get(row, "nps pb", "nps bs")), auftraege: get(row, "anzahl") || "—", quelle: "smsfeedback" };
      if (fmt === "smsfeedbackschalten") return { name, standort: "5335", cc_rate: parsePercent(get(row, "courtesy call")), termintreue: parsePercent(get(row, "termintreue mit st vo", "termintreue ohne st vo")), loesungsquote: null, nps: parsePercent(get(row, "nps")), auftraege: get(row, "anzahl") || "—", quelle: "smsfeedbackschalten" };
      if (fmt === "nftq") return { name, standort: "5335", cc_rate: null, termintreue: null, loesungsquote: null, nftq_b: parsePercent(get(row, "nftq b")), nftq_s: parsePercent(get(row, "nftq s")), nftq_m: parsePercent(get(row, "nftq m")), nftq_p: parsePercent(get(row, "nftq p")), auftraege: get(row, "anzahl") || "—", quelle: "nftq" };
      return { name, standort: String(get(row, "standort") || "5335"), cc_rate: parsePercent(get(row, "cc_rate")), termintreue: parsePercent(get(row, "termintreue")), loesungsquote: parsePercent(get(row, "loesungsquote")), nps: parsePercent(get(row, "nps")), auftraege: get(row, "auftraege") || "—", quelle: "standard" };
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

const STATUS_STYLE = {
  gut:      { bg: "#0f2e1a", color: "#4ade80", label: "GUT" },
  warnung:  { bg: "#2e1f00", color: "#fbbf24", label: "WARNUNG" },
  kritisch: { bg: "#2e0f0f", color: "#f87171", label: "KRITISCH" },
  unbekannt:{ bg: "#1a1a2e", color: "#6b7280", label: "—" },
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
  const a1 = tech.a1 || 0;
  const a2 = tech.a2 || 0;
  const a2plus = tech.a2plus || 0;
  const ax = tech.ax || 0;
  const a0 = tech.a0 || 0;
  const segments = [
    { key: "A1", val: a1, color: "#4ade80" },
    { key: "A2", val: a2, color: "#60a5fa" },
    { key: "A2+", val: a2plus, color: "#818cf8" },
    { key: "AX", val: ax, color: "#fbbf24" },
    { key: "A0", val: a0, color: "#f87171" },
  ];
  return (
    <div style={{ marginBottom: 10 }}>
      <div style={{ fontSize: 11, color: "#9ca3af", marginBottom: 4 }}>Auftragsverteilung</div>
      <div style={{ display: "flex", height: 10, borderRadius: 3, overflow: "hidden", background: "#1f2937" }}>
        {segments.map(s => s.val > 0 ? (
          <div key={s.key} style={{ width: `${s.val}%`, background: s.color }} title={`${s.key}: ${s.val.toFixed(1)}%`} />
        ) : null)}
      </div>
      <div style={{ display: "flex", gap: 10, marginTop: 5, flexWrap: "wrap" }}>
        {segments.map(s => (
          <span key={s.key} style={{ fontSize: 10, color: s.val > 0 ? s.color : "#374151" }}>
            {s.key} {s.val.toFixed(0)}%
          </span>
        ))}
      </div>
    </div>
  );
}

function TechCard({ tech }) {
  const isNFTQ = tech.quelle === "nftq";
  const isOT = tech.quelle === "onetouch";
  const bl = String(tech.standort) === "5336" ? BASELINE_FS5336 : BASELINE_FS5335;
  let worst = "gut";
  if (isOT) {
    worst = getOTStatus(tech);
  } else if (isNFTQ) {
    const vals = [tech.nftq_b, tech.nftq_s, tech.nftq_m, tech.nftq_p].filter(v => v !== null);
    worst = vals.some(v => v > 10) ? "kritisch" : vals.some(v => v > 5) ? "warnung" : "gut";
  } else {
    const statuses = [
      tech.cc_rate !== null ? getStatus(tech.cc_rate, bl.cc_rate) : null,
      tech.termintreue !== null ? getStatus(tech.termintreue, bl.termintreue) : null,
      tech.loesungsquote !== null ? getStatus(tech.loesungsquote, bl.loesungsquote) : null,
    ].filter(Boolean);
    worst = statuses.includes("kritisch") ? "kritisch" : statuses.includes("warnung") ? "warnung" : "gut";
  }
  const borderColor = worst === "kritisch" ? "#7f1d1d" : worst === "warnung" ? "#78350f" : "#14532d";
  const quelleLabel = { smsfeedback: "SMS-Feedback", smsfeedbackschalten: "Schalten", nftq: "NFTQ", standard: "Manuell", onetouch: "OneTouch" }[tech.quelle] || "";
  return (
    <div style={{ background: "#111827", border: `1px solid ${borderColor}`, borderRadius: 8, padding: "16px 18px", marginBottom: 12 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
        <div>
          <div style={{ fontWeight: 700, fontSize: 15, color: "#f9fafb" }}>{tech.name}</div>
          <div style={{ fontSize: 11, color: "#6b7280", marginTop: 2 }}>
            FS{tech.standort} · {tech.auftraege} Aufträge
            {isOT && tech.tage ? <span style={{ marginLeft: 6 }}>· {tech.tage} Tage</span> : null}
            <span style={{ marginLeft: 8, color: "#374151", background: "#1f2937", padding: "1px 6px", borderRadius: 3 }}>{quelleLabel}</span>
          </div>
        </div>
        <StatusBadge status={worst} />
      </div>
      {isOT && (
        <>
          <OTStackedBar tech={tech} />
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
            <KPIBar value={tech.a_ges} baseline={OT_BASELINE.a_ges} label="Gesamterfolg (A Ges.)" />
            <KPIBar value={tech.a1} baseline={OT_BASELINE.a1} label="Erstlösung (A1)" />
          </div>
          {tech.a0 > 0 ? <div style={{ marginTop: 6, fontSize: 11, color: "#f87171" }}>⚠ A0 (nicht erledigt): {tech.a0.toFixed(1)}%</div> : null}
        </>
      )}
      {isNFTQ && (
        <>
          <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 8 }}>NFTQ Fehlerquoten (niedriger = besser)</div>
          <NFTQBar value={tech.nftq_b} label="NFTQ Bereitstellung" />
          <NFTQBar value={tech.nftq_s} label="NFTQ Schalten" />
          <NFTQBar value={tech.nftq_m} label="NFTQ Montage" />
          <NFTQBar value={tech.nftq_p} label="NFTQ Problembehebung" />
        </>
      )}
      {!isOT && !isNFTQ && (
        <>
          <KPIBar value={tech.cc_rate} baseline={bl.cc_rate} label="CC-Rate" />
          <KPIBar value={tech.termintreue} baseline={bl.termintreue} label="Termintreue" />
          <KPIBar value={tech.loesungsquote} baseline={bl.loesungsquote} label="Lösungsquote" />
          {tech.nps !== null ? (
            <div style={{ marginTop: 8, fontSize: 11, color: "#9ca3af" }}>
              NPS: <span style={{ color: tech.nps >= 50 ? "#4ade80" : tech.nps >= 0 ? "#fbbf24" : "#f87171", fontWeight: 700 }}>{tech.nps.toFixed(0)}</span>
            </div>
          ) : null}
        </>
      )}
    </div>
  );
}

function renderMarkdown(text) {
  return text
    .replace(/## (.*)/g, '<h3 style="color:#f9fafb;margin:20px 0 8px;font-size:14px">$1</h3>')
    .replace(/\*\*(.*?)\*\*/g, '<strong style="color:#e5e7eb">$1</strong>')
    .replace(/\n/g, "<br/>");
}

function FormatBadge({ format }) {
  const labels = { smsfeedback: "SMS-Feedback", smsfeedbackschalten: "Schalten", nftq: "NFTQ", onetouch: "OneTouch", standard: "Standard CSV", mixed: "Gemischt" };
  return <span style={{ fontSize: 11, color: "#60a5fa", background: "#0d1f3c", border: "1px solid #1e3a5f", padding: "2px 8px", borderRadius: 3, fontFamily: "monospace" }}>{labels[format] || format}</span>;
}

export default function KPIAgent() {
  const [techniker, setTechniker] = useState([]);
  const [aiAnalysis, setAiAnalysis] = useState("");
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [error, setError] = useState("");
  const [activeTab, setActiveTab] = useState("dashboard");
  const [fileName, setFileName] = useState("");
  const [detectedFormat, setDetectedFormat] = useState(null);
  const [sessions, setSessions] = useState([]);
  const [showSessions, setShowSessions] = useState(false);
  const dashboardRef = useRef(null);

  // Beim Start gespeicherte Sessions laden
  useEffect(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) setSessions(JSON.parse(saved));
    } catch(e) {}
  }, []);

  // Sessions speichern wenn sie sich ändern
  useEffect(() => {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(sessions));
    } catch(e) {}
  }, [sessions]);

  const handleRows = useCallback((rows, name, skipSave) => {
    if (!rows.length) { setError("Keine verwertbaren Daten gefunden."); return; }
    const formats = [...new Set(rows.map(r => r.quelle))];
    const fmt = formats.length === 1 ? formats[0] : "mixed";
    setDetectedFormat(fmt);
    setTechniker(rows);
    setAiAnalysis("");
    setFileName(name);
    setError("");
    setActiveTab("dashboard");

    // Session speichern
    if (!skipSave) {
      const session = {
        id: Date.now(),
        name,
        format: fmt,
        datum: new Date().toLocaleDateString("de-DE"),
        uhrzeit: new Date().toLocaleTimeString("de-DE", { hour: "2-digit", minute: "2-digit" }),
        techniker: rows,
      };
      setSessions(prev => [session, ...prev.slice(0, 19)]); // max 20 Sessions
    }
  }, []);

  const loadSession = (session) => {
    handleRows(session.techniker, session.name, true);
    setDetectedFormat(session.format);
    setShowSessions(false);
  };

  const deleteSession = (id, e) => {
    e.stopPropagation();
    setSessions(prev => prev.filter(s => s.id !== id));
  };

  const processXLSX = useCallback(async (file) => {
    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rawRows = XLSX.utils.sheet_to_json(ws, { defval: "" });
      handleRows(normalizeRows(rawRows), file.name);
    } catch (e) { setError("Fehler beim Lesen der Excel-Datei: " + e.message); }
  }, [handleRows]);

  const handleFile = useCallback((e) => {
    const file = e.target.files[0];
    if (!file) return;
    if (file.name.match(/\.xlsx?$/i)) { processXLSX(file); }
    else {
      const reader = new FileReader();
      reader.onload = (ev) => handleRows(parseCSV(ev.target.result), file.name);
      reader.readAsText(file, "utf-8");
    }
  }, [processXLSX, handleRows]);

  const loadExample = () => { handleRows(parseCSV(EXAMPLE_CSV), "Beispieldaten"); setDetectedFormat("standard"); };

  const runAnalysis = async () => {
    if (!techniker.length) return;
    setLoading(true); setError(""); setAiAnalysis("");
    const dataStr = techniker.map(t => {
      if (t.quelle === "onetouch") return `${t.name}: A-Ges=${t.a_ges?.toFixed(1) ?? "—"}%, A1=${t.a1?.toFixed(1) ?? "—"}%, AX=${t.ax?.toFixed(1) ?? "—"}%, A0=${t.a0?.toFixed(1) ?? "—"}%, Aufträge=${t.auftraege}`;
      if (t.quelle === "nftq") return `${t.name}: NFTQ-B=${t.nftq_b?.toFixed(2) ?? "—"}%, NFTQ-S=${t.nftq_s?.toFixed(2) ?? "—"}%, NFTQ-M=${t.nftq_m?.toFixed(2) ?? "—"}%, NFTQ-P=${t.nftq_p?.toFixed(2) ?? "—"}%, Aufträge=${t.auftraege}`;
      return `${t.name} (FS${t.standort}): CC=${t.cc_rate?.toFixed(1) ?? "—"}%, Termintreue=${t.termintreue?.toFixed(1) ?? "—"}%, Lösungsquote=${t.loesungsquote?.toFixed(1) ?? "—"}%, NPS=${t.nps?.toFixed(0) ?? "—"}, Aufträge=${t.auftraege}`;
    }).join("\n");
    try {
      const res = await fetch("/api/analyse", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ model: "claude-sonnet-4-6", max_tokens: 1500, system: SYSTEM_PROMPT, messages: [{ role: "user", content: `Analysiere diese Techniker-KPIs:\n\n${dataStr}` }] }),
      });
      const data = await res.json();
      setAiAnalysis(data.content?.map(b => b.text || "").join("") || "Keine Antwort.");
      setActiveTab("analyse");
    } catch (e) { setError("Fehler bei der KI-Analyse."); }
    finally { setLoading(false); }
  };

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

  const criticalCount = techniker.filter(t => {
    if (t.quelle === "onetouch") return getOTStatus(t) === "kritisch";
    if (t.quelle === "nftq") return [t.nftq_b, t.nftq_s, t.nftq_m, t.nftq_p].filter(Boolean).some(v => v > 10);
    const bl = String(t.standort) === "5336" ? BASELINE_FS5336 : BASELINE_FS5335;
    return [
      t.cc_rate !== null ? getStatus(t.cc_rate, bl.cc_rate) : null,
      t.termintreue !== null ? getStatus(t.termintreue, bl.termintreue) : null,
    ].includes("kritisch");
  }).length;

  const avg = (key) => {
    const vals = techniker.map(t => t[key]).filter(v => v !== null && !isNaN(v));
    return vals.length ? (vals.reduce((s, v) => s + v, 0) / vals.length).toFixed(1) : "—";
  };

  const isOTView = detectedFormat === "onetouch";

  return (
    <div style={{ background: "#0a0e1a", minHeight: "100vh", fontFamily: "system-ui, sans-serif", color: "#e5e7eb" }}>
      {/* Header */}
      <div style={{ borderBottom: "1px solid #1f2937", padding: "14px 24px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 8, height: 8, borderRadius: "50%", background: "#4ade80", boxShadow: "0 0 6px #4ade80" }} />
          {/* KPI Agent Button — klickbar zum Aktualisieren */}
          <button onClick={() => window.location.reload()} style={{ background: "none", border: "none", cursor: "pointer", padding: 0, display: "flex", alignItems: "center", gap: 6 }}>
            <span style={{ fontWeight: 700, fontSize: 13, letterSpacing: 2, color: "#9ca3af", textTransform: "uppercase" }}>KPI Agent</span>
            <span style={{ fontSize: 10, color: "#374151" }}>↻</span>
          </button>
          <span style={{ color: "#374151" }}>·</span>
          <span style={{ fontSize: 12, color: "#6b7280" }}>Techniker-Kontrolle FiberNC</span>
          {detectedFormat ? <FormatBadge format={detectedFormat} /> : null}
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          {/* Sessions Button */}
          <div style={{ position: "relative" }}>
            <button onClick={() => setShowSessions(!showSessions)}
              style={{ background: "#1f2937", color: sessions.length > 0 ? "#60a5fa" : "#6b7280", border: "1px solid #374151", padding: "6px 14px", borderRadius: 6, cursor: "pointer", fontSize: 12 }}>
              📂 Verlauf {sessions.length > 0 ? `(${sessions.length})` : ""}
            </button>
            {showSessions && (
              <div style={{ position: "absolute", right: 0, top: 36, width: 320, background: "#111827", border: "1px solid #1f2937", borderRadius: 8, zIndex: 100, maxHeight: 400, overflowY: "auto" }}>
                <div style={{ padding: "12px 16px", borderBottom: "1px solid #1f2937", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <span style={{ fontSize: 12, fontWeight: 700, color: "#f9fafb" }}>Gespeicherte Auswertungen</span>
                  {sessions.length > 0 ? <button onClick={() => { setSessions([]); setShowSessions(false); }} style={{ background: "none", border: "none", color: "#6b7280", cursor: "pointer", fontSize: 11 }}>Alle löschen</button> : null}
                </div>
                {sessions.length === 0 ? (
                  <div style={{ padding: 20, textAlign: "center", color: "#6b7280", fontSize: 12 }}>Noch keine Auswertungen gespeichert</div>
                ) : sessions.map(s => (
                  <div key={s.id} onClick={() => loadSession(s)}
                    style={{ padding: "10px 16px", borderBottom: "1px solid #1f2937", cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center" }}
                    onMouseEnter={e => e.currentTarget.style.background = "#1f2937"}
                    onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                    <div>
                      <div style={{ fontSize: 12, color: "#f9fafb", fontWeight: 600 }}>{s.name}</div>
                      <div style={{ fontSize: 10, color: "#6b7280", marginTop: 2 }}>{s.datum} {s.uhrzeit} · {s.techniker.length} Techniker · <FormatBadge format={s.format} /></div>
                    </div>
                    <button onClick={(e) => deleteSession(s.id, e)} style={{ background: "none", border: "none", color: "#4b5563", cursor: "pointer", fontSize: 14, padding: "0 4px" }}>✕</button>
                  </div>
                ))}
              </div>
            )}
          </div>
          {techniker.length > 0 ? <button onClick={exportPDF} disabled={exporting} style={{ background: "#1f2937", color: "#9ca3af", border: "1px solid #374151", padding: "6px 14px", borderRadius: 6, cursor: "pointer", fontSize: 12 }}>📄 PDF</button> : null}
          <span style={{ fontSize: 11, color: "#4b5563", fontFamily: "monospace" }}>Baseline KW13–19</span>
        </div>
      </div>

      <div ref={dashboardRef} style={{ maxWidth: 720, margin: "0 auto", padding: "24px 20px" }}>
        {!techniker.length && (
          <div style={{ textAlign: "center", padding: "60px 20px" }}>
            <div style={{ fontSize: 36, marginBottom: 12 }}>📁</div>
            <div style={{ fontSize: 18, fontWeight: 700, color: "#f9fafb", marginBottom: 8 }}>Telekom-Export hochladen</div>
            <div style={{ display: "flex", justifyContent: "center", gap: 8, marginBottom: 24, flexWrap: "wrap" }}>
              {["SMS-Feedback", "Schalten", "NFTQ", "OneTouch", "Excel (.xlsx)", "Standard CSV"].map(f => (
                <span key={f} style={{ fontSize: 11, color: "#60a5fa", background: "#0d1f3c", border: "1px solid #1e3a5f", padding: "3px 10px", borderRadius: 3, fontFamily: "monospace" }}>{f}</span>
              ))}
            </div>
            <label style={{ display: "inline-block", background: "#1d4ed8", color: "#fff", padding: "10px 24px", borderRadius: 6, cursor: "pointer", fontSize: 13, fontWeight: 600, marginRight: 12 }}>
              📂 Datei wählen
              <input type="file" accept=".csv,.xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
            </label>
            <button onClick={loadExample} style={{ background: "#1f2937", color: "#9ca3af", border: "1px solid #374151", padding: "10px 24px", borderRadius: 6, cursor: "pointer", fontSize: 13, fontWeight: 600 }}>Beispieldaten</button>
            {error && <div style={{ marginTop: 16, color: "#f87171", fontSize: 13 }}>{error}</div>}
            {sessions.length > 0 && (
              <div style={{ marginTop: 32, textAlign: "left", maxWidth: 400, margin: "32px auto 0" }}>
                <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 12, textAlign: "center" }}>— oder letzte Auswertung laden —</div>
                {sessions.slice(0, 3).map(s => (
                  <div key={s.id} onClick={() => loadSession(s)}
                    style={{ background: "#111827", border: "1px solid #1f2937", borderRadius: 8, padding: "12px 16px", marginBottom: 8, cursor: "pointer" }}
                    onMouseEnter={e => e.currentTarget.style.borderColor = "#3b82f6"}
                    onMouseLeave={e => e.currentTarget.style.borderColor = "#1f2937"}>
                    <div style={{ fontSize: 13, color: "#f9fafb", fontWeight: 600 }}>{s.name}</div>
                    <div style={{ fontSize: 11, color: "#6b7280", marginTop: 2 }}>{s.datum} {s.uhrzeit} · {s.techniker.length} Techniker</div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        {techniker.length > 0 && (
          <>
            {fileName ? <div style={{ fontSize: 11, color: "#4b5563", marginBottom: 12, fontFamily: "monospace" }}>📂 {fileName} · {techniker.length} Techniker</div> : null}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 10, marginBottom: 20 }}>
              {(isOTView ? [
                { label: "Techniker", value: techniker.length, color: "#60a5fa" },
                { label: "Kritisch", value: criticalCount, color: criticalCount > 0 ? "#f87171" : "#4ade80" },
                { label: "Ø A1-Rate", value: avg("a1") !== "—" ? `${avg("a1")}%` : "—", color: "#4ade80" },
                { label: "Ø A-Gesamt", value: avg("a_ges") !== "—" ? `${avg("a_ges")}%` : "—", color: "#fbbf24" },
              ] : [
                { label: "Techniker", value: techniker.length, color: "#60a5fa" },
                { label: "Kritisch", value: criticalCount, color: criticalCount > 0 ? "#f87171" : "#4ade80" },
                { label: "Ø CC-Rate", value: avg("cc_rate") !== "—" ? `${avg("cc_rate")}%` : "—", color: "#fbbf24" },
                { label: "Ø Termintreue", value: avg("termintreue") !== "—" ? `${avg("termintreue")}%` : "—", color: "#fbbf24" },
              ]).map(s => (
                <div key={s.label} style={{ background: "#111827", border: "1px solid #1f2937", borderRadius: 8, padding: "12px 14px" }}>
                  <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 4 }}>{s.label}</div>
                  <div style={{ fontSize: 20, fontWeight: 700, color: s.color, fontFamily: "monospace" }}>{s.value}</div>
                </div>
              ))}
            </div>

            <div style={{ display: "flex", marginBottom: 16, borderBottom: "1px solid #1f2937" }}>
              {[{ id: "dashboard", label: "Dashboard" }, { id: "analyse", label: "KI-Analyse" + (aiAnalysis ? " ✓" : "") }].map(tab => (
                <button key={tab.id} onClick={() => setActiveTab(tab.id)}
                  style={{ background: "none", border: "none", borderBottom: activeTab === tab.id ? "2px solid #3b82f6" : "2px solid transparent", color: activeTab === tab.id ? "#f9fafb" : "#6b7280", padding: "8px 16px", cursor: "pointer", fontSize: 13, fontWeight: activeTab === tab.id ? 600 : 400, marginBottom: -1 }}>
                  {tab.label}
                </button>
              ))}
            </div>

            {activeTab === "dashboard" && (
              <>
                <div style={{ marginBottom: 16 }}>{techniker.map((t, i) => <TechCard key={i} tech={t} />)}</div>
                <button onClick={runAnalysis} disabled={loading} style={{ width: "100%", background: loading ? "#1f2937" : "#1d4ed8", color: loading ? "#6b7280" : "#fff", border: "none", borderRadius: 8, padding: "14px", fontSize: 14, fontWeight: 700, cursor: loading ? "not-allowed" : "pointer" }}>
                  {loading ? "KI analysiert..." : "🤖 KI-Analyse starten"}
                </button>
                <button onClick={() => { setTechniker([]); setAiAnalysis(""); setFileName(""); setError(""); setDetectedFormat(null); }}
                  style={{ width: "100%", marginTop: 8, background: "none", color: "#4b5563", border: "1px solid #1f2937", borderRadius: 8, padding: "10px", fontSize: 12, cursor: "pointer" }}>
                  Neue Datei laden
                </button>
              </>
            )}

            {activeTab === "analyse" && (
              <div>
                {!aiAnalysis && !loading ? <div style={{ textAlign: "center", padding: "40px", color: "#6b7280" }}>Noch keine Analyse. Dashboard öffnen und starten.</div> : null}
                {loading ? <div style={{ textAlign: "center", padding: "40px", color: "#6b7280" }}>KI analysiert...</div> : null}
                {aiAnalysis ? <div style={{ background: "#111827", border: "1px solid #1f2937", borderRadius: 8, padding: "20px", fontSize: 13, lineHeight: 1.8, color: "#d1d5db" }} dangerouslySetInnerHTML={{ __html: renderMarkdown(aiAnalysis) }} /> : null}
                {error && <div style={{ background: "#2e0f0f", border: "1px solid #7f1d1d", borderRadius: 8, padding: 16, color: "#f87171", fontSize: 13 }}>{error}</div>}
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}
