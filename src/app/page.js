'use client';

import { useState, useCallback, useRef } from "react";

const BASELINE = { cc_rate: 97.6, termintreue: 97.7, loesungsquote: 96.0 };
const THRESHOLDS = { kritisch: 0.85, warnung: 0.93 };

const SYSTEM_PROMPT = `Du bist ein operativer KPI-Analyseagent für ein Telekommunikations-Subunternehmen (Telekom-Subunternehmer, Kupfer & FTTH, Bergheim NRW).

Baseline KW13-19: CC=97,6% | Termintreue=97,7% | Lösungsquote=96,0%
Warnsignale: KW20 schlechteste Woche (NPS 26, Termintreue 85,7%).

Du erhältst echte Techniker-Daten aus dem Telekom Auftragsinfo-Export. Bewerte jeden Techniker, gib Frühwarnungen bei >=7% Abweichung unter Baseline, formuliere konkrete Leitstellen-Empfehlungen. Antworte auf Deutsch, direkt und operativ.

## KPI-Übersicht
[Techniker | CC | Termintreue | Infoquote | NPS | Status]

## Frühwarnungen
[Nur kritische Fälle mit Name und Problem]

## Team-Durchschnitt vs Baseline
[Delta in %]

## Empfehlungen Leitstelle
[3-5 konkrete Maßnahmen]`;

function parseCSVTelekom(text) {
  const cleaned = text.replace(/^\uFEFF/, "").trim();
  const lines = cleaned.split(/\r?\n/).filter(l => l.trim() && !l.includes("Diese Datei muss"));
  if (lines.length < 2) return [];
  const parseRow = (line) => line.split(";").map(v => v.replace(/^"|"$/g, "").trim());
  const headers = parseRow(lines[0]);
  return lines.slice(1).map(line => {
    const values = parseRow(line);
    const obj = {};
    headers.forEach((h, i) => (obj[h] = values[i] || ""));
    return obj;
  }).filter(r => r["Name"] && r["Name"].length > 0);
}

function parsePercent(val) {
  if (!val) return 0;
  return parseFloat(val.replace("%", "").replace(",", ".")) || 0;
}

function getStatusVal(value, baseline) {
  const ratio = value / baseline;
  if (ratio < THRESHOLDS.kritisch) return "kritisch";
  if (ratio < THRESHOLDS.warnung) return "warnung";
  return "gut";
}

function StatusBadge({ status }) {
  const s = {
    gut: { bg: "#0f2e1a", color: "#4ade80", label: "GUT" },
    warnung: { bg: "#2e1f00", color: "#fbbf24", label: "WARNUNG" },
    kritisch: { bg: "#2e0f0f", color: "#f87171", label: "KRITISCH" }
  }[status] || { bg: "#0f2e1a", color: "#4ade80", label: "GUT" };
  return <span style={{ background: s.bg, color: s.color, padding: "2px 10px", borderRadius: 3, fontSize: 11, fontFamily: "monospace", fontWeight: 700 }}>{s.label}</span>;
}

function KPIBar({ value, baseline, label }) {
  const color = value / baseline < THRESHOLDS.kritisch ? "#f87171" : value / baseline < THRESHOLDS.warnung ? "#fbbf24" : "#4ade80";
  return (
    <div style={{ marginBottom: 6 }}>
      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11, color: "#9ca3af", marginBottom: 2 }}>
        <span>{label}</span>
        <span style={{ color }}>{value.toFixed(1)}% <span style={{ color: "#4b5563" }}>/ {baseline}%</span></span>
      </div>
      <div style={{ background: "#1f2937", borderRadius: 2, height: 6, position: "relative" }}>
        <div style={{ width: `${Math.min(100, value)}%`, background: color, height: "100%", borderRadius: 2 }} />
        <div style={{ position: "absolute", left: `${Math.min(100, baseline)}%`, top: -3, width: 2, height: 12, background: "#6b7280" }} />
      </div>
    </div>
  );
}

function TechCard({ tech }) {
  const cc = parsePercent(tech["CC"]);
  const tt = parsePercent(tech["Termintreue"]);
  const lq = parsePercent(tech["Erledigt B"]);
  const infoquote = parsePercent(tech["Infoquote P"]);
  const nps = parseFloat((tech["NPS PB"] || "0").replace(",", ".")) || 0;
  const ccStatus = getStatusVal(cc, BASELINE.cc_rate);
  const ttStatus = getStatusVal(tt, BASELINE.termintreue);
  const lqStatus = lq > 0 ? getStatusVal(lq, BASELINE.loesungsquote) : "gut";
  const worst = [ccStatus, ttStatus, lqStatus].includes("kritisch") ? "kritisch" : [ccStatus, ttStatus, lqStatus].includes("warnung") ? "warnung" : "gut";
  const borderColor = worst === "kritisch" ? "#7f1d1d" : worst === "warnung" ? "#78350f" : "#14532d";
  return (
    <div style={{ background: "#111827", border: `1px solid ${borderColor}`, borderRadius: 8, padding: "16px 18px", marginBottom: 12 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 12 }}>
        <div>
          <div style={{ fontWeight: 700, fontSize: 14, color: "#f9fafb" }}>{tech["Name"]}</div>
          <div style={{ fontSize: 11, color: "#6b7280", marginTop: 2 }}>OD {tech["OD"]} · {tech["Anzahl"]} Aufträge · {tech["Sterne"]} Sterne</div>
        </div>
        <StatusBadge status={worst} />
      </div>
      <KPIBar value={cc} baseline={BASELINE.cc_rate} label="CC-Rate" />
      <KPIBar value={tt} baseline={BASELINE.termintreue} label="Termintreue" />
      {lq > 0 && <KPIBar value={lq} baseline={BASELINE.loesungsquote} label="Lösungsquote (B)" />}
      <div style={{ display: "flex", gap: 16, marginTop: 8, fontSize: 11, color: "#6b7280" }}>
        <span>Infoquote: <span style={{ color: infoquote >= 90 ? "#4ade80" : "#fbbf24" }}>{infoquote.toFixed(0)}%</span></span>
        <span>NPS: <span style={{ color: nps >= 50 ? "#4ade80" : nps >= 0 ? "#fbbf24" : "#f87171" }}>{isNaN(nps) ? "—" : nps.toFixed(0)}</span></span>
        <span>Geplatzt: <span style={{ color: parsePercent(tech["T. Geplatz"]) > 5 ? "#f87171" : "#4ade80" }}>{tech["T. Geplatz"]}</span></span>
      </div>
    </div>
  );
}

function renderMarkdown(text) {
  return text.replace(/## (.*)/g, '<h3 style="color:#f9fafb;margin:20px 0 8px;font-size:14px">$1</h3>').replace(/\*\*(.*?)\*\*/g, '<strong style="color:#e5e7eb">$1</strong>').replace(/\n/g, "<br/>");
}

export default function KPIAgent() {
  const [techniker, setTechniker] = useState([]);
  const [aiAnalysis, setAiAnalysis] = useState("");
  const [loading, setLoading] = useState(false);
  const [exporting, setExporting] = useState(false);
  const [error, setError] = useState("");
  const [activeTab, setActiveTab] = useState("dashboard");
  const [fileName, setFileName] = useState("");
  const dashboardRef = useRef(null);

  const handleFile = useCallback((e) => {
    const file = e.target.files[0];
    if (!file) return;
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (ev) => {
      const parsed = parseCSVTelekom(ev.target.result);
      if (parsed.length === 0) { setError("Keine Daten gefunden. Bitte Auftragsinfo-CSV hochladen."); return; }
      setTechniker(parsed); setAiAnalysis(""); setError("");
    };
    reader.readAsText(file, "UTF-8");
  }, []);

  const runAnalysis = async () => {
    if (!techniker.length) return;
    setLoading(true); setError(""); setAiAnalysis("");
    const dataStr = techniker.map(t => `${t["Name"]} (OD${t["OD"]}): CC=${t["CC"]}, Termintreue=${t["Termintreue"]}, Lösungsquote=${t["Erledigt B"]}, Infoquote=${t["Infoquote P"]}, NPS=${t["NPS PB"]}, Aufträge=${t["Anzahl"]}, Sterne=${t["Sterne"]}, Geplatzt=${t["T. Geplatz"]}`).join("\n");
    try {
      const res = await fetch("/api/analyse", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ model: "claude-sonnet-4-6", max_tokens: 1500, system: SYSTEM_PROMPT, messages: [{ role: "user", content: `Analysiere diese Techniker-KPIs (Telekom Auftragsinfo):\n\n${dataStr}` }] }),
      });
      const data = await res.json();
      setAiAnalysis(data.content?.map(b => b.text || "").join("") || "Keine Antwort.");
      setActiveTab("analyse");
    } catch (e) { setError("Fehler bei der KI-Analyse."); }
    finally { setLoading(false); }
  };

  const exportScreenshot = async () => {
    setExporting(true);
    try {
      const html2canvas = (await import('https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/+esm')).default;
      const canvas = await html2canvas(dashboardRef.current, { backgroundColor: "#0a0e1a", scale: 2 });
      const link = document.createElement("a");
      link.download = `KPI-${new Date().toLocaleDateString("de-DE").replace(/\./g, "-")}.png`;
      link.href = canvas.toDataURL("image/png"); link.click();
    } catch (e) { setError("Screenshot fehlgeschlagen."); }
    finally { setExporting(false); }
  };

  const exportPDF = async () => {
    setExporting(true);
    try {
      const html2canvas = (await import('https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/+esm')).default;
      const { jsPDF } = await import('https://cdn.jsdelivr.net/npm/jspdf@2.5.1/+esm');
      const canvas = await html2canvas(dashboardRef.current, { backgroundColor: "#0a0e1a", scale: 2 });
      const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
      const w = pdf.internal.pageSize.getWidth();
      pdf.addImage(canvas.toDataURL("image/png"), "PNG", 0, 12, w, (canvas.height * w) / canvas.width);
      pdf.save(`KPI-${new Date().toLocaleDateString("de-DE").replace(/\./g, "-")}.pdf`);
    } catch (e) { setError("PDF fehlgeschlagen."); }
    finally { setExporting(false); }
  };

  const criticalCount = techniker.filter(t => [getStatusVal(parsePercent(t["CC"]), BASELINE.cc_rate), getStatusVal(parsePercent(t["Termintreue"]), BASELINE.termintreue)].includes("kritisch")).length;
  const avgCC = techniker.length ? (techniker.reduce((s, t) => s + parsePercent(t["CC"]), 0) / techniker.length).toFixed(1) : "—";
  const avgTT = techniker.length ? (techniker.reduce((s, t) => s + parsePercent(t["Termintreue"]), 0) / techniker.length).toFixed(1) : "—";
  const avgNPS = techniker.length ? (techniker.reduce((s, t) => s + (parseFloat((t["NPS PB"] || "0").replace(",", ".")) || 0), 0) / techniker.length).toFixed(0) : "—";

  return (
    <div style={{ background: "#0a0e1a", minHeight: "100vh", fontFamily: "system-ui, sans-serif", color: "#e5e7eb" }}>
      <div style={{ borderBottom: "1px solid #1f2937", padding: "14px 24px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 8, height: 8, borderRadius: "50%", background: "#4ade80", boxShadow: "0 0 6px #4ade80" }} />
          <span style={{ fontWeight: 700, fontSize: 13, letterSpacing: 2, color: "#9ca3af", textTransform: "uppercase" }}>KPI Agent</span>
          <span style={{ color: "#374151" }}>·</span>
          <span style={{ fontSize: 12, color: "#6b7280" }}>Auftragsinfo</span>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          {techniker.length > 0 && <>
            <button onClick={exportScreenshot} disabled={exporting} style={{ background: "#1f2937", color: "#9ca3af", border: "1px solid #374151", padding: "6px 14px", borderRadius: 6, cursor: "pointer", fontSize: 12 }}>PNG</button>
            <button onClick={exportPDF} disabled={exporting} style={{ background: "#1f2937", color: "#9ca3af", border: "1px solid #374151", padding: "6px 14px", borderRadius: 6, cursor: "pointer", fontSize: 12 }}>PDF</button>
          </>}
          <span style={{ fontSize: 11, color: "#4b5563", fontFamily: "monospace" }}>Baseline KW13-19</span>
        </div>
      </div>
      <div ref={dashboardRef} style={{ maxWidth: 760, margin: "0 auto", padding: "24px 20px" }}>
        {!techniker.length && (
          <div style={{ textAlign: "center", padding: "60px 20px" }}>
            <div style={{ fontSize: 36, marginBottom: 12 }}>📁</div>
            <div style={{ fontSize: 18, fontWeight: 700, color: "#f9fafb", marginBottom: 8 }}>Auftragsinfo-Export hochladen</div>
            <div style={{ fontSize: 13, color: "#6b7280", marginBottom: 24 }}>Telekom Auftragsinfo → Auswertungen → SMS-Feedback → CSV</div>
            <label style={{ display: "inline-block", background: "#1d4ed8", color: "#fff", padding: "10px 24px", borderRadius: 6, cursor: "pointer", fontSize: 13, fontWeight: 600 }}>
              CSV hochladen
              <input type="file" accept=".csv,.xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
            </label>
            {error && <div style={{ marginTop: 16, color: "#f87171", fontSize: 13 }}>{error}</div>}
          </div>
        )}
        {techniker.length > 0 && (
          <>
            {fileName && <div style={{ fontSize: 11, color: "#4b5563", marginBottom: 12, fontFamily: "monospace" }}>{fileName} · {techniker.length} Techniker</div>}
            <div style={{ display: "grid", gridTemplateColumns: "repeat(5, 1fr)", gap: 10, marginBottom: 20 }}>
              {[
                { label: "Techniker", value: techniker.length, color: "#60a5fa" },
                { label: "Kritisch", value: criticalCount, color: criticalCount > 0 ? "#f87171" : "#4ade80" },
                { label: "Ø CC", value: `${avgCC}%`, color: parseFloat(avgCC) >= BASELINE.cc_rate ? "#4ade80" : "#fbbf24" },
                { label: "Ø Termintreue", value: `${avgTT}%`, color: parseFloat(avgTT) >= BASELINE.termintreue ? "#4ade80" : "#fbbf24" },
                { label: "Ø NPS", value: avgNPS, color: parseFloat(avgNPS) >= 50 ? "#4ade80" : "#fbbf24" },
              ].map(s => (
                <div key={s.label} style={{ background: "#111827", border: "1px solid #1f2937", borderRadius: 8, padding: "12px 14px" }}>
                  <div style={{ fontSize: 10, color: "#6b7280", marginBottom: 4 }}>{s.label}</div>
                  <div style={{ fontSize: 18, fontWeight: 700, color: s.color, fontFamily: "monospace" }}>{s.value}</div>
                </div>
              ))}
            </div>
            <div style={{ display: "flex", marginBottom: 16, borderBottom: "1px solid #1f2937" }}>
              {[{ id: "dashboard", label: "Dashboard" }, { id: "analyse", label: "KI-Analyse" + (aiAnalysis ? " ✓" : "") }].map(tab => (
                <button key={tab.id} onClick={() => setActiveTab(tab.id)} style={{ background: "none", border: "none", borderBottom: activeTab === tab.id ? "2px solid #3b82f6" : "2px solid transparent", color: activeTab === tab.id ? "#f9fafb" : "#6b7280", padding: "8px 16px", cursor: "pointer", fontSize: 13, fontWeight: activeTab === tab.id ? 600 : 400, marginBottom: -1 }}>
                  {tab.label}
                </button>
              ))}
            </div>
            {activeTab === "dashboard" && (
              <>
                <div style={{ marginBottom: 16 }}>{techniker.map((t, i) => <TechCard key={i} tech={t} />)}</div>
                <button onClick={runAnalysis} disabled={loading} style={{ width: "100%", background: loading ? "#1f2937" : "#1d4ed8", color: loading ? "#6b7280" : "#fff", border: "none", borderRadius: 8, padding: "14px", fontSize: 14, fontWeight: 700, cursor: loading ? "not-allowed" : "pointer" }}>
                  {loading ? "KI analysiert..." : "KI-Analyse starten"}
                </button>
                <button onClick={() => { setTechniker([]); setAiAnalysis(""); setFileName(""); setError(""); }} style={{ width: "100%", marginTop: 8, background: "none", color: "#4b5563", border: "1px solid #1f2937", borderRadius: 8, padding: "10px", fontSize: 12, cursor: "pointer" }}>
                  Neue Datei laden
                </button>
              </>
            )}
            {activeTab === "analyse" && (
              <div>
                {!aiAnalysis && !loading && <div style={{ textAlign: "center", padding: "40px", color: "#6b7280" }}>Noch keine Analyse. Dashboard öffnen und starten.</div>}
                {loading && <div style={{ textAlign: "center", padding: "40px", color: "#6b7280" }}>KI analysiert {techniker.length} Techniker...</div>}
                {aiAnalysis && <div style={{ background: "#111827", border: "1px solid #1f2937", borderRadius: 8, padding: "20px", fontSize: 13, lineHeight: 1.8, color: "#d1d5db" }} dangerouslySetInnerHTML={{ __html: renderMarkdown(aiAnalysis) }} />}
                {error && <div style={{ background: "#2e0f0f", border: "1px solid #7f1d1d", borderRadius: 8, padding: 16, color: "#f87171", fontSize: 13 }}>{error}</div>}
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}
