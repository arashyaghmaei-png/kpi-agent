'use client';

import { useState, useCallback, useRef } from "react";

const BASELINE = { cc_rate: 97.6, termintreue: 97.7, loesungsquote: 96.0, nps: 69.9, zufriedenheit: 4.82 };
const BASELINE_FS5335 = { cc_rate: 99.6, termintreue: 99.1, loesungsquote: 96.9, nps: 74.4, zufriedenheit: 4.86 };
const BASELINE_FS5336 = { cc_rate: 95.7, termintreue: 96.7, loesungsquote: 97.2, nps: 66.7, zufriedenheit: 4.77 };
const THRESHOLDS = { kritisch: 0.85, warnung: 0.93 };

const SYSTEM_PROMPT = `Du bist ein operativer KPI-Analyseagent für ein Telekommunikations-Subunternehmen (Telekom-Subunternehmer, Kupfer & FTTH, Bergheim NRW).

Baseline-Werte aus echten Betriebsdaten (KW13–19, Wochendashboard FiberNC):
Gesamt FS53: CC=${BASELINE.cc_rate}% | Termintreue=${BASELINE.termintreue}% | Lösungsquote=${BASELINE.loesungsquote}% | NPS=${BASELINE.nps}
FS5335: CC=${BASELINE_FS5335.cc_rate}% | Termintreue=${BASELINE_FS5335.termintreue}% | Lösungsquote=${BASELINE_FS5335.loesungsquote}%
FS5336: CC=${BASELINE_FS5336.cc_rate}% | Termintreue=${BASELINE_FS5336.termintreue}% | Lösungsquote=${BASELINE_FS5336.loesungsquote}%

Bekannte Warnsignale: KW20 schlechteste Woche (NPS 26, Termintreue 85,7%). KW23–24 FS5336 kritisch: CC 70%, SearchCall 37%.

Aufgabe: Techniker-KPIs bewerten, Frühwarnungen bei ≥7% Abweichung, standortspezifischer Baseline-Vergleich, Leitstellen-Empfehlungen.
Antworte auf Deutsch, direkt und operativ.

## 📊 KPI-Übersicht
[Techniker, Wert, Baseline-Delta, Status]

## 🚨 Frühwarnungen
[Nur kritische Fälle mit Name und Problem]

## 📈 Baseline-Vergleich
[Teamdurchschnitt vs KW13-19]

## 💡 Empfehlungen Leitstelle
[3–5 konkrete Maßnahmen für heute]`;

const EXAMPLE_CSV = `name,standort,cc_rate,termintreue,loesungsquote,auftraege
Mehmet K.,5335,91,94,88,5
Tobias R.,5335,78,82,75,4
Sven L.,5336,65,70,68,3
Igor P.,5336,88,90,85,5
Kai B.,5335,82,85,79,4`;

function parseCSV(text) {
  const lines = text.trim().split("\n");
  if (lines.length < 2) return [];
  const headers = lines[0].split(",").map(h => h.trim());
  return lines.slice(1).filter(l => l.trim()).map(line => {
    const values = line.split(",").map(v => v.trim());
    const obj = {};
    headers.forEach((h, i) => (obj[h] = values[i]));
    return obj;
  });
}

function getStatusVal(value, baseline) {
  const ratio = value / baseline;
  if (ratio < THRESHOLDS.kritisch) return "kritisch";
  if (ratio < THRESHOLDS.warnung) return "warnung";
  return "gut";
}

function StatusBadge({ status }) {
  const s = { gut: { bg: "#0f2e1a", color: "#4ade80", label: "✓ GUT" }, warnung: { bg: "#2e1f00", color: "#fbbf24", label: "⚠ WARNUNG" }, kritisch: { bg: "#2e0f0f", color: "#f87171", label: "✕ KRITISCH" } }[status] || { bg: "#0f2e1a", color: "#4ade80", label: "✓ GUT" };
  return <span style={{ background: s.bg, color: s.color, padding: "2px 8px", borderRadius: 3, fontSize: 11, fontFamily: "monospace", fontWeight: 700, letterSpacing: 1 }}>{s.label}</span>;
}

function KPIBar({ value, baseline, label }) {
  const val = parseFloat(value);
  const color = val / baseline < THRESHOLDS.kritisch ? "#f87171" : val / baseline < THRESHOLDS.warnung ? "#fbbf24" : "#4ade80";
  return (
    <div style={{ marginBottom: 6 }}>
      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11, color: "#9ca3af", marginBottom: 2 }}>
        <span>{label}</span>
        <span style={{ color }}>{value}% <span style={{ color: "#4b5563" }}>/ {baseline}%</span></span>
      </div>
      <div style={{ background: "#1f2937", borderRadius: 2, height: 6, position: "relative" }}>
        <div style={{ width: `${Math.min(100, val)}%`, background: color, height: "100%", borderRadius: 2, transition: "width 0.6s ease" }} />
        <div style={{ position: "absolute", left: `${Math.min(100, baseline)}%`, top: -3, width: 2, height: 12, background: "#6b7280" }} />
      </div>
    </div>
  );
}

function TechCard({ tech }) {
  const bl = tech.standort === "5336" ? BASELINE_FS5336 : BASELINE_FS5335;
  const statuses = [getStatusVal(parseFloat(tech.cc_rate), bl.cc_rate), getStatusVal(parseFloat(tech.termintreue), bl.termintreue), getStatusVal(parseFloat(tech.loesungsquote), bl.loesungsquote)];
  const worst = statuses.includes("kritisch") ? "kritisch" : statuses.includes("warnung") ? "warnung" : "gut";
  const borderColor = worst === "kritisch" ? "#7f1d1d" : worst === "warnung" ? "#78350f" : "#14532d";
  return (
    <div style={{ background: "#111827", border: `1px solid ${borderColor}`, borderRadius: 8, padding: "16px 18px", marginBottom: 12 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
        <div>
          <div style={{ fontWeight: 700, fontSize: 15, color: "#f9fafb" }}>{tech.name}</div>
          <div style={{ fontSize: 11, color: "#6b7280", marginTop: 2 }}>FS{tech.standort || "5335"} · {tech.auftraege || "—"} Aufträge</div>
        </div>
        <StatusBadge status={worst} />
      </div>
      <KPIBar value={tech.cc_rate} baseline={bl.cc_rate} label="CC-Rate" />
      <KPIBar value={tech.termintreue} baseline={bl.termintreue} label="Termintreue" />
      <KPIBar value={tech.loesungsquote} baseline={bl.loesungsquote} label="Lösungsquote" />
    </div>
  );
}

function renderMarkdown(text) {
  return text
    .replace(/## (.*)/g, '<h3 style="color:#f9fafb;margin:20px 0 8px;font-size:14px;letter-spacing:0.5px">$1</h3>')
    .replace(/\*\*(.*?)\*\*/g, '<strong style="color:#e5e7eb">$1</strong>')
    .replace(/\n/g, "<br/>");
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

  const processXLSX = useCallback((file) => {
    const reader = new FileReader();
    reader.onload = async (ev) => {
      try {
        const XLSX = await import('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm');
        const wb = XLSX.read(ev.target.result, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
        if (!rows.length) { setError("Excel-Datei ist leer oder hat kein erkanntes Format."); return; }
        // Normalize column names (lowercase, trim)
        const normalized = rows.map(row => {
          const n = {};
          Object.keys(row).forEach(k => { n[k.toLowerCase().trim().replace(/\s+/g, "_")] = row[k]; });
          return n;
        });
        setTechniker(normalized);
        setAiAnalysis("");
        setError("");
      } catch (e) {
        setError("Fehler beim Lesen der Excel-Datei: " + e.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const handleFile = useCallback((e) => {
    const file = e.target.files[0];
    if (!file) return;
    setFileName(file.name);
    if (file.name.endsWith(".xlsx") || file.name.endsWith(".xls")) {
      processXLSX(file);
    } else {
      const reader = new FileReader();
      reader.onload = (ev) => { setTechniker(parseCSV(ev.target.result)); setAiAnalysis(""); setError(""); };
      reader.readAsText(file);
    }
  }, [processXLSX]);

  const loadExample = () => { setTechniker(parseCSV(EXAMPLE_CSV)); setAiAnalysis(""); setFileName("Beispieldaten"); setError(""); };

  const runAnalysis = async () => {
    if (!techniker.length) return;
    setLoading(true); setError(""); setAiAnalysis("");
    const dataStr = techniker.map(t =>
      `${t.name} (FS${t.standort}): CC=${t.cc_rate}%, Termintreue=${t.termintreue}%, Lösungsquote=${t.loesungsquote}%, Aufträge=${t.auftraege}`
    ).join("\n");
    try {
      const res = await fetch("/api/analyse", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-6", max_tokens: 1500,
          system: SYSTEM_PROMPT,
          messages: [{ role: "user", content: `Analysiere diese Techniker-KPIs für heute:\n\n${dataStr}` }],
        }),
      });
      const data = await res.json();
      const text = data.content?.map(b => b.text || "").join("") || "Keine Antwort erhalten.";
      setAiAnalysis(text); setActiveTab("analyse");
    } catch (e) { setError("Fehler bei der KI-Analyse. Bitte erneut versuchen."); }
    finally { setLoading(false); }
  };

  const exportScreenshot = async () => {
    setExporting(true);
    try {
      const html2canvas = (await import('https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/+esm')).default;
      const canvas = await html2canvas(dashboardRef.current, { backgroundColor: "#0a0e1a", scale: 2 });
      const link = document.createElement("a");
      link.download = `KPI-Report-${new Date().toLocaleDateString("de-DE").replace(/\./g, "-")}.png`;
      link.href = canvas.toDataURL("image/png");
      link.click();
    } catch (e) { setError("Screenshot-Export fehlgeschlagen."); }
    finally { setExporting(false); }
  };

  const exportPDF = async () => {
    setExporting(true);
    try {
      const html2canvas = (await import('https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/+esm')).default;
      const { jsPDF } = await import('https://cdn.jsdelivr.net/npm/jspdf@2.5.1/+esm');
      const canvas = await html2canvas(dashboardRef.current, { backgroundColor: "#0a0e1a", scale: 2 });
      const imgData = canvas.toDataURL("image/png");
      const pdf = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
      const pdfWidth = pdf.internal.pageSize.getWidth();
      const pdfHeight = (canvas.height * pdfWidth) / canvas.width;
      const date = new Date().toLocaleDateString("de-DE");
      pdf.setFontSize(10); pdf.setTextColor(150);
      pdf.text(`KPI Report FiberNC – ${date}`, 10, 8);
      pdf.addImage(imgData, "PNG", 0, 12, pdfWidth, pdfHeight);
      pdf.save(`KPI-Report-${date.replace(/\./g, "-")}.pdf`);
    } catch (e) { setError("PDF-Export fehlgeschlagen."); }
    finally { setExporting(false); }
  };

  const criticalCount = techniker.filter(t => {
    const bl = t.standort === "5336" ? BASELINE_FS5336 : BASELINE_FS5335;
    return [getStatusVal(parseFloat(t.cc_rate), bl.cc_rate), getStatusVal(parseFloat(t.termintreue), bl.termintreue), getStatusVal(parseFloat(t.loesungsquote), bl.loesungsquote)].includes("kritisch");
  }).length;

  const avgCC = techniker.length ? (techniker.reduce((s, t) => s + parseFloat(t.cc_rate || 0), 0) / techniker.length).toFixed(1) : "—";
  const avgTT = techniker.length ? (techniker.reduce((s, t) => s + parseFloat(t.termintreue || 0), 0) / techniker.length).toFixed(1) : "—";

  return (
    <div style={{ background: "#0a0e1a", minHeight: "100vh", fontFamily: "'Inter', system-ui, sans-serif", color: "#e5e7eb" }}>
      {/* Header */}
      <div style={{ borderBottom: "1px solid #1f2937", padding: "14px 24px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{ width: 8, height: 8, borderRadius: "50%", background: "#4ade80", boxShadow: "0 0 6px #4ade80" }} />
          <span style={{ fontWeight: 700, fontSize: 13, letterSpacing: 2, color: "#9ca3af", textTransform: "uppercase" }}>KPI Agent</span>
          <span style={{ color: "#374151", fontSize: 13 }}>·</span>
          <span style={{ fontSize: 12, color: "#6b7280" }}>Techniker-Kontrolle</span>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
          {techniker.length > 0 && (
            <>
              <button onClick={exportScreenshot} disabled={exporting} style={{ background: "#1f2937", color: "#9ca3af", border: "1px solid #374151", padding: "6px 14px", borderRadius: 6, cursor: "pointer", fontSize: 12, fontWeight: 600 }}>
                {exporting ? "..." : "📸 PNG"}
              </button>
              <button onClick={exportPDF} disabled={exporting} style={{ background: "#1f2937", color: "#9ca3af", border: "1px solid #374151", padding: "6px 14px", borderRadius: 6, cursor: "pointer", fontSize: 12, fontWeight: 600 }}>
                {exporting ? "..." : "📄 PDF"}
              </button>
            </>
          )}
          <div style={{ fontSize: 11, color: "#4b5563", fontFamily: "monospace" }}>Baseline KW13–19 · FiberNC</div>
        </div>
      </div>

      <div ref={dashboardRef} style={{ maxWidth: 720, margin: "0 auto", padding: "24px 20px" }}>

        {/* Upload Zone */}
        {!techniker.length && (
          <div style={{ textAlign: "center", padding: "60px 20px" }}>
            <div style={{ fontSize: 36, marginBottom: 12 }}>📁</div>
            <div style={{ fontSize: 18, fontWeight: 700, color: "#f9fafb", marginBottom: 8 }}>Excel oder CSV hochladen</div>
            <div style={{ fontSize: 13, color: "#6b7280", marginBottom: 6, lineHeight: 1.6 }}>
              Excel (.xlsx): Spalten müssen enthalten: <span style={{ fontFamily: "monospace", color: "#9ca3af" }}>name, standort, cc_rate, termintreue, loesungsquote, auftraege</span>
            </div>
            <div style={{ fontSize: 12, color: "#4b5563", marginBottom: 24 }}>Unterstützt: .xlsx, .xls, .csv</div>
            <label style={{ display: "inline-block", background: "#1d4ed8", color: "#fff", padding: "10px 24px", borderRadius: 6, cursor: "pointer", fontSize: 13, fontWeight: 600, marginRight: 12 }}>
              📂 Datei wählen (.xlsx / .csv)
              <input type="file" accept=".csv,.xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
            </label>
            <button onClick={loadExample} style={{ background: "#1f2937", color: "#9ca3af", border: "1px solid #374151", padding: "10px 24px", borderRadius: 6, cursor: "pointer", fontSize: 13, fontWeight: 600 }}>
              Beispieldaten laden
            </button>
            {error && <div style={{ marginTop: 20, background: "#2e0f0f", border: "1px solid #7f1d1d", borderRadius: 8, padding: 14, color: "#f87171", fontSize: 13 }}>{error}</div>}
          </div>
        )}

        {techniker.length > 0 && (
          <>
            {/* File info + stats */}
            {fileName && <div style={{ fontSize: 11, color: "#4b5563", marginBottom: 12, fontFamily: "monospace" }}>📂 {fileName} · {techniker.length} Einträge</div>}

            <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 10, marginBottom: 20 }}>
              {[
                { label: "Techniker", value: techniker.length, color: "#60a5fa" },
                { label: "Kritisch", value: criticalCount, color: criticalCount > 0 ? "#f87171" : "#4ade80" },
                { label: "Ø CC-Rate", value: `${avgCC}%`, color: parseFloat(avgCC) >= BASELINE.cc_rate ? "#4ade80" : "#fbbf24" },
                { label: "Ø Termintreue", value: `${avgTT}%`, color: parseFloat(avgTT) >= BASELINE.termintreue ? "#4ade80" : "#fbbf24" },
              ].map(s => (
                <div key={s.label} style={{ background: "#111827", border: "1px solid #1f2937", borderRadius: 8, padding: "12px 14px" }}>
                  <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 4 }}>{s.label}</div>
                  <div style={{ fontSize: 20, fontWeight: 700, color: s.color, fontFamily: "monospace" }}>{s.value}</div>
                </div>
              ))}
            </div>

            {/* Tabs */}
            <div style={{ display: "flex", marginBottom: 16, borderBottom: "1px solid #1f2937" }}>
              {[{ id: "dashboard", label: "Dashboard" }, { id: "analyse", label: "KI-Analyse" + (aiAnalysis ? " ✓" : "") }].map(tab => (
                <button key={tab.id} onClick={() => setActiveTab(tab.id)} style={{ background: "none", border: "none", borderBottom: activeTab === tab.id ? "2px solid #3b82f6" : "2px solid transparent", color: activeTab === tab.id ? "#f9fafb" : "#6b7280", padding: "8px 16px", cursor: "pointer", fontSize: 13, fontWeight: activeTab === tab.id ? 600 : 400, marginBottom: -1 }}>
                  {tab.label}
                </button>
              ))}
            </div>

            {activeTab === "dashboard" && (
              <>
                <div style={{ marginBottom: 16 }}>
                  {techniker.map((t, i) => <TechCard key={i} tech={t} />)}
                </div>
                <button onClick={runAnalysis} disabled={loading} style={{ width: "100%", background: loading ? "#1f2937" : "#1d4ed8", color: loading ? "#6b7280" : "#fff", border: "none", borderRadius: 8, padding: "14px", fontSize: 14, fontWeight: 700, cursor: loading ? "not-allowed" : "pointer" }}>
                  {loading ? "KI analysiert..." : "🤖 KI-Analyse starten"}
                </button>
                <button onClick={() => { setTechniker([]); setAiAnalysis(""); setFileName(""); setError(""); }} style={{ width: "100%", marginTop: 8, background: "none", color: "#4b5563", border: "1px solid #1f2937", borderRadius: 8, padding: "10px", fontSize: 12, cursor: "pointer" }}>
                  Neue Datei laden
                </button>
              </>
            )}

            {activeTab === "analyse" && (
              <div>
                {!aiAnalysis && !loading && (
                  <div style={{ textAlign: "center", padding: "40px 20px", color: "#6b7280" }}>
                    <div style={{ fontSize: 32, marginBottom: 12 }}>🤖</div>
                    <div style={{ fontSize: 14 }}>Noch keine Analyse. Dashboard öffnen und Analyse starten.</div>
                  </div>
                )}
                {loading && (
                  <div style={{ textAlign: "center", padding: "40px 20px" }}>
                    <div style={{ fontSize: 14, color: "#6b7280" }}>KI analysiert Daten...</div>
                    <div style={{ marginTop: 12, display: "flex", justifyContent: "center", gap: 6 }}>
                      {[0, 1, 2].map(i => <div key={i} style={{ width: 8, height: 8, borderRadius: "50%", background: "#3b82f6", animation: `pulse 1.2s ease-in-out ${i * 0.2}s infinite` }} />)}
                    </div>
                  </div>
                )}
                {aiAnalysis && (
                  <div style={{ background: "#111827", border: "1px solid #1f2937", borderRadius: 8, padding: "20px 22px", fontSize: 13, lineHeight: 1.8, color: "#d1d5db" }}
                    dangerouslySetInnerHTML={{ __html: renderMarkdown(aiAnalysis) }} />
                )}
                {error && <div style={{ background: "#2e0f0f", border: "1px solid #7f1d1d", borderRadius: 8, padding: 16, color: "#f87171", fontSize: 13 }}>{error}</div>}
              </div>
            )}
          </>
        )}
      </div>
      <style>{`@keyframes pulse { 0%, 100% { opacity: 0.3; transform: scale(0.8); } 50% { opacity: 1; transform: scale(1); } }`}</style>
    </div>
  );
}
