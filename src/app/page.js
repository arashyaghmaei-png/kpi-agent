'use client';

import { useState, useCallback, useRef } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import html2canvas from "html2canvas";
import { jsPDF } from "jspdf";

const BASELINE = {
  cc_rate: 97.6,
  termintreue: 97.7,
  loesungsquote: 96.0,
};

const SYSTEM_PROMPT = `Du bist ein operativer KPI-Analyseagent für ein Telekommunikations-Subunternehmen (Telekom-Subunternehmer, Kupfer & FTTH, Bergheim NRW).

Baseline KW13-19: CC=97,6% | Termintreue=97,7% | Lösungsquote=96,0%
Warnlogik: Warnung ab >=7 Prozentpunkte unter Baseline, kritisch ab >=15 Prozentpunkte unter Baseline.

Du erhältst echte Techniker-Daten aus dem Telekom Auftragsinfo-Export. Bewerte jeden Techniker, gib Frühwarnungen, formuliere konkrete Leitstellen-Empfehlungen. Antworte auf Deutsch, direkt und operativ.

## KPI-Übersicht
[Techniker | CC | Termintreue | Infoquote | NPS | Status]

## Frühwarnungen
[Nur kritische Fälle mit Name und Problem]

## Team-Durchschnitt vs Baseline
[Delta in Prozentpunkten]

## Empfehlungen Leitstelle
[3-5 konkrete Maßnahmen]`;

function parsePercent(val) {
  if (!val) return 0;
  return parseFloat(String(val).replace("%", "").replace(",", ".")) || 0;
}

function parseNumber(val) {
  if (!val) return 0;
  return parseFloat(String(val).replace(/\./g, "").replace(",", ".")) || 0;
}

function normalizeRows(rows) {
  return rows
    .filter(r => r && r["Name"] && String(r["Name"]).trim().length > 0)
    .filter(r => !String(r["Name"]).includes("Diese Datei muss"))
    .map(r => {
      const obj = {};
      Object.keys(r).forEach(k => {
        obj[String(k).trim()] = typeof r[k] === "string" ? r[k].trim() : r[k];
      });
      return obj;
    });
}

function parseCSVTelekom(text) {
  const cleaned = text.replace(/^\uFEFF/, "").trim();

  const result = Papa.parse(cleaned, {
    header: true,
    delimiter: ";",
    skipEmptyLines: true,
    transformHeader: h => h.trim(),
    transform: v => typeof v === "string" ? v.replace(/^"|"$/g, "").trim() : v,
  });

  return normalizeRows(result.data);
}

function parseExcelTelekom(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];

  const rows = XLSX.utils.sheet_to_json(sheet, {
    defval: "",
    raw: false,
  });

  return normalizeRows(rows);
}

function getStatusVal(value, baseline) {
  const delta = baseline - value;

  if (delta >= 15) return "kritisch";
  if (delta >= 7) return "warnung";
  return "gut";
}

function getWorstStatus(statuses) {
  if (statuses.includes("kritisch")) return "kritisch";
  if (statuses.includes("warnung")) return "warnung";
  return "gut";
}

function getTechnikerStatus(tech) {
  const cc = parsePercent(tech["CC"]);
  const tt = parsePercent(tech["Termintreue"]);
  const lq = parsePercent(tech["Erledigt B"]);

  const statuses = [
    getStatusVal(cc, BASELINE.cc_rate),
    getStatusVal(tt, BASELINE.termintreue),
  ];

  if (lq > 0) {
    statuses.push(getStatusVal(lq, BASELINE.loesungsquote));
  }

  return getWorstStatus(statuses);
}

function weightedAverage(rows, field) {
  const totalWeight = rows.reduce((sum, r) => sum + parseNumber(r["Anzahl"]), 0);

  if (!totalWeight) return 0;

  return rows.reduce((sum, r) => {
    return sum + parsePercent(r[field]) * parseNumber(r["Anzahl"]);
  }, 0) / totalWeight;
}

function weightedNPS(rows) {
  const totalWeight = rows.reduce((sum, r) => sum + parseNumber(r["Anzahl"]), 0);

  if (!totalWeight) return 0;

  return rows.reduce((sum, r) => {
    const nps = parseFloat(String(r["NPS PB"] || "0").replace(",", ".")) || 0;
    return sum + nps * parseNumber(r["Anzahl"]);
  }, 0) / totalWeight;
}

function StatusBadge({ status }) {
  const s = {
    gut: {
      bg: "#0f2e1a",
      color: "#4ade80",
      label: "GUT",
    },
    warnung: {
      bg: "#2e1f00",
      color: "#fbbf24",
      label: "WARNUNG",
    },
    kritisch: {
      bg: "#2e0f0f",
      color: "#f87171",
      label: "KRITISCH",
    },
  }[status] || {
    bg: "#0f2e1a",
    color: "#4ade80",
    label: "GUT",
  };

  return (
    <span
      style={{
        background: s.bg,
        color: s.color,
        padding: "2px 10px",
        borderRadius: 3,
        fontSize: 11,
        fontFamily: "monospace",
        fontWeight: 700,
      }}
    >
      {s.label}
    </span>
  );
}

function KPIBar({ value, baseline, label }) {
  const status = getStatusVal(value, baseline);

  const color =
    status === "kritisch"
      ? "#f87171"
      : status === "warnung"
      ? "#fbbf24"
      : "#4ade80";

  return (
    <div style={{ marginBottom: 6 }}>
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          fontSize: 11,
          color: "#9ca3af",
          marginBottom: 2,
        }}
      >
        <span>{label}</span>
        <span style={{ color }}>
          {value.toFixed(1)}%{" "}
          <span style={{ color: "#4b5563" }}>/ {baseline}%</span>
        </span>
      </div>

      <div
        style={{
          background: "#1f2937",
          borderRadius: 2,
          height: 6,
          position: "relative",
        }}
      >
        <div
          style={{
            width: `${Math.min(100, value)}%`,
            background: color,
            height: "100%",
            borderRadius: 2,
          }}
        />

        <div
          style={{
            position: "absolute",
            left: `${Math.min(100, baseline)}%`,
            top: -3,
            width: 2,
            height: 12,
            background: "#6b7280",
          }}
        />
      </div>
    </div>
  );
}

function TechCard({ tech }) {
  const cc = parsePercent(tech["CC"]);
  const tt = parsePercent(tech["Termintreue"]);
  const lq = parsePercent(tech["Erledigt B"]);
  const infoquote = parsePercent(tech["Infoquote P"]);
  const nps = parseFloat(String(tech["NPS PB"] || "0").replace(",", ".")) || 0;
  const geplatz = parsePercent(tech["T. Geplatz"]);

  const ccStatus = getStatusVal(cc, BASELINE.cc_rate);
  const ttStatus = getStatusVal(tt, BASELINE.termintreue);
  const lqStatus = lq > 0 ? getStatusVal(lq, BASELINE.loesungsquote) : "gut";

  const worst = getWorstStatus([ccStatus, ttStatus, lqStatus]);

  const borderColor =
    worst === "kritisch"
      ? "#7f1d1d"
      : worst === "warnung"
      ? "#78350f"
      : "#14532d";

  return (
    <div
      style={{
        background: "#111827",
        border: `1px solid ${borderColor}`,
        borderRadius: 8,
        padding: "16px 18px",
        marginBottom: 12,
      }}
    >
      <div
        style={{
          display: "flex",
          justifyContent: "space-between",
          alignItems: "flex-start",
          marginBottom: 12,
        }}
      >
        <div>
          <div style={{ fontWeight: 700, fontSize: 14, color: "#f9fafb" }}>
            {tech["Name"]}
          </div>

          <div style={{ fontSize: 11, color: "#6b7280", marginTop: 2 }}>
            OD {tech["OD"] || "—"} · {tech["Anzahl"] || "0"} Aufträge ·{" "}
            {tech["Sterne"] || "—"} Sterne
          </div>
        </div>

        <StatusBadge status={worst} />
      </div>

      <KPIBar value={cc} baseline={BASELINE.cc_rate} label="CC-Rate" />
      <KPIBar value={tt} baseline={BASELINE.termintreue} label="Termintreue" />

      {lq > 0 && (
        <KPIBar
          value={lq}
          baseline={BASELINE.loesungsquote}
          label="Lösungsquote / Erledigt B"
        />
      )}

      <div
        style={{
          display: "flex",
          gap: 16,
          marginTop: 8,
          fontSize: 11,
          color: "#6b7280",
          flexWrap: "wrap",
        }}
      >
        <span>
          Infoquote:{" "}
          <span style={{ color: infoquote >= 90 ? "#4ade80" : "#fbbf24" }}>
            {infoquote.toFixed(0)}%
          </span>
        </span>

        <span>
          NPS:{" "}
          <span
            style={{
              color: nps >= 50 ? "#4ade80" : nps >= 0 ? "#fbbf24" : "#f87171",
            }}
          >
            {nps.toFixed(0)}
          </span>
        </span>

        <span>
          Geplatzt:{" "}
          <span style={{ color: geplatz > 5 ? "#f87171" : "#4ade80" }}>
            {tech["T. Geplatz"] || "0%"}
          </span>
        </span>
      </div>
    </div>
  );
}

function renderSafeText(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/## (.*)/g, '<h3 style="color:#f9fafb;margin:20px 0 8px;font-size:14px">$1</h3>')
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
  const [fileType, setFileType] = useState("");
  const dashboardRef = useRef(null);

  const handleFile = useCallback(async e => {
    const file = e.target.files[0];
    if (!file) return;

    setFileName(file.name);
    setFileType(file.type || "");
    setError("");
    setAiAnalysis("");
    setTechniker([]);
    setLoading(true);

    const extension = file.name.split(".").pop()?.toLowerCase();

    try {
      if (extension === "csv") {
        const text = await file.text();
        const parsed = parseCSVTelekom(text);

        if (!parsed.length) {
          setError("Keine Daten gefunden. Bitte Telekom-Auftragsinfo-CSV prüfen.");
          return;
        }

        setTechniker(parsed);
        setActiveTab("dashboard");
        return;
      }

      if (extension === "xls" || extension === "xlsx") {
        const buffer = await file.arrayBuffer();
        const parsed = parseExcelTelekom(buffer);

        if (!parsed.length) {
          setError("Keine Daten gefunden. Bitte Excel-Auftragsinfo prüfen.");
          return;
        }

        setTechniker(parsed);
        setActiveTab("dashboard");
        return;
      }

      if (
        ["pdf", "doc", "docx", "png", "jpg", "jpeg", "webp"].includes(extension)
      ) {
        const formData = new FormData();
        formData.append("file", file);

        const res = await fetch("/api/extract", {
          method: "POST",
          body: formData,
        });

        if (!res.ok) {
          throw new Error("Backend-Extraktion fehlgeschlagen.");
        }

        const data = await res.json();

        if (!data.rows || !Array.isArray(data.rows) || !data.rows.length) {
          setError(
            "Datei wurde erkannt, aber keine KPI-Tabelle gefunden. PDF/Word/Bilder brauchen OCR/KI im Backend."
          );
          return;
        }

        setTechniker(normalizeRows(data.rows));
        setActiveTab("dashboard");
        return;
      }

      setError("Dateityp nicht unterstützt.");
    } catch (err) {
      setError(
        "Upload fehlgeschlagen. CSV/Excel gehen direkt. PDF/Word/Bilder brauchen /api/extract."
      );
    } finally {
      setLoading(false);
      e.target.value = "";
    }
  }, []);

  const runAnalysis = async () => {
    if (!techniker.length) return;

    setLoading(true);
    setError("");
    setAiAnalysis("");

    const dataStr = techniker
      .map(
        t =>
          `${t["Name"]} (OD${t["OD"] || "—"}): CC=${t["CC"]}, Termintreue=${
            t["Termintreue"]
          }, Lösungsquote=${t["Erledigt B"]}, Infoquote=${
            t["Infoquote P"]
          }, NPS=${t["NPS PB"]}, Aufträge=${t["Anzahl"]}, Sterne=${
            t["Sterne"]
          }, Geplatzt=${t["T. Geplatz"]}`
      )
      .join("\n");

    try {
      const res = await fetch("/api/analyse", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          model: "claude-sonnet-4-6",
          max_tokens: 1500,
          system: SYSTEM_PROMPT,
          messages: [
            {
              role: "user",
              content: `Analysiere diese Techniker-KPIs aus Telekom Auftragsinfo:\n\n${dataStr}`,
            },
          ],
        }),
      });

      if (!res.ok) {
        throw new Error("Analyse API fehlgeschlagen.");
      }

      const data = await res.json();

      const text =
        data.content?.map(b => b.text || "").join("") ||
        data.text ||
        data.answer ||
        "Keine Antwort.";

      setAiAnalysis(text);
      setActiveTab("analyse");
    } catch (e) {
      setError("Fehler bei der KI-Analyse. Bitte /api/analyse prüfen.");
    } finally {
      setLoading(false);
    }
  };

  const exportScreenshot = async () => {
    setExporting(true);

    try {
      const canvas = await html2canvas(dashboardRef.current, {
        backgroundColor: "#0a0e1a",
        scale: 2,
      });

      const link = document.createElement("a");
      link.download = `KPI-${new Date()
        .toLocaleDateString("de-DE")
        .replace(/\./g, "-")}.png`;
      link.href = canvas.toDataURL("image/png");
      link.click();
    } catch (e) {
      setError("Screenshot fehlgeschlagen.");
    } finally {
      setExporting(false);
    }
  };

  const exportPDF = async () => {
    setExporting(true);

    try {
      const canvas = await html2canvas(dashboardRef.current, {
        backgroundColor: "#0a0e1a",
        scale: 2,
      });

      const pdf = new jsPDF({
        orientation: "portrait",
        unit: "mm",
        format: "a4",
      });

      const w = pdf.internal.pageSize.getWidth();
      const h = (canvas.height * w) / canvas.width;

      pdf.addImage(canvas.toDataURL("image/png"), "PNG", 0, 10, w, h);
      pdf.save(
        `KPI-${new Date().toLocaleDateString("de-DE").replace(/\./g, "-")}.pdf`
      );
    } catch (e) {
      setError("PDF fehlgeschlagen.");
    } finally {
      setExporting(false);
    }
  };

  const criticalCount = techniker.filter(
    t => getTechnikerStatus(t) === "kritisch"
  ).length;

  const warningCount = techniker.filter(
    t => getTechnikerStatus(t) === "warnung"
  ).length;

  const avgCC = techniker.length
    ? weightedAverage(techniker, "CC").toFixed(1)
    : "—";

  const avgTT = techniker.length
    ? weightedAverage(techniker, "Termintreue").toFixed(1)
    : "—";

  const avgNPS = techniker.length ? weightedNPS(techniker).toFixed(0) : "—";

  const totalOrders = techniker.length
    ? techniker.reduce((s, t) => s + parseNumber(t["Anzahl"]), 0)
    : 0;

  return (
    <div
      style={{
        background: "#0a0e1a",
        minHeight: "100vh",
        fontFamily: "system-ui, sans-serif",
        color: "#e5e7eb",
      }}
    >
      <div
        style={{
          borderBottom: "1px solid #1f2937",
          padding: "14px 24px",
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          gap: 12,
        }}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div
            style={{
              width: 8,
              height: 8,
              borderRadius: "50%",
              background: "#4ade80",
              boxShadow: "0 0 6px #4ade80",
            }}
          />

          <span
            style={{
              fontWeight: 700,
              fontSize: 13,
              letterSpacing: 2,
              color: "#9ca3af",
              textTransform: "uppercase",
            }}
          >
            KPI Agent
          </span>

          <span style={{ color: "#374151" }}>·</span>

          <span style={{ fontSize: 12, color: "#6b7280" }}>
            CSV · Excel · PDF · Word · Screenshot
          </span>
        </div>

        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          {techniker.length > 0 && (
            <>
              <button
                onClick={exportScreenshot}
                disabled={exporting}
                style={{
                  background: "#1f2937",
                  color: "#9ca3af",
                  border: "1px solid #374151",
                  padding: "6px 14px",
                  borderRadius: 6,
                  cursor: "pointer",
                  fontSize: 12,
                }}
              >
                PNG
              </button>

              <button
                onClick={exportPDF}
                disabled={exporting}
                style={{
                  background: "#1f2937",
                  color: "#9ca3af",
                  border: "1px solid #374151",
                  padding: "6px 14px",
                  borderRadius: 6,
                  cursor: "pointer",
                  fontSize: 12,
                }}
              >
                PDF
              </button>
            </>
          )}

          <span
            style={{
              fontSize: 11,
              color: "#4b5563",
              fontFamily: "monospace",
            }}
          >
            Baseline KW13-19
          </span>
        </div>
      </div>

      <div
        ref={dashboardRef}
        style={{
          maxWidth: 820,
          margin: "0 auto",
          padding: "24px 20px",
        }}
      >
        {!techniker.length && (
          <div style={{ textAlign: "center", padding: "60px 20px" }}>
            <div style={{ fontSize: 36, marginBottom: 12 }}>📁</div>

            <div
              style={{
                fontSize: 18,
                fontWeight: 700,
                color: "#f9fafb",
                marginBottom: 8,
              }}
            >
              Telekom Auftragsinfo hochladen
            </div>

            <div
              style={{
                fontSize: 13,
                color: "#6b7280",
                marginBottom: 24,
              }}
            >
              CSV und Excel werden direkt gelesen. PDF, Word und Bilder laufen
              über /api/extract.
            </div>

            <label
              style={{
                display: "inline-block",
                background: "#1d4ed8",
                color: "#fff",
                padding: "10px 24px",
                borderRadius: 6,
                cursor: "pointer",
                fontSize: 13,
                fontWeight: 600,
              }}
            >
              Datei hochladen
              <input
                type="file"
                accept=".csv,.xls,.xlsx,.pdf,.doc,.docx,.png,.jpg,.jpeg,.webp"
                onChange={handleFile}
                style={{ display: "none" }}
              />
            </label>

            {loading && (
              <div style={{ marginTop: 16, color: "#6b7280", fontSize: 13 }}>
                Datei wird verarbeitet...
              </div>
            )}

            {error && (
              <div style={{ marginTop: 16, color: "#f87171", fontSize: 13 }}>
                {error}
              </div>
            )}
          </div>
        )}

        {techniker.length > 0 && (
          <>
            {fileName && (
              <div
                style={{
                  fontSize: 11,
                  color: "#4b5563",
                  marginBottom: 12,
                  fontFamily: "monospace",
                }}
              >
                {fileName} · {techniker.length} Techniker · {totalOrders}{" "}
                Aufträge
              </div>
            )}

            <div
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(6, 1fr)",
                gap: 10,
                marginBottom: 20,
              }}
            >
              {[
                {
                  label: "Techniker",
                  value: techniker.length,
                  color: "#60a5fa",
                },
                {
                  label: "Aufträge",
                  value: totalOrders,
                  color: "#60a5fa",
                },
                {
                  label: "Kritisch",
                  value: criticalCount,
                  color: criticalCount > 0 ? "#f87171" : "#4ade80",
                },
                {
                  label: "Warnung",
                  value: warningCount,
                  color: warningCount > 0 ? "#fbbf24" : "#4ade80",
                },
                {
                  label: "Ø CC",
                  value: `${avgCC}%`,
                  color:
                    parseFloat(avgCC) >= BASELINE.cc_rate
                      ? "#4ade80"
                      : "#fbbf24",
                },
                {
                  label: "Ø Termin",
                  value: `${avgTT}%`,
                  color:
                    parseFloat(avgTT) >= BASELINE.termintreue
                      ? "#4ade80"
                      : "#fbbf24",
                },
              ].map(s => (
                <div
                  key={s.label}
                  style={{
                    background: "#111827",
                    border: "1px solid #1f2937",
                    borderRadius: 8,
                    padding: "12px 14px",
                  }}
                >
                  <div
                    style={{
                      fontSize: 10,
                      color: "#6b7280",
                      marginBottom: 4,
                    }}
                  >
                    {s.label}
                  </div>

                  <div
                    style={{
                      fontSize: 18,
                      fontWeight: 700,
                      color: s.color,
                      fontFamily: "monospace",
                    }}
                  >
                    {s.value}
                  </div>
                </div>
              ))}
            </div>

            <div
              style={{
                display: "flex",
                marginBottom: 16,
                borderBottom: "1px solid #1f2937",
              }}
            >
              {[
                { id: "dashboard", label: "Dashboard" },
                {
                  id: "analyse",
                  label: "KI-Analyse" + (aiAnalysis ? " ✓" : ""),
                },
              ].map(tab => (
                <button
                  key={tab.id}
                  onClick={() => setActiveTab(tab.id)}
                  style={{
                    background: "none",
                    border: "none",
                    borderBottom:
                      activeTab === tab.id
                        ? "2px solid #3b82f6"
                        : "2px solid transparent",
                    color: activeTab === tab.id ? "#f9fafb" : "#6b7280",
                    padding: "8px 16px",
                    cursor: "pointer",
                    fontSize: 13,
                    fontWeight: activeTab === tab.id ? 600 : 400,
                    marginBottom: -1,
                  }}
                >
                  {tab.label}
                </button>
              ))}
            </div>

            {activeTab === "dashboard" && (
              <>
                <div style={{ marginBottom: 16 }}>
                  {techniker.map((t, i) => (
                    <TechCard key={i} tech={t} />
                  ))}
                </div>

                <button
                  onClick={runAnalysis}
                  disabled={loading}
                  style={{
                    width: "100%",
                    background: loading ? "#1f2937" : "#1d4ed8",
                    color: loading ? "#6b7280" : "#fff",
                    border: "none",
                    borderRadius: 8,
                    padding: "14px",
                    fontSize: 14,
                    fontWeight: 700,
                    cursor: loading ? "not-allowed" : "pointer",
                  }}
                >
                  {loading ? "KI analysiert..." : "KI-Analyse starten"}
                </button>

                <button
                  onClick={() => {
                    setTechniker([]);
                    setAiAnalysis("");
                    setFileName("");
                    setFileType("");
                    setError("");
                  }}
                  style={{
                    width: "100%",
                    marginTop: 8,
                    background: "none",
                    color: "#4b5563",
                    border: "1px solid #1f2937",
                    borderRadius: 8,
                    padding: "10px",
                    fontSize: 12,
                    cursor: "pointer",
                  }}
                >
                  Neue Datei laden
                </button>

                {error && (
                  <div
                    style={{
                      marginTop: 12,
                      background: "#2e0f0f",
                      border: "1px solid #7f1d1d",
                      borderRadius: 8,
                      padding: 16,
                      color: "#f87171",
                      fontSize: 13,
                    }}
                  >
                    {error}
                  </div>
                )}
              </>
            )}

            {activeTab === "analyse" && (
              <div>
                {!aiAnalysis && !loading && (
                  <div
                    style={{
                      textAlign: "center",
                      padding: "40px",
                      color: "#6b7280",
                    }}
                  >
                    Noch keine Analyse. Dashboard öffnen und starten.
                  </div>
                )}

                {loading && (
                  <div
                    style={{
                      textAlign: "center",
                      padding: "40px",
                      color: "#6b7280",
                    }}
                  >
                    KI analysiert {techniker.length} Techniker...
                  </div>
                )}

                {aiAnalysis && (
                  <div
                    style={{
                      background: "#111827",
                      border: "1px solid #1f2937",
                      borderRadius: 8,
                      padding: "20px",
                      fontSize: 13,
                      lineHeight: 1.8,
                      color: "#d1d5db",
                    }}
                    dangerouslySetInnerHTML={{
                      __html: renderSafeText(aiAnalysis),
                    }}
                  />
                )}

                {error && (
                  <div
                    style={{
                      background: "#2e0f0f",
                      border: "1px solid #7f1d1d",
                      borderRadius: 8,
                      padding: 16,
                      color: "#f87171",
                      fontSize: 13,
                    }}
                  >
                    {error}
                  </div>
                )}
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}
