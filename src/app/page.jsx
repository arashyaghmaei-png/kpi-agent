'use client'; // v2.0

import React, { useState, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";
import jsPDF from "jspdf";
import html2canvas from "html2canvas";

const BASELINE_KEY = "fibernc_baselines";
const ARCHIV_KEY = "fibernc_archiv";
// Die Befunde aus ursachen_bericht.py ([A] im Menue). Eigener Speicher, weil es
// keine Kennzahlen sind, sondern Kundentexte, NFT-Vermerke und Anrufnachweise -
// eine Zeile je Befund. Der Agent RECHNET damit nichts; er zeigt sie zu dem
// Techniker an, dessen Karte man aufmacht.
const URSACHEN_KEY = "fibernc_ursachen";

// ---------------------------------------------------------------------------
// ORDNER VERBINDEN
// Arash musste bisher jede CSV einzeln suchen und hochladen - fuenf Stueck,
// verteilt auf zwei Ordner (Reports in <KW>\, Ursachen in Pipeline\). Mit der
// File System Access API waehlt er EINMAL den Ordner Auftragsinfo_Downloads
// aus, und der Agent holt sich alles selbst.
//
// WARUM NICHT ANDERSHERUM (Vikuline OS schickt hoch): Dann laegen
// Telekom-Kundentexte auf einem Server. Heute liegen sie nur in seinem
// Browser. Diese Frage haengt bei Thomas (Geschaeftsfuehrer) und ist nicht
// meine, sie durch eine Bauentscheidung zu beantworten. Hier liest der
// Browser lokal - die Daten verlassen den Rechner nicht.
//
// NUR CHROME/EDGE. Firefox und Safari koennen das nicht - deshalb erscheint
// der Knopf dort gar nicht erst und der Upload bleibt, wie er war.
// ---------------------------------------------------------------------------
const ORDNER_DB = "fibernc_ordner";

// Der Ordner-Zugriff wird als "Handle" gemerkt - der passt in kein
// localStorage (nur Text), deshalb IndexedDB. Ohne das muesste Arash den
// Ordner nach jedem Neuladen wieder heraussuchen.
function ordnerMerken(handle) {
  return new Promise((ok, fehler) => {
    const a = indexedDB.open(ORDNER_DB, 1);
    a.onupgradeneeded = () => a.result.createObjectStore("h");
    a.onsuccess = () => {
      const t = a.result.transaction("h", "readwrite");
      t.objectStore("h").put(handle, "pipeline");
      t.oncomplete = () => { a.result.close(); ok(); };
      t.onerror = () => fehler(t.error);
    };
    a.onerror = () => fehler(a.error);
  });
}

function ordnerHolen() {
  return new Promise((ok) => {
    try {
      const a = indexedDB.open(ORDNER_DB, 1);
      a.onupgradeneeded = () => a.result.createObjectStore("h");
      a.onsuccess = () => {
        const t = a.result.transaction("h", "readonly");
        const g = t.objectStore("h").get("pipeline");
        g.onsuccess = () => { a.result.close(); ok(g.result || null); };
        g.onerror = () => { a.result.close(); ok(null); };
      };
      a.onerror = () => ok(null);
    } catch (e) { ok(null); }
  });
}

// Welche Datei ist wofuer gut. WICHTIG - hier wird nach dem NAMEN gefiltert,
// nicht nach dem Inhalt: Im selben Ordner liegen auch die Detail-CSVs
// (*_techniker_details.csv usw.). Die haben eine Zeile je AUFTRAG, nicht je
// Techniker - wuerde man sie einlesen, haelt detectFormat sie fuer einen
// Report (die Schalten-Details haben z.B. eine Spalte "Abschluss Call") und
// baut daraus Unsinn. Deshalb: nur was hier steht, wird angefasst.
const DATEI_ROLLEN = [
  { endung: "_sms_feedback_schalten.csv", rolle: "report" },
  { endung: "_sms_feedback.csv", rolle: "report" },
  { endung: "_one_touch.csv", rolle: "report" },
  { endung: "_nftq.csv", rolle: "report" },
  { endung: "_ursachen.csv", rolle: "ursachen" },
];

function dateiRolle(name) {
  const n = String(name || "").toLowerCase();
  for (const r of DATEI_ROLLEN) {
    if (n.endsWith(r.endung)) {
      return { rolle: r.rolle, label: name.slice(0, name.length - r.endung.length) };
    }
  }
  return null;
}

// Alle Dateien einsammeln, zwei Ebenen tief: Auftragsinfo_Downloads enthaelt
// <KW>\ (Reports) und Pipeline\ (Ursachen).
async function* dateienImOrdner(dir, tiefe = 0) {
  for await (const e of dir.values()) {
    if (e.kind === "file") yield e;
    else if (e.kind === "directory" && tiefe < 2) yield* dateienImOrdner(e, tiefe + 1);
  }
}
const FIRMA = (typeof window !== "undefined" && localStorage.getItem("firma_name")) || "Vikuline";   // Firmenname - im Agenten ueber den 'Firma'-Knopf aenderbar
const DEFAULT_BASELINES = {
  gesamt: {
    // SMS-Feedback
    cc_rate: 95,           // Courtesy Calls Zielwert
    termintreue: 96,       // Termintreue Zielwert
    loesungsquote: 95,     // Loesungsquote Bereitstellung
    nps_montage: 68,       // NPS Montage (Spalte 'NPS BS')       - Portal ZW 68,0%
    nps_pb: 68,            // NPS Problembehebung ('NPS PB')      - Portal ZW 68,0%
    nps: 68,               // NPS Schalten (Report 'SMS Fb Schalten')
    // NFTQ
    nftq_montage: 4,       // NFTQ Montage Zielwert
    nftq_schalten: 6.6,    // NFTQ Schalten Zielwert              - Portal ZW 6,6%
    nftq_pb: 8.5,          // NFTQ Problembehebung Zielwert       - Portal ZW 8,5%
    nftq_bereitstellung: 4,// NFTQ Bereitstellung Zielwert
    // Sonstige
    geplatzte_termine: 0.6,// Geplatzte Termine Zielwert
    info_quote_pb: 90,     // Informationsquote PB                - Portal ZW 90,0%
    so_quote: 2,           // SO-Quote
    service_calls: 93,     // Service Calls Carrier
  },
  fs5335: {
    cc_rate: 95, termintreue: 96, loesungsquote: 95, nps_montage: 68, nps_pb: 68, nps: 68,
    nftq_montage: 4, nftq_schalten: 6.6, nftq_pb: 8.5, nftq_bereitstellung: 4,
    geplatzte_termine: 0.6, info_quote_pb: 90, so_quote: 2, service_calls: 93,
  },
  fs5336: {
    cc_rate: 95, termintreue: 96, loesungsquote: 95, nps_montage: 68, nps_pb: 68, nps: 68,
    nftq_montage: 4, nftq_schalten: 6.6, nftq_pb: 8.5, nftq_bereitstellung: 4,
    geplatzte_termine: 0.6, info_quote_pb: 90, so_quote: 2, service_calls: 93,
  },
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
Aufgabe: Techniker-KPIs bewerten, Leitstellen-Empfehlungen. Wenn Vorperioden-Daten vorhanden, Trend je Techniker angeben.

Telekom-Zielwerte. Quelle: Portal > Gesamtsicht Qualitaet, dort steht bei jedem
Diagramm die Zielwertmarkierung "ZW". NICHTS DAVON SCHAETZEN oder ergaenzen.
Diese Werte MUESSEN mit der Excel-Auswertung (kpi_uebersicht.py) uebereinstimmen -
wenn Analyse und Excel verschiedene Ampeln zeigen, ist das schlimmer als gar keine.

SMS-Feedback/Schalten:
- Courtesy Calls (CC-Rate): Ziel >= 95%, Warnung >= 85.5%, Kritisch < 85.5%
- Termintreue: Ziel >= 96%, Warnung >= 86.4%, Kritisch < 86.4%
- Loesungsquote: Ziel >= 95%
- NPS Montage (Report-Spalte "NPS BS"): Ziel >= 68, Warnung >= 40, Kritisch < 40
- NPS Problembehebung ("NPS PB"): Ziel >= 68, Warnung >= 40, Kritisch < 40
- NPS SCHALTEN: NICHT BEWERTEN. Dafuer gibt es in der Gesamtsicht Qualitaet
  keinen Zielwert. Wert nennen, kein Urteil, keine Ampel.
- Informationsquote PB: Ziel >= 87.5%
- Geplatzte Termine: Ziel <= 0.6%, Kritisch > 2%

NFTQ (Fehlerquote - NIEDRIGER ist besser). "NFT" = Nachfolgeticket: ein Auftrag,
bei dem nochmal jemand raus musste. Schreibe das Wort aus, nicht abkuerzen.
- NFTQ Montage: Ziel <= 4%, Warnung <= 8%, Kritisch > 8%
- NFTQ Schalten: Ziel <= 6.6%, Warnung <= 10%, Kritisch > 10%
- NFTQ Problembehebung: Ziel <= 8.5%, Warnung <= 12%, Kritisch > 12%
- NFTQ BEREITSTELLUNG: NICHT BEWERTEN. Bereitstellung = Schalten + Montage,
  ist also nur die Sammelquote der beiden. Wer sie mitzaehlt, zaehlt dieselben
  Nachfolgetickets zweimal und der Techniker sieht doppelt so schlecht aus,
  wie er ist. Telekom hat dafuer folgerichtig auch keinen Zielwert.

MINDESTMENGEN (Vikuline-Regel, keine Telekom-Vorgabe - so aber sagen):
- NFTQ erst ab 10 Auftraegen in DER Kategorie bewerten. Darunter Wert nennen,
  kein Urteil: bei 3 Schalten-Auftraegen sind 33,3% EIN Nachfolgeticket - das
  misst den Zufall, nicht die Arbeit.
- NPS erst ab 2 Rueckmeldungen bewerten. Bei einer ist der Wert nur +100 oder -100.
- NENNE BEI JEDER QUOTE DIE BASIS: "NFTQ Montage 22,2% (4 von 18)", nicht nur
  "22,2%". Eine Quote ohne ihre Basis ist im Gespraech mit dem Monteur wertlos.
- Bei einem Extremwert aus wenigen Faellen: sag ausdruecklich, dass er auf
  wenigen Faellen beruht, statt ihn als Befund zu verkaufen.

OneTouch (im Agenten hinterlegt; ob Telekom dafuer ZW vorgibt, ist UNGEPRUEFT -
also nicht als Telekom-Vorgabe darstellen):
- A1-Quote: Ziel >= 60%, Warnung 45-59%, Kritisch < 45%
- A0-Quote: Ziel <= 5%, Kritisch > 10%

Antworte auf Deutsch, direkt und operativ. Erfinde keine Zielwerte fuer
Kennzahlen, die oben nicht stehen - schreibe dann "kein Zielwert bekannt".

PFLICHT - IMMER AM ENDE - JEDEN TECHNIKER AUFLISTEN - AUCH GUTE:
<MASSNAHMEN>
{"massnahmen":[{"name":"Vollständiger Name","status":"kritisch|warnung|gut","massnahme":"Bei gut: Lob in einem Satz. Bei warnung/kritisch: Konkrete Maßnahme.","betreff":"Bei gut: Lob KPI-Werte. Bei kritisch: KPI Maßnahme"}]}
</MASSNAHMEN>
Status gut = Lob aussprechen. Kein Markdown im JSON.

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
  // Der Ursachenbericht: eine Zeile je Befund, keine Kennzahlen-Tabelle.
  // Muss VOR allen anderen stehen - er hat auch eine Spalte "Kennzahl", und
  // die wuerde sonst faelschlich als Standard-Report durchgehen.
  if (h.includes("einstufung") && h.some(x => x.startsWith("kunde sagt"))) return "ursachen";
  if (h.some(x => x === "a1" || x === "a ges." || x === "a ges" || x === "a0")) return "onetouch";
  if (h.some(x => x.includes("nftq b") || x.includes("nftq s"))) return "nftq";
  if (h.some(x => x.includes("courtesy call") || x.includes("abschluss call"))) return "smsfeedbackschalten";
  if (h.some(x => x.includes("cc anzahl") || x.includes("nps bs") || x.includes("nps pb"))) return "smsfeedback";
  if (h.some(x => x === "cc_rate" || x === "loesungsquote")) return "standard";
  return "standard";
}

function aggregateOneTouch(rawRows) {
  const gruppen = {};
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
    // ATS 35/36 -> Standort (35 = Koeln 5335, 36 = Bonn 5336)
    const ats = String(get("ats") || "").trim();
    const standort = ats === "36" ? "5336" : "5335";
    const key = name + "#" + standort;   // getrennt nach Name + ATS
    if (!gruppen[key]) gruppen[key] = { name, standort, days: [] };
    gruppen[key].days.push({
      gesamt: parseFloat(get("gesamt") || 0) || 0,
      a_ges: parsePercent(get("a ges.", "a ges")),
      a1: parsePercent(get("a1")),
      a2: parsePercent(get("a2")),
      a2plus: parsePercent(get("a2+")),
      ax: parsePercent(get("ax")),
      a0: parsePercent(get("a0")),
    });
  });
  return Object.values(gruppen).map(({ name, standort, days }) => {
    const total = days.reduce((s, d) => s + d.gesamt, 0);
    const wavg = (key) => {
      const num = days.reduce((s, d) => d[key] !== null ? s + d[key] * d.gesamt : s, 0);
      const den = days.reduce((s, d) => d[key] !== null ? s + d.gesamt : s, 0);
      return den > 0 ? Math.round(num / den * 10) / 10 : null;
    };
    return {
      name, standort, auftraege: total,
      a_ges: wavg("a_ges"), a1: wavg("a1"), a2: wavg("a2"),
      a2plus: wavg("a2plus"), ax: wavg("ax"), a0: wavg("a0"),
      tage: days.length, quelle: "onetouch",
      cc_rate: null, termintreue: null, loesungsquote: null, nps: null,
    };
  });
}

function autoDetectType(headers) {
  const h = headers.map(x => String(x || "").toLowerCase());
  if (h.some(x => x.includes("nftq") || x.includes("fehlerquote"))) return "nftq";
  if (h.some(x => x.includes("a1") || x.includes("onetouch") || x.includes("erstlosung"))) return "onetouch";
  if (h.some(x => x.includes("courtesy call") || x.includes("abschluss call") || x.includes("termintreue mit st"))) return "smsfeedbackschalten";
  if (h.some(x => x.includes("schalten") || x.includes("schalt"))) return "smsfeedbackschalten";
  if (h.some(x => x.includes("cc anzahl") || x.includes("nps bs") || x.includes("nps pb") || x.includes("infoquote"))) return "smsfeedback";
  if (h.some(x => x.includes("nps") || x.includes("feedback") || x.includes("sms"))) return "smsfeedback";
  if (h.some(x => x.includes("bemerkung") || x.includes("sterne") || x.includes("anliegen"))) return "smsfeedback";
  return null;
}

function normalizeRows(rawRows) {
  if (!rawRows || !rawRows.length) return [];
  const filtered = rawRows.filter(row => !isJunkRow(row));
  if (!filtered.length) return [];
  const rawHeaders = Object.keys(filtered[0]);
  const headers = rawHeaders.map(cleanHeader);
  const fmt = detectFormat(headers);

  // Befunde sind keine Techniker: sie werden nicht gemittelt, nicht bewertet
  // und nicht zusammengefasst - nur weitergereicht. Deshalb hier raus, bevor
  // die Kennzahlen-Logik darueber laeuft.
  if (fmt === "ursachen") {
    const f = (row, name) => {
      const k = Object.keys(row).find(x => cleanHeader(x).toLowerCase() === name);
      return k ? String(row[k] ?? "").trim() : "";
    };
    return rawRows
      .filter(row => f(row, "techniker"))
      .map(row => ({
        quelle: "ursachen",
        name: f(row, "techniker"),
        datum: f(row, "datum"),
        ats: f(row, "ats"),
        bereich: f(row, "bereich"),
        kennzahl: f(row, "kennzahl"),
        wert: f(row, "wert"),
        einstufung: f(row, "einstufung"),
        auftrag: f(row, "auftrag"),
        kundeSagt: f(row, "kunde sagt"),
        kundeUeber: f(row, "kunde ueber techniker"),
        technikerSagt: f(row, "techniker sagt"),
        weiteres: f(row, "weiteres"),
      }));
  }
  if (fmt === "onetouch") return aggregateOneTouch(filtered);

  const get = (row, ...keys) => {
    for (const key of keys) {
      const raw = rawHeaders.find(h => cleanHeader(h).toLowerCase() === key.toLowerCase());
      if (raw !== undefined && row[raw] !== undefined && row[raw] !== "") return row[raw];
    }
    return null;
  };

  // ATS 35/36 -> Standort-Schluessel der Baselines (5335 = Koeln, 5336 = Bonn)
  const standortAus = (row) => {
    const ats = String(get(row, "ats") || "").trim();
    if (ats === "36") return "5336";
    if (ats === "35") return "5335";
    const od = String(get(row, "od") || "").trim();
    if (od === "5336" || od === "5335") return od;
    const st = String(get(row, "standort") || "").trim();
    if (st === "5336" || st === "5335") return st;
    return "5335";
  };

  return filtered
    .filter(row => { const name = get(row, "name", "techniker"); return name && String(name).trim().length > 2; })
    .map(row => {
      const name = String(get(row, "name", "techniker") || "").trim();
      const standort = standortAus(row);

      if (fmt === "smsfeedback") {
        // Zwei getrennte NPS: BS = Montage/Bereitstellung, PB = Problembehebung
        const npsZahl = (r) => {
          if (r === null || r === undefined || r === "") return null;
          const v = parseFloat(String(r).replace(",", ".").replace("%", ""));
          return isNaN(v) ? null : v;
        };
        return {
          name, standort,
          cc_rate: parsePercent(get(row, "cc")),
          termintreue: parsePercent(get(row, "termintreue")),
          loesungsquote: parsePercent(get(row, "erledigt b") ?? get(row, "erledigt")),
          infoquote_p: parsePercent(get(row, "infoquote p")),
          geplatzte_termine: parsePercent(get(row, "t. geplatz")),
          sterne: parseFloat(String(get(row, "sterne") || "").replace(",", ".")) || null,
          anzahl_nps: parseInt(get(row, "anzahl nps gesamt")) || null,
          nps_montage: npsZahl(get(row, "nps bs")),
          nps_pb: npsZahl(get(row, "nps pb")),
          anzahl_nps_montage: parseInt(get(row, "anzahl nps bs")) || null,
          anzahl_nps_pb: parseInt(get(row, "anzahl nps pb")) || null,
          nps: null,
          auftraege: get(row, "anzahl") || "-",
          quelle: "smsfeedback", standortUnbekannt: false
        };
      }

      if (fmt === "smsfeedbackschalten") {
        const npsRaw = get(row, "nps");
        const npsVal = npsRaw !== null && npsRaw !== undefined
          ? parseFloat(String(npsRaw).replace(",", ".")) : null;
        return {
          name, standort,
          cc_rate: parsePercent(get(row, "courtesy call")),
          termintreue: parsePercent(get(row, "termintreue mit st vo") ?? get(row, "termintreue ohne st vo")),
          loesungsquote: parsePercent(get(row, "erledigt")),
          nps: isNaN(npsVal) ? null : npsVal,
          auftraege: get(row, "anzahl") || "-",
          anzahl_nps: parseInt(get(row, "anzahl nps")) || null,
          quelle: "smsfeedbackschalten", standortUnbekannt: false
        };
      }

      if (fmt === "nftq") {
        return {
          name, standort, cc_rate: null, termintreue: null, loesungsquote: null,
          nftq_b: parsePercent(get(row, "nftq b")),
          nftq_s: parsePercent(get(row, "nftq s")),
          nftq_m: parsePercent(get(row, "nftq m")),
          nftq_p: parsePercent(get(row, "nftq p")),
          auftraege: get(row, "anzahl") || "-",
          menge_b: parseInt(get(row, "bereitstellung")) || null,
          menge_s: parseInt(get(row, "schalten")) || null,
          menge_m: parseInt(get(row, "montage")) || null,
          menge_p: parseInt(get(row, "problembehebung")) || null,
          quelle: "nftq", standortUnbekannt: false
        };
      }

      const rawStandort = String(get(row, "standort") || "").trim();
      const standortKlar = rawStandort === "5335" || rawStandort === "5336";
      return {
        name, standort: standortKlar ? rawStandort : standort,
        cc_rate: parsePercent(get(row, "cc_rate")),
        termintreue: parsePercent(get(row, "termintreue")),
        loesungsquote: parsePercent(get(row, "loesungsquote")),
        nps: parsePercent(get(row, "nps")),
        auftraege: get(row, "auftraege") || "-",
        quelle: "standard", standortUnbekannt: !standortKlar
      };
    });
}

function parseCSV(text) {
  const clean = String(text).replace(/^\uFEFF/, "");
  const lines = clean.split(/\r?\n/).filter(l => l.trim().length);
  if (lines.length < 2) return [];
  const sep = lines[0].includes(";") ? ";" : ",";

  const splitLine = (line) => {
    const out = [];
    let cur = "", inQuotes = false;
    for (let i = 0; i < line.length; i++) {
      const c = line[i];
      if (c === '"') {
        if (inQuotes && line[i + 1] === '"') { cur += '"'; i++; }
        else inQuotes = !inQuotes;
      } else if (c === sep && !inQuotes) {
        out.push(cur); cur = "";
      } else {
        cur += c;
      }
    }
    out.push(cur);
    return out.map(v => v.trim());
  };

  const headers = splitLine(lines[0]);
  const rows = lines.slice(1).map(line => {
    const values = splitLine(line);
    const obj = {};
    headers.forEach((h, i) => (obj[h] = values[i] !== undefined ? values[i] : ""));
    return obj;
  });
  return normalizeRows(rows);
}

// Mindestmengen - keine Telekom-Vorgabe, sondern dieselbe Entscheidung wie in
// kpi_uebersicht.py ([U]). Anlass KW28: Ali Sodjajy stand mit NFTQ-S 33,3 % als
// kritisch da - dahinter steckte EIN NFT bei DREI Schalten-Auftraegen. Bei so
// kleinen Mengen misst die Quote den Zufall, nicht die Arbeit.
// Diese zwei Zahlen MUESSEN mit kpi_uebersicht.py uebereinstimmen, sonst sagen
// Excel und Agent Verschiedenes ueber denselben Mann.
const MINDEST_NFTQ = 10;   // Auftraege in DER Kategorie
const MINDEST_NPS = 2;     // Rueckmeldungen. Bei einer ist NPS nur +100 oder -100.

function getNPSStatus(nps, ziel = 68, menge = null) {
  if (nps === null || nps === undefined || isNaN(nps) || nps === "undefined") return null;
  // Zu duenne Datenlage: Wert anzeigen, aber nicht bewerten.
  if (menge !== null && menge !== undefined && menge < MINDEST_NPS) return null;
  if (nps >= ziel) return "gut";
  if (nps >= 40) return "warnung";   // wie [U]: kritisch erst unter 40.
  return "kritisch";                 // Frueher stand hier 20 - der Agent sagte
                                     // dann "warnung", wo die Excel "kritisch"
                                     // sagte (Tsoukalas KW28, NPS PB 33,3).
}

// NFTQ ist eine FEHLERQUOTE: niedriger ist besser, und jede Kategorie hat einen
// EIGENEN Zielwert aus dem Portal. Vorher stand an beiden Status-Stellen hart
// "<= 4 gut, <= 8 warnung" fuer alle vier Spalten - die Zielwerte oben in den
// Baselines wurden nie benutzt. Ein NFTQ-P von 9,1 % galt damit als kritisch,
// obwohl der ZW 8,5 mit Warnbereich bis 12 ist.
function getNFTQStatus(wert, ziel, warn, menge) {
  if (wert === null || wert === undefined || isNaN(wert)) return null;
  if (menge !== null && menge !== undefined && menge < MINDEST_NFTQ) return null;
  if (wert <= ziel) return "gut";
  if (wert <= warn) return "warnung";
  return "kritisch";
}

// Die drei bewerteten NFTQ-Kategorien mit ihren Portal-Zielwerten.
// NFTQ-B (Bereitstellung) steht bewusst NICHT hier: Bereitstellung = Schalten +
// Montage, B ist also nur die Sammelquote der beiden. Wer B mitzaehlt, zaehlt
// dieselben NFTs zweimal - der Techniker sieht doppelt so schlecht aus, wie er
// ist. Telekom hat fuer B folgerichtig auch keinen Zielwert. B wird weiter
// ANGEZEIGT, aber nicht bewertet.
function nftqStatusListe(t, bl) {
  return [
    getNFTQStatus(t.nftq_s, bl.nftq_schalten ?? 6.6, 10, t.menge_s),
    getNFTQStatus(t.nftq_m, bl.nftq_montage ?? 4, 8, t.menge_m),
    getNFTQStatus(t.nftq_p, bl.nftq_pb ?? 8.5, 12, t.menge_p),
  ].filter(x => x !== null);
}

function getStatus(value, baseline) {
  if (value === null || value === undefined || isNaN(value)) return "unbekannt";
  if (value >= baseline) return "gut";            // Ziel erreicht oder besser
  if (value >= baseline * 0.9) return "warnung";  // bis 10% unter Ziel
  return "kritisch";                              // mehr als 10% unter Ziel
}

function getOTStatus(tech) {
  if (tech.a0 !== null && tech.a0 > 10) return "kritisch";
  if (tech.a_ges !== null && tech.a_ges < 85) return "kritisch";
  if (tech.a1 !== null && tech.a1 < 45) return "kritisch";
  if (tech.a1 !== null && tech.a1 < 60) return "warnung";
  if (tech.ax !== null && tech.ax > 20) return "warnung";
  return "gut";
}

function techWorst(t, baselines) {
  const bl = String(t.standort) === "5336" ? baselines.fs5336 : baselines.fs5335;
  const s = [];
  if (t.a1 != null || t.a_ges != null || t.a0 != null) s.push(getOTStatus(t));
  nftqStatusListe(t, bl).forEach(x => s.push(x));
  if (t.cc_rate != null) s.push(getStatus(t.cc_rate, bl.cc_rate));
  if (t.termintreue != null) s.push(getStatus(t.termintreue, bl.termintreue));
  if (t.loesungsquote != null) s.push(getStatus(t.loesungsquote, bl.loesungsquote));
  // NPS Schalten (Feld "nps") wird NICHT bewertet: dafuer gibt es in der
  // Gesamtsicht Qualitaet keinen Zielwert. Adil Kheder stand deswegen mit
  // -100 aus EINER Bewertung als kritisch da, waehrend sein NPS PB 100 war.
  if (t.nps_montage != null) s.push(getNPSStatus(t.nps_montage, bl.nps_montage ?? 68, t.anzahl_nps_montage));
  if (t.nps_pb != null) s.push(getNPSStatus(t.nps_pb, bl.nps_pb ?? 68, t.anzahl_nps_pb));
  return s.includes("kritisch") ? "kritisch" : s.includes("warnung") ? "warnung" : "gut";
}

function getTrend(current, previous) {
  if (current === null || current === undefined || previous === null || previous === undefined) return null;
  const diff = current - previous;
  if (Math.abs(diff) < 0.5) return { symbol: "=", color: "#6b7280", diff: 0 };
  return diff > 0
    ? { symbol: "+" + diff.toFixed(1), color: "#4ade80", diff }
    : { symbol: diff.toFixed(1), color: "#f87171", diff };
}

function parseMassnahmen(text) {
  // Versuch 1: <MASSNAHMEN> Block
  try {
    const match = text.match(/<MASSNAHMEN>([\s\S]*?)<\/MASSNAHMEN>/);
    if (match) {
      const clean = match[1].trim().replace(/```json|```/g, "").trim();
      const json = JSON.parse(clean);
      if (Array.isArray(json.massnahmen) && json.massnahmen.length > 0)
        return { massnahmen: json.massnahmen, fehler: null };
    }
  } catch(e) {}

  // Versuch 2: JSON irgendwo im Text
  try {
    const jsonMatch = text.match(/\{\s*"massnahmen"\s*:\s*\[[\s\S]*?\]\s*\}/);
    if (jsonMatch) {
      const json = JSON.parse(jsonMatch[0]);
      if (Array.isArray(json.massnahmen) && json.massnahmen.length > 0)
        return { massnahmen: json.massnahmen, fehler: null };
    }
  } catch(e) {}

  // Versuch 3: Text-Parsing - Namen aus Analyse extrahieren
  try {
    const massnahmen = [];
    const lines = text.split("\n");
    let currentName = null;
    let currentStatus = "warnung";
    for (const line of lines) {
      const boldName = line.match(/\*\*([A-Z][a-z]+ [A-Z][a-z]+.*?)\*\*/);
      if (boldName) {
        currentName = boldName[1].replace(/\s*[--].*/, "").trim();
        currentStatus = line.toLowerCase().includes("gut") || line.toLowerCase().includes("best performer") || line.toLowerCase().includes("unauffaellig") || line.toLowerCase().includes("unauffällig") ? "gut" : line.toLowerCase().includes("kritisch") ? "kritisch" : "warnung";
      }
      if (currentName && line.includes("- ") && line.length > 20 && !line.includes("**")) {
        const massnahme = line.replace(/^[-\s]+/, "").trim();
        if (massnahme.length > 10) {
          const existing = massnahmen.find(m => m.name === currentName);
          if (!existing) {
            massnahmen.push({ name: currentName, status: currentStatus, massnahme, betreff: currentStatus === "gut" ? "Lob: Sehr gute KPI-Werte" : "KPI Massnahme " + currentName });
          }
        }
      }
    }
    if (massnahmen.length > 0) return { massnahmen, fehler: null };
  } catch(e) {}

  return { massnahmen: [], fehler: "Kein Massnahmen-Block gefunden." };
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
  return `KW${String(kw).padStart(2, "0")} ${jahr} . ${datum}`;
}

const STATUS_STYLE = {
  gut:       { bg: "#0f2e1a", color: "#4ade80", label: "GUT" },
  warnung:   { bg: "#2e1f00", color: "#fbbf24", label: "WARNUNG" },
  kritisch:  { bg: "#2e0f0f", color: "#f87171", label: "KRITISCH" },
  unbekannt: { bg: "#1a1a2e", color: "#6b7280", label: "-" },
};

function StatusBadge({ status }) {
  const s = STATUS_STYLE[status] || STATUS_STYLE.gut;
  return <span style={{ background: s.bg, color: s.color, padding: "2px 8px", borderRadius: 3, fontSize: 11, fontFamily: "monospace", fontWeight: 700 }}>{s.label}</span>;
}

function KPIBar({ value, baseline, label, trend }) {
  if (value === null || value === undefined || isNaN(value)) return null;
  const color = value / baseline < 0.85 ? "#f87171" : value / baseline < 0.93 ? "#fbbf24" : "#4ade80";
  return (
    <div style={{ marginBottom: 6 }}>
      <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11, color: "#9ca3af", marginBottom: 2 }}>
        <span>{label}</span>
        <span style={{ display: "flex", gap: 6, alignItems: "center" }}>
          {trend && <span style={{ color: trend.color, fontSize: 10, fontWeight: 700 }}>{trend.symbol}%</span>}
          <span style={{ color }}>{value.toFixed(1)}% / {baseline}%</span>
        </span>
      </div>
      <div style={{ background: "#1f2937", borderRadius: 2, height: 6, position: "relative" }}>
        <div style={{ width: `${Math.min(100, value)}%`, background: color, height: "100%", borderRadius: 2 }} />
        <div style={{ position: "absolute", left: `${Math.min(100, baseline)}%`, top: -3, width: 2, height: 12, background: "#6b7280" }} />
      </div>
    </div>
  );
}

function NFTQBar({ value, label, ziel, warn }) {
  if (value === null || isNaN(value)) return null;
  const z = ziel || 4;
  const w = warn || 8;
  const color = value > w ? "#f87171" : value > z ? "#fbbf24" : "#4ade80";
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

function TechCard({ tech, baselines, vorperiode, ursachen }) {
  const bl = String(tech.standort) === "5336" ? baselines.fs5336 : baselines.fs5335;
  // Bereiche nach Daten-Vorhandensein (so zeigt eine Kombi-Karte alle gleichzeitig)
  const isOT = tech.a1 != null || tech.a_ges != null || tech.a0 != null || tech.quelle === "onetouch";
  const isNFTQ = [tech.nftq_b, tech.nftq_s, tech.nftq_m, tech.nftq_p].some(v => v != null) || tech.quelle === "nftq";
  const isSMS = tech.cc_rate != null || tech.termintreue != null || tech.loesungsquote != null || tech.nps != null || tech.nps_montage != null || tech.nps_pb != null;
  const _stat = [];
  if (isOT) _stat.push(getOTStatus(tech));
  if (isNFTQ) {
    // VIERTE Stelle, die frueher selbst gerechnet hat - wieder pauschal 4/8 fuer
    // alle vier Spalten, NFTQ-B eingeschlossen. Jetzt dieselbe Regel wie ueberall.
    nftqStatusListe(tech, bl).forEach(x => _stat.push(x));
  }
  if (isSMS) {
    if (tech.cc_rate != null) _stat.push(getStatus(tech.cc_rate, bl.cc_rate));
    if (tech.termintreue != null) _stat.push(getStatus(tech.termintreue, bl.termintreue));
    if (tech.loesungsquote != null) _stat.push(getStatus(tech.loesungsquote, bl.loesungsquote));
    if (tech.nps != null) _stat.push(getNPSStatus(tech.nps, bl.nps ?? 67));
    if (tech.nps_montage != null) _stat.push(getNPSStatus(tech.nps_montage, bl.nps_montage ?? 67));
    if (tech.nps_pb != null) _stat.push(getNPSStatus(tech.nps_pb, bl.nps_pb ?? 67));
  }
  const worst = _stat.includes("kritisch") ? "kritisch" : _stat.includes("warnung") ? "warnung" : "gut";
  const borderColor = worst === "kritisch" ? "#7f1d1d" : worst === "warnung" ? "#78350f" : "#14532d";
  const quelleLabel = { smsfeedback: "SMS-Feedback", smsfeedbackschalten: "Schalten", nftq: "NFTQ", standard: "Manuell", onetouch: "OneTouch", alle: "Alle" }[tech.quelle] || "";
  const npsStatus = tech.nps !== null ? getNPSStatus(tech.nps) : null;
  const npsColor = npsStatus === "kritisch" ? "#f87171" : npsStatus === "warnung" ? "#fbbf24" : "#4ade80";
  return (
    <div style={{ background: "#111827", border: `1px solid ${borderColor}`, borderRadius: 8, padding: "16px 18px", marginBottom: 12 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12 }}>
        <div>
          <div style={{ fontWeight: 700, fontSize: 15, color: "#f9fafb" }}>{tech.name}</div>
          <div style={{ fontSize: 11, color: "#6b7280", marginTop: 2 }}>
            {tech.standortUnbekannt ? <span style={{ color: "#fbbf24" }}>! Standort unbekannt - FS5335</span> : `FS${tech.standort}`}
            {" . "}{tech.auftraege} Aufträge
            {isOT && tech.tage ? <span style={{ marginLeft: 6 }}>. {tech.tage} Tage</span> : null}
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
        {tech.a0 > 0 ? <div style={{ marginTop: 6, fontSize: 11, color: "#f87171" }}>! A0: {tech.a0.toFixed(1)}%</div> : null}
      </>)}
      {isNFTQ && (<>
        <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 8 }}>NFTQ Fehlerquoten</div>
        {(tech.menge_b !== null && tech.menge_b !== undefined) && (
          <div style={{ display: "flex", gap: 12, fontSize: 10, color: "#6b7280", marginBottom: 6, padding: "4px 0" }}>
            <span>B: <strong style={{ color: "#9ca3af" }}>{tech.menge_b}</strong></span>
            <span>S: <strong style={{ color: "#9ca3af" }}>{tech.menge_s ?? "-"}</strong></span>
            <span>M: <strong style={{ color: "#9ca3af" }}>{tech.menge_m ?? "-"}</strong></span>
            <span>P: <strong style={{ color: "#9ca3af" }}>{tech.menge_p ?? "-"}</strong></span>
          </div>
        )}
        <NFTQBar value={tech.nftq_b} label="Bereitstellung" />
        <NFTQBar value={tech.nftq_s} label="Schalten" />
        <NFTQBar value={tech.nftq_m} label="Montage" />
        <NFTQBar value={tech.nftq_p} label="Problembehebung" />
      </>)}
      {isSMS && (<>
        <KPIBar value={tech.cc_rate} baseline={bl.cc_rate} label="CC-Rate"
          trend={vorperiode ? getTrend(tech.cc_rate, vorperiode.cc_rate) : null} />
        <KPIBar value={tech.termintreue} baseline={bl.termintreue} label="Termintreue"
          trend={vorperiode ? getTrend(tech.termintreue, vorperiode.termintreue) : null} />
        {tech.termintreue !== null && (() => { const pts = getTermintreeuPunkte(tech.termintreue); return (
          <div style={{ fontSize: 10, color: pts >= 164 ? "#4ade80" : pts >= 0 ? "#fbbf24" : "#f87171", marginTop: -6, marginBottom: 4, paddingLeft: 2 }}>
            Auftragsinfo Punkte: <b>{pts > 0 ? "+" : ""}{pts}</b>
          </div>
        ); })()}
        <KPIBar value={tech.loesungsquote} baseline={bl.loesungsquote} label="Lösungsquote"
          trend={vorperiode ? getTrend(tech.loesungsquote, vorperiode.loesungsquote) : null} />
        {[
          { wert: tech.nps_montage, ziel: bl.nps_montage ?? 67, label: "NPS Montage", anz: tech.anzahl_nps_montage },
          { wert: tech.nps_pb, ziel: bl.nps_pb ?? 67, label: "NPS Problembeh.", anz: tech.anzahl_nps_pb },
          { wert: tech.nps, ziel: bl.nps ?? 67, label: "NPS", anz: tech.anzahl_nps },
        ].filter(n => n.wert !== null && n.wert !== undefined && !isNaN(n.wert)).map(n => {
          const st = getNPSStatus(n.wert, n.ziel);
          const c = st === "kritisch" ? "#f87171" : st === "warnung" ? "#fbbf24" : "#4ade80";
          return (
            <div key={n.label} style={{ marginTop: 8, display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
              <span style={{ fontSize: 11, color: "#9ca3af" }}>{n.label}:</span>
              <span style={{ color: c, fontWeight: 700, fontSize: 13, fontFamily: "monospace" }}>{n.wert.toFixed(0)}</span>
              <span style={{ background: STATUS_STYLE[st]?.bg, color: c, padding: "1px 6px", borderRadius: 3, fontSize: 10, fontFamily: "monospace", fontWeight: 700 }}>
                {st === "kritisch" ? "KRITISCH" : st === "warnung" ? "WARNUNG" : "GUT"}
              </span>
              <span style={{ fontSize: 10, color: "#4b5563" }}>Ziel: {n.ziel}{n.anz ? ` (${n.anz} Rueckmeldungen)` : ""}</span>
            </div>
          );
        })}
      </>)}
      <UrsachenBlock befunde={ursachen} />
    </div>
  );
}

// Der Ursachenbericht in der Technikerkarte: was hinter den Zahlen steht.
// Die Zahlen sagen, WORUEBER zu reden ist - die Texte sagen, WARUM. Deshalb
// steht der Block unter den Balken und ist zugeklappt: wer nur den Ueberblick
// will, wird nicht mit Kundentexten zugeschuettet; wer den Grund sucht, klickt.
// Hier wird NICHTS gerechnet und NICHTS bewertet - die Einstufung steht schon
// in der Datei, die ursachen_bericht.py geschrieben hat.
const URSACHEN_FARBE = {
  "Detraktor": { rand: "#7f1d1d", bg: "#2e0f0f", text: "#f87171" },
  "kein erfolgreicher Anruf": { rand: "#78350f", bg: "#2a1a05", text: "#fbbf24" },
  "NFT": { rand: "#78350f", bg: "#2a1a05", text: "#fb923c" },
  "Passiv": { rand: "#374151", bg: "#161b26", text: "#9ca3af" },
  "Promotor": { rand: "#14532d", bg: "#0f2e1a", text: "#4ade80" },
};
const URSACHEN_RANG = { "Detraktor": 0, "kein erfolgreicher Anruf": 1, "NFT": 2, "Passiv": 3, "Promotor": 4 };

function UrsachenBlock({ befunde }) {
  const [offen, setOffen] = useState(false);
  if (!befunde || !befunde.length) return null;
  const sortiert = [...befunde].sort(
    (a, b) => (URSACHEN_RANG[a.einstufung] ?? 9) - (URSACHEN_RANG[b.einstufung] ?? 9));
  const zaehl = (e) => befunde.filter(b => b.einstufung === e).length;
  return (
    <div style={{ marginTop: 14, borderTop: "1px solid #1f2937", paddingTop: 10 }}>
      <div onClick={() => setOffen(!offen)}
        style={{ display: "flex", alignItems: "center", gap: 8, cursor: "pointer", flexWrap: "wrap" }}>
        <span style={{ fontSize: 11, color: "#9ca3af", fontWeight: 700 }}>
          {offen ? "▾" : "▸"} Was dahintersteckt ({befunde.length})
        </span>
        {[["Detraktor", "unzufrieden"], ["NFT", "Nachfolgetickets"],
          ["kein erfolgreicher Anruf", "nicht erreicht"]].map(([e, label]) =>
          zaehl(e) ? (
            <span key={e} style={{ fontSize: 10, fontFamily: "monospace", padding: "1px 6px",
              borderRadius: 3, background: URSACHEN_FARBE[e].bg, color: URSACHEN_FARBE[e].text }}>
              {zaehl(e)} {label}
            </span>
          ) : null)}
      </div>
      {offen && sortiert.map((b, i) => {
        const f = URSACHEN_FARBE[b.einstufung] || URSACHEN_FARBE["Passiv"];
        return (
          <div key={i} style={{ marginTop: 8, background: f.bg, borderLeft: `3px solid ${f.rand}`,
            borderRadius: 4, padding: "8px 10px" }}>
            <div style={{ fontSize: 10, color: "#6b7280", fontFamily: "monospace" }}>
              {b.datum}{b.ats ? ` · ATS ${b.ats}` : ""} · {b.bereich}
              {b.auftrag ? ` · ${b.auftrag}` : ""}
              {b.wert ? ` · ${b.kennzahl}: ${b.wert}` : ""}
              <span style={{ color: f.text, fontWeight: 700, marginLeft: 6 }}>{b.einstufung}</span>
            </div>
            {[["Kunde", b.kundeSagt], ["Kunde über den Techniker", b.kundeUeber],
              ["Techniker (Abschlussvermerk)", b.technikerSagt]].map(([wer, txt]) => txt ? (
              <div key={wer} style={{ marginTop: 6, borderLeft: "2px solid #374151", paddingLeft: 8 }}>
                <div style={{ fontSize: 9, color: "#4b5563", textTransform: "uppercase", fontWeight: 700 }}>{wer}</div>
                <div style={{ fontSize: 12, color: "#d1d5db", fontStyle: "italic" }}>{txt}</div>
              </div>
            ) : null)}
            {b.weiteres ? (
              <div style={{ marginTop: 5, fontSize: 10, color: "#6b7280" }}>{b.weiteres}</div>
            ) : null}
          </div>
        );
      })}
    </div>
  );
}


// ---------------------------------------------------------------------------
// Reiter "Berichte": ein Techniker pro Seite - zum Lesen, Ausdrucken und
// Verschicken. Bewusst getrennt von der Technikerkarte: dort stehen Zahl und
// Grund nebeneinander zum Ueberblicken, hier ist es ein Blatt fuer EIN
// Gespraech. Arash geht damit zum Monteur, nicht ans Dashboard.
//
// ZUM PDF: Der Browser kann keine Datei an eine Mail haengen - das verbietet
// jeder Browser aus Sicherheitsgruenden, das ist keine Bequemlichkeit meinerseits.
// Deshalb: "Als PDF" oeffnet den Druckdialog (dort "Als PDF speichern"), die
// Mail wird getrennt vorbereitet, und der Anhang kommt per Hand dran. Zwei
// Klicks statt keiner - aber ehrlich.
// ---------------------------------------------------------------------------
function berichtText(name, meine, tech, bl) {
  // Der GANZE Bericht als Text - nicht ein Anschreiben mit Anhang.
  // Grund (Arash, 16.07.): "mach das bericht als mail fertig wenn man nichts
  // anhaengen darf". Ein Browser kann keine Datei an eine Mail haengen, also
  // wandert der Inhalt in die Mail statt daneben.
  //
  // Gebaut aus den ECHTEN Befunden, nicht aus Textbausteinen: "Sofortgespraech
  // mit der Leitstelle" sagt einem Monteur nichts, "Auftrag 200097480208 am
  // 10.07., der Router kam nicht" sagt ihm alles.
  const detr = meine.filter(b => b.einstufung === "Detraktor");
  const nft = meine.filter(b => b.einstufung === "NFT");
  const ohne = meine.filter(b => b.einstufung === "kein erfolgreicher Anruf");
  const gut = meine.filter(b => b.einstufung === "Promotor");
  const ats = [...new Set(meine.map(b => b.ats).filter(Boolean))].join(", ");
  const z = [];
  const trenner = "-".repeat(58);

  z.push(`Hallo ${name.split(" ")[0]},`);
  z.push("");
  z.push("hier deine Rueckmeldung aus dem Auftragsinfo-Portal der Telekom.");
  z.push("Ich schicke dir alles, was zu deinen Auftraegen zurueckgekommen ist -");
  z.push("das Gute und das, worueber wir reden sollten.");
  z.push("");

  // Kennzahlen mit Ampel - dieselbe Regel wie ueberall im Agenten. Ohne die
  // Zahlen fehlt der Mail der Anlass; ohne die Texte weiter unten fehlt ihr
  // der Grund. Beides gehoert in eine Mail.
  if (tech && bl) {
    const zeilen = [];
    const wort = { gut: "im Ziel", warnung: "Warnung", kritisch: "kritisch" };
    const f = (v) => (v === null || v === undefined || isNaN(v)) ? null : v.toFixed(1).replace(".", ",");
    const nimm = (label, wert, status, basis) => {
      if (wert === null) return;
      zeilen.push(`   ${label.padEnd(24)} ${String(wert).padStart(7)}` +
        (status ? `   ${wort[status]}` : "   (kein Zielwert / zu wenig Daten)") +
        (basis ? `   [Basis: ${basis}]` : ""));
    };
    nimm("Termintreue", f(tech.termintreue), tech.termintreue != null ? getStatus(tech.termintreue, bl.termintreue) : null);
    nimm("Courtesy Call", f(tech.cc_rate), tech.cc_rate != null ? getStatus(tech.cc_rate, bl.cc_rate) : null);
    nimm("NPS Problembehebung", f(tech.nps_pb), getNPSStatus(tech.nps_pb, bl.nps_pb ?? 68, tech.anzahl_nps_pb), tech.anzahl_nps_pb);
    nimm("NPS Montage", f(tech.nps_montage), getNPSStatus(tech.nps_montage, bl.nps_montage ?? 68, tech.anzahl_nps_montage), tech.anzahl_nps_montage);
    nimm("NFTQ Montage %", f(tech.nftq_m), getNFTQStatus(tech.nftq_m, bl.nftq_montage ?? 4, 8, tech.menge_m), tech.menge_m);
    nimm("NFTQ Schalten %", f(tech.nftq_s), getNFTQStatus(tech.nftq_s, bl.nftq_schalten ?? 6.6, 10, tech.menge_s), tech.menge_s);
    nimm("NFTQ Problembeh. %", f(tech.nftq_p), getNFTQStatus(tech.nftq_p, bl.nftq_pb ?? 8.5, 12, tech.menge_p), tech.menge_p);
    if (zeilen.length) {
      z.push(`DEINE ZAHLEN${ats ? ` (ATS ${ats})` : ""}`);
      z.push(trenner);
      zeilen.forEach(x => z.push(x));
      z.push("");
      z.push("   Wo kein Urteil steht, gibt es entweder keinen Telekom-Zielwert");
      z.push("   oder zu wenige Faelle, um daraus etwas abzulesen.");
      z.push("");
    }
  }

  if (gut.length) {
    z.push(`GUT GELAUFEN - ${gut.length} Kunden haben dich mit 9 oder 10 bewertet`);
    z.push(trenner);
    gut.forEach(b => {
      const t = b.kundeUeber || b.kundeSagt;
      z.push(`   ${b.datum}${t ? `: "${t}"` : ""}`);
    });
    z.push("");
  }

  if (detr.length) {
    z.push(`UNZUFRIEDENE KUNDEN - ${detr.length}`);
    z.push(trenner);
    detr.forEach(b => {
      z.push(`   ${b.datum} | Auftrag ${b.auftrag} | Bewertung ${b.wert} von 10`);
      if (b.kundeSagt) z.push(`      Kunde: "${b.kundeSagt}"`);
      if (b.kundeUeber) z.push(`      Kunde ueber dich: "${b.kundeUeber}"`);
      if (b.weiteres) z.push(`      ${b.weiteres}`);
      z.push("");
    });
  }

  if (nft.length) {
    z.push(`NACHFOLGETICKETS - ${nft.length}`);
    z.push(trenner);
    z.push("   Auftraege, bei denen nochmal jemand raus musste.");
    z.push("");
    nft.forEach(b => {
      z.push(`   ${b.datum} | Auftrag ${b.auftrag}`);
      if (b.technikerSagt) z.push(`      Dein Abschlussvermerk: "${b.technikerSagt}"`);
      if (b.weiteres) z.push(`      ${b.weiteres}`);
      z.push("");
    });
  }

  if (ohne.length) {
    z.push(`AUFTRAEGE OHNE ERFOLGREICHEN ANRUF - ${ohne.length}`);
    z.push(trenner);
    ohne.forEach(b => z.push(`   ${b.datum} | Auftrag ${b.auftrag} | ${b.wert}`));
    z.push("");
  }

  z.push(trenner);
  z.push("Lass uns kurz drueber sprechen - besonders ueber die Faelle, bei denen");
  z.push("du selbst siehst, dass es nicht an dir lag. Die will ich wissen:");
  z.push("nicht gelieferte Router und Nachfolgeauftraege, die die Telekom selbst");
  z.push("erledigt hat, zaehlen trotzdem gegen dich. Das aendere ich nur, wenn");
  z.push("ich es weiss.");
  z.push("");
  z.push("Viele Gruesse");
  return z.join("\n");
}

function BerichtTab({ ursachen, techs, baselines, kontakte }) {
  const namen = [...new Set((ursachen || []).map(u => u.name))].sort();
  const [gewaehlt, setGewaehlt] = useState(namen[0] || "");
  const [mailOffen, setMailOffen] = useState(false);
  const [kopiert, setKopiert] = useState(false);
  if (!namen.length) {
    return (
      <div style={{ background: "#111827", border: "1px solid #1f2937", borderRadius: 8, padding: 24, textAlign: "center" }}>
        <div style={{ fontSize: 13, color: "#9ca3af", fontWeight: 700 }}>Noch kein Ursachenbericht geladen</div>
        <div style={{ fontSize: 12, color: "#6b7280", marginTop: 8, lineHeight: 1.6 }}>
          Im Vikuline OS [A] druecken (oder [W] laufen lassen), dann liegt in<br />
          <code style={{ color: "#9ca3af" }}>Auftragsinfo_Downloads\Pipeline\</code> eine Datei
          <code style={{ color: "#9ca3af" }}> &lt;KW&gt;_ursachen.csv</code>.<br />
          Die hier hochladen - wie die anderen CSVs auch.
        </div>
      </div>
    );
  }
  const name = namen.includes(gewaehlt) ? gewaehlt : namen[0];
  const meine = (ursachen || []).filter(u => u.name === name);
  const tech = (techs || []).find(t => t.name === name);
  const bl = tech && String(tech.standort) === "5336" ? baselines.fs5336 : baselines.fs5335;
  const sortiert = [...meine].sort(
    (a, b) => (URSACHEN_RANG[a.einstufung] ?? 9) - (URSACHEN_RANG[b.einstufung] ?? 9));
  const zaehl = (e) => meine.filter(b => b.einstufung === e).length;
  const ats = [...new Set(meine.map(b => b.ats).filter(Boolean))].join(", ");
  const text = berichtText(name, meine, tech, bl);
  const kontakt = (kontakte || {})[name] || {};
  const betreff = `Rueckmeldung zu deinen Auftraegen${ats ? ` (ATS ${ats})` : ""}`;
  const mailto = `mailto:${kontakt.email || ""}?subject=${encodeURIComponent(betreff)}&body=${encodeURIComponent(text)}`;
  // mailto: geht als Adresse ans Mailprogramm - und Adressen sind begrenzt.
  // Outlook schneidet bei rund 2000 Zeichen ab, ANDERE MELDEN DAS NICHT: die
  // Mail geht mitten im Satz zu Ende und niemand merkt es. Deshalb wird hier
  // gerechnet statt gehofft; ist es zu lang, ist Kopieren der ehrliche Weg.
  const zuLang = mailto.length > 1900;

  return (
    <div>
      <style>{`@media print {
        body { background: #fff !important; }
        .nicht-drucken { display: none !important; }
        .druckblatt { background: #fff !important; color: #000 !important; border: none !important; }
        .druckblatt * { color: #000 !important; background: #fff !important; border-color: #999 !important; }
      }`}</style>

      <div className="nicht-drucken" style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap", alignItems: "center" }}>
        <select value={name} onChange={e => { setGewaehlt(e.target.value); setMailOffen(false); }}
          style={{ background: "#1f2937", color: "#f9fafb", border: "1px solid #374151", borderRadius: 6, padding: "7px 10px", fontSize: 12 }}>
          {namen.map(n => <option key={n} value={n}>{n}</option>)}
        </select>
        <button onClick={() => window.print()}
          style={{ background: "#1f2937", color: "#f9fafb", border: "1px solid #374151", borderRadius: 6, padding: "7px 12px", fontSize: 12, cursor: "pointer" }}>
          Drucken / Als PDF
        </button>
        <button onClick={() => setMailOffen(!mailOffen)}
          style={{ background: "#1e3a5f", color: "#dbeafe", border: "1px solid #2563eb", borderRadius: 6, padding: "7px 12px", fontSize: 12, cursor: "pointer" }}>
          Mail vorbereiten
        </button>
        <span style={{ fontSize: 10, color: "#6b7280", maxWidth: 420, lineHeight: 1.4 }}>
          Die Mail enthaelt den ganzen Bericht als Text - kein Anhang noetig.
          "Als PDF" ist fuer die Ablage oder wenn du das Blatt ausgedruckt mitnehmen willst.
        </span>
      </div>

      {mailOffen && (
        <div className="nicht-drucken" style={{ marginBottom: 14, background: "#0f1729", border: "1px solid #1e3a5f", borderRadius: 8, padding: 14 }}>
          <div style={{ fontSize: 11, color: "#93c5fd", fontWeight: 700, marginBottom: 2 }}>
            Der ganze Bericht als Mail an {name}
          </div>
          <div style={{ fontSize: 10, color: "#6b7280", marginBottom: 6 }}>
            Nichts anzuhaengen: Zahlen, Ampel und Kundentexte stehen in der Mail selbst.
          </div>
          <textarea readOnly value={text} rows={18}
            style={{ width: "100%", background: "#111827", color: "#d1d5db", border: "1px solid #374151",
              borderRadius: 6, padding: 10, fontSize: 11, fontFamily: "monospace", lineHeight: 1.5 }} />
          <div style={{ display: "flex", gap: 8, marginTop: 8, flexWrap: "wrap", alignItems: "center" }}>
            <button onClick={() => { navigator.clipboard && navigator.clipboard.writeText(text); setKopiert(true); setTimeout(() => setKopiert(false), 2000); }}
              style={{ background: zuLang ? "#2563eb" : "#1f2937", color: zuLang ? "#fff" : "#f9fafb",
                border: zuLang ? "none" : "1px solid #374151", borderRadius: 6, padding: "7px 12px", fontSize: 12, cursor: "pointer", fontWeight: zuLang ? 600 : 400 }}>
              {kopiert ? "kopiert" : "Ganzen Bericht kopieren"}
            </button>
            <a href={mailto} style={{ background: zuLang ? "#1f2937" : "#2563eb", color: zuLang ? "#9ca3af" : "#fff",
              border: zuLang ? "1px solid #374151" : "none", borderRadius: 6, padding: "7px 12px",
              fontSize: 12, textDecoration: "none" }}>In Mailprogramm oeffnen</a>
            {zuLang ? (
              <span style={{ fontSize: 10, color: "#fbbf24", maxWidth: 380, lineHeight: 1.4 }}>
                Der Bericht ist {text.length} Zeichen lang. Manche Mailprogramme schneiden
                ihn beim direkten Oeffnen ab, ohne es zu sagen. Sicherer: kopieren und
                in eine leere Mail einfuegen.
              </span>
            ) : (
              <span style={{ fontSize: 10, color: "#6b7280" }}>{text.length} Zeichen - passt.</span>
            )}
            {!kontakt.email && (
              <span style={{ fontSize: 10, color: "#fbbf24" }}>
                Keine Mailadresse hinterlegt (Knopf "Kontakte").
              </span>
            )}
          </div>
        </div>
      )}

      <div className="druckblatt" style={{ background: "#111827", border: "1px solid #1f2937", borderRadius: 8, padding: "20px 24px" }}>
        <div style={{ fontSize: 20, fontWeight: 700, color: "#f9fafb" }}>{name}</div>
        <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 14 }}>
          Ursachenbericht{ats ? ` · ATS ${ats}` : ""}{tech ? ` · ${tech.auftraege} Auftraege` : ""}
        </div>

        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 16 }}>
          {[[meine.filter(b => b.kennzahl === "NPS").length, "Bewertungen"],
            [zaehl("Detraktor"), "Unzufrieden"],
            [zaehl("NFT"), "Nachfolgetickets"],
            [zaehl("kein erfolgreicher Anruf"), "Nicht erreicht"]].map(([z, l]) => (
            <div key={l} style={{ border: "1px solid #374151", borderRadius: 6, padding: "8px 14px", minWidth: 90 }}>
              <div style={{ fontSize: 19, fontWeight: 700, color: "#f9fafb" }}>{z}</div>
              <div style={{ fontSize: 9, color: "#6b7280", textTransform: "uppercase" }}>{l}</div>
            </div>
          ))}
        </div>

        {[["Detraktor", "Unzufriedene Kunden (Bewertung 0-6)",
           "Die Zahl sagt nicht, woran es lag. Der Text schon."],
          ["NFT", "Nachfolgetickets",
           "\"Geloest durch DTA/DTS\" heisst: die Telekom hat den Nachfolgeauftrag selbst erledigt."],
          ["kein erfolgreicher Anruf", "Auftraege ohne erfolgreichen Anruf",
           "Mehrere Versuche sprechen fuer den Techniker, ein einziger wirft eine Frage auf."],
          ["Promotor", "Zufriedene Kunden", ""]].map(([e, titel, hinweis]) => {
          const liste = sortiert.filter(b => b.einstufung === e);
          if (!liste.length) return null;
          const f = URSACHEN_FARBE[e] || URSACHEN_FARBE["Passiv"];
          return (
            <div key={e} style={{ marginBottom: 18 }}>
              <div style={{ fontSize: 13, fontWeight: 700, color: "#f9fafb", borderBottom: "1px solid #374151", paddingBottom: 5, marginBottom: 4 }}>{titel}</div>
              {hinweis ? <div style={{ fontSize: 10, color: "#6b7280", marginBottom: 8 }}>{hinweis}</div> : null}
              {liste.map((b, i) => (
                <div key={i} style={{ background: f.bg, borderLeft: `3px solid ${f.rand}`, borderRadius: 4, padding: "8px 10px", marginBottom: 6 }}>
                  <div style={{ fontSize: 10, color: "#6b7280", fontFamily: "monospace" }}>
                    {b.datum}{b.ats ? ` · ATS ${b.ats}` : ""} · {b.bereich}
                    {b.auftrag ? ` · ${b.auftrag}` : ""}{b.wert ? ` · ${b.kennzahl}: ${b.wert}` : ""}
                  </div>
                  {[["Kunde", b.kundeSagt], ["Kunde über den Techniker", b.kundeUeber],
                    ["Techniker (Abschlussvermerk)", b.technikerSagt]].map(([wer, txt]) => txt ? (
                    <div key={wer} style={{ marginTop: 5, borderLeft: "2px solid #374151", paddingLeft: 8 }}>
                      <div style={{ fontSize: 9, color: "#4b5563", textTransform: "uppercase", fontWeight: 700 }}>{wer}</div>
                      <div style={{ fontSize: 12, color: "#d1d5db", fontStyle: "italic" }}>{txt}</div>
                    </div>
                  ) : null)}
                  {b.weiteres ? <div style={{ marginTop: 4, fontSize: 10, color: "#6b7280" }}>{b.weiteres}</div> : null}
                </div>
              ))}
            </div>
          );
        })}

        <div style={{ marginTop: 20, paddingTop: 10, borderTop: "1px solid #374151", fontSize: 9, color: "#4b5563" }}>
          Quelle: Auftragsinfo-Portal der Telekom. Enthaelt Kundentexte im Klartext -
          nach dem Gespraech loeschen (Telekom-Vorgabe: sobald der Zweck der Auswertung erfuellt ist).
        </div>
      </div>
    </div>
  );
}

function renderMarkdown(text) {
  return text
    .replace(/<MASSNAHMEN>[\s\S]*?<\/MASSNAHMEN>/g, "")
    .replace(/\{\s*"massnahmen"[\s\S]*?\}(?=\s*$|\s*<|\s*##)/g, "")
    .replace(/\{\s*"massnahmen"[\s\S]*$/g, "")
    .replace(/## (.*)/g, '<h3 style="color:#f9fafb;margin:20px 0 8px;font-size:14px">$1</h3>')
    .replace(/\*\*(.*?)\*\*/g, '<strong style="color:#e5e7eb">$1</strong>')
    .replace(/\n/g, "<br/>");
}

function MassnahmenPanel({ massnahmen, parseError, kontakte }) {
  if (parseError) {
    return (
      <div style={{ marginTop: 16, background: "#2e0f0f", border: "1px solid #7f1d1d", borderRadius: 8, padding: "12px 16px" }}>
        <div style={{ fontSize: 12, color: "#f87171", fontWeight: 700 }}>! Maßnahmen konnten nicht geladen werden</div>
        <div style={{ fontSize: 11, color: "#9ca3af", marginTop: 4 }}>{parseError}</div>
      </div>
    );
  }
  if (!massnahmen.length) return null;
  const statusBg = { kritisch: "#2e0f0f", warnung: "#2e1f00", gut: "#0f2e1a" };
  const statusColor = { kritisch: "#f87171", warnung: "#fbbf24", gut: "#4ade80" };
  return (
    <div style={{ marginTop: 24 }}>
      <div style={{ fontSize: 13, fontWeight: 700, color: "#f9fafb", marginBottom: 12, borderBottom: "1px solid #1f2937", paddingBottom: 8 }}> Maßnahmen pro Techniker</div>
      {massnahmen.map((m, i) => {
        const k = kontakte[m.name] || {};
        const body = m.status === "gut"
          ? `Hallo ${m.name.split(" ")[0]},\n\nwir möchten Ihnen ein herzliches Lob für Ihre hervorragenden KPI-Werte aussprechen!\n\n${m.massnahme}\n\nWeiter so - Sie sind ein wertvoller Teil unseres Teams!\n\nMit freundlichen Grüßen\n${FIRMA} Leitstelle`
          : `Hallo ${m.name.split(" ")[0]},\n\nfolgende Maßnahme wurde für Sie festgelegt:\n\n${m.massnahme}\n\nBitte bestätigen Sie die Umsetzung.\n\nMit freundlichen Grüßen\n${FIRMA} Leitstelle`;
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
                <a href={mailto} style={{ background: "#1d4ed8", color: "#fff", padding: "5px 10px", borderRadius: 5, fontSize: 11, textDecoration: "none", fontWeight: 600 }}> Email</a>
                {waLink ? <a href={waLink} target="_blank" rel="noreferrer" style={{ background: "#15803d", color: "#fff", padding: "5px 10px", borderRadius: 5, fontSize: 11, textDecoration: "none", fontWeight: 600 }}> WA</a> : null}
              </div>
            </div>
            {(!k.email && !k.mobil) && <div style={{ marginTop: 6, fontSize: 10, color: "#6b7280" }}>! Keine Kontaktdaten - unter " Kontakte" eintragen</div>}
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
  const kpiLabels = { cc_rate: "CC-Rate %", termintreue: "Termintreue %", loesungsquote: "Lösungsquote %", nps_montage: "NPS Montage", nps_pb: "NPS Problembehebung", nps: "NPS Schalten", nftq_montage: "NFTQ Montage %", nftq_schalten: "NFTQ Schalten %", nftq_pb: "NFTQ Problembeh. %", nftq_bereitstellung: "NFTQ Bereitstellung %", geplatzte_termine: "Geplatzte Termine %", info_quote_pb: "Infoquote PB %", so_quote: "SO-Quote %", service_calls: "Service Calls %" };
  const standortLabels = { gesamt: "Gesamt (KW13-19)", fs5335: "FS5335", fs5336: "FS5336" };
  const inputStyle = { background: "#1f2937", border: "1px solid #374151", borderRadius: 5, padding: "5px 8px", color: "#e5e7eb", fontSize: 12, width: "80px", textAlign: "right" };
  return (
    <div style={{ position: "fixed", top: 0, left: 0, right: 0, bottom: 0, background: "rgba(0,0,0,0.85)", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ background: "#111827", border: "1px solid #1f2937", borderRadius: 12, padding: 24, width: 620, maxHeight: "85vh", overflowY: "auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
          <span style={{ fontWeight: 700, fontSize: 15, color: "#f9fafb" }}> Baseline-Werte verwalten</span>
          <button onClick={onClose} style={{ background: "none", border: "none", color: "#6b7280", cursor: "pointer", fontSize: 18 }}>x</button>
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
                  style={{ background: "#2e0f0f", color: "#f87171", border: "1px solid #7f1d1d", borderRadius: 4, cursor: "pointer", fontSize: 11, padding: "3px 8px" }}>x</button>
              </div>
            ))}
            <AddKPIRow onAdd={(kpi, val) => setLocal(prev => ({ ...prev, [standort]: { ...prev[standort], [kpi]: val } }))} />
          </div>
        ))}
        <div style={{ display: "flex", gap: 10, marginTop: 4 }}>
          <button onClick={() => setLocal(JSON.parse(JSON.stringify(DEFAULT_BASELINES)))}
            style={{ flex: 1, background: "#1f2937", color: "#9ca3af", border: "1px solid #374151", borderRadius: 8, padding: "10px", fontSize: 12, cursor: "pointer" }}>Reset Standard</button>
          <button onClick={() => { onSave(local); onClose(); }}
            style={{ flex: 2, background: "#1d4ed8", color: "#fff", border: "none", borderRadius: 8, padding: "10px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}> Speichern</button>
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
          <span style={{ fontWeight: 700, fontSize: 15, color: "#f9fafb" }}> Techniker-Einträge verwalten</span>
          <button onClick={onClose} style={{ background: "none", border: "none", color: "#6b7280", cursor: "pointer", fontSize: 18 }}>x</button>
        </div>
        {Object.keys(local).length === 0 && <div style={{ textAlign: "center", padding: "30px", color: "#6b7280" }}>Keine Daten geladen.</div>}
        {Object.entries(local).map(([kat, rows]) => (
          <div key={kat} style={{ marginBottom: 20 }}>
            <div style={{ fontSize: 11, fontWeight: 700, color: "#60a5fa", textTransform: "uppercase", letterSpacing: 1, marginBottom: 8 }}>{kategorieLabels[kat] || kat} - {rows.length} Einträge</div>
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
                  style={{ background: "#2e0f0f", color: "#f87171", border: "1px solid #7f1d1d", borderRadius: 4, cursor: "pointer", fontSize: 11, padding: "3px 10px" }}>x Löschen</button>
              </div>
            ))}
          </div>
        ))}
        <button onClick={() => { onUpdate(local); onClose(); }}
          style={{ width: "100%", marginTop: 8, background: "#1d4ed8", color: "#fff", border: "none", borderRadius: 8, padding: "12px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}> Änderungen speichern</button>
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
          <span style={{ fontWeight: 700, fontSize: 15, color: "#f9fafb" }}> Techniker Stammdaten</span>
          <button onClick={onClose} style={{ background: "none", border: "none", color: "#6b7280", cursor: "pointer", fontSize: 18 }}>x</button>
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
            <button onClick={() => { const n = { ...local }; delete n[name]; setLocal(n); }} style={{ background: "none", border: "none", color: "#4b5563", cursor: "pointer", fontSize: 16, padding: 0 }}>x</button>
          </div>
        ))}
        <div style={{ borderTop: "1px solid #1f2937", paddingTop: 16, marginTop: 8 }}>
          <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 8 }}>+ Neuen Techniker hinzufügen:</div>
          <div style={{ display: "grid", gridTemplateColumns: "150px 1fr 130px 40px", gap: 8 }}>
            <input value={neuerName} onChange={e => setNeuerName(e.target.value)} placeholder="Vor- Nachname" style={inputStyle} onKeyDown={e => e.key === "Enter" && hinzufuegen()} />
            <input value={neuerEmail} onChange={e => setNeuerEmail(e.target.value)} placeholder="email@beispiel.de" style={inputStyle} onKeyDown={e => e.key === "Enter" && hinzufuegen()} />
            <input value={neuerMobil} onChange={e => setNeuerMobil(e.target.value)} placeholder="+4915..." style={inputStyle} onKeyDown={e => e.key === "Enter" && hinzufuegen()} />
            <button onClick={hinzufuegen} style={{ background: "#1d4ed8", color: "#fff", border: "none", borderRadius: 5, cursor: "pointer", fontSize: 16, fontWeight: 700 }}>+</button>
          </div>
        </div>
        <button onClick={() => { onSave(local); onClose(); }}
          style={{ width: "100%", marginTop: 20, background: "#1d4ed8", color: "#fff", border: "none", borderRadius: 8, padding: "12px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}> Speichern</button>
      </div>
    </div>
  );
}


function PeriodDialog({ onConfirm, onCancel }) {
  const today = new Date();
  const fmt = (d) => d.toISOString().split("T")[0];
  const [von, setVon] = React.useState(fmt(new Date(today.getTime() - 6*24*60*60*1000)));
  const [bis, setBis] = React.useState(fmt(today));

  const berechneKW = (vonStr, bisStr) => {
    const vonD = new Date(vonStr);
    const bisD = new Date(bisStr);
    const getKWNr = (d) => {
      const tmp = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
      const dayNum = tmp.getUTCDay() || 7;
      tmp.setUTCDate(tmp.getUTCDate() + 4 - dayNum);
      const yearStart = new Date(Date.UTC(tmp.getUTCFullYear(), 0, 1));
      return { kw: Math.ceil((((tmp - yearStart) / 86400000) + 1) / 7), jahr: tmp.getUTCFullYear() };
    };
    const kwVon = getKWNr(vonD);
    const kwBis = getKWNr(bisD);
    const datumVon = vonD.toLocaleDateString("de-DE", { day: "2-digit", month: "2-digit" });
    const datumBis = bisD.toLocaleDateString("de-DE", { day: "2-digit", month: "2-digit", year: "numeric" });
    if (kwVon.kw === kwBis.kw && kwVon.jahr === kwBis.jahr) {
      return { label: `KW${String(kwVon.kw).padStart(2,"0")} ${kwVon.jahr} (${datumVon} - ${datumBis})`, kw: kwVon.kw, jahr: kwVon.jahr };
    }
    return { label: `KW${String(kwVon.kw).padStart(2,"0")}-KW${String(kwBis.kw).padStart(2,"0")} ${kwBis.jahr} (${datumVon} - ${datumBis})`, kw: kwBis.kw, jahr: kwBis.jahr };
  };

  const info = berechneKW(von, bis);

  return (
    <div style={{ position: "fixed", top: 0, left: 0, right: 0, bottom: 0, background: "rgba(0,0,0,0.8)", zIndex: 400, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ background: "#111827", border: "1px solid #1f2937", borderRadius: 12, padding: 28, width: 420 }}>
        <div style={{ fontSize: 16, fontWeight: 700, color: "#f9fafb", marginBottom: 6 }}>Auswertungszeitraum</div>
        <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 20 }}>Von wann bis wann ist diese Auswertung?</div>
        
        <div style={{ display: "flex", gap: 12, marginBottom: 16 }}>
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 11, color: "#9ca3af", marginBottom: 4 }}>Von</div>
            <input type="date" value={von} onChange={e => setVon(e.target.value)}
              style={{ width: "100%", background: "#0f172a", border: "1px solid #374151", borderRadius: 6, color: "#e5e7eb", padding: "8px 10px", fontSize: 13, boxSizing: "border-box" }} />
          </div>
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 11, color: "#9ca3af", marginBottom: 4 }}>Bis</div>
            <input type="date" value={bis} onChange={e => setBis(e.target.value)}
              style={{ width: "100%", background: "#0f172a", border: "1px solid #374151", borderRadius: 6, color: "#e5e7eb", padding: "8px 10px", fontSize: 13, boxSizing: "border-box" }} />
          </div>
        </div>

        <div style={{ background: "#0f172a", borderRadius: 8, padding: "10px 14px", marginBottom: 20 }}>
          <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 2 }}>Berechnete Kalenderwoche:</div>
          <div style={{ fontSize: 14, fontWeight: 700, color: "#60a5fa" }}>{info.label}</div>
          <div style={{ fontSize: 11, color: "#4ade80", marginTop: 4 }}>Wird automatisch archiviert nach der Analyse</div>
        </div>

        <div style={{ display: "flex", gap: 10 }}>
          <button onClick={onCancel} style={{ flex: 1, background: "#1f2937", color: "#9ca3af", border: "1px solid #374151", borderRadius: 8, padding: "10px", cursor: "pointer", fontSize: 13 }}>
            Abbrechen
          </button>
          <button onClick={() => onConfirm({ von, bis, ...info })} style={{ flex: 2, background: "#2563eb", color: "#fff", border: "none", borderRadius: 8, padding: "10px", cursor: "pointer", fontSize: 13, fontWeight: 700 }}>
            Hochladen
          </button>
        </div>
      </div>
    </div>
  );
}

function VerlaufPanel({ techName, archiv, onClose }) {
  const eintraege = archiv
    .map(e => {
      const alle = Object.values(e.daten).flat();
      const tech = alle.find(t => t.name === techName);
      if (!tech) return null;
      return { label: e.label, datum: e.datum, tech };
    })
    .filter(Boolean)
    .sort((a, b) => new Date(a.datum) - new Date(b.datum));

  if (!eintraege.length) return (
    <div style={{ position: "fixed", top: 0, left: 0, right: 0, bottom: 0, background: "rgba(0,0,0,0.85)", zIndex: 300, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ background: "#111827", border: "1px solid #1f2937", borderRadius: 12, padding: 24, width: 600 }}>
        <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 16 }}>
          <span style={{ fontWeight: 700, color: "#f9fafb" }}>Verlauf: {techName}</span>
          <button onClick={onClose} style={{ background: "none", border: "none", color: "#6b7280", cursor: "pointer", fontSize: 18 }}>x</button>
        </div>
        <div style={{ color: "#6b7280", textAlign: "center", padding: 40 }}>Noch keine archivierten Daten fuer diesen Techniker.</div>
      </div>
    </div>
  );

  const kpiKeys = [
    { key: "cc_rate", label: "CC-Rate", baseline: 96, einheit: "%" },
    { key: "termintreue", label: "Termintreue", baseline: 97, einheit: "%" },
    { key: "nps", label: "NPS", baseline: 50, einheit: "" },
    { key: "a1", label: "A1-Quote", baseline: 60, einheit: "%" },
    { key: "nftq_s", label: "NFTQ-S", baseline: 4, einheit: "%", invert: true },
    { key: "nftq_p", label: "NFTQ-P", baseline: 4, einheit: "%", invert: true },
  ].filter(k => eintraege.some(e => e.tech[k.key] !== null && e.tech[k.key] !== undefined));

  return (
    <div style={{ position: "fixed", top: 0, left: 0, right: 0, bottom: 0, background: "rgba(0,0,0,0.85)", zIndex: 300, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ background: "#111827", border: "1px solid #1f2937", borderRadius: 12, padding: 24, width: 680, maxHeight: "88vh", overflowY: "auto" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 20 }}>
          <span style={{ fontWeight: 700, fontSize: 15, color: "#f9fafb" }}>Verlauf: {techName}</span>
          <button onClick={onClose} style={{ background: "none", border: "none", color: "#6b7280", cursor: "pointer", fontSize: 18 }}>x</button>
        </div>

        {kpiKeys.map(({ key, label, baseline, einheit, invert }) => (
          <div key={key} style={{ marginBottom: 20, background: "#0f172a", borderRadius: 8, padding: "14px 16px" }}>
            <div style={{ fontSize: 12, fontWeight: 700, color: "#60a5fa", marginBottom: 10 }}>{label} (Ziel: {invert ? "<=" : ">="}{baseline}{einheit})</div>
            <div style={{ display: "flex", gap: 8, alignItems: "flex-end", flexWrap: "wrap" }}>
              {eintraege.map((e, i) => {
                const val = e.tech[key];
                if (val === null || val === undefined) return null;
                const gut = invert ? val <= baseline : val >= baseline;
                const warn = invert ? val <= baseline * 2 : val >= baseline * 0.93;
                const color = gut ? "#4ade80" : warn ? "#fbbf24" : "#f87171";
                const maxVal = invert ? baseline * 3 : 100;
                const height = Math.max(20, Math.min(80, (val / maxVal) * 80));
                return (
                  <div key={i} style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 4 }}>
                    <span style={{ fontSize: 10, color, fontWeight: 700 }}>{typeof val === "number" ? val.toFixed(1) : val}{einheit}</span>
                    <div style={{ width: 32, height, background: color, borderRadius: "3px 3px 0 0", opacity: 0.8 }} />
                    <span style={{ fontSize: 9, color: "#6b7280", textAlign: "center", maxWidth: 40 }}>{e.label.split(" ")[0]}</span>
                  </div>
                );
              })}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function ArchivPanel({ archiv, onDelete, onClose }) {
  const [aufgeklappt, setAufgeklappt] = useState(null);
  const kategorieLabels = { smsfeedback: "SMS-Feedback", smsfeedbackschalten: "Schalten", nftq: "NFTQ", standard: "Manuell", onetouch: "OneTouch" };
  const exportCSV = (eintrag) => {
    const rows = Object.entries(eintrag.daten).flatMap(([kat, techs]) =>
      techs.map(t => ({ Kategorie: kategorieLabels[kat] || kat, Name: t.name, Standort: `FS${t.standort}`, CC: t.cc_rate ?? "", Termintreue: t.termintreue ?? "", Loesungsquote: t.loesungsquote ?? "", NPS: t.nps ?? "", NPS_Montage: t.nps_montage ?? "", NPS_PB: t.nps_pb ?? "", A1: t.a1 ?? "", A0: t.a0 ?? "" }))
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
          <span style={{ fontWeight: 700, fontSize: 15, color: "#f9fafb" }}> KPI-Archiv</span>
          <button onClick={onClose} style={{ background: "none", border: "none", color: "#6b7280", cursor: "pointer", fontSize: 18 }}>x</button>
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
                    {kritisch > 0 ? <span style={{ color: "#f87171", marginLeft: 8 }}>. {kritisch} kritisch</span> : null}
                    {eintrag.analyse ? <span style={{ color: "#4ade80", marginLeft: 8 }}>. KI ok</span> : null}
                  </div>
                </div>
                <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
                  <button onClick={e => { e.stopPropagation(); exportCSV(eintrag); }}
                    style={{ background: "#1f2937", color: "#9ca3af", border: "1px solid #374151", borderRadius: 4, cursor: "pointer", fontSize: 10, padding: "3px 8px" }}>CSV CSV</button>
                  <button onClick={e => { e.stopPropagation(); if (window.confirm("Eintrag löschen?")) onDelete(idx); }}
                    style={{ background: "#2e0f0f", color: "#f87171", border: "1px solid #7f1d1d", borderRadius: 4, cursor: "pointer", fontSize: 10, padding: "3px 8px" }}>x</button>
                  <span style={{ color: "#6b7280", fontSize: 14 }}>{isOpen ? "^" : "v"}</span>
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


const getTermintreeuPunkte = (prozent) => {
  if (prozent === null || prozent === undefined) return null;
  if (prozent < 80) return -200;
  if (prozent < 90) return -100;
  if (prozent < 93) return 0;
  if (prozent < 94) return 84;
  if (prozent < 95) return 105;
  if (prozent < 96) return 131;
  if (prozent < 97) return 164;
  if (prozent < 98) return 205;
  if (prozent < 99) return 256;
  if (prozent < 100) return 320;
  return 400;
};

export default function KPIAgent() {
  const [gespeichert, setGespeichert] = useState({});
  const [ursachen, setUrsachen] = useState([]);
  const [ordner, setOrdner] = useState(null);       // FileSystemDirectoryHandle
  const [ordnerLaeuft, setOrdnerLaeuft] = useState(false);
  const kannOrdner = typeof window !== "undefined" && "showDirectoryPicker" in window;
  const [kontakte, setKontakte] = useState({});
  const [baselines, setBaselines] = useState(DEFAULT_BASELINES);
  const [archiv, setArchiv] = useState([]);
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
  const [isLoggedIn, setIsLoggedIn] = useState(false);
  const [loginUser, setLoginUser] = useState("");
  const [loginPass, setLoginPass] = useState("");
  const [loginError, setLoginError] = useState("");
  const [currentUser, setCurrentUser] = useState("");
  const [minAuftraege, setMinAuftraege] = useState(1);
  const [nurKritisch, setNurKritisch] = useState(false);
  const [showVerlauf, setShowVerlauf] = useState(null);
  const [uploadPeriod, setUploadPeriod] = useState(null); // { von, bis, kw, label }
  const [showPeriodDialog, setShowPeriodDialog] = useState(false);
  const [pendingFile, setPendingFile] = useState(null);
  const dashboardRef = useRef(null);

  useEffect(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) setGespeichert(JSON.parse(saved));
      const savedK = localStorage.getItem(KONTAKTE_KEY);
      if (savedK) setKontakte(JSON.parse(savedK));
      const savedB = localStorage.getItem(BASELINE_KEY);
      if (savedB) setBaselines(JSON.parse(savedB));
      const savedA = localStorage.getItem(ARCHIV_KEY);
      if (savedA) setArchiv(JSON.parse(savedA));
      const savedU = localStorage.getItem(URSACHEN_KEY);
      if (savedU) setUrsachen(JSON.parse(savedU));
      // Gemerkten Ordner wiederfinden - aber NICHT ungefragt lesen. Der Browser
      // will die Erlaubnis je Sitzung neu bestaetigt haben, und das darf nur
      // ein Klick von Arash ausloesen, kein Seitenaufruf.
      ordnerHolen().then(h => { if (h) setOrdner(h); });
    } catch(e) {}
  }, []);

  useEffect(() => { try { localStorage.setItem(STORAGE_KEY, JSON.stringify(gespeichert)); } catch(e) {} }, [gespeichert]);
  useEffect(() => { try { localStorage.setItem(KONTAKTE_KEY, JSON.stringify(kontakte)); } catch(e) {} }, [kontakte]);
  useEffect(() => { try { localStorage.setItem(BASELINE_KEY, JSON.stringify(baselines)); } catch(e) {} }, [baselines]);
  useEffect(() => { try { localStorage.setItem(ARCHIV_KEY, JSON.stringify(archiv)); } catch(e) {} }, [archiv]);
  useEffect(() => { try { localStorage.setItem(URSACHEN_KEY, JSON.stringify(ursachen)); } catch(e) {} }, [ursachen]);

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

  // Alles aus dem verbundenen Ordner lesen. Nimmt IMMER den neuesten Zeitraum,
  // den es findet - liegen KW27 und KW28 nebeneinander, waere ein Mischen aus
  // beiden das Schlimmste, was passieren koennte: Zahlen aus zwei Wochen unter
  // einem Namen. Welcher genommen wurde, steht danach in der Meldung.
  const ausOrdnerLesen = useCallback(async (handle) => {
    setOrdnerLaeuft(true);
    setError("");
    try {
      let recht = await handle.queryPermission({ mode: "read" });
      if (recht !== "granted") recht = await handle.requestPermission({ mode: "read" });
      if (recht !== "granted") {
        setError("Zugriff auf den Ordner wurde nicht erlaubt.");
        setOrdnerLaeuft(false);
        return;
      }
      const gefunden = [];
      for await (const eintrag of dateienImOrdner(handle)) {
        const r = dateiRolle(eintrag.name);
        if (r) gefunden.push({ eintrag, ...r });
      }
      if (!gefunden.length) {
        setError("Keine passenden CSVs gefunden. Ist das der Ordner Auftragsinfo_Downloads?");
        setOrdnerLaeuft(false);
        return;
      }
      const labels = [...new Set(gefunden.map(g => g.label))].sort();
      const label = labels[labels.length - 1];
      const nehmen = gefunden.filter(g => g.label === label);

      const neu = {};
      let neueUrsachen = null;
      const namen = [];
      for (const g of nehmen) {
        const text = await (await g.eintrag.getFile()).text();
        const rows = parseCSV(text);
        if (!rows.length) continue;
        namen.push(g.eintrag.name);
        if (g.rolle === "ursachen") { neueUrsachen = rows; continue; }
        const quelle = rows[0].quelle || "standard";
        neu[quelle] = rows;
      }
      if (Object.keys(neu).length) {
        setGespeichert(prev => ({ ...prev, ...neu }));
        setAktiveKategorie("alle");
        setAiAnalysis(""); setMassnahmen([]); setMassnahmenFehler(null); setTechBewertungen({});
      }
      if (neueUrsachen) setUrsachen(neueUrsachen);
      const uebrig = labels.length > 1 ? ` (${labels.length - 1} aeltere Zeitraeume liegen daneben und wurden NICHT geladen)` : "";
      setError(`ok ${label}: ${namen.length} Dateien gelesen${uebrig}.`);
    } catch (e) {
      setError("Ordner konnte nicht gelesen werden: " + e.message);
    }
    setOrdnerLaeuft(false);
  }, []);

  const ordnerWaehlen = useCallback(async () => {
    try {
      const h = await window.showDirectoryPicker({ mode: "read" });
      setOrdner(h);
      await ordnerMerken(h);
      await ausOrdnerLesen(h);
    } catch (e) {
      if (e && e.name !== "AbortError") setError("Ordner nicht gewaehlt: " + e.message);
    }
  }, [ausOrdnerLesen]);

  const handleRows = useCallback((rows) => {
    if (!rows.length) { setError("Keine Daten gefunden."); return; }
    // Der Ursachenbericht geht einen eigenen Weg: er ersetzt keine
    // Kennzahlen-Kategorie und darf die Ansicht nicht umschalten. Er legt sich
    // neben die Zahlen und taucht in den Technikerkarten auf.
    if (rows[0] && rows[0].quelle === "ursachen") {
      setUrsachen(rows);
      setError("ok " + rows.length + " Befunde geladen - Reiter \"Berichte\" und unten in jeder Technikerkarte.");
      if (!Object.keys(gespeichert).length) setActiveTab("berichte");
      return;
    }
    const quellen = [...new Set(rows.map(r => r.quelle))];
    const quelle = quellen.length === 1 ? quellen[0] : "standard";
    if (loading) {
      setPending({ rows, quelle });
      setError("ok Gespeichert - wird nach Analyse geladen");
    } else {
      setGespeichert(prev => ({ ...prev, [quelle]: rows }));
      setAktiveKategorie(quelle);
      setAiAnalysis(""); setMassnahmen([]); setMassnahmenFehler(null); setTechBewertungen({});
      setActiveTab("dashboard"); setError("");
    }
  }, [loading, gespeichert]);

  const processXLSX = useCallback(async (file) => {
    try {
      const buffer = await file.arrayBuffer();
      const wb = XLSX.read(buffer, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      handleRows(normalizeRows(XLSX.utils.sheet_to_json(ws, { defval: "" })));
    } catch (e) { setError("Fehler: " + e.message); }
  }, [handleRows]);

  const handleImageOCR = useCallback(async (file) => {
    setLoading(true);
    setError("");
    try {
      const base64 = await new Promise((res, rej) => {
        const r = new FileReader();
        r.onload = () => res(r.result.split(",")[1]);
        r.onerror = rej;
        r.readAsDataURL(file);
      });
      const mediaType = file.type || "image/jpeg";
      const resp = await fetch("/api/analyse", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-6",
          max_tokens: 3000,
          system: `Du bist ein Datenextraktor fuer Telekom-KPI-Reports von ${FIRMA}. Extrahiere ALLE Techniker-Daten aus dem Bild als CSV.

WICHTIG: Erkenne den Report-Typ automatisch:

1. SMS-Feedback / Schalten (hat Spalten: Courtesy Calls, Termintreue, NPS, Loesungsquote):
   CSV-Format: Name,Auftraege,CC-Rate,Termintreue,NPS,Loesungsquote,Standort
   Beispiel: Ali Sodjajy,112,100.0,100.0,67,96.0,FS5335

2. NFTQ Mitarbeitersicht (hat Techniker-Namen mit Monatswerten, Titel enthaelt "NFTQ Mitarbeiter"):
   CSV-Format: Name,Auftraege,NFTQ-B,NFTQ-S,NFTQ-M,NFTQ-P,Standort
   Nimm den aktuellsten Monatswert (letzte Spalte vor Gesamtergebnis)
   Beispiel: Pejman Nazem,142,1.77,0.00,1.96,0.00,FS5335

3. OneTouch (hat A1, A2, AX, A0 Spalten):
   CSV-Format: Name,Auftraege,Tage,A1,A2,AX,A0,A-Ges,Standort

4. NFTQ Detailsicht (hat KPIName Spalte, KEINE Techniker-Namen):
   Antworte: "KEIN_TECHNIKER_REPORT - Das ist eine Firmen-Gesamtansicht ohne einzelne Techniker."

Gib NUR die CSV aus (inkl. Header-Zeile), keinen Text davor oder danach.
Standort ist FS5335 wenn nicht anders erkennbar.`,
          messages: [{ role: "user", content: [
            { type: "image", source: { type: "base64", media_type: mediaType, data: base64 } },
            { type: "text", text: "Extrahiere alle Techniker-Daten aus diesem Telekom-KPI-Screenshot als CSV. Wenn keine einzelnen Techniker erkennbar sind, schreibe KEIN_TECHNIKER_REPORT." }
          ]}]
        })
      });
      const data = await resp.json();
      const csvText = data.content?.map(b => b.text || "").join("") || "";
      if (csvText.includes("KEIN_TECHNIKER_REPORT")) {
        setError("Dieses Bild zeigt keine Techniker-Einzeldaten (z.B. Firmen-Gesamtansicht). Bitte Screenshot der Techniker-Tabelle hochladen.");
      } else if (csvText.trim()) {
        handleRows(parseCSV(csvText));
        setError("Bild erfolgreich verarbeitet!");
        setTimeout(() => setError(""), 3000);
      } else {
        setError("Bild konnte nicht verarbeitet werden. Bitte deutlicheren Screenshot versuchen.");
      }
    } catch(e) {
      setError("Fehler bei Bild-Verarbeitung: " + e.message);
    } finally {
      setLoading(false);
    }
  }, [handleRows]);

  const processFile = useCallback(async (file) => {
    // Gibt jetzt ein Versprechen zurueck, damit mehrere Dateien NACHEINANDER
    // laufen koennen. Der alte FileReader arbeitete mit Rueckrufen - darauf
    // kann man nicht warten. Fuenf Dateien waeren gleichzeitig losgelaufen und
    // haetten sich gegenseitig ueberschrieben, ohne dass es jemand merkt.
    if (file.name.match(/\.xlsx?$/i)) { await processXLSX(file); }
    else if (file.name.match(/\.(png|jpg|jpeg|gif|webp)$/i)) { await handleImageOCR(file); }
    else {
      const text = await file.text();
      handleRows(parseCSV(text));
    }
  }, [processXLSX, handleRows, handleImageOCR]);

  const handleFile = useCallback((e) => {
    // Vorher: e.target.files[0] - also genau EINE Datei je Klick. Bei fuenf
    // CSVs hiess das fuenfmal suchen, fuenfmal den Zeitraum bestaetigen.
    // Der Zeitraum wird EINMAL gefragt und gilt fuer alle: wer mehrere Dateien
    // zusammen auswaehlt, meint denselben Zeitraum. Waeren es verschiedene,
    // waere das Zusammenwerfen ohnehin falsch.
    const dateien = Array.from(e.target.files || []);
    if (!dateien.length) return;
    e.target.value = "";
    setPendingFile(dateien);
    setShowPeriodDialog(true);
  }, []);

  const angezeigt = (aktiveKategorie === "alle"
    ? (() => {
        const merged = {};
        const felder = ["cc_rate", "termintreue", "loesungsquote", "nps", "sterne", "infoquote_p",
          "geplatzte_termine", "anzahl_nps", "a_ges", "a1", "a2", "a2plus", "ax", "a0", "tage",
          "nftq_b", "nftq_s", "nftq_m", "nftq_p", "menge_b", "menge_s", "menge_m", "menge_p",
          // Ohne die zwei greifen die Mindestmengen in der Ansicht "alle" nicht:
          // die Felder wurden eingelesen, aber beim Zusammenfuehren weggelassen.
          "nps_montage", "nps_pb", "anzahl_nps_montage", "anzahl_nps_pb"];
        Object.values(gespeichert).flat().forEach(t => {
          const key = t.name + "#" + String(t.standort);
          if (!merged[key]) merged[key] = { name: t.name, standort: t.standort, quelle: "alle", auftraege: 0, standortUnbekannt: t.standortUnbekannt };
          const m = merged[key];
          felder.forEach(f => { if ((m[f] === undefined || m[f] === null) && t[f] !== undefined && t[f] !== null) m[f] = t[f]; });
          const auf = typeof t.auftraege === "number" ? t.auftraege : parseInt(t.auftraege) || 0;
          if (auf > (m.auftraege || 0)) m.auftraege = auf;
        });
        return Object.values(merged);
      })()
    : (gespeichert[aktiveKategorie] || [])
  ).filter(t => {
    const auftr = typeof t.auftraege === "number" ? t.auftraege : parseInt(t.auftraege) || 0;
    return auftr >= minAuftraege;
  });

  const hatDaten = Object.keys(gespeichert).length > 0;

  const berechneMassnahmen = (techs, bl_baselines) => {
    return techs.map(t => {
      const bl = String(t.standort) === "5336" ? bl_baselines.fs5336 : bl_baselines.fs5335;
      const sl = [];
      if (t.cc_rate !== null) sl.push(getStatus(t.cc_rate, bl.cc_rate));
      if (t.termintreue !== null) sl.push(getStatus(t.termintreue, bl.termintreue));
      if (t.loesungsquote !== null) sl.push(getStatus(t.loesungsquote, bl.loesungsquote));
      // NPS Schalten (t.nps) bleibt hier bewusst aussen vor - kein Zielwert.
      if (t.nps_montage != null) sl.push(getNPSStatus(t.nps_montage, bl.nps_montage ?? 68, t.anzahl_nps_montage));
      if (t.nps_pb != null) sl.push(getNPSStatus(t.nps_pb, bl.nps_pb ?? 68, t.anzahl_nps_pb));
      if (t.a1 !== null) sl.push(t.a1 >= 60 ? "gut" : t.a1 >= 45 ? "warnung" : "kritisch");
      if (t.a0 !== null && t.a0 > 10) sl.push("kritisch");
      // Dieselbe Rechnung wie in techWorst - EINE Stelle, sonst laufen die zwei
      // Ansichten des Agenten frueher oder spaeter auseinander.
      nftqStatusListe(t, bl).forEach(x => sl.push(x));
      // null heisst "nicht bewertet" (kein Zielwert oder zu wenig Daten) und
      // darf nicht als "gut" durchgehen.
      const bewertet = sl.filter(x => x !== null && x !== undefined);
      const worst = bewertet.length === 0 ? "gut" : bewertet.includes("kritisch") ? "kritisch" : bewertet.includes("warnung") ? "warnung" : "gut";
      // Frueher stand hier "weit ueber Zielwert 50" - ein dritter, erfundener
      // NPS-Zielwert neben den 68 aus dem Portal und den 20 in getNPSStatus.
      const lob = t.nps_pb !== null && t.nps_pb !== undefined && t.nps_pb >= (bl.nps_pb ?? 68)
        ? "Sehr gut! NPS Problembehebung " + Math.round(t.nps_pb) + " ueber Zielwert " + (bl.nps_pb ?? 68) + "."
        : t.cc_rate !== null && t.cc_rate >= 96
        ? "Sehr gute CC-Rate " + t.cc_rate.toFixed(1) + "% - Zielwert erreicht!"
        : t.nftq_b !== null && worst === "gut"
        ? "Alle NFTQ-Werte im Zielbereich - ausgezeichnete Qualitaetsarbeit!"
        : t.a1 !== null && t.a1 >= 60
        ? "Erstloesungsquote " + t.a1.toFixed(1) + "% - Zielwert erreicht!"
        : "Hervorragende Leistung! Alle KPI-Werte im Zielbereich.";
      return {
        name: t.name,
        status: worst,
        massnahme: worst === "gut" ? lob : worst === "warnung" ? "KPI-Werte beobachten und gezielt verbessern." : "Sofortgesprach mit Leitstelle - Verbesserungsmassnahmen festlegen.",
        betreff: worst === "gut" ? "Lob: Sehr gute KPI-Leistung" : "KPI Massnahme erforderlich"
      };
    });
  };

  const berechneTechScore = useCallback((tech) => {
    const v = (x) => x !== null && x !== undefined && !isNaN(x);
    const bl = String(tech.standort) === "5336" ? baselines.fs5336 : baselines.fs5335;
    const scores = [];
    if (v(tech.cc_rate)) scores.push(Math.min(10, (tech.cc_rate / bl.cc_rate) * 10));
    if (v(tech.termintreue)) scores.push(Math.min(10, (tech.termintreue / bl.termintreue) * 10));
    if (v(tech.loesungsquote)) scores.push(Math.min(10, (tech.loesungsquote / bl.loesungsquote) * 10));
    if (v(tech.nps)) scores.push(Math.min(10, Math.max(0, (tech.nps + 100) / 20)));
    if (v(tech.a1)) scores.push(Math.min(10, (tech.a1 / 60) * 10));
    if (v(tech.a_ges)) scores.push(Math.min(10, (tech.a_ges / 95) * 10));
    // NFTQ Score - niedrigere Werte sind besser, Zielwert = 10/10
    if (v(tech.nftq_b)) scores.push(Math.max(0, Math.min(10, (1 - tech.nftq_b / 20) * 10)));
    if (v(tech.nftq_s)) scores.push(Math.max(0, Math.min(10, (1 - tech.nftq_s / 20) * 10)));
    if (v(tech.nftq_m)) scores.push(Math.max(0, Math.min(10, (1 - tech.nftq_m / 20) * 10)));
    if (v(tech.nftq_p)) scores.push(Math.max(0, Math.min(10, (1 - tech.nftq_p / 20) * 10)));
    if (!scores.length) return null;
    return Math.round((scores.reduce((s, v) => s + v, 0) / scores.length) * 10) / 10;
  }, [baselines]);

  const scoreColor = (s) => !s ? "#6b7280" : s >= 8.5 ? "#4ade80" : s >= 7 ? "#fbbf24" : "#f87171";
  const scoreLabel = (s) => !s ? "-" : s >= 9 ? "Ausgezeichnet" : s >= 8 ? "Gut" : s >= 7 ? "Befriedigend" : s >= 5 ? "Verbesserungsbedarf" : "Kritisch";

  const bewerteEinzelTechniker = useCallback(async (tech) => {
    const bl = String(tech.standort) === "5336" ? baselines.fs5336 : baselines.fs5335;
    const score = berechneTechScore(tech);
    const kpiText = tech.quelle === "onetouch"
      ? `A-Gesamt=${tech.a_ges?.toFixed(1) ?? "-"}%, A1=${tech.a1?.toFixed(1) ?? "-"}%, AX=${tech.ax?.toFixed(1) ?? "-"}%, A0=${tech.a0?.toFixed(1) ?? "-"}%`
      : tech.quelle === "nftq"
      ? `NFTQ-B=${tech.nftq_b?.toFixed(2) ?? "-"}% (Ziel<=4%), NFTQ-S=${tech.nftq_s?.toFixed(2) ?? "-"}% (Ziel<=7%), NFTQ-M=${tech.nftq_m?.toFixed(2) ?? "-"}% (Ziel<=4%), NFTQ-P=${tech.nftq_p?.toFixed(2) ?? "-"}% (Ziel<=8.7%), Mengen: B=${tech.menge_b ?? "-"} S=${tech.menge_s ?? "-"} M=${tech.menge_m ?? "-"} P=${tech.menge_p ?? "-"}`
      : `CC=${tech.cc_rate?.toFixed(1) ?? "-"}% (Ziel>=95%), Termintreue=${tech.termintreue?.toFixed(1) ?? "-"}% (Ziel>=96%), Loesungsquote=${tech.loesungsquote?.toFixed(1) ?? "-"}%, NPS-Montage=${tech.nps_montage?.toFixed(0) ?? "-"}, NPS-Problembehebung=${tech.nps_pb?.toFixed(0) ?? "-"}, NPS-Schalten=${tech.nps?.toFixed(0) ?? "-"} (NPS-Ziel jeweils >=67)`;
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
      if (t.quelle === "onetouch") return `${t.name} (FS${t.standort}): A-Ges=${t.a_ges?.toFixed(1) ?? "-"}%, A1=${t.a1?.toFixed(1) ?? "-"}%, AX=${t.ax?.toFixed(1) ?? "-"}%, A0=${t.a0?.toFixed(1) ?? "-"}%, Aufträge=${t.auftraege}`;
      if (t.quelle === "nftq") return `${t.name} (FS${t.standort}): NFTQ-B=${t.nftq_b?.toFixed(2) ?? "-"}%, NFTQ-S=${t.nftq_s?.toFixed(2) ?? "-"}%, NFTQ-M=${t.nftq_m?.toFixed(2) ?? "-"}%, NFTQ-P=${t.nftq_p?.toFixed(2) ?? "-"}%, Aufträge=${t.auftraege}`;
      return `${t.name} (FS${t.standort}): CC=${t.cc_rate?.toFixed(1) ?? "-"}%, Termintreue=${t.termintreue?.toFixed(1) ?? "-"}%, Lösungsquote=${t.loesungsquote?.toFixed(1) ?? "-"}%, NPS-Montage=${t.nps_montage?.toFixed(0) ?? "-"}, NPS-PB=${t.nps_pb?.toFixed(0) ?? "-"}, NPS-Schalten=${t.nps?.toFixed(0) ?? "-"}, Aufträge=${t.auftraege}`;
    }).join("\n");
    // Letzte archivierte KW fuer Vergleich
    let vorperiodeStr = "";
    if (archiv.length > 0) {
      const letzteKW = archiv[archiv.length - 1];
      const vorTechs = Object.values(letzteKW.daten).flat();
      vorperiodeStr = "\n\nVorperiode (" + letzteKW.label + "):\n" + vorTechs.map(t => {
        if (t.quelle === "onetouch") return `${t.name} (FS${t.standort}): A1=${t.a1?.toFixed(1) ?? "-"}%, A-Ges=${t.a_ges?.toFixed(1) ?? "-"}%`;
        if (t.quelle === "nftq") return `${t.name} (FS${t.standort}): NFTQ-B=${t.nftq_b?.toFixed(2) ?? "-"}%, NFTQ-S=${t.nftq_s?.toFixed(2) ?? "-"}%, NFTQ-M=${t.nftq_m?.toFixed(2) ?? "-"}%, NFTQ-P=${t.nftq_p?.toFixed(2) ?? "-"}%`;
        return `${t.name}: CC=${t.cc_rate?.toFixed(1) ?? "-"}%, TT=${t.termintreue?.toFixed(1) ?? "-"}%, NPS-M=${t.nps_montage?.toFixed(0) ?? "-"}, NPS-PB=${t.nps_pb?.toFixed(0) ?? "-"}, NPS-S=${t.nps?.toFixed(0) ?? "-"}`;
      }).join("\n");
    }

    try {
      const res = await fetch("/api/analyse", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ model: "claude-sonnet-4-6", max_tokens: 2000, system: SYSTEM_PROMPT_FN(baselines), messages: [{ role: "user", content: `Analysiere diese Techniker-KPIs:\n\n${dataStr}${vorperiodeStr}\n\nBitte zeige bei jedem Techniker den Trend zur Vorperiode mit Pfeil (gestiegen/gesunken/gleich) und Delta.` }] }),
      });
      const data = await res.json();
      const text = data.content?.map(b => b.text || "").join("") || "";
      setAiAnalysis(text);

      // Massnahmen fuer jeden Techniker berechnen
      const alleMassnahmen = angezeigt.map(t => {
        const bl = String(t.standort) === "5336" ? (baselines.fs5336 || DEFAULT_BASELINES.fs5336) : (baselines.fs5335 || DEFAULT_BASELINES.fs5335);
        const statusList = [];
        const v = (x) => x !== null && x !== undefined && !isNaN(x);
        if (v(t.cc_rate) && bl.cc_rate) statusList.push(getStatus(t.cc_rate, bl.cc_rate));
        if (v(t.termintreue) && bl.termintreue) statusList.push(getStatus(t.termintreue, bl.termintreue));
        if (v(t.loesungsquote) && bl.loesungsquote) statusList.push(getStatus(t.loesungsquote, bl.loesungsquote));
        // NPS Schalten (t.nps) bleibt aussen vor - kein Portal-Zielwert.
        if (v(t.nps_montage)) statusList.push(getNPSStatus(t.nps_montage, bl.nps_montage ?? 68, t.anzahl_nps_montage));
        if (v(t.nps_pb)) statusList.push(getNPSStatus(t.nps_pb, bl.nps_pb ?? 68, t.anzahl_nps_pb));
        if (v(t.a1)) statusList.push(t.a1 >= 60 ? "gut" : t.a1 >= 45 ? "warnung" : "kritisch");
        if (v(t.a0) && t.a0 > 10) statusList.push("kritisch");
        // DRITTE Stelle, die frueher selbst gerechnet hat - und mit WIEDER
        // ANDEREN Schwellen als die zwei oben: Schalten 7 statt 6,6,
        // Problembehebung 8,7 statt 8,5. Derselbe Agent gab also je nach
        // Ansicht verschiedene Ampeln aus. Jetzt rechnet nftqStatusListe()
        // an allen drei Stellen - eine Regel, ein Ergebnis.
        nftqStatusListe(t, bl).forEach(x => statusList.push(x));
        const bewertet2 = statusList.filter(x => x !== null && x !== undefined);
        const worst = bewertet2.length === 0 ? "gut" : bewertet2.includes("kritisch") ? "kritisch" : bewertet2.includes("warnung") ? "warnung" : "gut";
        const nps_val = t.nps !== null ? t.nps : 0;
        const cc_val = t.cc_rate !== null ? t.cc_rate : 0;
        const a1_val = t.a1 !== null ? t.a1 : 0;
        const lob = t.quelle === "nftq" && worst === "gut"
          ? "Alle bewerteten NFTQ-Werte im Zielbereich - ausgezeichnete Qualitaetsarbeit!"
          : v(t.nps_pb) && t.nps_pb >= (bl.nps_pb ?? 68)
          ? "Ausgezeichnet! NPS Problembehebung " + t.nps_pb.toFixed(0) + " ueber Zielwert " + (bl.nps_pb ?? 68) + " - Vorbild im Team!"
          : t.cc_rate !== null && t.cc_rate >= 96
          ? "Sehr gute CC-Rate " + cc_val.toFixed(1) + "% - Zielwert 96% erreicht!"
          : t.a1 !== null && t.a1 >= 60
          ? "Erstloesungsquote " + a1_val.toFixed(1) + "% - Zielwert 60% erreicht!"
          : "Hervorragende Leistung! Alle KPI-Werte im Zielbereich - weiter so!";
        return {
          name: t.name,
          status: worst,
          massnahme: worst === "gut" ? lob : worst === "warnung" ? "KPI-Werte beobachten und gezieltes Coaching einleiten." : "Sofortgespraech mit Leitstelle erforderlich - Verbesserungsmassnahmen festlegen.",
          betreff: worst === "gut" ? "Lob: Sehr gute KPI-Leistung" : worst === "warnung" ? "KPI Verbesserung" : "Dringend: KPI kritisch"
        };
      });
      setMassnahmen(alleMassnahmen);
      setMassnahmenFehler(null);

      setActiveTab("analyse");
      // Auto-Archivierung nach Analyse
      if (uploadPeriod) {
        const archivEintrag = {
          label: uploadPeriod.label,
          datum: new Date().toISOString(),
          daten: { ...gespeichert }
        };
        setArchiv(prev => {
          const exists = prev.find(a => a.label === uploadPeriod.label);
          if (exists) return prev.map(a => a.label === uploadPeriod.label ? archivEintrag : a);
          return [...prev, archivEintrag];
        });
      }
    } catch (e) { setError("Fehler bei der KI-Analyse."); }
    finally { setLoading(false); }
  };

  const archivieren = useCallback(() => {
    if (!hatDaten) return;
    const now = new Date();
    const { kw, jahr } = getKW(now);
    const kwInput = window.prompt(
      "Kalenderwoche eingeben (Format: KW25 2026 oder KW25/26 2026):",
      `KW${String(kw).padStart(2, "0")} ${jahr}`
    );
    if (kwInput === null) return; // Abgebrochen
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
            t.nps_montage != null ? getNPSStatus(t.nps_montage) : null,
            t.nps_pb != null ? getNPSStatus(t.nps_pb) : null,
          ].filter(Boolean);
          status = statuses.includes("kritisch") ? "kritisch" : statuses.includes("warnung") ? "warnung" : "gut";
        }
        return { ...t, _status: status, _score: berechneTechScore(t) };
      });
    });
    const datum = now.toLocaleDateString("de-DE", { day: "2-digit", month: "2-digit", year: "numeric" });
    const archivLabel = kwInput.trim() ? `${kwInput.trim()} · ${datum}` : formatArchivLabel(now);
    setArchiv(prev => [...prev, {
      label: archivLabel, datum: now.toISOString(),
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

  const criticalCount = angezeigt.filter(t => techWorst(t, baselines) === "kritisch").length;

  const avg = (key) => {
    const vals = angezeigt.map(t => t[key]).filter(v => v !== null && !isNaN(v));
    return vals.length ? (vals.reduce((s, v) => s + v, 0) / vals.length).toFixed(1) : "-";
  };

  const teamAvgScore = () => {
    const scores = angezeigt.map(t => berechneTechScore(t)).filter(s => s !== null && !isNaN(s));
    return scores.length ? (scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(1) : "-";
  };

  const isOTView = aktiveKategorie === "onetouch";

  const FirmendashboardTab = () => {
    if (!angezeigt || !angezeigt.length) return null;
    const sorted = [...angezeigt].filter(t => !nurKritisch || techWorst(t, baselines) === "kritisch").sort((a, b) => (berechneTechScore(b) || 0) - (berechneTechScore(a) || 0));
    return (
      <div>
        <div style={{ background: "#0f172a", border: "1px solid #1f2937", borderRadius: 8, padding: "16px", marginBottom: 20 }}>
          <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: 1 }}>Team-Übersicht</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12 }}>
            {[
              { label: "Avg Team-Score", value: teamAvgScore() + "/10", color: scoreColor(parseFloat(teamAvgScore())) },
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
            ? `Hallo ${tech.name.split(" ")[0]},\n\nhier ist Ihre persönliche KPI-Bewertung:\n\nScore: ${score}/10 - ${scoreLabel(score)}\n\n${bew.kommentar}\n\n${bew.staerken?.length ? `Stärken:\n${bew.staerken.map(s => `• ${s}`).join("\n")}\n\n` : ""}${bew.schwaechen?.length ? `Verbesserungsbedarf:\n${bew.schwaechen.map(s => `• ${s}`).join("\n")}\n\n` : ""}Maßnahme: ${bew.massnahme || ""}\n\nMit freundlichen Grüßen\n${FIRMA} Leitstelle`
            : "";
          const mailto = `mailto:${k.email || ""}?subject=${encodeURIComponent(`KPI-Bewertung ${tech.name}`)}&body=${encodeURIComponent(mailBody)}`;
          return (
            <div key={tech.name} style={{ background: "#111827", border: `1px solid ${score >= 8 ? "#14532d" : score >= 6 ? "#78350f" : "#7f1d1d"}`, borderRadius: 8, padding: "14px 16px", marginBottom: 10 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <div style={{ fontSize: 16, fontWeight: 700, color: scoreColor(score), fontFamily: "monospace", minWidth: 28 }}>#{i + 1}</div>
                  <div>
                    <div style={{ fontWeight: 700, fontSize: 14, color: "#f9fafb" }}>{tech.name}</div>
                    <div style={{ fontSize: 11, color: "#6b7280" }}>FS{tech.standort} . {tech.auftraege} Aufträge</div>
                  </div>
                </div>
                {score !== null && (
                  <div style={{ textAlign: "center" }}>
                    <div style={{ fontSize: 22, fontWeight: 700, color: scoreColor(score), fontFamily: "monospace" }}>{score}</div>
                    <div style={{ fontSize: 9, color: scoreColor(score) }}>/10 . {scoreLabel(score)}</div>
                  </div>
                )}
              </div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 10 }}>
                {tech.cc_rate !== null && tech.cc_rate !== undefined && (() => { const z = (String(tech.standort)==="5336"?baselines.fs5336:baselines.fs5335).cc_rate; const s = getStatus(tech.cc_rate, z); const c = s==="kritisch"?"#f87171":s==="warnung"?"#fbbf24":"#4ade80"; return <span style={{ fontSize: 10, background: "#1f2937", color: c, padding: "2px 8px", borderRadius: 3, fontWeight: 700 }}>CC {tech.cc_rate.toFixed(1)}% / &gt;={z}%</span>; })()}
                {tech.termintreue !== null && tech.termintreue !== undefined && (() => { const z = (String(tech.standort)==="5336"?baselines.fs5336:baselines.fs5335).termintreue; const s = getStatus(tech.termintreue, z); const c = s==="kritisch"?"#f87171":s==="warnung"?"#fbbf24":"#4ade80"; return <span style={{ fontSize: 10, background: "#1f2937", color: c, padding: "2px 8px", borderRadius: 3, fontWeight: 700 }}>TT {tech.termintreue.toFixed(1)}% / &gt;={z}%</span>; })()}
                {tech.loesungsquote !== null && tech.loesungsquote !== undefined && (() => { const z = (String(tech.standort)==="5336"?baselines.fs5336:baselines.fs5335).loesungsquote; const s = getStatus(tech.loesungsquote, z); const c = s==="kritisch"?"#f87171":s==="warnung"?"#fbbf24":"#4ade80"; return <span style={{ fontSize: 10, background: "#1f2937", color: c, padding: "2px 8px", borderRadius: 3, fontWeight: 700 }}>LQ {tech.loesungsquote.toFixed(1)}% / &gt;={z}%</span>; })()}
                {(() => { const b = String(tech.standort)==="5336"?baselines.fs5336:baselines.fs5335; const kand = [[tech.nps, b.nps ?? 67], [tech.nps_pb, b.nps_pb ?? 67], [tech.nps_montage, b.nps_montage ?? 67]].find(([v]) => v !== null && v !== undefined && !isNaN(v)); if (!kand) return null; const [nv, z] = kand; const s = getNPSStatus(nv, z); const c = s==="kritisch"?"#f87171":s==="warnung"?"#fbbf24":"#4ade80"; return <span style={{ fontSize: 10, background: "#1f2937", color: c, padding: "2px 8px", borderRadius: 3, fontWeight: 700 }}>NPS {nv.toFixed(0)} / &gt;={z}</span>; })()}
                {tech.a1 !== null && tech.a1 !== undefined && (() => { const c = tech.a1>=60?"#4ade80":tech.a1>=45?"#fbbf24":"#f87171"; return <span style={{ fontSize: 10, background: "#1f2937", color: c, padding: "2px 8px", borderRadius: 3, fontWeight: 700 }}>A1 {tech.a1.toFixed(1)}% / &gt;={OT_BASELINE.a1}%</span>; })()}
                {tech.nftq_b !== null && tech.nftq_b !== undefined && (() => { const c = tech.nftq_b<=4?"#4ade80":tech.nftq_b<=8?"#fbbf24":"#f87171"; return <span style={{ fontSize: 10, background: "#1f2937", color: c, padding: "2px 8px", borderRadius: 3, fontWeight: 700 }}>NFTQ-B {tech.nftq_b.toFixed(1)}%</span>; })()}
                {tech.nftq_s !== null && tech.nftq_s !== undefined && (() => { const c = tech.nftq_s<=4?"#4ade80":tech.nftq_s<=8?"#fbbf24":"#f87171"; return <span style={{ fontSize: 10, background: "#1f2937", color: c, padding: "2px 8px", borderRadius: 3, fontWeight: 700 }}>NFTQ-S {tech.nftq_s.toFixed(1)}%</span>; })()}
                {tech.nftq_m !== null && tech.nftq_m !== undefined && (() => { const c = tech.nftq_m<=4?"#4ade80":tech.nftq_m<=8?"#fbbf24":"#f87171"; return <span style={{ fontSize: 10, background: "#1f2937", color: c, padding: "2px 8px", borderRadius: 3, fontWeight: 700 }}>NFTQ-M {tech.nftq_m.toFixed(1)}%</span>; })()}
                {tech.nftq_p !== null && tech.nftq_p !== undefined && (() => { const c = tech.nftq_p<=4?"#4ade80":tech.nftq_p<=8?"#fbbf24":"#f87171"; return <span style={{ fontSize: 10, background: "#1f2937", color: c, padding: "2px 8px", borderRadius: 3, fontWeight: 700 }}>NFTQ-P {tech.nftq_p.toFixed(1)}%</span>; })()}
                {tech.a0 !== null && tech.a0 !== undefined && tech.a0 > 0 && <span style={{ fontSize: 10, background: "#2e0f0f", color: "#f87171", padding: "2px 8px", borderRadius: 3, fontWeight: 700 }}>A0 {tech.a0.toFixed(1)}%</span>}
              </div>
              {isLoadingThis && <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 8 }}>... KI bewertet...</div>}
              {bew && !isLoadingThis && (
                <div style={{ background: "#0f172a", borderRadius: 6, padding: "10px 12px", marginBottom: 10 }}>
                  <div style={{ fontSize: 12, color: "#d1d5db", lineHeight: 1.6, marginBottom: 8 }}>{bew.kommentar}</div>
                  <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
                    {bew.staerken?.map((s, si) => <span key={si} style={{ fontSize: 10, color: "#4ade80" }}>ok {s}</span>)}
                    {bew.schwaechen?.map((s, si) => <span key={si} style={{ fontSize: 10, color: "#f87171" }}>x {s}</span>)}
                  </div>
                  {bew.massnahme && <div style={{ fontSize: 11, color: "#fbbf24", marginTop: 6 }}>Massnahme: {bew.massnahme}</div>}
                </div>
              )}
              <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                {!bew && !isLoadingThis && (
                  <button onClick={() => bewerteEinzelTechniker(tech)}
                    style={{ background: "#1d4ed8", color: "#fff", border: "none", borderRadius: 5, cursor: "pointer", fontSize: 11, padding: "5px 12px", fontWeight: 600 }}> Bewerten</button>
                )}
                {bew && k.email && (
                  <a href={mailto} style={{ background: "#1d4ed8", color: "#fff", padding: "5px 12px", borderRadius: 5, fontSize: 11, textDecoration: "none", fontWeight: 600 }}> Mail senden</a>
                )}
                {bew && !k.email && <span style={{ fontSize: 10, color: "#6b7280" }}>! Keine Email - unter  eintragen</span>}
                {archiv.length > 0 && <button onClick={() => setShowVerlauf(tech.name)} style={{ background: "#0f172a", color: "#60a5fa", border: "1px solid #1e3a5f", borderRadius: 5, cursor: "pointer", fontSize: 11, padding: "5px 12px" }}>Verlauf</button>}
                {k.mobil && bew && (
                  <a href={`https://wa.me/${k.mobil.replace(/[^0-9]/g, "")}?text=${encodeURIComponent(`Hallo ${tech.name.split(" ")[0]}, Ihr KPI-Score: ${score}/10 - ${scoreLabel(score)}. ${bew?.massnahme || ""}`)}`}
                    target="_blank" rel="noreferrer"
                    style={{ background: "#15803d", color: "#fff", padding: "5px 12px", borderRadius: 5, fontSize: 11, textDecoration: "none", fontWeight: 600 }}> WhatsApp</a>
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
      {showPeriodDialog && <PeriodDialog
        onConfirm={(period) => {
          setUploadPeriod(period);
          setShowPeriodDialog(false);
          if (pendingFile) {
            // Nacheinander, nicht gleichzeitig: die Verarbeitung schreibt in
            // denselben Zustand, und zwei Dateien parallel wuerden sich
            // gegenseitig ueberschreiben - leise, ohne Fehlermeldung.
            const liste = Array.isArray(pendingFile) ? pendingFile : [pendingFile];
            (async () => {
              for (const f of liste) { await processFile(f); }
              if (liste.length > 1) setError(`ok ${liste.length} Dateien gelesen.`);
            })();
            setPendingFile(null);
          }
        }}
        onCancel={() => { setShowPeriodDialog(false); setPendingFile(null); }}
      />}
      {showVerlauf && <VerlaufPanel techName={showVerlauf} archiv={archiv} onClose={() => setShowVerlauf(null)} />}
      {showArchiv && <ArchivPanel archiv={archiv} onDelete={(idx) => setArchiv(prev => prev.filter((_, i) => i !== idx))} onClose={() => setShowArchiv(false)} />}

      {/* KPI Warnungsleiste */}
      {(() => {
        const alleTechs = Object.values(gespeichert).flat().filter((t, idx, arr) => arr.findLastIndex(x => x.name === t.name && String(x.standort) === String(t.standort)) === idx);
        const kritisch = alleTechs.filter(t => t.overallStatus === "kritisch");
        const warnung = alleTechs.filter(t => t.overallStatus === "warnung");
        if (kritisch.length === 0 && warnung.length === 0) return null;
        return (
          <div style={{ background: kritisch.length > 0 ? "#2e0f0f" : "#2e1f00", borderBottom: `1px solid ${kritisch.length > 0 ? "#7f1d1d" : "#92400e"}`, padding: "8px 24px", display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
            <span style={{ fontSize: 11, fontWeight: 700, color: kritisch.length > 0 ? "#f87171" : "#fbbf24", textTransform: "uppercase", letterSpacing: 1 }}>
              {kritisch.length > 0 ? `⚠ ${kritisch.length} KRITISCH` : `⚡ ${warnung.length} WARNUNG`}
            </span>
            {kritisch.map(t => (
              <span key={t.name} style={{ background: "#450a0a", border: "1px solid #7f1d1d", borderRadius: 4, padding: "2px 8px", fontSize: 11, color: "#f87171" }}>{t.name}</span>
            ))}
            {warnung.map(t => (
              <span key={t.name} style={{ background: "#431407", border: "1px solid #92400e", borderRadius: 4, padding: "2px 8px", fontSize: 11, color: "#fbbf24" }}>{t.name}</span>
            ))}
          </div>
        );
      })()}
            <div style={{ borderBottom: "1px solid #1f2937", padding: "12px 24px", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 8 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" }}>
          <div style={{ width: 8, height: 8, borderRadius: "50%", background: "#4ade80", boxShadow: "0 0 6px #4ade80", flexShrink: 0 }} />
          <button onClick={() => window.location.reload()} style={{ background: "none", border: "none", cursor: "pointer", padding: 0 }}>
            <span style={{ fontWeight: 700, fontSize: 13, letterSpacing: 2, color: "#9ca3af", textTransform: "uppercase" }}>KPI Agent </span>
          </button>
          <span style={{ color: "#374151" }}> · {FIRMA}</span>
          <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
            {KATEGORIEN.map(k => {
              const anzahl = k.id === "alle"
                ? Object.values(gespeichert).flat().filter((t, idx, arr) => arr.findLastIndex(x => x.name === t.name && String(x.standort) === String(t.standort)) === idx).length
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
            {/* Der Ursachenbericht ist KEINE Kennzahlen-Kategorie - man kann ihn
                nicht "ansehen" wie SMS-Feedback, er steht in den Technikerkarten
                und im Reiter "Berichte". Trotzdem gehoert er hierher: Arash hat
                zweimal gefragt, ob die Datei ueberhaupt drin ist. Die Meldung
                beim Hochladen ist weg, sobald man woanders hinklickt - das hier
                bleibt. Deshalb kein Knopf, sondern ein Merkzettel. */}
            {ursachen.length > 0 && (
              <span title="Ursachenbericht geladen - steht im Reiter Berichte und in den Technikerkarten"
                style={{ padding: "3px 10px", borderRadius: 5, fontSize: 11,
                  background: "#0f2e1a", color: "#4ade80", border: "1px solid #14532d" }}>
                Ursachen ({new Set(ursachen.map(u => u.name)).size} Techniker, {ursachen.length} Befunde)
              </span>
            )}
          </div>
        </div>
        <div style={{ display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" }}>
          <button onClick={() => { const n = window.prompt("Firmenname (erscheint in Kopfzeile, Mails und KI-Analyse):", FIRMA); if (n && n.trim()) { localStorage.setItem("firma_name", n.trim()); window.location.reload(); } }} style={{ background: "#111827", color: "#9ca3af", border: "1px solid #374151", padding: "6px 12px", borderRadius: 6, cursor: "pointer", fontSize: 11 }}> Firma</button>
          <button onClick={() => setShowKontakte(true)} style={{ background: "#111827", color: "#9ca3af", border: "1px solid #374151", padding: "6px 12px", borderRadius: 6, cursor: "pointer", fontSize: 11 }}> Kontakte</button>
          <button onClick={() => setShowBaseline(true)} style={{ background: "#111827", color: "#9ca3af", border: "1px solid #374151", padding: "6px 12px", borderRadius: 6, cursor: "pointer", fontSize: 11 }}> Baselines</button>
          <button onClick={() => setShowTechVerwaltung(true)} style={{ background: "#111827", color: "#9ca3af", border: "1px solid #374151", padding: "6px 12px", borderRadius: 6, cursor: "pointer", fontSize: 11 }}> Techniker</button>
          <button onClick={() => setShowArchiv(true)} style={{ background: "#111827", color: "#9ca3af", border: "1px solid #374151", padding: "6px 12px", borderRadius: 6, cursor: "pointer", fontSize: 11 }}> Archiv{archiv.length > 0 ? ` (${archiv.length})` : ""}</button>
          {uploadPeriod && <div style={{ fontSize: 11, color: "#60a5fa", background: "#0f172a", border: "1px solid #1e3a5f", borderRadius: 6, padding: "4px 10px" }}>{uploadPeriod.label}</div>}
          <div style={{ display: "flex", alignItems: "center", gap: 4, background: "#111827", border: "1px solid #374151", borderRadius: 6, padding: "4px 8px" }}>
            <span style={{ fontSize: 10, color: "#6b7280" }}>Min.</span>
            <input type="number" min="1" max="50" value={minAuftraege}
              onChange={e => setMinAuftraege(parseInt(e.target.value) || 1)}
              style={{ width: 36, background: "transparent", border: "none", color: "#e5e7eb", fontSize: 11, textAlign: "center", outline: "none" }} />
            <span style={{ fontSize: 10, color: "#6b7280" }}>Auftr.</span>
          </div>
          <button onClick={() => setNurKritisch(v => !v)} title="Zwischen allen und nur kritischen Technikern umschalten"
            style={{ background: nurKritisch ? "#7f1d1d" : "#111827", color: nurKritisch ? "#fecaca" : "#9ca3af", border: `1px solid ${nurKritisch ? "#b91c1c" : "#374151"}`, padding: "6px 12px", borderRadius: 6, cursor: "pointer", fontSize: 11, fontWeight: 600 }}>
            {nurKritisch ? "Nur kritische" : "Alle"}
          </button>
          {/* Ordner verbinden: EIN Klick statt fuenf Dateien einzeln suchen.
              Der Knopf erscheint nur in Chrome/Edge - Firefox und Safari
              koennen die File System Access API nicht, und ein Knopf, der
              nichts tut, ist schlimmer als keiner. */}
          {kannOrdner && (
            ordner ? (
              <button onClick={() => ausOrdnerLesen(ordner)} disabled={ordnerLaeuft}
                title="Liest die CSVs aus dem verbundenen Ordner neu ein - immer den neuesten Zeitraum"
                style={{ background: ordnerLaeuft ? "#1a2e1a" : "#0f2e1a", color: "#4ade80",
                  border: "1px solid #14532d", padding: "6px 12px", borderRadius: 6,
                  cursor: ordnerLaeuft ? "default" : "pointer", fontSize: 11, fontWeight: 600 }}>
                {ordnerLaeuft ? "liest..." : "Aus Ordner aktualisieren"}
              </button>
            ) : (
              <button onClick={ordnerWaehlen} disabled={ordnerLaeuft}
                title="Einmal den Ordner Auftragsinfo_Downloads auswaehlen - danach holt der Agent die CSVs selbst. Die Daten bleiben auf deinem Rechner."
                style={{ background: "#1e3a5f", color: "#dbeafe", border: "1px solid #2563eb",
                  padding: "6px 12px", borderRadius: 6, cursor: "pointer", fontSize: 11, fontWeight: 600 }}>
                Ordner verbinden
              </button>
            )
          )}
          <label style={{ background: loading ? "#1a2e1a" : "#1f2937", color: loading ? "#4ade80" : "#9ca3af", border: `1px solid ${loading ? "#14532d" : "#374151"}`, padding: "6px 12px", borderRadius: 6, cursor: "pointer", fontSize: 11 }}>
            {loading ? "... Nächste" : " Upload"}
            <input type="file" multiple accept=".csv,.xlsx,.xls,.png,.jpg,.jpeg,.pdf" onChange={handleFile} style={{ display: "none" }} />
          </label>
          {angezeigt.length > 0 && <button onClick={exportPDF} disabled={exporting} style={{ background: "#1f2937", color: "#9ca3af", border: "1px solid #374151", padding: "6px 12px", borderRadius: 6, cursor: "pointer", fontSize: 11 }}> PDF</button>}
          {aktiveKategorie !== "alle" && gespeichert[aktiveKategorie] && <button onClick={() => loescheKategorie(aktiveKategorie)} style={{ background: "#2e0f0f", color: "#f87171", border: "1px solid #7f1d1d", padding: "6px 12px", borderRadius: 6, cursor: "pointer", fontSize: 11, fontWeight: 700 }}> {KATEGORIEN.find(k => k.id === aktiveKategorie)?.label} löschen</button>}
          <button onClick={async () => { await fetch("/api/logout", { method: "POST" }); window.location.href = "/login"; }} style={{ background: "#2e0f0f", color: "#f87171", border: "1px solid #7f1d1d", padding: "6px 12px", borderRadius: 6, cursor: "pointer", fontSize: 11 }}>Logout</button>
        </div>
      </div>

      <div ref={dashboardRef} style={{ maxWidth: 720, margin: "0 auto", padding: "24px 20px" }}>
        {/* Frueher stand hier nur !hatDaten - also "keine Kennzahlen". Damit hat
            dieser Bildschirm auch die Berichte verdeckt, obwohl die geladen waren:
            Arash hatte 67 Befunde drin und sah trotzdem ueberall "hochladen".
            Der Ursachenbericht braucht die Kennzahlen aber gar nicht. */}
        {!hatDaten && ursachen.length === 0 && (
          <div style={{ textAlign: "center", padding: "60px 20px" }}>
            <div style={{ fontSize: 36, marginBottom: 12 }}></div>
            <div style={{ fontSize: 18, fontWeight: 700, color: "#f9fafb", marginBottom: 8 }}>Telekom-Export hochladen</div>
            <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 24 }}>Wird automatisch der richtigen Kategorie zugeordnet</div>
            <label style={{ display: "inline-block", background: "#1d4ed8", color: "#fff", padding: "10px 24px", borderRadius: 6, cursor: "pointer", fontSize: 13, fontWeight: 600 }}>
               Datei wählen (.csv / .xlsx)
              <input type="file" multiple accept=".csv,.xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
            </label>
            {error ? <div style={{ marginTop: 16, color: "#4ade80", fontSize: 13 }}>{error}</div> : null}
          </div>
        )}

        {hatDaten && angezeigt.length === 0 && ursachen.length === 0 && (
          <div style={{ textAlign: "center", padding: "40px 20px" }}>
            <div style={{ fontSize: 14, color: "#6b7280", marginBottom: 16 }}>Noch keine Daten für "{KATEGORIEN.find(k => k.id === aktiveKategorie)?.label}"</div>
            <label style={{ display: "inline-block", background: "#1d4ed8", color: "#fff", padding: "10px 24px", borderRadius: 6, cursor: "pointer", fontSize: 13, fontWeight: 600 }}>
               Export hochladen
              <input type="file" multiple accept=".csv,.xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
            </label>
          </div>
        )}

        {(angezeigt.length > 0 || ursachen.length > 0) && (
          <>
            {pending && <div style={{ fontSize: 11, color: "#4ade80", marginBottom: 8 }}>... Nächste Datei bereit</div>}
            {error && <div style={{ fontSize: 11, color: "#4ade80", marginBottom: 8 }}>{error}</div>}

            <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 10, marginBottom: 20 }}>
              {(isOTView ? [
                { label: "Techniker", value: angezeigt.length, color: "#60a5fa" },
                { label: "Kritisch", value: criticalCount, color: criticalCount > 0 ? "#f87171" : "#4ade80" },
                { label: "Avg A1-Rate", value: avg("a1") !== "-" ? avg("a1") + "%" : "-", color: "#4ade80" },
                { label: "Avg Score", value: teamAvgScore() + "/10", color: scoreColor(parseFloat(teamAvgScore())) },
              ] : [
                { label: "Techniker", value: angezeigt.length, color: "#60a5fa" },
                { label: "Kritisch", value: criticalCount, color: criticalCount > 0 ? "#f87171" : "#4ade80" },
                { label: "Avg CC-Rate", value: avg("cc_rate") !== "-" ? avg("cc_rate") + "%" : "-", color: "#fbbf24" },
                { label: "Avg Score", value: teamAvgScore() + "/10", color: scoreColor(parseFloat(teamAvgScore())) },
              ]).map(s => {
                const klickbar = s.label === "Techniker" || s.label === "Kritisch";
                const zielFilter = s.label === "Kritisch";
                const aktiv = klickbar && zielFilter === nurKritisch;
                return (
                  <div key={s.label} onClick={klickbar ? () => setNurKritisch(zielFilter) : undefined}
                    title={klickbar ? (zielFilter ? "Nur kritische Techniker anzeigen" : "Alle Techniker anzeigen") : undefined}
                    style={{ background: "#111827", border: `1px solid ${aktiv ? "#3b82f6" : "#1f2937"}`, borderRadius: 8, padding: "12px 14px", cursor: klickbar ? "pointer" : "default" }}>
                    <div style={{ fontSize: 11, color: "#6b7280", marginBottom: 4 }}>{s.label}</div>
                    <div style={{ fontSize: 20, fontWeight: 700, color: s.color, fontFamily: "monospace" }}>{s.value}</div>
                  </div>
                );
              })}
            </div>

            <div style={{ display: "flex", marginBottom: 16, borderBottom: "1px solid #1f2937" }}>
              {[
                { id: "dashboard", label: "Dashboard" },
                { id: "firmendashboard", label: ` Firmendashboard${Object.keys(techBewertungen).length > 0 ? ` (${Object.keys(techBewertungen).length})` : ""}` },
                { id: "analyse", label: "KI-Analyse" + (aiAnalysis ? " ok" : "") },
                { id: "berichte", label: "Berichte" + (ursachen.length ? ` (${new Set(ursachen.map(u => u.name)).size})` : "") },
              ].map(tab => (
                <button key={tab.id} onClick={() => setActiveTab(tab.id)}
                  style={{ background: "none", border: "none", borderBottom: activeTab === tab.id ? "2px solid #3b82f6" : "2px solid transparent", color: activeTab === tab.id ? "#f9fafb" : "#6b7280", padding: "8px 14px", cursor: "pointer", fontSize: 12, fontWeight: activeTab === tab.id ? 600 : 400, marginBottom: -1, whiteSpace: "nowrap" }}>
                  {tab.label}
                </button>
              ))}
            </div>

            {activeTab === "dashboard" && (
              <>
                <div style={{ marginBottom: 16 }}>{angezeigt.filter(t => !nurKritisch || techWorst(t, baselines) === "kritisch").map((t, i) => {
                  const vorTechs = archiv.length > 0 ? Object.values(archiv[archiv.length-1].daten).flat() : [];
                  const vorperiode = vorTechs.find(v => v.name === t.name) || null;
                  return <TechCard key={i} tech={t} baselines={baselines} vorperiode={vorperiode}
                    ursachen={(ursachen || []).filter(u => u.name === t.name)} />;
                })}</div>
                <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                  <button onClick={runAnalysis} disabled={loading}
                    style={{ width: "100%", background: loading ? "#1f2937" : "#1d4ed8", color: loading ? "#6b7280" : "#fff", border: "none", borderRadius: 8, padding: "14px", fontSize: 14, fontWeight: 700, cursor: loading ? "not-allowed" : "pointer" }}>
                    {loading ? "... KI analysiert..." : " Team-Analyse starten"}
                  </button>
                  <button onClick={bewerteAlle}
                    style={{ width: "100%", background: "#0f172a", color: "#60a5fa", border: "1px solid #1e3a5f", borderRadius: 8, padding: "12px", fontSize: 13, fontWeight: 700, cursor: "pointer" }}>
                     Alle bewerten - Firmendashboard
                  </button>
                  <button onClick={archivieren}
                    style={{ width: "100%", background: "#0f172a", color: "#6b7280", border: "1px solid #1f2937", borderRadius: 8, padding: "10px", fontSize: 12, cursor: "pointer" }}>
                     Archivieren & Dashboard leeren
                  </button>
                </div>
              </>
            )}

            {activeTab === "berichte" && (
              <BerichtTab ursachen={ursachen} techs={angezeigt} baselines={baselines}
                kontakte={kontakte} />
            )}
            {activeTab === "firmendashboard" && <FirmendashboardTab />}

            {activeTab === "analyse" && (
              <div>
                {!aiAnalysis && !loading && <div style={{ textAlign: "center", padding: "40px", color: "#6b7280" }}>Noch keine Analyse. Dashboard öffnen und starten.</div>}
                {loading && <div style={{ textAlign: "center", padding: "40px", color: "#6b7280" }}>... KI analysiert...</div>}
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
