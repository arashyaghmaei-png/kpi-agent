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
// Die Uebersicht aus [U] - Ampel, Ziel und Schwelle je Kennzahl. Eigener
// Speicher, weil sie KEINE Kennzahlen-Kategorie ist: sie ersetzt keinen Report
// und darf nicht als Reiter auftauchen. Sie legt sich neben die Zahlen und
// sagt, wie sie zu lesen sind.
const UEBERSICHT_KEY = "fibernc_uebersicht";

// ---------------------------------------------------------------------------
// WIE ALT SIND DIE GESPEICHERTEN DATEN?
// Die Zahlen liegen fertig eingelesen im Browser. Repariere ich das EINLESEN,
// wirkt das NICHT rueckwirkend - was drin liegt, bleibt falsch, bis es neu
// hochgeladen wird.
// Passiert am 16.07.2026 genau so: "0 Rueckmeldungen" wurde beim Einlesen zu
// "unbekannt" (parseInt("0") || null). Nach der Reparatur zeigte der Agent bei
// Alae@36 weiter KRITISCH - die Datei war richtig, der Speicher nicht. Arash
// konnte das nicht wissen.
// Deshalb: Wird das Einlesen geaendert, DATEN_VERSION hochzaehlen. Dann sagt
// der Agent es von selbst, statt still falsche Zahlen zu zeigen.
// ---------------------------------------------------------------------------
const VERSION_KEY = "fibernc_daten_version";
const DATEN_VERSION = 3;   // 3 = seit dem Umbau zum Fenster (17.07.2026):
                           // parseCSV/normalizeRows haben sich geaendert, und der
                           // Agent rechnet nicht mehr selbst. Was vor dem Umbau
                           // im Browser lag, ist nach anderen Regeln eingelesen
                           // worden - deshalb hochzaehlen, sonst zeigt er still
                           // alte Zahlen. 2 = Reparatur "0 bleibt 0" (16.07.).

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
  // Die Uebersicht aus [U]. Bis 17.07.2026 stand sie hier ausdruecklich auf der
  // Ignorieren-Liste - der Agent rechnete ja selbst. Jetzt ist sie die Quelle
  // fuer JEDE Ampel und JEDEN Zielwert; ohne sie zeigt der Agent Zahlen ohne
  // Urteil.
  { endung: "_kpi_uebersicht.csv", rolle: "uebersicht" },
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
// OT_BASELINE = { a_ges: 95.0, a1: 60.0 } ist raus - beide Zahlen erfunden,
// One Touch steht in keinem Telekom-Papier mit Zielwert.
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

// HIER STANDEN BIS 17.07.2026 DOCH NOCH ZIELWERTE - getarnt als "Baseline":
//
//   Baseline KW13-19: CC=${bl.gesamt.cc_rate}% | Termintreue=${...} | ...
//   KW20 schlechteste Woche (NPS 26, Termintreue 85,7%). KW23-24 FS5336 ...
//
// Zwei Probleme auf einmal. ERSTENS waren das gar keine Baselines, sondern
// DEFAULT_BASELINES - also die Zielwerte des Baseline-Editors (cc_rate: 95
// "Courtesy Calls Zielwert"). Die KI las "Lösungsquote=95%" und schrieb
// "Lösungsquote ✅ Ziel >= 95%". Damit war der Prompt die letzte Stelle, an
// der der Editor noch urteilte - und mit veralteten Zahlen (NPS 68 statt
// 67,1 laut Vertrag, NFTQ-S 6,6 statt 7,0).
// ZWEITENS: KW13-19 und KW20 liegen VOR dem 01.07.2026. Vikuline faehrt erst
// seit diesem Tag fuer die Telekom - diese Wochen koennen keine Zahlen der
// Firma sein. Sie stammen aus dem Vorgaenger-Agenten. Die KI bekam also die
// Historie einer anderen Firma als Massstab fuer Arashs Monteure.
//
// Der Parameter bl bleibt in der Signatur, damit der Aufruf unveraendert
// bleibt - benutzt wird er nicht mehr.
const SYSTEM_PROMPT_FN = (bl) => `Du bist ein operativer KPI-Analyseagent für ein Telekommunikations-Subunternehmen (Telekom-Subunternehmer, Kupfer & FTTH, Bergheim NRW).
Vikuline faehrt erst seit dem 01.07.2026 fuer die Telekom. Es gibt KEINE aeltere
Historie und keine Baseline - fehlt eine Vorperiode, sag das und lass den Trend weg.
OneTouch: A1=erster Besuch erledigt, AX=Abbruch, A0=nicht erledigt. KEIN Zielwert - Werte nennen, nicht bewerten.
Aufgabe: Techniker-KPIs bewerten, Leitstellen-Empfehlungen. Wenn Vorperioden-Daten vorhanden, Trend je Techniker angeben.

ZIELWERTE UND AMPELN STEHEN NICHT IN DIESEM PROMPT.

Bis 17.07.2026 standen sie hier - und liefen mit der Excel auseinander. Die
Analyse schrieb dann "NFTQ-Bereitstellung 18,18% Kritisch (Ziel <=4%)", wo die
Excel gar kein Urteil hatte. Ein Prompt ist eine Kopie wie jede andere.

Jede Kennzahl kommt dir unter "## KPI-Uebersicht ([U])" FERTIG BEURTEILT zu:
wert, menge, ampel, ziel, schwelle, grund. Gerechnet hat kpi_uebersicht.py
gegen die Zusatzvereinbarung Bonus-Malus (gueltig ab 01.04.2026).

REGELN:
- Uebernimm die Ampel, wie sie dasteht. Rechne nichts nach.
- ampel = null heisst NICHT "gut". Es heisst: kein Urteil. Der Grund steht in
  "grund" - schreib ihn hin, statt selbst zu entscheiden.
- Nenne IMMER die Basis: "22,2 % = 4 von 18". Eine Quote ohne Menge ist wertlos.
- ERFINDE KEINE ZIELWERTE. Steht bei einer Kennzahl kein "ziel", dann gibt es
  keinen - schreibe "kein Zielwert bekannt" und beurteile sie nicht.
- Kennzahlen ohne Zielwert: NFTQ Bereitstellung (= Schalten + Montage, zaehlt
  dieselben Nachfolgetickets zweimal), NPS Schalten, One Touch (A1/A2/AX/A0),
  Sterne. Wert nennen, kein Urteil.
- "NFT" heisst Nachfolgeticket: ein Auftrag, bei dem nochmal jemand raus musste.
  Schreib das Wort aus.
- Fehlt die KPI-Uebersicht ganz, sag das und beurteile NICHTS.

MINDESTMENGEN (Vikuline-Regel, keine Telekom-Vorgabe - so aber sagen):
- NFTQ erst ab 10 Auftraegen in DER Kategorie bewerten. Darunter Wert nennen,
  kein Urteil: bei 3 Schalten-Auftraegen sind 33,3% EIN Nachfolgeticket - das
  misst den Zufall, nicht die Arbeit.
- NPS erst ab 2 Rueckmeldungen bewerten. Bei einer ist der Wert nur +100 oder -100.
- NENNE BEI JEDER QUOTE DIE BASIS: "NFTQ Montage 22,2% (4 von 18)", nicht nur
  "22,2%". Eine Quote ohne ihre Basis ist im Gespraech mit dem Monteur wertlos.
- Bei einem Extremwert aus wenigen Faellen: sag ausdruecklich, dass er auf
  wenigen Faellen beruht, statt ihn als Befund zu verkaufen.
- Die Basis steht in den Daten hinter jedem Wert. Steht dort "unbekannt", dann
  ist sie UNBEKANNT - schreibe das so und erfinde keine Zahl. "unbekannt" ist
  NICHT dasselbe wie "0 Rueckmeldungen": bei 0 gibt es nichts zu bewerten, bei
  unbekannt weisst du es schlicht nicht.
- Ein Wert "-" heisst: fuer diese Kennzahl wurde nichts geliefert. Das ist kein
  schlechter Wert und keine Null - dazu sagst du gar nichts.

ABKUERZUNGEN - RATE SIE NICHT. Du hast am 17.07.2026 "NFTQ Portiern" und
"NPS Privatbereitstellung" geschrieben. Beides heisst PROBLEMBEHEBUNG. Solche
Texte gehen per Mail an den Monteur.
- NFT = Nachfolgeticket (ein Auftrag, bei dem nochmal jemand raus musste)
- NFTQ-M = Nachfolgetickets Montage
- NFTQ-S = Nachfolgetickets Schalten
- NFTQ-P = Nachfolgetickets PROBLEMBEHEBUNG (nicht Portieren)
- NFTQ-B = Nachfolgetickets Bereitstellung (= Schalten + Montage zusammen)
- NPS BS / NPS Montage = Kundenzufriedenheit bei Montage/Bereitstellung
- NPS PB = Kundenzufriedenheit bei PROBLEMBEHEBUNG (nicht Privatbereitstellung)
- CC = Courtesy Call (Rueckruf nach dem Termin)
- A1/AX/A0 = One Touch: beim ersten Besuch erledigt / abgebrochen / nicht erledigt
Steht eine Abkuerzung nicht in dieser Liste, schreibe sie so hin, wie sie
dasteht - erfinde keine Bedeutung.

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

// "0" muss 0 bleiben. Vorher stand ueberall parseInt(...) || null - und in
// JavaScript ist 0 || null gleich null. Damit verschwand die Null: Adil Kheder
// hatte in KW28 NULL Montage-Rueckmeldungen (Anzahl NPS BS = 0), das Portal
// schreibt dann NPS BS = 0,00 - und der Agent machte daraus "NPS Montage 0,
// KRITISCH". Ein Mann ohne eine einzige Bewertung stand als kritisch da.
// Die Mindestmengen-Pruefung konnte nicht greifen, weil sie die 0 nie sah.
// Der One-Touch-Report schreibt "Kheder Adil", die anderen "Adil Kheder" -
// derselbe Mann, zwei Schreibweisen. Der Agent hat daraus ZWEI Techniker
// gemacht: 16 statt 8. Deshalb wird zum Zusammenfuehren nicht der Name
// verglichen, sondern die SORTIERTEN Namensteile. So passt jede Reihenfolge
// zusammen. (Im Ursachenbericht steht dieselbe Regel - ohne sie haetten die
// One-Touch-Befunde nie zu einem Techniker gefunden.)
function namensSchluessel(name) {
  return String(name || "")
    .toLowerCase()
    .replace(/[^a-zäöüß\s-]/g, " ")
    .split(/[\s-]+/)
    .filter(Boolean)
    .sort()
    .join(" ");
}

function ganzzahl(v) {
  const n = parseInt(String(v ?? "").trim(), 10);
  return isNaN(n) ? null : n;   // 0 bleibt 0, nur echter Unsinn wird null
}

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
  // Die kpi_uebersicht.csv MUSS vor "smsfeedback" geprueft werden: sie hat eine
  // Spalte "NPS PB", und die Pruefung unten wuerde sie dafuer halten. Erkennbar
  // ist sie an den Ampel-Spalten - die hat sonst keine Datei.
  if (h.some(x => x.startsWith("ampel "))) return "uebersicht";
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
  // Die Uebersicht aus [U] wird DURCHGEREICHT, nicht verrechnet. Genau das ist
  // der Punkt: [U] hat gerechnet, der Agent liest ab. Aus jeder Zeile wird ein
  // Eintrag je Kennzahl - Wert, Menge, Ampel, Ziel, Schwelle nebeneinander, so
  // wie sie in der Datei stehen.
  if (fmt === "uebersicht") {
    const f = (row, name) => {
      const k = Object.keys(row).find(x => cleanHeader(x).toLowerCase() === name.toLowerCase());
      return k ? String(row[k] ?? "").trim() : "";
    };
    // "1.234,5" -> 1234.5 ; "" -> null. Die CSV kommt mit Dezimalkomma.
    const z = (v) => {
      const t = String(v ?? "").replace(/\./g, "").replace(",", ".").trim();
      if (t === "" || t === "-") return null;
      const n = parseFloat(t);
      return isNaN(n) ? null : n;
    };
    // [U] schreibt seit 17.07.2026 WOERTER statt eines Strichs. Drei davon sind
    // kein Urteil, sondern ein Grund - der wird hier in Klartext uebersetzt.
    const urteil = (wort, mindest) => {
      const w = String(wort || "").trim().toUpperCase();
      if (w === "GUT" || w === "WARNUNG" || w === "KRITISCH") return { ampel: w.toLowerCase(), grund: null };
      if (w === "KEIN_ZIEL") return { ampel: null, grund: "kein Telekom-Zielwert - Wert ohne Urteil" };
      if (w === "ZU_WENIG_DATEN") return { ampel: null, grund: mindest ? `unter ${mindest} - daraus laesst sich nichts ablesen` : "zu wenig Daten fuer ein Urteil" };
      if (w === "KEINE_DATEN") return { ampel: null, grund: "keine Rueckmeldungen" };
      return { ampel: null, grund: null };   // auch der alte "-" landet hier
    };

    const KENNZAHLEN = [
      ["termintreue", "Termintreue", null, "Termintreue", null],
      ["cc", "CC", null, "CC", null],
      ["nps_montage", "NPS Montage", "Menge NPS Montage", "NPS Montage", "nps"],
      ["nps_pb", "NPS PB", "Menge NPS PB", "NPS PB", "nps"],
      ["nps_schalten", "NPS Schalten", "Menge NPS Schalten", "NPS Schalten", "nps"],
      // NFTQ-B hat KEINE eigene Mengenspalte - B = Bereitstellung = Schalten +
      // Montage, die Basis ist also die Summe der beiden. (Kheder KW28:
      // 4 + 18 = 22 Auftraege, 4 NFT -> 18,18 %. Stuende hier nur die
      // Montage-Menge, waere die Basis 18 und die Prozentzahl daneben gelogen.)
      ["nftq_b", "NFTQ-B", "@s+m", "NFTQ-B", "nftq"],
      ["nftq_s", "NFTQ-S", "Menge NFTQ-S", "NFTQ-S", "nftq"],
      ["nftq_m", "NFTQ-M", "Menge NFTQ-M", "NFTQ-M", "nftq"],
      ["nftq_p", "NFTQ-P", "Menge NFTQ-P", "NFTQ-P", "nftq"],
    ];

    return rawRows
      .filter(row => f(row, "techniker"))
      .map(row => {
        const mind = { nps: z(f(row, "Mindestmenge NPS")), nftq: z(f(row, "Mindestmenge NFTQ")) };
        const kpi = {};
        KENNZAHLEN.forEach(([id, wSp, mSp, suffix, mindArt]) => {
          const u = urteil(f(row, "Ampel " + suffix), mindArt ? mind[mindArt] : null);
          const menge = mSp === "@s+m"
            ? (z(f(row, "Menge NFTQ-S")) ?? 0) + (z(f(row, "Menge NFTQ-M")) ?? 0)
            : (mSp ? z(f(row, mSp)) : null);
          kpi[id] = {
            wert: z(f(row, wSp)),
            menge,
            ampel: u.ampel,
            grund: u.grund,
            ziel: z(f(row, "Ziel " + suffix)),
            schwelle: z(f(row, "Schwelle " + suffix)),
          };
        });
        return {
          quelle: "uebersicht",
          name: f(row, "Techniker"),
          ats: f(row, "ATS"),
          auftraege: z(f(row, "Auftraege")),
          sterne: z(f(row, "Sterne")),
          status: f(row, "Status").toLowerCase(),   // GUT/WARNUNG/KRITISCH/-
          kpi,
        };
      });
  }

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
          anzahl_nps: ganzzahl(get(row, "anzahl nps gesamt")),
          nps_montage: npsZahl(get(row, "nps bs")),
          nps_pb: npsZahl(get(row, "nps pb")),
          anzahl_nps_montage: ganzzahl(get(row, "anzahl nps bs")),
          anzahl_nps_pb: ganzzahl(get(row, "anzahl nps pb")),
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
          anzahl_nps: ganzzahl(get(row, "anzahl nps")),
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
          menge_b: ganzzahl(get(row, "bereitstellung")),
          menge_s: ganzzahl(get(row, "schalten")),
          menge_m: ganzzahl(get(row, "montage")),
          menge_p: ganzzahl(get(row, "problembehebung")),
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
// MINDEST_NFTQ und MINDEST_NPS standen hier bis 17.07.2026 - "diese zwei
// Zahlen MUESSEN mit kpi_uebersicht.py uebereinstimmen", stand als Kommentar
// dabei. Genau das ist das Problem gewesen: zwei Stellen, die jemand von Hand
// gleich halten muss. Jetzt entscheidet [U], ob eine Menge reicht, und schreibt
// "ZU_WENIG_DATEN" in die Ampel - samt Mindestmenge in einer eigenen Spalte.

// ===========================================================================
// AB HIER RECHNET DER AGENT NICHT MEHR.
//
// Bis 17.07.2026 hatte er NEUN Stellen mit eigenen Regeln - techWorst,
// berechneMassnahmen, die Massnahmen-Ansicht, TechCard, den KI-Prompt, NFTQBar,
// das Firmendashboard, getOTStatus und getStatus. Jede mit eigenen Zahlen, und
// sie liefen auseinander: Am 17.07. sagte die Excel zu Adil Kheder
// "ZU_WENIG_DATEN" (eine einzige Rueckmeldung), der Agent machte ihn zum
// "Vorbild im Team". Gleiche Datei, gleicher Mann, gleiche Woche.
//
// Jetzt gibt es genau eine Quelle: Pipeline\<KW>_kpi_uebersicht.csv aus [U].
// Dort steht je Kennzahl Wert, Menge, Ampel, Zielwert und Schwellwert. Der
// Agent liest ab. Wer hier eine Bedingung braucht, aendert kpi_uebersicht.py -
// und NICHT diese Datei.
// ===========================================================================

// Die Uebersicht-Zeile zu einem Techniker. Schluessel wie ueberall:
// sortierte Namensteile + ATS, damit "Adil Kheder" und "Kheder Adil"
// zusammenfinden (siehe namensSchluessel).
function uebersichtIndex(zeilen) {
  const m = new Map();
  (zeilen || []).forEach(z => {
    const ats = String(z.ats || "").trim();
    const st = ats === "36" ? "5336" : ats === "35" ? "5335" : ats;
    m.set(namensSchluessel(z.name) + "#" + st, z);
  });
  return m;
}

// EINE Kennzahl eines Technikers, so wie [U] sie beurteilt hat.
// Liefert immer ein Objekt - nie undefined -, damit die Aufrufer nicht jedes
// Mal pruefen muessen. Fehlt die Uebersicht, ist ampel null und grund sagt,
// warum: dann zeigt der Agent Zahlen ohne Urteil, statt sich eines auszudenken.
const KEINE_UEBERSICHT = { wert: null, menge: null, ampel: null, ziel: null,
  schwelle: null, grund: "keine Uebersicht geladen - kpi_uebersicht.csv fehlt" };

// Haengt einem Techniker seine Uebersicht-Zeile an. Alles danach liest nur
// noch t.u - deshalb steht dieser Schritt an EINER Stelle und nicht verteilt.
function mitUebersicht(t, index) {
  const st = String(t.standort || "");
  return { ...t, u: index.get(namensSchluessel(t.name) + "#" + st) || null };
}

function kpiAusU(t, kennzahl) {
  const k = t && t.u && t.u.kpi && t.u.kpi[kennzahl];
  return k || KEINE_UEBERSICHT;
}

// Nur echte Urteile. null heisst "nicht bewertet" und darf NIE als "gut"
// durchgehen - genau daran ist der Agent frueher gescheitert.
function ampelU(t, kennzahl) {
  return kpiAusU(t, kennzahl).ampel;
}

// Die schlechteste Ampel des Technikers - aus den Urteilen, die [U] gefaellt
// hat. Nicht bewertete Kennzahlen (kein Zielwert, zu wenig Daten, keine Daten)
// zaehlen NICHT mit; sie duerfen weder nach "gut" noch nach "kritisch" kippen.
// Hat [U] zu keiner einzigen Kennzahl ein Urteil, gibt es auch keins: null.
//
// One Touch steht bewusst nicht mehr drin. Der Agent hat A1/A0/A Ges. frueher
// nach selbst erfundenen Schwellen gefaerbt (a0 > 10 kritisch, a1 < 60 warnung).
// In der Zusatzvereinbarung Bonus-Malus kommt One Touch NICHT vor, in keiner
// Telekom-Folie steht ein Zielwert dafuer. Also wird es angezeigt, nicht
// beurteilt.
const BEWERTETE_KENNZAHLEN = ["termintreue", "cc", "nps_montage", "nps_pb",
  "nftq_s", "nftq_m", "nftq_p"];

function techWorst(t) {
  const rang = { kritisch: 0, warnung: 1, gut: 2 };
  const urteile = BEWERTETE_KENNZAHLEN.map(k => ampelU(t, k)).filter(x => x !== null);
  if (!urteile.length) return null;
  return urteile.sort((a, b) => rang[a] - rang[b])[0];
}

// Die drei bewerteten NFTQ-Kategorien. NFTQ-B fehlt hier weiterhin, und jetzt
// steht der Grund in der Datei: B = Schalten + Montage, Telekom hat dafuer
// keinen Zielwert - [U] schreibt "KEIN_ZIEL", und ampelU liefert null.
function nftqStatusListe(t) {
  return ["nftq_s", "nftq_m", "nftq_p"].map(k => ampelU(t, k)).filter(x => x !== null);
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


// SECHSTE Stelle, die selbst geurteilt hat. Vorher stand hier
// "const z = ziel || 4; const w = warn || 8" - und weil NIEMAND ziel/warn
// uebergeben hat, faerbte der Balken alle vier Kategorien nach dem pauschalen
// 4/8. Auch die Bereitstellung, die wir ueberall sonst nicht bewerten: Adil
// Kheder hatte einen knallroten Balken bei 18,18 % fuer eine Kennzahl, die
// dieselben Nachfolgetickets doppelt zaehlt.
// Jetzt rechnet der Balken NICHT mehr - er bekommt den Status fertig geliefert.
// ziel kommt seit 17.07.2026 mit: Der Balken zeigt an, wogegen gemessen wird.
// Vorher hatte NFTQBar "const z = ziel || 4" - und NIEMAND uebergab ziel, also
// faerbte er alle vier Kategorien nach dem pauschalen 4/8. Jetzt rechnet er gar
// nicht mehr, er bekommt das Urteil fertig geliefert.

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

// KPI-KACHEL (17.07.2026, Arashs Entwurf: "Zielwert Telekom oben, berechneter
// Wert der Techniker unten").
//
// Vier Zeilen, immer dieselben, fuer JEDE Kennzahl:
//   1 Name
//   2 was die Telekom vorgibt  - Ziel und der gelbe Bereich bis zum Schwellwert
//   3 was wir gerechnet haben  - der Wert
//   4 worauf er beruht         - die Menge, und das Urteil
//
// Warum die Trennung wichtig ist: Zeile 2 ist fremd (Zusatzvereinbarung
// Bonus-Malus), Zeile 3 ist unsere Rechnung. Wer die beiden nicht auseinander
// halten kann, streitet mit dem Monteur ueber die falsche Zahl.
//
// Der Schwellwert war bisher UNSICHTBAR: [U] schreibt ihn in die CSV, der
// Agent hat ihn stumm weggeworfen. Ein gelbes Feld konnte deshalb niemand
// erklaeren. Alae@35 hat NFT Problembehebung 9,09 % bei einem Schwellwert von
// 9,1 - das sieht man erst, wenn beide Zahlen dastehen.
//
// Ohne Zielwert bleibt die Kachel grau und sagt es. Ohne Urteil auch. Grau
// heisst nie "gut".
function KPIKachel({ label, k, hoch, trend, einheit }) {
  if ((k.wert === null || k.wert === undefined || isNaN(k.wert)) && !k.grund) return null;
  const st = k.ampel;
  const c = st === "kritisch" ? "#f87171" : st === "warnung" ? "#fbbf24" : st === "gut" ? "#4ade80" : "#9ca3af";
  const bg = st === "kritisch" ? "#1f1113" : st === "warnung" ? "#1f1a0f" : st === "gut" ? "#0f1a12" : "#161c26";
  const e = einheit === undefined ? "%" : einheit;
  const vgl = hoch ? ">=" : "<=";
  const hatZiel = k.ziel !== null && k.ziel !== undefined && k.ziel !== "";
  const hatSchwelle = k.schwelle !== null && k.schwelle !== undefined && k.schwelle !== "";
  const wort = st === "kritisch" ? "KRITISCH" : st === "warnung" ? "WARNUNG" : st === "gut" ? "im Ziel" : "kein Urteil";
  return (
    <div style={{ background: bg, borderRadius: 6, padding: "8px 10px" }}>
      <div style={{ fontSize: 11, color: "#9ca3af", marginBottom: 1 }}>{label}</div>
      <div style={{ fontSize: 10, color: "#6b7280" }}>
        {hatZiel
          ? <>Ziel {vgl} {k.ziel}{e}{hatSchwelle ? ` . gelb bis ${k.schwelle}${e}` : ""}</>
          : "kein Telekom-Zielwert"}
      </div>
      <div style={{ display: "flex", alignItems: "baseline", gap: 6, marginTop: 3 }}>
        <span style={{ fontSize: 17, fontWeight: 700, fontFamily: "monospace", color: c }}>
          {k.wert === null || k.wert === undefined || isNaN(k.wert) ? "-" : k.wert.toFixed(1) + e}
        </span>
        {trend && <span style={{ color: trend.color, fontSize: 10, fontWeight: 700 }}>{trend.symbol}{e}</span>}
      </div>
      <div style={{ fontSize: 10, color: st ? c : "#6b7280", marginTop: 1 }}>
        {k.menge !== null && k.menge !== undefined ? `Basis ${k.menge} . ` : ""}
        {st ? wort : (k.grund || "kein Urteil")}
      </div>
    </div>
  );
}

function TechCard({ tech, baselines, vorperiode, ursachen }) {
  // baselines wird hier NICHT mehr gelesen. Ziele und Schwellen kommen aus der
  // kpi_uebersicht.csv. Der Parameter bleibt, damit die Aufrufstellen
  // unveraendert bleiben.
  // Bereiche nach Daten-Vorhandensein (so zeigt eine Kombi-Karte alle gleichzeitig)
  const isOT = tech.a1 != null || tech.a_ges != null || tech.a0 != null || tech.quelle === "onetouch";
  const isNFTQ = [tech.nftq_b, tech.nftq_s, tech.nftq_m, tech.nftq_p].some(v => v != null) || tech.quelle === "nftq";
  const isSMS = tech.cc_rate != null || tech.termintreue != null || tech.loesungsquote != null || tech.nps != null || tech.nps_montage != null || tech.nps_pb != null;
  // Die Karte RECHNET NICHT MEHR. Sie fragt dieselbe Funktion wie die Zaehlung
  // oben ("Kritisch: 6") und die Massnahmenliste.
  //
  // Warum das noetig war: Hier stand eine eigene Kopie der Regeln - und die war
  // beim Aufraeumen nur halb erwischt. Ergebnis bei Adil Kheder (Schalten,
  // KW28): Das Schild sagte KRITISCH, die Zaehlung daneben sagte "Kritisch: 0",
  // und auf der Karte war nichts rot. Die Karte bewertete noch NPS Schalten
  // (-100 aus EINER Bewertung), was es nirgends sonst mehr tut.
  //
  // Wer hier kuenftig eine Bedingung braucht: NICHT nachbauen, sondern
  // techWorst() erweitern. Eine Regel, ein Ergebnis.
  const worst = techWorst(tech);
  const borderColor = worst === "kritisch" ? "#7f1d1d" : worst === "warnung" ? "#78350f" : "#14532d";
  const quelleLabel = { smsfeedback: "SMS-Feedback", smsfeedbackschalten: "Schalten", nftq: "NFTQ", standard: "Manuell", onetouch: "OneTouch", alle: "Alle" }[tech.quelle] || "";
  // NPS Schalten hat keinen Zielwert (steht so in der Uebersicht: KEIN_ZIEL),
  // also gibt es hier nie eine Farbe. Frueher rief diese Zeile getNPSStatus()
  // ohne Menge und ohne Ziel auf - und faerbte eine Zahl, die niemand bewertet.
  const npsStatus = ampelU(tech, "nps_schalten");
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
        {/* Hier stand KPIBar mit OT_BASELINE = { a_ges: 95.0, a1: 60.0 }. Die 60
            hat niemand je belegt, und One Touch kommt in der Zusatzvereinbarung
            gar nicht vor. Jetzt: Zahlen zeigen, nicht bewerten. */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 8 }}>
          <KPIKachel label="Gesamterfolg" k={{ wert: tech.a_ges, menge: tech.auftraege, ampel: null, ziel: null, schwelle: null, grund: "kein Telekom-Zielwert" }} />
          <KPIKachel label="Erstlösung (A1)" k={{ wert: tech.a1, menge: tech.auftraege, ampel: null, ziel: null, schwelle: null, grund: "kein Telekom-Zielwert" }} />
          {tech.a0 !== null && tech.a0 !== undefined ? <KPIKachel label="Nicht erledigt (A0)" k={{ wert: tech.a0, menge: tech.auftraege, ampel: null, ziel: null, schwelle: null, grund: "kein Telekom-Zielwert" }} /> : null}
        </div>
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
        {/* Der Status kommt aus derselben Funktion wie ueberall sonst.
            Bereitstellung bekommt bewusst KEINEN - mit Begruendung darunter,
            damit niemand denkt, da sei etwas vergessen worden. */}
        {/* Wert, Menge, Ampel und die Begruendung kommen jetzt ALLE aus der
            Uebersicht. Frueher stand hier viermal ein Zielwert im Code
            (6,6 / 4 / 8,5) - und die Bereitstellung bekam ihre Begruendung als
            festen Satz, obwohl [U] laengst "KEIN_ZIEL" dazu schreibt. */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 8 }}>
          {[["nftq_b", "Nachfolgetickets Bereitstellung"], ["nftq_s", "Nachfolgetickets Schalten"],
            ["nftq_m", "Nachfolgetickets Montage"], ["nftq_p", "Nachfolgetickets Problembehebung"]].map(([id, label]) => (
            <KPIKachel key={id} label={label} k={kpiAusU(tech, id)} hoch={false} />
          ))}
        </div>
      </>)}
      {isSMS && (<>
        {/* KPIBar ist raus. Sie faerbte nach value/baseline: unter 0,93 gelb,
            unter 0,85 rot. Bei Termintreue hiess das Warnung erst ab 89,3 % -
            der Vertrag sagt 95,5 %. Die Karte urteilte also milder als [U]
            direkt daneben. */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(150px, 1fr))", gap: 8, marginBottom: 8 }}>
          <KPIKachel label="Courtesy Calls" k={kpiAusU(tech, "cc")} hoch
            trend={vorperiode ? getTrend(tech.cc_rate, vorperiode.cc_rate) : null} />
          <KPIKachel label="Termintreue" k={kpiAusU(tech, "termintreue")} hoch
            trend={vorperiode ? getTrend(tech.termintreue, vorperiode.termintreue) : null} />
          <KPIKachel label="Lösungsquote Bereitstellung" k={{ wert: tech.loesungsquote, menge: null, ampel: null, ziel: null, schwelle: null, grund: "[U] liest sie noch nicht - kein Urteil" }} hoch />
          {[["nps_montage", "NPS Montage"], ["nps_pb", "NPS Problembehebung"],
            ["nps_schalten", "NPS Schalten"]].map(([id, label]) => (
            <KPIKachel key={id} label={label} k={kpiAusU(tech, id)} hoch einheit="" />
          ))}
        </div>
        {tech.termintreue !== null && (() => { const pts = getTermintreeuPunkte(tech.termintreue); return (
          <div style={{ fontSize: 10, color: pts >= 164 ? "#4ade80" : pts >= 0 ? "#fbbf24" : "#f87171", marginTop: -6, marginBottom: 4, paddingLeft: 2 }}>
            Auftragsinfo Punkte: <b>{pts > 0 ? "+" : ""}{pts}</b>
          </div>
        ); })()}
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
  // One-Touch-Befunde tragen keine ATS - der Report liefert keine. Sie stehen
  // deshalb bei BEIDEN Karten eines Technikers, der in zwei Bereichen
  // arbeitet. Das gehoert dazugesagt, sonst zaehlt man sie doppelt.
  const ohneAts = befunde.filter(b => !b.ats).length;
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
      {offen && ohneAts > 0 ? (
        <div style={{ fontSize: 9.5, color: "#6b7280", marginTop: 6, fontStyle: "italic" }}>
          {ohneAts} davon ohne ATS-Angabe (One Touch liefert keine) - die stehen bei
          jeder Karte dieses Technikers.
        </div>
      ) : null}
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
function berichtText(name, meine, tech) {
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
  if (tech) {
    const zeilen = [];
    // GROSSBUCHSTABEN, nicht "kritisch": Farbe kann jedes Mailprogramm
    // wegwerfen, Grossbuchstaben nicht. Der Techniker muss die Stelle auch
    // dann finden, wenn die Mail als nackter Text bei ihm ankommt.
    const wort = { gut: "im Ziel", warnung: "WARNUNG", kritisch: "KRITISCH" };
    const f = (v) => (v === null || v === undefined || isNaN(v)) ? null : v.toFixed(1).replace(".", ",");
    const nimm = (label, wert, status, basis) => {
      if (wert === null) return;
      zeilen.push(`   ${label.padEnd(24)} ${String(wert).padStart(7)}` +
        (status ? `   ${wort[status]}` : "   (kein Zielwert / zu wenig Daten)") +
        (basis ? `   [Basis: ${basis}]` : ""));
    };
    // Der Bericht geht an den Techniker. Deshalb steht hier genau das, was auch
    // in der Excel steht - Wert, Urteil und Basis aus der Uebersicht. Bis
    // 17.07.2026 rechnete diese Stelle selbst und konnte einem Mann etwas
    // anderes schreiben, als seine Zeile in der Uebersicht sagt.
    [["termintreue", "Termintreue"], ["cc", "Courtesy Call"],
     ["nps_pb", "NPS Problembehebung"], ["nps_montage", "NPS Montage"],
     ["nftq_m", "NFTQ Montage %"], ["nftq_s", "NFTQ Schalten %"],
     ["nftq_p", "NFTQ Problembeh. %"]].forEach(([id, label]) => {
      const k = kpiAusU(tech, id);
      nimm(label, f(k.wert), k.ampel, k.menge);
    });
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

// Der Agent ist dunkel und steckt in Kacheln, die scrollen. window.print()
// erwischt davon nur, was gerade sichtbar ist - Arash bekam ein PDF in
// Bildschirmlaenge. Statt zu suchen, welches Element abschneidet, bekommt der
// Bericht ein EIGENES Fenster: nur das Blatt, weisses Papier, volle Laenge.
// Das ist unabhaengig davon, was der Agent drumherum treibt.
const DRUCK_CSS = `
  @page { size: A4; margin: 1.6cm; }
  body { background: #fff; color: #111; margin: 0;
         font-family: Calibri, Arial, sans-serif; font-size: 11pt; line-height: 1.45; }
  /* Die Farben des Agenten sind fuer den Bildschirm gedacht - auf Papier
     kaeme eine schwarze Seite heraus. Raender bleiben farbig, die sind der
     einzige Hinweis auf die Einstufung. */
  * { background-color: transparent !important; color: #111 !important;
      box-shadow: none !important; max-height: none !important; overflow: visible !important; }
  div[style*="border-left"] { padding-left: 10px !important; margin: 8px 0 !important;
                              page-break-inside: avoid; }
  div[style*="border-bottom"] { border-bottom: 1px solid #999 !important; }
  div[style*="border: 1px solid"], div[style*="border:1px solid"] {
      border: 1px solid #999 !important; }
`;

// Farbe im Mailtext.
// WICHTIG: Das hier ist KEIN zweiter Bericht - es faerbt nur ein, was
// berichtText() geschrieben hat. Zwei Fassungen desselben Textes waeren genau
// der Fehler, den wir im ganzen Projekt bekaempfen: sie laufen auseinander,
// und dann steht im Bildschirm etwas anderes als in der Mail.
//
// Eine mailto-Mail ist immer reiner Text - da geht keine Farbe rein. Beim
// KOPIEREN aber schon: die Zwischenablage kann Text UND HTML tragen, und
// Outlook nimmt das HTML. Da Arashs Berichte fuer mailto ohnehin zu lang sind
// und er sie kopiert, ist das der Weg, den er wirklich geht.
function berichtAlsHtml(text) {
  const e = String(text || "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  return e
    // Ueberschriften des Berichts fett - erkennbar an Grossbuchstaben am Zeilenanfang
    .replace(/^(DEINE ZAHLEN[^\n]*|GUT GELAUFEN[^\n]*|UNZUFRIEDENE KUNDEN[^\n]*|NACHFOLGETICKETS[^\n]*|AUFTRAEGE OHNE[^\n]*)$/gm,
      '<b style="color:#111">$1</b>')
    // Farbe UND Fettschrift UND Grossbuchstaben - drei Wege zum selben Ziel.
    // Faellt einer aus (Schwarzweissdruck, Mailprogramm ohne HTML, Rotblindheit
    // - das betrifft jeden zwoelften Mann), tragen die anderen zwei.
    .replace(/\bKRITISCH\b/g, '<span style="color:#c00;font-weight:900;font-size:1.08em;letter-spacing:0.5px">KRITISCH</span>')
    .replace(/\bWARNUNG\b/g, '<span style="color:#b26b00;font-weight:900;letter-spacing:0.5px">WARNUNG</span>')
    .replace(/\bim Ziel\b/g, '<span style="color:#1b7a2f;font-weight:bold">im Ziel</span>')
    .replace(/\n/g, "<br>");
}

function BerichtTab({ ursachen, techs, baselines, kontakte, nurTechniker, onWaehlen }) {
  const namen = [...new Set((ursachen || []).map(u => u.name))].sort();
  // Die Auswahl oben in der Kopfleiste gilt auch hier - sonst waehlt man
  // zweimal denselben Mann und wundert sich, warum die Reiter
  // Verschiedenes zeigen.
  const vonOben = nurTechniker
    ? namen.find(n => namensSchluessel(n) === nurTechniker) || ""
    : "";
  const [gewaehlt, setGewaehlt] = useState(namen[0] || "");
  const [mailOffen, setMailOffen] = useState(false);
  const [kopiert, setKopiert] = useState(false);
  const blattRef = useRef(null);

  const drucken = (wenName) => {
    const inhalt = blattRef.current ? blattRef.current.innerHTML : "";
    const w = window.open("", "_blank", "width=900,height=1100");
    if (!w) {
      // Popup-Blocker. Dann wenigstens der alte Weg - lieber ein kurzes PDF
      // als gar keins, aber der Grund gehoert gesagt.
      window.alert("Der Browser hat das Druckfenster blockiert.\n\n" +
        "Erlaube Pop-ups fuer diese Seite - sonst druckt er nur den sichtbaren " +
        "Bildschirmausschnitt statt des ganzen Berichts.");
      window.print();
      return;
    }
    w.document.write('<!DOCTYPE html><html lang="de"><head><meta charset="utf-8">' +
      '<title>' + String(wenName || "Bericht").replace(/[<>]/g, "") + '</title>' +
      '<style>' + DRUCK_CSS + '</style></head><body>' + inhalt + '</body></html>');
    w.document.close();
    w.focus();
    // Kurz warten, sonst druckt Chrome ein leeres Blatt - das Fenster ist
    // noch nicht fertig aufgebaut.
    setTimeout(() => { w.print(); }, 400);
  };

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
  const name = vonOben || (namen.includes(gewaehlt) ? gewaehlt : namen[0]);
  const meine = (ursachen || []).filter(u => u.name === name);
  // Auch hier ueber die Namensteile: der Techniker heisst in den Kennzahlen
  // vielleicht "Kheder Adil" und im Bericht "Adil Kheder".
  const tech = (techs || []).find(t => namensSchluessel(t.name) === namensSchluessel(name));
  const sortiert = [...meine].sort(
    (a, b) => (URSACHEN_RANG[a.einstufung] ?? 9) - (URSACHEN_RANG[b.einstufung] ?? 9));
  const zaehl = (e) => meine.filter(b => b.einstufung === e).length;
  const ats = [...new Set(meine.map(b => b.ats).filter(Boolean))].join(", ");
  const text = berichtText(name, meine, tech);
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
        <select value={name} onChange={e => {
          setGewaehlt(e.target.value);
          setMailOffen(false);
          // Auch nach oben durchreichen: eine Auswahl, alle Reiter.
          if (onWaehlen) onWaehlen(namensSchluessel(e.target.value));
        }}
          style={{ background: "#1f2937", color: "#f9fafb", border: "1px solid #374151", borderRadius: 6, padding: "7px 10px", fontSize: 12 }}>
          {namen.map(n => <option key={n} value={n}>{n}</option>)}
        </select>
        <button onClick={() => drucken(name)}
          title="Oeffnet den Bericht als sauberes weisses Blatt in voller Laenge - dort Strg+P oder 'Als PDF speichern'"
          style={{ background: "#1f2937", color: "#f9fafb", border: "1px solid #374151", borderRadius: 6, padding: "7px 12px", fontSize: 12, cursor: "pointer" }}>
          Drucken / Als PDF
        </button>
        <button onClick={() => setMailOffen(!mailOffen)}
          style={{ background: "#1e3a5f", color: "#dbeafe", border: "1px solid #2563eb", borderRadius: 6, padding: "7px 12px", fontSize: 12, cursor: "pointer" }}>
          Mail vorbereiten
        </button>
        <span style={{ fontSize: 10, color: "#6b7280", maxWidth: 430, lineHeight: 1.4 }}>
          Die Mail enthaelt den ganzen Bericht als Text - kein Anhang noetig.
          "Als PDF" oeffnet ein eigenes weisses Blatt in voller Laenge; dort im
          Druckdialog "Als PDF speichern". Falls nichts aufgeht: Pop-ups erlauben.
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
          {/* Kein Textfeld mehr: ein Textfeld kann nur EINE Farbe. So siehst du
              schon vor dem Kopieren, was rot ist - und genau das landet auch
              in der Mail. */}
          <div style={{ width: "100%", background: "#0b1220", color: "#d1d5db",
            border: "1px solid #374151", borderRadius: 6, padding: 12, fontSize: 11,
            fontFamily: "Consolas, monospace", lineHeight: 1.55, maxHeight: 420,
            overflowY: "auto", whiteSpace: "pre-wrap" }}
            dangerouslySetInnerHTML={{ __html: berichtAlsHtml(text)
              .replace(/color:#111/g, "color:#f9fafb")
              .replace(/color:#c00/g, "color:#f87171")
              .replace(/color:#b26b00/g, "color:#fbbf24")
              .replace(/color:#1b7a2f/g, "color:#4ade80") }} />
          <div style={{ display: "flex", gap: 8, marginTop: 8, flexWrap: "wrap", alignItems: "center" }}>
            <button onClick={async () => {
              // Zwei Fassungen in die Zwischenablage: HTML fuer Outlook (mit
              // Farbe), reiner Text fuer alles andere. Das Mailprogramm nimmt
              // sich, was es versteht - kann es kein HTML, kommt der Text an
              // und nichts geht verloren.
              const html = '<div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;' +
                'line-height:1.4;color:#111">' + berichtAlsHtml(text) + '</div>';
              try {
                if (navigator.clipboard && window.ClipboardItem) {
                  await navigator.clipboard.write([new window.ClipboardItem({
                    "text/html": new Blob([html], { type: "text/html" }),
                    "text/plain": new Blob([text], { type: "text/plain" }),
                  })]);
                } else if (navigator.clipboard) {
                  await navigator.clipboard.writeText(text);   // ohne Farbe, aber vollstaendig
                }
                setKopiert(true); setTimeout(() => setKopiert(false), 2000);
              } catch (e) {
                try { await navigator.clipboard.writeText(text); setKopiert(true); setTimeout(() => setKopiert(false), 2000); }
                catch (e2) { window.alert("Kopieren hat nicht geklappt - bitte den Text oben von Hand markieren."); }
              }
            }}
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

      <div className="druckblatt" ref={blattRef} style={{ background: "#111827", border: "1px solid #1f2937", borderRadius: 8, padding: "20px 24px" }}>
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
  const [uebersicht, setUebersicht] = useState([]);
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
  // Ein Techniker, ueber alle Reiter hinweg. Bisher konnte man nur im Bericht
  // einen auswaehlen - in den Kennzahlen-Reitern musste man ihn suchen.
  // Gespeichert wird der NAMENSSCHLUESSEL, nicht der Name: sonst greift der
  // Filter im One-Touch-Reiter nicht, wo er "Kheder Adil" heisst.
  const [nurTechniker, setNurTechniker] = useState("");
  const [datenAlt, setDatenAlt] = useState(false);
  const [nurKritisch, setNurKritisch] = useState(false);
  // Die KI-Analyse ist eine Momentaufnahme - sie rechnet nicht mit, wenn man
  // danach filtert. Arash sah eine Analyse ueber DREI Techniker, waehrend das
  // Dashboard ACHT zeigte, und nichts sagte ihm das. Deshalb merken wir uns,
  // worauf sie lief, und vergleichen es mit dem, was gerade eingestellt ist.
  const [analyseBasis, setAnalyseBasis] = useState(null);
  const [showVerlauf, setShowVerlauf] = useState(null);
  const [uploadPeriod, setUploadPeriod] = useState(null); // { von, bis, kw, label }
  const [showPeriodDialog, setShowPeriodDialog] = useState(false);
  const [pendingFile, setPendingFile] = useState(null);
  const dashboardRef = useRef(null);

  useEffect(() => {
    try {
      const saved = localStorage.getItem(STORAGE_KEY);
      if (saved) {
        setGespeichert(JSON.parse(saved));
        // Aelterer Stand? Dann wurde er mit einem Einlesen erzeugt, das Fehler
        // hatte. Nicht heimlich weiterrechnen - sagen.
        const v = parseInt(localStorage.getItem(VERSION_KEY) || "1", 10);
        if (v !== DATEN_VERSION && Object.keys(JSON.parse(saved)).length) setDatenAlt(true);
      }
      const savedK = localStorage.getItem(KONTAKTE_KEY);
      if (savedK) setKontakte(JSON.parse(savedK));
      const savedB = localStorage.getItem(BASELINE_KEY);
      if (savedB) setBaselines(JSON.parse(savedB));
      const savedA = localStorage.getItem(ARCHIV_KEY);
      if (savedA) setArchiv(JSON.parse(savedA));
      const savedU = localStorage.getItem(URSACHEN_KEY);
      if (savedU) setUrsachen(JSON.parse(savedU));
      const savedUe = localStorage.getItem(UEBERSICHT_KEY);
      if (savedUe) setUebersicht(JSON.parse(savedUe));
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
  useEffect(() => { try { localStorage.setItem(UEBERSICHT_KEY, JSON.stringify(uebersicht)); } catch(e) {} }, [uebersicht]);

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
    // Die Uebersicht ist keine Kategorie: kein eigener Reiter, kein Umschalten
    // der Ansicht. Sie liefert nur die Urteile zu den Zahlen, die schon da sind.
    if (rows[0] && rows[0].quelle === "uebersicht") {
      setUebersicht(rows);
      setError("ok Uebersicht aus [U] geladen - " + rows.length
        + " Techniker. Ampel und Zielwerte kommen ab jetzt von dort.");
      return;
    }
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
      // Frisch eingelesen - also mit der jetzigen Fassung.
      try { localStorage.setItem(VERSION_KEY, String(DATEN_VERSION)); } catch (e) {}
      setDatenAlt(false);
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

  // Beim "Ordner verbinden" filtert dateiRolle die Detail-CSVs raus - beim
  // Hochladen tat es das NICHT. Arash waehlte am 17.07. acht Dateien aus; die
  // *_details.csv landeten als "Manuell (320)" (eine Zeile je Anrufversuch) und
  // verstopften die Ansicht. Derselbe Filter, beide Wege.
  const processFile = useCallback(async (file) => {
    if (/_details\.csv$/i.test(file.name)) {
      setError(`"${file.name}" ist eine Detail-Datei - die gehoert in den Ursachenbericht, nicht in den Agenten. Uebersprungen.`);
      return;
    }
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

  // Der Nachschlag-Index. Einmal gebaut, von allen benutzt.
  const uIndex = React.useMemo(() => uebersichtIndex(uebersicht), [uebersicht]);

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
          // Schluessel ueber die sortierten Namensteile - siehe namensSchluessel().
          const key = namensSchluessel(t.name) + "#" + String(t.standort);
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
    if (nurTechniker && namensSchluessel(t.name) !== nurTechniker) return false;
    // Der Mindest-Auftrags-Filter darf nur greifen, wenn es ueberhaupt eine
    // Auftragszahl GIBT. Bis 17.07.2026 stand hier parseInt(...) || 0 - Formate
    // ohne Auftragsspalte lieferten "-", das wurde zu 0, und der Filter warf
    // sie alle raus. Arash sah dann "SMS-Feedback (8)" im Reiter und "Noch
    // keine Daten" in der Ansicht. Ein Filter sortiert aus, weil ein Wert zu
    // klein ist - nicht, weil er fehlt.
    const roh = t.auftraege;
    if (roh === null || roh === undefined || roh === "-" || roh === "") return true;
    const auftr = typeof roh === "number" ? roh : parseInt(roh);
    if (isNaN(auftr)) return true;
    return auftr >= minAuftraege;
  }).map(t => mitUebersicht(t, uIndex));

  // Alle Namen, die es irgendwo gibt - aus den Kennzahlen UND aus dem
  // Ursachenbericht. Ein Techniker, der diese Woche nur Befunde hat, soll
  // trotzdem waehlbar sein.
  const alleNamen = (() => {
    const m = new Map();
    [...Object.values(gespeichert).flat(), ...ursachen].forEach(t => {
      const k = namensSchluessel(t.name);
      if (k && !m.has(k)) m.set(k, t.name);
    });
    return [...m.entries()].sort((a, b) => a[1].localeCompare(b[1]));
  })();

  const hatDaten = Object.keys(gespeichert).length > 0;

  // berechneMassnahmen() stand hier bis 17.07.2026 - 25 Zeilen mit EIGENEN
  // Schwellen (a1 >= 60, a0 > 10, NPS gegen 68). Beim Aufraeumen fiel auf: die
  // Funktion wurde NIE aufgerufen. Sie hat also nie etwas angezeigt und war
  // trotzdem eine Stelle, an der jemand haette nachziehen muessen. Geloescht.
  // Die Massnahmen kommen aus der Ansicht weiter unten.

  // -------------------------------------------------------------------------
  // BILANZ STATT SCORE (17.07.2026)
  //
  // Hier stand berechneTechScore: eine Zehnerskala aus a1/60, a_ges/95,
  // (nps+100)/20 und NFTQ (1 - x/20). JEDE dieser Zahlen war frei erfunden -
  // keine steht in der Zusatzvereinbarung, keine in einer Telekom-Folie. Die
  // Rechnung kannte keine Mindestmenge und zog ihre Bezugswerte aus den
  // Baselines; damit war sie die letzte Stelle, an der der Baseline-Editor
  // noch etwas bewirkte, ohne dass man es sah.
  //
  // Was sie im Betrieb angerichtet hat (KW28): Adil Kheder stand mit
  // "Avg Score 10.0/10" auf der Startseite. Sein NPS PB von +100 kam aus
  // EINER Rueckmeldung, und seine NFTQ Montage lag bei 22,2 % (4 von 18) -
  // dem einzigen echten KRITISCH der Woche. Die Skala hat es weggemittelt.
  //
  // Jetzt wird nur gezaehlt, was [U] beurteilt hat. Kein Mittelwert, keine
  // Gewichtung, keine Skala: "5 von 7 im Ziel" laesst sich nachrechnen, und
  // der Nenner sagt gleich mit, wie viel ueberhaupt beurteilbar war. Wer
  // keine Uebersicht hat, bekommt "-" und keine Note.
  const techBilanz = (t) => {
    const urteile = BEWERTETE_KENNZAHLEN.map(k => ampelU(t, k)).filter(x => x !== null);
    return {
      bewertet: urteile.length,
      imZiel: urteile.filter(a => a === "gut").length,
      warnung: urteile.filter(a => a === "warnung").length,
      kritisch: urteile.filter(a => a === "kritisch").length,
    };
  };
  const bilanzText = (t) => {
    const b = techBilanz(t);
    return b.bewertet ? `${b.imZiel} von ${b.bewertet}` : "-";
  };
  const teamBilanz = () => {
    const b = angezeigt.map(techBilanz).filter(x => x.bewertet);
    if (!b.length) return "-";
    return `${b.reduce((s, x) => s + x.imZiel, 0)} von ${b.reduce((s, x) => s + x.bewertet, 0)}`;
  };

  // Farbe und Wort kommen aus derselben Quelle wie die Ampel selbst - nicht
  // aus einer zweiten Schwelle. null ist grau: kein Urteil ist kein "gut".
  const ampelFarbe = (a) => a === "kritisch" ? "#f87171" : a === "warnung" ? "#fbbf24" : a === "gut" ? "#4ade80" : "#6b7280";
  const ampelWort = (a) => a === "kritisch" ? "KRITISCH" : a === "warnung" ? "WARNUNG" : a === "gut" ? "im Ziel" : "kein Urteil";

  // Die Einzelbewertung war die dreizehnte Stelle mit eigenen Zahlen: sie tippte
  // "Ziel>=96%", "NPS-Ziel jeweils >=67" und "NFTQ-S Ziel<=7%" fest in den Text,
  // den die KI zu sehen bekam - am Umbau vorbei und teils falsch (der Vertrag
  // sagt 96,1 und 67,1). Jetzt bekommt die KI dasselbe wie ueberall: Wert,
  // Menge, Ampel, Ziel und den Grund, alles aus [U].
  const bewerteEinzelTechniker = useCallback(async (tech) => {
    const b = techBilanz(tech);
    const zeile = (id, label) => {
      const k = kpiAusU(tech, id);
      if (k.wert === null && !k.grund) return null;
      const wert = k.wert === null ? "-" : k.wert.toFixed(1);
      const basis = k.menge !== null && k.menge !== undefined ? `Basis ${k.menge}` : "Basis unbekannt";
      const urteil = k.ampel ? k.ampel.toUpperCase() : `kein Urteil (${k.grund || "Grund unbekannt"})`;
      const ziel = k.ziel !== null && k.ziel !== undefined && k.ziel !== "" ? `, Ziel ${k.ziel}` : ", kein Zielwert";
      return `${label}=${wert} (${basis}${ziel}) -> ${urteil}`;
    };
    const zeilen = [
      zeile("termintreue", "Termintreue"), zeile("cc", "Courtesy Calls"),
      zeile("nps_montage", "NPS Montage"), zeile("nps_pb", "NPS Problembehebung"),
      zeile("nps_schalten", "NPS Schalten"),
      zeile("nftq_s", "Nachfolgetickets Schalten"),
      zeile("nftq_m", "Nachfolgetickets Montage"),
      zeile("nftq_p", "Nachfolgetickets Problembehebung"),
      zeile("nftq_b", "Nachfolgetickets Bereitstellung"),
    ].filter(Boolean);
    const kpiText = zeilen.length
      ? zeilen.join(" | ")
      : "KEINE KPI-Uebersicht geladen - es liegt kein einziges Urteil vor.";
    setBewertungLoading(prev => ({ ...prev, [tech.name]: true }));
    try {
      const res = await fetch("/api/analyse", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          model: "claude-sonnet-4-6", max_tokens: 500,
          system: `Du bist KPI-Bewerter für Telekom-Techniker. Antworte NUR mit JSON ohne Backticks:
{"kommentar": "1-2 Sätze persönliche Bewertung mit Namen", "staerken": ["max 2 Stärken"], "schwaechen": ["max 2 Schwächen"], "massnahme": "Eine konkrete Maßnahme"}`,
          messages: [{ role: "user", content: `Techniker: ${tech.name}, im Ziel: ${bilanzText(tech)} bewerteten Kennzahlen, KPIs: ${kpiText}, Aufträge: ${tech.auftraege}` }]
        }),
      });
      const data = await res.json();
      const text = data.content?.map(b => b.text || "").join("") || "";
      const clean = text.replace(/```json|```/g, "").trim();
      const parsed = JSON.parse(clean);
      setTechBewertungen(prev => ({ ...prev, [tech.name]: { ...parsed, bilanz: b } }));
    } catch(e) {
      setTechBewertungen(prev => ({ ...prev, [tech.name]: { kommentar: "Bewertung fehlgeschlagen.", bilanz: b } }));
    } finally {
      setBewertungLoading(prev => ({ ...prev, [tech.name]: false }));
    }
  }, [angezeigt]);

  const bewerteAlle = useCallback(async () => {
    for (const tech of angezeigt) {
      await bewerteEinzelTechniker(tech);
    }
    setActiveTab("firmendashboard");
  }, [angezeigt, bewerteEinzelTechniker]);

  const runAnalysis = async () => {
    if (!angezeigt.length) return;
    setLoading(true); setError(""); setAiAnalysis(""); setMassnahmen([]); setMassnahmenFehler(null);
    setAnalyseBasis({ anzahl: angezeigt.length, kategorie: aktiveKategorie, techniker: nurTechniker });
    // DIE MENGEN MUESSEN MIT. Vorher standen hier nur die Quoten - die
    // Mindestmengenregel steht im Prompt, aber ohne die Zahlen dahinter musste
    // die KI raten. Sie schrieb dann "Basis unbekannt" und erfand "0
    // Rueckmeldungen". Eine Quote ohne ihre Basis ist genau die halbe Wahrheit,
    // gegen die dieses Projekt gebaut ist.
    // Was an die KI geht, ist jetzt die Uebersicht - fertig beurteilt, mit
    // Menge, Zielwert und Grund. Vorher gingen die Rohquoten hin, und die
    // Zielwerte standen als Prosa im Prompt: zwei Kopien derselben Zahl, die
    // auseinanderliefen. Die KI schrieb dann "NFTQ-Bereitstellung kritisch
    // (Ziel <=4%)" - ein Urteil, das es nirgends gibt.
    const zz = (x) => (x === null || x === undefined ? "-" : String(x));
    const dataStr = angezeigt.map(t => {
      if (!t.u) return `${t.name} (FS${t.standort}): keine Zeile in der KPI-Uebersicht - NICHT beurteilen.`;
      const zeilen = Object.entries(t.u.kpi)
        .filter(([, k]) => k.wert !== null || k.grund)
        .map(([id, k]) => `    ${id}: wert=${zz(k.wert)} menge=${zz(k.menge)} `
          + `ampel=${k.ampel ?? "null"} ziel=${zz(k.ziel)} schwelle=${zz(k.schwelle)}`
          + (k.grund ? ` grund="${k.grund}"` : ""));
      return `${t.name} (FS${t.standort}) - Gesamt: ${t.u.status}, ${zz(t.u.auftraege)} Auftraege\n`
        + zeilen.join("\n");
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
        // Hier standen bis 17.07.2026 eigene Schwellen - und zwar WIEDER andere
        // als an den zwei Stellen darueber (Schalten 7 statt 6,6, PB 8,7 statt
        // 8,5, A1 >= 60, A0 > 10). Derselbe Agent gab je nach Ansicht
        // verschiedene Ampeln aus. Jetzt: eine Quelle, ein Urteil.
        // Kein Urteil -> "gut" waere gelogen; die Massnahme sagt es dann selbst.
        const worst = techWorst(t) ?? "gut";
        // Das Lob. Hier standen die letzten drei erfundenen Zielwerte des
        // Agenten: "?? 68" fuer den NPS, "Zielwert 96%" fuer die CC-Rate und
        // "Zielwert 60%" fuer A1 - letzteres fuer eine Kennzahl, fuer die es
        // ueberhaupt keinen Zielwert gibt. Jetzt kommt jede Zahl aus der Uebersicht.
        const pb = kpiAusU(t, "nps_pb");
        const ccU = kpiAusU(t, "cc");
        const lob = pb.ampel === "gut" && pb.wert != null
          ? "Ausgezeichnet! NPS Problembehebung " + pb.wert.toFixed(0) + " ueber Zielwert " + pb.ziel + "!"
          : ccU.ampel === "gut" && ccU.wert != null
          ? "Sehr gute Courtesy-Call-Rate " + ccU.wert.toFixed(1) + "% - Zielwert " + ccU.ziel + "% erreicht!"
          : nftqStatusListe(t).length > 0 && nftqStatusListe(t).every(x => x === "gut")
          ? "Alle bewerteten NFTQ-Werte im Zielbereich - ausgezeichnete Qualitaetsarbeit!"
          : "Alle bewerteten Kennzahlen im Zielbereich - weiter so!";
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
      // Fuer das Archiv. Hier stand bis 17.07.2026 die ZEHNTE Regelquelle des
      // Agenten, und sie war die schlechteste von allen:
      //   const vals = [t.nftq_b, t.nftq_s, t.nftq_m, t.nftq_p].filter(Boolean);
      //   status = vals.some(v => v > 10) ? "kritisch" : vals.some(v => v > 5) ...
      // Schwellen 5 und 10 - frei erfunden, nirgends sonst im Agenten, in keinem
      // Telekom-Papier. NFTQ-B mitgezaehlt. Und filter(Boolean) wirft jede 0
      // weg: wer keinen einzigen Fehler hatte, fiel aus der Bewertung.
      // Jetzt kommt der Status aus derselben Quelle wie ueberall.
      datatenMitStatus[kat] = techs.map(t => {
        const status = techWorst(mitUebersicht(t, uIndex));
        // _score ist raus (der erfundene Zehnerscore). Nebenbei war er hier
        // doppelt falsch: er lief auf t OHNE angehaengte Uebersicht, waehrend
        // _status direkt daneben mitUebersicht(t) benutzt.
        return { ...t, _status: status };
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
  }, [gespeichert, aiAnalysis, hatDaten, techBewertungen, uebersicht]);

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

  const criticalCount = angezeigt.filter(t => techWorst(t) === "kritisch").length;
  // Wie viele Techniker haben ueberhaupt ein Urteil? Ist das 0, sagt
  // "Kritisch: 0" nichts ueber die Lage aus - dann gehoert dort ein Strich hin.
  const bewertetCount = angezeigt.filter(t => techWorst(t) !== null).length;

  const avg = (key) => {
    const vals = angezeigt.map(t => t[key]).filter(v => v !== null && !isNaN(v));
    return vals.length ? (vals.reduce((s, v) => s + v, 0) / vals.length).toFixed(1) : "-";
  };

  const isOTView = aktiveKategorie === "onetouch";

  const FirmendashboardTab = () => {
    if (!angezeigt || !angezeigt.length) return null;
    // Sortierung ohne erfundene Gewichtung: erst wer die wenigsten KRITISCH
    // hat, dann die wenigsten WARNUNG, dann wer mehr im Ziel steht. Wer gar
    // kein Urteil hat, steht am Ende - nicht oben, wo Leere wie Bestnote
    // aussieht. Frueher sortierte hier der Zehnerscore.
    const sorted = [...angezeigt].filter(t => !nurKritisch || techWorst(t) === "kritisch").sort((a, b) => {
      const x = techBilanz(a), y = techBilanz(b);
      if (!x.bewertet !== !y.bewertet) return x.bewertet ? -1 : 1;
      return (x.kritisch - y.kritisch) || (x.warnung - y.warnung) || (y.imZiel - x.imZiel);
    });
    return (
      <div>
        <div style={{ background: "#0f172a", border: "1px solid #1f2937", borderRadius: 8, padding: "16px", marginBottom: 20 }}>
          <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 12, fontWeight: 700, textTransform: "uppercase", letterSpacing: 1 }}>Team-Übersicht</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 12 }}>
            {[
              { label: "Im Ziel", value: teamBilanz(), color: "#9ca3af" },
              { label: "Kritisch", value: bewertetCount ? criticalCount : "-", color: !bewertetCount ? "#6b7280" : criticalCount > 0 ? "#f87171" : "#4ade80" },
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
          const ampel = techWorst(tech);
          const bilanz = techBilanz(tech);
          const bew = techBewertungen[tech.name];
          const isLoadingThis = bewertungLoading[tech.name];
          const k = kontakte[tech.name] || {};
          const mailBody = bew
            ? `Hallo ${tech.name.split(" ")[0]},\n\nhier ist Ihre persönliche KPI-Bewertung:\n\nStand: ${ampelWort(ampel)} - ${bilanzText(tech)} bewerteten Kennzahlen im Ziel\n\n${bew.kommentar}\n\n${bew.staerken?.length ? `Stärken:\n${bew.staerken.map(s => `• ${s}`).join("\n")}\n\n` : ""}${bew.schwaechen?.length ? `Verbesserungsbedarf:\n${bew.schwaechen.map(s => `• ${s}`).join("\n")}\n\n` : ""}Maßnahme: ${bew.massnahme || ""}\n\nMit freundlichen Grüßen\n${FIRMA} Leitstelle`
            : "";
          const mailto = `mailto:${k.email || ""}?subject=${encodeURIComponent(`KPI-Bewertung ${tech.name}`)}&body=${encodeURIComponent(mailBody)}`;
          return (
            <div key={tech.name} style={{ background: "#111827", border: `1px solid ${ampel === "gut" ? "#14532d" : ampel === "warnung" ? "#78350f" : ampel === "kritisch" ? "#7f1d1d" : "#1f2937"}`, borderRadius: 8, padding: "14px 16px", marginBottom: 10 }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <div style={{ fontSize: 16, fontWeight: 700, color: ampelFarbe(ampel), fontFamily: "monospace", minWidth: 28 }}>#{i + 1}</div>
                  <div>
                    <div style={{ fontWeight: 700, fontSize: 14, color: "#f9fafb" }}>{tech.name}</div>
                    <div style={{ fontSize: 11, color: "#6b7280" }}>FS{tech.standort} . {tech.auftraege} Aufträge</div>
                  </div>
                </div>
                <div style={{ textAlign: "center" }} title={bilanz.bewertet ? `${bilanz.bewertet} von 7 Kennzahlen sind beurteilbar - der Rest hat keinen Zielwert oder zu wenig Daten` : "Keine KPI-Uebersicht geladen"}>
                  <div style={{ fontSize: 22, fontWeight: 700, color: ampelFarbe(ampel), fontFamily: "monospace" }}>{bilanzText(tech)}</div>
                  <div style={{ fontSize: 9, color: ampelFarbe(ampel) }}>im Ziel . {ampelWort(ampel)}</div>
                </div>
              </div>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 10 }}>
                {(() => { const k = kpiAusU(tech, "cc"); if (k.wert === null) return null; const c = k.ampel === "kritisch" ? "#f87171" : k.ampel === "warnung" ? "#fbbf24" : k.ampel === "gut" ? "#4ade80" : "#9ca3af"; return <span title={k.grund || undefined} style={{ fontSize: 10, background: "#1f2937", color: c, padding: "2px 8px", borderRadius: 3, fontWeight: k.ampel ? 700 : 400 }}>CC {k.wert.toFixed(1)}% / &gt;={k.ziel}%</span>; })()}
                {(() => { const k = kpiAusU(tech, "termintreue"); if (k.wert === null) return null; const c = k.ampel === "kritisch" ? "#f87171" : k.ampel === "warnung" ? "#fbbf24" : k.ampel === "gut" ? "#4ade80" : "#9ca3af"; return <span title={k.grund || undefined} style={{ fontSize: 10, background: "#1f2937", color: c, padding: "2px 8px", borderRadius: 3, fontWeight: k.ampel ? 700 : 400 }}>Termintreue {k.wert.toFixed(1)}% / &gt;={k.ziel}%</span>; })()}
                {/* Loesungsquote: [U] liest sie noch nicht ein (Spalte "Erledigt B"
                    im SMS-Feedback-Report, Zuordnung bei Telekom angefragt). Bis das
                    geklaert ist: Zahl zeigen, nicht bewerten. Der Agent hat sie hier
                    frueher gegen 95 gemessen - ohne die Menge zu kennen. Bei Ali
                    Sodjajy standen dahinter ZWEI Faelle. */}
                {tech.loesungsquote !== null && tech.loesungsquote !== undefined && <span style={{ fontSize: 10, background: "#1f2937", color: "#9ca3af", padding: "2px 8px", borderRadius: 3 }}>Loesungsquote {tech.loesungsquote.toFixed(1)}% (ohne Urteil)</span>}
                {/* SIEBTE Stelle mit eigener Regel. Vorher: der erstbeste Wert
                    aus [nps, nps_pb, nps_montage] gegen 67, ohne Mindestmenge -
                    und an erster Stelle stand ausgerechnet NPS SCHALTEN, die
                    einzige NPS-Zahl OHNE Telekom-Zielwert. Jetzt nur die zwei
                    Kennzahlen, die einen Zielwert haben, und nur wenn genug
                    Rueckmeldungen dahinterstehen. */}
                {/* Das NPS-Schild. Hier stand frueher ein IIFE mit eigenen
                    Zielwerten ("?? 68") - und davor nahm es sogar den
                    ERSTBESTEN Wert aus [nps, nps_pb, nps_montage] gegen 67,
                    also ausgerechnet NPS Schalten, das gar keinen Zielwert hat.
                    Jetzt: erstes NPS, zu dem [U] ueberhaupt ein Urteil hat. */}
                {(() => {
                  const kand = [["nps_pb", "PB"], ["nps_montage", "Montage"]]
                    .map(([id, was]) => ({ ...kpiAusU(tech, id), was }))
                    .find(k => k.ampel !== null);
                  if (!kand) return null;
                  const c = kand.ampel === "kritisch" ? "#f87171" : kand.ampel === "warnung" ? "#fbbf24" : "#4ade80";
                  return <span title={`NPS ${kand.was} aus ${kand.menge} Rueckmeldungen`}
                    style={{ fontSize: 10, background: "#1f2937", color: c, padding: "2px 8px", borderRadius: 3, fontWeight: 700 }}>
                    NPS {kand.was} {kand.wert.toFixed(0)} / &gt;={kand.ziel}</span>;
                })()}
                {/* One Touch hat keinen Zielwert - im Bonus-Malus-Vertrag
                    kommt es nicht vor. Also Zahl ohne Farbe, ohne Urteil. */}
                {tech.a1 !== null && tech.a1 !== undefined && <span style={{ fontSize: 10, background: "#1f2937", color: "#9ca3af", padding: "2px 8px", borderRadius: 3 }}>A1 {tech.a1.toFixed(1)}%</span>}
                {(() => { const k = kpiAusU(tech, "nftq_b"); if (k.wert === null) return null; const c = k.ampel === "kritisch" ? "#f87171" : k.ampel === "warnung" ? "#fbbf24" : k.ampel === "gut" ? "#4ade80" : "#9ca3af"; return <span title={k.grund || undefined} style={{ fontSize: 10, background: "#1f2937", color: c, padding: "2px 8px", borderRadius: 3, fontWeight: k.ampel ? 700 : 400 }}>NFTQ-B {k.wert.toFixed(1)}%</span>; })()}
                {(() => { const k = kpiAusU(tech, "nftq_s"); if (k.wert === null) return null; const c = k.ampel === "kritisch" ? "#f87171" : k.ampel === "warnung" ? "#fbbf24" : k.ampel === "gut" ? "#4ade80" : "#9ca3af"; return <span title={k.grund || undefined} style={{ fontSize: 10, background: "#1f2937", color: c, padding: "2px 8px", borderRadius: 3, fontWeight: k.ampel ? 700 : 400 }}>NFTQ-S {k.wert.toFixed(1)}%</span>; })()}
                {(() => { const k = kpiAusU(tech, "nftq_m"); if (k.wert === null) return null; const c = k.ampel === "kritisch" ? "#f87171" : k.ampel === "warnung" ? "#fbbf24" : k.ampel === "gut" ? "#4ade80" : "#9ca3af"; return <span title={k.grund || undefined} style={{ fontSize: 10, background: "#1f2937", color: c, padding: "2px 8px", borderRadius: 3, fontWeight: k.ampel ? 700 : 400 }}>NFTQ-M {k.wert.toFixed(1)}%</span>; })()}
                {(() => { const k = kpiAusU(tech, "nftq_p"); if (k.wert === null) return null; const c = k.ampel === "kritisch" ? "#f87171" : k.ampel === "warnung" ? "#fbbf24" : k.ampel === "gut" ? "#4ade80" : "#9ca3af"; return <span title={k.grund || undefined} style={{ fontSize: 10, background: "#1f2937", color: c, padding: "2px 8px", borderRadius: 3, fontWeight: k.ampel ? 700 : 400 }}>NFTQ-P {k.wert.toFixed(1)}%</span>; })()}
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
                  <a href={`https://wa.me/${k.mobil.replace(/[^0-9]/g, "")}?text=${encodeURIComponent(`Hallo ${tech.name.split(" ")[0]}, Ihre KPI: ${bilanzText(tech)} bewerteten Kennzahlen im Ziel - ${ampelWort(ampel)}. ${bew?.massnahme || ""}`)}`}
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

      {/* Alte Daten im Speicher - die auffaelligste Stelle, die es gibt.
          Wer das uebersieht, sucht den Fehler in der Datei, obwohl er im
          Browser liegt. */}
      {datenAlt && (
        <div style={{ background: "#2e1f00", borderBottom: "2px solid #b45309", padding: "10px 16px",
          display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
          <span style={{ fontSize: 12, color: "#fbbf24", fontWeight: 700 }}>
            Die gespeicherten Zahlen stammen aus einer aelteren Fassung
          </span>
          <span style={{ fontSize: 11, color: "#d1d5db", maxWidth: 620, lineHeight: 1.5 }}>
            Sie wurden eingelesen, als der Agent "0 Rueckmeldungen" noch als "unbekannt"
            gespeichert hat - dadurch bewertet er Techniker, ueber die es gar keine
            Rueckmeldung gibt. Die CSV-Dateien auf deinem Rechner sind in Ordnung; nur
            was hier liegt, ist es nicht. Einmal neu hochladen, dann stimmt es.
          </span>
          <label style={{ background: "#b45309", color: "#fff", padding: "6px 12px", borderRadius: 6,
            cursor: "pointer", fontSize: 11, fontWeight: 700, whiteSpace: "nowrap" }}>
            Report-CSVs neu laden
            <input type="file" multiple accept=".csv,.xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
          </label>
          <button onClick={() => setDatenAlt(false)}
            style={{ background: "none", border: "1px solid #78350f", color: "#9ca3af",
              padding: "6px 10px", borderRadius: 6, cursor: "pointer", fontSize: 11 }}>
            Spaeter
          </button>
        </div>
      )}

      {/* KPI Warnungsleiste */}
      {(() => {
        const alleTechs = Object.values(gespeichert).flat().filter((t, idx, arr) => arr.findLastIndex(x => namensSchluessel(x.name) === namensSchluessel(t.name) && String(x.standort) === String(t.standort)) === idx);
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
                ? Object.values(gespeichert).flat().filter((t, idx, arr) => arr.findLastIndex(x => namensSchluessel(x.name) === namensSchluessel(t.name) && String(x.standort) === String(t.standort)) === idx).length
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
                style={{ padding: "3px 4px 3px 10px", borderRadius: 5, fontSize: 11,
                  background: "#0f2e1a", color: "#4ade80", border: "1px solid #14532d",
                  display: "inline-flex", alignItems: "center", gap: 6 }}>
                Ursachen ({new Set(ursachen.map(u => u.name)).size} Techniker, {ursachen.length} Befunde)
                {/* Loeschen muss gehen wie bei jeder Kategorie - und hier
                    besonders: das sind Kundentexte im Klartext. Telekom schreibt
                    vor, sie zu loeschen, sobald der Zweck erfuellt ist. Nach dem
                    Gespraech mit dem Techniker also weg. Ohne diesen Knopf
                    blieben sie im Browser liegen. */}
                <button onClick={() => {
                  if (window.confirm("Ursachenbericht aus dem Agenten entfernen?\n\n" +
                    ursachen.length + " Befunde mit Kundentexten. Die Datei auf deinem " +
                    "Rechner bleibt - hier im Browser ist sie dann weg.")) {
                    setUrsachen([]);
                    setError("Ursachenbericht entfernt.");
                    if (activeTab === "berichte") setActiveTab("dashboard");
                  }
                }}
                  title="Ursachenbericht entfernen - Kundentexte gehoeren nach dem Gespraech geloescht"
                  style={{ background: "none", border: "none", color: "#4ade80", cursor: "pointer",
                    fontSize: 13, lineHeight: 1, padding: "0 4px", opacity: 0.6 }}>
                  ×
                </button>
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
          {/* Techniker-Auswahl - gilt fuer ALLE Reiter, nicht nur fuer den
              Bericht. Wer ueber einen Mann redet, will ihn in allen vier
              Bereichen sehen, ohne ihn viermal zu suchen. */}
          {alleNamen.length > 0 && (
            <select value={nurTechniker} onChange={e => setNurTechniker(e.target.value)}
              title="Einen Techniker ueberall anzeigen - Dashboard, alle Kennzahlen und Bericht"
              style={{ background: nurTechniker ? "#1e3a5f" : "#111827",
                color: nurTechniker ? "#dbeafe" : "#9ca3af",
                border: `1px solid ${nurTechniker ? "#2563eb" : "#374151"}`,
                padding: "6px 10px", borderRadius: 6, cursor: "pointer", fontSize: 11,
                fontWeight: nurTechniker ? 600 : 400, maxWidth: 200 }}>
              <option value="">Alle Techniker</option>
              {alleNamen.map(([k, name]) => <option key={k} value={k}>{name}</option>)}
            </select>
          )}
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

        {/* Diese Kachel MUSS auch erscheinen, wenn nur der Ursachenbericht
            geladen ist. Vorher stand hier "ursachen.length === 0" - damit war
            in genau dem Fall (Befunde da, Kennzahlen weg) KEIN Hochlade-Knopf
            mehr auf der Seite ausser dem kleinen in der Kopfleiste. Arash sah
            nur noch "Team-Analyse starten" und kam nicht mehr weiter. */}
        {angezeigt.length === 0 && (hatDaten || ursachen.length > 0) && (
          <div style={{ textAlign: "center", padding: "40px 20px", background: "#111827",
            border: "1px solid #1f2937", borderRadius: 8, marginBottom: 16 }}>
            <div style={{ fontSize: 14, color: "#f9fafb", fontWeight: 700, marginBottom: 6 }}>
              {ursachen.length > 0 && !hatDaten
                ? "Kennzahlen fehlen - der Ursachenbericht ist da"
                : (gespeichert[aktiveKategorie] || []).length > 0
                ? `${(gespeichert[aktiveKategorie] || []).length} Techniker vorhanden - alle durch die Filter oben ausgeblendet`
                : `Noch keine Daten für "${KATEGORIEN.find(k => k.id === aktiveKategorie)?.label}"`}
            </div>
            {ursachen.length > 0 && !hatDaten && (
              <div style={{ fontSize: 12, color: "#6b7280", marginBottom: 16, lineHeight: 1.6 }}>
                Die Berichte je Techniker stehen im Reiter <b style={{ color: "#9ca3af" }}>Berichte</b> - die
                brauchen keine Kennzahlen.<br />
                Für Dashboard, KI-Analyse und den Zahlen-Teil der Mail fehlen die vier Report-CSVs
                aus <code style={{ color: "#9ca3af" }}>Auftragsinfo_Downloads\&lt;KW&gt;\</code>:<br />
                <span style={{ color: "#9ca3af" }}>_sms_feedback · _sms_feedback_schalten · _nftq · _one_touch</span><br />
                Keine Datei mit <code style={{ color: "#9ca3af" }}>_details</code> im Namen - die gehören nicht hierher.
              </div>
            )}
            <label style={{ display: "inline-block", background: "#1d4ed8", color: "#fff", padding: "10px 24px", borderRadius: 6, cursor: "pointer", fontSize: 13, fontWeight: 600 }}>
              Report-CSVs hochladen (mehrere auf einmal)
              <input type="file" multiple accept=".csv,.xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
            </label>
          </div>
        )}

        {(angezeigt.length > 0 || ursachen.length > 0) && (
          <>
            {pending && <div style={{ fontSize: 11, color: "#4ade80", marginBottom: 8 }}>... Nächste Datei bereit</div>}
            {error && <div style={{ fontSize: 11, color: "#4ade80", marginBottom: 8 }}>{error}</div>}

            <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 10, marginBottom: 20 }}>
              {/* DIE KOPFKACHELN LOGEN (17.07.2026): "Kritisch" zeigte eine 0 und
                  faerbte sie GRUEN, auch wenn gar keine Uebersicht geladen war -
                  dann liefert techWorst ueberall null, criticalCount ist 0, und
                  oben stand in Gruen "Kritisch 0". Das heisst aber nicht "keiner
                  ist kritisch", sondern "ich weiss es nicht". Daneben stand
                  "Avg Score 10.0/10" fuer Adil Kheder, dessen NPS PB auf EINER
                  Rueckmeldung beruht. Jetzt: ohne Urteil ein grauer Strich. */}
              {(isOTView ? [
                { label: "Techniker", value: angezeigt.length, color: "#60a5fa" },
                { label: "Kritisch", value: bewertetCount ? criticalCount : "-", color: !bewertetCount ? "#6b7280" : criticalCount > 0 ? "#f87171" : "#4ade80" },
                { label: "Avg A1-Rate", value: avg("a1") !== "-" ? avg("a1") + "%" : "-", color: "#9ca3af" },
                { label: "Im Ziel", value: teamBilanz(), color: "#9ca3af" },
              ] : [
                { label: "Techniker", value: angezeigt.length, color: "#60a5fa" },
                { label: "Kritisch", value: bewertetCount ? criticalCount : "-", color: !bewertetCount ? "#6b7280" : criticalCount > 0 ? "#f87171" : "#4ade80" },
                { label: "Avg CC-Rate", value: avg("cc_rate") !== "-" ? avg("cc_rate") + "%" : "-", color: "#fbbf24" },
                { label: "Im Ziel", value: teamBilanz(), color: "#9ca3af" },
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
                    ursachen={(ursachen || []).filter(u => {
                      if (namensSchluessel(u.name) !== namensSchluessel(t.name)) return false;
                      // Alae arbeitet in ATS 35 UND 36 - er hat zwei Karten.
                      // Ohne diese Pruefung standen auf beiden dieselben 15
                      // Befunde, auch die aus dem jeweils anderen Bereich.
                      if (!u.ats || !t.standort) return true;
                      return String(t.standort).endsWith(String(u.ats));  // 5335 endet auf 35
                    })} />;
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
                kontakte={kontakte} nurTechniker={nurTechniker} onWaehlen={setNurTechniker} />
            )}
            {activeTab === "firmendashboard" && <FirmendashboardTab />}

            {/* Warnung, wenn die gezeigte Analyse zu einer anderen Auswahl
                gehoert als der, die gerade eingestellt ist. Eine alte Analyse
                sieht genauso aus wie eine frische - man erkennt es nur an den
                Zahlen darin, und die liest niemand gegen. */}
            {activeTab === "analyse" && aiAnalysis && analyseBasis
              && (analyseBasis.anzahl !== angezeigt.length
                || analyseBasis.kategorie !== aktiveKategorie
                || analyseBasis.techniker !== nurTechniker) && (
              <div style={{ background: "#2e1f00", border: "1px solid #78350f", borderRadius: 8,
                padding: "10px 14px", marginBottom: 12, fontSize: 12, color: "#fbbf24" }}>
                <b>Diese Analyse gehoert zu einer anderen Auswahl.</b><br />
                <span style={{ color: "#d1d5db", fontSize: 11 }}>
                  Gerechnet wurde ueber {analyseBasis.anzahl} Techniker
                  ({KATEGORIEN.find(k => k.id === analyseBasis.kategorie)?.label || analyseBasis.kategorie}
                  {analyseBasis.techniker ? ", ein Techniker gefiltert" : ""}).
                  Jetzt eingestellt: {angezeigt.length} Techniker
                  ({KATEGORIEN.find(k => k.id === aktiveKategorie)?.label || aktiveKategorie}
                  {nurTechniker ? ", ein Techniker gefiltert" : ""}).
                  Die KI rechnet nicht mit - im Dashboard "Team-Analyse starten" druecken.
                </span>
              </div>
            )}
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
