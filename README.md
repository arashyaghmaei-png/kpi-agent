# KPI Agent – FiberNC
## Deployment auf Vercel (Schritt für Schritt)

---

### Was du brauchst
- Einen kostenlosen Account auf vercel.com
- Einen kostenlosen Account auf github.com
- Deinen Anthropic API-Key (von platform.anthropic.com)

---

### Schritt 1: GitHub-Repository erstellen
1. Geh auf github.com → "New repository"
2. Name: `kpi-agent`
3. Klick "Create repository"
4. Lade alle Dateien aus diesem Ordner hoch (drag & drop)

---

### Schritt 2: Vercel verbinden
1. Geh auf vercel.com → "Add New Project"
2. Wähle dein GitHub-Repository `kpi-agent`
3. Klick "Deploy" — Vercel erkennt Next.js automatisch

---

### Schritt 3: API-Key hinterlegen (WICHTIG)
1. In Vercel: Dein Projekt → "Settings" → "Environment Variables"
2. Name: `ANTHROPIC_API_KEY`
3. Value: dein API-Key (beginnt mit `sk-ant-...`)
4. Klick "Save"
5. Dann: "Deployments" → "Redeploy"

---

### Schritt 4: Fertig
Du bekommst eine URL wie: `https://kpi-agent-fibernc.vercel.app`
Diese URL kannst du auf jedem Gerät öffnen — Handy, Tablet, PC.

---

### CSV-Format für tägliche Eingabe
```
name,standort,cc_rate,termintreue,loesungsquote,auftraege
Max M.,5335,95,97,92,5
Anna K.,5336,88,91,89,4
```

---

### Kosten
- Vercel: kostenlos (Hobby-Plan reicht)
- GitHub: kostenlos
- Anthropic API: ca. 0,01–0,05 EUR pro Analyse
