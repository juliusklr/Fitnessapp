# Training Tracker 🏋️

Eine mobile-optimierte Web-App (PWA), die Trainingsdaten direkt in deine Excel-Datei auf OneDrive schreibt.

## Setup

### Voraussetzungen
- Node.js (v18+) installiert
- Azure App Registration mit Client ID `812777cb-5b2f-4d0b-9e23-5708c6df6948`
- OneDrive Excel-Datei mit "Masterplan"-Sheet

### Lokal starten

```bash
npm install
npm run dev
```

Die App läuft dann unter `http://localhost:5173`

### Auf Vercel deployen (kostenlos)

1. Erstelle ein GitHub-Repository und pushe den Code
2. Gehe zu [vercel.com](https://vercel.com) und verbinde das Repo
3. Deploy klicken – fertig!
4. **Wichtig**: Füge die Vercel-URL als Redirect URI in der Azure App Registration hinzu:
   - Azure Portal → App registrations → Training Tracker → Authentication
   - "Add URI" → `https://dein-projekt.vercel.app`

### Als PWA auf iPhone installieren

1. Öffne die App-URL in Safari
2. Tippe auf das **Teilen**-Symbol (Quadrat mit Pfeil nach oben)
3. Wähle **"Zum Home-Bildschirm"**
4. Die App erscheint als Icon auf deinem Homescreen

## Funktionen

- 📅 **Datumsauswahl** – Standard: heute, Quick-Chips für die letzten Tage
- 🏃 **Tägliche Aktivitäten** – Run, Bike, VB, DGR, Mob, etc.
- 💪 **Trainingsplan-Auswahl** – zwischen deinen Plänen wechseln
- 📝 **Übungs-Tracking** – Freitexteingabe pro Übung mit Sätzen, Wiederholungen, Timing, Pause
- ☁️ **Direkte Excel-Sync** – Daten werden sofort in OneDrive geschrieben
- 📱 **PWA** – installierbar als Home-Screen-App

## Architektur

- **React 18** + Vite
- **MSAL** für Microsoft OAuth
- **Microsoft Graph API** für Excel-Operationen
- **PWA** mit vite-plugin-pwa
