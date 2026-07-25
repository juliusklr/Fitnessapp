# Training – Garmin Epix 2 App

Connect-IQ-Watch-App zum Loggen von Sätzen direkt am Handgelenk. Nutzt dasselbe
Supabase-Backend wie die Web-App (Login per E-Mail/Passwort, RLS-geschützt).

## Funktionsumfang

- Zeigt beim Start den für heute geplanten Trainingstag (Planung → `schedule_entries`
  bzw. Wochenmuster), sonst die Liste aller Trainingstage.
- Pro Übung: Ziel (Sätze/Wdh/Runden), letzte Session als Vorbelegung,
  KG/WDH einstellen, START loggt den Satz in `log_sets`.
- Bedienung: UP/DOWN = Wert ±, langes UP (Menü) = Feld wechseln,
  Touch: Box antippen = Feld wählen, gewählte Box oben/unten antippen = ±,
  START = Satz loggen, BACK = zurück.

Internet läuft über das gekoppelte Handy (Garmin Connect Mobile muss laufen).

## Build

Läuft automatisch in GitHub Actions (`.github/workflows/garmin-build.yml`):
SDK + Gerätedateien werden per `connect-iq-sdk-manager-cli` geladen, die
Supabase-Zugangsdaten kommen aus den Repo-Secrets `SUPABASE_EMAIL` /
`SUPABASE_PASSWORD` und werden zur Buildzeit in `resources/properties.xml`
eingesetzt (Sideload-Apps können keine Einstellungen über Garmin Connect
Mobile beziehen).

## Installation (Sideload)

1. Im Actions-Run das Artefakt **Training-Garmin** herunterladen.
2. Uhr per USB verbinden, `Training-epix2.prg` nach `GARMIN/Apps/` kopieren
   (Epix 2 Pro: passende `epix2pro…`-Variante).
3. Uhr neu starten – die App erscheint in der Aktivitäten-/App-Liste.
