# Backend-Umstieg: Excel → Supabase

Entwurf des Datenmodells und Migrationsplans. **Noch nichts umgesetzt** — erst zum Durchschauen.

---

## 1. Grundprinzip

Das heutige Problem: Bibliothek und Pläne hängen nur über den **Übungs-Namen als Text** zusammen. Umbenennen bricht die Verknüpfung.

Lösung: stabile **IDs** (UUIDs) als Fremdschlüssel. Zwei Datenarten werden bewusst unterschiedlich behandelt:

- **Lebende, editierbare Daten** (Übungen, Pläne, Plan-Einträge) → echte Fremdschlüssel. Umbenennen wirkt überall sofort.
- **Historie / Log** (geloggte Sätze) → Fremdschlüssel **plus** Namens-Snapshot. So folgt die Auswertung einem Umbenennen, eine *gelöschte* Übung verliert aber ihre Historie nicht.

---

## 2. Tabellen

### `exercises` (Bibliothek)
| Spalte | Typ | Hinweis |
|---|---|---|
| `id` | uuid PK | `default gen_random_uuid()` |
| `name` | text NOT NULL | |
| `cues` | text | Hinweise, eine pro Zeile |
| `category` | text | optional (heute „Kategorie") |
| `equipment` | text | optional |
| `video_url` | text | optional |
| `owner` | uuid | `default auth.uid()` (RLS) |
| `created_at` | timestamptz | `default now()` |
| `updated_at` | timestamptz | per Trigger aktualisiert |

### `plans` (Pläne)
| Spalte | Typ | Hinweis |
|---|---|---|
| `id` | uuid PK | |
| `name` | text NOT NULL | |
| `position` | int | Sortierung der Pläne in der UI |
| `owner` | uuid | `default auth.uid()` |
| `created_at` / `updated_at` | timestamptz | |

### `plan_items` (eine Zeile pro Übung-im-Plan)
| Spalte | Typ | Hinweis |
|---|---|---|
| `id` | uuid PK | |
| `plan_id` | uuid FK → `plans.id` | `on delete cascade` |
| `exercise_id` | uuid FK → `exercises.id` | `on delete restrict` (Übung in Plan kann nicht versehentlich verwaisen) |
| `position` | int | heute „Reihenfolge" |
| `gruppe` | text | Superset-Buchstabe A/B/… (leer = Einzelübung) |
| `runden` | int | Runden für Superset |
| `ziel_saetze` | int | |
| `ziel_wdh` | text | z. B. „8-12" → bleibt Text |
| `tempo` | text | |
| `pause` | text | |
| `notiz` | text | überschreibt Cues im Workout |

### `log_sets` (Trainings-Log, append-only, eine Zeile pro Satz)
| Spalte | Typ | Hinweis |
|---|---|---|
| `id` | uuid PK | |
| `datum` | date NOT NULL | |
| `plan_name` | text | Snapshot inkl. „Aktivität" — bleibt korrekt, auch wenn Plan später umbenannt/gelöscht wird |
| `exercise_id` | uuid FK → `exercises.id` | nullable, `on delete set null` |
| `exercise_name` | text NOT NULL | Snapshot/Fallback; bei „Aktivität" steht hier das Label |
| `satz` | int | |
| `gewicht` | numeric | |
| `wdh` | numeric | |
| `dauer` | text | |
| `rpe` | numeric | |
| `notiz` | text | |
| `owner` | uuid | `default auth.uid()` |
| `created_at` | timestamptz | |

**Anzeige-Logik:** Übungsname = aktueller Name aus `exercises` (via `exercise_id`), sonst Fallback `exercise_name`. → Umbenennen folgt der Verlaufs-Auswertung; gelöschte Übung behält ihren Snapshot-Namen.

---

## 3. SQL (DDL-Auszug, zur Veranschaulichung)

```sql
create extension if not exists pgcrypto;

create table exercises (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  cues text,
  category text,
  equipment text,
  video_url text,
  owner uuid not null default auth.uid(),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table plans (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  position int default 0,
  owner uuid not null default auth.uid(),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table plan_items (
  id uuid primary key default gen_random_uuid(),
  plan_id uuid not null references plans(id) on delete cascade,
  exercise_id uuid not null references exercises(id) on delete restrict,
  position int not null default 0,
  gruppe text,
  runden int,
  ziel_saetze int,
  ziel_wdh text,
  tempo text,
  pause text,
  notiz text
);

create table log_sets (
  id uuid primary key default gen_random_uuid(),
  datum date not null,
  plan_name text,
  exercise_id uuid references exercises(id) on delete set null,
  exercise_name text not null,
  satz int,
  gewicht numeric,
  wdh numeric,
  dauer text,
  rpe numeric,
  notiz text,
  owner uuid not null default auth.uid(),
  created_at timestamptz not null default now()
);

create index on log_sets (owner, datum);
create index on log_sets (exercise_id);
create index on plan_items (plan_id, position);

-- Row Level Security: nur der Eigentümer sieht/ändert seine Daten
alter table exercises  enable row level security;
alter table plans      enable row level security;
alter table plan_items enable row level security;
alter table log_sets   enable row level security;

create policy own_rows on exercises for all using (owner = auth.uid()) with check (owner = auth.uid());
create policy own_rows on plans     for all using (owner = auth.uid()) with check (owner = auth.uid());
create policy own_rows on log_sets  for all using (owner = auth.uid()) with check (owner = auth.uid());
-- plan_items erbt Schutz über plan_id; Policy prüft den Besitz des Plans
create policy own_rows on plan_items for all
  using (exists (select 1 from plans p where p.id = plan_id and p.owner = auth.uid()))
  with check (exists (select 1 from plans p where p.id = plan_id and p.owner = auth.uid()));
```

---

## 4. Migration der bestehenden Daten (einmalig)

Quelle: aktuelle `Masterplan.xlsx` (`tblUebungen`, `tblPlaene`, `tblLog`).

1. **Übungen einlesen** → `exercises`. Zusätzlich Namen sammeln, die nur in Plänen/Log vorkommen, aber nicht in der Bibliothek → als Stub-Übung anlegen, damit Fremdschlüssel auflösbar sind.
2. **Name → id Mapping** im Speicher aufbauen.
3. **Pläne**: distinkte Plannamen → `plans`.
4. **Plan-Einträge** → `plan_items`, `exercise_id` per Mapping aufgelöst, `position`/`gruppe`/`runden`/Ziele übernommen.
5. **Log** → `log_sets`: `exercise_id` per Mapping (wo möglich), `exercise_name` = Snapshot, `plan_name` = Snapshot (inkl. „Aktivität").

Skript in Python (openpyxl + `supabase-py` oder direkter Postgres-Insert). Die alte Excel bleibt unangetastet als Backup.

---

## 5. Frontend-Änderungen

- **Neu:** `src/supabaseService.js` ersetzt `graphService.js`, gleiche Funktions-Signaturen wo möglich:
  `getExercises / addExercise / updateExercise / deleteExercise`,
  `getPlans (inkl. items) / savePlan / deletePlan`,
  `getLog / addLogRow`,
  `exportToExcel`.
- **Auth:** MSAL/Azure raus, Supabase-Auth rein (Magic-Link an `julius.keller@yahoo.de` oder Google). Kein Azure-Redirect-URI-Theater mehr.
- **App.jsx:** Lookups laufen über `exercise_id` statt Name; sonst bleibt die UI nahezu gleich. Plan-Editor speichert beim Hinzufügen die `exercise_id`.
- **Dashboard:** Gruppierung der Progression über `exercise_id` (folgt Umbenennungen korrekt).

---

## 6. Excel-Export bleibt erhalten

Button „Export nach Excel" erzeugt **on-demand** eine `.xlsx` mit drei Blättern (Übungen / Pläne / Sessions), Layout wie heute — clientseitig per [SheetJS](https://sheetjs.com). Excel ist damit Export-Ziel statt Live-Speicher.

---

## 7. Was du tun musst

1. Auf [supabase.com](https://supabase.com) kostenloses Konto + Projekt anlegen (~5 Min).
2. Mir **Project URL** und **anon public key** geben (aus Project Settings → API). Der anon-Key ist für Client-Apps gedacht; RLS schützt die Daten.
3. Den Rest (Schema-SQL ausführen, Daten migrieren, Code umstellen) übernehme ich.

**Kosten:** Free-Tier (500 MB DB, reichlich für diese App). Kein Server-Betrieb nötig.

---

## 8. Zukunft: Weitergeben & Teilen

Wir bauen **jetzt nur für Julius**, aber mehrnutzer-fähig vorbereitet — ohne Mehraufwand heute:

- `owner`-Spalte + RLS sind bereits das Fundament. Jeder neue Login bekommt automatisch einen isolierten Datenbereich. **Kein Umbau** nötig, wenn die App weitergegeben wird.
- **Teilen** (z. B. einen Plan an einen Trainingspartner geben) und **Coach-Sicht** (Trainer sieht Logs seiner Athleten) sind später *additiv* nachrüstbar: eine `shares`-Tabelle (`resource_id`, `shared_with_user`, `role`) plus erweiterte RLS-Policies. Kein Neuschreiben des Modells.
- Praktische Konsequenz schon jetzt: keine globalen Annahmen „es gibt nur einen Nutzer" in den Code bauen — alles läuft über `auth.uid()`. Ist im Plan so vorgesehen.

---

## 9. Offene Entscheidungen

- **Auth-Methode:** Magic-Link (am einfachsten, gut für PWA) vs. Google-Login. Empfehlung: Magic-Link.
- **Aktivitäten:** als eigene Übungen mit `category = 'Aktivität'` führen, oder wie heute nur als Log-Einträge ohne `exercise_id`? Empfehlung: nur als Log (weniger Umbau).
- **Echtzeit-Sync** zwischen Geräten: Supabase kann das, aber optional. Empfehlung: später, erst mit optimistischen Updates wie heute starten.
