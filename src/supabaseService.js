// Supabase data layer. Replaces the old Microsoft-Graph/Excel service.
// Tables: exercises · programs · phases · plans (= Trainingstage) · plan_items
// · log_sets · schedule_entries · weekly_pattern (see BACKEND_PLAN.md).
// Hierarchie: programs (Trainingspläne) → phases → plans (Tage) → plan_items.
// RLS scopes every row to owner = auth.uid(); owner is set by column default.
import { supabase } from './supabaseClient';

// ── coercion helpers ────────────────────────────────────────────
const s = (v) => (v === null || v === undefined ? '' : String(v)); // null → '' for the UI
const toInt = (v) => {
  if (v === '' || v === null || v === undefined) return null;
  const n = parseInt(v, 10);
  return Number.isFinite(n) ? n : null;
};
const toNum = (v) => {
  if (v === '' || v === null || v === undefined) return null;
  const n = parseFloat(String(v).replace(',', '.'));
  return Number.isFinite(n) ? n : null;
};

function fail(error, ctx) {
  if (error) throw new Error(`${ctx}: ${error.message}`);
}

// ═══════════════════════════════════════════════════════════════
//  LIBRARY (exercises)
// ═══════════════════════════════════════════════════════════════
const mapExercise = (r) => ({
  id: r.id,
  uebung: s(r.name),
  cues: s(r.cues),
  kategorie: s(r.category),
  equipment: s(r.equipment),
  video: s(r.video_url),
});

export async function getExercises() {
  const { data, error } = await supabase
    .from('exercises')
    .select('id, name, cues, category, equipment, video_url')
    .order('name');
  fail(error, 'getExercises');
  return (data || []).map(mapExercise);
}

export async function addExercise(e) {
  const { data, error } = await supabase
    .from('exercises')
    .insert({ name: e.uebung, cues: e.cues || null })
    .select()
    .single();
  fail(error, 'addExercise');
  return mapExercise(data);
}

export async function updateExercise(id, e) {
  const patch = {};
  if ('uebung' in e) patch.name = e.uebung;
  if ('cues' in e) patch.cues = e.cues || null;
  if ('kategorie' in e) patch.category = e.kategorie || null;
  if ('equipment' in e) patch.equipment = e.equipment || null;
  if ('video' in e) patch.video_url = e.video || null;
  const { data, error } = await supabase
    .from('exercises')
    .update(patch)
    .eq('id', id)
    .select()
    .single();
  fail(error, 'updateExercise');
  return mapExercise(data);
}

export async function deleteExercise(id) {
  const { error } = await supabase.from('exercises').delete().eq('id', id);
  // on delete restrict: a 23503 means the exercise is still used in a plan.
  if (error && error.code === '23503') {
    throw new Error('Übung wird noch in einem Plan verwendet — erst dort entfernen.');
  }
  fail(error, 'deleteExercise');
}

// ═══════════════════════════════════════════════════════════════
//  PROGRAMS (programs → phases → plans/Tage + plan_items)
// ═══════════════════════════════════════════════════════════════
const byPos = (a, b) => (a.position || 0) - (b.position || 0);

const mapItem = (r) => ({
  id: r.id,
  exercise_id: r.exercise_id,
  uebung: s(r.exercise?.name),
  gruppe: s(r.gruppe),
  runden: r.runden ?? '',
  zielSaetze: r.ziel_saetze ?? '',
  zielWdh: s(r.ziel_wdh),
  tempo: s(r.tempo),
  pause: s(r.pause),
  notiz: s(r.notiz),
});

const mapDay = (p) => ({
  id: p.id,
  name: p.name,
  notiz: s(p.notiz),
  position: p.position || 0,
  items: (p.plan_items || []).sort(byPos).map(mapItem),
});

// Whole hierarchy in one query: programs → phases → days → items.
export async function getPrograms() {
  const { data, error } = await supabase
    .from('programs')
    .select(
      'id, name, notiz, position, created_at, phases(id, name, notiz, position, created_at, ' +
        'plans(id, name, notiz, position, created_at, ' +
        'plan_items(id, exercise_id, position, gruppe, runden, ziel_saetze, ziel_wdh, tempo, pause, notiz, exercise:exercises(name))))'
    )
    .order('position');
  fail(error, 'getPrograms');
  return (data || []).map((pr) => ({
    id: pr.id,
    name: pr.name,
    notiz: s(pr.notiz),
    position: pr.position || 0,
    phases: (pr.phases || []).sort(byPos).map((ph) => ({
      id: ph.id,
      name: ph.name,
      notiz: s(ph.notiz),
      position: ph.position || 0,
      days: (ph.plans || []).sort(byPos).map(mapDay),
    })),
  }));
}

export async function saveProgram({ id, name, notiz }) {
  const row = { name, notiz: notiz || null };
  const q = id
    ? supabase.from('programs').update(row).eq('id', id)
    : supabase.from('programs').insert(row);
  const { error } = await q;
  fail(error, 'saveProgram');
}

export async function deleteProgram(id) {
  const { error } = await supabase.from('programs').delete().eq('id', id); // cascades phases → days → items
  fail(error, 'deleteProgram');
}

export async function savePhase({ id, program_id, name, notiz, position }) {
  const row = { name, notiz: notiz || null };
  if (position !== undefined) row.position = position;
  let q;
  if (id) q = supabase.from('phases').update(row).eq('id', id);
  else q = supabase.from('phases').insert({ ...row, program_id });
  const { error } = await q;
  fail(error, 'savePhase');
}

export async function deletePhase(id) {
  const { error } = await supabase.from('phases').delete().eq('id', id); // cascades days → items
  fail(error, 'deletePhase');
}

// Create or replace a day (plan) and all its items in one go.
export async function savePlan({ id, phase_id, name, notiz, position, items }) {
  let planId = id;
  if (planId) {
    const { error } = await supabase.from('plans').update({ name, notiz: notiz || null }).eq('id', planId);
    fail(error, 'savePlan(update)');
    const { error: delErr } = await supabase.from('plan_items').delete().eq('plan_id', planId);
    fail(delErr, 'savePlan(clear items)');
  } else {
    const { data, error } = await supabase
      .from('plans')
      .insert({ name, notiz: notiz || null, phase_id, position: position ?? 0 })
      .select('id')
      .single();
    fail(error, 'savePlan(insert)');
    planId = data.id;
  }
  if (items.length) {
    const rows = items.map((it, i) => ({
      plan_id: planId,
      exercise_id: it.exercise_id,
      position: i + 1,
      gruppe: it.gruppe || null,
      runden: toInt(it.runden),
      ziel_saetze: it.zielSaetze || null,
      ziel_wdh: it.zielWdh || null,
      tempo: it.tempo || null,
      pause: it.pause || null,
      notiz: it.notiz || null,
    }));
    const { error } = await supabase.from('plan_items').insert(rows);
    fail(error, 'savePlan(items)');
  }
  return planId;
}

export async function deletePlan(id) {
  const { error } = await supabase.from('plans').delete().eq('id', id); // cascades to plan_items
  fail(error, 'deletePlan');
}

// ═══════════════════════════════════════════════════════════════
//  LOG (log_sets)
// ═══════════════════════════════════════════════════════════════
const mapLog = (r) => ({
  id: r.id,
  datum: s(r.datum),
  plan: s(r.plan_name),
  plan_id: r.plan_id || null,
  exercise_id: r.exercise_id,
  // current name follows renames; falls back to the snapshot for deleted exercises
  uebung: r.exercise?.name || s(r.exercise_name),
  satz: r.satz ?? '',
  gewicht: r.gewicht ?? '',
  wdh: r.wdh ?? '',
  dauer: s(r.dauer),
  rpe: r.rpe ?? '',
  notiz: s(r.notiz),
});

export async function getLog() {
  const { data, error } = await supabase
    .from('log_sets')
    .select(
      'id, datum, plan_name, plan_id, exercise_id, exercise_name, satz, gewicht, wdh, dauer, rpe, notiz, exercise:exercises(name)'
    )
    .order('datum');
  fail(error, 'getLog');
  return (data || []).map(mapLog);
}

export async function addLogRow(row) {
  const { data, error } = await supabase
    .from('log_sets')
    .insert({
      datum: row.datum,
      plan_name: row.plan || null,
      plan_id: row.plan_id || null,
      exercise_id: row.exercise_id || null,
      exercise_name: row.uebung,
      satz: toInt(row.satz),
      gewicht: toNum(row.gewicht),
      wdh: toNum(row.wdh),
      dauer: row.dauer || null,
      rpe: toNum(row.rpe),
      notiz: row.notiz || null,
    })
    .select('id, datum, plan_name, plan_id, exercise_id, exercise_name, satz, gewicht, wdh, dauer, rpe, notiz')
    .single();
  fail(error, 'addLogRow');
  return mapLog(data);
}

export async function deleteLogRow(id) {
  const { error } = await supabase.from('log_sets').delete().eq('id', id);
  fail(error, 'deleteLogRow');
}

// ── Most recent logged session for an exercise, before a given date ──
// Matches by exercise name (already resolved to the current name in mapLog),
// so renames keep history attached.
export function lastSession(log, uebung, beforeISO) {
  const rows = log.filter((r) => r.uebung === uebung && r.datum && (!beforeISO || r.datum < beforeISO));
  if (!rows.length) return null;
  const maxDate = rows.reduce((m, r) => (r.datum > m ? r.datum : m), rows[0].datum);
  const sets = rows
    .filter((r) => r.datum === maxDate)
    .sort((a, b) => (parseInt(a.satz) || 0) - (parseInt(b.satz) || 0));
  return { datum: maxDate, sets };
}

// ═══════════════════════════════════════════════════════════════
//  SCHEDULE (schedule_entries + weekly_pattern)
// ═══════════════════════════════════════════════════════════════
// Auflösung für ein Datum: expliziter Eintrag gewinnt (plan_id null =
// bewusst trainingsfrei), sonst greift das Wochenmuster (weekday 0 = Montag).

export async function getSchedule() {
  const [entriesRes, patternRes] = await Promise.all([
    supabase.from('schedule_entries').select('id, datum, plan_id, notiz').order('datum'),
    supabase.from('weekly_pattern').select('id, weekday, plan_id'),
  ]);
  fail(entriesRes.error, 'getSchedule(entries)');
  fail(patternRes.error, 'getSchedule(pattern)');
  return {
    entries: (entriesRes.data || []).map((r) => ({ id: r.id, datum: s(r.datum), plan_id: r.plan_id, notiz: s(r.notiz) })),
    pattern: (patternRes.data || []).map((r) => ({ id: r.id, weekday: r.weekday, plan_id: r.plan_id })),
  };
}

// planId: uuid = Trainingstag, null = explizit frei (überschreibt Muster).
export async function setScheduleEntry(datum, planId) {
  const { error: delErr } = await supabase.from('schedule_entries').delete().eq('datum', datum);
  fail(delErr, 'setScheduleEntry(clear)');
  const { error } = await supabase.from('schedule_entries').insert({ datum, plan_id: planId });
  fail(error, 'setScheduleEntry');
}

// Eintrag entfernen → Datum fällt zurück aufs Wochenmuster.
export async function clearScheduleEntry(datum) {
  const { error } = await supabase.from('schedule_entries').delete().eq('datum', datum);
  fail(error, 'clearScheduleEntry');
}

// planId null = Wochentag im Muster frei lassen (Zeile löschen).
export async function setPatternDay(weekday, planId) {
  const { error: delErr } = await supabase.from('weekly_pattern').delete().eq('weekday', weekday);
  fail(delErr, 'setPatternDay(clear)');
  if (planId) {
    const { error } = await supabase.from('weekly_pattern').insert({ weekday, plan_id: planId });
    fail(error, 'setPatternDay');
  }
}

// ═══════════════════════════════════════════════════════════════
//  EXCEL EXPORT (on-demand .xlsx, three sheets) — SheetJS
// ═══════════════════════════════════════════════════════════════
export async function exportToExcel({ exercises, plans, log }) {
  const XLSX = await import('xlsx');
  const wb = XLSX.utils.book_new();

  const ueb = exercises.map((e) => ({ Übung: e.uebung, Kategorie: e.kategorie, Equipment: e.equipment, Cues: e.cues, Video: e.video }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(ueb), 'Übungen');

  const pl = [];
  plans.forEach((p) =>
    p.items.forEach((it, i) =>
      pl.push({ Plan: p.name, Reihenfolge: i + 1, Gruppe: it.gruppe, Runden: it.runden, Übung: it.uebung, 'Ziel-Sätze': it.zielSaetze, 'Ziel-Wdh': it.zielWdh, Tempo: it.tempo, Pause: it.pause, Notiz: it.notiz })
    )
  );
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(pl), 'Pläne');

  const lg = log.map((r) => ({ Datum: r.datum, Plan: r.plan, Übung: r.uebung, Satz: r.satz, Gewicht: r.gewicht, Wdh: r.wdh, Dauer: r.dauer, RPE: r.rpe, Notiz: r.notiz }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(lg), 'Log');

  XLSX.writeFile(wb, `Training-Export-${new Date().toISOString().slice(0, 10)}.xlsx`);
}
