import { useState, useEffect, useCallback, useMemo, useRef } from 'react';
import { DndContext, closestCenter, PointerSensor, useSensor, useSensors } from '@dnd-kit/core';
import { SortableContext, verticalListSortingStrategy, useSortable, arrayMove } from '@dnd-kit/sortable';
import { CSS } from '@dnd-kit/utilities';
import { supabase } from './supabaseClient';
import {
  getExercises, addExercise, updateExercise, deleteExercise,
  getPlans, savePlan, deletePlan,
  getLog, addLogRow, lastSession, exportToExcel,
} from './supabaseService';

// ── Date helpers ────────────────────────────────────────────────
const isoDate = (d) => {
  const dt = new Date(d);
  return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}-${String(dt.getDate()).padStart(2, '0')}`;
};
const fmtDate = (iso) => {
  if (!iso) return '';
  const dt = new Date(iso + (iso.length === 10 ? 'T00:00:00' : ''));
  return dt.toLocaleDateString('de-DE', { weekday: 'short', day: '2-digit', month: '2-digit' });
};

// ── Icons (inline, stroke) ──────────────────────────────────────
const Icon = ({ d, fill }) => (
  <svg viewBox="0 0 24 24" width="22" height="22" fill={fill ? 'currentColor' : 'none'}
    stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round">
    {d}
  </svg>
);
const IconDumbbell = () => <Icon d={<><path d="M6.5 6.5l11 11" /><path d="M3 8l2-2 3 3-2 2zM21 16l-2 2-3-3 2-2z" /><path d="M2 9l2 2M22 15l-2-2" /></>} />;
const IconPlan = () => <Icon d={<><rect x="4" y="3" width="16" height="18" rx="2" /><path d="M8 8h8M8 12h8M8 16h5" /></>} />;
const IconBook = () => <Icon d={<><path d="M4 5a2 2 0 0 1 2-2h13v16H6a2 2 0 0 0-2 2z" /><path d="M4 19V5" /></>} />;
const IconChart = () => <Icon d={<><path d="M5 21V11M12 21V4M19 21v-6" /></>} />;
const IconPlus = () => <Icon d={<><path d="M12 5v14M5 12h14" /></>} />;
const IconCheck = () => <Icon d={<path d="M5 13l4 4L19 7" />} />;
const IconBack = () => <Icon d={<path d="M15 18l-6-6 6-6" />} />;
const IconTrash = () => <Icon d={<><path d="M4 7h16M9 7V5a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2M6 7l1 13h10l1-13" /></>} />;
const IconGrip = () => <Icon d={<><circle cx="9" cy="6" r="1" /><circle cx="9" cy="12" r="1" /><circle cx="9" cy="18" r="1" /><circle cx="15" cy="6" r="1" /><circle cx="15" cy="12" r="1" /><circle cx="15" cy="18" r="1" /></>} fill />;
const IconExport = () => <Icon d={<><path d="M12 3v12M8 11l4 4 4-4M5 21h14" /></>} />;

// ── Small bits ──────────────────────────────────────────────────
const num = (v) => (v === '' || v == null ? '' : v);

// Group consecutive plan items that share a superset letter into blocks.
function buildBlocks(items) {
  const blocks = [];
  let cur = null;
  for (const it of items) {
    if (it.gruppe) {
      if (cur && cur.gruppe === it.gruppe) cur.items.push(it);
      else { cur = { gruppe: it.gruppe, runden: it.runden, items: [it] }; blocks.push(cur); }
    } else { blocks.push({ gruppe: '', items: [it] }); cur = null; }
  }
  return blocks;
}

// ═══════════════════════════════════════════════════════════════
//  WORKOUT TAB
// ═══════════════════════════════════════════════════════════════
function ActivityCard({ item, last, onLog }) {
  const [val, setVal] = useState('');
  const [st, setSt] = useState(null);
  return (
    <div className="card">
      <div className="ex-head"><h3>{item.uebung}</h3></div>
      {last?.sets?.[0]?.notiz && <div className="last">zuletzt {fmtDate(last.datum)} · {last.sets[0].notiz}</div>}
      <div className="row-inline">
        <input className="in flex" value={val} onChange={(e) => setVal(e.target.value)} placeholder="Eintrag…" />
        <button className={`btn-log ${st === 'saved' ? 'done' : ''}`} disabled={!val || st === 'saving'} onClick={async () => {
          setSt('saving');
          try { await onLog({ uebung: item.uebung, satz: 1, notiz: val }); setSt('saved'); setVal(''); setTimeout(() => setSt(null), 1200); }
          catch { setSt('error'); }
        }}><IconCheck /></button>
      </div>
    </div>
  );
}

const fmtSet = (s) => (`${num(s.gewicht)}${s.gewicht ? '×' : ''}${num(s.wdh)}`.trim() || num(s.dauer) || s.notiz || '');

function ExerciseLogCard({ item, libEntry, last, loggedToday, onLog }) {
  const note = (item.notiz && item.notiz !== 'None' ? item.notiz : '')
    || (libEntry?.cues && libEntry.cues !== 'None' ? libEntry.cues : '');
  const seed = last?.sets?.[0];
  const [g, setG] = useState(seed?.gewicht || '');
  const [w, setW] = useState(seed?.wdh || item.zielWdh || '');
  const [du, setDu] = useState('');
  const [rpe, setRpe] = useState('');
  const [sessNote, setSessNote] = useState('');
  const [saving, setSaving] = useState(null);

  const target = [item.zielSaetze && `${item.zielSaetze}×`, item.zielWdh, item.tempo && `T ${item.tempo}`, item.pause && `P ${item.pause}`]
    .filter(Boolean).join('  ');
  const nextSatz = loggedToday.length + 1;

  const logSet = async () => {
    if (!g && !w && !du && !rpe && !sessNote) return;
    setSaving('saving');
    try {
      await onLog({ uebung: item.uebung, satz: nextSatz, gewicht: g, wdh: w, dauer: du, rpe, notiz: sessNote });
      setDu(''); setRpe(''); // keep weight/reps for the next straight set
      setSaving('saved'); setTimeout(() => setSaving(null), 1000);
    } catch { setSaving('error'); }
  };

  const lastTxt = last ? last.sets.map(fmtSet).filter(Boolean).join(', ') : '';

  return (
    <div className="card">
      <div className="ex-head"><h3>{item.uebung}</h3></div>
      {target && <div className="target">{target}</div>}
      {lastTxt && <div className="last">zuletzt {fmtDate(last.datum)} · {lastTxt}</div>}
      {note && <div className="cue-box">{note.split('\n').map((l, i) => <p key={i}>{l.replace(/^- /, '• ')}</p>)}</div>}

      {loggedToday.length > 0 && (
        <div className="logged">
          {loggedToday.map((s, i) => (
            <div className="logged-row" key={i}>
              <span className="set-n">{s.satz || i + 1}</span>
              <span className="logged-val">{fmtSet(s) || '–'}</span>
              {s.rpe && <span className="logged-rpe">RPE {s.rpe}</span>}
              {s.notiz && <span className="logged-note">{s.notiz}</span>}
            </div>
          ))}
        </div>
      )}

      <div className="set-head"><span>#</span><span>kg</span><span>Wdh</span><span>Zeit</span><span>RPE</span><span /></div>
      <div className="set-row">
        <span className="set-n">{nextSatz}</span>
        <input className="in" inputMode="decimal" value={g} onChange={(e) => setG(e.target.value)} placeholder="–" />
        <input className="in" inputMode="numeric" value={w} onChange={(e) => setW(e.target.value)} placeholder="–" />
        <input className="in" value={du} onChange={(e) => setDu(e.target.value)} placeholder="–" />
        <input className="in" inputMode="decimal" value={rpe} onChange={(e) => setRpe(e.target.value)} placeholder="–" />
        <button className={`btn-log ${saving === 'saved' ? 'done' : ''}`} disabled={saving === 'saving'} onClick={logSet}>
          {saving === 'error' ? '✕' : <IconCheck />}
        </button>
      </div>
      <input className="in note-in" value={sessNote} onChange={(e) => setSessNote(e.target.value)} placeholder="Notiz zu dieser Übung…" />
    </div>
  );
}

function WorkoutTab({ plans, libMap, log, selectedDate, setSelectedDate, onLog }) {
  const [planName, setPlanName] = useState(plans[0]?.name);
  useEffect(() => { if (!plans.find((p) => p.name === planName)) setPlanName(plans[0]?.name); }, [plans]);
  const plan = plans.find((p) => p.name === planName);

  return (
    <div className="tab-body">
      <input type="date" className="date-input" value={selectedDate} onChange={(e) => setSelectedDate(e.target.value)} />
      <div className="chips wrap">
        {plans.map((p) => (
          <button key={p.name} className={`chip ${p.name === planName ? 'on' : ''}`} onClick={() => setPlanName(p.name)}>{p.name}</button>
        ))}
      </div>
      {plan && buildBlocks(plan.items).map((b, bi) => {
        const cards = b.items.map((it, ii) => {
          const key = `${bi}-${ii}-${it.uebung}`;
          const common = { item: it, last: lastSession(log, it.uebung, selectedDate), onLog: (set) => onLog({ ...set, plan: plan.name, datum: selectedDate, exercise_id: it.exercise_id }) };
          if (plan.name === 'Aktivität') return <ActivityCard key={key} {...common} />;
          const loggedToday = log
            .filter((r) => r.uebung === it.uebung && r.datum === selectedDate && r.plan === plan.name)
            .sort((a, b2) => (parseInt(a.satz) || 0) - (parseInt(b2.satz) || 0));
          return <ExerciseLogCard key={key} libEntry={libMap.get(it.uebung)} loggedToday={loggedToday} {...common} />;
        });
        if (!b.gruppe) return cards;
        return (
          <div className="superset" key={bi}>
            <div className="superset-head"><span className="ss-chip">{b.gruppe}</span><span>Superset · {b.runden || '?'} Runden</span></div>
            {cards}
          </div>
        );
      })}
      {(!plan || plan.items.length === 0) && <p className="empty">Keine Übungen in diesem Plan.</p>}
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
//  PLANS TAB
// ═══════════════════════════════════════════════════════════════
function SortableExercise({ it, i, groupPos, upd, setGroup, setRunden, onRemove }) {
  const { attributes, listeners, setNodeRef, setActivatorNodeRef, transform, transition, isDragging } = useSortable({ id: it._id });
  const style = { transform: CSS.Transform.toString(transform), transition, opacity: isDragging ? 0.65 : 1, zIndex: isDragging ? 20 : undefined, position: 'relative' };
  return (
    <div className="card" ref={setNodeRef} style={style}>
      <div className="ex-head">
        <h3>{it.gruppe && <span className="ss-chip">{it.gruppe}{groupPos}</span>}{it.uebung}</h3>
        <div className="reorder">
          <button className="drag-handle" ref={setActivatorNodeRef} {...attributes} {...listeners} aria-label="Verschieben"><IconGrip /></button>
          <button onClick={onRemove} aria-label="Entfernen"><IconTrash /></button>
        </div>
      </div>
      <div className="grp-row">
        <select className="select-bare grow" value={it.gruppe || ''} onChange={(e) => setGroup(i, e.target.value)}>
          <option value="">Einzelübung</option>
          {['A', 'B', 'C', 'D', 'E', 'F'].map((g) => <option key={g} value={g}>Superset {g}</option>)}
        </select>
        {it.gruppe
          ? <label className="mini">Runden<input className="in" inputMode="numeric" value={it.runden || ''} onChange={(e) => setRunden(it.gruppe, e.target.value)} /></label>
          : <label className="mini">Sätze<input className="in" value={it.zielSaetze || ''} onChange={(e) => upd(i, 'zielSaetze', e.target.value)} /></label>}
      </div>
      <div className="target-grid three">
        <label>Wdh<input className="in" value={it.zielWdh || ''} onChange={(e) => upd(i, 'zielWdh', e.target.value)} /></label>
        <label>Tempo<input className="in" value={it.tempo || ''} onChange={(e) => upd(i, 'tempo', e.target.value)} /></label>
        <label>Pause<input className="in" value={it.pause || ''} onChange={(e) => upd(i, 'pause', e.target.value)} /></label>
      </div>
    </div>
  );
}

function PlanEditor({ plan, library, onSave, onCancel }) {
  const idRef = useRef(0);
  const newId = () => ++idRef.current;
  const [name, setName] = useState(plan?.name || '');
  const [items, setItems] = useState(() => (plan?.items || []).map((it) => ({ _id: newId(), ...it })));
  const [picking, setPicking] = useState(false);
  const [q, setQ] = useState('');
  const [saving, setSaving] = useState(false);
  const sensors = useSensors(useSensor(PointerSensor, { activationConstraint: { distance: 6 } }));

  const add = (lib) => { setItems((s) => [...s, { _id: newId(), exercise_id: lib.id, uebung: lib.uebung, gruppe: '', runden: '', zielSaetze: '', zielWdh: '', tempo: '', pause: '', notiz: '' }]); setPicking(false); setQ(''); };
  const upd = (i, k, v) => setItems((s) => s.map((x, j) => (j === i ? { ...x, [k]: v } : x)));
  const setGroup = (i, g) => setItems((s) => {
    const r = g ? (s.find((x) => x.gruppe === g)?.runden || '3') : '';
    return s.map((x, j) => (j === i ? { ...x, gruppe: g, runden: r } : x));
  });
  const setRunden = (g, v) => setItems((s) => s.map((x) => (x.gruppe === g ? { ...x, runden: v } : x)));
  const onDragEnd = ({ active, over }) => {
    if (!over || active.id === over.id) return;
    setItems((s) => arrayMove(s, s.findIndex((x) => x._id === active.id), s.findIndex((x) => x._id === over.id)));
  };

  if (picking) {
    const filtered = library.filter((l) => l.uebung.toLowerCase().includes(q.toLowerCase()));
    return (
      <div className="tab-body">
        <div className="editor-top"><button className="icon-btn" onClick={() => setPicking(false)}><IconBack /></button><h2>Übung wählen</h2></div>
        <input className="in flex search" autoFocus value={q} onChange={(e) => setQ(e.target.value)} placeholder="Suchen…" />
        {filtered.map((l) => (
          <button key={l.id} className="pick-row" onClick={() => add(l)}>
            <span>{l.uebung}</span><span className="muted">{l.kategorie}</span>
          </button>
        ))}
      </div>
    );
  }

  return (
    <div className="tab-body">
      <div className="editor-top">
        <button className="icon-btn" onClick={onCancel}><IconBack /></button>
        <input className="in flex" value={name} onChange={(e) => setName(e.target.value)} placeholder="Planname" />
      </div>
      <DndContext sensors={sensors} collisionDetection={closestCenter} onDragEnd={onDragEnd}>
        <SortableContext items={items.map((x) => x._id)} strategy={verticalListSortingStrategy}>
          {(() => {
            const groupCounts = {};
            return items.map((it, i) => {
              const groupPos = it.gruppe ? (groupCounts[it.gruppe] = (groupCounts[it.gruppe] || 0) + 1) : null;
              return (
                <SortableExercise key={it._id} it={it} i={i} groupPos={groupPos} upd={upd} setGroup={setGroup} setRunden={setRunden}
                  onRemove={() => setItems((s) => s.filter((_, j) => j !== i))} />
              );
            });
          })()}
        </SortableContext>
      </DndContext>
      <button className="add-set" onClick={() => setPicking(true)}><IconPlus /> Übung hinzufügen</button>
      <button className="btn-primary" disabled={!name || saving} onClick={async () => { setSaving(true); try { await onSave({ id: plan?.id, name, items }); } catch (e) { alert(e.message); setSaving(false); } }}>
        {saving ? 'Speichern…' : 'Plan speichern'}
      </button>
    </div>
  );
}

function PlansTab({ plans, library, onSave, onDelete }) {
  const [editing, setEditing] = useState(null); // {plan} | {new:true}
  if (editing) return <PlanEditor plan={editing.plan} library={library} onCancel={() => setEditing(null)} onSave={async (p) => { await onSave(p); setEditing(null); }} />;
  return (
    <div className="tab-body">
      {plans.map((p) => (
        <div className="card list-card" key={p.id}>
          <div onClick={() => setEditing({ plan: p })} className="list-main">
            <h3>{p.name}</h3><span className="muted">{p.items.length} Übungen</span>
          </div>
          <button className="icon-btn ghost" onClick={() => { if (confirm(`Plan "${p.name}" löschen?`)) onDelete(p.id); }}><IconTrash /></button>
        </div>
      ))}
      <button className="btn-primary" onClick={() => setEditing({ plan: null })}><IconPlus /> Neuer Plan</button>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
//  LIBRARY TAB
// ═══════════════════════════════════════════════════════════════
function LibraryEditor({ entry, onSave, onDelete, onCancel }) {
  const [uebung, setUebung] = useState(entry?.uebung || '');
  const [cues, setCues] = useState(entry?.cues && entry.cues !== 'None' ? entry.cues : '');
  const [saving, setSaving] = useState(false);
  return (
    <div className="tab-body">
      <div className="editor-top">
        <button className="icon-btn" onClick={onCancel}><IconBack /></button>
        <h2>{entry ? 'Übung bearbeiten' : 'Neue Übung'}</h2>
      </div>
      <input className="in flex" value={uebung} onChange={(e) => setUebung(e.target.value)} placeholder="Name der Übung" autoFocus />
      <textarea className="in flex area" value={cues} onChange={(e) => setCues(e.target.value)} placeholder="Cues / Hinweise — eine pro Zeile" />
      <button className="btn-primary" disabled={!uebung || saving} onClick={async () => { setSaving(true); try { await onSave({ uebung, cues }); } catch (e) { alert(e.message); setSaving(false); } }}>
        {saving ? 'Speichern…' : 'Speichern'}
      </button>
      {entry && <button className="btn-text-danger" onClick={() => { if (confirm('Übung löschen?')) onDelete(); }}>Übung löschen</button>}
    </div>
  );
}

function LibraryTab({ library, onAdd, onUpdate, onDelete }) {
  const [q, setQ] = useState('');
  const [editing, setEditing] = useState(null); // library row | { new: true }

  if (editing) {
    const entry = editing.new ? null : editing;
    return (
      <LibraryEditor
        entry={entry}
        onCancel={() => setEditing(null)}
        onSave={async (d) => { if (entry) await onUpdate(entry.id, d); else await onAdd(d); setEditing(null); }}
        onDelete={async () => { try { await onDelete(entry.id); setEditing(null); } catch (e) { alert(e.message); } }}
      />
    );
  }

  const filtered = library.filter((l) => l.uebung.toLowerCase().includes(q.toLowerCase()));
  return (
    <div className="tab-body">
      <input className="in flex search" value={q} onChange={(e) => setQ(e.target.value)} placeholder="Übung suchen…" />
      {filtered.map((l) => (
        <div className="card lib-card" key={l.id} onClick={() => setEditing(l)}>
          <div className="ex-head"><h3>{l.uebung}</h3><span className="muted">bearbeiten ›</span></div>
          {l.cues && l.cues !== 'None' && <div className="cue-box">{l.cues.split('\n').map((x, i) => <p key={i}>{x.replace(/^- /, '• ')}</p>)}</div>}
        </div>
      ))}
      {filtered.length === 0 && <p className="empty">Keine Übung gefunden.</p>}
      <button className="btn-primary" onClick={() => setEditing({ new: true })}><IconPlus /> Neue Übung</button>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
//  DASHBOARD TAB
// ═══════════════════════════════════════════════════════════════
const mondayOf = (iso) => {
  const d = new Date(iso + 'T00:00:00');
  d.setDate(d.getDate() - ((d.getDay() + 6) % 7));
  return d.toISOString().slice(0, 10);
};
const lastNWeeks = (n) => {
  const out = [];
  const base = new Date();
  base.setDate(base.getDate() - ((base.getDay() + 6) % 7));
  for (let i = n - 1; i >= 0; i--) {
    const w = new Date(base); w.setDate(base.getDate() - i * 7);
    out.push(w.toISOString().slice(0, 10));
  }
  return out;
};
const shortDay = (iso) => { const d = new Date(iso + 'T00:00:00'); return `${d.getDate()}.${d.getMonth() + 1}`; };

function BarChart({ data }) {
  const max = Math.max(1, ...data.map((d) => d.value));
  const W = 320, H = 148, base = H - 22, top = 16;
  const slot = W / data.length, bw = Math.min(22, slot * 0.5);
  return (
    <svg viewBox={`0 0 ${W} ${H}`} width="100%" style={{ display: 'block' }}>
      {data.map((d, i) => {
        const h = (d.value / max) * (base - top);
        const x = i * slot + (slot - bw) / 2;
        return (
          <g key={i}>
            {d.value > 0 && <text x={x + bw / 2} y={base - h - 5} textAnchor="middle" fontSize="9" fontFamily="Space Grotesk" fill="#111" fontWeight="600">{d.value}</text>}
            <rect x={x} y={base - h} width={bw} height={Math.max(h, d.value > 0 ? 2 : 0)} rx="3" fill={d.on ? '#d81413' : '#e6e6e6'} />
            <text x={x + bw / 2} y={H - 6} textAnchor="middle" fontSize="8" fill="#9b9b9b">{d.label}</text>
          </g>
        );
      })}
    </svg>
  );
}

function LineChart({ points }) {
  if (points.length < 2) return <p className="chart-empty">Noch zu wenig Daten — logge Sätze mit Gewicht, dann erscheint hier dein Verlauf.</p>;
  const W = 320, H = 150, padL = 8, padR = 30, top = 18, bot = 26;
  const ys = points.map((p) => p.y), min = Math.min(...ys), max = Math.max(...ys), span = max - min || 1;
  const x = (i) => padL + i * ((W - padL - padR) / (points.length - 1));
  const y = (v) => top + (1 - (v - min) / span) * (H - top - bot);
  const path = points.map((p, i) => `${i ? 'L' : 'M'}${x(i).toFixed(1)},${y(p.y).toFixed(1)}`).join(' ');
  const last = points[points.length - 1];
  return (
    <svg viewBox={`0 0 ${W} ${H}`} width="100%" style={{ display: 'block' }}>
      <path d={path} fill="none" stroke="#d81413" strokeWidth="2" strokeLinejoin="round" strokeLinecap="round" />
      {points.map((p, i) => <circle key={i} cx={x(i)} cy={y(p.y)} r={i === points.length - 1 ? 4 : 2.5} fill="#d81413" />)}
      <text x={x(points.length - 1) + 5} y={y(last.y) + 3} fontSize="11" fontFamily="Space Grotesk" fontWeight="600" fill="#111">{last.y}</text>
      <text x={padL} y={H - 8} fontSize="8" fill="#9b9b9b">{points[0].label}</text>
      <text x={W - padR} y={H - 8} textAnchor="end" fontSize="8" fill="#9b9b9b">{last.label}</text>
    </svg>
  );
}

function DashboardTab({ log }) {
  const strength = useMemo(() => log.filter((r) => r.plan !== 'Aktivität'), [log]);
  const weeks = useMemo(() => lastNWeeks(10), []);
  const thisWeek = weeks[weeks.length - 1];
  const setsByWeek = useMemo(() => {
    const m = new Map();
    for (const r of strength) { if (!r.datum) continue; const w = mondayOf(r.datum); m.set(w, (m.get(w) || 0) + 1); }
    return m;
  }, [strength]);
  const barData = weeks.map((w) => ({ label: shortDay(w), value: setsByWeek.get(w) || 0, on: w === thisWeek }));

  const sessionDates = useMemo(() => [...new Set(strength.map((r) => r.datum).filter(Boolean))], [strength]);
  const sessionsThisWeek = sessionDates.filter((d) => mondayOf(d) === thisWeek).length;
  const sets4w = weeks.slice(-4).reduce((s, w) => s + (setsByWeek.get(w) || 0), 0);
  const streak = useMemo(() => {
    let n = 0;
    for (let i = weeks.length - 1; i >= 0; i--) { if (setsByWeek.get(weeks[i])) n++; else if (i !== weeks.length - 1) break; else continue; }
    return n;
  }, [weeks, setsByWeek]);

  // Per-exercise weight progression
  const exWithWeight = useMemo(() => [...new Set(strength.filter((r) => parseFloat(r.gewicht) > 0).map((r) => r.uebung))].sort(), [strength]);
  const [exName, setExName] = useState('');
  useEffect(() => { if (!exName && exWithWeight.length) setExName(exWithWeight[0]); }, [exWithWeight]);
  const progression = useMemo(() => {
    const byDate = new Map();
    for (const r of strength) {
      if (r.uebung !== exName) continue;
      const g = parseFloat(r.gewicht); if (!(g > 0)) continue;
      byDate.set(r.datum, Math.max(byDate.get(r.datum) || 0, g));
    }
    return [...byDate.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([d, g]) => ({ label: shortDay(d), y: g }));
  }, [strength, exName]);

  return (
    <div className="tab-body">
      <div className="stat-grid">
        <div className="stat"><div className="v accent">{sessionsThisWeek}</div><div className="l">Sessions / Woche</div></div>
        <div className="stat"><div className="v">{sets4w}</div><div className="l">Sätze / 4 Wo.</div></div>
        <div className="stat"><div className="v">{streak}</div><div className="l">Wochen-Serie</div></div>
      </div>

      <div className="panel">
        <div className="panel-head"><span className="eyebrow">Trainingsvolumen</span><span className="muted">Sätze pro Woche</span></div>
        <BarChart data={barData} />
      </div>

      <div className="panel">
        <div className="panel-head">
          <span className="eyebrow">Fortschritt</span>
          {exWithWeight.length > 0 && (
            <select className="select-bare" value={exName} onChange={(e) => setExName(e.target.value)}>
              {exWithWeight.map((n) => <option key={n} value={n}>{n}</option>)}
            </select>
          )}
        </div>
        {exWithWeight.length === 0
          ? <p className="chart-empty">Sobald du Sätze mit Gewicht loggst, zeigt sich hier deine Gewichtsentwicklung pro Übung.</p>
          : <LineChart points={progression} />}
        {exWithWeight.length > 0 && <p className="legend">Höchstes Gewicht (kg) pro Session</p>}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
//  AUTH SCREEN (Supabase email + password)
// ═══════════════════════════════════════════════════════════════
function AuthScreen() {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [state, setState] = useState(null); // 'sending' | 'error'
  const [msg, setMsg] = useState('');

  const login = async () => {
    if (!email || !password) return;
    setState('sending'); setMsg('');
    const { error } = await supabase.auth.signInWithPassword({ email, password });
    if (error) { setState('error'); setMsg(error.message); }
    // success: onAuthStateChange picks up the session automatically
  };

  return (
    <div className="center">
      <div className="glass auth">
        <div className="mark">T</div>
        <h1>Training</h1>
        <p>Anmelden mit E-Mail und Passwort.</p>
        <input className="in flex" type="email" inputMode="email" autoComplete="email" value={email}
          onChange={(e) => setEmail(e.target.value)} placeholder="deine@email.de" />
        <input className="in flex" type="password" autoComplete="current-password" value={password}
          onChange={(e) => setPassword(e.target.value)} placeholder="Passwort"
          onKeyDown={(e) => e.key === 'Enter' && login()} />
        <button className="btn-primary" disabled={!email || !password || state === 'sending'} onClick={login}>
          {state === 'sending' ? 'Anmelden…' : 'Anmelden'}
        </button>
        {state === 'error' && <p className="err">{msg}</p>}
      </div>
    </div>
  );
}

// ═══════════════════════════════════════════════════════════════
//  APP SHELL
// ═══════════════════════════════════════════════════════════════
export default function App() {
  const [session, setSession] = useState(null);
  const [authReady, setAuthReady] = useState(false);

  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);
  const [lib, setLib] = useState([]);
  const [plans, setPlans] = useState([]);
  const [log, setLog] = useState([]);
  const [tab, setTab] = useState('dashboard');
  const [selectedDate, setSelectedDate] = useState(isoDate(new Date()));

  // ── Auth session ──
  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => { setSession(data.session); setAuthReady(true); });
    const { data: sub } = supabase.auth.onAuthStateChange((_e, sess) => setSession(sess));
    return () => sub.subscription.unsubscribe();
  }, []);

  const loadAll = useCallback(async () => {
    setLoading(true); setError(null);
    try {
      const [l, p, g] = await Promise.all([getExercises(), getPlans(), getLog()]);
      setLib(l); setPlans(p); setLog(g);
    } catch (e) { setError(e.message); } finally { setLoading(false); }
  }, []);

  const uid = session?.user?.id;
  useEffect(() => { if (uid) loadAll(); }, [uid, loadAll]);

  const libMap = useMemo(() => new Map(lib.map((l) => [l.uebung, l])), [lib]);
  const workoutPlans = useMemo(() => {
    const base = [...plans];
    const acts = [...new Set(log.filter((r) => r.plan === 'Aktivität').map((r) => r.uebung))];
    if (acts.length) base.push({ name: 'Aktivität', items: acts.map((a) => ({ uebung: a })) });
    return base;
  }, [plans, log]);
  // ── Mutations ──
  const handleLog = async (row) => {
    const saved = await addLogRow(row);
    setLog((prev) => [...prev, saved]); // optimistic; avoids a full refetch
  };
  const handleSavePlan = async (p) => { await savePlan(p); setPlans(await getPlans()); };
  const handleDeletePlan = async (id) => { await deletePlan(id); setPlans(await getPlans()); };
  const handleAddLib = async (e) => {
    const saved = await addExercise(e);
    setLib((prev) => [...prev, saved].sort((a, b) => a.uebung.localeCompare(b.uebung)));
  };
  const handleUpdateLib = async (id, d) => {
    const saved = await updateExercise(id, d);
    setLib((prev) => prev.map((x) => (x.id === id ? saved : x)).sort((a, b) => a.uebung.localeCompare(b.uebung)));
    if ('uebung' in d) setPlans(await getPlans()); // name change must propagate to cached plan items
  };
  const handleDeleteLib = async (id) => {
    await deleteExercise(id);
    setLib((prev) => prev.filter((x) => x.id !== id));
  };
  const handleExport = async () => {
    try { await exportToExcel({ exercises: lib, plans, log }); }
    catch (e) { alert('Export fehlgeschlagen: ' + e.message); }
  };

  // ── Gate screens ──
  if (!authReady) return <div className="center"><div className="spinner" /></div>;
  if (!session) return <AuthScreen />;
  if (loading && !lib.length) return <div className="center"><div className="spinner" /></div>;
  if (error && !lib.length) return (
    <div className="center"><div className="glass auth"><h1>Fehler</h1><p className="err">{error}</p><button className="btn-primary" onClick={loadAll}>Erneut versuchen</button></div></div>
  );

  const titles = { workout: 'Workout', dashboard: 'Übersicht', plans: 'Pläne', library: 'Bibliothek' };
  return (
    <div className="app">
      <header className="topbar glass">
        <h1>{titles[tab]}</h1>
        <div className="topbar-actions">
          <button className="icon-btn ghost" onClick={handleExport} title="Nach Excel exportieren"><IconExport /></button>
          <button className="icon-btn ghost" onClick={() => supabase.auth.signOut()} title="Abmelden">⏻</button>
        </div>
      </header>

      <main className="main">
        {tab === 'workout' && <WorkoutTab plans={workoutPlans} libMap={libMap} log={log} selectedDate={selectedDate} setSelectedDate={setSelectedDate} onLog={handleLog} />}
        {tab === 'dashboard' && <DashboardTab log={log} />}
        {tab === 'plans' && <PlansTab plans={plans} library={lib} onSave={handleSavePlan} onDelete={handleDeletePlan} />}
        {tab === 'library' && <LibraryTab library={lib} onAdd={handleAddLib} onUpdate={handleUpdateLib} onDelete={handleDeleteLib} />}
      </main>

      <nav className="tabbar glass">
        <button className={tab === 'dashboard' ? 'on' : ''} onClick={() => setTab('dashboard')}><IconChart /><span>Übersicht</span></button>
        <button className={tab === 'workout' ? 'on' : ''} onClick={() => setTab('workout')}><IconDumbbell /><span>Workout</span></button>
        <button className={tab === 'plans' ? 'on' : ''} onClick={() => setTab('plans')}><IconPlan /><span>Pläne</span></button>
        <button className={tab === 'library' ? 'on' : ''} onClick={() => setTab('library')}><IconBook /><span>Bibliothek</span></button>
      </nav>
      {loading && lib.length > 0 && <div className="reload-bar" />}
    </div>
  );
}
