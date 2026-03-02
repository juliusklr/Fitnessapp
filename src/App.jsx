import { useState, useEffect, useCallback } from 'react';
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import { loginRequest } from './authConfig';
import {
  loadTrainingData,
  getValuesForDate,
  getLastEntriesPerExercise,
  updateRange,
  colLetter,
} from './graphService';

// ── Date helpers ────────────────────────────────────────────────
const isoDate = (d) => {
  const dt = new Date(d);
  return `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}-${String(dt.getDate()).padStart(2, '0')}`;
};

const formatDate = (d) => {
  const dt = new Date(d);
  return dt.toLocaleDateString('de-DE', { weekday: 'short', day: '2-digit', month: '2-digit' });
};

const sameDay = (a, b) => {
  const da = new Date(a), db = new Date(b);
  return da.getFullYear() === db.getFullYear() && da.getMonth() === db.getMonth() && da.getDate() === db.getDate();
};

// ── Save Indicator ──────────────────────────────────────────────
function SaveIndicator({ state }) {
  if (!state) return null;
  const map = {
    saving: { text: '…', cls: 'saving' },
    saved: { text: '✓', cls: 'saved' },
    error: { text: '✕', cls: 'error' },
  };
  const s = map[state] || {};
  return <span className={`save-ind ${s.cls}`}>{s.text}</span>;
}

// ── Exercise Card ───────────────────────────────────────────────
function ExerciseCard({ exercise, value, dateExists, saving, onSave, lastEntry }) {
  const [noteOpen, setNoteOpen] = useState(false);
  const [val, setVal] = useState(value);

  useEffect(() => { setVal(value); }, [value]);

  const hasNote = exercise.note && exercise.note !== 'None' && exercise.note !== '';
  const params = [
    { label: 'S', value: exercise.sets },
    { label: 'W', value: exercise.reps },
    { label: 'T', value: exercise.timing },
    { label: 'P', value: exercise.pause },
  ].filter((p) => p.value && p.value !== 'None' && p.value !== '-' && p.value !== '');

  return (
    <div className="ex-card">
      <div className="ex-header">
        <h3 className="ex-name">{exercise.name}</h3>
        {hasNote && (
          <button className="note-toggle" onClick={() => setNoteOpen(!noteOpen)}>
            {noteOpen ? '✕' : 'ℹ'}
          </button>
        )}
      </div>

      {noteOpen && hasNote && (
        <div className="note-box">
          {exercise.note.split('\n').map((line, i) => (
            <p key={i} className="note-line">{line.replace(/^- /, '• ')}</p>
          ))}
        </div>
      )}

      <div className="param-row">
        {params.map((p) => (
          <div key={p.label} className="param-chip">
            <span className="param-label">{p.label}</span>
            <span className="param-value">{p.value}</span>
          </div>
        ))}
      </div>

      {lastEntry && (
        <div className="last-entry">
          <span className="last-entry-icon">↩</span>
          <span className="last-entry-date">{formatDate(lastEntry.date)}</span>
          <span className="last-entry-sep">·</span>
          <span className="last-entry-value">{lastEntry.value}</span>
        </div>
      )}

      <div className="ex-input-wrap">
        <input
          className="ex-input"
          value={val}
          onChange={(e) => setVal(e.target.value)}
          placeholder="Ergebnis eingeben..."
          onBlur={() => {
            if (val !== value) onSave(val);
          }}
          disabled={!dateExists}
        />
        <SaveIndicator state={saving} />
      </div>
    </div>
  );
}

// ── Main App ────────────────────────────────────────────────────
export default function App() {
  const { instance, accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  const [loading, setLoading] = useState(false);
  const [error, setError] = useState(null);

  // Data
  const [dates, setDates] = useState([]);
  const [activityLabels, setActivityLabels] = useState([]);
  const [exercises, setExercises] = useState([]);
  const [plans, setPlans] = useState([]);
  const [rawRows, setRawRows] = useState(null);

  // UI state
  const [selectedDate, setSelectedDate] = useState(isoDate(new Date()));
  const [selectedPlan, setSelectedPlan] = useState(null);
  const [actValues, setActValues] = useState({});
  const [exValues, setExValues] = useState({});
  const [lastEntries, setLastEntries] = useState({});
  const [saving, setSaving] = useState({});
  const [actCollapsed, setActCollapsed] = useState(false);

  // ── Get token ─────────────────────────────────────────────
  const getToken = useCallback(async () => {
    const account = accounts[0];
    if (!account) throw new Error('Nicht eingeloggt');
    try {
      const resp = await instance.acquireTokenSilent({ ...loginRequest, account });
      return resp.accessToken;
    } catch {
      const resp = await instance.acquireTokenPopup({ ...loginRequest, account });
      return resp.accessToken;
    }
  }, [instance, accounts]);

  // ── Load all data ─────────────────────────────────────────
  const loadData = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const token = await getToken();
      const data = await loadTrainingData(token);
      setDates(data.dates);
      setActivityLabels(data.activityLabels);
      setExercises(data.exercises);
      setPlans(data.plans);
      setRawRows(data.rawRows);
      if (!selectedPlan && data.plans.length > 0) {
        setSelectedPlan(data.plans[0]);
      }
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }, [getToken, selectedPlan]);

  // ── Update values when date or rawRows change ─────────────
  useEffect(() => {
    if (!rawRows || dates.length === 0) return;
    const dateCol = dates.find((d) => sameDay(d.date, selectedDate));
    if (!dateCol) {
      setActValues({});
      setExValues({});
    } else {
      const { activityValues, exerciseValues } = getValuesForDate(rawRows, dateCol.col);
      setActValues(activityValues);
      setExValues(exerciseValues);
    }
    setLastEntries(getLastEntriesPerExercise(rawRows, dates, exercises, selectedDate));
  }, [rawRows, dates, selectedDate, exercises]);

  // ── Auto-load on auth ─────────────────────────────────────
  useEffect(() => {
    if (isAuthenticated && !rawRows && !loading) {
      loadData();
    }
  }, [isAuthenticated]);

  // ── Login ─────────────────────────────────────────────────
  const handleLogin = async () => {
    try {
      await instance.loginPopup(loginRequest);
    } catch (e) {
      setError(e.message);
    }
  };

  // ── Logout ────────────────────────────────────────────────
  const handleLogout = () => {
    instance.logoutPopup();
  };

  // ── Save a single cell ────────────────────────────────────
  const saveCell = async (row, value, key) => {
    const dateCol = dates.find((d) => sameDay(d.date, selectedDate));
    if (!dateCol) return;

    const cellAddr = `${colLetter(dateCol.col)}${row}`;
    setSaving((s) => ({ ...s, [key]: 'saving' }));
    try {
      const token = await getToken();
      await updateRange(token, cellAddr, [[value]]);
      setSaving((s) => ({ ...s, [key]: 'saved' }));
      setTimeout(() => setSaving((s) => ({ ...s, [key]: null })), 1500);
    } catch (e) {
      setSaving((s) => ({ ...s, [key]: 'error' }));
      console.error('Save error:', e);
    }
  };

  // ── Date change handler ───────────────────────────────────
  const handleDateChange = (newDate) => {
    setSelectedDate(newDate);
  };

  // ── Render: Not authenticated ─────────────────────────────
  if (!isAuthenticated) {
    return (
      <div className="screen-center">
        <div className="auth-card">
          <div className="logo">🏋️</div>
          <h1>Training Tracker</h1>
          <p className="subtitle">
            Verbinde dich mit deinem Microsoft-Account, um deine Trainingsdaten direkt in Excel zu tracken.
          </p>
          <button className="btn-primary" onClick={handleLogin}>
            Mit Microsoft anmelden
          </button>
          {error && <p className="error-text">{error}</p>}
        </div>
      </div>
    );
  }

  // ── Render: Loading ───────────────────────────────────────
  if (loading && !rawRows) {
    return (
      <div className="screen-center">
        <div className="loading-wrap">
          <div className="spinner" />
          <p className="loading-text">Lade Trainingsdaten...</p>
        </div>
      </div>
    );
  }

  // ── Render: Error ─────────────────────────────────────────
  if (error && !rawRows) {
    return (
      <div className="screen-center">
        <div className="auth-card">
          <h1 style={{ color: 'var(--error)', marginBottom: 12, fontSize: 20 }}>Fehler</h1>
          <p className="error-text">{error}</p>
          <button className="btn-primary" onClick={loadData} style={{ marginTop: 16 }}>
            Erneut versuchen
          </button>
        </div>
      </div>
    );
  }

  // ── Render: Main ──────────────────────────────────────────
  const dateExists = dates.some((d) => sameDay(d.date, selectedDate));
  const filteredExercises = exercises.filter((e) => e.plan === selectedPlan);

  return (
    <div className="app">
      {/* Header */}
      <header className="header">
        <div className="header-top">
          <h1 className="header-title">Training</h1>
          <div className="header-actions">
            <button className="icon-btn" onClick={loadData} title="Refresh">
              ↻
            </button>
            <button className="icon-btn" onClick={handleLogout} title="Logout">
              ⏻
            </button>
          </div>
        </div>

        {/* Date picker */}
        <div className="date-row">
          <input
            type="date"
            className="date-input"
            value={selectedDate}
            onChange={(e) => handleDateChange(e.target.value)}
          />
          {!dateExists && <span className="date-warning">⚠ Nicht in Excel</span>}
        </div>

        {/* Date chips */}
        <div className="date-chips">
          {dates.slice(-10).map((d) => (
            <button
              key={d.col}
              className={`date-chip ${sameDay(d.date, selectedDate) ? 'active' : ''}`}
              onClick={() => handleDateChange(isoDate(d.date))}
            >
              {formatDate(d.date)}
            </button>
          ))}
        </div>
      </header>

      <main className="main">
        {/* Daily Activities */}
        <section className="section">
          <button className="section-header" onClick={() => setActCollapsed(!actCollapsed)}>
            <span className="section-title">Tägliche Aktivitäten</span>
            <span className="chevron">{actCollapsed ? '▸' : '▾'}</span>
          </button>

          {!actCollapsed && (
            <div className="act-grid">
              {activityLabels.map((act) => (
                <div key={act.row} className="act-row">
                  <label className="act-label">{act.label}</label>
                  <div className="act-input-wrap">
                    <input
                      className="act-input"
                      defaultValue={actValues[act.row] || ''}
                      key={`${act.row}-${selectedDate}`}
                      placeholder="—"
                      onBlur={(e) => {
                        const current = actValues[act.row] || '';
                        if (e.target.value !== current) {
                          saveCell(act.row, e.target.value, `act-${act.row}`);
                          setActValues((v) => ({ ...v, [act.row]: e.target.value }));
                        }
                      }}
                      disabled={!dateExists}
                    />
                    <SaveIndicator state={saving[`act-${act.row}`]} />
                  </div>
                </div>
              ))}
            </div>
          )}
        </section>

        {/* Training Plan */}
        <section className="section">
          <div className="section-header" style={{ cursor: 'default' }}>
            <span className="section-title">Trainingsplan</span>
          </div>

          <div className="plan-tabs">
            {plans.map((p) => (
              <button
                key={p}
                className={`plan-tab ${p === selectedPlan ? 'active' : ''}`}
                onClick={() => setSelectedPlan(p)}
              >
                {p}
              </button>
            ))}
          </div>

          <div className="exercise-list">
            {filteredExercises.map((ex) => (
              <ExerciseCard
                key={`${ex.row}-${ex.name}`}
                exercise={ex}
                value={exValues[ex.row] || ''}
                dateExists={dateExists}
                saving={saving[`ex-${ex.row}`]}
                onSave={(val) => {
                  saveCell(ex.row, val, `ex-${ex.row}`);
                  setExValues((v) => ({ ...v, [ex.row]: val }));
                }}
                lastEntry={lastEntries[ex.row]}
              />
            ))}
            {filteredExercises.length === 0 && (
              <p className="empty-text">Keine Übungen für diesen Plan.</p>
            )}
          </div>
        </section>
      </main>

      {/* Reload indicator */}
      {loading && rawRows && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          height: 3,
          background: 'var(--accent)',
          zIndex: 100,
          animation: 'pulse 1s ease-in-out infinite',
        }} />
      )}
    </div>
  );
}
