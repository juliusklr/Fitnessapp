const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const FILE_NAME = 'Masterplan.xlsx';
const SHEET_NAME = 'Masterplan';

let cachedItemPath = null;

async function graphFetch(token, url, options = {}) {
  const res = await fetch(`${GRAPH_BASE}${url}`, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
      ...options.headers,
    },
  });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Graph API ${res.status}: ${text}`);
  }
  return res.status === 204 ? null : res.json();
}

async function findFile(token) {
  if (cachedItemPath) return cachedItemPath;
  
  // Search for the file by name
  const result = await graphFetch(
    token,
    `/me/drive/root/search(q='${FILE_NAME}')?$select=id,name,parentReference`
  );
  
  const file = result.value?.find(f => f.name === FILE_NAME);
  if (!file) throw new Error(`Datei "${FILE_NAME}" nicht gefunden auf OneDrive`);
  
  cachedItemPath = `/me/drive/items/${file.id}`;
  return cachedItemPath;
}

// ── Read a range ────────────────────────────────────────────────
export async function getRange(token, address) {
  const itemPath = await findFile(token);
  return graphFetch(
    token,
    `${itemPath}/workbook/worksheets('${SHEET_NAME}')/range(address='${address}')`
  );
}

// ── Write a range ───────────────────────────────────────────────
export async function updateRange(token, address, values) {
  const itemPath = await findFile(token);
  return graphFetch(
    token,
    `${itemPath}/workbook/worksheets('${SHEET_NAME}')/range(address='${address}')`,
    {
      method: 'PATCH',
      body: JSON.stringify({ values }),
    }
  );
}

// ── Helper: column number to letter(s) ──────────────────────────
export function colLetter(n) {
  let s = '';
  while (n > 0) {
    n--;
    s = String.fromCharCode(65 + (n % 26)) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

// ── Helper: detect Excel date serial / ISO string ───────────────
export function isExcelDate(v) {
  if (v === null || v === undefined || v === '' || v === false) return false;
  // Excel serial number (dates range roughly 1900-2100 → 1 to 73000)
  if (typeof v === 'number' && v > 25000 && v < 73000) return true;
  // ISO date string from Graph API
  if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}(T|$)/.test(v)) {
    // But not if it looks like a normal value that happens to match (e.g. "2025-06-04" as reps)
    return true;
  }
  return false;
}

export function excelDateToDisplay(v) {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number' && v > 25000 && v < 73000) {
    const d = new Date((v - 25569) * 86400000);
    return `${d.getUTCDate()}-${d.getUTCMonth() + 1}`;
  }
  if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}/.test(v)) {
    const parts = v.split(/[-T]/);
    return `${parseInt(parts[2])}-${parseInt(parts[1])}`;
  }
  return String(v ?? '');
}

// ── Load all training data ──────────────────────────────────────
export async function loadTrainingData(token) {
  // Load a large range to cover all possible data
  // Rows 3-150 (enough for exercises), Cols A through CZ (104 cols, enough for dates)
  const range = await getRange(token, 'A3:CZ150');
  const rows = range.values;

  // Row 0 in array = Row 3 in sheet (dates), from col I (index 8)
  const dates = [];
  for (let c = 8; c < rows[0].length; c++) {
    const v = rows[0][c];
    if (v !== null && v !== undefined && v !== '' && v !== false) {
      let dateObj = null;
      
      if (typeof v === 'number') {
        // Excel serial number: days since 1899-12-30
        // e.g. 46083 = 2026-03-02
        const utcDays = v - 25569; // Convert to Unix days
        dateObj = new Date(utcDays * 86400000);
      } else if (typeof v === 'string') {
        // Try ISO format first (2026-03-02T00:00:00.000)
        const parsed = new Date(v);
        if (!isNaN(parsed.getTime()) && parsed.getFullYear() > 2000) {
          dateObj = parsed;
        }
      }
      
      // Debug: log first few date values
      if (c < 12) {
        console.log(`Date col ${c}: raw="${v}" (${typeof v}) → parsed=${dateObj?.toISOString()}`);
      }
      
      if (dateObj && !isNaN(dateObj.getTime()) && dateObj.getFullYear() > 2000) {
        dates.push({ col: c + 1, date: dateObj });
      }
    }
  }
  dates.sort((a, b) => a.date - b.date);
  console.log(`Found ${dates.length} dates, range: ${dates[0]?.date?.toISOString()} - ${dates[dates.length-1]?.date?.toISOString()}`);

  // Rows 1-11 in array = Rows 4-14 in sheet (daily activities)
  const activityLabels = [];
  for (let r = 1; r <= 11; r++) {
    const label = rows[r]?.[7] || '';
    if (String(label).trim()) {
      activityLabels.push({ row: r + 3, label: String(label) });
    }
  }

  // Rows 13+ in array = Rows 16+ in sheet (exercises)
  const exercises = [];
  const planSet = new Set();
  for (let r = 13; r < rows.length; r++) {
    const name = rows[r]?.[1];
    const plan = rows[r]?.[7];
    if (!name || !plan) continue;
    // Skip if name is empty string
    if (String(name).trim() === '') continue;
    
    planSet.add(String(plan));

    const rawReps = rows[r]?.[4];
    const reps = isExcelDate(rawReps) ? excelDateToDisplay(rawReps) : String(rawReps ?? '');

    exercises.push({
      row: r + 3, // actual sheet row
      name: String(name),
      note: String(rows[r]?.[2] ?? ''),
      sets: String(rows[r]?.[3] ?? ''),
      reps,
      timing: String(rows[r]?.[5] ?? ''),
      pause: String(rows[r]?.[6] ?? ''),
      plan: String(plan),
    });
  }

  const plans = [...planSet].sort();
  console.log(`Found ${exercises.length} exercises, plans:`, plans);
  console.log(`Total rows in response: ${rows.length}, total cols: ${rows[0]?.length}`);

  return { dates, activityLabels, exercises, plans, rawRows: rows };
}

// ── Find last entry per exercise (excluding the selected date) ───
export function getLastEntriesPerExercise(rawRows, dates, exercises, selectedDate) {
  const result = {};

  // Sort dates descending, skip the currently selected date
  const sortedDates = [...dates]
    .filter(d => {
      const dt = d.date;
      const iso = `${dt.getFullYear()}-${String(dt.getMonth() + 1).padStart(2, '0')}-${String(dt.getDate()).padStart(2, '0')}`;
      return iso !== selectedDate;
    })
    .sort((a, b) => b.date - a.date);

  for (const ex of exercises) {
    const arrayRow = ex.row - 3;
    for (const d of sortedDates) {
      const colIdx = d.col - 1;
      const val = rawRows[arrayRow]?.[colIdx];
      if (val !== null && val !== undefined && val !== '' && val !== false) {
        const display = isExcelDate(val) ? excelDateToDisplay(val) : String(val);
        result[ex.row] = { date: d.date, value: display };
        break;
      }
    }
  }
  return result;
}

// ── Load values for a specific date column ──────────────────────
export function getValuesForDate(rawRows, dateCol) {
  const colIdx = dateCol - 1; // 0-indexed

  // Activities: rows 1-11 in array
  const activityValues = {};
  for (let r = 1; r <= 11; r++) {
    const row = r + 3;
    const val = rawRows[r]?.[colIdx];
    const display = isExcelDate(val) ? excelDateToDisplay(val) : String(val ?? '');
    activityValues[row] = display;
  }

  // Exercises: rows 13+ in array
  const exerciseValues = {};
  for (let r = 13; r < rawRows.length; r++) {
    const row = r + 3;
    const val = rawRows[r]?.[colIdx];
    const display = isExcelDate(val) ? excelDateToDisplay(val) : String(val ?? '');
    exerciseValues[row] = display;
  }

  return { activityValues, exerciseValues };
}
