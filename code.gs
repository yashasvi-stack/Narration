const SHEET_ID = '16fL0ksNgjmDHXchBnV3Axkrm3OvleT5stC70B9aC8B4';
const HOUSEHOLD_TAB = 'Household Datasets';
const INDUSTRIAL_TAB = 'Industrial Datasets';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Narration - Atlas Video Performance')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Re-reads both tabs on every call. Columns are found by heading name, never by letter.
function getData() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const rows = []
    .concat(readTab_(ss, HOUSEHOLD_TAB, 'Household', 'File name', 'Video Duration', false))
    .concat(readTab_(ss, INDUSTRIAL_TAB, 'Industrial', 'Session', 'Duration (min)', true));
  return { rows: rows, sheetUrl: ss.getUrl() };
}

function readTab_(ss, tab, ds, idH, durH, isMin) {
  const sh = ss.getSheetByName(tab);
  if (!sh) throw new Error('Tab not found: ' + tab);
  const rng = sh.getDataRange();
  const v = rng.getDisplayValues();
  const raw = rng.getValues(); // exact numbers (display text can be rounded)
  if (v.length < 2) return [];

  const h = {};
  v[0].forEach(function (x, i) {
    const k = String(x).trim().toLowerCase().replace(/\s+/g, ' ');
    if (k && !(k in h)) h[k] = i;
  });
  const col = function (n) { const k = n.toLowerCase(); return k in h ? h[k] : -1; };
  const c = {
    id: col(idH), dur: col(durH), cat: col('Category'),
    poc: col('Assigned To'), dt: col('Assigned Date'),
    st: col('Start time'), en: col('End time')
  };
  const need = { id: idH, dur: durH, poc: 'Assigned To', dt: 'Assigned Date' };
  Object.keys(need).forEach(function (k) {
    if (c[k] < 0) throw new Error('Heading "' + need[k] + '" not found in tab "' + tab + '"');
  });

  const seen = {}, out = [];
  for (let r = 1; r < v.length; r++) {
    const g = function (i) { return i < 0 ? '' : String(v[r][i]).trim(); };
    const id = g(c.id), poc = g(c.poc), date = isoDate_(g(c.dt));
    if (!id) continue; // rows without POC/date are kept (for sheet totals); the tables skip them
    const done = g(c.en) !== '';
    const rec = {
      ds: ds, id: id, poc: poc, date: date,
      cat: g(c.cat) || (ds === 'Household' ? 'Household Data' : 'Uncategorised'),
      dur: isMin ? minToSecExact_(raw[r][c.dur], g(c.dur)) : hmsToSec_(g(c.dur)),
      taken: done ? span_(g(c.st), g(c.en)) : 0,
      done: done
    };
    const k = id.toLowerCase();
    if (k in seen) {
      // duplicate ID: count once, prefer the completed row, then the assigned row
      const score = function (x) { return (x.done ? 2 : 0) + (x.poc && x.date ? 1 : 0); };
      if (score(rec) > score(out[seen[k]])) out[seen[k]] = rec;
    } else {
      seen[k] = out.length;
      out.push(rec);
    }
  }
  return out;
}

function isoDate_(s) {
  let m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + ('0' + m[1]).slice(-2) + '-' + ('0' + m[2]).slice(-2);
  m = s.match(/(\d{4})-(\d{2})-(\d{2})/);
  return m ? m[0] : '';
}

function hmsToSec_(s) {
  if (!s) return 0;
  const p = s.split(':').map(Number);
  if (p.some(isNaN)) return 0;
  if (p.length === 3) return p[0] * 3600 + p[1] * 60 + p[2];
  if (p.length === 2) return p[0] * 60 + p[1];
  return 0;
}

// Duration (min) -> seconds: minutes / 1440 = days, x 86400 = seconds (same as minutes x 60), using the exact cell value
function minToSecExact_(rawVal, shown) {
  if (shown && shown.indexOf(':') > -1) return hmsToSec_(shown);
  if (typeof rawVal === 'number') return Math.round(rawVal / 1440 * 86400);
  return minToSec_(shown);
}

function minToSec_(s) {
  if (!s) return 0;
  if (s.indexOf(':') > -1) return hmsToSec_(s);
  const n = parseFloat(s.replace(/,/g, ''));
  return isNaN(n) ? 0 : Math.round(n * 60);
}

function timeOfDay_(s) {
  const m = s.match(/(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?/i);
  if (!m) return null;
  let h = +m[1];
  if (m[4]) {
    const p = m[4].toUpperCase();
    if (p === 'PM' && h < 12) h += 12;
    if (p === 'AM' && h === 12) h = 0;
  }
  return h * 3600 + (+m[2]) * 60 + (+m[3] || 0);
}

// Time taken = End time - Start time (handles date+time values and overnight wrap)
function span_(st, en) {
  const a = timeOfDay_(st), b = timeOfDay_(en);
  if (a === null || b === null) return 0;
  const da = isoDate_(st), db = isoDate_(en);
  if (da && db) {
    const t = function (d, x) { const q = d.split('-'); return Date.UTC(+q[0], +q[1] - 1, +q[2]) / 1000 + x; };
    return Math.max(0, t(db, b) - t(da, a));
  }
  let d = b - a;
  if (d < 0) d += 86400;
  return d;
}