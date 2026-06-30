
        function toggleFiltersWrap() {
          const chips = document.getElementById('screenerFilterChips');
          const toggles = document.getElementById('screenerTogglesRow');
          const btn = document.getElementById('filtersToggleBtn');
          const collapsed = chips.style.display === 'none';
          chips.style.display = collapsed ? '' : 'none';
          toggles.style.display = collapsed ? '' : 'none';
          document.getElementById('filtersToggleLabel').textContent = collapsed ? 'Hide Filters' : 'Filters';
        }
        // Default to collapsed on mobile so the screener has more vertical
        // room right away; desktop is untouched (filters stay visible).
        if (window.innerWidth <= 768) {
          const chips = document.getElementById('screenerFilterChips');
          const toggles = document.getElementById('screenerTogglesRow');
          if (chips && toggles) {
            chips.style.display = 'none';
            toggles.style.display = 'none';
            document.getElementById('filtersToggleLabel').textContent = 'Filters';
          }
        }
      


// ===== EMBEDDED DATA =====
const SOURCE_DATA = [];

// ===== SECTOR_DATA: read directly from SectorAnalysis sheet on upload =====
function parseSectorSheet(wb) {
  const sheetName = wb.SheetNames.find(n => n.toLowerCase() === 'sectoranalysis');
  if (!sheetName) return [];
  const ws = wb.Sheets[sheetName];
  const allRows = XLSX.utils.sheet_to_json(ws, {header: 1, defval: null});

  // Find header row (the row where col 0 is "Sector")
  const headerIdx = allRows.findIndex(r => r[0] && String(r[0]).trim().toLowerCase() === 'sector');
  if (headerIdx < 0) return [];
  const headerRow = allRows[headerIdx];
  const dataRows  = allRows.slice(headerIdx + 1);

  // Log the full header so we can verify column names in console
  console.log('[SectorAnalysis] Header row:', headerRow.map((h, i) => `${i}:${h}`).join(' | '));

  const toNum = v => (v != null && v !== '-' && v !== '' && !isNaN(Number(v))) ? Number(v) : null;

  // Build a map of normalised header name → column index
  const colIdx = {};
  headerRow.forEach((h, i) => {
    if (h != null) colIdx[String(h).trim().toLowerCase()] = i;
  });

  // Helper: find column index by trying a list of candidate names (exact, then partial)
  const fc = (...names) => {
    for (const n of names) {
      const nl = n.trim().toLowerCase();
      if (colIdx[nl] != null) return colIdx[nl];           // exact match
      const partial = Object.keys(colIdx).find(k => k.includes(nl) || nl.includes(k));
      if (partial) return colIdx[partial];                  // partial match
    }
    return null; // not found — will result in null values
  };

  // Fixed columns 0-12 (always the same positions in your sheet)
  // Price columns: search by every plausible name variant
  const C = {
    p1d:          fc('day change %',  '1d %',  '1d%',  'day change',  '1 day %'),
    p1w:          fc('current week return %',  '1w %',  '1w%',  'week return %', 'weekly return %', '1 week %', 'week %'),
    p1m:          fc('current month return %', '1m %',  '1m%',  'month return %','monthly return %','1 month %','month %'),
    p3m:          fc('past 3 months return %', '3m %',  '3m%',  '3 month return %','3months %','3 months %','3m return %','quarterly return %'),
    pYTD:         fc('ytd return %',  'ytd %', 'ytd%',  'year to date %', 'ytd'),
    relVol:       fc('relative vol', 'rel vol', 'rel_vol', 'relative volume', 'relvol'),
    discRatio:    fc('discount ratio', 'discount_ratio', 'discountratio', 'disc ratio'),
    divYield:     fc('dividend yield %', 'dividend yield', 'div yield %', 'div yield', 'div y'),
    peRatio:      fc('p/e ratio', 'pe ratio', 'p/e', 'pe'),
  };

  // Log which column indices were resolved
  console.log('[SectorAnalysis] Price column indices resolved:', JSON.stringify(C));

  return dataRows
    .filter(r => r[0] && typeof r[0] === 'string' && r[0].trim() !== '')
    .map(r => ({
      sector:     r[0].trim(),
      companies:  toNum(r[1]) || 0,
      epsQ:       toNum(r[2]),
      epsTTM:     toNum(r[3]),
      opMargin:   toNum(r[4]),
      roe:        toNum(r[6]),
      de:         toNum(r[7]),
      cfo:        toNum(r[8]),
      rev:        toNum(r[9]),
      qtrScore:   toNum(r[10]) || 0,
      ttmScore:   toNum(r[11]) || 0,
      totalScore: toNum(r[12]) || 0,
      p1d:        C.p1d      != null ? toNum(r[C.p1d])      : null,
      p1w:        C.p1w      != null ? toNum(r[C.p1w])      : null,
      p1m:        C.p1m      != null ? toNum(r[C.p1m])      : null,
      p3m:        C.p3m      != null ? toNum(r[C.p3m])      : null,
      pYTD:       C.pYTD     != null ? toNum(r[C.pYTD])     : null,
      relVol:     C.relVol   != null ? toNum(r[C.relVol])   : null,
      discRatio:  C.discRatio!= null ? toNum(r[C.discRatio]): null,
      divYield:   C.divYield != null ? toNum(r[C.divYield]) : null,
      peRatio:    C.peRatio  != null ? toNum(r[C.peRatio])  : null,
    }));
}
let SECTOR_DATA = [];

// ===== CHART INSTANCES =====
let charts = {};
let allTickers = [];
let screenerData = [];
let screenerPage = 1;
const PAGE_SIZE = 100;
let screenerSort = {col: 14, dir: -1};
let sectorSort = {col: 12, dir: -1};

// ===== INIT =====
function init() {
  buildColMap(SOURCE_DATA);
  // Build ticker list
  allTickers = SOURCE_DATA
    .filter(d => d.Ticker && d.Ticker !== '0' && d.Ticker !== 0)
    .map(d => ({ticker: String(d.Ticker), name: String(d.Name || ''), score: d['total improvement'] || 0}))
    .sort((a,b) => b.score - a.score);

  // Populate sector & index filters (multi-select)
  allSectors = [...new Set(SOURCE_DATA.filter(d=>d.Sector && d.Sector!=='0').map(d=>d.Sector))].sort();
  allTickersList = [...new Set(SOURCE_DATA.filter(d=>d.Ticker).map(d=>String(d.Ticker)))].sort();
  const allIndicesSet = new Set();
  SOURCE_DATA.forEach(d => { if(d.Index) String(d.Index).split(',').forEach(i => allIndicesSet.add(i.trim())); });
  allIndicesList = [...allIndicesSet].filter(i=>i&&i!=='0').sort();
  resetMselFilters();

  // Load first ticker
  if (allTickers.length > 0) loadTicker(allTickers[0].ticker);

  // Build sector chart
  buildSectorChart();
  buildSectorTable();

  // Init screener
  screenerData = SOURCE_DATA
    .filter(d => d.Ticker && d.Ticker !== '0' && d.Ticker !== 0)
    .map(d => ({...d, Ticker: String(d.Ticker), Name: String(d.Name||''), Index: String(d.Index||'')}));
  renderScreener();
  initWatchlist();
}

let acIndex = -1;
let acFiltered = [];

function onTickerInput() {
  acIndex = -1;
  openAC();
}

function openAC() {
  const q = document.getElementById('tickerSearch').value.toLowerCase().trim();
  acFiltered = q
    ? allTickers.filter(t => t.ticker.toLowerCase().includes(q) || t.name.toLowerCase().includes(q)).slice(0, 50)
    : allTickers.slice(0, 50);
  renderAC();
  document.getElementById('acList').classList.add('open');
}

function renderAC() {
  const list = document.getElementById('acList');
  if (acFiltered.length === 0) {
    list.innerHTML = '<div class="ac-item"><span class="ac-name">No results found</span></div>';
    return;
  }
  list.innerHTML = acFiltered.map((t, i) => {
    const score = t.score;
    const sc = score >= 80 ? 'var(--success)' : score >= 40 ? 'var(--warn)' : 'var(--danger)';
    return `<div class="ac-item ${i===acIndex?'selected':''}" onmousedown="pickTicker('${t.ticker}')">
      <span class="ac-ticker">${t.ticker}</span>
      <span class="ac-score" style="color:${sc}">${score || '—'}</span>
      <div class="ac-name">${t.name || ''}</div>
    </div>`;
  }).join('');
}

function onTickerKey(e) {
  const list = document.getElementById('acList');
  if (!list.classList.contains('open')) { openAC(); return; }
  if (e.key === 'ArrowDown') { e.preventDefault(); acIndex = Math.min(acIndex+1, acFiltered.length-1); renderAC(); scrollAC(); }
  else if (e.key === 'ArrowUp') { e.preventDefault(); acIndex = Math.max(acIndex-1, 0); renderAC(); scrollAC(); }
  else if (e.key === 'Enter') { e.preventDefault(); if (acIndex >= 0 && acFiltered[acIndex]) pickTicker(acFiltered[acIndex].ticker); else if (acFiltered.length > 0) pickTicker(acFiltered[0].ticker); }
  else if (e.key === 'Escape') { closeAC(); }
}

function scrollAC() {
  const list = document.getElementById('acList');
  const sel = list.querySelector('.selected');
  if (sel) sel.scrollIntoView({block:'nearest'});
}

function pickTicker(ticker) {
  currentTicker = ticker;
  document.getElementById('tickerSearch').value = ticker;
  document.getElementById('currentTickerLabel').innerHTML = 'Selected: <span>' + ticker + '</span>';
  closeAC();
  loadTicker(ticker);
}

function closeAC() {
  document.getElementById('acList').classList.remove('open');
}

// Close AC when clicking outside
document.addEventListener('click', e => {
  if (!e.target.closest('.ticker-selector')) closeAC();
});

// ===== COLUMN NAME NORMALIZER =====
// After SheetJS upload, column names are preserved exactly. But build a lookup map
// so we can find columns even if spacing differs slightly.
let COL_MAP = {};
function buildColMap(data) {
  COL_MAP = {};
  if (!data || data.length === 0) return;
  Object.keys(data[0]).forEach(k => {
    // key: normalized (collapsed spaces, trimmed) → original
    COL_MAP[k.trim().replace(/\s+/g,' ')] = k;
  });
}
function col(name) {
  // Return the actual column key in the data, falling back to the name itself
  return COL_MAP[name.trim().replace(/\s+/g,' ')] || name;
}
function dget(row, name) {
  return row[col(name)];
}

// ===== EXCEL SERIAL DATE HELPER =====
function excelSerialToDateStr(serial) {
  // Excel serial: days since 1900-01-01 (with leap year bug)
  const utc = (serial - 25569) * 86400 * 1000;
  const d = new Date(utc);
  const y = d.getUTCFullYear();
  const m = String(d.getUTCMonth()+1).padStart(2,'0');
  const dd = String(d.getUTCDate()).padStart(2,'0');
  return `${y}-${m}-${dd}`;
}

function getDateStr(val) {
  if (!val) return '';
  if (typeof val === 'number') return excelSerialToDateStr(val);
  return String(val).substring(0, 10);
}

// ===== QUARTER KEY HELPER =====
// Maps a company's Last Period End Date to the correct CFO/ROE/D/E column keys
// and derives the 4 quarter labels for charts
function getQuarterKeys(lastDate) {
  const dateStr = getDateStr(lastDate);

  // All available quarter columns in the data (in chronological order)
  const allQuarters = [
    { label: '2025-Q3', date: '2025-09-30', suffix: '2025-Q3' },
    { label: '2025-Q4', date: '2025-12-31', suffix: '2025-Q4' },
    { label: '2026-Q1', date: '2026-03-31', suffix: '2026-Q1' },
  ];

  // Map end date → which 3 hardcoded columns to use (Q-3 relative position, Q-2, latest)
  // The EPS/Rev/NI cols are already relative (Q-3, Q-2, Q-1, Latest) so we only need
  // to figure out which 3 of the fixed-name CFO/ROE/DE cols match Q-3, Q-2, Q-1(latest)
  // for this company's reporting cycle.
  //
  // Latest date       → Latest col     → Q-2 col       → Q-3 col
  // 2026-03-31 (Q1)   → 2026-Q1        → 2025-Q4       → 2025-Q3
  // 2025-12-31 (Q4)   → 2025-Q4        → 2025-Q3       → (no Q-5, use Q3 again)
  // 2025-09-30 (Q3)   → 2025-Q3        → (no earlier)  → (no earlier)
  // 2025-06-30 (Q2)   → 2025-Q3 best   → —             → —

  let idxLatest;
  if (dateStr >= '2026-03-01') idxLatest = 2;       // 2026-Q1
  else if (dateStr >= '2025-12-01') idxLatest = 1;  // 2025-Q4
  else idxLatest = 0;                               // 2025-Q3 or earlier

  // Pick 3 indices: [Q-3, Q-2, latest]
  const i2 = idxLatest;
  const i1 = Math.max(0, idxLatest - 1);
  const i0 = Math.max(0, idxLatest - 2);

  const qAt = (i) => allQuarters[i].suffix;

  // Chart labels: derive 4 quarters ending at lastDate
  const chartLabels = quartersBefore(dateStr, 4);

  return {
    cfo:    [`CFO ${qAt(i0)}`,          `CFO ${qAt(i1)}`,          `CFO ${qAt(i2)}`],
    roe:    [`ROE ${qAt(i0)}`,          `ROE ${qAt(i1)}`,          `ROE ${qAt(i2)}`],
    de:     [`Debt/Equity ${qAt(i0)}`,  `Debt/Equity ${qAt(i1)}`,  `Debt/Equity ${qAt(i2)}`],
    labels: chartLabels,
  };
}

function quartersBefore(endDateStr, n) {
  // Returns n quarter-end labels ending at endDateStr (YYYY-MM-DD)
  const qEnds = ['03-31','06-30','09-30','12-31'];
  const qLabels = ['Q1','Q2','Q3','Q4'];
  const d = new Date(endDateStr);
  const year = d.getFullYear();
  const month = d.getMonth() + 1; // 1-based
  // find which quarter index this is
  let qi = month <= 3 ? 0 : month <= 6 ? 1 : month <= 9 ? 2 : 3;
  let y = year;
  const result = [];
  for (let i = 0; i < n; i++) {
    result.unshift(`${y}-${qLabels[qi]}`);
    qi--;
    if (qi < 0) { qi = 3; y--; }
  }
  return result;
}

// ===== LOAD TICKER =====
function loadTicker(ticker) {
  const d = SOURCE_DATA.find(r => String(r.Ticker) === String(ticker));
  if (!d) return;

  // Single source of truth for the "Selected: TICKER" label. Previously this
  // was only set inside pickTicker() (the manual search-selection path), so
  // the automatic first-ticker load in init() — which calls loadTicker()
  // directly — never updated it, leaving it stuck on its static HTML
  // placeholder forever. That stale value then fed into the EPS Trend / KPI
  // Summary header badges (which read from this label), showing the wrong
  // ticker even though the rest of Company View was correct.
  const tickerLabelEl = document.getElementById('currentTickerLabel');
  if (tickerLabelEl) tickerLabelEl.innerHTML = 'Selected: <span>' + ticker + '</span>';

  // Same badges shown next to the ticker in the Stock Screener table
  setHTML('ciTickerBadges', tickerBadges(d));

  // Company info
  setEl('ciName',   d.Name   || '—');
  setEl('ciSector', d.Sector || '—');
  setEl('ciIndex',  d.Index  || '—');
  const rawDate = dget(d,'Last Period End Date');
  setEl('ciPeriod', getDateStr(rawDate) || '—');
  const price = dget(d,'Price');
  setEl('ciPrice', price != null ? Number(price).toFixed(2) : '—');

  const score = dget(d,'total improvement');
  setEl('ciScore', score != null ? score : '—');
  setClass('ciScore', 'score-value ' + scoreColor(score));
  setEl('ciScoreHint', score >= 80 ? '🟢 Strong' : score >= 40 ? '🟡 Moderate' : '🔴 Weak');

  // Signal info
  const sigDate = dget(d,'Signal date');
  const sigStatus = dget(d,'Signal Status');
  const sigReturn = dget(d,'Signal Return %');
  const sigDateStr = sigDate ? fmtSignalDate(sigDate) : '—';
  setEl('ciSignalDate', sigDateStr);
  const sigStatusStr = sigStatusLabel(sigStatus) || 'No trade';
  const sigCode = sigStatusCode(sigStatus);
  const sigPillClass = (sigCode === 0 || sigCode == null) ? '' : sigStatusPillClass(sigStatus);
  setHTML('ciSignalStatus', `<span class="pill ${sigPillClass}" style="font-size:12px;padding:3px 9px;">${sigStatusStr}</span>`);
  const sigRet = toNum(sigReturn);
  const sigRetStr = sigRet != null ? ((sigRet >= 0 ? '+' : '') + sigRet.toFixed(2) + '%') : '—';
  setEl('ciSignalReturn', sigRetStr);
  setClass('ciSignalReturn', 'info-value mono ' + (sigRet != null ? (sigRet >= 0 ? 'positive' : 'negative') : ''));

  // KPI Cards
  const epsQ  = dget(d,'Latest EPS  Q');
  const epsQ1 = dget(d,'EPS  Q-1');
  setEl('kpiEPS', fmt(epsQ, 3));
  setClass('kpiEPS', 'kpi-value mono ' + (toNum(epsQ) > 0 ? 'positive' : toNum(epsQ) < 0 ? 'negative' : ''));
  setHTML('kpiEPSTrend', trendIcon(epsQ, epsQ1) + ` vs prev: ${fmt(epsQ1, 3)}`);
  setEl('kpiTTM', fmt(dget(d,'Latest TTM EPS Q'), 3));

  const rev     = dget(d,'Revenue - Q');
  const revPrev = dget(d,'Revenue - Q-1');
  setEl('kpiRev', fmtBig(rev));
  setHTML('kpiRevTrend', trendIcon(rev, revPrev) + ` prev: ${fmtBig(revPrev)}`);
  setEl('kpiOpMgn', fmtPct(dget(d,'Op Income-Q')));

  // Dynamic quarter keys based on Last Period End Date
  const qKeys = getQuarterKeys(rawDate);

  const roe  = dget(d, qKeys.roe[2]);
  const roeP = dget(d, qKeys.roe[1]);
  setEl('kpiROE', fmtPct(roe));
  setClass('kpiROE', 'kpi-value mono ' + (toNum(roe) > 0 ? 'positive' : toNum(roe) < 0 ? 'negative' : ''));
  setHTML('kpiROETrend', trendIcon(roe, roeP) + ` prev: ${fmtPct(roeP)}`);
  setEl('kpiDE', fmt(dget(d, qKeys.de[2]), 3));

  const cfo  = dget(d, qKeys.cfo[2]);
  const cfoP = dget(d, qKeys.cfo[1]);
  setEl('kpiCFO', fmtBig(cfo));
  setClass('kpiCFO', 'kpi-value mono ' + (toNum(cfo) > 0 ? 'positive' : toNum(cfo) < 0 ? 'negative' : ''));
  setHTML('kpiCFOTrend', trendIcon(cfo, cfoP) + ` prev: ${fmtBig(cfoP)}`);

  const sig = toNum(dget(d,'Signal Return %'));
  setEl('kpiSig', sig != null ? (sig >= 0 ? '+' : '') + sig.toFixed(2) + '%' : '—');
  setClass('kpiSig', 'mono ' + (sig > 0 ? 'positive' : sig < 0 ? 'negative' : ''));

  const periods = qKeys.labels;

  buildLineChart('chartEPS', periods,
    [dget(d,'EPS  Q-3'), dget(d,'EPS  Q-2'), dget(d,'EPS  Q-1'), dget(d,'Latest EPS  Q')],
    [dget(d,'TTM EPS Q-3'), dget(d,'TTM EPS Q-2'), dget(d,'TTM EPS Q-1'), dget(d,'Latest TTM EPS Q')],
    'EPS (Q)', 'EPS (TTM)'
  );
  buildBarChart('chartRev', periods,
    [dget(d,'Revenue - Q-3'), dget(d,'Revenue - Q-2'), dget(d,'Revenue - Q-1'), dget(d,'Revenue - Q')],
    'var(--accent3)'
  );
  buildLineChart('chartMargins', periods,
    [dget(d,'Op Income-Q-3'), dget(d,'Op Income-Q-2'), dget(d,'Op Income-Q-1'), dget(d,'Op Income-Q')],
    [dget(d,'Net Income -Q-3'), dget(d,'Net Income -Q-2'), dget(d,'Net Income -Q-1'), dget(d,'Net Income -Q')],
    'Op Margin', 'Net Margin', true
  );

  const sectorInfo = SECTOR_DATA.find(s => s.sector === d.Sector);

  // KPI table header
  setHTML('kpiTableHead', `
    <th>Metric</th>
    <th>${periods[0]}</th><th>${periods[1]}</th><th>${periods[2]}</th>
    <th style="color:var(--accent)">${periods[3]} ★</th>
    <th style="color:var(--accent2)">Sector Avg</th>
  `);

  const tbody = document.getElementById('kpiTableBody');
  if (tbody) {
    tbody.innerHTML = '';
    const kpiRows = [
      {label:'EPS (Q)',    vals:[dget(d,'EPS  Q-3'),dget(d,'EPS  Q-2'),dget(d,'EPS  Q-1'),dget(d,'Latest EPS  Q')], sectorVal:sectorInfo?.epsQ, pct:false},
      {label:'EPS (TTM)', vals:[dget(d,'TTM EPS Q-3'),dget(d,'TTM EPS Q-2'),dget(d,'TTM EPS Q-1'),dget(d,'Latest TTM EPS Q')], sectorVal:sectorInfo?.epsTTM, pct:false},
      {label:'Revenue',   vals:[dget(d,'Revenue - Q-3'),dget(d,'Revenue - Q-2'),dget(d,'Revenue - Q-1'),dget(d,'Revenue - Q')], sectorVal:null, pct:false, big:true},
      {label:'Op Margin', vals:[dget(d,'Op Income-Q-3'),dget(d,'Op Income-Q-2'),dget(d,'Op Income-Q-1'),dget(d,'Op Income-Q')], sectorVal:sectorInfo?.opMargin, pct:true},
      {label:'Net Margin',vals:[dget(d,'Net Income -Q-3'),dget(d,'Net Income -Q-2'),dget(d,'Net Income -Q-1'),dget(d,'Net Income -Q')], sectorVal:null, pct:true},
      {label:'ROE',       vals:[null,dget(d,qKeys.roe[0]),dget(d,qKeys.roe[1]),dget(d,qKeys.roe[2])], sectorVal:sectorInfo?.roe, pct:true},
      {label:'CFO',       vals:[null,dget(d,qKeys.cfo[0]),dget(d,qKeys.cfo[1]),dget(d,qKeys.cfo[2])], sectorVal:sectorInfo?.cfo, pct:false, big:true},
      {label:'D/E',       vals:[null,dget(d,qKeys.de[0]),dget(d,qKeys.de[1]),dget(d,qKeys.de[2])], sectorVal:sectorInfo?.de, pct:false},
      {label:'Div Yield', vals:[null,dget(d,'Div Y Q-2'),dget(d,'Div Y Q-1'),dget(d,'Latest Div Y Q')], sectorVal:sectorInfo?.divYield, pct:true},
    ];
    kpiRows.forEach(row => {
      const tr = document.createElement('tr');
      const fmtFn = row.big ? fmtBig : (row.pctRaw ? (v => { const n = toNum(v); return n == null ? '—' : n.toFixed(1) + '%'; }) : (row.pct ? fmtPct : v => fmt(v,3)));
      tr.innerHTML = `
        <td class="metric-name">${row.label}</td>
        ${row.vals.slice(0,3).map(v => `<td class="q-val ${valColor(v)}">${fmtFn(v)}</td>`).join('')}
        <td class="q-val ${valColor(row.vals[3])}" style="font-weight:700">${fmtFn(row.vals[3])}</td>
        <td class="sector-val">${row.sectorVal != null ? fmtFn(row.sectorVal) : '—'}</td>
      `;
      tbody.appendChild(tr);
    });
  }

  // Fin. Score Breakdown (hidden but still computed safely)
  const grid = document.getElementById('breakdownGrid');
  if (grid) grid.innerHTML = '';

  // YTD peer comparison chart
  buildPeerScoreChart(ticker, d.Sector);
  buildPeerYTDChart(ticker, d.Sector);

  // Sync the EPS Trend / KPI Summary header badges immediately rather than
  // waiting for their background poll (every 800ms-1000ms) to catch up —
  // that poll is now just a defensive fallback, not the primary mechanism.
  if (typeof updateKpiBadge === 'function') updateKpiBadge();
  if (typeof updateEPSHeader === 'function') updateEPSHeader();
}

// ===== CHARTS =====

// Shared chart defaults
const CHART_FONT = "'IBM Plex Mono', 'Courier New', monospace";

// Read live CSS variable (resolves per active theme at call time)
function cssVar(name, fallback) {
  return getComputedStyle(document.documentElement).getPropertyValue(name).trim() || fallback;
}
function getChartTheme() {
  const t = document.documentElement.getAttribute('data-theme') || 'dark';
  if (t === 'light') return {
    grid:    'rgba(0,0,0,0.07)',
    tick:    '#64748b',
    tooltip: { bg: '#ffffff', border: '#d3daea', title: '#0f172a', body: '#475569' },
    accent:  '#0d9e7e',
    accent2: '#ea6b10',
    accent3: '#2563eb',
    good:    '#16a34a',
    bad:     '#dc2626',
    warn:    '#d97706',
    zero:    'rgba(0,0,0,0.15)',
    barPos:  'rgba(37,99,235,0.65)',
    barNeg:  'rgba(234,107,16,0.65)',
    barLast: 'rgba(13,158,126,0.85)',
    barPosBorder: '#2563eb',
    barNegBorder: '#ea6b10',
    barLastBorder:'#0d9e7e',
    scoreHigh: 'rgba(22,163,74,0.70)',
    scoreMid:  'rgba(217,119,6,0.70)',
    scoreLow:  'rgba(220,38,38,0.70)',
    scoreHighBorder: '#16a34a',
    scoreMidBorder:  '#d97706',
    scoreLowBorder:  '#dc2626',
    selectedBar: 'rgba(13,158,126,0.90)',
    selectedBarBorder: '#0d9e7e',
    lineGrad1Top: 'rgba(13,158,126,0.28)', lineGrad1Bot: 'rgba(13,158,126,0)',
    lineGrad2Top: 'rgba(234,107,16,0.20)',  lineGrad2Bot: 'rgba(234,107,16,0)',
    line1Color: '#0d9e7e', line2Color: '#ea6b10',
    pt1Pos: '#0d9e7e', pt1Neg: '#dc2626',
    pt2Pos: '#ea6b10', pt2Neg: '#dc2626',
  };
  if (t === 'bloomberg') return {
    grid:    'rgba(255,102,0,0.08)',
    tick:    '#cc5500',
    tooltip: { bg: '#111111', border: '#ff6600', title: '#ff6600', body: '#cc5500' },
    accent:  '#ff6600',
    accent2: '#ffcc00',
    accent3: '#00aaff',
    good:    '#00cc44',
    bad:     '#ff3333',
    warn:    '#ffcc00',
    zero:    'rgba(255,102,0,0.25)',
    barPos:  'rgba(0,170,255,0.55)',
    barNeg:  'rgba(255,51,51,0.55)',
    barLast: 'rgba(255,102,0,0.90)',
    barPosBorder: '#00aaff',
    barNegBorder: '#ff3333',
    barLastBorder:'#ff6600',
    scoreHigh: 'rgba(0,204,68,0.65)',
    scoreMid:  'rgba(255,204,0,0.65)',
    scoreLow:  'rgba(255,51,51,0.65)',
    scoreHighBorder: '#00cc44',
    scoreMidBorder:  '#ffcc00',
    scoreLowBorder:  '#ff3333',
    selectedBar: 'rgba(255,102,0,0.90)',
    selectedBarBorder: '#ff6600',
    lineGrad1Top: 'rgba(255,102,0,0.30)', lineGrad1Bot: 'rgba(255,102,0,0)',
    lineGrad2Top: 'rgba(255,204,0,0.20)',  lineGrad2Bot: 'rgba(255,204,0,0)',
    line1Color: '#ff6600', line2Color: '#ffcc00',
    pt1Pos: '#ff6600', pt1Neg: '#ff3333',
    pt2Pos: '#ffcc00', pt2Neg: '#ff3333',
  };
  if (t === 'tradingview') return {
    grid:    'rgba(255,255,255,0.04)',
    tick:    '#787b86',
    tooltip: { bg: '#2a2e39', border: '#363a45', title: '#d1d4dc', body: '#787b86' },
    accent:  '#2962ff',
    accent2: '#ff9800',
    accent3: '#00bcd4',
    good:    '#089981',
    bad:     '#f23645',
    warn:    '#ff9800',
    zero:    'rgba(255,255,255,0.12)',
    barPos:  'rgba(8,153,129,0.65)',
    barNeg:  'rgba(242,54,69,0.65)',
    barLast: 'rgba(41,98,255,0.88)',
    barPosBorder: '#089981',
    barNegBorder: '#f23645',
    barLastBorder:'#2962ff',
    scoreHigh: 'rgba(8,153,129,0.60)',
    scoreMid:  'rgba(255,152,0,0.60)',
    scoreLow:  'rgba(242,54,69,0.55)',
    scoreHighBorder: '#089981',
    scoreMidBorder:  '#ff9800',
    scoreLowBorder:  '#f23645',
    selectedBar: 'rgba(41,98,255,0.90)',
    selectedBarBorder: '#2962ff',
    lineGrad1Top: 'rgba(41,98,255,0.28)', lineGrad1Bot: 'rgba(41,98,255,0)',
    lineGrad2Top: 'rgba(8,153,129,0.20)',  lineGrad2Bot: 'rgba(8,153,129,0)',
    line1Color: '#2962ff', line2Color: '#089981',
    pt1Pos: '#089981', pt1Neg: '#f23645',
    pt2Pos: '#2962ff', pt2Neg: '#f23645',
  };
  // dark (default)
  return {
    grid:    'rgba(255,255,255,0.05)',
    tick:    '#8a9bc0',
    tooltip: { bg: '#202c42', border: '#2e3a52', title: '#f0f4ff', body: '#8a9bc0' },
    accent:  '#00c49a',
    accent2: '#f97316',
    accent3: '#3b82f6',
    good:    '#22c55e',
    bad:     '#ef4444',
    warn:    '#f59e0b',
    zero:    'rgba(255,255,255,0.18)',
    barPos:  'rgba(59,130,246,0.68)',
    barNeg:  'rgba(249,115,22,0.68)',
    barLast: 'rgba(0,196,154,0.88)',
    barPosBorder: '#3b82f6',
    barNegBorder: '#f97316',
    barLastBorder:'#00c49a',
    scoreHigh: 'rgba(34,197,94,0.60)',
    scoreMid:  'rgba(245,158,11,0.60)',
    scoreLow:  'rgba(239,68,68,0.55)',
    scoreHighBorder: '#22c55e',
    scoreMidBorder:  '#f59e0b',
    scoreLowBorder:  '#ef4444',
    selectedBar: 'rgba(0,196,154,0.90)',
    selectedBarBorder: '#00c49a',
    lineGrad1Top: 'rgba(0,196,154,0.26)', lineGrad1Bot: 'rgba(0,196,154,0)',
    lineGrad2Top: 'rgba(249,115,22,0.18)',  lineGrad2Bot: 'rgba(249,115,22,0)',
    line1Color: '#00c49a', line2Color: '#f97316',
    pt1Pos: '#00c49a', pt1Neg: '#ef4444',
    pt2Pos: '#f97316', pt2Neg: '#ef4444',
  };
}

function makeGradient(ctx, canvas, hexTop, hexBot) {
  const gr = ctx.createLinearGradient(0, 0, 0, canvas.offsetHeight || 220);
  gr.addColorStop(0, hexTop);
  gr.addColorStop(1, hexBot);
  return gr;
}

function sharedScales(yFmt) {
  const th = getChartTheme();
  return {
    x: {
      grid: { color: th.grid, drawBorder: false },
      ticks: { color: th.tick, font: { size: 10, family: CHART_FONT }, maxRotation: 0 },
      border: { display: false }
    },
    y: {
      grid: { color: th.grid, drawBorder: false },
      ticks: { color: th.tick, font: { size: 10, family: CHART_FONT }, callback: yFmt || (v => v) },
      border: { display: false }
    }
  };
}

function sharedTooltip(valueFmt) {
  const th = getChartTheme();
  return {
    backgroundColor: th.tooltip.bg,
    borderColor: th.tooltip.border,
    borderWidth: 1,
    titleColor: th.tooltip.title,
    bodyColor: th.tooltip.body,
    padding: 10,
    cornerRadius: 6,
    titleFont: { size: 11, family: CHART_FONT },
    bodyFont:  { size: 11, family: CHART_FONT },
    callbacks: valueFmt ? { label: ctx => ` ${ctx.dataset.label}: ${valueFmt(ctx.raw)}` } : {}
  };
}

function buildLineChart(id, labels, data1, data2, label1, label2, isPercent) {
  if (charts[id]) { charts[id].destroy(); }
  const canvas = document.getElementById(id);
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  const th = getChartTheme();

  // Gradient fills — resolved to real hex at render time
  const grad1 = ctx.createLinearGradient(0, 0, 0, 260);
  grad1.addColorStop(0,   th.lineGrad1Top);
  grad1.addColorStop(0.7, th.lineGrad1Bot);
  grad1.addColorStop(1,   th.lineGrad1Bot);

  const grad2 = ctx.createLinearGradient(0, 0, 0, 260);
  grad2.addColorStop(0,   th.lineGrad2Top);
  grad2.addColorStop(1,   th.lineGrad2Bot);

  const ptColor1 = (data1||[]).map(v => (v == null ? th.tick : v >= 0 ? th.pt1Pos : th.pt1Neg));
  const ptColor2 = (data2||[]).map(v => (v == null ? th.tick : v >= 0 ? th.pt2Pos : th.pt2Neg));

  charts[id] = new Chart(ctx, {
    type: 'line',
    data: {
      labels,
      datasets: [
        {
          label: label1,
          data: data1,
          borderColor: th.line1Color,
          backgroundColor: grad1,
          tension: 0.35,
          fill: true,
          borderWidth: 2.5,
          pointRadius: 5,
          pointHoverRadius: 8,
          pointBackgroundColor: ptColor1,
          pointBorderColor: th.tooltip.bg,
          pointBorderWidth: 2,
          pointHoverBorderWidth: 2,
          pointHoverBackgroundColor: th.line1Color,
        },
        {
          label: label2,
          data: data2,
          borderColor: th.line2Color,
          backgroundColor: grad2,
          tension: 0.35,
          fill: true,
          borderWidth: 2,
          borderDash: [5, 4],
          pointRadius: 4,
          pointHoverRadius: 7,
          pointBackgroundColor: ptColor2,
          pointBorderColor: th.tooltip.bg,
          pointBorderWidth: 2,
          pointHoverBackgroundColor: th.line2Color,
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 400, easing: 'easeOutCubic' },
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: {
          labels: {
            color: th.tick,
            font: { size: 10, family: CHART_FONT },
            usePointStyle: true,
            pointStyleWidth: 10,
            padding: 18,
            boxHeight: 2,
            generateLabels: chart => chart.data.datasets.map((ds, i) => ({
              text: ds.label,
              fillStyle: ds.borderColor,
              strokeStyle: ds.borderColor,
              fontColor: ds.borderColor,
              lineWidth: 2,
              pointStyle: 'line',
              datasetIndex: i,
              hidden: !chart.isDatasetVisible(i),
            })),
          }
        },
        tooltip: sharedTooltip(v => v == null ? '—' : (isPercent ? (v*100).toFixed(1) + '%' : fmt(v, 3)))
      },
      scales: {
        x: {
          grid: { color: th.grid, drawBorder: false },
          ticks: { color: th.tick, font: { size: 10, family: CHART_FONT }, maxRotation: 0 },
          border: { display: false }
        },
        y: {
          grid: { color: th.grid, drawBorder: false },
          border: { display: false },
          ticks: {
            color: th.tick, font: { size: 10, family: CHART_FONT },
            callback: v => isPercent ? (v*100).toFixed(0) + '%' : v
          },
          afterBuildTicks: axis => {
            if (!axis.ticks.find(t => t.value === 0)) axis.ticks.push({ value: 0 });
            axis.ticks.sort((a,b) => a.value - b.value);
          }
        }
      }
    }
  });

  // Zero-line plugin
  const zeroId = 'zeroLine_' + id;
  const zeroPlugin = {
    id: zeroId,
    afterDraw(chart) {
      const yScale = chart.scales.y;
      if (!yScale) return;
      const y0 = yScale.getPixelForValue(0);
      if (y0 < yScale.top || y0 > yScale.bottom) return;
      const c = chart.ctx;
      c.save();
      c.beginPath();
      c.moveTo(chart.chartArea.left, y0);
      c.lineTo(chart.chartArea.right, y0);
      c.strokeStyle = getChartTheme().zero;
      c.lineWidth = 1;
      c.setLineDash([4, 4]);
      c.stroke();
      c.restore();
    }
  };
  charts[id].options._zeroPlugin = zeroPlugin;
  Chart.register(zeroPlugin);
}

function buildBarChart(id, labels, data, color) {
  if (charts[id]) { charts[id].destroy(); }
  const canvas = document.getElementById(id);
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  const th = getChartTheme();

  const barColors = (data||[]).map(v =>
    v == null ? th.barPos : v >= 0 ? th.barPos : th.barNeg
  );
  const barBorders = (data||[]).map(v =>
    v == null ? th.barPosBorder : v >= 0 ? th.barPosBorder : th.barNegBorder
  );

  // Highlight the most recent bar with accent color
  if (data && data.length > 0) {
    const last = data[data.length - 1];
    barColors[data.length - 1]  = last >= 0 ? th.barLast : th.barNeg;
    barBorders[data.length - 1] = last >= 0 ? th.barLastBorder : th.barNegBorder;
  }

  charts[id] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Revenue',
        data: data,
        backgroundColor: barColors,
        borderColor: barBorders,
        borderWidth: 1.5,
        borderRadius: 5,
        borderSkipped: false,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 400, easing: 'easeOutCubic' },
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { display: false },
        tooltip: sharedTooltip(v => v == null ? '—' : fmtBig(v))
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: { color: th.tick, font: { size: 10, family: CHART_FONT }, maxRotation: 0 },
          border: { display: false }
        },
        y: {
          grid: { color: th.grid, drawBorder: false },
          ticks: { color: th.tick, font: { size: 10, family: CHART_FONT }, callback: v => fmtBig(v) },
          border: { display: false }
        }
      }
    }
  });
}

function buildPeerScoreChart(ticker, sector) {
  if (charts['peerScore']) { charts['peerScore'].destroy(); delete charts['peerScore']; }
  const canvas = document.getElementById('chartPeerScore');
  if (!canvas) return;

  const peers = SOURCE_DATA
    .filter(r => r.Sector === sector && r.Ticker && r.Ticker !== '0')
    .map(r => ({
      ticker: String(r.Ticker),
      name:   String(r.Name || r.Ticker),
      score:  toNum(dget(r, 'total improvement'))
    }))
    .filter(p => p.score != null)
    .sort((a, b) => b.score - a.score);

  const labelEl = document.getElementById('peerScoreLabel');
  if (labelEl) labelEl.textContent = `${peers.length} companies · ${sector}`;

  const card = document.getElementById('peerScoreCard');
  if (peers.length === 0) { if (card) card.style.display = 'none'; return; }
  if (card) card.style.display = '';

  const barH = Math.max(200, peers.length * 26);
  canvas.parentElement.style.height = barH + 'px';

  const selectedIdx = peers.findIndex(p => p.ticker === String(ticker));

  const th = getChartTheme();
  const bgColors = peers.map((p, i) => {
    if (i === selectedIdx) return th.selectedBar;
    return p.score >= 80 ? th.scoreHigh : p.score >= 40 ? th.scoreMid : th.scoreLow;
  });
  const borderColors = peers.map((p, i) => {
    if (i === selectedIdx) return th.selectedBarBorder;
    return p.score >= 80 ? th.scoreHighBorder : p.score >= 40 ? th.scoreMidBorder : th.scoreLowBorder;
  });
  const borderWidths = peers.map((_, i) => i === selectedIdx ? 2.5 : 1);

  const ctx = canvas.getContext('2d');
  charts['peerScore'] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: peers.map(p => p.ticker),
      datasets: [{
        label: 'Financial Score',
        data: peers.map(p => p.score),
        backgroundColor: bgColors,
        borderColor: borderColors,
        borderWidth: borderWidths,
        borderRadius: 4,
        borderSkipped: false,
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 400, easing: 'easeOutCubic' },
      plugins: {
        legend: { display: false },
        tooltip: {
          ...sharedTooltip(),
          callbacks: {
            label: ctx => ` Financial Score: ${ctx.parsed.x.toFixed(1)}`,
            title: ctx => {
              const p = peers[ctx[0].dataIndex];
              return `${p.ticker}  ${p.name.length > 32 ? p.name.substring(0,30)+'…' : p.name}`;
            }
          }
        }
      },
      scales: {
        x: {
          grid: { color: th.grid, drawBorder: false },
          ticks: { color: th.tick, font: { size: 10, family: CHART_FONT }, callback: v => v.toFixed(0) },
          border: { display: false }
        },
        y: {
          grid: { display: false },
          ticks: {
            color: ctx => ctx.index === selectedIdx ? th.accent : th.tick,
            font: ctx => ({ size: 10, family: CHART_FONT, weight: ctx.index === selectedIdx ? '700' : '400' })
          },
          border: { display: false }
        }
      }
    }
  });
}

function buildPeerYTDChart(ticker, sector) {
  if (charts['peerYTD']) { charts['peerYTD'].destroy(); delete charts['peerYTD']; }
  const canvas = document.getElementById('chartPeerYTD');
  if (!canvas) return;

  // Gather all peers in the same sector with a valid YTD value
  const peers = SOURCE_DATA
    .filter(r => r.Sector === sector && r.Ticker && r.Ticker !== '0')
    .map(r => ({
      ticker: String(r.Ticker),
      name:   String(r.Name || r.Ticker),
      ytd:    toNum(dget(r, 'YTD Return %'))
    }))
    .filter(p => p.ytd != null)
    .sort((a, b) => b.ytd - a.ytd);

  // Update label
  const labelEl = document.getElementById('peerYTDLabel');
  if (labelEl) labelEl.textContent = `${peers.length} companies · ${sector}`;

  // Hide card if no data
  const card = document.getElementById('peerYTDCard');
  if (peers.length === 0) { if (card) card.style.display = 'none'; return; }
  if (card) card.style.display = '';

  // Dynamic height: 26px per bar, min 200px
  const barH = Math.max(200, peers.length * 26);
  canvas.parentElement.style.height = barH + 'px';

  const labels = peers.map(p => p.ticker);
  const values = peers.map(p => p.ytd);
  const selectedIdx = peers.findIndex(p => p.ticker === String(ticker));

  // Colour each bar using theme
  const th2 = getChartTheme();
  const bgColors = peers.map((p, i) => {
    if (i === selectedIdx) return th2.selectedBar;
    return p.ytd >= 0 ? th2.scoreHigh : th2.scoreLow;
  });
  const borderColors = peers.map((p, i) => {
    if (i === selectedIdx) return th2.selectedBarBorder;
    return p.ytd >= 0 ? th2.scoreHighBorder : th2.scoreLowBorder;
  });
  const borderWidths = peers.map((p, i) => i === selectedIdx ? 2.5 : 1);

  const ctx = canvas.getContext('2d');
  charts['peerYTD'] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'YTD Return %',
        data: values,
        backgroundColor: bgColors,
        borderColor: borderColors,
        borderWidth: borderWidths,
        borderRadius: 4,
        borderSkipped: false,
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 400, easing: 'easeOutCubic' },
      plugins: {
        legend: { display: false },
        tooltip: {
          ...sharedTooltip(),
          callbacks: {
            label: ctx => {
              const v = ctx.parsed.x;
              return ` ${v >= 0 ? '+' : ''}${v.toFixed(2)}%`;
            },
            title: ctx => {
              const p = peers[ctx[0].dataIndex];
              return `${p.ticker}  ${p.name.length > 32 ? p.name.substring(0,30)+'…' : p.name}`;
            }
          }
        }
      },
      scales: {
        x: {
          grid: { color: th2.grid, drawBorder: false },
          ticks: {
            color: th2.tick,
            font: { size: 10, family: CHART_FONT },
            callback: v => (v >= 0 ? '+' : '') + v.toFixed(1) + '%'
          },
          border: { display: false }
        },
        y: {
          grid: { display: false },
          ticks: {
            color: (ctx) => ctx.index === selectedIdx ? th2.accent : th2.tick,
            font: (ctx) => ({
              size: 10,
              family: CHART_FONT,
              weight: ctx.index === selectedIdx ? '700' : '400'
            })
          },
          border: { display: false }
        }
      }
    }
  });
}

function buildSectorChart() {
  if (charts['sectors']) { charts['sectors'].destroy(); delete charts['sectors']; }
  const canvas = document.getElementById('chartSectors');
  if (!canvas) return;
  const sorted = [...SECTOR_DATA].sort((a,b) => b.totalScore - a.totalScore).slice(0, 20);
  const ctx = canvas.getContext('2d');
  const th = getChartTheme();

  const bgColors  = sorted.map(s => s.totalScore >= 80 ? th.scoreHigh  : s.totalScore >= 40 ? th.scoreMid  : th.scoreLow);
  const brdColors = sorted.map(s => s.totalScore >= 80 ? th.scoreHighBorder : s.totalScore >= 40 ? th.scoreMidBorder : th.scoreLowBorder);

  charts['sectors'] = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: sorted.map(s => s.sector.length > 30 ? s.sector.substring(0,28)+'…' : s.sector),
      datasets: [{
        label: 'Financial Score',
        data: sorted.map(s => s.totalScore),
        backgroundColor: bgColors,
        borderColor: brdColors,
        borderWidth: 1.5,
        borderRadius: 4,
        borderSkipped: false,
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      animation: { duration: 400, easing: 'easeOutCubic' },
      plugins: {
        legend: { display: false },
        tooltip: {
          ...sharedTooltip(),
          callbacks: {
            label: ctx => ` Financial Score: ${ctx.parsed.x.toFixed(1)}`
          }
        }
      },
      scales: {
        x: {
          grid: { color: th.grid, drawBorder: false },
          ticks: { color: th.tick, font: { size: 10, family: CHART_FONT }, callback: v => v.toFixed(0) },
          border: { display: false },
          max: 110
        },
        y: {
          grid: { display: false },
          ticks: { color: th.tick, font: { size: 10, family: CHART_FONT } },
          border: { display: false }
        }
      }
    }
  });
}

// ===== SORT ARROW HELPER =====
function updateSortArrows(headRowId, activeCol, dir) {
  const headRow = document.getElementById(headRowId);
  if (!headRow) return;
  headRow.querySelectorAll('th').forEach((th, i) => {
    th.classList.remove('sort-asc', 'sort-desc');
    if (i === activeCol) th.classList.add(dir === 1 ? 'sort-asc' : 'sort-desc');
  });
}

// ===== SECTOR TABLE =====
let sectorTableData = [];
function buildSectorTable() {
  sectorTableData = [...SECTOR_DATA].sort((a,b) => b.totalScore - a.totalScore);
  updateSortArrows('sectorTableHead', sectorSort.col, sectorSort.dir);
  renderSectorTable(sectorTableData);
}
function renderSectorTable(data) {
  const tbody = document.getElementById('sectorTableBody');
  tbody.innerHTML = '';
  const fmtChg = v => {
    if (v == null) return '<span style="color:var(--text3)">—</span>';
    const pct = Number(v).toFixed(2);
    const cls = v > 0 ? 'positive' : v < 0 ? 'negative' : '';
    const arrow = v > 0 ? '▲' : v < 0 ? '▼' : '';
    return `<span class="mono ${cls}">${arrow}${Math.abs(pct)}%</span>`;
  };
  data.forEach(s => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="sector-name-cell">${s.sector}</td>
      <td class="sector-hide-mobile mono" style="text-align:center">${s.companies}</td>
      <td class="sector-hide-mobile mono ${valColor(s.epsQ)}">${fmt(s.epsQ,2)}</td>
      <td class="sector-hide-mobile mono ${valColor(s.epsTTM)}">${fmt(s.epsTTM,2)}</td>
      <td class="sector-hide-mobile mono ${valColor(s.opMargin)}">${s.opMargin!=null?fmtPct(s.opMargin):'—'}</td>
      <td class="sector-hide-mobile mono ${valColor(s.roe)}">${fmtPct(s.roe)}</td>
      <td class="sector-hide-mobile mono">${s.de!=null?fmt(s.de,2):'—'}</td>
      <td class="sector-hide-mobile mono">${fmt(s.cfo,1)}</td>
      <td class="sector-hide-mobile mono ${valColor(s.divYield)}">${s.divYield!=null?fmtPct(s.divYield):'—'}</td>
      <td class="sector-hide-mobile mono">${s.peRatio!=null?fmt(s.peRatio,2):'—'}</td>
      <td class="mono"><span class="${scoreColor(s.totalScore)}" style="font-weight:700">${fmt(s.totalScore,1)}</span>
        <div class="prog-bar"><div class="prog-fill" style="width:${Math.min(100,s.totalScore)}%; background:${s.totalScore>=60?'var(--success)':s.totalScore>=40?'var(--warn)':'var(--danger)'}"></div></div>
      </td>
      <td class="mono" style="text-align:center">${s.relVol!=null?fmt(s.relVol,2):'—'}</td>
      <td class="mono ${s.discRatio!=null?(s.discRatio>0?'positive':'negative'):''}" style="text-align:center">${s.discRatio!=null?fmt(s.discRatio,1)+'%':'—'}</td>
      <td class="sector-hide-mobile" style="text-align:center">${fmtChg(s.p1d)}</td>
      <td class="sector-hide-mobile" style="text-align:center">${fmtChg(s.p1w)}</td>
      <td style="text-align:center">${fmtChg(s.p1m)}</td>
      <td style="text-align:center">${fmtChg(s.p3m)}</td>
      <td style="text-align:center">${fmtChg(s.pYTD)}</td>
    `;
    tbody.appendChild(tr);
  });
}
function filterSectorTable() {
  const q = document.getElementById('sectorSearch').value.toLowerCase();
  const filtered = q ? SECTOR_DATA.filter(s => s.sector.toLowerCase().includes(q)) : [...SECTOR_DATA];
  renderSectorTable(filtered.sort((a,b) => b.totalScore - a.totalScore));
}
function sortSectorTable(col) {
  const cols = ['sector','companies','epsQ','epsTTM','opMargin','roe','de','cfo','divYield','peRatio','totalScore','relVol','discRatio','p1d','p1w','p1m','p3m','pYTD'];
  const key = cols[col];
  if (sectorSort.col === col) sectorSort.dir *= -1; else { sectorSort.col = col; sectorSort.dir = -1; }
  updateSortArrows('sectorTableHead', sectorSort.col, sectorSort.dir);
  const sorted = [...SECTOR_DATA].sort((a,b) => {
    const av = a[key], bv = b[key];
    if (av == null) return 1; if (bv == null) return -1;
    return typeof av === 'string' ? av.localeCompare(bv) * sectorSort.dir : (av - bv) * sectorSort.dir;
  });
  renderSectorTable(sorted);
}

// ===== SCREENER =====
let currentTicker = null;
let filteredScreener = [];
let allSectors = [];
let allIndicesList = [];
let allTickersList = [];

const SCORE_OPTIONS = [
  {value:'80,Infinity',   label:'Fin. Scores > 80'},
  {value:'50,80',         label:'Fin. Scores > 50 to 80'},
  {value:'40,50',         label:'Fin. Scores > 40 to 50'},
  {value:'20,40',         label:'Fin. Scores > 20 to 40'},
  {value:'-Infinity,20',  label:'Fin. Scores < 20'}
];
const STATUS_OPTIONS = [
  {value:'2', label:'Continuation Buy Signal'},
  {value:'2.5', label:'Extended Buy Signal (Cautious)'},
  {value:'1.5', label:'Initial Buy Signal'},
  {value:'1', label:'Hold Trade'},
  {value:'3', label:'Take Some Profit'},
  {value:'4', label:'Be Cautious'},
  {value:'9', label:'Buy Call Closed'},
  {value:'0', label:'No Trade'}
];
const EXTRA_OPTIONS = [
  {value:'mcap_gt_1b',     label:'Market Cap > 1 Billion'},
  {value:'mcap_gt_500m',   label:'Market Cap > 500 Million'},
  {value:'mcap_gt_250m',   label:'Market Cap > 250 Million'},
  {value:'mcap_lt_250m',   label:'Market Cap < 250 Million'}
];
const OTHERS_OPTIONS = [
  {value:'turnaround', label:'🔁 Turnaround Candidates'},
  {value:'net_gt_op',  label:'📊 Net Inc > Op. Inc'}
];
const LIQUID_OPTIONS = [
  {value:'1',  label:'🟢 Liquid Stock (+)'},
  {value:'-1', label:'🔴 Non-Liquid Stock (−)'}
];
const VOLPHASE_OPTIONS = [
  {value:'1',  label:'🟢 High Phase (+1)'},
  {value:'-1', label:'⚪ Normal Phase (-1)'}
];

// Generic multi-select registry. Each entry: options() returns [{value,label}], selected: Set of values, ids + labels for the button.
const mselRegistry = {
  sector: {
    options: () => allSectors.map(s => ({value:s, label:s})),
    selected: new Set(),
    searchable: true,
    allLabel: 'Sectors',
    oneLabel: v => v,
    manyLabel: n => `Sectors`,
  },
  index: {
    options: () => allIndicesList.map(s => ({value:s, label:s})),
    selected: new Set(),
    searchable: true,
    allLabel: 'Indices',
    oneLabel: v => v,
    manyLabel: n => `Indices`,
  },
  ticker: {
    options: () => allTickersList.map(s => ({value:s, label:s})),
    selected: new Set(),
    searchable: true,
    allLabel: 'Comparison',
    oneLabel: v => v,
    manyLabel: n => `Comparison`,
  },
  score: {
    options: () => SCORE_OPTIONS,
    selected: new Set(),
    searchable: false,
    allLabel: 'Fin. Scores',
    oneLabel: v => (SCORE_OPTIONS.find(o=>o.value===v)||{}).label || v,
    manyLabel: n => `Fin. Scores`,
  },
  status: {
    options: () => STATUS_OPTIONS,
    selected: new Set(),
    searchable: false,
    allLabel: 'Signals',
    oneLabel: v => (STATUS_OPTIONS.find(o=>o.value===v)||{}).label || v,
    manyLabel: n => `Signals`,
  },
  extra: {
    options: () => EXTRA_OPTIONS,
    selected: new Set(),
    searchable: false,
    allLabel: 'Market Cap',
    oneLabel: v => (EXTRA_OPTIONS.find(o=>o.value===v)||{}).label || v,
    manyLabel: n => `Filters`,
  },
  others: {
    options: () => OTHERS_OPTIONS,
    selected: new Set(),
    searchable: false,
    allLabel: 'Other Filters',
    oneLabel: v => (OTHERS_OPTIONS.find(o=>o.value===v)||{}).label || v,
    manyLabel: n => `Filters`,
  },
  liquid: {
    options: () => LIQUID_OPTIONS,
    selected: new Set(),
    searchable: false,
    allLabel: 'Liquidity',
    oneLabel: v => (LIQUID_OPTIONS.find(o=>o.value===v)||{}).label || v,
    manyLabel: n => `Liquidity`,
  },
  volPhase: {
    options: () => VOLPHASE_OPTIONS,
    selected: new Set(),
    searchable: false,
    allLabel: 'Volume Phase',
    oneLabel: v => (VOLPHASE_OPTIONS.find(o=>o.value===v)||{}).label || v,
    manyLabel: n => `Phases`,
  }
};

function escapeHtml(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function mselUpdateLabel(key) {
  const cfg = mselRegistry[key];
  const label = document.getElementById(key + 'MselLabel');
  const btn = document.getElementById(key + 'MselBtn');
  if (!label || !cfg) return;
  const sel = cfg.selected;
  if (sel.size === 0) {
    label.innerHTML = cfg.allLabel;
    btn.classList.remove('active');
  } else if (sel.size === 1) {
    label.innerHTML = escapeHtml(cfg.oneLabel([...sel][0]));
    btn.classList.add('active');
  } else {
    label.innerHTML = `${cfg.manyLabel(sel.size)} <span class="msel-count">${sel.size}</span>`;
    btn.classList.add('active');
  }
}

function mselRenderList(key) {
  const cfg = mselRegistry[key];
  const list = document.getElementById(key + 'MselList');
  if (!list || !cfg) return;
  const searchEl = document.getElementById(key + 'MselSearch');
  const q = (searchEl && searchEl.value || '').toLowerCase();
  let opts = cfg.options();
  if (q) opts = opts.filter(o => String(o.label).toLowerCase().includes(q));
  if (opts.length === 0) {
    list.innerHTML = '<div class="msel-empty">No matches found</div>';
    return;
  }
  list.innerHTML = opts.map(o => {
    const checked = cfg.selected.has(o.value) ? 'checked' : '';
    const safeVal = escapeHtml(o.value).replace(/'/g, "\\'");
    const safeLabel = escapeHtml(o.label);
    return `<label class="msel-item"><input type="checkbox" ${checked} onchange="mselToggleItem('${key}', '${safeVal}', this.checked)"><span class="msel-label" title="${safeLabel}">${safeLabel}</span></label>`;
  }).join('');
}

function mselToggleItem(key, value, checked) {
  const cfg = mselRegistry[key];
  if (!cfg) return;
  if (checked) cfg.selected.add(value); else cfg.selected.delete(value);
  mselUpdateLabel(key);
  filterScreener();
}

function mselSelectAll(key) {
  const cfg = mselRegistry[key];
  if (!cfg) return;
  const searchEl = document.getElementById(key + 'MselSearch');
  const q = (searchEl && searchEl.value || '').toLowerCase();
  let opts = cfg.options();
  if (q) opts = opts.filter(o => String(o.label).toLowerCase().includes(q));
  opts.forEach(o => cfg.selected.add(o.value));
  mselRenderList(key);
  mselUpdateLabel(key);
  filterScreener();
}

function mselClear(key) {
  const cfg = mselRegistry[key];
  if (!cfg) return;
  const searchEl = document.getElementById(key + 'MselSearch');
  const q = (searchEl && searchEl.value || '').toLowerCase();
  if (q) {
    let opts = cfg.options().filter(o => String(o.label).toLowerCase().includes(q));
    opts.forEach(o => cfg.selected.delete(o.value));
  } else {
    cfg.selected.clear();
  }
  mselRenderList(key);
  mselUpdateLabel(key);
  filterScreener();
}

let mselOpenedAt = 0;
function mselTogglePanel(key, e) {
  e.stopPropagation();
  const panel = document.getElementById(key + 'MselPanel');
  const btn = document.getElementById(key + 'MselBtn');
  if (!panel || !btn) return;
  const opening = !panel.classList.contains('open');
  // close all other panels first
  Object.keys(mselRegistry).forEach(k => {
    if (k !== key) {
      document.getElementById(k + 'MselPanel')?.classList.remove('open');
      document.getElementById(k + 'MselBtn')?.classList.remove('open');
    }
  });
  panel.classList.toggle('open', opening);
  btn.classList.toggle('open', opening);
  if (opening) {
    // Panels with a search box (index/sector) autofocus it below, which pops
    // the mobile keyboard and makes the browser fire a scroll/resize event as
    // the viewport adjusts. That event would otherwise hit the listeners below
    // and close this panel the instant it opens. Stamp the open time so those
    // listeners can ignore scroll/resize for a brief moment right after opening.
    mselOpenedAt = Date.now();
    const search = document.getElementById(key + 'MselSearch');
    if (search) search.value = '';
    mselRenderList(key);
    positionMselPanel(key);
    if (search) setTimeout(() => search.focus(), 0);
  }
}

function positionMselPanel(key) {
  // The filter panels live inside .table-card, which has overflow:hidden
  // for rounded corners. If the table shrinks (e.g. a filter returns 0 rows),
  // a position:absolute panel extending below the filter bar gets clipped.
  // Using position:fixed (computed from the button's real screen position)
  // escapes that clipping entirely, regardless of table height.
  const panel = document.getElementById(key + 'MselPanel');
  const btn = document.getElementById(key + 'MselBtn');
  if (!panel || !btn) return;
  const rect = btn.getBoundingClientRect();
  const panelWidth = panel.offsetWidth || 240;
  // Prefer aligning with the button's own left edge, like a normal dropdown.
  // Only shift it left if that would push the panel off the right edge of
  // the screen — and never push it further left than the button itself
  // needs to go to stay on-screen.
  let left = rect.left;
  if (left + panelWidth > window.innerWidth - 8) {
    left = window.innerWidth - panelWidth - 8;
  }
  left = Math.max(8, left);
  const top = rect.bottom + 4;
  panel.style.position = 'fixed';
  panel.style.top = top + 'px';
  panel.style.left = left + 'px';
  panel.style.right = 'auto';
}

document.addEventListener('click', function(e) {
  Object.keys(mselRegistry).forEach(key => {
    const root = document.getElementById(key + 'Msel');
    if (!root) return;
    if (!root.contains(e.target)) {
      document.getElementById(key + 'MselPanel')?.classList.remove('open');
      document.getElementById(key + 'MselBtn')?.classList.remove('open');
    }
  });
});
document.addEventListener('keydown', function(e) {
  if (e.key === 'Escape') {
    Object.keys(mselRegistry).forEach(key => {
      document.getElementById(key + 'MselPanel')?.classList.remove('open');
      document.getElementById(key + 'MselBtn')?.classList.remove('open');
    });
  }
});
// Close any open filter panel on scroll/resize rather than letting it drift
// out of place now that it's positioned relative to the viewport.
// Exception: index/sector panels have long scrollable checkbox lists for
// multi-selecting several items (common on mobile), so scrolling inside them
// must NOT close them. They still close via outside click, the toggle
// button, or Escape — just not from scroll/resize.
const NO_SCROLL_CLOSE = new Set(['index', 'sector', 'ticker']);
window.addEventListener('scroll', function() {
  if (Date.now() - mselOpenedAt < 400) return;
  Object.keys(mselRegistry).forEach(key => {
    if (NO_SCROLL_CLOSE.has(key)) return;
    document.getElementById(key + 'MselPanel')?.classList.remove('open');
    document.getElementById(key + 'MselBtn')?.classList.remove('open');
  });
}, true);
window.addEventListener('resize', function() {
  if (Date.now() - mselOpenedAt < 400) return;
  Object.keys(mselRegistry).forEach(key => {
    if (NO_SCROLL_CLOSE.has(key)) return;
    document.getElementById(key + 'MselPanel')?.classList.remove('open');
    document.getElementById(key + 'MselBtn')?.classList.remove('open');
  });
});

function resetMselFilters() {
  Object.keys(mselRegistry).forEach(key => {
    mselRegistry[key].selected.clear();
    mselUpdateLabel(key);
  });
}

function filterScreener() {
  const q = document.getElementById('screenerSearch').value.toLowerCase();
  const selLiquid = mselRegistry.liquid.selected;
  const selVolPhase = mselRegistry.volPhase.selected;
  const selSectors = mselRegistry.sector.selected;
  const selIndices = mselRegistry.index.selected;
  const selTickers = mselRegistry.ticker.selected;
  const selScores = mselRegistry.score.selected;
  const selStatuses = mselRegistry.status.selected;
  const selExtra = mselRegistry.extra.selected;
  const selOthers = mselRegistry.others.selected;

  const matchesExtra = (d, v) => {
    const mcap = parseFloat(d['Market Cap']) || 0;
    switch (v) {
      case 'mcap_gt_1b':   return mcap > 1000000000;
      case 'mcap_gt_500m': return mcap > 500000000;
      case 'mcap_gt_250m': return mcap > 250000000;
      case 'mcap_lt_250m': return mcap < 250000000;
      default: return false;
    }
  };

  const matchesOthers = (d, v) => {
    switch (v) {
      case 'turnaround': return (parseFloat(d['Loss narrow'])||0) > 0;
      case 'net_gt_op':  return (parseFloat(d['NI > OI'])||0) > 0;
      default: return false;
    }
  };

  filteredScreener = screenerData.filter(d => {
    if (q && !String(d.Ticker||'').toLowerCase().includes(q) && !String(d.Name||'').toLowerCase().includes(q)) return false;
    if (selExtra.size > 0) {
      let matched = false;
      for (const v of selExtra) { if (matchesExtra(d, v)) { matched = true; break; } }
      if (!matched) return false;
    }
    if (selOthers.size > 0) {
      let matched = false;
      for (const v of selOthers) { if (matchesOthers(d, v)) { matched = true; break; } }
      if (!matched) return false;
    }
    if (selSectors.size > 0 && !selSectors.has(d.Sector)) return false;
    if (selTickers.size > 0 && !selTickers.has(String(d.Ticker))) return false;
    if (selIndices.size > 0) {
      const dIdx = String(d.Index||'');
      let matched = false;
      for (const v of selIndices) { if (dIdx.includes(v)) { matched = true; break; } }
      if (!matched) return false;
    }
    if (selScores.size > 0) {
      const score = d['total improvement'] || 0;
      let matched = false;
      for (const v of selScores) {
        const [min, max] = v.split(',').map(Number);
        if (score > min && score <= max) { matched = true; break; }
      }
      if (!matched) return false;
    }
    if (selLiquid.size > 0) {
      let matched = false;
      for (const v of selLiquid) { if (d['Liquid Stock'] === Number(v)) { matched = true; break; } }
      if (!matched) return false;
    }
    if (selVolPhase.size > 0) {
      let matched = false;
      for (const v of selVolPhase) { if (Number(d['Accumulation']) === Number(v)) { matched = true; break; } }
      if (!matched) return false;
    }
    if (selStatuses.size > 0 && !selStatuses.has(String(sigStatusCode(d['Signal Status'])))) return false;
    return true;
  });

  // Sort
  const sortKeys = ['Ticker','Name','Sector','Latest EPS  Q','Latest TTM EPS Q','Revenue - Q','Op Income-Q','Net Income -Q','ROE 2026-Q1','Debt/Equity 2026-Q1','CFO 2026-Q1','Latest Div Y Q','P/E Ratio','Market Cap','total improvement','Signal date','Signal Price','Signal Return %','Signal Status','Price','Relative Vol','Volume','Day Change %','Current Week Return %','Current Month Return %','Past 3 Months Return %','YTD Return %'];
  const key = sortKeys[screenerSort.col];
  filteredScreener.sort((a,b) => {
  const av = a[key];
  const bv = b[key];

  if (av == null) return 1;
  if (bv == null) return -1;

  if (key === 'Signal Status') {
    return ((sigStatusCode(av) ?? -Infinity) - (sigStatusCode(bv) ?? -Infinity)) * screenerSort.dir;
  }

  // columns that must be numeric
  if (key === 'Signal date' || key === 'Signal Price' || key === 'Signal Return %' || key === 'Latest EPS  Q' || key === 'Latest TTM EPS Q' || key === 'Revenue - Q' || key === 'ROE 2026-Q1' || key === 'Debt/Equity 2026-Q1' || key === 'CFO 2026-Q1' || key === 'Latest Div Y Q' || key === 'P/E Ratio' || key === 'Market Cap' || key === 'total improvement' || key === 'Price' || key === 'Relative Vol' || key === 'Relative Volume' || key === 'Rel Vol' || key === 'Volume' || key === 'Day Change %' || key === 'Current Week Return %' || key === 'Current Month Return %' || key === 'Past 3 Months Return %' || key === 'YTD Return %') {
    return (Number(av) - Number(bv)) * screenerSort.dir;
  }

  // default string sort
  return String(av).localeCompare(String(bv)) * screenerSort.dir;
});


  screenerPage = 1;
  document.getElementById('screenerCount').textContent = `${filteredScreener.length} companies`;
  updateSortArrows('screenerTableHead', screenerSort.col, screenerSort.dir);
  renderScreenerPage();
  alignScreenerToggles();
}

function alignScreenerToggles() {
  const toggles = document.querySelector('.table-toggles');
  if (!toggles) return;
  if (window.innerWidth <= 768) {
    toggles.style.paddingLeft = '16px';
    return;
  }
  const target = document.getElementById('indexMsel');
  const header = document.getElementById('screenerSearch') ? document.getElementById('screenerSearch').closest('.table-header') : null;
  if (!target || !header) return;
  const headerRect = header.getBoundingClientRect();
  const targetRect = target.getBoundingClientRect();
  if (headerRect.width === 0) return; // tab hidden, skip
  const offset = Math.max(16, Math.round(targetRect.left - headerRect.left));
  toggles.style.paddingLeft = offset + 'px';
}
window.addEventListener('resize', () => { if (typeof alignScreenerToggles === 'function') alignScreenerToggles(); });

function sortScreener(col) {
  if (screenerSort.col === col) screenerSort.dir *= -1; else { screenerSort.col = col; screenerSort.dir = -1; }
  filterScreener();
}
function toggleFinancials() {
  const table = document.getElementById('screenerTable');
  const btn = document.getElementById('finToggleBtn');
  const hidden = table.classList.toggle('hide-financials');
  btn.classList.toggle('active', !hidden);
  document.getElementById('finToggleLabel').textContent = hidden ? 'Show Financials' : 'Hide Financials';
}

function toggleTechnical() {
  const table = document.getElementById('screenerTable');
  const btn = document.getElementById('techToggleBtn');
  const hidden = table.classList.toggle('hide-technical');
  btn.classList.toggle('active', !hidden);
  document.getElementById('techToggleLabel').textContent = hidden ? 'Show Technical' : 'Hide Technical';
}

function toggleDaily() {
  const table = document.getElementById('screenerTable');
  const btn = document.getElementById('dailyToggleBtn');
  const hidden = table.classList.toggle('hide-daily');
  btn.classList.toggle('active', !hidden);
  document.getElementById('dailyToggleLabel').textContent = hidden ? 'Show Daily Movement' : 'Hide Daily Movement';
}

// ===== WATCHLIST — three modes: localStorage (guest) + Firestore (signed-in account) + GitHub Gist (legacy sync option) =====
const WL_GIST_FILENAME = 'psx_watchlist.json';
const WL_LOCAL_KEY     = 'psx_watchlist_local';
let wlToken    = null;
let wlGistId   = null;
let wlUsername = null;
let wlList     = [];
let wlSaving   = false;
let wlMode     = 'local'; // 'local' | 'firestore' | 'github'
let wlUID      = null;
let wlEmail    = null;
// Resolves once the post-sign-in Firestore sync (wlOnSignIn) has finished, so
// that an add/remove/clear right after signing in doesn't race ahead of it
// and silently save to local storage instead of the user's account. Starts
// pre-resolved (guest state needs no waiting); reset to a fresh in-flight
// promise each time a sign-in sync begins.
let wlReadyPromise = Promise.resolve();

function wlLocalGet(k)    { try { return localStorage.getItem(k); }    catch { return null; } }
function wlLocalSet(k, v) { try { localStorage.setItem(k, v); }        catch {} }
function wlLocalDel(k)    { try { localStorage.removeItem(k); }        catch {} }

function wlReadLocal()       { try { return JSON.parse(wlLocalGet(WL_LOCAL_KEY) || '[]'); } catch { return []; } }
function wlWriteLocal(list)  { wlLocalSet(WL_LOCAL_KEY, JSON.stringify(list)); }

async function initWatchlist() {
  wlToken  = wlLocalGet('psx_gh_token');
  wlGistId = wlLocalGet('psx_gh_gist');
  if (wlToken) {
    try {
      const r = await fetch('https://api.github.com/user', {
        headers: { Authorization: `token ${wlToken}`, Accept: 'application/vnd.github+json' }
      });
      if (!r.ok) throw new Error('invalid');
      const u = await r.json();
      wlUsername = u.login;
      wlMode = 'github';
    } catch {
      wlToken = null; wlGistId = null; wlUsername = null;
      wlLocalDel('psx_gh_token'); wlLocalDel('psx_gh_gist');
      wlMode = 'local';
    }
  } else {
    wlMode = 'local';
  }
  if (wlMode === 'local') wlList = wlReadLocal();
  updateWatchlistBadge();
}

// Called from the Firebase Auth module script when a user signs in/out.
// Firestore takes over as the sync mode automatically — no token, no setup
// step. If the user already had a local/guest watchlist (e.g. items added
// before signing in), it's uploaded once so nothing is lost.
async function wlSaveFirestore() {
  if (!wlUID || typeof window.fsSaveWatchlist !== 'function') return false;
  return await window.fsSaveWatchlist(wlUID, wlList);
}

window.wlOnSignIn = function (uid, email) {
  wlReadyPromise = (async () => {
    wlUID = uid; wlEmail = email || '';
    if (typeof window.fsLoadWatchlist !== 'function') return; // Firestore bridge not ready yet
    const remote = await window.fsLoadWatchlist(uid);
    if (remote === undefined) {
      // Firestore read failed (offline, rules misconfigured, etc.) — fall back
      // to local storage rather than silently showing an empty watchlist.
      wlMode = 'local';
      wlList = wlReadLocal();
    } else if (remote === null) {
      // No Firestore watchlist yet for this account — migrate any existing
      // local/guest watchlist up so first-time sign-ins don't lose items.
      const localList = wlReadLocal();
      wlList = localList;
      wlMode = 'firestore';
      if (localList.length) await wlSaveFirestore();
    } else {
      wlList = remote;
      wlMode = 'firestore';
    }
    updateWatchlistBadge();
    if (document.getElementById('tab-watchlist')?.classList.contains('active')) showWatchlistPanel();
    // If the "only my watchlist" alert filter is on, the very first alert check
    // (fired synchronously on sign-in, before this async load finished) may
    // have used a stale/empty watchlist. Re-check now that it's loaded. Skipped
    // when the filter is off, since the result wouldn't change and would just
    // re-fire a duplicate toast/notification.
    if (document.getElementById('alertsWatchlistOnlyToggle')?.checked) {
      window.checkFreshSignalsToday?.();
    }
  })();
  return wlReadyPromise;
};

window.wlOnSignOut = function () {
  wlUID = null; wlEmail = null;
  wlMode = 'local';
  wlList = wlReadLocal();
  wlReadyPromise = Promise.resolve(); // nothing to wait for once signed out
  updateWatchlistBadge();
};

// Bridge for the Alerts panel (defined in the separate Firebase Auth module
// script, which can't see this script's `let`/`const` variables directly).
window.getWatchlistTickers = function () { return wlList.slice(); };

function showWatchlistPanel() {
  const setup   = document.getElementById('wlSetupPanel');
  const github  = document.getElementById('wlConnectedPanel');
  const guest   = document.getElementById('wlGuestPanel');
  const syncNote = document.getElementById('wlSyncNote');
  if (wlMode === 'github') {
    setup.style.display  = 'none';
    github.style.display = 'block';
    guest.style.display  = 'none';
    document.getElementById('wlUserLabel').textContent = '@' + (wlUsername || '');
    loadWatchlistFromGist();
    return;
  }
  setup.style.display  = 'none';
  github.style.display = 'none';
  guest.style.display  = 'block';
  if (wlMode === 'local') wlList = wlReadLocal();
  updateWatchlistBadge();
  renderWatchlist('guest');
  if (syncNote) {
    if (wlMode === 'firestore') {
      syncNote.innerHTML = `✓ Synced to your account${wlEmail ? ' (' + wlEmail + ')' : ''}`;
    } else {
      syncNote.innerHTML = 'Saved in this browser only — <span onclick="wlShowConnect()" style="color:var(--accent3); cursor:pointer; text-decoration:underline;">Connect GitHub to sync across devices →</span>';
    }
  }
}

function wlShowConnect() {
  document.getElementById('wlGuestPanel').style.display = 'none';
  document.getElementById('wlSetupPanel').style.display = 'block';
}

async function connectGitHub() {
  const token = document.getElementById('wlTokenInput').value.trim();
  const errEl = document.getElementById('wlTokenError');
  const btn   = document.getElementById('wlConnectBtn');
  errEl.style.display = 'none';
  if (!token) { errEl.textContent = 'Please paste your token first.'; errEl.style.display = 'block'; return; }
  btn.textContent = 'Connecting…'; btn.disabled = true;
  try {
    const ur = await fetch('https://api.github.com/user', {
      headers: { Authorization: `token ${token}`, Accept: 'application/vnd.github+json' }
    });
    if (!ur.ok) throw new Error('Invalid token — check it was copied correctly.');
    const user   = await ur.json();
    const scopes = ur.headers.get('x-oauth-scopes') || '';
    if (!scopes.includes('gist')) throw new Error('Token is missing "gist" scope. Please regenerate with gist access.');
    wlToken = token; wlUsername = user.login; wlMode = 'github';
    wlLocalSet('psx_gh_token', token);
    await findOrCreateGist();
    // Migrate local stocks into Gist
    const local = wlReadLocal();
    if (local.length > 0) {
      await loadWatchlistFromGist(true);
      wlList = [...new Set([...wlList, ...local])];
      await saveWatchlistToGist();
      wlWriteLocal([]);
      showToast(`✓ ${local.length} local stock(s) migrated to GitHub`);
    }
    document.getElementById('wlTokenInput').value = '';
    showWatchlistPanel();
  } catch(e) {
    errEl.textContent = '❌ ' + e.message;
    errEl.style.display = 'block';
  }
  btn.textContent = 'Connect to GitHub'; btn.disabled = false;
}

async function findOrCreateGist() {
  if (wlGistId) {
    try {
      const r = await fetch(`https://api.github.com/gists/${wlGistId}`, { headers: { Authorization: `token ${wlToken}` } });
      if (r.ok) return;
    } catch {}
    wlGistId = null; wlLocalDel('psx_gh_gist');
  }
  const r     = await fetch('https://api.github.com/gists?per_page=100', { headers: { Authorization: `token ${wlToken}`, Accept: 'application/vnd.github+json' } });
  const gists = await r.json();
  const found = gists.find(g => g.files && g.files[WL_GIST_FILENAME]);
  if (found) {
    wlGistId = found.id;
  } else {
    const cr = await fetch('https://api.github.com/gists', {
      method: 'POST',
      headers: { Authorization: `token ${wlToken}`, Accept: 'application/vnd.github+json', 'Content-Type': 'application/json' },
      body: JSON.stringify({ description: 'Nexus PSX Watchlist', public: false, files: { [WL_GIST_FILENAME]: { content: '[]' } } })
    });
    const g = await cr.json();
    wlGistId = g.id;
  }
  wlLocalSet('psx_gh_gist', wlGistId);
}

async function loadWatchlistFromGist(silent = false) {
  if (!silent) {
    document.getElementById('wlLoading').style.display        = 'block';
    document.getElementById('watchlistEmptyGH').style.display = 'none';
    document.getElementById('watchlistTableGH').style.display = 'none';
    setSyncStatus('⏳ Syncing…', '#f5a623');
  }
  try {
    if (!wlGistId) await findOrCreateGist();
    const r = await fetch(`https://api.github.com/gists/${wlGistId}`, {
      headers: { Authorization: `token ${wlToken}`, Accept: 'application/vnd.github+json' }, cache: 'no-store'
    });
    if (!r.ok) throw new Error('Could not load gist');
    const g = await r.json();
    wlList = JSON.parse(g.files[WL_GIST_FILENAME]?.content || '[]');
    if (!silent) { setSyncStatus('✓ Synced', 'var(--positive)'); updateWatchlistBadge(); renderWatchlist('github'); }
  } catch(e) {
    if (!silent) { setSyncStatus('⚠ Sync failed', 'var(--danger)'); showToast('Could not load: ' + e.message); }
  }
  if (!silent) document.getElementById('wlLoading').style.display = 'none';
}

async function saveWatchlistToGist() {
  if (!wlToken || !wlGistId || wlSaving) return;
  wlSaving = true;
  setSyncStatus('⏳ Saving…', '#f5a623');
  try {
    const r = await fetch(`https://api.github.com/gists/${wlGistId}`, {
      method: 'PATCH',
      headers: { Authorization: `token ${wlToken}`, Accept: 'application/vnd.github+json', 'Content-Type': 'application/json' },
      body: JSON.stringify({ files: { [WL_GIST_FILENAME]: { content: JSON.stringify(wlList) } } })
    });
    if (!r.ok) throw new Error('Save failed');
    setSyncStatus('✓ Saved', 'var(--positive)');
  } catch(e) {
    setSyncStatus('⚠ Save failed', 'var(--danger)'); showToast('Could not save: ' + e.message);
  }
  wlSaving = false;
}

function setSyncStatus(msg, color) {
  const el = document.getElementById('wlSyncStatus');
  if (el) { el.textContent = msg; el.style.color = color; }
}

function disconnectGitHub() {
  if (!confirm('Disconnect from GitHub? Your watchlist stays in the Gist for next time.')) return;
  wlWriteLocal(wlList); // copy to local so nothing is lost
  wlToken = null; wlGistId = null; wlUsername = null;
  wlLocalDel('psx_gh_token'); wlLocalDel('psx_gh_gist');
  wlMode = 'local';
  updateWatchlistBadge();
  showWatchlistPanel();
  showToast('Disconnected — watchlist copied to browser storage');
}

async function addToWatchlist(ticker) {
  // Wait for any in-flight post-sign-in Firestore sync to finish first, so a
  // tap right after signing in doesn't race ahead of it (wlMode would still
  // read 'local' for a brief moment otherwise, silently saving to this
  // browser only instead of the user's account).
  await wlReadyPromise;
  ticker = String(ticker);
  if (wlList.includes(ticker)) { showToast(`${ticker} already in watchlist`); return; }
  wlList.push(ticker);
  updateWatchlistBadge();
  showToast(`⭐ ${ticker} added to watchlist`);
  if (wlMode === 'github') { renderWatchlist('github'); await saveWatchlistToGist(); }
  else if (wlMode === 'firestore') { renderWatchlist('guest'); await wlSaveFirestore(); }
  else { wlWriteLocal(wlList); renderWatchlist('guest'); }
}

async function removeFromWatchlist(ticker) {
  await wlReadyPromise;
  wlList = wlList.filter(t => t !== String(ticker));
  updateWatchlistBadge();
  if (wlMode === 'github') { renderWatchlist('github'); await saveWatchlistToGist(); }
  else if (wlMode === 'firestore') { renderWatchlist('guest'); await wlSaveFirestore(); }
  else { wlWriteLocal(wlList); renderWatchlist('guest'); }
}

async function clearWatchlist() {
  await wlReadyPromise;
  if (!wlList.length) return;
  if (!confirm('Clear your entire watchlist?')) return;
  wlList = [];
  updateWatchlistBadge();
  if (wlMode === 'github') { renderWatchlist('github'); await saveWatchlistToGist(); }
  else if (wlMode === 'firestore') { renderWatchlist('guest'); await wlSaveFirestore(); }
  else { wlWriteLocal(wlList); renderWatchlist('guest'); }
}

function updateWatchlistBadge() {
  const badge = document.getElementById('watchlistCountBadge');
  if (badge) {
    badge.style.display = wlList.length > 0 ? 'inline' : 'none';
    badge.textContent   = wlList.length;
  }
  // sync mobile nav watchlist badge
  const mbadge = document.getElementById('mnavWlBadge');
  if (mbadge) {
    mbadge.style.display = wlList.length > 0 ? 'inline' : 'none';
    mbadge.textContent   = wlList.length;
  }
}

// Watchlist sort state: { col: 0-6, dir: 1 | -1 }
const wlSort = { col: null, dir: 1 };

function sortWatchlist(col, mode) {
  if (wlSort.col === col) wlSort.dir *= -1;
  else { wlSort.col = col; wlSort.dir = 1; }
  renderWatchlist(mode);
}

function renderWatchlist(mode) {
  const emptyId = mode === 'github' ? 'watchlistEmptyGH'    : 'watchlistEmptyGuest';
  const tableId = mode === 'github' ? 'watchlistTableGH'    : 'watchlistTableGuest';
  const tbodyId = mode === 'github' ? 'watchlistBodyGH'     : 'watchlistBodyGuest';
  const headId  = mode === 'github' ? 'wlHeadGH'            : 'wlHeadGuest';
  const empty = document.getElementById(emptyId);
  const table = document.getElementById(tableId);
  const tbody = document.getElementById(tbodyId);
  if (!tbody) return;
  if (wlList.length === 0) { empty.style.display = 'block'; table.style.display = 'none'; tbody.innerHTML = ''; return; }
  empty.style.display = 'none'; table.style.display = '';

  // Update sort indicators on headers
  const headRow = document.getElementById(headId);
  if (headRow) {
    headRow.querySelectorAll('th').forEach((th, i) => {
      th.classList.remove('sort-asc','sort-desc');
      if (i === wlSort.col) th.classList.add(wlSort.dir === 1 ? 'sort-asc' : 'sort-desc');
    });
  }

  // Build row objects for sorting
  let rows = wlList.map(ticker => {
    const d = SOURCE_DATA.find(r => String(r.Ticker) === ticker);
    return { ticker, d };
  });

  // Sort
  if (wlSort.col !== null) {
    rows.sort((a, b) => {
      const da = a.d, db = b.d;
      let va, vb;
      switch (wlSort.col) {
        case 0: va = a.ticker; vb = b.ticker; break;
        case 1: va = da?.Name||''; vb = db?.Name||''; break;
        case 2: va = da?.Sector||''; vb = db?.Sector||''; break;
        case 3: va = da ? (da['total improvement'] ?? -Infinity) : -Infinity;
                vb = db ? (db['total improvement'] ?? -Infinity) : -Infinity; break;
        case 4: va = da ? (dget(da,'Signal date')||'') : '';
                vb = db ? (dget(db,'Signal date')||'') : ''; break;
        case 5: va = da ? (toNum(dget(da,'Signal Price')) ?? -Infinity) : -Infinity;
                vb = db ? (toNum(dget(db,'Signal Price')) ?? -Infinity) : -Infinity; break;
        case 6: va = da ? (sigStatusCode(dget(da,'Signal Status')) ?? -Infinity) : -Infinity;
                vb = db ? (sigStatusCode(dget(db,'Signal Status')) ?? -Infinity) : -Infinity; break;
        case 7: va = da ? (toNum(dget(da,'Signal Return %')) ?? -Infinity) : -Infinity;
                vb = db ? (toNum(dget(db,'Signal Return %')) ?? -Infinity) : -Infinity; break;
        default: return 0;
      }
      if (va < vb) return -wlSort.dir;
      if (va > vb) return wlSort.dir;
      return 0;
    });
  }

  tbody.innerHTML = '';
  rows.forEach(({ ticker, d }) => {
    const tr = document.createElement('tr');
    if (!d) {
      tr.innerHTML = `<td class="ticker-link" onclick="switchTab('company');pickTicker('${ticker}')">${ticker}</td><td colspan="8" style="color:var(--text2);font-size:12px;">Data not available</td><td><button onclick="removeFromWatchlist('${ticker}')" style="background:rgba(255,80,80,0.1);border:1px solid rgba(255,80,80,0.25);border-radius:4px;color:#ff6b6b;width:26px;height:26px;cursor:pointer;font-size:14px;line-height:1;padding:0;">✕</button></td>`;
    } else {
      const score  = d['total improvement'];
      const ret    = toNum(dget(d, 'Signal Return %'));
      const retStr = ret != null ? (ret >= 0 ? '+' : '') + ret.toFixed(1) + '%' : '—';
      const sigRaw = dget(d, 'Signal Status');
      const sig    = sigStatusLabel(sigRaw);
      const sigPill = sig != null ? sigStatusPillClass(sigRaw) : '';
      const sigHtml = sig != null
        ? `<span class="pill ${sigPill}" style="font-size:10px;padding:2px 7px;text-transform:none;">${sig}</span>`
        : '<span style="color:var(--text3)">—</span>';
      tr.innerHTML = `
        <td class="ticker-link" onclick="switchTab('company');pickTicker('${ticker}')">${ticker}</td>
        <td style="max-width:180px;overflow:hidden;text-overflow:ellipsis">${d.Name||'—'}</td>
        <td class="screener-hide-mobile" style="max-width:140px;overflow:hidden;text-overflow:ellipsis;color:var(--text2)">${d.Sector||'—'}</td>
        <td class="mono"><span class="pill ${score>=80?'pill-good':score>=50?'pill-neutral':'pill-bad'}">${score!=null?score:'—'}</span></td>
        <td class="mono">${fmtSignalDate(dget(d,'Signal date'))}</td>
        <td class="mono">${(()=>{const n=toNum(dget(d,'Signal Price'));return n!=null?n.toFixed(2):'—';})()}</td>
        <td class="mono">${sigHtml}</td>
        <td class="mono ${valColor(ret)}">${retStr}</td>
        <td><button onclick="removeFromWatchlist('${ticker}')" title="Remove" style="background:rgba(255,80,80,0.1);border:1px solid rgba(255,80,80,0.25);border-radius:4px;color:#ff6b6b;width:26px;height:26px;cursor:pointer;font-size:14px;line-height:1;padding:0;">✕</button></td>
      `;
    }
    tbody.appendChild(tr);
  });
}

function showToast(msg) {
  let t = document.getElementById('wlToast');
  if (!t) {
    t = document.createElement('div'); t.id = 'wlToast';
    t.style.cssText = 'position:fixed;bottom:28px;left:50%;transform:translateX(-50%);background:#1e2d45;border:1px solid var(--accent3);color:var(--text);padding:9px 20px;border-radius:8px;font-size:13px;font-family:"Inter",sans-serif;z-index:9999;transition:opacity 0.3s;pointer-events:none;white-space:nowrap;box-shadow:0 4px 16px rgba(0,0,0,0.4)';
    document.body.appendChild(t);
  }
  t.textContent = msg; t.style.opacity = '1';
  clearTimeout(t._timeout);
  t._timeout = setTimeout(() => { t.style.opacity = '0'; }, 2500);
}
function tickerBadges(d) {
  let badges = '';
  if (String(d.Index||'').includes('KMI')) {
    badges += '<span class="ticker-badge ticker-badge-kmi" title="KMI Index member"><svg viewBox="0 0 24 24" width="13" height="13" fill="currentColor"><path d="M12 2c-1.1 0-2 .9-2 2 0 .74.4 1.38 1 1.72V7H9a3 3 0 0 0-3 3v1H4v9h16v-9h-2v-1a3 3 0 0 0-3-3h-2V5.72c.6-.34 1-.98 1-1.72 0-1.1-.9-2-2-2zM7 11a1 1 0 0 1 1-1h8a1 1 0 0 1 1 1v1H7v-1zM5 13h2v6H5v-6zm12 0h2v6h-2v-6zM9 14h6v5H9v-5z"></path></svg></span>';
  }
  if ((parseFloat(d['NI > OI'])||0) > 0) {
    badges += '<span class="ticker-badge ticker-badge-ni" title="Net Inc > Op Inc">N</span>';
  }
  if ((parseFloat(d['Loss narrow'])||0) > 0) {
    badges += '<span class="ticker-badge ticker-badge-turn" title="Turnaround Candidate">T</span>';
  }
  return badges;
}

function renderScreenerPage() {
  const start = (screenerPage - 1) * PAGE_SIZE;
  const page = filteredScreener.slice(start, start + PAGE_SIZE);
  const tbody = document.getElementById('screenerBody');
  tbody.innerHTML = '';
  tbody.classList.remove('pre-render');
  page.forEach(d => {
    const score = d['total improvement'];
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="ticker-link" onclick="switchTab('company');pickTicker('${String(d.Ticker)}')">${d.Ticker}${tickerBadges(d)}</td>
      <td style="max-width:180px; overflow:hidden; text-overflow:ellipsis; color:var(--text); font-weight:500">${d.Name||'—'}</td>
      <td class="screener-sector-cell" style="max-width:140px; overflow:hidden; text-overflow:ellipsis">${d.Sector||'—'}</td>
      <td class="screener-fin-col mono ${valColor(dget(d,'Latest EPS  Q'))}">${fmt(dget(d,'Latest EPS  Q'),3)}</td>
      <td class="screener-hide-mobile screener-fin-col mono ${valColor(dget(d,'Latest TTM EPS Q'))}">${fmt(dget(d,'Latest TTM EPS Q'),3)}</td>
      <td class="screener-hide-mobile screener-fin-col mono">${fmtBig(dget(d,'Revenue - Q'))}</td>
      <td class="screener-hide-mobile screener-fin-col mono ${valColor(dget(d,'Op Income-Q'))}">${fmtPct(dget(d,'Op Income-Q'))}</td>
      <td class="screener-hide-mobile screener-fin-col mono ${valColor(dget(d,'Net Income -Q'))}">${fmtPct(dget(d,'Net Income -Q'))}</td>
      <td class="screener-fin-col mono ${valColor(dget(d,'ROE 2026-Q1'))}">${fmtPct(dget(d,'ROE 2026-Q1'))}</td>
      <td class="screener-hide-mobile screener-fin-col mono">${fmt(dget(d,'Debt/Equity 2026-Q1'),2)}</td>
      <td class="screener-hide-mobile screener-fin-col mono ${valColor(dget(d,'CFO 2026-Q1'))}">${fmtBig(dget(d,'CFO 2026-Q1'))}</td>
      <td class="screener-fin-col mono ${valColor(dget(d,'Latest Div Y Q'))}">${fmtPct(dget(d,'Latest Div Y Q'))}</td>
      <td class="screener-fin-col mono">${fmt(dget(d,'P/E Ratio'),2)}</td>
      <td class="screener-fin-col mono">${fmtMarketCap(dget(d,'Market Cap'))}</td>
      <td class="mono"><span class="pill ${score>=80?'pill-good':score>=50?'pill-neutral':'pill-bad'}">${score!=null?score:'—'}</span></td>
      <td class="mono screener-tech-col">${fmtSignalDate(dget(d,'Signal date'))}</td>   
      <td class="mono screener-tech-col">${(()=>{const n=toNum(dget(d,'Signal Price'));return n!=null?n.toFixed(2):'—';})()}</td>
      <td class="mono screener-tech-col ${valColor(dget(d,'Signal Return %'))}">${(()=>{const n=toNum(dget(d,'Signal Return %'));return n!=null?(n>=0?'+':'')+n.toFixed(1)+'%':'—'})()}</td>
      <td class="mono screener-tech-col">${(()=>{const raw=dget(d,'Signal Status');const s=sigStatusLabel(raw);if(s==null)return '—';const pc=sigStatusPillClass(raw);return `<span class="pill ${pc}" style="font-size:10px;padding:2px 7px;text-transform:none;">${s}</span>`;})()}</td>
      <td class="mono screener-daily-col">${(()=>{const n=toNum(dget(d,'Price'));return n!=null?n.toFixed(2):'—';})()}</td>
      <td class="mono screener-daily-col ${(()=>{const n=toNum(dget(d,'Relative Vol')??dget(d,'Rel Vol'));return n==null?'':n>1.5?'positive':n<0.5?'negative':'';})()}">${(()=>{const n=toNum(dget(d,'Relative Vol')??dget(d,'Rel Vol'));return n!=null?n.toFixed(2):'—';})()}</td>
      <td class="mono screener-daily-col ${(()=>{const n=toNum(dget(d,'Volume'));return n==null?'':n>0?'positive':n<0?'negative':'';})()}">${(()=>{const n=toNum(dget(d,'Volume'));return n!=null?Math.round(n).toLocaleString():'—';})()}</td>
      <td class="mono screener-daily-col ${(()=>{const n=toNum(dget(d,'Day Change %'));return n==null||n===0?'':n>0?'positive':'negative';})()}">${(()=>{const n=toNum(dget(d,'Day Change %'));return n!=null&&n!==0?(n>0?'+':'')+n.toFixed(2)+'%':'—';})()}</td>
      <td class="mono screener-perf-col ${(()=>{const n=toNum(dget(d,'Current Week Return %'));return n==null?'':n>0?'positive':'negative';})()}">${(()=>{const n=toNum(dget(d,'Current Week Return %'));return n!=null?(n>=0?'+':'')+n.toFixed(2)+'%':'—';})()}</td>
      <td class="mono screener-perf-col ${(()=>{const n=toNum(dget(d,'Current Month Return %'));return n==null?'':n>0?'positive':'negative';})()}">${(()=>{const n=toNum(dget(d,'Current Month Return %'));return n!=null?(n>=0?'+':'')+n.toFixed(2)+'%':'—';})()}</td>
      <td class="mono screener-perf-col ${(()=>{const n=toNum(dget(d,'Past 3 Months Return %'));return n==null?'':n>0?'positive':'negative';})()}">${(()=>{const n=toNum(dget(d,'Past 3 Months Return %'));return n!=null?(n>=0?'+':'')+n.toFixed(2)+'%':'—';})()}</td>
      <td class="mono screener-perf-col ${(()=>{const n=toNum(dget(d,'YTD Return %'));return n==null?'':n>0?'positive':'negative';})()}">${(()=>{const n=toNum(dget(d,'YTD Return %'));return n!=null?(n>=0?'+':'')+n.toFixed(2)+'%':'—';})()}</td>
      <td style="padding:4px 8px;"><button onclick="addToWatchlist('${String(d.Ticker)}')" title="Add to Watchlist" style="background:var(--accent3-dim); border:1px solid rgba(59,130,246,0.35); border-radius:4px; color:var(--accent3); width:26px; height:26px; font-size:16px; cursor:pointer; line-height:1; padding:0;">＋</button></td>
    `;
    tbody.appendChild(tr);
  });
  renderPagination();
}

function togglePerformance() {
  const tbl=document.getElementById('screenerTable');
  const btn=document.getElementById('perfToggleBtn');
  const hidden = tbl.classList.toggle('hide-performance');
  btn.classList.toggle('active', !hidden);
  document.getElementById('perfToggleLabel').textContent = hidden ? 'Show Performance' : 'Hide Performance';
}
function renderPagination() {
  const total = filteredScreener.length;
  const pages = Math.ceil(total / PAGE_SIZE);
  const pag = document.getElementById('screenerPagination');
  let html = '';
  const start = Math.max(1, screenerPage - 2);
  const end = Math.min(pages, screenerPage + 2);
  if (screenerPage > 1) html += `<button class="page-btn" onclick="goPage(${screenerPage-1})">←</button>`;
  if (start > 1) html += `<button class="page-btn" onclick="goPage(1)">1</button>`;
  for (let i = start; i <= end; i++) html += `<button class="page-btn ${i===screenerPage?'active':''}" onclick="goPage(${i})">${i}</button>`;
  if (end < pages) html += `<button class="page-btn" onclick="goPage(${pages})">${pages}</button>`;
  if (screenerPage < pages) html += `<button class="page-btn" onclick="goPage(${screenerPage+1})">→</button>`;
  html += `<span class="page-info">Showing ${(screenerPage-1)*PAGE_SIZE+1}–${Math.min(screenerPage*PAGE_SIZE,total)} of ${total}</span>`;
  pag.innerHTML = html;
}
function goPage(p) { screenerPage = p; renderScreenerPage(); }
function renderScreener() { filteredScreener = [...screenerData]; filterScreener(); }

// ===== TABS =====
function switchTab(name) {
  const tabNames = ['company','sector','screener','watchlist','top'];
  document.querySelectorAll('.tab').forEach((t,i) => t.classList.toggle('active', tabNames[i]===name));
  document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
  const el = document.getElementById('tab-'+name);
  if (el) el.classList.add('active');
  if (name === 'watchlist') showWatchlistPanel();
  if (name === 'top') buildTopTab();
  if (name === 'screener') setTimeout(alignScreenerToggles, 0);
  // sync mobile nav
  tabNames.forEach(t => {
    const btn = document.getElementById('mnav-'+t);
    if (btn) btn.classList.toggle('active', t === name);
  });
  // scroll to top on mobile tab switch
  if (window.innerWidth <= 768) window.scrollTo({top: 0, behavior: 'smooth'});
}

// ===== SAFE DOM HELPERS =====
function setEl(id, text) {
  const el = document.getElementById(id);
  if (el) el.textContent = text;
}
function setHTML(id, html) {
  const el = document.getElementById(id);
  if (el) el.innerHTML = html;
}
function setClass(id, cls) {
  const el = document.getElementById(id);
  if (el) el.className = cls;
}
function setStyle(id, prop, val) {
  const el = document.getElementById(id);
  if (el) el.style[prop] = val;
}

// ===== FORMATTING HELPERS =====
function toNum(v) {
  if (v == null) return null;
  const n = parseFloat(v);
  return isNaN(n) ? null : n;
}
function fmt(v, dec=2) {
  const n = toNum(v);
  if (n == null) return '—';
  return n.toFixed(dec);
}
// Signal date is stored as an 8-digit YYYYMMDD number (e.g. 20260511 → 11 May 2026).
// Shared by the Stock Screener table and Company View so both read it identically.
function fmtSignalDate(v) {
  if (v == null) return '—';
  const sd = String(v);
  if (sd.length !== 8) return sd || '—';
  const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const yr = sd.slice(0,4), mo = parseInt(sd.slice(4,6),10)-1, dy = sd.slice(6,8);
  if (mo < 0 || mo > 11) return sd;
  return `${parseInt(dy,10)} ${months[mo]} ${yr}`;
}
// Signal Status is now encoded as a numeric score in the source data.
// Some rows may still carry the older text labels — map those to the same codes too.
const SIGNAL_STATUS_MAP = {
  9: 'Buy call closed',
  4: 'Be cautious',
  3: 'Take some profit',
  2.5: 'Extended buy signal (cautious)',
  2: 'Continuation buy signal',
  1.5: 'Initial buy signal',
  1: 'Hold trade',
  0: 'No trade'
};
const LEGACY_SIGNAL_STATUS_TO_CODE = {
  'Buy call closed': 9,
  'Be cautious': 4,
  'Take some profit': 3,
  'Hold trade': 1,
  'No trade': 0,
  // code 1.5 — Initial buy signal
  'Initial buy signal': 1.5,
  // code 2 — has gone by several names; map them all
  'Fresh buy signal': 2,
  'Fresh signal in an initial buy call': 2,
  'Continuation buy signal': 2,
  // code 2.5 — cautious variant of the above
  'Fresh buy signal (cautious)': 2.5,
  'Fresh signal in an initial buy call (cautious)': 2.5,
  'Extended buy signal (cautious)': 2.5
};
function sigStatusCode(raw) {
  if (raw == null || raw === '') return null;
  if (typeof raw === 'number') return raw;
  const s = String(raw).trim();
  if (LEGACY_SIGNAL_STATUS_TO_CODE.hasOwnProperty(s)) return LEGACY_SIGNAL_STATUS_TO_CODE[s];
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}
function sigStatusLabel(code) {
  const v = sigStatusCode(code);
  if (v == null) return (code == null || code === '') ? null : String(code);
  return SIGNAL_STATUS_MAP.hasOwnProperty(v) ? SIGNAL_STATUS_MAP[v] : String(code);
}
function sigStatusPillClass(code) {
  const v = sigStatusCode(code);
  if (v === 1.5 || v === 2 || v === 1) return 'pill-good';   // Initial buy, Continuation buy, Hold trade
  if (v === 2.5 || v === 4 || v === 9) return 'pill-bad';     // Extended buy (cautious), Be cautious, Buy call closed
  return 'pill-neutral'; // 0 (No trade), 3 (Take some profit), and unrecognized values — orange
}
function fmtMarketCap(v) {
  const n = toNum(v);
  if (n == null) return '—';
  if (n >= 1e9) return Math.round(n/1e9).toLocaleString() + 'B';
  if (n >= 1e6) return Math.round(n/1e6).toLocaleString() + 'M';
  if (n >= 1e3) return Math.round(n/1e3).toLocaleString() + 'K';
  return Math.round(n).toLocaleString();
}
function fmtBig(v) {
  const n = toNum(v);
  if (n == null) return '—';
  const abs = Math.abs(n);
  if (abs >= 1e12) return (n/1e12).toFixed(2) + 'T';
  if (abs >= 1e9)  return (n/1e9).toFixed(2)  + 'B';
  if (abs >= 1e6)  return (n/1e6).toFixed(2)  + 'M';
  if (abs >= 1e3)  return (n/1e3).toFixed(1)  + 'K';
  return n.toFixed(0);
}
function fmtPct(v) {
  const n = toNum(v);
  if (n == null) return '—';
  return (n * 100).toFixed(1) + '%';
}
function valColor(v) {
  const n = toNum(v);
  if (n == null) return '';
  return n > 0 ? 'positive' : n < 0 ? 'negative' : '';
}
function scoreColor(s) {
  const n = toNum(s);
  if (n == null) return '';
  return n >= 80 ? 'score-high' : n >= 40 ? 'score-mid' : 'score-low';
}
function trendIcon(curr, prev) {
  const a = toNum(curr), b = toNum(prev);
  if (a == null || b == null) return '';
  return a > b ? '<span class="up">▲</span>' : a < b ? '<span class="down">▼</span>' : '→';
}

/* ===================== PWA INSTALL (Add to Home Screen) ===================== */
let deferredInstallPrompt = null;
const APP_ICON_DATAURI = document.querySelector('link[rel="apple-touch-icon"]') ? document.querySelector('link[rel="apple-touch-icon"]').href : '';

function isIOSDevice() {
  const ua = navigator.userAgent || '';
  const isIOS = /iPad|iPhone|iPod/.test(ua);
  const isIPadOS13Up = navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1;
  return isIOS || isIPadOS13Up;
}
function isStandaloneMode() {
  return window.matchMedia('(display-mode: standalone)').matches || window.navigator.standalone === true;
}
function isSafariBrowser() {
  const ua = navigator.userAgent || '';
  return /^((?!chrome|android|crios|fxios|edgios).)*safari/i.test(ua);
}

function showInstallButton() {
  const btn = document.getElementById('installAppBtn');
  if (btn && !isStandaloneMode()) btn.classList.add('show');
}
function hideInstallButton() {
  const btn = document.getElementById('installAppBtn');
  if (btn) btn.classList.remove('show');
}

window.addEventListener('beforeinstallprompt', (e) => {
  e.preventDefault();
  deferredInstallPrompt = e;
  showInstallButton();
});
window.addEventListener('appinstalled', () => {
  deferredInstallPrompt = null;
  hideInstallButton();
});

document.addEventListener('DOMContentLoaded', () => {
  if (isStandaloneMode()) { hideInstallButton(); return; }
  // iOS Safari never fires beforeinstallprompt — show our button so users
  // can get manual "Add to Home Screen" instructions.
  if (isIOSDevice() && isSafariBrowser()) showInstallButton();
  // Register service worker (required by Chrome/Edge/Android for installability)
  if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('./sw.js').catch(() => { /* sw.js not hosted alongside index.html — install button still works on iOS */ });
  }
});

function handleInstallClick() {
  const iconImg = document.getElementById('installIconPreview');
  if (iconImg) iconImg.src = APP_ICON_DATAURI;

  const title = document.getElementById('installModalTitle');
  const body = document.getElementById('installModalBody');
  const steps = document.getElementById('installSteps');
  const cancelBtn = document.getElementById('installCancelBtn');
  const confirmBtn = document.getElementById('installConfirmBtn');

  if (isIOSDevice()) {
    // iOS has no programmatic install API — guide the user manually.
    title.textContent = 'Add Nexus PSX to Home Screen';
    body.textContent = 'iOS doesn\'t allow apps to install themselves automatically — just follow these steps in Safari:';
    steps.style.display = 'block';
    steps.innerHTML = '1. Tap the <b>Share</b> icon (square with an arrow) in Safari\'s toolbar<br>'
      + '2. Scroll down and tap <b>"Add to Home Screen"</b><br>'
      + '3. Tap <b>"Add"</b> in the top-right corner<br><br>'
      + 'The Nexus PSX icon will then appear on your home screen.';
    cancelBtn.style.display = 'none';
    confirmBtn.textContent = 'Got it';
    confirmBtn.onclick = closeInstallModal;
  } else if (deferredInstallPrompt) {
    // Chrome / Edge / Android — native install prompt available after confirmation.
    title.textContent = 'Install Nexus PSX?';
    body.textContent = 'Add Nexus PSX to your home screen or desktop for quick, full-screen access — just like a native app.';
    steps.style.display = 'none';
    cancelBtn.style.display = '';
    confirmBtn.textContent = 'Install';
    confirmBtn.onclick = confirmInstall;
  } else {
    // Browser supports installing but hasn't fired the event yet, or doesn't support it at all.
    title.textContent = 'Install Nexus PSX';
    body.textContent = 'Use your browser\'s built-in install option to add Nexus PSX to your device:';
    steps.style.display = 'block';
    steps.innerHTML = '<b>Desktop Chrome/Edge:</b> click the install icon (⊕ or computer icon) in the address bar.<br><br>'
      + '<b>Android Chrome:</b> open the ⋮ menu → "Install app" or "Add to Home screen".';
    cancelBtn.style.display = 'none';
    confirmBtn.textContent = 'Got it';
    confirmBtn.onclick = closeInstallModal;
  }

  document.getElementById('installModal').classList.add('show');
}

function closeInstallModal() {
  document.getElementById('installModal').classList.remove('show');
}

async function confirmInstall() {
  closeInstallModal();
  if (!deferredInstallPrompt) return;
  deferredInstallPrompt.prompt();
  const choice = await deferredInstallPrompt.userChoice;
  if (choice && choice.outcome === 'accepted') {
    hideInstallButton();
  }
  deferredInstallPrompt = null;
}




function handleExcelUpload(event) {
  const file = event.target.files[0];
  if (!file) return;
  event.target.value = '';
  showOverlay('Reading Excel file...');

  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      showOverlay('Parsing SourceData sheet...');
      const wb = XLSX.read(e.target.result, {type: 'binary'});

      const sheetName = wb.SheetNames.find(n => n.toLowerCase() === 'sourcedata');
      if (!sheetName) {
        hideOverlay();
        showModalError('Sheet Not Found', `Could not find a sheet named "SourceData" in this file.`,
          `Sheets found: ${wb.SheetNames.join(', ')}`);
        return;
      }

      showOverlay('Processing ' + sheetName + '...');
      const ws = wb.Sheets[sheetName];

      // raw:true keeps numbers as numbers, dates as Date objects
      const rows = XLSX.utils.sheet_to_json(ws, {raw: true, defval: null});

      if (!rows || rows.length === 0) {
        hideOverlay();
        showModalError('Empty Sheet', 'The SourceData sheet appears to be empty.', '');
        return;
      }

      showOverlay('Updating dashboard...');

      // Normalize values: convert Date objects to YYYY-MM-DD strings, keep numbers as-is
      const cleaned = rows.map(row => {
        const r = {};
        for (const [k, v] of Object.entries(row)) {
          if (v === null || v === undefined || v === '') { r[k] = null; continue; }
          if (v instanceof Date) {
            // Format as YYYY-MM-DD
            const y = v.getFullYear();
            const m = String(v.getMonth()+1).padStart(2,'0');
            const d2 = String(v.getDate()).padStart(2,'0');
            r[k] = `${y}-${m}-${d2}`;
            continue;
          }
          // SheetJS date serial numbers for date columns (e.g. Last Period End Date)
          if (typeof v === 'number' && k.toLowerCase().includes('date')) {
            try {
              const d2 = XLSX.SSF.parse_date_code(v);
              r[k] = `${d2.y}-${String(d2.m).padStart(2,'0')}-${String(d2.d).padStart(2,'0')}`;
            } catch(e) { r[k] = v; }
            continue;
          }
          r[k] = v;
        }
        return r;
      });

      SOURCE_DATA.length = 0;
      cleaned.forEach(r => {
        delete r['Addon Signal'];
        SOURCE_DATA.push(r);
      });

      // Read SectorAnalysis sheet directly — Excel has already evaluated all formulas
      showOverlay('Parsing SectorAnalysis sheet...');
      const parsedSectors = parseSectorSheet(wb);
      SECTOR_DATA.length = 0;
      parsedSectors.forEach(s => SECTOR_DATA.push(s));

      reinitDashboard(file.name);
      hideOverlay();

      // Show Save button
      document.getElementById('saveBtn').style.display = 'inline-flex';

      // Gather stats for confirmation modal
      const companies = cleaned.filter(d => d.Ticker && d.Ticker !== '0' && d.Ticker !== 0).length;
      const sectors = new Set(cleaned.map(d => d.Sector).filter(Boolean)).size;
      const cols = Object.keys(cleaned[0] || {}).length;
      const periods = cleaned.map(d => d['Last Period End Date']).filter(Boolean);
      const latestPeriod = periods.sort().reverse()[0] || '—';
      const periodShort = String(latestPeriod).substring(0,10);

      showModalSuccess(
        `${file.name}`,
        companies, sectors, cols, periodShort
      );

    } catch(err) {
      hideOverlay();
      showModalError('Upload Failed', 'An error occurred while reading the file.', err.message);
      console.error(err);
    }
  };
  reader.onerror = () => { hideOverlay(); showModalError('Read Error', 'Could not read the file.', ''); };
  reader.readAsBinaryString(file);
}

function updateDataBadges(timestampMs) {
  const tickerCount = SOURCE_DATA.filter(d => d.Ticker && d.Ticker !== '0' && d.Ticker !== 0).length;
  setEl('lastUpdated', `DATA: ${tickerCount} COMPANIES`);

  const now = timestampMs ? new Date(timestampMs) : new Date();
  const pad = n => String(n).padStart(2,'0');
  const months2 = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const dateStr = `${pad(now.getDate())} ${months2[now.getMonth()]} ${now.getFullYear()}`;
  const timeStr = `${pad(now.getHours())}:${pad(now.getMinutes())}`;
  const badge = document.getElementById('lastUploadedBadge');
  if (badge) { badge.textContent = `🕐 Updated: ${dateStr} ${timeStr}`; badge.style.display = ''; }
}

function reinitDashboard(filename) {
  // Switch to company tab first so all DOM elements exist and are visible
  switchTab('company');

  buildColMap(SOURCE_DATA);
  allTickers = SOURCE_DATA
    .filter(d => d.Ticker && d.Ticker !== '0' && d.Ticker !== 0)
    .map(d => ({ticker: String(d.Ticker), name: String(d.Name || ''), score: d['total improvement'] || 0}))
    .sort((a,b) => b.score - a.score);

  updateDataBadges();
  // Admin just uploaded fresh data — re-check for any stocks that changed
  // into a buy signal as of this update.
  if (typeof window.checkFreshSignalsToday === 'function') window.checkFreshSignalsToday();

  const match = (filename||'').match(/(\d{4})(\d{2})(\d{2})/);
  if (match) {
    const months = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
    setEl('dataDateBadge', `PSX · ${months[parseInt(match[2])-1]} ${match[1]}`);
  }

  screenerData = SOURCE_DATA
    .filter(d => d.Ticker && d.Ticker !== '0' && d.Ticker !== 0)
    .map(d => {
      const row = {...d, Ticker: String(d.Ticker), Name: String(d.Name||''), Index: String(d.Index||'')};
      delete row['Addon Signal'];
      return row;
    });

  allSectors = [...new Set(SOURCE_DATA.filter(d=>d.Sector && d.Sector!=='0').map(d=>d.Sector))].sort();
  allTickersList = [...new Set(SOURCE_DATA.filter(d=>d.Ticker).map(d=>String(d.Ticker)))].sort();
  const allIdx = new Set();
  SOURCE_DATA.forEach(d => { if(d.Index) String(d.Index).split(',').forEach(i => allIdx.add(i.trim())); });
  allIndicesList = [...allIdx].filter(i=>i&&i!=='0').sort();
  resetMselFilters();

  const ts = document.getElementById('tickerSearch');
  if (ts) ts.value = '';
  const cl = document.getElementById('currentTickerLabel');
  if (cl) cl.innerHTML = 'Selected: <span>—</span>';

  if (allTickers.length > 0) pickTicker(allTickers[0].ticker);

  filterScreener();

  if (charts['sectors']) { charts['sectors'].destroy(); delete charts['sectors']; }
  buildSectorChart();
  buildSectorTable();
}

function showOverlay(msg) {
  document.getElementById('uploadMsg').textContent = msg;
  document.getElementById('uploadOverlay').classList.add('show');
}
function hideOverlay() {
  document.getElementById('uploadOverlay').classList.remove('show');
}

function showModalSuccess(filename, companies, sectors, cols, latestPeriod) {
  document.getElementById('modalIcon').textContent = '✅';
  document.getElementById('modalTitle').className = 'modal-title success';
  document.getElementById('modalTitle').textContent = 'Data Loaded Successfully';
  document.getElementById('modalBody').textContent = `File "${filename}" has been parsed and the dashboard is now showing the latest data.`;
  document.getElementById('statCompanies').textContent = companies;
  document.getElementById('statSectors').textContent = sectors;
  document.getElementById('statColumns').textContent = cols;
  document.getElementById('statLatestPeriod').textContent = latestPeriod;
  document.getElementById('modalStats').style.display = 'grid';
  document.getElementById('modalErrorDetail').style.display = 'none';
  document.getElementById('confirmModal').classList.add('show');
}

function showModalError(title, body, detail) {
  document.getElementById('modalIcon').textContent = '❌';
  document.getElementById('modalTitle').className = 'modal-title error';
  document.getElementById('modalTitle').textContent = title;
  document.getElementById('modalBody').textContent = body;
  document.getElementById('modalStats').style.display = 'none';
  if (detail) {
    document.getElementById('modalErrorDetail').textContent = detail;
    document.getElementById('modalErrorDetail').style.display = 'block';
  } else {
    document.getElementById('modalErrorDetail').style.display = 'none';
  }
  document.getElementById('confirmModal').classList.add('show');
}

// ===== DATA MENU =====
let dataMenuOpen = false;

function positionDataMenuMobile(menu, btn) {
  // On narrow screens the header wraps, so the menu's normal CSS
  // (position:absolute; right:0 relative to the button) can end up
  // partially off-screen. Switch to position:fixed and clamp it
  // within the viewport instead.
  if (window.innerWidth > 768) {
    menu.style.position = '';
    menu.style.top = '';
    menu.style.left = '';
    menu.style.right = '';
    menu.style.width = '';
    menu.style.minWidth = '';
    return;
  }
  const rect = btn.getBoundingClientRect();
  const menuWidth = Math.min(240, window.innerWidth - 24);
  let left = rect.right - menuWidth;
  left = Math.max(12, Math.min(left, window.innerWidth - menuWidth - 12));
  let top = rect.bottom + 8;
  const maxTop = window.innerHeight - 16;
  menu.style.position = 'fixed';
  menu.style.width = menuWidth + 'px';
  menu.style.minWidth = 'unset';
  menu.style.left = left + 'px';
  menu.style.right = 'auto';
  menu.style.top = Math.min(top, maxTop) + 'px';
}

function toggleDataMenu() {
  dataMenuOpen = !dataMenuOpen;
  const menu = document.getElementById('dataMenu');
  const chevron = document.getElementById('dataMenuChevron');
  const btn = document.getElementById('dataMenuBtn');
  if (dataMenuOpen) {
    menu.style.display = 'block';
    positionDataMenuMobile(menu, btn);
    chevron.style.transform = 'rotate(180deg)';
    btn.style.borderColor = 'var(--accent)';
    btn.style.color = 'var(--accent)';
  } else {
    menu.style.display = 'none';
    chevron.style.transform = '';
    btn.style.borderColor = 'var(--border)';
    btn.style.color = 'var(--text2)';
  }
}

function closeDataMenu() {
  dataMenuOpen = false;
  const menu = document.getElementById('dataMenu');
  const chevron = document.getElementById('dataMenuChevron');
  const btn = document.getElementById('dataMenuBtn');
  if (menu) menu.style.display = 'none';
  if (chevron) chevron.style.transform = '';
  if (btn) { btn.style.borderColor = 'var(--border)'; btn.style.color = 'var(--text2)'; }
}

window.addEventListener('resize', () => {
  if (dataMenuOpen) {
    const menu = document.getElementById('dataMenu');
    const btn = document.getElementById('dataMenuBtn');
    if (menu && btn) positionDataMenuMobile(menu, btn);
  }
});

// Close menu when clicking outside
document.addEventListener('click', function(e) {
  const btn = document.getElementById('dataMenuBtn');
  const menu = document.getElementById('dataMenu');
  if (dataMenuOpen && btn && menu && !btn.contains(e.target) && !menu.contains(e.target)) {
    closeDataMenu();
  }
});

function saveDataJson() {
  // Primary, lightweight save: just the data, not the whole app shell.
  // Push the resulting data.json to replace the one your live site fetches —
  // the HTML itself never needs to change for a routine data update.
  closeDataMenu();
  const payload = JSON.stringify({ source: SOURCE_DATA, sector: SECTOR_DATA, updatedAt: Date.now() });
  const blob = new Blob([payload], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'data.json';
  a.click();
  URL.revokeObjectURL(url);
}

function saveHtmlFile() {
  // Close the data menu before capturing HTML so it doesn't stay open on reopen
  closeDataMenu();

  // Get the current page HTML
  const html = document.documentElement.outerHTML;

  // Find the SOURCE_DATA placeholder and replace with current data
  const dataJson = JSON.stringify(SOURCE_DATA, null, 0);
  let updated = html.replace(
    /const SOURCE_DATA\s*=\s*\[[\s\S]*?\];/,
    `const SOURCE_DATA = ${dataJson};`
  );

  // Also embed SECTOR_DATA so it persists across reopens
  const sectorJson = JSON.stringify(SECTOR_DATA, null, 0);
  updated = updated.replace(
    /let SECTOR_DATA\s*=\s*\[[\s\S]*?\];/,
    `let SECTOR_DATA = ${sectorJson};`
  );

  // Trigger download
  const blob = new Blob([updated], {type: 'text/html'});
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  const badge = (document.getElementById('dataDateBadge')?.textContent || '').replace('PSX · ','').replace(' ','').toLowerCase();
  a.href     = url;
  a.download = `FS_Dashboard_${badge || 'updated'}.html`;
  a.click();
  URL.revokeObjectURL(url);
}

function closeModal() {
  document.getElementById('confirmModal').classList.remove('show');
}

// ===== STARTUP =====
fetch('./data.json')
  .then(r => r.json())
  .then(data => {
    SOURCE_DATA.length = 0;
    (data.source || []).forEach(d => SOURCE_DATA.push(d));
    SECTOR_DATA.length = 0;
    (data.sector || []).forEach(s => SECTOR_DATA.push(s));
    init();
    updateDataBadges(data.updatedAt);
    // Data wasn't necessarily ready yet at sign-in time (this fetch is async
    // and can resolve after or before auth state settles) — re-check alerts
    // now that real data definitely exists.
    if (typeof window.checkFreshSignalsToday === 'function') window.checkFreshSignalsToday();
  })
  .catch(err => {
    console.error('Failed to load data.json — dashboard will show with no data.', err);
    init();
    updateDataBadges();
  });

// ===== SHARE FUNCTIONALITY =====
const PORTAL_URL = 'https://Nexus-PSX.github.io/Nexus-PSX/';

function getShareText() {
  const ticker = (document.getElementById('currentTickerLabel')?.querySelector('span')?.textContent || '').trim();
  const score  = document.getElementById('ciScore')?.textContent?.trim() || '';
  const name   = document.getElementById('ciName')?.textContent?.trim() || '';
  const sector = document.getElementById('ciSector')?.textContent?.trim() || '';
  const period = document.getElementById('ciPeriod')?.textContent?.trim() || '';
  const hint   = document.getElementById('ciScoreHint')?.textContent?.trim() || '';
  return { ticker, score, name, sector, period, hint };
}

function toggleShareMenu() {
  const menu = document.getElementById('shareMenu');
  const isOpen = menu.style.display !== 'none';
  if (isOpen) { menu.style.display = 'none'; return; }

  // Refresh share menu header
  const { ticker, score } = getShareText();
  const sc = parseFloat(score) >= 70 ? 'var(--success)' : parseFloat(score) >= 40 ? 'var(--warn)' : 'var(--danger)';
  document.getElementById('shareMenuTitle').innerHTML =
    `Share <span style="color:var(--accent3)">${ticker}</span> · Financial Score <span style="color:${sc}">${score}</span>`;
  menu.style.display = 'block';

  // Close on outside click
  setTimeout(() => {
    document.addEventListener('click', function _close(e) {
      if (!e.target.closest('#shareMenu') && !e.target.closest('#shareBtn')) {
        menu.style.display = 'none';
        document.removeEventListener('click', _close);
      }
    });
  }, 10);
}

function buildShareMessage() {
  const { ticker, score, name, sector, period, hint } = getShareText();
  return `📊 ${ticker} — ${name}\n🏭 ${sector}\n📅 Period: ${period}\n⭐ Financial Score: ${score} ${hint}\n\nFull analysis: ${PORTAL_URL}`;
}

// ===== SCREENSHOT ENGINE =====
// Captures the live #tab-company element directly so canvas charts render correctly.
// Returns a Promise<dataURL string>.
async function captureCompanyScreenshot() {
  const target = document.getElementById('tab-company');
  if (!target) throw new Error('Company tab not found');

  const { ticker, score, name, sector, period, hint } = getShareText();
  const sc = parseFloat(score) >= 70 ? '#2ed573' : parseFloat(score) >= 40 ? '#f5a623' : '#ff4757';

  // --- Prep: temporarily make the full tab visible and remove scroll clamps ---
  const prevDisplay  = target.style.display;
  const prevOverflow = target.style.overflow;
  target.style.display  = 'block';
  target.style.overflow = 'visible';

  // Expand any scroll-clipped children
  const scrollEls = target.querySelectorAll('[style*="overflow"]');
  const scrollSaved = [];
  scrollEls.forEach(el => {
    scrollSaved.push({ el, ov: el.style.overflow, mh: el.style.maxHeight, oy: el.style.overflowY });
    el.style.overflow  = 'visible';
    el.style.maxHeight = 'none';
    el.style.overflowY = 'visible';
  });

  // Hide the share dropdown while capturing
  const shareMenu = document.getElementById('shareMenu');
  shareMenu.style.display = 'none';

  // Inject a temporary footer watermark inside the tab
  const footer = document.createElement('div');
  footer.id = '_ssFooter';
  footer.style.cssText = `
    display:flex; justify-content:space-between; align-items:center;
    padding:10px 4px 6px; margin-top:12px;
    border-top:1px solid #21262d;
    font-family:'IBM Plex Mono',monospace; font-size:11px;
  `;
  footer.innerHTML = `
    <span style="color:#484f58;">PSX Corporate Financial Performance Dashboard · ${ticker} · Financial Score <span style="color:${sc}">${score}</span> ${hint}</span>
    <span style="color:#4d94ff;">${PORTAL_URL}</span>
  `;
  target.appendChild(footer);

  // Give browser a frame to settle layout
  await new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)));

  let dataUrl;
  try {
    const canvas = await html2canvas(target, {
      backgroundColor: '#0d1117',
      scale: 1.5,
      useCORS: true,
      allowTaint: true,
      logging: false,
      scrollX: 0,
      scrollY: 0,
      windowWidth:  target.scrollWidth  + 48,
      windowHeight: target.scrollHeight + 48,
    });
    dataUrl = canvas.toDataURL('image/png');
  } finally {
    // --- Restore everything regardless of success/failure ---
    target.style.display  = prevDisplay;
    target.style.overflow = prevOverflow;
    scrollEls.forEach((el, i) => {
      el.style.overflow  = scrollSaved[i].ov;
      el.style.maxHeight = scrollSaved[i].mh;
      el.style.overflowY = scrollSaved[i].oy;
    });
    target.removeChild(footer);
  }

  return dataUrl;
}

// Shared loading state helper
function setShareBtnLoading(id, loading) {
  const btn = document.getElementById(id);
  if (!btn) return;
  if (loading) {
    btn._origHtml = btn.innerHTML;
    btn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="var(--warn)" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg> Capturing…`;
    btn.disabled = true;
  } else {
    if (btn._origHtml) btn.innerHTML = btn._origHtml;
    btn.disabled = false;
  }
}

async function shareScreenshot() {
  document.getElementById('shareMenu').style.display = 'none';
  setShareBtnLoading('screenshotBtn', true);
  try {
    const { ticker, period } = getShareText();
    const dataUrl = await captureCompanyScreenshot();
    const link = document.createElement('a');
    link.download = `PSX_${ticker}_${period || 'analysis'}.png`;
    link.href = dataUrl;
    link.click();
    const btn = document.getElementById('screenshotBtn');
    if (btn) {
      btn.innerHTML = `<svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="var(--success)" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="20 6 9 17 4 12"/></svg> Saved!`;
      setTimeout(() => setShareBtnLoading('screenshotBtn', false), 2000);
    }
  } catch(err) {
    console.error('Screenshot error:', err);
    setShareBtnLoading('screenshotBtn', false);
    showToast('Screenshot failed — try in Chrome/Edge');
  }
}

async function shareToWhatsApp() {
  document.getElementById('shareMenu').style.display = 'none';
  setShareBtnLoading('whatsappBtn', true);
  try {
    const { ticker, score, name, sector, period, hint } = getShareText();
    const dataUrl = await captureCompanyScreenshot();
    const blob = await (await fetch(dataUrl)).blob();
    const file = new File([blob], `PSX_${ticker}.png`, { type: 'image/png' });
    const msg  = `📊 ${ticker} — ${name}\n🏭 ${sector}\n📅 ${period} · Financial Score ${score} ${hint}\n🔗 ${PORTAL_URL}`;
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      await navigator.share({ files: [file], text: msg });
    } else {
      // Fallback: download image + open WhatsApp with text
      const a = document.createElement('a');
      a.href = dataUrl; a.download = `PSX_${ticker}.png`; a.click();
      setTimeout(() => window.open('https://wa.me/?text=' + encodeURIComponent(msg), '_blank'), 600);
    }
  } catch(err) {
    if (err.name !== 'AbortError') showToast('WhatsApp share failed');
  } finally {
    setShareBtnLoading('whatsappBtn', false);
  }
}

async function shareToTwitter() {
  document.getElementById('shareMenu').style.display = 'none';
  setShareBtnLoading('twitterBtn', true);
  try {
    const { ticker, score, hint } = getShareText();
    const dataUrl = await captureCompanyScreenshot();
    // Download screenshot then open Twitter compose
    const a = document.createElement('a');
    a.href = dataUrl; a.download = `PSX_${ticker}.png`; a.click();
    const text = `📊 ${ticker} · Financial Score ${score} ${hint}\n${PORTAL_URL}`;
    setTimeout(() => window.open('https://twitter.com/intent/tweet?text=' + encodeURIComponent(text), '_blank'), 700);
  } catch(err) {
    showToast('Screenshot failed');
  } finally {
    setShareBtnLoading('twitterBtn', false);
  }
}

async function shareToLinkedIn() {
  document.getElementById('shareMenu').style.display = 'none';
  setShareBtnLoading('linkedinBtn', true);
  try {
    const { ticker, score, name, sector, period, hint } = getShareText();
    const dataUrl = await captureCompanyScreenshot();
    const a = document.createElement('a');
    a.href = dataUrl; a.download = `PSX_${ticker}.png`; a.click();
    setTimeout(() => window.open('https://www.linkedin.com/sharing/share-offsite/?url=' + encodeURIComponent(PORTAL_URL), '_blank'), 700);
  } catch(err) {
    showToast('Screenshot failed');
  } finally {
    setShareBtnLoading('linkedinBtn', false);
  }
}

async function copyShareLink() {
  document.getElementById('shareMenu').style.display = 'none';
  setShareBtnLoading('copyLinkBtn', true);
  try {
    const { ticker, score, name, sector, period, hint } = getShareText();
    const dataUrl = await captureCompanyScreenshot();
    // Try clipboard write with image
    try {
      const blob = await (await fetch(dataUrl)).blob();
      await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]);
      showToast('📋 Screenshot copied to clipboard!');
    } catch {
      // Fallback: copy text + download image
      const msg = `📊 ${ticker} — ${name}\n🏭 ${sector}\n📅 ${period} · Financial Score ${score} ${hint}\n🔗 ${PORTAL_URL}`;
      await navigator.clipboard.writeText(msg);
      const a = document.createElement('a');
      a.href = dataUrl; a.download = `PSX_${ticker}.png`; a.click();
      showToast('Link copied · Screenshot saved');
    }
  } catch(err) {
    showToast('Copy failed');
  } finally {
    setShareBtnLoading('copyLinkBtn', false);
  }
}

// ===== KPI SUMMARY TICKER + SCORE BADGE =====
function updateKpiBadge() {
  try {
    const ticker = (document.getElementById('currentTickerLabel')?.querySelector('span')?.textContent || '').trim();
    const score  = parseFloat(document.getElementById('ciScore')?.textContent) || 0;
    const tickerEl = document.getElementById('kpiTickerBadge');
    const scoreEl  = document.getElementById('kpiScoreBadge');
    if (!tickerEl || !scoreEl) return;
    tickerEl.textContent = ticker;
    const sc = score >= 80 ? 'var(--success)' : score >= 40 ? 'var(--warn)' : 'var(--danger)';
    scoreEl.innerHTML = `• <span style="color:${sc};font-weight:700">Fin. Score ${score}</span>`;

    // Also sync share menu button ticker badge
    const shareMenuTitle = document.getElementById('shareMenuTitle');
    if (shareMenuTitle && document.getElementById('shareMenu').style.display === 'none') {
      shareMenuTitle.innerHTML = `Share <span style="color:var(--accent3)">${ticker}</span> · Fin. Score <span style="color:${sc}">${score}</span>`;
    }
  } catch(e) {}
}
updateKpiBadge();
setInterval(updateKpiBadge, 800);
window.addEventListener('load', updateKpiBadge);


// ===== THEME TOGGLE =====
(function() {
  const saved = localStorage.getItem('psx_theme') || 'light';
  applyTheme(saved);
  // Sync label on load
  const label = document.getElementById('themeLabel');
  const labelMap = { light: 'LIGHT', bloomberg: 'BBG', tradingview: 'TV', dark: 'DARK', cream: 'CREAM' };
  if (label) label.textContent = labelMap[saved] || 'DARK';
})();

// ===== MOBILE NAV INIT — set screener active on load =====
(function() {
  const btn = document.getElementById('mnav-screener');
  if (btn) btn.classList.add('active');
})();

function applyTheme(theme) {
  const attrMap = { light: 'light', bloomberg: 'bloomberg', tradingview: 'tradingview', cream: 'cream' };
  document.documentElement.setAttribute('data-theme', attrMap[theme] || '');
  const icon = document.getElementById('themeIcon');
  const label = document.getElementById('themeLabel');
  const iconMap = { light: '🌙', bloomberg: '📟', tradingview: '📈', cream: '🌿' };
  const labelMap = { light: 'LIGHT', bloomberg: 'BBG', tradingview: 'TV', cream: 'CREAM' };
  if (icon) icon.textContent = iconMap[theme] || '☀️';
  if (label) label.textContent = labelMap[theme] || 'DARK';
  try { localStorage.setItem('psx_theme', theme); } catch(e) {}
  // Rebuild all charts with new theme colors
  setTimeout(() => {
    if (typeof buildSectorChart === 'function') buildSectorChart();
    // Rebuild company charts if a ticker is loaded
    const cur = document.getElementById('infoName');
    if (cur && cur.textContent && currentTicker) {
      if (typeof showTicker === 'function') {
        const d = SOURCE_DATA && SOURCE_DATA.find(r => String(r.Ticker) === String(currentTicker));
        if (d) {
          const qKeys = buildQuarterKeys(d);
          const periods = qKeys.labels;
          buildLineChart('chartEPS', periods,
            [dget(d,'EPS  Q-3'), dget(d,'EPS  Q-2'), dget(d,'EPS  Q-1'), dget(d,'Latest EPS  Q')],
            [dget(d,'TTM EPS Q-3'), dget(d,'TTM EPS Q-2'), dget(d,'TTM EPS Q-1'), dget(d,'Latest TTM EPS Q')],
            'EPS (Q)', 'EPS (TTM)'
          );
          buildBarChart('chartRev', periods,
            [dget(d,'Revenue - Q-3'), dget(d,'Revenue - Q-2'), dget(d,'Revenue - Q-1'), dget(d,'Revenue - Q')],
            null
          );
          buildLineChart('chartMargins', periods,
            [dget(d,'Op Income-Q-3'), dget(d,'Op Income-Q-2'), dget(d,'Op Income-Q-1'), dget(d,'Op Income-Q')],
            [dget(d,'Net Income -Q-3'), dget(d,'Net Income -Q-2'), dget(d,'Net Income -Q-1'), dget(d,'Net Income -Q')],
            'Op Margin', 'Net Margin', true
          );
          buildPeerScoreChart(currentTicker, d.Sector);
          buildPeerYTDChart(currentTicker, d.Sector);
        }
      }
    }
  }, 50);
}

function toggleTheme() {
  const current = document.documentElement.getAttribute('data-theme') || '';
  // Cycle: dark → light → bloomberg → tradingview → cream → dark
  const next = current === '' ? 'light' : current === 'light' ? 'bloomberg' : current === 'bloomberg' ? 'tradingview' : current === 'tradingview' ? 'cream' : 'dark';
  applyTheme(next);
}



function updateEPSHeader() {
  try{
    const ticker = (document.getElementById('currentTickerLabel')?.querySelector('span')?.textContent || '').trim();
    const scoreEl = document.getElementById('ciScore');
    const labelEl = document.getElementById('chartEPSLabel');
    if(!labelEl || !scoreEl) return;

    const score = parseFloat(scoreEl.textContent) || 0;
    let color = '#ff4757';
    if(score >= 80) color = '#2ed573';
    else if(score >= 40) color = '#f5a623';

    labelEl.innerHTML =
      `<span style="color:#4d94ff">${ticker}</span> • <span style="color:${color};font-weight:700">Fin. Score ${score}</span>`;
  }catch(e){}
}
updateEPSHeader();
setInterval(updateEPSHeader, 1000);
window.addEventListener('load', updateEPSHeader);



// ===== TOP & COMPARE TAB — all helpers at module scope =====

// ── Shared formatting helpers ────────────────────────────────────────────
function topFmtPct(v) {
  if (v == null) return '—';
  const n = parseFloat(v);
  if (isNaN(n)) return '—';
  return (n >= 0 ? '+' : '') + n.toFixed(2) + '%';
}
function topFmtNum(v) {
  if (v == null) return '—';
  const n = parseFloat(v);
  if (isNaN(n)) return '—';
  if (Math.abs(n) >= 1e9) return (n/1e9).toFixed(2) + 'B';
  if (Math.abs(n) >= 1e6) return (n/1e6).toFixed(2) + 'M';
  if (Math.abs(n) >= 1e3) return (n/1e3).toFixed(1) + 'K';
  return n.toFixed(2);
}
function topScoreColor(sc) {
  const n = parseFloat(sc);
  return n >= 80 ? 'var(--success)' : n >= 40 ? 'var(--warn)' : 'var(--danger)';
}
function topPctColor(v) {
  return parseFloat(v) >= 0 ? 'var(--success)' : 'var(--danger)';
}

// ── Sortable table registry ──────────────────────────────────────────────
window._topTables = {};

function topTableHTML(tableId, rows, cols) {
  window._topTables[tableId] = { rows: rows.slice(), cols: cols, sortCol: null, sortDir: -1 };
  return topTableRender(tableId);
}

function topTableRender(tableId) {
  const st = window._topTables[tableId];
  if (!st) return '<em>no data</em>';
  const { rows, cols, sortCol, sortDir } = st;

  const sorted = rows.slice();
  if (sortCol === 'ticker') {
    sorted.sort((a,b) => String(a.Ticker||'').localeCompare(String(b.Ticker||'')) * sortDir);
  } else if (sortCol === 'name') {
    sorted.sort((a,b) => String(a.Name||'').localeCompare(String(b.Name||'')) * sortDir);
  } else if (sortCol !== null && sortCol !== 'rank') {
    const key = cols[sortCol].key;
    sorted.sort((a,b) => {
      const av = parseFloat(a[key]), bv = parseFloat(b[key]);
      if (!isNaN(av) && !isNaN(bv)) return (av - bv) * sortDir;
      return String(a[key]||'').localeCompare(String(b[key]||'')) * sortDir;
    });
  }

  const ths = `padding:7px 11px;background:var(--surface2);font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;cursor:pointer;user-select:none;white-space:nowrap;border-bottom:2px solid var(--border2);color:var(--text2);`;
  const thA = `padding:7px 11px;background:var(--accent-dim);font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.07em;cursor:pointer;user-select:none;white-space:nowrap;border-bottom:2px solid var(--accent);color:var(--accent);`;
  const arr = col => sortCol===col ? (sortDir===-1?'▼':'▲') : '<span style="opacity:.3">⇅</span>';

  let html = `<table class="top-mini-table" id="${tableId}" style="width:100%;border-collapse:collapse;">
    <thead><tr>
      <th style="${sortCol==='rank'?thA:ths}text-align:center;" onclick="sortTopTable('${tableId}','rank')"># ${arr('rank')}</th>
      <th style="${sortCol==='ticker'?thA:ths}" onclick="sortTopTable('${tableId}','ticker')">Ticker ${arr('ticker')}</th>
      <th style="${sortCol==='name'?thA:ths}" onclick="sortTopTable('${tableId}','name')">Name ${arr('name')}</th>
      ${cols.map((c,i)=>`<th style="${sortCol===i?thA:ths}text-align:${c.key==='Sector'?'left':'right'};" onclick="sortTopTable('${tableId}',${i})">${c.label} ${arr(i)}</th>`).join('')}
    </tr></thead><tbody>`;

  sorted.forEach((d,i) => {
    html += `<tr>
      <td class="rk" style="text-align:center;padding:7px 8px;">${i+1}</td>
      <td class="tk" style="cursor:pointer;padding:7px 11px;" onclick="switchTab('company');pickTicker('${(d.Ticker||'').replace(/'/g,"\\'")}');">${d.Ticker||'—'}</td>
      <td class="nm" style="padding:7px 11px;" title="${(d.Name||'').replace(/"/g,'&quot;')}">${(d.Name||'').substring(0,26)}${(d.Name||'').length>26?'…':''}</td>
      ${cols.map((c,ci)=>{
        const raw = d[c.key];
        const val = c.fmt ? c.fmt(raw) : (raw!=null?raw:'—');
        const color = c.colorFn ? c.colorFn(raw) : 'var(--text)';
        return `<td class="vl" style="padding:7px 11px;text-align:${c.key==='Sector'?'left':'right'};${sortCol===ci?'background:var(--accent-dim);':''}color:${color};">${val}</td>`;
      }).join('')}
    </tr>`;
  });
  html += `</tbody></table>`;
  return html;
}

function sortTopTable(tableId, colIdx) {
  const st = window._topTables && window._topTables[tableId];
  if (!st) return;
  st.sortDir = (st.sortCol === colIdx) ? (st.sortDir === -1 ? 1 : -1) : -1;
  st.sortCol = colIdx;
  const tbl = document.getElementById(tableId);
  if (!tbl) return;
  const wrap = tbl.parentElement;
  if (!wrap) return;
  wrap.innerHTML = topTableRender(tableId);
}

// ── buildTopTab ──────────────────────────────────────────────────────────
function buildTopTab() {
  const el = document.getElementById('tab-top');
  if (!el) return;
  if (!SOURCE_DATA || !SOURCE_DATA.length) {
    el.innerHTML = '<div style="padding:40px;text-align:center;color:var(--text2);">No data loaded yet. Please upload your data file.</div>';
    return;
  }
  el.innerHTML = '';
  el.style.display = '';

  const liquid = SOURCE_DATA.filter(d => parseFloat(d['Price']) > 0 && parseFloat(d['Volume']) > 0);

  // 1. Top per sector by score
  const sectorScoreMap = {};
  SOURCE_DATA.forEach(d => {
    const s = d['Sector']; if (!s) return;
    const sc = parseFloat(d['total improvement']) || 0;
    if (!sectorScoreMap[s] || sc > (parseFloat(sectorScoreMap[s]['total improvement'])||0)) sectorScoreMap[s] = d;
  });
  const topByScoreFiltered = Object.values(sectorScoreMap)
    .filter(d => (parseFloat(d['total improvement'])||0) > 80)
    .sort((a,b) => (parseFloat(b['total improvement'])||0) - (parseFloat(a['total improvement'])||0));

  // 2. Top per sector by YTD
  const sectorYTDMap = {};
  liquid.forEach(d => {
    const s = d['Sector']; if (!s) return;
    const ytd = parseFloat(d['YTD Return %']) || 0;
    if (!sectorYTDMap[s] || ytd > (parseFloat(sectorYTDMap[s]['YTD Return %'])||0)) sectorYTDMap[s] = d;
  });
  const topByYTDFiltered = Object.values(sectorYTDMap)
    .filter(d => (parseFloat(d['YTD Return %'])||0) > 10)
    .sort((a,b) => (parseFloat(b['YTD Return %'])||0) - (parseFloat(a['YTD Return %'])||0));

  // 3. Daily upper/lower caps — use full SOURCE_DATA (not liquid) because PSX
  // stocks can hit the ±10% circuit breaker even with zero volume (offered at
  // limit with no trades clearing), so the Volume > 0 liquid filter would
  // incorrectly drop them.
  const dayGainers = SOURCE_DATA.filter(d => {
    const v = parseFloat(d['Day Change %']);
    return !isNaN(v) && v >= 9.9;  // 9.9 threshold catches PSX ±10% circuit breaker stocks
  }).sort((a,b)=>(parseFloat(b['Day Change %'])||0)-(parseFloat(a['Day Change %'])||0));   // even if stored as 9.999... due to floating point
  const dayLosers = SOURCE_DATA.filter(d => {
    const v = parseFloat(d['Day Change %']);
    return !isNaN(v) && v <= -9.9;
  }).sort((a,b)=>(parseFloat(a['Day Change %'])||0)-(parseFloat(b['Day Change %'])||0));

  // 4. Weekly gainers/losers — same ±9.9% cutoff as the daily caps, scrollable.
  const weekGainers = SOURCE_DATA.filter(d => {
    const v = parseFloat(d['Current Week Return %']);
    return !isNaN(v) && v >= 9.9;
  }).sort((a,b)=>(parseFloat(b['Current Week Return %'])||0)-(parseFloat(a['Current Week Return %'])||0));
  const weekLosers = SOURCE_DATA.filter(d => {
    const v = parseFloat(d['Current Week Return %']);
    return !isNaN(v) && v <= -9.9;
  }).sort((a,b)=>(parseFloat(a['Current Week Return %'])||0)-(parseFloat(b['Current Week Return %'])||0));

  // 5. Top RVOL
  const topRVOL = [...liquid].filter(d=>(parseFloat(d['Relative Vol'])||0)>0)
    .sort((a,b)=>(parseFloat(b['Relative Vol'])||0)-(parseFloat(a['Relative Vol'])||0)).slice(0,10);

  // 6. Top discount (most negative)
  const topDisc = [...liquid].filter(d=>d['Discount Ratio']!=null && (parseFloat(d['Discount Ratio'])||0)<0)
    .sort((a,b)=>(parseFloat(a['Discount Ratio'])||0)-(parseFloat(b['Discount Ratio'])||0)).slice(0,10);

  // 7. Turnaround candidates — Loss narrow > 0, ranked by total score
  const turnaroundList = [...SOURCE_DATA].filter(d=>(parseFloat(d['Loss narrow'])||0)>0)
    .sort((a,b)=>(parseFloat(b['total improvement'])||0)-(parseFloat(a['total improvement'])||0));

  // col definitions
  const scoreColsDef = [
    {key:'Sector',            label:'Sector',  fmt:v=>(v||'').substring(0,20), colorFn:()=>'var(--text2)'},
    {key:'total improvement', label:'Fin. Score',   fmt:v=>parseFloat(v||0).toFixed(0), colorFn:v=>topScoreColor(v)},
    {key:'Price',             label:'Price',   fmt:v=>'PKR '+(parseFloat(v||0).toFixed(2)), colorFn:()=>'var(--text)'},
    {key:'YTD Return %',      label:'YTD %',   fmt:topFmtPct, colorFn:v=>topPctColor(v)}
  ];
  const ytdColsDef = [
    {key:'Sector',            label:'Sector',  fmt:v=>(v||'').substring(0,20), colorFn:()=>'var(--text2)'},
    {key:'YTD Return %',      label:'YTD %',   fmt:topFmtPct, colorFn:v=>topPctColor(v)},
    {key:'Price',             label:'Price',   fmt:v=>'PKR '+(parseFloat(v||0).toFixed(2)), colorFn:()=>'var(--text)'},
    {key:'total improvement', label:'Fin. Score',   fmt:v=>parseFloat(v||0).toFixed(0), colorFn:v=>topScoreColor(v)}
  ];
  const dayColsDef = [
    {key:'Day Change %',      label:'Day %',   fmt:topFmtPct, colorFn:v=>topPctColor(v)},
    {key:'Price',             label:'Price',   fmt:v=>parseFloat(v||0).toFixed(2), colorFn:()=>'var(--text)'},
    {key:'total improvement', label:'Fin. Score',   fmt:v=>parseFloat(v||0).toFixed(0), colorFn:v=>topScoreColor(v)}
  ];
  const weekColsDef = [
    {key:'Current Week Return %', label:'Week %', fmt:topFmtPct, colorFn:v=>topPctColor(v)},
    {key:'Price',                 label:'Price',  fmt:v=>parseFloat(v||0).toFixed(2), colorFn:()=>'var(--text)'},
    {key:'total improvement',     label:'Fin. Score',  fmt:v=>parseFloat(v||0).toFixed(0), colorFn:v=>topScoreColor(v)}
  ];
  const rvolColsDef = [
    {key:'Relative Vol',  label:'RVOL',   fmt:v=>parseFloat(v||0).toFixed(2)+'x', colorFn:v=>parseFloat(v||0)>=2?'var(--warn)':'var(--text)'},
    {key:'Volume',        label:'Volume', fmt:topFmtNum, colorFn:()=>'var(--text2)'},
    {key:'Day Change %',  label:'Day %',  fmt:topFmtPct, colorFn:v=>topPctColor(v)}
  ];
  const discColsDef = [
    {key:'Discount Ratio',    label:'Disc %', fmt:topFmtPct, colorFn:()=>'var(--danger)'},
    {key:'Price',             label:'Price',  fmt:v=>parseFloat(v||0).toFixed(2), colorFn:()=>'var(--text)'},
    {key:'total improvement', label:'Fin. Score',  fmt:v=>parseFloat(v||0).toFixed(0), colorFn:v=>topScoreColor(v)}
  ];
  const turnColsDef = [
    {key:'total improvement', label:'Fin. Score', fmt:v=>parseFloat(v||0).toFixed(0), colorFn:v=>topScoreColor(v)},
    {key:'YTD Return %',      label:'YTD %',       fmt:topFmtPct, colorFn:v=>topPctColor(v)}
  ];

  const row2 = (card1, card2) => `<div class="top-grid-2" style="margin-bottom:16px;">${card1}${card2}</div>`;
  const row3 = (card1, card2, card3) => `<div class="top-grid-3" style="margin-bottom:16px;">${card1}${card2}${card3}</div>`;
  const card = (icon, title, tableId, rows, cols) => `
    <div class="top-table-card">
      <div class="top-table-head"><span class="th-icon">${icon}</span>${title}</div>
      <div style="overflow-x:auto;" id="wrap-${tableId}">${topTableHTML(tableId, rows, cols)}</div>
    </div>`;
  const cardScroll = (icon, title, tableId, rows, cols) => `
    <div class="top-table-card">
      <div class="top-table-head"><span class="th-icon">${icon}</span>${title} <span style="margin-left:auto;font-weight:600;color:var(--text3);">${rows.length}</span></div>
      <div style="overflow-x:auto;max-height:360px;overflow-y:auto;" id="wrap-${tableId}">${topTableHTML(tableId, rows, cols)}</div>
    </div>`;

  el.innerHTML = `<div style="padding-bottom:32px;">
    ${row2(
      cardScroll('🟢','Daily Upper Caps','dayGain',dayGainers,dayColsDef),
      cardScroll('🔴','Daily Lower Caps','dayLoss',dayLosers,dayColsDef)
    )}
    ${row2(
      cardScroll('🚀','Weekly Top Gainers','weekGain',weekGainers,weekColsDef),
      cardScroll('📉','Weekly Top Losers','weekLoss',weekLosers,weekColsDef)
    )}
    ${row3(
      card('🔊','Top 10 Relative Volume','rvol',topRVOL,rvolColsDef),
      cardScroll('🔁','Turnaround Candidates (Loss Narrowing)','turn',turnaroundList,turnColsDef),
      card('💰','Top 10 Discount (Most Undervalued)','disc',topDisc,discColsDef)
    )}
    ${row2(
      card('🏅','Top per Sector — Fin. Score &gt; 80','score',topByScoreFiltered,scoreColsDef),
      card('📈','Top per Sector — YTD &gt; 10%','ytd',topByYTDFiltered,ytdColsDef)
    )}
  </div>`;
}


// ===== DISCLAIMER MODAL =====
function showDisclaimer(callback) {
  const overlay = document.getElementById('disclaimerOverlay');
  overlay._callback = callback || null;
  overlay.classList.add('open');
  document.getElementById('disclaimerModal').querySelector('.disclaimer-ok').focus();
}
function closeDisclaimer() {
  const overlay = document.getElementById('disclaimerOverlay');
  overlay.classList.remove('open');
  if (typeof overlay._callback === 'function') {
    overlay._callback();
    overlay._callback = null;
  }
}
// Close on overlay click (outside modal)
document.addEventListener('DOMContentLoaded', function() {
  document.getElementById('disclaimerOverlay').addEventListener('click', function(e) {
    if (e.target === this) closeDisclaimer();
  });
});

// Wrap toggleTechnical to show disclaimer first (only when revealing)
const _origToggleTechnical = toggleTechnical;
(function() {
  const origFn = toggleTechnical;
  window.toggleTechnical = function() {
    const table = document.getElementById('screenerTable');
    const isCurrentlyHidden = table.classList.contains('hide-technical');
    if (isCurrentlyHidden) {
      // About to SHOW technical — show disclaimer first, then toggle
      showDisclaimer(function() { origFn(); });
    } else {
      // About to HIDE — no disclaimer needed
      origFn();
    }
  };
})();



// ── Mobile chart tooltip dismiss ─────────────────────────────────────────────
// On touch devices, Chart.js tooltips appear on tap but have no built-in
// "tap elsewhere to dismiss" behaviour — they stay stuck on screen until
// the user taps the exact same spot again. This listener fires on every
// touchstart: if the touch lands outside all Company View chart canvases,
// it hides the active tooltip on every chart in the registry.
(function () {
  const COMPANY_CHART_IDS = ['chartEPS', 'chartRev', 'chartMargins',
                              'chartPeerScore', 'chartPeerYTD'];

  function dismissChartTooltips(touchedEl) {
    if (typeof charts === 'undefined') return;
    COMPANY_CHART_IDS.forEach(id => {
      const instance = charts[id];
      if (!instance) return;
      const canvas = document.getElementById(id);
      // Only dismiss if the touch was outside this specific canvas
      if (canvas && canvas.contains(touchedEl)) return;
      instance.tooltip.setActiveElements([], {x:0, y:0});
      instance.update('none');
    });
  }

  document.addEventListener('touchstart', function (e) {
    dismissChartTooltips(e.target);
  }, { passive: true });
})();
