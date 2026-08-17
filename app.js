/* ============================================================
   DataPulse Dashboard — app.js
   Full interactive Excel analytics dashboard
   ============================================================ */

'use strict';

// ── STATE ──────────────────────────────────────────────────
const STATE = {
  rawData: [],       // original parsed rows
  filteredData: [],  // after slicers/timeline
  headers: [],       // column names
  colTypes: {},      // 'numeric' | 'date' | 'string'
  fileName: '',
  charts: {},        // Chart.js instances
  sortCol: null,
  sortDir: 'asc',
  page: 0,
  pageSize: 50,
  analysisRun: false,
  insightsRun: false,
};

// ── CHART COLORS ───────────────────────────────────────────
const COLORS = [
  '#f5a623','#4a9eff','#00d4aa','#9b72f7',
  '#ff6b35','#ff4757','#2ed573','#ffd32a',
  '#7bed9f','#a4b0be','#eccc68','#70a1ff',
];
const gradient = (ctx, c1, c2) => {
  const g = ctx.createLinearGradient(0, 0, 0, 300);
  g.addColorStop(0, c1); g.addColorStop(1, c2);
  return g;
};

// ── UTILS ──────────────────────────────────────────────────
const fmt = (n, digits = 2) => {
  if (n === null || n === undefined || isNaN(n)) return '—';
  if (Math.abs(n) >= 1e9) return (n / 1e9).toFixed(1) + 'B';
  if (Math.abs(n) >= 1e6) return (n / 1e6).toFixed(1) + 'M';
  if (Math.abs(n) >= 1e3) return (n / 1e3).toFixed(1) + 'K';
  return Number(n).toFixed(digits);
};
const fmtFull = n => (isNaN(n) ? '—' : Number(n).toLocaleString(undefined, { maximumFractionDigits: 2 }));
const capitalize = s => s.charAt(0).toUpperCase() + s.slice(1);
const isDate = v => {
  if (v instanceof Date) return true;
  if (typeof v === 'string') {
    const d = Date.parse(v);
    return !isNaN(d) && /\d{4}|\d{2}[\/-]\d{2}/.test(v);
  }
  return false;
};
const toDate = v => v instanceof Date ? v : new Date(v);
const mean = arr => arr.reduce((a, b) => a + b, 0) / arr.length;
const median = arr => {
  const s = [...arr].sort((a, b) => a - b);
  const m = Math.floor(s.length / 2);
  return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
};
const std = arr => {
  const m = mean(arr);
  return Math.sqrt(arr.reduce((s, v) => s + (v - m) ** 2, 0) / arr.length);
};
const correlation = (xs, ys) => {
  const n = xs.length, mx = mean(xs), my = mean(ys);
  const num = xs.reduce((s, x, i) => s + (x - mx) * (ys[i] - my), 0);
  const den = Math.sqrt(xs.reduce((s, x) => s + (x - mx) ** 2, 0) *
                        ys.reduce((s, y) => s + (y - my) ** 2, 0));
  return den === 0 ? 0 : num / den;
};

// ── DOM HELPERS ────────────────────────────────────────────
const $  = id => document.getElementById(id);
const qs = sel => document.querySelector(sel);

function showSpinner(msg = 'Processing…') {
  $('spinnerMsg').textContent = msg;
  $('spinner').classList.remove('hidden');
}
function hideSpinner() { $('spinner').classList.add('hidden'); }

function showScreen(name) {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  $(`${name}-screen`).classList.add('active');
}

// ── UPLOAD / FILE HANDLING ─────────────────────────────────
$('dropZone').addEventListener('dragover', e => { e.preventDefault(); $('dropZone').classList.add('dragover'); });
$('dropZone').addEventListener('dragleave', () => $('dropZone').classList.remove('dragover'));
$('dropZone').addEventListener('drop', e => {
  e.preventDefault();
  $('dropZone').classList.remove('dragover');
  const f = e.dataTransfer.files[0];
  if (f) processFile(f);
});
$('fileInput').addEventListener('change', e => { if (e.target.files[0]) processFile(e.target.files[0]); });
$('dropZone').addEventListener('click', e => { if (e.target !== qs('.btn-upload')) $('fileInput').click(); });
$('newFileBtn').addEventListener('click', () => showScreen('upload'));
$('loadSample').addEventListener('click', loadSampleData);

async function processFile(file) {
  STATE.fileName = file.name;
  $('upload-progress').classList.remove('hidden');
  setProgress(10, 'Reading file…');
  const reader = new FileReader();
  reader.onload = async e => {
    try {
      setProgress(40, 'Parsing Excel…');
      const data = e.target.result;
      const wb = XLSX.read(data, { type: 'array', cellDates: true });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: null, raw: false, dateNF: 'yyyy-mm-dd' });
      setProgress(70, 'Analysing columns…');
      await delay(100);
      ingestData(rows, file.name);
      setProgress(100, 'Done!');
      await delay(300);
      $('upload-progress').classList.add('hidden');
      showScreen('dashboard');
      buildDashboard();
    } catch(err) {
      alert('Error reading file: ' + err.message);
      $('upload-progress').classList.add('hidden');
    }
  };
  reader.readAsArrayBuffer(file);
}

function setProgress(pct, msg) {
  $('progressFill').style.width = pct + '%';
  $('progressLabel').textContent = msg;
}
const delay = ms => new Promise(r => setTimeout(r, ms));

function ingestData(rows, fname) {
  if (!rows.length) return;
  STATE.rawData = rows;
  STATE.filteredData = [...rows];
  STATE.headers = Object.keys(rows[0]);
  STATE.fileName = fname || STATE.fileName;
  STATE.colTypes = {};
  STATE.headers.forEach(h => {
    const vals = rows.map(r => r[h]).filter(v => v !== null && v !== '');
    const numVals = vals.map(v => parseFloat(v)).filter(v => !isNaN(v));
    if (numVals.length / Math.max(vals.length, 1) > 0.7) {
      STATE.colTypes[h] = 'numeric';
    } else if (vals.some(v => isDate(v))) {
      STATE.colTypes[h] = 'date';
    } else {
      STATE.colTypes[h] = 'string';
    }
  });
}

// ── SAMPLE DATA ────────────────────────────────────────────
function loadSampleData() {
  const categories = ['Electronics','Clothing','Food','Sports','Books','Furniture','Beauty'];
  const regions = ['North','South','East','West','Central'];
  const reps = ['Alice','Bob','Charlie','Diana','Evan','Fiona'];
  const rows = [];
  const base = new Date('2023-01-01');
  for (let i = 0; i < 500; i++) {
    const d = new Date(base);
    d.setDate(d.getDate() + Math.floor(i * 365 / 500));
    const cat = categories[i % categories.length];
    const qty = Math.floor(Math.random() * 100) + 1;
    const price = +(Math.random() * 300 + 10).toFixed(2);
    rows.push({
      Date: d.toISOString().split('T')[0],
      Category: cat,
      Region: regions[i % regions.length],
      SalesRep: reps[i % reps.length],
      Quantity: qty,
      UnitPrice: price,
      Revenue: +(qty * price).toFixed(2),
      Cost: +(qty * price * (0.4 + Math.random() * 0.2)).toFixed(2),
      Profit: +(qty * price * (0.1 + Math.random() * 0.3)).toFixed(2),
      CustomerRating: +(Math.random() * 2 + 3).toFixed(1),
    });
  }
  ingestData(rows, 'sample_sales_data.xlsx');
  showScreen('dashboard');
  buildDashboard();
}

// ── BUILD DASHBOARD ────────────────────────────────────────
function buildDashboard() {
  showSpinner('Building dashboard…');
  setTimeout(() => {
    updateHeaderInfo();
    buildSlicers();
    buildKPIs();
    populateChartSelects();
    buildAllCharts();
    buildTable();
    hideSpinner();
    STATE.analysisRun = false;
    STATE.insightsRun = false;
    $('analysisBody').innerHTML = '<div class="analysis-placeholder">Click <strong>Run Analysis</strong> to generate deep statistical insights from your data.</div>';
    $('insightsBody').innerHTML = '<div class="analysis-placeholder">Click <strong>Generate Insights</strong> to surface key findings and recommendations.</div>';
  }, 50);
}

function updateHeaderInfo() {
  $('fileName').textContent = STATE.fileName;
  $('fileRows').textContent = `${STATE.rawData.length.toLocaleString()} rows · ${STATE.headers.length} cols`;
}

// ── SLICERS ────────────────────────────────────────────────
function buildSlicers() {
  const wrap = $('slicersWrap');
  wrap.innerHTML = '';
  const tSel = $('timelineCol');
  tSel.innerHTML = '<option value="">— Select Date Col —</option>';

  const strCols = STATE.headers.filter(h => STATE.colTypes[h] === 'string').slice(0, 4);
  const dateCols = STATE.headers.filter(h => STATE.colTypes[h] === 'date');

  dateCols.forEach(h => {
    const opt = document.createElement('option');
    opt.value = h; opt.textContent = h;
    tSel.appendChild(opt);
  });
  if (dateCols.length) tSel.value = dateCols[0];

  strCols.forEach(col => {
    const vals = [...new Set(STATE.rawData.map(r => r[col]).filter(Boolean))].sort();
    const grp = document.createElement('div');
    grp.className = 'slicer-group';
    grp.innerHTML = `<div class="slicer-group-label">${col}</div>`;
    const sel = document.createElement('select');
    sel.className = 'slicer-select';
    sel.id = `slicer_${col}`;
    sel.innerHTML = `<option value="">All</option>`;
    vals.forEach(v => { const o = document.createElement('option'); o.value = v; o.textContent = v; sel.appendChild(o); });
    grp.appendChild(sel);
    wrap.appendChild(grp);
  });
}

$('applyFilters').addEventListener('click', applyFilters);
$('clearFilters').addEventListener('click', clearFilters);

function applyFilters() {
  let data = [...STATE.rawData];
  // String slicers
  document.querySelectorAll('[id^="slicer_"]').forEach(sel => {
    if (sel.value) {
      const col = sel.id.replace('slicer_', '');
      data = data.filter(r => String(r[col]) === sel.value);
    }
  });
  // Timeline
  const dateCol = $('timelineCol').value;
  const from = $('dateFrom').value;
  const to = $('dateTo').value;
  if (dateCol && from) data = data.filter(r => r[dateCol] && new Date(r[dateCol]) >= new Date(from));
  if (dateCol && to)   data = data.filter(r => r[dateCol] && new Date(r[dateCol]) <= new Date(to));
  STATE.filteredData = data;
  STATE.page = 0;
  buildKPIs();
  buildAllCharts();
  buildTable();
}

function clearFilters() {
  document.querySelectorAll('[id^="slicer_"]').forEach(s => s.value = '');
  $('dateFrom').value = ''; $('dateTo').value = '';
  STATE.filteredData = [...STATE.rawData];
  STATE.page = 0;
  buildKPIs();
  buildAllCharts();
  buildTable();
}

// ── KPIs ───────────────────────────────────────────────────
const KPI_DEFS = [
  { icon: '📦', label: 'Total Records',   fn: d => d.length,                                fmt: n => n.toLocaleString(),      color: '#f5a623', delta: 'count' },
  { icon: '💰', label: 'Total Revenue',   fn: d => sumCol(d, numCols(d)[0]),                fmt: n => '$' + fmt(n),            color: '#4a9eff', delta: 'numeric' },
  { icon: '📈', label: 'Avg Value',       fn: d => avgCol(d, numCols(d)[0]),                fmt: n => fmt(n, 2),               color: '#00d4aa', delta: 'numeric' },
  { icon: '🏆', label: 'Max Value',       fn: d => maxCol(d, numCols(d)[0]),                fmt: n => fmtFull(n),              color: '#9b72f7', delta: 'numeric' },
  { icon: '📊', label: 'Unique Categories', fn: d => uniqueCount(d, strCols(d)[0]),         fmt: n => n.toLocaleString(),      color: '#ff6b35', delta: 'neutral' },
];

function numCols(d) { return STATE.headers.filter(h => STATE.colTypes[h] === 'numeric'); }
function strCols(d) { return STATE.headers.filter(h => STATE.colTypes[h] === 'string'); }
function sumCol(d, col) { if (!col) return 0; return d.reduce((s, r) => s + (parseFloat(r[col]) || 0), 0); }
function avgCol(d, col) { if (!col || !d.length) return 0; return sumCol(d, col) / d.length; }
function maxCol(d, col) { if (!col || !d.length) return 0; return Math.max(...d.map(r => parseFloat(r[col]) || 0)); }
function uniqueCount(d, col) { if (!col) return 0; return new Set(d.map(r => r[col])).size; }

function buildKPIs() {
  const grid = $('kpiGrid');
  grid.innerHTML = '';
  const numC = numCols();
  const strC = strCols();
  const d = STATE.filteredData;

  // Dynamic KPIs based on actual columns
  const kpis = [];

  // KPI 1: Total rows
  kpis.push({ icon: '📦', label: 'Total Records', value: d.length.toLocaleString(), sub: `of ${STATE.rawData.length.toLocaleString()} total`, color: '#f5a623', delta: null });

  // KPI 2–4: First 3 numeric columns summed
  numC.slice(0, 3).forEach((col, i) => {
    const colors = ['#4a9eff','#00d4aa','#9b72f7'];
    const icons  = ['💰','📈','🏆'];
    const sum = sumCol(d, col);
    const rawSum = sumCol(STATE.rawData, col);
    const pct = rawSum ? ((sum / rawSum - 1) * 100).toFixed(1) : 0;
    kpis.push({
      icon: icons[i], label: `Total ${col}`, value: '$' + fmt(sum),
      sub: `Avg: ${fmt(avgCol(d, col), 2)}`,
      color: colors[i],
      delta: { pct, dir: pct >= 0 ? 'up' : 'down' }
    });
  });

  // KPI 5: Unique string column
  if (strC.length) {
    const cnt = uniqueCount(d, strC[0]);
    kpis.push({ icon: '📊', label: `Unique ${strC[0]}`, value: cnt, sub: `column: ${strC[0]}`, color: '#ff6b35', delta: null });
  } else {
    const col = numC[3] || numC[0];
    if (col) kpis.push({ icon: '📉', label: `Max ${col}`, value: fmtFull(maxCol(d, col)), sub: `Min: ${fmt(Math.min(...d.map(r=>parseFloat(r[col])||0)),2)}`, color: '#ff6b35', delta: null });
  }

  kpis.slice(0, 5).forEach(k => {
    const el = document.createElement('div');
    el.className = 'kpi-card';
    el.style.setProperty('--kpi-color', k.color);
    let deltaHtml = '';
    if (k.delta) {
      const cls = k.delta.dir === 'up' ? 'up' : 'down';
      const arrow = k.delta.dir === 'up' ? '▲' : '▼';
      deltaHtml = `<div class="kpi-delta ${cls}">${arrow} ${Math.abs(k.delta.pct)}% vs total</div>`;
    }
    el.innerHTML = `
      <div class="kpi-icon">${k.icon}</div>
      <div class="kpi-label">${k.label}</div>
      <div class="kpi-value animate">${k.value}</div>
      <div class="kpi-sub">${k.sub}</div>
      ${deltaHtml}
    `;
    grid.appendChild(el);
  });
}

// ── CHART SELECTS ──────────────────────────────────────────
function populateChartSelects() {
  const numC = numCols();
  const strC = strCols();
  const dateC = STATE.headers.filter(h => STATE.colTypes[h] === 'date');
  const allC  = STATE.headers;

  const fill = (id, cols, defaultIdx = 0) => {
    const el = $(id); if (!el) return;
    el.innerHTML = cols.map((c, i) => `<option value="${c}" ${i === defaultIdx ? 'selected' : ''}>${c}</option>`).join('');
  };

  fill('barColX', [...strC, ...dateC], 0);
  fill('barColY', numC, 0);
  fill('lineColX', [...dateC, ...strC], 0);
  fill('lineColY', numC, 0);
  fill('pieCol', strC, 0);
  fill('scatterX', numC, 0);
  fill('scatterY', numC, Math.min(1, numC.length - 1));
  fill('hbarColX', [...strC, ...dateC], 0);
  fill('hbarColY', numC, 0);

  // Re-render on change
  ['barColX','barColY'].forEach(id => $(id)?.addEventListener('change', () => buildBar()));
  ['lineColX','lineColY'].forEach(id => $(id)?.addEventListener('change', () => buildLine()));
  $('pieCol')?.addEventListener('change', () => buildPie());
  ['scatterX','scatterY'].forEach(id => $(id)?.addEventListener('change', () => buildScatter()));
  ['hbarColX','hbarColY'].forEach(id => $(id)?.addEventListener('change', () => buildHBar()));
}

// ── CHART BUILDER ──────────────────────────────────────────
const chartDefaults = () => ({
  plugins: {
    legend: { labels: { color: '#8892a8', font: { family: 'IBM Plex Mono', size: 11 } } },
    tooltip: {
      backgroundColor: '#181c24', borderColor: '#2e3650', borderWidth: 1,
      titleColor: '#e8eaf2', bodyColor: '#8892a8',
      titleFont: { family: 'Syne', weight: '700' },
    }
  },
  scales: {
    x: { ticks: { color: '#555e75', font: { family: 'IBM Plex Mono', size: 10 } }, grid: { color: '#1e2330' } },
    y: { ticks: { color: '#555e75', font: { family: 'IBM Plex Mono', size: 10 } }, grid: { color: '#1e2330' } }
  },
  animation: { duration: 600, easing: 'easeInOutQuart' },
  responsive: true, maintainAspectRatio: false,
});

function destroyChart(key) {
  if (STATE.charts[key]) { STATE.charts[key].destroy(); delete STATE.charts[key]; }
}

function buildAllCharts() {
  buildBar(); buildLine(); buildPie(); buildScatter(); buildHBar();
}

// ── BAR CHART ──────────────────────────────────────────────
function buildBar() {
  destroyChart('bar');
  const xCol = $('barColX')?.value, yCol = $('barColY')?.value;
  if (!xCol || !yCol) return;
  const agg = groupSum(STATE.filteredData, xCol, yCol, 12);
  const ctx = $('barChart').getContext('2d');
  const opts = chartDefaults();
  STATE.charts.bar = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: agg.map(a => a.key),
      datasets: [{
        label: yCol,
        data: agg.map(a => a.val),
        backgroundColor: agg.map((_, i) => COLORS[i % COLORS.length] + 'cc'),
        borderColor: agg.map((_, i) => COLORS[i % COLORS.length]),
        borderWidth: 1, borderRadius: 5,
      }]
    },
    options: { ...opts, plugins: { ...opts.plugins, legend: { ...opts.plugins.legend, display: false } } }
  });
  $('title-bar').textContent = `${yCol} by ${xCol}`;
}

// ── LINE CHART ─────────────────────────────────────────────
function buildLine() {
  destroyChart('line');
  const xCol = $('lineColX')?.value, yCol = $('lineColY')?.value;
  if (!xCol || !yCol) return;
  const agg = groupSum(STATE.filteredData, xCol, yCol, 30);
  const ctx = $('lineChart').getContext('2d');
  const opts = chartDefaults();
  const grad = gradient(ctx, COLORS[0] + '80', COLORS[0] + '00');
  STATE.charts.line = new Chart(ctx, {
    type: 'line',
    data: {
      labels: agg.map(a => a.key),
      datasets: [{
        label: yCol, data: agg.map(a => a.val),
        borderColor: COLORS[0], backgroundColor: grad,
        borderWidth: 2, pointRadius: 3, pointBackgroundColor: COLORS[0],
        fill: true, tension: 0.4,
      }]
    },
    options: opts
  });
  $('title-line').textContent = `${yCol} Trend over ${xCol}`;
}

// ── PIE CHART ──────────────────────────────────────────────
function buildPie() {
  destroyChart('pie');
  const col = $('pieCol')?.value;
  if (!col) return;
  const counts = {};
  STATE.filteredData.forEach(r => { const v = r[col] || 'Unknown'; counts[v] = (counts[v] || 0) + 1; });
  const entries = Object.entries(counts).sort((a, b) => b[1] - a[1]).slice(0, 10);
  const ctx = $('pieChart').getContext('2d');
  const opts = chartDefaults();
  delete opts.scales;
  STATE.charts.pie = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: entries.map(e => e[0]),
      datasets: [{ data: entries.map(e => e[1]), backgroundColor: COLORS, borderColor: '#0a0b0e', borderWidth: 2, hoverOffset: 6 }]
    },
    options: { ...opts, cutout: '60%' }
  });
}

// ── SCATTER CHART ──────────────────────────────────────────
function buildScatter() {
  destroyChart('scatter');
  const xCol = $('scatterX')?.value, yCol = $('scatterY')?.value;
  if (!xCol || !yCol) return;
  const pts = STATE.filteredData
    .map(r => ({ x: parseFloat(r[xCol]), y: parseFloat(r[yCol]) }))
    .filter(p => !isNaN(p.x) && !isNaN(p.y))
    .slice(0, 500);
  const ctx = $('scatterChart').getContext('2d');
  const opts = chartDefaults();
  STATE.charts.scatter = new Chart(ctx, {
    type: 'scatter',
    data: { datasets: [{ label: `${xCol} vs ${yCol}`, data: pts, backgroundColor: COLORS[1] + '99', pointRadius: 4 }] },
    options: {
      ...opts,
      plugins: { ...opts.plugins, legend: { ...opts.plugins.legend, display: false } },
      scales: {
        x: { ...opts.scales.x, title: { display: true, text: xCol, color: '#8892a8', font: { size: 10 } } },
        y: { ...opts.scales.y, title: { display: true, text: yCol, color: '#8892a8', font: { size: 10 } } }
      }
    }
  });
}

// ── HORIZONTAL BAR ─────────────────────────────────────────
function buildHBar() {
  destroyChart('hbar');
  const xCol = $('hbarColX')?.value, yCol = $('hbarColY')?.value;
  if (!xCol || !yCol) return;
  const agg = groupSum(STATE.filteredData, xCol, yCol, 15).sort((a, b) => b.val - a.val);
  const ctx = $('hbarChart').getContext('2d');
  const opts = chartDefaults();
  STATE.charts.hbar = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: agg.map(a => a.key),
      datasets: [{
        label: yCol, data: agg.map(a => a.val),
        backgroundColor: COLORS.map(c => c + 'cc'), borderColor: COLORS, borderWidth: 1, borderRadius: 4,
      }]
    },
    options: {
      ...opts, indexAxis: 'y',
      plugins: { ...opts.plugins, legend: { ...opts.plugins.legend, display: false } }
    }
  });
}

// ── AGGREGATION HELPER ─────────────────────────────────────
function groupSum(data, groupCol, valueCol, limit = 20) {
  const agg = {};
  data.forEach(r => {
    const k = String(r[groupCol] || 'Unknown').substring(0, 20);
    const v = parseFloat(r[valueCol]) || 0;
    agg[k] = (agg[k] || 0) + v;
  });
  return Object.entries(agg)
    .sort((a, b) => b[1] - a[1])
    .slice(0, limit)
    .map(([key, val]) => ({ key, val }));
}

// ── TABLE ──────────────────────────────────────────────────
let tableData = [];
let sortState = { col: null, dir: 'asc' };

function buildTable(search = '') {
  const raw = STATE.filteredData;
  tableData = search
    ? raw.filter(r => STATE.headers.some(h => String(r[h] ?? '').toLowerCase().includes(search.toLowerCase())))
    : [...raw];

  if (sortState.col) {
    tableData.sort((a, b) => {
      const va = a[sortState.col], vb = b[sortState.col];
      const na = parseFloat(va), nb = parseFloat(vb);
      if (!isNaN(na) && !isNaN(nb)) return sortState.dir === 'asc' ? na - nb : nb - na;
      return sortState.dir === 'asc' ? String(va).localeCompare(String(vb)) : String(vb).localeCompare(String(va));
    });
  }

  $('tableCount').textContent = `${tableData.length.toLocaleString()} rows`;
  renderTable();
  renderPagination();
}

function renderTable() {
  const wrap = $('tableWrap');
  const page = STATE.page, size = STATE.pageSize;
  const slice = tableData.slice(page * size, (page + 1) * size);
  const cols = STATE.headers.slice(0, 12);

  let html = '<table><thead><tr>';
  cols.forEach(h => {
    const arrow = sortState.col === h ? (sortState.dir === 'asc' ? '▲' : '▼') : '▲';
    const cls = sortState.col === h ? `sorted-${sortState.dir}` : '';
    html += `<th class="${cls}" data-col="${h}">${h} <span class="sort-arrow">${arrow}</span></th>`;
  });
  html += '</tr></thead><tbody>';
  slice.forEach(row => {
    html += '<tr>' + cols.map(h => `<td title="${row[h] ?? ''}">${row[h] ?? ''}</td>`).join('') + '</tr>';
  });
  html += '</tbody></table>';
  wrap.innerHTML = html;

  wrap.querySelectorAll('th[data-col]').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.dataset.col;
      sortState.dir = sortState.col === col && sortState.dir === 'asc' ? 'desc' : 'asc';
      sortState.col = col;
      buildTable($('tableSearch').value);
    });
  });
}

function renderPagination() {
  const total = Math.ceil(tableData.length / STATE.pageSize);
  const p = $('tablePagination');
  if (total <= 1) { p.innerHTML = ''; return; }
  let html = `<button class="page-btn" onclick="goPage(0)" ${STATE.page === 0 ? 'disabled' : ''}>«</button>`;
  html += `<button class="page-btn" onclick="goPage(${STATE.page - 1})" ${STATE.page === 0 ? 'disabled' : ''}>‹</button>`;
  const start = Math.max(0, STATE.page - 2), end = Math.min(total - 1, STATE.page + 2);
  for (let i = start; i <= end; i++) {
    html += `<button class="page-btn ${i === STATE.page ? 'active' : ''}" onclick="goPage(${i})">${i + 1}</button>`;
  }
  html += `<button class="page-btn" onclick="goPage(${STATE.page + 1})" ${STATE.page >= total - 1 ? 'disabled' : ''}>›</button>`;
  html += `<button class="page-btn" onclick="goPage(${total - 1})" ${STATE.page >= total - 1 ? 'disabled' : ''}>»</button>`;
  p.innerHTML = html;
}

window.goPage = function(p) {
  STATE.page = Math.max(0, Math.min(p, Math.ceil(tableData.length / STATE.pageSize) - 1));
  renderTable(); renderPagination();
};

$('tableSearch').addEventListener('input', e => {
  STATE.page = 0;
  buildTable(e.target.value);
});

// ── DEEP ANALYSIS ──────────────────────────────────────────
$('runAnalysis').addEventListener('click', () => {
  showSpinner('Running deep analysis…');
  setTimeout(() => {
    runDeepAnalysis();
    hideSpinner();
  }, 200);
});

function runDeepAnalysis() {
  const d = STATE.filteredData;
  const numC = numCols();
  let html = '';

  numC.slice(0, 6).forEach(col => {
    const vals = d.map(r => parseFloat(r[col])).filter(v => !isNaN(v));
    if (!vals.length) return;
    const mn = mean(vals), med = median(vals), s = std(vals);
    const q1 = vals.sort((a,b)=>a-b)[Math.floor(vals.length*0.25)];
    const q3 = vals[Math.floor(vals.length*0.75)];
    html += `<div class="analysis-section-title">📐 ${col}</div>`;
    html += stat('Count', vals.length.toLocaleString());
    html += stat('Sum', fmtFull(vals.reduce((a,b)=>a+b,0)));
    html += stat('Mean', fmtFull(mn));
    html += stat('Median', fmtFull(med));
    html += stat('Std Dev', fmtFull(s));
    html += stat('Min', fmtFull(Math.min(...vals)));
    html += stat('Max', fmtFull(Math.max(...vals)));
    html += stat('Q1 / Q3', `${fmtFull(q1)} / ${fmtFull(q3)}`);
    html += stat('Skewness', ((3*(mn-med))/s).toFixed(3));
  });

  // Correlations
  if (numC.length >= 2) {
    html += `<div class="analysis-section-title">🔗 Correlations</div>`;
    for (let i = 0; i < Math.min(numC.length, 4); i++) {
      for (let j = i + 1; j < Math.min(numC.length, 4); j++) {
        const xs = d.map(r => parseFloat(r[numC[i]])).filter(v=>!isNaN(v));
        const ys = d.map(r => parseFloat(r[numC[j]])).filter(v=>!isNaN(v));
        const n = Math.min(xs.length, ys.length);
        const r = correlation(xs.slice(0,n), ys.slice(0,n));
        const strength = Math.abs(r) > 0.7 ? '🔴 Strong' : Math.abs(r) > 0.4 ? '🟡 Moderate' : '🟢 Weak';
        html += stat(`${numC[i]} ↔ ${numC[j]}`, `${r.toFixed(3)} (${strength})`);
      }
    }
  }

  // Missing values
  html += `<div class="analysis-section-title">⚠️ Data Quality</div>`;
  STATE.headers.forEach(h => {
    const missing = d.filter(r => r[h] === null || r[h] === '' || r[h] === undefined).length;
    if (missing > 0) html += stat(`${h} (missing)`, `${missing} (${(missing/d.length*100).toFixed(1)}%)`);
  });
  if (!html.includes('missing')) html += stat('Missing Values', '0 — Clean dataset ✓');

  $('analysisBody').innerHTML = html || '<div class="analysis-placeholder">Not enough numeric data for analysis.</div>';
  STATE.analysisRun = true;
}

function stat(label, value) {
  return `<div class="analysis-stat"><span class="analysis-stat-label">${label}</span><span class="analysis-stat-value">${value}</span></div>`;
}

// ── AI INSIGHTS ────────────────────────────────────────────
$('runInsights').addEventListener('click', () => {
  showSpinner('Generating insights…');
  setTimeout(() => {
    generateInsights();
    hideSpinner();
  }, 300);
});

function generateInsights() {
  const d = STATE.filteredData;
  const numC = numCols();
  const strC = strCols();
  const insights = [];

  // Top performer
  if (strC.length && numC.length) {
    const agg = groupSum(d, strC[0], numC[0], 20).sort((a,b) => b.val - a.val);
    if (agg.length >= 2) {
      const top = agg[0], bot = agg[agg.length - 1];
      const ratio = (top.val / Math.max(bot.val, 1)).toFixed(1);
      insights.push({
        icon: '🏆', priority: 'high',
        text: `<strong>${top.key}</strong> leads <em>${strC[0]}</em> with ${fmtFull(top.val)} in <em>${numC[0]}</em> — <strong>${ratio}×</strong> more than the lowest performer (${bot.key}).`
      });
    }
  }

  // Trend direction
  if (numC.length) {
    const half = Math.floor(d.length / 2);
    const h1 = mean(d.slice(0, half).map(r => parseFloat(r[numC[0]]) || 0));
    const h2 = mean(d.slice(half).map(r => parseFloat(r[numC[0]]) || 0));
    const chg = ((h2 / h1 - 1) * 100).toFixed(1);
    const dir = h2 > h1 ? 'upward 📈' : 'downward 📉';
    insights.push({
      icon: h2 > h1 ? '📈' : '📉', priority: Math.abs(parseFloat(chg)) > 10 ? 'high' : 'med',
      text: `<strong>${numC[0]}</strong> shows a <strong>${dir}</strong> trend — second half of data averages <strong>${chg > 0 ? '+' : ''}${chg}%</strong> vs first half.`
    });
  }

  // Data volume
  const pct = ((d.length / STATE.rawData.length) * 100).toFixed(0);
  insights.push({
    icon: '🔍', priority: 'low',
    text: `Current view shows <strong>${d.length.toLocaleString()}</strong> records (<strong>${pct}%</strong> of total dataset). ${pct < 100 ? 'Filters are active — try clearing to see full picture.' : 'No filters applied — full dataset in view.'}`
  });

  // Outlier detection
  if (numC.length) {
    const vals = d.map(r => parseFloat(r[numC[0]])).filter(v => !isNaN(v));
    const m = mean(vals), s = std(vals);
    const outliers = vals.filter(v => Math.abs(v - m) > 2.5 * s);
    if (outliers.length) {
      insights.push({
        icon: '⚠️', priority: 'med',
        text: `Detected <strong>${outliers.length}</strong> potential outliers in <em>${numC[0]}</em> (values beyond 2.5σ from mean of ${fmtFull(m)}). These may warrant investigation.`
      });
    }
  }

  // Category concentration
  if (strC.length) {
    const counts = {};
    d.forEach(r => { const v = r[strC[0]] || 'Unknown'; counts[v] = (counts[v]||0)+1; });
    const total = d.length;
    const topEntry = Object.entries(counts).sort((a,b)=>b[1]-a[1])[0];
    if (topEntry) {
      const conc = (topEntry[1] / total * 100).toFixed(1);
      const badge = conc > 40 ? 'high' : conc > 20 ? 'med' : 'low';
      insights.push({
        icon: '🎯', priority: badge,
        text: `<strong>${topEntry[0]}</strong> dominates <em>${strC[0]}</em> with <strong>${conc}%</strong> concentration. ${conc > 40 ? 'High dependency — diversification recommended.' : 'Healthy distribution across categories.'}`
      });
    }
  }

  // Numeric correlation recommendation
  if (numC.length >= 2) {
    const xs = d.map(r => parseFloat(r[numC[0]])).filter(v=>!isNaN(v));
    const ys = d.map(r => parseFloat(r[numC[1]])).filter(v=>!isNaN(v));
    const n = Math.min(xs.length, ys.length);
    const r = Math.abs(correlation(xs.slice(0,n), ys.slice(0,n)));
    if (r > 0.6) {
      insights.push({
        icon: '🔗', priority: 'med',
        text: `Strong correlation detected between <em>${numC[0]}</em> and <em>${numC[1]}</em> (r = ${r.toFixed(2)}). Consider using one as a predictor for the other in forecasting models.`
      });
    }
  }

  const html = insights.map(ins => `
    <div class="insight-item">
      <div class="insight-bullet">${ins.icon}</div>
      <div class="insight-text">${ins.text}<br/><span class="insight-badge ${ins.priority}">${ins.priority.toUpperCase()} PRIORITY</span></div>
    </div>
  `).join('');

  $('insightsBody').innerHTML = html || '<div class="analysis-placeholder">Not enough data to generate insights.</div>';
  STATE.insightsRun = true;
}

// ── REFRESH ────────────────────────────────────────────────
$('refreshBtn').addEventListener('click', () => {
  showSpinner('Refreshing…');
  setTimeout(() => {
    buildAllCharts();
    buildKPIs();
    buildTable();
    hideSpinner();
  }, 200);
});

// ── REPORT ─────────────────────────────────────────────────
$('reportBtn').addEventListener('click', () => $('reportModal').classList.remove('hidden'));
$('closeModal').addEventListener('click', () => $('reportModal').classList.add('hidden'));
$('genPdfBtn').addEventListener('click', generateReport);

async function generateReport() {
  $('reportModal').classList.add('hidden');
  showSpinner('Generating PDF report…');
  await delay(100);

  // Run analysis/insights if not done
  if (!STATE.analysisRun) runDeepAnalysis();
  if (!STATE.insightsRun) generateInsights();
  await delay(100);

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
  const W = 210, H = 297;
  let y = 0;

  const addPage = () => { doc.addPage(); y = 20; };
  const checkY = (needed = 20) => { if (y + needed > H - 15) addPage(); };

  // Cover page
  doc.setFillColor(10, 11, 14); doc.rect(0, 0, W, H, 'F');
  doc.setFillColor(245, 166, 35); doc.rect(0, 0, 6, H, 'F');
  doc.setTextColor(245, 166, 35); doc.setFont('helvetica', 'bold'); doc.setFontSize(28);
  doc.text($('reportTitle').value || 'Analytics Report', 20, 60);
  doc.setTextColor(136, 146, 168); doc.setFontSize(12); doc.setFont('helvetica', 'normal');
  doc.text(`File: ${STATE.fileName}`, 20, 80);
  doc.text(`Records: ${STATE.filteredData.length.toLocaleString()} of ${STATE.rawData.length.toLocaleString()}`, 20, 90);
  doc.text(`Generated: ${new Date().toLocaleString()}`, 20, 100);
  doc.setTextColor(85, 94, 117); doc.setFontSize(10);
  doc.text('Powered by DataPulse Analytics', 20, H - 20);

  // KPI Page
  if ($('rptKpi').checked) {
    addPage();
    doc.setTextColor(245,166,35); doc.setFont('helvetica','bold'); doc.setFontSize(16);
    doc.text('KPI Summary', 20, y); y += 12;
    const cards = document.querySelectorAll('.kpi-card');
    cards.forEach(card => {
      checkY(20);
      const label = card.querySelector('.kpi-label')?.textContent || '';
      const value = card.querySelector('.kpi-value')?.textContent || '';
      const sub   = card.querySelector('.kpi-sub')?.textContent || '';
      doc.setFillColor(24,28,36); doc.roundedRect(20, y, W-40, 16, 2, 2, 'F');
      doc.setTextColor(232,234,242); doc.setFont('helvetica','bold'); doc.setFontSize(11);
      doc.text(label, 26, y+7);
      doc.setTextColor(245,166,35); doc.setFontSize(12);
      doc.text(value, W-26, y+7, { align: 'right' });
      doc.setTextColor(136,146,168); doc.setFontSize(9);
      doc.text(sub, 26, y+13);
      y += 20;
    });
  }

  // Analysis page
  if ($('rptAnalysis').checked && STATE.analysisRun) {
    addPage();
    doc.setTextColor(245,166,35); doc.setFont('helvetica','bold'); doc.setFontSize(16);
    doc.text('Deep Analysis', 20, y); y += 10;
    const items = $('analysisBody').querySelectorAll('.analysis-stat');
    items.forEach(item => {
      checkY(8);
      const lbl = item.querySelector('.analysis-stat-label')?.textContent || '';
      const val = item.querySelector('.analysis-stat-value')?.textContent || '';
      doc.setTextColor(136,146,168); doc.setFont('helvetica','normal'); doc.setFontSize(9);
      doc.text(lbl, 22, y);
      doc.setTextColor(232,234,242);
      doc.text(val, W-22, y, { align: 'right' });
      y += 7;
    });
  }

  // Insights page
  if ($('rptInsights').checked && STATE.insightsRun) {
    addPage();
    doc.setTextColor(0,212,170); doc.setFont('helvetica','bold'); doc.setFontSize(16);
    doc.text('AI Insights', 20, y); y += 10;
    const items = $('insightsBody').querySelectorAll('.insight-item');
    items.forEach(item => {
      const txt = item.querySelector('.insight-text')?.innerText || '';
      checkY(24);
      doc.setFillColor(24,28,36); doc.roundedRect(20, y, W-40, 20, 2, 2, 'F');
      doc.setTextColor(232,234,242); doc.setFont('helvetica','normal'); doc.setFontSize(9);
      const lines = doc.splitTextToSize(txt.replace(/\n/g,' '), W-50);
      doc.text(lines.slice(0,2), 26, y+7);
      y += 24;
    });
  }

  // Data table
  if ($('rptTable').checked) {
    addPage();
    doc.setTextColor(245,166,35); doc.setFont('helvetica','bold'); doc.setFontSize(16);
    doc.text('Data Preview (top 50)', 20, y); y += 6;
    const cols = STATE.headers.slice(0, 8);
    const rows = STATE.filteredData.slice(0, 50).map(r => cols.map(c => String(r[c] ?? '')));
    doc.autoTable({
      startY: y + 4, head: [cols], body: rows,
      theme: 'grid', pageBreak: 'auto',
      headStyles: { fillColor: [30, 35, 48], textColor: [245,166,35], fontSize: 8, fontStyle: 'bold' },
      bodyStyles: { fillColor: [24,28,36], textColor: [136,146,168], fontSize: 7 },
      alternateRowStyles: { fillColor: [20,24,32] },
      styles: { cellPadding: 2 },
    });
  }

  doc.save(`datapulse_report_${Date.now()}.pdf`);
  hideSpinner();
}

// ── INIT ───────────────────────────────────────────────────
(function init() {
  showScreen('upload');
})();
