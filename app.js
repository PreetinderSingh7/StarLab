// ── APP STATE ────────────────────────────────────────
const AppState = {
  sheets: { sheet1: null, sheet2: null, sheet3: null },
  headers: { sheet1: [], sheet2: [], sheet3: [] },
  filename: '',
  analysed: false
};

// ── PAGES ────────────────────────────────────────────
const PAGES = {
  home:    { title: 'Home & Guide',             render: renderHome },
  upload:  { title: 'Upload Data',              render: renderUpload },
  sheet1:  { title: 'Sheet 1 — Demographics',   render: () => renderSheetPage('sheet1') },
  sheet2:  { title: 'Sheet 2 — Scale Analysis', render: () => renderSheetPage('sheet2') },
  sheet3:  { title: 'Sheet 3 — Scale Analysis', render: () => renderSheetPage('sheet3') },
  chisq:   { title: 'Chi-Square Tests',         render: renderChiSquarePage },
  charts:  { title: 'All Charts',               render: renderChartsPage }
};

// ── ROUTER ───────────────────────────────────────────
function navigate(page) {
  if (!PAGES[page]) page = 'home';
  document.querySelectorAll('.nav-item').forEach(el =>
    el.classList.toggle('active', el.dataset.page === page));
  document.getElementById('pageTitle').textContent = PAGES[page].title;
  const main = document.getElementById('mainContent');
  main.innerHTML = '';
  const div = document.createElement('div');
  div.className = 'fade-in';
  div.innerHTML = PAGES[page].render();
  main.appendChild(div);
  attachPageEvents(page);
  window._currentPage = page;
  const hasData = AppState.sheets.sheet1 || AppState.sheets.sheet2 || AppState.sheets.sheet3;
  document.getElementById('exportBtn').style.display = hasData ? 'flex' : 'none';
}

document.querySelectorAll('.nav-item').forEach(el =>
  el.addEventListener('click', e => { e.preventDefault(); navigate(el.dataset.page); }));

document.getElementById('menuToggle').addEventListener('click', () =>
  document.getElementById('sidebar').classList.toggle('open'));

// ── HOME PAGE ─────────────────────────────────────────
function renderHome() {
  return `
  <div class="page-header">
    <h1>Welcome to StatLab Pro</h1>
    <p>Your complete data analysis toolkit — Frequency, Descriptives, Chi-Square & Charts</p>
  </div>
  <div class="grid-2" style="margin-bottom:28px">
    <div class="card" style="border-color:#2a3a6a">
      <div style="font-size:36px;margin-bottom:12px">📊</div>
      <div class="card-title" style="margin-bottom:8px">What this tool does</div>
      <p style="color:var(--text2);font-size:13.5px;line-height:1.7">
        Upload an Excel file with up to 3 sheets. Get instant frequency tables,
        percentages, mean, <strong>median</strong>, standard deviation, above/below average
        categorisation, chi-square tests, degrees of freedom, p-values, and
        editable charts. Export everything as a <strong>Word document</strong> with formatted tables.
      </p>
    </div>
    <div class="card" style="border-color:#1e4a38">
      <div style="font-size:36px;margin-bottom:12px">🚀</div>
      <div class="card-title" style="margin-bottom:8px">Get started in 3 steps</div>
      <ol style="color:var(--text2);font-size:13.5px;line-height:2;padding-left:18px">
        <li>Upload your Excel file (.xlsx)</li>
        <li>Map your sheets to Sheet 1 / 2 / 3</li>
        <li>Click any analysis section in the sidebar</li>
      </ol>
      <br>
      <a href="#" onclick="navigate('upload');return false" class="btn btn-success btn-sm">Upload Data Now →</a>
    </div>
  </div>

  <div class="section-title">How Above / Below Average is Calculated</div>
  <div class="card" style="margin-bottom:24px;border-left:3px solid var(--accent)">
    <div class="card-title">Mean ± 1 SD Method</div>
    <p style="font-size:13.5px;color:var(--text2);line-height:1.8">
      Each respondent's scores across all numeric columns are summed to give a <strong>Total Score</strong>.
      The <strong>Grand Mean (M)</strong> and <strong>Standard Deviation (σ)</strong> of these totals are computed.<br><br>
      • <span class="badge-above">Above Average</span> → Total Score &gt; M + σ<br>
      • <span class="badge-avg">Average</span> → M − σ ≤ Total Score ≤ M + σ<br>
      • <span class="badge-below">Below Average</span> → Total Score &lt; M − σ<br><br>
      This is the standard <strong>Mean ± 1 SD</strong> method widely used in Indian university research theses for Likert scale data.
    </p>
  </div>

  <div class="section-title">How to use each section</div>
  <div class="guide-step"><div class="step-num">1</div><div>
    <h4>Upload Data</h4>
    <p>Upload a .xlsx file. The app reads up to 3 sheets automatically. Headers must be in the first row.</p>
  </div></div>
  <div class="guide-step"><div class="step-num">2</div><div>
    <h4>Sheet 1 — Demographic Variables</h4>
    <p>Calculates frequency and percentage for every column. Choose a column to see its breakdown, then pick your chart type (Bar, Pie, Doughnut, etc.) and edit values before downloading.</p>
  </div></div>
  <div class="guide-step"><div class="step-num">3</div><div>
    <h4>Sheet 2 & 3 — Scale / Likert Data</h4>
    <p>Calculates frequency, percentage, mean, <strong>median</strong>, SD. Respondents are categorised Above/Average/Below using Mean ± 1 SD. Sheet 3 can be divided by Sheet 2 groupings.</p>
  </div></div>
  <div class="guide-step"><div class="step-num">4</div><div>
    <h4>Chi-Square Tests</h4>
    <p>Select any two columns to run chi-square. Results show χ², df, p-value, contingency table, and interpretation.</p>
  </div></div>
  <div class="guide-step"><div class="step-num">5</div><div>
    <h4>Export to Word</h4>
    <p>Click <strong>Export Results</strong> to download a .docx file with all tables formatted exactly like your thesis — with frequency, percentage, mean, median, SD columns.</p>
  </div></div>
  <div class="alert alert-warn" style="margin-top:16px">
    ⚠ All analysis runs locally in your browser. No data is sent to any server.
  </div>`;
}

// ── UPLOAD PAGE ───────────────────────────────────────
function renderUpload() {
  const hasData = AppState.sheets.sheet1 || AppState.sheets.sheet2;
  return `
  <div class="page-header">
    <h1>Upload Your Data</h1>
    <p>Upload an Excel (.xlsx) file with 1–3 sheets</p>
  </div>
  ${hasData ? `<div class="alert alert-success">✓ Data loaded: <strong>${AppState.filename}</strong>. Re-upload to replace.</div>` : ''}
  <div class="upload-zone" id="uploadZone" onclick="document.getElementById('fileInput').click()">
    <div class="upload-icon">📂</div>
    <h3>Click to upload or drag & drop</h3>
    <p>Supports .xlsx and .xls files • Max 10MB</p>
    <br>
    <span class="btn btn-primary">Choose File</span>
  </div>
  <input type="file" id="fileInput" accept=".xlsx,.xls" style="display:none" />
  <div id="uploadStatus" style="margin-top:20px"></div>
  ${hasData ? renderSheetMapper() : ''}`;
}

function renderSheetMapper() {
  const sheetNames = AppState._sheetNames || [];
  const opts = sheetNames.map((s,i) => `<option value="${i}">${s}</option>`).join('');
  const noneOpt = `<option value="-1">— not used —</option>`;
  return `
  <div class="card" style="margin-top:24px">
    <div class="card-title">Map Sheets to Analysis Slots</div>
    <div class="grid-3">
      <div>
        <label style="font-size:12px;color:var(--text3);display:block;margin-bottom:6px">SHEET 1 — Demographics</label>
        <select id="mapSheet1" onchange="AppUpload.remapSheets()">${noneOpt}${opts}</select>
      </div>
      <div>
        <label style="font-size:12px;color:var(--text3);display:block;margin-bottom:6px">SHEET 2 — Scale A</label>
        <select id="mapSheet2" onchange="AppUpload.remapSheets()">${noneOpt}${opts}</select>
      </div>
      <div>
        <label style="font-size:12px;color:var(--text3);display:block;margin-bottom:6px">SHEET 3 — Scale B</label>
        <select id="mapSheet3" onchange="AppUpload.remapSheets()">${noneOpt}${opts}</select>
      </div>
    </div>
    <div id="mapStatus" style="margin-top:12px;font-size:13px;color:var(--accent3)"></div>
  </div>`;
}

// ── SHEET PAGE ────────────────────────────────────────
function renderSheetPage(sheetKey) {
  const data = AppState.sheets[sheetKey];
  if (!data || data.length === 0) {
    return `
    <div class="page-header"><h1>${PAGES[sheetKey].title}</h1></div>
    <div class="empty-state">
      <div class="empty-icon">📋</div>
      <p>No data loaded for this sheet yet.</p><br>
      <a href="#" onclick="navigate('upload');return false" class="btn btn-primary">Upload Data →</a>
    </div>`;
  }

  const headers = AppState.headers[sheetKey];
  const isDemographic = sheetKey === 'sheet1';
  const colOpts = headers.map(h => `<option value="${h}">${h}</option>`).join('');
  const summaryStats = isDemographic ? '' : AppDescriptive.renderSummaryStats(data, headers);

  return `
  <div class="page-header">
    <h1>${PAGES[sheetKey].title}</h1>
    <p>${data.length} respondents · ${headers.length} variables</p>
  </div>

  ${summaryStats}
  ${!isDemographic ? renderAboveBelowSection(sheetKey) : ''}

  <div class="section-title">Column Analysis</div>
  <div style="display:flex;gap:12px;align-items:center;margin-bottom:20px;flex-wrap:wrap">
    <select id="colSelect_${sheetKey}" style="min-width:200px">${colOpts}</select>
    <button class="btn btn-primary btn-sm" onclick="onColChange('${sheetKey}')">Analyse ▶</button>
    ${!isDemographic ? `<button class="btn btn-sm" onclick="AppDescriptive.analyseAll('${sheetKey}')">Analyse All Columns</button>` : ''}
    <button class="btn btn-sm" onclick="AppExport.exportSheetWord('${sheetKey}')">⬇ Export Sheet to Word</button>
  </div>
  <div id="colAnalysis_${sheetKey}"></div>
  ${!isDemographic && AppState.sheets.sheet2 && sheetKey === 'sheet3' ? renderDivisionSection() : ''}`;
}

function renderAboveBelowSection(sheetKey) {
  const data  = AppState.sheets[sheetKey];
  const headers = AppState.headers[sheetKey];
  const result = AppDescriptive.aboveBelowAvg(data, headers);
  return `
  <div class="card" style="margin-bottom:24px">
    <div class="card-title">Respondent Categorisation <span class="badge">Mean ± 1 SD Method</span></div>
    <div style="font-size:12.5px;color:var(--text3);margin-bottom:14px;line-height:1.7">
      Each respondent's total score is computed. Grand Mean = <strong style="color:var(--accent)">${result.mean.toFixed(2)}</strong>,
      SD = <strong style="color:var(--accent)">${result.sd.toFixed(2)}</strong>.
      Above = score &gt; ${(result.mean+result.sd).toFixed(2)} |
      Average = ${(result.mean-result.sd).toFixed(2)} to ${(result.mean+result.sd).toFixed(2)} |
      Below = score &lt; ${(result.mean-result.sd).toFixed(2)}
    </div>
    <div class="stat-grid">
      <div class="stat-pill"><span class="stat-val" style="color:var(--accent3)">${result.above}</span><span class="stat-lbl">Above Average</span></div>
      <div class="stat-pill"><span class="stat-val">${result.average}</span><span class="stat-lbl">Average</span></div>
      <div class="stat-pill"><span class="stat-val" style="color:var(--danger)">${result.below}</span><span class="stat-lbl">Below Average</span></div>
      <div class="stat-pill"><span class="stat-val">${result.mean.toFixed(2)}</span><span class="stat-lbl">Grand Mean</span></div>
      <div class="stat-pill"><span class="stat-val">${result.sd.toFixed(2)}</span><span class="stat-lbl">Std Deviation</span></div>
    </div>
    <div class="chart-box" style="margin-top:14px">
      <canvas id="aboveBelow_${sheetKey}" height="180"></canvas>
    </div>
  </div>`;
}

function renderDivisionSection() {
  const s2Headers = AppState.headers.sheet2 || [];
  const opts = s2Headers.map(h => `<option value="${h}">${h}</option>`).join('');
  return `
  <div class="card" style="margin-top:24px">
    <div class="card-title">Divide Sheet 3 by Sheet 2 Grouping</div>
    <p style="font-size:13px;color:var(--text2);margin-bottom:14px">
      Select a grouping column from Sheet 2 to cross-tabulate Sheet 3 results by group.
    </p>
    <div style="display:flex;gap:12px;align-items:center;flex-wrap:wrap">
      <select id="divisionCol" style="min-width:200px">${opts}</select>
      <button class="btn btn-primary btn-sm" onclick="AppDescriptive.dividedAnalysis()">Run Division Analysis ▶</button>
    </div>
    <div id="divisionResult" style="margin-top:20px"></div>
  </div>`;
}

// ── CHI-SQUARE PAGE ────────────────────────────────────
function renderChiSquarePage() {
  const allHeaders = [];
  ['sheet1','sheet2','sheet3'].forEach(sk => {
    (AppState.headers[sk]||[]).forEach(h => allHeaders.push({ label:`[${sk}] ${h}`, sheet:sk, col:h }));
  });
  if (!allHeaders.length) {
    return `<div class="page-header"><h1>Chi-Square Tests</h1></div>
    <div class="empty-state"><div class="empty-icon">χ²</div><p>Upload data first.</p><br>
    <a href="#" onclick="navigate('upload');return false" class="btn btn-primary">Upload Data →</a></div>`;
  }
  const opts = allHeaders.map((h,i) => `<option value="${i}">${h.label}</option>`).join('');
  window._chiHeaders = allHeaders;
  return `
  <div class="page-header">
    <h1>Chi-Square Test of Association</h1>
    <p>Test whether two categorical variables are statistically associated</p>
  </div>
  <div class="card" style="margin-bottom:24px">
    <div class="card-title">Select Variables</div>
    <div class="grid-2">
      <div>
        <label style="font-size:12px;color:var(--text3);display:block;margin-bottom:6px">VARIABLE 1</label>
        <select id="chiVar1" style="width:100%">${opts}</select>
      </div>
      <div>
        <label style="font-size:12px;color:var(--text3);display:block;margin-bottom:6px">VARIABLE 2</label>
        <select id="chiVar2" style="width:100%">${opts}</select>
      </div>
    </div>
    <br>
    <button class="btn btn-primary" onclick="AppInferential.runChiSquare()">Run Chi-Square Test ▶</button>
  </div>
  <div id="chiResult"></div>
  <div class="section-title" style="margin-top:32px">Batch Chi-Square — All Sheet 2 vs Sheet 1 Demographics</div>
  <p style="font-size:13px;color:var(--text2);margin-bottom:14px">Run chi-square for every Sheet 2 column vs every Sheet 1 column at once.</p>
  <button class="btn" onclick="AppInferential.batchChiSquare()">Run Batch Test ▶</button>
  <div id="batchChiResult" style="margin-top:20px"></div>`;
}

// ── CHARTS PAGE ────────────────────────────────────────
function renderChartsPage() {
  const hasAny = AppState.sheets.sheet1||AppState.sheets.sheet2||AppState.sheets.sheet3;
  if (!hasAny) {
    return `<div class="page-header"><h1>All Charts</h1></div>
    <div class="empty-state"><div class="empty-icon">◈</div><p>Upload data to generate charts.</p></div>`;
  }
  return `
  <div class="page-header">
    <h1>All Charts</h1>
    <p>Visual overview of all variables across all sheets — select chart type per variable</p>
  </div>
  <div class="tab-bar" id="chartTabBar">
    ${AppState.sheets.sheet1?'<button class="tab-btn active" data-tab="ct1">Sheet 1</button>':''}
    ${AppState.sheets.sheet2?'<button class="tab-btn" data-tab="ct2">Sheet 2</button>':''}
    ${AppState.sheets.sheet3?'<button class="tab-btn" data-tab="ct3">Sheet 3</button>':''}
  </div>
  ${AppState.sheets.sheet1?`<div class="tab-pane active" id="ct1">${AppCharts.renderAllForSheet('sheet1')}</div>`:''}
  ${AppState.sheets.sheet2?`<div class="tab-pane" id="ct2">${AppCharts.renderAllForSheet('sheet2')}</div>`:''}
  ${AppState.sheets.sheet3?`<div class="tab-pane" id="ct3">${AppCharts.renderAllForSheet('sheet3')}</div>`:''}`;
}

// ── EVENT BINDING ──────────────────────────────────────
function attachPageEvents(page) {
  if (page === 'upload') AppUpload.bind();
  if (['sheet1','sheet2','sheet3'].includes(page)) {
    setTimeout(() => {
      onColChange(page);
      if (page !== 'sheet1') {
        const data = AppState.sheets[page];
        const headers = AppState.headers[page];
        if (data) AppCharts.renderAboveBelow(page, data, headers);
      }
    }, 50);
  }
  if (page === 'charts') {
    setTimeout(() => { AppCharts.bindTabs(); AppCharts.renderAllCharts(); }, 50);
  }
}

function onColChange(sheetKey) {
  const sel = document.getElementById(`colSelect_${sheetKey}`);
  if (!sel) return;
  const col = sel.value;
  const data = AppState.sheets[sheetKey];
  const isDemographic = sheetKey === 'sheet1';
  const container = document.getElementById(`colAnalysis_${sheetKey}`);
  if (!container || !data) return;
  const vals = data.map(r => r[col]).filter(v => v !== undefined && v !== null && v !== '');
  container.innerHTML = AppDescriptive.renderColAnalysis(col, vals, isDemographic, sheetKey);
  setTimeout(() => AppCharts.renderEditableChart(`editChart_${col.replace(/\W/g,'_')}`, col, vals, isDemographic), 60);
}

navigate('home');