const AppDescriptive = (() => {

  // ── FREQUENCY ─────────────────────────────────────
  function frequency(values) {
    const counts = {};
    values.forEach(v => { const k = String(v).trim(); counts[k] = (counts[k] || 0) + 1; });
    const total = values.length;
    return Object.entries(counts)
      .sort((a, b) => b[1] - a[1])
      .map(([val, count]) => ({ val, count, pct: ((count / total) * 100).toFixed(1) }));
  }

  // ── NUMERIC STATS ──────────────────────────────────
  function numericStats(values) {
    const nums = values.map(Number).filter(v => !isNaN(v));
    if (!nums.length) return null;
    const n      = nums.length;
    const mean   = nums.reduce((a, b) => a + b, 0) / n;
    const sorted = [...nums].sort((a, b) => a - b);
    const mid    = Math.floor(n / 2);
    const median = n % 2 === 0 ? (sorted[mid-1] + sorted[mid]) / 2 : sorted[mid];
    const variance = nums.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / n;
    const sd     = Math.sqrt(variance);
    return { n, mean, median, sd, min: sorted[0], max: sorted[n-1], variance };
  }

  // ── ABOVE / BELOW AVERAGE (Mean ± 1 SD) ───────────
  function aboveBelowAvg(data, headers) {
    const numericHeaders = headers.filter(h => {
      const sample = data.slice(0, 5).map(r => Number(r[h]));
      return sample.some(v => !isNaN(v));
    });
    const totals = data.map(row =>
      numericHeaders.reduce((s, h) => {
        const v = Number(row[h]); return s + (isNaN(v) ? 0 : v);
      }, 0)
    );
    const stats = numericStats(totals);
    if (!stats) return { above: 0, average: 0, below: 0, mean: 0, median: 0, sd: 0, totals: [] };
    const { mean, median, sd } = stats;
    let above = 0, average = 0, below = 0;
    totals.forEach(t => {
      if (t > mean + sd)      above++;
      else if (t < mean - sd) below++;
      else                    average++;
    });
    return { above, average, below, mean, median, sd, totals };
  }

  // ── COLUMN ANALYSIS HTML ───────────────────────────
  function renderColAnalysis(col, values, isDemographic, sheetKey) {
    const freq     = frequency(values);
    const isNumeric = values.some(v => !isNaN(Number(v)) && v !== '');
    const stats    = isNumeric ? numericStats(values.map(Number).filter(v => !isNaN(v))) : null;
    const safeId   = col.replace(/\W/g, '_');

    // Stats pills
    let statsHtml = '';
    if (!isDemographic && stats) {
      statsHtml = `
      <div class="stat-grid" style="margin-bottom:16px">
        <div class="stat-pill"><span class="stat-val">${stats.mean.toFixed(2)}</span><span class="stat-lbl">Mean</span></div>
        <div class="stat-pill"><span class="stat-val">${stats.median.toFixed(2)}</span><span class="stat-lbl">Median</span></div>
        <div class="stat-pill"><span class="stat-val">${stats.sd.toFixed(2)}</span><span class="stat-lbl">Std Dev</span></div>
        <div class="stat-pill"><span class="stat-val">${stats.min}</span><span class="stat-lbl">Min</span></div>
        <div class="stat-pill"><span class="stat-val">${stats.max}</span><span class="stat-lbl">Max</span></div>
        <div class="stat-pill"><span class="stat-val">${stats.n}</span><span class="stat-lbl">N</span></div>
      </div>`;
    }

    // Frequency table rows
    const tableRows = freq.map(({ val, count, pct }) => {
      let badge = '';
      if (!isDemographic && stats) {
        const num = Number(val);
        if (!isNaN(num)) {
          if      (num > stats.mean) badge = `<span class="badge-above">above avg</span>`;
          else if (num < stats.mean) badge = `<span class="badge-below">below avg</span>`;
          else                       badge = `<span class="badge-avg">at avg</span>`;
        }
      }
      return `<tr>
        <td>${val}</td>
        <td>${count}</td>
        <td>${pct}%</td>
        <td><div class="progress-bar-wrap" style="width:100px"><div class="progress-bar" style="width:${pct}%"></div></div></td>
        ${!isDemographic ? `<td>${badge}</td>` : ''}
      </tr>`;
    }).join('');

    // Edit table rows (for chart editor)
    const editRows = freq.map(({ val, count }) => `
      <tr>
        <td><input type="text"   class="edit-label" value="${val}"   data-id="${safeId}" style="width:130px"/></td>
        <td><input type="number" class="edit-val"   value="${count}" data-id="${safeId}" style="width:80px"/></td>
        <td><button onclick="AppCharts.removeEditRow(this)" style="background:none;border:none;color:var(--danger);cursor:pointer;font-size:14px">✕</button></td>
      </tr>`).join('');

    return `
    ${statsHtml}

    <div class="section-title">${col} — Frequency Distribution</div>
    <div class="table-wrap" style="margin-bottom:20px">
      <table>
        <thead><tr>
          <th>Value</th><th>Frequency</th><th>Percentage</th><th>Distribution</th>
          ${!isDemographic ? '<th>Comparison</th>' : ''}
        </tr></thead>
        <tbody>${tableRows}</tbody>
      </table>
    </div>

    <div class="card" style="margin-bottom:0">
      <div class="card-title">Chart Editor <span class="badge">Editable</span></div>

      <div style="display:flex;gap:12px;align-items:flex-end;flex-wrap:wrap;margin-bottom:16px">
        <div>
          <label style="font-size:11px;color:var(--text3);display:block;margin-bottom:4px">CHART TYPE</label>
          <select id="chartType_${safeId}" style="min-width:150px">
            <option value="bar">Bar Chart</option>
            <option value="pie">Pie Chart</option>
            <option value="doughnut">Doughnut Chart</option>
            <option value="polarArea">Polar Area</option>
            <option value="line">Line Chart</option>
            <option value="radar">Radar Chart</option>
          </select>
        </div>
        <div>
          <label style="font-size:11px;color:var(--text3);display:block;margin-bottom:4px">CHART TITLE</label>
          <input type="text" id="chartTitle_${safeId}" value="${col}" style="width:200px"/>
        </div>
        <button class="btn btn-sm btn-primary"
          onclick="AppCharts.refreshEditableChart('${safeId}','${col.replace(/'/g,"\\'")}')">
          ▶ Update Chart
        </button>
        <button class="btn btn-sm btn-success"
          onclick="AppCharts.downloadChart('${safeId}','${col.replace(/'/g,"\\'")}')">
          ⬇ Download PNG
        </button>
      </div>

      <div class="grid-2" style="gap:20px;align-items:start">
        <div>
          <div class="section-title" style="margin-bottom:8px">
            Edit Labels &amp; Values
            <button class="btn btn-sm" style="float:right;margin-top:-4px"
              onclick="AppCharts.addEditRow('${safeId}')">+ Add Row</button>
          </div>
          <div class="table-wrap" style="margin-bottom:10px">
            <table>
              <thead><tr><th>Label (Editable)</th><th>Value</th><th></th></tr></thead>
              <tbody id="editTable_${safeId}">${editRows}</tbody>
            </table>
          </div>
          <button class="btn btn-sm btn-primary"
            onclick="AppCharts.refreshEditableChart('${safeId}','${col.replace(/'/g,"\\'")}')">
            ▶ Apply &amp; Refresh Chart
          </button>
        </div>
        <div class="chart-box">
          <canvas id="editChart_${safeId}"></canvas>
        </div>
      </div>
    </div>`;
  }

  // ── SUMMARY STATS CARD ─────────────────────────────
  function renderSummaryStats(data, headers) {
    const numericHeaders = headers.filter(h => {
      const v = Number(data[0]?.[h]);
      return !isNaN(v) && data[0]?.[h] !== '';
    });
    if (!numericHeaders.length) return '';
    const allVals = data.flatMap(row =>
      numericHeaders.map(h => Number(row[h])).filter(v => !isNaN(v))
    );
    const stats = numericStats(allVals);
    if (!stats) return '';
    return `
    <div class="card" style="margin-bottom:24px">
      <div class="card-title">Sheet Summary Statistics</div>
      <div class="stat-grid">
        <div class="stat-pill"><span class="stat-val">${data.length}</span><span class="stat-lbl">Respondents</span></div>
        <div class="stat-pill"><span class="stat-val">${numericHeaders.length}</span><span class="stat-lbl">Numeric Cols</span></div>
        <div class="stat-pill"><span class="stat-val">${stats.mean.toFixed(2)}</span><span class="stat-lbl">Grand Mean</span></div>
        <div class="stat-pill"><span class="stat-val">${stats.median.toFixed(2)}</span><span class="stat-lbl">Grand Median</span></div>
        <div class="stat-pill"><span class="stat-val">${stats.sd.toFixed(2)}</span><span class="stat-lbl">Grand SD</span></div>
      </div>
    </div>`;
  }

  // ── ANALYSE ALL COLUMNS ────────────────────────────
  function analyseAll(sheetKey) {
    const data    = AppState.sheets[sheetKey];
    const headers = AppState.headers[sheetKey];
    if (!data) return;
    const container = document.getElementById(`colAnalysis_${sheetKey}`);
    if (!container) return;
    let html = '';
    headers.forEach(h => {
      const vals = data.map(r => r[h]).filter(v => v !== undefined && v !== null && v !== '');
      html += `<div class="card" style="margin-bottom:20px">
        <div class="card-title">${h}</div>
        ${renderColAnalysis(h, vals, false, sheetKey)}
      </div>`;
    });
    container.innerHTML = html;
    setTimeout(() => {
      headers.forEach(h => {
        const vals = data.map(r => r[h]).filter(v => v !== undefined && v !== null && v !== '');
        AppCharts.renderEditableChart(`editChart_${h.replace(/\W/g,'_')}`, h, vals, false);
      });
    }, 60);
  }

  // ── DIVIDED ANALYSIS (Sheet3 ÷ Sheet2) ────────────
  function dividedAnalysis() {
    const groupCol = document.getElementById('divisionCol')?.value;
    const data2    = AppState.sheets.sheet2;
    const data3    = AppState.sheets.sheet3;
    const headers3 = AppState.headers.sheet3;
    if (!groupCol || !data2 || !data3) return;
    const groups = {};
    data2.forEach((row2, i) => {
      const grpVal = String(row2[groupCol] ?? 'Unknown');
      if (!groups[grpVal]) groups[grpVal] = [];
      if (data3[i]) groups[grpVal].push(data3[i]);
    });
    let html = '';
    Object.entries(groups).forEach(([grp, rows]) => {
      if (!rows.length) return;
      const numH  = headers3.filter(h => !isNaN(Number(rows[0]?.[h])));
      const stats = numH.map(h => {
        const vals = rows.map(r => Number(r[h])).filter(v => !isNaN(v));
        return { h, ...numericStats(vals) };
      });
      html += `
      <div class="card" style="margin-bottom:16px">
        <div class="card-title">${groupCol}: <em>${grp}</em> <span class="badge">n=${rows.length}</span></div>
        <div class="table-wrap"><table>
          <thead><tr><th>Column</th><th>Mean</th><th>Median</th><th>SD</th><th>Min</th><th>Max</th></tr></thead>
          <tbody>${stats.map(s => `<tr>
            <td>${s.h}</td>
            <td>${s.mean?.toFixed(2)??'-'}</td>
            <td>${s.median?.toFixed(2)??'-'}</td>
            <td>${s.sd?.toFixed(2)??'-'}</td>
            <td>${s.min??'-'}</td>
            <td>${s.max??'-'}</td>
          </tr>`).join('')}</tbody>
        </table></div>
      </div>`;
    });
    document.getElementById('divisionResult').innerHTML =
      html || '<p style="color:var(--text3)">No groups found.</p>';
  }

  return {
    frequency, numericStats, aboveBelowAvg,
    renderColAnalysis, renderSummaryStats,
    analyseAll, dividedAnalysis
  };
})();