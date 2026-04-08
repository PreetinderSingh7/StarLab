const AppCharts = (() => {

  const _instances = {};

  const PALETTE = [
    '#4f8ef7','#36c5a0','#f0a500','#e05252','#7b5ea7',
    '#4bc9f0','#a3e05c','#f47ef7','#f7a04f','#5ec4f7'
  ];

  function destroyChart(id) {
    if (_instances[id]) { _instances[id].destroy(); delete _instances[id]; }
  }

  // ── READ EDITABLE DATA FROM TABLE ─────────────────
  function getEditData(safeId) {
    const table = document.getElementById(`editTable_${safeId}`);
    if (!table) return { labels: [], data: [] };
    const labels = [], data = [];
    table.querySelectorAll('tr').forEach(tr => {
      const lInput = tr.querySelector('.edit-label');
      const vInput = tr.querySelector('.edit-val');
      if (lInput && vInput) {
        labels.push(lInput.value.trim() || '?');
        data.push(parseFloat(vInput.value) || 0);
      }
    });
    return { labels, data };
  }

  // ── RENDER EDITABLE CHART (called on col change) ──
  function renderEditableChart(canvasId, col, values, isDemographic) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    const safeId = canvasId.replace('editChart_', '');
    const freq   = AppDescriptive.frequency(values);
    const labels = freq.map(f => f.val);
    const data   = freq.map(f => f.count);

    // Populate edit table
    const editTableBody = document.getElementById(`editTable_${safeId}`);
    if (editTableBody) {
      editTableBody.innerHTML = freq.map(({ val, count }) => `
        <tr>
          <td><input type="text"   class="edit-label" value="${val}"   data-id="${safeId}" style="width:130px"/></td>
          <td><input type="number" class="edit-val"   value="${count}" data-id="${safeId}" style="width:80px"/></td>
          <td><button onclick="AppCharts.removeEditRow(this)" style="background:none;border:none;color:var(--danger);cursor:pointer;font-size:14px">✕</button></td>
        </tr>`).join('');
    }

    const typeSel    = document.getElementById(`chartType_${safeId}`);
    const titleInput = document.getElementById(`chartTitle_${safeId}`);
    const type       = typeSel   ? typeSel.value   : 'bar';
    const titleText  = titleInput ? titleInput.value : col;

    destroyChart(canvasId);
    _instances[canvasId] = buildChart(canvas, type, titleText, labels, data, col);
  }

  // ── REFRESH (called by Update button) ─────────────
  function refreshEditableChart(safeId, col) {
    const canvasId   = `editChart_${safeId}`;
    const canvas     = document.getElementById(canvasId);
    if (!canvas) return;
    const typeSel    = document.getElementById(`chartType_${safeId}`);
    const titleInput = document.getElementById(`chartTitle_${safeId}`);
    const type       = typeSel   ? typeSel.value   : 'bar';
    const titleText  = titleInput ? titleInput.value : col;
    const { labels, data } = getEditData(safeId);
    destroyChart(canvasId);
    _instances[canvasId] = buildChart(canvas, type, titleText, labels, data, col);
  }

  // ── ADD / REMOVE EDIT ROWS ─────────────────────────
  function addEditRow(safeId) {
    const table = document.getElementById(`editTable_${safeId}`);
    if (!table) return;
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><input type="text"   class="edit-label" value="New Label" data-id="${safeId}" style="width:130px"/></td>
      <td><input type="number" class="edit-val"   value="0"         data-id="${safeId}" style="width:80px"/></td>
      <td><button onclick="AppCharts.removeEditRow(this)" style="background:none;border:none;color:var(--danger);cursor:pointer;font-size:14px">✕</button></td>`;
    table.appendChild(tr);
  }

  function removeEditRow(btn) {
    btn.closest('tr').remove();
  }

  // ── DOWNLOAD PNG ───────────────────────────────────
  function downloadChart(safeId, col) {
    const canvasId = `editChart_${safeId}`;
    const chart    = _instances[canvasId];
    if (!chart) { alert('Generate the chart first.'); return; }

    // Temporarily set white background for download
    const canvas = document.getElementById(canvasId);
    const ctx    = canvas.getContext('2d');
    const w = canvas.width, h = canvas.height;
    const imageData = ctx.getImageData(0, 0, w, h);
    ctx.globalCompositeOperation = 'destination-over';
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, w, h);
    const img = canvas.toDataURL('image/png', 1.0);
    ctx.clearRect(0, 0, w, h);
    ctx.putImageData(imageData, 0, 0);

    const a = document.createElement('a');
    a.href     = img;
    a.download = `${(col||'chart').replace(/\W/g, '_')}_chart.png`;
    a.click();
  }

  // ── CHART BUILDER ──────────────────────────────────
  function buildChart(canvas, type, titleText, labels, data, col) {
    const isRound = ['pie','doughnut','polarArea','radar'].includes(type);
    return new Chart(canvas, {
      type,
      data: {
        labels,
        datasets: [{
          label: col || titleText,
          data,
          backgroundColor: PALETTE.map(c => c + 'cc'),
          borderColor:     PALETTE,
          borderWidth:     1.5,
          borderRadius:    type === 'bar' ? 4 : 0,
          fill:            type === 'line'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: true,
        plugins: {
          title: {
            display:  !!titleText,
            text:     titleText,
            color:    '#e8ecf7',
            font:     { size: 13, family: 'Syne', weight: '600' }
          },
          legend: {
            display:  isRound,
            position: 'bottom',
            labels:   { color: '#8e97b8', font: { size: 11 } }
          },
          tooltip: {
            backgroundColor: '#1e2435',
            titleColor:      '#e8ecf7',
            bodyColor:       '#8e97b8',
            borderColor:     '#2a3250',
            borderWidth:     1
          }
        },
        scales: isRound ? {} : {
          x: { ticks: { color: '#566080', font: { size: 11 } }, grid: { color: '#1e2435' } },
          y: { ticks: { color: '#566080', font: { size: 11 } }, grid: { color: '#1e2435' }, beginAtZero: true }
        }
      }
    });
  }

  // ── ABOVE/BELOW DOUGHNUT ──────────────────────────
  function renderAboveBelow(sheetKey, data, headers) {
    const canvasId = `aboveBelow_${sheetKey}`;
    const canvas   = document.getElementById(canvasId);
    if (!canvas) return;
    const result = AppDescriptive.aboveBelowAvg(data, headers);
    destroyChart(canvasId);
    _instances[canvasId] = new Chart(canvas, {
      type: 'doughnut',
      data: {
        labels:   ['Above Average', 'Average', 'Below Average'],
        datasets: [{
          data:            [result.above, result.average, result.below],
          backgroundColor: ['#36c5a0cc', '#4f8ef7cc', '#e05252cc'],
          borderColor:     ['#36c5a0',   '#4f8ef7',   '#e05252'],
          borderWidth:     2
        }]
      },
      options: {
        responsive: true,
        plugins: {
          legend: { position: 'bottom', labels: { color: '#8e97b8', font: { size: 12 } } }
        }
      }
    });
  }

  // ── ALL-CHARTS GRID (Charts page) ─────────────────
  function renderAllForSheet(sheetKey) {
    const headers = AppState.headers[sheetKey] || [];
    if (!headers.length) return '<div class="empty-state"><p>No data.</p></div>';
    return `<div class="grid-auto" id="chartsGrid_${sheetKey}">${
      headers.map(h => {
        const safeH = h.replace(/\W/g, '_') + '_' + sheetKey;
        return `
        <div class="chart-box">
          <div style="font-size:12px;font-weight:600;color:var(--text3);margin-bottom:6px">${h}</div>
          <div style="display:flex;gap:6px;align-items:center;margin-bottom:8px;flex-wrap:wrap">
            <select class="all-chart-type" id="alltype_${safeH}" style="font-size:11px;padding:3px 8px;flex:1">
              <option value="bar">Bar</option>
              <option value="pie">Pie</option>
              <option value="doughnut">Doughnut</option>
              <option value="polarArea">Polar Area</option>
              <option value="line">Line</option>
            </select>
            <button class="btn btn-sm" style="padding:3px 8px;font-size:11px"
              onclick="AppCharts.downloadAllChart('${safeH}','${h.replace(/'/g,"\\'")}')">⬇</button>
          </div>
          <canvas id="allchart_${safeH}"></canvas>
        </div>`;
      }).join('')
    }</div>`;
  }

  function renderAllCharts() {
    ['sheet1', 'sheet2', 'sheet3'].forEach(sk => {
      const data    = AppState.sheets[sk];
      const headers = AppState.headers[sk] || [];
      if (!data) return;
      headers.forEach(h => {
        const safeH    = h.replace(/\W/g, '_') + '_' + sk;
        const canvasId = `allchart_${safeH}`;
        const canvas   = document.getElementById(canvasId);
        if (!canvas) return;
        const vals  = data.map(r => r[h]).filter(v => v !== undefined && v !== null && v !== '');
        const freq  = AppDescriptive.frequency(vals);
        const typeSel = document.getElementById(`alltype_${safeH}`);
        const type  = typeSel ? typeSel.value : 'bar';
        destroyChart(canvasId);
        _instances[canvasId] = buildChart(canvas, type, h, freq.map(f=>f.val), freq.map(f=>f.count), h);

        // Re-render on type change
        if (typeSel && !typeSel._bound) {
          typeSel._bound = true;
          typeSel.addEventListener('change', () => {
            destroyChart(canvasId);
            _instances[canvasId] = buildChart(canvas, typeSel.value, h, freq.map(f=>f.val), freq.map(f=>f.count), h);
          });
        }
      });
    });
  }

  function downloadAllChart(safeH, col) {
    const canvasId = `allchart_${safeH}`;
    const chart    = _instances[canvasId];
    if (!chart) return;
    const canvas = document.getElementById(canvasId);
    const ctx    = canvas.getContext('2d');
    const w = canvas.width, h = canvas.height;
    const imageData = ctx.getImageData(0, 0, w, h);
    ctx.globalCompositeOperation = 'destination-over';
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, w, h);
    const img = canvas.toDataURL('image/png', 1.0);
    ctx.clearRect(0, 0, w, h);
    ctx.putImageData(imageData, 0, 0);
    const a = document.createElement('a');
    a.href = img; a.download = `${col.replace(/\W/g,'_')}_chart.png`; a.click();
  }

  // ── TABS ───────────────────────────────────────────
  function bindTabs() {
    const bar = document.getElementById('chartTabBar');
    if (!bar) return;
    bar.querySelectorAll('.tab-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        bar.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
        const pane = document.getElementById(btn.dataset.tab);
        if (pane) pane.classList.add('active');
        setTimeout(renderAllCharts, 50);
      });
    });
  }

  return {
    renderEditableChart, refreshEditableChart,
    addEditRow, removeEditRow, downloadChart,
    renderAboveBelow,
    renderAllForSheet, renderAllCharts, downloadAllChart,
    bindTabs
  };
})();