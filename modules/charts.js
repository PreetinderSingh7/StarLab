const AppCharts = (() => {

  const _instances = {};

  const PALETTE = [
    '#4f8ef7','#36c5a0','#f0a500','#e05252','#7b5ea7',
    '#4bc9f0','#a3e05c','#f47ef7','#f7a04f','#5ec4f7'
  ];

  function destroyChart(id) {
    if (_instances[id]) { _instances[id].destroy(); delete _instances[id]; }
  }

  function getEditData(safeId) {
    const table = document.getElementById(`editTable_${safeId}`);
    if (!table) return { labels:[], data:[] };
    const labels = [], data = [];
    table.querySelectorAll('tr').forEach(tr => {
      const lInput = tr.querySelector('.edit-label');
      const vInput = tr.querySelector('.edit-val');
      if (lInput && vInput) {
        labels.push(lInput.value.trim() || '?');
        data.push(parseFloat(vInput.value)||0);
      }
    });
    return { labels, data };
  }

  // ── MAIN EDITABLE CHART ───────────────────────────
  function renderEditableChart(canvasId, col, values, isDemographic) {
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    const safeId = canvasId.replace('editChart_','');
    const isNumeric = values.some(v => !isNaN(Number(v)) && v !== '');
    const freq = AppDescriptive.frequency(values);
    const labels = freq.map(f => f.val);
    const data   = freq.map(f => f.count);

    const typeSel = document.getElementById(`chartType_${safeId}`);
    const type = typeSel ? typeSel.value : 'bar';
    const titleInput = document.getElementById(`chartTitle_${safeId}`);
    const titleText = titleInput ? titleInput.value : col;

    destroyChart(canvasId);
    _instances[canvasId] = new Chart(canvas, {
      type,
      data: {
        labels,
        datasets: [{
          label: col,
          data,
          backgroundColor: PALETTE.map(c => c + 'cc'),
          borderColor: PALETTE,
          borderWidth: 1.5,
          borderRadius: type === 'bar' ? 4 : 0,
          fill: type === 'line'
        }]
      },
      options: buildOptions(type, titleText)
    });
  }

  function refreshEditableChart(safeId, col) {
    const canvasId = `editChart_${safeId}`;
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    const typeSel  = document.getElementById(`chartType_${safeId}`);
    const titleInput = document.getElementById(`chartTitle_${safeId}`);
    const type = typeSel ? typeSel.value : 'bar';
    const titleText = titleInput ? titleInput.value : col;
    const { labels, data } = getEditData(safeId);

    destroyChart(canvasId);
    _instances[canvasId] = new Chart(canvas, {
      type,
      data: {
        labels,
        datasets: [{
          label: col,
          data,
          backgroundColor: PALETTE.map(c => c + 'cc'),
          borderColor: PALETTE,
          borderWidth: 1.5,
          borderRadius: type === 'bar' ? 4 : 0,
          fill: type === 'line'
        }]
      },
      options: buildOptions(type, titleText)
    });
  }

  function addEditRow(safeId) {
    const table = document.getElementById(`editTable_${safeId}`);
    if (!table) return;
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><input type="text" value="New Label" class="edit-label" data-id="${safeId}" style="width:120px"/></td>
      <td><input type="number" value="0" class="edit-val" data-id="${safeId}" style="width:70px"/></td>`;
    table.appendChild(tr);
  }

  function downloadChart(safeId, col) {
    const canvasId = `editChart_${safeId}`;
    const chart = _instances[canvasId];
    if (!chart) { alert('Generate the chart first.'); return; }
    const a = document.createElement('a');
    a.href = chart.toBase64Image('image/png', 1.0);
    a.download = `${col.replace(/\W/g,'_')}_chart.png`;
    a.click();
  }

  // ── ABOVE/BELOW DOUGHNUT ──────────────────────────
  function renderAboveBelow(sheetKey, data, headers) {
    const canvasId = `aboveBelow_${sheetKey}`;
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    const result = AppDescriptive.aboveBelowAvg(data, headers);
    destroyChart(canvasId);
    _instances[canvasId] = new Chart(canvas, {
      type: 'doughnut',
      data: {
        labels: ['Above Average','Average','Below Average'],
        datasets: [{
          data: [result.above, result.average, result.below],
          backgroundColor: ['#36c5a0cc','#4f8ef7cc','#e05252cc'],
          borderColor: ['#36c5a0','#4f8ef7','#e05252'],
          borderWidth: 2
        }]
      },
      options: {
        responsive: true,
        plugins: {
          legend: { position:'bottom', labels:{ color:'#8e97b8', font:{size:12} } },
          title: { display:false }
        }
      }
    });
  }

  // ── ALL CHARTS GRID ────────────────────────────────
  function renderAllForSheet(sheetKey) {
    const headers = AppState.headers[sheetKey]||[];
    if (!headers.length) return '<div class="empty-state"><p>No data.</p></div>';
    return `<div class="grid-auto" id="chartsGrid_${sheetKey}">${
      headers.map(h => `
      <div class="chart-box">
        <div style="font-size:12px;font-weight:600;color:var(--text3);margin-bottom:6px">${h}</div>
        <select class="all-chart-type" data-sheet="${sheetKey}" data-col="${h.replace(/'/g,"\\'")}__${sheetKey}" style="font-size:11px;padding:3px 6px;margin-bottom:6px">
          <option value="bar">Bar</option>
          <option value="pie">Pie</option>
          <option value="doughnut">Doughnut</option>
          <option value="polarArea">Polar</option>
          <option value="line">Line</option>
        </select>
        <canvas id="allchart_${sheetKey}_${h.replace(/\W/g,'_')}"></canvas>
      </div>`).join('')
    }</div>`;
  }

  function renderAllCharts() {
    ['sheet1','sheet2','sheet3'].forEach(sk => {
      const data = AppState.sheets[sk];
      const headers = AppState.headers[sk]||[];
      if (!data) return;
      headers.forEach(h => {
        const canvasId = `allchart_${sk}_${h.replace(/\W/g,'_')}`;
        const canvas = document.getElementById(canvasId);
        if (!canvas) return;
        const vals = data.map(r=>r[h]).filter(v=>v!==undefined&&v!==null&&v!=='');
        const freq = AppDescriptive.frequency(vals);
        const typeSel = canvas.parentElement?.querySelector('.all-chart-type');
        const type = typeSel ? typeSel.value : 'bar';
        destroyChart(canvasId);
        _instances[canvasId] = new Chart(canvas, {
          type,
          data: {
            labels: freq.map(f=>f.val),
            datasets: [{
              label: h, data: freq.map(f=>f.count),
              backgroundColor: PALETTE.map(c=>c+'cc'),
              borderColor: PALETTE, borderWidth:1.5,
              borderRadius: type==='bar'?4:0
            }]
          },
          options: buildOptions(type, h)
        });
        // re-render on type change
        typeSel?.addEventListener('change', () => {
          destroyChart(canvasId);
          _instances[canvasId] = new Chart(canvas, {
            type: typeSel.value,
            data: {
              labels: freq.map(f=>f.val),
              datasets: [{
                label: h, data: freq.map(f=>f.count),
                backgroundColor: PALETTE.map(c=>c+'cc'),
                borderColor: PALETTE, borderWidth:1.5,
                borderRadius: typeSel.value==='bar'?4:0
              }]
            },
            options: buildOptions(typeSel.value, h)
          });
        });
      });
    });
  }

  // ── OPTIONS BUILDER ────────────────────────────────
  function buildOptions(type, title) {
    const isRound = ['pie','doughnut','polarArea','radar'].includes(type);
    return {
      responsive: true,
      maintainAspectRatio: true,
      plugins: {
        title: { display: !!title, text: title, color:'#e8ecf7', font:{size:13,family:'Syne',weight:'600'} },
        legend: { display: isRound, position:'bottom', labels:{color:'#8e97b8',font:{size:11}} },
        tooltip: {
          backgroundColor:'#1e2435', titleColor:'#e8ecf7',
          bodyColor:'#8e97b8', borderColor:'#2a3250', borderWidth:1
        }
      },
      scales: isRound ? {} : {
        x: { ticks:{color:'#566080',font:{size:11}}, grid:{color:'#1e2435'} },
        y: { ticks:{color:'#566080',font:{size:11}}, grid:{color:'#1e2435'} }
      }
    };
  }

  function bindTabs() {
    const bar = document.getElementById('chartTabBar');
    if (!bar) return;
    bar.querySelectorAll('.tab-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        bar.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
        btn.classList.add('active');
        document.querySelectorAll('.tab-pane').forEach(p=>p.classList.remove('active'));
        const pane = document.getElementById(btn.dataset.tab);
        if (pane) pane.classList.add('active');
        setTimeout(renderAllCharts, 50);
      });
    });
  }

  return { renderEditableChart, refreshEditableChart, addEditRow, downloadChart, renderAboveBelow, renderAllForSheet, renderAllCharts, bindTabs };
})();