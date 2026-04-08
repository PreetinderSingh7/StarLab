const AppCharts = (() => {

  const _instances = {};

  const PALETTE = [
    '#4f8ef7','#36c5a0','#7b5ea7','#f0a500','#e05252',
    '#4bc9f0','#a3e05c','#f47ef7','#f7a04f','#5ec4f7'
  ];

  function destroyChart(id) {
    if (_instances[id]) { _instances[id].destroy(); delete _instances[id]; }
  }

  function getOrCreateCanvas(id, container) {
    let canvas = document.getElementById(id);
    if (!canvas) {
      canvas = document.createElement('canvas');
      canvas.id = id;
      if (container) container.appendChild(canvas);
    }
    return canvas;
  }

  // ── COLUMN CHART ──────────────────────────────────
  function renderColChart(col, values, isDemographic) {
    const safeId = `colChart_${col.replace(/\W/g,'_')}`;
    const canvas = document.getElementById(safeId);
    if (!canvas) return;

    const isNumeric = values.some(v => !isNaN(Number(v)) && v !== '');
    destroyChart(safeId);

    if (isNumeric && !isDemographic) {
      // Histogram
      const nums = values.map(Number).filter(v => !isNaN(v));
      const min = Math.min(...nums), max = Math.max(...nums);
      const bins = Math.min(10, new Set(nums).size);
      const binSize = (max - min) / bins || 1;
      const labels = [], data = [];
      for (let i = 0; i < bins; i++) {
        const lo = min + i * binSize, hi = lo + binSize;
        labels.push(`${lo.toFixed(1)}–${hi.toFixed(1)}`);
        data.push(nums.filter(n => n >= lo && (i === bins-1 ? n <= hi : n < hi)).length);
      }
      _instances[safeId] = new Chart(canvas, {
        type: 'bar',
        data: { labels, datasets: [{ label: col, data, backgroundColor: PALETTE[0] + 'cc', borderColor: PALETTE[0], borderWidth: 1.5, borderRadius: 4 }] },
        options: chartOptions('Histogram: ' + col, 'Frequency')
      });
    } else {
      // Frequency bar + pie
      const freq = AppDescriptive.frequency(values);
      const labels = freq.map(f => f.val);
      const data = freq.map(f => f.count);
      _instances[safeId] = new Chart(canvas, {
        type: 'bar',
        data: { labels, datasets: [{ label: 'Frequency', data, backgroundColor: PALETTE.map(c => c + 'cc'), borderColor: PALETTE, borderWidth: 1.5, borderRadius: 4 }] },
        options: chartOptions(col, 'Frequency')
      });
    }
  }

  // ── ABOVE/BELOW AVG CHART ─────────────────────────
  function renderAboveBelow(sheetKey, data, headers) {
    const canvasId = `aboveBelow_${sheetKey}`;
    const canvas = document.getElementById(canvasId);
    if (!canvas) return;
    const result = AppDescriptive.aboveBelowAvg(data, headers);
    destroyChart(canvasId);
    _instances[canvasId] = new Chart(canvas, {
      type: 'doughnut',
      data: {
        labels: ['Above Average', 'Average', 'Below Average'],
        datasets: [{ data: [result.above, result.average, result.below], backgroundColor: ['#36c5a0cc','#4f8ef7cc','#e05252cc'], borderColor: ['#36c5a0','#4f8ef7','#e05252'], borderWidth: 2 }]
      },
      options: {
        ...chartOptions('Respondent Distribution'),
        plugins: { legend: { position: 'bottom', labels: { color: '#8e97b8', font: { size: 12 } } } }
      }
    });
  }

  // ── ALL CHARTS FOR A SHEET ─────────────────────────
  function renderAllForSheet(sheetKey) {
    const headers = AppState.headers[sheetKey] || [];
    if (!headers.length) return '<div class="empty-state"><p>No data.</p></div>';
    const grid = headers.map(h => `
      <div class="chart-box">
        <div style="font-size:12px;font-weight:600;color:var(--text3);margin-bottom:8px;letter-spacing:0.5px">${h}</div>
        <canvas id="allchart_${sheetKey}_${h.replace(/\W/g,'_')}"></canvas>
      </div>`).join('');
    return `<div class="grid-auto" id="chartsGrid_${sheetKey}">${grid}</div>`;
  }

  function renderAllCharts() {
    ['sheet1','sheet2','sheet3'].forEach(sk => {
      const data = AppState.sheets[sk];
      const headers = AppState.headers[sk] || [];
      if (!data) return;
      headers.forEach(h => {
        const canvasId = `allchart_${sk}_${h.replace(/\W/g,'_')}`;
        const canvas = document.getElementById(canvasId);
        if (!canvas) return;
        const vals = data.map(r => r[h]).filter(v => v !== undefined && v !== null && v !== '');
        const isNumeric = vals.some(v => !isNaN(Number(v)) && v !== '');
        const freq = AppDescriptive.frequency(vals);
        destroyChart(canvasId);
        _instances[canvasId] = new Chart(canvas, {
          type: isNumeric ? 'bar' : 'pie',
          data: {
            labels: freq.map(f => f.val),
            datasets: [{ label: h, data: freq.map(f => f.count), backgroundColor: PALETTE.map(c => c + 'cc'), borderColor: PALETTE, borderWidth: 1.5, borderRadius: isNumeric ? 4 : 0 }]
          },
          options: { ...chartOptions(h), plugins: { legend: { display: !isNumeric, position: 'bottom', labels: { color: '#8e97b8', font: { size: 11 } } } } }
        });
      });
    });
  }

  // ── CHART OPTIONS ─────────────────────────────────
  function chartOptions(title, yLabel) {
    return {
      responsive: true,
      maintainAspectRatio: true,
      plugins: {
        title: { display: false },
        legend: { display: false },
        tooltip: {
          backgroundColor: '#1e2435',
          titleColor: '#e8ecf7',
          bodyColor: '#8e97b8',
          borderColor: '#2a3250',
          borderWidth: 1
        }
      },
      scales: {
        x: { ticks: { color: '#566080', font: { size: 11 } }, grid: { color: '#1e2435' } },
        y: { ticks: { color: '#566080', font: { size: 11 } }, grid: { color: '#1e2435' }, title: { display: !!yLabel, text: yLabel, color: '#566080' } }
      }
    };
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

  return { renderColChart, renderAboveBelow, renderAllForSheet, renderAllCharts, bindTabs };
})();