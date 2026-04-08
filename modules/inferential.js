const AppInferential = (() => {

  // ── CHI-SQUARE CALCULATION ────────────────────────
  function chiSquare(col1Vals, col2Vals) {
    const n = Math.min(col1Vals.length, col2Vals.length);
    const c1 = col1Vals.slice(0, n).map(String);
    const c2 = col2Vals.slice(0, n).map(String);

    const cats1 = [...new Set(c1)];
    const cats2 = [...new Set(c2)];

    // Build observed frequency table
    const observed = {};
    cats1.forEach(a => {
      observed[a] = {};
      cats2.forEach(b => observed[a][b] = 0);
    });
    for (let i = 0; i < n; i++) observed[c1[i]][c2[i]]++;

    // Row/col totals
    const rowTotals = {}, colTotals = {};
    cats1.forEach(a => rowTotals[a] = cats2.reduce((s,b) => s + observed[a][b], 0));
    cats2.forEach(b => colTotals[b] = cats1.reduce((s,a) => s + observed[a][b], 0));

    // Chi-square statistic
    let chi2 = 0;
    cats1.forEach(a => {
      cats2.forEach(b => {
        const E = (rowTotals[a] * colTotals[b]) / n;
        if (E > 0) chi2 += Math.pow(observed[a][b] - E, 2) / E;
      });
    });

    const df = (cats1.length - 1) * (cats2.length - 1);
    const pVal = 1 - chiCDF(chi2, df);

    return { chi2, df, pVal, n, cats1, cats2, observed, rowTotals, colTotals };
  }

  // ── CHI DISTRIBUTION CDF (approximation) ─────────
  function chiCDF(x, df) {
    if (x <= 0) return 0;
    return regularizedGammaP(df / 2, x / 2);
  }

  function regularizedGammaP(a, x) {
    if (x < 0) return 0;
    if (x === 0) return 0;
    if (x >= a + 1) return 1 - regularizedGammaQ(a, x);
    let sum = 1 / a, term = 1 / a;
    for (let i = 1; i <= 200; i++) {
      term *= x / (a + i);
      sum += term;
      if (Math.abs(term) < 1e-10 * Math.abs(sum)) break;
    }
    return sum * Math.exp(-x + a * Math.log(x) - logGamma(a));
  }

  function regularizedGammaQ(a, x) {
    let f = 1, c = 1, d = 1 - a + x;
    if (Math.abs(d) < 1e-30) d = 1e-30;
    d = 1/d; let h = d;
    for (let i = 1; i <= 200; i++) {
      let an = -i*(i-a); let b = x + 2*i + 1 - a;
      d = an*d + b; c = b + an/c;
      if (Math.abs(d) < 1e-30) d = 1e-30;
      if (Math.abs(c) < 1e-30) c = 1e-30;
      d = 1/d; let del = d*c; h *= del;
      if (Math.abs(del-1) < 1e-10) break;
    }
    return Math.exp(-x + a*Math.log(x) - logGamma(a)) * h;
  }

  function logGamma(z) {
    const g = 7;
    const c = [0.99999999999980993,676.5203681218851,-1259.1392167224028,771.32342877765313,-176.61502916214059,12.507343278686905,-0.13857109526572012,9.9843695780195716e-6,1.5056327351493116e-7];
    if (z < 0.5) return Math.log(Math.PI) - Math.log(Math.sin(Math.PI*z)) - logGamma(1-z);
    z--;
    let x = c[0];
    for (let i = 1; i < g+2; i++) x += c[i]/(z+i);
    const t = z + g + 0.5;
    return 0.5*Math.log(2*Math.PI) + (z+0.5)*Math.log(t) - t + Math.log(x);
  }

  // ── RENDER CHI RESULT ──────────────────────────────
  function renderChiResult(label1, label2, result) {
    const sig = result.pVal < 0.05;
    const sigClass = sig ? 'sig-yes' : 'sig-no';
    const sigText = sig ? '✓ Significant (p < 0.05)' : '✗ Not Significant (p ≥ 0.05)';

    // Contingency table
    const headerCells = result.cats2.map(b => `<th>${b}</th>`).join('') + '<th>Total</th>';
    const bodyRows = result.cats1.map(a => `
      <tr>
        <td>${a}</td>
        ${result.cats2.map(b => `<td>${result.observed[a][b]}</td>`).join('')}
        <td><strong>${result.rowTotals[a]}</strong></td>
      </tr>`).join('');
    const footRow = `<tr>
      <td><strong>Total</strong></td>
      ${result.cats2.map(b => `<td><strong>${result.colTotals[b]}</strong></td>`).join('')}
      <td><strong>${result.n}</strong></td>
    </tr>`;

    return `
    <div class="chi-result">
      <div style="margin-bottom:16px">
        <span class="chi-stat"><span class="val">${result.chi2.toFixed(4)}</span><span class="lbl">Chi-Square (χ²)</span></span>
        <span class="chi-stat"><span class="val">${result.df}</span><span class="lbl">Degrees of Freedom</span></span>
        <span class="chi-stat"><span class="val">${result.pVal.toFixed(4)}</span><span class="lbl">p-value</span></span>
        <span class="chi-stat"><span class="val">${result.n}</span><span class="lbl">N</span></span>
      </div>
      <div style="font-size:14px;font-weight:600" class="${sigClass}">${sigText}</div>
      <p style="font-size:12.5px;color:var(--text3);margin-top:6px">
        The association between <strong>${label1}</strong> and <strong>${label2}</strong> is 
        ${sig ? 'statistically significant' : 'not statistically significant'} 
        at the 0.05 significance level.
      </p>
    </div>
    <div class="section-title" style="margin-top:20px">Contingency Table</div>
    <div class="table-wrap">
      <table>
        <thead><tr><th>${label1} \\ ${label2}</th>${headerCells}</tr></thead>
        <tbody>${bodyRows}${footRow}</tbody>
      </table>
    </div>`;
  }

  // ── RUN FROM UI ────────────────────────────────────
  function runChiSquare() {
    const v1i = parseInt(document.getElementById('chiVar1').value);
    const v2i = parseInt(document.getElementById('chiVar2').value);
    const h1 = window._chiHeaders[v1i];
    const h2 = window._chiHeaders[v2i];
    if (v1i === v2i) {
      document.getElementById('chiResult').innerHTML = `<div class="alert alert-warn">⚠ Please select two different variables.</div>`;
      return;
    }
    const d1 = AppState.sheets[h1.sheet];
    const d2 = AppState.sheets[h2.sheet];
    if (!d1 || !d2) return;
    const vals1 = d1.map(r => r[h1.col]);
    const vals2 = d2.map(r => r[h2.col]);
    try {
      const result = chiSquare(vals1, vals2);
      document.getElementById('chiResult').innerHTML = renderChiResult(h1.label, h2.label, result);
    } catch(e) {
      document.getElementById('chiResult').innerHTML = `<div class="alert alert-warn">⚠ Error: ${e.message}</div>`;
    }
  }

  // ── BATCH CHI-SQUARE ──────────────────────────────
  function batchChiSquare() {
    const s1Headers = AppState.headers.sheet1 || [];
    const s2Headers = AppState.headers.sheet2 || [];
    const d1 = AppState.sheets.sheet1;
    const d2 = AppState.sheets.sheet2;
    if (!d1 || !d2 || !s1Headers.length || !s2Headers.length) {
      document.getElementById('batchChiResult').innerHTML = `<div class="alert alert-warn">⚠ Need both Sheet 1 and Sheet 2 data.</div>`;
      return;
    }
    let rows = '';
    s2Headers.forEach(h2 => {
      s1Headers.forEach(h1 => {
        try {
          const v1 = d1.map(r => r[h1]);
          const v2 = d2.map(r => r[h2]);
          const res = chiSquare(v1, v2);
          const sig = res.pVal < 0.05;
          rows += `<tr>
            <td>${h2}</td><td>${h1}</td>
            <td>${res.chi2.toFixed(3)}</td>
            <td>${res.df}</td>
            <td>${res.pVal.toFixed(4)}</td>
            <td><span class="${sig ? 'badge-above' : 'badge-below'}">${sig ? 'Yes' : 'No'}</span></td>
          </tr>`;
        } catch(e) { /* skip invalid pairs */ }
      });
    });
    document.getElementById('batchChiResult').innerHTML = `
    <div class="table-wrap">
      <table>
        <thead><tr><th>Sheet 2 Variable</th><th>Sheet 1 (Demographic)</th><th>χ²</th><th>df</th><th>p-value</th><th>Significant?</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>`;
  }

  return { chiSquare, runChiSquare, batchChiSquare, renderChiResult };
})();