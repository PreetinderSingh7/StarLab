const AppExport = (() => {

  // ── WORD EXPORT — uses HTML → .doc trick (opens in Word) ─────────
  function exportSheetWord(sheetKey) {
    const data    = AppState.sheets[sheetKey];
    const headers = AppState.headers[sheetKey];
    if (!data || !data.length) { alert('No data to export.'); return; }

    const isDemographic = sheetKey === 'sheet1';
    let tableRows = '';
    let sNo = 1;

    // Build grouped frequency rows like the thesis format
    headers.forEach(col => {
      const vals = data.map(r => r[col]).filter(v => v !== undefined && v !== null && v !== '');
      const freq = AppDescriptive.frequency(vals);
      const stats = AppDescriptive.numericStats(vals.map(Number).filter(v => !isNaN(v)));

      // Merge-row for the variable name (S.No | Variable | ... )
      tableRows += `
        <tr style="background:#f0f4ff">
          <td rowspan="${freq.length}" style="text-align:center;vertical-align:top;font-weight:bold;border:1px solid #ccc;padding:6px 8px">${sNo++}</td>
          <td colspan="4" style="font-weight:bold;border:1px solid #ccc;padding:6px 8px">${col}</td>
        </tr>`;

      freq.forEach(({val, count, pct}, idx) => {
        const isNum = !isNaN(Number(val));
        const aboveBelow = (!isDemographic && stats && isNum)
          ? (Number(val) > stats.mean ? 'Above Avg' : Number(val) < stats.mean ? 'Below Avg' : 'At Mean')
          : '';
        tableRows += `
          <tr>
            <td style="border:1px solid #ccc;padding:5px 8px;padding-left:20px">
              ${idx+1}. ${val}
            </td>
            <td style="text-align:center;border:1px solid #ccc;padding:5px 8px">${count}</td>
            <td style="text-align:center;border:1px solid #ccc;padding:5px 8px">${pct}%</td>
            ${!isDemographic ? `<td style="text-align:center;border:1px solid #ccc;padding:5px 8px">${aboveBelow}</td>` : ''}
          </tr>`;
      });

      // Stats footer row for scale sheets
      if (!isDemographic && stats) {
        tableRows += `
          <tr style="background:#f9f9f9">
            <td colspan="2" style="border:1px solid #ccc;padding:5px 8px;font-style:italic;color:#555">
              Mean: ${stats.mean.toFixed(2)} &nbsp;|&nbsp;
              Median: ${stats.median.toFixed(2)} &nbsp;|&nbsp;
              SD: ${stats.sd.toFixed(2)}
            </td>
            <td style="border:1px solid #ccc;padding:5px 8px"></td>
            <td style="border:1px solid #ccc;padding:5px 8px"></td>
          </tr>`;
      }
    });

    // Above/below summary for scale sheets
    let aboveBelowSummary = '';
    if (!isDemographic) {
      const ab = AppDescriptive.aboveBelowAvg(data, headers);
      aboveBelowSummary = `
      <h3 style="font-family:Calibri,sans-serif;font-size:13pt;margin-top:24pt">
        Respondent Categorisation (Mean ± 1 SD Method)
      </h3>
      <table style="border-collapse:collapse;width:60%;font-family:Calibri,sans-serif;font-size:11pt">
        <thead>
          <tr style="background:#1e3a5f;color:white">
            <th style="border:1px solid #ccc;padding:6px 12px">Category</th>
            <th style="border:1px solid #ccc;padding:6px 12px">Frequency</th>
            <th style="border:1px solid #ccc;padding:6px 12px">Percentage</th>
          </tr>
        </thead>
        <tbody>
          <tr><td style="border:1px solid #ccc;padding:5px 12px">Above Average (Score &gt; ${(ab.mean+ab.sd).toFixed(2)})</td>
              <td style="border:1px solid #ccc;padding:5px 12px;text-align:center">${ab.above}</td>
              <td style="border:1px solid #ccc;padding:5px 12px;text-align:center">${((ab.above/data.length)*100).toFixed(1)}%</td></tr>
          <tr><td style="border:1px solid #ccc;padding:5px 12px">Average (${(ab.mean-ab.sd).toFixed(2)} – ${(ab.mean+ab.sd).toFixed(2)})</td>
              <td style="border:1px solid #ccc;padding:5px 12px;text-align:center">${ab.average}</td>
              <td style="border:1px solid #ccc;padding:5px 12px;text-align:center">${((ab.average/data.length)*100).toFixed(1)}%</td></tr>
          <tr><td style="border:1px solid #ccc;padding:5px 12px">Below Average (Score &lt; ${(ab.mean-ab.sd).toFixed(2)})</td>
              <td style="border:1px solid #ccc;padding:5px 12px;text-align:center">${ab.below}</td>
              <td style="border:1px solid #ccc;padding:5px 12px;text-align:center">${((ab.below/data.length)*100).toFixed(1)}%</td></tr>
          <tr style="background:#f5f5f5">
            <td style="border:1px solid #ccc;padding:5px 12px;font-weight:bold">Total</td>
            <td style="border:1px solid #ccc;padding:5px 12px;text-align:center;font-weight:bold">${data.length}</td>
            <td style="border:1px solid #ccc;padding:5px 12px;text-align:center;font-weight:bold">100%</td>
          </tr>
        </tbody>
      </table>
      <p style="font-family:Calibri,sans-serif;font-size:10pt;color:#555;margin-top:6pt">
        Grand Mean = ${ab.mean.toFixed(2)}, SD = ${ab.sd.toFixed(2)}.
        Above Average = Score &gt; M+SD = ${(ab.mean+ab.sd).toFixed(2)};
        Below Average = Score &lt; M−SD = ${(ab.mean-ab.sd).toFixed(2)}.
      </p>`;
    }

    const colHeader = isDemographic
      ? '<th style="border:1px solid #ccc;padding:6px 10px">S.No</th><th style="border:1px solid #ccc;padding:6px 10px">Sample Characteristics</th><th style="border:1px solid #ccc;padding:6px 10px">Frequency (f)</th><th style="border:1px solid #ccc;padding:6px 10px">Percentage (%)</th>'
      : '<th style="border:1px solid #ccc;padding:6px 10px">S.No</th><th style="border:1px solid #ccc;padding:6px 10px">Sample Characteristics</th><th style="border:1px solid #ccc;padding:6px 10px">Frequency (f)</th><th style="border:1px solid #ccc;padding:6px 10px">Percentage (%)</th><th style="border:1px solid #ccc;padding:6px 10px">Category</th>';

    const sheetTitle = PAGES[sheetKey].title;
    const sectionName = isDemographic
      ? 'Section 1: Frequency and Frequency Percentage of Sample Characteristics'
      : `Section: Frequency, Percentage and Descriptive Statistics — ${sheetTitle}`;

    const html = `
    <html xmlns:o="urn:schemas-microsoft-com:office:office"
          xmlns:w="urn:schemas-microsoft-com:office:word"
          xmlns="http://www.w3.org/TR/REC-html40">
    <head>
      <meta charset="UTF-8">
      <!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View></w:WordDocument></xml><![endif]-->
      <style>
        body { font-family: Calibri, sans-serif; font-size: 12pt; margin: 2cm; }
        h1 { font-size: 16pt; text-align: center; }
        h2 { font-size: 13pt; text-align: center; margin-bottom: 4pt; }
        h3 { font-size: 12pt; }
        table { border-collapse: collapse; width: 100%; font-size: 11pt; }
        th { background-color: #1e3a5f; color: white; font-weight: bold; }
        td, th { border: 1px solid #999; padding: 5px 8px; }
        tr:nth-child(even) { background-color: #f2f2f2; }
        p { font-size: 11pt; }
      </style>
    </head>
    <body>
      <h1>StatLab Pro — Analysis Report</h1>
      <p style="text-align:center;color:#555;font-size:10pt">
        File: ${AppState.filename} &nbsp;|&nbsp; Generated: ${new Date().toLocaleString()}
      </p>
      <hr/>

      <h2>${sectionName}</h2>
      <p style="text-align:center;font-style:italic;font-size:10pt">
        Table: Frequency &amp; Frequency Percentage Distribution of the Sample Characteristics
      </p>

      <table>
        <thead><tr>${colHeader}</tr></thead>
        <tbody>${tableRows}</tbody>
      </table>

      ${aboveBelowSummary}

      <p style="margin-top:24pt;font-size:10pt;color:#777">
        Note: Above/Below Average classification uses the Mean ± 1 SD method.
        Above Average = Total Score &gt; Grand Mean + 1SD;
        Below Average = Total Score &lt; Grand Mean − 1SD.
      </p>
    </body>
    </html>`;

    const blob = new Blob(['\ufeff', html], { type: 'application/msword' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `${sheetKey}_analysis_table.doc`;
    a.click();
  }

  // ── FULL PDF EXPORT ────────────────────────────────
  function exportAll() {
    try {
      const { jsPDF } = window.jspdf;
      const doc = new jsPDF({ orientation:'portrait', unit:'mm', format:'a4' });
      let y = 20;
      const margin = 15;

      doc.setFillColor(30,36,53);
      doc.rect(0,0,210,30,'F');
      doc.setTextColor(232,236,247);
      doc.setFontSize(18); doc.setFont('helvetica','bold');
      doc.text('StatLab Pro — Analysis Report', margin, 18);
      doc.setFontSize(9); doc.setFont('helvetica','normal');
      doc.setTextColor(142,151,184);
      doc.text(`Generated: ${new Date().toLocaleString()} | File: ${AppState.filename}`, margin, 26);
      y = 40;

      ['sheet1','sheet2','sheet3'].forEach((sk,si) => {
        const data = AppState.sheets[sk];
        const headers = AppState.headers[sk];
        if (!data||!data.length) return;
        if (y>250) { doc.addPage(); y=20; }
        doc.setFontSize(13); doc.setFont('helvetica','bold');
        doc.setTextColor(79,142,247);
        doc.text(`Sheet ${si+1}: ${PAGES[sk].title}`, margin, y); y+=7;
        doc.setFontSize(9); doc.setFont('helvetica','normal');
        doc.setTextColor(100,110,140);
        doc.text(`${data.length} respondents, ${headers.length} columns`, margin, y); y+=8;

        headers.slice(0,10).forEach(h => {
          if (y>260) { doc.addPage(); y=20; }
          const vals = data.map(r=>r[h]).filter(v=>v!==undefined&&v!=='');
          const freq = AppDescriptive.frequency(vals);
          const stats = AppDescriptive.numericStats(vals.map(Number).filter(v=>!isNaN(v)));
          doc.setFontSize(10); doc.setFont('helvetica','bold'); doc.setTextColor(200,210,240);
          doc.text(h, margin, y); y+=5;
          if (stats) {
            doc.setFontSize(8); doc.setFont('helvetica','normal'); doc.setTextColor(100,180,140);
            doc.text(`Mean: ${stats.mean.toFixed(2)}  Median: ${stats.median.toFixed(2)}  SD: ${stats.sd.toFixed(2)}`, margin+4, y); y+=5;
          }
          freq.slice(0,8).forEach(({val,count,pct}) => {
            if (y>270) { doc.addPage(); y=20; }
            doc.setFontSize(8.5); doc.setFont('helvetica','normal'); doc.setTextColor(142,151,184);
            doc.text(`  ${val}`, margin+4, y);
            doc.text(`${count}`, margin+80, y);
            doc.text(`${pct}%`, margin+100, y); y+=5;
          }); y+=3;
        });

        if (sk!=='sheet1') {
          const ab = AppDescriptive.aboveBelowAvg(data,headers);
          if (y>250) { doc.addPage(); y=20; }
          doc.setFontSize(9); doc.setFont('helvetica','bold'); doc.setTextColor(54,197,160);
          doc.text(`Above: ${ab.above}  Average: ${ab.average}  Below: ${ab.below}  Mean: ${ab.mean.toFixed(2)}  SD: ${ab.sd.toFixed(2)}`, margin, y); y+=10;
        }
      });

      doc.save(`StatLab_Report_${Date.now()}.pdf`);
    } catch(e) { alert('Export error: '+e.message); }
  }

  return { exportSheetWord, exportAll };
})();