const AppExport = (() => {

  // ── WORD EXPORT — formatted table matching thesis style ──
  function exportSheetWord(sheetKey) {
    const data    = AppState.sheets[sheetKey];
    const headers = AppState.headers[sheetKey];
    if (!data || !data.length) { alert('No data to export for this sheet.'); return; }

    const isDemographic = sheetKey === 'sheet1';
    const sheetTitle    = PAGES[sheetKey].title;
    let   tableRows     = '';
    let   sNo           = 1;

    headers.forEach(col => {
      const vals  = data.map(r => r[col]).filter(v => v !== undefined && v !== null && v !== '');
      const freq  = AppDescriptive.frequency(vals);
      const nums  = vals.map(Number).filter(v => !isNaN(v));
      const stats = nums.length ? AppDescriptive.numericStats(nums) : null;
      const total = vals.length;

      // Header row for each variable
      const extraCols = isDemographic ? '' : '<td colspan="3"></td>';
      tableRows += `
      <tr style="background:#dce6f1">
        <td style="border:1px solid #999;padding:6px 10px;font-weight:bold;text-align:center;vertical-align:middle">${sNo++}</td>
        <td colspan="${isDemographic ? 3 : 5}" style="border:1px solid #999;padding:6px 10px;font-weight:bold">${col}</td>
      </tr>`;

      // Sub-rows for each value
      freq.forEach(({ val, count, pct }, idx) => {
        const isNum = !isNaN(Number(val)) && val !== '';
        let category = '';
        if (!isDemographic && stats && isNum) {
          const n = Number(val);
          if      (n > stats.mean) category = 'Above Average';
          else if (n < stats.mean) category = 'Below Average';
          else                     category = 'Average';
        }
        tableRows += `
        <tr>
          <td style="border:1px solid #999;padding:5px 10px;text-align:center"></td>
          <td style="border:1px solid #999;padding:5px 10px;padding-left:20px">${idx+1}. ${val}</td>
          <td style="border:1px solid #999;padding:5px 10px;text-align:center">${count}</td>
          <td style="border:1px solid #999;padding:5px 10px;text-align:center">${pct}%</td>
          ${!isDemographic ? `<td style="border:1px solid #999;padding:5px 10px;text-align:center">${category}</td>` : ''}
        </tr>`;
      });

      // Stats footer row
      if (!isDemographic && stats) {
        tableRows += `
        <tr style="background:#f2f2f2">
          <td style="border:1px solid #999;padding:5px 10px"></td>
          <td style="border:1px solid #999;padding:5px 10px;font-style:italic;color:#444" colspan="4">
            Mean = ${stats.mean.toFixed(2)} &nbsp;|&nbsp;
            Median = ${stats.median.toFixed(2)} &nbsp;|&nbsp;
            SD = ${stats.sd.toFixed(2)} &nbsp;|&nbsp;
            Min = ${stats.min} &nbsp;|&nbsp;
            Max = ${stats.max}
          </td>
        </tr>`;
      }
    });

    // ── ABOVE / BELOW TABLE ──────────────────────────
    let aboveBelowTable = '';
    if (!isDemographic) {
      const ab = AppDescriptive.aboveBelowAvg(data, headers);
      const n  = data.length;
      aboveBelowTable = `
      <br/><br/>
      <h3 style="font-family:Calibri,sans-serif;font-size:13pt;margin:0 0 6pt 0">
        Respondent Categorisation — Mean ± 1 SD Method
      </h3>
      <p style="font-family:Calibri,sans-serif;font-size:10pt;color:#555;margin:0 0 8pt 0">
        Grand Mean = ${ab.mean.toFixed(2)} &nbsp;|&nbsp;
        Median = ${ab.median.toFixed(2)} &nbsp;|&nbsp;
        SD = ${ab.sd.toFixed(2)}<br/>
        Above Average: score &gt; ${(ab.mean+ab.sd).toFixed(2)} &nbsp;|&nbsp;
        Average: ${(ab.mean-ab.sd).toFixed(2)} to ${(ab.mean+ab.sd).toFixed(2)} &nbsp;|&nbsp;
        Below Average: score &lt; ${(ab.mean-ab.sd).toFixed(2)}
      </p>
      <table style="border-collapse:collapse;width:70%;font-family:Calibri,sans-serif;font-size:11pt">
        <thead>
          <tr style="background:#1f3864;color:white">
            <th style="border:1px solid #999;padding:7px 12px;text-align:left">Category</th>
            <th style="border:1px solid #999;padding:7px 12px;text-align:center">Frequency (f)</th>
            <th style="border:1px solid #999;padding:7px 12px;text-align:center">Percentage (%)</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style="border:1px solid #999;padding:6px 12px">Above Average (Score &gt; ${(ab.mean+ab.sd).toFixed(2)})</td>
            <td style="border:1px solid #999;padding:6px 12px;text-align:center">${ab.above}</td>
            <td style="border:1px solid #999;padding:6px 12px;text-align:center">${((ab.above/n)*100).toFixed(1)}%</td>
          </tr>
          <tr style="background:#f9f9f9">
            <td style="border:1px solid #999;padding:6px 12px">Average (${(ab.mean-ab.sd).toFixed(2)} to ${(ab.mean+ab.sd).toFixed(2)})</td>
            <td style="border:1px solid #999;padding:6px 12px;text-align:center">${ab.average}</td>
            <td style="border:1px solid #999;padding:6px 12px;text-align:center">${((ab.average/n)*100).toFixed(1)}%</td>
          </tr>
          <tr>
            <td style="border:1px solid #999;padding:6px 12px">Below Average (Score &lt; ${(ab.mean-ab.sd).toFixed(2)})</td>
            <td style="border:1px solid #999;padding:6px 12px;text-align:center">${ab.below}</td>
            <td style="border:1px solid #999;padding:6px 12px;text-align:center">${((ab.below/n)*100).toFixed(1)}%</td>
          </tr>
          <tr style="background:#dce6f1;font-weight:bold">
            <td style="border:1px solid #999;padding:6px 12px">Total</td>
            <td style="border:1px solid #999;padding:6px 12px;text-align:center">${n}</td>
            <td style="border:1px solid #999;padding:6px 12px;text-align:center">100%</td>
          </tr>
        </tbody>
      </table>`;
    }

    // ── COLUMN HEADERS ────────────────────────────────
    const colHeaders = isDemographic
      ? `<th style="border:1px solid #999;padding:7px 10px;text-align:center">S.No</th>
         <th style="border:1px solid #999;padding:7px 10px">Sample Characteristics</th>
         <th style="border:1px solid #999;padding:7px 10px;text-align:center">Frequency (f)</th>
         <th style="border:1px solid #999;padding:7px 10px;text-align:center">Percentage (%)</th>`
      : `<th style="border:1px solid #999;padding:7px 10px;text-align:center">S.No</th>
         <th style="border:1px solid #999;padding:7px 10px">Sample Characteristics</th>
         <th style="border:1px solid #999;padding:7px 10px;text-align:center">Frequency (f)</th>
         <th style="border:1px solid #999;padding:7px 10px;text-align:center">Percentage (%)</th>
         <th style="border:1px solid #999;padding:7px 10px;text-align:center">Category</th>`;

    const sectionTitle = isDemographic
      ? 'Section 1: Frequency and Frequency Percentage of Sample Characteristics'
      : `Frequency, Percentage and Descriptive Statistics — ${sheetTitle}`;

    const html = `
    <html xmlns:o="urn:schemas-microsoft-com:office:office"
          xmlns:w="urn:schemas-microsoft-com:office:word"
          xmlns="http://www.w3.org/TR/REC-html40">
    <head>
      <meta charset="UTF-8"/>
      <!--[if gte mso 9]>
      <xml><w:WordDocument><w:View>Print</w:View><w:Zoom>90</w:Zoom></w:WordDocument></xml>
      <![endif]-->
      <style>
        body        { font-family: Calibri, sans-serif; font-size: 11pt; margin: 2.5cm; }
        h1          { font-size: 16pt; text-align: center; color: #1f3864; margin-bottom: 4pt; }
        h2          { font-size: 13pt; text-align: center; margin-bottom: 2pt; }
        p.sub       { font-size: 10pt; text-align: center; color: #555; font-style: italic; margin-top: 0; }
        table       { border-collapse: collapse; width: 100%; font-size: 10.5pt; margin-top: 10pt; }
        th          { background-color: #1f3864; color: white; font-weight: bold; }
        td, th      { border: 1px solid #999; padding: 5px 8px; }
        tr:nth-child(even) td { background-color: #f7f7f7; }
        .note       { font-size: 9.5pt; color: #666; margin-top: 10pt; }
      </style>
    </head>
    <body>
      <h1>StatLab Pro — Analysis Report</h1>
      <p class="sub">File: ${AppState.filename} &nbsp;|&nbsp; Generated: ${new Date().toLocaleString()}</p>
      <hr style="border:none;border-top:2px solid #1f3864;margin:10pt 0"/>

      <h2>${sectionTitle}</h2>
      <p class="sub">Table: Frequency &amp; Frequency Percentage Distribution of the Sample Characteristics</p>

      <table>
        <thead><tr>${colHeaders}</tr></thead>
        <tbody>${tableRows}</tbody>
      </table>

      ${aboveBelowTable}

      <p class="note">
        <b>Note on Categorisation:</b> Above/Below Average uses Mean ± 1 SD method.
        Above Average = Total Score &gt; Grand Mean + 1SD;
        Below Average = Total Score &lt; Grand Mean − 1SD;
        Average = between the two thresholds.
      </p>
    </body>
    </html>`;

    const blob = new Blob(['\ufeff', html], { type: 'application/msword' });
    const a    = document.createElement('a');
    a.href     = URL.createObjectURL(blob);
    a.download = `${sheetKey}_table.doc`;
    a.click();
  }

  // ── PDF EXPORT (full report) ──────────────────────
  function exportAll() {
    try {
      const { jsPDF } = window.jspdf;
      const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
      let y = 20;
      const margin = 15;

      doc.setFillColor(30, 36, 53);
      doc.rect(0, 0, 210, 30, 'F');
      doc.setTextColor(232, 236, 247);
      doc.setFontSize(18); doc.setFont('helvetica', 'bold');
      doc.text('StatLab Pro — Analysis Report', margin, 18);
      doc.setFontSize(9); doc.setFont('helvetica', 'normal');
      doc.setTextColor(142, 151, 184);
      doc.text(`Generated: ${new Date().toLocaleString()} | File: ${AppState.filename}`, margin, 26);
      y = 40;

      ['sheet1', 'sheet2', 'sheet3'].forEach((sk, si) => {
        const data    = AppState.sheets[sk];
        const headers = AppState.headers[sk];
        if (!data || !data.length) return;
        if (y > 250) { doc.addPage(); y = 20; }
        doc.setFontSize(13); doc.setFont('helvetica', 'bold');
        doc.setTextColor(79, 142, 247);
        doc.text(`Sheet ${si+1}: ${PAGES[sk].title}`, margin, y); y += 7;
        doc.setFontSize(9); doc.setFont('helvetica', 'normal');
        doc.setTextColor(100, 110, 140);
        doc.text(`${data.length} respondents, ${headers.length} columns`, margin, y); y += 8;

        headers.slice(0, 10).forEach(h => {
          if (y > 260) { doc.addPage(); y = 20; }
          const vals  = data.map(r => r[h]).filter(v => v !== undefined && v !== '');
          const freq  = AppDescriptive.frequency(vals);
          const stats = AppDescriptive.numericStats(vals.map(Number).filter(v => !isNaN(v)));
          doc.setFontSize(10); doc.setFont('helvetica', 'bold');
          doc.setTextColor(200, 210, 240);
          doc.text(h, margin, y); y += 5;
          if (stats) {
            doc.setFontSize(8); doc.setFont('helvetica', 'normal');
            doc.setTextColor(100, 180, 140);
            doc.text(`Mean: ${stats.mean.toFixed(2)}  Median: ${stats.median.toFixed(2)}  SD: ${stats.sd.toFixed(2)}`, margin+4, y); y += 5;
          }
          freq.slice(0, 8).forEach(({ val, count, pct }) => {
            if (y > 270) { doc.addPage(); y = 20; }
            doc.setFontSize(8.5); doc.setFont('helvetica', 'normal');
            doc.setTextColor(142, 151, 184);
            doc.text(`  ${val}`, margin+4, y);
            doc.text(`${count}`, margin+80, y);
            doc.text(`${pct}%`, margin+100, y); y += 5;
          }); y += 3;
        });

        if (sk !== 'sheet1') {
          const ab = AppDescriptive.aboveBelowAvg(data, headers);
          if (y > 250) { doc.addPage(); y = 20; }
          doc.setFontSize(9); doc.setFont('helvetica', 'bold');
          doc.setTextColor(54, 197, 160);
          doc.text(
            `Above: ${ab.above}  Average: ${ab.average}  Below: ${ab.below}  Mean: ${ab.mean.toFixed(2)}  Median: ${ab.median.toFixed(2)}  SD: ${ab.sd.toFixed(2)}`,
            margin, y
          ); y += 10;
        }
      });

      doc.save(`StatLab_Report_${Date.now()}.pdf`);
    } catch(e) {
      alert('Export error: ' + e.message);
    }
  }

  return { exportSheetWord, exportAll };
})();