const AppExport = (() => {

  function exportAll() {
    try {
      const { jsPDF } = window.jspdf;
      const doc = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
      let y = 20;
      const margin = 15;
      const pageW = 210;

      // Header
      doc.setFillColor(30, 36, 53);
      doc.rect(0, 0, pageW, 30, 'F');
      doc.setTextColor(232, 236, 247);
      doc.setFontSize(18);
      doc.setFont('helvetica', 'bold');
      doc.text('StatLab Pro — Analysis Report', margin, 18);
      doc.setFontSize(9);
      doc.setFont('helvetica', 'normal');
      doc.setTextColor(142, 151, 184);
      doc.text(`Generated: ${new Date().toLocaleString()} | File: ${AppState.filename}`, margin, 26);
      y = 40;

      ['sheet1','sheet2','sheet3'].forEach((sk, si) => {
        const data = AppState.sheets[sk];
        const headers = AppState.headers[sk];
        if (!data || !data.length) return;

        if (y > 250) { doc.addPage(); y = 20; }
        doc.setFontSize(13);
        doc.setFont('helvetica', 'bold');
        doc.setTextColor(79, 142, 247);
        doc.text(`Sheet ${si+1}: ${PAGES[sk].title}`, margin, y); y += 7;

        doc.setFontSize(9);
        doc.setFont('helvetica', 'normal');
        doc.setTextColor(100, 110, 140);
        doc.text(`${data.length} respondents, ${headers.length} columns`, margin, y); y += 8;

        headers.slice(0, 10).forEach(h => {
          if (y > 260) { doc.addPage(); y = 20; }
          const vals = data.map(r => r[h]).filter(v => v !== undefined && v !== '');
          const freq = AppDescriptive.frequency(vals);

          doc.setFontSize(10);
          doc.setFont('helvetica', 'bold');
          doc.setTextColor(200, 210, 240);
          doc.text(h, margin, y); y += 5;

          freq.slice(0, 8).forEach(({ val, count, pct }) => {
            if (y > 270) { doc.addPage(); y = 20; }
            doc.setFontSize(8.5);
            doc.setFont('helvetica', 'normal');
            doc.setTextColor(142, 151, 184);
            doc.text(`  ${val}`, margin + 4, y);
            doc.text(`${count}`, margin + 80, y);
            doc.text(`${pct}%`, margin + 100, y);
            y += 5;
          });
          y += 3;
        });

        // Stats for scale sheets
        if (sk !== 'sheet1') {
          const ab = AppDescriptive.aboveBelowAvg(data, headers);
          if (y > 250) { doc.addPage(); y = 20; }
          doc.setFontSize(9);
          doc.setFont('helvetica', 'bold');
          doc.setTextColor(54, 197, 160);
          doc.text(`Above Avg: ${ab.above}  Average: ${ab.average}  Below Avg: ${ab.below}  Mean: ${ab.mean.toFixed(2)}  SD: ${ab.sd.toFixed(2)}`, margin, y);
          y += 10;
        }
      });

      doc.save(`StatLab_Report_${Date.now()}.pdf`);
    } catch(e) {
      alert('Export error: ' + e.message);
    }
  }

  function exportCSV(sheetKey) {
    const data = AppState.sheets[sheetKey];
    if (!data) return;
    const headers = AppState.headers[sheetKey];
    const rows = [headers.join(','), ...data.map(r => headers.map(h => `"${r[h] ?? ''}"`).join(','))];
    const blob = new Blob([rows.join('\n')], { type: 'text/csv' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `${sheetKey}_data.csv`;
    a.click();
  }

  return { exportAll, exportCSV };
})();