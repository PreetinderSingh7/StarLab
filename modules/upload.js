const AppUpload = (() => {

  function bind() {
    const zone = document.getElementById('uploadZone');
    const input = document.getElementById('fileInput');
    if (!zone || !input) return;

    input.addEventListener('change', e => handleFile(e.target.files[0]));

    zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('drag-over'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('drag-over'));
    zone.addEventListener('drop', e => {
      e.preventDefault();
      zone.classList.remove('drag-over');
      if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
    });
  }

  function handleFile(file) {
    if (!file) return;
    const status = document.getElementById('uploadStatus');
    if (!file.name.match(/\.xlsx?$/i)) {
      status.innerHTML = `<div class="alert alert-warn">⚠ Please upload an .xlsx or .xls file.</div>`;
      return;
    }
    status.innerHTML = `<div class="alert alert-info"><span class="spinner"></span> Reading file...</div>`;
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'binary' });
        AppState._sheetNames = wb.SheetNames;
        AppState._workbook = wb;
        AppState.filename = file.name;

        // Auto-map first 3 sheets
        wb.SheetNames.slice(0, 3).forEach((name, idx) => {
          const key = `sheet${idx+1}`;
          const ws = wb.Sheets[name];
          const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
          AppState.sheets[key] = json;
          AppState.headers[key] = json.length ? Object.keys(json[0]) : [];
        });

        status.innerHTML = `<div class="alert alert-success">✓ Loaded <strong>${file.name}</strong> — ${wb.SheetNames.length} sheet(s) found.</div>`;
        // Re-render to show mapper
        navigate('upload');
      } catch(err) {
        status.innerHTML = `<div class="alert alert-warn">⚠ Error reading file: ${err.message}</div>`;
      }
    };
    reader.readAsBinaryString(file);
  }

  function remapSheets() {
    const wb = AppState._workbook;
    if (!wb) return;
    ['1','2','3'].forEach(n => {
      const sel = document.getElementById(`mapSheet${n}`);
      if (!sel) return;
      const idx = parseInt(sel.value);
      const key = `sheet${n}`;
      if (idx < 0) { AppState.sheets[key] = null; AppState.headers[key] = []; return; }
      const ws = wb.Sheets[wb.SheetNames[idx]];
      const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
      AppState.sheets[key] = json;
      AppState.headers[key] = json.length ? Object.keys(json[0]) : [];
    });
    const ms = document.getElementById('mapStatus');
    if (ms) ms.textContent = '✓ Mapping updated.';
  }

  return { bind, handleFile, remapSheets };
})();