/* === CartAudit Smart Script v13.5 — robust parsing, correct math & wave order,
       strict highlights (value==1 → yellow, value==3 → orange on Bags/OVs ONLY),
       single CSV, 1-wave-per-page PDF (header-based highlighting) === */

/* -------------------- CONFIG -------------------- */
const HIGHLIGHT_MODE = 'strict13'; // Highlight only when the cell value equals 1 (yellow) or 3 (orange)

/* -------------------- DOM refs / globals -------------------- */
const logBox          = document.getElementById('log-box');
const bufferContainer = document.getElementById('buffer-cards');
const bagContainer    = document.getElementById('bag-cards');

let pick = [];        // PickOrder.csv rows (header at [0])
let scc  = [];        // SCCPick.csv rows  (header at [0])
let sccQR = [];       // [["Route Code","Carts","Bags","OVs"], ...]
let bufferWaves = []; // BufferPlan waves (36 rows each)
let bagWaves    = []; // BagCount waves  (36 rows each)

/* -------------------- logging -------------------- */
function log(msg) {
  const p = document.createElement('div');
  p.textContent = `${new Date().toLocaleTimeString()}  ${msg}`;
  logBox.prepend(p);
}

/* -------------------- utilities -------------------- */
function normalizeWeirdCSV(text) {
  const lines = text.split(/\r?\n/);
  const out = [];
  for (let i = 0; i < lines.length; i++) {
    let s = lines[i];
    if (s == null) { out.push(''); continue; }
    s = s.replace(/^\uFEFF/, ''); // strip BOM
    const t = s.trim();
    if (t.length >= 2 && t[0] === '"' && t[t.length - 1] === '"') {
      s = t.slice(1, -1).replace(/""/g, '"');
    }
    out.push(s);
  }
  return out.join('\n');
}

function parseCSV(text, delim = ',') {
  const rows = [];
  let row = [], cell = '', inQ = false;

  for (let i = 0; i < text.length; i++) {
    const ch = text[i], next = text[i + 1];

    if (inQ) {
      if (ch === '"') {
        if (next === '"') { cell += '"'; i++; } else { inQ = false; }
      } else cell += ch;
    } else {
      if (ch === '"') inQ = true;
      else if (ch === delim) { row.push(cell.trim()); cell = ''; }
      else if (ch === '\n' || ch === '\r') {
        if (cell.length || row.length) { row.push(cell.trim()); rows.push(row); row = []; cell = ''; }
        if (ch === '\r' && next === '\n') i++;
      } else cell += ch;
    }
  }
  if (cell.length || row.length) { row.push(cell.trim()); rows.push(row); }
  return rows.filter(r => r.length && r.some(v => v !== ''));
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, m => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;" }[m]));
}

function downloadCSV(name, rows) {
  const csv = rows
    .map(r => r.map(v => String(v).replace(/"/g, '""')).map(v => `"${v}"`).join(','))
    .join('\n');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv' }));
  a.download = name;
  a.click();
}

function resetAll() {
  pick = []; scc = []; sccQR = [];
  bufferWaves = []; bagWaves = [];
  if (bufferContainer) bufferContainer.innerHTML = '';
  if (bagContainer)    bagContainer.innerHTML = '';
  log('Reset all in-memory data.');
}

/* -------------------- file loaders -------------------- */
// PickOrder (dispatchTime, routeCode, routeID, dispatchArea)
document.getElementById('file-pick').addEventListener('change', e => {
  const f = e.target.files[0];
  if (!f) return;
  f.text().then(t => {
    let raw = normalizeWeirdCSV(t);
    let head = raw.split(/\r?\n/, 1)[0] || '';
    let delimiter = head.includes(';') ? ';' : (head.includes('\t') ? '\t' : ',');
    pick = parseCSV(raw, delimiter);

    if ((pick[0]?.length || 0) === 1 && /",\s*"/.test(head)) {
      raw = normalizeWeirdCSV(t);
      head = raw.split(/\r?\n/, 1)[0] || '';
      delimiter = head.includes(';') ? ';' : (head.includes('\t') ? '\t' : ',');
      pick = parseCSV(raw, delimiter);
    }

    log(`Loaded PickOrder: ${pick.length - 1} data rows, delimiter="${delimiter}", cols=${pick[0]?.length || 0}`);
  });
});

// SCCPick (Route Code, Picklist Code, …, Bags, OVs, SPR, Progress, Associate, Duration, Type)
document.getElementById('file-scc').addEventListener('change', e => {
  const f = e.target.files[0];
  if (!f) return;
  f.text().then(t => {
    let raw = normalizeWeirdCSV(t);
    let head = raw.split(/\r?\n/, 1)[0] || '';
    let delimiter = head.includes(';') ? ';' : (head.includes('\t') ? '\t' : ',');
    scc = parseCSV(raw, delimiter);

    if ((scc[0]?.length || 0) === 1 && /",\s*"/.test(head)) {
      raw = normalizeWeirdCSV(t);
      head = raw.split(/\r?\n/, 1)[0] || '';
      delimiter = head.includes(';') ? ';' : (head.includes('\t') ? '\t' : ',');
      scc = parseCSV(raw, delimiter);
    }

    log(`Loaded SCCPick: ${scc.length - 1} data rows, delimiter="${delimiter}", cols=${scc[0]?.length || 0}`);

    const H = scc[0] || [];
    if (H[0] !== 'Route Code' || H[1] !== 'Picklist Code' || H[10] !== 'Bags' || H[11] !== 'OVs') {
      log(`⚠️ SCCPick header mismatch. Got: [${H.slice(0, 16).join(' | ')}]`);
      log('   Expected indices: A=Route Code (0), B=Picklist Code (1), K=Bags (10), L=OVs (11).');
    }
  });
});

/* -------------------- processing -------------------- */
function processSCC() {
  if (scc.length < 2) { log('Error: Load SCCPick file first.'); return; }

  const H = scc[0] || [];
  if (H[0] !== 'Route Code' || H[1] !== 'Picklist Code' || H[10] !== 'Bags' || H[11] !== 'OVs') {
    log(`⚠️ SCCPick header check failed before processing. Got: [${H.slice(0, 16).join(' | ')}]`);
    log('   Fix CSV parsing (quotes) or adjust indices if export changed.');
  }

  const colA = 0, colB = 1, colK = 10, colL = 11;

  log(`Processing SCCPick. Using columns: A=${colA}, B=${colB}, K=${colK}, L=${colL}`);

  const routeSet = new Set();
  for (let i = 1; i < scc.length; i++) {
    const r = scc[i];
    const route = (r?.[colA] || '').trim();
    if (route) routeSet.add(route);
  }

  const out = [['Route Code','Carts','Bags','OVs']];
  const rows = scc.slice(1);

  for (const route of routeSet) {
    let carts = 0, bags = 0, ovs = 0;

    for (const r of rows) {
      if (!r) continue;
      const picklist = (r[colB] || '').trim();

      if (picklist && picklist.startsWith(route + '#')) {
        carts++;
        const b = Number(String(r[colK] || '0').replace(/[^0-9.-]/g, '')) || 0;
        const o = Number(String(r[colL] || '0').replace(/[^0-9.-]/g, '')) || 0;
        bags += b;
        ovs  += o;
      }
    }

    out.push([route, carts, bags, ovs]);
  }

  out.splice(1, out.length - 1, ...out.slice(1).sort((x, y) => String(x[0]).localeCompare(String(y[0]))));
  sccQR = out;
  log(`SCC processed → ${sccQR.length - 1} routes.`);
}

/* -------------------- builders / rendering -------------------- */
function buildBuffer() {
  if (pick.length < 2) { log('Error: Load PickOrder first.'); return; }
  if (sccQR.length < 2) { log('Run SCC processing first.'); return; }

  const hdr = pick[0] || [];
  const idxRoute = hdr.indexOf('routeCode');
  const idxArea  = hdr.indexOf('dispatchArea');
  if (idxRoute === -1 || idxArea === -1) {
    log('⚠️ PickOrder header names not found. Falling back to routeCode=1, dispatchArea=3.');
  }

  const cartsMap = new Map();
  for (let i = 1; i < sccQR.length; i++) {
    const [route, carts] = sccQR[i];
    cartsMap.set(route, Number(carts) || 0);
  }

  const seen = new Set();
  const uniq = [];
  for (let i = 1; i < pick.length; i++) {
    const r = pick[i];
    const route = (r[idxRoute >= 0 ? idxRoute : 1] || '').trim();
    const loc   = (r[idxArea  >= 0 ? idxArea  : 3] || '').trim();
    if (!route || seen.has(route)) continue;
    seen.add(route);
    uniq.push([route, loc, cartsMap.get(route) ?? 0]);
  }

  bufferWaves = [];
  for (let i = 0; i < uniq.length; i += 36) {
    bufferWaves.push(uniq.slice(i, i + 36));
  }

  renderWaves(bufferWaves, false, bufferContainer);
  log(`Buffer built: ${bufferWaves.length} wave block(s), ${uniq.length} total rows.`);
}

function buildBagCount() {
  if (pick.length < 2) { log('Error: Load PickOrder first.'); return; }
  if (sccQR.length < 2) { log('Run SCC processing first.'); return; }

  const hdr = pick[0] || [];
  const idxRoute = hdr.indexOf('routeCode');
  const idxArea  = hdr.indexOf('dispatchArea');
  if (idxRoute === -1 || idxArea === -1) {
    log('⚠️ PickOrder header names not found. Falling back to routeCode=1, dispatchArea=3.');
  }

  const cartsMap = new Map();
  const bagsMap  = new Map();
  const ovsMap   = new Map();
  for (let i = 1; i < sccQR.length; i++) {
    const [route, carts, bags, ovs] = sccQR[i];
    cartsMap.set(route, Number(carts) || 0);
    bagsMap.set(route,  Number(bags)  || 0);
    ovsMap.set(route,   Number(ovs)   || 0);
  }

  const seen = new Set();
  const rows = [];
  for (let i = 1; i < pick.length; i++) {
    const r = pick[i];
    const route = (r[idxRoute >= 0 ? idxRoute : 1] || '').trim();
    const loc   = (r[idxArea  >= 0 ? idxArea  : 3] || '').trim();
    if (!route || seen.has(route)) continue;
    seen.add(route);
    rows.push([route, loc, cartsMap.get(route) ?? 0, bagsMap.get(route) ?? 0, ovsMap.get(route) ?? 0]);
  }

  bagWaves = [];
  for (let i = 0; i < rows.length; i += 36) {
    bagWaves.push(rows.slice(i, i + 36));
  }

  renderWaves(bagWaves, true, bagContainer);
  log(`BagCount built: ${bagWaves.length} wave block(s), ${rows.length} total rows.`);
}

/* -------------------- highlight helper -------------------- */
function cellClassByValue(v) {
  if (Number(v) === 1) return 'hl-bag';   // yellow for value==1
  if (Number(v) === 3) return 'hl-ovs';   // orange for value==3
  return '';
}

/* -------------------- render -------------------- */
function renderWaves(waves, showAllCols, targetEl) {
  const cont = targetEl;
  if (!cont) return;
  let html = '';

  for (let w = 0; w < waves.length; w++) {
    const rows = waves[w];

    html += `<div class="wave">
      <div class="wave-head">Wave ${w + 1}</div>
      <div class="wave-table">
        <table>
          <thead><tr>
            <th>Route Code</th><th>Location</th><th>Carts</th>${showAllCols ? '<th>Bags</th><th>OVs</th><th>Departed</th>' : '' }
          </tr></thead><tbody>`;

    rows.forEach(r => {
      if (showAllCols) {
        const bags = Number(r[3] ?? 0);
        const ovs  = Number(r[4] ?? 0);
        html += `<tr>
          <td>${escapeHtml(r[0] || '')}</td>
          <td>${escapeHtml(r[1] || '')}</td>
          <td>${escapeHtml(r[2] ?? 0)}</td>
          <td class="${cellClassByValue(bags)}">${escapeHtml(bags)}</td>
          <td class="${cellClassByValue(ovs)}">${escapeHtml(ovs)}</td>
          <td></td>
        </tr>`;
      } else {
        html += `<tr>
          <td>${escapeHtml(r[0] || '')}</td>
          <td>${escapeHtml(r[1] || '')}</td>
          <td>${escapeHtml(r[2] ?? 0)}</td>
        </tr>`;
      }
    });

    html += `</tbody></table></div></div>`;
  }

  cont.innerHTML = html;
}

/* -------------------- export (single wide CSV) -------------------- */
function exportCSV() {
  if (!bagWaves.length) {
    if (pick.length < 2 || sccQR.length < 2) {
      log('Nothing to export. Load files and build BagCount first.');
      return;
    }
    buildBagCount();
  }

  const waves = bagWaves;
  const perWaveCols = 6;
  const maxRows = Math.max(...waves.map(w => w.length), 36);

  const rows = [];

  const titleRow = [];
  for (let w = 0; w < waves.length; w++) {
    titleRow.push(`Wave ${w+1}`);
    for (let c = 1; c < perWaveCols; c++) titleRow.push('');
    titleRow.push('');
  }
  rows.push(titleRow);

  const headerBlock = ['Route Code','Location','Carts','Bags','OVs','Departed'];
  const headerRow = [];
  for (let w = 0; w < waves.length; w++) headerRow.push(...headerBlock, '');
  rows.push(headerRow);

  for (let r = 0; r < maxRows; r++) {
    const line = [];
    for (let w = 0; w < waves.length; w++) {
      const item = waves[w][r];
      if (item) {
        line.push(item[0] ?? '', item[1] ?? '', item[2] ?? 0, item[3] ?? 0, item[4] ?? 0, '');
      } else {
        line.push('', '', '', '', '', '');
      }
      line.push('');
    }
    rows.push(line);
  }

  downloadCSV('CartAudit_BagCount_View.csv', rows);
  log('Exported: CartAudit_BagCount_View.csv');
}

/* -------------------- export PDF (1 wave per page) -------------------- */
/* Uses header-based highlighting so we never miss Bags due to column index quirks */
async function exportPDF() {
  if (!bagWaves.length) {
    if (pick.length < 2 || sccQR.length < 2) {
      log('Nothing to export. Load files and build BagCount first.');
      return;
    }
    buildBagCount();
  }

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'portrait', unit: 'pt', format: 'a4' });

  const waves = bagWaves;

  const left = 40;
  const topTitle = 40;
  const topTable = 60;

  for (let w = 0; w < waves.length; w++) {
    if (w > 0) doc.addPage('a4', 'portrait');

    const rows = waves[w] || [];

    doc.setFont('helvetica', 'bold');
    doc.setFontSize(14);
    doc.text(`Wave ${w + 1}`, left, topTitle);

    const head = [['Route Code','Location','Carts','Bags','OVs','Departed']];
    const body = rows.map(r => [
      String(r[0] ?? ''), String(r[1] ?? ''),
      String(r[2] ?? 0),  String(r[3] ?? 0),
      String(r[4] ?? 0),  ''
    ]);

    for (let fill = rows.length; fill < 36; fill++) body.push(['','','','','','']);

    doc.autoTable({
      startY: topTable,
      margin: { left },
      head,
      body,
      styles: {
        font: 'helvetica',
        fontSize: 10,
        lineColor: [150, 150, 150],
        lineWidth: 0.5,
        fillColor: [255, 255, 255],
        textColor: [0, 0, 0],
        cellPadding: { top: 3, right: 6, bottom: 3, left: 6 },
        halign: 'left',
        valign: 'middle'
      },
      headStyles: { fillColor: [230, 230, 230], textColor: [0,0,0], fontStyle: 'bold' },
      columnStyles: {
        0: { cellWidth: 140 }, 1: { cellWidth: 140 },
        2: { cellWidth: 55, halign: 'center' },
        3: { cellWidth: 55, halign: 'center' },
        4: { cellWidth: 55, halign: 'center' },
        5: { cellWidth: 70, halign: 'center' }
      },
      didParseCell: function (data) {
        if (data.section !== 'body') return;
        const raw = String(data.cell.raw ?? '');
        const val = Number(raw);
        // Read header text for this column to avoid index issues
        const headRow = data.table?.head?.[0];
        const headCell = headRow ? headRow.cells?.[data.column.index] : null;
        const header = headCell ? String(headCell.raw ?? headCell.content ?? '').toLowerCase() : '';

        if (header.includes('bags') && val === 1) {            // Bags == 1 → yellow
          data.cell.styles.fillColor = [255, 223, 110];
          data.cell.styles.textColor = [17, 17, 17];
          data.cell.styles.fontStyle = 'bold';
        }
        if (header.includes('ov') && val === 3) {               // OVs == 3 → orange
          data.cell.styles.fillColor = [255, 210, 77];
          data.cell.styles.textColor = [17, 17, 17];
          data.cell.styles.fontStyle = 'bold';
        }
      }
    });
  }

  doc.save('CartAudit_BagCount.pdf');
  log('Exported: CartAudit_BagCount.pdf (1 wave per page).');
}

/* -------------------- buttons -------------------- */
document.getElementById('btn-pick').onclick      = () => document.getElementById('file-pick').click();
document.getElementById('btn-scc').onclick       = () => document.getElementById('file-scc').click();
document.getElementById('btn-reset').onclick     = resetAll;
document.getElementById('btn-sccproc').onclick   = processSCC;
document.getElementById('btn-buffer').onclick    = buildBuffer;
document.getElementById('btn-bagcount').onclick  = buildBagCount;
document.getElementById('btn-export').onclick    = exportCSV;
document.getElementById('btn-export-pdf').onclick = exportPDF;
