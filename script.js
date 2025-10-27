/* === CartAudit Smart Script v13.7 — REMOVED OLD BUTTON LISTENERS === */
/* === Based on v13.6, but removes listeners for non-existent buttons === */

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
      // Basic outer quote stripping
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

/* --- ADDED: Robust header normalization helper --- */
function normHeaderCell(x) {
  let s = String(x ?? '');
  s = s.replace(/\ufeff/g, '');   // 1. strip BOM
  s = s.trim();                   // 2. TRIM WHITESPACE FIRST
  s = s.replace(/^"+|"+$/g, ''); // 3. THEN STRIP QUOTES
  s = s.replace(/\s+/g, ' ');    // 4. collapse ALL internal whitespace (newlines, tabs, spaces) to a single space
  s = s.trim();                   // 5. trim again, just in case
  s = s.toLowerCase();            // 6. lowercase
  return s;
}

/* --- NEW: Header Index Finder --- */
// Searches cleaned headers for a list of possible variations
function findHeaderIndex(cleanedHeaders, variations) {
  for (const variation of variations) {
    const index = cleanedHeaders.indexOf(variation);
    if (index !== -1) {
      return index; // Return index of first match found
    }
  }
  return -1; // Not found
}


/* -------------------- file loaders (Original v13.5 Logic) -------------------- */
// PickOrder (dispatchTime, routeCode, routeID, dispatchArea)
document.getElementById('file-pick').addEventListener('change', e => {
  const f = e.target.files[0];
  if (!f) return;
  f.text().then(t => {
    let raw = normalizeWeirdCSV(t); // Using original normalize
    let head = raw.split(/\r?\n/, 1)[0] || '';
    let delimiter = head.includes(';') ? ';' : (head.includes('\t') ? '\t' : ',');
    pick = parseCSV(raw, delimiter); // Using original parse

    // Original re-parse logic if needed
    if ((pick[0]?.length || 0) === 1 && /",\s*"/.test(head)) {
       log('Re-parsing PickOrder with potential quote issue.');
       raw = t; // Use original text if normalize failed
       head = raw.split(/\r?\n/, 1)[0] || '';
       delimiter = head.includes(';') ? ';' : (head.includes('\t') ? '\t' : ',');
       pick = parseCSV(raw, delimiter);
    }

    log(`Loaded PickOrder: ${pick.length - 1} data rows, delimiter="${delimiter}", cols=${pick[0]?.length || 0}`);
  });
  e.target.value = null; // Clear input
});

// SCCPick (Route Code, Picklist Code, …, Bags, OVs, SPR, Progress, Associate, Duration, Type)
document.getElementById('file-scc').addEventListener('change', e => {
  const f = e.target.files[0];
  if (!f) return;
  f.text().then(t => {
    let raw = normalizeWeirdCSV(t); // Using original normalize
    let head = raw.split(/\r?\n/, 1)[0] || '';
    let delimiter = head.includes(';') ? ';' : (head.includes('\t') ? '\t' : ',');
    scc = parseCSV(raw, delimiter); // Using original parse

    // Original re-parse logic if needed
    if ((scc[0]?.length || 0) === 1 && /",\s*"/.test(head)) {
       log('Re-parsing SCCPick with potential quote issue.');
       raw = t; // Use original text if normalize failed
       head = raw.split(/\r?\n/, 1)[0] || '';
       delimiter = head.includes(';') ? ';' : (head.includes('\t') ? '\t' : ',');
       scc = parseCSV(raw, delimiter);
    }

    log(`Loaded SCCPick: ${scc.length - 1} data rows, delimiter="${delimiter}", cols=${scc[0]?.length || 0}`);
    // Removed old useless header check log
  });
  e.target.value = null; // Clear input
});

/* -------------------- processing (MODIFIED TO FIND COLUMNS BY NAME) -------------------- */
function processSCC() {
  if (scc.length < 2) { log('Error: Load SCCPick file first.'); return; }

  const H = scc[0] || [];
  // Ensure H is treated as an array even if parsing resulted in a single string
  const headerArray = Array.isArray(H) ? H : [H]; 
  const normH = headerArray.map(normHeaderCell); // Normalize headers

  // === SEARCH FOR VARIATIONS ===
  const routeVariations    = ['route code', 'routecode'];
  const picklistVariations = ['picklist code', 'picklistcode'];
  const bagsVariations     = ['bags']; // Assuming 'bags' is consistent
  const ovsVariations      = ['ovs', 'ov'];  // Added 'ov' just in case

  const idxRoute    = findHeaderIndex(normH, routeVariations);
  const idxPicklist = findHeaderIndex(normH, picklistVariations);
  const idxBags     = findHeaderIndex(normH, bagsVariations);
  const idxOVs      = findHeaderIndex(normH, ovsVariations);
  // === END SEARCH ===

  // --- Error Checking ---
  const missing = [];
  if (idxRoute === -1)    missing.push(routeVariations.join(' / '));
  if (idxPicklist === -1) missing.push(picklistVariations.join(' / '));
  if (idxBags === -1)     missing.push(bagsVariations.join(' / '));
  if (idxOVs === -1)      missing.push(ovsVariations.join(' / '));

  if (missing.length > 0) {
    log(`Error: SCCPick file is missing required columns: ${missing.join(', ')}`);
    log(`       (This is what the script found after cleaning the headers:)`);
    log(`       Found headers: [${normH.join(' | ')}]`);
    sccQR = []; // Clear previous results
    return; // Stop processing
  }
  // --- End Error Checking ---

  log(`Processing SCCPick. Using dynamic columns: route code=${idxRoute}, picklist code=${idxPicklist}, bags=${idxBags}, ovs=${idxOVs}`);

  const routeSet = new Set();
  for (let i = 1; i < scc.length; i++) {
    const r = scc[i];
    // Add safety check: ensure 'r' is an array before accessing indices
    if (!Array.isArray(r) || r.length <= idxRoute) continue; 
    const route = (r[idxRoute] || '').trim(); // Use DYNAMIC index
    if (route) routeSet.add(route);
  }

  const out = [['Route Code','Carts','Bags','OVs']];
  const rows = scc.slice(1);

  for (const route of routeSet) {
    let carts = 0, bags = 0, ovs = 0;

    for (const r of rows) {
      if (!Array.isArray(r) || r.length <= Math.max(idxPicklist, idxBags, idxOVs)) continue; 
      const picklist = (r[idxPicklist] || '').trim(); // Use DYNAMIC index

      if (picklist && picklist.startsWith(route + '#')) {
        carts++;
        // Use DYNAMIC indices
        const b = Number(String(r[idxBags] || '0').replace(/[^0-9.-]/g, '')) || 0; 
        const o = Number(String(r[idxOVs] || '0').replace(/[^0-9.-]/g, '')) || 0; 
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

/* -------------------- builders / rendering (Original v13.5 Logic, uses indexOf) -------------------- */
// Added normHeaderCell to PickOrder header lookup
function buildBuffer() {
  if (pick.length < 2) { log('Error: Load PickOrder first.'); return; }
  if (sccQR.length < 2) { log('Run SCC processing first.'); return; }

  const hdr = pick[0] || [];
  const headerArray = Array.isArray(hdr) ? hdr : [hdr]; 
  const normH = headerArray.map(normHeaderCell); // Normalize PickOrder headers too

  // === SEARCH FOR VARIATIONS in PickOrder ===
  const routeVariations = ['route code', 'routecode'];
  const areaVariations  = ['dispatch area', 'dispatcharea'];

  let idxRoute = findHeaderIndex(normH, routeVariations);
  let idxArea  = findHeaderIndex(normH, areaVariations);
  // === END SEARCH ===
  
  let routeCol = idxRoute;
  let areaCol = idxArea;

  if (idxRoute === -1 || idxArea === -1) {
    const missing = [];
    if (idxRoute === -1) missing.push(routeVariations.join(' / '));
    if (idxArea === -1) missing.push(areaVariations.join(' / '));
    log(`⚠️ PickOrder header names not found (${missing.join(', ')}). Falling back to columns 1 and 3.`);
    routeCol = 1; // Fallback
    areaCol = 3;  // Fallback
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
     if (!Array.isArray(r) || r.length <= Math.max(routeCol, areaCol)) continue; 
    const route = (r[routeCol] || '').trim(); // Use found/fallback index
    const loc   = (r[areaCol] || '').trim();  // Use found/fallback index
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
   const headerArray = Array.isArray(hdr) ? hdr : [hdr]; 
  const normH = headerArray.map(normHeaderCell); // Normalize PickOrder headers too

  // === SEARCH FOR VARIATIONS in PickOrder ===
  const routeVariations = ['route code', 'routecode'];
  const areaVariations  = ['dispatch area', 'dispatcharea'];

  let idxRoute = findHeaderIndex(normH, routeVariations);
  let idxArea  = findHeaderIndex(normH, areaVariations);
  // === END SEARCH ===
  
  let routeCol = idxRoute;
  let areaCol = idxArea;

  if (idxRoute === -1 || idxArea === -1) {
    const missing = [];
    if (idxRoute === -1) missing.push(routeVariations.join(' / '));
    if (idxArea === -1) missing.push(areaVariations.join(' / '));
    log(`⚠️ PickOrder header names not found (${missing.join(', ')}). Falling back to columns 1 and 3.`);
    routeCol = 1; // Fallback
    areaCol = 3;  // Fallback
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
     if (!Array.isArray(r) || r.length <= Math.max(routeCol, areaCol)) continue; 
    const route = (r[routeCol] || '').trim(); // Use found/fallback index
    const loc   = (r[areaCol] || '').trim();  // Use found/fallback index
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

/* -------------------- highlight helper (Original) -------------------- */
function cellClassByValue(v) {
  if (Number(v) === 1) return 'hl-bag';   // yellow for value==1
  if (Number(v) === 3) return 'hl-ovs';   // orange for value==3
  return '';
}

/* -------------------- render (Original) -------------------- */
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
      // Add safety check for 'r' being an array inside the loop too
      if (!Array.isArray(r)) return; 
      
      if (showAllCols) {
        // Assume r has at least 5 elements after processing
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
         // Assume r has at least 3 elements after processing
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

/* -------------------- export (Original, maybe needs check?) -------------------- */
// This *might* fail if buildBagCount hasn't run, let's add the checks from later versions
function exportCSV() {
  if (pick.length < 2 || scc.length < 2) {
    log('Nothing to export. Load files and generate CartAudit first.');
    return;
  }
  if (!bagWaves.length) {
    log('Building data before export...');
    processSCC(); // Ensure SCC is processed
    if (sccQR.length === 0) { log('Processing failed. Stopping export.'); return; }
    buildBagCount(); // Ensure BagCount is built
     if (!bagWaves.length) { log('Build failed. Stopping export.'); return; }
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
      const item = waves[w]?.[r]; // Add safety check
      if (item && Array.isArray(item)) { // Add safety check
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

/* -------------------- export PDF (Original, maybe needs check?) -------------------- */
// Add safety checks
async function exportPDF() {
   if (pick.length < 2 || scc.length < 2) {
    log('Nothing to export. Load files and generate CartAudit first.');
    return;
  }
  if (!bagWaves.length) {
    log('Building data before export...');
    processSCC();
    if (sccQR.length === 0) { log('Processing failed. Stopping export.'); return; }
    buildBagCount();
     if (!bagWaves.length) { log('Build failed. Stopping export.'); return; }
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
    const body = rows.map(r => Array.isArray(r) ? [ // Add safety check
      String(r[0] ?? ''), String(r[1] ?? ''),
      String(r[2] ?? 0),  String(r[3] ?? 0),
      String(r[4] ?? 0),  ''
    ] : ['','','','','','']); // Provide default if row is bad

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
// Attach clicks only to buttons that exist in the *current* HTML
document.getElementById('btn-pick').onclick      = () => document.getElementById('file-pick').click();
document.getElementById('btn-scc').onclick       = () => document.getElementById('file-scc').click();
document.getElementById('btn-reset').onclick     = resetAll;

// Combined generate button logic
document.getElementById('btn-generate').onclick = () => {
  if (pick.length === 0 || scc.length === 0) {
    log('Error: Both files must be loaded first.');
    return;
  }
  processSCC(); // This now finds columns dynamically
  if (sccQR.length === 0) {
    log('Processing failed. Stopping build.');
    return;
  }
  buildBuffer();
   if (bufferWaves.length === 0 && pick.length > 1) { 
      log('Build Buffer failed. Stopping.');
      return;
  }
  buildBagCount();
};

document.getElementById('btn-export').onclick    = exportCSV;
document.getElementById('btn-export-pdf').onclick = exportPDF;

// Keep Debug button if it exists in the HTML
const debugBtn = document.getElementById('btn-debug');
if (debugBtn) {
    debugBtn.onclick = function debugHeaders() {
      log('--- DEBUGGING HEADERS ---');
      
      if (pick.length > 0) {
        const H = pick[0] || [];
        const headerArray = Array.isArray(H) ? H : [H]; 
        const normH = headerArray.map(normHeaderCell);
        log('PickOrder Headers (Cleaned):');
        log(`[${normH.join(' | ')}]`);
      } else {
        log('PickOrder file is not loaded or is empty.');
      }

      if (scc.length > 0) {
        const H = scc[0] || [];
        const headerArray = Array.isArray(H) ? H : [H]; 
        const normH = headerArray.map(normHeaderCell);
        log('SCCPick Headers (Cleaned):');
        log(`[${normH.join(' | ')}]`);
      } else {
        log('SCCPick file is not loaded or is empty.');
      }
      log('--- END DEBUG ---');
    }
} else {
    log('Debug button not found in HTML.');
}