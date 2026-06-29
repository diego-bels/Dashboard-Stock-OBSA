"""
Paso 2: Lee data.json y genera Dashboard.html interactivo multi-rubro.
"""
import json, sys
from pathlib import Path
from datetime import date

BASE = Path(__file__).parent
DATA_FILE = BASE / 'data.json'

if not DATA_FILE.exists():
    print("ERROR: data.json no encontrado. Ejecutá primero '1_extraer_datos.py'.")
    sys.exit(1)

with open(DATA_FILE, encoding='utf-8') as f:
    data = json.load(f)

data_json = json.dumps(data, ensure_ascii=False)
today_str = date.today().strftime('%d/%m/%Y')
source_file = data.get('source_file', '')
n_products = len(data['products'])
n_rubros = len(data.get('rubros', []))

HTML = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dashboard Stock — Última Unidad Sin Reposición</title>
<script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
<style>
  :root{
    --navy:#1F3864;--blue:#2E75B6;--lightblue:#4472C4;
    --red:#C00000;--red-bg:#FFC7CE;--yellow-bg:#FFEB9C;
    --green-bg:#C6EFCE;--gray-bg:#F2F2F2;--alt-row:#F2F7FF;
    --border:#D0D7E4;--text:#1a1a2e;--sub:#595959;
  }
  *{box-sizing:border-box;margin:0;padding:0}
  body{font-family:Arial,sans-serif;font-size:13px;background:#F4F6FB;color:var(--text)}

  .header{background:var(--navy);color:#fff;padding:12px 20px;display:flex;align-items:center;justify-content:space-between;gap:12px}
  .header h1{font-size:16px;font-weight:bold;white-space:nowrap}
  .header .meta{font-size:11px;color:#a0b4cc;margin-top:2px}

  .layout{display:flex;height:calc(100vh - 52px)}
  .sidebar{width:250px;min-width:200px;background:#fff;border-right:1px solid var(--border);overflow-y:auto;padding:12px;flex-shrink:0}
  .main{flex:1;display:flex;flex-direction:column;overflow:hidden}

  /* SIDEBAR FILTERS */
  .filter-section{margin-bottom:11px}
  .filter-label{font-size:10px;font-weight:bold;color:var(--navy);text-transform:uppercase;letter-spacing:.5px;margin-bottom:4px}
  .search-box{width:100%;border:1px solid var(--border);border-radius:5px;padding:5px 8px;font-size:12px;color:var(--text)}
  .search-box:focus{outline:none;border-color:var(--blue)}
  .ac-wrap{position:relative}
  .ac-dropdown{position:absolute;top:100%;left:0;right:0;background:#fff;border:1px solid var(--blue);border-top:none;border-radius:0 0 5px 5px;max-height:160px;overflow-y:auto;z-index:999;display:none;box-shadow:0 4px 8px rgba(0,0,0,.1)}
  .ac-dropdown.open{display:block}
  .ac-item{padding:5px 8px;font-size:12px;cursor:pointer;color:var(--text)}
  .ac-item:hover,.ac-item.active{background:var(--blue);color:#fff}
  .ac-item em{font-style:normal;font-weight:700}
  .cb-list{display:flex;flex-direction:column;gap:3px;max-height:220px;overflow-y:auto}
  .cb-item{display:flex;align-items:center;gap:5px;padding:2px 0}
  .cb-item input{accent-color:var(--blue);cursor:pointer;flex-shrink:0}
  .cb-item label{cursor:pointer;font-size:11px;flex:1;line-height:1.3}
  .cb-item .badge{font-size:10px;background:var(--alt-row);border-radius:8px;padding:1px 5px;color:var(--navy);font-weight:bold;min-width:20px;text-align:center;flex-shrink:0}
  .cb-item .badge.red{background:var(--red-bg);color:var(--red)}
  .toggle-row{display:flex;align-items:center;gap:7px;margin-bottom:7px}
  .toggle{position:relative;display:inline-block;width:34px;height:18px;flex-shrink:0}
  .toggle input{opacity:0;width:0;height:0}
  .slider{position:absolute;cursor:pointer;inset:0;background:#ccc;border-radius:18px;transition:.25s}
  .slider:before{content:'';position:absolute;width:12px;height:12px;left:3px;bottom:3px;background:#fff;border-radius:50%;transition:.25s}
  input:checked+.slider{background:var(--blue)}
  input:checked+.slider:before{transform:translateX(16px)}
  .toggle-label{font-size:11px;line-height:1.3}
  .divider{height:1px;background:var(--border);margin:8px 0}
  .filter-label-row{display:flex;justify-content:space-between;align-items:center;margin-bottom:4px}
  .filter-label-row .filter-label{margin-bottom:0}
  .lnk{font-size:10px;color:var(--blue);cursor:pointer;font-weight:normal}
  .lnk:hover{text-decoration:underline}
  .reset-btn{width:100%;padding:6px;border-radius:5px;border:1.5px solid var(--border);background:#fff;cursor:pointer;font-size:11px;color:var(--sub);margin-top:4px}
  .reset-btn:hover{background:var(--alt-row);color:var(--navy)}

  /* KPI BAR */
  .kpi-bar{display:flex;gap:8px;padding:8px 14px;background:#fff;border-bottom:1px solid var(--border);flex-shrink:0;flex-wrap:wrap}
  .kpi{background:#E8F0FE;border-radius:7px;padding:7px 12px;min-width:110px}
  .kpi .kpi-val{font-size:18px;font-weight:bold;color:var(--navy)}
  .kpi .kpi-lbl{font-size:10px;color:var(--sub);margin-top:1px}

  /* TOOLBAR */
  .toolbar{display:flex;align-items:center;gap:8px;padding:6px 14px;background:#fff;border-bottom:1px solid var(--border);flex-shrink:0}
  .results-count{font-size:11px;color:var(--sub);flex:1}
  .btn{padding:5px 12px;border-radius:5px;border:none;cursor:pointer;font-size:11px;font-weight:bold;display:flex;align-items:center;gap:5px}
  .btn-outline{background:#fff;color:var(--navy);border:1.5px solid var(--navy)}
  .btn-outline:hover{background:var(--alt-row)}
  .btn-green{background:#197A3E;color:#fff}
  .btn-green:hover{background:#145c2e}
  .view-toggle{display:flex;border:1.5px solid var(--border);border-radius:5px;overflow:hidden}
  .view-btn{padding:4px 10px;border:none;background:#fff;cursor:pointer;font-size:11px;color:var(--sub)}
  .view-btn.active{background:var(--navy);color:#fff}

  /* TABLE */
  .table-wrap{flex:1;overflow:auto}
  table{border-collapse:collapse;width:100%;font-size:11px}
  thead th{background:var(--navy);color:#fff;padding:6px 8px;text-align:center;font-size:10px;position:sticky;top:0;z-index:2;white-space:nowrap;cursor:pointer;user-select:none}
  thead th:hover{background:var(--lightblue)}
  thead th.sort-asc::after{content:' ↑'}
  thead th.sort-desc::after{content:' ↓'}
  tbody tr:hover{background:#ddeeff!important}
  td{padding:4px 7px;border-bottom:1px solid var(--border);white-space:nowrap}
  td.art{white-space:normal;max-width:280px;font-size:11px}
  .stk-cell{text-align:center;border-radius:3px;font-weight:bold;padding:2px 6px;display:inline-block;min-width:24px;font-size:11px}
  .stk-1{background:var(--red-bg);color:var(--red)}
  .stk-0{background:var(--gray-bg);color:#bbb}
  .stk-few{background:var(--yellow-bg);color:#7a5a00}
  .stk-ok{background:var(--green-bg);color:#145c2e}
  .rubro-tag{font-size:9px;background:#E8F0FE;color:var(--blue);padding:1px 5px;border-radius:8px;white-space:nowrap}

  /* SUMMARY CARDS */
  .summary-grid{display:flex;flex-wrap:wrap;gap:10px;padding:14px;overflow-y:auto;flex:1;align-content:flex-start}
  .branch-card{background:#fff;border:1px solid var(--border);border-radius:9px;min-width:210px;flex:1;max-width:320px;overflow:hidden}
  .card-header{background:var(--blue);color:#fff;padding:9px 12px}
  .card-header h3{font-size:13px;font-weight:bold}
  .card-meta{font-size:10px;opacity:.85;margin-top:2px}
  .card-body{padding:6px 10px;max-height:200px;overflow-y:auto}
  .prod-row{display:flex;justify-content:space-between;align-items:baseline;padding:3px 0;border-bottom:1px solid #f0f0f0;gap:6px}
  .prod-row:last-child{border-bottom:none}
  .prod-name{font-size:10px;flex:1;line-height:1.3}
  .prod-rubro{font-size:9px;color:var(--blue);background:#E8F0FE;padding:1px 5px;border-radius:8px;white-space:nowrap}

  /* MODAL */
  .modal-overlay{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:1000;align-items:center;justify-content:center}
  .modal-overlay.open{display:flex}
  .modal{background:#fff;border-radius:10px;padding:22px;max-width:440px;width:90%}
  .modal h2{font-size:14px;color:var(--navy);margin-bottom:12px}
  .form-group{margin-bottom:10px}
  .form-group label{font-size:11px;color:var(--sub);display:block;margin-bottom:3px}
  .form-group input,.form-group select{width:100%;border:1px solid var(--border);border-radius:5px;padding:6px 8px;font-size:12px}
  .modal-actions{display:flex;gap:8px;justify-content:flex-end;margin-top:14px}
  .no-results{text-align:center;padding:50px 20px;color:var(--sub)}
  .no-results .icon{font-size:36px;margin-bottom:8px}
</style>
</head>
<body>

<div class="header">
  <div>
    <h1>📦 Dashboard Stock — Última Unidad Sin Reposición</h1>
    <div class="meta">SOURCE_FILE &nbsp;|&nbsp; Actualizado: TODAY_STR</div>
  </div>
  <button class="btn btn-green" style="flex-shrink:0" onclick="openExportModal()">⬇ Exportar Excel</button>
</div>

<div class="layout">

  <!-- SIDEBAR -->
  <div class="sidebar">

    <div class="filter-section">
      <div class="filter-label">Código / Descripción</div>
      <input class="search-box" id="searchInput" type="text" placeholder="Código o artículo…" oninput="applyFilters()">
    </div>

    <div class="filter-section">
      <div class="filter-label">Familia</div>
      <div class="ac-wrap">
        <input class="search-box" id="familiaInput" type="text" placeholder="Buscar familia…" oninput="acUpdate('familia')" onfocus="acUpdate('familia')" onkeydown="acKey(event,'familia')" autocomplete="off">
        <div class="ac-dropdown" id="familiaDropdown"></div>
      </div>
    </div>

    <div class="filter-section">
      <div class="filter-label">Rubro</div>
      <div class="ac-wrap">
        <input class="search-box" id="rubroInput" type="text" placeholder="Buscar rubro…" oninput="acUpdate('rubro')" onfocus="acUpdate('rubro')" onkeydown="acKey(event,'rubro')" autocomplete="off">
        <div class="ac-dropdown" id="rubroDropdown"></div>
      </div>
    </div>

    <div class="filter-section">
      <div class="filter-label">Marca</div>
      <div class="ac-wrap">
        <input class="search-box" id="marcaInput" type="text" placeholder="Buscar marca…" oninput="acUpdate('marca')" onfocus="acUpdate('marca')" onkeydown="acKey(event,'marca')" autocomplete="off">
        <div class="ac-dropdown" id="marcaDropdown"></div>
      </div>
    </div>

    <div class="divider"></div>

    <div class="filter-section">
      <div class="filter-label-row">
        <span class="filter-label">Sucursales</span>
        <span class="lnk" id="selTodoLbl" onclick="toggleAllBranches()">Desel. todo</span>
      </div>
      <div class="cb-list" id="branchList"></div>
    </div>

    <div class="divider"></div>


    <button class="reset-btn" onclick="resetFilters()">↺ Limpiar todos los filtros</button>
  </div>

  <!-- MAIN -->
  <div class="main">

    <div class="kpi-bar">
      <div class="kpi"><div class="kpi-val" id="kpiProds">—</div><div class="kpi-lbl">Productos</div></div>
      <div class="kpi"><div class="kpi-val" id="kpiRubros">—</div><div class="kpi-lbl">Rubros</div></div>
      <div class="kpi"><div class="kpi-val" id="kpiBranches">—</div><div class="kpi-lbl">Sucursales c/ alerta</div></div>
      <div class="kpi"><div class="kpi-val" id="kpiAlerts">—</div><div class="kpi-lbl">Registros última unidad</div></div>
    </div>

    <div class="toolbar">
      <span class="results-count" id="resultsCount"></span>
      <div class="view-toggle">
        <button class="view-btn active" id="btnTable"   onclick="setView('table')">≡ Tabla</button>
        <button class="view-btn"        id="btnSummary" onclick="setView('summary')">⊞ Por Sucursal</button>
      </div>
    </div>

    <div class="table-wrap" id="tableView">
      <table id="dataTable">
        <thead id="tableHead"></thead>
        <tbody id="tableBody"></tbody>
      </table>
      <div class="no-results" id="noResults" style="display:none">
        <div class="icon">🔍</div><div>Sin resultados para los filtros aplicados.</div>
      </div>
    </div>

    <div class="summary-grid" id="summaryView" style="display:none"></div>
  </div>
</div>

<!-- EXPORT MODAL -->
<div class="modal-overlay" id="exportModal">
  <div class="modal">
    <h2>⬇ Exportar a Excel</h2>
    <div class="form-group">
      <label>Título del informe</label>
      <input type="text" id="exportTitle" value="Stock - Última Unidad Sin Reposición">
    </div>
    <div class="form-group">
      <label>Destinatario (aparece en el nombre del archivo)</label>
      <input type="text" id="exportDest" placeholder="Ej: Responsable Famailla">
    </div>
    <div class="form-group">
      <label>Sucursales a incluir</label>
      <select id="exportMode">
        <option value="selected">Solo sucursales seleccionadas en filtros</option>
        <option value="alerts">Solo sucursales con última unidad (filtro actual)</option>
        <option value="all">Todas las sucursales</option>
      </select>
    </div>
    <div class="modal-actions">
      <button class="btn btn-outline" onclick="closeExportModal()">Cancelar</button>
      <button class="btn btn-green"   onclick="exportExcel()">⬇ Descargar</button>
    </div>
  </div>
</div>

<script>
const RAW = DATA_PLACEHOLDER;
const ALL_BRANCHES = RAW.branches;
const PRODUCTS = RAW.products;

let selectedBranches = new Set(ALL_BRANCHES);
let sortCol = null, sortDir = 1, currentView = 'table';

// ── INIT ─────────────────────────────────────────────────────────────────────
function initFilters() { buildBranchList(); }

function buildBranchList(filteredProds) {
  const el = document.getElementById('branchList');
  el.innerHTML = '';
  const prods = filteredProds || PRODUCTS;
  const counts = {};
  prods.forEach(p => ALL_BRANCHES.forEach(b => {
    if ((p.branch_stocks[b]||0) === 1) counts[b] = (counts[b]||0) + 1;
  }));
  ALL_BRANCHES.forEach(branch => {
    const cnt = counts[branch] || 0;
    const safe = branch.replace(/'/g,"\\'");
    const row = document.createElement('div');
    row.className = 'cb-item';
    row.innerHTML = `<input type="checkbox" id="cb_${branch}" ${selectedBranches.has(branch)?'checked':''}
      onchange="toggleBranch('${safe}')"><label for="cb_${branch}">${branch}</label>
      <span class="badge ${cnt>0?'red':''}">${cnt}</span>`;
    el.appendChild(row);
  });
}

function toggleBranch(b) {
  selectedBranches.has(b) ? selectedBranches.delete(b) : selectedBranches.add(b);
  updateSelTodoLbl(); applyFilters();
}
function updateSelTodoLbl() {
  document.getElementById('selTodoLbl').textContent =
    selectedBranches.size === ALL_BRANCHES.length ? 'Desel. todo' : 'Sel. todo';
}
function toggleAllBranches() {
  if (selectedBranches.size === ALL_BRANCHES.length) {
    selectedBranches.clear();
    document.querySelectorAll('#branchList input').forEach(c => c.checked = false);
  } else {
    ALL_BRANCHES.forEach(b => selectedBranches.add(b));
    document.querySelectorAll('#branchList input').forEach(c => c.checked = true);
  }
  updateSelTodoLbl(); applyFilters();
}

// ── AUTOCOMPLETE ──────────────────────────────────────────────────────────────
// Selecciones exactas (via autocomplete click) vs. texto libre
const acExact = { rubro: false, familia: false, marca: false };

// Mapa familia → rubros para filtrado en cascada
const FAMILIA_RUBROS = {};
PRODUCTS.forEach(p => {
  if (!p.familia || !p.rubro) return;
  if (!FAMILIA_RUBROS[p.familia]) FAMILIA_RUBROS[p.familia] = new Set();
  FAMILIA_RUBROS[p.familia].add(p.rubro);
});

const AC_FIELDS = {
  rubro:   { inputId:'rubroInput',   dropId:'rubroDropdown',   values: [...new Set(PRODUCTS.map(p=>p.rubro).filter(Boolean))].sort() },
  familia: { inputId:'familiaInput', dropId:'familiaDropdown', values: [...new Set(PRODUCTS.map(p=>p.familia).filter(Boolean))].sort() },
  marca:   { inputId:'marcaInput',   dropId:'marcaDropdown',   values: [...new Set(PRODUCTS.map(p=>p.marca).filter(Boolean))].sort() },
};
let acActiveIdx = {};

function acHighlight(q, text) {
  if (!q) return text;
  const i = text.toLowerCase().indexOf(q.toLowerCase());
  if (i < 0) return text;
  return text.slice(0,i) + '<em>' + text.slice(i, i+q.length) + '</em>' + text.slice(i+q.length);
}

function acGetRubroValues() {
  const familiaVal = document.getElementById('familiaInput').value.trim();
  if (familiaVal && FAMILIA_RUBROS[familiaVal]) {
    return [...FAMILIA_RUBROS[familiaVal]].sort();
  }
  return AC_FIELDS.rubro.values;
}

function acUpdate(field) {
  const f = AC_FIELDS[field];
  const input = document.getElementById(f.inputId);
  const drop  = document.getElementById(f.dropId);
  const q = input.value.trim();
  const pool = field === 'rubro' ? acGetRubroValues() : f.values;
  const matches = q ? pool.filter(v => v.toLowerCase().includes(q.toLowerCase())) : pool;
  acActiveIdx[field] = -1;
  if (!matches.length) { drop.classList.remove('open'); applyFilters(); return; }
  drop.innerHTML = matches.slice(0,50).map((v,i) =>
    `<div class="ac-item" data-val="${v.replace(/"/g,'&quot;')}" onmousedown="acSelect('${field}','${v.replace(/'/g,"\\'")}')">
      ${acHighlight(q, v)}
    </div>`
  ).join('');
  drop.classList.add('open');
  acExact[field] = false;  // escritura libre → no es exact
  applyFilters();
}

function acSelect(field, value) {
  const f = AC_FIELDS[field];
  document.getElementById(f.inputId).value = value;
  document.getElementById(f.dropId).classList.remove('open');
  acExact[field] = true;  // selección desde dropdown → exact match
  // Al cambiar familia, limpiar rubro si ya no pertenece a esa familia
  if (field === 'familia') {
    const rubroVal = document.getElementById('rubroInput').value.trim();
    if (rubroVal && FAMILIA_RUBROS[value] && !FAMILIA_RUBROS[value].has(rubroVal)) {
      document.getElementById('rubroInput').value = '';
    }
  }
  applyFilters();
}

function acClose(field) {
  document.getElementById(AC_FIELDS[field].dropId).classList.remove('open');
}

function acKey(e, field) {
  const drop = document.getElementById(AC_FIELDS[field].dropId);
  const items = drop.querySelectorAll('.ac-item');
  if (!items.length) return;
  if (e.key === 'ArrowDown') {
    e.preventDefault();
    acActiveIdx[field] = Math.min((acActiveIdx[field]||0)+1, items.length-1);
  } else if (e.key === 'ArrowUp') {
    e.preventDefault();
    acActiveIdx[field] = Math.max((acActiveIdx[field]||0)-1, 0);
  } else if (e.key === 'Enter') {
    e.preventDefault();
    const idx = acActiveIdx[field];
    if (idx >= 0 && items[idx]) { acSelect(field, items[idx].dataset.val); return; }
    drop.classList.remove('open'); applyFilters(); return;
  } else if (e.key === 'Escape') {
    drop.classList.remove('open'); return;
  } else { return; }
  items.forEach((it,i) => it.classList.toggle('active', i === acActiveIdx[field]));
  if (items[acActiveIdx[field]]) items[acActiveIdx[field]].scrollIntoView({block:'nearest'});
}

document.addEventListener('click', e => {
  ['rubro','familia','marca'].forEach(field => {
    if (!document.getElementById(AC_FIELDS[field].inputId).contains(e.target) &&
        !document.getElementById(AC_FIELDS[field].dropId).contains(e.target)) {
      acClose(field);
    }
  });
});

// ── FILTER ────────────────────────────────────────────────────────────────────
function getFilteredProducts() {
  const search  = document.getElementById('searchInput').value.toLowerCase().trim();
  const rubro   = document.getElementById('rubroInput').value.toLowerCase().trim();
  const familia = document.getElementById('familiaInput').value.toLowerCase().trim();
  const marca   = document.getElementById('marcaInput').value.toLowerCase().trim();

  return PRODUCTS.filter(p => {
    if (search  && !p.codigo.toLowerCase().includes(search)  && !p.articulo.toLowerCase().includes(search))  return false;
    if (rubro)   { const rv = p.rubro.toLowerCase();   if (acExact.rubro   ? rv !== rubro   : !rv.includes(rubro))   return false; }
    if (familia) { const fv = p.familia.toLowerCase(); if (acExact.familia ? fv !== familia : !fv.includes(familia)) return false; }
    if (marca)   { const mv = p.marca.toLowerCase();   if (acExact.marca   ? mv !== marca   : !mv.includes(marca))   return false; }
    return true;
  });
}

function applyFilters() {
  const prods = getFilteredProducts();
  updateKPIs(prods);
  buildBranchList(prods);
  if (currentView === 'table') renderTable(prods);
  else renderSummary(prods);
}

// ── KPIs ──────────────────────────────────────────────────────────────────────
function updateKPIs(prods) {
  const active = [...selectedBranches];
  let alerts = 0; const bset = new Set(), rset = new Set();
  prods.forEach(p => {
    rset.add(p.rubro);
    active.forEach(b => { if ((p.branch_stocks[b]||0)===1){ alerts++; bset.add(b); } });
  });
  document.getElementById('kpiProds').textContent    = prods.length.toLocaleString('es-AR');
  document.getElementById('kpiRubros').textContent   = rset.size;
  document.getElementById('kpiBranches').textContent = bset.size;
  document.getElementById('kpiAlerts').textContent   = alerts.toLocaleString('es-AR');
  document.getElementById('resultsCount').textContent = prods.length.toLocaleString('es-AR') + ' producto(s) encontrado(s)';
}

// ── TABLE ─────────────────────────────────────────────────────────────────────
function stkCell(v) {
  const cls = v===0?'stk-0':v===1?'stk-1':v<5?'stk-few':'stk-ok';
  return `<span class="stk-cell ${cls}">${v}</span>`;
}
function getColVal(p, i, branches) {
  switch(i){case 0:return p.codigo;case 1:return p.rubro;case 2:return p.articulo;
    case 3:return p.marca;case 4:return p.col_stock;default:return p.branch_stocks[branches[i-5]]||0;}
}
function sortBy(i){ sortCol===i?sortDir*=-1:(sortCol=i,sortDir=1); applyFilters(); }

function renderTable(prods) {
  const active = [...selectedBranches];
  const cols = ['Código','Rubro','Artículo','Marca','Stock COL',...active];
  document.getElementById('tableHead').innerHTML = '<tr>'+cols.map((c,i)=>
    `<th onclick="sortBy(${i})" class="${sortCol===i?(sortDir>0?'sort-asc':'sort-desc'):''}">${c}</th>`
  ).join('')+'</tr>';

  const noRes = document.getElementById('noResults');
  const tbl   = document.getElementById('dataTable');
  if (!prods.length) {
    document.getElementById('tableBody').innerHTML = '';
    noRes.style.display = 'block'; tbl.style.display = 'none'; return;
  }
  noRes.style.display = 'none'; tbl.style.display = '';

  let sorted = [...prods];
  if (sortCol !== null) sorted.sort((a,b) => {
    const av = getColVal(a,sortCol,active), bv = getColVal(b,sortCol,active);
    return (av>bv?1:av<bv?-1:0)*sortDir;
  });

  document.getElementById('tableBody').innerHTML = sorted.map((p,ri) =>
    `<tr style="${ri%2?'background:#F2F7FF':''}">
      <td>${p.codigo}</td>
      <td><span class="rubro-tag">${p.rubro}</span></td>
      <td class="art">${p.articulo}</td>
      <td style="text-align:center">${p.marca}</td>
      <td style="text-align:center">${stkCell(0)}</td>
      ${active.map(b=>`<td style="padding:2px 4px;text-align:center">${stkCell(p.branch_stocks[b]||0)}</td>`).join('')}
    </tr>`
  ).join('');
}

// ── SUMMARY ───────────────────────────────────────────────────────────────────
function renderSummary(prods) {
  const active = [...selectedBranches];
  const byBranch = {}; active.forEach(b => byBranch[b] = []);
  prods.forEach(p => active.forEach(b => { if((p.branch_stocks[b]||0)===1) byBranch[b].push(p); }));
  const entries = active.filter(b=>byBranch[b].length>0).sort((a,b)=>byBranch[b].length-byBranch[a].length);
  const container = document.getElementById('summaryView');
  if (!entries.length) {
    container.innerHTML='<div class="no-results"><div class="icon">🔍</div><div>Sin alertas para los filtros aplicados.</div></div>'; return;
  }
  container.innerHTML = entries.map(branch => {
    const ps = byBranch[branch];
    return `<div class="branch-card">
      <div class="card-header"><h3>${branch}</h3><div class="card-meta">${ps.length} producto(s)</div></div>
      <div class="card-body">${ps.map(p=>`<div class="prod-row">
        <span class="prod-name">${p.articulo.substring(0,52)}${p.articulo.length>52?'…':''}</span>
        <span class="prod-rubro">${p.rubro}</span>
      </div>`).join('')}</div>
    </div>`;
  }).join('');
}

// ── VIEW TOGGLE ───────────────────────────────────────────────────────────────
function setView(v) {
  currentView = v;
  document.getElementById('tableView').style.display   = v==='table'  ?'flex':'none';
  document.getElementById('summaryView').style.display  = v==='summary'?'flex':'none';
  document.getElementById('btnTable').classList.toggle('active',   v==='table');
  document.getElementById('btnSummary').classList.toggle('active', v==='summary');
  applyFilters();
}

// ── EXPORT ────────────────────────────────────────────────────────────────────
function openExportModal()  { document.getElementById('exportModal').classList.add('open'); }
function closeExportModal() { document.getElementById('exportModal').classList.remove('open'); }

function exportExcel() {
  const title = document.getElementById('exportTitle').value || 'Stock';
  const dest  = document.getElementById('exportDest').value;
  const mode  = document.getElementById('exportMode').value;
  const prods = getFilteredProducts();

  let exportBranches;
  if (mode==='selected') exportBranches = [...selectedBranches];
  else if (mode==='all') exportBranches = ALL_BRANCHES;
  else exportBranches = ALL_BRANCHES.filter(b => prods.some(p=>(p.branch_stocks[b]||0)===1));

  const wb = XLSX.utils.book_new();

  // ── Hoja única: una fila por sucursal × producto con última unidad ──
  const rows = [['Sucursal', 'Cant. Registros', 'Código', 'Descripción', 'Rubro', 'Marca']];

  exportBranches.forEach(branch => {
    const ps = prods.filter(p => (p.branch_stocks[branch]||0) === 1);
    if (!ps.length) return;
    const cnt = ps.length;
    ps.forEach(p => rows.push([branch, cnt, p.codigo, p.articulo, p.rubro, p.marca]));
  });

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [{wch:16},{wch:16},{wch:10},{wch:54},{wch:18},{wch:14}];
  XLSX.utils.book_append_sheet(wb, ws, 'Última Unidad por Sucursal');

  const dateStr = new Date().toISOString().slice(0,10);
  const destPart = dest ? `_${dest.replace(/[^a-zA-Z0-9áéíóúÁÉÍÓÚñÑ\s]/g,'').replace(/\s+/g,'_')}` : '';
  XLSX.writeFile(wb, `${title.replace(/[\\/:*?"<>|]/g,'').replace(/\s+/g,'_')}${destPart}_${dateStr}.xlsx`);
  closeExportModal();
}

// ── RESET ─────────────────────────────────────────────────────────────────────
function resetFilters() {
  ['searchInput','rubroInput','familiaInput','marcaInput'].forEach(id => document.getElementById(id).value = '');
  ['rubro','familia','marca'].forEach(acClose);
  selectedBranches = new Set(ALL_BRANCHES);
  sortCol = null; sortDir = 1;
  buildBranchList(); applyFilters();
}

initFilters();
applyFilters();
</script>
</body>
</html>"""

HTML = (HTML
    .replace('TODAY_STR', today_str)
    .replace('SOURCE_FILE', source_file)
    .replace('DATA_PLACEHOLDER', data_json))

out_path = BASE / 'Dashboard.html'
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(HTML)

print(f"Dashboard generado: {out_path.name}")
print(f"  Productos: {n_products:,}  |  Rubros: {n_rubros}")
