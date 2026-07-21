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

# ── CREDENCIALES POR SUCURSAL ──────────────────────────────────────────────────
# usuario → { password, sucursal }
# sucursal: nombre exacto del depósito, o None para ver TODAS (admin)
CREDENCIALES = {
    'admin':        {'password': 'obsa2025',      'sucursal': None,  'familias': None, 'titulo': None},
    'hogar':        {'password': 'hogar2025',     'sucursal': None,  'familias': [
                        'APARATOS DE GIMNASIA','AUDIO','BEBES','CAMPING Y JARDIN',
                        'CLIMATIZACION','COLCHONES Y SOMMIERS','EQUIPAMIENTOS COMERCIALES',
                        'FOTOGRAFIA','HERRAMIENTAS','HOGAR','INFORMATICA','JUGUETERIA',
                        'MUEBLES','MUEBLES DE OFICINA','PEQUEÑOS ELECTRODOMESTICOS',
                        'RODADOS','T.V. / DVD','TELEFONIA','VIDEOJUEGOS',
                    ], 'titulo': 'Dashboard Ú.UNIDAD — Hogar'},
    'indumen':      {'password': 'indumen2025',   'sucursal': None,  'familias': [
                        'INDUMENTARIA','CALZADO','BAZAR','BLANCO',
                    ], 'titulo': 'Dashboard Ú.UNIDAD — Indumen/Calz/Otros'},
    '25demayo':     {'password': 'mayo2025',       'sucursal': '25 de Mayo'},
    'tafi':         {'password': 'tafi2025',       'sucursal': 'Tafi del Valle'},
    'banda':        {'password': 'banda2025',      'sucursal': 'Banda'},
    'maipu':        {'password': 'maipu2025',      'sucursal': 'Maipú'},
    'concepcion':   {'password': 'conc2025',       'sucursal': 'Concepción'},
    'lules':        {'password': 'lules2025',      'sucursal': 'Lules'},
    'outlet':       {'password': 'outlet2025',     'sucursal': 'Outlet'},
    'rosario':      {'password': 'rosario2025',    'sucursal': 'Rosario'},
    'cafayate':     {'password': 'cafa2025',       'sucursal': 'Cafayate'},
    'yerbabuena':   {'password': 'yb2025',         'sucursal': 'Yerba Buena'},
    'monteros':     {'password': 'mont2025',       'sucursal': 'Monteros'},
    'tafviejo':     {'password': 'tafv2025',       'sucursal': 'Tafi Viejo'},
    'famailla':     {'password': 'fama2025',       'sucursal': 'Famailla'},
    'santamaria':   {'password': 'santa2025',      'sucursal': 'Santa María'},
    'big':          {'password': 'big2025',        'sucursal': 'Big'},
    'aguilares':    {'password': 'agui2025',       'sucursal': 'Aguilares'},
    'alberdi':      {'password': 'albe2025',       'sucursal': 'Alberdi'},
}
creds_json = json.dumps(CREDENCIALES, ensure_ascii=False)

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

  /* LOGIN */
  .login-overlay{position:fixed;inset:0;background:var(--navy);z-index:2000;display:flex;align-items:center;justify-content:center}
  .login-overlay.hidden{display:none}
  .login-box{background:#fff;border-radius:14px;padding:36px 32px;width:340px;max-width:95vw;box-shadow:0 8px 32px rgba(0,0,0,.3)}
  .login-logo{text-align:center;margin-bottom:20px}
  .login-logo h2{color:var(--navy);font-size:20px;font-weight:bold;margin-top:8px}
  .login-logo p{color:var(--sub);font-size:12px;margin-top:4px}
  .login-field{margin-bottom:14px}
  .login-field label{display:block;font-size:11px;font-weight:bold;color:var(--navy);text-transform:uppercase;letter-spacing:.5px;margin-bottom:5px}
  .login-field input{width:100%;border:1.5px solid var(--border);border-radius:7px;padding:9px 12px;font-size:13px;color:var(--text);outline:none;transition:.2s}
  .login-field input:focus{border-color:var(--blue)}
  .login-btn{width:100%;padding:10px;background:var(--navy);color:#fff;border:none;border-radius:7px;font-size:13px;font-weight:bold;cursor:pointer;margin-top:6px;transition:.2s}
  .login-btn:hover{background:var(--blue)}
  .login-error{color:#C00000;font-size:12px;text-align:center;margin-top:10px;min-height:18px}
  .logout-btn{padding:5px 12px;border-radius:5px;border:1.5px solid rgba(255,255,255,.4);background:transparent;color:#fff;cursor:pointer;font-size:11px}
  .logout-btn:hover{background:rgba(255,255,255,.15)}
  .session-info{font-size:11px;color:#a0b4cc;margin-right:8px}
  .tab-bar{display:flex;gap:0;border-bottom:2px solid var(--border);background:#fff;padding:0 14px;flex-shrink:0}
  .tab-btn{padding:8px 16px;border:none;background:none;cursor:pointer;font-size:12px;color:var(--sub);border-bottom:2px solid transparent;margin-bottom:-2px;font-weight:bold}
  .tab-btn.active{color:var(--navy);border-bottom-color:var(--navy)}
  .tab-btn:hover{color:var(--navy)}

  /* PANEL VENTAS */
  .ventas-wrap{flex:1;overflow:auto;padding:14px}
  .ventas-summary{display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap}
  .ventas-kpi{background:#E8F0FE;border-radius:7px;padding:6px 14px;font-size:12px;color:var(--navy)}
  .ventas-kpi b{font-size:16px;display:block}

  /* DATE RANGE PICKER */
  .drp-trigger{background:#E8F0FE;color:var(--navy);border:1px solid var(--border);border-radius:7px;padding:7px 10px;font-size:12px;font-weight:600;display:flex;align-items:center;gap:6px;width:100%;cursor:pointer;white-space:nowrap}
  .drp-trigger:hover{background:#d0defa}
  .drp-wrap{position:relative}
  .drp-panel{position:fixed;z-index:9999;background:#fff;border-radius:12px;box-shadow:0 8px 32px rgba(0,0,0,.22);display:none;flex-direction:column;border:1px solid var(--border)}
  .drp-right{padding:12px;display:flex;flex-direction:column;gap:8px}
  .drp-cals{display:flex;gap:16px}
  .drp-mhdr{display:flex;justify-content:space-between;align-items:center;font-weight:700;font-size:12px;margin-bottom:5px;color:var(--navy)}
  .drp-grid{display:grid;grid-template-columns:repeat(7,30px);gap:2px}
  .drp-dh{width:30px;height:22px;display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;color:#90a4ae}
  .drp-day{width:30px;height:30px;display:flex;align-items:center;justify-content:center;font-size:11px;border-radius:50%;cursor:pointer;transition:background .1s}
  .drp-day:hover{background:#e8eaf6}
  .drp-day.sel{background:var(--navy);color:#fff;font-weight:700}
  .drp-day.in-range{background:#E8F0FE;border-radius:0}
  .drp-day.today{font-weight:700;text-decoration:underline}
  .drp-foot{display:flex;justify-content:flex-end;gap:6px;border-top:1px solid #e8eaf6;padding-top:8px}
  .drp-foot button{border-radius:6px;padding:6px 14px;font-size:12px;font-weight:600}
  .drp-cancel{background:#f5f5f5;color:#555}
  .drp-apply{background:var(--navy);color:#fff}

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

  /* ÚLTIMA UNIDAD EN TODA LA EMPRESA */
  .empresa-panel{background:#fff;border-bottom:2px solid var(--red);flex-shrink:0}
  .empresa-header{display:flex;align-items:center;gap:8px;padding:7px 14px;cursor:pointer;background:#FFF0F0;user-select:none}
  .empresa-header:hover{background:#FFE8E8}
  .empresa-title{font-size:12px;font-weight:bold;color:var(--red);flex:1}
  .empresa-badge-count{font-size:11px;background:var(--red);color:#fff;border-radius:10px;padding:1px 8px;font-weight:bold}
  .empresa-body{display:none;overflow-x:auto;max-height:220px;overflow-y:auto}
  .empresa-body.open{display:block}
  .empresa-table{width:100%;border-collapse:collapse;font-size:11px}
  .empresa-table th{background:var(--red);color:#fff;padding:4px 8px;text-align:left;font-size:10px;position:sticky;top:0}
  .empresa-table td{padding:3px 8px;border-bottom:1px solid #FFE0E0;white-space:nowrap}
  .empresa-table td.art{white-space:normal;max-width:280px}
  .empresa-table tr:hover td{background:#FFF5F5}
  tr.empresa-row{background:#FFF5F5!important;border-left:3px solid var(--red)}
  tr.empresa-row:hover{background:#FFE8E8!important}
  .emp-badge{display:inline-block;background:var(--red);color:#fff;font-size:8px;font-weight:bold;padding:1px 4px;border-radius:3px;margin-left:5px;vertical-align:middle;white-space:nowrap}
</style>
</head>
<body>

<!-- LOGIN OVERLAY -->
<div class="login-overlay" id="loginOverlay">
  <div class="login-box">
    <div class="login-logo">
      <div style="font-size:40px">📦</div>
      <h2>Stock — Última Unidad</h2>
      <p>Oscar Barbieri S.A.</p>
    </div>
    <div class="login-field">
      <label>Usuario</label>
      <input type="text" id="loginUser" placeholder="Ingresá tu usuario" autocomplete="username" onkeydown="if(event.key==='Enter')document.getElementById('loginPass').focus()">
    </div>
    <div class="login-field">
      <label>Contraseña</label>
      <input type="password" id="loginPass" placeholder="Ingresá tu contraseña" autocomplete="current-password" onkeydown="if(event.key==='Enter')doLogin()">
    </div>
    <button class="login-btn" onclick="doLogin()">Ingresar</button>
    <div class="login-error" id="loginError"></div>
  </div>
</div>

<div class="header">
  <div>
    <h1>📦 Dashboard Stock — Última Unidad Sin Reposición</h1>
    <div class="meta">SOURCE_FILE &nbsp;|&nbsp; Actualizado: TODAY_STR</div>
  </div>
  <div style="display:flex;align-items:center;gap:10px;flex-shrink:0">
    <span class="session-info" id="sessionInfo"></span>
    <button class="btn btn-green" onclick="exportContextual()">⬇ Exportar Excel</button>
    <button class="logout-btn" onclick="doLogout()">Cerrar sesión</button>
  </div>
</div>

<div class="layout">

  <!-- SIDEBAR -->
  <div class="sidebar">

    <div class="filter-section">
      <div class="filter-label">Código / Descripción</div>
      <input class="search-box" id="searchInput" type="text" placeholder="Código o artículo…" oninput="applyFilters()">
    </div>

    <div class="filter-section" id="fechaSection" style="display:none">
      <div class="filter-label">Período</div>
      <div class="drp-wrap" id="drpWrap">
        <button class="drp-trigger" id="drpTrigger" onclick="drpToggle()">
          📅 <span id="drpLabel">Seleccionar fechas</span>
        </button>
        <div class="drp-panel" id="drpPanel" onclick="event.stopPropagation()">
          <div class="drp-right">
            <div class="drp-cals">
              <div id="drpCal0"></div>
            </div>
            <div class="drp-foot">
              <button class="drp-cancel" onclick="drpCancel()">Cancelar</button>
              <button class="drp-apply"  onclick="drpApply()">Aplicar</button>
            </div>
          </div>
        </div>
      </div>
      <!-- inputs ocultos que leen el resto del código -->
      <input type="hidden" id="filtroDesde">
      <input type="hidden" id="filtroHasta">
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

    <div class="filter-section" id="branchSection">
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
      <div class="kpi" id="kpiBranchesCard"><div class="kpi-val" id="kpiBranches">—</div><div class="kpi-lbl">Sucursales c/ alerta</div></div>
      <div class="kpi"><div class="kpi-val" id="kpiAlerts">—</div><div class="kpi-lbl">Registros última unidad</div></div>
      <div class="kpi" style="background:#FFF0F0;border:1px solid var(--red-bg)"><div class="kpi-val" id="kpiEmpresa" style="color:var(--red)">—</div><div class="kpi-lbl" style="color:var(--red)">Única unidad en empresa</div></div>
    </div>

    <div class="tab-bar" id="tabBar">
      <button class="tab-btn active" id="tabStock" onclick="setTab('stock')">📦 Stock Última Unidad</button>
      <button class="tab-btn" id="tabVentas" onclick="setTab('ventas')" style="display:none">📉 Ventas Detectadas</button>
    </div>

    <!-- PANEL ÚNICA UNIDAD EMPRESA (admin: panel colapsable; todos: marca en tabla) -->
    <div class="empresa-panel" id="empresaPanel" style="display:none">
      <div class="empresa-header" onclick="toggleEmpresaPanel()">
        <span style="font-size:15px">🔴</span>
        <span class="empresa-title" id="empresaTitle">Última unidad en toda la empresa</span>
        <span class="empresa-badge-count" id="empresaCount">0</span>
        <span style="font-size:11px;color:var(--red);margin-left:6px" id="empresaChevron">▼</span>
      </div>
      <div class="empresa-body" id="empresaBody">
        <table class="empresa-table">
          <thead id="empresaThead"><tr><th>Código</th><th>Artículo</th><th>Rubro</th><th>Marca</th><th>Sucursal</th></tr></thead>
          <tbody id="empresaTbody"></tbody>
        </table>
      </div>
    </div>

    <div class="toolbar" id="stockToolbar">
      <span class="results-count" id="resultsCount"></span>
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

    <!-- PANEL VENTAS (solo admin) -->
    <div id="ventasView" style="display:none;flex:1;overflow:hidden;display:none;flex-direction:column">
      <div class="ventas-wrap">
        <div class="ventas-summary" id="ventasSummary"></div>
        <div id="ventasTable"></div>
      </div>
    </div>
  </div>
</div>


<script>
const RAW = DATA_PLACEHOLDER;
const VENTAS_HISTORIAL = VENTAS_PLACEHOLDER;
const ALL_BRANCHES = RAW.branches;
const PRODUCTS = RAW.products;
const CREDS = CREDS_PLACEHOLDER;

let selectedBranches = new Set(ALL_BRANCHES);
let sortCol = null, sortDir = 1, currentView = 'table';
let currentUser = null, currentSucursal = null, currentFamilias = null;

// ── LOGIN ─────────────────────────────────────────────────────────────────────
function doLogin() {
  const user = document.getElementById('loginUser').value.trim().toLowerCase();
  const pass = document.getElementById('loginPass').value;
  const cred = CREDS[user];
  if (!cred || cred.password !== pass) {
    document.getElementById('loginError').textContent = 'Usuario o contraseña incorrectos.';
    document.getElementById('loginPass').value = '';
    return;
  }
  currentUser = user;
  currentSucursal = cred.sucursal;
  currentFamilias = cred.familias || null;
  document.getElementById('loginOverlay').classList.add('hidden');
  document.getElementById('loginError').textContent = '';
  const label = currentSucursal
    ? currentSucursal
    : (cred.titulo || 'Administrador (todas las sucursales)');
  document.getElementById('sessionInfo').textContent = label;
  if (cred.titulo) document.title = cred.titulo;
  // Si es sucursal específica: filtrar, ocultar selector y card de sucursales
  if (currentSucursal) {
    selectedBranches = new Set([currentSucursal]);
    document.getElementById('branchSection').style.display = 'none';
    document.getElementById('kpiBranchesCard').style.display = 'none';
  } else {
    // Admin: mostrar tab de ventas si hay historial
    if (VENTAS_HISTORIAL.length > 0) {
      document.getElementById('tabVentas').style.display = '';
    }
  }
  refreshAcValues();
  initFilters();
  applyFilters();
}

function doLogout() {
  currentUser = null; currentSucursal = null; currentFamilias = null;
  document.title = 'Dashboard Stock — Última Unidad Sin Reposición';
  selectedBranches = new Set(ALL_BRANCHES);
  document.getElementById('loginUser').value = '';
  document.getElementById('loginPass').value = '';
  document.getElementById('sessionInfo').textContent = '';
  document.getElementById('branchSection').style.display = '';
  document.getElementById('kpiBranchesCard').style.display = '';
  document.getElementById('tabVentas').style.display = 'none';
  setTab('stock');
  document.getElementById('loginOverlay').classList.remove('hidden');
  refreshAcValues();
  resetFilters();
}

// ── TABS ──────────────────────────────────────────────────────────────────────
let currentTab = 'stock';
function setTab(tab) {
  currentTab = tab;
  document.getElementById('tabStock').classList.toggle('active', tab==='stock');
  document.getElementById('tabVentas').classList.toggle('active', tab==='ventas');
  document.getElementById('stockToolbar').style.display  = tab==='stock'  ? '' : 'none';
  document.getElementById('tableView').style.display     = tab==='stock'  ? '' : 'none';
  document.getElementById('summaryView').style.display   = 'none';
  document.getElementById('ventasView').style.display    = tab==='ventas' ? 'flex' : 'none';
  document.getElementById('kpiProds').closest('.kpi-bar').style.display = tab==='stock' ? '' : 'none';
  document.getElementById('fechaSection').style.display  = tab==='ventas' ? '' : 'none';
  if (tab==='ventas') { drpInit(); renderVentas(); }
}

// ── VENTAS ────────────────────────────────────────────────────────────────────

function getVentasFiltradas() {
  // Lee filtros del sidebar unificado
  const texto = (document.getElementById('searchInput').value || '').toLowerCase().trim();
  const fam   = acExact.familia ? document.getElementById('familiaInput').value.trim() : '';
  const rubro = acExact.rubro   ? document.getElementById('rubroInput').value.trim()   : '';
  const marca = acExact.marca   ? document.getElementById('marcaInput').value.trim()   : '';
  const desde = document.getElementById('filtroDesde').value;
  const hasta = document.getElementById('filtroHasta').value;
  // Sucursal: si hay una sola seleccionada la usamos como filtro
  const sucFiltro = (currentSucursal && !currentSucursal.includes('todas')) ? currentSucursal : '';
  const sucSet = new Set([...selectedBranches]);

  return VENTAS_HISTORIAL.filter(v => {
    if (currentFamilias && !currentFamilias.includes((v.familia||'').toUpperCase())) return false;
    if (sucFiltro && v.sucursal !== sucFiltro) return false;
    if (!sucFiltro && sucSet.size < ALL_BRANCHES.length && !sucSet.has(v.sucursal)) return false;
    if (fam   && (v.familia||'') !== fam)  return false;
    if (rubro && (v.rubro||'')   !== rubro) return false;
    if (marca && (v.marca||'')   !== marca) return false;
    if (texto && !v.articulo.toLowerCase().includes(texto) && !v.codigo.includes(texto)) return false;
    if (desde || hasta) {
      const [d,m,a] = v.fecha.split('/');
      const iso = `${a}-${m}-${d}`;
      if (desde && iso < desde) return false;
      if (hasta && iso > hasta) return false;
    }
    return true;
  });
}

function renderVentas() {
  const ventas = getVentasFiltradas();
  const sucSet = new Set(ventas.map(v=>v.sucursal));
  buildBranchList();
  updateEmpresaPanelVentas(ventas);
  const totalUnidades = ventas.reduce((s,v) => s + (v.unidades||1), 0);
  document.getElementById('ventasSummary').innerHTML =
    `<div class="ventas-kpi"><b>${ventas.length.toLocaleString('es-AR')}</b>Registros</div>
     <div class="ventas-kpi"><b>${sucSet.size}</b>Sucursales</div>
     <div class="ventas-kpi"><b>${new Set(ventas.map(v=>v.codigo)).size}</b>Productos distintos</div>
     <div class="ventas-kpi"><b>${totalUnidades.toLocaleString('es-AR')}</b>Unidades vendidas</div>`;
  if (!ventas.length) {
    document.getElementById('ventasTable').innerHTML =
      '<div class="no-results"><div class="icon">📭</div><div>Sin ventas detectadas para los filtros aplicados.</div></div>';
    return;
  }
  const cols = ['Fecha','Sucursal','Código','Artículo','Rubro','Familia','Marca','Uds.',''];
  document.getElementById('ventasTable').innerHTML = `
    <table style="border-collapse:collapse;width:100%;font-size:11px">
      <thead><tr>${cols.map(c=>`<th style="background:var(--navy);color:#fff;padding:6px 8px;text-align:left;white-space:nowrap">${c}</th>`).join('')}</tr></thead>
      <tbody>${ventas.map((v,i) => {
        const emp = v.empresa_last;
        const bg = emp ? 'background:#FFF5F5' : (i%2 ? 'background:#F2F7FF' : '');
        const bd = emp ? 'border-left:3px solid var(--red)' : '';
        const td = `padding:4px 8px;border-bottom:1px solid var(--border)`;
        return `<tr style="${bg};${bd}">
          <td style="${td}">${v.fecha}</td>
          <td style="${td};font-weight:bold">${v.sucursal}</td>
          <td style="${td}">${v.codigo}</td>
          <td style="${td};max-width:260px">${v.articulo}</td>
          <td style="${td}"><span class="rubro-tag">${v.rubro}</span></td>
          <td style="${td}">${v.familia}</td>
          <td style="${td}">${v.marca}</td>
          <td style="${td};text-align:center;font-weight:bold">${v.unidades||1}</td>
          <td style="${td};text-align:center">${emp ? '<span class="emp-badge">EMPRESA</span>' : ''}</td>
        </tr>`;
      }).join('')}</tbody>
    </table>`;
}


function exportVentas() {
  const ventas = getVentasFiltradas();
  if (!ventas.length) { alert('Sin datos para exportar.'); return; }
  const wb = XLSX.utils.book_new();
  const rows = [['Fecha','Sucursal','Código','Artículo','Rubro','Familia','Marca','Uds.','Única unidad empresa']];
  ventas.forEach(v => rows.push([v.fecha,v.sucursal,v.codigo,v.articulo,v.rubro,v.familia,v.marca,v.unidades||1,v.empresa_last?'SÍ':'']));
  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [{wch:12},{wch:16},{wch:10},{wch:50},{wch:18},{wch:16},{wch:14},{wch:7},{wch:20}];
  XLSX.utils.book_append_sheet(wb, ws, 'Ventas Última Unidad');
  XLSX.writeFile(wb, `Ventas_Ultima_Unidad_${new Date().toISOString().slice(0,10)}.xlsx`);
}

// ── INIT ─────────────────────────────────────────────────────────────────────
function initFilters() { buildBranchList(); }

function buildBranchList(filteredProds) {
  const el = document.getElementById('branchList');
  el.innerHTML = '';
  const counts = {};
  if (currentTab === 'ventas') {
    // Contar ventas del período filtrado por sucursal
    const ventas = getVentasFiltradas();
    ventas.forEach(v => { counts[v.sucursal] = (counts[v.sucursal]||0) + 1; });
  } else {
    const prods = filteredProds || PRODUCTS;
    prods.forEach(p => ALL_BRANCHES.forEach(b => {
      if ((p.branch_stocks[b]||0) === 1) counts[b] = (counts[b]||0) + 1;
    }));
  }
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
  rubro:   { inputId:'rubroInput',   dropId:'rubroDropdown',   values:[] },
  familia: { inputId:'familiaInput', dropId:'familiaDropdown', values:[] },
  marca:   { inputId:'marcaInput',   dropId:'marcaDropdown',   values:[] },
};

// Recalcula las opciones disponibles según los productos con stock en la sucursal activa
function refreshAcValues() {
  let pool = currentSucursal
    ? PRODUCTS.filter(p => (p.branch_stocks[currentSucursal]||0) > 0)
    : PRODUCTS;
  if (currentFamilias) pool = pool.filter(p => currentFamilias.includes((p.familia||'').toUpperCase()));
  AC_FIELDS.familia.values = [...new Set(pool.map(p=>p.familia).filter(Boolean))].sort();
  AC_FIELDS.marca.values   = [...new Set(pool.map(p=>p.marca).filter(Boolean))].sort();
  AC_FIELDS.rubro.values   = [...new Set(pool.map(p=>p.rubro).filter(Boolean))].sort();
  // Reconstruir mapa familia→rubros
  Object.keys(FAMILIA_RUBROS).forEach(k => delete FAMILIA_RUBROS[k]);
  pool.forEach(p => {
    if (!p.familia || !p.rubro) return;
    if (!FAMILIA_RUBROS[p.familia]) FAMILIA_RUBROS[p.familia] = new Set();
    FAMILIA_RUBROS[p.familia].add(p.rubro);
  });
}
let acActiveIdx = {};

function acHighlight(q, text) {
  if (!q) return text;
  const i = text.toLowerCase().indexOf(q.toLowerCase());
  if (i < 0) return text;
  return text.slice(0,i) + '<em>' + text.slice(i, i+q.length) + '</em>' + text.slice(i+q.length);
}

// Devuelve el pool de productos activo según familia+rubro seleccionados
function acBasePool() {
  let base = currentSucursal
    ? PRODUCTS.filter(p => (p.branch_stocks[currentSucursal]||0) > 0)
    : PRODUCTS;
  if (currentFamilias) base = base.filter(p => currentFamilias.includes((p.familia||'').toUpperCase()));
  const familiaVal = document.getElementById('familiaInput').value.trim();
  const rubroVal   = document.getElementById('rubroInput').value.trim();
  return base.filter(p =>
    (!familiaVal || (acExact.familia ? p.familia === familiaVal : p.familia.toLowerCase().includes(familiaVal.toLowerCase()))) &&
    (!rubroVal   || (acExact.rubro   ? p.rubro   === rubroVal   : p.rubro.toLowerCase().includes(rubroVal.toLowerCase())))
  );
}

function acGetPoolFor(field) {
  if (field === 'familia') {
    let base = currentSucursal ? PRODUCTS.filter(p => (p.branch_stocks[currentSucursal]||0) > 0) : PRODUCTS;
    if (currentFamilias) base = base.filter(p => currentFamilias.includes((p.familia||'').toUpperCase()));
    return [...new Set(base.map(p=>p.familia).filter(Boolean))].sort();
  }
  if (field === 'rubro') {
    const familiaVal = document.getElementById('familiaInput').value.trim();
    let base = currentSucursal ? PRODUCTS.filter(p => (p.branch_stocks[currentSucursal]||0) > 0) : PRODUCTS;
    if (currentFamilias) base = base.filter(p => currentFamilias.includes((p.familia||'').toUpperCase()));
    const pool = familiaVal ? base.filter(p => acExact.familia ? p.familia === familiaVal : p.familia.toLowerCase().includes(familiaVal.toLowerCase())) : base;
    return [...new Set(pool.map(p=>p.rubro).filter(Boolean))].sort();
  }
  if (field === 'marca') {
    return [...new Set(acBasePool().map(p=>p.marca).filter(Boolean))].sort();
  }
  return AC_FIELDS[field].values;
}

function acUpdate(field) {
  const f = AC_FIELDS[field];
  const input = document.getElementById(f.inputId);
  const drop  = document.getElementById(f.dropId);
  const q = input.value.trim();
  const pool = acGetPoolFor(field);
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
  // Cascada: al cambiar familia limpiar rubro y marca; al cambiar rubro limpiar marca
  if (field === 'familia') {
    document.getElementById('rubroInput').value = '';
    document.getElementById('marcaInput').value = '';
    acExact.rubro = false; acExact.marca = false;
  }
  if (field === 'rubro') {
    document.getElementById('marcaInput').value = '';
    acExact.marca = false;
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
    if (currentFamilias && !currentFamilias.includes(p.familia.toUpperCase())) return false;
    if (search  && !p.codigo.toLowerCase().includes(search)  && !p.articulo.toLowerCase().includes(search))  return false;
    if (rubro)   { const rv = p.rubro.toLowerCase();   if (acExact.rubro   ? rv !== rubro   : !rv.includes(rubro))   return false; }
    if (familia) { const fv = p.familia.toLowerCase(); if (acExact.familia ? fv !== familia : !fv.includes(familia)) return false; }
    if (marca)   { const mv = p.marca.toLowerCase();   if (acExact.marca   ? mv !== marca   : !mv.includes(marca))   return false; }
    if (currentSucursal) {
      if ((p.branch_stocks[currentSucursal] || 0) === 0) return false;
    } else {
      if (![...selectedBranches].some(b => (p.branch_stocks[b] || 0) > 0)) return false;
    }
    return true;
  });
}

// ── ÚLTIMA UNIDAD EMPRESA ─────────────────────────────────────────────────────
function isSingleCompany(p) {
  return Object.values(p.branch_stocks).reduce((a,b)=>a+b,0) === 1;
}

function toggleEmpresaPanel() {
  const body = document.getElementById('empresaBody');
  const chev = document.getElementById('empresaChevron');
  const open = body.classList.toggle('open');
  chev.textContent = open ? '▲' : '▼';
}

function updateEmpresaPanel(prods) {
  // Modo STOCK: productos con única unidad en empresa
  const checkProds = currentSucursal
    ? prods.filter(p => (p.branch_stocks[currentSucursal]||0) > 0)
    : prods;
  const singles = checkProds.filter(isSingleCompany);

  const panel = document.getElementById('empresaPanel');
  document.getElementById('empresaTitle').textContent = 'Última unidad en toda la empresa — stock actual';
  document.getElementById('empresaCount').textContent = singles.length;

  if (!singles.length) { panel.style.display = 'none'; return; }
  panel.style.display = '';

  if (!currentSucursal) {
    document.getElementById('empresaThead').innerHTML =
      '<tr><th>Código</th><th>Artículo</th><th>Rubro</th><th>Marca</th><th>Sucursal</th></tr>';
    document.getElementById('empresaTbody').innerHTML = singles.map(p => {
      const suc = Object.entries(p.branch_stocks).find(([,v])=>v===1)?.[0] || '';
      return `<tr><td>${p.codigo}</td><td class="art">${p.articulo}</td><td>${p.rubro}</td><td>${p.marca}</td><td><b>${suc}</b></td></tr>`;
    }).join('');
  }
}

function updateEmpresaPanelVentas(ventas) {
  // Modo VENTAS: ventas donde era la única unidad en empresa
  const singles = ventas.filter(v => v.empresa_last);
  const panel = document.getElementById('empresaPanel');
  document.getElementById('empresaTitle').textContent = 'Ventas de última unidad en toda la empresa';
  document.getElementById('empresaCount').textContent = singles.length;

  if (!singles.length) { panel.style.display = 'none'; return; }
  panel.style.display = '';

  document.getElementById('empresaThead').innerHTML =
    '<tr><th>Fecha</th><th>Sucursal</th><th>Código</th><th>Artículo</th><th>Rubro</th><th>Marca</th></tr>';
  document.getElementById('empresaTbody').innerHTML = singles.map(v =>
    `<tr><td>${v.fecha}</td><td><b>${v.sucursal}</b></td><td>${v.codigo}</td><td class="art">${v.articulo}</td><td>${v.rubro}</td><td>${v.marca}</td></tr>`
  ).join('');
}

function applyFilters() {
  if (currentTab === 'ventas') { renderVentas(); return; }
  const prods = getFilteredProducts();
  updateKPIs(prods);
  buildBranchList(prods);
  updateEmpresaPanel(prods);
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
  document.getElementById('kpiEmpresa').textContent  = prods.filter(isSingleCompany).length.toLocaleString('es-AR');
  document.getElementById('resultsCount').textContent = prods.length.toLocaleString('es-AR') + ' producto(s) encontrado(s)';
}

// ── TABLE ─────────────────────────────────────────────────────────────────────
function stkCell(v) {
  if (v === 0) return `<span class="stk-cell stk-0">-</span>`;
  const cls = v===1?'stk-1':v<5?'stk-few':'stk-ok';
  return `<span class="stk-cell ${cls}">${v}</span>`;
}
function getColVal(p, i, branches) {
  switch(i){case 0:return p.codigo;case 1:return p.articulo;
    case 2:return p.marca;default:return p.branch_stocks[branches[i-3]]||0;}
}
function sortBy(i){ sortCol===i?sortDir*=-1:(sortCol=i,sortDir=1); applyFilters(); }

function renderTable(prods) {
  // Solo mostrar columnas de sucursales que tienen al menos un producto con stock
  const active = [...selectedBranches].filter(b => prods.some(p => (p.branch_stocks[b]||0) > 0));

  const noRes = document.getElementById('noResults');
  const tbl   = document.getElementById('dataTable');
  if (!prods.length) {
    document.getElementById('tableBody').innerHTML = '';
    document.getElementById('tableHead').innerHTML = '';
    noRes.style.display = 'block'; tbl.style.display = 'none'; return;
  }
  noRes.style.display = 'none'; tbl.style.display = '';

  document.getElementById('tableHead').innerHTML = '<tr>'+['Código','Artículo','Marca',...active].map((c,i)=>
    `<th onclick="sortBy(${i})" class="${sortCol===i?(sortDir>0?'sort-asc':'sort-desc'):''}">${c}</th>`
  ).join('')+'</tr>';

  let sorted = [...prods];
  if (sortCol !== null) sorted.sort((a,b) => {
    const av = getColVal(a,sortCol,active), bv = getColVal(b,sortCol,active);
    return (av>bv?1:av<bv?-1:0)*sortDir;
  });

  document.getElementById('tableBody').innerHTML = sorted.map((p,ri) => {
    const emp = isSingleCompany(p);
    const rowCls = emp ? 'empresa-row' : '';
    const rowStyle = emp ? '' : (ri%2 ? 'background:#F2F7FF' : '');
    const badge = emp ? '<span class="emp-badge">EMPRESA</span>' : '';
    return `<tr class="${rowCls}" style="${rowStyle}">
      <td>${p.codigo}</td>
      <td class="art">${p.articulo}${badge}</td>
      <td style="text-align:center">${p.marca}</td>
      ${active.map(b=>`<td style="padding:2px 4px;text-align:center">${stkCell(p.branch_stocks[b]||0)}</td>`).join('')}
    </tr>`;
  }).join('');
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
function exportContextual() {
  if (currentTab === 'ventas') exportVentas();
  else exportStock();
}

function exportStock() {
  const prods = getFilteredProducts();
  const branches = [...selectedBranches].filter(b => prods.some(p => (p.branch_stocks[b]||0) === 1));
  if (!branches.length) { alert('Sin datos para exportar en las sucursales seleccionadas.'); return; }

  const wb = XLSX.utils.book_new();
  const rows = [['Sucursal', 'Cant. Registros', 'Código', 'Descripción', 'Rubro', 'Marca']];
  branches.forEach(branch => {
    const ps = prods.filter(p => (p.branch_stocks[branch]||0) === 1);
    if (!ps.length) return;
    ps.forEach(p => rows.push([branch, ps.length, p.codigo, p.articulo, p.rubro, p.marca]));
  });

  const ws = XLSX.utils.aoa_to_sheet(rows);
  ws['!cols'] = [{wch:16},{wch:16},{wch:10},{wch:54},{wch:18},{wch:14}];
  XLSX.utils.book_append_sheet(wb, ws, 'Última Unidad por Sucursal');

  const label = currentSucursal ? currentSucursal.replace(/\s+/g,'_') : 'Todas';
  XLSX.writeFile(wb, `Stock_Ultima_Unidad_${label}_${new Date().toISOString().slice(0,10)}.xlsx`);
}

// ── DATE RANGE PICKER ─────────────────────────────────────────────────────────
const DRP_MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
const DRP_DIAS  = ['Lu','Ma','Mi','Ju','Vi','Sa','Do'];
const TODAY_ISO = new Date().toISOString().slice(0,10);
let drpFrom='', drpTo='', drpSelecting=false, drpViewYear, drpViewMonth;

function drpInit() {
  const d = new Date();
  drpViewYear = d.getFullYear(); drpViewMonth = d.getMonth();
  drpRender();
}
function drpToggle() {
  const p = document.getElementById('drpPanel');
  const open = p.style.display === 'flex';
  if (open) { p.style.display = 'none'; return; }
  const btn = document.getElementById('drpTrigger');
  const rect = btn.getBoundingClientRect();
  p.style.top  = (rect.bottom + 6) + 'px';
  p.style.left = rect.left + 'px';
  p.style.display = 'flex';
  drpFrom = document.getElementById('filtroDesde').value || '';
  drpTo   = document.getElementById('filtroHasta').value || '';
  drpSelecting = false;
  drpRender();
}
function drpClose() { document.getElementById('drpPanel').style.display = 'none'; }
function drpApply() {
  document.getElementById('filtroDesde').value = drpFrom;
  document.getElementById('filtroHasta').value = drpTo;
  const label = drpFrom && drpTo
    ? (drpFrom === drpTo ? drpFmt(drpFrom) : drpFmt(drpFrom) + ' → ' + drpFmt(drpTo))
    : 'Seleccionar fechas';
  document.getElementById('drpLabel').textContent = label;
  drpClose();
  applyFilters();
}
function drpCancel() {
  drpFrom = document.getElementById('filtroDesde').value || '';
  drpTo   = document.getElementById('filtroHasta').value || '';
  drpClose();
}
function drpFmt(iso) {
  if (!iso) return '';
  const [y,m,d] = iso.split('-');
  return `${d}/${m}/${y}`;
}
function drpPreset(p) {
  const hoy = new Date(TODAY_ISO + 'T12:00:00');
  const iso = d => d.toISOString().slice(0,10);
  const addD = (d,n) => { const r=new Date(d); r.setDate(r.getDate()+n); return r; };
  let f, t = TODAY_ISO;
  if      (p==='hoy')     { f = TODAY_ISO; }
  else if (p==='ayer')    { f = t = iso(addD(hoy,-1)); }
  else if (p==='7d')      { f = iso(addD(hoy,-6)); }
  else if (p==='15d')     { f = iso(addD(hoy,-14)); }
  else if (p==='30d')     { f = iso(addD(hoy,-29)); }
  else if (p==='mes_act') { f = `${TODAY_ISO.slice(0,7)}-01`; }
  else if (p==='mes_ant') { const d=new Date(hoy.getFullYear(),hoy.getMonth()-1,1);
                            f=iso(d); t=iso(new Date(hoy.getFullYear(),hoy.getMonth(),0)); }
  else if (p==='todo')    { f=''; t=''; }
  else return;
  drpFrom=f; drpTo=t; drpSelecting=false;
  document.querySelectorAll('.drp-pre').forEach(b=>b.classList.toggle('active',b.dataset.p===p));
  drpApply();
}
function drpRender() {
  document.getElementById('drpCal0').innerHTML = drpBuildMonth(drpViewYear, drpViewMonth);
}
function drpBuildMonth(y, mo) {
  const first = new Date(y, mo, 1);
  let dow = first.getDay(); dow = dow===0 ? 6 : dow-1;
  const days = new Date(y, mo+1, 0).getDate();
  let h = `<div class="drp-mhdr">
    <button style="background:none;border:none;cursor:pointer;font-size:.9rem;color:var(--navy)" onclick="drpNav(-1)">◀</button>
    <span>${DRP_MESES[mo]} ${y}</span>
    <button style="background:none;border:none;cursor:pointer;font-size:.9rem;color:var(--navy)" onclick="drpNav(1)">▶</button>
  </div><div class="drp-grid">`;
  DRP_DIAS.forEach(d => h += `<div class="drp-dh">${d}</div>`);
  for (let i=0; i<dow; i++) h += `<div></div>`;
  for (let d=1; d<=days; d++) {
    const iso = `${y}-${String(mo+1).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
    let cls = 'drp-day';
    if (iso === TODAY_ISO) cls += ' today';
    if (iso===drpFrom||iso===drpTo) cls += ' sel';
    else if (drpFrom && drpTo && iso>drpFrom && iso<drpTo) cls += ' in-range';
    h += `<div class="${cls}" onclick="drpClick('${iso}')">${d}</div>`;
  }
  return h + '</div>';
}
function drpNav(dir) {
  drpViewMonth += dir;
  if (drpViewMonth > 11) { drpViewMonth=0; drpViewYear++; }
  if (drpViewMonth < 0)  { drpViewMonth=11; drpViewYear--; }
  drpRender();
}
function drpClick(iso) {
  if (!drpSelecting || iso < drpFrom) {
    drpFrom=iso; drpTo=iso; drpSelecting=true;
    document.querySelectorAll('.drp-pre').forEach(b=>b.classList.toggle('active',b.dataset.p==='custom'));
    drpRender();
  } else { drpTo=iso; drpSelecting=false; drpRender(); }
}
// Cerrar al hacer click fuera
document.addEventListener('click', e => {
  const w = document.getElementById('drpWrap');
  if (w && !w.contains(e.target)) drpClose();
});

// ── RESET ─────────────────────────────────────────────────────────────────────
function resetFilters() {
  ['searchInput','rubroInput','familiaInput','marcaInput','filtroDesde','filtroHasta']
    .forEach(id => document.getElementById(id).value = '');
  document.getElementById('drpLabel').textContent = 'Seleccionar fechas';
  ['rubro','familia','marca'].forEach(acClose);
  selectedBranches = new Set(ALL_BRANCHES);
  sortCol = null; sortDir = 1;
  buildBranchList(); applyFilters();
}

// No iniciar hasta login
document.getElementById('loginUser').focus();
</script>
</body>
</html>"""

# Leer historial de ventas si existe
ventas_path = BASE / 'ventas_historial.json'
ventas_json = '[]'
if ventas_path.exists():
    with open(ventas_path, encoding='utf-8') as f:
        ventas_json = f.read()

HTML = (HTML
    .replace('TODAY_STR', today_str)
    .replace('SOURCE_FILE', source_file)
    .replace('DATA_PLACEHOLDER', data_json)
    .replace('CREDS_PLACEHOLDER', creds_json)
    .replace('VENTAS_PLACEHOLDER', ventas_json))

out_path = BASE / 'Dashboard.html'
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(HTML)

print(f"Dashboard generado: {out_path.name}")
print(f"  Productos: {n_products:,}  |  Rubros: {n_rubros}")
