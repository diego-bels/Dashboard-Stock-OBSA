"""
Paso 1: Procesa todos los archivos de stock y venta diaria en orden cronológico.
  - Carpeta Stock/        → archivos "Stock Total DD MM.xlsx"
  - Carpeta Venta Diaria/ → archivos "Venta Diaria del DD MM YYYY.xlsx"
  - procesados.json       → registro de pares ya procesados (evita duplicados)
"""
import pandas as pd, json, re, sys
from pathlib import Path
from datetime import datetime, date, timedelta

BASE        = Path(__file__).parent
DIR_STOCK   = BASE / 'Stock'
DIR_VENTAS  = BASE / 'Venta Diaria'
PROC_PATH   = BASE / 'procesados.json'
HIST_PATH   = BASE / 'ventas_historial.json'
DATA_PATH   = BASE / 'data.json'

# ── CONFIGURACIÓN ─────────────────────────────────────────────────────────────
DEPOSITOS = {
    'D01': '25 de Mayo',   'D02': 'Tafi del Valle', 'D04': 'Banda',
    'D05': 'Maipú',        'D06': 'Concepción',     'D07': 'Lules',
    'D09': 'Outlet',       'D10': 'Rosario',        'D11': 'Cafayate',
    'D12': 'COL',          'D13': 'Yerba Buena',    'D15': 'Monteros',
    'D16': 'Tafi Viejo',   'D17': 'Famailla',       'D18': 'Santa María',
    'D23': 'Big',          'D24': 'Aguilares',      'D34': 'Alberdi',
}
COL_KEY    = 'D12'
SUCURSALES = [v for k, v in DEPOSITOS.items() if k != COL_KEY]

FAMILIAS_EXCLUIDAS = {
    'INSUMOS', 'FLETES - PROMO REG.',
}
RUBROS_EXCLUIDOS = {
    'BOLSAS BARBIERI', 'FLETE - CONCEPTOS VARIOS', 'GARANTIAS',
    'INSUMOS INTERNOS', 'PROMO REGALO', 'REGALOS EMPRESARIALES',
}

# ── HELPERS ───────────────────────────────────────────────────────────────────
def clean_text(s):
    if not isinstance(s, str): return str(s)
    return s.replace('_x000D_', '').replace('\n', ' ').strip()

def safe_int(v):
    try:    return int(float(v)) if pd.notna(v) else 0
    except: return 0

def safe_str(v):
    if pd.isna(v): return ''
    return str(v).strip()

def fecha_desde_nombre(nombre):
    """'Venta Diaria del 16 07 2026.xlsx' → date(2026,7,16)
       'Stock Total 17 07.xlsx'           → date(2026,7,17)"""
    m = re.search(r'\b(\d{1,2})\s+(\d{1,2})\s+(\d{4})\b', nombre)
    if m:
        try: return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except: pass
    nums = re.findall(r'\b(\d{1,2})\b', nombre)
    if len(nums) >= 2:
        try: return date(datetime.now().year, int(nums[-1]), int(nums[-2]))
        except: pass
    return None

def fecha_str(d):
    return d.strftime('%d/%m/%Y')

def leer_stock(path):
    """Lee un Excel de stock y devuelve (df, desc_col, cod_col, costo_col)."""
    df = pd.read_excel(path, header=0)
    # Normalizar columna descripción
    rename = {}
    for col in df.columns:
        if 'escripci' in col.encode('utf-8','replace').decode('utf-8').lower():
            rename[col] = 'Descripción'
    if rename: df = df.rename(columns=rename)
    dep_cols = [d for d in DEPOSITOS if d in df.columns]
    if not dep_cols:
        print(f"  ERROR: sin columnas D en {path.name}")
        return None, None, None, None
    desc_col = next((c for c in df.columns if any(x in c.lower() for x in ['escripci','mercader','articul'])), df.columns[12])
    cod_col  = 'Codigo' if 'Codigo' in df.columns else df.columns[10]
    costo_col = next((c for c in df.columns if 'costo' in c.lower() and 'repo' in c.lower()), None)
    df = df[df[cod_col].notna()].copy()
    return df, desc_col, cod_col, costo_col

def construir_snapshot(df, desc_col, cod_col):
    """Construye dict codigo→{col, stocks, articulo, rubro, familia, marca}."""
    snap = {}
    for _, row in df.iterrows():
        try:    codigo = str(int(float(row[cod_col]))).zfill(6)
        except: codigo = clean_text(str(row[cod_col]))
        snap[codigo] = {
            'col':      safe_int(row.get(COL_KEY, 0)),
            'stocks':   {DEPOSITOS[d]: safe_int(row.get(d, 0)) for d in DEPOSITOS if d != COL_KEY},
            'articulo': clean_text(row.get(desc_col, '')),
            'rubro':    clean_text(safe_str(row.get('Rubro', ''))),
            'familia':  clean_text(safe_str(row.get('Familia', ''))),
            'marca':    clean_text(safe_str(row.get('Marca', ''))),
        }
    return snap

def leer_venta_diaria(path):
    """Devuelve dict codigo→unidades para los productos vendidos."""
    sold = {}
    try:
        dv = pd.read_excel(path)
        for _, row in dv.iterrows():
            try:    cod = str(int(float(row['Cod.']))).zfill(6)
            except: continue
            u = safe_int(row.get('Unid.', 0))
            if u > 0:
                sold[cod] = u
    except Exception as e:
        print(f"  AVISO: no se pudo leer {path.name}: {e}")
    return sold

def detectar_ventas(snap_prev, snap_act, sold_map, fec, hora):
    """Compara dos snapshots y devuelve lista de ventas detectadas."""
    ventas = []
    for codigo, prev in snap_prev.items():
        stocks_prev = prev.get('stocks', {})
        col_prev    = prev.get('col', 0)
        act         = snap_act.get(codigo, {})
        stocks_act  = act.get('stocks', {})
        col_act     = act.get('col', -1)

        for suc in SUCURSALES:
            if not (stocks_prev.get(suc,0)==1 and col_prev==0
                    and stocks_act.get(suc,0)==0 and col_act==0):
                continue
            if sold_map and codigo not in sold_map:
                continue
            total_prev = sum(stocks_prev.get(s,0) for s in SUCURSALES)
            ventas.append({
                'fecha':        fec,
                'hora':         hora,
                'sucursal':     suc,
                'codigo':       codigo,
                'articulo':     prev.get('articulo',''),
                'rubro':        prev.get('rubro',''),
                'familia':      prev.get('familia',''),
                'marca':        prev.get('marca',''),
                'empresa_last': total_prev == 1,
                'unidades':     sold_map.get(codigo, 1) if sold_map else 1,
                'confirmada':   bool(sold_map),
            })
    return ventas

# ── CREAR CARPETAS SI NO EXISTEN ──────────────────────────────────────────────
DIR_STOCK.mkdir(exist_ok=True)
DIR_VENTAS.mkdir(exist_ok=True)

# ── LISTAR Y ORDENAR ARCHIVOS DE STOCK ────────────────────────────────────────
stock_files = []
for pat in ['Stock Total*.xlsx', 'stock total*.xlsx']:
    stock_files += list(DIR_STOCK.glob(pat))
stock_files = [f for f in stock_files if 'Ultima Unidad' not in f.name]

dated = []
for f in stock_files:
    d = fecha_desde_nombre(f.name)
    if d:
        dated.append((d, f))
dated.sort(key=lambda x: x[0])

if len(dated) < 1:
    print("ERROR: No hay archivos en la carpeta Stock/")
    print("Copiá los archivos 'Stock Total DD MM.xlsx' en la carpeta Stock/")
    sys.exit(1)

print(f"  Archivos de stock encontrados: {len(dated)}")
for d, f in dated:
    print(f"    {fecha_str(d)}  →  {f.name}")

# ── LISTAR VENTAS DIARIAS ─────────────────────────────────────────────────────
vd_files = {}
for pat in ['Venta Diaria*.xlsx', 'Venta diaria*.xlsx']:
    for f in DIR_VENTAS.glob(pat):
        d = fecha_desde_nombre(f.name)
        if d:
            vd_files[d] = f

print(f"  Ventas diarias encontradas: {len(vd_files)}")

# ── CARGAR REGISTRO DE PARES PROCESADOS ───────────────────────────────────────
procesados = set()
if PROC_PATH.exists():
    with open(PROC_PATH, encoding='utf-8') as f:
        procesados = set(tuple(x) for x in json.load(f))

# ── PROCESAR PARES CONSECUTIVOS ───────────────────────────────────────────────
historial = []
if HIST_PATH.exists():
    with open(HIST_PATH, encoding='utf-8') as f:
        historial = json.load(f)
existentes = {(v['fecha'], v['sucursal'], v['codigo']) for v in historial}

pares_nuevos = 0
for i in range(len(dated) - 1):
    d_prev, f_prev = dated[i]
    d_act,  f_act  = dated[i+1]
    par_key = (f_prev.name, f_act.name)

    if par_key in procesados:
        continue

    print(f"\n  Procesando: {f_prev.name} → {f_act.name}")

    # Leer ambos stocks
    df_prev, desc_prev, cod_prev, _ = leer_stock(f_prev)
    df_act,  desc_act,  cod_act,  _ = leer_stock(f_act)
    if df_prev is None or df_act is None:
        continue

    snap_prev = construir_snapshot(df_prev, desc_prev, cod_prev)
    snap_act  = construir_snapshot(df_act,  desc_act,  cod_act)

    # Buscar venta diaria: fecha = día anterior al stock actual
    # (el stock del 18/07 refleja ventas del 17/07)
    fec_venta  = d_act - timedelta(days=1)
    vd_file    = vd_files.get(fec_venta)

    # Si no hay venta diaria exacta, buscar dentro del rango (ej: viernes→lunes)
    if not vd_file:
        for delta in range(1, (d_act - d_prev).days + 1):
            candidate = d_act - timedelta(days=delta)
            if candidate in vd_files:
                vd_file   = vd_files[candidate]
                fec_venta = candidate
                break

    sold_map = {}
    if vd_file:
        print(f"    Venta Diaria: {vd_file.name}")
        sold_map = leer_venta_diaria(vd_file)
        print(f"    Productos vendidos: {len(sold_map):,}")
    else:
        print(f"    Sin Venta Diaria para {fecha_str(fec_venta)} — detección solo por snapshot")

    hora = datetime.now().strftime('%H:%M')
    ventas = detectar_ventas(snap_prev, snap_act, sold_map, fecha_str(fec_venta), hora)

    nuevas = [v for v in ventas if (v['fecha'], v['sucursal'], v['codigo']) not in existentes]
    dup    = len(ventas) - len(nuevas)
    if dup:
        print(f"    Duplicados ignorados: {dup}")

    historial.extend(nuevas)
    existentes.update((v['fecha'], v['sucursal'], v['codigo']) for v in nuevas)
    procesados.add(par_key)
    pares_nuevos += 1
    print(f"    Ventas detectadas: {len(nuevas):,}")

if pares_nuevos == 0:
    print("\n  Todos los pares ya fueron procesados. Sin novedades.")
else:
    with open(HIST_PATH, 'w', encoding='utf-8') as f:
        json.dump(historial, f, ensure_ascii=False, indent=2)
    print(f"\n  Historial actualizado: {len(historial):,} ventas acumuladas")

# Guardar registro de procesados
with open(PROC_PATH, 'w', encoding='utf-8') as f:
    json.dump([list(p) for p in procesados], f, ensure_ascii=False)

# ── GENERAR data.json CON EL STOCK MÁS RECIENTE ──────────────────────────────
print(f"\n  Generando data.json con stock más reciente: {dated[-1][1].name}")
df_last, desc_last, cod_last, costo_last = leer_stock(dated[-1][1])
if df_last is None:
    print("ERROR: no se pudo leer el archivo de stock más reciente.")
    sys.exit(1)

records   = []
rubros_set = set()
marcas_set = set()

for _, row in df_last.iterrows():
    rubro  = clean_text(safe_str(row.get('Rubro', '')))
    if rubro in RUBROS_EXCLUIDOS: continue
    familia = clean_text(safe_str(row.get('Familia', '')))
    if familia in FAMILIAS_EXCLUIDAS: continue
    if safe_int(row.get(COL_KEY, 0)) != 0: continue

    try:    codigo = str(int(float(row[cod_last]))).zfill(6)
    except: codigo = clean_text(str(row[cod_last]))

    branch_stocks = {DEPOSITOS[d]: safe_int(row.get(d,0)) for d in DEPOSITOS if d != COL_KEY}
    if not any(v == 1 for v in branch_stocks.values()): continue

    art  = clean_text(row.get(desc_last, ''))
    marca = clean_text(safe_str(row.get('Marca', '')))
    rubros_set.add(rubro)
    marcas_set.add(marca)
    records.append({
        'codigo':       codigo,
        'articulo':     art,
        'rubro':        rubro,
        'familia':      familia,
        'marca':        marca,
        'size':         'N/A',
        'costo_repo':   float(row[costo_last]) if costo_last and pd.notna(row.get(costo_last)) else 0,
        'col_stock':    0,
        'branch_stocks': branch_stocks,
    })

out = {
    'branches':    SUCURSALES,
    'deposito_map': DEPOSITOS,
    'products':    records,
    'source_file': dated[-1][1].name,
    'rubros':      sorted(rubros_set),
    'marcas':      sorted(marcas_set),
}
with open(DATA_PATH, 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, indent=2)

print(f"  Productos COL=0 con última unidad: {len(records):,}")
print(f"  Rubros: {len(rubros_set)}")
print(f"  Guardado: data.json")
