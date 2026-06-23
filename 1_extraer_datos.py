"""
Paso 1: Lee el archivo de stock total y genera data.json.
Formato esperado: Stock Total DDMM.xlsx (exportado del sistema con todos los depósitos)
"""
import pandas as pd, json, re, sys, glob, os
from pathlib import Path

BASE = Path(__file__).parent

# Mapeo de códigos de depósito → nombre de sucursal
DEPOSITOS = {
    'D01': '25 de Mayo',
    'D02': 'Tafi del Valle',
    'D04': 'Banda',
    'D05': 'Maipú',
    'D06': 'Concepción',
    'D07': 'Lules',
    'D09': 'Outlet',
    'D10': 'Rosario',
    'D11': 'Cafayate',
    'D12': 'COL',
    'D13': 'Yerba Buena',
    'D15': 'Monteros',
    'D16': 'Tafi Viejo',
    'D17': 'Famailla',
    'D18': 'Santa María',
    'D23': 'Big',
    'D24': 'Aguilares',
    'D34': 'Alberdi',
}

COL_KEY = 'D12'

# Sucursales comerciales (las que aparecen en el dashboard - excluye COL)
SUCURSALES = [v for k, v in DEPOSITOS.items() if k != COL_KEY]

# Familias excluidas del análisis
FAMILIAS_EXCLUIDAS = {
    'INDUMENTARIA', 'CALZADO', 'JUGUETERIA', 'INSUMOS',
    'FOTOGRAFIA', 'FLETES - PROMO REG.', 'BLANCO', 'BAZAR',
}

# Rubros excluidos del análisis (indumentaria, calzado, internos/seguros)
RUBROS_EXCLUIDOS = {
    # Indumentaria
    'ACC.DE INDUMENTARIA','BERMUDA','BODY','BUZO','CALZA','CAMISA','CAMISETAS',
    'CAMPERA','CHALECO','CHOMBA','CONJUNTO DEP.','GORRA','GUANTE','GUARDAPOLVO',
    'JEANS','MEDIAS','PANTALON','POLAR','POLERA','POLLERA','REMERA','ROPA INTERIOR',
    'SHORT','SWEATER','TOP','UNIFORMES',
    # Calzado
    'BOTAS','BOTINES','OJOTAS','SANDALIA','ZAPATILLAS','ZAPATOS',
    # Seguros / conceptos internos
    'BOLSAS BARBIERI','FLETE - CONCEPTOS VARIOS','GARANTIAS',
    'INSUMOS INTERNOS','PROMO REGALO','REGALOS EMPRESARIALES',
}

def find_excel():
    """Busca el archivo de stock más reciente en la carpeta del proyecto."""
    patterns = ['Stock Total*.xlsx', 'stock total*.xlsx', 'Stock*.xlsx']
    for pat in patterns:
        files = sorted(BASE.glob(pat), key=lambda f: f.stat().st_mtime, reverse=True)
        # Excluir archivos generados por nosotros
        files = [f for f in files if 'Ultima Unidad' not in f.name and 'Dashboard' not in f.name]
        if files:
            return files[0]
    return None

excel_file = find_excel()
if not excel_file:
    print("ERROR: No se encontró archivo de stock en esta carpeta.")
    print("Copiá el archivo 'Stock Total DDMM.xlsx' aquí y volvé a ejecutar.")
    sys.exit(1)

print(f"Leyendo: {excel_file.name} ...")
df = pd.read_excel(excel_file, header=0)

# Verificar columnas clave
required = ['Codigo', 'Descripción', 'Rubro', 'Marca', COL_KEY]
# Normalizar nombres de columna con encoding issues
col_rename = {}
for col in df.columns:
    clean = col.replace('ó','ó').strip()
    # Descripción puede venir con encoding roto
    if 'escripci' in col:
        col_rename[col] = 'Descripción'
col_rename_fixed = {}
for col in df.columns:
    normalized = col.encode('utf-8', 'replace').decode('utf-8')
    if 'escripci' in normalized.lower():
        col_rename_fixed[col] = 'Descripción'
df = df.rename(columns=col_rename_fixed)

# Verificar que tengamos las columnas D
dep_cols_present = [d for d in DEPOSITOS.keys() if d in df.columns]
if not dep_cols_present:
    print(f"ERROR: No se encontraron columnas de depósito (D01, D02...) en el archivo.")
    print(f"Columnas encontradas: {list(df.columns)}")
    sys.exit(1)

# Columna de descripción (puede llamarse 'Descripción', 'Descripcion', 'Mercaderia', etc.)
desc_col = None
for c in df.columns:
    if any(x in c.lower() for x in ['escripci', 'mercader', 'articul']):
        desc_col = c
        break
if not desc_col:
    desc_col = df.columns[12]  # fallback por posición

# Columna código
cod_col = 'Codigo' if 'Codigo' in df.columns else df.columns[10]

# Columna costo repo (puede no existir en este formato)
costo_col = None
for c in df.columns:
    if 'costo' in c.lower() and 'repo' in c.lower():
        costo_col = c
        break

def clean_text(s):
    if not isinstance(s, str): return str(s)
    return s.replace('_x000D_', '').replace('\n', ' ').strip()

def safe_int(v):
    try:
        return int(float(v)) if pd.notna(v) else 0
    except (ValueError, TypeError):
        return 0

def safe_str(v):
    if pd.isna(v): return ''
    return str(v).strip()

# Filtrar filas sin código
df = df[df[cod_col].notna()].copy()

records = []
rubros_set = set()
marcas_set = set()

for _, row in df.iterrows():
    rubro = clean_text(safe_str(row.get('Rubro', '')))
    if rubro in RUBROS_EXCLUIDOS:
        continue

    familia = clean_text(safe_str(row.get('Familia', '')))
    if familia in FAMILIAS_EXCLUIDAS:
        continue

    col_stock = safe_int(row.get(COL_KEY, 0))
    if col_stock != 0:
        continue  # solo productos sin reposición posible

    branch_stocks = {}
    for dep_code, branch_name in DEPOSITOS.items():
        if dep_code == COL_KEY:
            continue
        branch_stocks[branch_name] = safe_int(row.get(dep_code, 0))

    # Tamaño desde descripción (especialmente para LED)
    art = clean_text(row.get(desc_col, ''))
    size_match = re.search(r"(\d{2,3})['\"]", art) or re.search(r'\b(\d{2,3})\b', art)
    size = (size_match.group(1) + '"') if size_match else 'N/A'

    marca = clean_text(safe_str(row.get('Marca', '')))

    try:
        codigo = str(int(float(row[cod_col]))).zfill(6)
    except Exception:
        codigo = clean_text(str(row[cod_col]))

    rubros_set.add(rubro)
    marcas_set.add(marca)

    # Solo guardamos productos con al menos una sucursal con última unidad
    if not any(v == 1 for v in branch_stocks.values()):
        continue

    rec = {
        'codigo': codigo,
        'articulo': art,
        'rubro': rubro,
        'familia': familia,
        'marca': marca,
        'size': size,
        'costo_repo': float(row[costo_col]) if costo_col and pd.notna(row.get(costo_col)) else 0,
        'col_stock': col_stock,
        'branch_stocks': branch_stocks,
    }
    records.append(rec)

last_unit_count = sum(
    1 for r in records if any(v == 1 for v in r['branch_stocks'].values())
)

out = {
    'branches': SUCURSALES,
    'deposito_map': DEPOSITOS,
    'products': records,
    'source_file': excel_file.name,
    'rubros': sorted(rubros_set),
    'marcas': sorted(marcas_set),
}

out_path = BASE / 'data.json'
with open(out_path, 'w', encoding='utf-8') as f:
    json.dump(out, f, ensure_ascii=False, indent=2)

print(f"  Productos COL=0: {len(records):,}")
print(f"  Con última unidad: {last_unit_count:,}")
print(f"  Rubros: {len(rubros_set)}")
print(f"  Guardado: data.json")
