# Dashboard Stock — Última Unidad Sin Reposición

Dashboard interactivo para detectar productos con **última unidad en sucursal y sin stock en COL** (sin posibilidad de reposición), permitiendo acción comercial rápida.

---

## Uso rápido

1. Copiá el archivo de stock exportado del sistema a esta carpeta  
   (nombre: `Stock Total DDMM.xlsx`)
2. Doble clic en **`Actualizar Dashboard.bat`**
3. Se abre el `Dashboard.html` en el navegador

---

## Requisitos

- **Python 3.x** instalado con `pip` ([python.org](https://python.org/downloads))
- Librerías: `pandas` y `openpyxl`

```
pip install pandas openpyxl
```

---

## Estructura del proyecto

```
Dashboard Stock LED/
├── 1_extraer_datos.py        ← Lee el Excel y genera data.json
├── 2_generar_dashboard.py    ← Genera Dashboard.html desde data.json
├── Actualizar Dashboard.bat  ← Ejecuta los dos pasos con doble clic
├── .gitignore
└── README.md
```

> `data.json` y `Dashboard.html` se generan localmente y **no se suben al repo**.  
> El archivo Excel fuente tampoco se sube (datos internos).

---

## Filtros disponibles en el dashboard

- **Código / Descripción** — búsqueda de texto libre
- **Rubro** — ej: LED, HELADERA, SPLIT
- **Familia** — ej: T.V. / DVD, CLIMATIZACIÓN
- **Marca** — ej: PHILIPS, LG, SAMSUNG
- **Sucursales** — selección múltiple con contador de alertas
- **Toggles** — "Última unidad" y "COL = 0"

---

## Exportar Excel

El botón **Exportar Excel** genera un `.xlsx` con una sola hoja:

| Sucursal | Cant. Registros | Código | Descripción | Rubro | Marca |
|---|---|---|---|---|---|

Se puede filtrar por sucursales seleccionadas o solo las que tienen alerta.

---

## Líneas excluidas del análisis

- **Indumentaria** (bermuda, camisa, campera, etc.)
- **Calzado** (zapatillas, botas, sandalias, etc.)
- **Internos / Seguros** (garantías, flete, insumos internos, etc.)

Para modificar, editá el set `RUBROS_EXCLUIDOS` en `1_extraer_datos.py`.

---

## Depósitos / Sucursales

| Código | Sucursal | | Código | Sucursal |
|---|---|---|---|---|
| D01 | 25 de Mayo | | D13 | Yerba Buena |
| D02 | Tafi del Valle | | D15 | Monteros |
| D03 | Post Venta | | D16 | Tafi Viejo |
| D04 | Banda | | D17 | Famailla |
| D05 | Maipú | | D18 | Santa María |
| D06 | Concepción | | D23 | Big |
| D07 | Lules | | D24 | Aguilares |
| D08 | YB Web | | D25 | PV Web |
| D09 | Outlet | | D26 | PV Muebles |
| D10 | Rosario | | D30 | Almacén Bs.As. |
| D11 | Cafayate | | D34 | Alberdi |
| D12 | **COL** (referencia) | | | |
