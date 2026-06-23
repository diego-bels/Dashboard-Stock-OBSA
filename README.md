# Dashboard Stock — Última Unidad Sin Reposición

Herramienta para detectar productos con **última unidad en sucursal y sin stock en COL** (sin posibilidad de reposición), para poder accionar comercialmente sobre esas unidades.

---

## Instalación (primera vez)

### Paso 1 — Aceptar la invitación de GitHub

1. Revisá tu email — te llegó una invitación de GitHub de parte de Diego
2. Hacé clic en **"Accept invitation"**

---

### Paso 2 — Descargar el proyecto con GitHub Desktop

1. Abrí **GitHub Desktop**
2. `File → Clone repository`
3. Buscá **Dashboard-Stock-OBSA** en la lista y hacé clic en **Clone**
4. Anotá la carpeta donde se guardó (ej: `C:\Users\Victoria\Documents\GitHub\Dashboard-Stock-OBSA`)

---

### Paso 3 — Instalar las librerías de Python

> Si ya instalaste las librerías para el proyecto STOCK-ALERTAS, este paso ya está listo — podés saltearlo.

1. Abrí el **Explorador de archivos** y navegá a la carpeta donde clonaste el proyecto
2. En la barra de dirección escribí `cmd` y presioná Enter (se abre una ventana negra)
3. Copiá este comando, pegalo y presioná Enter:
```
pip install pandas openpyxl
```
4. Esperá que termine (puede tardar un par de minutos)

---

## Uso diario

### Paso 1 — Obtener el archivo de stock actualizado

Pedile a Diego el archivo Excel de stock exportado del sistema.  
El archivo se llama algo como: **`Stock Total 2206.xlsx`**

Copiá ese archivo dentro de la carpeta del proyecto (la misma donde está `Actualizar Dashboard.bat`).

> ⚠️ Si hay un archivo de stock anterior en la carpeta, reemplazalo con el nuevo.

---

### Paso 2 — Actualizar el dashboard

Hacé **doble clic** en el archivo:

```
Actualizar Dashboard.bat
```

Se abre una ventana negra que muestra el progreso. Cuando termine, el dashboard se abre automáticamente en el navegador.

> Si la ventana negra muestra un error, mandále una captura de pantalla a Diego.

---

### Paso 3 — Usar el dashboard

El dashboard tiene dos vistas:

- **≡ Tabla** — todos los productos con última unidad, con columnas por sucursal
- **⊞ Por Sucursal** — tarjetas agrupadas por sucursal

**Filtros disponibles** (panel izquierdo):

| Filtro | Cómo usarlo |
|---|---|
| Código / Descripción | Escribí parte del código o nombre del producto |
| Rubro | Escribí ej: `LED`, `HELADERA`, `SPLIT` |
| Familia | Escribí ej: `T.V. / DVD`, `CLIMATIZACIÓN` |
| Marca | Escribí ej: `PHILIPS`, `LG`, `SAMSUNG` |
| Sucursales | Tildá/destildá para ver solo las que te interesan |

**Colores en la tabla:**

| Color | Significado |
|---|---|
| 🔴 Rojo (1) | Última unidad — hay que accionar |
| 🟡 Amarillo (2-4) | Stock bajo |
| 🟢 Verde (5+) | Stock normal |
| ⬜ Gris (0) | Sin stock |

---

### Paso 4 — Exportar a Excel para enviar a los responsables

1. Hacé clic en el botón **⬇ Exportar Excel** (arriba a la derecha)
2. Completá:
   - **Título** del informe (podés dejarlo como está)
   - **Destinatario** — ej: `Responsable Famailla` (aparece en el nombre del archivo)
   - **Sucursales a incluir** — elegí la opción que necesites
3. Hacé clic en **⬇ Descargar**

El archivo Excel se descarga con una sola hoja que muestra:

| Sucursal | Cant. Registros | Código | Descripción | Rubro | Marca |
|---|---|---|---|---|---|

---

## Sugerir cambios o mejoras

Si querés proponer un cambio en los scripts o en cómo funciona el dashboard:

1. En **GitHub Desktop**, hacé clic en `Branch → New branch`
2. Poné un nombre descriptivo, ej: `mejora-filtro-rubro`
3. Hacé los cambios en los archivos que corresponda
4. En GitHub Desktop: escribí un mensaje describiendo el cambio y hacé clic en **Commit**
5. Hacé clic en **Push origin**
6. Hacé clic en **Create Pull Request** → se le notifica a Diego para que lo revise y apruebe

---

## Contacto

Cualquier duda o problema: **Diego Ruiz** — diegorcmr@gmail.com
