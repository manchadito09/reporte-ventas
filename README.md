# Informe de ventas automático: de Excel a PDF

**Convierte tu Excel de ventas en un informe PDF listo para enviar, en segundos y sin tocarlo a mano.**

Sin errores de copiar y pegar. Y el informe sale igual cada vez que lo generas.

## Antes y después

| Antes: Excel en bruto (filas vacías, duplicados, nombres mal escritos) | Después: informe PDF de 1 página |
|---|---|
| ![Excel de ventas sin procesar](docs/antes.png) | ![Informe PDF generado](docs/despues.png) |

## Qué hace

Lee un Excel de ventas (probado con 25.000 filas, unos 10 segundos) y genera un PDF con:

- **KPIs**: total vendido, número de pedidos y ticket medio
- **Gráfico**: ventas por mes
- **Top 5 productos** por ingresos
- **Ventas por categoría**, con su porcentaje
- **Nota de calidad**: cuántas filas se limpiaron y por qué

Antes de calcular nada, **limpia los datos**:

| Problema en el Excel | Qué hace el programa |
|---|---|
| Filas vacías | Las descarta |
| Filas duplicadas (copiar y pegar) | Las descarta |
| `"  PACK 500 FOLIOS "`, `"pack 500 folios"` | Las unifica en `"Pack 500 folios"` |
| Cantidades a 0 o negativas, precios vacíos, fechas rotas | Descarta la fila |
| Fechas en texto (`2025-01-05` o `05/01/2025`) | Las entiende sin confundir día y mes |

Si el Excel no existe o le faltan columnas, lo avisa con un mensaje claro en español.

## Instalación (Windows, PowerShell)

Necesitas [Python 3.10 o superior](https://www.python.org/downloads/).

```powershell
# 1. Descargar el proyecto
git clone https://github.com/manchadito09/reporte-ventas.git
cd reporte-ventas

# 2. Crear y activar un entorno virtual (una "caja" aislada para las librerías)
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# 3. Instalar las librerías
pip install -r requirements.txt
```

> Si `Activate.ps1` da un error de permisos, ejecuta una vez
> `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` y vuelve a intentarlo.

## Uso

```powershell
# Generar el informe a partir del Excel de ejemplo
python generate_report.py data/sample_sales.xlsx --output output/report.pdf

# Abrirlo
start output\report.pdf
```

Para usar tu propio Excel, basta con que tenga estas columnas:

| fecha | pedido_id | producto | categoria | cantidad | precio_unitario | ciudad |
|---|---|---|---|---|---|---|

### Datos de ejemplo

`data/sample_sales.xlsx` contiene **datos inventados** (ninguna persona ni empresa real):
**un año completo (2025)** con unas 25.000 líneas de pedido, 14.000 pedidos y 60 productos en 6 categorías,
con temporada realista (rebajas, verano flojo, Black Friday, Navidad) y suciedad puesta a propósito.
Para volver a crearlo:

```powershell
python generate_data.py
```

Siempre sale idéntico (usa una semilla fija). Tarda unos 10 segundos.

## Tests

```powershell
pytest -v
```

10 tests comprueban la limpieza de datos y los cálculos (pedidos únicos, ticket medio, top 5, porcentajes, los 12 meses en orden...).

## Tecnologías

| Herramienta | Para qué |
|---|---|
| **Python 3** | Lenguaje del proyecto |
| **pandas** | Leer, limpiar y resumir los datos |
| **openpyxl** | Leer y escribir archivos `.xlsx` |
| **matplotlib** | Dibujar el gráfico |
| **fpdf2** | Montar el PDF |
| **pytest** | Tests automáticos |
| **DejaVu Sans** | Fuente libre incluida en `fonts/` para mostrar `€` y tildes en el PDF |

## Estructura

```
reporte-ventas/
├── generate_report.py   # programa principal: Excel -> PDF
├── generate_data.py     # crea el Excel de ejemplo con datos falsos
├── data/                # Excel de ejemplo
├── fonts/               # fuente DejaVu Sans + licencia
├── tests/               # tests con pytest
├── docs/                # imágenes del README
└── requirements.txt
```

## Mejoras futuras

- Enviar el informe por email automáticamente
- Programar la ejecución (por ejemplo, cada lunes a las 8:00)
- Aceptar CSV y otros formatos además de `.xlsx`
- Pequeña interfaz gráfica para elegir el archivo sin usar la terminal
- Filtros por ciudad o por rango de fechas
- Comparar con el periodo anterior (por ejemplo, +12 % frente al año anterior)

## Licencia de la fuente

DejaVu Sans se distribuye bajo su propia licencia libre: ver [`fonts/LICENSE_DEJAVU`](fonts/LICENSE_DEJAVU).
