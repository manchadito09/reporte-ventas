"""
Genera un informe de ventas en PDF (1 página) a partir de un Excel.

El informe incluye:
- KPIs: total vendido, número de pedidos y ticket medio
- Gráfico de barras con las ventas por mes
- Tabla con el top 5 de productos por ingresos
- Tabla con las ventas por categoría

Uso:
    python generate_report.py data/sample_sales.xlsx --output output/report.pdf
"""

import argparse
import io
import sys
from datetime import date
from pathlib import Path

import matplotlib

matplotlib.use("Agg")  # dibuja en memoria, sin abrir ventanas
import matplotlib.pyplot as plt
import pandas as pd
from fpdf import FPDF

# Columnas que el Excel tiene que traer sí o sí
REQUIRED_COLUMNS = [
    "fecha", "pedido_id", "producto", "categoria",
    "cantidad", "precio_unitario", "ciudad",
]
TEXT_COLUMNS = ["pedido_id", "producto", "categoria", "ciudad"]

MONTH_NAMES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun",
               "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
MONTH_NAMES_LONG = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio",
                    "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

# Carpeta base de los recursos (fuentes).
# - Ejecutado como script: la carpeta de este archivo.
# - Dentro del .exe (PyInstaller): la carpeta temporal donde se descomprime,
#   que PyInstaller guarda en sys._MEIPASS.
BASE_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).parent))

# Fuente TTF incluida en el repo: soporta el símbolo € y las tildes
FONT_DIR = BASE_DIR / "fonts"

# Tamaño del gráfico en pulgadas (ancho, alto). El PDF respeta esta proporción
CHART_SIZE = (7.5, 3.6)

# Colores (hex -> RGB). Un solo azul para los datos; grises para el texto
BLUE = "#2a78d6"
INK = "#0b0b0b"         # texto principal
INK_SOFT = "#52514e"    # texto secundario
LINE = "#e4e3df"        # líneas finas y bordes
PANEL = "#f4f4f2"       # fondo de las cajas de KPI


class ReportError(Exception):
    """Error pensado para el usuario: se muestra tal cual, sin traza técnica."""


# --------------------------------------------------------------------------
# 1) Cargar
# --------------------------------------------------------------------------
def load_data(path: Path) -> pd.DataFrame:
    """Lee el Excel y comprueba que tiene las columnas necesarias."""
    if not path.exists():
        raise ReportError(f"No se encuentra el archivo: {path}")
    if path.suffix.lower() != ".xlsx":
        raise ReportError(f"El archivo tiene que ser un Excel .xlsx: {path.name}")

    try:
        df = pd.read_excel(path)
    except Exception as exc:  # archivo corrupto, abierto por otro programa, etc.
        raise ReportError(f"No se pudo leer el Excel ({path.name}): {exc}") from exc

    # Normalizamos las cabeceras por si vienen con espacios o mayúsculas
    df.columns = [str(c).strip().lower() for c in df.columns]

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ReportError(
            "Al Excel le faltan columnas obligatorias: " + ", ".join(missing)
            + "\nColumnas esperadas: " + ", ".join(REQUIRED_COLUMNS)
        )
    return df[REQUIRED_COLUMNS]


# --------------------------------------------------------------------------
# 2) Limpiar
# --------------------------------------------------------------------------
def _canonical_names(names: pd.Series) -> pd.Series:
    """Unifica nombres que solo se diferencian en mayúsculas.

    Agrupa "PACK 500 FOLIOS", "pack 500 folios" y "Pack 500 folios"
    y se queda con la forma que más se repite (la que escribe casi todo el mundo).
    """
    key = names.str.lower()
    most_common = (
        pd.DataFrame({"key": key, "name": names})
        .groupby("key")["name"]
        .agg(lambda s: s.value_counts().index[0])
    )
    return key.map(most_common)


def _parse_dates(dates: pd.Series) -> pd.Series:
    """Convierte la columna de fechas, venga como venga.

    - Si Excel ya las guarda como fecha, se dejan tal cual.
    - Si vienen como texto, se aceptan dos formatos:
        "2026-01-05"  (internacional: año-mes-día)
        "05/01/2026"  (español: día/mes/año)
    Probamos cada formato por separado para no confundir día y mes.
    Lo que no encaje en ninguno se queda vacío (NaT) y se descarta después.
    """
    if pd.api.types.is_datetime64_any_dtype(dates):
        return dates
    iso = pd.to_datetime(dates, format="ISO8601", errors="coerce")
    spanish = pd.to_datetime(dates, format="%d/%m/%Y", errors="coerce")
    return iso.fillna(spanish)


def clean_data(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """Limpia el Excel y devuelve (datos_limpios, resumen_de_la_limpieza).

    Pasos, en este orden:
    1. Quitar filas vacías
    2. Arreglar textos: espacios de más y mayúsculas mezcladas
    3. Quitar duplicados exactos (después del paso 2, para pillar
       duplicados que solo se diferenciaban en un espacio)
    4. Convertir tipos y descartar filas con datos imposibles
    5. Calcular el ingreso de cada línea
    """
    df = df.copy()  # no tocamos el DataFrame original
    report = {"rows_in": len(df)}

    # 1) Filas totalmente vacías
    empty = df.isna().all(axis=1)
    df = df[~empty]
    report["empty_rows"] = int(empty.sum())

    # 2) Textos: quitar espacios en los bordes y dejar uno solo entre palabras
    before = df["producto"].astype("string")  # copia para contar lo corregido
    for col in TEXT_COLUMNS:
        df[col] = (
            df[col].astype("string")
            .str.strip()
            .str.replace(r"\s+", " ", regex=True)
        )
    df["producto"] = _canonical_names(df["producto"])
    df["categoria"] = _canonical_names(df["categoria"])
    df["ciudad"] = _canonical_names(df["ciudad"])
    report["names_fixed"] = int((before != df["producto"]).sum())

    # 3) Duplicados exactos (misma fila entera)
    dup = df.duplicated()
    df = df[~dup]
    report["duplicates"] = int(dup.sum())

    # 4) Tipos. errors="coerce" = si algo no se puede convertir, se queda vacío (NaN)
    df["fecha"] = _parse_dates(df["fecha"])
    df["cantidad"] = pd.to_numeric(df["cantidad"], errors="coerce")
    df["precio_unitario"] = pd.to_numeric(df["precio_unitario"], errors="coerce")

    invalid = (
        df[REQUIRED_COLUMNS].isna().any(axis=1)
        | (df["cantidad"] <= 0)
        | (df["precio_unitario"] < 0)
    )
    df = df[~invalid]
    report["invalid_rows"] = int(invalid.sum())

    df["cantidad"] = df["cantidad"].astype(int)

    # 5) Ingreso de cada línea
    df["ingreso"] = df["cantidad"] * df["precio_unitario"]

    report["rows_out"] = len(df)
    return df.reset_index(drop=True), report


# --------------------------------------------------------------------------
# 3) Resumir
# --------------------------------------------------------------------------
def compute_summary(df: pd.DataFrame) -> dict:
    """Calcula los KPIs y las tablas del informe a partir de datos ya limpios."""
    if df.empty:
        raise ReportError("No quedan filas válidas después de limpiar el Excel.")

    total_sales = float(df["ingreso"].sum())
    n_orders = int(df["pedido_id"].nunique())  # pedidos ÚNICOS, no filas

    # Ventas por mes. "period" agrupa por año+mes (ene-2026 != ene-2027)
    monthly = df.groupby(df["fecha"].dt.to_period("M"))["ingreso"].sum().sort_index()

    top_products = (
        df.groupby("producto")
        .agg(unidades=("cantidad", "sum"), ingresos=("ingreso", "sum"))
        .sort_values("ingresos", ascending=False)
        .head(5)
        .reset_index()
    )

    by_category = (
        df.groupby("categoria")
        .agg(pedidos=("pedido_id", "nunique"), ingresos=("ingreso", "sum"))
        .sort_values("ingresos", ascending=False)
        .reset_index()
    )
    by_category["porcentaje"] = by_category["ingresos"] / total_sales * 100

    return {
        "total_sales": total_sales,
        "n_orders": n_orders,
        "avg_ticket": total_sales / n_orders,
        "date_from": df["fecha"].min().date(),
        "date_to": df["fecha"].max().date(),
        "monthly": monthly,
        "top_products": top_products,
        "by_category": by_category,
    }


# --------------------------------------------------------------------------
# Formato español de números
# --------------------------------------------------------------------------
def format_number(value: float, decimals: int = 0) -> str:
    """1234567.891 -> '1.234.567,89' (punto para miles, coma para decimales)."""
    text = f"{value:,.{decimals}f}"  # formato inglés: 1,234,567.89
    # Intercambiamos , y . usando un carácter temporal
    return text.replace(",", "#").replace(".", ",").replace("#", ".")


def format_eur(value: float, decimals: int = 2) -> str:
    return f"{format_number(value, decimals)} €"


def format_period(date_from: date, date_to: date) -> str:
    """Ej.: 'año 2025', 'enero – junio 2026' o 'noviembre 2025 – febrero 2026'."""
    # Año natural completo: de enero a diciembre del mismo año
    if (date_from.year == date_to.year
            and (date_from.month, date_from.day) == (1, 1)
            and (date_to.month, date_to.day) == (12, 31)):
        return f"año {date_to.year}"
    m1 = MONTH_NAMES_LONG[date_from.month - 1]
    m2 = MONTH_NAMES_LONG[date_to.month - 1]
    if date_from.year == date_to.year:
        if date_from.month == date_to.month:
            return f"{m1} {date_to.year}"
        return f"{m1} – {m2} {date_to.year}"
    return f"{m1} {date_from.year} – {m2} {date_to.year}"


# --------------------------------------------------------------------------
# 4) Gráfico
# --------------------------------------------------------------------------
def build_chart(monthly: pd.Series) -> io.BytesIO:
    """Dibuja las ventas por mes y devuelve la imagen PNG en memoria.

    En memoria (BytesIO) = no se crea ningún archivo temporal en disco.
    """
    labels = [
        MONTH_NAMES[p.month - 1] + ("" if p.year == monthly.index[-1].year else f" {p.year % 100}")
        for p in monthly.index
    ]
    values = monthly.to_numpy()

    fig, ax = plt.subplots(figsize=CHART_SIZE, dpi=200)
    bars = ax.bar(labels, values, width=0.55, color=BLUE)

    # Valor encima de cada barra, en miles (la unidad va en el título).
    # Así sobran el eje Y y la cuadrícula. Con cifras grandes, sin decimales
    # para que las 12 etiquetas quepan sin pisarse.
    decimals = 0 if values.max() >= 100_000 else 1
    for bar, value in zip(bars, values):
        ax.annotate(
            format_number(value / 1000, decimals),
            xy=(bar.get_x() + bar.get_width() / 2, bar.get_height()),
            xytext=(0, 4), textcoords="offset points",
            ha="center", va="bottom", fontsize=8.5, color=INK,
        )

    # Menos tinta: fuera marco, eje Y y marcas; solo la línea base
    for side in ("top", "right", "left"):
        ax.spines[side].set_visible(False)
    ax.spines["bottom"].set_color(LINE)
    ax.set_yticks([])
    ax.tick_params(axis="x", length=0, labelsize=9, colors=INK_SOFT, pad=6)
    ax.set_ylim(0, values.max() * 1.18)  # hueco para las etiquetas de arriba

    fig.tight_layout(pad=0.4)
    buffer = io.BytesIO()
    fig.savefig(buffer, format="png")
    plt.close(fig)  # liberamos memoria
    buffer.seek(0)
    return buffer


# --------------------------------------------------------------------------
# 5) PDF
# --------------------------------------------------------------------------
def _rgb(hex_color: str) -> tuple[int, int, int]:
    hex_color = hex_color.lstrip("#")
    return tuple(int(hex_color[i:i + 2], 16) for i in (0, 2, 4))


class ReportPDF(FPDF):
    """PDF A4 con nuestra fuente y un pie de página fijo."""

    def __init__(self, source_name: str):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.source_name = source_name
        self.add_font("DejaVu", "", str(FONT_DIR / "DejaVuSans.ttf"))
        self.add_font("DejaVu", "B", str(FONT_DIR / "DejaVuSans-Bold.ttf"))
        self.set_margins(18, 16, 18)
        self.set_auto_page_break(False)  # queremos 1 sola página, sin saltos

    def footer(self):
        self.set_y(-14)
        self.set_draw_color(*_rgb(LINE))
        self.line(self.l_margin, self.get_y(), self.w - self.r_margin, self.get_y())
        self.ln(2)
        self.set_font("DejaVu", "", 7.5)
        self.set_text_color(*_rgb(INK_SOFT))
        today = date.today().strftime("%d/%m/%Y")
        self.cell(0, 5, f"Fuente: {self.source_name}", align="L")
        self.set_x(self.l_margin)
        self.cell(0, 5, f"Generado automáticamente el {today}", align="R")


def _section_title(pdf: FPDF, text: str) -> None:
    pdf.set_font("DejaVu", "B", 11)
    pdf.set_text_color(*_rgb(INK))
    pdf.cell(0, 7, text, new_x="LMARGIN", new_y="NEXT")


def _kpi_box(pdf: FPDF, x: float, y: float, w: float, label: str, value: str) -> None:
    pdf.set_fill_color(*_rgb(PANEL))
    pdf.rect(x, y, w, 22, style="F")
    pdf.set_fill_color(*_rgb(BLUE))
    pdf.rect(x, y, 1.2, 22, style="F")  # rayita azul a la izquierda
    pdf.set_xy(x + 5, y + 3.5)
    pdf.set_font("DejaVu", "", 8.5)
    pdf.set_text_color(*_rgb(INK_SOFT))
    pdf.cell(w - 8, 5, label.upper())
    pdf.set_xy(x + 5, y + 10)
    pdf.set_font("DejaVu", "B", 16)
    pdf.set_text_color(*_rgb(INK))
    pdf.cell(w - 8, 8, value)


def _table(pdf: FPDF, x: float, w: float, headers: list[str], rows: list[list[str]],
           col_widths: list[float]) -> None:
    """Tabla sencilla: cabecera en gris, líneas finas, números alineados a la derecha."""
    widths = [w * c for c in col_widths]
    aligns = ["L"] + ["R"] * (len(headers) - 1)

    pdf.set_x(x)
    pdf.set_font("DejaVu", "B", 8)
    pdf.set_text_color(*_rgb(INK_SOFT))
    for header, cw, al in zip(headers, widths, aligns):
        pdf.cell(cw, 6, header, align=al)
    pdf.ln(6)
    pdf.set_draw_color(*_rgb(LINE))
    pdf.line(x, pdf.get_y(), x + w, pdf.get_y())

    pdf.set_font("DejaVu", "", 8.5)
    pdf.set_text_color(*_rgb(INK))
    for row in rows:
        pdf.set_x(x)
        for value, cw, al in zip(row, widths, aligns):
            pdf.cell(cw, 7, value, align=al)
        pdf.ln(7)
        pdf.line(x, pdf.get_y(), x + w, pdf.get_y())


def build_pdf(summary: dict, chart_png: io.BytesIO, cleaning: dict,
              source_name: str, output_path: Path) -> None:
    """Monta el informe de 1 página y lo guarda en output_path."""
    pdf = ReportPDF(source_name)
    pdf.add_page()
    content_w = pdf.w - pdf.l_margin - pdf.r_margin

    # --- Cabecera ---
    pdf.set_font("DejaVu", "B", 20)
    pdf.set_text_color(*_rgb(INK))
    pdf.cell(0, 10, "Informe de ventas", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("DejaVu", "", 10.5)
    pdf.set_text_color(*_rgb(INK_SOFT))
    period = format_period(summary["date_from"], summary["date_to"])
    pdf.cell(0, 6, period[0].upper() + period[1:], new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)

    # --- KPIs: 3 cajas en fila ---
    gap = 5
    box_w = (content_w - 2 * gap) / 3
    y = pdf.get_y()
    kpis = [
        ("Total vendido", format_eur(summary["total_sales"], 0)),  # sin céntimos: son ruido en millones
        ("Pedidos", format_number(summary["n_orders"])),
        ("Ticket medio", format_eur(summary["avg_ticket"])),
    ]
    for i, (label, value) in enumerate(kpis):
        _kpi_box(pdf, pdf.l_margin + i * (box_w + gap), y, box_w, label, value)
    pdf.set_y(y + 22 + 8)

    # --- Gráfico ---
    _section_title(pdf, "Ventas por mes (miles de €)")
    chart_h = content_w * CHART_SIZE[1] / CHART_SIZE[0]  # misma proporción que la figura
    pdf.image(chart_png, x=pdf.l_margin, w=content_w, h=chart_h)
    pdf.ln(6)

    # --- Dos tablas lado a lado ---
    y = pdf.get_y()
    col_gap = 8
    left_w = content_w * 0.56
    right_w = content_w - left_w - col_gap
    right_x = pdf.l_margin + left_w + col_gap

    top = summary["top_products"]
    _section_title(pdf, "Top 5 productos por ingresos")
    _table(
        pdf, pdf.l_margin, left_w,
        ["Producto", "Unidades", "Ingresos"],
        [[r.producto, format_number(r.unidades), format_eur(r.ingresos, 0)]
         for r in top.itertuples()],
        [0.52, 0.18, 0.30],
    )
    y_after_left = pdf.get_y()

    pdf.set_xy(right_x, y)
    pdf.set_font("DejaVu", "B", 11)
    pdf.set_text_color(*_rgb(INK))
    pdf.cell(right_w, 7, "Ventas por categoría")
    pdf.set_xy(right_x, y + 7)
    cat = summary["by_category"]
    _table(
        pdf, right_x, right_w,
        ["Categoría", "Ingresos", "%"],
        [[r.categoria, format_eur(r.ingresos, 0), format_number(r.porcentaje, 1) + " %"]
         for r in cat.itertuples()],
        [0.40, 0.38, 0.22],
    )
    # --- Nota de calidad de datos: enseña al cliente que se limpió el Excel ---
    # Va anclada abajo, justo encima del pie, como la "letra pequeña"
    pdf.set_y(max(y_after_left, pdf.get_y(), pdf.h - 36))
    pdf.set_font("DejaVu", "B", 8.5)
    pdf.set_text_color(*_rgb(INK_SOFT))
    pdf.cell(0, 5, "Calidad de los datos", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("DejaVu", "", 8)
    note = (
        f"Se leyeron {format_number(cleaning['rows_in'])} filas del Excel. "
        f"Se descartaron {format_number(cleaning['empty_rows'])} filas vacías, "
        f"{format_number(cleaning['duplicates'])} duplicadas y "
        f"{format_number(cleaning['invalid_rows'])} con datos incompletos, "
        f"y se corrigieron {format_number(cleaning['names_fixed'])} nombres de producto mal escritos. "
        f"El informe usa {format_number(cleaning['rows_out'])} líneas de pedido válidas."
    )
    pdf.multi_cell(0, 4.5, note, align="L")  # "L" evita huecos raros entre palabras

    output_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        pdf.output(str(output_path))
    except PermissionError as exc:
        raise ReportError(
            f"No se pudo guardar {output_path}. ¿Está abierto en otro programa?"
        ) from exc


# --------------------------------------------------------------------------
# 6) Cadena completa: la usan la terminal (main) y el .exe (app.py)
# --------------------------------------------------------------------------
def generate_report(input_path: Path, output_path: Path) -> tuple[dict, dict]:
    """Excel -> PDF. Devuelve (resumen, informe_de_limpieza) por si se quieren mostrar."""
    raw = load_data(input_path)
    clean, cleaning = clean_data(raw)
    summary = compute_summary(clean)
    chart = build_chart(summary["monthly"])
    build_pdf(summary, chart, cleaning, input_path.name, output_path)
    return summary, cleaning


# --------------------------------------------------------------------------
# 7) Programa principal (terminal)
# --------------------------------------------------------------------------
def main() -> int:
    parser = argparse.ArgumentParser(description="Genera un informe de ventas en PDF desde un Excel.")
    parser.add_argument("input", help="Excel de ventas (.xlsx)")
    parser.add_argument(
        "--output", default="output/report.pdf",
        help="Ruta del PDF a crear (por defecto: output/report.pdf)",
    )
    args = parser.parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output)

    try:
        summary, _ = generate_report(input_path, output_path)
    except ReportError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1  # código de salida != 0 -> "algo fue mal" (útil en scripts)

    print(f"Informe creado: {output_path}")
    print(f"  Total vendido: {format_eur(summary['total_sales'])}")
    print(f"  Pedidos: {format_number(summary['n_orders'])}")
    print(f"  Ticket medio: {format_eur(summary['avg_ticket'])}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
