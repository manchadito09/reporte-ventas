"""
Genera Excels de 4 empresas FICTICIAS para comprobar que el informe
funciona con datos distintos al Excel de ejemplo.

Cada empresa pone a prueba un punto débil distinto:
  1. ropa_tienda.xlsx           -> periodo que cruza de año (nov 2024 - abr 2025)
  2. bebidas_distribuidora.xlsx -> venta al por mayor, cifras de millones
  3. electronica_tienda.xlsx    -> 15 categorías y nombres de producto muy largos
  4. oficina_excel_manual.xlsx  -> Excel "hecho a mano": cabeceras raras, columnas
                                   extra, fechas y precios como texto, devoluciones

Uso:
    python generate_company_samples.py
"""

import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

from generate_data import make_reproducible

OUTPUT_DIR = Path("data/empresas")

CITIES = ["Madrid", "Barcelona", "Valencia", "Sevilla", "Bilbao", "Zaragoza", "Málaga"]


def make_sales(rng: random.Random, start: date, end: date, catalog: dict,
               n_rows: int, quantities: list[int], cities: list[str]) -> pd.DataFrame:
    """Crea líneas de pedido aleatorias (limpias) a partir de un catálogo.

    catalog: producto -> (categoría, precio, peso)
    """
    names = list(catalog)
    weights = [w for _, _, w in catalog.values()]
    total_days = (end - start).days + 1
    rows = []
    order = 1
    while len(rows) < n_rows:
        day = start + timedelta(days=rng.randrange(total_days))
        city = rng.choice(cities)
        n_lines = rng.choices([1, 2, 3], weights=[60, 30, 10])[0]
        chosen = {rng.choices(names, weights=weights)[0] for _ in range(n_lines)}
        for product in sorted(chosen):
            category, price, _ = catalog[product]
            rows.append({
                "fecha": day,
                "pedido_id": f"P{order:06d}",
                "producto": product,
                "categoria": category,
                "cantidad": rng.choice(quantities),
                "precio_unitario": price,
                "ciudad": city,
            })
        order += 1
    df = pd.DataFrame(rows).sort_values(["fecha", "pedido_id"]).reset_index(drop=True)
    return df


def add_basic_dirt(rng: random.Random, df: pd.DataFrame, n_empty: int, n_dup: int) -> pd.DataFrame:
    """Filas vacías, duplicados y algún nombre en mayúsculas (lo típico)."""
    df = df.copy()
    for i in rng.sample(range(len(df)), k=len(df) // 15):
        df.loc[i, "producto"] = str(df.loc[i, "producto"]).upper()
    dups = df.iloc[rng.sample(range(len(df)), k=n_dup)]
    df = pd.concat([df, dups]).sort_index(kind="stable").reset_index(drop=True)
    empty = pd.DataFrame([{c: None for c in df.columns}] * n_empty)
    for _, row in empty.iterrows():
        pos = rng.randrange(1, len(df))
        df = pd.concat([df.iloc[:pos], row.to_frame().T, df.iloc[pos:]]).reset_index(drop=True)
    return df


def save(df: pd.DataFrame, filename: str) -> Path:
    """Guarda el Excel con columnas anchas y fecha legible, sin horas internas."""
    path = OUTPUT_DIR / filename
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="ventas")
        ws = writer.sheets["ventas"]
        for col_cells in ws.columns:
            longest = max(len(str(c.value)) if c.value is not None else 0 for c in col_cells[:200])
            ws.column_dimensions[col_cells[0].column_letter].width = min(max(10, longest + 2), 60)
        for cell in ws["A"][1:]:
            if isinstance(cell.value, (date, pd.Timestamp)):
                cell.number_format = "DD/MM/YYYY"
    make_reproducible(path)
    return path


# --------------------------------------------------------------------------
# 1) Tienda de ropa: cruza de año, pocos datos
# --------------------------------------------------------------------------
def company_clothing() -> pd.DataFrame:
    rng = random.Random(1)
    catalog = {
        "Camiseta básica": ("Camisetas", 12.99, 10),
        "Camiseta estampada": ("Camisetas", 17.99, 6),
        "Vaquero slim": ("Pantalones", 39.95, 6),
        "Chino beige": ("Pantalones", 34.95, 4),
        "Abrigo de lana": ("Abrigos", 119.00, 3),
        "Plumífero ligero": ("Abrigos", 79.90, 4),
        "Jersey de punto": ("Punto", 29.95, 7),
        "Cárdigan": ("Punto", 35.00, 4),
        "Zapatillas urbanas": ("Calzado", 59.90, 5),
        "Botines de piel": ("Calzado", 89.00, 3),
    }
    df = make_sales(rng, date(2024, 11, 1), date(2025, 4, 30), catalog,
                    n_rows=600, quantities=[1, 1, 1, 2, 3], cities=CITIES[:3])
    return add_basic_dirt(rng, df, n_empty=5, n_dup=6)


# --------------------------------------------------------------------------
# 2) Distribuidora de bebidas: al por mayor, cifras enormes
# --------------------------------------------------------------------------
def company_drinks() -> pd.DataFrame:
    rng = random.Random(2)
    catalog = {
        "Agua mineral 1,5 L (pack 6)": ("Agua", 2.10, 10),
        "Agua con gas 1 L (pack 6)": ("Agua", 3.40, 5),
        "Refresco de cola 2 L": ("Refrescos", 1.35, 9),
        "Refresco de naranja 2 L": ("Refrescos", 1.25, 6),
        "Cerveza lata 33 cl (pack 24)": ("Cerveza", 14.90, 8),
        "Cerveza sin alcohol (pack 24)": ("Cerveza", 13.50, 4),
        "Vino tinto crianza 75 cl": ("Vinos", 6.80, 5),
        "Vino blanco verdejo 75 cl": ("Vinos", 5.90, 4),
        "Zumo de naranja 1 L": ("Zumos", 1.95, 6),
        "Bebida isotónica 50 cl": ("Refrescos", 0.95, 5),
    }
    df = make_sales(rng, date(2025, 1, 1), date(2025, 12, 31), catalog,
                    n_rows=8000, quantities=[200, 500, 800, 1000, 2500], cities=CITIES)
    return add_basic_dirt(rng, df, n_empty=20, n_dup=30)


# --------------------------------------------------------------------------
# 3) Tienda de electrónica: 15 categorías y nombres muy largos
# --------------------------------------------------------------------------
def company_electronics() -> pd.DataFrame:
    rng = random.Random(3)
    categories = [
        "Televisores", "Portátiles", "Sobremesa", "Tablets", "Móviles", "Fotografía",
        "Audio y HiFi", "Pequeño electrodoméstico", "Gran electrodoméstico", "Climatización",
        "Consolas y videojuegos", "Redes y conectividad", "Impresión", "Smart home", "Accesorios",
    ]
    catalog = {}
    for i, cat in enumerate(categories):
        catalog[f"{cat} modelo básico"] = (cat, round(30 + i * 17.5, 2), 6)
        catalog[f"{cat} modelo premium de alta gama con garantía extendida de 3 años"] = (
            cat, round(400 + i * 85.0, 2), 2)
    # Un producto con nombre larguísimo que además vende mucho (saldrá en el top 5)
    catalog["Televisor OLED 77 pulgadas 4K 120 Hz con HDR, Dolby Vision y barra de sonido incluida"] = (
        "Televisores", 3299.00, 4)
    df = make_sales(rng, date(2025, 1, 1), date(2025, 12, 31), catalog,
                    n_rows=5000, quantities=[1, 1, 1, 2], cities=CITIES)
    return add_basic_dirt(rng, df, n_empty=10, n_dup=10)


# --------------------------------------------------------------------------
# 4) Excel "hecho a mano" en una oficina
# --------------------------------------------------------------------------
def company_manual_excel() -> pd.DataFrame:
    rng = random.Random(4)
    catalog = {
        "Tóner negro": ("Consumibles", 64.90, 6),
        "Cartucho color": ("Consumibles", 32.50, 6),
        "Papel A4 caja 5 paquetes": ("Papelería", 24.95, 10),
        "Sobres americanos (500)": ("Papelería", 18.40, 4),
        "Silla de oficina": ("Mobiliario", 145.00, 2),
        "Armario metálico": ("Mobiliario", 289.00, 1),
        "Calculadora de sobremesa": ("Equipos", 22.90, 4),
        "Plastificadora A3": ("Equipos", 79.00, 2),
    }
    df = make_sales(rng, date(2025, 1, 1), date(2025, 6, 30), catalog,
                    n_rows=900, quantities=[1, 2, 3, 5, 10], cities=["Madrid", "Getafe", "Alcorcón"])

    # Devoluciones: cantidad negativa (así las apuntan muchas oficinas)
    for i in rng.sample(range(len(df)), k=15):
        df.loc[i, "cantidad"] = -df.loc[i, "cantidad"]

    # Todo a texto, como si alguien lo hubiera tecleado
    df = df.astype(object)
    df["fecha"] = [d.strftime("%d/%m/%Y") for d in df["fecha"]]                # "05/01/2025"
    df["precio_unitario"] = [
        f"{p:.2f}".replace(".", ",") if rng.random() < 0.4 else p              # "64,90" (texto)
        for p in df["precio_unitario"]
    ]
    euro_rows = rng.sample(range(len(df)), k=10)
    df.loc[euro_rows, "precio_unitario"] = [
        f"{float(str(p).replace(',', '.')):.2f} €".replace(".", ",")          # "64,90 €"
        for p in df.loc[euro_rows, "precio_unitario"]
    ]

    # Columnas extra que el informe no usa
    df["vendedor"] = [rng.choice(["Ana", "Luis", "Marta", "Jorge"]) for _ in range(len(df))]
    df["observaciones"] = ["" if rng.random() < 0.9 else "urgente" for _ in range(len(df))]

    df = add_basic_dirt(rng, df, n_empty=8, n_dup=5)

    # Cabeceras escritas a mano y columnas en otro orden
    df = df.rename(columns={
        "fecha": "Fecha ", "pedido_id": "PEDIDO_ID", "producto": "Producto",
        "categoria": "Categoria", "cantidad": "Cantidad",
        "precio_unitario": "Precio_Unitario", "ciudad": " Ciudad",
    })
    return df[["PEDIDO_ID", "Fecha ", "vendedor", "Producto", "Categoria",
               "Cantidad", "Precio_Unitario", " Ciudad", "observaciones"]]


def main() -> None:
    companies = {
        "ropa_tienda.xlsx": company_clothing,
        "bebidas_distribuidora.xlsx": company_drinks,
        "electronica_tienda.xlsx": company_electronics,
        "oficina_excel_manual.xlsx": company_manual_excel,
    }
    for filename, build in companies.items():
        path = save(build(), filename)
        print(f"Creado: {path}")


if __name__ == "__main__":
    main()
