"""
Genera un Excel de ventas FALSO para probar el generador de reportes.

Los datos son inventados: no hay personas reales.
Lleva "suciedad" a propósito (filas vacías, duplicados, nombres mal escritos)
para demostrar que generate_report.py sabe limpiarla.

Uso:
    python generate_data.py
    python generate_data.py --output data/otro_archivo.xlsx
"""

import argparse
import random
import re
import zipfile
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

# Semilla fija: el azar sale siempre igual, así el Excel es reproducible
SEED = 42

# Periodo de ventas: año 2025 completo
START_DATE = date(2025, 1, 1)
END_DATE = date(2025, 12, 31)

# Fecha fija que se escribe "por dentro" del .xlsx (ver make_reproducible)
FIXED_TIMESTAMP = "2026-01-01T00:00:00Z"
FIXED_ZIP_TIME = (2026, 1, 1, 0, 0, 0)

# Cuántas líneas de pedido limpias queremos (luego se suma la suciedad)
TARGET_ROWS = 25_000

# Catálogo: producto -> (categoría, precio unitario en euros, peso)
# Peso = cuántas "papeletas" tiene en el sorteo. Los baratos se venden más a menudo.
PRODUCTS = {
    # Informática
    "Portátil 14 pulgadas": ("Informática", 649.00, 3),
    "Portátil 16 pulgadas": ("Informática", 999.00, 2),
    "Monitor 24 pulgadas": ("Informática", 139.90, 5),
    "Monitor 27 pulgadas": ("Informática", 229.90, 4),
    "Teclado mecánico": ("Informática", 79.95, 6),
    "Ratón inalámbrico": ("Informática", 24.99, 10),
    "Disco SSD 1TB": ("Informática", 89.50, 6),
    "Memoria USB 128GB": ("Informática", 14.99, 9),
    "Webcam Full HD": ("Informática", 49.90, 5),
    "Hub USB-C 7 en 1": ("Informática", 39.99, 6),
    # Oficina
    "Silla ergonómica": ("Oficina", 189.00, 3),
    "Mesa elevable": ("Oficina", 349.00, 2),
    "Lámpara LED de escritorio": ("Oficina", 34.90, 6),
    "Pack 500 folios": ("Oficina", 5.49, 12),
    "Archivador de palanca": ("Oficina", 3.95, 9),
    "Pizarra blanca 90x60": ("Oficina", 44.90, 3),
    "Destructora de papel": ("Oficina", 89.00, 2),
    "Reposapiés regulable": ("Oficina", 29.90, 4),
    "Bolígrafos pack 10": ("Oficina", 6.50, 11),
    "Organizador de cables": ("Oficina", 12.99, 7),
    # Audio
    "Auriculares Bluetooth": ("Audio", 59.99, 8),
    "Auriculares con cancelación": ("Audio", 199.00, 3),
    "Altavoz portátil": ("Audio", 45.00, 6),
    "Barra de sonido": ("Audio", 179.00, 2),
    "Micrófono USB": ("Audio", 69.90, 4),
    "Tocadiscos Bluetooth": ("Audio", 129.00, 1),
    "Radio despertador": ("Audio", 34.95, 3),
    "Cascos de estudio": ("Audio", 99.00, 2),
    "Altavoz inteligente": ("Audio", 54.99, 5),
    "Cable jack 3,5 mm": ("Audio", 7.99, 7),
    # Hogar
    "Cafetera espresso": ("Hogar", 119.00, 4),
    "Hervidor eléctrico": ("Hogar", 29.95, 6),
    "Purificador de aire": ("Hogar", 159.00, 3),
    "Freidora de aire": ("Hogar", 89.99, 6),
    "Robot aspirador": ("Hogar", 249.00, 3),
    "Batidora de vaso": ("Hogar", 59.90, 4),
    "Tostadora 2 ranuras": ("Hogar", 27.50, 5),
    "Ventilador de torre": ("Hogar", 49.00, 4),
    "Calefactor cerámico": ("Hogar", 39.90, 4),
    "Báscula de cocina": ("Hogar", 15.99, 6),
    # Telefonía
    "Smartphone gama media": ("Telefonía", 299.00, 4),
    "Smartphone gama alta": ("Telefonía", 849.00, 2),
    "Funda de silicona": ("Telefonía", 12.99, 12),
    "Protector de pantalla": ("Telefonía", 9.99, 12),
    "Cargador rápido 30W": ("Telefonía", 24.90, 9),
    "Batería externa 10000": ("Telefonía", 29.99, 8),
    "Soporte de coche": ("Telefonía", 16.50, 6),
    "Smartwatch": ("Telefonía", 179.00, 3),
    "Pulsera de actividad": ("Telefonía", 39.99, 5),
    "Cable USB-C 2 m": ("Telefonía", 8.99, 11),
    # Gaming
    "Consola portátil": ("Gaming", 329.00, 2),
    "Mando inalámbrico": ("Gaming", 59.99, 6),
    "Silla gaming": ("Gaming", 219.00, 2),
    "Alfombrilla XL": ("Gaming", 19.99, 7),
    "Ratón gaming": ("Gaming", 49.90, 5),
    "Teclado gaming RGB": ("Gaming", 89.00, 4),
    "Auriculares gaming": ("Gaming", 69.00, 5),
    "Tarjeta regalo 50 €": ("Gaming", 50.00, 6),
    "Volante de carreras": ("Gaming", 279.00, 1),
    "Base de carga mandos": ("Gaming", 24.99, 4),
}

CITIES = ["Madrid", "Barcelona", "Valencia", "Sevilla", "Bilbao", "Zaragoza", "Málaga"]
CITY_WEIGHTS = [30, 25, 12, 10, 8, 8, 7]

# Efecto temporada: multiplica cuántos pedidos hay cada mes (1 = normal)
# Rebajas en enero, bajón en verano, vuelta al cole en septiembre,
# Black Friday en noviembre y Navidad en diciembre.
MONTH_FACTOR = {
    1: 1.10, 2: 0.80, 3: 0.85, 4: 0.90, 5: 0.95, 6: 1.00,
    7: 0.85, 8: 0.70, 9: 1.00, 10: 0.95, 11: 1.40, 12: 1.55,
}

# Cantidad de suciedad a meter (proporcional al volumen: ~1,4 % de filas)
N_EMPTY_ROWS = 150
N_DUPLICATES = 200
MESSY_NAME_RATE = 0.08  # 8 % de filas con el nombre del producto mal escrito


def random_order_date() -> date:
    """Elige un día al azar, con más probabilidad en los meses fuertes."""
    total_days = (END_DATE - START_DATE).days + 1
    while True:
        day = START_DATE + timedelta(days=random.randrange(total_days))
        # Truco de "aceptar o rechazar": cuanto más alto el factor del mes,
        # más fácil es que el día se quede. Así cada mes tiene su peso.
        if random.random() < MONTH_FACTOR[day.month] / max(MONTH_FACTOR.values()):
            return day


def mess_up_name(name: str) -> str:
    """Estropea un nombre como lo haría una persona tecleando a mano."""
    variants = [
        f"  {name}",                  # espacios delante
        f"{name}   ",                 # espacios detrás
        name.upper(),                 # TODO EN MAYÚSCULAS
        name.lower(),                 # todo en minúsculas
        name.replace(" ", "  ", 1),   # doble espacio en medio
    ]
    return random.choice(variants)


def build_clean_rows() -> list[dict]:
    """Crea pedidos correctos. Un pedido puede tener varias líneas (productos)."""
    rows = []
    order_number = 1
    product_names = list(PRODUCTS)
    product_weights = [weight for _, _, weight in PRODUCTS.values()]

    while len(rows) < TARGET_ROWS:
        order_id = f"PED-{order_number:05d}"  # id provisional, se renumera al final
        order_date = random_order_date()
        city = random.choices(CITIES, weights=CITY_WEIGHTS)[0]

        # Cada pedido lleva de 1 a 4 productos distintos (casi siempre pocos)
        n_lines = random.choices([1, 2, 3, 4], weights=[50, 30, 15, 5])[0]
        chosen = set()
        while len(chosen) < n_lines:
            chosen.add(random.choices(product_names, weights=product_weights)[0])

        for product in sorted(chosen):
            category, price, _ = PRODUCTS[product]
            rows.append({
                "fecha": order_date,
                "pedido_id": order_id,
                "producto": product,
                "categoria": category,
                "cantidad": random.choices([1, 2, 3, 5, 10], weights=[60, 20, 10, 7, 3])[0],
                "precio_unitario": price,
                "ciudad": city,
            })
        order_number += 1

    # Ordenamos por fecha: así parece un registro real, día a día
    rows.sort(key=lambda r: (r["fecha"], r["pedido_id"]))

    # Renumeramos los pedidos en orden de fecha (el primero de enero es PED-00001),
    # como haría un programa de ventas real
    new_ids = {}
    for row in rows:
        if row["pedido_id"] not in new_ids:
            new_ids[row["pedido_id"]] = f"PED-{len(new_ids) + 1:05d}"
        row["pedido_id"] = new_ids[row["pedido_id"]]
    return rows


def add_dirt(rows: list[dict]) -> list[dict]:
    """Mete suciedad realista: nombres mal escritos, duplicados y filas vacías."""
    dirty = [dict(r) for r in rows]  # copia, para no tocar la lista original

    # 1) Nombres de producto mal escritos
    for row in dirty:
        if random.random() < MESSY_NAME_RATE:
            row["producto"] = mess_up_name(row["producto"])

    # 2) Duplicados: la misma fila copiada justo debajo (típico copiar-pegar)
    for _ in range(N_DUPLICATES):
        pos = random.randrange(len(dirty))
        dirty.insert(pos + 1, dict(dirty[pos]))

    # 3) Filas vacías sueltas por el medio
    empty_row = {col: None for col in dirty[0]}
    for _ in range(N_EMPTY_ROWS):
        pos = random.randrange(1, len(dirty))
        dirty.insert(pos, dict(empty_row))

    return dirty


def save_excel(rows: list[dict], output_path: Path) -> None:
    """Guarda las filas en Excel con columnas legibles."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    df = pd.DataFrame(rows)

    # Escribimos y damos formato en una sola pasada (sin volver a abrir el archivo:
    # con 25.000 filas, reabrirlo costaría varios segundos más)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="ventas")

        # Retoque visual: ancho de columnas y formato de fecha/precio.
        # No cambia los datos, solo cómo se ven al abrir el Excel.
        ws = writer.sheets["ventas"]
        widths = {"A": 12, "B": 12, "C": 32, "D": 14, "E": 10, "F": 16, "G": 12}
        for col, width in widths.items():
            ws.column_dimensions[col].width = width
        for row in ws.iter_rows(min_row=2):
            row[0].number_format = "DD/MM/YYYY"   # fecha
            row[5].number_format = "#,##0.00"     # precio_unitario
        ws.freeze_panes = "A2"  # la cabecera se queda fija al bajar
    # Al salir del "with", pandas guarda el archivo

    make_reproducible(output_path)


def make_reproducible(xlsx_path: Path) -> None:
    """Quita la hora actual de dentro del .xlsx para que salga igual byte a byte.

    Un .xlsx es un ZIP con archivos XML. La hora se cuela en dos sitios:
    - docProps/core.xml guarda "creado" y "modificado" = ahora
    - cada archivo del ZIP lleva su propia hora de guardado
    Si no lo arreglamos, git ve el Excel "cambiado" cada vez que lo regeneramos,
    aunque los datos sean idénticos.
    """
    with zipfile.ZipFile(xlsx_path) as zin:
        entries = [(info.filename, zin.read(info.filename)) for info in zin.infolist()]

    with zipfile.ZipFile(xlsx_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in entries:
            if name == "docProps/core.xml":
                # Cambiamos el contenido de <dcterms:created> y <dcterms:modified>
                text = data.decode("utf-8")
                text = re.sub(
                    r"(<dcterms:(?:created|modified)[^>]*>)[^<]*(<)",
                    rf"\g<1>{FIXED_TIMESTAMP}\g<2>",
                    text,
                )
                data = text.encode("utf-8")
            info = zipfile.ZipInfo(name, date_time=FIXED_ZIP_TIME)
            info.compress_type = zipfile.ZIP_DEFLATED
            zout.writestr(info, data)


def main() -> None:
    parser = argparse.ArgumentParser(description="Genera un Excel de ventas de ejemplo (datos falsos).")
    parser.add_argument(
        "--output",
        default="data/sample_sales.xlsx",
        help="Ruta del Excel a crear (por defecto: data/sample_sales.xlsx)",
    )
    args = parser.parse_args()

    random.seed(SEED)
    clean_rows = build_clean_rows()
    dirty_rows = add_dirt(clean_rows)
    save_excel(dirty_rows, Path(args.output))

    n_orders = len({r["pedido_id"] for r in clean_rows})
    print(f"Excel creado: {args.output}")
    print(f"  Filas totales: {len(dirty_rows)}")
    print(f"  Líneas limpias: {len(clean_rows)} ({n_orders} pedidos)")
    print(f"  Suciedad: {N_DUPLICATES} duplicados, {N_EMPTY_ROWS} filas vacías, "
          f"~{MESSY_NAME_RATE:.0%} nombres mal escritos")


if __name__ == "__main__":
    main()
