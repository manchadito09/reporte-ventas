"""
Tests de clean_data y compute_summary.

Usamos datos diminutos hechos a mano: así sabemos la respuesta correcta
de antemano y, si un test falla, se ve enseguida por qué.

Ejecutar:
    pytest
"""

import pandas as pd
import pytest

from generate_report import clean_data, compute_summary


def make_df(rows: list[tuple]) -> pd.DataFrame:
    """Crea un DataFrame con las columnas del Excel a partir de tuplas.

    Orden de cada tupla: fecha, pedido_id, producto, categoria,
    cantidad, precio_unitario, ciudad
    """
    columns = ["fecha", "pedido_id", "producto", "categoria",
               "cantidad", "precio_unitario", "ciudad"]
    return pd.DataFrame(rows, columns=columns)


# Una fila vacía, como las que vienen en el Excel sucio
EMPTY = (None, None, None, None, None, None, None)


# --------------------------------------------------------------------------
# clean_data
# --------------------------------------------------------------------------
def test_clean_removes_empty_rows():
    df = make_df([
        ("2026-01-05", "PED-1", "Ratón", "Informática", 1, 20.0, "Madrid"),
        EMPTY,
        ("2026-01-06", "PED-2", "Teclado", "Informática", 1, 50.0, "Madrid"),
        EMPTY,
    ])

    clean, report = clean_data(df)

    assert len(clean) == 2
    assert report["empty_rows"] == 2


def test_clean_removes_duplicates_but_keeps_lines_of_same_order():
    df = make_df([
        # Pedido con 2 productos distintos: NO es un duplicado
        ("2026-01-05", "PED-1", "Ratón", "Informática", 1, 20.0, "Madrid"),
        ("2026-01-05", "PED-1", "Teclado", "Informática", 1, 50.0, "Madrid"),
        # Copia exacta de la primera fila: SÍ es duplicado
        ("2026-01-05", "PED-1", "Ratón", "Informática", 1, 20.0, "Madrid"),
        # Copia que solo se diferencia en espacios: también es duplicado
        ("2026-01-05", "PED-1", "  Ratón ", "Informática", 1, 20.0, "Madrid"),
    ])

    clean, report = clean_data(df)

    assert len(clean) == 2
    assert report["duplicates"] == 2
    assert sorted(clean["producto"]) == ["Ratón", "Teclado"]


def test_clean_fixes_messy_product_names():
    # La forma bien escrita es la mayoría; las rotas deben unirse a ella
    df = make_df([
        ("2026-01-05", "PED-1", "Pack 500 folios", "Oficina", 1, 5.0, "Madrid"),
        ("2026-01-06", "PED-2", "Pack 500 folios", "Oficina", 1, 5.0, "Madrid"),
        ("2026-01-07", "PED-3", "Pack 500 folios", "Oficina", 1, 5.0, "Madrid"),
        ("2026-01-08", "PED-4", "  PACK 500 FOLIOS ", "Oficina", 1, 5.0, "Madrid"),
        ("2026-01-09", "PED-5", "pack  500 folios", "Oficina", 1, 5.0, "Madrid"),
    ])

    clean, report = clean_data(df)

    assert clean["producto"].unique().tolist() == ["Pack 500 folios"]
    assert report["names_fixed"] == 2


def test_clean_drops_impossible_rows():
    df = make_df([
        ("2026-01-05", "PED-1", "Ratón", "Informática", 1, 20.0, "Madrid"),   # válida
        ("2026-01-05", "PED-2", "Ratón", "Informática", 0, 20.0, "Madrid"),   # cantidad 0
        ("2026-01-05", "PED-3", "Ratón", "Informática", -2, 20.0, "Madrid"),  # cantidad negativa
        ("2026-01-05", "PED-4", "Ratón", "Informática", 1, None, "Madrid"),   # sin precio
        ("no es fecha", "PED-5", "Ratón", "Informática", 1, 20.0, "Madrid"),  # fecha rota
    ])

    clean, report = clean_data(df)

    assert clean["pedido_id"].tolist() == ["PED-1"]
    assert report["invalid_rows"] == 4


def test_clean_reads_text_dates_without_swapping_day_and_month():
    # Fallo real que encontraron los tests: "2026-01-05" se leía como 1 de mayo
    df = make_df([
        ("2026-01-05", "PED-1", "Ratón", "Informática", 1, 20.0, "Madrid"),  # internacional
        ("05/01/2026", "PED-2", "Ratón", "Informática", 1, 20.0, "Madrid"),  # español
        ("2026-01-20", "PED-3", "Ratón", "Informática", 1, 20.0, "Madrid"),
    ])

    clean, _ = clean_data(df)

    # Las tres son de enero de 2026, día 5, 5 y 20
    assert clean["fecha"].dt.month.tolist() == [1, 1, 1]
    assert clean["fecha"].dt.day.tolist() == [5, 5, 20]


def test_clean_computes_line_revenue():
    df = make_df([
        ("2026-01-05", "PED-1", "Ratón", "Informática", 3, 20.5, "Madrid"),
    ])

    clean, _ = clean_data(df)

    assert clean.loc[0, "ingreso"] == pytest.approx(61.5)  # 3 × 20,5


# --------------------------------------------------------------------------
# compute_summary
# --------------------------------------------------------------------------
def test_summary_counts_unique_orders_and_avg_ticket():
    # 4 filas pero solo 2 pedidos: PED-1 tiene 3 productos
    df = make_df([
        ("2026-01-05", "PED-1", "Ratón", "Informática", 1, 20.0, "Madrid"),
        ("2026-01-05", "PED-1", "Teclado", "Informática", 1, 50.0, "Madrid"),
        ("2026-01-05", "PED-1", "Monitor", "Informática", 1, 130.0, "Madrid"),
        ("2026-02-10", "PED-2", "Silla", "Oficina", 1, 100.0, "Sevilla"),
    ])
    clean, _ = clean_data(df)

    summary = compute_summary(clean)

    assert summary["total_sales"] == pytest.approx(300.0)
    assert summary["n_orders"] == 2                         # no 4
    assert summary["avg_ticket"] == pytest.approx(150.0)    # 300 / 2, no 300 / 4


def test_summary_top_products_sorted_and_limited_to_five():
    # 6 productos con ingresos distintos: el más barato debe quedar fuera
    df = make_df([
        ("2026-01-05", f"PED-{i}", f"Producto {i}", "Oficina", 1, float(i * 10), "Madrid")
        for i in range(1, 7)
    ])
    clean, _ = clean_data(df)

    top = compute_summary(clean)["top_products"]

    assert len(top) == 5
    assert top["producto"].tolist() == [
        "Producto 6", "Producto 5", "Producto 4", "Producto 3", "Producto 2",
    ]


def test_summary_monthly_and_category_percentages():
    df = make_df([
        ("2026-01-05", "PED-1", "Ratón", "Informática", 1, 30.0, "Madrid"),
        ("2026-01-20", "PED-2", "Silla", "Oficina", 1, 70.0, "Madrid"),
        ("2026-02-03", "PED-3", "Ratón", "Informática", 1, 100.0, "Madrid"),
    ])
    clean, _ = clean_data(df)

    summary = compute_summary(clean)

    # Ventas por mes: enero = 30 + 70, febrero = 100
    assert summary["monthly"].tolist() == pytest.approx([100.0, 100.0])

    # Porcentajes por categoría: Informática 130/200, Oficina 70/200
    pct = dict(zip(summary["by_category"]["categoria"], summary["by_category"]["porcentaje"]))
    assert pct == pytest.approx({"Informática": 65.0, "Oficina": 35.0})
