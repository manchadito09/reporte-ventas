"""
Tests de clean_data y compute_summary.

Usamos datos diminutos hechos a mano: así sabemos la respuesta correcta
de antemano y, si un test falla, se ve enseguida por qué.

Ejecutar:
    pytest
"""

import pandas as pd
import pytest

from generate_report import (
    ReportPDF, clean_data, compute_summary, fit_text, fold_small_rows,
    format_quantity, month_labels, quality_warning,
)


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
        ("2026-01-05", "PED-3", "Ratón", "Informática", 1, -5.0, "Madrid"),  # precio negativo
        ("2026-01-05", "PED-4", "Ratón", "Informática", 1, None, "Madrid"),   # sin precio
        ("no es fecha", "PED-5", "Ratón", "Informática", 1, 20.0, "Madrid"),  # fecha rota
        ("2026-01-05", "PED-6", "Ratón", "Informática", 1, "abc", "Madrid"),  # precio ilegible
    ])

    clean, report = clean_data(df)

    assert clean["pedido_id"].tolist() == ["PED-1"]
    assert report["invalid_rows"] == 5


def test_clean_reads_prices_written_as_spanish_text():
    # Precios tecleados a mano en una oficina española
    df = make_df([
        ("2025-01-05", "PED-1", "A", "X", 1, "22,90", "Madrid"),
        ("2025-01-05", "PED-2", "B", "X", 1, "22,90 €", "Madrid"),
        ("2025-01-05", "PED-3", "C", "X", 1, "1.234,56", "Madrid"),
        ("2025-01-05", "PED-4", "D", "X", 1, "1.234", "Madrid"),
        ("2025-01-05", "PED-5", "E", "X", 1, 22.9, "Madrid"),
        ("2025-01-05", "PED-6", "F", "X", "3", "10,00", "Madrid"),     # cantidad en texto
    ])

    clean, report = clean_data(df)

    assert report["invalid_rows"] == 0
    assert clean["precio_unitario"].tolist() == pytest.approx([22.9, 22.9, 1234.56, 1234.0, 22.9, 10.0])
    assert clean["cantidad"].tolist() == [1, 1, 1, 1, 1, 3]


def test_returns_are_kept_and_subtracted_from_total():
    # Opción elegida: las devoluciones (cantidad negativa) restan del total
    df = make_df([
        ("2025-01-05", "PED-1", "Silla", "Oficina", 3, 100.0, "Madrid"),
        ("2025-01-20", "PED-2", "Silla", "Oficina", -1, 100.0, "Madrid"),   # devuelve 1
    ])
    clean, report = clean_data(df)

    summary = compute_summary(clean)

    assert report["returns"] == 1
    assert report["invalid_rows"] == 0
    assert summary["total_sales"] == pytest.approx(200.0)           # 300 - 100
    assert summary["top_products"].loc[0, "unidades"] == 2           # 3 - 1


def test_quality_warning_only_when_many_rows_are_discarded():
    # 1 de 10 filas ilegible (10 %) -> avisa; 0 de 10 -> no avisa
    good = [("2025-01-05", f"PED-{i}", "A", "X", 1, 10.0, "Madrid") for i in range(9)]
    bad = [("2025-01-05", "PED-X", "A", "X", 1, "abc", "Madrid")]

    _, report_bad = clean_data(make_df(good + bad))
    _, report_ok = clean_data(make_df(good + [("2025-01-05", "PED-Y", "A", "X", 1, 10.0, "Madrid")]))

    assert report_bad["invalid_share"] == pytest.approx(0.10)
    assert "Revisa tu Excel" in quality_warning(report_bad)
    assert quality_warning(report_ok) is None


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


def test_clean_keeps_decimal_quantities():
    # Venta por kilos o metros: 2,5 × 10 € = 25 €. Antes se cortaba a 2 (20 €) sin avisar
    df = make_df([
        ("2025-01-05", "PED-1", "Cable", "Ferretería", "2,5", 10.0, "Madrid"),
        ("2025-01-05", "PED-2", "Queso", "Alimentación", 0.75, 20.0, "Madrid"),
    ])

    clean, report = clean_data(df)

    assert report["invalid_rows"] == 0
    assert clean["cantidad"].tolist() == pytest.approx([2.5, 0.75])
    assert clean["ingreso"].tolist() == pytest.approx([25.0, 15.0])


def test_quantity_shows_decimals_only_when_needed():
    assert format_quantity(147.0) == "147"
    assert format_quantity(1255100.0) == "1.255.100"
    assert format_quantity(2.5) == "2,5"
    assert format_quantity(0.75) == "0,75"


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


def test_summary_full_year_months_in_order():
    # Un pedido por mes de 2025, metidos en el Excel desordenados
    months = [7, 12, 1, 5, 9, 3, 11, 2, 8, 4, 10, 6]
    df = make_df([
        (f"2025-{m:02d}-15", f"PED-{m}", "Ratón", "Informática", 1, float(m), "Madrid")
        for m in months
    ])
    clean, _ = clean_data(df)

    monthly = compute_summary(clean)["monthly"]

    # Salen los 12 meses, de enero a diciembre, cada uno con su importe
    assert [p.month for p in monthly.index] == list(range(1, 13))
    assert monthly.tolist() == pytest.approx([float(m) for m in range(1, 13)])


# --------------------------------------------------------------------------
# Presentación del PDF
# --------------------------------------------------------------------------
def test_many_categories_fold_into_others():
    # 15 categorías -> 6 más grandes + "Otras (9)" = 7 filas, sin perder dinero
    table = pd.DataFrame({
        "categoria": [f"Cat {i}" for i in range(15)],
        "ingresos": [float(100 - i) for i in range(15)],
    })
    table["porcentaje"] = table["ingresos"] / table["ingresos"].sum() * 100

    folded = fold_small_rows(table, "categoria", max_rows=7)

    assert len(folded) == 7
    assert folded["categoria"].iloc[-1] == "Otras (9)"
    assert folded["ingresos"].sum() == pytest.approx(table["ingresos"].sum())
    assert folded["porcentaje"].sum() == pytest.approx(100.0)


def test_few_categories_are_not_folded():
    table = pd.DataFrame({"categoria": ["A", "B"], "ingresos": [2.0, 1.0], "porcentaje": [66.7, 33.3]})

    assert fold_small_rows(table, "categoria", max_rows=7).equals(table)


def test_month_labels_show_year_only_when_period_crosses_years():
    one_year = pd.period_range("2025-01", "2025-03", freq="M")
    two_years = pd.period_range("2024-11", "2025-02", freq="M")

    assert month_labels(one_year) == ["Ene", "Feb", "Mar"]
    assert month_labels(two_years) == ["Nov 24", "Dic 24", "Ene 25", "Feb 25"]


def test_long_text_is_cut_with_ellipsis_to_fit():
    pdf = ReportPDF("prueba.xlsx")
    pdf.set_font("DejaVu", "", 8.5)
    long_name = "Televisor OLED 77 pulgadas 4K 120 Hz con HDR, Dolby Vision y barra de sonido"

    cut = fit_text(pdf, long_name, max_width=40)

    assert cut.endswith("…")
    assert pdf.get_string_width(cut) <= 40
    assert fit_text(pdf, "Ratón", max_width=40) == "Ratón"   # lo corto no se toca
