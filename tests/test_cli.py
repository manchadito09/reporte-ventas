"""Tests de principio a fin: ejecutan el programa como lo haría una persona.

Los otros tests prueban piezas sueltas. Estos comprueban que las piezas
encajan: al unir dos ramas, main() llegó a usar una variable que ya no existía
y ningún test de piezas lo detectó.
"""

import sys

import pandas as pd

import generate_report


def write_excel(path, price_for_row):
    """Crea un Excel pequeño de 20 filas; price_for_row(i) decide cada precio."""
    rows = [
        {"fecha": "2025-01-15", "pedido_id": f"P{i}", "producto": "Silla",
         "categoria": "Oficina", "cantidad": 1, "precio_unitario": price_for_row(i),
         "ciudad": "Madrid"}
        for i in range(20)
    ]
    pd.DataFrame(rows).to_excel(path, index=False)


def run_cli(monkeypatch, *args) -> int:
    """Ejecuta main() como si escribiéramos: python generate_report.py <args>."""
    monkeypatch.setattr(sys, "argv", ["generate_report.py", *map(str, args)])
    return generate_report.main()


def test_cli_creates_pdf(tmp_path, monkeypatch, capsys):
    excel = tmp_path / "ventas.xlsx"
    pdf = tmp_path / "informe.pdf"
    write_excel(excel, lambda i: 10.0)

    code = run_cli(monkeypatch, excel, "--output", pdf)

    assert code == 0
    assert pdf.exists() and pdf.stat().st_size > 0
    assert "Revisa tu Excel" not in capsys.readouterr().out


def test_cli_warns_when_many_rows_are_unreadable(tmp_path, monkeypatch, capsys):
    excel = tmp_path / "ventas.xlsx"
    write_excel(excel, lambda i: "abc" if i < 4 else 10.0)   # 4 de 20 = 20 %

    code = run_cli(monkeypatch, excel, "--output", tmp_path / "informe.pdf")

    assert code == 0
    assert "Revisa tu Excel" in capsys.readouterr().out


def test_cli_missing_file_returns_error(tmp_path, monkeypatch, capsys):
    code = run_cli(monkeypatch, tmp_path / "no_existe.xlsx")

    assert code == 1
    assert "No se encuentra el archivo" in capsys.readouterr().err
