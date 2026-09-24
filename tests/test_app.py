"""Tests del punto de entrada del ejecutable (app.py)."""

from pathlib import Path

from app import default_output_path


def test_output_pdf_goes_next_to_excel_with_suffix():
    # El PDF va en la misma carpeta que el Excel, aunque el nombre tenga espacios
    excel = Path("C:/Oficina/Ventas 2025/ventas enero.xlsx")

    assert default_output_path(excel) == Path("C:/Oficina/Ventas 2025/ventas enero_informe.pdf")
