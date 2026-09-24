"""
Punto de entrada del ejecutable InformeVentas.exe (para usuarios sin terminal).

Formas de uso:
- Arrastrar un Excel encima de InformeVentas.exe
- Hacer doble clic en InformeVentas.exe y elegir el Excel en la ventana

El PDF se guarda junto al Excel (<nombre>_informe.pdf) y se abre solo.
Los errores se muestran en una ventana, no en una consola.

También se puede probar sin fabricar el .exe:
    python app.py
    python app.py data/sample_sales.xlsx
"""

import os
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from generate_report import ReportError, generate_report, quality_warning

APP_TITLE = "Informe de ventas"


def default_output_path(input_path: Path) -> Path:
    """ventas.xlsx -> ventas_informe.pdf, en la misma carpeta que el Excel."""
    return input_path.with_name(f"{input_path.stem}_informe.pdf")


def ask_for_excel(root: tk.Tk) -> Path | None:
    """Abre la ventana normal de Windows para elegir un Excel. None si cancela."""
    filename = filedialog.askopenfilename(
        parent=root,
        title="Elige el Excel de ventas",
        filetypes=[("Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
    )
    return Path(filename) if filename else None


def run_with_progress(root: tk.Tk, input_path: Path, work) -> dict:
    """Ejecuta work() en un hilo aparte mientras enseña una ventanita de espera.

    Hilo aparte = la ventana sigue respondiendo mientras se calcula el informe.
    Si lo hiciéramos todo en el mismo hilo, Windows diría "No responde".
    Devuelve {"value": resultado} o {"error": excepción}.
    """
    # Ventanita de espera
    root.deiconify()
    root.title(APP_TITLE)
    root.resizable(False, False)
    frame = ttk.Frame(root, padding=20)
    frame.pack()
    ttk.Label(frame, text="Generando el informe de:").pack(anchor="w")
    ttk.Label(frame, text=input_path.name, font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(2, 10))
    bar = ttk.Progressbar(frame, mode="indeterminate", length=280)  # barra que va y viene
    bar.pack()
    bar.start(12)
    ttk.Label(frame, text="Tarda unos segundos, no cierres esta ventana.").pack(anchor="w", pady=(10, 0))

    # Centrar la ventana en la pantalla
    root.update_idletasks()
    x = (root.winfo_screenwidth() - root.winfo_width()) // 2
    y = (root.winfo_screenheight() - root.winfo_height()) // 3
    root.geometry(f"+{x}+{y}")

    result = {}

    def target():
        try:
            result["value"] = work()
        except Exception as exc:  # cualquier fallo se enseña luego en una ventana
            result["error"] = exc

    worker = threading.Thread(target=target, daemon=True)
    worker.start()

    # Cada 100 ms miramos si el trabajador ha terminado; si sí, cerramos el bucle
    def check():
        if worker.is_alive():
            root.after(100, check)
        else:
            root.quit()

    root.after(100, check)
    root.mainloop()
    root.withdraw()  # escondemos la ventanita de espera
    return result


def main() -> int:
    root = tk.Tk()
    root.withdraw()  # empezamos con la ventana escondida

    # 1) ¿Qué Excel? El arrastrado encima del .exe o el elegido en la ventana
    if len(sys.argv) > 1:
        input_path = Path(sys.argv[1])
    else:
        input_path = ask_for_excel(root)
        if input_path is None:
            return 0  # ha cancelado: salimos sin decir nada

    output_path = default_output_path(input_path)

    # 2) Generar el informe con la ventanita de espera
    result = run_with_progress(root, input_path, lambda: generate_report(input_path, output_path))

    # 3) Resultado
    error = result.get("error")
    if isinstance(error, ReportError):
        messagebox.showerror(APP_TITLE, str(error), parent=root)
        return 1
    if error is not None:
        messagebox.showerror(
            APP_TITLE,
            "Ha ocurrido un error inesperado al generar el informe.\n\n"
            f"Detalle técnico: {type(error).__name__}: {error}",
            parent=root,
        )
        return 1

    # Si se descartaron muchas filas, avisamos antes de abrir el PDF:
    # el contable no ve la terminal y el total podría estar incompleto
    _, cleaning = result["value"]
    warning = quality_warning(cleaning)
    if warning:
        messagebox.showwarning(
            APP_TITLE,
            warning.lstrip("⚠ ") + "\n\nEl informe se ha creado igualmente.",
            parent=root,
        )

    os.startfile(output_path)  # abre el PDF con el programa por defecto (solo Windows)
    return 0


if __name__ == "__main__":
    sys.exit(main())
