# Plan: Generador de reportes Excel → PDF

## Qué
Un script de Python lee un Excel de ventas y saca un PDF de 1 página:
KPIs (total, pedidos, ticket medio), gráfico de ventas por mes,
top 5 productos y ventas por categoría.

## Por qué
Es una pieza de portfolio freelance.
Enseña a un cliente algo que hoy hace a mano: abrir Excel, limpiar, resumir, maquetar.
Con esto: un comando y listo.

## Cómo (ideas clave)
- **Datos falsos con semilla fija**: `generate_data.py` crea ~1.200 filas (ene–jun 2026)
  con suciedad a propósito (filas vacías, duplicados, nombres mal escritos).
  Así el cliente ve que el script limpia datos reales, no solo datos bonitos.
- **Cadena de funciones pequeñas**:
  `load_data -> clean_data -> compute_summary -> build_chart -> build_pdf -> main`.
  Cada una hace una cosa. Fácil de probar y de cambiar.
- **Símbolo €**: fuente TTF libre **DejaVu Sans** dentro del repo (`fonts/`).
  Se ve profesional (“1.234,56 €”) y no depende de las fuentes del PC del cliente.
  Alternativa descartada: escribir “EUR” (funciona, pero queda menos pulido).
- **Números en formato español**: `1.234,56 €` (punto miles, coma decimales).
- **Tests con pytest** solo para `clean_data` y `compute_summary`:
  ahí está la lógica que puede fallar en silencio.
- **Errores claros en español**: archivo no encontrado, faltan columnas.

## Librerías
pandas, openpyxl, matplotlib, fpdf2, pytest. Nada más.

## Ramas
Commit inicial (este plan) en `main`. Luego una rama por tarea:
`funcionalidad/generar-datos`, `funcionalidad/generar-reporte`,
`funcionalidad/tests`, `funcionalidad/readme`. Fusión a `main` al acabar cada una.

## Fuera de alcance (v1)
Email, interfaz gráfica, varios formatos de Excel → “Mejoras futuras” en el README.
