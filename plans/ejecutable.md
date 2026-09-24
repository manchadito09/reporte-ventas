# Plan: ejecutable para usuarios sin conocimientos técnicos

## Qué
Un `InformeVentas.exe` que un contable pueda usar sin instalar Python ni abrir la terminal:
arrastra su Excel encima del icono (o hace doble clic y lo elige) y se abre el PDF.

## Por qué
El script de terminal exige Python, entorno virtual y comandos. Para el cliente real
(un contable, una oficina) eso es una barrera. El `.exe` la quita.

## Cómo (ideas clave)
- **`app.py`**: punto de entrada del `.exe`. Recibe el Excel arrastrado o abre el diálogo
  de Windows "Abrir archivo" (tkinter, ya viene con Python). Muestra un aviso de "Generando..."
  mientras trabaja, guarda el PDF junto al Excel (`<nombre>_informe.pdf`) y lo abre.
  Los errores salen en una ventanita, no en una consola.
- **Sin duplicar lógica**: la cadena load → clean → summary → chart → pdf pasa a una función
  `generate_report()` que usan tanto la terminal como el `.exe`.
- **PyInstaller** (solo para fabricar el `.exe`, en `requirements-dev.txt`) con la fuente
  DejaVu dentro. Un script `build_exe.ps1` lo fabrica con un comando.
- **Reparto**: el `.exe` no se sube a git (pesa ~80 MB); va en una Release de GitHub.
- **README**: sección nueva al principio para el usuario no técnico, con el aviso de
  SmartScreen ("Más información → Ejecutar de todas formas").

## Lo que no cambia
El script de terminal y los tests actuales siguen igual.
