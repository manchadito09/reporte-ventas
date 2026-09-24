# Fabrica dist\InformeVentas.exe a partir de app.py
#
# Uso (con el entorno virtual activado):
#   pip install -r requirements-dev.txt
#   .\build_exe.ps1

# --onefile   : todo en un único .exe (más fácil de repartir)
# --windowed  : sin la ventana negra de consola
# --add-data  : mete la carpeta fonts dentro del .exe (origen;destino)
# --clean     : borra restos de fabricaciones anteriores
pyinstaller app.py `
    --name InformeVentas `
    --onefile `
    --windowed `
    --add-data "fonts;fonts" `
    --clean `
    --noconfirm

# PyInstaller escribe sus avisos normales por el canal de errores,
# así que para saber si falló miramos su código de salida (0 = todo bien)
if ($LASTEXITCODE -ne 0) {
    Write-Host "ERROR: no se pudo fabricar el .exe (mira los mensajes de arriba)" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Listo: dist\InformeVentas.exe" -ForegroundColor Green
