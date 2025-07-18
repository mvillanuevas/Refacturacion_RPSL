@echo off
REM Cambia al directorio donde está el script (opcional pero recomendable)
cd /d "C:\Users\HE678HU\OneDrive - EY\Documents\UiPath\Refacturacion_RPSL\.scripts"

REM Ejecuta el script con Python
python "get_tdc_banxico.py"

if %errorlevel% neq 0 (
    echo Hubo un error al ejecutar el script Python. Código de error: %errorlevel% > "C:\Users\HE678HU\OneDrive - EY\Documents\UiPath\Refacturacion_RPSL\.scripts\output.txt"
)