@echo off
setlocal EnableExtensions EnableDelayedExpansion

rem =====================================================
rem = RUTAS DEL PROYECTO (NO CAMBIAN)
rem =====================================================
set "PROY=C:\Proyecto TICs\Estructurador de verdad"
set "PY=%PROY%\estructurador.py"
set "IN_LOCAL=%PROY%\Estructurador de Información.pdf"
set "OUT_LOCAL=%PROY%\Estructurado en tabla.xlsx"

rem =====================================================
rem = RUTAS DEL SITIO (VÍA UNIDAD REAL X:, MONTADA POR run_with_ps.ps1)
rem =====================================================
set "IN_SP_LONG=X:\EntradaPDF"
set "OUT_SP_LONG=X:\SalidaExcel"

if not exist "%IN_SP_LONG%" (
  echo [ERROR] X:\EntradaPDF no existe. Ejecuta primero el launcher run_with_ps.ps1
  pause
  exit /b 91
)
if not exist "%OUT_SP_LONG%" mkdir "%OUT_SP_LONG%"

rem =====================================================
rem = 1) ELEGIR EL PDF A PROCESAR
rem =    - Si existe el nombre fijo, usarlo
rem =    - Si no, tomar el mas reciente (*.pdf)
rem =====================================================
set "PICKED="
if exist "%IN_SP_LONG%\Estructurador de Información.pdf" (
  set "PICKED=Estructurador de Información.pdf"
) else (
  for /f "delims=" %%F in ('dir /b /a:-d /o:-d "%IN_SP_LONG%\*.pdf" 2^>nul') do (
    set "PICKED=%%~nxF"
    goto :HAVE_PDF
  )
)

:HAVE_PDF
if not defined PICKED (
  echo [ERROR] No hay PDF en "%IN_SP_LONG%"
  pause
  exit /b 1
)

echo [DEBUG] EntradaPDF contiene:
dir /b "%IN_SP_LONG%\*.pdf"

copy /Y "%IN_SP_LONG%\%PICKED%" "%IN_LOCAL%" >nul
echo [OK] PDF copiado a trabajo: %PICKED%

rem (OPCIONAL) Mover a Procesados para evitar reprocesar:
rem if not exist "%IN_SP_LONG%\Procesados" mkdir "%IN_SP_LONG%\Procesados"
rem move /Y "%IN_SP_LONG%\%PICKED%" "%IN_SP_LONG%\Procesados" >nul

rem =====================================================
rem = 2) EJECUTAR PYTHON COMO SIEMPRE
rem =====================================================
pushd "%PROY%"
call "%PROY%\.venv\Scripts\activate.bat" 2>nul
python "%PY%" -i "%IN_LOCAL%" -o "%OUT_LOCAL%"
set "RC=%ERRORLEVEL%"
popd

if not "%RC%"=="0" (
  echo [ERROR] Python RC=%RC%
  pause
  exit /b %RC%
)

if not exist "%OUT_LOCAL%" (
  echo [ERROR] No se generó salida: %OUT_LOCAL%
  pause
  exit /b 2
)

rem =====================================================
rem = 3) TIMESTAMP SEGURO (NO USAR %TIME%)
rem =====================================================
for /f %%T in ('powershell -NoProfile -Command "(Get-Date).ToString(\"yyyyMMdd_HHmmss\")"') do set "TS=%%T"

rem =====================================================
rem = 4) COPIAR SALIDA (FIJO + CON TIMESTAMP)
rem =====================================================
copy /Y "%OUT_LOCAL%" "%OUT_SP_LONG%\Estructurado en tabla.xlsx" >nul
copy /Y "%OUT_LOCAL%" "%OUT_SP_LONG%\Estructurado_en_tabla_%TS%.xlsx" >nul
echo [OK] Entregado en: "%OUT_SP_LONG%"

rem (OPCIONAL) Candado anti-doble clic:
rem set "LOCK=%TEMP%\estructurador.lock"
rem if exist "%LOCK%" (echo Ya hay un proceso en curso & exit /b 99)
rem echo 1>"%LOCK%"
rem ... (tu proceso)
rem del "%LOCK%" >nul 2>&1

echo.
echo PROCESO COMPLETO. Presione una tecla para salir...
pause
endlocal
