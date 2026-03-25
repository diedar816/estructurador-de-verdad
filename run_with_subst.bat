@echo off
setlocal

REM === 1) Montar X: al sitio con tilde (evita problemas de codificación en CMD) ===
REM     Ajusta la ruta EXACTA si cambió:
REM     "C:\Users\dieda\Fiscalia General de la Nacion\Estructurador de Información - Documentos"
subst X: "C:\Users\dieda\Fiscalia General de la Nacion\Estructurador de Información - Documentos"
if errorlevel 1 (
  echo [ERROR] No se pudo montar X:
  pause
  exit /b 90
)

REM Validar que EntradaPDF existe
if not exist "X:\EntradaPDF" (
  echo [ERROR] X:\EntradaPDF no existe o no esta sincronizada
  echo Abre la carpeta en el Explorador y marca "Conservar siempre en este dispositivo".
  subst X: /D
  pause
  exit /b 91
)

REM Mostrar lo que hay (diagnostico)
echo [DEBUG] Contenido de X:\EntradaPDF:
dir /b "X:\EntradaPDF\*.pdf"

REM === 2) Ejecutar tu bridge principal (el que ya funciona) ===
call "C:\Proyecto TICs\Estructurador de verdad\run_sharepoint_bridge.bat"

REM === 3) Desmontar X: (opcional; si prefieres dejarlo, comenta la siguiente linea) ===
subst X: /D

endlocal
``