@echo on
setlocal EnableExtensions EnableDelayedExpansion

echo OK#1

rem ==== Ponemos la base del sitio (sin tildes ni puntos raros) ====
set "BASE=C:\Users\dieda\Fiscalia General de la Nacion"
echo OK#2 BASE=%BASE%

rem ==== Detectar la carpeta "Estructurador de Información - Documentos" con comodín ====
set "SITE_DIR="
for /d %%D in ("%BASE%\Estructurador de Informaci* - Documentos") do (
  set "SITE_DIR=%%~fD"
)

echo OK#3 SITE_DIR=!SITE_DIR!

if not defined SITE_DIR (
  echo [ERROR] No se detecto el sitio. Revisa que exista en el Explorador.
  pause
  endlocal
  exit /b 1
)

rem ==== Ver si podemos entrar a EntradaPDF (primero con ruta corta 8.3) ====
set "IN_SP=!SITE_DIR!\EntradaPDF"
set "IN_SP_S="
for %%I in ("!IN_SP!") do set "IN_SP_S=%%~sI"
echo OK#4 IN_SP=!IN_SP!
echo OK#4s IN_SP_S=!IN_SP_S!

if defined IN_SP_S (
  pushd "!IN_SP_S!" || (echo [ERROR] pushd 8.3 fallo & pause & endlocal & exit /b 2)
) else (
  pushd "!IN_SP!"   || (echo [ERROR] pushd normal fallo & pause & endlocal & exit /b 3)
)

echo OK#5 Dentro de EntradaPDF
dir *.pdf
popd

echo OK#6 Probe finalizada correctamente
pause
endlocal