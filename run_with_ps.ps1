# -*- coding: utf-8 -*-
# Launcher robusto: detecta la carpeta del sitio con comodín, crea unidad real X: con SUBST,
# ejecuta el .BAT y desmonta X:.

$base    = 'C:\Users\dieda\Fiscalia General de la Nacion'
$pattern = 'Estructurador de Informaci* - Documentos'

# 1) Detectar carpeta del sitio (evitamos escribir la tilde literal)
$siteDir = Get-ChildItem -LiteralPath $base -Directory -ErrorAction SilentlyContinue |
           Where-Object { $_.Name -like $pattern } |
           Select-Object -First 1

if (-not $siteDir) {
    Write-Host "[ERROR] No se encontró una carpeta que coincida con: $base\$pattern"
    exit 80
}

$root = $siteDir.FullName
$in   = Join-Path $root 'EntradaPDF'
$out  = Join-Path $root 'SalidaExcel'

if (-not (Test-Path -LiteralPath $in)) {
    Write-Host "[ERROR] Falta la carpeta EntradaPDF en: $root"
    Write-Host "Abre la carpeta en el Explorador y marca 'Conservar siempre en este dispositivo'."
    exit 81
}
if (-not (Test-Path -LiteralPath $out)) {
    New-Item -ItemType Directory -Path $out | Out-Null
}

# 2) Asegurar unidad REAL X: para que CMD la vea (SUBST)
#    - si ya existe, desmontar; luego montar a $root
cmd /c subst X: /D >nul 2>&1
$substCmd = 'subst X: "{0}"' -f $root
$rc = (cmd /c $substCmd).ExitCode
if ($LASTEXITCODE -ne 0) {
    Write-Host "[ERROR] No se pudo montar X: a $root"
    exit 90
}

# 3) Diagnóstico: listar PDFs en X:\EntradaPDF (ahora CMD/PowerShell ven la misma unidad)
Write-Host "[DEBUG] PDFs en X:\EntradaPDF:"
Get-ChildItem -LiteralPath 'X:\EntradaPDF' -Filter *.pdf -File |
    Select-Object -ExpandProperty Name |
    ForEach-Object { Write-Host "  - $_" }
if (-not (Get-ChildItem -LiteralPath 'X:\EntradaPDF' -Filter *.pdf -File)) {
    Write-Host "[ADVERTENCIA] No se encontraron *.pdf en X:\EntradaPDF"
}

# 4) Ejecutar el .BAT principal (debe usar X:\EntradaPDF y X:\SalidaExcel)
$bat = 'C:\Proyecto TICs\Estructurador de verdad\run_sharepoint_bridge.bat'
$proc = Start-Process -FilePath 'cmd.exe' -ArgumentList "/c `"$bat`"" -NoNewWindow -PassThru -Wait

# 5) Desmontar X: (comenta la línea si prefieres dejarla montada)
cmd /c subst X: /D >nul 2>&1

exit ($proc.ExitCode)