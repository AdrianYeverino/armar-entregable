# instalar.ps1
#
# Deja la herramienta lista en una maquina nueva. Es idempotente: se puede
# correr las veces que haga falta.
#
#   1. Comprueba pandoc y ofrece instalarlo con winget.
#   2. Comprueba que Microsoft Word este instalado (el pipeline lo automatiza
#      por COM, no hay forma de sustituirlo por LibreOffice sin reescribirlo).
#   3. Crea config.json con el nombre y la matricula, si todavia no existe.
#
# USO
#   .\instalar.ps1
#
# Si PowerShell se niega a correrlo por la politica de ejecucion:
#   powershell -ExecutionPolicy Bypass -File .\instalar.ps1

$ErrorActionPreference = "Stop"
$BASE = Split-Path -Parent $MyInvocation.MyCommand.Path

function Titulo($t) { Write-Host ""; Write-Host $t -ForegroundColor Cyan }
function Ok($t)     { Write-Host "  OK   $t" -ForegroundColor Green }
function Falta($t)  { Write-Host "  FALTA $t" -ForegroundColor Yellow }

Write-Host "Instalacion de armar-entregable" -ForegroundColor White
Write-Host "Carpeta: $BASE"

# ---------- 1. pandoc ---------------------------------------------------------
Titulo "1. pandoc"

$pandoc = "$env:LOCALAPPDATA\Pandoc\pandoc.exe"
if (-not (Test-Path $pandoc)) {
  $c = Get-Command pandoc -ErrorAction SilentlyContinue
  if ($c) { $pandoc = $c.Source } else { $pandoc = $null }
}

if ($pandoc) {
  $v = (& $pandoc --version | Select-Object -First 1)
  Ok "$v"
  Ok "en $pandoc"
}
else {
  Falta "pandoc no esta instalado."
  $r = Read-Host "  Instalarlo ahora con winget? (s/n)"
  if ($r -eq "s") {
    winget install --id JohnMacFarlane.Pandoc -e --accept-package-agreements --accept-source-agreements
    Write-Host "  Cierra y vuelve a abrir PowerShell para que tome el PATH." -ForegroundColor Yellow
  }
  else {
    Write-Host "  Instalalo despues con: winget install --id JohnMacFarlane.Pandoc -e" -ForegroundColor Yellow
  }
}

# ---------- 2. Microsoft Word -------------------------------------------------
Titulo "2. Microsoft Word"

$word = $null
try {
  $word = New-Object -ComObject Word.Application
  Ok ("Word " + $word.Version)
}
catch {
  Falta "no encuentro Microsoft Word por COM."
  Write-Host "  La herramienta lo necesita: Word es quien pone la portada," -ForegroundColor Yellow
  Write-Host "  actualiza el indice y exporta el PDF. LibreOffice no sirve aqui." -ForegroundColor Yellow
}
finally {
  if ($word) {
    $word.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
  }
}

# ---------- 3. config.json ----------------------------------------------------
Titulo "3. config.json (nombre y matricula)"

$cfgFile = Join-Path $BASE "config.json"
if (Test-Path $cfgFile) {
  $j = Get-Content -LiteralPath $cfgFile -Raw -Encoding UTF8 | ConvertFrom-Json
  Ok ("ya existe: " + $j.nombre + " / " + $j.matricula)
  Write-Host "  Para cambiarlo, borra config.json y vuelve a correr esto."
}
else {
  Write-Host "  Estos datos solo salen en el encabezado del perfil 'martinez'."
  Write-Host "  config.json esta en el .gitignore: no se sube al repositorio."
  $nombre = Read-Host "  Nombre completo"
  $matric = Read-Host "  Matricula"
  $obj = [ordered]@{ nombre = $nombre; matricula = $matric }
  # UTF-8 con BOM: sin el, PowerShell 5.1 puede leer mal los acentos.
  $utf8bom = New-Object System.Text.UTF8Encoding($true)
  [System.IO.File]::WriteAllText($cfgFile, ($obj | ConvertTo-Json), $utf8bom)
  Ok "creado config.json"
}

# ---------- 4. comprobacion final --------------------------------------------
Titulo "4. Archivos de la herramienta"

$necesarios = @(
  "armar-entregable.ps1",
  "plantillas\reference-martinez.docx",
  "plantillas\reference-fime.docx",
  "plantillas\reference-formemp.docx",
  "plantillas\portada-fime.docx",
  "plantillas\portada-formemp.docx",
  "plantillas\portada-topicos.docx"
)
$faltan = 0
foreach ($n in $necesarios) {
  if (Test-Path (Join-Path $BASE $n)) { Ok $n } else { Falta $n; $faltan++ }
}

Write-Host ""
if ($faltan -eq 0) {
  Write-Host "Listo. Prueba con:" -ForegroundColor Green
  Write-Host "  .\armar-entregable.ps1 -Md ejemplos\tarea-portada.md -Out prueba.docx -Perfil fime"
}
else {
  Write-Host "Faltan $faltan archivos. Revisa que el clon este completo." -ForegroundColor Red
}
