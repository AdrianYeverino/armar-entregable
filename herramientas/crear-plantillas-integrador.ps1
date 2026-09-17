# crear-plantillas-integrador.ps1
#
# Genera las dos plantillas del perfil "integrador" a partir de un entregable
# real del equipo 5 de Proyecto Integrador I (hecho en Google Docs):
#
#   plantillas\portada-integrador.docx    portada + indice + seccion del cuerpo
#   plantillas\reference-integrador.docx  estilos que usa pandoc para el cuerpo
#
# Se corre una sola vez, o cuando cambie la portada del equipo:
#   .\herramientas\crear-plantillas-integrador.ps1 -Origen "...\Fase 4.docx"
#
# Que hace con la portada:
#   - cambia el subtitulo subrayado ("Fase IV. ...") por {{ACTIVIDAD}}
#   - cambia la fecha subrayada por {{FECHA}}
#   - deja la seccion 2 (la que numera paginas) con un solo parrafo {{TITULO}}
#     en estilo Titulo, que es el "Fase N" grande con el que abre el cuerpo
#   - redefine los estilos del cuerpo igual que en la referencia, porque al
#     insertar el cuerpo Word se queda con los estilos del documento destino

param(
  [Parameter(Mandatory=$true)][string]$Origen
)
$ErrorActionPreference = "Stop"
$BASE = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$portada = Join-Path $BASE "plantillas\portada-integrador.docx"
$refer   = Join-Path $BASE "plantillas\reference-integrador.docx"

$pandoc = "$env:LOCALAPPDATA\Pandoc\pandoc.exe"
if (-not (Test-Path $pandoc)) { $pandoc = (Get-Command pandoc).Source }

# colores de Word: R + G*256 + B*65536
$AZUL_TITULO = 0x0f + 0x47*256 + 0x61*65536   # #0f4761, titulos del equipo

function Ajustar-Estilos($doc) {
  # Los estilos de Word se piden por su constante (Normal -1, Titulo 1 -2...),
  # porque en un Word en espanol el nombre "Heading 1" no existe.
  function E($nombre) { try { return $doc.Styles.Item($nombre) } catch { return $null } }

  # Normal NO se toca: la tabla de integrantes de la portada lo hereda, y con
  # interlineado 1.5 la portada se desborda a una segunda hoja.
  foreach ($n in @(-67, "First Paragraph", "Compact")) {
    $s = E $n
    if (-not $s) { continue }
    $s.Font.Name = "Arial"; $s.Font.Size = 12; $s.Font.Color = 0
    $s.ParagraphFormat.Alignment = 3                 # justificado
    $s.ParagraphFormat.LineSpacingRule = 1           # 1.5 lineas
    $s.ParagraphFormat.SpaceBefore = 0
    $s.ParagraphFormat.SpaceAfter = $(if ($n -eq "Compact") { 6 } else { 10 })
  }
  $t = @(
    @(-2, "Play",  20, 18, 4),
    @(-3, "Play",  16, 8,  4),
    @(-4, "Aptos", 14, 8,  4)
  )
  foreach ($x in $t) {
    $s = E $x[0]; if (-not $s) { continue }
    $s.Font.Name = $x[1]; $s.Font.Size = $x[2]; $s.Font.Bold = $false; $s.Font.Italic = $false
    $s.Font.Color = $AZUL_TITULO
    $s.ParagraphFormat.Alignment = 0
    $s.ParagraphFormat.LineSpacingRule = 0
    $s.ParagraphFormat.SpaceBefore = $x[3]; $s.ParagraphFormat.SpaceAfter = $x[4]
    $s.ParagraphFormat.KeepWithNext = $true
  }
  $s = E -63   # wdStyleTitle
  if ($s) {
    $s.Font.Name = "Play"; $s.Font.Size = 28; $s.Font.Bold = $false; $s.Font.Color = 0
    $s.ParagraphFormat.Alignment = 0; $s.ParagraphFormat.SpaceAfter = 4
    $s.ParagraphFormat.LineSpacingRule = 0
    try { $s.ParagraphFormat.Borders.Enable = $false } catch { }
  }
  # indice: primer nivel en negrita, como el que genera Google Docs
  foreach ($x in @(@(-20, $true), @(-21, $false), @(-22, $false))) {   # wdStyleTOC1..3
    $s = E $x[0]; if (-not $s) { continue }
    $s.Font.Name = "Arial"; $s.Font.Size = 11; $s.Font.Bold = $x[1]; $s.Font.Color = 0
  }
  $s = E "Image Caption"
  if ($s) {
    $s.Font.Name = "Arial"; $s.Font.Size = 10; $s.Font.Italic = $true; $s.Font.Color = 0x404040
    $s.ParagraphFormat.Alignment = 1; $s.ParagraphFormat.SpaceAfter = 10
  }
  $s = E "Captioned Figure"
  if ($s) { $s.ParagraphFormat.Alignment = 1; $s.ParagraphFormat.SpaceAfter = 2 }
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  # ---------- referencia: la de pandoc con los estilos del equipo ----------
  # --print-default-data-file escribe a stdout; cmd lo redirige en binario
  # (el > de PowerShell lo pasaria a UTF-16 y dejaria el zip inservible)
  cmd /c "`"$pandoc`" --print-default-data-file reference.docx > `"$refer`""
  $doc = $word.Documents.Open($refer)
  Ajustar-Estilos $doc
  $doc.Save(); $doc.Close()
  Write-Host "OK -> $refer"

  # ---------- portada: el entregable real recortado ----------
  Copy-Item $Origen $portada -Force
  $doc = $word.Documents.Open($portada)
  if ($doc.Sections.Count -lt 2) { throw "El origen no tiene seccion 2 (cuerpo)." }

  # cuerpo: borrar todo y dejar un parrafo {{TITULO}} en estilo Titulo
  $s2 = $doc.Sections.Item(2).Range
  $r = $doc.Range($s2.Start, $doc.Content.End - 1)
  $r.Delete() | Out-Null
  $r = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
  $r.InsertAfter("{{TITULO}}")
  $p = $doc.Paragraphs.Item($doc.Paragraphs.Count)
  $p.Style = -63   # wdStyleTitle: por numero, el nombre cambia con el idioma de Word
  $p.Range.ParagraphFormat.TabStops.ClearAll()

  # subtitulo y fecha de la portada
  function Reemplazar($de, $a) {
    $f = $doc.Content.Find
    $f.ClearFormatting(); $f.Replacement.ClearFormatting()
    $ok = $f.Execute([ref]$de,[ref]$false,[ref]$false,[ref]$false,[ref]$false,[ref]$false,[ref]$true,[ref]0,[ref]$false,[ref]$a,[ref]1)
    if (-not $ok) { Write-Warning "no encontre '$de' en la portada" }
  }
  foreach ($p in 1..$doc.Paragraphs.Count) {
    $par = $doc.Paragraphs.Item($p)
    if ($par.Range.Tables.Count -gt 0) { continue }
    $txt = $par.Range.Text
    if ($txt -match '^\s*Fase\s+[IVX]+\.\s') {
      $m = [regex]::Match($txt, 'Fase\s+[IVX]+\.[^\r]*').Value.TrimEnd()
      Reemplazar $m "{{ACTIVIDAD}}"
    }
    if ($txt -match 'N\.L\., a\s+(.+?)\s*\r') {
      Reemplazar $Matches[1].Trim() "{{FECHA}}"
    }
  }
  Ajustar-Estilos $doc

  # El campo TOC de Google Docs filtra por nombre de estilo en ingles
  # (\t "Heading 1,1,..."). En un Word en espanol esos nombres no existen y el
  # indice sale vacio; con \o toma los niveles de esquema, en cualquier idioma.
  foreach ($fi in 1..$doc.Fields.Count) {
    $campo = $doc.Fields.Item($fi)
    if ($campo.Code.Text -match '^\s*TOC\b') { $campo.Code.Text = ' TOC \o "1-3" \h \z \u ' }
  }
  $doc.EmbedTrueTypeFonts = $true
  $doc.SaveSubsetFonts = $false
  $doc.Save()
  Write-Host ("  marcadores: " + (([regex]::Matches($doc.Content.Text, '\{\{\w+\}\}') | ForEach-Object Value) -join ", "))
  $doc.Close()
  Write-Host "OK -> $portada"
}
finally {
  try { $word.Quit() } catch { }
}
