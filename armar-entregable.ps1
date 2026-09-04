# armar-entregable.ps1
#
# Pipeline: tarea.md  ->  .docx con formato NATIVO  ->  .pdf
#
#   1. pandoc convierte el cuerpo usando un .docx de referencia con TUS estilos.
#      Los parrafos quedan con estilos con nombre y HEREDAN el espaciado,
#      igual que un documento hecho a mano en Word.
#   2. Word (COM) abre la plantilla de portada, rellena los marcadores {{...}}
#      y le inserta el cuerpo. Word hace el trabajo, asi que el resultado es nativo.
#   3. Word actualiza el indice automatico y exporta el PDF.
#
# Ver "00 - Sistema/Convencion - Generar el Word desde Obsidian.md" en el vault.
#
# USO
#   .\armar-entregable.ps1 -Md tarea.md -Out Actividad1.docx
#   .\armar-entregable.ps1 -Md tarea.md -Out AF1.docx -Perfil fime
#   .\armar-entregable.ps1 -Md "nota del vault.md" -Out "EQ#3_ACT1_FORMEMP_034.docx" `
#                          -Perfil formemp -Seccion "Documento entregable"
#
# PERFILES
#   martinez  sin portada, con encabezado (tarea / nombre / matricula).
#             Estilos: Arial 12 justificado 1.15, titulos 16/14/12.
#   fime      portada estandar de la facultad. Ver "Formato - Portada UANL FIME".
#   topicos   portada del equipo en Topicos Selectos de Ciencias de la Ingenieria I.
#             Renglones: materia, actividad subrayada, Instructor, Semestre,
#             Grupo | Hora | Frecuencia, equipo y tabla Estudiante/Matricula/Carrera.
#             La fecha va centrada y en minusculas, precedida de "Ciudad
#             Universitaria, ". Usa -BordesTabla.
#   formemp   portada del M.A. Guillermo Marin Rangel, con indice automatico.
#             Estilos: Arial 12 justificado 1.5, titulos 18/16/14.
#             Ver "Formato - M.A. Guillermo Marin Rangel".
#
# EQUIPO O INDIVIDUAL
#   Lo decide el frontmatter, no un parametro: si trae "equipo:" se pone el
#   renglon del equipo; si no lo trae, se omite y el hueco se cierra solo.
#
# FRONTMATTER
#   Los datos de la portada van en un bloque anidado "portada:", para que una
#   nota del vault pueda llevar su propio frontmatter (materia, tipo, estado)
#   sin que choque con los campos del entregable. Los campos de "portada:"
#   tienen precedencia sobre los de primer nivel.
#
#   ---
#   materia: Formacion De Emprendedores      <- del vault, lo usa Dataview
#   tipo: tarea
#   portada:
#     materia: Formación de Emprendedores    <- el que va en la portada
#     actividad: Actividad Fundamental 1
#     tema: "“FODA”"
#     equipo: 3
#     docente: M.A. Guillermo Marín Rangel
#     grupo: 034 / E2026
#     semestre: AGOSTO – DICIEMBRE 2026
#     plan: 401
#     modalidad:
#     ciudad: San Nicolás de los Garza, N.L
#     fecha: 24 de agosto del 2026
#     integrantes:
#       - 1234567 | Nombre Apellido Apellido | ITS
#   ---
#
#   Tambien acepta el formato viejo, con los campos de portada en el primer
#   nivel del frontmatter.
#
# -Seccion
#   Toma solo lo que esta debajo de ese encabezado, hasta el siguiente
#   encabezado del mismo nivel o mas alto. Sirve para que una nota de Obsidian
#   sea a la vez nota (enunciado, dudas, notas de diseño) y fuente del Word.
#   Los encabezados de dentro se suben de nivel para que el primero quede
#   como Titulo 1: si la seccion es "## Documento entregable", entonces
#   ### -> Titulo 1, #### -> Titulo 2, ##### -> Titulo 3.

param(
  [Parameter(Mandatory=$true)][string]$Md,
  [Parameter(Mandatory=$true)][string]$Out,
  [ValidateSet("martinez","fime","formemp","topicos","lbtssi")][string]$Perfil = "",
  [string]$Seccion = "",
  [switch]$ConPortada,
  [string]$Encabezado = "",
  [string]$Referencia = "",
  [string]$Plantilla  = "",
  [int]$SubirNivel = -1,
  [switch]$SinPdf,
  [switch]$BordesTabla,
  [switch]$TablaCompacta
)

$ErrorActionPreference = "Stop"
$BASE   = Split-Path -Parent $MyInvocation.MyCommand.Path

# ---------- 0a. datos del autor ----------------------------------------------
# Viven en config.json, junto al script, NO en el codigo: el repositorio es
# publico y la matricula no tiene por que quedar en el. config.json esta en el
# .gitignore; config.ejemplo.json es la plantilla que si se versiona.
# Solo los usa el perfil "martinez", que los imprime en el encabezado.
$NOMBRE = ""
$MATRIC = ""
$cfgFile = Join-Path $BASE "config.json"
if (Test-Path $cfgFile) {
  $json = Get-Content -LiteralPath $cfgFile -Raw -Encoding UTF8 | ConvertFrom-Json
  if ($json.nombre)    { $NOMBRE = $json.nombre }
  if ($json.matricula) { $MATRIC = $json.matricula }
}

# ---------- 0. perfil ---------------------------------------------------------
# -ConPortada se conserva por compatibilidad: equivale a -Perfil fime.
if (-not $Perfil) { $Perfil = if ($ConPortada) { "fime" } else { "martinez" } }

$PERFILES = @{
  "martinez" = @{ portada = $false; ref = "reference-martinez.docx"; plantilla = ""                    }
  "fime"     = @{ portada = $true;  ref = "reference-fime.docx";     plantilla = "portada-fime.docx"    }
  "formemp"  = @{ portada = $true;  ref = "reference-formemp.docx";  plantilla = "portada-formemp.docx" }
  "topicos"  = @{ portada = $true;  ref = "reference-martinez.docx"; plantilla = "portada-topicos.docx" }
  "lbtssi"   = @{ portada = $true;  ref = "reference-fime.docx";     plantilla = "portada-lbtssi.docx"  }
}
$cfg = $PERFILES[$Perfil]

if (-not $Referencia) { $Referencia = Join-Path $BASE ("plantillas\" + $cfg.ref) }
if (-not $Plantilla -and $cfg.plantilla) { $Plantilla = Join-Path $BASE ("plantillas\" + $cfg.plantilla) }

$pandoc = "$env:LOCALAPPDATA\Pandoc\pandoc.exe"
if (-not (Test-Path $pandoc)) {
  $c = Get-Command pandoc -ErrorAction SilentlyContinue
  if ($c) { $pandoc = $c.Source } else { throw "No encuentro pandoc.exe" }
}
if (-not (Test-Path $Md))         { throw "No encuentro el markdown: $Md" }
if (-not (Test-Path $Referencia)) { throw "No encuentro el reference.docx: $Referencia" }

$Md  = (Resolve-Path $Md).Path
# Ojo: GetFullPath resuelve las rutas relativas contra el directorio actual DEL
# PROCESO, que no es el mismo que la ubicacion de PowerShell. Sin combinar con
# $PWD, un -Out relativo escribe en donde arranco el proceso (paso una vez: el
# entregable termino en la raiz del vault en lugar de la carpeta de la tarea).
$Out = [System.IO.Path]::GetFullPath([System.IO.Path]::Combine($PWD.Path, $Out))

Write-Host ("      perfil: " + $Perfil)

# ---------- 1. separar frontmatter del cuerpo --------------------------------
# $meta = claves de primer nivel. $port = claves del bloque "portada:", que
# tienen precedencia. Se separan para que "materia" del vault (sin acentos,
# la que resuelve Dataview) no pise la de la portada (con acentos).
function Desentrecomillar([string]$v) {
  $v = $v.Trim()
  if ($v.Length -ge 2) {
    if (($v[0] -eq '"' -and $v[-1] -eq '"') -or ($v[0] -eq "'" -and $v[-1] -eq "'")) {
      return $v.Substring(1, $v.Length - 2)
    }
  }
  return $v
}

$lineas = Get-Content -LiteralPath $Md -Encoding UTF8
$meta = @{}; $port = @{}; $integrantes = @(); $cuerpo = $lineas

if ($lineas.Count -gt 0 -and $lineas[0].Trim() -eq "---") {
  $i = 1; $enPortada = $false; $enInt = $false
  while ($i -lt $lineas.Count -and $lineas[$i].Trim() -ne "---") {
    $l = $lineas[$i]
    if ($l.Trim() -eq "") { $i++; continue }
    $sangria = $l.Length - $l.TrimStart().Length

    if ($sangria -eq 0) {
      # una clave sin sangria cierra cualquier bloque anidado
      $enPortada = $false; $enInt = $false
      if ($l -match '^([A-Za-z_][\w]*):\s*(.*)$') {
        $k = $Matches[1]; $v = Desentrecomillar $Matches[2]
        if     ($k -eq "portada"     -and $v -eq "") { $enPortada = $true }
        elseif ($k -eq "integrantes" -and $v -eq "") { $enInt = $true }
        else { $meta[$k] = $v }
      }
    }
    else {
      if ($enInt -and $l -match '^\s*-\s*(.+?)\s*$') { $integrantes += $Matches[1] }
      elseif ($enPortada -and $l -match '^\s*([A-Za-z_][\w]*):\s*(.*)$') {
        $k = $Matches[1]; $v = Desentrecomillar $Matches[2]
        $enInt = $false
        if ($k -eq "integrantes" -and $v -eq "") { $enInt = $true }
        else { $port[$k] = $v }
      }
    }
    $i++
  }
  if ($i + 1 -lt $lineas.Count) { $cuerpo = $lineas[($i+1)..($lineas.Count-1)] } else { $cuerpo = @() }
}

# M busca primero en el bloque portada y luego en el primer nivel
function M([string]$k) {
  if ($port.ContainsKey($k) -and $port[$k]) { return $port[$k] }
  if ($meta.ContainsKey($k) -and $meta[$k]) { return $meta[$k] }
  return ""
}

# ---------- 1b. recortar la seccion pedida -----------------------------------
$subir = 0
if ($Seccion) {
  $ini = -1; $nivel = 0
  for ($k = 0; $k -lt $cuerpo.Count; $k++) {
    if ($cuerpo[$k] -match '^(#+)\s+(.*)$') {
      if ($Matches[2].Trim() -eq $Seccion.Trim()) { $ini = $k; $nivel = $Matches[1].Length; break }
    }
  }
  if ($ini -lt 0) { throw "No encuentro la seccion '$Seccion' en $Md" }

  $fin = $cuerpo.Count
  for ($k = $ini + 1; $k -lt $cuerpo.Count; $k++) {
    if ($cuerpo[$k] -match '^(#+)\s+' -and $Matches[1].Length -le $nivel) { $fin = $k; break }
  }
  if ($fin -le $ini + 1) { throw "La seccion '$Seccion' esta vacia" }

  $cuerpo = $cuerpo[($ini+1)..($fin-1)]

  # Limpiar los bordes del recorte. En una nota de Obsidian es normal separar
  # secciones con una regla "---", y esa regla cae dentro del recorte. pandoc
  # la convierte en linea horizontal y Word truena despues con "This member
  # cannot be accessed on a horizontal line".
  $esBorde = { param($x) $x.Trim() -eq "" -or $x -match '^\s*([-*_])\1{2,}\s*$' }
  $a = 0; $b = $cuerpo.Count - 1
  while ($a -le $b -and (& $esBorde $cuerpo[$a])) { $a++ }
  while ($b -ge $a -and (& $esBorde $cuerpo[$b])) { $b-- }
  if ($b -lt $a) { throw "La seccion '$Seccion' esta vacia" }
  $cuerpo = $cuerpo[$a..$b]

  $subir = $nivel
  Write-Host ("      seccion: '" + $Seccion + "' (nivel " + $nivel + ") -> " + $cuerpo.Count + " lineas")
}
if ($SubirNivel -ge 0) { $subir = $SubirNivel }

# ---------- 1c. expandir marcadores de imagen ![] y ![Pie] -------------------
# Toma la siguiente imagen de la carpeta img/ en orden alfabetico, para no
# tener que nombrar ni numerar capturas.
$imgDir = Join-Path (Split-Path -Parent $Md) "img"
$pool = @()
if (Test-Path $imgDir) {
  $pool = Get-ChildItem -LiteralPath $imgDir -File |
          Where-Object { $_.Extension -match '^\.(png|jpg|jpeg)$' } |
          Sort-Object Name
}
$pi = 0
$cuerpo = @($cuerpo | ForEach-Object {
  $l = $_
  if ($l -match '^!\[([^\]]*)\]\s*$') {
    $cap = $Matches[1]
    if ($pi -lt $pool.Count) {
      $ruta = $pool[$pi].FullName.Replace('\', '/')
      $pi++
      '![' + $cap + '](' + $ruta + ')'
    } else {
      Write-Warning "no quedan imagenes en img/ para un marcador ![] - se omite"
      ''
    }
  } else { $l }
})
if ($pool.Count -gt 0) { Write-Host ("      imagenes: " + $pi + " de " + $pool.Count + " usadas de img\") }

# ---------- 2. pandoc: cuerpo.docx con los estilos de la referencia ----------
$tmpMd  = [System.IO.Path]::GetTempFileName() + ".md"
$tmpDoc = [System.IO.Path]::GetTempFileName() + ".docx"
$cuerpo -join "`n" | Out-File -FilePath $tmpMd -Encoding utf8

Write-Host "[1/4] pandoc: markdown -> docx con tus estilos..."
$argsPandoc = @($tmpMd, "-o", $tmpDoc, "--reference-doc=$Referencia", "--wrap=none")
if ($subir -gt 0) {
  $argsPandoc += "--shift-heading-level-by=-$subir"
  Write-Host ("      encabezados subidos " + $subir + " nivel(es): el primero queda como Titulo 1")
}
& $pandoc $argsPandoc
if (-not (Test-Path $tmpDoc)) { throw "pandoc no genero el cuerpo" }

# ---------- 3. Word: portada + cuerpo + indice + pdf -------------------------
Write-Host "[2/4] Word: armando el documento..."
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
  if ($cfg.portada) {
    if (-not (Test-Path $Plantilla)) { throw "No encuentro la plantilla de portada: $Plantilla" }
    Copy-Item $Plantilla $Out -Force
    $doc = $word.Documents.Open($Out)

    # {{CURSO}} es de la portada estandar: junta grupo, semestre y modalidad en
    # un renglon. La portada de formemp los lleva en renglones separados.
    $curso = @()
    if (M "grupo")     { $curso += "Grupo: "     + (M "grupo") }
    if (M "semestre")  { $curso += "Semestre: "  + (M "semestre") }
    if (M "modalidad") { $curso += "Modalidad: " + (M "modalidad") }

    if ($Perfil -eq "topicos") {
      # Esta portada no lleva tema ni renglon de curso unico: el grupo, la hora
      # y la frecuencia van en un solo renglon con sus etiquetas en negrita, ya
      # puestas en la plantilla. La fecha va tal cual, sin mayusculas.
      $equipo = ""; if (M "equipo") { $equipo = "Equipo " + (M "equipo") }
      # La ciudad va en el mismo renglon que la fecha, separada por ", a ".
      $fecha = (M "fecha")
      if ($fecha -and (M "ciudad")) { $fecha = (M "ciudad") + ", a " + $fecha }
      $subs = @(
        ,@("{{MATERIA}}",    (M "materia"))
        ,@("{{ACTIVIDAD}}",  (M "actividad"))
        ,@("{{INSTRUCTOR}}", (M "docente"))
        ,@("{{SEMESTRE}}",   (M "semestre"))
        ,@("{{GRUPO}}",      (M "grupo"))
        ,@("{{HORA}}",       (M "hora"))
        ,@("{{FRECUENCIA}}", (M "frecuencia"))
        ,@("{{EQUIPO}}",     $equipo)
        ,@("{{FECHA}}",      $fecha)
      )
    }
    elseif ($Perfil -eq "formemp") {
      # esta portada NO lleva mayusculas ni el prefijo "A", y parte la fecha
      # en dos renglones. Ver "Formato - M.A. Guillermo Marin Rangel".
      $equipo = ""; if (M "equipo") { $equipo = "Equipo #" + (M "equipo") }
      $subs = @(
        ,@("{{MATERIA}}",   (M "materia"))
        ,@("{{ACTIVIDAD}}", (M "actividad"))
        ,@("{{TEMA}}",      (M "tema"))
        ,@("{{EQUIPO}}",    $equipo)
        ,@("{{GRUPO}}",     $(if (M "grupo")     { "Grupo: "     + (M "grupo") }     else { "" }))
        ,@("{{DOCENTE}}",   $(if (M "docente")   { "Docente: "   + (M "docente") }   else { "" }))
        ,@("{{SEMESTRE}}",  $(if (M "semestre")  { "Semestre: "  + (M "semestre") }  else { "" }))
        ,@("{{PLAN}}",      $(if (M "plan")      { "Plan: "      + (M "plan") }      else { "" }))
        ,@("{{CIUDAD}}",    (M "ciudad"))
        ,@("{{FECHA}}",     (M "fecha"))
      )
    }
    elseif ($Perfil -eq "lbtssi") {
      # Portada del Laboratorio de Temas Selectos de Sistemas Inteligentes: igual
      # a la estandar de FIME, pero el grupo se llama "Brigada", hay un renglon
      # extra para la hora y la tabla lleva una cuarta columna, "INSCRITO EN LAB".
      $fecha = ""
      if (M "fecha") { $fecha = "SAN NICOLÁS DE LOS GARZA, N.L. A " + (M "fecha").ToUpper() }
      $equipo = ""; if (M "equipo") { $equipo = "Equipo: " + (M "equipo") }
      $curso = @()
      if (M "brigada")   { $curso += "Brigada: "   + (M "brigada") }
      if (M "grupo")     { $curso += "Grupo: "     + (M "grupo") }
      if (M "semestre")  { $curso += "Semestre: "  + (M "semestre") }
      if (M "modalidad") { $curso += "Modalidad: " + (M "modalidad") }
      $subs = @(
        ,@("{{MATERIA}}",   (M "materia"))
        ,@("{{ACTIVIDAD}}", (M "actividad"))
        ,@("{{TEMA}}",      (M "tema"))
        ,@("{{EQUIPO}}",    $equipo)
        ,@("{{DOCENTE}}",   $(if (M "docente") { "Docente: " + (M "docente") } else { "" }))
        ,@("{{CURSO}}",     ($curso -join " "))
        ,@("{{HORA}}",      $(if (M "hora") { "Hora: " + (M "hora") } else { "" }))
        ,@("{{FECHA}}",     $fecha)
      )
    }
    else {
      $fecha = ""
      if (M "fecha") { $fecha = "SAN NICOLÁS DE LOS GARZA, N.L. A " + (M "fecha").ToUpper() }
      $equipo = ""; if (M "equipo") { $equipo = "Equipo: " + (M "equipo") }
      $subs = @(
        ,@("{{MATERIA}}",   (M "materia"))
        ,@("{{ACTIVIDAD}}", (M "actividad"))
        ,@("{{TEMA}}",      (M "tema"))
        ,@("{{EQUIPO}}",    $equipo)
        ,@("{{DOCENTE}}",   $(if (M "docente") { "Docente: " + (M "docente") } else { "" }))
        ,@("{{CURSO}}",     ($curso -join " "))
        ,@("{{FECHA}}",     $fecha)
      )
    }

    foreach ($s in $subs) {
      $f = $doc.Content.Find
      $f.ClearFormatting(); $f.Replacement.ClearFormatting()
      $null = $f.Execute([ref]$s[0],[ref]$false,[ref]$false,[ref]$false,[ref]$false,[ref]$false,[ref]$true,[ref]1,[ref]$false,[ref]$s[1],[ref]2)
    }
    # borrar renglones que quedaron vacios (campos no usados)
    for ($i = $doc.Paragraphs.Count; $i -ge 1; $i--) {
      $t = $doc.Paragraphs.Item($i).Range.Text.Trim()
      if ($t -eq "" -and $doc.Paragraphs.Item($i).Range.Tables.Count -eq 0) {
        $prev = if ($i -gt 1) { $doc.Paragraphs.Item($i-1).Range.Text.Trim() } else { "x" }
        if ($prev -eq "") { $doc.Paragraphs.Item($i).Range.Delete() | Out-Null }
      }
    }
    # tabla de integrantes: clonar la fila molde
    if ($integrantes.Count -gt 0 -and $doc.Tables.Count -gt 0) {
      $t = $doc.Tables.Item(1)
      $ncol = $t.Columns.Count
      foreach ($fila in $integrantes) {
        $campos = $fila -split '\|' | ForEach-Object { $_.Trim() }
        # El frontmatter siempre lista "matricula | nombre | carrera". La portada
        # de Topicos imprime "Estudiante | Matricula | Carrera", asi que aqui se
        # intercambian las dos primeras columnas y el frontmatter no cambia.
        if ($Perfil -eq "topicos" -and $campos.Count -ge 2) {
          $resto = @()
          if ($campos.Count -ge 3) { $resto = $campos[2..($campos.Count - 1)] }
          $campos = @($campos[1], $campos[0]) + $resto
        }
        $r = $t.Rows.Add()
        for ($c = 1; $c -le $ncol; $c++) {
          $v = if ($c -le $campos.Count) { $campos[$c-1] } else { "" }
          $t.Cell($r.Index, $c).Range.Text = $v
        }
      }
      $t.Rows.Item(2).Delete()   # quitar el molde {{C1}}...
      Write-Host ("      integrantes: " + $integrantes.Count)
    }
    # insertar el cuerpo despues de la portada (y del indice, si la plantilla
    # lo trae). InsertFile respeta los estilos del destino, que son los mismos.
    #
    # Si la plantilla termina con el campo TOC, Content.End cae DENTRO del
    # campo y InsertBreak truena con "the current selection is outside of a
    # block-level XML element". Se agrega un parrafo al final para tener un
    # punto de insercion fuera del campo.
    if ($doc.TablesOfContents.Count -gt 0) { $doc.Content.InsertParagraphAfter() }
    $fin = $doc.Content.End - 1
    $rng = $doc.Range($fin, $fin)
    $rng.InsertBreak(7)                      # wdPageBreak
    $rng.Collapse(0)
    $rng.InsertFile($tmpDoc)
  }
  else {
    Copy-Item $tmpDoc $Out -Force
    $doc = $word.Documents.Open($Out)
    # encabezado: tarea <tab> nombre <tab> matricula
    if (-not $NOMBRE -and -not $MATRIC) {
      Write-Warning "Sin config.json: el encabezado sale sin nombre ni matricula. Copia config.ejemplo.json a config.json y rellenalo."
    }
    $tarea = if ($Encabezado) { $Encabezado } else { [System.IO.Path]::GetFileNameWithoutExtension($Out) }
    $hdr = $doc.Sections.Item(1).Headers.Item(1).Range
    $hdr.Text = "$tarea`t$NOMBRE`t$MATRIC"
    $hdr.Font.Name = "Arial"; $hdr.Font.Size = 10; $hdr.Font.Bold = $false
    $hdr.ParagraphFormat.Alignment = 0
    # tabulaciones explicitas: tarea a la izquierda, nombre centrado, matricula a la derecha.
    # Sin esto los tres datos quedan pegados con la tabulacion por defecto.
    $ancho = $doc.PageSetup.PageWidth - $doc.PageSetup.LeftMargin - $doc.PageSetup.RightMargin
    $pf = $hdr.ParagraphFormat
    $pf.TabStops.ClearAll()
    $null = $pf.TabStops.Add(($ancho / 2), 1)   # 1 = wdAlignTabCenter
    $null = $pf.TabStops.Add($ancho, 2)         # 2 = wdAlignTabRight
  }

  # Bordes completos en las tablas. pandoc deja solo las lineas horizontales del
  # estilo "Table"; con -BordesTabla toda tabla queda con linea solida negra en
  # todos sus bordes, exteriores e interiores, sin lineas que haya que deducir.
  if ($BordesTabla -and $doc.Tables.Count -gt 0) {
    foreach ($ti in 1..$doc.Tables.Count) {
      $tb = $doc.Tables.Item($ti)
      $tb.Borders.InsideLineStyle  = 1    # wdLineStyleSingle
      $tb.Borders.OutsideLineStyle = 1
      $tb.Borders.InsideLineWidth  = 8    # 1 pt
      $tb.Borders.OutsideLineWidth = 8
      $tb.Borders.InsideColor  = 0        # negro
      $tb.Borders.OutsideColor = 0
    }
    Write-Host ("      bordes completos aplicados a " + $doc.Tables.Count + " tablas")
  }

  # Formato interno de las tablas. pandoc deja los parrafos de las celdas con el
  # mismo justificado del cuerpo, y en una columna angosta eso estira las
  # palabras hasta hacerlas ilegibles. Aqui se alinean a la izquierda, se bajan
  # a 10 pt y la tabla se ajusta al ancho de la pagina.
  if ($TablaCompacta -and $doc.Tables.Count -gt 0) {
    foreach ($ti in 1..$doc.Tables.Count) {
      $tb = $doc.Tables.Item($ti)
      $tb.Range.Font.Name = "Arial"
      $tb.Range.Font.Size = 10
      $tb.Range.ParagraphFormat.Alignment  = 0    # izquierda
      $tb.Range.ParagraphFormat.SpaceAfter = 0
      $tb.Range.ParagraphFormat.SpaceBefore = 0
      $tb.Range.ParagraphFormat.LineSpacingRule = 0   # sencillo
      $tb.Range.Cells.VerticalAlignment = 0        # wdCellAlignVerticalTop
      $tb.AutoFitBehavior(2)                      # wdAutoFitWindow
      $tb.Rows.Item(1).HeadingFormat = $true      # repetir encabezado al cambiar de pagina
      $tb.Rows.Item(1).Range.Font.Bold = $true
      $tb.Rows.AllowBreakAcrossPages = $false
    }
    Write-Host ("      formato compacto aplicado a " + $doc.Tables.Count + " tablas")
  }

  # Bloques de codigo. pandoc los deja con el estilo "Source Code", que hereda
  # el justificado del cuerpo: en un bloque de codigo eso estira los espacios
  # hasta separar las palabras media linea y lo vuelve ilegible. Se alinean a
  # la izquierda, a Consolas 10, interlineado sencillo y sin separacion entre
  # renglones, que es como se lee un listado.
  $codigo = 0
  foreach ($pi in 1..$doc.Paragraphs.Count) {
    $p = $doc.Paragraphs.Item($pi)
    $nom = ""
    try { $nom = $p.Style.NameLocal } catch { }
    if ($nom -eq "Source Code") {
      $p.Format.Alignment       = 0    # izquierda
      $p.Format.LineSpacingRule = 0    # sencillo
      $p.Format.SpaceAfter      = 0
      $p.Format.SpaceBefore     = 0
      $p.Range.Font.Name = "Consolas"
      $p.Range.Font.Size = 10
      $codigo++
    }
  }
  if ($codigo -gt 0) { Write-Host ("      bloques de codigo: " + $codigo + " renglones alineados a la izquierda") }

  # ajustar imagenes que se salen de la columna (capturas a tamano nativo)
  if ($doc.InlineShapes.Count -gt 0) {
    $maxW = $doc.PageSetup.PageWidth - $doc.PageSetup.LeftMargin - $doc.PageSetup.RightMargin
    $maxH = 312   # ~11 cm en puntos
    $ajust = 0
    foreach ($i in 1..$doc.InlineShapes.Count) {
      $s = $doc.InlineShapes.Item($i)
      $s.LockAspectRatio = -1
      if ($s.Width -gt $maxW)  { $s.Width  = $maxW; $ajust++ }
      if ($s.Height -gt $maxH) { $s.Height = $maxH; $ajust++ }
    }
    if ($ajust -gt 0) { Write-Host ("      $ajust imagenes reescaladas al ancho de columna") }
  }

  # indice automatico: la plantilla ya trae el campo TOC, aqui solo se refresca
  # para que tome los titulos del cuerpo recien insertado y sus paginas.
  if ($doc.TablesOfContents.Count -gt 0) {
    Write-Host "[3/4] Word: actualizando el indice..."
    $doc.Fields.Update() | Out-Null
    $doc.TablesOfContents.Item(1).Update()
    Write-Host ("      entradas del indice: " + $doc.TablesOfContents.Item(1).Range.Paragraphs.Count)
  }

  $doc.Save()
  $paginas = $doc.ComputeStatistics(2)
  $palabras = $doc.ComputeStatistics(0)
  $imgs = $doc.InlineShapes.Count + $doc.Shapes.Count

  if (-not $SinPdf) {
    Write-Host "[4/4] exportando PDF..."
    $pdf = [System.IO.Path]::ChangeExtension($Out, ".pdf")
    $doc.ExportAsFixedFormat($pdf, 17)
  }
  $doc.Close($false)
  Write-Host ""
  Write-Host ("OK -> " + $Out)
  if (-not $SinPdf) { Write-Host ("OK -> " + [System.IO.Path]::ChangeExtension($Out, ".pdf")) }
  Write-Host ("     $paginas paginas, $palabras palabras, $imgs imagenes")
}
finally {
  # Si algo trueno a media construccion, el documento sigue abierto y Quit()
  # se queda esperando: eso deja instancias de WINWORD.EXE huerfanas que
  # bloquean el archivo en el siguiente intento. Se cierra sin guardar antes
  # de salir.
  # Ojo: Quit() no acepta argumentos por COM desde PowerShell 5.1 ("should be
  # a System.Management.Automation.PSReference"), asi que se llama sin ellos.
  try { if ($doc) { $doc.Close($false) } } catch { }
  try { $word.Quit() } catch { }
  Remove-Item $tmpMd, $tmpDoc -ErrorAction SilentlyContinue
}
