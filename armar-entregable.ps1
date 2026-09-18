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
#   lbtssi    portada del Laboratorio de Temas Sel. de Sistemas Inteligentes.
#             Es la de FIME con tres cambios: el grupo se llama Brigada, la hora
#             lleva renglon propio y la tabla de integrantes tiene una cuarta
#             columna, INSCRITO EN LAB. Frontmatter: brigada, hora, y cada
#             integrante con cuatro campos (matricula | nombre | carrera | SI/NO).
#             Estilos: los mismos de fime. Usa -BordesTabla -TablaCompacta.
#   agronomia portada de la Facultad de Agronomia (UANL). Es la de fime con el
#             escudo de FAUANL y su renglon de facultad. La ciudad va en el
#             frontmatter ("ciudad:") porque el campus no esta en San Nicolas,
#             y el renglon de curso junta grupo, hora y semestre. Trae indice
#             automatico y numera las paginas a partir del cuerpo, despues del
#             indice. Estilos: los mismos de fime.
#   integrador portada del equipo 5 de Proyecto Integrador I (Dr. Juvencio
#             Jaramillo), recortada de la Fase 4 que armo el equipo en Google
#             Docs: logos, tabla de integrantes, indice automatico y una seccion
#             aparte para el cuerpo, que es la que numera paginas. El cuerpo abre
#             con "Fase N" en estilo Titulo (frontmatter "titulo:"). Estilos del
#             equipo: Arial 12 justificado 1.5, Titulo 1 y 2 en Play azul #0f4761.
#             Trae ==texto== como resaltado amarillo. Usa -TablaIntegrador.
#             Las plantillas salen de herramientas\crear-plantillas-integrador.ps1.
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
  [ValidateSet("martinez","fime","formemp","topicos","lbtssi","agronomia","integrador")][string]$Perfil = "",
  # alto maximo de una imagen, en puntos. 312 (~11 cm) sirve para capturas;
  # un diagrama con texto necesita mas alto para que la letra se lea.
  [int]$AltoMaxImagen = 312,
  [string]$Seccion = "",
  [switch]$ConPortada,
  [string]$Encabezado = "",
  [string]$Referencia = "",
  [string]$Plantilla  = "",
  [int]$SubirNivel = -1,
  [switch]$SinPdf,
  [switch]$BordesTabla,
  [switch]$TablaCompacta,
  # tablas como las del equipo de Proyecto Integrador: encabezado azul marino
  # con letra blanca, bordes gris claro, 10 pt
  [switch]$TablaIntegrador
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
  "agronomia"= @{ portada = $true;  ref = "reference-fime.docx";     plantilla = "portada-agronomia.docx" }
  "integrador"=@{ portada = $true;  ref = "reference-integrador.docx"; plantilla = "portada-integrador.docx" }
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
  # Un bloque de codigo puede traer renglones que empiezan con "#" (un
  # comentario de Python, una directiva de C). Si se leen como encabezado, el
  # recorte de la seccion se corta ahi y el entregable sale truncado sin avisar.
  # Por eso las dos busquedas saltan lo que este dentro de un cerco ``` o ~~~.
  $ini = -1; $nivel = 0; $cerco = $false
  for ($k = 0; $k -lt $cuerpo.Count; $k++) {
    if ($cuerpo[$k] -match '^\s*(```|~~~)') { $cerco = -not $cerco; continue }
    if ($cerco) { continue }
    if ($cuerpo[$k] -match '^(#+)\s+(.*)$') {
      if ($Matches[2].Trim() -eq $Seccion.Trim()) { $ini = $k; $nivel = $Matches[1].Length; break }
    }
  }
  if ($ini -lt 0) { throw "No encuentro la seccion '$Seccion' en $Md" }

  $fin = $cuerpo.Count; $cerco = $false
  for ($k = $ini + 1; $k -lt $cuerpo.Count; $k++) {
    if ($cuerpo[$k] -match '^\s*(```|~~~)') { $cerco = -not $cerco; continue }
    if ($cerco) { continue }
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
$mdDir = Split-Path -Parent $Md
$argsPandoc = @($tmpMd, "-f", "markdown+mark", "-o", $tmpDoc, "--reference-doc=$Referencia", "--wrap=none", "--resource-path=$mdDir")
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

    if ($Perfil -eq "integrador") {
      # La portada ya trae escritos instructor, semestre, grupo, hora y la
      # tabla del equipo, copiados tal cual del entregable del equipo. Solo
      # cambian el subtitulo subrayado, la fecha y el "Fase N" del cuerpo.
      $subs = @(
        ,@("{{ACTIVIDAD}}", (M "actividad"))
        ,@("{{FECHA}}",     (M "fecha"))
        ,@("{{TITULO}}",    (M "titulo"))
      )
    }
    elseif ($Perfil -eq "agronomia") {
      # Portada de la Facultad de Agronomia. Es la de FIME con el escudo y el
      # renglon de la facultad cambiados, mas dos diferencias de contenido:
      # la ciudad no esta clavada en el codigo (el campus no esta en San
      # Nicolas) y el renglon de curso junta grupo, hora y semestre, que es
      # como se identifica ahi un grupo. Lleva indice automatico.
      $curso = @()
      if (M "grupo")    { $curso += "Grupo: "    + (M "grupo") }
      if (M "hora")     { $curso += "Hora: "     + (M "hora") }
      if (M "semestre") { $curso += "Semestre: " + (M "semestre") }
      $ciudad = if (M "ciudad") { (M "ciudad") } else { "San Nicolás de los Garza, N.L." }
      $fecha = ""
      if (M "fecha") { $fecha = $ciudad.ToUpper() + " A " + (M "fecha").ToUpper() }
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
    elseif ($Perfil -eq "lbtssi") {
      # Portada del Laboratorio de Temas Sel. de Sistemas Inteligentes. Es la de
      # FIME con tres diferencias: el grupo se llama Brigada, hay un renglon
      # propio para la hora, y la tabla de integrantes lleva una cuarta columna
      # (INSCRITO EN LAB) porque no todos los del equipo estan inscritos al
      # laboratorio. El frontmatter usa "brigada:" y "hora:".
      $curso = @()
      if (M "brigada")   { $curso += "Brigada: "   + (M "brigada") }
      if (M "semestre")  { $curso += "Semestre: "  + (M "semestre") }
      if (M "modalidad") { $curso += "Modalidad: " + (M "modalidad") }
      $fecha = ""
      if (M "fecha") { $fecha = "SAN NICOLÁS DE LOS GARZA, N.L. A " + (M "fecha").ToUpper() }
      $equipo = ""; if (M "equipo") { $equipo = "Equipo: " + (M "equipo") }
      $subs = @(
        ,@("{{MATERIA}}",   (M "materia"))
        ,@("{{ACTIVIDAD}}", (M "actividad"))
        ,@("{{DOCENTE}}",   $(if (M "docente") { "Docente: " + (M "docente") } else { "" }))
        ,@("{{EQUIPO}}",    $equipo)
        ,@("{{CURSO}}",     ($curso -join " "))
        ,@("{{HORA}}",      $(if (M "hora") { "Hora: " + (M "hora") } else { "" }))
        ,@("{{FECHA}}",     $fecha)
      )
    }
    elseif ($Perfil -eq "topicos") {
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
    # borrar renglones que quedaron vacios (campos no usados). La portada de
    # integrador usa parrafos vacios a proposito para acomodar los logos, y
    # no tiene campos opcionales, asi que ahi no se toca.
    for ($i = $doc.Paragraphs.Count; $i -ge 1 -and $Perfil -ne "integrador"; $i--) {
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
    if ($Perfil -eq "integrador") {
      # La plantilla termina en la seccion del cuerpo, con el "Fase N" como
      # ultimo parrafo: el cuerpo se pega despues, sin salto, para que quede
      # dentro de esa seccion y herede su pie con el numero de pagina.
      $doc.Content.InsertParagraphAfter()
      $ult = $doc.Paragraphs.Item($doc.Paragraphs.Count)
      $ult.Style = -1   # wdStyleNormal
      $fin = $doc.Content.End - 1
      $rng = $doc.Range($fin, $fin)
      $rng.InsertFile($tmpDoc)
      # InsertFile deja un parrafo vacio al final; se quita si sobra
      $ult = $doc.Paragraphs.Item($doc.Paragraphs.Count)
      if ($ult.Range.Text.Trim() -eq "" -and $ult.Range.InlineShapes.Count -eq 0) {
        try { $doc.Range($ult.Range.Start - 1, $ult.Range.End - 1).Delete() | Out-Null } catch { }
      }
    }
    if ($Perfil -ne "integrador") {
    if ($doc.TablesOfContents.Count -gt 0) { $doc.Content.InsertParagraphAfter() }
    $fin = $doc.Content.End - 1
    $rng = $doc.Range($fin, $fin)
    # agronomia pide que la numeracion de pagina aparezca despues del indice,
    # asi que ahi el corte es de SECCION y no de pagina: el cuerpo queda en una
    # seccion aparte con su propio pie de pagina. lbtssi tambien numera desde
    # ahi. El resto de los perfiles no numeran, asi que les basta el salto de
    # pagina.
    if ($Perfil -eq "agronomia" -or $Perfil -eq "lbtssi") { $rng.InsertBreak(2) }  # wdSectionBreakNextPage
    else                                                  { $rng.InsertBreak(7) }  # wdPageBreak
    $rng.Collapse(0)
    $rng.InsertFile($tmpDoc)
    }

    if ($Perfil -eq "agronomia" -and $doc.Sections.Count -ge 2) {
      # La portada y el indice se cuentan pero no se imprimen: la numeracion
      # sigue corrida (la introduccion sale como pagina 3), que es lo que pide
      # APA. Va en el encabezado y alineada a la derecha, que es donde la
      # coloca APA 7, y se desliga para que la seccion 1 no la herede.
      $enc = $doc.Sections.Item(2).Headers.Item(1)
      $enc.LinkToPrevious = $false
      $doc.Sections.Item(1).Headers.Item(1).Range.Text = ""
      $r = $enc.Range
      $r.Text = ""
      $r.Font.Name = "Arial"; $r.Font.Size = 12; $r.Font.Bold = $false
      $null = $doc.Fields.Add($r, 33)        # wdFieldPage
      # la alineacion va DESPUES de insertar el campo: puesta antes, Fields.Add
      # la revierte a la del estilo Normal y el numero sale a la izquierda
      $enc.Range.ParagraphFormat.Alignment = 2   # wdAlignParagraphRight
      Write-Host "      numeracion de pagina desde la seccion 2 (despues del indice)"
    }

    if ($Perfil -eq "lbtssi" -and $doc.Sections.Count -ge 2) {
      # El equipo del laboratorio numera en el PIE de pagina, a la derecha, y la
      # cuenta arranca en 1 en el cuerpo: portada e indice no cuentan. Asi quedo
      # la version que el equipo entrego de la Practica 2.1. Se desliga el pie
      # antes de vaciar el de la seccion 1, o el vaciado se hereda.
      $pie = $doc.Sections.Item(2).Footers.Item(1)
      $pie.LinkToPrevious = $false
      $doc.Sections.Item(1).Footers.Item(1).Range.Text = ""
      $pie.PageNumbers.RestartNumberingAtSection = $true
      $pie.PageNumbers.StartingNumber = 1
      $r = $pie.Range
      $r.Text = ""
      $r.Font.Name = "Arial"; $r.Font.Size = 12; $r.Font.Bold = $false
      $null = $doc.Fields.Add($r, 33)        # wdFieldPage
      $pie.Range.ParagraphFormat.Alignment = 2   # wdAlignParagraphRight, despues del campo
      Write-Host "      numeracion de pagina en el pie desde la seccion 2 (despues del indice)"
    }
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

  # Tablas del equipo de Proyecto Integrador: encabezado #073763 con letra
  # blanca en negrita, bordes #c4c7c5 y celdas a 10 pt alineadas a la
  # izquierda. Se salta la tabla de integrantes de la portada (la primera).
  if ($TablaIntegrador -and $doc.Tables.Count -gt 0) {
    $azul = 0x07 + 0x37*256 + 0x63*65536
    $gris = 0xc4 + 0xc7*256 + 0xc5*65536
    $ini = if ($cfg.portada) { 2 } else { 1 }
    $n = 0
    for ($ti = $ini; $ti -le $doc.Tables.Count; $ti++) {
      $tb = $doc.Tables.Item($ti)
      $tb.Range.Font.Name = "Arial"
      $tb.Range.Font.Size = 10
      $tb.Range.ParagraphFormat.Alignment = 0
      $tb.Range.ParagraphFormat.SpaceAfter = 2
      $tb.Range.ParagraphFormat.SpaceBefore = 2
      $tb.Range.ParagraphFormat.LineSpacingRule = 0
      $tb.Range.Cells.VerticalAlignment = 1          # centrado vertical
      $tb.Borders.InsideLineStyle = 1; $tb.Borders.OutsideLineStyle = 1
      $tb.Borders.InsideLineWidth = 4; $tb.Borders.OutsideLineWidth = 4
      $tb.Borders.InsideColor = $gris; $tb.Borders.OutsideColor = $gris
      $tb.TopPadding = 3; $tb.BottomPadding = 3; $tb.LeftPadding = 5; $tb.RightPadding = 5
      $tb.AutoFitBehavior(2)
      $f1 = $tb.Rows.Item(1)
      $f1.HeadingFormat = $true
      $f1.Shading.BackgroundPatternColor = $azul
      $f1.Range.Font.Color = 0xFFFFFF
      $f1.Range.Font.Bold = $true
      $f1.Range.ParagraphFormat.Alignment = 1
      $tb.Rows.AllowBreakAcrossPages = $true
      # el parrafo que sigue a la tabla queda pegado al borde inferior; se separa
      try {
        $sig = $doc.Range($tb.Range.End, $tb.Range.End)
        if ($sig.Paragraphs.Count -gt 0 -and $sig.Paragraphs.Item(1).Range.Tables.Count -eq 0) {
          $sig.Paragraphs.Item(1).SpaceBefore = 12
        }
      } catch { }
      $n++
    }
    Write-Host ("      tablas con formato del equipo: " + $n)
  }

  # Bloques de codigo. pandoc los deja con el estilo "Source Code", que hereda
  # el justificado del cuerpo: en un bloque de codigo eso estira los espacios
  # hasta separar las palabras media linea y lo vuelve ilegible. Se alinean a
  # la izquierda, a Consolas 10, interlineado sencillo y sin separacion entre
  # renglones, que es como se lee un listado.
  $codigo = 0
  $anteriorEraCodigo = $false
  foreach ($pi in 1..$doc.Paragraphs.Count) {
    $p = $doc.Paragraphs.Item($pi)
    $nom = ""
    try { $nom = $p.Style.NameLocal } catch { }
    # integrador: el parrafo que sigue a un bloque de codigo queda pegado al fondo gris
    if ($Perfil -eq "integrador" -and $anteriorEraCodigo -and $nom -ne "Source Code") { $p.Format.SpaceBefore = 10 }
    $anteriorEraCodigo = ($nom -eq "Source Code")
    if ($nom -eq "Source Code") {
      $p.Format.Alignment       = 0    # izquierda
      $p.Format.LineSpacingRule = 0    # sencillo
      $p.Format.SpaceAfter      = 0
      $p.Format.SpaceBefore     = 0
      $p.Range.Font.Name = "Consolas"
      $p.Range.Font.Size = $(if ($Perfil -eq "integrador") { 8 } else { 10 })
      if ($Perfil -eq "integrador") { $p.Shading.BackgroundPatternColor = 0xF2F2F2 }
      $codigo++
    }
  }
  if ($codigo -gt 0) { Write-Host ("      bloques de codigo: " + $codigo + " renglones alineados a la izquierda") }

  # ajustar imagenes que se salen de la columna (capturas a tamano nativo)
  if ($doc.InlineShapes.Count -gt 0) {
    $maxW = $doc.PageSetup.PageWidth - $doc.PageSetup.LeftMargin - $doc.PageSetup.RightMargin
    $maxH = $AltoMaxImagen   # 312 pt (~11 cm) por omision
    $ajust = 0
    foreach ($i in 1..$doc.InlineShapes.Count) {
      $s = $doc.InlineShapes.Item($i)
      $s.LockAspectRatio = -1
      if ($s.Width -gt $maxW)  { $s.Width  = $maxW; $ajust++ }
      if ($s.Height -gt $maxH) { $s.Height = $maxH; $ajust++ }
      # en integrador las capturas y diagramas van centrados
      if ($Perfil -eq "integrador") { $s.Range.ParagraphFormat.Alignment = 1 }
    }
    if ($ajust -gt 0) { Write-Host ("      $ajust imagenes reescaladas al ancho de columna") }
  }

  # saltos de pagina pedidos desde el markdown: un renglon que diga exactamente
  # [salto de pagina] se convierte en un salto manual. pandoc no tiene forma de
  # expresarlo, y hay profesores que piden secciones en hoja aparte (por ejemplo
  # la bibliografia). El marcador se borra junto con su parrafo.
  # Se recorren los parrafos, NO se usa Find: Find con "^p" en el texto buscado
  # y reemplazo por "^m" deja a Word girando sin terminar (comprobado el
  # 2026-09-11, se colgo con 200 s de CPU y hubo que matar el proceso).
  # Recorrer de atras hacia adelante mantiene validos los indices al borrar.
  $nSaltos = 0
  for ($i = $doc.Paragraphs.Count; $i -ge 1; $i--) {
    $par = $doc.Paragraphs.Item($i)
    if ($par.Range.Text.Trim() -eq "[salto de pagina]") {
      $r = $par.Range
      $r.MoveEnd(1, -1) | Out-Null    # wdCharacter: dejar fuera la marca de parrafo
      $r.Text = ""                    # borrar el marcador, conservar el parrafo
      $r.InsertBreak(7)               # wdPageBreak: el parrafo pasa a ser el salto
      $nSaltos++
    }
  }
  if ($nSaltos -gt 0) { Write-Host "      $nSaltos saltos de pagina del markdown" }

  # sangria francesa de las referencias (APA): las entradas van alineadas a la
  # izquierda y con la segunda linea en adelante recorrida 1.27 cm. Se aplica a
  # los parrafos que van despues del titulo de referencias, hasta el final o
  # hasta el siguiente titulo. Solo el perfil agronomia lo pide por escrito.
  if ($Perfil -eq "agronomia") {
    $enRef = $false; $nRef = 0
    foreach ($i in 1..$doc.Paragraphs.Count) {
      $par = $doc.Paragraphs.Item($i)
      $txt = $par.Range.Text.Trim()
      $est = [string]$par.Style.NameLocal
      if ($est -like "T*tulo 1" -or $est -eq "Heading 1") {
        $enRef = ($txt -like "Referencias*" -or $txt -like "Bibliograf*")
        continue
      }
      if ($enRef -and $txt -ne "") {
        $par.Alignment = 0                        # izquierda, no justificado
        $par.LeftIndent = 36                      # 1.27 cm en puntos
        $par.FirstLineIndent = -36                # sangria francesa
        $nRef++
      }
    }
    if ($nRef -gt 0) { Write-Host "      $nRef referencias con sangria francesa" }
  }

  # indice automatico: la plantilla ya trae el campo TOC, aqui solo se refresca
  # para que tome los titulos del cuerpo recien insertado y sus paginas.
  if ($doc.TablesOfContents.Count -gt 0) {
    Write-Host "[3/4] Word: actualizando el indice..."
    $doc.Fields.Update() | Out-Null
    $doc.TablesOfContents.Item(1).Update()
    Write-Host ("      entradas del indice: " + $doc.TablesOfContents.Item(1).Range.Paragraphs.Count)
  }

  # Propiedades del archivo. La plantilla de portada salio de un entregable
  # real, asi que el .docx hereda el autor de quien lo hizo. Para un trabajo
  # individual eso es un dato mal puesto, y el PDF se lleva el mismo. El autor
  # correcto es el alumno que entrega: se toma del primer integrante.
  # UserName se cambia antes de guardar para que "modificado por" tampoco
  # delate la maquina en la que se armo, y se restaura al terminar.
  $autor = ""
  if ($integrantes.Count -gt 0) {
    $campos = $integrantes[0] -split '\|'
    if ($campos.Count -ge 2) { $autor = $campos[1].Trim() }
  }
  if (-not $autor) { $autor = $NOMBRE }
  $userPrevio = $null
  if ($autor) {
    try {
      $userPrevio = $word.UserName
      $word.UserName = $autor
      # BuiltInDocumentProperties es una propiedad con parametro y PowerShell
      # 5.1 no la sabe enlazar directo ("Object reference not set..."): hay que
      # llegarle por reflexion.
      $props = $doc.BuiltInDocumentProperties
      function Set-PropDoc([string]$nombre, [string]$valor) {
        $p = [System.__ComObject].InvokeMember("Item","GetProperty",$null,$props,@($nombre))
        [void][System.__ComObject].InvokeMember("Value","SetProperty",$null,$p,@($valor))
      }
      Set-PropDoc "Author"  $autor
      Set-PropDoc "Company" ""
      $tituloProp = (M "tema"); if (-not $tituloProp) { $tituloProp = (M "actividad") }
      if ($tituloProp) { Set-PropDoc "Title" $tituloProp }
      Write-Host ("      autor del archivo: " + $autor)
    } catch {
      Write-Warning ("no se pudieron fijar las propiedades del archivo: " + $_.Exception.Message)
    }
  }

  # integrador: la fuente Play de los titulos no viene con Windows ni Office;
  # se incrusta para que el documento se vea igual en la computadora de otro
  # integrante del equipo.
  if ($Perfil -eq "integrador") { $doc.EmbedTrueTypeFonts = $true; $doc.SaveSubsetFonts = $false }

  $doc.Save()
  $paginas = $doc.ComputeStatistics(2)
  $palabras = $doc.ComputeStatistics(0)
  $imgs = $doc.InlineShapes.Count + $doc.Shapes.Count

  if (-not $SinPdf) {
    Write-Host "[4/4] exportando PDF..."
    $pdf = [System.IO.Path]::ChangeExtension($Out, ".pdf")
    # El parametro 11 es CreateBookmarks: con 1 (wdExportCreateHeadingBookmarks)
    # los Titulo 1/2/3 salen como marcadores en el panel de navegacion del PDF,
    # ademas del indice con hipervinculos que ya viene del campo TOC.
    $doc.ExportAsFixedFormat($pdf, 17, $false, 0, 0, 0, 0, 0, $true, $true, 1)
  }
  $doc.Close($false)
  $doc = $null

  # "Modificado por" lo escribe Word al guardar, tomandolo del usuario
  # registrado de la instalacion, y no obedece a UserName ni a las propiedades
  # del documento. En un trabajo individual eso deja el nombre de otra persona
  # dentro del archivo, que es justo lo que prohibe la nota de redaccion de
  # entregables. Se corrige sobre el .docx ya cerrado, que es un zip.
  if ($autor) {
    try {
      Add-Type -AssemblyName System.IO.Compression.FileSystem
      $zip = [System.IO.Compression.ZipFile]::Open($Out, "Update")
      $core = $zip.GetEntry("docProps/core.xml")
      if ($core) {
        $sr = New-Object System.IO.StreamReader($core.Open())
        $xml = $sr.ReadToEnd(); $sr.Close()
        $xml = [regex]::Replace($xml, "<cp:lastModifiedBy>.*?</cp:lastModifiedBy>", "<cp:lastModifiedBy>$autor</cp:lastModifiedBy>")
        $st = $core.Open(); $st.SetLength(0)
        $sw = New-Object System.IO.StreamWriter($st, (New-Object System.Text.UTF8Encoding($false)))
        $sw.Write($xml); $sw.Flush(); $sw.Close(); $st.Close()
      }
      $zip.Dispose()
    } catch {
      Write-Warning ("no se pudo corregir 'modificado por' en el .docx: " + $_.Exception.Message)
    }
  }

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
  try { if ($userPrevio) { $word.UserName = $userPrevio } } catch { }
  try { $word.Quit() } catch { }
  Remove-Item $tmpMd, $tmpDoc -ErrorAction SilentlyContinue
}
