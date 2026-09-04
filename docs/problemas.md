# Problemas conocidos

Todos estos salieron corriendo el pipeline de punta a punta, no leyendo el
código. Están ordenados por probabilidad de que te toquen.

## Errores al ejecutar

### `No encuentro pandoc.exe`

No está instalado, o está fuera del `PATH` y de `%LOCALAPPDATA%\Pandoc\`.

```powershell
winget install --id JohnMacFarlane.Pandoc -e
```

Después hay que **cerrar y volver a abrir PowerShell** para que tome el `PATH`.

### `No se puede cargar el archivo ... porque la ejecución de scripts está deshabilitada`

Política de ejecución de PowerShell. Para una corrida:

```powershell
powershell -ExecutionPolicy Bypass -File .\armar-entregable.ps1 -Md tarea.md -Out salida.docx
```

Permanente, solo para tu usuario:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

### `the current selection is outside of a block-level XML element`

La plantilla termina con el campo TOC, así que `Content.End` cae **dentro** del
campo y no se puede escribir ahí. El script ya inserta un párrafo antes
(`Content.InsertParagraphAfter()`). Si aparece con una plantilla nueva, es que
esa plantilla necesita el mismo tratamiento.

### `This member cannot be accessed on a horizontal line`

Hay una regla `---` en el Markdown que pandoc convirtió en línea horizontal, y
Word truena al recorrer los párrafos. Con `-Seccion` el script ya limpia las
reglas de los dos bordes del recorte; si está a media sección, quítala.

### `El archivo está en uso` o `no se puede acceder`

Dos causas:

- El `.docx` está abierto en Word. Ciérralo.
- Quedó un `WINWORD.EXE` huérfano de una corrida que falló a la mitad. También
  deja archivos de bloqueo `~$*.docx`.

```powershell
Stop-Process -Name WINWORD -Force -ErrorAction SilentlyContinue
```

### `No encuentro la seccion '...'`

El texto de `-Seccion` tiene que coincidir con el encabezado **sin los `#`** y
sin distinguir espacios sobrantes. `-Seccion "Documento entregable"` para
`## Documento entregable`.

## Cosas que no fallan, pero salen mal

Estas son peores que un error, porque no avisan.

### El entregable aparece en la carpeta equivocada

`[System.IO.Path]::GetFullPath` resuelve las rutas relativas contra el
directorio del **proceso**, que no es la ubicación de PowerShell. Ya está
arreglado combinando con `$PWD.Path`, pero si mueves esa parte del código,
vuelve. Ante la duda, pasa `-Out` con ruta absoluta.

### Los acentos salen como `JesÃºs` o `SAN NICOLÃS`

**PowerShell 5.1 lee los `.ps1` como ANSI si no traen BOM UTF-8.** Los archivos
del repositorio ya lo traen. Si editas el script con un editor que lo quita,
hay que reponerlo:

```powershell
$p = ".\armar-entregable.ps1"
$t = [System.IO.File]::ReadAllText($p)
[System.IO.File]::WriteAllText($p, $t, (New-Object System.Text.UTF8Encoding($true)))
```

Esta trampa ya mordió dos veces. Si tocas un `.ps1`, compruébalo.

### Los títulos salen en azul y con otra tipografía

Al insertar el cuerpo, Word resuelve los estilos **contra el documento
destino**: los `Título 1/2/3` de la portada le ganan a los del `reference-doc`.
Si la portada trae los de fábrica, salen en Aptos Display y en el azul del tema.

Se arregla **en la portada**, no en el reference: abrir el `portada-*.docx` en
Word y redefinir Título 1/2/3 a Arial negrita negro. Los tres que vienen en el
repositorio ya están corregidos.

Este fue el bug más caro de ver, porque el archivo donde estaba el problema no
era el que se estaba mirando.

### Las palabras se estiran dentro de las tablas

Los párrafos de las celdas heredan el **justificado** del cuerpo, y en una
columna angosta el justificado separa las palabras hasta el absurdo. Cuatro
líneas para un nombre.

Se arregla con `-TablaCompacta`. En un entregable real eso solo bajó el
documento de 18 páginas a 13, sin quitar una palabra.

### La portada se desborda a una segunda página

La tabla de integrantes creció más de lo que cabe. Los anchos de columna están
en el script como porcentaje del ancho útil (52 / 26 / 22 %); si tu plantilla
necesita otros, ahí se ajustan.

### El índice sale vacío o con las páginas mal

El script refresca el campo TOC, pero si el documento no tiene ningún párrafo
con estilo Título 1/2/3, no hay nada que indexar. Revisa que el cuerpo use
encabezados Markdown y, con `-Seccion`, que la subida de niveles esté dejando
un Título 1.

### Quedaron marcadores `{{...}}` sin rellenar

Falta el campo en el frontmatter. Los campos vacíos se omiten y el hueco se
cierra, pero un marcador que no está en la tabla de sustituciones del perfil se
queda tal cual. Ver [`frontmatter.md`](frontmatter.md).

## Notas de implementación

Por si tocas el script.

- **Los estilos de Word están localizados.** El Word de esta máquina está en
  español (`Título 1`), así que buscar `"Heading 1"` falla. Se direccionan por
  constante numérica: `Styles.Item(-1)` = Normal, `-2/-3/-4` = Título 1/2/3.
- **Geometría.** Carta 8.5×11" con márgenes de 1" da una columna útil de
  **468 pt**. De ahí salen el tope de ancho de las imágenes (468 pt), el de
  alto (312 pt ≈ 11 cm) y las tabulaciones del encabezado (centrada en 234,
  derecha en 468).
- **`Quit()` no acepta argumentos por COM desde PowerShell 5.1.** El bloque
  `finally` cierra el documento sin guardar *antes* de llamarlo.
- **Verifica leyendo el archivo, no confiando en que se generó bien.**
  Descomprimir el `.docx` y mirar el `document.xml` es lo que detectó el
  problema de espaciado que hundió el primer intento.
