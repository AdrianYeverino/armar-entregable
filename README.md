# armar-entregable

De una nota en Markdown a un `.docx` y un `.pdf` con formato nativo de Word,
listos para entregar. Sin armar el formato a mano y sin que se note que salió
de un archivo de texto.

```
tarea.md  ──►  pandoc  ──►  cuerpo.docx  ──►  Word  ──►  entregable.docx + .pdf
(Obsidian)     aplica       con estilos      pone la portada,
               tus estilos  nativos          inserta el cuerpo,
                                             actualiza el índice,
                                             exporta el PDF
```

Está pensado para trabajos de la UANL FIME, pero el mecanismo sirve para
cualquier formato: **las reglas tipográficas no viven en el código, viven en
archivos de Word que se editan en Word.**

---

## Por qué así y no generando el .docx desde cero

Este proyecto tuvo un primer intento que escribía el XML de OpenXML a mano.
Falló de forma visible: espaciados e interlineados mal. La causa fue concreta —
el script clavaba `w:spacing` en cada párrafo, sobrescribiendo el sistema de
estilos de Word. Un documento hecho a mano deja esos valores vacíos y los
**hereda del estilo**.

| Origen del párrafo | Espaciado en el XML |
|---|---|
| `.docx` hecho a mano en Word | *(hereda del estilo)* |
| pandoc + `--reference-doc` | *(hereda del estilo)* |
| generador de XML (descartado) | `w:after="0" w:line="240"` |

La regla que salió de ahí, y que ordena todo lo demás:

> **No generar formato: heredarlo.** El formato vive en un `.docx` que se edita
> en Word. El script solo *rellena marcadores* y *aplica estilos con nombre*.
> Nunca escribe espaciados, tamaños ni márgenes.

El detalle histórico está en [`docs/historia.md`](docs/historia.md).

---

## Requisitos

| | Por qué | Se versiona |
|---|---|---|
| **Windows** | El pipeline automatiza Word por COM | — |
| **Microsoft Word** | Pone la portada, actualiza el índice, exporta el PDF | No |
| **pandoc** | Convierte el Markdown aplicando los estilos | No |
| **PowerShell 5.1** | El que trae Windows; no hace falta instalar nada | — |

LibreOffice no sustituye a Word aquí: el paso de portada, índice y PDF usa
automatización COM de Word.

---

## Instalación en un equipo nuevo

```powershell
git clone https://github.com/AdrianYeverino/armar-entregable.git C:\FIME-Herramientas
cd C:\FIME-Herramientas
.\instalar.ps1
```

`instalar.ps1` comprueba pandoc y Word, ofrece instalar pandoc con winget y
crea tu `config.json`. Es idempotente: se puede volver a correr cuando sea.

Si PowerShell se niega a ejecutarlo:

```powershell
powershell -ExecutionPolicy Bypass -File .\instalar.ps1
```

A mano, si prefieres:

```powershell
winget install --id JohnMacFarlane.Pandoc -e
copy config.ejemplo.json config.json   # y editar nombre y matrícula
```

### La carpeta del clon da igual

El script resuelve todas sus rutas contra **su propia ubicación**, así que
funciona desde donde lo clones. `C:\FIME-Herramientas` es solo la convención
que uso yo.

### Qué no se sincroniza y por qué

- **`config.json`** — tu nombre y matrícula. Está en el `.gitignore` porque
  este repositorio es público. Cada equipo crea el suyo.
- **pandoc y Word** — se instalan por máquina, no se versionan.
- **Los entregables generados** (`*.docx`, `*.pdf` fuera de `plantillas/`) — se
  rehacen corriendo el script. Los `.docx` de `plantillas/` sí se versionan:
  esos *son* la herramienta.

Todo lo demás —script, plantillas, logos, documentación— viaja en el repo.
Clonas, corres `instalar.ps1` y estás en el mismo punto que en la otra máquina.

---

## Uso

```powershell
# El caso normal: sin portada, con encabezado en cada página
.\armar-entregable.ps1 -Md tarea.md -Out Actividad1.docx `
  -Encabezado "Actividad 1 - IA en The Creator"

# Con la portada estándar de FIME
.\armar-entregable.ps1 -Md tarea.md -Out AF5.docx -Perfil fime

# Leyendo solo una sección de una nota de Obsidian
.\armar-entregable.ps1 `
  -Md "C:\Second-Brain-Yeve\...\Actividad 1 - FODA.md" `
  -Out "EQ3_ACT1_034.docx" -Perfil formemp -Seccion "Documento entregable"

# Con tablas de borde completo y compactas
.\armar-entregable.ps1 -Md nota.md -Out salida.docx `
  -Perfil topicos -Seccion "Documento entregable" -BordesTabla -TablaCompacta
```

Cada corrida deja el `.docx` **y** el `.pdf`. `-SinPdf` omite el segundo.

### Los cuatro perfiles

El formato se elige con `-Perfil`, no pasando rutas a mano.

| Perfil | Portada | Índice | Cuerpo | Títulos |
|---|---|---|---|---|
| `martinez` *(por omisión)* | No, encabezado | No | Arial 12 justificado **1.15** | 16 / 14 / 12 |
| `fime` | Estándar de la facultad | No | Arial 12 justificado 1.5 | 16 / 14 / 12 |
| `formemp` | Propia del profesor | **Sí, automático** | Arial 12 justificado 1.5 | **18 / 16 / 14** |
| `topicos` | Propia del equipo | **Sí, automático** | Arial 12 justificado 1.15 | 16 / 14 / 12 |
| `lbtssi` | Estándar de FIME, con brigada, hora y columna de inscripción | **Sí, automático** | Arial 12 justificado 1.5 | 16 / 14 / 12 |

`-ConPortada` sigue funcionando por compatibilidad: equivale a `-Perfil fime`.

### Todos los parámetros

| Parámetro | Qué hace |
|---|---|
| `-Md` *(obligatorio)* | Markdown de entrada |
| `-Out` *(obligatorio)* | `.docx` de salida. El `.pdf` sale con el mismo nombre |
| `-Perfil` | `martinez` · `fime` · `formemp` · `topicos` · `lbtssi` |
| `-Seccion` | Toma solo lo que está debajo de ese encabezado |
| `-Encabezado` | Texto del encabezado en el perfil `martinez`. Por omisión, el nombre del archivo |
| `-BordesTabla` | Línea negra en todos los bordes de todas las tablas |
| `-TablaCompacta` | Celdas a Arial 10, alineadas a la izquierda y arriba, encabezado repetido entre páginas |
| `-SinPdf` | Solo genera el `.docx` |
| `-Referencia` | Un `reference-*.docx` distinto al del perfil |
| `-Plantilla` | Una portada distinta a la del perfil |
| `-SubirNivel` | Fuerza cuántos niveles suben los encabezados con `-Seccion` |
| `-ConPortada` | Compatibilidad: igual que `-Perfil fime` |

`-TablaCompacta` no es cosmético. Sin él las celdas heredan el justificado del
cuerpo, y en una columna angosta el justificado estira las palabras hasta
hacerlas ilegibles. La primera corrida de un entregable así costó cinco páginas
de más.

---

## El Markdown de entrada

Los datos de la portada van en un bloque anidado `portada:`, para que una nota
de Obsidian pueda llevar su propio frontmatter sin que choque:

```yaml
---
materia: Formacion De Emprendedores      # del vault, lo resuelve Dataview
tipo: tarea
portada:
  materia: Formación de Emprendedores    # el que se imprime en la portada
  actividad: Actividad Fundamental 1
  tema: "FODA"
  equipo: 3
  docente: M.A. Nombre Del Docente
  grupo: 034 / E2026
  semestre: AGOSTO – DICIEMBRE 2026
  plan: 401
  ciudad: San Nicolás de los Garza, N.L
  fecha: 24 de agosto del 2026
  integrantes:
    - 1234567 | Nombre Apellido Apellido | ITS
---
```

Las claves de `portada:` **tienen precedencia** sobre las de primer nivel. Ese
es el punto: `materia` existe dos veces a propósito.

**Equipo o individual lo decide el frontmatter, no un parámetro.** Si trae
`equipo:` se pone el renglón; si no, se omite y el hueco se cierra solo. Lo
mismo con `modalidad:` y `plan:`. El mismo `.md` sirve para los dos casos.

Referencia completa de campos por plantilla:
[`docs/frontmatter.md`](docs/frontmatter.md).

### `-Seccion`: una nota que es a la vez apunte y fuente

Toma solo lo que está debajo de ese encabezado, hasta el siguiente encabezado
del mismo nivel o más alto. Así el entregable vive dentro de la nota de la
tarea, junto al enunciado y a tus dudas, sin que nada de eso acabe en el Word.

Hace dos cosas solo:

- **Sube los encabezados de nivel** para que el primero quede como Título 1. Si
  la sección es `## Documento entregable`, entonces `###` → Título 1, `####` →
  Título 2, `#####` → Título 3.
- **Limpia los bordes del recorte:** líneas en blanco y reglas `---` al
  principio y al final.

> Cuidado: **todo** lo que esté debajo de ese encabezado entra al Word,
> incluidos comentarios y notas al margen. Esas van *arriba* del encabezado.

### Imágenes: el modo secuencial

El problema real de una práctica con cien capturas no es insertarlas: es
**nombrarlas y mantener la correspondencia**, que se rompe en cuanto metes una
captura a la mitad. La solución es **no nombrar nada**.

1. Crea una carpeta `img/` junto al `.md`.
2. Tira ahí las capturas tal como salen, sin renombrar. Windows las nombra
   `Captura 2026-08-17 210101.png`: el orden alfabético ya es el orden en que
   las tomaste.
3. Pon el marcador `![]` donde va cada una.

```markdown
Se colocan los componentes en el área de trabajo.

![]

Se realizan las conexiones entre las terminales.

![Conexión entre el transformador y el puente rectificador]
```

| Sintaxis | Qué hace |
|---|---|
| `![]` | Siguiente imagen de `img/`. **El caso normal** |
| `![Pie de figura]` | Siguiente imagen, con pie |
| `![](ruta/imagen.png)` | Imagen específica (Markdown estándar) |

Después de armar el documento, **Word reescala** las que excedan el ancho de
columna (468 pt) o 11 cm de alto, conservando proporción. Una captura de
1920 px insertada a tamaño nativo mediría 20 pulgadas y se saldría de la hoja.

---

## Cómo cambiar el formato

**Se abre el `.docx` en Word y se editan sus estilos. No se toca ningún
script.** Los cambios aplican a todos los entregables futuros de ese perfil.

```
plantillas/
├── reference-martinez.docx   ← SOLO estilos: 1.15, títulos 16/14/12
├── reference-fime.docx       ← SOLO estilos: 1.5,  títulos 16/14/12
├── reference-formemp.docx    ← SOLO estilos: 1.5,  títulos 18/16/14
├── portada-fime.docx         ← portada estándar, marcadores {{MATERIA}} etc.
├── portada-formemp.docx      ← portada del profesor + índice automático
└── portada-topicos.docx      ← portada del equipo + índice automático
```

Dos tipos de archivo con dos papeles distintos:

- **`reference-*.docx`: solo estilos.** Su contenido se ignora; pandoc lee de
  ellos únicamente Normal, Título 1/2/3 e `Image Caption` (el estilo de los
  pies de figura).
- **`portada-*.docx`: la portada real, recortada.** Salieron de tomar un
  entregable real, recortarlo a su primera página y sustituir los valores por
  marcadores `{{...}}`. **Conservan los logos originales con su posición y
  tamaño exactos**, porque nunca se regeneraron.

> **Trampa que ya costó un bug:** los estilos de título los define la
> **plantilla de portada**, no el `reference-doc`. Al insertar el cuerpo, Word
> resuelve los estilos contra el documento destino. Si un `portada-*.docx` trae
> los Título 1/2/3 de fábrica, los títulos salen en Aptos Display y en el azul
> del tema de Word aunque el `reference-doc` diga otra cosa. Los tres ya están
> redefinidos a Arial negrita negro.

Para crear un perfil nuevo, ver
[`docs/plantillas.md`](docs/plantillas.md).

---

## Estructura del repositorio

```
armar-entregable/
├── armar-entregable.ps1     ← el generador
├── instalar.ps1             ← puesta a punto en un equipo nuevo
├── config.ejemplo.json      ← plantilla de config.json (que no se versiona)
├── plantillas/              ← donde vive el formato: se editan en Word
├── logos/                   ← escudos UANL y FIME, por si hay que rehacer una portada
├── ejemplos/                ← markdown de prueba, incluido el modo secuencial de imágenes
└── docs/
    ├── plantillas.md        ← cómo crear un perfil nuevo
    ├── frontmatter.md       ← todos los campos, por plantilla
    ├── problemas.md         ← errores conocidos y qué significan
    └── historia.md          ← el intento descartado y la lección
```

---

## Prueba rápida

```powershell
.\armar-entregable.ps1 -Md ejemplos\tarea-portada.md -Out prueba.docx -Perfil fime
```

Debe dejar `prueba.docx` y `prueba.pdf` con portada, tres integrantes en la
tabla, cuerpo en Arial 12 justificado y ningún marcador `{{...}}` sin rellenar.

Si algo falla, [`docs/problemas.md`](docs/problemas.md).

---

## Lo que esto no resuelve

El generador se encarga del **formato**, que en una rúbrica típica son 20 de
100 puntos. Los otros 80 son contenido, y ese no sale de aquí.

Automatizar el formato existe para tener más tiempo de pensar el contenido, no
menos.
