# Campos del frontmatter

Todos los campos son opcionales. **Los que no vengan se omiten y el hueco se
cierra solo** — así el mismo `.md` sirve para un trabajo individual y para uno
en equipo, sin parámetros extra.

Las claves dentro de `portada:` **tienen precedencia** sobre las de primer
nivel. Eso permite que una nota de Obsidian lleve su propio frontmatter
(`materia`, `tipo`, `estado`) sin chocar con lo que se imprime en la portada.

El formato viejo, con los campos de portada en el primer nivel, sigue
funcionando.

## Tabla de campos

| Campo | `fime` | `formemp` | `topicos` | `lbtssi` | `martinez` |
|---|:--:|:--:|:--:|:--:|:--:|
| `materia` | ✓ | ✓ | ✓ | ✓ | — |
| `actividad` | ✓ | ✓ | ✓ | ✓ | — |
| `tema` | ✓ | ✓ | — | ✓ | — |
| `equipo` | ✓ | ✓ | ✓ | ✓ | — |
| `docente` | ✓ | ✓ | ✓ *(como Instructor)* | ✓ | — |
| `brigada` | — | — | — | ✓ | — |
| `grupo` | ✓ | ✓ | ✓ | ✓ | — |
| `semestre` | ✓ | ✓ | ✓ | ✓ | — |
| `modalidad` | ✓ | — | — | ✓ | — |
| `plan` | — | ✓ | — | — | — |
| `ciudad` | — | ✓ | ✓ | — | — |
| `hora` | — | — | ✓ | ✓ | — |
| `frecuencia` | — | — | ✓ | — | — |
| `fecha` | ✓ | ✓ | ✓ | ✓ | — |
| `integrantes` | ✓ | ✓ | ✓ | ✓ | — |

El perfil `martinez` no usa portada: el encabezado sale de `config.json` y del
parámetro `-Encabezado`.

## Cómo se combinan algunos campos

No todo va tal cual a la portada. Estas son las excepciones:

| Perfil | Combinación |
|---|---|
| `fime` | `grupo`, `semestre` y `modalidad` se juntan en **un solo renglón** (`{{CURSO}}`), cada uno con su etiqueta. La fecha va en mayúsculas y a la derecha |
| `formemp` | Van en **renglones separados**, cada uno con su etiqueta (`Grupo: `, `Docente: `, `Semestre: `, `Plan: `) |
| `topicos` | `ciudad` y `fecha` se unen como `Ciudad, a fecha`, centrado y en minúsculas. `grupo`, `hora` y `frecuencia` comparten renglón |
| `lbtssi` | Igual que `fime`, pero `brigada` va antes que `grupo` en el renglón de curso y `hora` lleva su propio renglón debajo |

`equipo: 3` se imprime como `Equipo 3`; el número solo, en el frontmatter.

## `integrantes`

Una lista, un renglón por persona, con los tres datos separados por `|`:

```yaml
integrantes:
  - 1234567 | Nombre Apellido Apellido | ITS
  - 7654321 | Otro Nombre Apellido | IMA
```

La tabla de la portada **crece sola**: tres integrantes dan cuatro filas
(encabezado más tres). No hay que tocar la plantilla.

`portada-lbtssi.docx` lleva una cuarta columna, `INSCRITO EN LAB`, así que en
ese perfil cada renglón trae un dato más:

```yaml
integrantes:
  - 1234567 | Nombre Apellido Apellido | ITS | SI
```

Las columnas de más que no se llenen quedan vacías, y los datos de más que no
tengan columna se ignoran. El mismo `.md` sirve para los dos casos.

## Marcadores de cada plantilla

Esto es lo que hay dentro de los `portada-*.docx`. Solo importa si vas a crear
una plantilla nueva — ver [`plantillas.md`](plantillas.md).

| Plantilla | Marcadores |
|---|---|
| `portada-fime.docx` | `{{MATERIA}}` `{{ACTIVIDAD}}` `{{TEMA}}` `{{EQUIPO}}` `{{DOCENTE}}` `{{CURSO}}` `{{FECHA}}` |
| `portada-formemp.docx` | `{{MATERIA}}` `{{ACTIVIDAD}}` `{{TEMA}}` `{{EQUIPO}}` `{{GRUPO}}` `{{DOCENTE}}` `{{SEMESTRE}}` `{{PLAN}}` `{{CIUDAD}}` `{{FECHA}}` |
| `portada-topicos.docx` | `{{MATERIA}}` `{{ACTIVIDAD}}` `{{INSTRUCTOR}}` `{{SEMESTRE}}` `{{GRUPO}}` `{{HORA}}` `{{FRECUENCIA}}` `{{EQUIPO}}` `{{FECHA}}` |
| `portada-lbtssi.docx` | `{{MATERIA}}` `{{ACTIVIDAD}}` `{{TEMA}}` `{{DOCENTE}}` `{{EQUIPO}}` `{{CURSO}}` `{{HORA}}` `{{FECHA}}` |

Más la fila molde `{{C1}} | {{C2}} | {{C3}}` de la tabla de integrantes, que el
script clona una vez por persona y luego borra.

`portada-formemp.docx`, `portada-topicos.docx` y `portada-lbtssi.docx` traen además el
campo TOC del índice automático; el script solo lo refresca.

## Ejemplo completo

```yaml
---
materia: Formacion De Emprendedores      # del vault, sin acentos, para Dataview
tipo: tarea
estado: pendiente
portada:
  materia: Formación de Emprendedores    # con acentos, es lo que se imprime
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

`materia` aparece dos veces **a propósito**: arriba sin acentos, porque así se
llama el archivo del MOC y es lo que Dataview compara; dentro de `portada:` con
acentos, porque es lo que va impreso. Sin la separación, una de las dos tenía
que ceder.
