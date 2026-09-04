# Historia del proyecto

Por qué está hecho así y qué se descartó. Sirve para no repetir los errores.

## 2026-08-17 — El primer intento, descartado

Dos scripts en Perl (`md2docx-martinez.pl` y `md2docx-portada.pl`) generaban el
`.docx` **escribiendo el XML de OpenXML a mano**.

Falló de forma visible: espaciados e interlineados mal. La causa no fue "el
código no puede hacer archivos nativos", que era la conclusión fácil. Fue más
específica: **el script clavaba `w:spacing` en cada párrafo**, sobrescribiendo
el sistema de estilos. Un `.docx` hecho a mano en Word deja esos valores en
blanco y los hereda del estilo.

```
Portada.docx (hecha a mano)  ->  (hereda del estilo)
generador de XML             ->  w:after="0" w:line="240"
```

Se descubrió descomprimiendo los dos archivos y comparando el `document.xml`
párrafo por párrafo. **No mirando cómo se veían en pantalla.**

> **La regla que salió de ahí:** generar formato desde cero obliga a reinventar
> decisiones tipográficas que Word ya tiene resueltas, y siempre se erra
> alguna. El formato vive en un archivo de Word que se edita en Word; las
> herramientas solo *rellenan* y *aplican estilos con nombre*.

Aplica a cualquier formato nuevo: antes de programarlo, hacer la plantilla en
Word.

Los dos scripts en Perl no están en este repositorio. Contienen datos
personales y no aportan nada que no diga este documento.

## 2026-08-17 — El segundo intento, el actual

Pipeline híbrido en tres pasos, donde **ninguna pieza inventa formato**:

1. **pandoc** convierte el Markdown con `--reference-doc`, un `.docx` que
   contiene solo estilos definidos a mano en Word. pandoc aplica estilos *con
   nombre*; no escribe espaciados.
2. **Word**, por automatización COM, abre la plantilla de portada, sustituye
   los marcadores `{{...}}` e inserta el cuerpo. Word hace el trabajo, así que
   el resultado es nativo por definición.
3. **Word** actualiza el índice y exporta el PDF.

La prueba de que funciona está en el XML: los párrafos generados heredan el
espaciado del estilo, igual que los de un documento hecho a mano.

## 2026-08-23 — Perfiles, secciones y frontmatter anidado

Una segunda máquina y un tercer formato obligaron a tres cambios:

- **Perfiles en vez de rutas a mano.** `-Perfil formemp` en lugar de pasar
  `-Referencia` y `-Plantilla` cada vez.
- **`-Seccion`**, para que la nota de Obsidian sea a la vez apunte y fuente del
  entregable, sin que el enunciado ni las dudas acaben en el Word.
- **Frontmatter anidado `portada:`**, porque `materia` tenía que existir dos
  veces: sin acentos para Dataview, con acentos para la portada. Se resolvió
  con precedencia, no con orden de aparición, que era la solución fácil y
  frágil.

Cuatro bugs salieron al probar de punta a punta, ninguno leyendo el código.
Están en [`problemas.md`](problemas.md).

## 2026-08-26 — El perfil `topicos` y las tablas

Un cuarto perfil y dos modificadores de tabla que sirven para cualquiera.

**La plantilla no se dibujó, se derivó.** `portada-topicos.docx` salió de copiar
`portada-fime.docx` y reescribir sus renglones por COM. Es la misma lección de
agosto 17 un nivel más arriba: así como el `.docx` no se arma escribiendo OOXML
a mano, la portada no se arma volviendo a colocar los logos. Los dos escudos
conservan posición, tamaño y anclaje originales porque nunca se tocaron.

**El frontmatter no cambió de forma.** Los integrantes se siguen escribiendo
`matrícula | nombre | carrera` como en todos los perfiles; el script intercambia
las dos primeras columnas solo para este. Cambiar la convención del frontmatter
por una diferencia de maquetado habría contaminado los cuatro formatos.

Efecto medible de `-TablaCompacta` y los anchos corregidos: el mismo documento
pasó de **18 páginas a 13**. Ninguna palabra menos.

## 2026-08-31 — El bug de los títulos azules

Hasta esa fecha el perfil `fime` reusaba `reference-martinez.docx`, así que los
entregables con portada salían a interlineado 1.15 cuando debían ir a 1.5. Se
creó `reference-fime.docx`.

Y se redefinieron los `Título 1/2/3` dentro de `portada-fime.docx`: los traía de
fábrica, y por eso todos los entregables con portada salían con los títulos en
Aptos Display y en el azul del tema de Word.

## Cómo registrar un cambio futuro

Al modificar el sistema, agregar aquí una entrada con:

1. **Qué se hizo** y **por qué**, en una línea.
2. Los **síntomas** que se vieron, no solo el arreglo.
3. La **causa raíz**, si costó encontrarla.
4. **Si se descartó algo, la causa.** Es lo que evita repetir el error.
