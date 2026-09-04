# Las plantillas: dónde vive el formato

La idea de fondo: **el formato no está en el código.** Está en archivos de Word
que se editan en Word. El script solo rellena marcadores y aplica estilos con
nombre.

## Los dos tipos de archivo

### `reference-*.docx` — solo estilos

Su **contenido se ignora por completo**. pandoc lee de ellos únicamente las
definiciones de estilo:

| Estilo | Qué controla |
|---|---|
| `Normal` | Cuerpo: fuente, tamaño, justificado, interlineado, espacio posterior |
| `Título 1` / `2` / `3` | Los encabezados |
| `Image Caption` | Los pies de figura |
| `Table` | Las tablas, antes de `-BordesTabla` |

Para cambiar cómo salen **todas las tareas futuras** de un perfil: abre el
archivo en Word, entra al panel de estilos, edítalos, guarda. No se toca ningún
script.

Los tres que hay:

| Archivo | Cuerpo | Títulos |
|---|---|---|
| `reference-martinez.docx` | Arial 12 justificado **1.15**, 8 pt posterior | Arial 16/14/12 negrita negro |
| `reference-fime.docx` | Arial 12 justificado **1.5**, 8 pt posterior | Arial 16/14/12 negrita negro |
| `reference-formemp.docx` | Arial 12 justificado **1.5**, 8 pt posterior | Arial **18/16/14** negrita negro |

`reference-formemp.docx` se derivó copiando la de Martínez y ajustándole los
estilos, para partir de un `reference-doc` ya probado con pandoc. Es el camino
recomendado para uno nuevo: **copiar uno que ya funcione**, no hacerlo de cero.

### `portada-*.docx` — la portada real, recortada

No se dibujaron: salieron de **tomar un entregable real ya aceptado, recortarlo
a su primera página y sustituir los valores por marcadores `{{...}}`**.

Por eso conservan los logos originales con su posición y tamaño exactos.
Rehacerlos a mano es justo lo que se descartó al principio del proyecto.

## Crear un perfil nuevo

Digamos que aparece un profesor con otro formato.

**1. Haz la portada en Word, no en código.**
Toma un entregable real de esa materia, bórrale todo menos la primera página y
sustituye cada dato por su marcador: `{{MATERIA}}`, `{{ACTIVIDAD}}`, etc. Usa
los nombres que ya existen ([`frontmatter.md`](frontmatter.md)) siempre que
puedas; solo inventa uno si el dato no existía.

Si la portada lleva tabla de integrantes, deja una fila molde con
`{{C1}} | {{C2}} | {{C3}}`. El script la clona una vez por integrante y luego
borra el molde.

Si lleva índice automático, insértalo en Word como campo TOC. El script solo lo
refresca; no lo crea.

> **Redefine los Título 1/2/3 dentro de la portada.** Este es el error que ya
> costó un bug: al insertar el cuerpo, Word resuelve los estilos contra el
> documento **destino**. Si la portada trae los títulos de fábrica, salen en
> Aptos Display y en el azul del tema de Word aunque el `reference-doc` diga
> Arial negro. Hay que arreglarlo en la portada, no en el reference.

**2. Haz el `reference-*.docx`.**
Copia el más parecido de los tres, ábrelo en Word y ajusta Normal y los
títulos. Su contenido da igual: bórralo o déjalo, no se usa.

**3. Da de alta el perfil en el script.**
Es una línea en la tabla `$PERFILES` de `armar-entregable.ps1`:

```powershell
$PERFILES = @{
  ...
  "nuevo" = @{ portada = $true; ref = "reference-nuevo.docx"; plantilla = "portada-nuevo.docx" }
}
```

Y agregar `"nuevo"` al `[ValidateSet(...)]` del parámetro `-Perfil`, arriba.

**4. Si la portada combina campos de forma distinta**, hay que agregar su bloque
de sustituciones en la sección `3. Word: portada + cuerpo + indice + pdf`. Los
tres que existen sirven de modelo: `fime` junta grupo/semestre/modalidad en un
renglón, `formemp` los pone separados, `topicos` une ciudad y fecha.

**5. Prueba con `ejemplos/tarea-portada.md`** antes de usarlo en algo real. Que
no quede ningún `{{...}}` sin rellenar.

## Truco: copiar una portada por COM en vez de rehacerla

`portada-topicos.docx` salió de **copiar `portada-fime.docx` y reescribir sus
renglones por automatización COM**, no de colocar los logos otra vez. Los dos
escudos son los mismos y así conservaron su posición y tamaño originales.

Si el formato nuevo comparte los logos, ese es el camino más seguro.

## Los logos

`logos/logo-uanl.png` y `logos/logo-fime.png` están en el repositorio solo por
si alguna vez hay que armar una portada desde cero. **El flujo normal no los
usa**: las portadas ya los traen incrustados.
