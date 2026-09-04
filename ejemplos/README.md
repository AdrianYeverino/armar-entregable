# Ejemplos

Markdown de prueba para verificar que la instalación quedó bien. Los `.docx` y
`.pdf` que generen **no se versionan**: se rehacen corriendo el script.

| Archivo | Para qué |
|---|---|
| `tarea-portada.md` | Frontmatter completo con bloque `portada:`, tres integrantes, listas y una tabla |
| `cuerpo-prueba.md` | Solo cuerpo, sin frontmatter. Para el perfil `martinez` |
| `con-imagenes.md` | El modo secuencial de imágenes, con la carpeta `img/` |

## Prueba de cada perfil

```powershell
# Portada estándar de FIME
..\armar-entregable.ps1 -Md tarea-portada.md -Out prueba-fime.docx -Perfil fime

# Sin portada, con encabezado en cada página
..\armar-entregable.ps1 -Md cuerpo-prueba.md -Out prueba-martinez.docx `
  -Encabezado "Actividad 1 - Prueba"

# Modo secuencial de imágenes
..\armar-entregable.ps1 -Md con-imagenes.md -Out prueba-imagenes.docx
```

## Qué revisar en el resultado

- Ningún marcador `{{...}}` sin rellenar en la portada.
- La tabla de integrantes con cuatro filas: encabezado y tres personas.
- Cuerpo en Arial 12 justificado, con el interlineado que toca al perfil.
- Títulos en Arial negrita **negro**, no azul.
- En `prueba-imagenes.docx`, las dos capturas en el orden del archivo y
  reescaladas para caber en la hoja.
- El `.pdf` junto a cada `.docx`.

Los datos de `tarea-portada.md` son ficticios a propósito: el repositorio es
público.
