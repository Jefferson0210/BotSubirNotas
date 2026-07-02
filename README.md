# Subir Notas · EducarEcuador — extensión Chrome/Edge (Manifest V3)

Producto para docentes de Ecuador: inyecta un panel en la pantalla de **Calificación
ordinaria** de EducarEcuador para subir notas desde **tu propio Excel/CSV** (cualquier
formato) emparejando por **cédula** o por **nombre**.

> 🔒 **Privacidad:** todo se procesa en tu navegador. No se sube tu archivo ni las notas a
> ningún servidor. La extensión **nunca** pide ni guarda tu usuario/contraseña: **el login
> lo haces tú** en la página oficial, y no intenta evadir el "Comprueba que eres tú".

## Estructura

```
BotSubirNotas/
├─ manifest.json                 Manifest V3, permisos mínimos
├─ content.js                    Panel + Excel/CSV + mapeo + emparejado + llenado/paginación
├─ panel.css                     Estilos (colores desde BRAND, vía variables CSS)
├─ vendor/xlsx.full.min.js       SheetJS (lee .xlsx/.xls)
├─ icons/icon16|48|128.png       Iconos placeholder (reemplázalos por tu logo)
├─ tools/generate-icons.cjs      Regenera los iconos placeholder
├─ tools/generar-licencia.cjs    Genera claves de licencia (si activas la licencia)
├─ ejemplo-notas.csv             CSV de prueba
├─ PRIVACIDAD.md  ·  TIENDA.md   Política de privacidad y material para la Web Store
└─ README.md
```

## Cargar en modo desarrollador

1. `chrome://extensions` (o `edge://extensions`).
2. Activa **Modo de desarrollador**.
3. **Cargar descomprimida** → selecciona la carpeta `BotSubirNotas`.
4. Tras cualquier cambio de código, pulsa **Recargar** (⟳) en la tarjeta de la extensión
   y **refresca** la página de EducarEcuador.

## Usar (la extensión se autoconfigura)

1. Inicia sesión **tú** y entra a la **pantalla de ingreso de notas** de una asignatura
   (la que tiene las pestañas TRIMESTRE 1/2/3 y la tabla de estudiantes).
2. Abre el botón flotante **📋 Subir Notas** (abajo a la derecha). Arriba verás
   **"Estás llenando: MATERIA · CURSO · Trimestre N"** detectado de la propia página.
3. **1 · Sube tu archivo** (.xlsx/.xls/.csv). La extensión **encuentra sola**:
   - la **hoja** que corresponde a la materia (por nombre/encabezados, con sinónimos:
     CCNN, LENGUA, MATE, EESS, ECA, EEFF, INGLÉS…),
   - la columna del **Nombre**, la **Cédula** (si existe) y las columnas de **promedio de
     T1/T2/T3** (bloques "PRIMER/SEGUNDO/TERCER TRIMESTRE" o etiquetas "PROM 1ER…").
4. Al subir el archivo, la extensión **lee en silencio la lista completa del curso** (pasa
   una vez por las páginas de la plataforma, "Leyendo estudiantes… página 3 de 5", y vuelve
   a la 1). Tarda ~5–8 s para 5 páginas; ajustable con *Pausa al leer la lista* en opciones.
5. **Revisión de TODO el curso a la vez:** tabla #/Estudiante/Nota (editable)/✓ con scroll
   dentro del panel (los 21, no solo los 5 de la página). Edita cualquier nota ahí mismo;
   los **sin dato** salen en rojo. Si en tu Excel hay alguien que **no está en el curso**,
   se avisa abajo (no se sube).
   - Arriba: selector **Trimestre a subir** (Trimestre 1/2/3). Arranca en el de la pestaña
     activa de la página; al cambiarlo se rearma la tabla con esa columna.
   - Si el trimestre elegido **no coincide** con la pestaña abierta en la página, sale un
     **aviso** y se **bloquea** el botón de subir hasta que coincidan.
6. ¿La materia/hoja no calzó? Enlace **"Elegir manualmente"**: cambia hoja, fila de datos o
   reasigna Nombre / Cédula / T1 / T2 / T3.
7. **✅ Listo, subir todas las notas**: va sola a la página 1, llena+guarda fila por fila,
   pasa de página con "Siguiente" hasta "Página X de X", con progreso en vivo
   ("Guardando… 14 de 21"). **Probar con 1 estudiante** guarda solo el primero como prueba.
   **■ Detener** para en el acto.

## Emparejado y normalización

- **Nombre** (principal): se normaliza sin acentos, en minúsculas y con espacios colapsados;
  la Ñ no rompe nada. **No se reordenan los tokens** (para no confundir homónimos): se
  compara el nombre completo.
- **Cédula** (si el archivo la trae): empareja por los 10 dígitos (más seguro); si una no
  calza, intenta por nombre.
- **Notas**: la coma decimal se convierte a punto (`7,74 → 7.74`).

## Personalizar marca (branding)

Arriba de [content.js](content.js) está la constante `BRAND` (nombre, subtítulo, emoji y
colores) y `LICENCIA`. El CSS toma los colores desde variables, así que con cambiar `BRAND`
se reestiliza el panel. Reemplaza los PNG de `icons/` por tu logo (mismos nombres/tamaños),
o ejecuta `node tools/generate-icons.cjs` para regenerar los placeholder.

## Licencia (opcional, barrera básica)

`LICENCIA.activa = false` por defecto. Si la pones en `true`, el panel exige una clave para
habilitar los botones de llenado y la guarda en `localStorage`.

```
node tools/generar-licencia.cjs            # clave nueva
node tools/generar-licencia.cjs ABCD 2345  # clave a partir de 2 grupos
```

⚠ Una licencia dentro de una extensión es **fácil de saltar** (el código corre en el equipo
del usuario). Sirve solo como barrera básica para activar a quien paga, **no** como
seguridad real.

## Publicar en la Chrome Web Store

Ver [TIENDA.md](TIENDA.md) (nombre, descripción, permisos justificados, capturas y
checklist) y [PRIVACIDAD.md](PRIVACIDAD.md) (publícala en una URL y enlázala en la ficha).

## Si quieres soporte solo CSV (sin Excel)

Borra `vendor/xlsx.full.min.js` y su entrada en `manifest.json`; el CSV seguirá funcionando.
