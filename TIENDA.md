# Material para la Chrome Web Store

## Nombre
Subir Notas - EducarEcuador

## Descripción corta (máx. 132 caracteres)
Sube calificaciones a EducarEcuador desde tu Excel/CSV emparejando por cédula o nombre. Todo local; tú inicias sesión.

## Descripción larga
¿Cansado de escribir nota por nota en la plataforma del Ministerio de Educación?

**Subir Notas** agrega un panel en la pantalla de Calificación ordinaria de EducarEcuador
para que cargues tu propio Excel (o CSV) y la extensión coloque y guarde cada nota por ti.

Funciona con CUALQUIER formato de archivo gracias a un paso de mapeo:
• Elige la hoja del Excel y la fila de encabezados.
• Indica qué columna es el nombre del estudiante y cuál es la nota a subir
  (la materia y el trimestre que estés llenando ahora).
• Opcionalmente, indica la columna de cédula para emparejar de forma más segura.

Características:
• Emparejado por cédula (recomendado) o por nombre (tolera acentos, la Ñ y el orden).
• Tabla de revisión antes de guardar: marca en rojo a quien le falta nota.
• "Llenar solo esta página" para probar, o "Llenar y guardar todas las páginas".
• Pausas configurables y botón Detener.
• Cambia la columna de la nota para pasar a la siguiente materia sin volver a subir el Excel.

Privacidad primero:
• Todo se procesa en TU navegador. No se sube tu archivo ni las notas a ningún servidor.
• La extensión NUNCA pide ni guarda tu usuario o contraseña: tú inicias sesión.
• No intenta evadir las verificaciones de seguridad del Ministerio.

Importante: tú eres responsable de revisar las notas antes de guardarlas.
Esta extensión no está afiliada al Ministerio de Educación del Ecuador.

## Categoría sugerida
Productividad (Productivity)

## Idioma
Español (Ecuador) / Español

## Justificación de permisos (para la revisión de Google)
• host_permissions academico.educarecuador.gob.ec: la extensión solo actúa en la página
  de calificaciones; necesita inyectar el panel y rellenar los campos ahí.
• No se solicitan otros permisos. El archivo se lee con <input type=file> + FileReader en
  el navegador; no hay peticiones de red salientes.

## Política de privacidad (URL)
Publica PRIVACIDAD.md en una URL pública (por ejemplo GitHub Pages) y pégala aquí.

## Capturas (recomendadas 1280x800)
1. Panel abierto sobre la tabla de calificaciones.
2. Paso de mapeo de columnas.
3. Tabla de revisión con un "sin dato" en rojo.

## Checklist antes de publicar
[ ] Reemplazar iconos placeholder (icons/) por el logo definitivo.
[ ] Completar correo de soporte en PRIVACIDAD.md y en BRAND.soporte (content.js).
[ ] Publicar la política de privacidad en una URL.
[ ] Comprimir la carpeta en .zip SIN node_modules ni tools/ (opcionales) y subirla.
[ ] Verificar que vendor/xlsx.full.min.js está incluido.
