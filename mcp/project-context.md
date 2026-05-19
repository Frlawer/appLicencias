# Contexto del proyecto App Licencias

## Qué es
Web App de Google Apps Script para gestión de licencias, justificaciones, novedades, jornada y razones particulares.

## Estructura funcional
- `funciones.js`: punto de entrada de la Web App, `doGet`, guardado de solicitudes y consultas públicas.
- `admin.service.gs` / `admin.service.js`: consultas y acciones administrativas sobre solicitudes, justificaciones y reportes.
- `agentes.service.gs`: búsqueda de agentes, cargos y emails desde hojas externas.
- `siso.service.gs`: carga y filtrado de datos CSV de SISO.
- `email.service.gs`: creación de borradores de Gmail a partir de plantillas.
- `novedades.service.gs`: registro de actividad reciente.
- `auth.admin.gs`: whitelist administrativa por propiedad `ADMIN_ALLOWED_EMAILS`.
- `config.gs`: configuración central desde `ScriptProperties` con valores por defecto.
- HTML: `main.html`, `web.html`, `admin.html` y parciales `admin.*.html`, `header.html`, `footer.html`, `css.html`, `js.html`.

## Reglas importantes
- Es un proyecto de Apps Script, no un backend Node.
- El despliegue web está en `doGet`; el admin se abre con `?page=admin`.
- El acceso administrativo puede restringirse por whitelist.
- Usa `SpreadsheetApp`, `DriveApp` y `GmailApp`, por lo que cualquier cambio debe cuidar permisos y serialización.
- `.claspignore` excluye los duplicados `.js` que acompañan a los `.gs` para evitar conflictos.

## Criterios para la IA
- Priorizar cambios pequeños y compatibles con Apps Script V8.
- No romper el flujo de `clasp push`.
- Evitar introducir dependencias innecesarias en el runtime principal.
- Revisar siempre hojas, IDs y plantillas antes de cambiar lógica administrativa.