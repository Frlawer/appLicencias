# MCP de App Licencias

Este MCP expone el contexto del proyecto y utilidades de solo lectura para que una IA pueda trabajar sobre el repo sin depender de memoria externa.

## Qué expone
- Resumen del proyecto y reglas de trabajo.
- Lista de archivos del workspace.
- Lectura de archivos individuales.
- Búsqueda textual básica en el proyecto.

## Cómo usarlo
1. Instalar dependencias dentro de esta carpeta:

```bash
cd mcp
npm install
```

2. Registrar el servidor en tu cliente MCP con el comando:

```bash
node C:/xampp/htdocs/App Licencias/mcp/server.js
```

3. Si tu cliente soporta recursos MCP, usa:
- `applicencias://summary`
- `applicencias://files`

## Nota
El servidor es de solo lectura. No modifica archivos ni reemplaza `clasp push`.