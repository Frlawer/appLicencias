# Dashboard Administrativo - Sistema de Licencias CPEM 25

## 🔐 Acceso al Dashboard

Para acceder al panel administrativo, usa la URL de tu Web App agregando el parámetro `?page=admin`:

```
https://script.google.com/macros/s/TU_DEPLOYMENT_ID/exec?page=admin
```

**Importante:** Solo usuarios con permisos de edición en el script de Google Apps Script podrán acceder.

## 📊 Funcionalidades del Dashboard

### 1. **Gestión de Solicitudes**
- ✅ Ver todas las solicitudes de licencias
- ✅ Filtrar por estado (Pendiente, Autorizado, Justificado)
- ✅ Editar el estado de cada solicitud
- ✅ Ver detalles completos (docente, fechas, tipo de licencia, curso/cargo)
- ✅ Enviar emails personalizados a los docentes

### 2. **Gestión de Justificaciones**
- ✅ Ver todas las justificaciones cargadas
- ✅ Acceder a los archivos justificatorios (click en el link)
- ✅ Ver IDs de solicitudes asociadas
- ✅ Enviar emails de seguimiento a los docentes

### 3. **Comunicación con Docentes**
- ✅ Modal para redactar emails personalizados
- ✅ Creación automática de borradores en Gmail
- ✅ Plantilla HTML profesional

## 🎨 Características de Diseño

- **Menú lateral** con navegación entre Solicitudes y Justificaciones
- **Tablas responsivas** con scroll horizontal
- **Estados con colores**:
  - 🟡 Pendiente (amarillo)
  - 🟢 Autorizado (verde)
  - 🔵 Justificado (azul)
- **Modales** para edición y envío de emails
- **Botón de recarga** para actualizar datos en tiempo real

## 📋 Pasos para Desplegar

1. **Subir código:**
   ```bash
   clasp push
   ```

2. **Desplegar Web App:**
   - En Apps Script: Deploy > Manage deployments
   - New deployment
   - Type: Web app
   - Execute as: Me
   - Who has access: Anyone (el script verificará permisos internamente)

3. **Copiar URL del deployment** y agregar `?page=admin`

4. **Opcional - Personalizar plantilla de email:**
   - En `funciones.js`, localiza `PLANTILLA_ADMIN`
   - Si tienes un borrador de Gmail con plantilla, reemplaza el ID

## 🔧 Configuración Adicional

### Email con Plantilla Personalizada

Si quieres usar un borrador de Gmail como plantilla:

1. Crea un borrador en Gmail con tu diseño
2. En la consola de Apps Script, ejecuta:
   ```javascript
   function obtenerIDsBorradores() {
     const borradores = GmailApp.getDrafts();
     borradores.forEach(b => {
       Logger.log('Asunto: ' + b.getMessage().getSubject() + ' | ID: ' + b.getId());
     });
   }
   ```
3. Copia el ID del borrador que quieres usar
4. Reemplaza en `funciones.js`:
   ```javascript
   const PLANTILLA_ADMIN = 'r-TU_ID_AQUI';
   ```

## 🎯 Mejoras Futuras Sugeridas

- [ ] Filtros avanzados por fecha, docente, tipo de licencia
- [ ] Exportar a Excel/PDF
- [ ] Estadísticas y gráficos
- [ ] Notificaciones automáticas
- [ ] Historial de cambios de estado
- [ ] Búsqueda en tiempo real

## 🐛 Solución de Problemas

**Problema:** "No se cargan los datos"
- **Solución:** Verifica que las hojas "Solicitudes" y "Justificaciones" existan en el spreadsheet activo

**Problema:** "No puedo acceder al dashboard"
- **Solución:** Asegúrate de tener permisos de edición en el script de Apps Script

**Problema:** "Los emails no se envían"
- **Solución:** Revisa los permisos de Gmail en el script y verifica que el borrador se cree correctamente

---

**Desarrollado para CPEM N° 25**  
Sistema de Gestión de Licencias Docentes
