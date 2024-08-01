# Script VBS para Limpiar Archivos Temporales y Cachés Seguros

Este script en VBScript está diseñado para limpiar archivos temporales y cachés de manera segura en ubicaciones comunes del sistema. Está orientado a eliminar archivos en carpetas específicas, minimizando el riesgo de eliminar datos importantes.

## Descripción

1. **Eliminar Archivos y Carpetas Recursivamente**: La función `DeleteFilesAndFolders` elimina archivos y carpetas en una ruta especificada de manera recursiva, manejando errores para evitar interrupciones.
2. **Rutas de Limpieza**:
   - **Temp del Usuario**: `%TEMP%`
   - **Temp del Sistema**: `C:\Windows\Temp`
   - **Cachés de Navegadores**:
     - Google Chrome: `%LOCALAPPDATA%\Google\Chrome\User Data\Default\Cache`
     - Microsoft Edge: `%LOCALAPPDATA%\Microsoft\Edge\User Data\Default\Cache`
   - **Cachés de Aplicaciones**:
     - Discord: `%APPDATA%\discord\Cache`
     - Spotify: `%LOCALAPPDATA%\Spotify\Storage`
   - **Otros Archivos Temporales**:
     - Logs de Windows: `C:\Windows\Logs`
     - Informes de errores de Windows: `C:\ProgramData\Microsoft\Windows\WER\ReportArchive`
     - Caché de thumbnails: `%LOCALAPPDATA%\Microsoft\Windows\Explorer`
3. **Aviso al Usuario**: Al finalizar la limpieza, se muestra un mensaje informativo al usuario.

## Instrucciones de Uso

1. **Guardar el script**: Guarda el código en un archivo con extensión `.vbs`, por ejemplo, `limpiar_archivos_temporales.vbs`.
2. **Ejecutar el script**: Ejecuta el archivo `.vbs` haciendo doble clic en él o desde la línea de comandos.

## Advertencias

- Este script está diseñado para eliminar archivos en ubicaciones comunes seguras. Asegúrate de revisar las rutas especificadas para evitar la eliminación de archivos importantes.
- Las carpetas de caché y temporales especificadas son generalmente seguras para limpiar, pero revisa el contenido si tienes dudas sobre la eliminación de archivos.

## Licencia

Este script se proporciona sin garantía. Utilízalo bajo tu propio riesgo.
