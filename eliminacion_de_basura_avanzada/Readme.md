# Script VBS para Eliminar Archivos Temporales y Cachés

Este script en VBScript está diseñado para eliminar archivos temporales y cachés en varias ubicaciones comunes del sistema.

## Descripción

1. **Eliminar Archivos y Carpetas Recursivamente**: La función `DeleteFilesAndFolders` elimina archivos y carpetas en una ruta especificada de manera recursiva.
2. **Rutas de Limpieza**:
   - **Temp del Usuario**: `%TEMP%`
   - **Temp del Sistema**: `C:\Windows\Temp`
   - **Cachés de Navegadores**:
     - Google Chrome: `%LOCALAPPDATA%\Google\Chrome\User Data\Default\Cache`
     - Microsoft Edge: `%LOCALAPPDATA%\Microsoft\Edge\User Data\Default\Cache`
     - Mozilla Firefox: Caché en perfiles, ruta genérica `%APPDATA%\Mozilla\Firefox\Profiles`
   - **Cachés de Aplicaciones**:
     - Microsoft Office: `%LOCALAPPDATA%\Microsoft\Office\16.0\OfficeFileCache`
     - Discord: `%APPDATA%\discord\Cache`
     - Spotify: `%LOCALAPPDATA%\Spotify\Storage`
   - **Otros Archivos Temporales**:
     - Logs de Windows: `C:\Windows\Logs`
     - Informes de errores de Windows: `C:\ProgramData\Microsoft\Windows\WER\ReportArchive`
     - Descargas: `%USERPROFILE%\Downloads`
     - Prefetch: `C:\Windows\Prefetch`
     - Caché de thumbnails: `%LOCALAPPDATA%\Microsoft\Windows\Explorer`
3. **Aviso al Usuario**: Al finalizar la limpieza, se muestra un mensaje informativo al usuario.

## Instrucciones de Uso

1. **Guardar el script**: Guarda el código en un archivo con extensión `.vbs`, por ejemplo, `eliminar_archivos_temporales.vbs`.
2. **Ejecutar el script**: Ejecuta el archivo `.vbs` haciendo doble clic en él o desde la línea de comandos.

## Advertencias

- El script elimina archivos y carpetas de manera recursiva. Asegúrate de revisar las rutas especificadas para evitar la eliminación de archivos importantes.
- La carpeta de Descargas se limpia completamente; asegúrate de revisar su contenido antes de ejecutar el script.

## Licencia

Este script se proporciona sin garantía. Utilízalo bajo tu propio riesgo.
