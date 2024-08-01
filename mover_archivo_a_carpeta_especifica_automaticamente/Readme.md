# Script VBS para Mover el Archivo Más Reciente

Este script en VBScript está diseñado para buscar el archivo más reciente en una carpeta de origen y moverlo a una carpeta de destino. El script verifica la existencia de las carpetas especificadas y maneja errores durante el proceso de movimiento del archivo.

## Descripción

1. **Declaración de Rutas**: Se definen las rutas de la carpeta de origen y de la carpeta de destino.
2. **Verificación de Carpetas**:
   - Verifica si la carpeta de origen existe.
   - Verifica si la carpeta de destino existe; si no, la crea.
3. **Selección del Archivo Más Reciente**:
   - Obtiene la colección de archivos en la carpeta de origen.
   - Busca el archivo más reciente basándose en la fecha de última modificación.
4. **Movimiento del Archivo**:
   - Mueve el archivo más reciente a la carpeta de destino.
   - Maneja errores si ocurre algún problema al mover el archivo.
5. **Mensajes de Estado**:
   - Informa al usuario sobre el éxito o el fracaso del movimiento del archivo.

## Instrucciones de Uso

1. **Guardar el script**: Guarda el código en un archivo con extensión `.vbs`, por ejemplo, `mover_archivo_mas_reciente.vbs`.
2. **Ejecutar el script**: Ejecuta el archivo `.vbs` haciendo doble clic en él o desde la línea de comandos.

## Advertencias

- Asegúrate de que las rutas especificadas en el script sean correctas y de que tengas permisos adecuados para crear carpetas y mover archivos en esas ubicaciones.
- El script está configurado para mover el archivo más reciente basado en la fecha de modificación. Asegúrate de que esto se ajuste a tus necesidades.

## Licencia

Este script se proporciona sin garantía. Utilízalo bajo tu propio riesgo.
