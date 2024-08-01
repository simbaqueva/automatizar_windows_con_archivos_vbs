# Script VBS para Cerrar Procesos Secundarios

Este script en VBScript está diseñado para cerrar procesos secundarios innecesarios en el sistema, manteniendo los procesos críticos y el proceso de la ventana en primer plano.

## Descripción

1. **Definir Procesos Críticos**: El script define una lista de procesos críticos que no deben cerrarse, como `explorer.exe`, `svchost.exe`, y otros esenciales para el funcionamiento del sistema.
2. **Incluir Proceso en Primer Plano**: Obtiene el nombre del proceso que está asociado con la ventana en primer plano y lo agrega a la lista de procesos críticos para asegurarse de que no se cierre.
3. **Cerrar Procesos Secundarios**: Itera sobre todos los procesos en ejecución y cierra aquellos que no están en la lista de procesos críticos.
4. **Manejo de Errores**: Intenta cerrar los procesos secundarios y maneja posibles errores durante el proceso.

## Instrucciones de Uso

1. **Guardar el script**: Guarda el código en un archivo con extensión `.vbs`, por ejemplo, `cerrar_procesos_secundarios.vbs`.
2. **Ejecutar el script**: Ejecuta el archivo `.vbs` haciendo doble clic en él o desde la línea de comandos.

## Advertencias

- Asegúrate de que los procesos críticos definidos en el script sean correctos para tu sistema.
- El script intentará cerrar procesos sin pedir confirmación adicional. Utilízalo con precaución para evitar cerrar procesos importantes accidentalmente.

## Licencia

Este script se proporciona sin garantía. Utilízalo bajo tu propio riesgo.
