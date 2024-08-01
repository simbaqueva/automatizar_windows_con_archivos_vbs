# Monitor de Procesos de Alto Consumo de Memoria

Este script en VBScript identifica y muestra los 10 procesos que más memoria consumen en un sistema Windows.

## Descripción

1. **Conexión WMI**: Establece una conexión con el servicio WMI para acceder a la información de procesos.
2. **Recopilación de Datos**: Obtiene información sobre todos los procesos en ejecución.
3. **Procesamiento de Datos**:
   - Almacena la información de cada proceso en un array dinámico.
   - Maneja errores potenciales al obtener el uso de memoria.
4. **Ordenamiento**: Ordena los procesos por uso de memoria en orden descendente.
5. **Presentación de Resultados**: Muestra los 10 procesos principales, incluyendo:
   - Nombre del proceso
   - ID del proceso (PID)
   - Uso de memoria en MB

## Instrucciones de Uso

1. **Guardar el script**: Guarda el código en un archivo con extensión `.vbs`, por ejemplo, `monitor_procesos_memoria.vbs`.
2. **Ejecutar el script**: Ejecuta el archivo `.vbs` haciendo doble clic en él o desde la línea de comandos usando `cscript.exe nombre_del_script.vbs`.

## Advertencias

- Asegúrate de tener los permisos necesarios para ejecutar consultas WMI en el sistema.
- El script muestra el uso de memoria en un momento específico, que puede variar rápidamente.

## Requisitos

- Sistema operativo Windows
- Permisos para ejecutar consultas WMI

## Licencia

Este script se proporciona sin garantía. Utilízalo bajo tu propio riesgo.