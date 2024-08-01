# Script VBScript para Cerrar Procesos en Segundo Plano de Alto Consumo de Memoria

Este script en VBScript está diseñado para identificar y cerrar procesos en segundo plano que consumen mucha memoria, mientras se preservan los procesos de primer plano.

## Descripción

1. **Identificación de Procesos de Primer Plano**: El script identifica los procesos asociados con las ventanas de primer plano y los excluye de la terminación.
2. **Umbral de Memoria**: Establece un umbral de uso de memoria (100 MB en KB por defecto) para determinar qué procesos de segundo plano se consideran de alto consumo de memoria.
3. **Terminación de Procesos**: Cierra los procesos en segundo plano que consumen más memoria que el umbral configurado y no están en primer plano.
4. **Manejo de Errores**: Intenta cerrar los procesos y maneja posibles errores durante la terminación.

## Instrucciones de Uso

1. **Guardar el script**: Guarda el código en un archivo con extensión `.vbs`, por ejemplo, `cerrar_procesos_alto_consumo.vbs`.
2. **Ejecutar el script**: Ejecuta el archivo `.vbs` haciendo doble clic en él o desde la línea de comandos.

## Advertencias

- Asegúrate de que el umbral de memoria (`MemoryThreshold`) sea adecuado para tus necesidades. Puedes ajustar el valor según sea necesario.
- El script cierra procesos sin pedir confirmación adicional. Utilízalo con precaución para evitar cerrar procesos importantes accidentalmente.

## Licencia

Este script se proporciona sin garantía. Utilízalo bajo tu propio riesgo.
