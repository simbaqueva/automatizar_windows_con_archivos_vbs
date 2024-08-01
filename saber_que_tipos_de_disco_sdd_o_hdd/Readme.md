# Detector de Tipos de Discos

Este script en VBScript está diseñado para identificar y mostrar información sobre los discos físicos en un sistema Windows, incluyendo el nombre del disco, modelo, tipo de disco (HDD, SSD, RAID) y letras de unidad asociadas.

## Descripción

1. **Conexión WMI**: Establece una conexión con el servicio WMI para acceder a la información del sistema.
2. **Consulta de Discos**: Ejecuta una consulta WMI para obtener información sobre los discos físicos.
3. **Procesamiento de Discos**:
   - Recorre cada disco encontrado.
   - Muestra el nombre y modelo del disco.
   - Determina el tipo de disco basándose en MediaType, InterfaceType y Model.
   - Obtiene y muestra las letras de unidad asociadas.
4. **Identificación del Tipo de Disco**:
   - Utiliza lógica de inferencia para determinar si el disco es HDD, SSD o RAID.
   - Maneja casos donde el tipo no puede ser identificado con certeza.
5. **Presentación de Resultados**: Muestra la información recopilada para cada disco.

## Instrucciones de Uso

1. **Guardar el script**: Guarda el código en un archivo con extensión `.vbs`, por ejemplo, `detector_discos.vbs`.
2. **Ejecutar el script**: Ejecuta el archivo `.vbs` haciendo doble clic en él o desde la línea de comandos usando `cscript.exe nombre_del_script.vbs`.

## Advertencias

- Asegúrate de tener los permisos necesarios para ejecutar consultas WMI en el sistema.
- La precisión en la identificación del tipo de disco depende de la información proporcionada por WMI y puede variar según el hardware.

## Requisitos

- Sistema operativo Windows
- Permisos para ejecutar consultas WMI

## Licencia

Este script se proporciona sin garantía. Utilízalo bajo tu propio riesgo.