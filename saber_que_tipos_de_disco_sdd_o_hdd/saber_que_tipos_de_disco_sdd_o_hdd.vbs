' Crea un objeto WMI para acceder a la información del sistema
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

' Ejecuta la consulta WMI para obtener los discos físicos
Set colDisks = objWMIService.ExecQuery("SELECT * FROM Win32_DiskDrive")

' Recorre cada disco encontrado
For Each objDisk in colDisks
    ' Muestra el nombre del disco y el modelo
    WScript.Echo "Nombre: " & objDisk.Name
    WScript.Echo "Modelo: " & objDisk.Model

    ' Determina el tipo de medio utilizando la propiedad MediaType o BusType si está disponible
    On Error Resume Next
    mediaType = objDisk.MediaType
    busType = objDisk.InterfaceType

    ' Utiliza la propiedad MediaType o InterfaceType para inferir el tipo de disco
    If mediaType = "Fixed hard disk media" Or mediaType = "Removable Media" Then
        WScript.Echo "Tipo: Disco Mecánico (HDD)"
    ElseIf InStr(busType, "IDE") > 0 Or InStr(busType, "SCSI") > 0 Or InStr(busType, "SATA") > 0 Then
        WScript.Echo "Tipo: Disco Mecánico (HDD)"
    ElseIf InStr(busType, "RAID") > 0 Then
        WScript.Echo "Tipo: Disco RAID (Depende de la configuración)"
    ElseIf InStr(busType, "NVMe") > 0 Or InStr(busType, "USB") > 0 Or InStr(busType, "SD") > 0 Then
        WScript.Echo "Tipo: Disco de Estado Sólido (SSD)"
    ElseIf mediaType = "Solid State Disk" Then
        WScript.Echo "Tipo: Disco de Estado Sólido (SSD)"
    Else
        ' Verificación adicional para SSD basado en la interfaz o por exclusión
        If InStr(objDisk.Model, "SSD") > 0 Then
            WScript.Echo "Tipo: Disco de Estado Sólido (SSD)"
        ElseIf InStr(objDisk.Model, "HDD") > 0 Then
            WScript.Echo "Tipo: Disco Mecánico (HDD)"
        Else
            WScript.Echo "Tipo: No identificado (Podría ser HDD o SSD)"
        End If
    End If
    
    ' Obtener la letra de unidad del disco
    For Each partition In objDisk.Partitions_
        For Each logicalDisk In partition.Associators_("Win32_LogicalDisk")
            WScript.Echo "Letra de unidad: " & logicalDisk.DeviceID
        Next
    Next
    
    WScript.Echo "------------------------------------"
Next

' Limpieza
Set colDisks = Nothing
Set objWMIService = Nothing
