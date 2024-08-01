Option Explicit

' Declarar las rutas de origen y destino
Dim objWMIService, colProcesses, objProcess, processesList, i
Dim processInfo, processesCount

' Conectar al servicio WMI
Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")

' Obtener la colección de procesos
Set colProcesses = objWMIService.ExecQuery("SELECT Name, ProcessId, WorkingSetSize FROM Win32_Process")

' Crear un array dinámico para almacenar la información de los procesos
processesCount = 0
ReDim processesList(processesCount)

' Iterar sobre los procesos para almacenar la información de uso de memoria
For Each objProcess in colProcesses
    On Error Resume Next
    ' Intentar obtener el uso de memoria y manejar errores
    Dim memUsage
    memUsage = 0 ' Inicializar a cero
    If Not IsNull(objProcess.WorkingSetSize) Then
        memUsage = CLng(objProcess.WorkingSetSize)
    End If
    On Error GoTo 0
    
    ' Almacenar la información del proceso en el array
    processesList(processesCount) = Array(objProcess.Name, objProcess.ProcessId, memUsage)
    processesCount = processesCount + 1
    ReDim Preserve processesList(processesCount)
Next

' Ordenar la lista de procesos por uso de memoria en orden descendente
For i = 0 To processesCount - 2
    Dim j
    For j = i + 1 To processesCount - 1
        If processesList(i)(2) < processesList(j)(2) Then
            Dim temp
            temp = processesList(i)
            processesList(i) = processesList(j)
            processesList(j) = temp
        End If
    Next
Next

' Mostrar los 10 principales procesos que más memoria consumen
WScript.Echo "Los 10 procesos que más memoria consumen:"
For i = 0 To 9
    If i < processesCount Then
        processInfo = processesList(i)
        WScript.Echo "Nombre: " & processInfo(0) & ", PID: " & processInfo(1) & ", Uso de memoria: " & FormatNumber(processInfo(2) / 1024 / 1024, 2) & " MB"
    End If
Next

' Liberar objetos
Set processesList = Nothing
Set colProcesses = Nothing
Set objWMIService = Nothing
