' Script VBScript para cerrar procesos en segundo plano que consumen mucha memoria pero no son esenciales

' Crea un objeto para interactuar con el sistema
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process")

' Crear un diccionario para almacenar los procesos de primer plano
Set dictForegroundProcesses = CreateObject("Scripting.Dictionary")

' Obtener la ventana de primer plano
Set objShell = CreateObject("WScript.Shell")
strApp = objShell.AppActivate("")

' Recorrer los procesos para identificar los procesos de primer plano
For Each objProcess In colProcesses
    ' Obtenemos el nombre de la ventana si está en primer plano
    If objShell.AppActivate(objProcess.Name) Then
        dictForegroundProcesses.Add objProcess.ProcessId, objProcess.Name
    End If
Next

' Configurar el umbral de uso de memoria (en KB) para determinar los procesos que consumen mucha memoria
Const MemoryThreshold = 100000 ' 100 MB en KB

' Recorrer nuevamente los procesos para identificar y cerrar los procesos de segundo plano que consumen mucha memoria
For Each objProcess In colProcesses
    ' Si el proceso no está en el diccionario de procesos de primer plano y consume más memoria que el umbral
    If Not dictForegroundProcesses.Exists(objProcess.ProcessId) Then
        ' Obtener el uso de memoria del proceso
        memUsage = objProcess.WorkingSetSize / 1024 ' Convertir de bytes a KB
        
        If memUsage > MemoryThreshold Then
            ' Intentar cerrar el proceso
            On Error Resume Next
            objProcess.Terminate()
            If Err.Number = 0 Then
                WScript.Echo "Proceso terminado: " & objProcess.Name & " (PID: " & objProcess.ProcessId & "), Uso de memoria: " & memUsage & " KB"
            Else
                WScript.Echo "No se pudo terminar el proceso: " & objProcess.Name & " (PID: " & objProcess.ProcessId & ")"
            End If
            On Error GoTo 0
        End If
    End If
Next

WScript.Echo "Procesos en segundo plano de alto consumo de memoria han sido gestionados."

' Liberar objetos
Set objWMIService = Nothing
Set colProcesses = Nothing
Set dictForegroundProcesses = Nothing
Set objShell = Nothing
