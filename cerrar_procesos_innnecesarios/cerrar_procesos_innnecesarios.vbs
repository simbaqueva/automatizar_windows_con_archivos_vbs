Option Explicit

' Crear objetos WMI y Shell
Dim objWMIService, objShell
Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
Set objShell = CreateObject("WScript.Shell")

' Definir las rutas de los procesos críticos que no deben cerrarse
Dim criticalProcesses
criticalProcesses = Array("explorer.exe", "svchost.exe", "lsass.exe", "winlogon.exe", _
                          "csrss.exe", "services.exe", "smss.exe", "taskmgr.exe", _
                          "spoolsv.exe", "dwm.exe")

' Obtener el proceso de la ventana en primer plano
Dim foregroundProcessName
foregroundProcessName = GetForegroundProcess()

' Agregar el proceso de la ventana en primer plano a la lista de procesos críticos
ReDim Preserve criticalProcesses(UBound(criticalProcesses) + 1)
criticalProcesses(UBound(criticalProcesses)) = foregroundProcessName

' Obtener la colección de todos los procesos en el sistema
Dim colProcesses, objProcess
Set colProcesses = objWMIService.ExecQuery("SELECT Name, ProcessId FROM Win32_Process")

' Iterar sobre los procesos y cerrarlos si no están en la lista de procesos críticos
For Each objProcess in colProcesses
    If Not IsInArray(LCase(objProcess.Name), criticalProcesses) Then
        On Error Resume Next
        ' Intentar cerrar el proceso
        objProcess.Terminate
        If Err.Number <> 0 Then
            WScript.Echo "Error al cerrar el proceso: " & objProcess.Name & " (PID: " & objProcess.ProcessId & ")"
            Err.Clear
        End If
        On Error GoTo 0
    End If
Next

WScript.Echo "Procesos secundarios innecesarios han sido cerrados."

' Función para verificar si un valor está en un array
Function IsInArray(val, arr)
    Dim i
    IsInArray = False
    For i = LBound(arr) To UBound(arr)
        If LCase(arr(i)) = val Then
            IsInArray = True
            Exit For
        End If
    Next
End Function

' Función para obtener el nombre del proceso de la ventana en primer plano
Function GetForegroundProcess()
    Dim objWMIService, colProcesses, objProcess, hwnd, pid
    Dim objLocator, objSWbemServices, colItems, objItem
    Dim processName

    ' Crear objeto para WMI
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objSWbemServices = objLocator.ConnectServer(".", "root\CIMV2")

    ' Obtener el handle de la ventana en primer plano
    hwnd = GetForegroundWindow()
    
    ' Obtener el PID del proceso que tiene el handle de la ventana en primer plano
    pid = GetWindowProcessId(hwnd)
    
    ' Obtener el proceso basado en el PID
    Set colItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId = " & pid)

    ' Devolver el nombre del proceso
    For Each objItem in colItems
        processName = objItem.Name
        Exit For
    Next

    GetForegroundProcess = processName
End Function

' Función para obtener el handle de la ventana en primer plano
Function GetForegroundWindow()
    Dim objShell
    Set objShell = CreateObject("WScript.Shell")
    GetForegroundWindow = objShell.Exec("powershell -command ""(Get-Process | Where-Object { $_.MainWindowHandle -ne 0 } | Sort-Object MainWindowHandle -Descending | Select-Object -First 1).MainWindowHandle""").StdOut.ReadLine
End Function

' Función para obtener el PID de un handle de ventana
Function GetWindowProcessId(hwnd)
    Dim objWMIService, colProcesses, objProcess
    Dim pid

    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Handle = '" & hwnd & "'")

    For Each objProcess in colProcesses
        pid = objProcess.ProcessId
        Exit For
    Next

    GetWindowProcessId = pid
End Function

' Liberar objetos
Set colProcesses = Nothing
Set objWMIService = Nothing
Set objShell = Nothing
