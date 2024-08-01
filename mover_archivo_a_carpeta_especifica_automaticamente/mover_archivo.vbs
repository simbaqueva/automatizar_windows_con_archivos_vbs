Option Explicit

' Declarar las rutas de origen y destino
Dim rutaOrigen, rutaDestino
rutaOrigen = "C:\Users\csarmiento\Downloads"
rutaDestino = "Z:\Publicidad1\Alex 2022\alex 2022\FOLLETOS 2024\JUNIO\20 AL 26 DE JUNIO"

' Crear el objeto FileSystemObject
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

' Verificar si la carpeta de origen existe
If Not fso.FolderExists(rutaOrigen) Then
    WScript.Echo "Error: La carpeta de origen no existe."
    WScript.Quit
End If

' Verificar si la carpeta de destino existe, si no, crearla
If Not fso.FolderExists(rutaDestino) Then
    On Error Resume Next
    fso.CreateFolder(rutaDestino)
    If Err.Number <> 0 Then
        WScript.Echo "Error: No se pudo crear la carpeta de destino. " & Err.Description
        WScript.Quit
    End If
    On Error GoTo 0
End If

' Obtener la colección de archivos en la carpeta de origen
Dim folder, files, file
Set folder = fso.GetFolder(rutaOrigen)
Set files = folder.Files

' Inicializar variables para el archivo más reciente
Dim archivoMasReciente, fechaMasReciente
fechaMasReciente = #1/1/1970# ' Fecha muy antigua

' Buscar el archivo más reciente
For Each file In files
    If file.DateLastModified > fechaMasReciente Then
        Set archivoMasReciente = file
        fechaMasReciente = file.DateLastModified
    End If
Next

' Verificar si se encontró un archivo
If Not archivoMasReciente Is Nothing Then
    ' Mover el archivo más reciente a la carpeta de destino
    Dim rutaArchivoDestino
    rutaArchivoDestino = rutaDestino & "\" & archivoMasReciente.Name

    On Error Resume Next
    archivoMasReciente.Move rutaArchivoDestino
    If Err.Number = 0 Then
        WScript.Echo "Archivo movido con éxito a " & rutaArchivoDestino
    Else
        WScript.Echo "Error al mover el archivo: " & Err.Description
    End If
    On Error GoTo 0
Else
    WScript.Echo "No se encontró ningún archivo en la carpeta de origen."
End If

' Liberar objetos
Set archivoMasReciente = Nothing
Set files = Nothing
Set folder = Nothing
Set fso = Nothing
