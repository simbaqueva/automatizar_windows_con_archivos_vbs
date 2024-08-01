' Script VBS para eliminar archivos temporales y cachés

Dim objShell, fso
Set objShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Función para eliminar archivos y carpetas recursivamente
Sub DeleteFilesAndFolders(path)
    On Error Resume Next ' Ignorar errores para evitar que el script se detenga si encuentra archivos en uso
    If fso.FolderExists(path) Then
        For Each file In fso.GetFolder(path).Files
            fso.DeleteFile file, True
        Next
        For Each folder In fso.GetFolder(path).SubFolders
            fso.DeleteFolder folder, True
        Next
    End If
    On Error GoTo 0 ' Volver a la gestión normal de errores
End Sub

' Lista de carpetas comunes para limpiar

' Carpeta Temp del Usuario
DeleteFilesAndFolders objShell.ExpandEnvironmentStrings("%TEMP%")

' Carpeta Temp del Sistema
DeleteFilesAndFolders "C:\Windows\Temp"

' Carpetas de caché de navegadores comunes

' Caché de Google Chrome
DeleteFilesAndFolders objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%\Google\Chrome\User Data\Default\Cache")

' Caché de Microsoft Edge
DeleteFilesAndFolders objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%\Microsoft\Edge\User Data\Default\Cache")

' Caché de Mozilla Firefox
' La ruta es genérica ya que los nombres de las carpetas de perfiles pueden variar
Dim firefoxProfilePath
firefoxProfilePath = objShell.ExpandEnvironmentStrings("%APPDATA%\Mozilla\Firefox\Profiles")
If fso.FolderExists(firefoxProfilePath) Then
    For Each profileFolder In fso.GetFolder(firefoxProfilePath).SubFolders
        DeleteFilesAndFolders profileFolder & "\cache2"
    Next
End If

' Carpetas de caché de aplicaciones comunes

' Caché de Microsoft Office
DeleteFilesAndFolders objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%\Microsoft\Office\16.0\OfficeFileCache")

' Caché de Discord
DeleteFilesAndFolders objShell.ExpandEnvironmentStrings("%APPDATA%\discord\Cache")

' Caché de Spotify
DeleteFilesAndFolders objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%\Spotify\Storage")

' Carpetas de archivos temporales de otros programas

' Carpetas de logs de Windows
DeleteFilesAndFolders "C:\Windows\Logs"

' Carpeta de informes de errores de Windows
DeleteFilesAndFolders "C:\ProgramData\Microsoft\Windows\WER\ReportArchive"

' Carpetas de descargas (eliminar sólo archivos seleccionados)
' Nota: Se eliminarán todos los archivos en la carpeta de descargas
DeleteFilesAndFolders objShell.ExpandEnvironmentStrings("%USERPROFILE%\Downloads")

' Carpetas de prefetch
DeleteFilesAndFolders "C:\Windows\Prefetch"

' Carpetas de caché de thumbnails
DeleteFilesAndFolders objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%\Microsoft\Windows\Explorer")

' Aviso al usuario
MsgBox "La limpieza de archivos temporales y cachés se ha completado.", vbInformation, "Limpieza de Sistema"

' Liberar los objetos
Set fso = Nothing
Set objShell = Nothing
