' Script VBS para abrir carpetas de archivos temporales y cachés

Dim objShell
Set objShell = CreateObject("WScript.Shell")

' Lista de carpetas comunes para limpiar

' Carpeta Temp del Usuario
objShell.Run "explorer.exe " & objShell.ExpandEnvironmentStrings("%TEMP%")

' Carpeta Temp del Sistema
objShell.Run "explorer.exe " & "C:\Windows\Temp"

' Carpetas de caché de navegadores comunes

' Caché de Google Chrome
objShell.Run "explorer.exe " & objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%\Google\Chrome\User Data\Default\Cache")

' Caché de Microsoft Edge
objShell.Run "explorer.exe " & objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%\Microsoft\Edge\User Data\Default\Cache")

' Caché de Mozilla Firefox
objShell.Run "explorer.exe " & objShell.ExpandEnvironmentStrings("%APPDATA%\Mozilla\Firefox\Profiles")
' Nota: Las carpetas de perfiles de Firefox pueden tener nombres diferentes, el usuario debe seleccionar el perfil correcto.

' Carpetas de caché de aplicaciones comunes

' Caché de Microsoft Office
objShell.Run "explorer.exe " & objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%\Microsoft\Office\16.0\OfficeFileCache")

' Caché de Discord
objShell.Run "explorer.exe " & objShell.ExpandEnvironmentStrings("%APPDATA%\discord\Cache")

' Caché de Spotify
objShell.Run "explorer.exe " & objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%\Spotify\Storage")

' Carpetas de archivos temporales de otros programas

' Carpetas de logs de Windows (a veces contienen mucha información innecesaria)
objShell.Run "explorer.exe " & "C:\Windows\Logs"

' Carpeta de informes de errores de Windows
objShell.Run "explorer.exe " & "C:\ProgramData\Microsoft\Windows\WER\ReportArchive"

' Carpeta de descargas (revisión manual sugerida)
objShell.Run "explorer.exe " & objShell.ExpandEnvironmentStrings("%USERPROFILE%\Downloads")

' Carpetas de prefetch
objShell.Run "explorer.exe " & "C:\Windows\Prefetch"

' Carpetas de caché de thumbnails
objShell.Run "explorer.exe " & objShell.ExpandEnvironmentStrings("%LOCALAPPDATA%\Microsoft\Windows\Explorer")

' Aviso al usuario
MsgBox "Las carpetas de archivos temporales y cachés se han abierto. Revisa y limpia manualmente según sea necesario.", vbInformation, "Limpieza de Sistema"

' Liberar el objeto shell
Set objShell = Nothing
