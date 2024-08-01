Dim rutaCarpeta
rutaCarpeta = "E:\BACKUP 2024 DE COMPARTIDAS\RECURSOS LIDER"  ' Aquí especifica la ruta de la carpeta que deseas abrir

' Crear un objeto Shell
Set objShell = CreateObject("Shell.Application")

' Abrir la carpeta especificada
objShell.Open rutaCarpeta
