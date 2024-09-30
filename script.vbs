Dim objFSO, objShell, desktopPath, newFile

' Create FileSystemObject to handle file operations
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Create Shell object to get special folders like Desktop
Set objShell = CreateObject("WScript.Shell")

' Get the path to the Desktop
desktopPath = objShell.SpecialFolders("Desktop")

' Define the new file path
newFile = desktopPath & "\Ganesh.txt"

' Check if the file already exists
If Not objFSO.FileExists(newFile) Then
    ' Create the file if it doesn't exist
    objFSO.CreateTextFile(newFile).WriteLine "This is the Ganesh file."
    WScript.Echo "File 'Ganesh.txt' created on Desktop."
Else
    WScript.Echo "File 'Ganesh.txt' already exists on Desktop."
End If
