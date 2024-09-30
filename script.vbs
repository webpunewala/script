Dim objFSO, objShell, desktopPath, newFolder

' Create FileSystemObject to handle file operations
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Create Shell object to get special folders like Desktop
Set objShell = CreateObject("WScript.Shell")

' Get the path to the Desktop
desktopPath = objShell.SpecialFolders("Desktop")

' Define the new folder path
newFolder = desktopPath & "\Vishal"

' Check if the folder already exists
If Not objFSO.FolderExists(newFolder) Then
    ' Create the folder if it doesn't exist
    objFSO.CreateFolder(newFolder)
    WScript.Echo "Folder 'Vishal' created on Desktop."
Else
    WScript.Echo "Folder 'Vishal' already exists on Desktop."
End If
