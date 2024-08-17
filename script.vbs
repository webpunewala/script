
' Define the directory path
directoryPath = "C:\Users\Administrator\Desktop"

' Create a FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

' Check if the directory exists
If fso.FolderExists(directoryPath) Then
    ' Get the folder object
    Set folder = fso.GetFolder(directoryPath)
    
    ' Loop through each file in the folder
    For Each file In folder.Files
        ' Delete the file
        file.Delete True
    Next
    
    MsgBox "All files in the folder have been deleted.", vbInformation
Else
    MsgBox "Directory not found. Please check the path and try again.", vbExclamation
End If
