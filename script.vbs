Option Explicit

Dim fso, folderPath

' Set the folder path and name
folderPath = "C:\Users\Administrator\Desktop\IT_call" ' Desired path

' Create a FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")

' Check if the folder already exists
If fso.FolderExists(folderPath) Then
    WScript.Echo "The folder """ & folderPath & """ already exists."
Else
    ' Create the folder
    On Error Resume Next ' Ignore errors temporarily
    fso.CreateFolder(folderPath)
    If Err.Number <> 0 Then
        WScript.Echo "Failed to create the folder """ & folderPath & """."
        Err.Clear
    Else
        WScript.Echo "Folder """ & folderPath & """ created successfully!"
    End If
    On Error GoTo 0 ' Reset error handling
End If

' Clean up
Set fso = Nothing
