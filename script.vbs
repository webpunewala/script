Dim shell
Dim batFilePath

' Set the path to your batch file
batFilePath = "C:\Users\Administrator\Desktop\master.bat"

' Create a WScript.Shell object
Set shell = CreateObject("WScript.Shell")

' Check if the batch file exists
If CreateObject("Scripting.FileSystemObject").FileExists(batFilePath) Then
    ' Run the batch file
    shell.Run """" & batFilePath & """", 1, True
    MsgBox "Batch file executed successfully!"
Else
    MsgBox "Batch file not found: " & batFilePath
End If

' Clean up
Set shell = Nothing
