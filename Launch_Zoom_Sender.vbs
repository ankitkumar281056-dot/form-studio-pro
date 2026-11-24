Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
strPath = fso.GetParentFolderName(WScript.ScriptFullName)
WshShell.CurrentDirectory = strPath

' Launch directly via pythonw.exe / pyw.exe (window style 0 = completely hidden, False = do not wait)
On Error Resume Next
Err.Clear
WshShell.Run "pyw -3 " & chr(34) & strPath & "\zoom_auto_sender.pyw" & chr(34) & " --detached-gui", 0, False
If Err.Number <> 0 Then
    Err.Clear
    WshShell.Run "pythonw " & chr(34) & strPath & "\zoom_auto_sender.pyw" & chr(34) & " --detached-gui", 0, False
End If
If Err.Number <> 0 Then
    Err.Clear
    WshShell.Run "py -3 " & chr(34) & strPath & "\zoom_auto_sender.py" & chr(34), 0, False
End If
If Err.Number <> 0 Then
    Err.Clear
    WshShell.Run "python " & chr(34) & strPath & "\zoom_auto_sender.py" & chr(34), 0, False
End If
Set WshShell = Nothing
