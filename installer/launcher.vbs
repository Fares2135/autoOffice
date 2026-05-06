Dim sh, fso, dir
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")
dir = fso.GetParentFolderName(WScript.ScriptFullName)
sh.Run Chr(34) & dir & "\autoOffice-server.exe" & Chr(34), 0, False
