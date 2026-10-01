Set sh = CreateObject("WScript.Shell")
d = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
sh.Run """" & d & "\auto_push.bat""", 0, False
