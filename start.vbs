Set shell = CreateObject("WScript.Shell")
scriptPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) & "\rightBtnMenu.ps1"
shell.Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File """ & scriptPath & """", 0, False
