Set WShell = CreateObject("WScript.Shell")
strPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
WShell.CurrentDirectory = strPath
WShell.Run Chr(34) & strPath & "\venv\Scripts\pythonw.exe" & Chr(34) & " VoskSTT\RDC_Vosk_STT.py", 0, False
Set WShell = Nothing
