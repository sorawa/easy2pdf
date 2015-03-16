f0 = WScript.ScriptFullName
Set fso = CreateObject("Scripting.FileSystemObject")
dp0 = fso.GetAbsolutePathName(f0 & "\..")
msgbox dp0