' VBScript to created shortcut
Const strProgramTitle = "Shortcut to Calculator"
Const strProgram = "%SystemRoot%\System32\calc.exe"
Const strWorkDir = "%USERPROFILE%"
Dim objShortcut, objShell
Set objShell = WScript.CreateObject ("Wscript.Shell")
strLPath = objShell.SpecialFolders ("Desktop")
Set objShortcut = objShell.CreateShortcut (strLPath & "\" & strProgramTitle & ".lnk")
objShortcut.TargetPath = strProgram
objShortcut.WorkingDirectory = strWorkDir
objShortcut.Description = strProgramTitle
objShortcut.Save
WScript.Quit
