Dim ErrorFlag
'Create a shell object
Set Shell = WScript.CreateObject("WScript.Shell")
'Create a file system object
Set FSO = WScript.CreateObject("Scripting.FileSystemObject")
'Get path to Diablo II EXE from registry
Reg1 = "HKLM\Software\Blizzard Entertainment\Diablo II"
Reg2 = "HKCU\Software\Blizzard Entertainment\Diablo II"
Reg1Use = True
Reg2Use = True
'Ensure script gets to clean up if setting up and launching the mod fails
On Error Resume Next
Set Folder = FSO.GetFolder(".")
D2Path = Folder.Path & "\"

Shell.RegWrite "HKCU\Software\Blizzard Entertainment\Diablo II\InstallPath", Folder.Path, "REG_SZ"
Shell.RegWrite "HKLM\Software\Blizzard Entertainment\Diablo II\InstallPath", Folder.Path, "REG_SZ"
Shell.RegWrite "HKCU\Software\Wow6432Node\Blizzard Entertainment\Diablo II\InstallPath", Folder.Path, "REG_SZ"
Shell.RegWrite "HKLM\Software\Wow6432Node\Blizzard Entertainment\Diablo II\InstallPath", Folder.Path, "REG_SZ"
WScript.Quit(0)
