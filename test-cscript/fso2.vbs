
Set oShell = CreateObject("WScript.Shell")
'FileName = oShell.ExpandEnvironmentStrings("%Temp%\d.txt")
Set oShell = Nothing 'Tidy up the Objects we no longer need

'This bit creates the file
Set fso = CreateObject("Scripting.FileSystemObject")

Set oFile = fso.CreateTextFile("abc")

oFile.WriteLine "here is my content"
oFile.Close

Set oFile = Nothing
Set fso = Nothing
