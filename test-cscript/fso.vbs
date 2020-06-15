Set oWSH = CreateObject("WScript.Shell")
vbsInterpreter = "cscript.exe"

Set dict = CreateObject("Scripting.Dictionary")
Set file = fso.OpenTextFile ("c:\test.txt", 1)
row = 0
Do Until file.AtEndOfStream
  line = file.Readline
  dict.Add row, line
  row = row + 1
Loop

file.Close

'Loop over it
For Each line in dict.Items
  WScript.Echo line
Next
