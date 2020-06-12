
' if want to debug under cscript.
' isWsh = true

isWsh = false

if isWsh then
 Set oWSH = CreateObject("WScript.Shell")
 vbsInterpreter = "cscript.exe"
 Call ForceConsole()
end if

Function ForceConsole()
   If InStr(LCase(WScript.FullName), vbsInterpreter) = 0 Then
       oWSH.Run vbsInterpreter & " //NoLogo " & Chr(34) & WScript.ScriptFullName & Chr(34)
       WScript.Quit
   End If
End Function

Function printf(txt)
  if isWsh Then
    WScript.StdOut.WriteLine txt
  else
    AddDebugString(txt)
  end if
End Function

Function printl(txt)
  if isWsh Then
    WScript.StdOut.WriteLine txt
  else
    AddDebugString(txt)
  end if
End Function

Function scanf()
  'scanf = LCase(WScript.StdIn.ReadLine)
End Function

Function wait(n)
  'WScript.Sleep Int(n * 1000)
End Function

Function cls()
  For i = 1 To 5
      printf ""
  Next
End Function

function regvalue(fc, address)
  if (isWsh) Then
    regvalue = 0
  else
    regvalue = getregistervalue(fc,address)
  end if
end function

Function setreg(fc, address, val)
  if (isWsh) Then
    printf "fc:" & fc & ",address: " & address & ", val:" & val
  else
    SetRegisterValue fc,address, val
  end if
end Function

' tranfer Numberic system
Function ToBase(ByVal n, b)

  ' Handle everything from binary to base 36...
  If b < 2 Or b > 36 Then Exit Function

  Const SYMBOLS = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"

  Do
      ToBase = Mid(SYMBOLS, n Mod b + 1, 1) & ToBase
      n = Int(n / b)
  Loop While n > 0

End Function

'binary string to decimal.
Function BinaryToDecimal(Binary)
  Dim n
  Dim s

  For s = 1 To Len(Binary)
    n = n + (Mid(Binary, Len(Binary) - s + 1, 1) * (2 ^ (s - 1)))
  Next

  BinaryToDecimal = n

End Function

' samas x = ToBase(dec, 2)
Function DecimalToBinary(DecimalNum)
  Dim tmp
  Dim n

  n = DecimalNum

  tmp = Trim(Str(n Mod 2))
  n = n \ 2

  Do While n <> 0
    tmp = Trim(Str(n Mod 2)) & tmp
    n = n \ 2
  Loop

  DecimalToBinary = tmp
End Function



device = "cc81"
printf "device name is: " & device



'word count
w = 15
start = 256
startreg = 0
nstart = start
for i = 0 to (2*w -1)
  'printf "i:" & i

  nend = nstart + 7

  'printf "nstart: " & nstart & ", " & "nend: " & nend
  bstr = ""
  for n=nstart to nend
     'x = getregistervalue(0,n)
     x=regvalue(0,n)
     'printf Cstr(n) & ":" & CStr(x)
     bstr = bstr & CStr(x)
  next
  dec = BinaryToDecimal(bstr)
  'printf "bstr(" & i & "): " & bstr
  'printf "bstr 10 base: " & BinaryToDecimal(bstr)
  'printf ToBase(dec, 10)
  'printf ToBase(dec, 16)
  'printf ToBase(dec, 2)

  setreg 3, i, dec
  nstart = nend + 1
next

printf "ToBase(45,2): " & ToBase(45, 2)
printf "ToBase(45,8): " & ToBase(45, 8)
printf "ToBase(45,16): " & ToBase(45, 16)

printf "END =============="
