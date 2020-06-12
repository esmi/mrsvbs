     CLASS cCONSOLE
 '= =================================================================
 '= 
 '=    This class provides automatic switch to CScript and has methods
 '=    to write to and read from the CSCript console. It transparently
 '=    switches to CScript if the script has been started in WScript.
 '=
 '= =================================================================

    Private oOUT
    Private oIN


    Private Sub Class_Initialize()
    '= Run on creation of the cCONSOLE object, checks for cScript operation


        '= Check to make sure we are running under CScript, if not restart
        '= then run using CScript and terminate this instance.
        dim oShell
        set oShell = CreateObject("WScript.Shell")

        If InStr( LCase( WScript.FullName ), "cscript.exe" ) = 0 Then
            '= Not running under CSCRIPT

            '= Get the arguments on the command line and build an argument list
            dim ArgList, IX
            ArgList = ""

            For IX = 0 to wscript.arguments.count - 1
                '= Add the argument to the list, enclosing it in quotes
                argList = argList & " """ & wscript.arguments.item(IX) & """"
            next

            '= Now restart with CScript and terminate this instance
            oShell.Run "cscript.exe //NoLogo """ & WScript.ScriptName & """ " & arglist
            WScript.Quit

        End If

        '= Running under CScript so OK to continue
        set oShell = Nothing

        '= Save references to stdout and stdin for use with Print, Read and Prompt
        set oOUT = WScript.StdOut
        set oIN = WScript.StdIn

        '= Print out the startup box 
            StartBox
            BoxLine Wscript.ScriptName
            BoxLine "Started at " & Now()
            EndBox


    End Sub

    '= Utility methods for writing a box to the console with text in it

            Public Sub StartBox()

                Print "  " & String(73, "_") 
                Print " |" & Space(73) & "|"
            End Sub

            Public Sub BoxLine(sText)

                Print Left(" |" & Centre( sText, 74) , 75) & "|"
            End Sub

            Public Sub EndBox()
                Print " |" & String(73, "_") & "|"
                Print ""
            End Sub

            Public Sub Box(sMsg)
                StartBox
                BoxLine sMsg
                EndBox
            End Sub

    '= END OF Box utility methods


            '= Utility to center given text padded out to a certain width of text
            '= assuming font is monospaced
            Public Function Centre(sText, nWidth)
                dim iLen
                iLen = len(sText)

                '= Check for overflow
                if ilen > nwidth then Centre = sText : exit Function

                '= Calculate padding either side
                iLen = ( nWidth - iLen ) / 2

                '= Generate text with padding
                Centre = left( space(iLen) & sText & space(ilen), nWidth )
            End Function



    '= Method to write a line of text to the console
    Public Sub Print( sText )

        oOUT.WriteLine sText
    End Sub

    '= Method to prompt user input from the console with a message
    Public Function Prompt( sText )
        oOUT.Write sText
        Prompt = Read()
    End Function

    '= Method to read input from the console with no prompting
    Public Function Read()
        Read = oIN.ReadLine
    End Function

    '= Method to provide wait for n seconds
    Public Sub Wait(nSeconds)
        WScript.Sleep  nSeconds * 1000 
    End Sub

    '= Method to pause for user to continue
    Public Sub Pause
        Prompt "Hit enter to continue..."
    End Sub


 END CLASS