Option Explicit

Public Const SPECIFIERS As String = "bBhHsScCdoxXeEfgGaAtT%n"
Public Const FLAG_CHARS As String = "-+ #"
Public Const DIGITS As String = ".0123456789"



' VBA Implementation of the Java printf function.
' https://docs.oracle.com/javase/8/docs/api/java/util/Formatter.html
' https://docs.oracle.com/javase/tutorial/java/data/numberformat.html
Function printf(ByRef s As String, ParamArray p()) As String
    Dim sRet As String
    Dim sSpec As String, sFlags As String, sWidth As String, sLength As String, sPrecision As String, sWidthPrec As String
    Dim sToken As String
    Dim iLen As Integer ' Original String length
    Dim iPos As Integer ' current position in the original string
    Dim iDigitStartPos As Integer ' start position in the original string of any digit series
    Dim bLengthComplete As Boolean
    Dim bArgIndexFound As Boolean, bFlagFound As Boolean, bDigitFound As Boolean
    Dim iTokenNum As Integer ' nth token in the string
    Dim iParamItem As Integer ' normally  equal to iTokenNum, but could be overridden
    Dim iParamOverrideNum As Integer
    iParamItem = 0
    iParamOverrideNum = -1
    Dim iComp As Integer
    Dim iVarType As Integer
    Dim iWidth As Integer, iPrecision As Integer
    
    
    s = Replace(s, "%%", Chr$(1) & Chr$(1)) ' escape any double percents. Use two Chr$(1) to keep the new length the same as the original.
    sRet = s
    iPos = 1
    
    Do
        sSpec = ""
        sFlags = ""
        sWidth = ""
        sLength = ""
        sPrecision = ""
        sWidthPrec = ""
        bLengthComplete = False
        bDigitFound = False
        iPos = InStr(iPos, s, "%") ' where was the token start "%" in the original string, starting at the end of the last token
        If iPos = 0 Then
            Exit Do ' no more tokens
        End If
        
        Dim i As Integer
        Dim thisChar As String
        
        iParamItem = iTokenNum
        
        ' find the entire token
        For i = iPos + 1 To Len(s)
            thisChar = Mid(s, i, 1)
            If (thisChar = "$") Then
                bArgIndexFound = True
                Dim ss As String
                ss = Mid(s, iPos + 1, i - 1 - iPos)
                iParamItem = CInt(ss) - 1
                
                ' bDigitFound was set to true when the first digit was found. Need to set this back to false
                ' so that we can find the width & precision if it is present
                bDigitFound = False
            ElseIf (InStr(1, FLAG_CHARS, thisChar, vbBinaryCompare)) Then
                bFlagFound = True
                sFlags = sFlags & Mid(s, i, 1)
            ElseIf (InStr(DIGITS, thisChar)) Then
                ' we foud a digit. It could be either for the argument number or for the width/precision
                If (bDigitFound = False) Then
                    bDigitFound = True
                    sWidthPrec = ""
                    'iDigitStartPos = i
                End If
                sWidthPrec = sWidthPrec & Mid(s, i, 1)
            ElseIf (InStr(1, SPECIFIERS, thisChar, vbBinaryCompare)) Then
                sSpec = Mid(s, i, 1)
                If (UCase(sSpec) = "T") Then
                    sSpec = sSpec & Mid(s, i + 1, 1)
                End If
                Exit For
            Else
                Debug.Print "Error. Invalid character found after token: " & thisChar
            End If
        Next
        ' i now holds the position of the specifier in the original string. Everything between the % and specifier needs to be evaluated
        sToken = Mid(s, iPos, i - iPos + Len(sSpec))
        
        If sWidthPrec = "" Then sWidthPrec = "0"
        
        If (CDbl(sWidthPrec) <> 0) Then
            sWidthPrec = ConvertToZeroes(sWidthPrec)
        End If
        
        If (InStr(sWidthPrec, ".")) Then
            iWidth = InStr(sWidthPrec, ".") - 1
        Else
            iWidth = Len(sWidthPrec)
        End If

        Dim j As Integer
        j = InStr(sRet, sToken)
        Dim s1 As String, s2 As String
        s1 = Left(sRet, j - 1)
        s2 = Mid(s, i + Len(sSpec))
        Dim sCase As String
        sCase = Left(UCase(sSpec), 1)
        
        iVarType = VarType(p(iParamItem))

        Select Case sCase
            Case Is = "B" ' Boolean. If the argument arg is null, then the result is "false".
            'If arg is a boolean or Boolean, then the result is the string returned by String.valueOf(arg). Otherwise, the result is "true".
                If (iVarType = 10) Then
                    sRet = "false"
                ElseIf (IsNull(p(iParamItem)) Or p(iParamItem) = vbNull) Then
                    sRet = "false"
                Else
                    sRet = CStr(CBool(p(iParamItem)))
                End If
                
                iComp = StrComp(sSpec, "b", 0)
                If (iComp = 0) Then
                    sRet = LCase(sRet)
                Else
                    sRet = UCase(sRet)
                End If
                
                If (InStr(sFlags, "#")) Then
                    sRet = "ERROR. Cannot use # flag with %B"
                End If
                
            Case Is = "H" ' Hexadecimal. If the argument arg is null, then the result is "null".
            ' Otherwise, the result is obtained by invoking Integer.toHexString(arg.hashCode()).
                
                If (IsNull(p(iParamItem)) Or p(iParamItem) = vbNull) Then
                    sRet = "null"
                Else
                    sRet = Hex(CInt(p(iParamItem)))
                End If
                
                iComp = StrComp(sSpec, "h", 0)
                If (iComp = 0) Then
                    sRet = LCase(sRet)
                Else
                    sRet = UCase(sRet)
                End If
            
                If (InStr(sFlags, "#")) Then
                    sRet = "ERROR. Cannot use # flag with %H"
                End If
            
            Case Is = "S" ' Strings. If the argument arg is null, then the result is "null".
            ' If arg implements Formattable, then arg.formatTo is invoked. Otherwise, the result is obtained by invoking arg.toString().
                On Error GoTo SErrorHandler
                If (IsNull(p(iParamItem)) Or p(iParamItem) = vbNull) Then
                    sRet = "null"
                Else
                    sRet = p(iParamItem)
                End If
                
                GoTo SCont
SErrorHandler:
                sRet = "null"
SCont:
                On Error GoTo 0
                
                iComp = StrComp(sSpec, "s", 0)
                If (iComp = 1) Then
                    sRet = UCase(sRet)
                End If
                
            Case Is = "C" ' Unicode single character
                iComp = StrComp(sSpec, "c", 0)
                iVarType = VarType(p(iParamItem))
                sRet = Replace(p(iParamItem), "\u", "&H")
                
                If (iVarType = vbString) Then
                    sRet = Left(sRet, 1)
                ElseIf (iComp = 0) Then
                    sRet = ChrW(sRet)
                Else
                    sRet = UCase(ChrW(sRet))
                End If
                
                If (InStr(sFlags, "#")) Then
                    sRet = "ERROR. Cannot use # flag with %H"
                End If
                
            Case Is = "D" ' Decimals
                sRet = Format(p(iParamItem), sWidthPrec)
                ' there should not be a precision specified. If there is, send an error
                If ((InStr(sWidthPrec, ".") > 0) And (InStr(sWidthPrec, ".") < Len(sWidthPrec))) Then
                    sRet = "Error. Should not specify precision with %D specifier"
                End If
                
                If (InStr(sFlags, ",")) Then
                    sRet = Format(sRet, ",")
                End If
                
                If (InStr(sFlags, "+") And p(iParamItem) >= 0) Then
                    sRet = "+" & sRet
                End If
                
                If (InStr(sFlags, " ") And p(iParamItem) >= 0) Then
                    sRet = " " & sRet
                End If
                
                If (InStr(sFlags, "(") And p(iParamItem) < 0) Then
                    sRet = Replace("(" & sRet & ")", "-", "")
                End If
                
                If (InStr(sFlags, "#")) Then
                    sRet = "ERROR. Cannot use # flag with %D"
                End If
                
                If (InStr(sFlags, "+") And InStr(sFlags, " ")) Then
                    sRet = "ERROR. Cannot use PLUS and SPACE flags together with %D"
                End If
                
            Case Is = "O" ' Octal
                If (IsNull(p(iParamItem)) Or p(iParamItem) = vbNull) Then
                    sRet = "null"
                Else
                    sRet = Oct(CInt(p(iParamItem)))
                End If
                
                If (InStr(sFlags, "#")) Then
                    sRet = "0" & sRet
                End If
                
                If (InStr(sFlags, "+")) Then
                    sRet = "+" & sRet
                End If
                
                If (InStr(sFlags, "(")) Then
                    sRet = "ERROR. Cannot use ( flag with %O"
                End If
                
                If (InStr(sFlags, " ")) Then
                    sRet = "ERROR. Cannot use SPACE flag with %O"
                End If
                
                If (InStr(sFlags, ",")) Then
                    sRet = "ERROR. Cannot use COMMA flag with %O"
                End If
        
            Case Is = "X" ' Hexadecimal
                sRet = Hex(CInt(p(iParamItem)))
                
                iComp = StrComp(sSpec, "h", 0)
                
                If (InStr(sFlags, "#")) Then
                    sRet = "0x" & sRet
                End If
                
                If (iComp = 0) Then
                    sRet = LCase(sRet)
                Else
                    sRet = UCase(sRet)
                End If
                
                If (InStr(sFlags, "+")) Then
                    sRet = "+" & sRet
                End If
                
                If (InStr(sFlags, "(")) Then
                    sRet = "ERROR. Cannot use ( flag with %X"
                End If
                
                If (InStr(sFlags, " ")) Then
                    sRet = "ERROR. Cannot use SPACE flag with %X"
                End If
                
                If (InStr(sFlags, ",")) Then
                    sRet = "ERROR. Cannot use COMMA flag with %X"
                End If
            
            Case Is = "E" ' Scientific
                ' we can specify the width & precision in two ways. "3.2" or "000.00". Both of these are
                ' equivalent. The VBA "Format" function wants the "000.00" format, so check to see which
                ' one was used. Convert sWidthPrec to a double and see if it equals 0. If not, then the
                ' 3.2 format was used.
                If sWidthPrec = "" Then sWidthPrec = "0"
                If (CDbl(sWidthPrec) <> 0) Then
                    sWidthPrec = ConvertToZeroes(sWidthPrec)
                End If
                
                sRet = Format(p(iParamItem), sWidthPrec & "e+")
                iComp = StrComp(sSpec, "E", 0)
                If (iComp = 1) Then
                    sRet = UCase(sRet)
                End If
                
                If (InStr(sFlags, ",")) Then
                    sRet = "ERROR. Cannot use COMMA flag with %E"
                End If
                
            Case Is = "F" ' Floats
                ' we can specify the width & precision in two ways. "3.2" or "000.00". Both of these are
                ' equivalent. The VBA "Format" function wants the "000.00" format, so check to see which
                ' one was used. Convert sWidthPrec to a double and see if it equals 0. If not, then the
                ' 3.2 format was used.
                If sWidthPrec = "" Then sWidthPrec = "0"
                If (CDbl(sWidthPrec) <> 0) Then
                    sWidthPrec = ConvertToZeroes(sWidthPrec)
                End If
                
                sRet = Format(p(iParamItem), sWidthPrec)
                If (InStr(sFlags, "+") And sRet >= 0) Then
                    sRet = "+" & sRet
                End If
                
                If (InStr(sFlags, "+") And InStr(sRet, ".")) Then
                    sRet = sRet & "."
                End If
                
            Case Is = "G" ' Scientific, but different
                ' not supported because I don't understand it
                
            Case Is = "A" ' Hexadecimal, but different
                ' not supported because I don't understand it
                
            Case Is = "N" ' New line
                sRet = vbCrLf
                
            Case Is = "T"
                sRet = TimeDate(p(iParamItem), sSpec)
        End Select
        
        iPos = iPos + Len(sSpec)
        iTokenNum = iTokenNum + 1
        'iParamItem = iParamItem + 1
        sRet = AddPadding(sRet, iWidth, sFlags)
        sRet = s1 & sRet & s2
    Loop
    
    printf = Replace(Replace(Replace(sRet, Chr$(1) & Chr$(1), "%"), "\n", vbCrLf), "\t", vbTab)
End Function
Function AddPadding(str As String, width As Integer, flags As String)
    Dim spacesNeeded As Integer
    Dim spaces As String
    Dim i As Integer
    spaces = ""
    spacesNeeded = width - Len(str)
    For i = 1 To spacesNeeded
        spaces = spaces & " "
    Next
    If (width < Len(str)) Then
        AddPadding = str
    ElseIf (InStr(flags, "-")) Then ' left justified, i.e., right padding
        AddPadding = str & spaces
    Else
        AddPadding = spaces & str
    End If
End Function

Function TimeDate(ByVal str As String, fmt As String)
    Dim sCase As String, i As Integer
    sCase = Mid(fmt, 2)
    Select Case sCase
        Case Is = "H" ' Hour, 00, 24 hour clock
            TimeDate = Format(str, "hh")
        
        Case Is = "I" ' Hour, 00, 12 hour clock
            i = CInt(Format(str, "hh"))
            If (i > 12) Then i = i - 12
            TimeDate = Format(i, "00")
        
        Case Is = "k" ' Hour, no leading zeroes, 24 hour clock
            TimeDate = Format(str, "h")
        
        Case Is = "l" ' Hour, no leading zeroes, 12 hour clock
            i = CInt(Format(str, "hh"))
            If (i > 12) Then i = i - 12
            TimeDate = i
        
        Case Is = "M" ' Minute, 00-60 (include leading 0)
            TimeDate = Format(str, "nn")
        
        Case "S" ' Second, 00-60 (include leading 0)
            TimeDate = Format(str, "ss")
        
        Case "L" ' Millisecond, 000-999 (include leading zeroes)
            TimeDate = "Milliseconds not supported"
        
        Case "N" ' Nanosecond, 000000000 - 999999999 (include leading zeroes)
            TimeDate = "Nanoseconds not supported"
        
        Case Is = "p" ' am/pm indicator
            If (fmt = "Tp") Then
                TimeDate = Format(str, "AM/PM")
            Else
                TimeDate = Format(str, "am/pm")
            End If
        
        Case "z" ' Time zone number, like +5
            TimeDate = "Time zone information not supported"
        
        Case "Z" ' Time zone abbrev, like EST
            TimeDate = "Time zone information not supported"
        
        Case "s" ' number of seconds between the epoch and the argument
            TimeDate = DateDiff("S", "1/1/1970", str)
    
        Case "Q" ' Milliseconds since the epoch,
            TimeDate = "Milliseconds not supported"
        
        Case "B" ' Full month name, like "January"
            TimeDate = Format(str, "mmmm")
        
        Case "b" ' Abbrev month name, like "Jan"
            TimeDate = Format(str, "mmm")
        
        Case "h" ' Abbrev month name, like "Jan"
            TimeDate = Format(str, "mmm") ' same as "b" above
        
        Case "A" ' Full weekday name, like "Monday"
            TimeDate = Format(str, "dddd")
        
        Case "a" ' Abbrev weekday name, like "Mon"
            TimeDate = Format(str, "ddd")
        
        Case "C" ' Century
            TimeDate = Format(Application.WorksheetFunction.Floor((CInt(Format(str, "yyyy")) / 100), 1), "00")
        
        Case "Y" ' Year, formatted as at least four digits with leading zeros as necessary
            TimeDate = Format(str, "yyyy")
        
        Case "y" ' Last two digits of the year, formatted with leading zeros as necessary, i.e. 00 - 99
            TimeDate = Format(str, "yy")
            
        Case "j" ' Day of year, formatted as three digits with leading zeros as necessary, e.g. 001 - 366 for the Gregorian calendar.
            TimeDate = Format(str, "y")
        
        Case "m" ' Month, formatted as two digits with leading zeros as necessary, i.e. 01 - 13.
            TimeDate = Format(str, "mm")
        
        Case "d" ' Day of month, formatted as two digits with leading zeros as necessary, i.e. 01 - 31
            TimeDate = Format(str, "dd")
        
        Case "e" ' Day of month, formatted as two digits, i.e. 1 - 31.
            TimeDate = Format(str, "d")
        
        Case "R" ' Time formatted for the 24-hour clock as "%tH:%tM"
            TimeDate = Format(str, "hh") & ":" & Format(str, "nn")
        
        Case "T" ' Time formatted for the 24-hour clock as "%tH:%tM:%tS".
            TimeDate = Format(str, "hh") & ":" & Format(str, "nn") & ":" & Format(str, "ss")
        
        Case "r" ' Time formatted for the 12-hour clock as "%tI:%tM:%tS %Tp".
            i = CInt(Format(str, "hh"))
            If (i > 12) Then i = i - 12
            
            Dim ampm As String
            If (fmt = "Tp") Then
                ampm = Format(str, "AM/PM")
            Else
                ampm = Format(str, "am/pm")
            End If

            TimeDate = Format(i, "00") & ":" & Format(str, "nn") & ":" & Format(str, "ss") & " " & ampm
        
        Case "D" ' Date formatted as "%tm/%td/%ty".
            TimeDate = Format(str, "mm") & "/" & Format(str, "d") & "/" & Format(str, "yy")
        
        Case "F" ' ISO 8601 complete date formatted as "%tY-%tm-%td".
            TimeDate = Format(str, "yyyy") & "-" & Format(str, "mm") & "-" & Format(str, "d")
        
        Case "c" ' Date and time formatted as "%ta %tb %td %tT %tZ %tY", e.g. "Sun Jul 20 16:17:00 EDT 1969".
            TimeDate = "Format not supported due to time zone information not supported in VBA"
        
    End Select
End Function

Function ConvertToZeroes(str As String)
' take a decimal string, like 3.2, and turn it into a string of all zeroes, such as "000.00"
    Dim l As Integer, r As Integer, iDecPos As Integer, i As Integer, s As String
    If (InStr(str, ".")) Then
        iDecPos = InStr(str, ".")
        l = CInt(Left(str, iDecPos))
        If (InStr(str, ".") = Len(str)) Then
            r = 0
        Else
            r = Mid(str, iDecPos + 1)
        End If
    Else
        l = CInt(str)
    End If
    
    s = ""
    For i = 1 To l
        s = s & "0"
    Next
    
    If (InStr(str, ".")) Then
        s = s & "."
    End If
    
    For i = 1 To r
        s = s & "0"
    Next
    
    ConvertToZeroes = s
End Function

Function test()
Debug.Print printf("The quick brown %-10S jumps over the lazy %-8S!", "foxxy fox", "dog")
Debug.Print printf("%8S The quick brown %S jumps over the lazy!", "foxxy fox", "dog")
Debug.Print printf("floats: %+4.2f %+.0e %+E %+0000.0f \n", 3.1416, 3.1416, 3.1416, 3.1416);
Debug.Print printf("The quick 10%% brown %S jumps over the\nlazy %s", "fox", "dog")
Debug.Print printf("Boolean tests (vbFalse) %b\n(nothing) %b\n(Null) %b\n(vbNull) %b\n(1=1) %b\n(1=0) %b\n(1) %b\n(0) %b XXX", vbFalse, , Null, vbNull, 1 = 1, 1 = 0, 1, 0)
Debug.Print printf("Hex test: %h XXX", 255)
Debug.Print printf("Param test, 3: %3$s, 2: %2$S, 1: %1$s, 4: %s XXX", "one", "two", "three", "four")
Debug.Print printf("Char test:\n&H63: %c\nf: %c XXX", &H63, "f")
Debug.Print printf("D %2$+5d XXX", 32, 16)
Debug.Print printf("D %2$00000.d XXX", 32, 16.5)
Debug.Print printf("Oct test: %o XXX", 64)
Debug.Print printf("Hex test: %h XXX", 255)
Debug.Print printf("E: %0.000E XXX", 256789125)
Debug.Print printf("H: %tF XXX", Now)

End Function

