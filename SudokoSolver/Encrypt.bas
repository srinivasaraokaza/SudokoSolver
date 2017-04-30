Attribute VB_Name = "Encrypt"
Option Explicit
Private Function WriteIntoFile(filename As String, ContentToWrite As String)
  Open App.Path & "\" & filename For Output As #1    ' Open file.
        Write #1, Trim(ContentToWrite)
        Close #1   ' Close file.
End Function
Private Function GettingDateFromFile(filename As String) As String
Dim strRegFile As String
Dim strCharByCharReg As String
Dim strChar As String
strServerFile = Dir(App.Path & "\" & filename)
If strServerFile = "" Then
Else
    Open App.Path & "\" & filename For Input As #1
        Do While Not EOF(1)   ' Loop until end of file.
           strChar = Input(1, #1)
           Select Case Asc(strChar)
             Case 34, 13, 10
             Case Else
               strCharByCharReg = strCharByCharReg & strChar   ' Get one character.
            End Select
            Loop
    Close #1
    Text1.Text = Trim(strCharByCharReg)
    GettingDateFromFile = strCharByCharReg
End If
End Function
Private Function Encryption(TextToEncrypt As String)
        Dim Encrypt
            For e = 1 To 5
                Encrypt = Encode(TextToEncrypt)
            Encryption = Encrypt
            Next e
        
End Function

Private Function Decryption(TextToDecrypt As String)
        Dim decrypt
            For d = 1 To 5
                decrypt = DeCode(TextToDecrypt)
                Decryption = decrypt
            Next d
End Function

Public Function DeCode(vText As String) As String
    On Error GoTo ErrHandler
    Dim CurSpc As Integer
    Dim varLen As Integer
    Dim varChr As String
    Dim varFin As String
    CurSpc = CurSpc + 1
    varLen = Len(vText)
    Do While CurSpc <= varLen
        DoEvents
        varChr = Mid(vText, CurSpc, 3)
        Select Case varChr
            'lower case
            Case "coe"
                varChr = "a"
            Case "wer"
                varChr = "b"
            Case "ibq"
                varChr = "c"
            Case "am7"
                varChr = "d"
            Case "pm1"
                varChr = "e"
            Case "mop"
                varChr = "f"
            Case "9v4"
                varChr = "g"
            Case "qu6"
                varChr = "h"
            Case "zxc"
                varChr = "i"
            Case "4mp"
                varChr = "j"
            Case "f88"
                varChr = "k"
            Case "qe2"
                varChr = "l"
            Case "vbn"
                varChr = "m"
            Case "qwt"
                varChr = "n"
            Case "pl5"
                varChr = "o"
            Case "13s"
                varChr = "p"
            Case "c%l"
                varChr = "q"
            Case "w$w"
                varChr = "r"
            Case "6a@"
                varChr = "s"
            Case "!2&"
                varChr = "t"
            Case "(=c"
                varChr = "u"
            Case "wvf"
                varChr = "v"
            Case "dp0"
                varChr = "w"
            Case "w$-"
                varChr = "x"
            Case "vn&"
                varChr = "y"
            Case "c*4"
                varChr = "z"
            'numbers
            Case "aq@"
                varChr = "1"
            Case "902"
                varChr = "2"
            Case "2.&"
                varChr = "3"
            Case "/w!"
                varChr = "4"
            Case "|pq"
                varChr = "5"
            Case "ml|"
                varChr = "6"
            Case "t'?"
                varChr = "7"
            Case ">^s"
                varChr = "8"
            Case "<s^"
                varChr = "9"
            Case ";&c"
                varChr = "0"
            'caps
            Case "$)c"
                varChr = "A"
            Case "-gt"
                varChr = "B"
            Case "|p*"
                varChr = "C"
            Case "1" & Chr(34) & "r"
                varChr = "D"
            Case "c>:"
                varChr = "E"
            Case "@+x"
                varChr = "F"
            Case "v^a"
                varChr = "G"
            Case "]eE"
                varChr = "H"
            Case "aP0"
                varChr = "I"
            Case "{=1"
                varChr = "J"
            Case "cWv"
                varChr = "K"
            Case "cDc"
                varChr = "L"
            Case "*,!"
                varChr = "M"
            Case "fW" & Chr(34)
                varChr = "N"
            Case ".?T"
                varChr = "O"
            Case "%<8"
                varChr = "P"
            Case "@:a"
                varChr = "Q"
            Case "&c$"
                varChr = "R"
            Case "WnY"
                varChr = "S"
            Case "{Sh"
                varChr = "T"
            Case "_%M"
                varChr = "U"
            Case "}'$"
                varChr = "V"
            Case "QlU"
                varChr = "W"
            Case "Im^"
                varChr = "X"
            Case "l|P"
                varChr = "Y"
            Case ".>#"
                varChr = "Z"
            'Special characters
            Case "\" & Chr(34) & "]"
                varChr = "!"
            Case "cY,"
                varChr = "@"
            Case "x%B"
                varChr = "#"
            Case "a*v"
                varChr = "$"
            Case "'&T"
                varChr = "%"
            Case ";%R"
                varChr = "^"
            Case "eG_"
                varChr = "&"
            Case "Z/e"
                varChr = "*"
            Case "rG\"
                varChr = "("
            Case "]*F"
                varChr = ")"
            Case "@B*"
                varChr = "_"
            Case "+Hc"
                varChr = "-"
            Case "&|D"
                varChr = "="
            Case "(:#"
                varChr = "+"
            Case "SlW"
                varChr = "["
            Case "'QB"
                varChr = "]"
            Case "{D>"
                varChr = "{"
            Case "+c%"
                varChr = "}"
            Case "(s:"
                varChr = ":"
            Case "^a("
                varChr = ";"
            Case "16."
                varChr = "'"
            Case "s.*"
                varChr = Chr(34)
            Case "&?W"
                varChr = ","
            Case "GPQ"
                varChr = "."
            Case "SK*"
                varChr = "<"
            Case "RL^"
                varChr = ">"
            Case "40C"
                varChr = "/"
            Case "?#9"
                varChr = "?"
            Case "_?/"
                varChr = "\"
            Case "(_@"
                varChr = "|"
            Case "=#B"
                varChr = " "
        End Select
        varFin = varFin & varChr
        CurSpc = CurSpc + 3
        DoEvents
    Loop
    DeCode = varFin
    Exit Function
ErrHandler:
    Dim ErrNum, ErrDesc, ErrSource
    ErrNum = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    MsgBox "Error# = " & ErrNum & vbCrLf & "Description = " & ErrDesc & vbCrLf & "Source = " & ErrSource, vbCritical + vbOKOnly, "Program Error!"
    Err.Clear
    Exit Function
End Function

Public Function Encode(vText As String)
    On Error GoTo ErrHandler
    Dim CurSpc As Integer
    Dim varLen As Integer
    Dim varChr As String
    Dim varFin As String
    varLen = Len(vText)
    Do While CurSpc <= varLen
        DoEvents
        CurSpc = CurSpc + 1
        varChr = Mid(vText, CurSpc, 1)
        Select Case varChr
            'lower case
            Case "a"
                varChr = "coe"
            Case "b"
                varChr = "wer"
            Case "c"
                varChr = "ibq"
            Case "d"
                varChr = "am7"
            Case "e"
                varChr = "pm1"
            Case "f"
                varChr = "mop"
            Case "g"
                varChr = "9v4"
            Case "h"
                varChr = "qu6"
            Case "i"
                varChr = "zxc"
            Case "j"
                varChr = "4mp"
            Case "k"
                varChr = "f88"
            Case "l"
                varChr = "qe2"
            Case "m"
                varChr = "vbn"
            Case "n"
                varChr = "qwt"
            Case "o"
                varChr = "pl5"
            Case "p"
                varChr = "13s"
            Case "q"
                varChr = "c%l"
            Case "r"
                varChr = "w$w"
            Case "s"
                varChr = "6a@"
            Case "t"
                varChr = "!2&"
            Case "u"
                varChr = "(=c"
            Case "v"
                varChr = "wvf"
            Case "w"
                varChr = "dp0"
            Case "x"
                varChr = "w$-"
            Case "y"
                varChr = "vn&"
            Case "z"
                varChr = "c*4"
            'numbers
            Case "1"
                varChr = "aq@"
            Case "2"
                varChr = "902"
            Case "3"
                varChr = "2.&"
            Case "4"
                varChr = "/w!"
            Case "5"
                varChr = "|pq"
            Case "6"
                varChr = "ml|"
            Case "7"
                varChr = "t'?"
            Case "8"
                varChr = ">^s"
            Case "9"
                varChr = "<s^"
            Case "0"
                varChr = ";&c"
            'caps
            Case "A"
                varChr = "$)c"
            Case "B"
                varChr = "-gt"
            Case "C"
                varChr = "|p*"
            Case "D"
                varChr = "1" & Chr(34) & "r"
            Case "E"
                varChr = "c>:"
            Case "F"
                varChr = "@+x"
            Case "G"
                varChr = "v^a"
            Case "H"
                varChr = "]eE"
            Case "I"
                varChr = "aP0"
            Case "J"
                varChr = "{=1"
            Case "K"
                varChr = "cWv"
            Case "L"
                varChr = "cDc"
            Case "M"
                varChr = "*,!"
            Case "N"
                varChr = "fW" & Chr(34)
            Case "O"
                varChr = ".?T"
            Case "P"
                varChr = "%<8"
            Case "Q"
                varChr = "@:a"
            Case "R"
                varChr = "&c$"
            Case "S"
                varChr = "WnY"
            Case "T"
                varChr = "{Sh"
            Case "U"
                varChr = "_%M"
            Case "V"
                varChr = "}'$"
            Case "W"
                varChr = "QlU"
            Case "X"
                varChr = "Im^"
            Case "Y"
                varChr = "l|P"
            Case "Z"
                varChr = ".>#"
            'Special characters
            Case "!"
                varChr = "\" & Chr(34) & "]"
            Case "@"
                varChr = "cY,"
            Case "#"
                varChr = "x%B"
            Case "$"
                varChr = "a*v"
            Case "%"
                varChr = "'&T"
            Case "^"
                varChr = ";%R"
            Case "&"
                varChr = "eG_"
            Case "*"
                varChr = "Z/e"
            Case "("
                varChr = "rG\"
            Case ")"
                varChr = "]*F"
            Case "_"
                varChr = "@B*"
            Case "-"
                varChr = "+Hc"
            Case "="
                varChr = "&|D"
            Case "+"
                varChr = "(:#"
            Case "["
                varChr = "SlW"
            Case "]"
                varChr = "'QB"
            Case "{"
                varChr = "{D>"
            Case "}"
                varChr = "+c%"
            Case ":"
                varChr = "(s:"
            Case ";"
                varChr = "^a("
            Case "'"
                varChr = "16."
            Case Chr(34)
                varChr = "s.*"
            Case ","
                varChr = "&?W"
            Case "."
                varChr = "GPQ"
            Case "<"
                varChr = "SK*"
            Case ">"
                varChr = "RL^"
            Case "/"
                varChr = "40C"
            Case "?"
                varChr = "?#9"
            Case "\"
                varChr = "_?/"
            Case "|"
                varChr = "(_@"
            Case " "
                varChr = "=#B"
        End Select
        varFin = varFin & varChr
        DoEvents
    Loop
    Encode = varFin
    Exit Function
ErrHandler:
    Dim ErrNum, ErrDesc, ErrSource
    ErrNum = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source
    MsgBox "Error# = " & ErrNum & vbCrLf & "Description = " & ErrDesc & vbCrLf & "Source = " & ErrSource, vbCritical + vbOKOnly, "Program Error!"
    Err.Clear
    Exit Function
End Function
