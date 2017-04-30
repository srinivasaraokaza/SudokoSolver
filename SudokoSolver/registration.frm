VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Encryption"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   3120
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2778
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"registration.frx":0000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Decrypt"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encrypt"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "register"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
        Dim Encrypt
            For e = 1 To 5
                Encrypt = Encode(Text1.Text)
            Text1.Text = Encrypt
            Next e
        
End Sub

Private Sub Command2_Click()
        Dim decrypt
            For d = 1 To 5
                decrypt = DeCode(Text1.Text)
                Text1.Text = decrypt
            Next d
End Sub

Private Sub Command3_Click()
Text1.Text = Date
End Sub
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
Private Sub Form_Load()
Dim EDate As String
Dim GameCount As String
EDate = GettingDateFromFile("reg.png")
If Date < EDate Then
  MsgBox ("You have changed the dates")
  Unload Me
ElseIf Date > EDate Then
  MsgBox ("New Date So,You Are allowed to solve only three puzzles")
    Call WriteIntoFile("reg.png", Date)
    Call WriteIntoFile("co.bmp", 1)
Else
  GameCount = GettingDateFromFile("co.bmp")
  GameCount = GameCount + 1
    Call WriteIntoFile("co.bmp", GameCount)
  If GameCount < 4 Then
      MsgBox ("Today u  r playing the game " & GameCount & " time")
  Else
    MsgBox ("You are not allowed to play game today,today limit is over")
    Unload Me
  End If
End If
End Sub
