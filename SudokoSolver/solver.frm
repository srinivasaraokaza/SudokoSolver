VERSION 5.00
Begin VB.Form frmsolver 
   Caption         =   "Game Solver"
   ClientHeight    =   5925
   ClientLeft      =   180
   ClientTop       =   570
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   5265
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   99
      Left            =   4320
      MaxLength       =   1
      TabIndex        =   100
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   98
      Left            =   3840
      MaxLength       =   1
      TabIndex        =   99
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   97
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   98
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   96
      Left            =   2880
      MaxLength       =   1
      TabIndex        =   97
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   95
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   96
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   94
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   95
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   93
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   94
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   92
      Left            =   960
      MaxLength       =   1
      TabIndex        =   93
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   91
      Left            =   480
      MaxLength       =   1
      TabIndex        =   92
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   90
      Left            =   0
      MaxLength       =   1
      TabIndex        =   91
      Text            =   "I"
      Top             =   4320
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   89
      Left            =   4320
      MaxLength       =   1
      TabIndex        =   90
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   88
      Left            =   3840
      MaxLength       =   1
      TabIndex        =   89
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   87
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   88
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   86
      Left            =   2880
      MaxLength       =   1
      TabIndex        =   87
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   85
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   86
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   84
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   85
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   83
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   84
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   82
      Left            =   960
      MaxLength       =   1
      TabIndex        =   83
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   81
      Left            =   480
      MaxLength       =   1
      TabIndex        =   82
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   80
      Left            =   0
      MaxLength       =   1
      TabIndex        =   81
      Text            =   "H"
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   79
      Left            =   4320
      MaxLength       =   1
      TabIndex        =   80
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   78
      Left            =   3840
      MaxLength       =   1
      TabIndex        =   79
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   77
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   78
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   76
      Left            =   2880
      MaxLength       =   1
      TabIndex        =   77
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   75
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   76
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   74
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   75
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   73
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   74
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   72
      Left            =   960
      MaxLength       =   1
      TabIndex        =   73
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   71
      Left            =   480
      MaxLength       =   1
      TabIndex        =   72
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   70
      Left            =   0
      MaxLength       =   1
      TabIndex        =   71
      Text            =   "F"
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   69
      Left            =   4320
      MaxLength       =   1
      TabIndex        =   70
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   68
      Left            =   3840
      MaxLength       =   1
      TabIndex        =   69
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   67
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   68
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   66
      Left            =   2880
      MaxLength       =   1
      TabIndex        =   67
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   65
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   66
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   64
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   65
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   63
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   64
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   62
      Left            =   960
      MaxLength       =   1
      TabIndex        =   63
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   61
      Left            =   480
      MaxLength       =   1
      TabIndex        =   62
      Top             =   2880
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   60
      Left            =   0
      MaxLength       =   1
      TabIndex        =   61
      Text            =   "G"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   59
      Left            =   4320
      MaxLength       =   1
      TabIndex        =   60
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   58
      Left            =   3840
      MaxLength       =   1
      TabIndex        =   59
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   57
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   58
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   56
      Left            =   2880
      MaxLength       =   1
      TabIndex        =   57
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   55
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   56
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   54
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   55
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   53
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   54
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   52
      Left            =   960
      MaxLength       =   1
      TabIndex        =   53
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   51
      Left            =   480
      MaxLength       =   1
      TabIndex        =   52
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   50
      Left            =   0
      MaxLength       =   1
      TabIndex        =   51
      Text            =   "E"
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   49
      Left            =   4320
      MaxLength       =   1
      TabIndex        =   50
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   48
      Left            =   3840
      MaxLength       =   1
      TabIndex        =   49
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   47
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   48
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   46
      Left            =   2880
      MaxLength       =   1
      TabIndex        =   47
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   45
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   46
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   44
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   45
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   43
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   44
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   42
      Left            =   960
      MaxLength       =   1
      TabIndex        =   43
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   41
      Left            =   480
      MaxLength       =   1
      TabIndex        =   42
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   40
      Left            =   0
      MaxLength       =   1
      TabIndex        =   41
      Text            =   "D"
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   39
      Left            =   4320
      MaxLength       =   1
      TabIndex        =   40
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   38
      Left            =   3840
      MaxLength       =   1
      TabIndex        =   39
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   37
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   38
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   36
      Left            =   2880
      MaxLength       =   1
      TabIndex        =   37
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   35
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   36
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   34
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   35
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   33
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   34
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   32
      Left            =   960
      MaxLength       =   1
      TabIndex        =   33
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   31
      Left            =   480
      MaxLength       =   1
      TabIndex        =   32
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   30
      Left            =   0
      MaxLength       =   1
      TabIndex        =   31
      Text            =   "C"
      Top             =   1440
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   29
      Left            =   4320
      MaxLength       =   1
      TabIndex        =   30
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   28
      Left            =   3840
      MaxLength       =   1
      TabIndex        =   29
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   27
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   28
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   26
      Left            =   2880
      MaxLength       =   1
      TabIndex        =   27
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   25
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   26
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   24
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   25
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   23
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   24
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   22
      Left            =   960
      MaxLength       =   1
      TabIndex        =   23
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   21
      Left            =   480
      MaxLength       =   1
      TabIndex        =   22
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   20
      Left            =   0
      MaxLength       =   1
      TabIndex        =   21
      Text            =   "B"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   19
      Left            =   4320
      MaxLength       =   1
      TabIndex        =   20
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   18
      Left            =   3840
      MaxLength       =   1
      TabIndex        =   19
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   17
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   18
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   2880
      MaxLength       =   1
      TabIndex        =   17
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   16
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   15
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   14
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   960
      MaxLength       =   1
      TabIndex        =   13
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   480
      MaxLength       =   1
      TabIndex        =   12
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   0
      MaxLength       =   1
      TabIndex        =   11
      Text            =   "A"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   0
      MaxLength       =   1
      TabIndex        =   10
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   4320
      MaxLength       =   1
      TabIndex        =   9
      Text            =   "9"
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   3840
      MaxLength       =   1
      TabIndex        =   8
      Text            =   "8"
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   7
      Text            =   "7"
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   2880
      MaxLength       =   1
      TabIndex        =   6
      Text            =   "6"
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   5
      Text            =   "5"
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   4
      Text            =   "4"
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   1440
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "3"
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   480
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "1"
      Top             =   0
      Width           =   495
   End
   Begin VB.TextBox txtbox 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   960
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "2"
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdsolver 
      BackColor       =   &H000000FF&
      Caption         =   "Solver"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Menu mainmenu 
      Caption         =   "main"
      Visible         =   0   'False
      Begin VB.Menu one 
         Caption         =   "1"
         Index           =   1
      End
      Begin VB.Menu two 
         Caption         =   "2"
         Index           =   2
      End
      Begin VB.Menu three 
         Caption         =   "3"
         Index           =   3
      End
      Begin VB.Menu four 
         Caption         =   "4"
         Index           =   4
      End
      Begin VB.Menu five 
         Caption         =   "5"
         Index           =   5
      End
      Begin VB.Menu six 
         Caption         =   "6"
         Index           =   6
      End
      Begin VB.Menu seven 
         Caption         =   "7"
         Index           =   7
      End
      Begin VB.Menu eight 
         Caption         =   "8"
         Index           =   8
      End
      Begin VB.Menu nine 
         Caption         =   "9"
         Index           =   9
      End
   End
End
Attribute VB_Name = "frmsolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sss(1 To 9, 1 To 9, 1 To 9) As Integer
Dim Graph(1 To 9, 1 To 9) As Integer       ' it consists 1-exist 0-not exist
Dim outputarray(1 To 9, 1 To 9) As Integer 'values exist in ouputarray
Public ind As Integer, element As Integer, si As Integer
Dim stack(1 To 82, 1 To 2)  As Integer
Public Gameover As Boolean
Public Gamenoanswer As Boolean, Allow As Boolean
Public Function Initialize()
Dim i As Integer, j As Integer, t As Integer
For i = 1 To 9
    For j = 1 To 9
      If txtbox(i & j).Text <> "" Then
          't = outputarray(i, j)
          t = txtbox(i & j).Text
          txtbox(i & j).ForeColor = vbRed
          't1 = 4
        Call Checksurround(i, j, t)
      End If
    Next
 Next
End Function
Private Sub cmdsolver_Click()
Dim p As Integer
If Checkrows Then
 If Checkcolumns Then
   If Checkblocks Then
     Allow = True
   End If
  End If
End If
If Allow = True Then
    Initialize
    Displaynumbers
    For p = 0 To 5
    Singleelmentrows
    Displaynumbers
    Singleelementcolumns
    Displaynumbers
    Cellcontainoneelement
    Displaynumbers
    Next
    Displaynumbers
    Generatesolution
Else
   MsgBox ("wrong input")
End If



End Sub
Private Sub Form_Load()

'Locktxtcontrols
Dim i As Integer, j As Integer, k As Integer

For i = 1 To 9
 For j = 1 To 9
   For k = 1 To 9
      sss(i, j, k) = k
   Next
 Next
Next
'Releasetxtcontrols
   si = 0
Gameover = False
Gamenoanswer = False
End Sub

Private Sub Form_Resize()
'Form1.Width = 5550
'Form1.Height = 6930
End Sub
Public Function Locktxtcontrols()
Dim i As Integer, j As Integer
For i = 1 To 9
    For j = 1 To 9
         txtbox(i & j).Locked = True
    Next
Next
End Function
Public Function Releasetxtcontrols()
Dim i As Integer, j As Integer
For i = 1 To 9
    For j = 1 To 9
         txtbox(i & j).Enabled = True
    Next
Next
End Function
Public Function Checksurround(ByVal i As Integer, ByVal j As Integer, ByVal t As Integer)
Dim t1 As Integer, t2 As Integer, k As Integer
outputarray(i, j) = t
'Graph(i, j) = 1
         For k = 1 To 9
           sss(i, k, t) = 0      ' removing all same elements in the same row
         Next
         For k = 1 To 9
           sss(k, j, t) = 0      ' removing all same elements in the same column
         Next
         
         If i < 4 Then
           t1 = 1
         ElseIf i < 7 Then
           t1 = 4
         Else
           t1 = 7
         End If
         
         If j < 4 Then
           t2 = 1
         ElseIf j < 7 Then
           t2 = 4
         Else
           t2 = 7
         End If
         For i1 = t1 To t1 + 2
           For i2 = t2 To t2 + 2
             sss(i1, i2, t) = 0
           Next
         Next
         For k = 1 To 9
           sss(i, j, k) = 0
         Next
End Function
Public Function Displaynumbers()
Dim i As Integer, j As Integer
For i = 1 To 9
    For j = 1 To 9
      If outputarray(i, j) <> 0 Then
      'If txtbox(i & j).text = "" Then
      
      'If Graph(i, j) = 1 Then
      If txtbox(i & j).Text = "" Then
        'txtbox(i & j).Text = outputarray(i, j)
          txtbox(i & j).Enabled = False
      End If
      End If
    Next
 Next
End Function
Public Function Checkrows() As Boolean
 Dim i As Integer, k As Integer, j As Integer
 Dim numexist(1 To 9) As Boolean
 
 For i = 1 To 9     ' for nine rows
   For k = 1 To 9
     numexist(k) = False
   Next
   For j = 1 To 9   'for each element in the column
    If txtbox(i & j).Text <> "" Then
      If numexist(Int(txtbox(i & j).Text)) = False Then
        numexist(Int(txtbox(i & j).Text)) = True
      Else
        'MsgBox ("Row check is fail")
        Checkrows = False
        Exit Function
      End If
    End If
   Next
Next
   'MsgBox ("Row check is succ")
 Checkrows = True
 End Function
Public Function Checkcolumns() As Boolean
 Dim numexist(1 To 9) As Boolean
 Dim i, k, j As Integer
 
 For i = 1 To 9     ' for nine rows
   For k = 1 To 9
     numexist(k) = False
   Next
   For j = 1 To 9   'for each element in the column
    If txtbox(j & i).Text <> "" Then
      If numexist(Int(txtbox(j & i).Text)) = False Then
        numexist(Int(txtbox(j & i).Text)) = True
      Else
       ' MsgBox ("column check is fail")
        Checkcolumns = False
        Exit Function
      End If
    End If
   Next
Next
  ' MsgBox ("column check is succ")
 Checkcolumns = True
End Function

Public Function Checkblocks() As Boolean
 Dim i As Integer, j As Integer, t As Integer, t1 As Integer, t2 As Integer, count As Integer
 Dim numexist(1 To 9) As Boolean
 For k = 1 To 9
   If k < 4 Then
      t1 = 1
   ElseIf k < 7 Then
      t1 = 4
   Else
      t1 = 7
   End If
   t2 = 1
   If (k Mod 3) = 0 Then
      t2 = 7
   ElseIf ((k + 1) Mod 3) = 0 Then
      t2 = 4
   End If
   For n = 1 To 9
    numexist(n) = False
   Next
    For i = t1 To t1 + 2
     For j = t2 To t2 + 2
       If txtbox(i & j).Text <> "" Then
          If numexist(Int(txtbox(i & j).Text)) = False Then
            numexist(Int(txtbox(i & j).Text)) = True
          Else
           ' MsgBox ("Block check is fail")
            Checkblocks = False
            Exit Function
          End If
       End If           'If inputarray(i, j) = n Then  'If Int(Command1(i & j).Caption) = n Then
     Next
    Next
 Next
  Checkblocks = True
    'MsgBox ("Block check is succ")
End Function

Public Function Allarefilled() As Boolean
Dim i As Integer, j As Integer
  For i = 1 To 9     ' for nine rows
   For j = 1 To 9   'for nine columns
     If txtbox(i & j).Text = "" Then
           MsgBox ("In " & i & " RoW  " & j & " Element  " & " is empty ")
            Allarefilled = False
             Exit Function
     End If
   Next
 Next
 'MsgBox ("Allfilled is succ")
 Allarefilled = True
End Function

Public Function GenerateStack(ind1 As Integer)
 Dim t As Integer, nextval As Integer, c As Integer
 Dim statckenable, popbool1 As Boolean
   For i = ind1 To 99
    If i < 11 Then
     i = 11
    End If
    stackenable = False
     If (i Mod 10) <> 0 Then
         If txtbox(i).Text <> "" Then
          'MsgBox (CMD1(i).Caption)
         ElseIf txtbox(i).Text = "" Then
            For k = 1 To 9
                If sss(Int(i / 10), i Mod 10, k) <> 0 Then
                   t = k
                  If Checkrowelement(i, t) Then
                    If Checkcolumnelement(i, t) Then
                      If Checkblockelement(i, t) Then
                         stackenable = True
                         Exit For
                      End If
                    End If
                  End If
                End If
            Next
        ' popbool1 = False
         If stackenable Then
             Call Pushelement(i, t)
         Else
            c = 0
             Do
             If c > 99 Then
               Gamenoanswer = True
               Exit Function
             End If
             Popelement
             nextval = Nextelement(ind, element)
             c = c + 1
             Loop Until nextval <> 10
             Call Pushelement(ind, nextval)
             i = ind - c
         End If
        End If
     End If
    Next
    If Allarefilled Then
      If Checkrows Then
        If Checkcolumns Then
             If Checkblocks Then
                 MsgBox ("Succ")
                 Gameover = True
                 Locktxtcontrols
                 Exit Function
             End If
        End If
      End If
    Else
     Print "no"
    End If
End Function

Public Function Popelement()
  If si = 0 Then
    Print "srinu"
  Else
    ind = stack(si, 1)
    element = stack(si, 2)
    txtbox(ind).Text = ""
    si = si - 1
 End If
End Function
Public Function Pushelement(ByVal index_no As Integer, ByVal ele As Integer)
            si = si + 1
            stack(si, 1) = index_no     ' for index
            stack(si, 2) = ele     ' for element
            txtbox(index_no).Text = ele
End Function
Public Function Generatesolution()
Dim t_index As Integer, t_element As Integer
Dim pushboll As Boolean
Static popbool As Boolean
GenerateStack (11)
If Gamenoanswer = False Then
If Gameover = False Then
Popelement
Do
  pushbool = False
'  'If Allarefilled Then
'  If Checkrows Then
'        If Checkcolumns Then
'             If Checkblocks Then
'                 MsgBox ("Succ")
'             End If
'        End If
'   '  End If
'   Else
     If popbool = True Then
           Popelement
           popbool = False
     End If
     
     
        For k = 1 To 9
            If sss(Int(ind / 10), ind Mod 10, k) <> 0 Then
                 t = k
                  If Checkrowelement(i, t) Then
                    If Checkcolumnelement(i, t) Then
                      If Checkblockelement(i, t) Then
                         pushbool = True
                         Exit For
                      End If
                    End If
                  End If
            End If
        Next
        
          popbool = True
        If k = 10 Then
          Popelement
          popbool = False
        End If
        If pushbool = True Then
            Call Pushelement(ind, t)
             If Gameover = False Then
                GenerateStack (ind)
             End If
        End If
   If Gameover = True Then
           Exit Function
   End If
Loop Until si = 0
End If
End If
End Function
Public Function Checkrowelement(ByVal id As Integer, ByVal el As Integer) As Boolean
 Dim i As Integer, j As Integer
   i = Int(id / 10)
   For j = 1 To 9   'for each element in the row
    If txtbox(i & j).Text <> "" Then
      If Int(txtbox(i & j).Text) = el Then
       ' MsgBox ("Row check is fail")
        Checkrowelement = False
        Exit Function
      End If
    End If
   Next
   'MsgBox ("CheckRowelement is succ")
 Checkrowelement = True
End Function
Public Function Checkcolumnelement(ByVal id As Integer, ByVal el As Integer) As Boolean
 Dim i As Integer, j As Integer
   i = Int(id Mod 10)
   For j = 1 To 9   'for each element in the column
    If txtbox(j & i).Text <> "" Then
      If Int(txtbox(j & i).Text) = el Then
        'MsgBox ("column check is fail")
        Checkcolumnelement = False
        Exit Function
      End If
    End If
   Next
   'MsgBox ("column check is succ")
 Checkcolumnelement = True
End Function
Public Function Checkblockelement(ByVal id As Integer, ByVal el As Integer) As Boolean
 Dim i As Integer, j As Integer, t1 As Integer, t2 As Integer

     If Int(id / 10) < 4 Then
           t1 = 1
         ElseIf Int(id / 10) < 7 Then
           t1 = 4
         Else
           t1 = 7
         End If
         
         If Int(id Mod 10) < 4 Then
           t2 = 1
         ElseIf Int(id Mod 10) < 7 Then
           t2 = 4
         Else
           t2 = 7
         End If
    For i = t1 To t1 + 2
     For j = t2 To t2 + 2
       If txtbox(i & j).Text <> "" Then
          If Int(txtbox(i & j).Text) = el Then
            'MsgBox ("Block check is fail")
             Checkblockelement = False
            Exit Function
          End If
       End If           'If inputarray(i, j) = n Then  'If Int(Command1(i & j).Caption) = n Then
     Next
    Next
  Checkblockelement = True
    'MsgBox ("Block check is succ")
End Function
Public Function Nextelement(ByVal curr_index As Integer, ByVal curr_no As Integer) As Integer
   'Popelement
   Dim k As Integer
     For k = (curr_no + 1) To 9
        If sss(Int(curr_index / 10), curr_index Mod 10, k) <> 0 Then
           If Checkrowelement(curr_index, k) Then
              If Checkcolumnelement(curr_index, k) Then
                 If Checkblockelement(curr_index, k) Then
                    Nextelement = k
                    Exit Function
                 End If
              End If
           End If
        End If
     Next
 Nextelement = k
End Function

Public Function Singleelmentrows()
' Checking single element rows
Dim t As Integer
 Dim count As Integer
For i = 1 To 9      ' for rowwise checking
    For j = 1 To 9  ' for numbers 1 to 9

            count = 0
         For k = 1 To 9  'for columnwise
           If sss(i, k, j) = j Then
              count = count + 1
              t = k
           End If
           ' sss(i, j, k) = t
         Next
         If count = 1 Then
          txtbox(i & t).Text = j
          'outputarray(i, t) = j
          Call Checksurround(i, t, j)
         End If

    Next
 Next

End Function
Public Function Singleelementcolumns()
'cheking single element columns
Dim t As Integer
 Dim count As Integer
For i = 1 To 9      ' for columnwise checking
    For j = 1 To 9  ' for numbers 1 to 9

            count = 0
         For k = 1 To 9  'for rowwise
           If sss(k, i, j) = j Then
              count = count + 1
              t = k
           End If
           ' sss(i, j, k) = t
         Next
         If count = 1 Then
          'outputarray(t, i) = j
          txtbox(t & i).Text = j
          Call Checksurround(t, i, j)
         End If

    Next
 Next


End Function
Public Function Cellcontainoneelement()
'checking cell contains only one element
Dim t As Integer
 Dim count As Integer
For i = 1 To 9
    For j = 1 To 9
      If Graph(i, j) = 0 Then
         ' t = outputarray(i, j)
            count = 0
         For k = 1 To 9
           If sss(i, j, k) = 0 Then
              count = count + 1
           Else
              t = k
           End If
           
           ' sss(i, j, k) = t
         Next
         If count = 8 Then
           'outputarray(i, j) = t
           txtbox(i & j).Text = t
           Call Checksurround(i, j, t)
         End If
      End If
    Next
 Next
End Function

Private Sub txtbox_KeyPress(Index As Integer, KeyAscii As Integer)
If Index < 99 Then
    If ((Index + 1) Mod 10) = 0 Then
        txtbox(Index + 2).SetFocus
    Else
        txtbox(Index + 1).SetFocus
    End If
End If
End Sub
