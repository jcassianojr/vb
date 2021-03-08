VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "xpcontrols.ocx"
Begin VB.Form Calculadora 
   Caption         =   "Calculadora Científica"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800080&
   Icon            =   "Calculator.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin XPControls.XPButton XPButton1 
      Height          =   375
      Left            =   5880
      TabIndex        =   54
      Top             =   120
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdnumber 
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   5520
      TabIndex        =   53
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton cmdnumber 
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   5040
      TabIndex        =   52
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton cmdnumber 
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   4560
      TabIndex        =   51
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton cmdnumber 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   4080
      TabIndex        =   50
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton cmdnumber 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   3600
      TabIndex        =   49
      Top             =   3120
      Width           =   495
   End
   Begin VB.CommandButton cmdnumber 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   3120
      TabIndex        =   48
      Top             =   3120
      Width           =   495
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Hex"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   47
      Top             =   3120
      Width           =   735
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Dec"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   46
      Top             =   3120
      Width           =   735
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Oct"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   45
      Top             =   3120
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Bin"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   44
      Top             =   3120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   360
      TabIndex        =   43
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmdlog 
      Caption         =   "Antilog"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   42
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdlog 
      Caption         =   "Log10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   41
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmde 
      Caption         =   "e"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   40
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   39
      Top             =   720
      Width           =   855
   End
   Begin VB.CommandButton cmddegrad 
      Caption         =   "R2D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   38
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmddegrad 
      Caption         =   "D2R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   37
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdoperation 
      Caption         =   "Y  rt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1560
      TabIndex        =   36
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmdoperation 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   4920
      TabIndex        =   35
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdmemory 
      Caption         =   "Rcl"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   120
      TabIndex        =   34
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmdmemory 
      Caption         =   "RM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   33
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmdmemory 
      Caption         =   "M-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   32
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmdmemory 
      Caption         =   "M+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   31
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdsqr 
      Caption         =   "n!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   1560
      TabIndex        =   30
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdpi 
      Caption         =   "pi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   29
      Top             =   720
      Width           =   495
   End
   Begin VB.CommandButton cmdoperation 
      Caption         =   "X^Y"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   840
      TabIndex        =   28
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmdsqr 
      Caption         =   "1/X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1560
      TabIndex        =   27
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmdsqr 
      Caption         =   "sqrt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   840
      TabIndex        =   26
      Top             =   2160
      Width           =   615
   End
   Begin VB.CommandButton cmdsqr 
      Caption         =   "X^2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   25
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdsqr 
      Caption         =   "X^3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   24
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton cmdpn 
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   23
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton cmdlog 
      Caption         =   "Loge"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   840
      TabIndex        =   22
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   21
      Top             =   720
      Width           =   735
   End
   Begin VB.CommandButton cmdtrigno 
      Caption         =   "Cot"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2280
      TabIndex        =   20
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton cmdtrigno 
      Caption         =   "Tan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   19
      Top             =   1680
      Width           =   735
   End
   Begin VB.CommandButton cmdtrigno 
      Caption         =   "Cos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   18
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton cmdtrigno 
      Caption         =   "Sin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   17
      Top             =   720
      Width           =   735
   End
   Begin VB.TextBox txtResultado 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   840
      TabIndex        =   16
      Top             =   120
      Width           =   4815
   End
   Begin VB.CommandButton cmdnumber 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   3720
      TabIndex        =   15
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton cmdoperation 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   5520
      TabIndex        =   14
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmdoperation 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   5520
      TabIndex        =   13
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton cmdoperation 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   4920
      TabIndex        =   12
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmdoperation 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   5520
      TabIndex        =   11
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdoperation 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   4920
      TabIndex        =   10
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton cmdnumber 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   3120
      TabIndex        =   9
      Top             =   2640
      Width           =   495
   End
   Begin VB.CommandButton cmdnumber 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   4320
      TabIndex        =   8
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdnumber 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   3720
      TabIndex        =   7
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdnumber 
      BackColor       =   &H00FF00FF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   3120
      TabIndex        =   6
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdnumber 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   4320
      TabIndex        =   5
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton cmdnumber 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3720
      TabIndex        =   4
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton cmdnumber 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3120
      TabIndex        =   3
      Top             =   1680
      Width           =   495
   End
   Begin VB.CommandButton cmdnumber 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4320
      TabIndex        =   2
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmdnumber 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   1
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmdnumber 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3120
      TabIndex        =   0
      Top             =   2160
      Width           =   495
   End
End
Attribute VB_Name = "Calculadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op1 As Double
Dim res As Double
Dim op2 As Double
Dim n As Double
Dim i As Integer
Dim operator As String
Dim s As String
Dim c As String

Private Sub cmdclear_Click()
    txtResultado.Text = ""
End Sub

Private Sub cmddegrad_Click(Index As Integer)
    op1 = Val(txtResultado.Text)
    Select Case (Index)
    Case 0
        res = op1 * 3.142 / 180
    Case 1
        res = 180 * op1 / 3.142
    End Select
    txtResultado.Text = res
End Sub

Private Sub cmde_Click()
    txtResultado.Text = "2.71828182845904523536"
End Sub

Private Sub cmdlog_Click(Index As Integer)

    op1 = Val(txtResultado.Text)

    Select Case (Index)
    Case 0
        res = Log(op1)
    Case 1
        res = Log(op1) / Log(10)
    Case 2
        res = 10 ^ op1
    End Select
    txtResultado.Text = res
End Sub

Private Sub cmdmemory_Click(Index As Integer)
    Static m As Double

    Select Case (Index)
    Case 0
        m = m + Val(txtResultado.Text)
    Case 1
        m = m - Val(txtResultado.Text)
    Case 2
        m = 0
        txtResultado.Text = ""
    Case 3
        txtResultado.Text = m
        MsgBox m
    End Select

    If m = 0 Then
        Text1.Text = ""
    Else
        Text1.Text = "M"
    End If
End Sub

Private Sub cmdnumber_Click(Index As Integer)
    Select Case (Index)
    Case 0
        txtResultado.Text = txtResultado + "1"
    Case 1
        txtResultado.Text = txtResultado + "2"
    Case 2
        txtResultado.Text = txtResultado + "3"
    Case 3
        txtResultado.Text = txtResultado + "4"
    Case 4
        txtResultado.Text = txtResultado + "5"
    Case 5
        txtResultado.Text = txtResultado + "6"
    Case 6
        txtResultado.Text = txtResultado + "7"
    Case 7
        txtResultado.Text = txtResultado + "8"
    Case 8
        txtResultado.Text = txtResultado + "9"
    Case 9
        txtResultado.Text = txtResultado + "0"
    Case 10
        txtResultado.Text = txtResultado + "."
    Case 11
        txtResultado.Text = txtResultado + "A"
    Case 12
        txtResultado.Text = txtResultado + "B"
    Case 13
        txtResultado.Text = txtResultado + "C"
    Case 14
        txtResultado.Text = txtResultado + "D"
    Case 15
        txtResultado.Text = txtResultado + "E"
    Case 16
        txtResultado.Text = txtResultado + "F"
    End Select
End Sub

Private Sub cmdoperation_Click(Index As Integer)

    On Error GoTo trataerro

    Select Case (Index)
    Case 0
        op1 = Val(txtResultado.Text)
        operator = "/"
        txtResultado.Text = ""
    Case 1
        op1 = Val(txtResultado.Text)
        operator = "*"
        txtResultado.Text = ""
    Case 2
        op1 = Val(txtResultado.Text)
        operator = "+"
        txtResultado.Text = ""
    Case 3
        op1 = Val(txtResultado.Text)
        operator = "-"
        txtResultado.Text = ""
    Case 5
        op1 = Val(txtResultado.Text)
        operator = "^"
        txtResultado.Text = ""
    Case 6
        op1 = Val(txtResultado.Text)
        operator = "%"
        txtResultado.Text = ""
    Case 7
        op1 = Val(txtResultado.Text)
        operator = "$"
        txtResultado.Text = ""
    Case 4
        op2 = Val(txtResultado.Text)
        Select Case (operator)
        Case "+"
            res = op1 + op2
        Case "-"
            res = op1 - op2
        Case "*"
            res = op1 * op2
        Case "/"
            res = op1 / op2
        Case "^"
            res = op1 ^ op2
        Case "%"
            res = op1 / op2 * 100
        Case "$"
            res = op1 ^ (1 / op2)
        End Select
        txtResultado.Text = res
        op1 = res
    End Select
    Exit Sub
trataerro:
    MsgBox "Erro no processamento."
End Sub

Private Sub cmdpercent_Click()
    res = res * 100
    txtResultado.Text = res
End Sub

Private Sub cmdpi_Click()
    txtResultado.Text = 62832 / 20000            '22 / 7
End Sub

Private Sub cmdpn_Click()
    res = Val(txtResultado.Text)
    res = -(res)
    txtResultado.Text = res
End Sub

Private Sub cmdsqr_Click(Index As Integer)
    op1 = Val(txtResultado.Text)
    Select Case (Index)
    Case 0
        res = op1 * op1 * op1
    Case 1
        res = op1 * op1
    Case 2
        res = Sqr(op1)
    Case 3
        res = 1 / op1
    Case 4
        n = 1
        For i = 1 To op1
            n = i * n
        Next i
        res = n
    End Select
    txtResultado.Text = res
End Sub

Private Sub cmdtrigno_Click(Index As Integer)
    op1 = Val(txtResultado.Text)
    Select Case (Index)
    Case 0
        res = Sin(op1 * 3.142 / 180)
    Case 1
        res = Cos(op1 * 3.142 / 180)
    Case 2
        res = Tan(op1 * 3.142 / 180)
    Case 3
        res = 1 / Tan(op1 * 3.142 / 180)
    End Select
    txtResultado.Text = res
End Sub

Private Sub Command1_Click()
    s = txtResultado.Text
    c = Len(s)
    If c = 0 Then
        txtResultado.Text = ""
    Else
        c = c - 1
        txtResultado.Text = Left(s, c)
    End If
End Sub

Private Sub Form_Load()
    Center Me
    txtResultado.Text = "0"
End Sub

