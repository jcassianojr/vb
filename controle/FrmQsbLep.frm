VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "xpcontrols.ocx"
Begin VB.Form FrmQsbLep 
   Caption         =   "QSBLEP"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdescIED 
      Height          =   315
      Index           =   403
      Left            =   7920
      Picture         =   "FrmQsbLep.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   73
      TabStop         =   0   'False
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton cmdescIED 
      Height          =   315
      Index           =   402
      Left            =   7920
      Picture         =   "FrmQsbLep.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdescIED 
      Height          =   315
      Index           =   401
      Left            =   7920
      Picture         =   "FrmQsbLep.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   315
      Index           =   5
      Left            =   6360
      Picture         =   "FrmQsbLep.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   315
      Index           =   4
      Left            =   6360
      Picture         =   "FrmQsbLep.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   315
      Index           =   3
      Left            =   5280
      Picture         =   "FrmQsbLep.frx":1BB2
      Style           =   1  'Graphical
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   3360
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Height          =   315
      Index           =   2
      Left            =   5280
      Picture         =   "FrmQsbLep.frx":213C
      Style           =   1  'Graphical
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   3000
      Width           =   375
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   16
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   3000
      Width           =   555
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   15
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   3000
      Width           =   555
   End
   Begin VB.CommandButton CMDSIG 
      Height          =   255
      Index           =   1
      Left            =   2160
      Picture         =   "FrmQsbLep.frx":26C6
      Style           =   1  'Graphical
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton CMDSIG 
      Height          =   255
      Index           =   2
      Left            =   2760
      Picture         =   "FrmQsbLep.frx":2C50
      Style           =   1  'Graphical
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton CMDSIG 
      Height          =   255
      Index           =   3
      Left            =   3360
      Picture         =   "FrmQsbLep.frx":31DA
      Style           =   1  'Graphical
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton CMDSIG 
      Height          =   255
      Index           =   4
      Left            =   3960
      Picture         =   "FrmQsbLep.frx":3764
      Style           =   1  'Graphical
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton CMDSIG 
      Height          =   255
      Index           =   5
      Left            =   4560
      Picture         =   "FrmQsbLep.frx":3CEE
      Style           =   1  'Graphical
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton CMDSIG 
      Height          =   255
      Index           =   6
      Left            =   5160
      Picture         =   "FrmQsbLep.frx":4278
      Style           =   1  'Graphical
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton CMDSIG 
      Height          =   255
      Index           =   7
      Left            =   5760
      Picture         =   "FrmQsbLep.frx":4802
      Style           =   1  'Graphical
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton CMDSIG 
      Height          =   255
      Index           =   8
      Left            =   6360
      Picture         =   "FrmQsbLep.frx":4D8C
      Style           =   1  'Graphical
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   14
      Left            =   1560
      TabIndex        =   54
      Top             =   2640
      Width           =   6255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "KCC"
      Height          =   315
      Index           =   5
      Left            =   2880
      TabIndex        =   53
      Top             =   3720
      Width           =   555
   End
   Begin VB.CommandButton Command3 
      Caption         =   "KCC"
      Height          =   315
      Index           =   2
      Left            =   1200
      TabIndex        =   52
      Top             =   3720
      Width           =   555
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PQC"
      Height          =   315
      Index           =   4
      Left            =   2880
      TabIndex        =   51
      Top             =   3360
      Width           =   555
   End
   Begin VB.CommandButton Command3 
      Caption         =   "KPC"
      Height          =   315
      Index           =   3
      Left            =   2880
      TabIndex        =   50
      Top             =   3000
      Width           =   555
   End
   Begin VB.CommandButton Command1 
      Height          =   315
      Index           =   0
      Left            =   4200
      Picture         =   "FrmQsbLep.frx":5316
      Style           =   1  'Graphical
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "PQC"
      Height          =   315
      Index           =   1
      Left            =   1200
      TabIndex        =   48
      Top             =   3360
      Width           =   555
   End
   Begin VB.CommandButton Command3 
      Caption         =   "KPC"
      Height          =   315
      Index           =   0
      Left            =   1200
      TabIndex        =   47
      Top             =   3000
      Width           =   555
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   345
      Index           =   2
      Left            =   120
      TabIndex        =   46
      Top             =   3000
      Width           =   915
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   13
      Left            =   6120
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   120
      Width           =   555
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   12
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   120
      Width           =   555
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   11
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   120
      Width           =   555
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   10
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   120
      Width           =   555
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   9
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   120
      Width           =   555
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   8
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   120
      Width           =   555
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   7
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   120
      Width           =   555
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   6
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   120
      Width           =   555
   End
   Begin VB.CommandButton CmdZera 
      Caption         =   "B"
      Height          =   195
      Index           =   7
      Left            =   6120
      TabIndex        =   37
      ToolTipText     =   "Limpar Posicao"
      Top             =   840
      Width           =   315
   End
   Begin VB.CommandButton CmdZera 
      Caption         =   "B"
      Height          =   195
      Index           =   6
      Left            =   5520
      TabIndex        =   36
      ToolTipText     =   "Limpar Posicao"
      Top             =   840
      Width           =   315
   End
   Begin VB.CommandButton CmdZera 
      Caption         =   "B"
      Height          =   195
      Index           =   5
      Left            =   4920
      TabIndex        =   35
      ToolTipText     =   "Limpar Posicao"
      Top             =   840
      Width           =   315
   End
   Begin VB.CommandButton CmdZera 
      Caption         =   "B"
      Height          =   195
      Index           =   4
      Left            =   4320
      TabIndex        =   34
      ToolTipText     =   "Limpar Posicao"
      Top             =   840
      Width           =   315
   End
   Begin VB.CommandButton CmdZera 
      Caption         =   "B"
      Height          =   195
      Index           =   3
      Left            =   3660
      TabIndex        =   33
      ToolTipText     =   "Limpar Posicao"
      Top             =   840
      Width           =   315
   End
   Begin VB.CommandButton CmdZera 
      Caption         =   "B"
      Height          =   195
      Index           =   2
      Left            =   3120
      TabIndex        =   32
      ToolTipText     =   "Limpar Posicao"
      Top             =   840
      Width           =   315
   End
   Begin VB.CommandButton CmdZera 
      Caption         =   "B"
      Height          =   195
      Index           =   1
      Left            =   2520
      TabIndex        =   31
      ToolTipText     =   "Limpar Posicao"
      Top             =   840
      Width           =   315
   End
   Begin VB.CommandButton CmdZera 
      Caption         =   "B"
      Height          =   195
      Index           =   0
      Left            =   1920
      TabIndex        =   30
      ToolTipText     =   "Limpar Posicao"
      Top             =   840
      Width           =   315
   End
   Begin VB.CommandButton CmdFluxo 
      Caption         =   "8"
      Height          =   255
      Index           =   7
      Left            =   6120
      Picture         =   "FrmQsbLep.frx":58A0
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   600
      Width           =   315
   End
   Begin VB.CommandButton CmdFluxo 
      Caption         =   "7"
      Height          =   255
      Index           =   6
      Left            =   5520
      Picture         =   "FrmQsbLep.frx":5E2A
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   600
      Width           =   315
   End
   Begin VB.CommandButton CmdFluxo 
      Caption         =   "6"
      Height          =   255
      Index           =   5
      Left            =   4920
      Picture         =   "FrmQsbLep.frx":63B4
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   600
      Width           =   315
   End
   Begin VB.CommandButton CmdFluxo 
      Caption         =   "5"
      Height          =   255
      Index           =   4
      Left            =   4320
      Picture         =   "FrmQsbLep.frx":693E
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   600
      Width           =   315
   End
   Begin VB.CommandButton CmdFluxo 
      Caption         =   "4"
      Height          =   255
      Index           =   3
      Left            =   3720
      Picture         =   "FrmQsbLep.frx":6EC8
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   600
      Width           =   315
   End
   Begin VB.CommandButton CmdFluxo 
      Caption         =   "3"
      Height          =   255
      Index           =   2
      Left            =   3120
      Picture         =   "FrmQsbLep.frx":7452
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   600
      Width           =   315
   End
   Begin VB.CommandButton CmdFluxo 
      Caption         =   "2"
      Height          =   255
      Index           =   1
      Left            =   2520
      Picture         =   "FrmQsbLep.frx":79DC
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   600
      Width           =   315
   End
   Begin VB.CommandButton CmdFluxo 
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   1920
      Picture         =   "FrmQsbLep.frx":7F66
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   600
      Width           =   315
   End
   Begin VB.CommandButton Command2 
      Caption         =   "(Q)ualidade"
      Height          =   255
      Index           =   8
      Left            =   4620
      TabIndex        =   19
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "(V)isual"
      Height          =   255
      Index           =   7
      Left            =   3780
      TabIndex        =   18
      Top             =   1560
      Width           =   795
   End
   Begin VB.CommandButton Command2 
      Caption         =   "(M)anual"
      Height          =   255
      Index           =   6
      Left            =   2940
      TabIndex        =   17
      Top             =   1560
      Width           =   795
   End
   Begin VB.CommandButton Command2 
      Caption         =   "(A)utomatica"
      Height          =   255
      Index           =   5
      Left            =   1680
      TabIndex        =   16
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "(I)dentificação"
      Height          =   255
      Index           =   4
      Left            =   4680
      TabIndex        =   15
      Top             =   1260
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ma(N)useio"
      Height          =   255
      Index           =   3
      Left            =   3720
      TabIndex        =   14
      Top             =   1260
      Width           =   915
   End
   Begin VB.CommandButton Command2 
      Caption         =   "(F)erramenta"
      Height          =   255
      Index           =   1
      Left            =   2580
      TabIndex        =   13
      Top             =   1260
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "(P)roduto"
      Height          =   255
      Index           =   0
      Left            =   1740
      TabIndex        =   12
      Top             =   1260
      Width           =   795
   End
   Begin VB.CommandButton Command1 
      Height          =   315
      Index           =   1
      Left            =   4200
      Picture         =   "FrmQsbLep.frx":84F0
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3360
      Width           =   375
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   4
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3000
      Width           =   555
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   5
      Left            =   1560
      TabIndex        =   2
      Top             =   2280
      Width           =   6255
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox txtItem 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   180
      Width           =   375
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   0
      Top             =   1920
      Width           =   6255
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   405
      Index           =   3
      Left            =   1920
      TabIndex        =   1
      Top             =   3000
      Width           =   855
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   6720
      TabIndex        =   69
      Top             =   600
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FrmQsbLep.frx":8A7A
      Caption         =   "Retornar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton cmdClose 
      Height          =   435
      Left            =   6720
      TabIndex        =   70
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FrmQsbLep.frx":9014
      Caption         =   "Salvar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblLabels 
      Caption         =   "Inspeção"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   8
      Left            =   780
      TabIndex        =   29
      Top             =   1500
      Width           =   795
   End
   Begin VB.Label lblLabels 
      Caption         =   "Mudança"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   4
      Left            =   780
      TabIndex        =   28
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "Requisitos Entrada"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   1395
   End
   Begin VB.Label lblLabels 
      Caption         =   "Fluxo"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   3
      Left            =   1260
      TabIndex        =   8
      Top             =   180
      Width           =   555
   End
   Begin VB.Label lblLabels 
      Caption         =   "ORD:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descrição:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Requisitos Saida"
      ForeColor       =   &H00C00000&
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   1275
   End
End
Attribute VB_Name = "FrmQsbLep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cARQ As String
Dim cSQL As String
Dim aVAL As Variant
Dim aFOR As Variant
Dim aCAM As Variant
Dim aPAD As Variant
Dim nCAMPOS As Integer
Dim iLOOP As Integer

Private Sub cmdClose_Click()
    On Error Resume Next
    If MDG("Gravar alteraçôes") Then
        For iLOOP = 0 To nCAMPOS - 1
            aVAL(iLOOP) = TXTFIELDS(iLOOP)
        Next iLOOP
        GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR
    End If
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub cmdescIED_Click(Index As Integer)
    iMU01 = Index
    escIED.Show vbModal, Me
    If lRETU Then
        Select Case Index
        Case 401
            TXTFIELDS(1) = TXTFIELDS(1) & " " & eRETU02
        Case 402
            TXTFIELDS(5) = TXTFIELDS(5) & " " & eRETU02
        Case 403
            TXTFIELDS(14) = TXTFIELDS(14) & " " & eRETU02
        End Select
    End If
End Sub

Private Sub CmdFluxo_Click(Index As Integer)
    escFLX.Show vbModal, Me
    If lRETU Then
        TXTFIELDS(Index + 6) = eRETU01
    End If
End Sub

Private Sub CMDSIG_Click(Index As Integer)
    lRETU = False
    eRETU01 = ""
    eRETU02 = ""
    ePASS01 = "isoqsymbol"
    ePASS02 = True
    frmCharacters.Show vbModal, Me
    If lRETU Then
        TXTFIELDS(Index + 6) = eRETU01
    End If
End Sub

Private Sub CmdZera_Click(Index As Integer)
    TXTFIELDS(Index + 6) = " "
End Sub

Private Sub Command1_Click(Index As Integer)
    lRETU = False
    eRETU01 = ""
    eRETU02 = ""
    If Index = 0 Or Index = 2 Or Index = 4 Then
        escFLX.Show vbModal, Me
    Else
        ePASS01 = "isoqsymbol"
        ePASS02 = True
        frmCharacters.Show vbModal, Me
    End If
    If lRETU Then
        If Index = 0 Or Index = 1 Then
            Me.TXTFIELDS(4) = eRETU01
        End If
        If Index = 2 Or Index = 3 Then
            Me.TXTFIELDS(15) = eRETU01
        End If
        If Index = 4 Or Index = 5 Then
            Me.TXTFIELDS(16) = eRETU01
        End If
        
    End If
End Sub

Private Sub Command2_Click(Index As Integer)
    Select Case Index
    Case 0
        TXTFIELDS(0) = "P"
    Case 1
        TXTFIELDS(0) = "F"
    Case 2
        TXTFIELDS(0) = "S"
    Case 3
        TXTFIELDS(0) = "N"
    Case 4
        TXTFIELDS(0) = "I"
    Case 5
        TXTFIELDS(0) = "A"
    Case 6
        TXTFIELDS(0) = "M"
    Case 7
        TXTFIELDS(0) = "V"
    Case 8
        TXTFIELDS(0) = "Q"
    
    End Select
End Sub

Private Sub Command3_Click(Index As Integer)
    Select Case Index
    Case 0
        TXTFIELDS(2) = "KPC"
    Case 1
        TXTFIELDS(2) = "PQC"
    Case 2
        TXTFIELDS(2) = "KCC"
    Case 3
        TXTFIELDS(3) = "KPC"
    Case 4
        TXTFIELDS(3) = "PQC"
    Case 5
        TXTFIELDS(3) = "KCC"
    End Select
End Sub

Private Sub Encerrar_Click()
    If Not MDG("Sair sem gravar") Then
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
    Dim X As Integer
    Center Me
    'cARQ = PegPath("PATH", "PF")
    cARQ = Sdb                                   ''Agora sdb pois pode ser pf pfp pfg
    txtItem.Text = nORD
    cSQL = "select * from PFQSBLEP WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ & " AND ITEM=" & nORD
    nCAMPOS = 17
    aCAM = Array("MUD", "DESCRICAO", "CLASS01", "CLASS02", "CARAC", "REQUER", _
                 "FLX01", "FLX02", "FLX03", "FLX04", "FLX05", "FLX06", "FLX07", "FLX08", "REQUERSAI", "CARA2", "CARA3")
    aFOR = Array("C", "C", "C", "C", "C", "C", "C", "C", "C", "C", "C", "C", "C", "C", "C", "C", "C")
    aPAD = Array("", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "")
    aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
    For iLOOP = 0 To nCAMPOS - 1
        TXTFIELDS(iLOOP) = aVAL(iLOOP)
    Next iLOOP
    TXTFIELDS(4).Font = "isoqsymbol"
    For X = 6 To 13
        TXTFIELDS(X).Font = "isoqsymbol"
    Next X
    TXTFIELDS(15).Font = "isoqsymbol"
    TXTFIELDS(16).Font = "isoqsymbol"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub


