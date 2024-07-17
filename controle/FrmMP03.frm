VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form FrmMP03 
   Caption         =   "Tratamentos"
   ClientHeight    =   3285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10275
   Icon            =   "FrmMP03.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   10275
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Mao (O)bra"
      Height          =   255
      Index           =   2
      Left            =   8280
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   660
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "(T)ermico"
      Height          =   255
      Index           =   1
      Left            =   8280
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "(S)uperficial"
      Height          =   255
      Index           =   0
      Left            =   8280
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   60
      Width           =   1095
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   9
      Left            =   7860
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   7
      Left            =   1260
      TabIndex        =   8
      Top             =   2280
      Width           =   6855
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   6
      Left            =   1260
      TabIndex        =   7
      Top             =   1980
      Width           =   6855
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   315
      Index           =   0
      Left            =   2280
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   4
      Left            =   6480
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   2
      Left            =   1260
      TabIndex        =   4
      Top             =   1080
      Width           =   6855
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   3
      Left            =   1260
      TabIndex        =   5
      Top             =   1380
      Width           =   6855
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   1
      Left            =   1260
      TabIndex        =   3
      Top             =   780
      Width           =   6855
   End
   Begin VB.TextBox TXTcodigo 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   5
      Left            =   1260
      TabIndex        =   6
      Top             =   1680
      Width           =   6855
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   315
      Index           =   8
      Left            =   4380
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   8280
      TabIndex        =   23
      Top             =   1560
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FrmMP03.frx":058A
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
      Left            =   8280
      TabIndex        =   24
      Top             =   1080
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FrmMP03.frx":0B24
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
   Begin XPControls.XPButton ESCMP03 
      Height          =   375
      Left            =   3480
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   0
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "FrmMP03.frx":10BE
      Caption         =   ""
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
   Begin VB.Label Label 
      Caption         =   "Tipo"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   4
      Left            =   7860
      TabIndex        =   18
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label 
      Caption         =   "Cognome"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   6
      Left            =   6480
      TabIndex        =   17
      Top             =   60
      Width           =   1335
   End
   Begin VB.Label Label 
      Caption         =   "Codigo"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label 
      Caption         =   "Nome Compras"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   780
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "Codigo Interno"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   4380
      TabIndex        =   14
      Top             =   60
      Width           =   1995
   End
   Begin VB.Label Label 
      Caption         =   "Nome"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   1140
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "OBS:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   12
      Top             =   1740
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "Aplicacao"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "Norma"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   8
      Left            =   2280
      TabIndex        =   10
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "FrmMP03"
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

Private Sub Command1_Click(Index As Integer)
  Select Case Index
  Case 0
    TXTFIELDS(9).tEXT = "S"
  Case 1
    TXTFIELDS(9).tEXT = "T"
  Case 2
    TXTFIELDS(9).tEXT = "O"
  End Select
End Sub

Private Sub Encerrar_Click()
  If Not MDG("Sair sem gravar") Then
    Exit Sub
  End If
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub ESCMP03_Click()
  Dim oFORM As New escmp
  iMU01 = 4
  oFORM.Show vbModal, Me
  If lRETU Then
    TXTFIELDS(0) = eRETU01
    TXTFIELDS(2) = eRETU02
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  cARQ = GeraConn(zMANA5EMP, "SDECDX")
  cSQL = "select * from MP03 WHERE CODIGO='" & Trim(ePASS01) & "'"
  TxtCodigo = ePASS01
  nCAMPOS = 10
  aCAM = Array("NORMA", "NOME", "NOM2", "APLICACAO", "COGNOME", "OBS01", "OBS02", "OBS03", "CODIGOINT", "TIPTRA")
  aFOR = Array("C", "C", "C", "C", "C", "C", "C", "C", "C", "C")
  aPAD = Array("", "", "", "", "", "", "", "", "", "")
  aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
  For iLOOP = 0 To nCAMPOS - 1
    TXTFIELDS(iLOOP) = aVAL(iLOOP)
  Next iLOOP
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

