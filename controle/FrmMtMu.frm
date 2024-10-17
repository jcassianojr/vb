VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form FrmMtMu 
   Caption         =   "Materia Prima / Componentes"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10200
   Icon            =   "FrmMtMu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   4
      Left            =   6480
      TabIndex        =   9
      Top             =   2100
      Width           =   1275
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   3
      Left            =   3900
      TabIndex        =   8
      Top             =   2100
      Width           =   1275
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   7
      Top             =   2100
      Width           =   1275
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   5
      Left            =   6300
      TabIndex        =   1
      Top             =   60
      Width           =   1275
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   9
      Left            =   3780
      TabIndex        =   0
      Top             =   60
      Width           =   615
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   7
      Left            =   1200
      TabIndex        =   5
      Top             =   1320
      Width           =   6855
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   0
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   2
      Top             =   420
      Width           =   6855
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   6
      Left            =   1200
      TabIndex        =   4
      Top             =   1020
      Width           =   6855
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   1
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   3
      Top             =   720
      Width           =   6855
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   8
      Left            =   1200
      TabIndex        =   6
      Top             =   1680
      Width           =   6855
   End
   Begin VB.TextBox txtcodigo 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   60
      Width           =   2055
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   8400
      TabIndex        =   21
      Top             =   840
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FrmMtMu.frx":058A
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
      Left            =   8400
      TabIndex        =   22
      Top             =   240
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FrmMtMu.frx":0B24
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
   Begin XPControls.XPButton Command1 
      Height          =   375
      Left            =   4560
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   0
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "FrmMtMu.frx":10BE
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
      Caption         =   "Espessura"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   9
      Left            =   5280
      TabIndex        =   19
      Top             =   2100
      Width           =   1275
   End
   Begin VB.Label Label 
      Caption         =   "Largura"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   8
      Left            =   2700
      TabIndex        =   18
      Top             =   2100
      Width           =   1275
   End
   Begin VB.Label Label 
      Caption         =   "Comprimento"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   17
      Top             =   2100
      Width           =   1275
   End
   Begin VB.Label Label 
      Caption         =   "Peso Liquido"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   5
      Left            =   5220
      TabIndex        =   16
      Top             =   60
      Width           =   1035
   End
   Begin VB.Label Label 
      Caption         =   "Unidade"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   4
      Left            =   3060
      TabIndex        =   15
      Top             =   60
      Width           =   675
   End
   Begin VB.Label Label 
      Caption         =   "Aplicacao"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   7
      Left            =   60
      TabIndex        =   14
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "Instrução"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   3
      Left            =   60
      TabIndex        =   13
      Top             =   1380
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "locacao"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   60
      TabIndex        =   12
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "Nome"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   420
      Width           =   675
   End
   Begin VB.Label Label 
      Caption         =   "Codigo"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   60
      Width           =   615
   End
End
Attribute VB_Name = "FrmMtMu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cARQ As String
Dim cARQMTMU As String
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
    For iLOOP = 0 To nCAMPOS - 1             ''tira o campos  nao gravar o numero,nome
      aVAL(iLOOP) = TXTFIELDS(iLOOP)       ''vb matriz zero
    Next iLOOP
    GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR
  End If
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub Command1_Click()
  ePASS01 = "MANA5EMPFOX"
  ePASS02 = "SELECT UNIDADE as codigo,UNIDDES as nome from MD07"
  ePASS03 = Array("UNIDADE", "UNIDDES")
  EscCodNomSim.Show vbModal, Me
  If lRETU Then
    TXTFIELDS(9) = eRETU01
  End If
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
  CenterFormToScreen Me
  TxtCodigo.tEXT = ePASS01
  cARQMTMU = ePASS02
  cARQ = GeraConn(zMANA5EMP, "JETFOX")
  cSQL = "select * from " & cARQMTMU & " WHERE CODIGO='" & Trim(ePASS01) & "'"
  nCAMPOS = 10
  aCAM = Array("NOME", "NOM2", "DIMX", "DIMY", "DIMZ", "PESLIQ", "APLICACAO", "INSTRU", "LOCACAO", "UNIDADE")
  aFOR = Array("C", "C", "N", "N", "N", "N", "C", "C", "C", "C")
  aPAD = Array("", "", 0, 0, 0, 0, "", "", "", "")
  aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
  For iLOOP = 0 To nCAMPOS - 1
    TXTFIELDS(iLOOP) = aVAL(iLOOP)
  Next iLOOP
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub txtFields_GotFocus(Index As Integer)
  FocusMe
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index = 2 Or Index = 3 Or Index = 4 Or Index = 5 Then
    KeyAscii = ValiText(KeyAscii, "#N")
  End If
End Sub

