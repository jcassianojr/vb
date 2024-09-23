VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form FrmMsrd 
   Caption         =   "Registro de Defeito"
   ClientHeight    =   3090
   ClientLeft      =   150
   ClientTop       =   510
   ClientWidth     =   8325
   Icon            =   "FrmMsrd.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtFIELDS 
      Height          =   330
      Index           =   5
      Left            =   1080
      TabIndex        =   10
      Top             =   2160
      Width           =   5895
   End
   Begin VB.TextBox txtFIELDS 
      Height          =   330
      Index           =   4
      Left            =   1080
      TabIndex        =   9
      Top             =   1800
      Width           =   5895
   End
   Begin VB.TextBox txtFIELDS 
      Height          =   330
      Index           =   3
      Left            =   1080
      TabIndex        =   8
      Top             =   1440
      Width           =   5895
   End
   Begin VB.TextBox txtFIELDS 
      Height          =   330
      Index           =   2
      Left            =   1080
      TabIndex        =   7
      Top             =   1080
      Width           =   5895
   End
   Begin VB.TextBox txtFIELDS 
      Height          =   330
      Index           =   1
      Left            =   1080
      TabIndex        =   6
      Top             =   720
      Width           =   3135
   End
   Begin VB.TextBox txtFIELDS 
      Height          =   330
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   5400
      TabIndex        =   11
      Top             =   600
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FrmMsrd.frx":058A
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
      Left            =   5400
      TabIndex        =   12
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FrmMsrd.frx":0B24
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
   Begin VB.Label Label5 
      Caption         =   "Acao"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Inv"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Descrição"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Ocorreu"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Data"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "FrmMsrd"
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
    For iLOOP = 0 To nCAMPOS - 1             ''tira o campos  nao gravar o numero,nome
      aVAL(iLOOP) = TXTFIELDS(iLOOP)       ''vb matriz zero
    Next iLOOP
    GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR
  End If
  Screen.MousePointer = vbDefault
  Unload Me
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
  cARQ = PegPath("PATH", "MANA5TGQ")
  cARQ = GeraConn(cARQ, "JETFOX")
  cSQL = "select * from MSRD WHERE MSRD=" & ePASS01
  nCAMPOS = 6
  aCAM = Array("DATA", "OCORREU", "DESC01", "DESC02", "INV01", "ACA01")
  aFOR = Array("D", "C", "C", "C", "C", "C")
  aPAD = Array("", "", "", "", "", "")
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

