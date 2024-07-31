VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form frmPro 
   Caption         =   "Itens "
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8775
   Icon            =   "frmPRO.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   8
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      HelpContextID   =   90
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   4
      Left            =   1920
      MaxLength       =   40
      TabIndex        =   0
      Top             =   240
      Width           =   5115
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   7080
      TabIndex        =   10
      Top             =   720
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmPRO.frx":058A
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
      Left            =   7080
      TabIndex        =   11
      Top             =   240
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmPRO.frx":0B24
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
   Begin XPControls.XPButton esciedx 
      Height          =   375
      Left            =   960
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "frmPRO.frx":10BE
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
   Begin VB.Label lblLabels 
      Caption         =   "Codigo"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   4
      Left            =   1800
      TabIndex        =   9
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descrição"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   7
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Caption         =   "Numero"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   720
      TabIndex        =   6
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cliente"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Caption         =   "Tipo"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frmPro"
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
      aVAL(iLOOP) = txtFields(iLOOP)
    Next iLOOP
    GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR, 2 ' PULA 0 NUMERO 1 TIPO Chaves
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

Private Sub esciedx_Click()
  iMU01 = 201
  escIED.Show vbModal, Me
  If lRETU Then
    txtFields(0) = eRETU02
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  cARQ = PegPath("PATH", "DESENHO")
  cSQL = "select * from PRODENG WHERE NUMERO=" & nPF & " AND TIPO='" & cCONJUNTO & "'"
  nCAMPOS = 5
  aCAM = Array("NUMERO", "TIPO", "RESCLI", "PRODUTO", "NOME")
  aFOR = Array("NI", "C", "C", "C", "C")
  aPAD = Array(0, "", "", "", "")
  aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
  For iLOOP = 0 To nCAMPOS - 1
    txtFields(iLOOP) = aVAL(iLOOP)
  Next iLOOP

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

