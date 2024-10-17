VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form FrmClii 
   Caption         =   "Itens "
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8430
   Icon            =   "FrmClii.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2085
   ScaleWidth      =   8430
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Codigo Cliente"
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Codigo IEB"
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   2
      Left            =   120
      MaxLength       =   24
      TabIndex        =   2
      ToolTipText     =   "Codigo"
      Top             =   840
      Width           =   1575
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   3
      Left            =   1920
      MaxLength       =   40
      TabIndex        =   1
      ToolTipText     =   "Descriçao do Codigo"
      Top             =   840
      Width           =   5115
   End
   Begin VB.CheckBox inativo 
      Caption         =   "Inativo"
      Height          =   195
      Left            =   2760
      TabIndex        =   0
      ToolTipText     =   "Se A Peça Esta Ativa"
      Top             =   240
      Width           =   1095
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   4560
      TabIndex        =   8
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "FrmClii.frx":058A
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
   Begin XPControls.XPButton CmdClose 
      Height          =   435
      Left            =   6360
      TabIndex        =   9
      Top             =   240
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FrmClii.frx":0B24
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
      Caption         =   "Cliente"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   7
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "IEB"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Código:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "FrmClii"
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
  For iLOOP = 0 To nCAMPOS - 2
    aVAL(iLOOP) = txtFields(iLOOP)
  Next iLOOP
  aVAL(4) = FixNumBol(inativo.Value)
  GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR, 2 '0 item 1 cliente chaves
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
  cARQ = PegPath("PATH", "DESENHO")
  cSQL = "select * from PRODUTO WHERE CLIENTE=" & nPF & " and ITEM=" & nREVI
  nCAMPOS = 5
  aCAM = Array("ITEM", "CLIENTE", "CODIGO", "NOME", "INATIVO")
  aFOR = Array("NI", "NI", "C", "C", "BN")
  aPAD = Array(0, 0, "", "", False)
  aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
  For iLOOP = 0 To nCAMPOS - 2
    txtFields(iLOOP) = aVAL(iLOOP)
  Next iLOOP
  inativo.Value = aVAL(4)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

