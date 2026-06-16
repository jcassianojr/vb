VERSION 5.00
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "vbccr18.ocx"
Begin VB.Form FrmClii 
   Caption         =   "Itens "
   ClientHeight    =   2088
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   8436
   Icon            =   "FrmClii.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2088
   ScaleWidth      =   8436
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
   Begin VBCCR18.CommandButtonW Encerrar 
      Height          =   612
      Left            =   5760
      TabIndex        =   8
      Top             =   120
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Retornar"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "FrmClii.frx":058A
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW cmdClose 
      Height          =   612
      Left            =   7080
      TabIndex        =   9
      Top             =   120
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Salvar"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "FrmClii.frx":0B24
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
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
    aVAL(iLOOP) = TXTFIELDS(iLOOP)
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
    TXTFIELDS(iLOOP) = aVAL(iLOOP)
  Next iLOOP
  inativo.Value = aVAL(4)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

