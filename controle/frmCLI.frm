VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form frmCLI 
   Caption         =   "Controle de Desenho"
   ClientHeight    =   3840
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10200
   Icon            =   "frmCLI.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3840
   ScaleWidth      =   10200
   StartUpPosition =   2  'CenterScreen
   Begin vbExtra.FlexFn FlexFn1 
      Height          =   405
      Left            =   5880
      TabIndex        =   9
      Top             =   480
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   714
      PageNumbersFormat=   ""
      PageNumbersFormatIndex=   0
      BeginProperty PageNumbersFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Numero Cliente Engenharia"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Nome do Cliente Engenharia"
      Top             =   480
      Width           =   4335
   End
   Begin VBFLXGRD18.VBFlexGrid GridAtual 
      Height          =   2655
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Lista de Desenhos"
      Top             =   960
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4683
      _Version        =   393216
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   8400
      TabIndex        =   5
      Top             =   240
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmCLI.frx":058A
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
   Begin XPControls.XPButton novo 
      Height          =   435
      Left            =   8400
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmCLI.frx":0B24
      Caption         =   "Novo"
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
   Begin XPControls.XPButton editar 
      Height          =   435
      Left            =   8400
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmCLI.frx":10BE
      Caption         =   "Editar"
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
   Begin XPControls.XPButton excluir 
      Height          =   435
      Left            =   8400
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmCLI.frx":1658
      Caption         =   "Apagar"
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
      Caption         =   "Cliente:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nome:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   5
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmCLI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cARQDES As String
Dim cSQLDES As String
Dim aVAL As Variant
Dim aFOR As Variant
Dim aCAM As Variant
Dim aPAD As Variant
Dim nCAMPOS As Integer
Dim iLOOP As Integer

Private Sub cmdClose_Click()
'na inclusao preenchido aqui nao salva
'On Error Resume Next
'    For iLOOP = 0 To nCAMPOS - 1
'        aVAL(iLOOP) = txtFields(iLOOP)
'    Next iLOOP
'    GrvSQL cARQDES, cSQLDES, nCAMPOS, aCAM, aVAL, aFOR
'    Screen.MousePointer = vbDefault
'    Unload Me
End Sub

Private Sub Editar_Click()
  GridAtual.Col = 0
  nREVI = FixInt(GridAtual)
  FrmClii.Show vbModal, Me
  Filgrid
End Sub

Private Sub Encerrar_Click()
  Screen.MousePointer = vbDefault
  Unload Me

End Sub

Private Sub excluir_Click()
  Dim sSQL As String
  GridAtual.Col = 0
  nREVI = FixInt(GridAtual)
  'Conjunto
  sSQL = "select * from PRODUTO WHERE CLIENTE=" & nPF & " and ITEM=" & nREVI
  If ApagaSQLP(cARQDES, sSQL) Then
    Filgrid
  End If
End Sub

Private Sub Filgrid()
  Dim cSQL As String
  cSQL = "SELECT ITEM,CODIGO,NOME FROM PRODUTO WHERE CLIENTE=" & nPF
  MontaGridFast GridAtual, 3, Array(400, 2000, 4000), Array("IEB", "Produto", "Descrição"), _
                Array("ITEM", "L$CODIGO", "NOME"), cARQDES, cSQL
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
  cARQDES = PegPath("PATH", "DESENHO")
  cSQLDES = "select * from cliente WHERE cliente=" & nPF
  CenterFormToScreen Me
  nCAMPOS = 2
  aCAM = Array("CLIENTE", "CLINOME")
  aFOR = Array("NI", "C")
  aPAD = Array(0, "")
  aVAL = PegSQL(cARQDES, cSQLDES, nCAMPOS, aCAM, aFOR, aPAD)
  For iLOOP = 0 To nCAMPOS - 1
    TXTFIELDS(iLOOP) = aVAL(iLOOP)
  Next iLOOP
  Filgrid
End Sub

Private Sub Novo_Click()
  Dim cDESENHO As String
  Dim sSQL As String
  nREVI = 0
  cDESENHO = InputBox("Digite o Produto", "Inclusão Produto", "__")
  sSQL = "PRODUTO WHERE CLIENTE=" & nPF
  nREVI = FixInt(PegMAXSQL(cARQDES, sSQL, "item", 0)) + 1
  sSQL = "SELECT CLIENTE,CODIGO,ITEM FROM PRODUTO WHERE CLIENTE=" & nPF & " AND CODIGO='" & cDESENHO & "'"
  IncluiSQL cARQDES, sSQL, 3, Array("codigo", "item", "cliente"), Array(cDESENHO, nREVI, nPF), True, False
  If Not lRETU Then
    Alert "Produgo Ja Cadastrado"
    Exit Sub
  Else
    Filgrid
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub


