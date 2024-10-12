VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form frmprot 
   Caption         =   "Controle de Prototipos"
   ClientHeight    =   4035
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10425
   Icon            =   "frmprot.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   10425
   StartUpPosition =   2  'CenterScreen
   Begin vbExtra.FlexFn FlexFn1 
      Height          =   405
      Left            =   6240
      TabIndex        =   10
      Top             =   360
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
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Nome do Cliente Engenharia"
      Top             =   480
      Width           =   4335
   End
   Begin VBFLXGRD18.VBFlexGrid GridAtual 
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Lista de Desenhos"
      Top             =   960
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5106
      _Version        =   393216
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   8520
      TabIndex        =   5
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmprot.frx":058A
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
      Left            =   8520
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmprot.frx":0B24
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
   Begin XPControls.XPButton Editar 
      Height          =   435
      Left            =   8520
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmprot.frx":10BE
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
      Left            =   8520
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmprot.frx":1658
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
   Begin XPControls.XPButton CmdLocalizar 
      Height          =   375
      Left            =   8520
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Picture         =   "frmprot.frx":1BF2
      Caption         =   "Localizar"
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
      Left            =   1680
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmprot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'DIM cARQ AS STRING
Dim cSQL As String
Dim aVAL As Variant
Dim aFOR As Variant
Dim aCAM As Variant
Dim aPAD As Variant
Dim nCAMPOS As Integer
Dim iLOOP As Integer
Dim cARQDES As String

Private Sub CmdLocalizar_Click()
  LocalizaGri1 GridAtual
End Sub

Private Sub Editar_Click()
  GridAtual.Col = 0
  cDESENHO = FixStr(GridAtual)
  frmproti.Show vbModal, Me
  Filgrid
End Sub

Private Sub Encerrar_Click()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub excluir_Click()
  Dim sSQL As String
  GridAtual.Col = 0
  cDESENHO = FixStr(GridAtual)
  '******************************************************************
  'Conjunto
  sSQL = "select * from PROTO WHERE PARTNUMBER='" & cDESENHO & "'"
  '******************************************************************
  If ApagaSQLP(cARQDES, sSQL) Then
    Filgrid
  End If
End Sub

Private Sub Filgrid()
  Dim cSQL As String
  cSQL = "SELECT PARTNUMBER,NOME,PROJETO FROM PROTO WHERE CLIENTE=" & nPF
  MontaGridFast GridAtual, 3, Array(2000, 4000, 2000), Array("Produto", "Descrição", "Projeto"), _
                Array("L$PARTNUMBER", "NOME", "PROJETO"), cARQDES, cSQL
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  cARQDES = PegPath("PATH", "DESENHO")
  cSQL = "select * from cliente WHERE cliente=" & nPF
  nCAMPOS = 2
  aCAM = Array("cliente", "clinome")
  aFOR = Array("NI", "C")
  aPAD = Array(0, "")
  aVAL = PegSQL(cARQDES, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
  For iLOOP = 0 To nCAMPOS - 1
    TXTFIELDS(iLOOP) = aVAL(iLOOP)
  Next iLOOP
  Filgrid
End Sub

Private Sub Novo_Click()
  Dim cDESENHO As String
  Dim sSQL As String

  cDESENHO = InputBox("Digite o Produto", "Inclusão Produto", "__")
  sSQL = "SELECT * FROM PROTO WHERE PARTNUMBER='" & cDESENHO & "'"
  If IncluiSQL(cARQDES, sSQL, 3, Array("PARTNUMBER", "CLIENTE", "CLINOME"), _
               Array(cDESENHO, nPF, cCLINOME), True, True) Then

    Filgrid
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault

End Sub


