VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{66E63055-5A66-4C79-9327-4BC077858695}#14.0#0"; "newtab01.OCX"
Object = "{075212A8-C1CF-444E-939D-F6046CCDBC08}#1.0#0"; "VBFLXGRD18.OCX"
Begin VB.Form frmDes 
   Caption         =   "Controle de Desenho"
   ClientHeight    =   6135
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10965
   Icon            =   "frmDes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   10965
   StartUpPosition =   2  'CenterScreen
   Begin NewTabCtl.NewTab SSTab1 
      Height          =   4095
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   7223
      ControlJustAdded=   0   'False
      Tabs            =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabsPerRow      =   2
      TabHeight       =   520
      ControlVersion  =   9
      TabCaption(0)   =   "Atual"
      Tab(0).ControlCount=   6
      Tab(0).Control(0)=   "arquivar"
      Tab(0).Control(1)=   "NovaREv"
      Tab(0).Control(2)=   "GridAtual"
      Tab(0).Control(3)=   "novo"
      Tab(0).Control(4)=   "editar"
      Tab(0).Control(5)=   "excluir"
      TabCaption(1)   =   "Arquivadas"
      Tab(1).ControlCount=   1
      Tab(1).Control(0)=   "GridArq"
      Begin XPControls.XPButton arquivar 
         Height          =   375
         Left            =   8520
         TabIndex        =   21
         Top             =   2040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Picture         =   "frmDes.frx":058A
         Caption         =   "Arquivar"
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
      Begin VB.CommandButton NovaREv 
         Caption         =   "Nova Revisão"
         Height          =   495
         Left            =   8640
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2880
         Width           =   1095
      End
      Begin VBFLXGRD18.VBFlexGrid GridAtual 
         Height          =   3135
         Left            =   240
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   480
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   5530
      End
      Begin VBFLXGRD18.VBFlexGrid GridArq 
         Height          =   2535
         Left            =   -74760
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   600
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   4471
      End
      Begin XPControls.XPButton novo 
         Height          =   435
         Left            =   8520
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmDes.frx":0B24
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
         Left            =   8520
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmDes.frx":10BE
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
         Left            =   8520
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmDes.frx":1658
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
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   24
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   2535
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   1
      Left            =   3240
      MaxLength       =   40
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   3855
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   3
      Left            =   3240
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Width           =   5535
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   420
      Left            =   9120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   741
      Picture         =   "frmDes.frx":1BF2
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
      Height          =   420
      Left            =   9120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   600
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   741
      Picture         =   "frmDes.frx":218C
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
   Begin XPControls.XPButton ESCCLI 
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   720
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "frmDes.frx":2726
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
   Begin XPControls.XPButton ESCCLI 
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Picture         =   "frmDes.frx":2CC0
      Caption         =   "LX Cli"
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
   Begin XPControls.XPButton ESCCLI 
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Picture         =   "frmDes.frx":325A
      Caption         =   "LX 2dig"
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
   Begin XPControls.XPButton Command7 
      Height          =   375
      Left            =   1320
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "frmDes.frx":37F4
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
      Caption         =   "Conjunto"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descrição"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cliente:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nome:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   5
      Left            =   3360
      TabIndex        =   4
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "frmDes"
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
Dim aTVAL As Variant
Dim aTCAM As Variant
Dim aTFOR As Variant
Dim aTPAD As Variant
Dim nCAMPOS As Integer
Dim iLOOP As Integer

Private Sub arquivar_Click()
  Dim sSQL As String
  GridAtual.Col = 0
  cDESENHO = GridAtual
  sSQL = "select * from atual WHERE conjunto='" & cCONJUNTO & "' AND DESENHO='" & cDESENHO & "'"  'atual->baixado
  aTVAL = PegSQL(cARQDES, sSQL, 11, aTCAM, aTFOR, aTPAD)
  If IncluiSQL(cARQDES, "select * from baixado WHERE conjunto='" & cCONJUNTO & "' AND DESENHO='" & cDESENHO & "'", 10, aTCAM, aTVAL, True, False) Then
    ApagaSQL cARQDES, sSQL
    Filgrid "atual", GridAtual
    Filgrid "baixado", GridArq
  End If
End Sub

Private Sub cmdClose_Click()
  On Error Resume Next
  If MDG("Gravar alteraçôes") Then
    For iLOOP = 0 To nCAMPOS - 1
      aVAL(iLOOP) = TXTFIELDS(iLOOP)
    Next iLOOP
    GrvSQL cARQDES, cSQLDES, nCAMPOS, aCAM, aVAL, aFOR, 1 '0 conjunto chave
  End If
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub Command7_Click()
  Dim cARQ As String
  Dim aRETU As Variant
  Dim sSQL As String
  Dim nNUMERO As Long
  nNUMERO = FixInt(TXTFIELDS(2), 0)
  cARQ = GeraConn(zMANA5EMP, "JETFOX")
  sSQL = "SELECT NOME FROM MA01 WHERE NUMERO=" & nNUMERO
  aRETU = PegSQL(cARQ, sSQL, 1, Array("NOME"), Array("C"), Array(""))
  If lRETU Then
    TXTFIELDS(3) = aRETU(0)
  End If
End Sub

Private Sub Editar_Click()
  GridAtual.Col = 0
  cDESENHO = GridAtual
  FrmDesI.Show vbModal, Me
  Filgrid "atual", GridAtual
End Sub

Private Sub Encerrar_Click()
  If Not MDG("Sair sem gravar") Then
    Exit Sub
  End If
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub ESCCLI_Click(Index As Integer)
  lRETU = False
  eRETU01 = ""
  eRETU02 = ""
  cARQESC = "MA01"
  Select Case Index
  Case 0
    cARQESC = "MA01"
  Case 1
    ePASS01 = "LOGIX"
  Case 2
    ePASS01 = "LOGIX"
    cARQESC = "2DIG"
  End Select

  escNUMNOM.Show vbModal, Me
  If lRETU Then
    TXTFIELDS(2) = eRETU01
    TXTFIELDS(3) = eRETU02
  End If
End Sub

Private Sub excluir_Click()
  Dim sSQL As String
  GridAtual.Col = 0
  cDESENHO = FixStr(GridAtual)
  sSQL = "select * from atual WHERE conjunto='" & cCONJUNTO & "' AND DESENHO='" & cDESENHO & "'"
  If ApagaSQLP(cARQDES, sSQL) Then
    Filgrid "atual", GridAtual
  End If
End Sub


Private Sub Filgrid(cARQ As String, oGRID As Variant)
  Dim cSQL As String
  cSQL = "SELECT * FROM " & cARQ & " WHERE CONJUNTO='" & cCONJUNTO & "' ORDER BY desenho "
  MontaGridFast oGRID, 4, Array(1600, 4000, 200, 1200), Array("Desenho", "Descricao", "Rev", "Data"), Array("L$DESENHO", "DESCRICAO", "REV", "C$DATA"), cARQDES, cSQL
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  cARQDES = PegPath("PATH", "DESENHO")
  cSQLDES = "select * from conjunto WHERE conjunto='" & cCONJUNTO & "'"
  nCAMPOS = 4
  aCAM = Array("CONJUNTO", "DESCRICAO", "CLIENTE", "CLINOME")
  aFOR = Array("C", "C", "NI", "C")
  aPAD = Array("", "", 0, "")
  aVAL = PegSQL(cARQDES, cSQLDES, nCAMPOS, aCAM, aFOR, aPAD)
  For iLOOP = 0 To nCAMPOS - 1
    TXTFIELDS(iLOOP) = aVAL(iLOOP)
  Next iLOOP
  Filgrid "atual", GridAtual
  Filgrid "baixado", GridArq
  aTCAM = Array("CONJUNTO", "DESENHO", "REV", "CDROM", "FITADAT", "DATA", "RECEBIDO", "DESCRICAO", "OBS", "REVC", "PADRONIZADO")
  aTFOR = Array("C", "C", "NI", "NI", "NI", "D", "D", "C", "C", "C", "B")
  aTPAD = Array(cCONJUNTO, cDESENHO, 0, 0, 0, Date, Date, "", "", "", True)
End Sub

Private Sub NovaRev_Click()
  Dim sSQL As String
  Dim nRTMP As Integer
  GridAtual.Col = 0
  cDESENHO = GridAtual
  sSQL = "select * from atual WHERE conjunto='" & cCONJUNTO & "' AND DESENHO='" & cDESENHO & "'"  ' atual->baixado
  aTVAL = PegSQL(cARQDES, sSQL, 11, aTCAM, aTFOR, aTPAD)
  If IncluiSQL(cARQDES, "select * from baixado WHERE conjunto='" & cCONJUNTO & "' AND DESENHO='" & cDESENHO & "'", 11, aTCAM, aTVAL, True, False) Then
    nRTMP = FixInt(aTVAL(2)) + 1
    GrvSQL cARQDES, sSQL, 1, Array("rev"), Array(nRTMP), Array("NI")
    Filgrid "atual", GridAtual
    Filgrid "baixado", GridArq
  End If
End Sub

Private Sub Novo_Click()
  Dim sSQL As String
  Dim cDESENHO As String
  cDESENHO = InputBox("Digite o Desenho", "Inclusão Desenho", "__")
  sSQL = "select * from atual WHERE conjunto='" & cCONJUNTO & "' AND DESENHO='" & cDESENHO & "'"
  If IncluiSQL(cARQDES, sSQL, 7, Array("CONJUNTO", "DESENHO", "REV", "CDROM", "FITADAT", "DATA", "RECEBIDO"), Array(cCONJUNTO, cDESENHO, 0, 0, 0, Date, Date), True, True) Then
    Filgrid "atual", GridAtual
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

