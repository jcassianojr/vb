VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{EA478B61-D9EC-47F6-BB21-95A533AF2251}#1.3#0"; "TabExt01.OCX"
Begin VB.Form frmME01 
   Caption         =   "Equipamento"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11280
   Icon            =   "frmME01.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   11280
   StartUpPosition =   2  'CenterScreen
   Begin NewTabCtl.NewTab SSTab1 
      Height          =   5895
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   10398
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
      TabHeight       =   520
      TabCaption(0)   =   "Preventiva"
      Tab(0).ControlCount=   26
      Tab(0).Control(0)=   "txtFields(11)"
      Tab(0).Control(1)=   "txtFields(10)"
      Tab(0).Control(2)=   "txtFields(9)"
      Tab(0).Control(3)=   "txtFields(8)"
      Tab(0).Control(4)=   "txtFields(7)"
      Tab(0).Control(5)=   "txtFields(6)"
      Tab(0).Control(6)=   "EscTipIns(1)"
      Tab(0).Control(7)=   "EscTipIns(0)"
      Tab(0).Control(8)=   "txtFields(5)"
      Tab(0).Control(9)=   "txtFields(4)"
      Tab(0).Control(10)=   "txtFields(0)"
      Tab(0).Control(11)=   "txtFields(1)"
      Tab(0).Control(12)=   "txtFields(2)"
      Tab(0).Control(13)=   "txtFields(3)"
      Tab(0).Control(14)=   "Label(13)"
      Tab(0).Control(15)=   "Label(12)"
      Tab(0).Control(16)=   "Label(11)"
      Tab(0).Control(17)=   "Label(10)"
      Tab(0).Control(18)=   "Label(9)"
      Tab(0).Control(19)=   "Label(8)"
      Tab(0).Control(20)=   "Label(7)"
      Tab(0).Control(21)=   "Label(6)"
      Tab(0).Control(22)=   "Label(2)"
      Tab(0).Control(23)=   "Label(3)"
      Tab(0).Control(24)=   "Label(4)"
      Tab(0).Control(25)=   "Label(5)"
      TabCaption(1)   =   "Femea"
      Tab(1).ControlCount=   14
      Tab(1).Control(0)=   "CmdOrdFem(1)"
      Tab(1).Control(1)=   "Command2(15)"
      Tab(1).Control(2)=   "EditSeq"
      Tab(1).Control(3)=   "CmdExcluir"
      Tab(1).Control(4)=   "CMdnovo"
      Tab(1).Control(5)=   "importar"
      Tab(1).Control(6)=   "CmdRenumerar"
      Tab(1).Control(7)=   "CmdClo(0)"
      Tab(1).Control(8)=   "CmdClo(1)"
      Tab(1).Control(9)=   "CmdREvi"
      Tab(1).Control(10)=   "CmdRevFemea"
      Tab(1).Control(11)=   "CmdOrdFem(0)"
      Tab(1).Control(12)=   "GridSeq"
      Tab(1).Control(13)=   "gridrevcau"
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   2280
         TabIndex        =   44
         Text            =   "0"
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   10
         Left            =   2280
         TabIndex        =   43
         Top             =   3480
         Width           =   3615
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   9
         Left            =   2280
         TabIndex        =   42
         Top             =   3120
         Width           =   3615
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   8
         Left            =   2280
         TabIndex        =   41
         Top             =   2760
         Width           =   3615
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   7
         Left            =   2280
         TabIndex        =   40
         Top             =   2400
         Width           =   3615
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   4800
         TabIndex        =   34
         Text            =   "0"
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton EscTipIns 
         Height          =   315
         Index           =   1
         Left            =   3120
         Picture         =   "frmME01.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton EscTipIns 
         Height          =   315
         Index           =   0
         Left            =   1800
         Picture         =   "frmME01.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1920
         Width           =   375
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   5
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1920
         Width           =   375
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   4
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton CmdOrdFem 
         Caption         =   "Ordenar Fx Seq,Ssq,Item"
         Height          =   255
         Index           =   1
         Left            =   -73200
         TabIndex        =   26
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Excluir Todos"
         Height          =   255
         Index           =   15
         Left            =   -65520
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton EditSeq 
         Caption         =   "Editar"
         Height          =   255
         Left            =   -65520
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton CmdExcluir 
         Caption         =   "Excluir"
         Height          =   255
         Left            =   -65520
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton CMdnovo 
         Caption         =   "Novo"
         Height          =   255
         Left            =   -65520
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton importar 
         Caption         =   "Importar"
         Height          =   255
         Left            =   -65520
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton CmdRenumerar 
         Caption         =   "Renumerar"
         Height          =   255
         Left            =   -65520
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton CmdClo 
         Caption         =   "Concluir Item"
         Height          =   255
         Index           =   0
         Left            =   -65520
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton CmdClo 
         Caption         =   "Concluir Todos"
         Height          =   495
         Index           =   1
         Left            =   -65520
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton CmdREvi 
         Caption         =   "Ver Revisoes"
         Height          =   255
         Left            =   -65520
         TabIndex        =   18
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton CmdRevFemea 
         Caption         =   "Nova Rev."
         Height          =   255
         Left            =   -65520
         TabIndex        =   17
         Top             =   4320
         Width           =   1095
      End
      Begin VB.CommandButton CmdOrdFem 
         Caption         =   "Ordenar Item"
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   14
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   9
         Text            =   "0"
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   3000
         TabIndex        =   8
         Text            =   "0"
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   3000
         TabIndex        =   7
         Text            =   "0"
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txtFields 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   3000
         TabIndex        =   6
         Text            =   "0"
         Top             =   1500
         Width           =   2055
      End
      Begin VBFLXGRD17.VBFlexGrid GridSeq 
         Height          =   3255
         Left            =   -74880
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   840
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   5741
         _Version        =   393216
      End
      Begin VBFLXGRD17.VBFlexGrid gridrevcau 
         Height          =   1515
         Left            =   -74880
         TabIndex        =   16
         Top             =   4080
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   2672
         _Version        =   393216
      End
      Begin VB.Label Label 
         Caption         =   "Ano"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   13
         Left            =   240
         TabIndex        =   39
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "Numero Fabricante/Serial"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   38
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label 
         Caption         =   "Modelo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   37
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "Fabricante"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   36
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "Contabil"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   35
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "Grupo Femea"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   8
         Left            =   3600
         TabIndex        =   33
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "Tipo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   7
         Left            =   2280
         TabIndex        =   30
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label 
         Caption         =   "Grupo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   29
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label 
         Caption         =   "Quantidade Manutenção Preventiva"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label 
         Caption         =   "Horas Manutenção Preventiva"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   12
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label 
         Caption         =   "Quantidade Vida Util"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   11
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label 
         Caption         =   "Horas Vida Util"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   10
         Top             =   1560
         Width           =   2655
      End
   End
   Begin VB.TextBox txtNome 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   4215
   End
   Begin VB.TextBox txtcodigo 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   2055
   End
   Begin XPControls.XPButton cmdClose 
      Height          =   435
      Left            =   9720
      TabIndex        =   45
      Top             =   120
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   767
      Picture         =   "frmME01.frx":109E
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
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   8400
      TabIndex        =   46
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   767
      Picture         =   "frmME01.frx":1638
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
   Begin XPControls.XPButton cmdFOTO 
      Height          =   435
      Left            =   6840
      TabIndex        =   47
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      Picture         =   "frmME01.frx":1BD2
      Caption         =   "Imagem"
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
      Caption         =   "Nome"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "Codigo"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmME01"
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
Dim nFEMMAQ As Integer
Private cORDFEMUSO As String
Private cARQFEMEA As String

Private Sub CmdClo_Click(Index As Integer)
  Dim dDATE As Date
  Dim DB As New ADODB.Connection
  Dim RSTAB As New ADODB.Recordset
  Dim sSQL As String



  dDATE = Date

  If GridSeq.Row = 0 Then
    Exit Sub
  End If

  DB.ConnectionTimeout = 120
  DB.Open GeracArq(cARQFEMEA)

  Select Case Index
  Case 0
    sSQL = "select * from FEMMAQ WHERE PF=" & nFEMMAQ & " ORDER BY ITEM"
  Case 1
    GridSeq.Col = 1
    nSEQ = GridSeq
    sSQL = "select * from FEMmaq WHERE PF=" & nFEMMAQ & " AND ITEM=" & nSEQ
  End Select



  RSTAB.Open sSQL, DB, adOpenDynamic, adLockOptimistic
  While Not RSTAB.EOF
    If RSTAB("ACAO") Then
      If IsDate(RSTAB("RESDAT")) Then
        dDATE = RSTAB("RESDAT")
      End If
      If IsDate(RSTAB("RESDAT2")) Then
        If dDATE < RSTAB("RESDAT2") Then
          dDATE = RSTAB("RESDAT2")
        End If
      End If
      If IsDate(RSTAB("RESDAT3")) Then
        If dDATE < RSTAB("RESDAT3") Then
          dDATE = RSTAB("RESDAT3")
        End If
      End If
      RSTAB("ACATOM") = "Concluida"
      RSTAB("ACADAT") = dDATE + 2
      RSTAB.Update
    End If
    RSTAB.MoveNext
  Wend
  RSTAB.Close
  DB.Close
  FilRelat

End Sub

Private Sub cmdClose_Click()
  On Error Resume Next
  If MDG("Gravar alteraçôes") Then
    For iLOOP = 0 To nCAMPOS - 1             ''tira o campos  nao gravar o numero,nome
      aVAL(iLOOP) = txtFields(iLOOP)       ''vb matriz zero
    Next iLOOP
    GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR, 1 '0 codigo chave
  End If
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub CmdExcluir_Click()
  Dim nSEQ As Integer
  Dim sSQL As String
  If GridSeq.Row = 0 Then
    Exit Sub
  End If
  GridSeq.Col = 1
  nSEQ = GridSeq
  sSQL = "select * from FEMMAQ WHERE PF=" & nFEMMAQ & " AND ITEM=" & nSEQ
  If ApagaSQLP(cARQFEMEA, sSQL) Then
    sSQL = "select * FEMrevi WHERE  TIPOAPU='M' AND PF=" & nFEMMAQ & " AND ITEM=" & nSEQ
    ApagaSQL cARQFEMEA, sSQL
  End If
  FilRelat
End Sub

Private Sub cmdFOTO_Click()
  Dim cSQL As String
  zgrp = txtcodigo.text
  iImage = 4
  cARQRTF = PegPath("PATH", "IMGME01")
  cSQL = "select * from IMAGENS WHERE CODIGO='" & zgrp & "'"
  IncluiSQL cARQRTF, cSQL, 1, Array("CODIGO"), Array(zgrp), True, False

  Load frmIMAGENS
  frmIMAGENS.txtFields(0).Enabled = False
  frmIMAGENS.Escolher(0).Visible = False
  frmIMAGENS.Show vbModal, Me

End Sub

Private Sub CmdNovo_Click()
  Dim nNUM
  nNUM = FixInt(PegMAXSQL(cARQFEMEA, "FEMMAQ WHERE PF=" & nFEMMAQ, "ITEM", 0)) + 1
  cSQL = "select * from FEMMAQ WHERE PF=" & nFEMMAQ & " AND ITEM=" & nNUM
  IncluiSQL cARQFEMEA, cSQL, 2, Array("PF", "ITEM"), Array(nFEMMAQ, nNUM), True, False
  FilRelat
End Sub

Private Sub CmdOrdFem_Click(Index As Integer)
  Select Case Index
  Case 0
    cORDFEMUSO = "ITEM"
  Case 1
    cORDFEMUSO = "FXSEQ,FXSSQ,FXITEM"
  End Select
  FilRelat
End Sub

Private Sub CmdRenumerar_Click()
  Dim nSEQ
  Dim nNUM As Integer
  Dim cSQL As String

  If GridSeq.Row = 0 Then
    Exit Sub
  End If

  nNUM = 0

  nNUM = FixInt(Val(Busca("Nº", "Renumerando", CStr(nNUM), 8)))
  If nNUM = 0 Then
    Alert ("Nº em Branco")
    Exit Sub
  End If
  GridSeq.Col = 1
  nSEQ = GridSeq
  cSQL = "select * from FEMmaq WHERE PF=" & nFEMMAQ & " AND ITEM=" & nSEQ
  GrvSQL cARQFEMEA, cSQL, 1, Array("ITEM"), Array(nNUM), Array("NI")
  FilRelat

End Sub

Private Sub CmdRevFemea_Click()
  Dim nREVI As Long
  Dim nSEQ As Long
  Dim cSQL As String
  Dim cARQ As String
  Dim aVAL As Variant
  Dim aCAM As Variant
  Dim aFOR As Variant
  Dim aPAD As Variant
  If GridSeq.Row = 0 Then
    Exit Sub
  End If
  GridSeq.Col = 1
  nSEQ = GridSeq
  cARQ = cARQFEMEA
  cSQL = "FEMrevi WHERE  TIPOAPU='M' AND PF=" & nFEMMAQ & " AND ITEM=" & nSEQ

  nREVI = FixInt(PegMAXSQL(cARQFEMEA, cSQL, "SEGGRA", 0))
  nREVI = nREVI + 1


  aCAM = Array("PF", "ITEM", "PROCESSO", "FALTIP", "FALEFE", _
               "FALCAU", "CRTATU", "INDOCO", "INDSEV", "INDDET", _
               "INDRIS", "ACAREC", "RESCOD", "RESNOM", "RESCOD2", _
               "ACATOM", "RESCOD3", "RINDOCO", "RINDSER", "RINDDET", _
               "RINDRIS", "RESNOM2", "RESNOM3", "SITUACAO", "SIGI", _
               "CARAPREV", "RESDAT", "RESDAT2", "RESDAT3", "ACADAT", _
               "ACAO", "EXCRPN", "ALTMAN", "MUDPAD", "PSA", "FXSEQ", "FXSSQ", "FXITEM", _
               "PRONUM", "EFENUM", "FALNUM", "CAUNUM", "SEGGRA", "TIPOAPU")
  aFOR = Array("NI", "NI", "C", "C", "C", _
               "C", "C", "NI", "NI", "NI", _
               "NI", "C", "NI", "C", "NI", _
               "C", "NI", "NI", "NI", "NI", _
               "NI", "C", "C", "C", "C", "C", _
               "DN", "DN", "DN", "DN", "BN", _
               "BN", "BN", "BN", "C", "NI", "NI", "NI", "NI", "NI", "NI", "NI", "NI", "C")
  aPAD = Array(0, 0, "", "", "", _
               "", "", 0, 0, 0, _
               0, "", 0, "", 0, _
               "", 0, 0, 0, 0, 0, _
               "", "", "", "", "", _
               "", "", "", "", False, _
               False, False, False, "", 0, 0, 0, 0, 0, 0, 0, 0, " ")


  cSQL = "select * from FEMMAQ WHERE PF=" & nFEMMAQ & " AND ITEM=" & nSEQ

  aVAL = PegSQL(cARQ, cSQL, 44, aCAM, aFOR, aPAD)

  'checa se foi feita a revisao
  If aVAL(17) = 0 Or aVAL(18) = 0 Or aVAL(19) = 0 Or aVAL(20) = 0 Then
    Alert ("Novos Indices Nao Preenchidos")
    Exit Sub
  End If

  'atualiza revisao e tipo
  aVAL(42) = nREVI
  aVAL(43) = "M"

  cSQL = "select * from FEMMAQ WHERE  TIPOAPU='M' AND PF=" & nFEMMAQ & " AND ITEM=" & nSEQ & " AND SEGGRA=" & nREVI

  IncluiSQL cARQFEMEA, cSQL, 44, aCAM, aVAL, True, False

  cSQL = "select * from FEMMAQ WHERE PF=" & nFEMMAQ & " AND ITEM=" & nSEQ

  GrvSQL cARQ, cSQL, 17, Array("ACAREC", "RESCOD", "RESNOM", "RESCOD2", "ACATOM", _
                               "RESCOD3", "RINDOCO", "RINDSER", "RINDDET", "RINDRIS", _
                               "RESNOM2", "RESNOM3", "INDOCO", "INDSEV", "INDDET", _
                               "INDRIS", "ACAO"), _
                               Array("", 0, "", 0, "", 0, 0, 0, 0, 0, "", "", aVAL(17), aVAL(18), aVAL(19), aVAL(20), False), _
                               Array("C", "NI", "C", "NI", "C", "NI", "NI", "NI", "NI", "NI", "C", "C", "NI", "NI", "NI", "NI", "")

  FILGRIDrevCAU
End Sub

Private Sub CmdREvi_Click()
  Dim nREVI As Long
  Dim nSEQ As Long

  If gridrevcau.Row = 0 Then
    Exit Sub
  End If
  gridrevcau.Col = 0
  nSEQ = gridrevcau
  gridrevcau.Col = 1
  nREVI = gridrevcau
  ePASS01 = "select * from FEMrevi WHERE  TIPOAPU='M' AND PF=" & nFEMMAQ & " AND ITEM=" & nSEQ & " AND SEGGRA=" & nREVI
  ePASS02 = cARQFEMEA

  Load frmFEMEI
  frmFEMEI.txtFields(2).Enabled = False
  frmFEMEI.txtFields(3).Enabled = False
  frmFEMEI.txtFields(4).Enabled = False
  frmFEMEI.txtFields(5).Enabled = False
  frmFEMEI.cmdClose.Enabled = False
  frmFEMEI.CmdAltman.Enabled = False
  frmFEMEI.Show vbModal, Me


End Sub

Private Sub Command2_Click(Index As Integer)
  Dim sSQL As String

  sSQL = "select * from FEMMAQ WHERE PF=" & nFEMMAQ
  If ApagaSQLP(cARQFEMEA, sSQL) Then
    sSQL = "select * FEMrevi WHERE  TIPOAPU='M' AND PF=" & nFEMMAQ
    ApagaSQL cARQFEMEA, sSQL
  End If
  FilRelat
End Sub

Private Sub EditSeq_Click()
  Dim nSEQ As Integer
  If GridSeq.Row = 0 Then
    Exit Sub
  End If
  GridSeq.Col = 1
  nSEQ = GridSeq
  ePASS01 = "select * from FEMMAQ WHERE PF=" & nFEMMAQ & " AND ITEM=" & nSEQ
  ePASS02 = cARQFEMEA
  frmFEMEI.Show vbModal, Me
  FilRelat
End Sub

Private Sub Encerrar_Click()
  If Not MDG("Sair sem gravar") Then
    Exit Sub
  End If
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub EscTipIns_Click(Index As Integer)
  Select Case Index
  Case 0
    iMD02 = 8
  Case 1
    iMD02 = 9
  End Select
  escMD02.Show vbModal, Me
  If lRETU Then
    Select Case Index
    Case 0
      txtFields(4).text = eRETU01
      txtFields(6).text = eRETU03
      '                txt(5).Text = eRETU02
    Case 1
      txtFields(5).text = eRETU01
    End Select
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub FILGRIDrevCAU()
  Dim cSQL As String
  cSQL = "select ITEM,SEGGRA,PROCESSO,FXSEQ,FXSSQ,FXITEM from FEMREVI WHERE  TIPOAPU='M' AND PF=" & nFEMMAQ & " ORDER BY ITEM,SEGGRA"
  DizerBarra "carregando femea maquina revisoes"
  MontaGridFast gridrevcau, 6, Array(400, 400, 400, 400, 400, 7000), Array("ITEM", "REV", "SEQ", "SSQ", "IT", "Descriçao"), Array("ITEM", "SEGGRA", "FXSEQ", "FXSSQ", "FXITEM", "PROCESSO"), cARQFEMEA, cSQL
  DizerBarra ""
End Sub

Private Sub Form_Load()
  Dim nTMPNUMERO
  CenterFormToScreen Me
  cARQ = GeraConn(zMANA5EMP, "SDECDX")
  cSQL = "select NUMERO,QTDEBASE,HRBAS,VDBAS,VDHBAS,GRUPO,TIPO,FEMEA,CONTABIL,FABRICANTE,MODELO,NUMFAB,ANO "
  cSQL = cSQL & " FROM ME01 WHERE NUMERO='" & Trim(ePASS01) & "'"
  cARQFEMEA = PegPath("PATH", "FEMEA")

  txtcodigo.text = ePASS01
  txtNome.text = ePASS02
  nCAMPOS = 12
  aCAM = Array("QTDEBASE", "HRBAS", "VDBAS", "VDHBAS", "GRUPO", "TIPO", "FEMEA", "CONTABIL", "FABRICANTE", "MODELO", "NUMFAB", "ANO")
  aFOR = Array("N", "N", "N", "N", "C", "C", "N", "C", "C", "C", "C", "N")
  aPAD = Array(0, 0, 0, 0, "", "", 0, "", "", "", "", 0)
  aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
  For iLOOP = 0 To nCAMPOS - 1
    txtFields(iLOOP) = aVAL(iLOOP)
  Next iLOOP

  cORDFEMUSO = "ITEM"

  nFEMMAQ = Val(txtFields(6))


  FilRelat

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub importar_Click()
  cARQIMP = "FEMMAQ"
  FrmImp.Show vbModal, Me
  FilRelat
End Sub
Private Sub txtFields_GotFocus(Index As Integer)
  FocusMe
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index <= 3 Or Index = 6 Then
    KeyAscii = ValiText(KeyAscii, "#NI")
  End If
End Sub

Private Sub FilRelat()
  Dim cSQL As String
  cSQL = "select PF,ITEM,PROCESSO,FXSEQ,FXSSQ,FXITEM from FEMMAQ WHERE PF=" & nFEMMAQ & " ORDER BY " & cORDFEMUSO
  DizerBarra "carregando femea adicional"
  MontaGridFast GridSeq, 6, Array(400, 400, 400, 400, 400, 7000), Array("PF", "ITEM", "SEQ", "SSQ", "IT", "Descriçao"), Array("PF", "ITEM", "FXSEQ", "FXSSQ", "FXITEM", "PROCESSO"), cARQFEMEA, cSQL
  DizerBarra ""
End Sub

Private Sub TXTFIELDS_LostFocus(Index As Integer)
  If Index = 6 Then
    nFEMMAQ = Val(txtFields(6))
  End If
End Sub


