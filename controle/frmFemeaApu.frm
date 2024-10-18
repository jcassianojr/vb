VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.1#0"; "vbccr18.ocx"
Object = "{66E63055-5A66-4C79-9327-4BC077858695}#9.0#0"; "newtab01.OCX"
Object = "{075212A8-C1CF-444E-939D-F6046CCDBC08}#1.0#0"; "VBFLXGRD18.OCX"
Begin VB.Form frmFemeaApu 
   Caption         =   "Apuracao Femea"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9615
   Icon            =   "frmFemeaApu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin VBCCR18.DTPicker DTPicker2 
      Height          =   375
      Left            =   4200
      TabIndex        =   47
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Value           =   44285
      AllowUserInput  =   -1  'True
   End
   Begin VBCCR18.DTPicker DTPicker1 
      Height          =   375
      Left            =   2520
      TabIndex        =   46
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Value           =   44285
      AllowUserInput  =   -1  'True
   End
   Begin VBCCR18.ProgressBar barra 
      Height          =   255
      Left            =   5760
      Top             =   960
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   450
   End
   Begin VB.CommandButton CmdApurar 
      Caption         =   "Checagem "
      Height          =   315
      Index           =   1
      Left            =   8400
      TabIndex        =   37
      Top             =   960
      Width           =   915
   End
   Begin VB.TextBox DIZAPU 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   480
      Width           =   1215
   End
   Begin NewTabCtl.NewTab SSTab1 
      Height          =   5055
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   8916
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
      ControlVersion  =   9
      TabCaption(0)   =   "SAC RDP FAIXAS"
      Tab(0).ControlCount=   22
      Tab(0).Control(0)=   "text(4)"
      Tab(0).Control(1)=   "text(5)"
      Tab(0).Control(2)=   "text(6)"
      Tab(0).Control(3)=   "text(7)"
      Tab(0).Control(4)=   "text(8)"
      Tab(0).Control(5)=   "text(9)"
      Tab(0).Control(6)=   "text(10)"
      Tab(0).Control(7)=   "text(11)"
      Tab(0).Control(8)=   "CmdApagaTudo"
      Tab(0).Control(9)=   "CmdCheckcli"
      Tab(0).Control(10)=   "Grid"
      Tab(0).Control(11)=   "EditSeq"
      Tab(0).Control(12)=   "Command2"
      Tab(0).Control(13)=   "Command1"
      Tab(0).Control(14)=   "lbl(8)"
      Tab(0).Control(15)=   "lbl(9)"
      Tab(0).Control(16)=   "lbl(10)"
      Tab(0).Control(17)=   "lbl(11)"
      Tab(0).Control(18)=   "lbl(12)"
      Tab(0).Control(19)=   "lbl(13)"
      Tab(0).Control(20)=   "lbl(14)"
      Tab(0).Control(21)=   "lbl(15)"
      TabCaption(1)   =   "10 Maiores"
      Tab(1).ControlCount=   2
      Tab(1).Control(0)=   "Cmdeditar"
      Tab(1).Control(1)=   "Grid10"
      TabCaption(2)   =   "Resumo Operacoes"
      Tab(2).ControlCount=   2
      Tab(2).Control(0)=   "Command3"
      Tab(2).Control(1)=   "GridOpe"
      Begin VB.TextBox text 
         Alignment       =   1  'Right Justify
         Height          =   330
         Index           =   4
         Left            =   240
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   840
         Width           =   750
      End
      Begin VB.TextBox text 
         Alignment       =   1  'Right Justify
         Height          =   330
         Index           =   5
         Left            =   1080
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   855
         Width           =   750
      End
      Begin VB.TextBox text 
         Alignment       =   1  'Right Justify
         Height          =   330
         Index           =   6
         Left            =   1920
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   855
         Width           =   750
      End
      Begin VB.TextBox text 
         Alignment       =   1  'Right Justify
         Height          =   330
         Index           =   7
         Left            =   2760
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   855
         Width           =   750
      End
      Begin VB.TextBox text 
         Alignment       =   1  'Right Justify
         Height          =   330
         Index           =   8
         Left            =   3600
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   855
         Width           =   750
      End
      Begin VB.TextBox text 
         Alignment       =   1  'Right Justify
         Height          =   330
         Index           =   9
         Left            =   4440
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   840
         Width           =   750
      End
      Begin VB.TextBox text 
         Alignment       =   1  'Right Justify
         Height          =   330
         Index           =   10
         Left            =   5280
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   855
         Width           =   750
      End
      Begin VB.TextBox text 
         Alignment       =   1  'Right Justify
         Height          =   330
         Index           =   11
         Left            =   6120
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   855
         Width           =   750
      End
      Begin VB.CommandButton CmdApagaTudo 
         Caption         =   "Excluir Tudo"
         Height          =   435
         Left            =   7500
         TabIndex        =   16
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton CmdCheckcli 
         Caption         =   "Checar Nome Clientes"
         Height          =   495
         Left            =   7500
         TabIndex        =   15
         Top             =   3180
         Width           =   1095
      End
      Begin VBFLXGRD18.VBFlexGrid Grid 
         Height          =   2535
         Left            =   240
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1320
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   4471
      End
      Begin VBFLXGRD18.VBFlexGrid Grid10 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   480
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   7646
      End
      Begin VBFLXGRD18.VBFlexGrid GridOpe 
         Height          =   4335
         Left            =   -74760
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   480
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   7646
      End
      Begin XPControls.XPButton Cmdeditar 
         Height          =   435
         Left            =   -67200
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   480
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   767
         Picture         =   "frmFemeaApu.frx":058A
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
      Begin XPControls.XPButton Command3 
         Height          =   435
         Left            =   -67200
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   480
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   767
         Picture         =   "frmFemeaApu.frx":0B24
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
      Begin XPControls.XPButton EditSeq 
         Height          =   435
         Left            =   7440
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   767
         Picture         =   "frmFemeaApu.frx":10BE
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
      Begin XPControls.XPButton Command2 
         Height          =   435
         Left            =   7440
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   767
         Picture         =   "frmFemeaApu.frx":1658
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
      Begin XPControls.XPButton Command1 
         Height          =   435
         Left            =   7440
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   767
         Picture         =   "frmFemeaApu.frx":1BF2
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
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0-20"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   33
         Top             =   615
         Width           =   315
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "21-40"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   9
         Left            =   1320
         TabIndex        =   32
         Top             =   615
         Width           =   405
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "41-60"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   10
         Left            =   2040
         TabIndex        =   31
         Top             =   615
         Width           =   405
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "61-80"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   11
         Left            =   2760
         TabIndex        =   30
         Top             =   600
         Width           =   405
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "81-100"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   12
         Left            =   3600
         TabIndex        =   29
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "101-120"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   13
         Left            =   4440
         TabIndex        =   28
         Top             =   615
         Width           =   585
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "121-140"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   14
         Left            =   5280
         TabIndex        =   27
         Top             =   600
         Width           =   585
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ">140"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   15
         Left            =   6120
         TabIndex        =   26
         Top             =   600
         Width           =   360
      End
   End
   Begin VB.CommandButton CmdApurar 
      Caption         =   "Apurar"
      Height          =   315
      Index           =   0
      Left            =   5760
      TabIndex        =   13
      Top             =   480
      Width           =   915
   End
   Begin VB.CheckBox Semestral 
      Caption         =   "Semestral"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   4560
      TabIndex        =   12
      Top             =   1020
      Width           =   1035
   End
   Begin VB.CheckBox Anual 
      Caption         =   "Anual"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   4560
      TabIndex        =   11
      Top             =   660
      Width           =   915
   End
   Begin VB.CheckBox apurado 
      Caption         =   "Apurado"
      Enabled         =   0   'False
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   5760
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   915
   End
   Begin VB.TextBox text 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   330
      Index           =   12
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   1230
   End
   Begin VB.TextBox text 
      Height          =   330
      Index           =   3
      Left            =   2340
      TabIndex        =   7
      Top             =   900
      Width           =   2010
   End
   Begin VB.TextBox text 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   330
      Index           =   2
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "2006"
      Top             =   600
      Width           =   570
   End
   Begin VB.TextBox text 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   330
      Index           =   1
      Left            =   8520
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "1"
      Top             =   600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.TextBox text 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   975
   End
   Begin XPControls.XPButton cmdClose 
      Height          =   435
      Left            =   8160
      TabIndex        =   38
      Top             =   120
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   767
      Picture         =   "frmFemeaApu.frx":218C
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
   Begin VBCCR18.SpinBox mes 
      Height          =   495
      Left            =   120
      TabIndex        =   39
      Top             =   840
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   1
      Max             =   12
      Value           =   1
   End
   Begin VBCCR18.SpinBox ano 
      Height          =   495
      Left            =   1080
      TabIndex        =   40
      Top             =   840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   1950
      Max             =   2200
      Value           =   1950
      ThousandsSeparator=   0   'False
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   5
      Left            =   2400
      TabIndex        =   8
      Top             =   660
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ano:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   4
      Left            =   1200
      TabIndex        =   6
      Top             =   600
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mes:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   4
      Top             =   615
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   1800
      TabIndex        =   2
      Top             =   195
      Width           =   540
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seq:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   450
   End
End
Attribute VB_Name = "frmFemeaApu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Const nFORMID = 1016te
Option Explicit
Const cFORMID = "Apurar Femea"
Private aCAM
Private aFOR
Private aVAL
Private aPAD As Variant
Private cARQPF
Private cARQFEMEA
Private cSQLUSO As String
Private nITEM
Private nCAMPOS As Long
Private nBARPOS
Private nROWREC As Long

Private Sub CmdApagaTudo_Click()
  Dim cSQL As String
  cSQL = "SELECT * FROM FEMEARPNI WHERE SEQ=" & FixInt(aVAL(0))
  If ApagaSQLP(cARQFEMEA, cSQL, "Apagar Todos Lançamentos") Then
    Filgrid
  End If
End Sub

Private Sub CmdApurar_Click(Index As Integer)
  Dim nFX01 As Integer
  Dim nFX02 As Integer
  Dim nFX03 As Integer
  Dim nFX04 As Integer
  Dim nFX05 As Integer
  Dim nFX06 As Integer
  Dim nFX07 As Integer
  'Dim nFX08 As Long
  Dim nINDICE As Integer
  Dim nSEQ As Integer
  Dim nMES As Integer
  Dim nANO As Integer
  Dim nSAC As Integer
  Dim nMSRD As Integer
  Dim nCLIENTE As Integer
  Dim nINDRIS As Long
  Dim nMESANT As Integer
  Dim nANOANT As Integer
  Dim nMESANT1 As Integer
  Dim nANOANT1 As Integer
  Dim nMESANT2  As Integer
  Dim nANOANT2 As Integer
  Dim nANOANT3  As Integer
  Dim nMESANT3 As Long
  Dim nPF  As Integer
  Dim nTOTRPN  As Integer
  Dim nTOTCAU  As Integer
  Dim nTOTM40  As Integer
  Dim nMAIRPN  As Integer
  Dim nPRONUM As Integer
  Dim x  As Integer
  Dim y  As Integer
  Dim nBARPOS As Long
  Dim oDBPF As ADODB.Connection
  Dim oDBFEMEA As ADODB.Connection
  Dim oRSPF As ADODB.Recordset
  Dim oRSFEMEA As ADODB.Recordset
  Dim oRSFEME2 As ADODB.Recordset
  Dim oCONSAC As ADODB.Connection
  Dim oRSSAC As ADODB.Recordset
  Dim oCONMSRD As ADODB.Connection
  Dim oRSMSRD As ADODB.Recordset
  Dim Orscom As ADODB.Command
  Dim cSQLMSRD As String
  Dim cSQLSAC, cSQL2 As String
  Dim cARQ As String
  Dim cSQL As String
  Dim cTabela As String
  Dim cCODIGO As String
'  Dim lPULA As Variant
  Dim dINI As Date
  Dim dFIM As Date
  Dim lSAC As Boolean
  Dim lRDP As Boolean
  Dim aGRVAL As Variant
  Dim cTIPO As String
  'On Error GoTo errhandler

  Screen.MousePointer = vbHourglass
  lSAC = False
  lRDP = False
  dINI = DTPicker1.Value
  dFIM = DTPicker2.Value
  nSEQ = FixInt(tEXT(0))
  nMES = FixInt(mes.Value)
  nANO = FixInt(ano.Value)

  If nMES = 0 Or nANO = 0 Then
    Alert "Mes ou Ano não Preenchido"
    Exit Sub
  End If

  If Index = 1 Then
    '   Set oDBFEMEA = New ADODB.Connection
    '   Set oRSFEMEA = New ADODB.Recordset
    '   oDBFEMEA.Open GeracArq(cARQFEMEA)
    '   For X = 1 To 42
    '        DizerBarra X
    '        Y = 1
    '        cSQL = "SELECT TOP 10 * FROM FEMRPNT WHERE SEQ=" & X & " ORDER BY INDRIS DESC"
    '        oRSFEMEA.Open cSQL, oDBFEMEA, adOpenDynamic, adLockOptimistic
    '        While Not oRSFEMEA.EOF
    '            oRSFEMEA("TIPOAPU") = "P"
    '            oRSFEMEA("SEGGRA") = Y
    '            oRSFEMEA.Update
    '            oRSFEMEA.MoveNext
    '            Y = Y + 1
    '        Wend
    '        oRSFEMEA.Close
    '   Next X
    '   oDBFEMEA.Close
    '   Set oRSFEMEA = Nothing
    '   Set oDBFEMEA = Nothing
    Set oDBPF = New ADODB.Connection
    Set oDBFEMEA = New ADODB.Connection
    Set oRSFEMEA = New ADODB.Recordset
    Set oRSFEME2 = New ADODB.Recordset
    Set oRSPF = New ADODB.Recordset
    Set Orscom = New ADODB.Command
    oDBFEMEA.ConnectionTimeout = 120
    oDBFEMEA.Open GeracArq(cARQFEMEA)
    'cSQL = "SELECT DISTINCT PF,PRONUM FROM FEMEA ORDER BY PF,PRONUM"
    cSQL = "SELECT DISTINCT PF,PRONUM FROM FEMEAPAD ORDER BY PF,PRONUM"
    DizerBarra "Checando PF/Processo"
    oRSFEMEA.Open cSQL, oDBFEMEA, adOpenForwardOnly, adLockReadOnly
    While Not oRSFEMEA.EOF
      nPF = oRSFEMEA("PF")
      nPRONUM = FixNum(oRSFEMEA("PRONUM"))
      DIZAPU.tEXT = nPF & "/" & nPRONUM
      DIZAPU.Refresh
      cSQL = "SELECT PF,PRONUM FROM FEMPF WHERE PF=" & nPF & " and pronum=" & nPRONUM
      oRSFEME2.Open cSQL, oDBFEMEA, adOpenForwardOnly, adLockReadOnly
      oRSFEMEA.MoveNext                    'sai do registro evitar travamento
      If oRSFEME2.EOF Then                 'nao achou
        'cSQL = "DELETE FROM FEMEA WHERE PF=" & nPF & " AND PRONUM=" & nPRONUM
        cSQL = "DELETE FROM FEMEAPAD WHERE PF=" & nPF & " AND PRONUM=" & nPRONUM
        Orscom.CommandText = cSQL
        Orscom.ActiveConnection = oDBFEMEA
        Orscom.Execute
        'cSQL = "SELECT DISTINCT PF,PRONUM FROM FEMEA ORDER BY PF,PRONUM"
        cSQL = "SELECT DISTINCT PF,PRONUM FROM FEMEAPAD ORDER BY PF,PRONUM"
      End If
      oRSFEME2.Close
    Wend
    oRSFEMEA.Close
    DizerBarra "Apagando Bloqueados"
    oDBPF.ConnectionTimeout = 120
    oDBPF.Open GeracArq(cARQPF)
    cSQL = "SELECT PF FROM PF WHERE BLOQUEADO"
    oRSPF.Open cSQL, oDBPF
    While Not oRSPF.EOF
      nPF = oRSPF("PF")
      DIZAPU.tEXT = nPF
      DIZAPU.Refresh
      'cSQL = "DELETE FROM FEMEA WHERE PF=" & nPF
      cSQL = "DELETE FROM FEMEAPAD WHERE PF=" & nPF
      Orscom.CommandText = cSQL
      Orscom.ActiveConnection = oDBFEMEA
      Orscom.Execute

      cSQL = "DELETE FROM FEMRPNT WHERE PF=" & nPF
      Orscom.CommandText = cSQL
      Orscom.ActiveConnection = oDBFEMEA
      Orscom.Execute

      cSQL = "DELETE FROM FEMRPNO WHERE PF=" & nPF
      Orscom.CommandText = cSQL
      Orscom.ActiveConnection = oDBFEMEA
      Orscom.Execute

      cSQL = "DELETE FROM FEMPF WHERE PF=" & nPF
      Orscom.CommandText = cSQL
      Orscom.ActiveConnection = oDBFEMEA
      Orscom.Execute

      cSQL = "DELETE FROM FEMPFHIS WHERE PF=" & nPF
      Orscom.CommandText = cSQL
      Orscom.ActiveConnection = oDBFEMEA
      Orscom.Execute


      oRSPF.MoveNext
    Wend
    DizerBarra "Apagando sem numero PF"
    'cSQL = "DELETE FROM FEMEA WHERE (PF=0 OR PF IS NULL)" 'sem numero pf
    cSQL = "DELETE FROM FEMEAPAD WHERE (PF=0 OR PF IS NULL)"  'sem numero pf
    Orscom.CommandText = cSQL
    Orscom.ActiveConnection = oDBFEMEA
    Orscom.Execute

    DizerBarra "Apagando sem avulso sem item"
    cSQL = "delete from femavu where item is null or item=0"
    Orscom.CommandText = cSQL
    Orscom.ActiveConnection = oDBFEMEA
    Orscom.Execute


    DizerBarra "Apagando sem numero PF-RNPT"
    cSQL = "DELETE FROM FEMRPNT WHERE (PF=0  OR PF IS NULL) AND TIPOAPU<>'P'"  'sem numero pf
    Orscom.CommandText = cSQL
    Orscom.ActiveConnection = oDBFEMEA
    Orscom.Execute

    DizerBarra "Apagando sem numero PF=RPNO"
    cSQL = "DELETE FROM FEMRPNO WHERE (PF=0 OR PF IS NULL)"  'sem numero pf
    Orscom.CommandText = cSQL
    Orscom.ActiveConnection = oDBFEMEA
    Orscom.Execute


    oDBFEMEA.Close
    oDBPF.Close
    DizerBarra ""
    Set oRSFEMEA = Nothing
    Set oRSFEME2 = Nothing
    Set oDBFEMEA = Nothing
    Set oDBPF = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
  End If


  'rpn faixas
  If MDG("Apurar Femea RPN-Faixas") Then
    '        cSQL = "SELECT PF,EXCRPN FROM PF WHERE PF=" & nPF checar excprn/bloqueados melhorar selects abaixo
    tEXT(4) = 0
    tEXT(5) = 0
    tEXT(6) = 0
    tEXT(7) = 0
    tEXT(8) = 0
    tEXT(9) = 0
    tEXT(10) = 0
    tEXT(11) = 0

    Set oDBFEMEA = New ADODB.Connection
    Set oRSFEMEA = New ADODB.Recordset
    For x = 1 To 3
      cTabela = "FEMEA"
      oDBFEMEA.ConnectionTimeout = 120
      If x = 1 Or x = 3 Then
        oDBFEMEA.Open GeracArq(cARQFEMEA)
      Else
        oDBFEMEA.Open GeracArq(cARQPF)
      End If
      If x = 1 Then
        cTabela = "FEMEAPAD"             '1 Femea Padrao
      End If
      If x = 2 Then
        cTabela = "FEMEA"                '2 femea normal
      End If
      If x = 3 Then
        cTabela = "FEMAVU"               '3 Femea Avulso
      End If
      DizerBarra "carregando femea,indices <21"
      cSQL = "SELECT count(indris) FROM " & cTabela & " where indris<21 AND NOT BLOQUEADO"
      oRSFEMEA.Open cSQL, oDBFEMEA, adOpenForwardOnly, adLockReadOnly
      tEXT(4) = tEXT(4) + FixInt(oRSFEMEA(0))
      oRSFEMEA.Close
      DizerBarra "carregando femea,indices >21 <41"
      cSQL = "SELECT count(indris) FROM " & cTabela & " where indris>21 AND indris<41 AND NOT BLOQUEADO"
      oRSFEMEA.Open cSQL, oDBFEMEA, adOpenForwardOnly, adLockReadOnly
      tEXT(5) = tEXT(5) + FixInt(oRSFEMEA(0))
      oRSFEMEA.Close
      DizerBarra "carregando femea,indices >40 <61"
      cSQL = "SELECT count(indris) FROM " & cTabela & " where indris>40 AND indris<61 AND NOT BLOQUEADO"
      oRSFEMEA.Open cSQL, oDBFEMEA, adOpenForwardOnly, adLockReadOnly
      tEXT(6) = tEXT(6) + FixInt(oRSFEMEA(0))
      oRSFEMEA.Close
      DizerBarra "carregando femea,indices >60 <81"
      cSQL = "SELECT count(indris) FROM " & cTabela & " where indris>60 AND indris<81 AND NOT BLOQUEADO"
      oRSFEMEA.Open cSQL, oDBFEMEA, adOpenForwardOnly, adLockReadOnly
      tEXT(7) = tEXT(7) + FixInt(oRSFEMEA(0))
      oRSFEMEA.Close
      DizerBarra "carregando femea,indices >80 <101"
      cSQL = "SELECT count(indris) FROM " & cTabela & " where indris>80 AND indris<101 AND NOT BLOQUEADO"
      oRSFEMEA.Open cSQL, oDBFEMEA, adOpenForwardOnly, adLockReadOnly
      tEXT(8) = tEXT(8) + FixInt(oRSFEMEA(0))
      oRSFEMEA.Close
      DizerBarra "carregando femea,indices >100 <121"
      cSQL = "SELECT count(indris) FROM " & cTabela & " where indris>100 AND indris<121  AND NOT BLOQUEADO"
      oRSFEMEA.Open cSQL, oDBFEMEA, adOpenForwardOnly, adLockReadOnly
      tEXT(9) = tEXT(9) + FixInt(oRSFEMEA(0))
      oRSFEMEA.Close
      DizerBarra "carregando femea,indices >120 <141"
      cSQL = "SELECT count(indris) FROM " & cTabela & " where indris>120 AND indris<141  AND NOT BLOQUEADO"
      oRSFEMEA.Open cSQL, oDBFEMEA, adOpenForwardOnly, adLockReadOnly
      tEXT(10) = tEXT(10) + FixInt(oRSFEMEA(0))
      oRSFEMEA.Close
      DizerBarra "carregando femea,indices >140"
      cSQL = "SELECT count(indris) FROM " & cTabela & " where indris>140  AND NOT BLOQUEADO"
      oRSFEMEA.Open cSQL, oDBFEMEA, adOpenForwardOnly, adLockReadOnly
      tEXT(11) = tEXT(11) + FixInt(oRSFEMEA(0))
      oRSFEMEA.Close
      oDBFEMEA.Close
    Next x

    Set oRSFEMEA = Nothing
    Set oDBFEMEA = Nothing
  End If
  'fim rpn faixas

  If MDG("Apura Femea Padrao- 10 Maiores") Then
    cSQL = "DELETE FROM FEMRPNT WHERE SEQ=" & nSEQ
    ADOComando GeracArq(cARQFEMEA), cSQL
    Set oDBFEMEA = New ADODB.Connection
    Set oDBPF = New ADODB.Connection
    Set oRSFEMEA = New ADODB.Recordset
    Set oRSFEME2 = New ADODB.Recordset
    Set oRSPF = New ADODB.Recordset
    oDBPF.ConnectionTimeout = 120
    oDBPF.Open GeracArq(cARQFEMEA)
    For x = 1 To 3
      y = 0
      oDBFEMEA.ConnectionTimeout = 120
      If x = 1 Or x = 3 Then
        oDBFEMEA.Open GeracArq(cARQFEMEA)
      Else
        oDBFEMEA.Open GeracArq(cARQPF)
      End If
      If x = 1 Then                        '1 Femea Padrao
        cTIPO = "P"
        cTabela = "FEMCAU"
      End If
      If x = 2 Then                        '2 femea normal
        cTIPO = "N"
        cTabela = "FEMEA"
      End If
      If x = 3 Then                        ' 3 Femea Avulso/complemtar
        cTabela = "FEMAVU"
        cTIPO = "C"
      End If
      DizerBarra "Apurando 10 Mais " & cTabela
      cSQL = "SELECT TOP 10 * FROM " & cTabela & " WHERE NOT BLOQUEADO ORDER BY INDRIS DESC "
      oRSFEMEA.Open cSQL, oDBFEMEA, adOpenForwardOnly, adLockReadOnly
      oRSFEME2.Open "FEMRPNT", oDBPF, adOpenDynamic, adLockOptimistic
      While Not oRSFEMEA.EOF
        If y < 11 Then
          oRSFEME2.AddNew
          oRSFEME2("SEQ") = nSEQ
          oRSFEME2("ANO") = nANO
          oRSFEME2("MES") = nMES
          oRSFEME2("TIPOAPU") = cTIPO
          oRSFEME2("PF") = oRSFEMEA("PF")
          oRSFEME2("SEGGRA") = FixInt(y)
          oRSFEME2("INDRIS") = oRSFEMEA("INDRIS")
          oRSFEME2("INDOCO") = oRSFEMEA("INDOCO")
          oRSFEME2("INDSEV") = oRSFEMEA("INDSEV")
          oRSFEME2("INDDET") = oRSFEMEA("INDDET")
          oRSFEME2("PRONUM") = oRSFEMEA("PRONUM")
          oRSFEME2("FALNUM") = oRSFEMEA("FALNUM")
          oRSFEME2("EFENUM") = oRSFEMEA("EFENUM")
          oRSFEME2("CAUNUM") = oRSFEMEA("CAUNUM")
          oRSFEME2("PROCESSO") = oRSFEMEA("PROCESSO")
          oRSFEME2("FALTIP") = oRSFEMEA("FALTIP")
          oRSFEME2("FALEFE") = oRSFEMEA("FALEFE")
          oRSFEME2("FALCAU") = oRSFEMEA("FALCAU")
          oRSFEME2("ACAo") = oRSFEMEA("ACAO")
          oRSFEME2("ACAREC") = oRSFEMEA("ACAREC")
          oRSFEME2("ACATOM") = oRSFEMEA("ACATOM")
          oRSFEME2("ACADAT") = oRSFEMEA("ACADAT")
          oRSFEME2("RESCOD") = oRSFEMEA("RESCOD")
          oRSFEME2("RESNOM") = oRSFEMEA("RESNOM")
          oRSFEME2("RESDAT") = oRSFEMEA("RESDAT")
          oRSFEME2("RESCOD2") = oRSFEMEA("RESCOD2")
          oRSFEME2("RESNOM2") = oRSFEMEA("RESNOM2")
          oRSFEME2("RESDAT2") = oRSFEMEA("RESDAT2")
          oRSFEME2("RESCOD3") = oRSFEMEA("RESCOD3")
          oRSFEME2("RESNOM3") = oRSFEMEA("RESNOM3")
          oRSFEME2("RESDAT3") = oRSFEMEA("RESDAT3")
          oRSFEME2("RINDRIS") = oRSFEMEA("RINDRIS")
          oRSFEME2("RINDOCO") = oRSFEMEA("RINDOCO")
          oRSFEME2("RINDSER") = oRSFEMEA("RINDSER")
          oRSFEME2("RINDDET") = oRSFEMEA("RINDDET")
          oRSFEME2("ITEM") = oRSFEMEA("ITEM")
          oRSFEME2.Update
        End If
        oRSFEMEA.MoveNext
        y = y + 1
      Wend
      oRSFEMEA.Close
      oRSFEME2.Close
      oDBFEMEA.Close
      DizerBarra ""
    Next x
    oDBPF.Close
    Set oDBFEMEA = Nothing
    Set oDBPF = Nothing
    Set oRSFEMEA = Nothing
    Set oRSFEME2 = Nothing
    Set oRSPF = Nothing
  End If



  'apurar sac
  If MDG("Apurar Sac") Then
    lSAC = True
    DizerBarra "Zerando Saldo"
    cSQL = "UPDATE FEMEARPNI SET QTDESAC=0 WHERE SEQ=" & nSEQ
    ADOComando GeracArq(cARQFEMEA), cSQL
    DizerBarra ""

    Set oDBFEMEA = New ADODB.Connection
    Set oRSFEMEA = New ADODB.Recordset
    Set oCONSAC = New ADODB.Connection
    Set oRSSAC = New ADODB.Recordset

    oDBFEMEA.ConnectionTimeout = 120
    oDBFEMEA.Open GeracArq(cARQFEMEA)
    cARQ = GeracArq(PegPath("PATH", "MANA5TGQ"), "JETFOX")
    DizerBarra "carregando sac"
    cSQLSAC = "SELECT SAC,CODIGO,DATA,cLIENTE FROM SAC ORDER BY CLIENTE"
    DizerBarra ""
    oCONSAC.ConnectionTimeout = 120
    oCONSAC.Open cARQ
    oRSSAC.Open cSQLSAC, oCONSAC, adOpenForwardOnly, adLockReadOnly


    Barra.Value = 0
    nBARPOS = 0
    nROWREC = oRSSAC.RecordCount
    oRSSAC.MoveFirst
    While Not oRSSAC.EOF
      BARPOS
      nSAC = FixInt(oRSSAC("SAC"))
      DIZAPU = FixStr(nSAC)
      DIZAPU.Refresh
      If oRSSAC("DATA") >= dINI And oRSSAC("DATA") <= dFIM Then
        nCLIENTE = oRSSAC("CLIENTE")
        cCODIGO = FixStr(oRSSAC("CODIGO"))
        If nCLIENTE = 0 Or Len(cCODIGO) = 0 Then
          Alert ("Erro SAC:" & CStr(nSAC) & " Cliente ou Codigo em Branco")
          oDBFEMEA.Close
          oRSSAC.Close
          oCONSAC.Close
          Set oRSFEMEA = Nothing
          Set oDBFEMEA = Nothing
          Screen.MousePointer = vbDefault
          Exit Sub
        End If
        cSQL = "SELECT * FROM FEMEARPNI"
        cSQL = cSQL & " WHERE CLIENTE=" & nCLIENTE
        cSQL = cSQL & " AND SEQ=" & nSEQ
        cSQL = cSQL & " AND CODIGO='" & cCODIGO & "'"
        oRSFEMEA.Open cSQL, oDBFEMEA, adOpenDynamic, adLockOptimistic
        If oRSFEMEA.EOF Then
          oRSFEMEA.AddNew
          oRSFEMEA("SEQ") = nSEQ
          oRSFEMEA("ANO") = nANO
          oRSFEMEA("MES") = nMES
          oRSFEMEA("CLIENTE") = nCLIENTE
          oRSFEMEA("CODIGO") = cCODIGO
          oRSFEMEA("QTDESAC") = 1
        Else
          oRSFEMEA("QTDESAC") = oRSFEMEA("QTDESAC") + 1
        End If
        oRSFEMEA.Update
        oRSFEMEA.Close
      End If
      oRSSAC.MoveNext
    Wend
    oDBFEMEA.Close
    oRSSAC.Close
    oCONSAC.Close
    Set oRSFEMEA = Nothing
    Set oDBFEMEA = Nothing
  End If
  'fim apuracao sac


  'apurar registro defeitos produto
  If MDG("Apurar Registro Defeito Produto") Then
    lRDP = False
    DizerBarra "Zerando indicadores"
    cSQL = "UPDATE FEMEARPNI SET QTDERDP=0 WHERE SEQ=" & nSEQ
    ADOComando GeracArq(cARQFEMEA), cSQL
    DizerBarra ""

    Set oRSFEMEA = New ADODB.Recordset
    Set oCONMSRD = New ADODB.Connection
    Set oRSMSRD = New ADODB.Recordset
    Set oDBFEMEA = New ADODB.Connection
    oDBFEMEA.ConnectionTimeout = 120
    oDBFEMEA.Open GeracArq(cARQFEMEA)

    cARQ = GeracArq(PegPath("PATH", "MANA5TGQ"), "JETFOX")
    oCONMSRD.ConnectionTimeout = 120
    oCONMSRD.Open cARQ
    DizerBarra "Carregando MRSD"
    cSQLMSRD = "SELECT MSRD,CODIGO,DATA,cLIENTE FROM MSRD ORDER BY CLIENTE"
    oRSMSRD.Open cSQLMSRD, cARQ, adOpenForwardOnly, adLockReadOnly
    DizerBarra ""

    Barra.Value = 0
    nBARPOS = 0
    nROWREC = oRSMSRD.RecordCount
    oRSMSRD.MoveFirst

    While Not oRSMSRD.EOF
      nMSRD = FixInt(oRSMSRD("MSRD"))
      BARPOS
      DIZAPU = FixStr(nMSRD)
      DIZAPU.Refresh
      If oRSMSRD("DATA") >= dINI And oRSMSRD("DATA") <= dFIM Then
        nCLIENTE = FixNum(oRSMSRD("CLIENTE"))
        cCODIGO = FixStr(oRSMSRD("CODIGO"))
        If nCLIENTE = 0 Or Len(cCODIGO) = 0 Then
          Alert ("Erro RDP:" & CStr(nMSRD) & " Cliente ou Codigo em Branco")
        End If
        cSQL = "SELECT * FROM FEMEARPNI"
        cSQL = cSQL & " WHERE CLIENTE=" & nCLIENTE
        cSQL = cSQL & " AND SEQ=" & nSEQ
        cSQL = cSQL & " AND CODIGO='" & cCODIGO & "'"
        oRSFEMEA.Open cSQL, oDBFEMEA, adOpenDynamic, adLockOptimistic
        If oRSFEMEA.EOF Then
          oRSFEMEA.AddNew
          oRSFEMEA("SEQ") = nSEQ
          oRSFEMEA("ANO") = nANO
          oRSFEMEA("MES") = nMES
          oRSFEMEA("CLIENTE") = nCLIENTE
          oRSFEMEA("CODIGO") = cCODIGO
          oRSFEMEA("QTDESAC") = 1
        Else
          oRSFEMEA("QTDERDP") = oRSFEMEA("QTDERDP") + 1
        End If
        oRSFEMEA.Update
        oRSFEMEA.Close
      End If
      oRSMSRD.MoveNext
    Wend
    oDBFEMEA.Close
    oRSMSRD.Close
    oCONMSRD.Close
    Set oCONMSRD = Nothing
    Set oRSMSRD = Nothing
    Set oRSFEMEA = Nothing
    Set oDBFEMEA = Nothing
  End If

  'atualizar saldo rpni
  If lSAC Or lRDP Then
    DizerBarra "Atualizando Saldo RPTNI"
    cSQL = "UPDATE FEMEARPNI SET QTDETOT= QTDESAC+QTDERDP  WHERE SEQ=" & nSEQ
    ADOComando GeracArq(cARQFEMEA), cSQL
    DizerBarra ""
    DizerBarra "Ajustando Clientes"
    arrumacliente (False)
    DizerBarra ""
    Filgrid
  End If



  'atualizar femea padrao operacoes
  If MDG("Apura Femea Padrao- Operacoes") Then

    DizerBarra "Apagando Apuração Anterior"
    cSQL = "DELETE FROM FEMRPNO WHERE SEQ=" & nSEQ
    ADOComando GeracArq(cARQFEMEA), cSQL
    DizerBarra ""

    Set oDBFEMEA = New ADODB.Connection
    Set oRSPF = New ADODB.Recordset

    DizerBarra "Apurando Operacoes"
    oDBFEMEA.ConnectionTimeout = 120
    oDBFEMEA.Open GeracArq(cARQFEMEA)
    cSQL = "SELECT DISTINCT PF,PRONUM FROM FEMPF ORDER BY PF,PRONUM"
    oRSPF.Open cSQL, oDBFEMEA, adOpenForwardOnly, adLockReadOnly
    DizerBarra ""

    Barra.Value = 0
    nBARPOS = 0
    nROWREC = oRSPF.RecordCount
    oRSPF.MoveFirst

    Do While Not oRSPF.EOF
      BARPOS
      nPF = FixInt(oRSPF("PF"))
      DIZAPU = FixStr(nPF)
      DIZAPU.Refresh
      nPRONUM = FixInt(oRSPF("PRONUM"))
      nTOTRPN = 0
      nTOTCAU = 0
      nTOTM40 = 0
      nMAIRPN = 0
      nFX01 = 0
      nFX02 = 0
      nFX03 = 0

      'cSQL = "SELECT INDRIS FROM FEMEA WHERE PF=" & nPF & " AND PRONUM=" & nPRONUM
      cSQL = "SELECT INDRIS FROM FEMEAPAD WHERE PF=" & nPF & " AND PRONUM=" & nPRONUM
      Set oRSFEMEA = New ADODB.Recordset
      oRSFEMEA.Open cSQL, oDBFEMEA, adOpenForwardOnly, adLockReadOnly
      Do While Not oRSFEMEA.EOF
        nINDRIS = FixInt(oRSFEMEA("INDRIS"))
        nTOTRPN = nTOTRPN + nINDRIS
        nTOTCAU = nTOTCAU + 1
        If nINDRIS > 40 Then nTOTM40 = nTOTM40 + 1
        If nINDRIS > nMAIRPN Then
          nMAIRPN = FixInt(oRSFEMEA("INDRIS"))
        End If
        If nINDRIS < 41 Then nFX01 = nFX01 + 1
        If nINDRIS > 40 And nINDRIS < 101 Then nFX02 = nFX02 + 1
        If nINDRIS > 100 Then nFX03 = nFX03 + 1
        oRSFEMEA.MoveNext
      Loop
      If nTOTCAU > 0 Then
        cSQL = "SELECT * FROM FEMRPNO WHERE SEQ=" & nSEQ & " AND PF=" & nPF & " AND PRONUM=" & nPRONUM
        Set oRSFEMEA = New ADODB.Recordset
        oRSFEMEA.Open cSQL, oDBFEMEA, adOpenDynamic, adLockOptimistic
        If oRSFEMEA.EOF Then
          oRSFEMEA.AddNew
          oRSFEMEA("SEQ") = nSEQ
          oRSFEMEA("ANO") = nANO
          oRSFEMEA("MES") = nMES
          oRSFEMEA("PF") = nPF
          oRSFEMEA("PRONUM") = nPRONUM
        End If
        oRSFEMEA("TOTRPN") = nTOTRPN
        oRSFEMEA("TOTCAU") = nTOTCAU
        oRSFEMEA("TOTM40") = nTOTM40
        oRSFEMEA("MAIRPN") = nMAIRPN
        oRSFEMEA("TOTFX01") = nFX01
        oRSFEMEA("TOTFX02") = nFX02
        oRSFEMEA("TOTFX03") = nFX03
        oRSFEMEA.Update
      End If
      oRSPF.MoveNext
    Loop
    oRSFEMEA.Close
    oRSPF.Close
    oDBFEMEA.Close
  End If



  If MDG("Trazer Saldo Anteriores") Then
    nMESANT = nMES - 1
    nANOANT = nANO
    If nMESANT < 1 Then
      nMESANT = nMESANT + 12
      nANOANT = nANOANT - 1
    End If

    Set oDBFEMEA = New ADODB.Connection
    oDBFEMEA.ConnectionTimeout = 120
    oDBFEMEA.Open GeracArq(cARQFEMEA)

    cSQL = "SELECT PF,PRONUM,TOTRPN03,TOTRPN01,TOTRPN02,TOTFXA1,TOTFXA2,TOTFXA3,TOTFXB1,TOTFXB2,TOTFXB3,TOTFXC1,TOTFXC2,TOTFXC3 FROM FEMRPNO WHERE SEQ=" & nSEQ & " ORDER BY PF,PRONUM"
    Set oRSFEMEA = New ADODB.Recordset
    oRSFEMEA.Open cSQL, oDBFEMEA, adOpenDynamic, adLockOptimistic


    Barra.Value = 0
    nBARPOS = 0
    nROWREC = oRSFEMEA.RecordCount
    oRSFEMEA.MoveFirst

    Do While Not oRSFEMEA.EOF
      nPF = FixInt(oRSFEMEA("PF"))
      nPRONUM = FixInt(oRSFEMEA("PRONUM"))
      DIZAPU = FixStr(nPF) & "/" & FixStr(nPRONUM)
      DIZAPU.Refresh
      BARPOS
      cSQL = "SELECT TOTFX01,TOTFX02,TOTFX03,TOTRPN,TOTRPN01,TOTRPN02,TOTFXA1,TOTFXA2,TOTFXA3,TOTFXB1,TOTFXB2,TOTFXB3,TOTFXC1,TOTFXC2,TOTFXC3 FROM FEMRPNO WHERE PF=" & nPF & " AND ANO=" & nANOANT & " AND MES=" & nMESANT & " AND PRONUM=" & nPRONUM
      Set oRSFEME2 = New ADODB.Recordset
      oRSFEME2.Open cSQL, oDBFEMEA, adOpenDynamic, adLockReadOnly
      If Not oRSFEME2.EOF Then
        oRSFEME2.MoveFirst
        oRSFEMEA("TOTRPN01") = oRSFEME2("TOTRPN")
        oRSFEMEA("TOTRPN02") = oRSFEME2("TOTRPN01")
        oRSFEMEA("TOTRPN03") = oRSFEME2("TOTRPN02")
        oRSFEMEA("TOTFXA1") = oRSFEME2("TOTFX01")
        oRSFEMEA("TOTFXB1") = oRSFEME2("TOTFXA1")
        oRSFEMEA("TOTFXC1") = oRSFEME2("TOTFXB1")
        oRSFEMEA("TOTFXA2") = oRSFEME2("TOTFX02")
        oRSFEMEA("TOTFXB2") = oRSFEME2("TOTFXA2")
        oRSFEMEA("TOTFXC2") = oRSFEME2("TOTFXB2")
        oRSFEMEA("TOTFXA3") = oRSFEME2("TOTFX03")
        oRSFEMEA("TOTFXB3") = oRSFEME2("TOTFXA3")
        oRSFEMEA("TOTFXC3") = oRSFEME2("TOTFXB3")
        oRSFEMEA.Update
      End If
      oRSFEMEA.MoveNext
    Loop
    oRSFEME2.Close
    oRSFEMEA.Close
    oDBFEMEA.Close
  End If


  If MDG("Gerar Dados Grafico") Then
    ''Fase 2 Graficos
    ''1 Anterior
    eRETU01 = MesAnt(nMES, nANO)
    nMESANT1 = eRETU01(0)
    nANOANT1 = eRETU01(1)
    ''2 Anterior
    eRETU01 = MesAnt(nMESANT1, nANOANT1)
    nMESANT2 = eRETU01(0)
    nANOANT2 = eRETU01(1)
    ''3 Anterior
    eRETU01 = MesAnt(nMESANT2, nANOANT2)
    nMESANT3 = eRETU01(0)
    nANOANT3 = eRETU01(1)

    cSQL = "SELECT * FROM FEMRPNG WHERE ANO=" & nANO & " AND MES=" & nMES
    ApagaSQL cARQFEMEA, cSQL

    Set oDBFEMEA = New ADODB.Connection
    oDBFEMEA.ConnectionTimeout = 120
    oDBFEMEA.Open GeracArq(cARQFEMEA)

    Set oDBPF = New ADODB.Connection
    oDBPF.ConnectionTimeout = 120
    oDBPF.Open GeracArq(cARQPF)

    cSQL = "SELECT PF FROM PF  WHERE not bloqueado ORDER BY PF"
    Set oRSPF = New ADODB.Recordset
    oRSPF.Open cSQL, oDBPF, adOpenForwardOnly, adLockReadOnly

    Barra.Value = 0
    nBARPOS = 0
    nROWREC = oRSPF.RecordCount
    oRSPF.MoveFirst


    Do While Not oRSPF.EOF
      nPF = FixInt(oRSPF("PF"))
      aGRVAL = Array(0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
      DIZAPU = FixStr(nPF)
      DIZAPU.Refresh
      BARPOS
      cSQL = "SELECT * FROM FEMRPNO WHERE ANO=" & nANO & " AND MES=" & nMES & " AND PF=" & nPF
      Set oRSFEMEA = New ADODB.Recordset
      oRSFEMEA.Open cSQL, oDBFEMEA, adOpenForwardOnly, adLockReadOnly
      Do While Not oRSFEMEA.EOF
        aGRVAL(0) = aGRVAL(0) + FixInt(oRSFEMEA("TOTFX01"))
        aGRVAL(1) = aGRVAL(1) + FixInt(oRSFEMEA("TOTFXA1"))
        aGRVAL(2) = aGRVAL(2) + FixInt(oRSFEMEA("TOTFXB1"))
        aGRVAL(3) = aGRVAL(3) + FixInt(oRSFEMEA("TOTFXC1"))
        aGRVAL(4) = aGRVAL(4) + FixInt(oRSFEMEA("TOTFX02"))
        aGRVAL(5) = aGRVAL(5) + FixInt(oRSFEMEA("TOTFXA2"))
        aGRVAL(6) = aGRVAL(6) + FixInt(oRSFEMEA("TOTFXB2"))
        aGRVAL(7) = aGRVAL(7) + FixInt(oRSFEMEA("TOTFXC2"))
        aGRVAL(8) = aGRVAL(8) + FixInt(oRSFEMEA("TOTFX03"))
        aGRVAL(9) = aGRVAL(9) + FixInt(oRSFEMEA("TOTFXA3"))
        aGRVAL(10) = aGRVAL(10) + FixInt(oRSFEMEA("TOTFXB3"))
        aGRVAL(11) = aGRVAL(11) + FixInt(oRSFEMEA("TOTFXC3"))
        oRSFEMEA.MoveNext
      Loop
      oRSPF.MoveNext
      cSQL2 = " SELECT * FROM FEMRPNG WHERE ANO=" & nANO & " AND NMES=" & nMES & " and fx01=" & aGRVAL(0) & " and fx02=" & aGRVAL(4) & " and fx03=" & aGRVAL(8)
      IncluiSQL cARQFEMEA, cSQL2, 8, Array("ANO", "MES", "PF", "FX01", "FX02", "FX03", "ANODIZ", "MESDIZ"), Array(nANO, nMES, nPF, aGRVAL(0), aGRVAL(4), aGRVAL(8), nANO, nMES), False, False
      cSQL2 = " SELECT * FROM FEMRPNG WHERE ANO=" & nANOANT1 & " AND NMES=" & nMESANT1 & " and fx01=" & aGRVAL(1) & " and fx02=" & aGRVAL(5) & " and fx03=" & aGRVAL(9)
      IncluiSQL cARQFEMEA, cSQL2, 8, Array("ANO", "MES", "PF", "FX01", "FX02", "FX03", "ANODIZ", "MESDIZ"), Array(nANO, nMES, nPF, aGRVAL(1), aGRVAL(5), aGRVAL(9), nANOANT1, nMESANT1), False, False
      cSQL2 = " SELECT * FROM FEMRPNG WHERE ANO=" & nANOANT2 & " AND NMES=" & nMESANT2 & " and fx01=" & aGRVAL(2) & " and fx02=" & aGRVAL(6) & " and fx03=" & aGRVAL(10)
      IncluiSQL cARQFEMEA, cSQL2, 8, Array("ANO", "MES", "PF", "FX01", "FX02", "FX03", "ANODIZ", "MESDIZ"), Array(nANO, nMES, nPF, aGRVAL(2), aGRVAL(6), aGRVAL(10), nANOANT2, nMESANT2), False, False
      cSQL2 = " SELECT * FROM FEMRPNG WHERE ANO=" & nANOANT3 & " AND NMES=" & nMESANT3 & " and fx01=" & aGRVAL(3) & " and fx02=" & aGRVAL(7) & " and fx03=" & aGRVAL(11)
      IncluiSQL cARQFEMEA, cSQL2, 8, Array("ANO", "MES", "PF", "FX01", "FX02", "FX03", "ANODIZ", "MESDIZ"), Array(nANO, nMES, nPF, aGRVAL(3), aGRVAL(7), aGRVAL(11), nANOANT3, nMESANT3), False, False
    Loop
    oRSFEMEA.Close
    oRSPF.Close
  End If

  tEXT(12) = Today()
  apurado = 1
  Screen.MousePointer = vbDefault
  Barra.Value = 0
  DIZAPU = " "
  DIZAPU.Refresh
  Filgrid
  filgrid10
  filgridope
  Alert ("Apuracao Encerrada")
  Exit Sub



errhandler:
  Select Case Err.Number
  Case 3421, 424
    ''SayErro "Apurando Femea"
    Resume Next
  Case Else
    SayErro "Apurando Femea"
    Screen.MousePointer = vbDefault
    Exit Sub
  End Select

End Sub

Private Function PEGCLI(ByVal nCLIENTE As Long) As Variant
  Dim cARQ As String
  Dim cSQL As String
  cARQ = GeraConn(PegPath("PATH", "MANA5EMP"), "JETFOX")
  cSQL = "SELECT NOME,COGNOME,GRUPOEMP FROM MA01 WHERE NUMERO=" & nCLIENTE
  PEGCLI = PegSQL(cARQ, cSQL, 3, Array("COGNOME", "NOME", "GRUPOEMP"), Array("C", "C", "C"), Array("", "", ""))
End Function

Private Sub CmdCheckcli_Click()
  arrumacliente (True)
End Sub

Private Sub arrumacliente(Optional ByVal lPER As Boolean = False)
  Dim oDBPF As New ADODB.Connection
  Dim oRSFEMEA As ADODB.Recordset
  Dim oCONMA01 As New ADODB.Connection
  Dim oRSMA01 As New ADODB.Recordset

  Dim cARQ As String
  Dim cSQL  As String
  Dim cCOGNOME  As String
  Dim cCLINOME  As String
  Dim cGRUPOEMP As String
  Dim lCONT As Boolean
  Dim nCLIENTE As Long
  '' On Error Resume Next

  If lPER Then
    If MDG("Somente esta Sequencia") Then
      cSQL = "SELECT CLIENTE,COGNOME,CLINOME,GRUPOEMP FROM FEMEARPNI WHERE SEQ=" & aVAL(0) & " ORDER BY CLIENTE"
    Else
      cSQL = "SELECT CLIENTE,COGNOME,CLINOME,GRUPOEMP FROM FEMEARPNI ORDER BY CLIENTE"
    End If
  Else
    cSQL = "SELECT CLIENTE,COGNOME,CLINOME,GRUPOEMP FROM FEMEARPNI WHERE SEQ=" & aVAL(0) & " ORDER BY CLIENTE"
  End If

  oDBPF.ConnectionTimeout = 120
  oDBPF.Open GeracArq(cARQFEMEA)
  oRSFEMEA.Open cSQL, oDBPF, adOpenDynamic, adLockOptimistic
  cARQ = GeracArq(PegPath("PATH", "MANA5EMP"), "JETFOX")
  oCONMA01.ConnectionTimeout = 120
  oCONMA01.Open cARQ
  While Not oRSFEMEA.EOF
    nCLIENTE = FixNum(oRSFEMEA("CLIENTE"))
    cSQL = "SELECT NOME,COGNOME,GRUPOEMP FROM MA01 WHERE NUMERO=" & nCLIENTE
    oRSMA01.Open cSQL, oCONMA01, adOpenForwardOnly, adLockReadOnly
    cCOGNOME = "-"
    cCLINOME = "-"
    cGRUPOEMP = "-"
    If Not oRSMA01.EOF Then
      cCOGNOME = FixStr(oRSMA01("COGNOME"))
      cCLINOME = FixStr(oRSMA01("NOME"))
      cGRUPOEMP = FixStr(oRSMA01("GRUPOEMP"))
    End If
    oRSMA01.Close
    lCONT = True
    While lCONT
      oRSFEMEA("COGNOME") = cCOGNOME
      oRSFEMEA("CLINOME") = cCLINOME
      oRSFEMEA("GRUPOEMP") = cGRUPOEMP
      oRSFEMEA.Update
      oRSFEMEA.MoveNext
      If oRSFEMEA.EOF Then
        lCONT = False
      Else
        If nCLIENTE <> FixNum(oRSFEMEA("CLIENTE")) Then
          lCONT = False
        End If
      End If
    Wend
  Wend
  oRSFEMEA.Close
  oCONMA01.Close
  oDBPF.Close
  If lPER Then
    Alert ("Checagem Concluida")
  End If
  Filgrid
End Sub

Private Sub cmdClose_Click()
  Unload Me
  Screen.MousePointer = vbDefault
End Sub

Private Sub CmdEditar_Click()
  Dim aVAL As Variant
  aVAL = Array(0, 0, 0, 0, 0)

  Grid10.Col = 3
  aVAL(0) = Grid10
  Grid10.Col = 4
  aVAL(1) = Grid10
  Grid10.Col = 5
  aVAL(2) = Grid10
  Grid10.Col = 6
  aVAL(3) = Grid10
  Grid10.Col = 0
  aVAL(4) = Grid10
  ePASS01 = "select * from FEMRPNT WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1) & " AND EFENUM=" & aVAL(2) & " AND CAUNUM=" & aVAL(3) & " AND PF=" & aVAL(4)
  ePASS02 = cARQFEMEA
  frmFEMEI.Show vbModal, Me

End Sub

Private Sub Command1_Click()

  Dim cCODIGO As String
  Dim nCLIENTE As Integer
  Dim cSQL As String
  Dim aCLI As Variant
  nCLIENTE = FixNum(Val(Busca("NÝ Cliente", "Incluindo RPN/ITEM", "1", 8)))

  cCODIGO = FixStr(InputBox("Digite o Codigo", "Incluindo", "__"), "")
  If nCLIENTE = 0 Then
    Alert ("Cliente em Branco")
    Exit Sub
  End If
  If Len(cCODIGO) = 0 Then
    Alert ("Codigo em Branco")
    Exit Sub
  End If
  cCODIGO = Left(cCODIGO, 24)
  cSQL = "SELECT * FROM FEMEARPNI"
  cSQL = cSQL & " WHERE CLIENTE=" & nCLIENTE
  cSQL = cSQL & " AND SEQ=" & FixNum(aVAL(0))
  cSQL = cSQL & " AND CODIGO='" & cCODIGO & "'"
  aCLI = PEGCLI(nCLIENTE)
  IncluiSQL cARQFEMEA, cSQL, 8, Array("SEQ", "MES", "ANO", "CLIENTE", "CODIGO", "COGNOME", "CLINOME", "GRUPOEMP"), Array(aVAL(0), aVAL(1), aVAL(2), nCLIENTE, cCODIGO, aCLI(0), aCLI(1), aCLI(2)), True, True
  Filgrid

End Sub

Private Sub Command2_Click()

  Dim cSQL As String
  cSQL = "SELECT * FROM FEMEARPNI"

  Grid.Col = 1
  cSQL = cSQL & " WHERE CLIENTE=" & FixNum(Grid)
  cSQL = cSQL & " AND SEQ=" & FixNum(aVAL(0))
  Grid.Col = 0
  cSQL = cSQL & " AND CODIGO='" & FixStr(Grid) & "'"
  If ApagaSQLP(cARQFEMEA, cSQL) Then
    Filgrid
  End If

End Sub

Private Sub Command3_Click()

  Dim nPFTMP As Integer
  Dim nPROTMP As Long
  GridOpe.Col = 0

  nPFTMP = FixInt(GridOpe)
  GridOpe.Col = 1
  nPROTMP = FixInt(GridOpe)
  ePASS01 = "SELECT * FROM FEMRPNO WHERE SEQ=" & FixInt(tEXT(0)) & " AND PF=" & nPFTMP & " AND PRONUM=" & nPROTMP
  FrmFEMRPNO.Show vbModal, Me

End Sub

Private Sub EditSeq_Click()

  Dim cSQL As String
  cSQL = "SELECT * FROM FEMEARPNI"

  Grid.Col = 1
  cSQL = cSQL & " WHERE CLIENTE=" & FixNum(Grid)
  cSQL = cSQL & " AND SEQ=" & FixNum(aVAL(0))
  Grid.Col = 0
  cSQL = cSQL & " AND CODIGO='" & FixStr(Grid) & "'"
  Alert ("Nao Disponivel")

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  If MDG("Gravar e Sair", "Gravando ") Then
    For nITEM = 0 To 12
      aVAL(nITEM) = tEXT(nITEM)
    Next nITEM
    aVAL(13) = DTPicker1
    aVAL(14) = DTPicker2
    aVAL(15) = Anual
    aVAL(16) = Semestral
    aVAL(17) = apurado
    GrvSQL cARQFEMEA, cSQLUSO, nCAMPOS, aCAM, aVAL, aFOR
    GravaLog 999, 0, "Gravou "
  Else
    GravaLog 999, 0, "Nao Gravou "
  End If
  Screen.MousePointer = vbDefault
  Unload Me

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  ''Configura Help
  Me.Caption = cFORMID
  ''HelpContextID = nFORMID
  ''Arquivo
  cARQPF = PegPath("PATH", "PF")
  cARQFEMEA = PegPath("PATH", "FEMEA")
  If Not FileConnExist(cARQFEMEA, True) Then
    End
  End If
  'Variaveis sql
  cSQLUSO = ePASS01
  nCAMPOS = 18
  aCAM = Array("SEQ", "MES", "ANO", "DESCRI", "FX01", "FX02", "FX03", "FX04", "FX05", "FX06", "FX07", "FX08", "APUDATA", "DIAINI", "DIAFIM", "ANUAL", "SEMES", "APURADO")
  aFOR = Array("N", "MES", "ANO", "C", "N", "N", "N", "N", "N", "N", "N", "N", "D", "D", "D", "BN", "BN", "BN")
  aPAD = Array(0, 0, 0, "", 0, 0, 0, 0, 0, 0, 0, 0, "", "", "", False, False, False)
  aVAL = PegSQL(cARQFEMEA, cSQLUSO, nCAMPOS, aCAM, aFOR, aPAD)
  For nITEM = 0 To 12
    tEXT(nITEM) = aVAL(nITEM)
  Next nITEM
  If aVAL(13) <> "" Then DTPicker1 = aVAL(13)
  If aVAL(14) <> "" Then DTPicker2 = aVAL(14)
  mes.Value = tEXT(1)
  ano.Value = tEXT(2)
  Anual = aVAL(15)
  Semestral = aVAL(16)
  apurado = aVAL(17)
  Filgrid
  filgrid10
  filgridope

End Sub

'Private Sub maisano_Click()
'    text(2) = CStr(IncDec(text(2), 1, 1900, 2300))
'End Sub

'Private Sub maismes_Click()
'    text(1) = CStr(IncDec(text(1), 1, 1, 12))
'End Sub

'Private Sub menosano_Click()
'    text(2) = CStr(IncDec(text(2), -1, 1900, 2300))
'End Sub

'Private Sub menosmes_Click()
'    text(1) = CStr(IncDec(text(1), -1, 1, 12))
'End Sub

Private Sub Filgrid()
  Dim cSQL As String
  DizerBarra "Carregando RPNI"
  cSQL = "SELECT CODIGO,CLIENTE,COGNOME,QTDESAC,QTDERDP FROM FEMEARPNI WHERE SEQ=" & aVAL(0)
  MontaGridUltra Grid, 5, Array(3000, 800, 1200, 800, 800), Array("Codigo", "Cliente", "Cognome", "SAC", "RDP"), Array("L$CODIGO", "CLIENTE", "L$COGNOME", "QTDESAC", "QTDERDP"), cARQFEMEA, cSQL
  DizerBarra ""
End Sub

Private Sub filgrid10()
  Dim cSQL As String
  DizerBarra "Carregando RPNT"
  cSQL = "SELECT TIPOAPU, PF, ITEM, INDRIS, PRONUM, FALNUM, EFENUM, CAUNUM, ALTMAN, FALEFE, FALCAU,50 FROM FEMRPNT WHERE SEQ=" & FixInt(tEXT(0)) & " ORDER BY INDRIS DESC"
  MontaGrid Grid10, 11, Array(400, 400, 400, 400, 400, 400, 400, 400, 300, 2500, 2500), Array("TIPOAPU", "PF", "ITEM", "RPN", "Pro", "Fal", "Efe", "Cau", "Man", "Efeito", "Causa"), Array("TIPOAPU", "PF", "ITEM", "INDRIS", "PRONUM", "FALNUM", "EFENUM", "CAUNUM", "ALTMAN", "FALEFE", "FALCAU"), cARQFEMEA, cSQL
  DizerBarra ""
End Sub

Private Sub filgridope()
  Dim cSQL As String
  DizerBarra "Carregando RPNO"
  cSQL = "SELECT PF, PRONUM, TOTRPN, TOTCAU, TOTM40, MAIRPN, TOTFX01, TOTFX02, TOTFX03, TOTRPN01, TOTRPN02, TOTRPN03 FROM FEMRPNO WHERE SEQ=" & FixInt(tEXT(0)) & " ORDER BY PF,PRONUM"
  MontaGridUltra GridOpe, 12, Array(500, 500, 800, 500, 500, 500, 500, 500, 500, 800, 800, 800), Array("PF", "Pro", "Tot RPN", "Cau", ">40", "Maior", "<40", "40-100", ">100", "Ant", "Ant", "Ant"), Array("PF", "PRONUM", "TOTRPN", "TOTCAU", "TOTM40", "MAIRPN", "TOTFX01", "TOTFX02", "TOTFX03", "TOTRPN01", "TOTRPN02", "TOTRPN03"), cARQFEMEA, cSQL
  DizerBarra ""
End Sub

Private Sub BARPOS()
  Dim nPOSATU As Long
  nPOSATU = 0
  nBARPOS = nBARPOS + 1
  If nROWREC > 0 Then
    nPOSATU = Int(nBARPOS / nROWREC * 100)
  End If
  If nPOSATU > 0 And nPOSATU < 101 Then
    Debug.Print nPOSATU
    Barra.Value = nPOSATU
  End If
  DIZAPU.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub Text_GotFocus(Index As Integer)
  FocusMe
End Sub

Private Sub tEXT_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index = 1 Or Index = 2 Then
    KeyAscii = ValiText(KeyAscii, "#NI")
  End If
  If Index >= 4 And Index <= 11 Then
    KeyAscii = ValiText(KeyAscii, "#NI")
  End If
End Sub

