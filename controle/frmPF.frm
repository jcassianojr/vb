VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.2#0"; "vbccr17.ocx"
Object = "{EA478B61-D9EC-47F6-BB21-95A533AF2251}#1.0#0"; "TabExC01.OCX"
Object = "{451B73A5-1563-45D5-A6AC-7B2B7D30B778}#1.0#0"; "BSPrin10.ocx"
Begin VB.Form frmPF 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Processo de Fabricação"
   ClientHeight    =   8265
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   12015
   Icon            =   "frmPF.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   12015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin BSPrinter.PrintPreview PrintPreview1 
      Left            =   11160
      Top             =   3000
      _ExtentX        =   1191
      _ExtentY        =   1191
      LcK1b           =   "yefT59bnyufI583n1ufV59HnyefN58nn0+fK58nnzefR58rn1+fW5w=="
      LcK2b           =   "reeI58zn+ue059jn1Of656fni+eN543nleeM54HnlueR5w=="
      AmbientBb       =   $"frmPF.frx":058A
   End
   Begin VB.TextBox txtfields 
      Height          =   285
      Index           =   66
      Left            =   3960
      TabIndex        =   198
      Top             =   1800
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "LX-->"
      Height          =   315
      Index           =   11
      Left            =   6840
      TabIndex        =   196
      TabStop         =   0   'False
      ToolTipText     =   "Clique Para Trazer Nome Cliente"
      Top             =   120
      Width           =   555
   End
   Begin VB.CommandButton Command7 
      Caption         =   "LX-->"
      Height          =   315
      Index           =   10
      Left            =   4680
      TabIndex        =   195
      TabStop         =   0   'False
      ToolTipText     =   "Clique Para Trazer Nome Cliente"
      Top             =   120
      Width           =   555
   End
   Begin VB.CommandButton ESCCLI 
      Caption         =   "LX2"
      Height          =   315
      Index           =   2
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   194
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton ESCCLI 
      Caption         =   "LX"
      Height          =   315
      Index           =   1
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   193
      TabStop         =   0   'False
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Cmdciclopeca 
      Caption         =   "codigo_lx"
      Height          =   315
      Left            =   9360
      TabIndex        =   192
      TabStop         =   0   'False
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton cad_cjt_lx 
      Caption         =   "cad_des_lx"
      Height          =   255
      Left            =   1680
      TabIndex        =   191
      TabStop         =   0   'False
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdcaddes 
      Caption         =   "cad_des_lx"
      Height          =   255
      Left            =   1680
      TabIndex        =   190
      TabStop         =   0   'False
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Esctipo 
      Height          =   315
      Left            =   7200
      Picture         =   "frmPF.frx":062A
      Style           =   1  'Graphical
      TabIndex        =   188
      TabStop         =   0   'False
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox TipoDescritivo 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   187
      TabStop         =   0   'False
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CommandButton ESCpro 
      Caption         =   "MC"
      Height          =   315
      Index           =   8
      Left            =   9840
      TabIndex        =   160
      TabStop         =   0   'False
      Top             =   2760
      Width           =   375
   End
   Begin VB.CommandButton ESCpro 
      Caption         =   "MC"
      Height          =   315
      Index           =   7
      Left            =   9360
      TabIndex        =   159
      TabStop         =   0   'False
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton ESCpro 
      Caption         =   "MC"
      Height          =   315
      Index           =   6
      Left            =   2160
      TabIndex        =   158
      TabStop         =   0   'False
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton ESCpro 
      Caption         =   "LX"
      Height          =   315
      Index           =   5
      Left            =   9480
      TabIndex        =   147
      TabStop         =   0   'False
      Top             =   2760
      Width           =   375
   End
   Begin VB.CommandButton ESCpro 
      Height          =   315
      Index           =   4
      Left            =   8640
      Picture         =   "frmPF.frx":0BB4
      Style           =   1  'Graphical
      TabIndex        =   146
      TabStop         =   0   'False
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton ESCpro 
      Caption         =   "LX"
      Height          =   315
      Index           =   3
      Left            =   9000
      TabIndex        =   145
      TabStop         =   0   'False
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton ESCpro 
      Caption         =   "LX"
      Height          =   315
      Index           =   2
      Left            =   1680
      TabIndex        =   144
      TabStop         =   0   'False
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtfields 
      Height          =   285
      Index           =   0
      Left            =   7560
      MaxLength       =   24
      TabIndex        =   8
      Text            =   "12345678901234"
      Top             =   1200
      Width           =   1755
   End
   Begin VB.CommandButton ESCpro 
      Height          =   315
      Index           =   1
      Left            =   9120
      Picture         =   "frmPF.frx":113E
      Style           =   1  'Graphical
      TabIndex        =   140
      TabStop         =   0   'False
      Top             =   2760
      Width           =   375
   End
   Begin VB.TextBox txtfields 
      Height          =   285
      Index           =   32
      Left            =   9240
      TabIndex        =   5
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "-->"
      Height          =   315
      Index           =   1
      Left            =   6480
      TabIndex        =   138
      TabStop         =   0   'False
      ToolTipText     =   "Clique Para Trazer Codigo Cliente"
      Top             =   120
      Width           =   315
   End
   Begin VB.CommandButton esciedx 
      Height          =   315
      Left            =   1560
      Picture         =   "frmPF.frx":16C8
      Style           =   1  'Graphical
      TabIndex        =   134
      TabStop         =   0   'False
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command8 
      Caption         =   "-->"
      Height          =   315
      Left            =   3840
      TabIndex        =   130
      TabStop         =   0   'False
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "-->"
      Height          =   315
      Index           =   0
      Left            =   4440
      TabIndex        =   129
      TabStop         =   0   'False
      ToolTipText     =   "Clique Para Trazer Nome Cliente"
      Top             =   120
      Width           =   315
   End
   Begin VB.TextBox txtfields 
      Height          =   285
      Index           =   34
      Left            =   7440
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtfields 
      Enabled         =   0   'False
      Height          =   285
      Index           =   33
      Left            =   8640
      MaxLength       =   2
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtfields 
      Height          =   285
      Index           =   14
      Left            =   8160
      MaxLength       =   50
      TabIndex        =   19
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Embalagem"
      Height          =   375
      Left            =   10440
      TabIndex        =   102
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox txtfields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   28
      Left            =   6600
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1800
      Width           =   495
   End
   Begin VB.CommandButton ESCpro 
      Height          =   315
      Index           =   0
      Left            =   1260
      Picture         =   "frmPF.frx":1C52
      Style           =   1  'Graphical
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton ESCCLI 
      Height          =   315
      Index           =   0
      Left            =   3060
      Picture         =   "frmPF.frx":21DC
      Style           =   1  'Graphical
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtfields 
      Height          =   285
      Index           =   5
      Left            =   3180
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   480
      Width           =   4215
   End
   Begin VB.TextBox txtfields 
      Height          =   285
      Index           =   18
      Left            =   5760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   2400
      Width           =   3735
   End
   Begin VB.TextBox txtfields 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   17
      Left            =   10080
      TabIndex        =   11
      Top             =   1800
      Width           =   1395
   End
   Begin VB.TextBox txtfields 
      Height          =   285
      Index           =   16
      Left            =   6960
      TabIndex        =   18
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtfields 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   15
      Left            =   5760
      TabIndex        =   17
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtfields 
      Height          =   285
      Index           =   13
      Left            =   2760
      TabIndex        =   16
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox txtfields 
      Height          =   285
      Index           =   12
      Left            =   120
      MaxLength       =   24
      TabIndex        =   15
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox txtfields 
      Height          =   285
      Index           =   10
      Left            =   2760
      TabIndex        =   13
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtfields 
      Height          =   285
      Index           =   9
      Left            =   120
      MaxLength       =   24
      TabIndex        =   12
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox txtfields 
      Height          =   285
      Index           =   7
      Left            =   2760
      TabIndex        =   10
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox txtfields 
      Height          =   285
      Index           =   6
      Left            =   120
      MaxLength       =   24
      TabIndex        =   9
      Top             =   1800
      Width           =   2535
   End
   Begin VB.TextBox txtfields 
      Height          =   285
      Index           =   4
      Left            =   2040
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtfields 
      Height          =   285
      Index           =   3
      Left            =   3720
      MaxLength       =   40
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1200
      Width           =   3735
   End
   Begin VB.TextBox txtfields 
      Height          =   285
      Index           =   2
      Left            =   120
      MaxLength       =   24
      TabIndex        =   6
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox txtfields 
      Height          =   285
      Index           =   1
      Left            =   1140
      MaxLength       =   10
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox TXTPF 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   480
      Width           =   915
   End
   Begin TabExCtl.SSTabEx SSTab1 
      Height          =   4455
      Index           =   2
      Left            =   0
      TabIndex        =   51
      Top             =   3720
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7858
      Tabs            =   7
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabsPerRow      =   7
      Tab             =   5
      TabHeight       =   520
      TabCaption(0)   =   "Matéria Prima 1"
      Tab(0).ControlCount=   28
      Tab(0).Control(0)=   "lblLabels(22)"
      Tab(0).Control(1)=   "lblLabels(23)"
      Tab(0).Control(2)=   "lblLabels(24)"
      Tab(0).Control(3)=   "lblLabels(25)"
      Tab(0).Control(4)=   "lblLabels(26)"
      Tab(0).Control(5)=   "lblLabels(21)"
      Tab(0).Control(6)=   "VerImg(0)"
      Tab(0).Control(7)=   "CmdSaveimg(0)"
      Tab(0).Control(8)=   "DelImg(0)"
      Tab(0).Control(9)=   "Incimg(0)"
      Tab(0).Control(10)=   "CmdPaste(0)"
      Tab(0).Control(11)=   "cmdimprimir(0)"
      Tab(0).Control(12)=   "cmdFOTO(2)"
      Tab(0).Control(13)=   "txtfields(24)"
      Tab(0).Control(14)=   "escmu(1)"
      Tab(0).Control(15)=   "txtfields(20)"
      Tab(0).Control(16)=   "txtfields(23)"
      Tab(0).Control(17)=   "txtfields(27)"
      Tab(0).Control(18)=   "txtfields(35)"
      Tab(0).Control(19)=   "txtfields(36)"
      Tab(0).Control(20)=   "txtfields(19)"
      Tab(0).Control(21)=   "txtfields(47)"
      Tab(0).Control(22)=   "txtfields(48)"
      Tab(0).Control(23)=   "Command9(0)"
      Tab(0).Control(24)=   "CmdApagar(1)"
      Tab(0).Control(25)=   "Picture1(0)"
      Tab(0).Control(26)=   "Picture2(0)"
      Tab(0).Control(27)=   "ESCpro(9)"
      TabCaption(1)   =   "Matéria Prima 2"
      Tab(1).ControlCount=   28
      Tab(1).Control(0)=   "lblLabels(27)"
      Tab(1).Control(1)=   "lblLabels(28)"
      Tab(1).Control(2)=   "lblLabels(29)"
      Tab(1).Control(3)=   "lblLabels(30)"
      Tab(1).Control(4)=   "lblLabels(31)"
      Tab(1).Control(5)=   "lblLabels(37)"
      Tab(1).Control(6)=   "VerImg(1)"
      Tab(1).Control(7)=   "CmdSaveimg(1)"
      Tab(1).Control(8)=   "DelImg(1)"
      Tab(1).Control(9)=   "Incimg(1)"
      Tab(1).Control(10)=   "CmdPaste(1)"
      Tab(1).Control(11)=   "cmdimprimir(1)"
      Tab(1).Control(12)=   "cmdFOTO(3)"
      Tab(1).Control(13)=   "txtfields(25)"
      Tab(1).Control(14)=   "escmu(2)"
      Tab(1).Control(15)=   "txtfields(21)"
      Tab(1).Control(16)=   "txtfields(37)"
      Tab(1).Control(17)=   "txtfields(38)"
      Tab(1).Control(18)=   "txtfields(39)"
      Tab(1).Control(19)=   "txtfields(40)"
      Tab(1).Control(20)=   "txtfields(41)"
      Tab(1).Control(21)=   "txtfields(49)"
      Tab(1).Control(22)=   "txtfields(50)"
      Tab(1).Control(23)=   "Command9(0)"
      Tab(1).Control(24)=   "CmdApagar(1)"
      Tab(1).Control(25)=   "Picture1(1)"
      Tab(1).Control(26)=   "Picture2(0)"
      Tab(1).Control(27)=   "ESCpro(9)"
      TabCaption(2)   =   "Matéria Prima 3"
      Tab(2).ControlCount=   28
      Tab(2).Control(0)=   "lblLabels(32)"
      Tab(2).Control(1)=   "lblLabels(33)"
      Tab(2).Control(2)=   "lblLabels(34)"
      Tab(2).Control(3)=   "lblLabels(35)"
      Tab(2).Control(4)=   "lblLabels(36)"
      Tab(2).Control(5)=   "lblLabels(38)"
      Tab(2).Control(6)=   "VerImg(2)"
      Tab(2).Control(7)=   "CmdSaveimg(2)"
      Tab(2).Control(8)=   "DelImg(2)"
      Tab(2).Control(9)=   "Incimg(2)"
      Tab(2).Control(10)=   "CmdPaste(2)"
      Tab(2).Control(11)=   "cmdimprimir(2)"
      Tab(2).Control(12)=   "cmdFOTO(4)"
      Tab(2).Control(13)=   "txtfields(26)"
      Tab(2).Control(14)=   "escmu(3)"
      Tab(2).Control(15)=   "txtfields(22)"
      Tab(2).Control(16)=   "txtfields(42)"
      Tab(2).Control(17)=   "txtfields(43)"
      Tab(2).Control(18)=   "txtfields(44)"
      Tab(2).Control(19)=   "txtfields(45)"
      Tab(2).Control(20)=   "txtfields(46)"
      Tab(2).Control(21)=   "txtfields(61)"
      Tab(2).Control(22)=   "txtfields(62)"
      Tab(2).Control(23)=   "Command9(2)"
      Tab(2).Control(24)=   "CmdApagar(3)"
      Tab(2).Control(25)=   "Picture1(2)"
      Tab(2).Control(26)=   "Picture2(2)"
      Tab(2).Control(27)=   "ESCpro(11)"
      TabCaption(3)   =   "Composição"
      Tab(3).ControlCount=   7
      Tab(3).Control(0)=   "Command5"
      Tab(3).Control(1)=   "Command4"
      Tab(3).Control(2)=   "Command6"
      Tab(3).Control(3)=   "cmdFOTO(5)"
      Tab(3).Control(4)=   "GRIDCOM"
      Tab(3).Control(5)=   "CmdcomIMp"
      Tab(3).Control(6)=   "ImpComPF"
      TabCaption(4)   =   "Elab.Rev.Bloq."
      Tab(4).ControlCount=   24
      Tab(4).Control(0)=   "Command10"
      Tab(4).Control(1)=   "txtfields(65)"
      Tab(4).Control(2)=   "txtfields(64)"
      Tab(4).Control(3)=   "Check2"
      Tab(4).Control(4)=   "txtfields(63)"
      Tab(4).Control(5)=   "txtfields(29)"
      Tab(4).Control(6)=   "CmdMotRev"
      Tab(4).Control(7)=   "NovaRev"
      Tab(4).Control(8)=   "txtfields(30)"
      Tab(4).Control(9)=   "escmp04a"
      Tab(4).Control(10)=   "txtfields(31)"
      Tab(4).Control(11)=   "txtfields(8)"
      Tab(4).Control(12)=   "Check1"
      Tab(4).Control(13)=   "txtfields(11)"
      Tab(4).Control(14)=   "gridrev"
      Tab(4).Control(15)=   "gridrevi"
      Tab(4).Control(16)=   "ComMotInc"
      Tab(4).Control(17)=   "ComMotAlt"
      Tab(4).Control(18)=   "ComMotExc"
      Tab(4).Control(19)=   "Label4"
      Tab(4).Control(20)=   "Label3"
      Tab(4).Control(21)=   "Label10"
      Tab(4).Control(22)=   "Label12"
      Tab(4).Control(23)=   "Label1"
      TabCaption(5)   =   "Sequencia"
      Tab(5).ControlCount=   11
      Tab(5).Control(0)=   "ESCpro(10)"
      Tab(5).Control(1)=   "Picture2(1)"
      Tab(5).Control(2)=   "CmdApagar(2)"
      Tab(5).Control(3)=   "Command9(1)"
      Tab(5).Control(4)=   "IMPSEQ"
      Tab(5).Control(5)=   "Command1"
      Tab(5).Control(6)=   "Command2"
      Tab(5).Control(7)=   "EditSeq"
      Tab(5).Control(8)=   "seqreo"
      Tab(5).Control(9)=   "GridSeq"
      Tab(5).Control(10)=   "Label5"
      TabCaption(6)   =   "Fluxo Mat Prima"
      Tab(6).ControlCount=   30
      Tab(6).Control(0)=   "EscCodFlu(10)"
      Tab(6).Control(1)=   "EscCodFlu(9)"
      Tab(6).Control(2)=   "EscCodFlu(8)"
      Tab(6).Control(3)=   "EscCodFlu(7)"
      Tab(6).Control(4)=   "txtfields(60)"
      Tab(6).Control(5)=   "txtfields(59)"
      Tab(6).Control(6)=   "txtfields(58)"
      Tab(6).Control(7)=   "txtfields(57)"
      Tab(6).Control(8)=   "EscCodFlu(6)"
      Tab(6).Control(9)=   "EscCodFlu(5)"
      Tab(6).Control(10)=   "EscCodFlu(4)"
      Tab(6).Control(11)=   "EscCodFlu(3)"
      Tab(6).Control(12)=   "EscCodFlu(2)"
      Tab(6).Control(13)=   "EscCodFlu(1)"
      Tab(6).Control(14)=   "txtfields(56)"
      Tab(6).Control(15)=   "txtfields(55)"
      Tab(6).Control(16)=   "txtfields(54)"
      Tab(6).Control(17)=   "txtfields(53)"
      Tab(6).Control(18)=   "txtfields(52)"
      Tab(6).Control(19)=   "txtfields(51)"
      Tab(6).Control(20)=   "CMDSIG(1)"
      Tab(6).Control(21)=   "CMDSIG(2)"
      Tab(6).Control(22)=   "CMDSIG(3)"
      Tab(6).Control(23)=   "CMDSIG(4)"
      Tab(6).Control(24)=   "CMDSIG(5)"
      Tab(6).Control(25)=   "CMDSIG(6)"
      Tab(6).Control(26)=   "CMDSIG(7)"
      Tab(6).Control(27)=   "CMDSIG(8)"
      Tab(6).Control(28)=   "CMDSIG(9)"
      Tab(6).Control(29)=   "CMDSIG(10)"
      Begin VB.CommandButton Command10 
         Caption         =   "Distribuicao"
         Height          =   495
         Left            =   -67320
         TabIndex        =   186
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   65
         Left            =   -72000
         TabIndex        =   183
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   64
         Left            =   -73800
         TabIndex        =   182
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Item Seguranca"
         Height          =   375
         Left            =   -70800
         TabIndex        =   181
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton ESCpro 
         Caption         =   "LX"
         Height          =   315
         Index           =   11
         Left            =   -68280
         TabIndex        =   180
         TabStop         =   0   'False
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton ESCpro 
         Caption         =   "LX"
         Height          =   315
         Index           =   10
         Left            =   6600
         TabIndex        =   179
         TabStop         =   0   'False
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton ESCpro 
         Caption         =   "LX"
         Height          =   315
         Index           =   9
         Left            =   -68400
         TabIndex        =   178
         TabStop         =   0   'False
         Top             =   840
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         Height          =   2535
         Index           =   2
         Left            =   -74760
         ScaleHeight     =   2475
         ScaleWidth      =   2835
         TabIndex        =   177
         Top             =   600
         Width           =   2895
      End
      Begin VB.PictureBox Picture1 
         Height          =   255
         Index           =   2
         Left            =   -71520
         ScaleHeight     =   195
         ScaleWidth      =   555
         TabIndex        =   176
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox Picture2 
         Height          =   2535
         Index           =   1
         Left            =   240
         ScaleHeight     =   2475
         ScaleWidth      =   2835
         TabIndex        =   175
         Top             =   480
         Width           =   2895
      End
      Begin VB.PictureBox Picture1 
         Height          =   255
         Index           =   1
         Left            =   -71640
         ScaleHeight     =   195
         ScaleWidth      =   315
         TabIndex        =   174
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         Height          =   2535
         Index           =   0
         Left            =   -74760
         ScaleHeight     =   2475
         ScaleWidth      =   2835
         TabIndex        =   173
         Top             =   480
         Width           =   2895
      End
      Begin VB.PictureBox Picture1 
         Height          =   255
         Index           =   0
         Left            =   -71760
         ScaleHeight     =   195
         ScaleWidth      =   555
         TabIndex        =   172
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   63
         Left            =   -67080
         Locked          =   -1  'True
         TabIndex        =   171
         TabStop         =   0   'False
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   255
         Index           =   10
         Left            =   -71640
         Picture         =   "frmPF.frx":2766
         Style           =   1  'Graphical
         TabIndex        =   170
         Top             =   2880
         Width           =   255
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   255
         Index           =   9
         Left            =   -71640
         Picture         =   "frmPF.frx":2CF0
         Style           =   1  'Graphical
         TabIndex        =   169
         TabStop         =   0   'False
         Top             =   2400
         Width           =   255
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   255
         Index           =   8
         Left            =   -71640
         Picture         =   "frmPF.frx":327A
         Style           =   1  'Graphical
         TabIndex        =   168
         TabStop         =   0   'False
         Top             =   1920
         Width           =   255
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   255
         Index           =   7
         Left            =   -71640
         Picture         =   "frmPF.frx":3804
         Style           =   1  'Graphical
         TabIndex        =   167
         TabStop         =   0   'False
         Top             =   1440
         Width           =   255
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   255
         Index           =   6
         Left            =   -71640
         Picture         =   "frmPF.frx":3D8E
         Style           =   1  'Graphical
         TabIndex        =   166
         TabStop         =   0   'False
         Top             =   960
         Width           =   255
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   255
         Index           =   5
         Left            =   -74280
         Picture         =   "frmPF.frx":4318
         Style           =   1  'Graphical
         TabIndex        =   165
         TabStop         =   0   'False
         Top             =   2880
         Width           =   255
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   255
         Index           =   4
         Left            =   -74280
         Picture         =   "frmPF.frx":48A2
         Style           =   1  'Graphical
         TabIndex        =   164
         TabStop         =   0   'False
         Top             =   2400
         Width           =   255
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   255
         Index           =   3
         Left            =   -74280
         Picture         =   "frmPF.frx":4E2C
         Style           =   1  'Graphical
         TabIndex        =   163
         TabStop         =   0   'False
         Top             =   1920
         Width           =   255
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   255
         Index           =   2
         Left            =   -74280
         Picture         =   "frmPF.frx":53B6
         Style           =   1  'Graphical
         TabIndex        =   162
         TabStop         =   0   'False
         Top             =   1440
         Width           =   255
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   255
         Index           =   1
         Left            =   -74280
         Picture         =   "frmPF.frx":5940
         Style           =   1  'Graphical
         TabIndex        =   161
         TabStop         =   0   'False
         Top             =   960
         Width           =   255
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   51
         Left            =   -73440
         TabIndex        =   52
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   52
         Left            =   -72840
         TabIndex        =   53
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   53
         Left            =   -73440
         TabIndex        =   54
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   54
         Left            =   -72840
         TabIndex        =   55
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   55
         Left            =   -73440
         TabIndex        =   56
         Top             =   1680
         Width           =   615
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   56
         Left            =   -72840
         TabIndex        =   57
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton EscCodFlu 
         Caption         =   "..."
         Height          =   255
         Index           =   1
         Left            =   -74040
         TabIndex        =   157
         TabStop         =   0   'False
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton EscCodFlu 
         Caption         =   "..."
         Height          =   255
         Index           =   2
         Left            =   -72120
         TabIndex        =   156
         TabStop         =   0   'False
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton EscCodFlu 
         Caption         =   "..."
         Height          =   255
         Index           =   3
         Left            =   -74040
         TabIndex        =   155
         TabStop         =   0   'False
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton EscCodFlu 
         Caption         =   "..."
         Height          =   255
         Index           =   4
         Left            =   -72120
         TabIndex        =   154
         TabStop         =   0   'False
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton EscCodFlu 
         Caption         =   "..."
         Height          =   255
         Index           =   5
         Left            =   -74040
         TabIndex        =   153
         TabStop         =   0   'False
         Top             =   1920
         Width           =   495
      End
      Begin VB.CommandButton EscCodFlu 
         Caption         =   "..."
         Height          =   255
         Index           =   6
         Left            =   -72120
         TabIndex        =   152
         TabStop         =   0   'False
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   57
         Left            =   -73440
         TabIndex        =   58
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   58
         Left            =   -72840
         TabIndex        =   59
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   59
         Left            =   -73440
         TabIndex        =   60
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox txtfields 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   18
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Index           =   60
         Left            =   -72840
         TabIndex        =   61
         Top             =   2640
         Width           =   615
      End
      Begin VB.CommandButton EscCodFlu 
         Caption         =   "..."
         Height          =   255
         Index           =   7
         Left            =   -74040
         TabIndex        =   151
         TabStop         =   0   'False
         Top             =   2400
         Width           =   495
      End
      Begin VB.CommandButton EscCodFlu 
         Caption         =   "..."
         Height          =   255
         Index           =   8
         Left            =   -72120
         TabIndex        =   150
         TabStop         =   0   'False
         Top             =   2400
         Width           =   495
      End
      Begin VB.CommandButton EscCodFlu 
         Caption         =   "..."
         Height          =   255
         Index           =   9
         Left            =   -74040
         TabIndex        =   149
         TabStop         =   0   'False
         Top             =   2880
         Width           =   495
      End
      Begin VB.CommandButton EscCodFlu 
         Caption         =   "..."
         Height          =   255
         Index           =   10
         Left            =   -72120
         TabIndex        =   148
         TabStop         =   0   'False
         Top             =   2880
         Width           =   495
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   29
         Left            =   -72600
         Locked          =   -1  'True
         TabIndex        =   143
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton CmdApagar 
         Caption         =   "Apagar"
         Height          =   375
         Index           =   3
         Left            =   -67080
         TabIndex        =   137
         TabStop         =   0   'False
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton CmdApagar 
         Caption         =   "Apagar"
         Height          =   375
         Index           =   2
         Left            =   7680
         TabIndex        =   136
         TabStop         =   0   'False
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton CmdApagar 
         Caption         =   "Apagar"
         Height          =   375
         Index           =   1
         Left            =   -67320
         TabIndex        =   135
         TabStop         =   0   'False
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         Caption         =   "-->"
         Height          =   255
         Index           =   2
         Left            =   -67800
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   1020
         Width           =   435
      End
      Begin VB.CommandButton Command9 
         Caption         =   "-->"
         Height          =   255
         Index           =   1
         Left            =   7200
         TabIndex        =   132
         TabStop         =   0   'False
         Top             =   900
         Width           =   435
      End
      Begin VB.CommandButton Command9 
         Caption         =   "-->"
         Height          =   255
         Index           =   0
         Left            =   -67800
         TabIndex        =   131
         TabStop         =   0   'False
         Top             =   900
         Width           =   435
      End
      Begin VB.CommandButton ImpComPF 
         Caption         =   "Importar/PF"
         Height          =   255
         Left            =   -66720
         TabIndex        =   128
         TabStop         =   0   'False
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CommandButton IMPSEQ 
         Caption         =   "Importar"
         Height          =   255
         Left            =   10080
         TabIndex        =   127
         TabStop         =   0   'False
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   62
         Left            =   -66240
         TabIndex        =   47
         Top             =   1980
         Width           =   375
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   61
         Left            =   -66660
         TabIndex        =   46
         Top             =   1980
         Width           =   375
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   50
         Left            =   -65940
         TabIndex        =   38
         Top             =   1860
         Width           =   375
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   49
         Left            =   -66360
         TabIndex        =   37
         Top             =   1860
         Width           =   375
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   48
         Left            =   -66000
         TabIndex        =   29
         Top             =   1860
         Width           =   375
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   47
         Left            =   -66420
         TabIndex        =   28
         Top             =   1860
         Width           =   375
      End
      Begin VB.CommandButton CmdMotRev 
         Caption         =   "Motivos"
         Height          =   375
         Left            =   -67320
         TabIndex        =   119
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Novo"
         Height          =   255
         Left            =   10080
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Excluir"
         Height          =   255
         Left            =   10080
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   1320
         Width           =   975
      End
      Begin VB.CommandButton EditSeq 
         Caption         =   "Editar"
         Height          =   255
         Left            =   10080
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton seqreo 
         Caption         =   "Reordenar"
         Height          =   255
         Left            =   10080
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton NovaRev 
         Caption         =   "Nova Revisão"
         Height          =   375
         Left            =   -71160
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   30
         Left            =   -73740
         Locked          =   -1  'True
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   540
         Width           =   1095
      End
      Begin VB.CommandButton escmp04a 
         Height          =   315
         Left            =   -72600
         Picture         =   "frmPF.frx":5ECA
         Style           =   1  'Graphical
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   540
         Width           =   375
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   31
         Left            =   -72120
         Locked          =   -1  'True
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   540
         Width           =   4335
      End
      Begin VB.TextBox txtfields 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   -73800
         Locked          =   -1  'True
         TabIndex        =   106
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Bloqueado/Inativo"
         Height          =   375
         Left            =   -69000
         TabIndex        =   49
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   11
         Left            =   -67080
         MaxLength       =   50
         TabIndex        =   50
         Top             =   960
         Width           =   2535
      End
      Begin VB.CommandButton CmdcomIMp 
         Caption         =   "Importar/Mana5"
         Height          =   255
         Left            =   -66720
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox txtfields 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   19
         Left            =   -68880
         TabIndex        =   26
         Top             =   1860
         Width           =   1095
      End
      Begin VB.TextBox txtfields 
         Height          =   525
         Index           =   46
         Left            =   -71400
         MaxLength       =   130
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   48
         Top             =   2580
         Width           =   6465
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   45
         Left            =   -67800
         TabIndex        =   45
         Top             =   1980
         Width           =   1065
      End
      Begin VB.TextBox txtfields 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   44
         Left            =   -69000
         TabIndex        =   44
         Top             =   1980
         Width           =   1065
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   43
         Left            =   -70200
         TabIndex        =   43
         Top             =   1980
         Width           =   1065
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   42
         Left            =   -71400
         TabIndex        =   42
         Top             =   1980
         Width           =   1065
      End
      Begin VB.TextBox txtfields 
         Height          =   525
         Index           =   41
         Left            =   -71400
         MaxLength       =   130
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   39
         Top             =   2460
         Width           =   6465
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   40
         Left            =   -67560
         TabIndex        =   36
         Top             =   1860
         Width           =   1065
      End
      Begin VB.TextBox txtfields 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Index           =   39
         Left            =   -68760
         TabIndex        =   35
         Top             =   1860
         Width           =   1065
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   38
         Left            =   -70080
         TabIndex        =   34
         Top             =   1860
         Width           =   1065
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   37
         Left            =   -71400
         TabIndex        =   33
         Top             =   1860
         Width           =   1065
      End
      Begin VB.TextBox txtfields 
         Height          =   525
         Index           =   36
         Left            =   -71520
         MaxLength       =   130
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   2460
         Width           =   6465
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   35
         Left            =   -67680
         TabIndex        =   27
         Top             =   1860
         Width           =   1065
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   27
         Left            =   -70200
         TabIndex        =   25
         Top             =   1860
         Width           =   1065
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   23
         Left            =   -71520
         TabIndex        =   24
         Top             =   1860
         Width           =   1065
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   20
         Left            =   -71520
         TabIndex        =   22
         Top             =   900
         Width           =   2500
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   21
         Left            =   -71400
         TabIndex        =   31
         Top             =   900
         Width           =   2500
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   22
         Left            =   -71400
         TabIndex        =   40
         Top             =   1020
         Width           =   2500
      End
      Begin VB.CommandButton escmu 
         Height          =   315
         Index           =   1
         Left            =   -68880
         Picture         =   "frmPF.frx":6454
         Style           =   1  'Graphical
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton escmu 
         Height          =   315
         Index           =   2
         Left            =   -68760
         Picture         =   "frmPF.frx":69DE
         Style           =   1  'Graphical
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton escmu 
         Height          =   315
         Index           =   3
         Left            =   -68760
         Picture         =   "frmPF.frx":6F68
         Style           =   1  'Graphical
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   24
         Left            =   -71520
         TabIndex        =   23
         Top             =   1260
         Width           =   5175
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   25
         Left            =   -71400
         TabIndex        =   32
         Top             =   1260
         Width           =   5175
      End
      Begin VB.TextBox txtfields 
         Height          =   285
         Index           =   26
         Left            =   -71400
         TabIndex        =   41
         Top             =   1380
         Width           =   5175
      End
      Begin MSFlexGridLib.MSFlexGrid GRIDCOM 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   480
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   6376
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridSeq 
         Height          =   3615
         Left            =   120
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   360
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   6376
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrev 
         Height          =   1095
         Left            =   -74880
         TabIndex        =   120
         TabStop         =   0   'False
         Top             =   1680
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   1931
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrevi 
         Height          =   1335
         Left            =   -74880
         TabIndex        =   121
         TabStop         =   0   'False
         Top             =   2880
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   2355
         _Version        =   393216
      End
      Begin XPControls.XPButton cmdFOTO 
         Height          =   375
         Index           =   2
         Left            =   -66000
         TabIndex        =   203
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         Picture         =   "frmPF.frx":74F2
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
      Begin XPControls.XPButton cmdFOTO 
         Height          =   495
         Index           =   3
         Left            =   -65880
         TabIndex        =   204
         Top             =   600
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Picture         =   "frmPF.frx":7A8C
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
      Begin XPControls.XPButton cmdFOTO 
         Height          =   495
         Index           =   4
         Left            =   -65640
         TabIndex        =   205
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Picture         =   "frmPF.frx":8026
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
      Begin XPControls.XPButton cmdFOTO 
         Height          =   495
         Index           =   5
         Left            =   -66720
         TabIndex        =   206
         TabStop         =   0   'False
         Top             =   3360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Picture         =   "frmPF.frx":85C0
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
      Begin XPControls.XPButton cmdimprimir 
         Height          =   435
         Index           =   0
         Left            =   -72960
         TabIndex        =   207
         Top             =   3120
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
         Picture         =   "frmPF.frx":8B5A
         Caption         =   "Imprimir"
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
      Begin XPControls.XPButton cmdimprimir 
         Height          =   435
         Index           =   1
         Left            =   -71640
         TabIndex        =   208
         Top             =   3120
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
         Picture         =   "frmPF.frx":90F4
         Caption         =   "Imprimir"
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
      Begin XPControls.XPButton cmdimprimir 
         Height          =   435
         Index           =   2
         Left            =   -73080
         TabIndex        =   209
         Top             =   3720
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
         Picture         =   "frmPF.frx":968E
         Caption         =   "Imprimir"
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
      Begin XPControls.XPButton CmdPaste 
         Height          =   435
         Index           =   0
         Left            =   -71280
         TabIndex        =   210
         Top             =   3120
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   767
         Picture         =   "frmPF.frx":9C28
         Caption         =   "Copia Area Transferencia"
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
      Begin XPControls.XPButton CmdPaste 
         Height          =   435
         Index           =   1
         Left            =   -70080
         TabIndex        =   211
         Top             =   3120
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   767
         Picture         =   "frmPF.frx":A1C2
         Caption         =   "Copia Area Transferencia"
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
      Begin XPControls.XPButton CmdPaste 
         Height          =   435
         Index           =   2
         Left            =   -71400
         TabIndex        =   212
         Top             =   3240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   767
         Picture         =   "frmPF.frx":A75C
         Caption         =   "Copia Area Transferencia"
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
      Begin XPControls.XPButton Command6 
         Height          =   435
         Left            =   -66600
         TabIndex        =   213
         TabStop         =   0   'False
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPF.frx":ACF6
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
      Begin XPControls.XPButton Command4 
         Height          =   435
         Left            =   -66600
         TabIndex        =   214
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPF.frx":B290
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
      Begin XPControls.XPButton Command5 
         Height          =   435
         Left            =   -66600
         TabIndex        =   215
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPF.frx":B82A
         Caption         =   "Excluir"
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
      Begin XPControls.XPButton ComMotInc 
         Height          =   435
         Left            =   -67320
         TabIndex        =   216
         TabStop         =   0   'False
         Top             =   2880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPF.frx":BDC4
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
      Begin XPControls.XPButton ComMotAlt 
         Height          =   435
         Left            =   -67320
         TabIndex        =   217
         TabStop         =   0   'False
         Top             =   3360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPF.frx":C35E
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
      Begin XPControls.XPButton ComMotExc 
         Height          =   435
         Left            =   -67320
         TabIndex        =   218
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPF.frx":C8F8
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
      Begin XPControls.XPButton Incimg 
         Height          =   435
         Index           =   1
         Left            =   -74760
         TabIndex        =   219
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPF.frx":CE92
         Caption         =   "Incluir Imagem"
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
      Begin XPControls.XPButton Incimg 
         Height          =   435
         Index           =   0
         Left            =   -74760
         TabIndex        =   220
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPF.frx":D32C
         Caption         =   "Incluir Imagem"
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
      Begin XPControls.XPButton Incimg 
         Height          =   435
         Index           =   2
         Left            =   -74760
         TabIndex        =   221
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPF.frx":D7C6
         Caption         =   "Incluir Imagem"
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
      Begin XPControls.XPButton DelImg 
         Height          =   435
         Index           =   1
         Left            =   -74760
         TabIndex        =   223
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPF.frx":DC60
         Caption         =   "Excluir Imagem"
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
      Begin XPControls.XPButton DelImg 
         Height          =   435
         Index           =   0
         Left            =   -74760
         TabIndex        =   224
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPF.frx":E0FA
         Caption         =   "Excluir Imagem"
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
      Begin XPControls.XPButton DelImg 
         Height          =   435
         Index           =   2
         Left            =   -74760
         TabIndex        =   225
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPF.frx":E594
         Caption         =   "Excluir Imagem"
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
      Begin XPControls.XPButton CmdSaveimg 
         Height          =   435
         Index           =   1
         Left            =   -72960
         TabIndex        =   226
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPF.frx":EA2E
         Caption         =   "Salvar Imagem"
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
      Begin XPControls.XPButton CmdSaveimg 
         Height          =   435
         Index           =   0
         Left            =   -72960
         TabIndex        =   227
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPF.frx":EEC8
         Caption         =   "Salvar Imagem"
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
      Begin XPControls.XPButton CmdSaveimg 
         Height          =   435
         Index           =   2
         Left            =   -73080
         TabIndex        =   228
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPF.frx":F362
         Caption         =   "Salvar Imagem"
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
      Begin XPControls.XPButton VerImg 
         Height          =   495
         Index           =   1
         Left            =   -71280
         TabIndex        =   229
         TabStop         =   0   'False
         Top             =   3600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         Picture         =   "frmPF.frx":F7FC
         Caption         =   "Navegar Imagens"
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
      Begin XPControls.XPButton VerImg 
         Height          =   375
         Index           =   0
         Left            =   -71280
         TabIndex        =   230
         TabStop         =   0   'False
         Top             =   3600
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         Picture         =   "frmPF.frx":FC96
         Caption         =   "Navegar Imagens"
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
      Begin XPControls.XPButton VerImg 
         Height          =   495
         Index           =   2
         Left            =   -71400
         TabIndex        =   222
         TabStop         =   0   'False
         Top             =   3720
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         Picture         =   "frmPF.frx":10130
         Caption         =   "Navegar Imagens"
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
         Caption         =   "SEQ=999 para operacao de controles"
         Height          =   255
         Left            =   240
         TabIndex        =   189
         Top             =   3960
         Width           =   2895
      End
      Begin VB.Label Label4 
         Caption         =   "takta"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -72600
         TabIndex        =   185
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "takt"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   184
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Sequencia"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   38
         Left            =   -66660
         TabIndex        =   126
         Top             =   1740
         Width           =   915
      End
      Begin VB.Label lblLabels 
         Caption         =   "Sequencia"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   37
         Left            =   -66420
         TabIndex        =   125
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label lblLabels 
         Caption         =   "Sequencia"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   21
         Left            =   -66480
         TabIndex        =   124
         Top             =   1620
         Width           =   915
      End
      Begin VB.Label Label10 
         Caption         =   "Elaborado"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   112
         Top             =   540
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Em"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -67560
         TabIndex        =   111
         Top             =   540
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Rev.Processo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   110
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Observações"
         Height          =   255
         Index           =   36
         Left            =   -71400
         TabIndex        =   101
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Aprov. %"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   35
         Left            =   -67680
         TabIndex        =   100
         Top             =   1740
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Peso Bruto"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   34
         Left            =   -69000
         TabIndex        =   99
         Top             =   1740
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Qtde"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   33
         Left            =   -70080
         TabIndex        =   98
         Top             =   1740
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Passo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   32
         Left            =   -71400
         TabIndex        =   97
         Top             =   1740
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Observações"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   31
         Left            =   -71400
         TabIndex        =   96
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Aprov. %"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   30
         Left            =   -67320
         TabIndex        =   95
         Top             =   1620
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Peso Bruto"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   29
         Left            =   -68760
         TabIndex        =   94
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Qtde"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   28
         Left            =   -69960
         TabIndex        =   93
         Top             =   1620
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Passo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   27
         Left            =   -71400
         TabIndex        =   92
         Top             =   1620
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Observações"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   26
         Left            =   -71520
         TabIndex        =   91
         Top             =   2220
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Aprov. %"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   25
         Left            =   -67440
         TabIndex        =   90
         Top             =   1620
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Peso Bruto"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   24
         Left            =   -68880
         TabIndex        =   89
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Qtde"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   23
         Left            =   -70080
         TabIndex        =   88
         Top             =   1620
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Passo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   22
         Left            =   -71520
         TabIndex        =   87
         Top             =   1620
         Width           =   735
      End
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   9960
      TabIndex        =   199
      TabStop         =   0   'False
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmPF.frx":105CA
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
      Left            =   9960
      TabIndex        =   200
      TabStop         =   0   'False
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmPF.frx":10B64
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
   Begin XPControls.XPButton CmdConfImp 
      Height          =   435
      Left            =   9960
      TabIndex        =   141
      TabStop         =   0   'False
      Top             =   2160
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmPF.frx":110FE
      Caption         =   "Configurar Impressora"
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
      Height          =   375
      Index           =   0
      Left            =   2640
      TabIndex        =   201
      TabStop         =   0   'False
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Picture         =   "frmPF.frx":11698
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
   Begin XPControls.XPButton cmdFOTO 
      Height          =   375
      Index           =   1
      Left            =   10320
      TabIndex        =   202
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Picture         =   "frmPF.frx":11C32
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
   Begin VBCCR17.DTPicker DTPicker1 
      Height          =   375
      Left            =   4920
      TabIndex        =   231
      Top             =   1800
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Value           =   44285
      CheckBox        =   -1  'True
      AllowUserInput  =   -1  'True
   End
   Begin VBCCR17.DTPicker DTPicker2 
      Height          =   375
      Left            =   3960
      TabIndex        =   232
      Top             =   2400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Value           =   44285
      CheckBox        =   -1  'True
      AllowUserInput  =   -1  'True
   End
   Begin VBCCR17.DTPicker DTPicker3 
      Height          =   375
      Left            =   3960
      TabIndex        =   233
      Top             =   3120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Value           =   44285
      CheckBox        =   -1  'True
      AllowUserInput  =   -1  'True
   End
   Begin VB.Label lblLabels 
      Caption         =   "No.Versao"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   41
      Left            =   3960
      TabIndex        =   197
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cod Integrado"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   40
      Left            =   7560
      TabIndex        =   142
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Caption         =   "Fabrica"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   39
      Left            =   9240
      TabIndex        =   139
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo Cliente"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   20
      Left            =   7440
      TabIndex        =   123
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Opção"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   8640
      TabIndex        =   122
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo Final"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   19
      Left            =   8040
      TabIndex        =   103
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Tipo"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   6600
      TabIndex        =   83
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label11 
      Caption         =   "PF"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   82
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "Motivo Revisão"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   18
      Left            =   5640
      TabIndex        =   79
      Top             =   2160
      Width           =   1395
   End
   Begin VB.Label lblLabels 
      Caption         =   "Peso Liquido"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   17
      Left            =   10080
      TabIndex        =   78
      Top             =   1560
      Width           =   915
   End
   Begin VB.Label lblLabels 
      Caption         =   "Qtde Conjunto"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   16
      Left            =   6840
      TabIndex        =   77
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Caption         =   "Peso Conjunto"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   15
      Left            =   5640
      TabIndex        =   76
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Caption         =   "Data Revisão Cjto"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   14
      Left            =   3960
      TabIndex        =   75
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cjto Rev"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   13
      Left            =   2760
      TabIndex        =   74
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Conjunto Desenho"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   73
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      Caption         =   "Data Nossa Revisão"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   11
      Left            =   3960
      TabIndex        =   72
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nossa Revisão"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   10
      Left            =   2760
      TabIndex        =   71
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nosso Desenho"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   70
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Data da Revisão"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   8
      Left            =   4920
      TabIndex        =   69
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblLabels 
      Caption         =   "Revisão Cliente"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   7
      Left            =   2760
      TabIndex        =   68
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Caption         =   "Desenho Cliente"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   67
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nome:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   5
      Left            =   5400
      TabIndex        =   66
      Top             =   240
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Caption         =   "CLIENTE:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   4
      Left            =   2040
      TabIndex        =   65
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descrição"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   3
      Left            =   4320
      TabIndex        =   64
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Caption         =   "CODIGO:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   63
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "CPF:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   62
      Top             =   240
      Width           =   375
   End
End
Attribute VB_Name = "frmPF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cARQPF As String
Dim cARQPFP As String
Dim cARQPFG As String
Dim cARQFEMEA As String
Dim lOLDBLD As Boolean
Dim cSQL As String
Dim aVAL As Variant
Dim aFOR As Variant
Dim aCAM As Variant
Dim aPAD As Variant
Dim nCAMPOS As Integer
Dim iLOOP As Integer
Dim lTROCOU As Variant
Dim aCAMIMG As Variant

Private Sub cad_cjt_lx_Click()
    gravacaddes txtFields(12).Text, txtFields(13).Text, DTPicker3.Value
End Sub

Private Sub CmdApagar_Click(Index As Integer)
    Select Case Index
    Case 1
        txtFields(20) = ""
        txtFields(24) = ""
        txtFields(23) = 0
        txtFields(27) = 0
        txtFields(19) = 0
        txtFields(35) = 0
        txtFields(47) = 0
        txtFields(48) = ""
    Case 2
        txtFields(21) = ""
        txtFields(25) = ""
        txtFields(37) = 0
        txtFields(38) = 0
        txtFields(39) = 0
        txtFields(40) = 0
        txtFields(49) = 0
        txtFields(50) = 0
        txtFields(41) = ""
    Case 3
        txtFields(22) = ""
        txtFields(26) = ""
        txtFields(42) = 0
        txtFields(43) = 0
        txtFields(44) = 0
        txtFields(45) = 0
        txtFields(61) = 0
        txtFields(62) = 0
        txtFields(46) = ""
    End Select
End Sub

Private Sub cmdcaddes_Click()
    If Len(txtFields(66).Text) > 0 Then
        gravacaddes txtFields(6).Text, txtFields(66).Text, DTPicker1.Value
    Else
        gravacaddes txtFields(6).Text, txtFields(7).Text, DTPicker1.Value
    End If
End Sub

Private Sub Cmdciclopeca_Click()
    TXTFIELDS_LostFocus 0
End Sub

Private Sub cmdClose_Click()
    On Error Resume Next
    Dim cSQLUP As String
    Dim cSQLIMG As String

    If Check1 And Len(txtFields(11)) = 0 Then
        Alert ("Preencher Motivo do Bloqueio")
        Exit Sub
    End If
    
    aVAL(67) = FixNumBol(Check2.Value)           'seguranca 66
    aVAL(68) = FixNumBol(Check1.Value)           'bloqueado 67
    
    aVAL(69) = DTPicker1.Value                   'clidat 68
    aVAL(70) = DTPicker2.Value                   'nosdat 69
    aVAL(71) = DTPicker3.Value                   'condat 70
       
    ATUPFPG
    If lOLDBLD <> FixNumBol(Check1) Then
        DizerBarra "Checando Status Femea"
        If Check1 Then
            cSQLUP = "UPDATE FEMEA SET BLOQUEADO=TRUE WHERE PF=" & nPF
            ADOComando cARQPF, cSQLUP
            ADOComando cARQFEMEA, cSQLUP
            cSQLUP = "UPDATE FEMAVU SET BLOQUEADO=TRUE WHERE PF=" & nPF
            ADOComando cARQFEMEA, cSQLUP
        Else
            cSQLUP = "UPDATE FEMEA SET BLOQUEADO=FALSE WHERE PF=" & nPF
            ADOComando cARQPF, cSQLUP
            ADOComando cARQFEMEA, cSQLUP
            cSQLUP = "UPDATE FEMAVU SET BLOQUEADO=FALSE WHERE PF=" & nPF
            ADOComando cARQFEMEA, cSQLUP
        End If
        DizerBarra ""
    End If
    
    For iLOOP = 0 To nCAMPOS - 5
        aVAL(iLOOP) = txtFields(iLOOP)
    Next iLOOP
    GrvSQL cARQPF, cSQL, nCAMPOS, aCAM, aVAL, aFOR
       
    For iLOOP = 0 To 2
        If lTROCOU(iLOOP) Then
            cSQLIMG = "select PF," & aCAMIMG(iLOOP) & " from PF WHERE PF=" & nPF
            ADOGrvBlob cARQPF, cSQLIMG, Picture1(iLOOP), aCAMIMG(iLOOP)
        End If
    Next iLOOP

    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub CmdcomIMp_Click()
    Dim DB1 As New ADODB.Connection
    Dim DB2 As New ADODB.Connection
    Dim cARQ As String
    Dim ms03 As New ADODB.Recordset
    Dim PFMS03 As New ADODB.Recordset
    Dim cCODIGO, sqlMS03, sqlpfms03 As String
    Dim aCAMPOS As Variant
    cCODIGO = frmPF.txtFields(14)

    If IsNull(cCODIGO) Then

        Alert "Codigo Final em Branco", "Erro Preenchimento"
        Exit Sub

    End If

    If cCODIGO = "" Then

        Alert "Codigo Final em Branco", "Erro Preenchimento"
        Exit Sub

    End If

    cARQ = GeracArq(zMANA5EMP, "JETFOX")
    DB1.ConnectionTimeout = 120
    DB1.Open cARQ
    DB2.ConnectionTimeout = 120
    DB2.Open GeracArq(cARQPF)
    sqlMS03 = "SELECT * FROM MS03 WHERE CODIGO='" & cCODIGO & "' AND TIPOENT='C'"
    
    ms03.Open sqlMS03, DB1, adOpenForwardOnly, adLockReadOnly
    While Not ms03.EOF

        aCAMPOS = Array(ms03!CODCOMP, ms03!QTDDE, ms03!PRECO, ms03!NOMECOMP, ms03!BAIXAC, ms03!BSEQ, ms03!BSSQ, ms03!OPCAO)

        If IsNull(aCAMPOS(7)) Then aCAMPOS(7) = "0"

        sqlpfms03 = "select * from PFMS03 WHERE PF=" & nPF & " AND CODCOMP='" & aCAMPOS(0) & "'"
        
        
        PFMS03.Open sqlpfms03, DB2, adOpenDynamic, adLockOptimistic

        If PFMS03.EOF Then

            PFMS03.AddNew
            PFMS03!PF = nPF
            PFMS03!TIPOENT = "C"
            PFMS03!CODCOMP = aCAMPOS(0)

        Else

            'PFMS03.Edit

        End If

        PFMS03!QTDDE = aCAMPOS(1)
        PFMS03!PRECO = aCAMPOS(2)
        PFMS03!total = PFMS03!QTDDE * PFMS03!PRECO
        PFMS03!DESCRI = aCAMPOS(3)
        PFMS03!BAIXAC = aCAMPOS(4)
        PFMS03!SEQ = aCAMPOS(5)
        PFMS03!SSQ = aCAMPOS(6)
        PFMS03!OPCAO = Val(aCAMPOS(7))
        PFMS03.Update
        ms03.MoveNext

    Wend
    ms03.Close
    PFMS03.Close
    DB1.Close
    DB2.Close

    Set DB1 = Nothing
    Set DB2 = Nothing
    Set ms03 = Nothing
    Set PFMS03 = Nothing
    FilCom
End Sub

Private Sub CmdConfImp_Click()
FrmPrintSetup.Show vbModal, Me
End Sub
Private Sub cmdFOTO_Click(Index As Integer)
    Select Case Index
    Case 0
        zgrp = txtFields(2)
        cARQRTF = PegPath("PATH", "IMGMS01")
        iImage = 2
    Case 1
        zgrp = txtFields(14)
        cARQRTF = PegPath("PATH", "IMGMS01")
        iImage = 2
    Case 2
        zgrp = txtFields(20)
        cARQRTF = PegPath("PATH", "IMGMU01")
        iImage = 7
    Case 3
        zgrp = txtFields(21)
        cARQRTF = PegPath("PATH", "IMGMU01")
        iImage = 7
    Case 4
        zgrp = txtFields(22)
        cARQRTF = PegPath("PATH", "IMGMU01")
        iImage = 7
    Case 5
        GRIDCOM.Col = 0
        Ctipoent = GRIDCOM
        GRIDCOM.Col = 1
        Ccodcomp = GRIDCOM
        If Ctipoent = "C" Then
            zgrp = Ccodcomp
            cARQRTF = PegPath("PATH", "IMGMT01")
            iImage = 8
        Else
            Exit Sub
        End If
    End Select
    If Len(Trim(zgrp)) = 0 Then
        Alert "codigo nao preenchido"
        Exit Sub
    End If
    If Index = 0 Or Index = 1 Then
        If Len(FixStr(txtFields(0))) > 0 Then
            zgrp = zgrp & "|" & txtFields(0)
        End If
    End If
    Load frmIMAGENS
    frmIMAGENS.txtFields(0).Enabled = False
    frmIMAGENS.Escolher(0).Visible = False
    frmIMAGENS.Show vbModal, Me
End Sub

Private Sub CmdImprimir_Click(Index As Integer)
   If Picture1(Index).Height = 0 Then
        Alert ("Sem Imagem")
    Else
        ePASS01 = Index
        PrintPreview1.ShowPreview
    End If
End Sub

Private Sub CmdMotRev_Click()
    If gridrev.Row > 0 Then
        gridrev.Col = 1
        nREV = gridrev
        ComMotAlt.Enabled = True
        ComMotExc.Enabled = True
        ComMotInc.Enabled = True
        Filgridrevi
    End If
End Sub

Private Sub CmdPaste_Click(Index As Integer)
    If Picture1(Index).Height = 0 Then
        Alert ("Sem Imagem")
    Else
        Clipboard.Clear
        Clipboard.SetData Picture1(Index).Image, (vbCFBitmap)

        ' CopyEntirePicture Picture1
    End If

End Sub

Private Sub CmdSaveimg_Click(Index As Integer)
    Dim cARQ As String


    cARQ = "PF" & StrZero(nPF, 6) & "_MP_" & StrZero(Index, 3)
    salvarpict Me, Picture1(Index), cARQ

End Sub

Private Sub CMDSIG_Click(Index As Integer)
    lRETU = False
    eRETU01 = ""
    eRETU02 = ""
    ePASS01 = "isoqsymbol"
    ePASS02 = True
    frmCharacters.Show vbModal, Me
    If lRETU Then
        txtFields(50 + Index) = eRETU01
    End If

End Sub

Private Sub Command1_Click()

    iMU01 = 1
    NewSeq.Show vbModal, Me
    FilRelat

End Sub

Private Sub Command10_Click()
    If nREV = 0 Then
        Alert "Revisão Nao Escolhida"
        Exit Sub
    End If
    ePASS01 = "select * from rev WHERE PF=" & nPF & " AND TIPO='PF' AND REVISAO=" & nREV
    FrmRev.Show vbModal, Me
End Sub

Private Sub Command2_Click()
    'Dim eRETU As Variant
    Dim cdb As String
    Dim sSQL As String
    If GridSeq.Row = 0 Then
        Exit Sub
    End If

    If Not MDG("Apagar Sequencia Fabricação", "Confirme Exclusão") Then
        Exit Sub
    End If

    GridSeq.Col = 1
    nSEQ = GridSeq
    GridSeq.Col = 2
    nSSQ = GridSeq
    
    'PFS
    sSQL = "select * from PFS WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ
    ApagaSQL cARQPF, sSQL
    If zPFTIPO = "PFPRE" Or zPFTIPO = "PFPRO" Then
        ApagaSQL cARQPFP, sSQL
    End If
    If zPFTIPO = "PFGP12" Or zPFTIPO = "PFPRO" Then
        ApagaSQL cARQPFG, sSQL
    End If
    'PFI
    sSQL = "select * from PFI WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ
    ApagaSQL cARQPF, sSQL
    If zPFTIPO = "PFPRE" Or zPFTIPO = "PFPRO" Then
        ApagaSQL cARQPFP, sSQL
    End If
    If zPFTIPO = "PFGP12" Or zPFTIPO = "PFPRO" Then
        ApagaSQL cARQPFG, sSQL
    End If
    'PFC
    sSQL = "select * from PFC WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ
    ApagaSQL cARQPF, sSQL
    If zPFTIPO = "PFPRE" Or zPFTIPO = "PFPRO" Then
        ApagaSQL cARQPFP, sSQL
    End If
    If zPFTIPO = "PFGP12" Or zPFTIPO = "PFPRO" Then
        ApagaSQL cARQPFG, sSQL
    End If
    'PFQSBLEP
    sSQL = "select * from PFQSBLEP WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ
    ApagaSQL cARQPF, sSQL
    If zPFTIPO = "PFPRE" Or zPFTIPO = "PFPRO" Then
        ApagaSQL cARQPFP, sSQL
    End If
    If zPFTIPO = "PFGP12" Or zPFTIPO = "PFPRO" Then
        ApagaSQL cARQPFG, sSQL
    End If
    FilRelat
    
    eRETU02 = "PF:" & TXTPF.Text & Chr(13) & Chr(10)
    eRETU02 = eRETU02 & " Desenho:" & txtFields(2) & Chr(13) & Chr(10)
    eRETU02 = eRETU02 & " Data:" & Fdata(Date) & Chr(13) & Chr(10)
    eRETU02 = eRETU02 & " SEQ=" & nSEQ & " SSQ=" & nSSQ
        
    MAILENV "PFS00002", eRETU02
    

End Sub

Private Sub Command3_Click()
    frmIE.Show vbModal, Me
End Sub

Private Sub Command4_Click()
    GRIDCOM.Col = 0
    Ctipoent = GRIDCOM
    GRIDCOM.Col = 1
    Ccodcomp = GRIDCOM
    
    GRIDCOM.Col = 2
    nSEQ = FixInt(GRIDCOM)

    GRIDCOM.Col = 3
    nSSQ = FixInt(GRIDCOM)
    
    
    frmPFMS03.Show vbModal, Me
    

    FilCom

End Sub

Private Sub Command5_Click()
    Dim sSQL As String
    GRIDCOM.Col = 0
    Ctipoent = GRIDCOM
    GRIDCOM.Col = 1
    Ccodcomp = GRIDCOM

    GRIDCOM.Col = 2
    nSEQ = GRIDCOM

    GRIDCOM.Col = 3
    nSSQ = GRIDCOM

    
    sSQL = "select * from PFMS03 WHERE PF=" & nPF & " AND TIPOENT='" & Ctipoent & "' AND CODCOMP='" & Ccodcomp & "' and seq=" & FixInt(nSEQ) & " and ssq=" & FixInt(nSSQ)
    

    ApagaSQL cARQPF, sSQL
    ApagaSQL cARQPFP, sSQL
    ApagaSQL cARQPFG, sSQL
    
    FilCom

End Sub

Private Sub Command6_Click()
    Ctipoent = "C"
    Ctipoent = InputBox("Tipo (C)omponente (T)ratamento (S)ubProdutos (M)at.Prima", "Inclusão", Ctipoent)
    Ctipoent = UCase$(Left$(Ctipoent, 1))
    If Ctipoent = "C" Or Ctipoent = "T" Or Ctipoent = "S" Or Ctipoent = "M" Then
    Else
        Exit Sub
    End If
    Ccodcomp = InputBox("Digite o Codigo", "Inclusão", "__")
    Ccodcomp = UCase$(Left$(Ccodcomp, 24))
    nSEQ = 0
    nSSQ = 0
    
    nSEQ = FixInt(Val(Busca("Nº SEQ", "Incluindo", CStr(nSEQ), 8)))
    nSSQ = FixInt(Val(Busca("Nº SSQ", "Incluindo", CStr(nSSQ), 8)))
    
    IncluiSQL cARQPF, "SELECT * FROM PFMS03 WHERE PF=" & nPF & " AND TIPOENT='" & Ctipoent & "' AND CODCOMP='" & Ccodcomp & "'", 5, Array("PF", "TIPOENT", "CODCOMP", "SEQ", "SSQ"), _
        Array(nPF, Ctipoent, Ccodcomp, FixInt(nSEQ), FixInt(nSSQ)), True, False
    frmPFMS03.Show vbModal, Me
    FilCom
End Sub

Private Sub Command7_Click(Index As Integer)
    Dim cARQ As String
    Dim sSQL As String
    Dim nNUMERO As Variant
    Dim aRETU As Variant

    
    If Index = 0 Or Index = 1 Then               '0 e 1 ma01 mana5
        nNUMERO = FixInt(txtFields(4))
        cARQ = GeraConn(zMANA5EMP, "JETFOX")
        sSQL = "SELECT NOME,CODIGO FROM MA01 WHERE NUMERO=" & nNUMERO
    Else                                         '10 11  logix clientes/duns
        nNUMERO = Trim(FixStr(txtFields(4)))
        cARQ = PegPath("PATH", "LOGIXODBC")
        sSQL = " SELECT clientes.nom_cliente as nome,duns_itaesbra.num_duns as codigo  FROM clientes"
        sSQL = sSQL & " LEFT JOIN duns_itaesbra ON clientes.cod_cliente=duns_itaesbra.cod_cliente AND duns_itaesbra.cod_empresa='01'"
        sSQL = sSQL & " Where clientes.cod_cliente = '" & nNUMERO & "'"
    End If
    aRETU = PegSQL(cARQ, sSQL, 2, Array("NOME", "CODIGO"), Array("C", "C"), Array("", ""))
    If lRETU Then
        If Index = 0 Or Index = 10 Then          '0 10 nome e codigo cliente
            txtFields(5) = aRETU(0)
            txtFields(53) = aRETU(1)
        Else
            txtFields(53) = aRETU(1)             '1,11 codigo cliente
        End If
    End If
    
    
    
    
    
    
End Sub

Private Sub Command8_Click()
    Dim cARQ As String
    Dim sSQL As String
    Dim cCODIGO As String
    Dim aRETU As Variant
    cARQ = GeraConn(zMANA5EMP, "JETFOX")
    cCODIGO = FixStr(txtFields(2), "", "TRIM")
    sSQL = "SELECT NOME,CODIGOINT FROM MS01 WHERE CODIGO='" & cCODIGO & "'"
    aRETU = PegSQL(cARQ, sSQL, 2, Array("NOME", "CODIGOINT"), Array("C", "C"), Array("", ""))
    If lRETU Then
        txtFields(3) = aRETU(0)
        txtFields(0) = aRETU(1)
    End If
End Sub

Private Sub Command9_Click(Index As Integer)
    Dim cSQL As String
    Dim cNOME As String
    Dim cCODIGO As String
    Dim cARQ As String
    Dim aRETU As Variant

    Select Case Index

    Case 0
        cCODIGO = FixStr(txtFields(20))

    Case 1
        cCODIGO = FixStr(txtFields(21))

    Case 2
        cCODIGO = FixStr(txtFields(22))

    End Select

    cARQ = GeraConn(zMANA5EMP, "JETFOX")
    cSQL = "SELECT NOME,NOM2 FROM mu01 WHERE CODIGO='" & cCODIGO & "'"
    aRETU = PegSQL(cARQ, cSQL, 2, Array("NOME", "NOM2"), Array("C", "C"), Array("", ""))
    If lRETU Then
        cNOME = aRETU(0) + " " + aRETU(1)
    
        Select Case Index
        Case 0
            txtFields(24) = cNOME
        Case 1
            txtFields(25) = cNOME
        Case 2
            txtFields(26) = cNOME
        End Select
    End If
End Sub

Private Sub ComMotAlt_Click()
    If nREV = 0 Then
        Alert "escolha a revisao e clik em motivos"
        Exit Sub
    End If
    If gridrevi.Row > 0 Then
        gridrevi.Col = 3
        nREVI = FixInt(gridrevi)
        lEDITAR = True
        ePASS01 = "select * from revi WHERE PF=" & nPF & " AND TIPO='PF' AND REVISAO=" & nREV & " AND ITEM=" & nREVI
        frmREVI.Show vbModal, Me
        Filgridrevi
    End If
End Sub

Private Sub ComMotExc_Click()
    Dim sSQL As String
    Dim cdb As String
    If nREV = 0 Then
        Alert "escolha a revisao e clik em motivos"
        Exit Sub
    End If
    If gridrevi.Row > 0 Then
        gridrevi.Col = 3
        nREVI = FixInt(gridrevi)
        sSQL = "select * from revi WHERE PF=" & nPF & " AND TIPO='PF'AND REVISAO=" & nREV & " AND  ITEM=" & nREVI
        cdb = Sdb
        If ApagaSQLP(cdb, sSQL) Then
            Filgridrevi
        End If
    End If
End Sub

Private Sub ComMotInc_Click()
    Dim nITEM As Long
    Dim sSQL As String
    If nREV = 0 Then
        Alert "escolha a revisao e clik em motivos"
        Exit Sub
    End If

    sSQL = "revI WHERE PF=" & nPF & " AND TIPO='PF' AND REVISAO=" & nREV
    nITEM = FixInt(PegMAXSQL(Sdb, sSQL, "ITEM", 0)) + 1
    IncluiSQL Sdb, "SELECT * FROM REVI  WHERE PF=" & nPF & " AND TIPO='PF' AND REVISAO=" & nREV & "AND ITEM=" & nITEM, 4, Array("PF", "REVISAO", "TIPO", "ITEM"), _
        Array(nPF, nREV, "PF", nITEM), True, False
    Filgridrevi
End Sub

Private Sub DelImg_Click(Index As Integer)
    Set Picture1(Index).Picture = Nothing
    Set Picture2(Index).Picture = Nothing
    lTROCOU(Index) = True
End Sub

Private Sub DTPicker1_LostFocus()
    If Len(txtFields(66).Text) > 0 Then
        gravacaddes txtFields(6).Text, txtFields(66).Text, DTPicker1.Value
    Else
        gravacaddes txtFields(6).Text, txtFields(7).Text, DTPicker1.Value
    End If
End Sub

Private Sub DTPicker3_LostFocus()
    gravacaddes txtFields(12).Text, txtFields(13).Text, DTPicker3.Value
End Sub

Private Sub gravacaddes(ByVal cDESCLI As Variant, ByVal cREVCLI As Variant, ByVal dDATACLI As Variant)
    Dim cCODLOGIX As String
    Dim nPOS

    If Not MDG("Gravar revisao do cliente no logix") Then
        Exit Sub
    End If
    
    If Len(cREVCLI) = 0 Then
        Alert "revisao nao preenchida"
        Exit Sub
    End If
     
    If Len(cDESCLI) = 0 Then
        Alert "codigo desenho cliente nao preenchido"
        Exit Sub
    End If
     
    If Len(txtFields(0)) = 0 Then
        Alert "codigo interno logix nao preenchido"
        Exit Sub
    End If
     
    If Len(txtFields(3)) = 0 Then
        Alert "descricao do item nao preenchido"
        Exit Sub
    End If
     
    If DataBranco(dDATACLI) Then
        Alert "data revisao cliente nao preenchida"
        Exit Sub
    End If
     
   
    cCODLOGIX = Trim(txtFields(0))
    If Len(cCODLOGIX) > 0 Then
        nPOS = InStr(cCODLOGIX, "/")
        If nPOS > 0 Then
            cCODLOGIX = Mid(cCODLOGIX, 1, nPOS - 1)
        End If
        nPOS = InStr(cCODLOGIX, "_")
        If nPOS > 0 Then
            cCODLOGIX = Mid(cCODLOGIX, 1, nPOS - 1)
        End If
        cCODLOGIX = FixStr(Busca("Codigo Logix", "Codigo Logix", cCODLOGIX, 14), "")
          
        If Len(cCODLOGIX) = 0 Then
            Alert "Codigo logix em branco"
            Exit Sub
        End If
          
          
        cREVCLI = FixStr(Busca("num_revisao", "Num_Revisao", cREVCLI, 4), "")
        If Len(cREVCLI) = 0 Then
            Alert "num_revisao em branco"
            Exit Sub
        End If
          
          
        cDESCLI = FixStr(Busca("Desenho Cliente", "Desenho Cliente", cDESCLI, 50), "") 'padrao logix 15 aumentado base cad_des para 50
        If Len(cDESCLI) = 0 Then
            Alert "Desenho do Cliente"
            Exit Sub
        End If
          
          
        IncluiSQLAdo PegPath("PATH", "LOGIXODBC"), "select * from cad_des where cod_empresa='01' and cod_item='" & cCODLOGIX & "' and num_versao='" & txtFields(7) & "'", 8, _
                                                                                                                                                                   Array("cod_empresa", "cod_item", "num_versao", "seq_versao", "cod_desen", "dat_desen", "den_desen", "cod_form_folha"), _
                                                                                                                                                                   Array("01", cCODLOGIX, cREVCLI, 1, cDESCLI, dDATACLI, txtFields(3).Text, "''"), _
                                                                                                                                                                   True, False

          
    End If
   


End Sub

Private Sub EditSeq_Click()
    If GridSeq.Row = 0 Then
        Exit Sub
    End If
    GridSeq.Col = 1
    nSEQ = GridSeq
    GridSeq.Col = 2
    nSSQ = GridSeq
    frmPFS.Show vbModal
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
        frmPF.txtFields(4) = eRETU01
        frmPF.txtFields(5) = eRETU02
        If Index = 0 Then
            Command7_Click (1)
        End If
        If Index = 1 Then
            Command7_Click (11)
        End If
    End If
End Sub

Private Sub EscCodFlu_Click(Index As Integer)
    lRETU = False
    eRETU01 = ""
    eRETU02 = ""
    escFLX.Show vbModal, Me
    If lRETU Then
        txtFields(50 + Index) = eRETU01
    End If
End Sub

Private Sub esciedx_Click()
    iMU01 = 201
    escIED.Show vbModal, Me
    If lRETU Then
        txtFields(1) = eRETU02
    End If
End Sub

Private Sub escmp04a_Click()
    If MDG("Gravar Elaborador", "Confirme Gravação") Then
        frmPF.txtFields(30) = zIDFOLHA
        frmPF.txtFields(31) = zNOMEFOLHA
        frmPF.txtFields(63) = Date
    End If
End Sub

Private Sub escmu_Click(Index As Integer)
    'mATERIA PRIMA
    iMU01 = 1
    escmu01.Show vbModal, Me
    If Not lRETU Then Exit Sub
    Select Case Index
    Case 1
        frmPF.txtFields(20) = eRETU01
        frmPF.txtFields(24) = eRETU02
    Case 2
        frmPF.txtFields(21) = eRETU01
        frmPF.txtFields(25) = eRETU02
    Case 3
        frmPF.txtFields(22) = eRETU01
        frmPF.txtFields(26) = eRETU02
    End Select
End Sub

Private Sub ESCpro_Click(Index As Integer)
    Dim cCHAVEBUS As String
    ''0=2=6 codigo
    ''4=3=7 integrado
    ''1=5=8 final
    ePASS01 = "MANA5"
    If Index = 2 Or Index = 3 Or Index = 5 Then
        ePASS01 = "LOGIX"
    End If
    If Index = 9 Or Index = 10 Or Index = 11 Then
        ePASS01 = "LOGIM"
    End If

    If Index = 6 Or Index = 7 Or Index = 8 Then
        ePASS01 = "MICRO"
    End If
    
    If lRETU Then
        If Index = 0 Or Index = 2 Or Index = 6 Or Index = 3 Or Index = 4 Or Index = 7 Or Index = 9 Or Index = 10 Or Index = 11 Then
            escms01.Show vbModal, Me
            If lRETU Then
                If Index = 2 Or Index = 0 Or Index = 6 Then
                    frmPF.txtFields(2) = eRETU01
                    frmPF.txtFields(3) = eRETU02
                    frmPF.txtFields(0) = eRETU03
                End If
                If Index = 3 Or Index = 4 Or Index = 7 Then
                    frmPF.txtFields(0) = eRETU03
                End If
                If Index = 9 Then
                    frmPF.txtFields(20) = eRETU01
                    frmPF.txtFields(24) = eRETU02
                End If
                If Index = 10 Then
                    frmPF.txtFields(21) = eRETU01
                    frmPF.txtFields(25) = eRETU02
                End If
                If Index = 11 Then
                    frmPF.txtFields(22) = eRETU01
                    frmPF.txtFields(26) = eRETU02
                End If
               
               
            End If
        Else                                     ''1-5-8
            cCHAVEBUS = txtFields(2)
            escms01.Show vbModal, Me
            If lRETU Then
                frmPF.txtFields(14) = eRETU01
            End If
        End If
    End If
    cCHAVEBUS = ""
End Sub

Private Sub FilCom()
    Dim cARQ As String
    Dim cSQL As String
    DizerBarra "Carregando Componentes"
    cARQ = Sdb
    cSQL = "select TIPOENT,CODCOMP,SEQ,SSQ,DESCRI from PFMS03 WHERE PF=" & nPF & " ORDER BY TIPOENT,CODCOMP,SEQ,SSQ"
    MontaGridFast GRIDCOM, 5, Array(400, 1600, 400, 400, 5000), Array("Tipo", "Codigo", "SEQ", "SSQ", "Nome"), _
        Array("TIPOENT", "L$CODCOMP", "SEQ", "SSQ", "DESCRI"), cARQ, cSQL
    DizerBarra ""
End Sub

Private Sub Filgridrev()
    Dim cARQ As String
    Dim cSQL As String
    Dim cTIPO As String
    cARQ = Sdb
    cTIPO = "PF"
    DizerBarra "Carregando Revisoes"
    cSQL = "select PF,REVISAO,TIPO,DATA from rev WHERE PF=" & nPF & " AND TIPO='" & cTIPO & "' ORDER BY REVISAO"
    MontaGridFast gridrev, 4, Array(400, 400, 400, 1200), Array("PF", "REV", "Tipo", "DATA"), _
        Array("PF", "REVISAO", "TIPO", "C$DATA"), cARQ, cSQL
    DizerBarra ""
End Sub

Private Sub Filgridrevi()
    Dim cARQ As String
    Dim cSQL As String
    DizerBarra "Carregando Dizeres Revisao"
    cARQ = Sdb
    cSQL = "select PF,REVISAO,TIPO,ITEM,MOTIVO from revI WHERE PF=" & nPF & " AND REVISAO=" & nREV & " AND TIPO='PF'ORDER BY ITEM"
    MontaGridFast gridrevi, 5, Array(400, 400, 400, 400, 5000), Array("PF", "Rev", "Tipo", "Item", "Motivo"), _
        Array("PF", "REVISAO", "TIPO", "ITEM", "MOTIVO"), cARQ, cSQL
    DizerBarra ""
End Sub

Private Sub FilRelat()
    Dim cARQ As String
    Dim cSQL As String
    DizerBarra "Carregando Operacoes"
    cARQ = Sdb
    cSQL = "select PF,SEQ,SSQ,DESCRI,CODINT  from PFS WHERE PF=" & nPF & " ORDER BY SEQ,SSQ"
    MontaGridFast GridSeq, 5, Array(400, 400, 400, 5000, 1600), Array("PF", "SEQ", "SSQ", "Descrição", "Cod Int"), _
        Array("PF", "SEQ", "SSQ", "DESCRI", "L$CODINT"), cARQ, cSQL
    DizerBarra ""
End Sub

Private Sub Esctipo_Click()
    If txtFields(28) = "P" Then
        Alert ("Ja e producao")
        Exit Sub
    End If

    If Not AcessaMenu("frmPFTROCA", 1) Then
        Exit Sub
    End If

    ePASS01 = Array("(P)roducao", "p(I)loto", "p(R)ototipo", "pre(L)iminar", "pre l(A)ncamento", "GP12 a(G)ressivo", "GP12 Pla(N)")
    escOrdem.Show vbModal, Me
    If Not lRETU Then
        Exit Sub
    End If
    Select Case eRETU01
    Case 0
        cTIPO = "P"
    Case 1
        cTIPO = "I"
    Case 2
        cTIPO = "R"
    Case 3
        cTIPO = "L"
    Case 4
        cTIPO = "A"
    Case 5
        cTIPO = "G"
    Case 6
        cTIPO = "N"
    End Select
    
    If cTIPO = "P" Then
        If txtFields(28) = "R" Or txtFields(28) = "I" Or txtFields(28) = "L" Or txtFields(28) = "P" _
                                                                                                Or txtFields(28) = "A" Or Len(txtFields(28)) = 0 Then
            If MDG("Mudar para Producao e gerar nova revisao") Then
                NovaRev_Click
            End If
        End If
    End If

    txtFields(28) = cTIPO
    TIPOSAY

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
    Dim X As Integer
    Dim cSQLIMG As String

    nREV = 0
    nREVI = 0
    CenterFormToScreen Me
    frmPF.Caption = "Processo Fabricao " & zPFTIPO
    DizerBarra "Selecionando PF"
    
    cARQPF = PegPath("PATH", "PF")
    cARQPFP = PegPath("PATH", "PFP")
    cARQPFG = PegPath("PATH", "PFG")
    cARQFEMEA = PegPath("PATH", "FEMEA")
    
    For X = 51 To 60
        txtFields(X).Font = "isoqsymbol"
    Next
    
    TXTPF = nPF
    
    lTROCOU = Array(False, False, False)
    aCAMIMG = Array("FIGMU011", "FIGMU012", "FIGMU013")
    
    cARQPF = PegPath("PATH", "PF")
    
    cSQL = "select * from PF WHERE PF=" & nPF
    
    
    nCAMPOS = 72                                 '71
    aCAM = Array("CODIGOINT", "CPF", "CODIGO", "DESCR", "CLIENTE", "CLINOME", "CLIDES", "CLIREV", "REVPRO", "NOSDES", _
                 "NOSREV", "BLOQMOTIVO", "CONDES", "CONREV", "CODFINAL", "CONPES", "CONQTD", "PESLIQ", "MOTREV", "PESMU011", _
                 "CODMU011", "CODMU012", "CODMU013", "PASMU011", "NOMMU011", "NOMMU012", "NOMMU013", "QTDMU011", "TIPO", "REVDAT", _
                 "ELANUM", "ELANOM", "FILIAL", "OPCAO", "CODCLIENTE", "PERMU011", "OBSMU011", "PASMU012", "QTDMU012", "PESMU012", _
                 "PERMU012", "OBSMU012", "PASMU013", "QTDMU013", "PESMU013", "PERMU013", "OBSMU013", "SEQMU011", "SSQMU011", "SEQMU012", _
                 "SSQMU012", "FLX01", "FLX02", "FLX03", "FLX04", "FLX05", "FLX06", "FLX07", "FLX08", "FLX09", _
                 "FLX10", "SEQMU013", "SSQMU013", "ELADAT", "TAKT", "TAKTAT", "numversaocli", "SEGURANCA", "BLOQUEADO", "CLIDAT", "NOSDAT", _
                 "CONDAT")
    aFOR = Array("C", "C", "C", "C", "NI", "C", "C", "C", "NI", "C", _
                 "C", "C", "C", "C", "C", "N", "NI", "N", "C", "N", _
                 "C", "C", "C", "N", "C", "C", "C", "N", "C", "D", _
                 "NI", "C", "C", "NI", "C", "NI", "C", "N", "NI", "NI", _
                 "NI", "C", "N", "NI", "N", "NI", "C", "NI", "NI", "NI", _
                 "NI", "C", "C", "C", "C", "C", "C", "C", "C", "C", _
                 "C", "NI", "NI", "D", "N", "N", "C", "B", "B", "DN", "DN", _
                 "DN")
    aPAD = Array("", "", "", "", 0, "", "", "", 0, "", _
                 "", "", "", "", "", 0, 0, 0, "", 0, _
                 "", "", "", 0, "", "", "", 0, "", "", _
                 0, "", "", 0, "", 0, "C", 0, 0, 0, _
                 0, "", 0, 0, 0, 0, "", 0, 0, 0, _
                 0, "", "", "", "", "", "", "", "", "", _
                 "", 0, 0, "", 0, 0, "", "", "", "", "", _
                 "")
    
    aVAL = PegSQL(cARQPF, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
    For iLOOP = 0 To nCAMPOS - 6                 '
        txtFields(iLOOP) = aVAL(iLOOP)           '
    Next iLOOP
    
    
    Check2.Value = FixBolNum(aVAL(67))           'seguranca 66
    Check1.Value = FixBolNum(aVAL(68))           'bloqueado 67
    If IsDate(aVAL(69)) Then
        DTPicker1.Value = aVAL(69)               'clidat 68
    End If
    If IsDate(aVAL(70)) Then
        DTPicker2.Value = aVAL(70)               'nosdat 69
    End If
    If IsDate(aVAL(71)) Then
        DTPicker3.Value = aVAL(71)               'condat 70
    End If
    
    
    
    For iLOOP = 0 To 2
        cSQLIMG = "select PF," & aCAMIMG(iLOOP) & " from PF WHERE PF=" & nPF
        If ADOPegBlob(cARQPF, cSQLIMG, Picture1(iLOOP), aCAMIMG(iLOOP)) Then
            StretchSourcePictureFromPicture Picture1(iLOOP), Picture2(iLOOP)
            If FixNum(eRETU01) > 500000 Then
                Alert ("Imagem Muito Grande,Ajuste o tamanho")
                salvarpict Me, Picture1(iLOOP), aCAMIMG(iLOOP) & "_PF_" & StrZero(nPF, 6)
                Set Picture1(iLOOP).Picture = Nothing
                Set Picture2(iLOOP).Picture = Nothing
                lTROCOU(iLOOP) = True
            End If
        Else
            Set Picture1(iLOOP).Picture = Nothing
            Set Picture2(iLOOP).Picture = Nothing
        End If
    Next iLOOP

    DizerBarra "Carregando PF"
    

    lOLDBLD = Check1.Value
    
    
    FilRelat
    FilCom
    Filgridrev
    
    If zUSER = "ADMIN" Then
        txtFields(8).Enabled = True
        txtFields(8).Locked = False
        txtFields(8).Visible = True
        txtFields(29).Enabled = True
        txtFields(29).Locked = False
        txtFields(29).Visible = True
    End If
    
    DizerBarra ""
    
    If Check1 And Len(txtFields(11)) = 0 Then
        Alert ("Preencher Motivo do Bloqueio")
    End If
    
    TIPOSAY
    
End Sub

Private Sub ImpComPF_Click()
    cARQIMP = "PFMS03"
    FrmImp.Show vbModal, Me
    FilCom
End Sub

Private Sub IMPSEQ_Click()
    If GridSeq.Rows > 1 Then                     'importacao nao precisa ter um item
        Alert ("Ja existem operacoes para este pf utilize a importacao dentro da sequencias")
        Exit Sub
    End If
    cARQIMP = "PFS"
    FrmImp.Show vbModal, Me
    FilRelat
End Sub

Private Sub IncImg_Click(Index As Integer)
    Dim STMPFILE
    STMPFILE = OpenArqExt(Me, "", "JPG", "JPEG *.JPG")
    If Len(STMPFILE) > 0 Then
        If FixInt(FileLen(STMPFILE)) > 500000 Then
            Alert ("Imagem Muito Grande,Ajuste o tamanho")
            If Not MDG("Anexar mesmo assim-NAO RECOMENDADO") Then
                Exit Sub
            Else
                Exit Sub                         'nao permitindo aumentando o banco e travando relatorio crystal
            End If
        End If
        Picture1(Index).Picture = LoadPicture(STMPFILE)
        StretchSourcePictureFromPicture Picture1(Index), Picture2(Index)
        lTROCOU(Index) = True
    End If

End Sub

Private Sub NovaRev_Click()
    Dim nrevisao As Long
    Dim cSQLTMP As String
    Dim aCAMTMP As Variant
    Dim aVALTMP As Variant
    Dim aFORTMP As Variant

    On Error Resume Next
    
    If Len(txtFields(29).Text) > 0 Then          ''corrige data protheus em branco
        If txtFields(29) > Today() Then
            Alert ("Data Revisao Anterior Maior que Data do Sistema")
            Exit Sub
        End If
    End If

    If MDG("Fazer Nova Revisao", "Confirme Gravação") Then
        Encerrar.Visible = False
        nrevisao = FixInt(txtFields(8)) + 1
        txtFields(8) = nrevisao
        txtFields(29) = Fdata(Date)
        
        
        cSQLTMP = "select * from PF WHERE PF=" & nPF
        aCAMTMP = Array("REVPRO", "RESNUM", "FEMEAF", "RESNOM", "FEMEAR", "REVDAT", "RESDAT", _
                        "FEMEAD", "PCLIBDATE", "STPFFE", "STPFPC", "PCLIBNOM", "PCLIBNUM")
        aVALTMP = Array(nrevisao, 0, 0, "", "", Date, "", "", "", True, True, "", 0)
        aFORTMP = Array("N", "N", "N", "C", "C", "D", "D", "D", "D", "B", "B", "C", "N")
        
        
        
        GrvSQL Sdb, cSQLTMP, 13, aCAMTMP, aVALTMP, aFORTMP
        
        
        ' CDate(Null)
        
        cSQLTMP = "select * from rev WHERE PF=" & nPF & " AND REVISAO=" & nrevisao & " and tipo='PF'"
        
        IncluiSQL Sdb, cSQLTMP, 4, Array("PF", "REVISAO", "TIPO", "DATA"), _
        Array(nPF, nrevisao, "PF", Date), True, False
                             
                             
        cSQLTMP = "select * from PFrev WHERE PF=" & nPF & " AND REV=" & nrevisao
        IncluiSQL cARQPF, cSQLTMP, 3, Array("PF", "REV", "PFDATA"), _
        Array(nPF, nrevisao, Date), True, False
                             
        Filgridrev
        

        eRETU02 = "PF:" & TXTPF.Text & Chr(13) & Chr(10)
        eRETU02 = eRETU02 & " Desenho:" & txtFields(2) & Chr(13) & Chr(10)
        eRETU02 = eRETU02 & " Data:" & Fdata(Date) & Chr(13) & Chr(10)
        
        MAILENV "PF000001", eRETU02
       

        
        escmp04a_Click
        cmdClose_Click

    End If
End Sub

Private Sub seqreo_Click()
    If GridSeq.Row = 0 Then
        Exit Sub
    End If
    GridSeq.Col = 1
    nSEQ = GridSeq
    GridSeq.Col = 2
    nSSQ = GridSeq
    iMU01 = 2
    NewSeq.Show vbModal, Me
    FilRelat
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim KeyPressed As Variant
    Select Case Index
    Case 4, 16, 30, 33, 47, 48, 49, 50, 61, 62, 27, 35, 38, 40, 43, 45
        KeyAscii = ValiText(KeyAscii, "#NI")
    Case 15, 17, 19, 23, 37, 39, 42, 44, 64, 65
        KeyAscii = ValiText(KeyAscii, "#N")
    Case 28
        KeyPressed = ValiText(KeyAscii, "PRIL", , "#CU")
    End Select
End Sub

Function valornum(Index As Integer)
    valornum = Val(Replace(txtFields(Index), ",", "."))
End Function

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

Private Sub TXTFIELDS_LostFocus(Index As Integer)
    Dim cSQLITEM
    Dim cCODLOGIX
    Dim nPOS
    If Index = 0 Then
        cCODLOGIX = Trim(txtFields(0))
        If Len(cCODLOGIX) > 0 Then
            nPOS = InStr(cCODLOGIX, "/")
            If nPOS > 0 Then
                cCODLOGIX = Mid(cCODLOGIX, 1, nPOS - 1)
            End If
            nPOS = InStr(cCODLOGIX, "_")
            If nPOS > 0 Then
                cCODLOGIX = Mid(cCODLOGIX, 1, nPOS - 1)
            End If
            cCODLOGIX = FixStr(Busca("Codigo Logix", "Codigo Logix", cCODLOGIX, 14), "")
            If Len(cCODLOGIX) > 0 Then
                cSQLITEM = "update item_funcional_970 set pf= " & nPF & " where cod_item_logix='" & cCODLOGIX & "' ;"
                ADOComando PegPath("PATH", "LOGIXODBC"), cSQLITEM
                cSQLITEM = "update ciclo_peca_970 set pf= " & nPF & " where cod_item='" & cCODLOGIX & "' ;"
                ADOComando PegPath("PATH", "LOGIXODBC"), cSQLITEM
            End If
        End If
    End If

End Sub

Private Sub TIPOSAY()
    Select Case txtFields(28)
    Case "P"
        TipoDescritivo.Text = "(P)roducao"
    Case "R"
        TipoDescritivo.Text = "p(R)ototipo"
    Case "I"
        TipoDescritivo.Text = "p(I)loto"
    Case "A"
        TipoDescritivo.Text = "pre l(A)ncamento"
    Case "L"
        TipoDescritivo.Text = "pre (L)iminar"
    Case "G"
        TipoDescritivo.Text = "GP12 a(G)ressivo"
    Case "N"
        TipoDescritivo.Text = "GP12 pla(N)"
    End Select
End Sub
Private Sub VerImg_Click(Index As Integer)
    frmPicViewer.Show vbModal, Me
    If lRETU Then
        If lerarquivoimagem(eRETU01, Picture1(Index), Picture2(Index)) Then
            lTROCOU(Index) = True
        End If
    End If
End Sub
Public Sub PrintPreview1_PrepareReport(Cancel As Boolean)
    On Error Resume Next
    Printer.Print
    Printer.PaintPicture Picture1(ePASS01), 0, 0
    Printer.EndDoc
End Sub
