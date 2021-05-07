VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.0#0"; "vbccr17.ocx"
Object = "{EA478B61-D9EC-47F6-BB21-95A533AF2251}#1.0#0"; "TabExC01.ocx"
Begin VB.Form frmPPAP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PPAP"
   ClientHeight    =   6360
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   12015
   Icon            =   "Frmppap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   12015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton EscMA01A 
      Caption         =   "LX2"
      Height          =   315
      Index           =   1
      Left            =   3000
      TabIndex        =   124
      TabStop         =   0   'False
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton EscMA01A 
      Caption         =   "MC"
      Height          =   315
      Index           =   6
      Left            =   3600
      TabIndex        =   122
      TabStop         =   0   'False
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton EscMA01A 
      Caption         =   "LX"
      Height          =   315
      Index           =   2
      Left            =   2520
      TabIndex        =   121
      TabStop         =   0   'False
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton ESCMS01A 
      Caption         =   "MC"
      Height          =   255
      Index           =   2
      Left            =   3540
      TabIndex        =   8
      Top             =   420
      Width           =   435
   End
   Begin VB.CommandButton ESCMS01A 
      Caption         =   "LX"
      Height          =   255
      Index           =   1
      Left            =   3180
      TabIndex        =   120
      Top             =   420
      Width           =   315
   End
   Begin VB.TextBox Txtfields 
      DataField       =   "PF"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   32
      Left            =   4440
      TabIndex        =   114
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox Txtdispo 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   8040
      Locked          =   -1  'True
      TabIndex        =   107
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.CheckBox check1 
      Caption         =   "Inativo"
      DataField       =   "INATIVO"
      DataSource      =   "datPrimaryRS"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   21
      Left            =   5760
      TabIndex        =   106
      Top             =   120
      Width           =   975
   End
   Begin TabExCtl.SSTabEx SSTab1 
      Height          =   4635
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   8176
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
      TabHeight       =   520
      TabCaption(0)   =   "A"
      Tab(0).ControlCount=   34
      Tab(0).Control(0)=   "Command4"
      Tab(0).Control(1)=   "Txtfields(20)"
      Tab(0).Control(2)=   "Txtfields(19)"
      Tab(0).Control(3)=   "check1(18)"
      Tab(0).Control(4)=   "check1(20)"
      Tab(0).Control(5)=   "check1(19)"
      Tab(0).Control(6)=   "check1(16)"
      Tab(0).Control(7)=   "Txtfields(15)"
      Tab(0).Control(8)=   "Txtfields(12)"
      Tab(0).Control(9)=   "Txtfields(9)"
      Tab(0).Control(10)=   "check1(17)"
      Tab(0).Control(11)=   "Txtfields(10)"
      Tab(0).Control(12)=   "check1(14)"
      Tab(0).Control(13)=   "ESCMC02A"
      Tab(0).Control(14)=   "Txtfields(11)"
      Tab(0).Control(15)=   "Txtfields(7)"
      Tab(0).Control(16)=   "Txtfields(4)"
      Tab(0).Control(17)=   "Txtfields(3)"
      Tab(0).Control(18)=   "Txtfields(16)"
      Tab(0).Control(19)=   "DTPicker2"
      Tab(0).Control(20)=   "DTPicker3"
      Tab(0).Control(21)=   "lblLabels(16)"
      Tab(0).Control(22)=   "lblLabels(10)"
      Tab(0).Control(23)=   "lblLabels(14)"
      Tab(0).Control(24)=   "lblLabels(12)"
      Tab(0).Control(25)=   "lblLabels(8)"
      Tab(0).Control(26)=   "lblLabels(6)"
      Tab(0).Control(27)=   "Label11"
      Tab(0).Control(28)=   "lblLabels(24)"
      Tab(0).Control(29)=   "lblLabels(11)"
      Tab(0).Control(30)=   "lblLabels(9)"
      Tab(0).Control(31)=   "lblLabels(7)"
      Tab(0).Control(32)=   "lblLabels(3)"
      Tab(0).Control(33)=   "lblLabels(31)"
      TabCaption(1)   =   "B"
      Tab(1).ControlCount=   20
      Tab(1).Control(0)=   "Txtfields(14)"
      Tab(1).Control(1)=   "Txtfields(13)"
      Tab(1).Control(2)=   "check1(0)"
      Tab(1).Control(3)=   "check1(1)"
      Tab(1).Control(4)=   "check1(2)"
      Tab(1).Control(5)=   "check1(3)"
      Tab(1).Control(6)=   "check1(4)"
      Tab(1).Control(7)=   "check1(5)"
      Tab(1).Control(8)=   "check1(6)"
      Tab(1).Control(9)=   "check1(7)"
      Tab(1).Control(10)=   "check1(8)"
      Tab(1).Control(11)=   "check1(9)"
      Tab(1).Control(12)=   "check1(10)"
      Tab(1).Control(13)=   "check1(11)"
      Tab(1).Control(14)=   "check1(12)"
      Tab(1).Control(15)=   "check1(13)"
      Tab(1).Control(16)=   "check1(15)"
      Tab(1).Control(17)=   "Txtfields(29)"
      Tab(1).Control(18)=   "lblLabels(13)"
      Tab(1).Control(19)=   "Label2"
      TabCaption(2)   =   "C"
      Tab(2).ControlCount=   14
      Tab(2).Control(0)=   "Txtfields(30)"
      Tab(2).Control(1)=   "Txtfields(31)"
      Tab(2).Control(2)=   "Txtfields(8)"
      Tab(2).Control(3)=   "Txtfields(0)"
      Tab(2).Control(4)=   "Txtfields(17)"
      Tab(2).Control(5)=   "Txtfields(18)"
      Tab(2).Control(6)=   "Command2"
      Tab(2).Control(7)=   "DTPicker1"
      Tab(2).Control(8)=   "lblLabels(23)"
      Tab(2).Control(9)=   "lblLabels(32)"
      Tab(2).Control(10)=   "lblLabels(4)"
      Tab(2).Control(11)=   "Label1"
      Tab(2).Control(12)=   "lblLabels(2)"
      Tab(2).Control(13)=   "lblLabels(15)"
      TabCaption(3)   =   "Ctrle Datas"
      Tab(3).ControlCount=   4
      Tab(3).Control(0)=   "Grid"
      Tab(3).Control(1)=   "CmdNovo(0)"
      Tab(3).Control(2)=   "CmdExcluir(0)"
      Tab(3).Control(3)=   "CmdEditar(0)"
      TabCaption(4)   =   "Fornecedor"
      Tab(4).ControlCount=   25
      Tab(4).Control(0)=   "lblLabels(17)"
      Tab(4).Control(1)=   "lblLabels(18)"
      Tab(4).Control(2)=   "lblLabels(19)"
      Tab(4).Control(3)=   "Cmdmw02"
      Tab(4).Control(4)=   "Command6"
      Tab(4).Control(5)=   "Command7"
      Tab(4).Control(6)=   "Txtfields(21)"
      Tab(4).Control(7)=   "Txtfields(22)"
      Tab(4).Control(8)=   "Txtfields(23)"
      Tab(4).Control(9)=   "Command8(0)"
      Tab(4).Control(10)=   "Command8(1)"
      Tab(4).Control(11)=   "Command8(2)"
      Tab(4).Control(12)=   "Command9(0)"
      Tab(4).Control(13)=   "Command10"
      Tab(4).Control(14)=   "Txtfields(24)"
      Tab(4).Control(15)=   "Txtfields(25)"
      Tab(4).Control(16)=   "Txtfields(26)"
      Tab(4).Control(17)=   "Txtfields(27)"
      Tab(4).Control(18)=   "Txtfields(28)"
      Tab(4).Control(19)=   "Cmdprg"
      Tab(4).Control(20)=   "Cmdped"
      Tab(4).Control(21)=   "cmdcopcod"
      Tab(4).Control(22)=   "Command11"
      Tab(4).Control(23)=   "Command12"
      Tab(4).Control(24)=   "Command9(1)"
      TabCaption(5)   =   "Cron. Layout"
      Tab(5).ControlCount=   5
      Tab(5).Control(0)=   "gridlay"
      Tab(5).Control(1)=   "CmdEdilay(1)"
      Tab(5).Control(2)=   "CmdExcLay(1)"
      Tab(5).Control(3)=   "CmdNovoLay(1)"
      Tab(5).Control(4)=   "CmdNovoLay(0)"
      TabCaption(6)   =   "Cjto"
      Tab(6).ControlCount=   5
      Tab(6).Control(0)=   "GridCjto"
      Tab(6).Control(1)=   "CmdNovCjto(2)"
      Tab(6).Control(2)=   "CmdExcjto(2)"
      Tab(6).Control(3)=   "Cmdimppf(0)"
      Tab(6).Control(4)=   "Cmdcjtcrn(0)"
      Begin VB.CommandButton Command9 
         Height          =   315
         Index           =   1
         Left            =   -69600
         Picture         =   "Frmppap.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   1200
         Width           =   495
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Localizar Pelo PPAP"
         Height          =   315
         Left            =   -71940
         TabIndex        =   119
         ToolTipText     =   "Preenche programa/pedido/item com base fornecedor/codigo"
         Top             =   2220
         Width           =   1755
      End
      Begin VB.CommandButton Command11 
         Caption         =   "^^"
         Height          =   375
         Left            =   -71880
         TabIndex        =   118
         ToolTipText     =   "copia para codigo e nome "
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton cmdcopcod 
         Caption         =   "^^"
         Height          =   375
         Left            =   -70080
         TabIndex        =   117
         ToolTipText     =   "copia para codigo e nome "
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton CmdNovoLay 
         Caption         =   "Nova"
         Height          =   255
         Index           =   0
         Left            =   -66600
         TabIndex        =   116
         Top             =   1680
         Width           =   795
      End
      Begin VB.CommandButton Cmdcjtcrn 
         Caption         =   "Incluir Cronograma"
         Height          =   615
         Index           =   0
         Left            =   -66840
         TabIndex        =   115
         Top             =   1680
         Width           =   1155
      End
      Begin VB.CommandButton Cmdimppf 
         Caption         =   "Imp PF "
         Height          =   255
         Index           =   0
         Left            =   -66480
         TabIndex        =   112
         Top             =   1320
         Width           =   795
      End
      Begin VB.CommandButton CmdExcjto 
         Caption         =   "Excluir"
         Height          =   255
         Index           =   2
         Left            =   -66480
         TabIndex        =   110
         Top             =   960
         Width           =   795
      End
      Begin VB.CommandButton CmdNovCjto 
         Caption         =   "Incluir"
         Height          =   255
         Index           =   2
         Left            =   -66480
         TabIndex        =   109
         Top             =   600
         Width           =   795
      End
      Begin VB.CommandButton CmdNovoLay 
         Caption         =   "Incluir"
         Height          =   255
         Index           =   1
         Left            =   -66600
         TabIndex        =   105
         Top             =   600
         Width           =   795
      End
      Begin VB.CommandButton CmdExcLay 
         Caption         =   "Excluir"
         Height          =   255
         Index           =   1
         Left            =   -66600
         TabIndex        =   104
         Top             =   960
         Width           =   795
      End
      Begin VB.CommandButton CmdEdilay 
         Caption         =   "Editar"
         Height          =   255
         Index           =   1
         Left            =   -66600
         TabIndex        =   103
         Top             =   1320
         Width           =   795
      End
      Begin VB.CommandButton Cmdped 
         Caption         =   "Prg>"
         Height          =   375
         Left            =   -73500
         TabIndex        =   101
         ToolTipText     =   "Preenche Pedido/Item com base no programa"
         Top             =   2760
         Width           =   495
      End
      Begin VB.CommandButton Cmdprg 
         Caption         =   "Localizar Pelo Produto Fornecedor"
         Height          =   315
         Left            =   -74700
         TabIndex        =   100
         ToolTipText     =   "Preenche programa/pedido/item com base fornecedor/codigo"
         Top             =   2220
         Width           =   2715
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "COMITEM"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   28
         Left            =   -71820
         TabIndex        =   99
         Top             =   2820
         Width           =   1095
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "COMPRAS"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   27
         Left            =   -72960
         TabIndex        =   98
         Top             =   2820
         Width           =   1095
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "PRGENT"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   26
         Left            =   -74640
         TabIndex        =   97
         Top             =   2820
         Width           =   1095
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "NOMCOM"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   25
         Left            =   -72960
         MaxLength       =   40
         TabIndex        =   94
         Top             =   1740
         Width           =   5115
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "CODCOM"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   24
         Left            =   -72960
         MaxLength       =   24
         TabIndex        =   93
         Top             =   1260
         Width           =   1575
      End
      Begin VB.CommandButton Command10 
         Caption         =   "-->"
         Height          =   255
         Left            =   -70680
         TabIndex        =   92
         Top             =   1260
         Width           =   495
      End
      Begin VB.CommandButton Command9 
         Height          =   315
         Index           =   0
         Left            =   -71280
         Picture         =   "Frmppap.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   1260
         Width           =   495
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Terceiros"
         Height          =   255
         Index           =   2
         Left            =   -74160
         TabIndex        =   90
         Top             =   1860
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Componente"
         Height          =   255
         Index           =   1
         Left            =   -74160
         TabIndex        =   89
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Mat.Prima"
         Height          =   255
         Index           =   0
         Left            =   -74160
         TabIndex        =   88
         Top             =   1260
         Width           =   1095
      End
      Begin VB.TextBox Txtfields 
         BackColor       =   &H00C0FFFF&
         DataField       =   "TIPCOM"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   23
         Left            =   -74580
         Locked          =   -1  'True
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   1380
         Width           =   255
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "FORNECEDOR"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   22
         Left            =   -73680
         TabIndex        =   85
         Top             =   780
         Width           =   1095
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "FORNOM"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   21
         Left            =   -71400
         MaxLength       =   50
         TabIndex        =   84
         Top             =   780
         Width           =   5715
      End
      Begin VB.CommandButton Command7 
         Height          =   315
         Left            =   -72480
         Picture         =   "Frmppap.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command6 
         Caption         =   "-->"
         Height          =   255
         Left            =   -72480
         TabIndex        =   82
         Top             =   900
         Width           =   495
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "Editar"
         Height          =   255
         Index           =   0
         Left            =   -66660
         TabIndex        =   81
         Top             =   1320
         Width           =   795
      End
      Begin VB.CommandButton CmdExcluir 
         Caption         =   "Excluir"
         Height          =   255
         Index           =   0
         Left            =   -66660
         TabIndex        =   80
         Top             =   960
         Width           =   795
      End
      Begin VB.CommandButton CmdNovo 
         Caption         =   "Incluir"
         Height          =   255
         Index           =   0
         Left            =   -66660
         TabIndex        =   79
         Top             =   600
         Width           =   795
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   2655
         Left            =   -74760
         TabIndex        =   78
         Top             =   480
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   4683
         _Version        =   393216
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "EXP01"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   30
         Left            =   -72840
         MaxLength       =   150
         TabIndex        =   71
         Top             =   1380
         Width           =   7095
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "EXP02"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   31
         Left            =   -72840
         MaxLength       =   150
         TabIndex        =   70
         Top             =   1680
         Width           =   7095
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "DISPO"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   8
         Left            =   -72540
         TabIndex        =   69
         Top             =   1020
         Width           =   315
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "APRO"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Index           =   0
         Left            =   -69960
         TabIndex        =   68
         Top             =   960
         Width           =   435
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "RESNOME"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   17
         Left            =   -73740
         MaxLength       =   24
         TabIndex        =   67
         Top             =   660
         Width           =   2115
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "RESCARGO"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   18
         Left            =   -70800
         MaxLength       =   24
         TabIndex        =   66
         Top             =   660
         Width           =   2115
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Left            =   -68520
         Picture         =   "Frmppap.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   660
         Width           =   495
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "OBSADC02"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   14
         Left            =   -73140
         MaxLength       =   80
         TabIndex        =   64
         Top             =   2940
         Width           =   7335
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "OBSADC01"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   13
         Left            =   -73140
         MaxLength       =   80
         TabIndex        =   63
         Top             =   2640
         Width           =   7335
      End
      Begin VB.CheckBox check1 
         Caption         =   "Submissão Inicial"
         DataField       =   "SUB01"
         DataSource      =   "datPrimaryRS"
         Height          =   255
         Index           =   0
         Left            =   -74820
         TabIndex        =   60
         Top             =   720
         Width           =   2595
      End
      Begin VB.CheckBox check1 
         Caption         =   "Alterações de Engenharia"
         DataField       =   "SUB02"
         DataSource      =   "datPrimaryRS"
         Height          =   255
         Index           =   1
         Left            =   -74820
         TabIndex        =   59
         Top             =   1020
         Width           =   2595
      End
      Begin VB.CheckBox check1 
         Caption         =   "Ferramental, Trans. Rep.Ref."
         DataField       =   "SUB03"
         DataSource      =   "datPrimaryRS"
         Height          =   255
         Index           =   2
         Left            =   -74820
         TabIndex        =   58
         Top             =   1320
         Width           =   2595
      End
      Begin VB.CheckBox check1 
         Caption         =   "Correção de Discrepancia"
         DataField       =   "SUB04"
         DataSource      =   "datPrimaryRS"
         Height          =   255
         Index           =   3
         Left            =   -74820
         TabIndex        =   57
         Top             =   1620
         Width           =   2475
      End
      Begin VB.CheckBox check1 
         Caption         =   "Ferra. Inativo >1ano"
         DataField       =   "SUB05"
         DataSource      =   "datPrimaryRS"
         Height          =   255
         Index           =   4
         Left            =   -74820
         TabIndex        =   56
         Top             =   1920
         Width           =   2595
      End
      Begin VB.CheckBox check1 
         Caption         =   "Medições Tridimensionais"
         DataField       =   "RES01"
         DataSource      =   "datPrimaryRS"
         Height          =   255
         Index           =   5
         Left            =   -68820
         TabIndex        =   55
         Top             =   720
         Width           =   2955
      End
      Begin VB.CheckBox check1 
         Caption         =   "Critérios de Aparencia"
         DataField       =   "RES03"
         DataSource      =   "datPrimaryRS"
         Height          =   255
         Index           =   6
         Left            =   -68820
         TabIndex        =   54
         Top             =   1320
         Width           =   2955
      End
      Begin VB.CheckBox check1 
         Caption         =   "Material/Construção Opcional"
         DataField       =   "SUB06"
         DataSource      =   "datPrimaryRS"
         Height          =   255
         Index           =   7
         Left            =   -72120
         TabIndex        =   53
         Top             =   720
         Width           =   3075
      End
      Begin VB.CheckBox check1 
         Caption         =   "Sub-Forncedor ou Mud.Fonte Mat."
         DataField       =   "SUB07"
         DataSource      =   "datPrimaryRS"
         Height          =   255
         Index           =   8
         Left            =   -72120
         TabIndex        =   52
         Top             =   1020
         Width           =   3075
      End
      Begin VB.CheckBox check1 
         Caption         =   "Mudança no Processo da Peça"
         DataField       =   "SUB08"
         DataSource      =   "datPrimaryRS"
         Height          =   255
         Index           =   9
         Left            =   -72120
         TabIndex        =   51
         Top             =   1260
         Width           =   3075
      End
      Begin VB.CheckBox check1 
         Caption         =   "Peças Produzidas em outra localidade"
         DataField       =   "SUB09"
         DataSource      =   "datPrimaryRS"
         Height          =   255
         Index           =   10
         Left            =   -72120
         TabIndex        =   50
         Top             =   1500
         Width           =   3075
      End
      Begin VB.CheckBox check1 
         Caption         =   "Materias e testes Funcionais"
         DataField       =   "RES02"
         DataSource      =   "datPrimaryRS"
         Height          =   255
         Index           =   11
         Left            =   -68820
         TabIndex        =   49
         Top             =   1020
         Width           =   2955
      End
      Begin VB.CheckBox check1 
         Caption         =   "Dados Estatisticos"
         DataField       =   "RES04"
         DataSource      =   "datPrimaryRS"
         Height          =   255
         Index           =   12
         Left            =   -68820
         TabIndex        =   48
         Top             =   1620
         Width           =   2955
      End
      Begin VB.CheckBox check1 
         Caption         =   "Atendem todos os requisitos de desenhos e especificações"
         DataField       =   "APLIC"
         DataSource      =   "datPrimaryRS"
         Height          =   255
         Index           =   13
         Left            =   -74820
         TabIndex        =   47
         Top             =   2280
         Width           =   4635
      End
      Begin VB.CheckBox check1 
         Caption         =   "Outros Especifique"
         DataField       =   "SUB10"
         DataSource      =   "datPrimaryRS"
         Height          =   255
         Index           =   15
         Left            =   -72120
         TabIndex        =   46
         Top             =   1800
         Width           =   3075
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "molde"
         DataSource      =   "datPrimaryRS"
         Height          =   315
         Index           =   29
         Left            =   -69360
         TabIndex        =   45
         Top             =   2220
         Width           =   3195
      End
      Begin VB.CommandButton Command4 
         Caption         =   "-->"
         Height          =   255
         Left            =   2340
         TabIndex        =   31
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "DIVISAO"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   20
         Left            =   4980
         MaxLength       =   24
         TabIndex        =   30
         Top             =   480
         Width           =   2115
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "PEDOM"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   19
         Left            =   1200
         MaxLength       =   24
         TabIndex        =   29
         Top             =   480
         Width           =   2115
      End
      Begin VB.CheckBox check1 
         Caption         =   "Marcação ISO"
         DataField       =   "NOT02"
         DataSource      =   "datPrimaryRS"
         Height          =   255
         Index           =   18
         Left            =   4320
         TabIndex        =   28
         Top             =   2580
         Width           =   2055
      End
      Begin VB.CheckBox check1 
         Caption         =   "Material/Funcional"
         DataField       =   "INF02"
         DataSource      =   "datPrimaryRS"
         Height          =   255
         Index           =   20
         Left            =   6540
         TabIndex        =   27
         Top             =   2580
         Width           =   2055
      End
      Begin VB.CheckBox check1 
         Caption         =   "Dimensional"
         DataField       =   "INF01"
         DataSource      =   "datPrimaryRS"
         Height          =   255
         Index           =   19
         Left            =   6540
         TabIndex        =   26
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CheckBox check1 
         Caption         =   "Aparência"
         DataField       =   "INF03"
         DataSource      =   "datPrimaryRS"
         Height          =   255
         Index           =   16
         Left            =   6540
         TabIndex        =   25
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "AUXILIAR"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   15
         Left            =   3720
         MaxLength       =   24
         TabIndex        =   24
         Top             =   1860
         Width           =   2115
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "EXPOSTO"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   12
         Left            =   780
         MaxLength       =   24
         TabIndex        =   23
         Top             =   1860
         Width           =   2115
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "NIVELALT"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   9
         Left            =   5400
         MaxLength       =   24
         TabIndex        =   22
         Top             =   1440
         Width           =   2115
      End
      Begin VB.CheckBox check1 
         Caption         =   "Substancia Uso Restrito"
         DataField       =   "NOT01"
         DataSource      =   "datPrimaryRS"
         Height          =   255
         Index           =   17
         Left            =   4320
         TabIndex        =   21
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "NIVEL"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   10
         Left            =   3300
         TabIndex        =   20
         Top             =   2580
         Width           =   495
      End
      Begin VB.CheckBox check1 
         Caption         =   "Item de Segurança"
         DataField       =   "ITEM"
         DataSource      =   "datPrimaryRS"
         Height          =   255
         Index           =   14
         Left            =   4320
         TabIndex        =   19
         Top             =   2880
         Width           =   2055
      End
      Begin VB.CommandButton ESCMC02A 
         Height          =   315
         Left            =   2340
         Picture         =   "Frmppap.frx":1BB2
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   780
         Width           =   495
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "PESO"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   11
         Left            =   180
         TabIndex        =   17
         Top             =   2460
         Width           =   1935
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "DESENHO"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   7
         Left            =   60
         MaxLength       =   24
         TabIndex        =   16
         Top             =   1440
         Width           =   3315
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "COMNOME"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   4
         Left            =   3000
         MaxLength       =   40
         TabIndex        =   15
         Top             =   840
         Width           =   6015
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "COMPRADOR"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   3
         Left            =   1200
         MaxLength       =   5
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Txtfields 
         DataField       =   "APLICACAO"
         DataSource      =   "datPrimaryRS"
         Height          =   285
         Index           =   16
         Left            =   6900
         MaxLength       =   24
         TabIndex        =   13
         Top             =   1860
         Width           =   2115
      End
      Begin MSFlexGridLib.MSFlexGrid gridlay 
         Height          =   2655
         Left            =   -74640
         TabIndex        =   102
         Top             =   480
         Width           =   7875
         _ExtentX        =   13891
         _ExtentY        =   4683
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridCjto 
         Height          =   2655
         Left            =   -74760
         TabIndex        =   111
         Top             =   600
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   4683
         _Version        =   393216
      End
      Begin XPControls.XPButton Cmdmw02 
         Height          =   435
         Left            =   -70560
         TabIndex        =   129
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
         Picture         =   "Frmppap.frx":213C
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
      Begin VBCCR17.DTPicker DTPicker1 
         Height          =   375
         Left            =   -68760
         TabIndex        =   130
         Top             =   960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Value           =   44285
         CheckBox        =   -1  'True
         AllowUserInput  =   -1  'True
      End
      Begin VBCCR17.DTPicker DTPicker2 
         Height          =   375
         Left            =   3480
         TabIndex        =   131
         Top             =   1440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Value           =   44285
         CheckBox        =   -1  'True
         AllowUserInput  =   -1  'True
      End
      Begin VBCCR17.DTPicker DTPicker3 
         Height          =   375
         Left            =   7560
         TabIndex        =   132
         Top             =   1440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Value           =   44285
         CheckBox        =   -1  'True
         AllowUserInput  =   -1  'True
      End
      Begin VB.Label lblLabels 
         Caption         =   "Pedido Compras"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   19
         Left            =   -72960
         TabIndex        =   96
         Top             =   2580
         Width           =   1875
      End
      Begin VB.Label lblLabels 
         Caption         =   "Prg. Entrega"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   18
         Left            =   -74640
         TabIndex        =   95
         Top             =   2580
         Width           =   915
      End
      Begin VB.Label lblLabels 
         Caption         =   "Fornecedor"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   17
         Left            =   -74700
         TabIndex        =   86
         Top             =   780
         Width           =   915
      End
      Begin VB.Label lblLabels 
         Caption         =   "Explicações Comentarios"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   23
         Left            =   -74820
         TabIndex        =   77
         Top             =   1500
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Data"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   32
         Left            =   -69420
         TabIndex        =   76
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Disposição de Cert. (A)(R)(O)"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   4
         Left            =   -74820
         TabIndex        =   75
         Top             =   1020
         Width           =   2235
      End
      Begin VB.Label Label1 
         Caption         =   "Aprovação Funcional (A)(D)"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -72120
         TabIndex        =   74
         Top             =   1020
         Width           =   2115
      End
      Begin VB.Label lblLabels 
         Caption         =   "Responsavel"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   -74820
         TabIndex        =   73
         Top             =   660
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Cargo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   15
         Left            =   -71580
         TabIndex        =   72
         Top             =   660
         Width           =   795
      End
      Begin VB.Label lblLabels 
         Caption         =   "Alteraçoes Adc. Eng."
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   13
         Left            =   -74820
         TabIndex        =   62
         Top             =   2580
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Molde"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -70080
         TabIndex        =   61
         Top             =   2280
         Width           =   675
      End
      Begin VB.Label lblLabels 
         Caption         =   "Cliente/Divisào"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   16
         Left            =   3480
         TabIndex        =   44
         Top             =   480
         Width           =   1395
      End
      Begin VB.Label lblLabels 
         Caption         =   "Pedido Cliente"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   43
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Auxiliar"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   14
         Left            =   3000
         TabIndex        =   42
         Top             =   1860
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Exposto"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   12
         Left            =   60
         TabIndex        =   41
         Top             =   1860
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nivel Alt.Data"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   8
         Left            =   7560
         TabIndex        =   40
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nivel Alteraçao"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   6
         Left            =   5400
         TabIndex        =   39
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label11 
         Caption         =   "12345"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   2280
         TabIndex        =   38
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nivel de Submissão"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   24
         Left            =   2280
         TabIndex        =   37
         Top             =   2340
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Caption         =   "Peso:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   11
         Left            =   300
         TabIndex        =   36
         Top             =   2220
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Data Alteração"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   9
         Left            =   3420
         TabIndex        =   35
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Caption         =   "Desenho:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   34
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Comprador:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   33
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Aplicação"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   31
         Left            =   6000
         TabIndex        =   32
         Top             =   1860
         Width           =   795
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PF"
      Height          =   255
      Left            =   2820
      TabIndex        =   11
      Top             =   420
      Width           =   315
   End
   Begin VB.CommandButton ESCMS01A 
      Caption         =   "Mana"
      Height          =   255
      Index           =   0
      Left            =   2820
      TabIndex        =   10
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton EscMA01A 
      Height          =   315
      Index           =   0
      Left            =   1800
      Picture         =   "Frmppap.frx":26D6
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox Txtfields 
      DataField       =   "NOME"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   6
      Left            =   4020
      MaxLength       =   40
      TabIndex        =   7
      Top             =   420
      Width           =   4635
   End
   Begin VB.TextBox Txtfields 
      DataField       =   "CODIGO"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   5
      Left            =   1200
      MaxLength       =   24
      TabIndex        =   6
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox Txtfields 
      DataField       =   "CLINOME"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   2
      Left            =   4560
      MaxLength       =   50
      TabIndex        =   4
      Top             =   960
      Width           =   5715
   End
   Begin VB.TextBox Txtfields 
      DataField       =   "CLIENTE"
      DataSource      =   "datPrimaryRS"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtPPAP 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   10320
      TabIndex        =   125
      Top             =   720
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "Frmppap.frx":2C60
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
      Left            =   10320
      TabIndex        =   126
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "Frmppap.frx":31FA
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
   Begin XPControls.XPButton Command3 
      Height          =   375
      Left            =   4080
      TabIndex        =   127
      Top             =   840
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "Frmppap.frx":3794
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
   Begin XPControls.XPButton Command5 
      Height          =   375
      Left            =   3480
      TabIndex        =   128
      Top             =   0
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "Frmppap.frx":3D2E
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
      Caption         =   "PF"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   21
      Left            =   4080
      TabIndex        =   113
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Caption         =   "Disposição"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   20
      Left            =   6960
      TabIndex        =   108
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Código:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   5
      Left            =   1260
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cliente:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "PPAP:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmPPAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cARQPF As String
Dim cSQL As String
Dim aVAL As Variant
Dim aFOR As Variant
Dim aCAM As Variant
Dim aPAD As Variant
Dim nCAMPOS As Integer
Dim iLOOP As Integer

Private Sub Cmdcjtcrn_Click(Index As Integer)
    Dim cDESENHO As String
    Dim sSQL As String
    GridCjto.Col = 1
    cDESENHO = GridCjto
    Select Case iPPAP
    Case 1, 11
        sSQL = "PPAPC"
    Case 2
        sSQL = "PPAFC"
    Case 3
        sSQL = "PPAGC"
    End Select
    sSQL = "SELECT * FROM " & sSQL & " WHERE CODIGO='" & cDESENHO & "' AND PPAP=" & nPPAP
    IncluiSQL cARQPF, sSQL, 4, Array("CODIGO", "PPAP", "DATA", "PREVISTO"), _
        Array(cDESENHO, nPPAP, Today(), Today() + 365 + 365), False, _
        False
    ''FALSE POIS varias datas
    FILGRIDlay
End Sub

Private Sub cmdClose_Click()
    On Error Resume Next
    If MDG("Gravar alteraçôes") Then
        For iLOOP = 0 To 32
            aVAL(iLOOP) = Txtfields(iLOOP)
        Next iLOOP
        aVAL(33) = DTPicker1.Value
        aVAL(34) = DTPicker2.Value
        aVAL(35) = DTPicker3.Value
        For iLOOP = 36 To 57
            aVAL(iLOOP) = FixNumBol(check1(iLOOP).Value)
        Next iLOOP
        GrvSQL cARQPF, cSQL, nCAMPOS, aCAM, aVAL, aFOR
    End If
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub cmdcopcod_Click()
    Txtfields(24) = Txtfields(5)
    Txtfields(25) = Txtfields(6)
End Sub

Private Sub CmdEdilay_Click(Index As Integer)
    gridlay.Col = 0
    nREVI = gridlay
    FrmPPAPC.Show vbModal, Me
    FILGRIDlay
End Sub

Private Sub CmdEditar_Click(Index As Integer)
    eRETU02 = " Desenho:" & Txtfields(5) & "-" & Replace(Txtfields(6), " ", "_")
    eRETU02 = eRETU02 & " Componente:" & Txtfields(24) & "-" & Replace(Txtfields(25), " ", "_")
    eRETU02 = eRETU02 & " Fornecedor:" & Txtfields(22) & "-" & Replace(Txtfields(21), " ", "_")
    eRETU02 = eRETU02 & " Cliente:" & Txtfields(1) & "-" & Replace(Txtfields(2), " ", "_")
    Grid.Col = 0
    nREVI = Grid
    frmPPAPI.Show vbModal
    Filgrid
    FILGRIDlay
End Sub

Private Sub CmdExcjto_Click(Index As Integer)
    'Dim cdb              As String
    Dim sSQL As String
    GridCjto.Col = 1
    cARQRTF = GridCjto

    Select Case iPPAP
    Case 1, 11
        sSQL = "select * from PPAPP WHERE PPAP=" & nPPAP & " AND CODIGO='" & cARQRTF & "'"
    Case 2
        sSQL = "select * from PPAFP WHERE PPAP=" & nPPAP & " AND CODIGO='" & cARQRTF & "'"
    Case 3
        sSQL = "select * from PPAGP WHERE PPAP=" & nPPAP & " AND CODIGO='" & cARQRTF & "'"
    End Select
    
    If ApagaSQLP(cARQPF, sSQL) Then
        filgridcjto
    End If

End Sub

Private Sub CmdExcLay_Click(Index As Integer)
    Dim sSQL As String
    gridlay.Col = 0
    nREVI = FixInt(gridlay)
    Select Case iPPAP
    Case 1, 11
        sSQL = "select * from PPAPC WHERE ITEM=" & nREVI
    Case 2
        sSQL = "select * from PPAFC WHERE ITEM=" & nREVI
    Case 3
        sSQL = "select * from PPAGC WHERE ITEM=" & nREVI
    End Select
    
    If ApagaSQLP(cARQPF, sSQL) Then
        FILGRIDlay
    End If

End Sub

Private Sub CmdExcluir_Click(Index As Integer)
    Dim sSQL As String
    Grid.Col = 0
    nREVI = FixInt(Grid)
    Select Case iPPAP
    Case 1, 11
        sSQL = "select * from PPAPI WHERE PPAP=" & nPPAP & " AND  ITEM=" & nREVI
    Case 2
        sSQL = "select * from PPAFI WHERE PPAP=" & nPPAP & " AND  ITEM=" & nREVI
    Case 3
        sSQL = "select * from PPAGI WHERE PPAP=" & nPPAP & " AND  ITEM=" & nREVI

    End Select
    If ApagaSQLP(cARQPF, sSQL) Then
        Filgrid
    End If

End Sub

Private Sub Cmdimppf_Click(Index As Integer)
    Dim DB As New ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim RSTAB As ADODB.Recordset
    Dim sSQL As String
    DB.ConnectionTimeout = 120
    DB.Open GeracArq(cARQPF)
    sSQL = "SELECT CODIGO,PF FROM PF WHERE CODFINAL='" & Trim(Txtfields(5)) & "' AND NOT BLOQUEADO"
    rs.Open sSQL, DB, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then
        While Not rs.EOF
            Select Case iPPAP
            Case 1, 11
                sSQL = "select * from PPAPP WHERE PPAP=" & nPPAP & " AND CODIGO='" & rs("CODIGO") & "'"
            Case 2
                sSQL = "select * from PPAFP WHERE PPAP=" & nPPAP & " AND CODIGO='" & rs("CODIGO") & "'"
            Case 3
                sSQL = "select * from PPAGP WHERE PPAP=" & nPPAP & " AND CODIGO='" & rs("CODIGO") & "'"
            End Select
            RSTAB.Open sSQL, DB, adOpenDynamic, adLockOptimistic
            If RSTAB.EOF Then
                RSTAB.AddNew
                RSTAB("codigo") = rs("CODIGO")
                RSTAB("PPAP") = nPPAP
                RSTAB("PF") = rs("PF")
                RSTAB.Update
            Else
                If rs("PF") > 0 Then
                    RSTAB("PF") = rs("PF")
                    RSTAB.Update
                End If
            End If
            RSTAB.Close
            rs.MoveNext
        Wend
    End If
    rs.Close
    DB.Close
    filgridcjto
End Sub

Private Sub Cmdmw02_Click()
    Dim cSQL As String
    If Not IsNumeric(Txtfields(27)) Then Exit Sub
    If Not IsNumeric(Txtfields(28)) Then Exit Sub
    If Not IsNumeric(Txtfields(0)) Then Exit Sub

    cSQL = "SELECT FROM MW02 WHERE ="
    cSQL = cSQL & PadLeft(Txtfields(27).tEXT, 8) + PadLeft(Txtfields(28).tEXT, 3)
     
    GrvSQL GeraConn(zMANA5EMP, "FOX"), cSQL, 2, Array("ppap", "ppapD"), Array(Txtfields(0), Date), Array("N", "D") ''GrvSQLSDE zMANA5EMP
 

End Sub

Private Sub CmdNovCjto_Click(Index As Integer)
    Dim cDESENHO         As String
    Dim sSQL As String
    cDESENHO = InputBox("Digite o Produto", "Inclusão Produto", "__")

    Select Case iPPAP
    Case 1, 11
        sSQL = "PPAPP"
    Case 2
        sSQL = "PPAFP"
    Case 3
        sSQL = "PPAGP"

    End Select
    
    IncluiSQL cARQPF, "SELECT * FROM " & sSQL & " where codigo='" & cDESENHO & "' aND PPAP=" & nPPAP, 2, Array("CODIGO", "PPAP"), Array(cDESENHO, nPPAP), True, False

    filgridcjto

End Sub

Private Sub CmdNovoLay_Click(Index As Integer)
    Dim cDESENHO         As String
    Dim dDATA As Date
    Dim sSQL As String
  
    If Index = 1 Then
        cDESENHO = InputBox("Digite o Produto", "Inclusão Produto", Txtfields(5).tEXT)
        dDATA = Today()
    Else
        gridlay.Row = (gridlay.Rows) - 1         ''UltimoItem
        gridlay.Col = 1
        cDESENHO = gridlay
        gridlay.Col = 4
        If Not IsDate(gridlay) Then
            Alert ("Ultima Data Efetivacao Nao Preenchida")
            Exit Sub
        Else
            dDATA = gridlay
        End If
    End If
    Select Case iPPAP
    Case 1, 11
        sSQL = "PPAPC"
    Case 2
        sSQL = "PPAFC"
    Case 3
        sSQL = "PPAGC"

    End Select
    IncluiSQL cARQPF, "SELECT * FROM " & sSQL & " where codigo='" & cDESENHO & "' and PPAP=" & nPPAP, 4, Array("CODIGO", "PPAP", "DATA", "PREVISTO"), _
        Array(cDESENHO, nPPAP, dDATA, dDATA + 365 + 365), False, _
        False
    ''busca false pois tem data
    FILGRIDlay
End Sub

Private Sub CmdNovo_Click(Index As Integer)
    Dim nITEM            As Long
    Dim sSQL As String
    Select Case iPPAP
    Case 1, 11
        nITEM = FixInt(PegMAXSQL(cARQPF, "PPAPI WHERE PPAP=" & nPPAP, "ITEM", 0))
    Case 2
        nITEM = FixInt(PegMAXSQL(cARQPF, "PPAFI WHERE PPAP=" & nPPAP, "ITEM", 0))
    Case 3
        nITEM = FixInt(PegMAXSQL(cARQPF, "PPAGI WHERE PPAP=" & nPPAP, "ITEM", 0))
    End Select
    nITEM = nITEM + 1
    Select Case iPPAP
    Case 1, 11
        sSQL = "PPAPI"
    Case 2
        sSQL = "PPAFI"
    Case 3
        sSQL = "PPAGI"
    End Select
    IncluiSQL cARQPF, "select * from " & sSQL & " where PPAP= " & nPPAP & " AND ITEM=" & nITEM, 3, Array("PPAP", "DATA", "ITEM"), _
        Array(nPPAP, Date, nITEM), True, _
        False
                            
    Filgrid
End Sub

Private Sub Cmdped_Click()
    Dim cARQ         As String
    Dim nPE As Long
    Dim aRETU As Variant
    Dim sSQL As String
    nPE = FixInt(Txtfields(26))
    cARQ = PegPath("PATH", "MANA5PCP")
    cARQ = GeraConn(cARQ, "JETFOX")
    sSQL = "SELECT PEDIDO,COMPRAS,COMITEM FROM PE WHERE PEDIDO=" & nPE
    aRETU = PegSQL(cARQ, sSQL, 2, Array("COMPRAS", "COMITEM"), Array("NI", "NI"), Array(0, 0))
    If lRETU Then
        Txtfields(27) = aRETU(0)
        Txtfields(28) = aRETU(1)
    End If
End Sub

Private Sub Cmdprg_Click()
    Dim cARQ   As String
    Dim nFORN As Long
    Dim cCODIGO As String
    Dim aRETU As Variant
    Dim sSQL As String
    nFORN = FixInt(Txtfields(22), 0)
    cCODIGO = FixStr(Txtfields(24), "", "TRIM")
    cARQ = PegPath("PATH", "MANA5PCP")
    cARQ = GeraConn(cARQ, "JETFOX")
    sSQL = "SELECT PEDIDO,COMPRAS,COMITEM FROM PE WHERE FORNECEDO=" & nFORN & " AND CODIGO='" & cCODIGO & "'"
    aRETU = PegSQL(cARQ, sSQL, 3, Array("PEDIDO", "COMPRAS", "COMITEM"), Array("NI", "NI", "NI"), Array(0, 0, 0))
    If lRETU Then
        Txtfields(26) = aRETU(0)
        Txtfields(27) = aRETU(1)
        Txtfields(28) = aRETU(2)
    Else
        Alert ("Nao achei pedido para este fornecedor/produto")
    End If
End Sub

Private Sub Command10_Click()
    Dim cARQ As String
    Dim aRETU As Variant
    Dim sSQL      As String
    Dim cCODIGO      As String
    cCODIGO = FixStr(Txtfields(24), "", "TRIM")
    Select Case Txtfields(23)

    Case "M"
        sSQL = "SELECT NOME FROM MU01 WHERE CODIGO='" & cCODIGO & "'"

    Case "C"
        sSQL = "SELECT NOME FROM MT01 WHERE CODIGO='" & cCODIGO & "'"

    Case "T"
        sSQL = "SELECT NOME FROM MP03 WHERE CODIGO='" & cCODIGO & "'"

    End Select

    aRETU = PegSQL(cARQ, sSQL, 1, Array("NOME"), Array("C"), Array(""))
    If lRETU Then
        Txtfields(25) = aRETU(0)
    End If

End Sub

Private Sub Command1_Click()
    Dim cSQL As String
    Dim aRETU As Variant
    escpffim.Show vbModal, Me
    If lRETU Then
        Txtfields(5) = eRETU02
        Txtfields(6) = eRETU03
        Txtfields(32) = eRETU04
        cSQL = "select CONPES from PF WHERE PF=" & eRETU04
        aRETU = PegSQL(cARQPF, cSQL, 1, Array("CONPES"), Array("N"), Array(0))
        If lRETU Then
            Txtfields(11) = aRETU(0)
        End If
    End If
End Sub

Private Sub Command11_Click()
    Txtfields(22) = Txtfields(1)
    Txtfields(21) = Txtfields(2)

End Sub

Private Sub Command12_Click()
    Dim cSQLTMP As String
    Dim aRETU As Variant
    cSQLTMP = "SELECT PRGENT,COMPED,ITEM FROM MW02 WHERE PPAP=" & Txtfields(0).tEXT
    aRETU = PegSQL(GeraConn(zMANA5EMP, "JETFOX"), cSQLTMP, 3, Array("PRGENT", "COMPED", "ITEM"), _
                   Array("NI", "NI", "NI"), Array(0, 0, 0))
    If lRETU Then
        Txtfields(26) = aRETU(0)
        Txtfields(27) = aRETU(1)
        Txtfields(28) = aRETU(2)
    End If
End Sub

Private Sub Command2_Click()
    escMP05.Show vbModal, Me
    If lRETU Then
        Txtfields(17) = eRETU01
        Txtfields(18) = eRETU02
    End If
End Sub

Private Sub Command3_Click()
    Dim cARQ As String
    Dim aRETU As Variant
    Dim sSQL             As String
    Dim nNUMERO          As Long
    nNUMERO = FixInt(Txtfields(1), 0)
    cARQ = GeraConn(zMANA5EMP, "JETFOX")
    sSQL = "SELECT NOME FROM MA01 WHERE NUMERO=" & nNUMERO
    If iPPAP = 2 Then
        sSQL = "SELECT NOME FROM MB01 WHERE NUMERO=" & nNUMERO
    End If
    aRETU = PegSQL(cARQ, sSQL, 1, Array("NOME"), Array("C"), Array(""))
    If lRETU Then
        Txtfields(2) = aRETU(0)
    End If
End Sub

Private Sub Command4_Click()
    Dim cARQ As String
    Dim aRETU As Variant
    Dim sSQL             As String
    Dim cCODIGO          As String
    cCODIGO = FixStr(Txtfields(3), "", "TRIM")
    cARQ = GeraConn(zMANA5EMP, "JETFOX")
    sSQL = "SELECT NOME FROM MC02 WHERE NUMERO='" & cCODIGO & "'"
    aRETU = PegSQL(cARQ, sSQL, 1, Array("NOME"), Array("C"), Array(""))
    If lRETU Then
        Txtfields(4) = aRETU(0)
    End If
End Sub

Private Sub Command5_Click()
    Dim cARQ As String
    Dim aRETU As Variant
    Dim sSQL             As String
    Dim cCODIGO          As String
    cCODIGO = FixStr(Txtfields(5), "", "TRIM")
    cARQ = GeraConn(zMANA5EMP, "JETFOX")
    sSQL = "SELECT NOME FROM MS01 WHERE CODIGO='" & cCODIGO & "'"
    aRETU = PegSQL(cARQ, sSQL, 1, Array("NOME"), Array("C"), Array(""))
    If lRETU Then
        Txtfields(6) = aRETU(0)
    End If
End Sub

Private Sub Command6_Click()
    Dim cARQ As String
    Dim aRETU As Variant
    Dim sSQL             As String
    Dim nNUMERO          As Long
    nNUMERO = FixInt(Txtfields(22), 0)
    cARQ = GeraConn(zMANA5EMP, "JETFOX")
    sSQL = "SELECT NOME FROM MB01 WHERE NUMERO=" & nNUMERO
    aRETU = PegSQL(cARQ, sSQL, 1, Array("NOME"), Array("C"), Array(""))
    If lRETU Then
        Txtfields(21) = aRETU(0)
    End If
End Sub

Private Sub Command7_Click()
    lRETU = False
    eRETU01 = ""
    eRETU02 = ""
    cARQESC = "MB01"
    escNUMNOM.Show vbModal, Me

    If lRETU Then
        Txtfields(22) = eRETU01
        Txtfields(21) = eRETU02
    End If

End Sub

Private Sub Command8_Click(Index As Integer)

    Select Case Index

    Case 0
        Txtfields(23) = "M"

    Case 1
        Txtfields(23) = "C"

    Case 2
        Txtfields(23) = "T"

    End Select

End Sub

Private Sub Command9_Click(Index As Integer)
    lRETU = False
    If Index = 0 Then
        ePASS01 = "MANA5"
        Select Case Txtfields(23)
        Case "C"
            iMU01 = 2
        Case "T"
            iMU01 = 7
        Case "M"
            iMU01 = 1
        Case Else
            Exit Sub
        End Select
        escmu01.Show vbModal, Me
    End If
    
    If Index = 1 Then
        Select Case Txtfields(23)
        Case "C"
            ePASS01 = "LOGIC"
        Case "T"
            Alert ("Nao disponivel logix")
        Case "M"
            ePASS01 = "LOGIM"
        Case Else
            Exit Sub
        End Select
        ePASS01 = "LOGIC"
        escms01.Show vbModal, Me
    End If
    
    
    If lRETU Then
    
        Txtfields(24) = eRETU01
        Txtfields(25) = eRETU02
    
    End If
    
End Sub

Private Sub Encerrar_Click()
    If Not MDG("Sair sem gravar") Then
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub EscMA01A_Click(Index As Integer)

    lRETU = False
    eRETU01 = ""
    eRETU02 = ""
    cARQESC = "MA01"
    If iPPAP = 2 Then
        cARQESC = "MB01"
    End If
    
    If Index = 1 Then
        cARQESC = "2DIG"
        ePASS01 = "LOGIX"
    End If
    
    If Index = 2 Then
        ePASS01 = "LOGIX"
    End If

    If Index = 6 Then
        ePASS01 = "MICRO"
    End If
    
    escNUMNOM.Show vbModal, Me

    If lRETU Then
        frmPPAP.Txtfields(1) = eRETU01
        frmPPAP.Txtfields(2) = eRETU02
    End If
    
    

End Sub

Private Sub ESCMC02A_Click()

    lRETU = False
    eRETU01 = ""
    eRETU02 = ""
    cARQESC = "MC02"
    escNUMNOM.Show vbModal, Me
    If lRETU Then
        frmPPAP.Txtfields(3) = eRETU01
        frmPPAP.Txtfields(4) = eRETU02
    End If
End Sub

Private Sub ESCMS01A_Click(Index As Integer)
    ePASS01 = "MANA5"
    If Index = 1 Then
        ePASS01 = "LOGIX"
    End If
    If Index = 1 Then
        ePASS01 = "MICRO"
    End If
    escms01.Show vbModal, Me
    If lRETU Then
        frmPPAP.Txtfields(5) = eRETU01
        frmPPAP.Txtfields(6) = eRETU02
    End If
End Sub

Private Sub Filgrid()
    Dim sSQL As String
    Select Case iPPAP
    Case 1, 11
        sSQL = "select * from PPAPI WHERE PPAP=" & nPPAP & " ORDER BY ITEM"
    Case 2
        sSQL = "select * from PPAFI WHERE PPAP=" & nPPAP & " ORDER BY ITEM"
    Case 3
        sSQL = "select * from PPAGI WHERE PPAP=" & nPPAP & " ORDER BY ITEM"
    End Select
    MontaGridFast Grid, 4, Array(400, 1200, 400, 2000), Array("Item", "Data", "Dispo", "Obs"), _
        Array("Item", "Data", "Dispo", "OBS"), cARQPF, sSQL
    Grid.Row = (Grid.Rows) - 1
    Grid.Col = 2
    Txtdispo = FixStr(Grid)                      ''Ultima Disposicao
End Sub

Private Sub filgridcjto()
    Dim cSQL             As String
    Select Case iPPAP
    Case 1, 11
        cSQL = "select * from PPAPP WHERE PPAP=" & nPPAP & " ORDER BY CODIGO"
    Case 2
        cSQL = "select * FROM PPAFP WHERE PPAP=" & nPPAP & " ORDER BY CODIGO"
    Case 3
        cSQL = "select * FROM PPAGP WHERE PPAP=" & nPPAP & " ORDER BY CODIGO"
    End Select

    MontaGridFast GridCjto, 3, Array(400, 1200, 800), Array("PPAP", "Produto", "PF"), _
        Array("PPAP", "L$CODIGO", "PF"), cARQPF, cSQL

End Sub

Private Sub FILGRIDlay()
    Dim sSQL As String
    Select Case iPPAP
    Case 1, 11
        sSQL = "select * from PPAPC WHERE PPAP=" & nPPAP & " ORDER BY ITEM"
    Case 2
        sSQL = "select * from PPAFC WHERE PPAP=" & nPPAP & " ORDER BY ITEM"
    Case 3
        sSQL = "select * from PPAGC WHERE PPAP=" & nPPAP & " ORDER BY ITEM"
    End Select
    MontaGridFast gridlay, 5, Array(400, 2400, 1200, 1200, 1200), Array("Item", "Codigo", "Data", "Prevista", "Efetuadas"), _
        Array("ITEM", "L$CODIGO", "DATA", "PREVISTO", "EFETUADO"), cARQPF, sSQL
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
    'Dim cSQLTMP As String
    
    txtPPAP = nPPAP
    cARQPF = PegPath("PATH", "PPAP")

    CenterFormToScreen Me
    Select Case iPPAP
    Case 1, 11
        cSQL = "select * from PPAP WHERE PPAP=" & nPPAP
    Case 2
        cSQL = "select * from PPAF WHERE PPAP=" & nPPAP
        lblLabels(1).Caption = "Fornecedor"
        Me.Caption = "PPAP Fornecedores"
    Case 3
        cSQL = "select * from PPAG WHERE PPAP=" & nPPAP
        Me.Caption = "GP11"
    Case 4                                       'verificar chamado
        'CSQL
        Me.Caption = "Cronograma de Layout"
    End Select
    
    nCAMPOS = 58
    aCAM = Array("APRO", "CLIENTE", "CLINOME", "COMPRADOR", "COMNOME", "CODIGO", "NOME", "DESENHO", "DISPO", "NIVELALT", _
                 "NIVEL", "PESO", "EXPOSTO", "OBSADC01", "OBSADC02", "AUXILIAR", "APLICACAO", "RESNOME", "RESCARGO", "PEDOM", _
                 "DIVISAO", "FORNOM", "FORNECEDOR", "TIPCOM", "CODCOM", "NOMCOM", "PRGENT", "COMPRAS", "COMITEM", "MOLDE", _
                 "EXP01", "EXP02", "PF", "DATA", "DATAALT", "NIVELDAT", "SUB01", "SUB02", "SUB03", "SUB04", _
                 "SUB05", "RES01", "RES03", "SUB06", "SUB07", "SUB08", "SUB09", "RES02", "RES04", "APLIC", _
                 "ITEM", "SUB10", "INF03", "NOT01", "NOT02", "INF01", "INF02", "INATIVO")
    aFOR = Array("C", "NI", "C", "C", "C", "C", "C", "C", "C", "C", _
                 "C", "N", "C", "C", "C", "C", "C", "C", "C", "C", _
                 "C", "C", "NI", "C", "C", "C", "NI", "N", "N", "C", _
                 "C", "C", "NI", "DN", "DN", "DN", "B", "B", "B", "B", _
                 "B", "B", "B", "B", "B", "B", "B", "B", "B", "B", _
                 "B", "B", "B", "B", "B", "B", "B", "B")
    aPAD = Array("", 0, "", "", "", "", "", "", "", "", _
                 "", 0, "", "", "", "", "", "", "", "", _
                 "", "", 0, "", "", "", 0, 0, 0, "", _
                 "", "", 0, "", "", "", "", "", "", "", _
                 "", "", "", "", "", "", "", "", "", "", _
                 "", "", "", "", "", "", "", "")
    aVAL = PegSQL(cARQPF, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
    For iLOOP = 0 To 32
        Txtfields(iLOOP) = aVAL(iLOOP)
    Next iLOOP
    If IsDate(aVAL(33)) Then
        DTPicker1.Value = aVAL(33)
    End If
    If IsDate(aVAL(34)) Then
        DTPicker2.Value = aVAL(34)
    End If
    If IsDate(aVAL(35)) Then
        DTPicker3.Value = aVAL(35)
    End If
    For iLOOP = 36 To 57
        check1(iLOOP - 36).Value = FixBolNum(aVAL(iLOOP))
    Next iLOOP
    
    
    Filgrid
    FILGRIDlay
    filgridcjto
    If iPPAP = 11 Then
        EnableControls Me, False
        SSTab1.Enabled = True
       ' SSTab1.Tab = 5
       ' sstab1.
        cmdClose.Enabled = True
        gridlay.Enabled = True
        CmdNovoLay(1).Enabled = True
        CmdExcLay(1).Enabled = True
        CmdEdilay(1).Enabled = True
        CmdNovoLay(0).Enabled = True
        GridCjto.Enabled = True
        CmdNovCjto(2).Enabled = True
        CmdExcjto(2).Enabled = True
        Cmdimppf(0).Enabled = True
        Cmdcjtcrn(0).Enabled = True
    End If
    Screen.MousePointer = vbDefault


End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault

End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
    Case 26, 27, 28
        KeyAscii = ValiText(KeyAscii, "#NI")
    End Select
End Sub


