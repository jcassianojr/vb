VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{EA478B61-D9EC-47F6-BB21-95A533AF2251}#1.0#0"; "TabExC01.OCX"
Object = "{451B73A5-1563-45D5-A6AC-7B2B7D30B778}#1.0#0"; "BSPrin10.ocx"
Begin VB.Form frmPFS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PFS"
   ClientHeight    =   6540
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   12180
   Icon            =   "frmPFS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   12180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin BSPrinter.PrintPreview PrintPreview1 
      Left            =   9240
      Top             =   120
      _ExtentX        =   1191
      _ExtentY        =   1191
      LcK1b           =   "yefT59bnyufI583n1ufV59HnyefN58nn0+fK58nnzefR58rn1+fW5w=="
      LcK2b           =   "reeI58zn+ue059jn1Of656fni+eN543nleeM54HnlueR5w=="
      AmbientBb       =   $"frmPFS.frx":058A
   End
   Begin VB.CommandButton cmdpeg 
      Height          =   360
      Index           =   0
      Left            =   8160
      Picture         =   "frmPFS.frx":062A
      Style           =   1  'Graphical
      TabIndex        =   319
      TabStop         =   0   'False
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   7560
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   318
      TabStop         =   0   'False
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton ESCpro 
      Caption         =   "LX"
      Height          =   315
      Index           =   0
      Left            =   4920
      TabIndex        =   252
      TabStop         =   0   'False
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   67
      Left            =   5400
      MaxLength       =   24
      TabIndex        =   248
      Top             =   360
      Width           =   1995
   End
   Begin TabExCtl.SSTabEx SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   9128
      Tabs            =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabsPerRow      =   6
      Tab             =   5
      TabHeight       =   706
      TabCaption(0)   =   "Mao de Obra"
      Tab(0).ControlCount=   63
      Tab(0).Control(0)=   "escArranjo(2)"
      Tab(0).Control(1)=   "CmdEscMe01"
      Tab(0).Control(2)=   "escArranjo(1)"
      Tab(0).Control(3)=   "escRecursoHomem(13)"
      Tab(0).Control(4)=   "escRecursoHomem(12)"
      Tab(0).Control(5)=   "escRecursoHomem(11)"
      Tab(0).Control(6)=   "escRecursoHomem(10)"
      Tab(0).Control(7)=   "escRecursoHomem(3)"
      Tab(0).Control(8)=   "escRecursoHomem(2)"
      Tab(0).Control(9)=   "escRecursoHomem(1)"
      Tab(0).Control(10)=   "escRecursoHomem(0)"
      Tab(0).Control(11)=   "escArranjo(0)"
      Tab(0).Control(12)=   "TXTFIELDS(89)"
      Tab(0).Control(13)=   "TXTFIELDS(88)"
      Tab(0).Control(14)=   "TXTFIELDS(87)"
      Tab(0).Control(15)=   "TXTFIELDS(86)"
      Tab(0).Control(16)=   "TXTFIELDS(85)"
      Tab(0).Control(17)=   "TXTFIELDS(8)"
      Tab(0).Control(18)=   "TXTFIELDS(9)"
      Tab(0).Control(19)=   "TXTFIELDS(10)"
      Tab(0).Control(20)=   "TXTFIELDS(11)"
      Tab(0).Control(21)=   "EscMP01"
      Tab(0).Control(22)=   "ESCMP03"
      Tab(0).Control(23)=   "TXTFIELDS(4)"
      Tab(0).Control(24)=   "TXTFIELDS(5)"
      Tab(0).Control(25)=   "TXTFIELDS(7)"
      Tab(0).Control(26)=   "TXTFIELDS(6)"
      Tab(0).Control(27)=   "TXTFIELDS(15)"
      Tab(0).Control(28)=   "TXTFIELDS(16)"
      Tab(0).Control(29)=   "TXTFIELDS(17)"
      Tab(0).Control(30)=   "TXTFIELDS(18)"
      Tab(0).Control(31)=   "TXTFIELDS(23)"
      Tab(0).Control(32)=   "TXTFIELDS(24)"
      Tab(0).Control(33)=   "Command7(0)"
      Tab(0).Control(34)=   "Command7(1)"
      Tab(0).Control(35)=   "TXTFIELDS(31)"
      Tab(0).Control(36)=   "Command8"
      Tab(0).Control(37)=   "TXTFIELDS(32)"
      Tab(0).Control(38)=   "TXTFIELDS(33)"
      Tab(0).Control(39)=   "escsetref"
      Tab(0).Control(40)=   "escsetins"
      Tab(0).Control(41)=   "TXTFIELDS(34)"
      Tab(0).Control(42)=   "TXTFIELDS(35)"
      Tab(0).Control(43)=   "escmon"
      Tab(0).Control(44)=   "escloc"
      Tab(0).Control(45)=   "Command9(0)"
      Tab(0).Control(46)=   "Command9(1)"
      Tab(0).Control(47)=   "Command9(2)"
      Tab(0).Control(48)=   "Command9(3)"
      Tab(0).Control(49)=   "Command9(4)"
      Tab(0).Control(50)=   "Command9(5)"
      Tab(0).Control(51)=   "Command11"
      Tab(0).Control(52)=   "Command12"
      Tab(0).Control(53)=   "CmdEditar"
      Tab(0).Control(54)=   "lblLabels(8)"
      Tab(0).Control(55)=   "lblLabels(9)"
      Tab(0).Control(56)=   "lblLabels(11)"
      Tab(0).Control(57)=   "lblLabels(5)"
      Tab(0).Control(58)=   "lblLabels(7)"
      Tab(0).Control(59)=   "lblLabels(15)"
      Tab(0).Control(60)=   "lblLabels(16)"
      Tab(0).Control(61)=   "lblLabels(17)"
      Tab(0).Control(62)=   "lblLabels(18)"
      TabCaption(1)   =   "FiguraProcesso"
      Tab(1).ControlCount=   8
      Tab(1).Control(0)=   "Picture2(0)"
      Tab(1).Control(1)=   "Picture1(0)"
      Tab(1).Control(2)=   "Incimg(0)"
      Tab(1).Control(3)=   "DelImg(0)"
      Tab(1).Control(4)=   "commandx(0)"
      Tab(1).Control(5)=   "cmdimprimir(0)"
      Tab(1).Control(6)=   "CmdPaste(0)"
      Tab(1).Control(7)=   "VerImg(0)"
      TabCaption(2)   =   "Ferramenta Fluxo"
      Tab(2).ControlCount=   75
      Tab(2).Control(0)=   "TXTFIELDS(65)"
      Tab(2).Control(1)=   "TXTFIELDS(12)"
      Tab(2).Control(2)=   "TXTFIELDS(66)"
      Tab(2).Control(3)=   "EscCodFlu(13)"
      Tab(2).Control(4)=   "CMDSIG(13)"
      Tab(2).Control(5)=   "EscCodFlu(12)"
      Tab(2).Control(6)=   "CMDSIG(12)"
      Tab(2).Control(7)=   "TXTFIELDS(63)"
      Tab(2).Control(8)=   "TXTFIELDS(62)"
      Tab(2).Control(9)=   "TXTFIELDS(13)"
      Tab(2).Control(10)=   "TXTFIELDS(14)"
      Tab(2).Control(11)=   "Command1(0)"
      Tab(2).Control(12)=   "TXTFIELDS(51)"
      Tab(2).Control(13)=   "TXTFIELDS(52)"
      Tab(2).Control(14)=   "TXTFIELDS(53)"
      Tab(2).Control(15)=   "TXTFIELDS(54)"
      Tab(2).Control(16)=   "TXTFIELDS(55)"
      Tab(2).Control(17)=   "TXTFIELDS(56)"
      Tab(2).Control(18)=   "EscCodFlu(1)"
      Tab(2).Control(19)=   "EscCodFlu(2)"
      Tab(2).Control(20)=   "EscCodFlu(3)"
      Tab(2).Control(21)=   "EscCodFlu(4)"
      Tab(2).Control(22)=   "EscCodFlu(5)"
      Tab(2).Control(23)=   "EscCodFlu(6)"
      Tab(2).Control(24)=   "TXTFIELDS(57)"
      Tab(2).Control(25)=   "TXTFIELDS(58)"
      Tab(2).Control(26)=   "TXTFIELDS(59)"
      Tab(2).Control(27)=   "TXTFIELDS(60)"
      Tab(2).Control(28)=   "EscCodFlu(7)"
      Tab(2).Control(29)=   "EscCodFlu(8)"
      Tab(2).Control(30)=   "EscCodFlu(9)"
      Tab(2).Control(31)=   "EscCodFlu(10)"
      Tab(2).Control(32)=   "TXTFIELDS(20)"
      Tab(2).Control(33)=   "TXTFIELDS(21)"
      Tab(2).Control(34)=   "Command6"
      Tab(2).Control(35)=   "TXTFIELDS(25)"
      Tab(2).Control(36)=   "TXTFIELDS(28)"
      Tab(2).Control(37)=   "TXTFIELDS(29)"
      Tab(2).Control(38)=   "TXTFIELDS(30)"
      Tab(2).Control(39)=   "Command1(1)"
      Tab(2).Control(40)=   "Command1(2)"
      Tab(2).Control(41)=   "Command1(3)"
      Tab(2).Control(42)=   "Command9(6)"
      Tab(2).Control(43)=   "Edifer(0)"
      Tab(2).Control(44)=   "Edifer(1)"
      Tab(2).Control(45)=   "Edifer(2)"
      Tab(2).Control(46)=   "Edifer(3)"
      Tab(2).Control(47)=   "CmdForFer(0)"
      Tab(2).Control(48)=   "CmdForFer(1)"
      Tab(2).Control(49)=   "CmdForFer(2)"
      Tab(2).Control(50)=   "CmdForFer(3)"
      Tab(2).Control(51)=   "CMDSIG(1)"
      Tab(2).Control(52)=   "CMDSIG(2)"
      Tab(2).Control(53)=   "CMDSIG(3)"
      Tab(2).Control(54)=   "CMDSIG(4)"
      Tab(2).Control(55)=   "CMDSIG(5)"
      Tab(2).Control(56)=   "CMDSIG(6)"
      Tab(2).Control(57)=   "CMDSIG(7)"
      Tab(2).Control(58)=   "CMDSIG(8)"
      Tab(2).Control(59)=   "CMDSIG(9)"
      Tab(2).Control(60)=   "CMDSIG(10)"
      Tab(2).Control(61)=   "CMDSIG(11)"
      Tab(2).Control(62)=   "EscCodFlu(11)"
      Tab(2).Control(63)=   "TXTFIELDS(61)"
      Tab(2).Control(64)=   "TXTFIELDS(19)"
      Tab(2).Control(65)=   "lblLabels(20)"
      Tab(2).Control(66)=   "lblLabels(21)"
      Tab(2).Control(67)=   "lblLabels(12)"
      Tab(2).Control(68)=   "lblLabels(13)"
      Tab(2).Control(69)=   "lblLabels(14)"
      Tab(2).Control(70)=   "Label2"
      Tab(2).Control(71)=   "lblLabels(4)"
      Tab(2).Control(72)=   "lblLabels(10)"
      Tab(2).Control(73)=   "Label13"
      Tab(2).Control(74)=   "lblLabels(19)"
      TabCaption(3)   =   "EPI/Observações"
      Tab(3).ControlCount=   14
      Tab(3).Control(0)=   "EPI07"
      Tab(3).Control(1)=   "EPI06"
      Tab(3).Control(2)=   "EPI05"
      Tab(3).Control(3)=   "EPI04"
      Tab(3).Control(4)=   "EPI03"
      Tab(3).Control(5)=   "EPI02"
      Tab(3).Control(6)=   "EPI01"
      Tab(3).Control(7)=   "TXTFIELDS(26)"
      Tab(3).Control(8)=   "TXTFIELDS(27)"
      Tab(3).Control(9)=   "Cmdepi"
      Tab(3).Control(10)=   "Cmdapgobs"
      Tab(3).Control(11)=   "Command10"
      Tab(3).Control(12)=   "Label1"
      Tab(3).Control(13)=   "Label3"
      TabCaption(4)   =   "Instrucoes"
      Tab(4).ControlCount=   8
      Tab(4).Control(0)=   "CmdRenumera"
      Tab(4).Control(1)=   "InsEdit"
      Tab(4).Control(2)=   "Command3"
      Tab(4).Control(3)=   "Command2(1)"
      Tab(4).Control(4)=   "Cmdimp"
      Tab(4).Control(5)=   "GridIns"
      Tab(4).Control(6)=   "txttaktat"
      Tab(4).Control(7)=   "txttak"
      TabCaption(5)   =   "Fluxo 2"
      Tab(5).ControlCount=   6
      Tab(5).Control(0)=   "Picture2(4)"
      Tab(5).Control(1)=   "Command13"
      Tab(5).Control(2)=   "Command14"
      Tab(5).Control(3)=   "Command15"
      Tab(5).Control(4)=   "Command16"
      Tab(5).Control(5)=   "GrdQsbLep"
      TabCaption(6)   =   "Almofada"
      Tab(6).ControlCount=   98
      Tab(6).Control(0)=   "TXTFIELDS(84)"
      Tab(6).Control(1)=   "TXTFIELDS(83)"
      Tab(6).Control(2)=   "TXTFIELDS(82)"
      Tab(6).Control(3)=   "TXTFIELDS(81)"
      Tab(6).Control(4)=   "TXTFIELDS(41)"
      Tab(6).Control(5)=   "TXTFIELDS(42)"
      Tab(6).Control(6)=   "TXTFIELDS(43)"
      Tab(6).Control(7)=   "TXTFIELDS(44)"
      Tab(6).Control(8)=   "TXTFIELDS(45)"
      Tab(6).Control(9)=   "TXTFIELDS(46)"
      Tab(6).Control(10)=   "TXTFIELDS(47)"
      Tab(6).Control(11)=   "TXTFIELDS(48)"
      Tab(6).Control(12)=   "TXTFIELDS(36)"
      Tab(6).Control(13)=   "Command20"
      Tab(6).Control(14)=   "esc1(1)"
      Tab(6).Control(15)=   "esc1(2)"
      Tab(6).Control(16)=   "esc1(3)"
      Tab(6).Control(17)=   "esc1(4)"
      Tab(6).Control(18)=   "esc1(5)"
      Tab(6).Control(19)=   "esc1(6)"
      Tab(6).Control(20)=   "esc1(7)"
      Tab(6).Control(21)=   "esc1(8)"
      Tab(6).Control(22)=   "esc1(9)"
      Tab(6).Control(23)=   "esc1(10)"
      Tab(6).Control(24)=   "esc1(11)"
      Tab(6).Control(25)=   "esc1(12)"
      Tab(6).Control(26)=   "esc1(13)"
      Tab(6).Control(27)=   "esc1(14)"
      Tab(6).Control(28)=   "esc1(15)"
      Tab(6).Control(29)=   "esc1(16)"
      Tab(6).Control(30)=   "esc1(17)"
      Tab(6).Control(31)=   "esc1(18)"
      Tab(6).Control(32)=   "esc1(19)"
      Tab(6).Control(33)=   "esc1(20)"
      Tab(6).Control(34)=   "esc1(21)"
      Tab(6).Control(35)=   "esc1(22)"
      Tab(6).Control(36)=   "esc1(23)"
      Tab(6).Control(37)=   "esc1(24)"
      Tab(6).Control(38)=   "esc1(25)"
      Tab(6).Control(39)=   "esc1(26)"
      Tab(6).Control(40)=   "esc1(27)"
      Tab(6).Control(41)=   "esc1(28)"
      Tab(6).Control(42)=   "esc1(29)"
      Tab(6).Control(43)=   "esc1(30)"
      Tab(6).Control(44)=   "esc1(31)"
      Tab(6).Control(45)=   "esc1(32)"
      Tab(6).Control(46)=   "esc1(33)"
      Tab(6).Control(47)=   "esc1(34)"
      Tab(6).Control(48)=   "esc1(35)"
      Tab(6).Control(49)=   "esc1(36)"
      Tab(6).Control(50)=   "esc1(37)"
      Tab(6).Control(51)=   "esc1(38)"
      Tab(6).Control(52)=   "esc1(39)"
      Tab(6).Control(53)=   "esc1(40)"
      Tab(6).Control(54)=   "esc1(41)"
      Tab(6).Control(55)=   "esc1(42)"
      Tab(6).Control(56)=   "esc1(43)"
      Tab(6).Control(57)=   "esc1(44)"
      Tab(6).Control(58)=   "esc1(45)"
      Tab(6).Control(59)=   "esc1(46)"
      Tab(6).Control(60)=   "esc1(47)"
      Tab(6).Control(61)=   "esc1(48)"
      Tab(6).Control(62)=   "esc1(49)"
      Tab(6).Control(63)=   "esc1(50)"
      Tab(6).Control(64)=   "esc1(51)"
      Tab(6).Control(65)=   "esc1(52)"
      Tab(6).Control(66)=   "esc1(53)"
      Tab(6).Control(67)=   "esc1(54)"
      Tab(6).Control(68)=   "esc1(55)"
      Tab(6).Control(69)=   "esc1(56)"
      Tab(6).Control(70)=   "esc1(57)"
      Tab(6).Control(71)=   "esc1(58)"
      Tab(6).Control(72)=   "esc1(59)"
      Tab(6).Control(73)=   "esc1(60)"
      Tab(6).Control(74)=   "esc1(61)"
      Tab(6).Control(75)=   "esc1(62)"
      Tab(6).Control(76)=   "esc1(63)"
      Tab(6).Control(77)=   "esc1(64)"
      Tab(6).Control(78)=   "TXTFIELDS(37)"
      Tab(6).Control(79)=   "TXTFIELDS(38)"
      Tab(6).Control(80)=   "TXTFIELDS(39)"
      Tab(6).Control(81)=   "TXTFIELDS(40)"
      Tab(6).Control(82)=   "TXTFIELDS(49)"
      Tab(6).Control(83)=   "TXTFIELDS(50)"
      Tab(6).Control(84)=   "TXTFIELDS(1)"
      Tab(6).Control(85)=   "TXTFIELDS(22)"
      Tab(6).Control(86)=   "Label17"
      Tab(6).Control(87)=   "Label16"
      Tab(6).Control(88)=   "Label15"
      Tab(6).Control(89)=   "Label5"
      Tab(6).Control(90)=   "Label14"
      Tab(6).Control(91)=   "Label7"
      Tab(6).Control(92)=   "Label8"
      Tab(6).Control(93)=   "Label9"
      Tab(6).Control(94)=   "Label6"
      Tab(6).Control(95)=   "Label10"
      Tab(6).Control(96)=   "Label11"
      Tab(6).Control(97)=   "Label12"
      TabCaption(7)   =   "Figura Ferramenta"
      Tab(7).ControlCount=   16
      Tab(7).Control(0)=   "Picture1(3)"
      Tab(7).Control(1)=   "Picture2(3)"
      Tab(7).Control(2)=   "Picture2(1)"
      Tab(7).Control(3)=   "Picture1(1)"
      Tab(7).Control(4)=   "Incimg(1)"
      Tab(7).Control(5)=   "Incimg(3)"
      Tab(7).Control(6)=   "DelImg(1)"
      Tab(7).Control(7)=   "DelImg(3)"
      Tab(7).Control(8)=   "commandx(1)"
      Tab(7).Control(9)=   "commandx(3)"
      Tab(7).Control(10)=   "cmdimprimir(1)"
      Tab(7).Control(11)=   "cmdimprimir(3)"
      Tab(7).Control(12)=   "CmdPaste(1)"
      Tab(7).Control(13)=   "CmdPaste(3)"
      Tab(7).Control(14)=   "VerImg(1)"
      Tab(7).Control(15)=   "VerImg(3)"
      TabCaption(8)   =   "Embalagem"
      Tab(8).ControlCount=   18
      Tab(8).Control(0)=   "ESCpro(1)"
      Tab(8).Control(1)=   "Cmdbusmr"
      Tab(8).Control(2)=   "Cmdescmr"
      Tab(8).Control(3)=   "TXTFIELDS(71)"
      Tab(8).Control(4)=   "TXTFIELDS(70)"
      Tab(8).Control(5)=   "TXTFIELDS(69)"
      Tab(8).Control(6)=   "TXTFIELDS(68)"
      Tab(8).Control(7)=   "Picture1(2)"
      Tab(8).Control(8)=   "Picture2(2)"
      Tab(8).Control(9)=   "Incimg(2)"
      Tab(8).Control(10)=   "DelImg(2)"
      Tab(8).Control(11)=   "commandx(2)"
      Tab(8).Control(12)=   "cmdimprimir(2)"
      Tab(8).Control(13)=   "CmdPaste(2)"
      Tab(8).Control(14)=   "VerImg(2)"
      Tab(8).Control(15)=   "lblLabels(25)"
      Tab(8).Control(16)=   "lblLabels(24)"
      Tab(8).Control(17)=   "lblLabels(22)"
      TabCaption(9)   =   "..."
      Tab(9).ControlCount=   0
      TabCaption(10)  =   "Vista Explodida"
      Tab(10).ControlCount=   37
      Tab(10).Control(0)=   "TXTFIELDS(95)"
      Tab(10).Control(1)=   "TXTFIELDS(94)"
      Tab(10).Control(2)=   "TXTFIELDS(93)"
      Tab(10).Control(3)=   "TXTFIELDS(92)"
      Tab(10).Control(4)=   "TXTFIELDS(91)"
      Tab(10).Control(5)=   "TXTFIELDS(90)"
      Tab(10).Control(6)=   "TXTFIELDS(80)"
      Tab(10).Control(7)=   "TXTFIELDS(79)"
      Tab(10).Control(8)=   "TXTFIELDS(78)"
      Tab(10).Control(9)=   "TXTFIELDS(77)"
      Tab(10).Control(10)=   "TXTFIELDS(76)"
      Tab(10).Control(11)=   "TXTFIELDS(75)"
      Tab(10).Control(12)=   "TXTFIELDS(74)"
      Tab(10).Control(13)=   "TXTFIELDS(73)"
      Tab(10).Control(14)=   "TXTFIELDS(72)"
      Tab(10).Control(15)=   "Picture1(4)"
      Tab(10).Control(16)=   "Incimg(4)"
      Tab(10).Control(17)=   "DelImg(4)"
      Tab(10).Control(18)=   "commandx(4)"
      Tab(10).Control(19)=   "cmdimprimir(4)"
      Tab(10).Control(20)=   "CmdPaste(4)"
      Tab(10).Control(21)=   "VerImg(4)"
      Tab(10).Control(22)=   "lblLabels(41)"
      Tab(10).Control(23)=   "lblLabels(38)"
      Tab(10).Control(24)=   "lblLabels(37)"
      Tab(10).Control(25)=   "lblLabels(36)"
      Tab(10).Control(26)=   "lblLabels(35)"
      Tab(10).Control(27)=   "lblLabels(34)"
      Tab(10).Control(28)=   "lblLabels(33)"
      Tab(10).Control(29)=   "lblLabels(32)"
      Tab(10).Control(30)=   "lblLabels(31)"
      Tab(10).Control(31)=   "lblLabels(30)"
      Tab(10).Control(32)=   "lblLabels(29)"
      Tab(10).Control(33)=   "lblLabels(28)"
      Tab(10).Control(34)=   "lblLabels(27)"
      Tab(10).Control(35)=   "lblLabels(26)"
      Tab(10).Control(36)=   "lblLabels(23)"
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   95
         Left            =   -70080
         MaxLength       =   50
         TabIndex        =   358
         Top             =   4320
         Width           =   5235
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   94
         Left            =   -70080
         MaxLength       =   50
         TabIndex        =   357
         Top             =   4080
         Width           =   5235
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   93
         Left            =   -70080
         MaxLength       =   50
         TabIndex        =   356
         Top             =   3840
         Width           =   5235
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   92
         Left            =   -70080
         MaxLength       =   50
         TabIndex        =   355
         Top             =   3600
         Width           =   5235
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   91
         Left            =   -70080
         MaxLength       =   50
         TabIndex        =   354
         Top             =   3360
         Width           =   5235
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   90
         Left            =   -70080
         MaxLength       =   50
         TabIndex        =   353
         Top             =   3120
         Width           =   5235
      End
      Begin VB.CommandButton CmdRenumera 
         Caption         =   "Renumerar"
         Height          =   375
         Left            =   -66240
         TabIndex        =   316
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton escArranjo 
         Caption         =   "LXR"
         Height          =   255
         Index           =   2
         Left            =   -70320
         TabIndex        =   315
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton CmdEscMe01 
         Caption         =   "..E"
         Height          =   255
         Left            =   -70800
         TabIndex        =   314
         ToolTipText     =   "Escolhe por Equipamento"
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton escArranjo 
         Caption         =   "LXE"
         Height          =   255
         Index           =   1
         Left            =   -70800
         TabIndex        =   313
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton escRecursoHomem 
         Caption         =   ".."
         Height          =   255
         Index           =   13
         Left            =   -70800
         TabIndex        =   312
         Top             =   2880
         Width           =   375
      End
      Begin VB.CommandButton escRecursoHomem 
         Caption         =   ".."
         Height          =   255
         Index           =   12
         Left            =   -70800
         TabIndex        =   311
         Top             =   2520
         Width           =   375
      End
      Begin VB.CommandButton escRecursoHomem 
         Caption         =   ".."
         Height          =   255
         Index           =   11
         Left            =   -70800
         TabIndex        =   310
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton escRecursoHomem 
         Caption         =   ".."
         Height          =   255
         Index           =   10
         Left            =   -70800
         TabIndex        =   309
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton escRecursoHomem 
         Caption         =   "LX"
         Height          =   255
         Index           =   3
         Left            =   -71160
         TabIndex        =   308
         Top             =   2880
         Width           =   375
      End
      Begin VB.CommandButton escRecursoHomem 
         Caption         =   "LX"
         Height          =   255
         Index           =   2
         Left            =   -71160
         TabIndex        =   307
         Top             =   2520
         Width           =   375
      End
      Begin VB.CommandButton escRecursoHomem 
         Caption         =   "LX"
         Height          =   255
         Index           =   1
         Left            =   -71160
         TabIndex        =   306
         Top             =   2160
         Width           =   375
      End
      Begin VB.CommandButton escRecursoHomem 
         Caption         =   "LX"
         Height          =   255
         Index           =   0
         Left            =   -71160
         TabIndex        =   305
         Top             =   1800
         Width           =   375
      End
      Begin VB.CommandButton escArranjo 
         Caption         =   "LXA"
         Height          =   255
         Index           =   0
         Left            =   -71280
         TabIndex        =   304
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   89
         Left            =   -71880
         MaxLength       =   2
         TabIndex        =   303
         Top             =   2880
         Width           =   615
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   88
         Left            =   -71880
         MaxLength       =   2
         TabIndex        =   302
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   87
         Left            =   -71880
         MaxLength       =   2
         TabIndex        =   301
         Top             =   2160
         Width           =   615
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   86
         Left            =   -71880
         MaxLength       =   2
         TabIndex        =   300
         Top             =   1800
         Width           =   615
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   85
         Left            =   -71880
         MaxLength       =   2
         TabIndex        =   299
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton ESCpro 
         Caption         =   "LX"
         Height          =   375
         Index           =   1
         Left            =   -66360
         TabIndex        =   298
         TabStop         =   0   'False
         Top             =   1080
         Width           =   315
      End
      Begin VB.CommandButton Cmdbusmr 
         Caption         =   "-->"
         Height          =   375
         Left            =   -66600
         TabIndex        =   297
         TabStop         =   0   'False
         Top             =   1080
         Width           =   315
      End
      Begin VB.CommandButton Cmdescmr 
         Height          =   375
         Left            =   -66960
         Picture         =   "frmPFS.frx":0BB4
         Style           =   1  'Graphical
         TabIndex        =   296
         TabStop         =   0   'False
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   84
         Left            =   -68640
         MaxLength       =   24
         TabIndex        =   0
         Top             =   4200
         Width           =   1995
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   83
         Left            =   -68640
         MaxLength       =   24
         TabIndex        =   295
         Top             =   3840
         Width           =   1995
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   82
         Left            =   -74160
         MaxLength       =   24
         TabIndex        =   294
         Top             =   4200
         Width           =   1995
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   81
         Left            =   -74760
         MaxLength       =   30
         TabIndex        =   293
         Text            =   "1234567890123456789012345ABCDE"
         Top             =   3720
         Width           =   3435
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   80
         Left            =   -70080
         MaxLength       =   50
         TabIndex        =   279
         Top             =   2880
         Width           =   5235
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   79
         Left            =   -70080
         MaxLength       =   50
         TabIndex        =   278
         Top             =   2640
         Width           =   5235
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   78
         Left            =   -70080
         MaxLength       =   50
         TabIndex        =   277
         Top             =   2400
         Width           =   5235
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   77
         Left            =   -70080
         MaxLength       =   50
         TabIndex        =   276
         Top             =   2160
         Width           =   5235
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   76
         Left            =   -70080
         MaxLength       =   50
         TabIndex        =   275
         Top             =   1920
         Width           =   5235
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   75
         Left            =   -70080
         MaxLength       =   50
         TabIndex        =   274
         Top             =   1680
         Width           =   5235
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   74
         Left            =   -70080
         MaxLength       =   50
         TabIndex        =   273
         Top             =   1440
         Width           =   5235
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   73
         Left            =   -70080
         MaxLength       =   50
         TabIndex        =   272
         Top             =   1200
         Width           =   5235
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   72
         Left            =   -70080
         MaxLength       =   50
         TabIndex        =   271
         Top             =   960
         Width           =   5235
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   71
         Left            =   -69840
         MaxLength       =   24
         TabIndex        =   270
         Top             =   3000
         Width           =   4515
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   70
         Left            =   -68880
         MaxLength       =   24
         TabIndex        =   269
         Top             =   2280
         Width           =   1035
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   69
         Left            =   -69720
         MaxLength       =   24
         TabIndex        =   268
         Top             =   1560
         Width           =   4515
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   68
         Left            =   -69000
         MaxLength       =   24
         TabIndex        =   267
         Top             =   1200
         Width           =   1995
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Index           =   3
         Left            =   -68040
         ScaleHeight     =   435
         ScaleWidth      =   195
         TabIndex        =   263
         Top             =   3600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   2535
         Index           =   3
         Left            =   -70200
         ScaleHeight     =   2475
         ScaleWidth      =   2835
         TabIndex        =   262
         Top             =   960
         Width           =   2895
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Index           =   4
         Left            =   -72480
         ScaleHeight     =   435
         ScaleWidth      =   195
         TabIndex        =   261
         Top             =   3600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Index           =   2
         Left            =   -72720
         ScaleHeight     =   435
         ScaleWidth      =   195
         TabIndex        =   260
         Top             =   3840
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   2535
         Index           =   4
         Left            =   120
         ScaleHeight     =   2475
         ScaleWidth      =   2835
         TabIndex        =   259
         Top             =   960
         Width           =   2895
      End
      Begin VB.PictureBox Picture2 
         Height          =   2535
         Index           =   2
         Left            =   -74880
         ScaleHeight     =   2475
         ScaleWidth      =   2835
         TabIndex        =   258
         Top             =   1200
         Width           =   2895
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   65
         Left            =   -70680
         TabIndex        =   257
         Top             =   2370
         Width           =   1215
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   12
         Left            =   -72000
         TabIndex        =   255
         Top             =   1770
         Width           =   1215
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   66
         Left            =   -72000
         TabIndex        =   251
         Top             =   2370
         Width           =   1215
      End
      Begin VB.CheckBox EPI07 
         Caption         =   "Macara e Capuz"
         Height          =   255
         Left            =   -71160
         TabIndex        =   247
         Top             =   3090
         Width           =   2175
      End
      Begin VB.CheckBox EPI06 
         Caption         =   "Mangote"
         Height          =   255
         Left            =   -73440
         TabIndex        =   246
         Top             =   3090
         Width           =   1215
      End
      Begin VB.CheckBox EPI05 
         Caption         =   "Avental"
         Height          =   255
         Left            =   -74880
         TabIndex        =   245
         Top             =   3090
         Width           =   1215
      End
      Begin VB.CheckBox EPI04 
         Caption         =   "Oculos"
         Height          =   255
         Left            =   -68760
         TabIndex        =   244
         Top             =   2730
         Width           =   1215
      End
      Begin VB.CheckBox EPI03 
         Caption         =   "Protetor Auricular"
         Height          =   255
         Left            =   -71160
         TabIndex        =   243
         Top             =   2730
         Width           =   1815
      End
      Begin VB.CheckBox EPI02 
         Caption         =   "Calcado Seguranca"
         Height          =   375
         Left            =   -73440
         TabIndex        =   242
         Top             =   2730
         Width           =   1935
      End
      Begin VB.CheckBox EPI01 
         Caption         =   "Luvas"
         Height          =   255
         Left            =   -74880
         TabIndex        =   241
         Top             =   2730
         Width           =   1215
      End
      Begin VB.CommandButton EscCodFlu 
         Caption         =   "..."
         Height          =   255
         Index           =   13
         Left            =   -66720
         TabIndex        =   240
         Top             =   3930
         Width           =   375
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   255
         Index           =   13
         Left            =   -66720
         Picture         =   "frmPFS.frx":113E
         Style           =   1  'Graphical
         TabIndex        =   239
         TabStop         =   0   'False
         Top             =   3690
         Width           =   255
      End
      Begin VB.CommandButton EscCodFlu 
         Caption         =   "..."
         Height          =   255
         Index           =   12
         Left            =   -66720
         TabIndex        =   238
         Top             =   3330
         Width           =   375
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   255
         Index           =   12
         Left            =   -66720
         Picture         =   "frmPFS.frx":16C8
         Style           =   1  'Graphical
         TabIndex        =   237
         TabStop         =   0   'False
         Top             =   3090
         Width           =   255
      End
      Begin VB.TextBox TXTFIELDS 
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
         Index           =   63
         Left            =   -67320
         TabIndex        =   236
         Top             =   3690
         Width           =   615
      End
      Begin VB.TextBox TXTFIELDS 
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
         Index           =   62
         Left            =   -67320
         TabIndex        =   235
         Top             =   3090
         Width           =   615
      End
      Begin VB.PictureBox Picture2 
         Height          =   2535
         Index           =   1
         Left            =   -74880
         ScaleHeight     =   2475
         ScaleWidth      =   2835
         TabIndex        =   234
         Top             =   960
         Width           =   2895
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Index           =   1
         Left            =   -72720
         ScaleHeight     =   435
         ScaleWidth      =   195
         TabIndex        =   233
         Top             =   3600
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   2535
         Index           =   0
         Left            =   -74760
         ScaleHeight     =   2475
         ScaleWidth      =   2835
         TabIndex        =   232
         Top             =   960
         Width           =   2895
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Index           =   0
         Left            =   -72480
         ScaleHeight     =   435
         ScaleWidth      =   315
         TabIndex        =   231
         Top             =   3600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   41
         Left            =   -74760
         MaxLength       =   24
         TabIndex        =   222
         Top             =   1320
         Width           =   1995
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   42
         Left            =   -74760
         MaxLength       =   24
         TabIndex        =   221
         Top             =   1560
         Width           =   1995
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   43
         Left            =   -74760
         MaxLength       =   24
         TabIndex        =   220
         Top             =   1800
         Width           =   1995
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   44
         Left            =   -74760
         MaxLength       =   24
         TabIndex        =   219
         Top             =   2040
         Width           =   1995
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   45
         Left            =   -74760
         MaxLength       =   24
         TabIndex        =   218
         Top             =   2280
         Width           =   1995
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   46
         Left            =   -74760
         MaxLength       =   24
         TabIndex        =   217
         Top             =   2520
         Width           =   1995
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   47
         Left            =   -74760
         MaxLength       =   24
         TabIndex        =   216
         Top             =   2760
         Width           =   1995
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   48
         Left            =   -74760
         MaxLength       =   24
         TabIndex        =   215
         Top             =   3000
         Width           =   1995
      End
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   36
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   214
         TabStop         =   0   'False
         Top             =   960
         Width           =   495
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Tipo Almofados 4x6/8x8"
         Height          =   255
         Left            =   -74160
         TabIndex        =   213
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -72720
         TabIndex        =   212
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -72480
         TabIndex        =   211
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -72240
         TabIndex        =   210
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -72000
         TabIndex        =   209
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   -71760
         TabIndex        =   208
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   -71520
         TabIndex        =   207
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   -71280
         TabIndex        =   206
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   -71040
         TabIndex        =   205
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   -72720
         TabIndex        =   204
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   -72480
         TabIndex        =   203
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   -72240
         TabIndex        =   202
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   -72000
         TabIndex        =   201
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   -71760
         TabIndex        =   200
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   -71520
         TabIndex        =   199
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   -71280
         TabIndex        =   198
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   -71040
         TabIndex        =   197
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   -72720
         TabIndex        =   196
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   -72480
         TabIndex        =   195
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   -72240
         TabIndex        =   194
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   -72000
         TabIndex        =   193
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   -71760
         TabIndex        =   192
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   -71520
         TabIndex        =   191
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   -71280
         TabIndex        =   190
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   24
         Left            =   -71040
         TabIndex        =   189
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   25
         Left            =   -72720
         TabIndex        =   188
         Top             =   2040
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   -72480
         TabIndex        =   187
         Top             =   2040
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   27
         Left            =   -72240
         TabIndex        =   186
         Top             =   2040
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   28
         Left            =   -72000
         TabIndex        =   185
         Top             =   2040
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   29
         Left            =   -71760
         TabIndex        =   184
         Top             =   2040
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   30
         Left            =   -71520
         TabIndex        =   183
         Top             =   2040
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   31
         Left            =   -71280
         TabIndex        =   182
         Top             =   2040
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   32
         Left            =   -71040
         TabIndex        =   181
         Top             =   2040
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   33
         Left            =   -72720
         TabIndex        =   180
         Top             =   2280
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   34
         Left            =   -72480
         TabIndex        =   179
         Top             =   2280
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   35
         Left            =   -72240
         TabIndex        =   178
         Top             =   2280
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   36
         Left            =   -72000
         TabIndex        =   177
         Top             =   2280
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   37
         Left            =   -71760
         TabIndex        =   176
         Top             =   2280
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   38
         Left            =   -71520
         TabIndex        =   175
         Top             =   2280
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   39
         Left            =   -71280
         TabIndex        =   174
         Top             =   2280
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   40
         Left            =   -71040
         TabIndex        =   173
         Top             =   2280
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   41
         Left            =   -72720
         TabIndex        =   172
         Top             =   2520
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   42
         Left            =   -72480
         TabIndex        =   171
         Top             =   2520
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   43
         Left            =   -72240
         TabIndex        =   170
         Top             =   2520
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   44
         Left            =   -72000
         TabIndex        =   169
         Top             =   2520
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   45
         Left            =   -71760
         TabIndex        =   168
         Top             =   2520
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   46
         Left            =   -71520
         TabIndex        =   167
         Top             =   2520
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   47
         Left            =   -71280
         TabIndex        =   166
         Top             =   2520
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   48
         Left            =   -71040
         TabIndex        =   165
         Top             =   2520
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   49
         Left            =   -72720
         TabIndex        =   164
         Top             =   2760
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   50
         Left            =   -72480
         TabIndex        =   163
         Top             =   2760
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   51
         Left            =   -72240
         TabIndex        =   162
         Top             =   2760
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   52
         Left            =   -72000
         TabIndex        =   161
         Top             =   2760
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   53
         Left            =   -71760
         TabIndex        =   160
         Top             =   2760
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   54
         Left            =   -71520
         TabIndex        =   159
         Top             =   2760
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   55
         Left            =   -71280
         TabIndex        =   158
         Top             =   2760
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   56
         Left            =   -71040
         TabIndex        =   157
         Top             =   2760
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   57
         Left            =   -72720
         TabIndex        =   156
         Top             =   3000
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   58
         Left            =   -72480
         TabIndex        =   155
         Top             =   3000
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   59
         Left            =   -72240
         TabIndex        =   154
         Top             =   3000
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   60
         Left            =   -72000
         TabIndex        =   153
         Top             =   3000
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   61
         Left            =   -71760
         TabIndex        =   152
         Top             =   3000
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   62
         Left            =   -71520
         TabIndex        =   151
         Top             =   3000
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   63
         Left            =   -71280
         TabIndex        =   150
         Top             =   3000
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "¡"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   64
         Left            =   -71040
         TabIndex        =   149
         Top             =   3000
         Width           =   255
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   37
         Left            =   -68640
         MaxLength       =   24
         TabIndex        =   148
         Top             =   960
         Width           =   1995
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   38
         Left            =   -68640
         MaxLength       =   24
         TabIndex        =   147
         Top             =   1320
         Width           =   1995
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   39
         Left            =   -68640
         MaxLength       =   24
         TabIndex        =   146
         Top             =   2040
         Width           =   1995
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   40
         Left            =   -68640
         MaxLength       =   24
         TabIndex        =   145
         Top             =   2400
         Width           =   1995
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   49
         Left            =   -68640
         MaxLength       =   24
         TabIndex        =   144
         Top             =   2760
         Width           =   1995
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   50
         Left            =   -68640
         MaxLength       =   24
         TabIndex        =   143
         Top             =   3120
         Width           =   1995
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   -68640
         MaxLength       =   24
         TabIndex        =   142
         Top             =   3480
         Width           =   1995
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   22
         Left            =   -68640
         MaxLength       =   24
         TabIndex        =   141
         Top             =   1680
         Width           =   1995
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Novo"
         Height          =   375
         Left            =   8760
         TabIndex        =   139
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Apaga"
         Height          =   375
         Left            =   8760
         TabIndex        =   138
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Editar"
         Height          =   375
         Left            =   8760
         TabIndex        =   137
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Importar"
         Height          =   375
         Left            =   8760
         TabIndex        =   136
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton InsEdit 
         Caption         =   "Editar"
         Height          =   375
         Left            =   -66240
         TabIndex        =   134
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Apaga"
         Height          =   375
         Left            =   -66240
         TabIndex        =   133
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Novo"
         Height          =   375
         Index           =   1
         Left            =   -66240
         TabIndex        =   132
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Cmdimp 
         Caption         =   "Importar"
         Height          =   375
         Left            =   -66240
         TabIndex        =   131
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   1125
         Index           =   26
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   128
         Top             =   1530
         Width           =   7935
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   1125
         Index           =   27
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   127
         Top             =   3690
         Width           =   7935
      End
      Begin VB.CommandButton Cmdepi 
         Caption         =   "EPI"
         Height          =   255
         Left            =   -66720
         TabIndex        =   126
         ToolTipText     =   "inclui Mensagem Padrao de EPI"
         Top             =   1770
         Width           =   495
      End
      Begin VB.CommandButton Cmdapgobs 
         Caption         =   "Apagar"
         Height          =   255
         Left            =   -66720
         TabIndex        =   125
         ToolTipText     =   "Apagar as Observacoes"
         Top             =   2130
         Width           =   735
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Apagar"
         Height          =   255
         Left            =   -66840
         TabIndex        =   124
         ToolTipText     =   "Apagar as Observacoes Regulagem Maquina"
         Top             =   3690
         Width           =   735
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   13
         Left            =   -69240
         TabIndex        =   115
         Top             =   1770
         Width           =   1335
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   14
         Left            =   -72000
         MaxLength       =   24
         TabIndex        =   114
         Top             =   3210
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "...."
         Height          =   255
         Index           =   0
         Left            =   -68880
         TabIndex        =   113
         Top             =   3210
         Width           =   375
      End
      Begin VB.TextBox TXTFIELDS 
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
         Left            =   -74160
         TabIndex        =   112
         Top             =   1770
         Width           =   615
      End
      Begin VB.TextBox TXTFIELDS 
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
         Left            =   -73560
         TabIndex        =   111
         Top             =   1770
         Width           =   615
      End
      Begin VB.TextBox TXTFIELDS 
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
         Left            =   -74160
         TabIndex        =   110
         Top             =   2250
         Width           =   615
      End
      Begin VB.TextBox TXTFIELDS 
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
         Left            =   -73560
         TabIndex        =   109
         Top             =   2250
         Width           =   615
      End
      Begin VB.TextBox TXTFIELDS 
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
         Left            =   -74160
         TabIndex        =   108
         Top             =   2730
         Width           =   615
      End
      Begin VB.TextBox TXTFIELDS 
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
         Left            =   -73560
         TabIndex        =   107
         Top             =   2730
         Width           =   615
      End
      Begin VB.CommandButton EscCodFlu 
         Caption         =   "..."
         Height          =   255
         Index           =   1
         Left            =   -74640
         TabIndex        =   106
         Top             =   2010
         Width           =   375
      End
      Begin VB.CommandButton EscCodFlu 
         Caption         =   "..."
         Height          =   255
         Index           =   2
         Left            =   -72840
         TabIndex        =   105
         Top             =   2010
         Width           =   375
      End
      Begin VB.CommandButton EscCodFlu 
         Caption         =   "..."
         Height          =   255
         Index           =   3
         Left            =   -74640
         TabIndex        =   104
         Top             =   2490
         Width           =   375
      End
      Begin VB.CommandButton EscCodFlu 
         Caption         =   "..."
         Height          =   255
         Index           =   4
         Left            =   -72840
         TabIndex        =   103
         Top             =   2490
         Width           =   375
      End
      Begin VB.CommandButton EscCodFlu 
         Caption         =   "..."
         Height          =   255
         Index           =   5
         Left            =   -74640
         TabIndex        =   102
         Top             =   2970
         Width           =   375
      End
      Begin VB.CommandButton EscCodFlu 
         Caption         =   "..."
         Height          =   255
         Index           =   6
         Left            =   -72840
         TabIndex        =   101
         Top             =   2970
         Width           =   375
      End
      Begin VB.TextBox TXTFIELDS 
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
         Left            =   -74160
         TabIndex        =   100
         Top             =   3210
         Width           =   615
      End
      Begin VB.TextBox TXTFIELDS 
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
         Left            =   -73560
         TabIndex        =   99
         Top             =   3210
         Width           =   615
      End
      Begin VB.TextBox TXTFIELDS 
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
         Left            =   -74160
         TabIndex        =   98
         Top             =   3690
         Width           =   615
      End
      Begin VB.TextBox TXTFIELDS 
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
         Left            =   -73560
         TabIndex        =   97
         Top             =   3690
         Width           =   615
      End
      Begin VB.CommandButton EscCodFlu 
         Caption         =   "..."
         Height          =   255
         Index           =   7
         Left            =   -74640
         TabIndex        =   96
         Top             =   3450
         Width           =   375
      End
      Begin VB.CommandButton EscCodFlu 
         Caption         =   "..."
         Height          =   255
         Index           =   8
         Left            =   -72840
         TabIndex        =   95
         Top             =   3450
         Width           =   375
      End
      Begin VB.CommandButton EscCodFlu 
         Caption         =   "..."
         Height          =   255
         Index           =   9
         Left            =   -74640
         TabIndex        =   94
         Top             =   3930
         Width           =   375
      End
      Begin VB.CommandButton EscCodFlu 
         Caption         =   "..."
         Height          =   255
         Index           =   10
         Left            =   -72840
         TabIndex        =   93
         Top             =   3930
         Width           =   375
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   20
         Left            =   -73320
         MaxLength       =   24
         TabIndex        =   92
         Top             =   4410
         Width           =   3015
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   21
         Left            =   -69840
         MaxLength       =   24
         TabIndex        =   91
         Top             =   4410
         Width           =   4335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "...."
         Height          =   255
         Left            =   -70320
         TabIndex        =   90
         Top             =   4290
         Width           =   375
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   25
         Left            =   -70680
         TabIndex        =   89
         Top             =   1770
         Width           =   1335
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   28
         Left            =   -72000
         MaxLength       =   24
         TabIndex        =   88
         Top             =   3450
         Width           =   3015
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   29
         Left            =   -72000
         MaxLength       =   24
         TabIndex        =   87
         Top             =   3690
         Width           =   3015
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   30
         Left            =   -72000
         MaxLength       =   24
         TabIndex        =   86
         Top             =   3930
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         Caption         =   "...."
         Height          =   255
         Index           =   1
         Left            =   -68880
         TabIndex        =   85
         Top             =   3450
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "...."
         Height          =   255
         Index           =   2
         Left            =   -68880
         TabIndex        =   84
         Top             =   3690
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "...."
         Height          =   255
         Index           =   3
         Left            =   -68880
         TabIndex        =   83
         Top             =   3930
         Width           =   375
      End
      Begin VB.CommandButton Command9 
         Caption         =   "-->"
         Height          =   255
         Index           =   6
         Left            =   -70320
         TabIndex        =   82
         Top             =   4530
         Width           =   315
      End
      Begin VB.CommandButton Edifer 
         Caption         =   "E.."
         Height          =   255
         Index           =   0
         Left            =   -68400
         TabIndex        =   81
         ToolTipText     =   "Editar Cadastro Ferramenta"
         Top             =   3210
         Width           =   375
      End
      Begin VB.CommandButton Edifer 
         Caption         =   "E.."
         Height          =   255
         Index           =   1
         Left            =   -68400
         TabIndex        =   80
         ToolTipText     =   "Editar Cadastro Ferramenta"
         Top             =   3450
         Width           =   375
      End
      Begin VB.CommandButton Edifer 
         Caption         =   "E.."
         Height          =   255
         Index           =   2
         Left            =   -68400
         TabIndex        =   79
         ToolTipText     =   "Editar Cadastro Ferramenta"
         Top             =   3690
         Width           =   375
      End
      Begin VB.CommandButton Edifer 
         Caption         =   "E.."
         Height          =   255
         Index           =   3
         Left            =   -68400
         TabIndex        =   78
         ToolTipText     =   "Editar Cadastro Ferramenta"
         Top             =   3930
         Width           =   375
      End
      Begin VB.CommandButton CmdForFer 
         Caption         =   "F."
         Height          =   255
         Index           =   0
         Left            =   -67920
         TabIndex        =   77
         ToolTipText     =   "Alterar Foto Ferramenta"
         Top             =   3210
         Width           =   315
      End
      Begin VB.CommandButton CmdForFer 
         Caption         =   "F."
         Height          =   255
         Index           =   1
         Left            =   -67920
         TabIndex        =   76
         ToolTipText     =   "Alterar Foto Ferramenta"
         Top             =   3450
         Width           =   315
      End
      Begin VB.CommandButton CmdForFer 
         Caption         =   "F."
         Height          =   255
         Index           =   2
         Left            =   -67920
         TabIndex        =   75
         ToolTipText     =   "Alterar Foto Ferramenta"
         Top             =   3690
         Width           =   315
      End
      Begin VB.CommandButton CmdForFer 
         Caption         =   "F."
         Height          =   255
         Index           =   3
         Left            =   -67920
         TabIndex        =   74
         ToolTipText     =   "Alterar Foto Ferramenta"
         Top             =   3930
         Width           =   315
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   255
         Index           =   1
         Left            =   -74880
         Picture         =   "frmPFS.frx":1C52
         Style           =   1  'Graphical
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   2010
         Width           =   255
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   255
         Index           =   2
         Left            =   -72480
         Picture         =   "frmPFS.frx":21DC
         Style           =   1  'Graphical
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   2010
         Width           =   255
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   255
         Index           =   3
         Left            =   -74880
         Picture         =   "frmPFS.frx":2766
         Style           =   1  'Graphical
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   2490
         Width           =   255
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   255
         Index           =   4
         Left            =   -72480
         Picture         =   "frmPFS.frx":2CF0
         Style           =   1  'Graphical
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   2490
         Width           =   255
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   255
         Index           =   5
         Left            =   -74880
         Picture         =   "frmPFS.frx":327A
         Style           =   1  'Graphical
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   2970
         Width           =   255
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   255
         Index           =   6
         Left            =   -72480
         Picture         =   "frmPFS.frx":3804
         Style           =   1  'Graphical
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   2970
         Width           =   255
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   255
         Index           =   7
         Left            =   -74880
         Picture         =   "frmPFS.frx":3D8E
         Style           =   1  'Graphical
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   3450
         Width           =   255
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   255
         Index           =   8
         Left            =   -72480
         Picture         =   "frmPFS.frx":4318
         Style           =   1  'Graphical
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   3450
         Width           =   255
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   255
         Index           =   9
         Left            =   -74880
         Picture         =   "frmPFS.frx":48A2
         Style           =   1  'Graphical
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   3930
         Width           =   255
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   255
         Index           =   10
         Left            =   -72480
         Picture         =   "frmPFS.frx":4E2C
         Style           =   1  'Graphical
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   3930
         Width           =   255
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   255
         Index           =   11
         Left            =   -66720
         Picture         =   "frmPFS.frx":53B6
         Style           =   1  'Graphical
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   2490
         Width           =   255
      End
      Begin VB.CommandButton EscCodFlu 
         Caption         =   "..."
         Height          =   255
         Index           =   11
         Left            =   -66720
         TabIndex        =   62
         Top             =   2730
         Width           =   375
      End
      Begin VB.TextBox TXTFIELDS 
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
         Index           =   61
         Left            =   -67320
         TabIndex        =   61
         Top             =   2490
         Width           =   615
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   19
         Left            =   -67800
         TabIndex        =   60
         Top             =   1770
         Width           =   1335
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   8
         Left            =   -73920
         MaxLength       =   12
         TabIndex        =   50
         Top             =   1410
         Width           =   1935
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   9
         Left            =   -73920
         MaxLength       =   12
         TabIndex        =   49
         Top             =   1770
         Width           =   1935
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   10
         Left            =   -73920
         MaxLength       =   12
         TabIndex        =   48
         Top             =   2130
         Width           =   1935
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   11
         Left            =   -73920
         MaxLength       =   12
         TabIndex        =   47
         Top             =   3210
         Width           =   1935
      End
      Begin VB.CommandButton EscMP01 
         Caption         =   "..G"
         Height          =   255
         Left            =   -71160
         TabIndex        =   46
         ToolTipText     =   "Escolhe por Grupo de Equipamento"
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton ESCMP03 
         Caption         =   "...."
         Height          =   255
         Left            =   -71940
         TabIndex        =   45
         Top             =   3210
         Width           =   255
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   4
         Left            =   -69960
         MaxLength       =   30
         TabIndex        =   44
         Text            =   "123456789012345678901234567890"
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   5
         Left            =   -69960
         MaxLength       =   30
         TabIndex        =   43
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   7
         Left            =   -69960
         MaxLength       =   30
         TabIndex        =   42
         Top             =   2130
         Width           =   3375
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   6
         Left            =   -69960
         MaxLength       =   30
         TabIndex        =   41
         Top             =   3240
         Width           =   3375
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   15
         Left            =   -73920
         MaxLength       =   12
         TabIndex        =   40
         Top             =   2490
         Width           =   1935
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   16
         Left            =   -73920
         MaxLength       =   12
         TabIndex        =   39
         Top             =   2850
         Width           =   1935
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   17
         Left            =   -69960
         MaxLength       =   30
         TabIndex        =   38
         Top             =   2520
         Width           =   3375
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   18
         Left            =   -69960
         MaxLength       =   30
         TabIndex        =   37
         Top             =   2850
         Width           =   3375
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   23
         Left            =   -73920
         MaxLength       =   12
         TabIndex        =   36
         Top             =   3570
         Width           =   1935
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   24
         Left            =   -73920
         MaxLength       =   12
         TabIndex        =   35
         Top             =   3930
         Width           =   1935
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Copia"
         Height          =   315
         Index           =   0
         Left            =   -71880
         TabIndex        =   34
         Top             =   3570
         Width           =   735
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Copia"
         Height          =   315
         Index           =   1
         Left            =   -71880
         TabIndex        =   33
         Top             =   3930
         Width           =   735
      End
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   405
         Index           =   31
         Left            =   -74880
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton Command8 
         Caption         =   "..."
         Height          =   255
         Left            =   -74280
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2160
         Width           =   255
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   32
         Left            =   -69960
         MaxLength       =   2
         TabIndex        =   30
         Top             =   3630
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   33
         Left            =   -69960
         MaxLength       =   2
         TabIndex        =   29
         Top             =   4050
         Width           =   495
      End
      Begin VB.CommandButton escsetref 
         Caption         =   "..."
         Height          =   255
         Left            =   -69420
         TabIndex        =   28
         Top             =   3630
         Width           =   315
      End
      Begin VB.CommandButton escsetins 
         Caption         =   "..."
         Height          =   255
         Left            =   -69420
         TabIndex        =   27
         Top             =   4050
         Width           =   315
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   34
         Left            =   -67680
         MaxLength       =   2
         TabIndex        =   26
         Top             =   3630
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   35
         Left            =   -67680
         MaxLength       =   2
         TabIndex        =   25
         Top             =   3990
         Width           =   495
      End
      Begin VB.CommandButton escmon 
         Caption         =   "..."
         Height          =   255
         Left            =   -67020
         TabIndex        =   24
         Top             =   3630
         Width           =   315
      End
      Begin VB.CommandButton escloc 
         Caption         =   "..."
         Height          =   255
         Left            =   -67020
         TabIndex        =   23
         Top             =   3990
         Width           =   315
      End
      Begin VB.CommandButton Command9 
         Caption         =   "-->"
         Height          =   255
         Index           =   0
         Left            =   -70320
         TabIndex        =   22
         Top             =   1440
         Width           =   315
      End
      Begin VB.CommandButton Command9 
         Caption         =   "-->"
         Height          =   255
         Index           =   1
         Left            =   -70320
         TabIndex        =   21
         Top             =   1800
         Width           =   315
      End
      Begin VB.CommandButton Command9 
         Caption         =   "-->"
         Height          =   255
         Index           =   2
         Left            =   -70320
         TabIndex        =   20
         Top             =   2160
         Width           =   315
      End
      Begin VB.CommandButton Command9 
         Caption         =   "-->"
         Height          =   255
         Index           =   3
         Left            =   -70320
         TabIndex        =   19
         Top             =   2520
         Width           =   315
      End
      Begin VB.CommandButton Command9 
         Caption         =   "-->"
         Height          =   255
         Index           =   4
         Left            =   -70320
         TabIndex        =   18
         Top             =   2880
         Width           =   315
      End
      Begin VB.CommandButton Command9 
         Caption         =   "-->"
         Height          =   255
         Index           =   5
         Left            =   -71700
         TabIndex        =   17
         Top             =   3210
         Width           =   315
      End
      Begin VB.CommandButton Command11 
         Caption         =   "ETI"
         Height          =   255
         Left            =   -74880
         TabIndex        =   16
         ToolTipText     =   "Cadastro de ETI"
         Top             =   4290
         Width           =   495
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Tratamentos"
         Height          =   255
         Left            =   -74040
         TabIndex        =   15
         ToolTipText     =   "Cadastro de Tratamentos"
         Top             =   4290
         Width           =   1155
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "E"
         Height          =   255
         Left            =   -71160
         TabIndex        =   14
         Tag             =   "Editar "
         Top             =   3210
         Width           =   255
      End
      Begin MSFlexGridLib.MSFlexGrid GridIns 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   135
         Top             =   930
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   6800
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GrdQsbLep 
         Height          =   3915
         Left            =   120
         TabIndex        =   140
         Top             =   930
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   6906
         _Version        =   393216
      End
      Begin XPControls.XPButton Incimg 
         Height          =   435
         Index           =   0
         Left            =   -71640
         TabIndex        =   323
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFS.frx":5940
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
         Index           =   1
         Left            =   -71880
         TabIndex        =   324
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFS.frx":5DDA
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
         Left            =   -71760
         TabIndex        =   325
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFS.frx":6274
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
         Index           =   3
         Left            =   -67200
         TabIndex        =   326
         TabStop         =   0   'False
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFS.frx":670E
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
         Index           =   4
         Left            =   -71880
         TabIndex        =   327
         TabStop         =   0   'False
         Top             =   960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   767
         Picture         =   "frmPFS.frx":6BA8
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
         Index           =   0
         Left            =   -71640
         TabIndex        =   328
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFS.frx":7042
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
         Index           =   1
         Left            =   -71880
         TabIndex        =   329
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFS.frx":74DC
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
         Left            =   -71760
         TabIndex        =   330
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFS.frx":7976
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
         Index           =   3
         Left            =   -67200
         TabIndex        =   331
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFS.frx":7E10
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
         Index           =   4
         Left            =   -71880
         TabIndex        =   332
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   767
         Picture         =   "frmPFS.frx":82AA
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
      Begin XPControls.XPButton commandx 
         Height          =   435
         Index           =   0
         Left            =   -71640
         TabIndex        =   333
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFS.frx":8744
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
      Begin XPControls.XPButton commandx 
         Height          =   435
         Index           =   1
         Left            =   -71880
         TabIndex        =   334
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFS.frx":8BDE
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
      Begin XPControls.XPButton commandx 
         Height          =   435
         Index           =   2
         Left            =   -71760
         TabIndex        =   335
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFS.frx":9078
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
      Begin XPControls.XPButton commandx 
         Height          =   435
         Index           =   3
         Left            =   -67200
         TabIndex        =   336
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFS.frx":9512
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
      Begin XPControls.XPButton commandx 
         Height          =   435
         Index           =   4
         Left            =   -71880
         TabIndex        =   337
         TabStop         =   0   'False
         Top             =   2160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   767
         Picture         =   "frmPFS.frx":99AC
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
      Begin XPControls.XPButton cmdimprimir 
         Height          =   435
         Index           =   0
         Left            =   -71640
         TabIndex        =   338
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFS.frx":9E46
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
         Left            =   -71880
         TabIndex        =   339
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFS.frx":A3E0
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
         Left            =   -71760
         TabIndex        =   340
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFS.frx":A97A
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
         Index           =   3
         Left            =   -67200
         TabIndex        =   341
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFS.frx":AF14
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
         Index           =   4
         Left            =   -71880
         TabIndex        =   342
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   767
         Picture         =   "frmPFS.frx":B4AE
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
         Left            =   -71640
         TabIndex        =   343
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFS.frx":BA48
         Caption         =   "Copia"
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
         Left            =   -71880
         TabIndex        =   344
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFS.frx":BFE2
         Caption         =   "Copia"
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
         Left            =   -71760
         TabIndex        =   345
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFS.frx":C57C
         Caption         =   "Copia"
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
         Index           =   3
         Left            =   -67200
         TabIndex        =   346
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFS.frx":CB16
         Caption         =   "Copia"
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
         Index           =   4
         Left            =   -71880
         TabIndex        =   347
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   767
         Picture         =   "frmPFS.frx":D0B0
         Caption         =   "Copia"
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
         Index           =   0
         Left            =   -71640
         TabIndex        =   348
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmPFS.frx":D64A
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
         Index           =   1
         Left            =   -71880
         TabIndex        =   349
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmPFS.frx":DAE4
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
         Left            =   -71760
         TabIndex        =   350
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmPFS.frx":DF7E
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
         Index           =   3
         Left            =   -67200
         TabIndex        =   351
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmPFS.frx":E418
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
         Index           =   4
         Left            =   -71880
         TabIndex        =   352
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Picture         =   "frmPFS.frx":E8B2
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
      Begin VB.Label lblLabels 
         Caption         =   "15:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   41
         Left            =   -70440
         TabIndex        =   364
         Top             =   4320
         Width           =   315
      End
      Begin VB.Label lblLabels 
         Caption         =   "14:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   38
         Left            =   -70440
         TabIndex        =   363
         Top             =   4080
         Width           =   315
      End
      Begin VB.Label lblLabels 
         Caption         =   "13:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   37
         Left            =   -70440
         TabIndex        =   362
         Top             =   3840
         Width           =   315
      End
      Begin VB.Label lblLabels 
         Caption         =   "12:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   36
         Left            =   -70440
         TabIndex        =   361
         Top             =   3600
         Width           =   315
      End
      Begin VB.Label lblLabels 
         Caption         =   "11:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   35
         Left            =   -70440
         TabIndex        =   360
         Top             =   3360
         Width           =   315
      End
      Begin VB.Label lblLabels 
         Caption         =   "09:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   34
         Left            =   -70440
         TabIndex        =   359
         Top             =   2880
         Width           =   315
      End
      Begin VB.Label Label17 
         Caption         =   "Passo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   292
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "Força de Extracao"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -70200
         TabIndex        =   291
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "Tipo de Fixacao"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -70200
         TabIndex        =   290
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Altura do  Martelo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   289
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Caption         =   "10:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   33
         Left            =   -70440
         TabIndex        =   288
         Top             =   3120
         Width           =   315
      End
      Begin VB.Label lblLabels 
         Caption         =   "08:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   32
         Left            =   -70440
         TabIndex        =   287
         Top             =   2640
         Width           =   315
      End
      Begin VB.Label lblLabels 
         Caption         =   "07:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   31
         Left            =   -70440
         TabIndex        =   286
         Top             =   2400
         Width           =   315
      End
      Begin VB.Label lblLabels 
         Caption         =   "06:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   30
         Left            =   -70440
         TabIndex        =   285
         Top             =   2160
         Width           =   315
      End
      Begin VB.Label lblLabels 
         Caption         =   "05:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   29
         Left            =   -70440
         TabIndex        =   284
         Top             =   1920
         Width           =   315
      End
      Begin VB.Label lblLabels 
         Caption         =   "04:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   28
         Left            =   -70440
         TabIndex        =   283
         Top             =   1680
         Width           =   315
      End
      Begin VB.Label lblLabels 
         Caption         =   "03:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   27
         Left            =   -70440
         TabIndex        =   282
         Top             =   1440
         Width           =   315
      End
      Begin VB.Label lblLabels 
         Caption         =   "02:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   26
         Left            =   -70440
         TabIndex        =   281
         Top             =   1200
         Width           =   315
      End
      Begin VB.Label lblLabels 
         Caption         =   "01:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   23
         Left            =   -70440
         TabIndex        =   280
         Top             =   960
         Width           =   315
      End
      Begin VB.Label lblLabels 
         Caption         =   "Metodo:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   25
         Left            =   -69720
         TabIndex        =   266
         Top             =   2760
         Width           =   795
      End
      Begin VB.Label lblLabels 
         Caption         =   "Qtde:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   24
         Left            =   -69720
         TabIndex        =   265
         Top             =   2280
         Width           =   795
      End
      Begin VB.Label lblLabels 
         Caption         =   "Codigo:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   22
         Left            =   -69720
         TabIndex        =   264
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label lblLabels 
         Caption         =   "Peca Ciclo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   20
         Left            =   -70680
         TabIndex        =   256
         Top             =   2130
         Width           =   975
      End
      Begin VB.Label txttaktat 
         Caption         =   "txttaktat"
         Height          =   255
         Left            =   -66120
         TabIndex        =   254
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label txttak 
         Caption         =   "txttak"
         Height          =   255
         Left            =   -66120
         TabIndex        =   253
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Ciclo Peca"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   21
         Left            =   -72000
         TabIndex        =   250
         Top             =   2130
         Width           =   975
      End
      Begin VB.Label Label14 
         Caption         =   "Ajuste Martelo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -70200
         TabIndex        =   230
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "Usar Calco inf de"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -70200
         TabIndex        =   229
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "Avanco da Fita"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -70200
         TabIndex        =   228
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Pinos Almofada"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -70200
         TabIndex        =   227
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Quant"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -70200
         TabIndex        =   226
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Diametro"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -70200
         TabIndex        =   225
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Comprimento"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -70200
         TabIndex        =   224
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Pressao Almofada"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -70200
         TabIndex        =   223
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Observações"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   -74880
         TabIndex        =   130
         Top             =   1290
         Width           =   1995
      End
      Begin VB.Label Label3 
         Caption         =   "Regulagem de Máquina"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   -74880
         TabIndex        =   129
         Top             =   3450
         Width           =   1995
      End
      Begin VB.Label lblLabels 
         Caption         =   "Fator de Batidas"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   12
         Left            =   -72000
         TabIndex        =   123
         Top             =   1530
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Caption         =   "PC/Hora"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   13
         Left            =   -69240
         TabIndex        =   122
         Top             =   1530
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Ferramenta"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   14
         Left            =   -72000
         TabIndex        =   121
         Top             =   2850
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Fluxo"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   -74760
         TabIndex        =   120
         Top             =   1650
         Width           =   495
      End
      Begin VB.Label lblLabels 
         Caption         =   "Equipamento Auxiliar"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   4
         Left            =   -74880
         TabIndex        =   119
         Top             =   4410
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         Caption         =   "Fator de Corte"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   10
         Left            =   -70680
         TabIndex        =   118
         Top             =   1530
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Simbologia"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -67320
         TabIndex        =   117
         Top             =   2250
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Takt Time"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   19
         Left            =   -67800
         TabIndex        =   116
         Top             =   1530
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Equipamento"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   8
         Left            =   -74880
         TabIndex        =   59
         Top             =   1410
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Operadores"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   9
         Left            =   -74880
         TabIndex        =   58
         Top             =   1770
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Terceiros"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   11
         Left            =   -74880
         TabIndex        =   57
         Top             =   3210
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Esquerda"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   5
         Left            =   -74880
         TabIndex        =   56
         Top             =   3570
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Direita"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   7
         Left            =   -74880
         TabIndex        =   55
         Top             =   3930
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Setor Refugo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   15
         Left            =   -71040
         TabIndex        =   54
         Top             =   3630
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Setor Fabrica"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   16
         Left            =   -71040
         TabIndex        =   53
         Top             =   3990
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Montagem"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   17
         Left            =   -69000
         TabIndex        =   52
         Top             =   3630
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Local/Montagem"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   18
         Left            =   -69000
         TabIndex        =   51
         Top             =   3990
         Width           =   1215
      End
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   64
      Left            =   2880
      MaxLength       =   24
      TabIndex        =   2
      Top             =   360
      Width           =   1995
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   3
      Left            =   960
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   720
      Width           =   9015
   End
   Begin VB.TextBox TXTSSQ 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox TXTSEQ 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox TXTPF 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   735
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   10080
      TabIndex        =   320
      Top             =   600
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmPFS.frx":ED4C
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
      Left            =   10080
      TabIndex        =   321
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmPFS.frx":F2E6
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
      Height          =   735
      Left            =   10680
      TabIndex        =   322
      Top             =   2760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      Picture         =   "frmPFS.frx":F880
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
   Begin VB.Label Label4 
      Caption         =   "Subtipo"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   7440
      TabIndex        =   317
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo Cliente"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   6
      Left            =   5400
      TabIndex        =   249
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo Integrado"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   40
      Left            =   2880
      TabIndex        =   12
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      Caption         =   "Fabrica"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   39
      Left            =   2160
      TabIndex        =   6
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descrição"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Caption         =   "SSQ:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   9
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Caption         =   "SEQ:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "PF:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmPFS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cARQPF As String
Dim cSQL As String
Dim cSQL2 As String
Dim aVAL As Variant
Dim aFOR As Variant
Dim aCAM As Variant
Dim aVAL2 As Variant
Dim aFOR2 As Variant
Dim aCAM2 As Variant
Dim aPAD As Variant
Dim aPAD2 As Variant
Dim nCAMPOS As Integer
Dim nCAMPOS2 As Integer
Dim iLOOP As Integer
Dim lTROCOU As Variant
Private Sub Cmdapgobs_Click()
    If MDG("Apagar Observacao") Then
        TXTFIELDS(26) = " "
    End If
End Sub
Private Sub Cmdbusmr_Click()
    Dim aRETU As Variant
    Dim cARQ, cSQL, cCODIGO As String
    cCODIGO = FixStr(TXTFIELDS(19))
    If Len(cCODIGO) > 0 Then
        cARQ = GeraConn(zMANA5EMP, "JETFOX")
        cSQL = "SELECT NOME,NOM2 FROM mR01 WHERE CODIGO='" & cCODIGO & "'"
        aRETU = PegSQL(cARQ, cSQL, 2, Array("NOME", "NOM2"), Array("C", "C"), Array("", ""))
        If lRETU Then
            TXTFIELDS(69) = aRETU(0) + " " + aRETU(1)
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    On Error Resume Next
    Dim cSQLIMG As String
    Dim cCAMIMG

    CONTAHOMEM
    If MDG("Gravar Sequencia", "Confirme Gravação ") Then
        'CAMPOS
        For iLOOP = 0 To nCAMPOS - 1 '- 8
            aVAL(iLOOP) = TXTFIELDS(iLOOP)
        Next iLOOP

        GrvSQL cARQPF, cSQL, nCAMPOS, aCAM, aVAL, aFOR
        
        'campos boleanos
        aVAL2(0) = FixNumBol(EPI01.Value)
        aVAL2(1) = FixNumBol(EPI02.Value)
        aVAL2(2) = FixNumBol(EPI03.Value)
        aVAL2(3) = FixNumBol(EPI04.Value)
        aVAL2(4) = FixNumBol(EPI05.Value)
        aVAL2(5) = FixNumBol(EPI06.Value)
        aVAL2(6) = FixNumBol(EPI07.Value)
        GrvSQL cARQPF, cSQL2, nCAMPOS2, aCAM2, aVAL2, aFOR2
        
        'campos imagens
        For iLOOP = 0 To 4
            If lTROCOU(iLOOP) Then
                Select Case iLOOP
                Case 0
                    cCAMIMG = "FIG01"            ' PROCESSO
                Case 1
                    cCAMIMG = "FIG02"            'ferramenta
                Case 2
                    cCAMIMG = "FIGEMBAL"         'embalagem
                Case 3
                    cCAMIMG = "FIGSEQ01"         'ferramenta 2
                Case 4
                    cCAMIMG = "FIGSEQ02"         'visa explodida
                End Select
                cSQLIMG = "select  PF,SEQ,SSQ," & cCAMIMG & " from PFS WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ
                ADOGrvBlob cARQPF, cSQLIMG, Picture1(iLOOP), cCAMIMG
            End If
        Next iLOOP
    End If
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub CmdConfImp_Click()
    FrmPrintSetup.Show vbModal, Me
End Sub

Private Sub CmdEditar_Click()
    ePASS01 = TXTFIELDS(11)
    frmETI.Show vbModal, Me
End Sub

Private Sub Cmdepi_Click()
    TXTFIELDS(26) = TXTFIELDS(26) & " É OBRIGATÓRIO O USO DOS EQUIPAMENTOS DE PROTEÇÃO INDIVIDUAIS (EPI's)" _
                                & " - ÓCULOS, PROTETOR AURICULAR, LUVAS E SAPATOS"
End Sub

Private Sub CmdEscMe01_Click()
    lRETU = False
    eRETU01 = ""
    eRETU02 = ""
    cARQESC = "ME01"
    escNUMNOM.Show vbModal, Me
    If lRETU Then
        frmPFS.TXTFIELDS(8) = eRETU01
        frmPFS.TXTFIELDS(4) = eRETU02
    End If
End Sub

Private Sub Cmdescmr_Click()
    iMU01 = 4
    escmu01.Show vbModal, Me
    If lRETU Then
        TXTFIELDS(68) = eRETU01
        TXTFIELDS(69) = eRETU02
    End If

End Sub

Private Sub CmdForFer_Click(Index As Integer)
    Select Case Index
    Case 0
        zgrp = TXTFIELDS(14)
    Case 1
        zgrp = TXTFIELDS(28)
    Case 2
        zgrp = TXTFIELDS(29)
    Case 3
        zgrp = TXTFIELDS(30)
    End Select
    iImage = 3
    cARQRTF = PegPath("PATH", "IMGFER")
    Load frmIMAGENS
    frmIMAGENS.TXTFIELDS(0).Enabled = False
    frmIMAGENS.TXTFIELDS(1).Enabled = False
    frmIMAGENS.Escolher(0).Visible = False
    frmIMAGENS.Show vbModal, Me
End Sub

Private Sub cmdimp_Click()
    cARQIMP = "PFI"
    FrmImp3.Show vbModal, Me
    FilRelat
End Sub

Private Sub CmdImprimir_Click(Index As Integer)
  If Picture1(Index).Height = 0 Then
        Alert ("Sem Imagem")
    Else
        ePASS01 = Index
        PrintPreview1.ShowPreview
    End If
End Sub

Private Sub CmdPaste_Click(Index As Integer)
    If Picture1(Index).Height = 0 Then
        Alert ("Sem Imagem")
    Else
        Clipboard.Clear
        Clipboard.SetData Picture1(Index).Image, (vbCFBitmap)
    End If
End Sub

Private Sub cmdpeg_Click(Index As Integer)

    Select Case Index
    Case 0
        iMU01 = 202
    End Select
    eRETU02 = ""                                 'Evita Cancelamento Escolha
    escIED.Show vbModal, Me
    If lRETU Then
        If Index = 0 Then
            TXTFIELDS(2) = Mid(eRETU02, 1, 1)
        End If
    End If

End Sub

Private Sub CmdRenumera_Click()
    Dim nNUM As Integer
    Dim sSQL As String
    If GridIns.Row = 0 Then
        Exit Sub
    End If
    nNUM = 0
    nNUM = FixInt(Val(Busca("Nº", "Renumerando", CStr(nNUM), 8)))
    If nNUM = 0 Then
        Alert ("Nº em Branco")
        Exit Sub
    End If
    GridIns.Col = 0
    nORD = FixInt(GridIns)
    sSQL = "select PF,SEQ,SSQ,ITEM from PFI WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ & " AND ITEM=" & nORD
    GrvSQL cARQPF, sSQL, 1, Array("ITEM"), Array(nNUM), Array("NI")
    FilRelat
End Sub

Private Sub CMDSIG_Click(Index As Integer)
    lRETU = False
    eRETU01 = ""
    eRETU02 = ""
    ePASS01 = "isoqsymbol"
    ePASS02 = True
    frmCharacters.Show vbModal, Me
    If lRETU Then
        TXTFIELDS(50 + Index) = eRETU01
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    lRETU = False
    eRETU01 = ""
    eRETU02 = ""
    escFER.Show vbModal, Me

    If lRETU Then

        Select Case Index

        Case 0
            frmPFS.TXTFIELDS(14) = eRETU01

        Case 1
            frmPFS.TXTFIELDS(28) = eRETU01

        Case 2
            frmPFS.TXTFIELDS(29) = eRETU01

        Case 3
            frmPFS.TXTFIELDS(30) = eRETU01

        End Select

    End If

End Sub

Private Sub Command10_Click()
    If MDG("Apagar Obs Regulagem") Then
        TXTFIELDS(27) = ""
    End If
End Sub

Private Sub Command11_Click()
    iMU01 = 4
    escmp.Show vbModal
End Sub

Private Sub Command12_Click()
    iMU01 = 7
    escmp.Show vbModal, Me
End Sub

Private Sub Command13_Click()
    nNOVAPOS = FixInt(PegUltGrid(GrdQsbLep), 0)
    nNOVAPOS = nNOVAPOS + 1
    iMU01 = 6
    NewOrd.Show vbModal, Me
    FilQSBLEP
End Sub

Private Sub Command14_Click()
    Dim sSQL As String
    GrdQsbLep.Col = 0
    nORD = FixInt(GrdQsbLep)
    sSQL = "select * from PFQSBLEP WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ & " AND ITEM=" & nORD
    If ApagaSQLP(Sdb, sSQL, "Apagar Item Qsb Lep") Then
        FilQSBLEP
    End If
End Sub

Private Sub Command15_Click()
    GrdQsbLep.Col = 0
    nORD = FixInt(GrdQsbLep)
    FrmQsbLep.Show vbModal, Me
    FilQSBLEP
End Sub

Private Sub Command16_Click()
    cARQIMP = "PFQSBLEP"
    FrmImp3.Show vbModal, Me
    FilQSBLEP
End Sub

Private Sub Command2_Click(Index As Integer)
    nNOVAPOS = FixInt(PegUltGrid(GridIns), 0)
    nNOVAPOS = nNOVAPOS + 1
    iMU01 = 1
    NewOrd.Show vbModal, Me
    FilRelat
End Sub

Private Sub Command20_Click()
    Dim X
    Dim cTEMP As String
    If Not MDG("Trocar Tipo Almofada") Then
        Exit Sub
    End If
    If TXTFIELDS(36).Text = "6" Then
        TXTFIELDS(36).Text = "8"
    Else
        TXTFIELDS(36).Text = "6"
    End If
    For X = 1 To 8
        If TXTFIELDS(36).Text = "6" Then
            cTEMP = String(6, "¡") + Space(2)
            If X >= 5 Then
                cTEMP = Space(8)
            End If
        Else
            cTEMP = String(8, "¡")
        End If
        TXTFIELDS(40 + X).Text = cTEMP
    Next
End Sub
Private Sub Command3_Click()
    Dim sSQL As String
    GridIns.Col = 0
    nORD = FixInt(GridIns)
    sSQL = "select * from PFI WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ & " AND ITEM=" & nORD
    If ApagaSQLP(Sdb, sSQL, "Apagar Item Instrucao") Then
        FilRelat
    End If
End Sub
Private Sub Command6_Click()
    iMU01 = 1
    escmp.Show vbModal
    If lRETU Then
        frmPFS.TXTFIELDS(20) = eRETU01
        frmPFS.TXTFIELDS(21) = eRETU02
    End If
End Sub
Private Sub Command7_Click(Index As Integer)
    Select Case Index
    Case 0
        frmPFS.TXTFIELDS(23) = frmPF.TXTFIELDS(2)
    Case 1
        frmPFS.TXTFIELDS(24) = frmPF.TXTFIELDS(2)
    End Select
End Sub
Private Sub Command8_Click()
    CONTAHOMEM
End Sub
Private Sub Command9_Click(Index As Integer)
    Dim aRETU As Variant
    Dim cARQ As String
    Dim sSQL As String
    Dim cCODIGO As Variant
    Select Case Index
    Case 0
        cCODIGO = FixStr(TXTFIELDS(8), "", "TRIM")
        sSQL = "SELECT NOME FROM MP01 WHERE CODIGO='" & cCODIGO & "'"

    Case 1
        cCODIGO = FixStr(TXTFIELDS(9), "", "TRIM")
        sSQL = "SELECT NOME FROM MP02 WHERE CODIGO='" & cCODIGO & "'"

    Case 2
        cCODIGO = FixStr(TXTFIELDS(10), "", "TRIM")
        sSQL = "SELECT NOME FROM MP02 WHERE CODIGO='" & cCODIGO & "'"

    Case 3
        cCODIGO = FixStr(TXTFIELDS(15), "", "TRIM")
        sSQL = "SELECT NOME FROM MP02 WHERE CODIGO='" & cCODIGO & "'"

    Case 4
        cCODIGO = FixStr(TXTFIELDS(16))
        sSQL = "SELECT NOME FROM MP02 WHERE CODIGO='" & cCODIGO & "'"

    Case 5
        cCODIGO = FixStr(TXTFIELDS(11), "", "TRIM")
        sSQL = "SELECT NOME FROM ETI WHERE CODIGO='" & cCODIGO & "'"

    Case 6
        cCODIGO = FixStr(TXTFIELDS(20), "", "TRIM")
        sSQL = "SELECT NOME FROM MP01 WHERE CODIGO='" & cCODIGO & "'"

    End Select
    
    cARQ = GeraConn(zMANA5EMP, "JETFOX")
    
    aRETU = PegSQL(cARQ, sSQL, 1, Array("NOME"), Array("C"), Array(""))
    If lRETU Then
       Select Case Index
        Case 0
            TXTFIELDS(4) = aRETU(0)
        Case 1
            TXTFIELDS(5) = aRETU(0)
        Case 2
            TXTFIELDS(7) = aRETU(0)
        Case 3
            TXTFIELDS(17) = aRETU(0)
        Case 4
            TXTFIELDS(18) = aRETU(0)
        Case 5
            TXTFIELDS(6) = aRETU(0)
        Case 6
            TXTFIELDS(21) = aRETU(0)
        End Select
    End If
End Sub

Private Sub CONTAHOMEM()
    Dim nHOMEM As Byte
    nHOMEM = 0

    If Len(TXTFIELDS(9)) > 0 Then
        If FixNum(TXTFIELDS(86)) > 0 Then
            nHOMEM = nHOMEM + FixNum(TXTFIELDS(86))
        Else
            nHOMEM = nHOMEM + 1
            TXTFIELDS(86) = 1
        End If
        If Len(TXTFIELDS(5)) = 0 Then
            Command9_Click (1)                   'tras descritivo
        End If
    Else
        TXTFIELDS(86) = 0
    End If

    If Len(TXTFIELDS(10)) > 0 Then
        If FixNum(TXTFIELDS(87)) > 0 Then
            nHOMEM = nHOMEM + FixNum(TXTFIELDS(87))
        Else
            nHOMEM = nHOMEM + 1
            TXTFIELDS(87) = 1
        End If
        If Len(TXTFIELDS(7)) = 0 Then
            Command9_Click (2)                   'tras descritivo
        End If
    Else
        TXTFIELDS(87) = 0
    End If

    If Len(TXTFIELDS(15)) > 0 Then
        If FixNum(TXTFIELDS(88)) > 0 Then
            nHOMEM = nHOMEM + FixNum(TXTFIELDS(88))
        Else
            nHOMEM = nHOMEM + 1
            TXTFIELDS(88) = 1
        End If
        If Len(TXTFIELDS(17)) = 0 Then
            Command9_Click (3)                   'tras descritivo
        End If
    Else
        TXTFIELDS(88) = 0
    End If
    

    If Len(TXTFIELDS(16)) > 0 Then
        If FixNum(TXTFIELDS(89)) > 0 Then
            nHOMEM = nHOMEM + FixNum(TXTFIELDS(89))
        Else
            nHOMEM = nHOMEM + 1
            TXTFIELDS(89) = 1
        End If
        If Len(TXTFIELDS(18)) = 0 Then
            Command9_Click (4)                   'tras descritivo
        End If
    Else
        TXTFIELDS(89) = 0
    End If

    If nHOMEM > 0 Then
        TXTFIELDS(31) = nHOMEM
    Else
        TXTFIELDS(31) = 0
    End If

End Sub

Private Sub CommandX_Click(Index As Integer)
    Dim cARQ As String
    cARQ = ""
    Select Case Index
    Case 0
        cARQ = "FiguraProcesso"
    Case 1
        cARQ = "FiguraFerramenta"
    Case 2
        cARQ = "FiguraEmbalagem"
    Case 3
        cARQ = "VistaExplodida"
    Case 4
        cARQ = "FolhaOperacao"
    End Select

    cARQ = cARQ & "_PF_"
    cARQ = cARQ & StrZero(FixInt(TXTPF.Text), 6)
    cARQ = cARQ & "_SEQ_"
    cARQ = cARQ & StrZero(FixInt(TXTSEQ.Text), 3)
    cARQ = cARQ & "_SSQ_"
    cARQ = cARQ & StrZero(FixInt(TXTSSQ.Text), 3)
    salvarpict Me, Picture1(Index), cARQ

End Sub

Private Sub DelImg_Click(Index As Integer)
    Set Picture1(Index).Picture = Nothing
    Set Picture2(Index).Picture = Nothing
    lTROCOU(Index) = True
End Sub

Private Sub Edifer_Click(Index As Integer)
    Dim cARQ As String
    Dim cSQL As String
    Dim aRETU As Variant
    Select Case Index
    Case 0
        ePASS01 = TXTFIELDS(14)
    Case 1
        ePASS01 = TXTFIELDS(28)
    Case 2
        ePASS01 = TXTFIELDS(29)
    Case 3
        ePASS01 = TXTFIELDS(30)
    End Select
    
    cARQ = GeraConn(PegPath("PATH", "MANA5FER"), "SDECDX")
    cSQL = "select * from FERRAM WHERE FERRAM='" & LTrim(RTrim(ePASS01)) & "'"
    aRETU = PegSQL(cARQ, cSQL, 1, Array("NOME"), Array("C"), Array(""))
    
    ePASS02 = aRETU(0)
    frmFER.Show vbModal, Me
End Sub

Private Sub Encerrar_Click()
    If Not MDG("Sair sem gravar") Then
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub esc1_Click(Index As Integer)
    Dim linha
    Dim coluna
    Dim cTEMP As String
    Dim cDIG As String
    Dim cGRV As String
    linha = (Int((Index - 1) / 8)) + 1
    coluna = (Index + 8) - (linha * 8)
    If TXTFIELDS(36).Text = "6" And linha > 4 Then
        Alert ("Linha Incorreta para Formato 4x6")
        Exit Sub
    End If
    If TXTFIELDS(36).Text = "6" And coluna > 6 Then
        Alert ("Coluna Incorreta para Formato 4x6")
        Exit Sub
    End If
    cTEMP = TXTFIELDS(40 + linha).Text
    cDIG = Mid(cTEMP, coluna, 1)
    If cDIG = "l" Then
        cDIG = "¡"
    Else
        cDIG = "l"
    End If
    Select Case coluna
    Case 1
        cGRV = cDIG & Mid(cTEMP, 2, 7)
    Case 8
        cGRV = Mid(cTEMP, 1, 7) + cDIG
    Case Else
        cGRV = Mid(cTEMP, 1, coluna - 1) + cDIG + Mid(cTEMP, coluna + 1)
    End Select
    TXTFIELDS(40 + linha).Text = cGRV
    TXTFIELDS(40 + linha).Refresh
End Sub

Private Sub escArranjo_Click(Index As Integer)
    iMU01 = 8
    If Index = 1 Then
        iMU01 = 11
    End If
    If Index = 2 Then
        iMU01 = 9
    End If
    
    escmp.Show vbModal

    If lRETU Then

        frmPFS.TXTFIELDS(8) = eRETU01
        frmPFS.TXTFIELDS(4) = eRETU02

    End If

End Sub

Private Sub EscCodFlu_Click(Index As Integer)
    lRETU = False
    eRETU01 = ""
    eRETU02 = ""
    escFLX.Show vbModal, Me
    If lRETU Then
        TXTFIELDS(50 + Index) = eRETU01
    End If
End Sub

Private Sub escloc_Click()
    iMD02 = 5
    escMD02.Show vbModal, Me
    frmPFS.TXTFIELDS(35) = Left(eRETU01, 1)
End Sub

Private Sub escmon_Click()

    iMD02 = 4
    escMD02.Show vbModal, Me
    frmPFS.TXTFIELDS(34) = Left(eRETU01, 1)

End Sub

Private Sub EscMP01_Click()

    iMU01 = 1
    escmp.Show vbModal

    If lRETU Then

        frmPFS.TXTFIELDS(8) = eRETU01
        frmPFS.TXTFIELDS(4) = eRETU02

    End If

End Sub

Private Sub ESCMP03_Click()
    iMU01 = 4
    escmp.Show vbModal
    If lRETU Then
        frmPFS.TXTFIELDS(11) = eRETU01
        frmPFS.TXTFIELDS(6) = eRETU02
    End If
End Sub

Private Sub ESCpro_Click(Index As Integer)
    Dim cCHAVEBUS As String
    ePASS01 = "MANA5"
    If Index = 0 Then
        ePASS01 = "LOGIS"                        ''logis pois e subproduto aqui
        escms01.Show vbModal, Me
        If lRETU Then
            TXTFIELDS(65) = eRETU01
            TXTFIELDS(64) = eRETU03
        End If
    End If
    If Index = 1 Then
        ePASS01 = "LXEMB"                        ''embalagem aqui
        escms01.Show vbModal, Me
        If lRETU Then
            TXTFIELDS(68) = eRETU01
            TXTFIELDS(69) = eRETU02
        End If
    End If
    cCHAVEBUS = ""
End Sub

Private Sub escRecursoHomem_Click(Index As Integer)
    iMU01 = 10                                   'lx recurso_970_homem
    If Index >= 10 Then
        iMU01 = 2                                'mp02
    End If
    escmp.Show vbModal
    If lRETU Then
        Select Case Index
        Case 0, 10
            frmPFS.TXTFIELDS(9) = eRETU01
            frmPFS.TXTFIELDS(5) = eRETU02
        Case 1, 11
            frmPFS.TXTFIELDS(10) = eRETU01
            frmPFS.TXTFIELDS(7) = eRETU02
        Case 2, 12
            frmPFS.TXTFIELDS(15) = eRETU01
            frmPFS.TXTFIELDS(17) = eRETU02
        Case 3, 13
            frmPFS.TXTFIELDS(16) = eRETU01
            frmPFS.TXTFIELDS(18) = eRETU02
        End Select
    End If
End Sub

Private Sub escsetins_Click()
    iMD02 = 2
    escMD02.Show vbModal, Me
    frmPFS.TXTFIELDS(33) = Left(eRETU01, 3)
End Sub

Private Sub escsetref_Click()
    iMD02 = 3
    escMD02.Show vbModal, Me
    frmPFS.TXTFIELDS(32) = Left(eRETU01, 1)
End Sub

Private Sub FilRelat()
    Dim sSQL As String
    Dim Nsoma As Integer
    On Error Resume Next
    sSQL = "select ITEM,NORMA,DESCR,TEMPOMAN,TEMPOMAQ from PFI WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ & " ORDER BY ITEM"

    MontaGridFast GridIns, 5, Array(400, 600, 4000, 600, 600), Array("Ord", "Norma", "Instrucao", "tman", "tmaq"), _
        Array("Item", "Norma", "descr", "Tempoman", "Tempomaq"), Sdb, sSQL
    
    txttak.Caption = SomaGrid(GridIns, 3, 1)
    txttaktat.Caption = SomaGrid(GridIns, 4, 1)
    
End Sub

Private Sub FilQSBLEP()
    Dim sSQL As String
    sSQL = "select ITEM,MUD,DESCRICAO from PFQSBLEP WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ & " ORDER BY ITEM"
    MontaGridUltra GrdQsbLep, 3, Array(400, 400, 5000), Array("Ord", "MUD", "descricao"), _
        Array("Item", "C$Mud", "L$descricao"), Sdb, sSQL
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
    Dim cCAMIMG
    Dim cSQLIMG
    Dim X
    CenterFormToScreen Me
    lTROCOU = Array(False, False, False, False, False)
    TXTPF.Text = nPF
    TXTSEQ.Text = nSEQ
    TXTSSQ.Text = nSSQ
    cARQPF = Sdb        'agora tem que usar sdb pois pode ser pf pfp pfg
    
'Usando 2 matrizes pois estava com erro string complex

cSQL = "  SELECT PF,SEQ,SSQ,FILIAL,FERPRE,SUBTIPO,DESCRI,NOMMP01,NOMMP02,NOMMP03,NOMMP02B,CODMP01,CODMP02,"
cSQL = cSQL & " CODMP02B,CODMP03,FATBAT,PCHORA,FERRAMEN,CODMP02C,CODMP02D,NOMMP02C,NOMMP02D,TAKTIME,"
cSQL = cSQL & " FERCOD,FERNOM,FERAVF,CODIGOE,CODIGOD,FATCORTE,OBS,REGULAR,FERRAME2,FERRAME3,"
cSQL = cSQL & " FERRAME4,QTHOMEM,SETORREF,SETORINS,MONTAGEM,FEITO,ESQTIP,FERAJM,FERUCI,FERPIN,"
cSQL = cSQL & " FERQTD,ESQL01,ESQL02,ESQL03,ESQL04,ESQL05,ESQL06,ESQL07,ESQL08,FERDIM,"
cSQL = cSQL & " FERCOM,FLX01,FLX02,FLX03,FLX04,FLX05,FLX06,FLX07,FLX08,FLX09,"
cSQL = cSQL & " FLX10,SIMBOLOSEG,SIMBOLOSE2,SIMBOLOSE3,CODINT,PCCICLO,CICLOPC,CODIGOCLI,"
cSQL = cSQL & " EMBALCOD,EMBALNOME,EMBALQTDE,EMBALMETODO,"
cSQL = cSQL & " COMPON01,COMPON02,COMPON03,COMPON04,COMPON05,COMPON06,COMPON07,COMPON08,COMPON09,"
cSQL = cSQL & " FERRALT,FERRPASSO,FERRFORCA,FERRFIXA,"
cSQL = cSQL & " QTDMP01,QTDMP02,QTDMP02B,QTDMP02C,QTDMP02D"
cSQL = cSQL & " ,COMPON10, COMPON11, COMPON12, COMPON13, COMPON14, COMPON15"
cSQL = cSQL & " from PFS WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ
    
    nCAMPOS = 96 'campos
    aCAM = Array("FILIAL", "FERPRE", "SUBTIPO", "DESCRI", "NOMMP01", "NOMMP02", "NOMMP03", "NOMMP02B", "CODMP01", "CODMP02", _
                 "CODMP02B", "CODMP03", "FATBAT", "PCHORA", "FERRAMEN", "CODMP02C", "CODMP02D", "NOMMP02C", "NOMMP02D", "TAKTIME", _
                 "FERCOD", "FERNOM", "FERAVF", "CODIGOE", "CODIGOD", "FATCORTE", "OBS", "REGULAR", "FERRAME2", "FERRAME3", _
                 "FERRAME4", "QTHOMEM", "SETORREF", "SETORINS", "MONTAGEM", "FEITO", "ESQTIP", "FERAJM", "FERUCI", "FERPIN", _
                 "FERQTD", "ESQL01", "ESQL02", "ESQL03", "ESQL04", "ESQL05", "ESQL06", "ESQL07", "ESQL08", "FERDIM", _
                 "FERCOM", "FLX01", "FLX02", "FLX03", "FLX04", "FLX05", "FLX06", "FLX07", "FLX08", "FLX09", _
                 "FLX10", "SIMBOLOSEG", "SIMBOLOSE2", "SIMBOLOSE3", "CODINT", "PCCICLO", "CICLOPC", "CODIGOCLI", _
                 "EMBALCOD", "EMBALNOME", "EMBALQTDE", "EMBALMETODO", _
                 "COMPON01", "COMPON02", "COMPON03", "COMPON04", "COMPON05", "COMPON06", "COMPON07", "COMPON08", "COMPON09", _
                 "FERRALT", "FERRPASSO", "FERRFORCA", "FERRFIXA", _
                 "QTDMP01", "QTDMP02", "QTDMP02B", "QTDMP02C", "QTDMP02D", _
                 "COMPON10", "COMPON11", "COMPON12", "COMPON13", "COMPON14", "COMPON15")
                 
    aFOR = Array("C", "C", "C", "C", "C", "C", "C", "C", "C", "C", _
                 "C", "C", "NI", "NI", "C", "C", "C", "C", "C", "N", _
                 "C", "C", "C", "C", "C", "N", "C", "C", "C", "C", _
                 "C", "NI", "C", "C", "C", "C", "C", "C", "C", "C", _
                 "C", "C", "C", "C", "C", "C", "C", "C", "C", "C", _
                 "C", "C", "C", "C", "C", "C", "C", "C", "C", "C", _
                 "C", "C", "C", "C", "C", "N", "N", "C", _
                 "C", "C", "N", "C", _
                 "C", "C", "C", "C", "C", "C", "C", "C", "C", _
                 "C", "N", "C", "C", _
                 "N", "N", "N", "N", "N", _
                 "C", "C", "C", "C", "C", "C")
    
    aPAD = Array("", "", "", "", "", "", "", "", "", "", _
                 "", "", 0, 0, "", "", "", "", "", 0, _
                 "", "", "", "", "", 0, "", "", "", "", _
                 "", 0, "", "", "", "", "", "", "", "", _
                 "", "", "", "", "", "", "", "", "", "", _
                 "", "", "", "", "", "", "", "", "", "", _
                 "", "", "", "", "", 0, 0, "", _
                 "", "", 0, "", _
                 "", "", "", "", "", "", "", "", "", _
                 "", 0, "", "", _
                 0, 0, 0, 0, 0, _
                 "", "", "", "", "", "")

    
    
    aVAL = PegSQL(cARQPF, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
    For iLOOP = 0 To nCAMPOS - 1
        TXTFIELDS(iLOOP) = aVAL(iLOOP)
    Next iLOOP

    nCAMPOS2 = 7 'campos boleanos
    cSQL2 = "  SELECT PF,SEQ,SSQ, "
    cSQL2 = cSQL2 & " EPI01,EPI02,EPI03,EPI04,EPI05,EPI06,EPI07"
    cSQL2 = cSQL2 & " from PFS WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ
    aCAM2 = Array("EPI01", "EPI02", "EPI03", "EPI04", "EPI05", "EPI06", "EPI07")
    aFOR2 = Array("B", "B", "B", "B", "B", "B", "B")
    aPAD2 = Array(False, False, False, False, False, False, False)
    aVAL2 = PegSQL(cARQPF, cSQL2, nCAMPOS2, aCAM2, aFOR2, aPAD2)
    EPI01.Value = FixBolNum(aVAL2(0))
    EPI02.Value = FixBolNum(aVAL2(1))
    EPI03.Value = FixBolNum(aVAL2(2))
    EPI04.Value = FixBolNum(aVAL2(3))
    EPI05.Value = FixBolNum(aVAL2(4))
    EPI06.Value = FixBolNum(aVAL2(5))
    EPI07.Value = FixBolNum(aVAL2(6))
    
    'campos imagens
    For iLOOP = 0 To 4
        Select Case iLOOP
        Case 0
            cCAMIMG = "FIG01"                    'processo
        Case 1
            cCAMIMG = "FIG02"                    'ferramenta
        Case 2
            cCAMIMG = "FIGEMBAL"                 'embalagem
        Case 3
            cCAMIMG = "FIGSEQ01"                 'ferramenta 2
        Case 4
            cCAMIMG = "FIGSEQ02"                 'vista explodida
        End Select
               
        cSQLIMG = "select " & cCAMIMG & " from PFS WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ
    
        If ADOPegBlob(cARQPF, cSQLIMG, Picture1(iLOOP), cCAMIMG) Then
            StretchSourcePictureFromPicture Picture1(iLOOP), Picture2(iLOOP)
            If FixNum(eRETU01) > 500000 Then
                Alert ("Imagem Muito Grande,Ajuste o tamanho")
                salvarpict Me, Picture1(iLOOP), cCAMIMG & "_PF_" & StrZero(nPF, 6) & "_SEQ_" & StrZero(nSEQ, 3) & "_SSQ_" & StrZero(nSSQ, 3)
                Set Picture1(iLOOP).Picture = Nothing
                Set Picture2(iLOOP).Picture = Nothing
                lTROCOU(iLOOP) = True
            End If
        Else
            Set Picture1(iLOOP).Picture = Nothing
            Set Picture2(iLOOP).Picture = Nothing
        End If
    Next iLOOP
    FilRelat
    FilQSBLEP
    For X = 51 To 63
        TXTFIELDS(X).Font = "isoqsymbol"
    Next
    CONTAHOMEM
    
     PrintPreview1.AuxiliaryButtonVisible = PrintPreview1.PrinterExists("Microsoft Print to PDF")
    PrintPreview1.AuxiliaryButtonToolTipText = "Salvar como PDF"
End Sub
Public Sub PrintPreview1_AuxiliaryButtonClick(UpdateReport As Boolean)
    PrintPreview1.ShowSaveToFile "Microsoft Print to PDF", "*.pdf"
    UpdateReport = False ' we don't need to update the report in the Print preview window after this action (the default value of UpdateReport parameter is True)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

Private Sub IncImg_Click(Index As Integer)
    Dim STMPFILE
    STMPFILE = OpenArqExt(Me, "", "JPG", "JPEG *.JPG")
    If lerarquivoimagem(STMPFILE, Picture1(Index), Picture2(Index)) Then
        lTROCOU(Index) = True
    End If
End Sub

Private Sub InsEdit_Click()
    GridIns.Col = 0
    nORD = FixInt(GridIns)
    frmPFI.Show vbModal, Me
    FilRelat
End Sub

Private Sub TXTFIELDS_Change(Index As Integer)
    Select Case Index
    Case 9, 10, 15, 16, 86, 87, 88, 89
        CONTAHOMEM
    End Select
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
    Case 12, 13, 70
        KeyAscii = ValiText(KeyAscii, "#NI")
    Case 19, 66, 65, 25, 82
        KeyAscii = ValiText(KeyAscii, "#N")
    End Select
End Sub

Private Sub TXTFIELDS_LostFocus(Index As Integer)
    Dim cSQLITEM
    Dim cCODLOGIX
    Dim nPOS
    If Index = 64 Then
        cCODLOGIX = Trim(TXTFIELDS(64))
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
End Sub

