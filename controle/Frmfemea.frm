VERSION 5.00
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "vbccr18.ocx"
Object = "{66E63055-5A66-4C79-9327-4BC077858695}#14.0#0"; "newtab01.OCX"
Object = "{075212A8-C1CF-444E-939D-F6046CCDBC08}#1.5#0"; "VBFLXGRD18.OCX"
Begin VB.Form frmfemea 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PF/FMEA"
   ClientHeight    =   9600
   ClientLeft      =   1092
   ClientTop       =   336
   ClientWidth     =   12768
   Icon            =   "Frmfemea.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   12768
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Txtfields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   26
      Left            =   8160
      Locked          =   -1  'True
      TabIndex        =   242
      TabStop         =   0   'False
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton ESCprofin 
      Caption         =   "MC"
      Height          =   315
      Index           =   2
      Left            =   3540
      Picture         =   "Frmfemea.frx":058A
      TabIndex        =   141
      TabStop         =   0   'False
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton ESCprofin 
      Caption         =   "LX"
      Height          =   315
      Index           =   1
      Left            =   3180
      Picture         =   "Frmfemea.frx":0B14
      TabIndex        =   138
      TabStop         =   0   'False
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Txtfields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   15
      Left            =   8580
      Locked          =   -1  'True
      MaxLength       =   24
      TabIndex        =   135
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox Txtfields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   9600
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Txtfields 
      Height          =   285
      Index           =   21
      Left            =   840
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1320
      Width           =   1935
   End
   Begin VB.CommandButton ESCprofin 
      Height          =   315
      Index           =   0
      Left            =   2820
      Picture         =   "Frmfemea.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1320
      Width           =   375
   End
   Begin VB.TextBox Txtfields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   20
      Left            =   6660
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Txtfields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   10
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox Txtfields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Txtfields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   360
      Width           =   4335
   End
   Begin VB.TextBox Txtfields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   19
      Left            =   1920
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   495
   End
   Begin VB.TextBox Txtfields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   11
      Left            =   7440
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox Txtfields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox Txtfields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   2880
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   960
      Width           =   3735
   End
   Begin VB.TextBox Txtfields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   24
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox txtPF 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      HideSelection   =   0   'False
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   360
      Width           =   1095
   End
   Begin NewTabCtl.NewTab SSTab1 
      Height          =   7575
      Left            =   240
      TabIndex        =   28
      Top             =   1680
      Width           =   12375
      _ExtentX        =   21823
      _ExtentY        =   13356
      Tabs            =   17
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabsPerRow      =   6
      TabHeight       =   529
      ControlVersion  =   13
      TabCaption(0)   =   "Producao"
      Tab(0).ControlCount=   19
      Tab(0).Control(0)=   "CmdOrdFem(11)"
      Tab(0).Control(1)=   "CmdOrdFem(10)"
      Tab(0).Control(2)=   "CmdREvisarTodas(16)"
      Tab(0).Control(3)=   "CmdRevFemea(16)"
      Tab(0).Control(4)=   "CmdREvisarTodas(0)"
      Tab(0).Control(5)=   "Command2(10)"
      Tab(0).Control(6)=   "CmdREvi(0)"
      Tab(0).Control(7)=   "CmdRevFemea(0)"
      Tab(0).Control(8)=   "CmdOrdFem(1)"
      Tab(0).Control(9)=   "CmdOrdFem(0)"
      Tab(0).Control(10)=   "EditSeq(0)"
      Tab(0).Control(11)=   "Command2(0)"
      Tab(0).Control(12)=   "Command1(0)"
      Tab(0).Control(13)=   "importar(0)"
      Tab(0).Control(14)=   "Command4(0)"
      Tab(0).Control(15)=   "CmdClo(0)"
      Tab(0).Control(16)=   "CmdClo(1)"
      Tab(0).Control(17)=   "GridSeq(0)"
      Tab(0).Control(18)=   "gridrevcau(0)"
      TabCaption(1)   =   "Grupo/Copia/Obs"
      Tab(1).ControlCount=   14
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(1)=   "Label4"
      Tab(1).Control(2)=   "Label6"
      Tab(1).Control(3)=   "Label13"
      Tab(1).Control(4)=   "Label14"
      Tab(1).Control(5)=   "Txtfields(7)"
      Tab(1).Control(6)=   "Txtfields(8)"
      Tab(1).Control(7)=   "Txtfields(18)"
      Tab(1).Control(8)=   "escidfolha(0)"
      Tab(1).Control(9)=   "escidfolha(1)"
      Tab(1).Control(10)=   "escidfolha(10)"
      Tab(1).Control(11)=   "escidfolha(11)"
      Tab(1).Control(12)=   "Txtfields(32)"
      Tab(1).Control(13)=   "Txtfields(33)"
      TabCaption(2)   =   "Elaborador/Revisão"
      Tab(2).ControlCount=   38
      Tab(2).Control(0)=   "Label10"
      Tab(2).Control(1)=   "Label9"
      Tab(2).Control(2)=   "Label5"
      Tab(2).Control(3)=   "Label8"
      Tab(2).Control(4)=   "Label7"
      Tab(2).Control(5)=   "Label12"
      Tab(2).Control(6)=   "gridrevi(0)"
      Tab(2).Control(7)=   "gridrev"
      Tab(2).Control(8)=   "femeacrg"
      Tab(2).Control(9)=   "Command5"
      Tab(2).Control(10)=   "CmdProx"
      Tab(2).Control(11)=   "Check1"
      Tab(2).Control(12)=   "CmdNovaREV"
      Tab(2).Control(13)=   "CmdMotRev"
      Tab(2).Control(14)=   "ComMotAlt(0)"
      Tab(2).Control(15)=   "ComMotExc"
      Tab(2).Control(16)=   "ComMotInc"
      Tab(2).Control(17)=   "Txtfields(6)"
      Tab(2).Control(18)=   "Txtfields(9)"
      Tab(2).Control(19)=   "Command3"
      Tab(2).Control(20)=   "Txtfields(12)"
      Tab(2).Control(21)=   "CmdLiberar"
      Tab(2).Control(22)=   "Txtfields(14)"
      Tab(2).Control(23)=   "Txtfields(13)"
      Tab(2).Control(24)=   "CmdApaRev"
      Tab(2).Control(25)=   "TxtRevisao"
      Tab(2).Control(26)=   "TxtRevNova"
      Tab(2).Control(27)=   "CmdTrocaREv"
      Tab(2).Control(28)=   "Txtfields(22)"
      Tab(2).Control(29)=   "Txtfields(23)"
      Tab(2).Control(30)=   "Txtfields(24)"
      Tab(2).Control(31)=   "Command15"
      Tab(2).Control(32)=   "Txtfields(25)"
      Tab(2).Control(33)=   "Command1(8)"
      Tab(2).Control(34)=   "Command16"
      Tab(2).Control(35)=   "Check2"
      Tab(2).Control(36)=   "DTPicker1"
      Tab(2).Control(37)=   "DTPicker2"
      TabCaption(3)   =   "Revisões PF"
      Tab(3).ControlCount=   6
      Tab(3).Control(0)=   "gridrevi(1)"
      Tab(3).Control(1)=   "ComMotAlt(1)"
      Tab(3).Control(2)=   "Txtfields(16)"
      Tab(3).Control(3)=   "Txtfields(17)"
      Tab(3).Control(4)=   "GridPFrev"
      Tab(3).Control(5)=   "CmdMarcaFemea"
      TabCaption(4)   =   "Sequencia"
      Tab(4).ControlCount=   7
      Tab(4).Control(0)=   "gridopr"
      Tab(4).Control(1)=   "Command1(2)"
      Tab(4).Control(2)=   "Command1(3)"
      Tab(4).Control(3)=   "CmdTotPad(0)"
      Tab(4).Control(4)=   "CmdTotPad(1)"
      Tab(4).Control(5)=   "CmdcheckOp(0)"
      Tab(4).Control(6)=   "CmdcheckOp(1)"
      TabCaption(5)   =   "Reg Defeito"
      Tab(5).ControlCount=   5
      Tab(5).Control(0)=   "gridmsrd"
      Tab(5).Control(1)=   "CmdVerDef"
      Tab(5).Control(2)=   "CmdIncDef"
      Tab(5).Control(3)=   "CmdDelDef"
      Tab(5).Control(4)=   "CmdEdiDef"
      TabCaption(6)   =   "Padrao Processo/Falha"
      Tab(6).ControlCount=   18
      Tab(6).Control(0)=   "dizfal"
      Tab(6).Control(1)=   "dizpro"
      Tab(6).Control(2)=   "Label2"
      Tab(6).Control(3)=   "Label1"
      Tab(6).Control(4)=   "GridPadPF"
      Tab(6).Control(5)=   "GRIDFAL"
      Tab(6).Control(6)=   "GridPRO"
      Tab(6).Control(7)=   "Cmdimp"
      Tab(6).Control(8)=   "CmdNome"
      Tab(6).Control(9)=   "Command8"
      Tab(6).Control(10)=   "Command6"
      Tab(6).Control(11)=   "CmdExcPFPad(0)"
      Tab(6).Control(12)=   "CmdIncPf"
      Tab(6).Control(13)=   "CmdEscFal"
      Tab(6).Control(14)=   "CmdEscPro"
      Tab(6).Control(15)=   "CmdExcPFPad(1)"
      Tab(6).Control(16)=   "Command13"
      Tab(6).Control(17)=   "Command14"
      TabCaption(7)   =   "Padrao"
      Tab(7).ControlCount=   28
      Tab(7).Control(0)=   "lblLabels(14)"
      Tab(7).Control(1)=   "lblLabels(0)"
      Tab(7).Control(2)=   "lblLabels(6)"
      Tab(7).Control(3)=   "Line1(0)"
      Tab(7).Control(4)=   "Line1(1)"
      Tab(7).Control(5)=   "Line1(2)"
      Tab(7).Control(6)=   "GridPad"
      Tab(7).Control(7)=   "CmdEdiPad(9)"
      Tab(7).Control(8)=   "CmdEdiPad(8)"
      Tab(7).Control(9)=   "CmdEdiPad(7)"
      Tab(7).Control(10)=   "CmdEdiPad(6)"
      Tab(7).Control(11)=   "txtsigi"
      Tab(7).Control(12)=   "CMDSIG(0)"
      Tab(7).Control(13)=   "CmdEdiPad(5)"
      Tab(7).Control(14)=   "CmdEdiPad(4)"
      Tab(7).Control(15)=   "CmdEdiPad(3)"
      Tab(7).Control(16)=   "CmdOrdPad(2)"
      Tab(7).Control(17)=   "CmdOrdPad(1)"
      Tab(7).Control(18)=   "CmdOrdPad(0)"
      Tab(7).Control(19)=   "CmdEdiPad(2)"
      Tab(7).Control(20)=   "CmdEdiPad(1)"
      Tab(7).Control(21)=   "CmdEdiPad(0)"
      Tab(7).Control(22)=   "txtRPN"
      Tab(7).Control(23)=   "CmdApagaRPN"
      Tab(7).Control(24)=   "CmdOrdPad(3)"
      Tab(7).Control(25)=   "importar(10)"
      Tab(7).Control(26)=   "CmdEdiPad(10)"
      Tab(7).Control(27)=   "CMDSIG(1)"
      TabCaption(8)   =   "Padrao-Complementar"
      Tab(8).ControlCount=   17
      Tab(8).Control(0)=   "gridrevcau(1)"
      Tab(8).Control(1)=   "GridSeq(1)"
      Tab(8).Control(2)=   "EditSeq(1)"
      Tab(8).Control(3)=   "Command2(1)"
      Tab(8).Control(4)=   "Command1(1)"
      Tab(8).Control(5)=   "importar(1)"
      Tab(8).Control(6)=   "Command4(1)"
      Tab(8).Control(7)=   "CmdClo(2)"
      Tab(8).Control(8)=   "CmdClo(3)"
      Tab(8).Control(9)=   "CmdOrdFem(2)"
      Tab(8).Control(10)=   "CmdOrdFem(3)"
      Tab(8).Control(11)=   "CmdRevFemea(1)"
      Tab(8).Control(12)=   "CmdREvi(1)"
      Tab(8).Control(13)=   "Command10(0)"
      Tab(8).Control(14)=   "Command11"
      Tab(8).Control(15)=   "Command12"
      Tab(8).Control(16)=   "Command10(1)"
      TabCaption(9)   =   "PreCadastro"
      Tab(9).ControlCount=   5
      Tab(9).Control(0)=   "Gridfemprecad"
      Tab(9).Control(1)=   "Command7"
      Tab(9).Control(2)=   "cmdedipre(2)"
      Tab(9).Control(3)=   "CmddelPre"
      Tab(9).Control(4)=   "CmdPreinc"
      TabCaption(10)  =   "Apuracoes"
      Tab(10).ControlCount=   14
      Tab(10).Control(0)=   "dizcomapu"
      Tab(10).Control(1)=   "lbl(3)"
      Tab(10).Control(2)=   "lbl(4)"
      Tab(10).Control(3)=   "GridOpe"
      Tab(10).Control(4)=   "Grid10"
      Tab(10).Control(5)=   "Command9"
      Tab(10).Control(6)=   "Cmdeditar"
      Tab(10).Control(7)=   "CmdVerApu"
      Tab(10).Control(8)=   "text(1)"
      Tab(10).Control(9)=   "text(2)"
      Tab(10).Control(10)=   "menosmes"
      Tab(10).Control(11)=   "maismes"
      Tab(10).Control(12)=   "menosano"
      Tab(10).Control(13)=   "maisano"
      TabCaption(11)  =   "Fluxo II"
      Tab(11).ControlCount=   13
      Tab(11).Control(0)=   "GrdQsbLep"
      Tab(11).Control(1)=   "Command1(5)"
      Tab(11).Control(2)=   "Command1(4)"
      Tab(11).Control(3)=   "CmdTotFlux(0)"
      Tab(11).Control(4)=   "CmdTotFlux(1)"
      Tab(11).Control(5)=   "Cmdcheckfx(0)"
      Tab(11).Control(6)=   "Cmdcheckfx(1)"
      Tab(11).Control(7)=   "CmdTotFlux(2)"
      Tab(11).Control(8)=   "CmdTotFlux(3)"
      Tab(11).Control(9)=   "Cmdcheckfx(2)"
      Tab(11).Control(10)=   "Cmdcheckfx(3)"
      Tab(11).Control(11)=   "Command1(12)"
      Tab(11).Control(12)=   "Command1(13)"
      TabCaption(12)  =   "Pre"
      Tab(12).ControlCount=   13
      Tab(12).Control(0)=   "gridrevcau(2)"
      Tab(12).Control(1)=   "GridSeq(2)"
      Tab(12).Control(2)=   "CmdREvi(2)"
      Tab(12).Control(3)=   "CmdRevFemea(2)"
      Tab(12).Control(4)=   "CmdOrdFem(4)"
      Tab(12).Control(5)=   "CmdOrdFem(5)"
      Tab(12).Control(6)=   "EditSeq(2)"
      Tab(12).Control(7)=   "Command2(2)"
      Tab(12).Control(8)=   "Command1(6)"
      Tab(12).Control(9)=   "importar(2)"
      Tab(12).Control(10)=   "Command4(2)"
      Tab(12).Control(11)=   "CmdClo(4)"
      Tab(12).Control(12)=   "CmdClo(5)"
      TabCaption(13)  =   "GP12"
      Tab(13).ControlCount=   13
      Tab(13).Control(0)=   "gridrevcau(3)"
      Tab(13).Control(1)=   "GridSeq(3)"
      Tab(13).Control(2)=   "CmdREvi(3)"
      Tab(13).Control(3)=   "CmdRevFemea(3)"
      Tab(13).Control(4)=   "CmdOrdFem(6)"
      Tab(13).Control(5)=   "CmdOrdFem(7)"
      Tab(13).Control(6)=   "EditSeq(3)"
      Tab(13).Control(7)=   "Command2(3)"
      Tab(13).Control(8)=   "Command1(7)"
      Tab(13).Control(9)=   "importar(3)"
      Tab(13).Control(10)=   "Command4(3)"
      Tab(13).Control(11)=   "CmdClo(6)"
      Tab(13).Control(12)=   "CmdClo(7)"
      TabCaption(14)  =   "Historico"
      Tab(14).ControlCount=   2
      Tab(14).Control(0)=   "GridSeq(4)"
      Tab(14).Control(1)=   "CmdREvi(4)"
      TabCaption(15)  =   "Padrao  Historico"
      Tab(15).ControlCount=   1
      Tab(15).Control(0)=   "gridpfpadhist"
      TabCaption(16)  =   "Adcionais"
      Tab(16).ControlCount=   14
      Tab(16).Control(0)=   "gridrevcau(5)"
      Tab(16).Control(1)=   "GridSeq(5)"
      Tab(16).Control(2)=   "Command2(15)"
      Tab(16).Control(3)=   "CmdOrdFem(8)"
      Tab(16).Control(4)=   "CmdOrdFem(9)"
      Tab(16).Control(5)=   "EditSeq(5)"
      Tab(16).Control(6)=   "Command2(5)"
      Tab(16).Control(7)=   "Command1(20)"
      Tab(16).Control(8)=   "importar(5)"
      Tab(16).Control(9)=   "Command4(5)"
      Tab(16).Control(10)=   "CmdClo(8)"
      Tab(16).Control(11)=   "CmdClo(9)"
      Tab(16).Control(12)=   "CmdREvi(5)"
      Tab(16).Control(13)=   "CmdRevFemea(5)"
      Begin VBCCR18.DTPicker DTPicker2 
         Height          =   375
         Left            =   -72960
         TabIndex        =   254
         Top             =   2400
         Width           =   1695
         _ExtentX        =   2985
         _ExtentY        =   656
         Value           =   44285
         CheckBox        =   -1  'True
         AllowUserInput  =   -1  'True
      End
      Begin VBCCR18.DTPicker DTPicker1 
         Height          =   375
         Left            =   -68880
         TabIndex        =   253
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2773
         _ExtentY        =   656
         Value           =   44285
         CheckBox        =   -1  'True
         AllowUserInput  =   -1  'True
      End
      Begin VB.TextBox Txtfields 
         Height          =   285
         Index           =   33
         Left            =   -73800
         MaxLength       =   100
         TabIndex        =   250
         Top             =   3840
         Width           =   8055
      End
      Begin VB.TextBox Txtfields 
         Height          =   285
         Index           =   32
         Left            =   -73800
         MaxLength       =   100
         TabIndex        =   249
         Top             =   3360
         Width           =   2415
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Desconsiderar Apuraçao RPN"
         Height          =   255
         Left            =   -67080
         TabIndex        =   248
         TabStop         =   0   'False
         Top             =   2760
         Width           =   2475
      End
      Begin VB.CommandButton escidfolha 
         Caption         =   "LX"
         Height          =   495
         Index           =   11
         Left            =   -65040
         Picture         =   "Frmfemea.frx":1628
         Style           =   1  'Graphical
         TabIndex        =   247
         TabStop         =   0   'False
         ToolTipText     =   "Escolher Nº Folha"
         Top             =   2040
         Width           =   615
      End
      Begin VB.CommandButton escidfolha 
         Caption         =   "LX"
         Height          =   495
         Index           =   10
         Left            =   -65040
         Picture         =   "Frmfemea.frx":1BB2
         Style           =   1  'Graphical
         TabIndex        =   246
         TabStop         =   0   'False
         ToolTipText     =   "Escolher Nº Folha"
         Top             =   1320
         Width           =   615
      End
      Begin VB.CommandButton CmdOrdFem 
         Caption         =   "Ordem Seq,Ssq,FX,Item"
         Height          =   255
         Index           =   11
         Left            =   1920
         TabIndex        =   245
         TabStop         =   0   'False
         Top             =   4680
         Width           =   1935
      End
      Begin VB.CommandButton CmdOrdFem 
         Caption         =   "Ordem Reg"
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   244
         TabStop         =   0   'False
         Top             =   4680
         Width           =   1575
      End
      Begin VB.CommandButton CmdREvisarTodas 
         Caption         =   "Rev Todas Sem Indicadores"
         Height          =   615
         Index           =   16
         Left            =   9360
         TabIndex        =   241
         TabStop         =   0   'False
         ToolTipText     =   "Revisa todos os itens e gerar um revisao mantendo os indicadores"
         Top             =   6240
         Width           =   1335
      End
      Begin VB.CommandButton CmdRevFemea 
         Caption         =   "Nova Rev Sem Indicado"
         Height          =   375
         Index           =   16
         Left            =   9360
         TabIndex        =   240
         TabStop         =   0   'False
         ToolTipText     =   "Gera nova revisao do item  mantendo os indicadore"
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton CmdREvisarTodas 
         Caption         =   "Rev. Todas"
         Height          =   375
         Index           =   0
         Left            =   9360
         TabIndex        =   239
         TabStop         =   0   'False
         ToolTipText     =   "Revisa todos os itens e gerar um revisao trocando os indicadores"
         Top             =   5880
         Width           =   1335
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Marcar Producao"
         Height          =   375
         Left            =   -71160
         TabIndex        =   238
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton CmdRevFemea 
         Caption         =   "Nova Rev."
         Height          =   255
         Index           =   5
         Left            =   -65520
         TabIndex        =   235
         Top             =   4800
         Width           =   1095
      End
      Begin VB.CommandButton CmdREvi 
         Caption         =   "Ver Revisoes"
         Height          =   255
         Index           =   5
         Left            =   -65520
         TabIndex        =   234
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton CmdClo 
         Caption         =   "Concluir Todos"
         Height          =   495
         Index           =   9
         Left            =   -65520
         TabIndex        =   232
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1095
      End
      Begin VB.CommandButton CmdClo 
         Caption         =   "Concluir Item"
         Height          =   255
         Index           =   8
         Left            =   -65520
         TabIndex        =   231
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Renumerar"
         Height          =   255
         Index           =   5
         Left            =   -65520
         TabIndex        =   230
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton importar 
         Caption         =   "Importar"
         Height          =   255
         Index           =   5
         Left            =   -65520
         TabIndex        =   229
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Novo"
         Height          =   255
         Index           =   20
         Left            =   -65520
         TabIndex        =   228
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Excluir"
         Height          =   255
         Index           =   5
         Left            =   -65520
         TabIndex        =   227
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CommandButton EditSeq 
         Caption         =   "Editar"
         Height          =   255
         Index           =   5
         Left            =   -65520
         TabIndex        =   226
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CommandButton CmdOrdFem 
         Caption         =   "Ordenar Item"
         Height          =   255
         Index           =   9
         Left            =   -74880
         TabIndex        =   225
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton CmdOrdFem 
         Caption         =   "Ordenar Fx Seq,Ssq,Item"
         Height          =   255
         Index           =   8
         Left            =   -73200
         TabIndex        =   224
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Excluir Todos"
         Height          =   255
         Index           =   15
         Left            =   -65520
         TabIndex        =   223
         TabStop         =   0   'False
         Top             =   4080
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Distribuicao"
         Height          =   495
         Index           =   8
         Left            =   -67440
         TabIndex        =   222
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Excluir Todos"
         Height          =   255
         Index           =   10
         Left            =   9480
         TabIndex        =   221
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox Txtfields 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   25
         Left            =   -71520
         Locked          =   -1  'True
         TabIndex        =   220
         TabStop         =   0   'False
         Top             =   1980
         Width           =   1215
      End
      Begin VB.CommandButton Command15 
         Caption         =   "PF->Revisao"
         Height          =   255
         Left            =   -72840
         TabIndex        =   219
         TabStop         =   0   'False
         ToolTipText     =   "Revisao do PF nao Gera Revisao Femea"
         Top             =   3000
         Width           =   1575
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Mover Falha historico"
         Height          =   495
         Left            =   -66120
         TabIndex        =   217
         TabStop         =   0   'False
         Top             =   2340
         Width           =   1455
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Mover Todos Historico"
         Height          =   495
         Left            =   -66120
         TabIndex        =   216
         TabStop         =   0   'False
         Top             =   2940
         Width           =   1455
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Mover Todos Pre"
         Height          =   255
         Index           =   1
         Left            =   -67440
         TabIndex        =   215
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1575
      End
      Begin VB.CommandButton Command12 
         Caption         =   "Mover Pre"
         Height          =   255
         Left            =   -68400
         TabIndex        =   214
         TabStop         =   0   'False
         Top             =   1020
         Width           =   975
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Mover Historico"
         Height          =   255
         Left            =   -71520
         TabIndex        =   213
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1335
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Mover Todos Historico"
         Height          =   255
         Index           =   0
         Left            =   -70200
         TabIndex        =   212
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1815
      End
      Begin VB.CommandButton CmdREvi 
         Caption         =   "Ver Historico"
         Height          =   255
         Index           =   4
         Left            =   -65400
         TabIndex        =   210
         TabStop         =   0   'False
         Top             =   1380
         Width           =   1095
      End
      Begin VB.TextBox Txtfields 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   24
         Left            =   -67200
         Locked          =   -1  'True
         TabIndex        =   209
         TabStop         =   0   'False
         Top             =   1500
         Width           =   1215
      End
      Begin VB.TextBox Txtfields 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   23
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   208
         TabStop         =   0   'False
         Top             =   1980
         Width           =   1215
      End
      Begin VB.TextBox Txtfields 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   22
         Left            =   -67200
         Locked          =   -1  'True
         TabIndex        =   207
         TabStop         =   0   'False
         Top             =   1140
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Novo GP12"
         Height          =   255
         Index           =   13
         Left            =   -65760
         TabIndex        =   205
         TabStop         =   0   'False
         Top             =   1980
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Novo PRE"
         Height          =   255
         Index           =   12
         Left            =   -65760
         TabIndex        =   204
         TabStop         =   0   'False
         Top             =   1740
         Width           =   1455
      End
      Begin VB.CommandButton Cmdcheckfx 
         Caption         =   "Checar gp12"
         Height          =   375
         Index           =   3
         Left            =   -65760
         TabIndex        =   203
         TabStop         =   0   'False
         Top             =   5460
         Width           =   1455
      End
      Begin VB.CommandButton Cmdcheckfx 
         Caption         =   "Checar pre"
         Height          =   375
         Index           =   2
         Left            =   -65760
         TabIndex        =   202
         TabStop         =   0   'False
         Top             =   5100
         Width           =   1455
      End
      Begin VB.CommandButton CmdTotFlux 
         Caption         =   "Todos p/ gp12"
         Height          =   375
         Index           =   3
         Left            =   -65760
         TabIndex        =   201
         TabStop         =   0   'False
         Top             =   3540
         Width           =   1455
      End
      Begin VB.CommandButton CmdTotFlux 
         Caption         =   "Todos p/ pre"
         Height          =   375
         Index           =   2
         Left            =   -65760
         TabIndex        =   200
         TabStop         =   0   'False
         Top             =   3180
         Width           =   1455
      End
      Begin VB.CommandButton CmdClo 
         Caption         =   "Concluir Todos"
         Height          =   495
         Index           =   7
         Left            =   -65400
         TabIndex        =   197
         TabStop         =   0   'False
         Top             =   3420
         Width           =   1095
      End
      Begin VB.CommandButton CmdClo 
         Caption         =   "Concluir Item"
         Height          =   255
         Index           =   6
         Left            =   -65400
         TabIndex        =   196
         TabStop         =   0   'False
         Top             =   3060
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Renumerar"
         Height          =   255
         Index           =   3
         Left            =   -65400
         TabIndex        =   195
         TabStop         =   0   'False
         Top             =   2700
         Width           =   1095
      End
      Begin VB.CommandButton importar 
         Caption         =   "Importar"
         Height          =   255
         Index           =   3
         Left            =   -65400
         TabIndex        =   194
         TabStop         =   0   'False
         Top             =   2340
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Novo"
         Height          =   255
         Index           =   7
         Left            =   -65400
         TabIndex        =   193
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Excluir"
         Height          =   255
         Index           =   3
         Left            =   -65400
         TabIndex        =   192
         TabStop         =   0   'False
         Top             =   1620
         Width           =   1095
      End
      Begin VB.CommandButton EditSeq 
         Caption         =   "Editar"
         Height          =   255
         Index           =   3
         Left            =   -65400
         TabIndex        =   191
         TabStop         =   0   'False
         Top             =   1980
         Width           =   1095
      End
      Begin VB.CommandButton CmdOrdFem 
         Caption         =   "Ordenar Item"
         Height          =   255
         Index           =   7
         Left            =   -74880
         TabIndex        =   190
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1575
      End
      Begin VB.CommandButton CmdOrdFem 
         Caption         =   "Ordenar Fx Seq,Ssq,Item"
         Height          =   255
         Index           =   6
         Left            =   -73200
         TabIndex        =   189
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1935
      End
      Begin VB.CommandButton CmdRevFemea 
         Caption         =   "Nova Rev."
         Height          =   255
         Index           =   3
         Left            =   -65400
         TabIndex        =   188
         TabStop         =   0   'False
         Top             =   4740
         Width           =   1095
      End
      Begin VB.CommandButton CmdREvi 
         Caption         =   "Ver Revisoes"
         Height          =   255
         Index           =   3
         Left            =   -65400
         TabIndex        =   187
         TabStop         =   0   'False
         Top             =   5100
         Width           =   1095
      End
      Begin VB.CommandButton CmdClo 
         Caption         =   "Concluir Todos"
         Height          =   495
         Index           =   5
         Left            =   -65520
         TabIndex        =   184
         TabStop         =   0   'False
         Top             =   3420
         Width           =   1095
      End
      Begin VB.CommandButton CmdClo 
         Caption         =   "Concluir Item"
         Height          =   255
         Index           =   4
         Left            =   -65520
         TabIndex        =   183
         TabStop         =   0   'False
         Top             =   3060
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Renumerar"
         Height          =   255
         Index           =   2
         Left            =   -65520
         TabIndex        =   182
         TabStop         =   0   'False
         Top             =   2700
         Width           =   1095
      End
      Begin VB.CommandButton importar 
         Caption         =   "Importar"
         Height          =   255
         Index           =   2
         Left            =   -65520
         TabIndex        =   181
         TabStop         =   0   'False
         Top             =   2340
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Novo"
         Height          =   255
         Index           =   6
         Left            =   -65520
         TabIndex        =   180
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Excluir"
         Height          =   255
         Index           =   2
         Left            =   -65520
         TabIndex        =   179
         TabStop         =   0   'False
         Top             =   1620
         Width           =   1095
      End
      Begin VB.CommandButton EditSeq 
         Caption         =   "Editar"
         Height          =   255
         Index           =   2
         Left            =   -65520
         TabIndex        =   178
         TabStop         =   0   'False
         Top             =   1980
         Width           =   1095
      End
      Begin VB.CommandButton CmdOrdFem 
         Caption         =   "Ordenar Item"
         Height          =   255
         Index           =   5
         Left            =   -74880
         TabIndex        =   177
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1575
      End
      Begin VB.CommandButton CmdOrdFem 
         Caption         =   "Ordenar Fx Seq,Ssq,Item"
         Height          =   255
         Index           =   4
         Left            =   -73200
         TabIndex        =   176
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1935
      End
      Begin VB.CommandButton CmdRevFemea 
         Caption         =   "Nova Rev."
         Height          =   255
         Index           =   2
         Left            =   -65520
         TabIndex        =   175
         TabStop         =   0   'False
         Top             =   4740
         Width           =   1095
      End
      Begin VB.CommandButton CmdREvi 
         Caption         =   "Ver Revisoes"
         Height          =   255
         Index           =   2
         Left            =   -65520
         TabIndex        =   174
         TabStop         =   0   'False
         Top             =   5100
         Width           =   1095
      End
      Begin VB.CommandButton CmdREvi 
         Caption         =   "Ver Revisoes"
         Height          =   255
         Index           =   1
         Left            =   -65400
         TabIndex        =   173
         TabStop         =   0   'False
         Top             =   5100
         Width           =   1095
      End
      Begin VB.CommandButton CmdREvi 
         Caption         =   "Ver Revisao"
         Height          =   255
         Index           =   0
         Left            =   9360
         TabIndex        =   172
         TabStop         =   0   'False
         Top             =   5040
         Width           =   1335
      End
      Begin VB.CommandButton CmdRevFemea 
         Caption         =   "Nova Rev."
         Height          =   255
         Index           =   1
         Left            =   -65400
         TabIndex        =   171
         TabStop         =   0   'False
         Top             =   4740
         Width           =   1095
      End
      Begin VB.CommandButton CmdRevFemea 
         Caption         =   "Nova Revisao"
         Height          =   255
         Index           =   0
         Left            =   9360
         TabIndex        =   170
         TabStop         =   0   'False
         ToolTipText     =   "Gera Revisao do itens com os novos indicadores"
         Top             =   5280
         Width           =   1335
      End
      Begin VB.CommandButton CmdEdiDef 
         Caption         =   "Editar"
         Height          =   375
         Left            =   -66000
         TabIndex        =   167
         TabStop         =   0   'False
         Top             =   2700
         Width           =   1335
      End
      Begin VB.CommandButton CmdDelDef 
         Caption         =   "Excluir"
         Height          =   375
         Left            =   -66000
         TabIndex        =   166
         TabStop         =   0   'False
         Top             =   2220
         Width           =   1335
      End
      Begin VB.CommandButton CmdIncDef 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   -66000
         TabIndex        =   165
         TabStop         =   0   'False
         Top             =   1740
         Width           =   1335
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   255
         Index           =   1
         Left            =   -70680
         Picture         =   "Frmfemea.frx":213C
         Style           =   1  'Graphical
         TabIndex        =   164
         TabStop         =   0   'False
         Top             =   1260
         Width           =   375
      End
      Begin VB.CommandButton CmdcheckOp 
         Caption         =   "Checar Comp"
         Height          =   375
         Index           =   1
         Left            =   -65760
         TabIndex        =   163
         TabStop         =   0   'False
         Top             =   3420
         Width           =   1455
      End
      Begin VB.CommandButton CmdcheckOp 
         Caption         =   "Checar Padrao"
         Height          =   375
         Index           =   0
         Left            =   -65760
         TabIndex        =   162
         TabStop         =   0   'False
         Top             =   2940
         Width           =   1455
      End
      Begin VB.CommandButton Cmdcheckfx 
         Caption         =   "Checar Compl"
         Height          =   375
         Index           =   1
         Left            =   -65760
         TabIndex        =   161
         TabStop         =   0   'False
         Top             =   4740
         Width           =   1455
      End
      Begin VB.CommandButton Cmdcheckfx 
         Caption         =   "Checar Normais"
         Height          =   375
         Index           =   0
         Left            =   -65760
         TabIndex        =   160
         TabStop         =   0   'False
         Top             =   4380
         Width           =   1455
      End
      Begin VB.CommandButton CmdTotFlux 
         Caption         =   "Todos p/ Compl"
         Height          =   375
         Index           =   1
         Left            =   -65760
         TabIndex        =   159
         TabStop         =   0   'False
         Top             =   2820
         Width           =   1455
      End
      Begin VB.CommandButton CmdTotFlux 
         Caption         =   "Todos p/ Normais"
         Height          =   375
         Index           =   0
         Left            =   -65760
         TabIndex        =   158
         TabStop         =   0   'False
         Top             =   2460
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Novo Padrao"
         Height          =   255
         Index           =   4
         Left            =   -65760
         TabIndex        =   157
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Novo Compl"
         Height          =   255
         Index           =   5
         Left            =   -65760
         TabIndex        =   156
         TabStop         =   0   'False
         Top             =   1500
         Width           =   1455
      End
      Begin VB.CommandButton CmdTotPad 
         Caption         =   "Todos Compl"
         Height          =   375
         Index           =   1
         Left            =   -65760
         TabIndex        =   154
         TabStop         =   0   'False
         Top             =   2460
         Width           =   1455
      End
      Begin VB.CommandButton CmdTotPad 
         Caption         =   "Todos Padrao"
         Height          =   375
         Index           =   0
         Left            =   -65760
         TabIndex        =   153
         TabStop         =   0   'False
         Top             =   1980
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Novo Compl"
         Height          =   255
         Index           =   3
         Left            =   -65760
         TabIndex        =   152
         TabStop         =   0   'False
         Top             =   1620
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Novo Padrao"
         Height          =   255
         Index           =   2
         Left            =   -65760
         TabIndex        =   151
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1455
      End
      Begin VB.CommandButton CmdOrdFem 
         Caption         =   "Ordenar Fx Seq,Ssq,Item"
         Height          =   255
         Index           =   3
         Left            =   -73680
         TabIndex        =   150
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1935
      End
      Begin VB.CommandButton CmdOrdFem 
         Caption         =   "Ordenar Item"
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   149
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1215
      End
      Begin VB.CommandButton CmdOrdFem 
         Caption         =   "Ordem Fx Seq,Ssq,Item"
         Height          =   255
         Index           =   1
         Left            =   1800
         TabIndex        =   148
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1935
      End
      Begin VB.CommandButton CmdOrdFem 
         Caption         =   "Ordem Reg"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   147
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1575
      End
      Begin VB.CommandButton CmdVerDef 
         Caption         =   "Ver"
         Height          =   375
         Left            =   -66000
         TabIndex        =   146
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1335
      End
      Begin VB.CommandButton CmdMarcaFemea 
         Caption         =   "Verificado"
         Height          =   375
         Left            =   -66000
         TabIndex        =   145
         TabStop         =   0   'False
         Top             =   3420
         Width           =   1455
      End
      Begin VBFLXGRD18.VBFlexGrid GridPFrev 
         Height          =   1215
         Left            =   -74760
         TabIndex        =   144
         TabStop         =   0   'False
         Top             =   3420
         Width           =   8655
         _ExtentX        =   15261
         _ExtentY        =   2138
      End
      Begin VB.CommandButton escidfolha 
         Height          =   495
         Index           =   1
         Left            =   -65640
         Picture         =   "Frmfemea.frx":26C6
         Style           =   1  'Graphical
         TabIndex        =   143
         TabStop         =   0   'False
         ToolTipText     =   "Escolher Nº Folha"
         Top             =   2100
         Width           =   495
      End
      Begin VB.CommandButton escidfolha 
         Height          =   495
         Index           =   0
         Left            =   -65640
         Picture         =   "Frmfemea.frx":2C50
         Style           =   1  'Graphical
         TabIndex        =   142
         TabStop         =   0   'False
         ToolTipText     =   "Escolher Nº Folha"
         Top             =   1380
         Width           =   495
      End
      Begin VB.TextBox Txtfields 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   17
         Left            =   -65280
         Locked          =   -1  'True
         TabIndex        =   140
         TabStop         =   0   'False
         Top             =   1620
         Width           =   975
      End
      Begin VB.TextBox Txtfields 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   16
         Left            =   -66000
         Locked          =   -1  'True
         TabIndex        =   139
         TabStop         =   0   'False
         Top             =   1620
         Width           =   615
      End
      Begin VB.CommandButton CmdEdiPad 
         Caption         =   "Todos"
         Height          =   375
         Index           =   10
         Left            =   -65100
         TabIndex        =   137
         TabStop         =   0   'False
         Top             =   1500
         Width           =   855
      End
      Begin VB.CommandButton importar 
         Caption         =   "Importar"
         Height          =   255
         Index           =   10
         Left            =   -69240
         TabIndex        =   136
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton CmdOrdPad 
         Caption         =   "RPN>"
         Height          =   375
         Index           =   3
         Left            =   -72360
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   1020
         Width           =   615
      End
      Begin VB.CommandButton CmdApagaRPN 
         Caption         =   "Apagar RPN"
         Height          =   375
         Left            =   -65040
         TabIndex        =   132
         TabStop         =   0   'False
         Top             =   1020
         Width           =   735
      End
      Begin VB.TextBox txtRPN 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.4
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -65640
         MaxLength       =   1
         TabIndex        =   131
         Top             =   1020
         Width           =   615
      End
      Begin VB.CommandButton CmdTrocaREv 
         Caption         =   "Trocar Revisao"
         Height          =   375
         Left            =   -66240
         TabIndex        =   130
         TabStop         =   0   'False
         Top             =   3660
         Width           =   1335
      End
      Begin VB.TextBox TxtRevNova 
         Height          =   375
         Left            =   -64920
         TabIndex        =   129
         TabStop         =   0   'False
         Top             =   3660
         Width           =   615
      End
      Begin VB.TextBox TxtRevisao 
         Height          =   375
         Left            =   -64920
         TabIndex        =   128
         TabStop         =   0   'False
         Top             =   3300
         Width           =   615
      End
      Begin VB.CommandButton CmdApaRev 
         Caption         =   "Apagar Revisao"
         Height          =   375
         Left            =   -66240
         TabIndex        =   127
         TabStop         =   0   'False
         Top             =   3300
         Width           =   1335
      End
      Begin VB.TextBox Txtfields 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   -73920
         Locked          =   -1  'True
         TabIndex        =   126
         TabStop         =   0   'False
         Top             =   1500
         Width           =   735
      End
      Begin VB.TextBox Txtfields 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   -72600
         Locked          =   -1  'True
         TabIndex        =   125
         TabStop         =   0   'False
         Top             =   1500
         Width           =   4335
      End
      Begin VB.CommandButton CmdLiberar 
         Caption         =   "Liberar"
         Height          =   255
         Left            =   -65880
         TabIndex        =   124
         TabStop         =   0   'False
         Top             =   1500
         Width           =   1095
      End
      Begin VB.CommandButton CmdExcPFPad 
         Caption         =   "Excluir Processo"
         Height          =   375
         Index           =   1
         Left            =   -66120
         TabIndex        =   123
         TabStop         =   0   'False
         Top             =   1380
         Width           =   1455
      End
      Begin VB.CommandButton EditSeq 
         Caption         =   "Editar"
         Height          =   255
         Index           =   0
         Left            =   9480
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Excluir"
         Height          =   255
         Index           =   0
         Left            =   9480
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Novo"
         Height          =   255
         Index           =   0
         Left            =   9480
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1095
      End
      Begin VB.TextBox Txtfields 
         Height          =   285
         Index           =   18
         Left            =   -73800
         MaxLength       =   100
         TabIndex        =   6
         Top             =   2880
         Width           =   8055
      End
      Begin VB.TextBox Txtfields 
         Height          =   645
         Index           =   8
         Left            =   -73800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   2160
         Width           =   8055
      End
      Begin VB.TextBox Txtfields 
         Height          =   645
         Index           =   7
         Left            =   -73800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1440
         Width           =   8055
      End
      Begin VB.TextBox Txtfields 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   -73920
         Locked          =   -1  'True
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   1980
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Height          =   255
         Left            =   -73080
         Picture         =   "Frmfemea.frx":31DA
         Style           =   1  'Graphical
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox Txtfields 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   -72600
         Locked          =   -1  'True
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   1080
         Width           =   4335
      End
      Begin VB.TextBox Txtfields 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   -73920
         Locked          =   -1  'True
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   1080
         Width           =   735
      End
      Begin VB.CommandButton ComMotInc 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   -67440
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   4800
         Width           =   1335
      End
      Begin VB.CommandButton ComMotExc 
         Caption         =   "Excluir"
         Height          =   375
         Left            =   -67440
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   5280
         Width           =   1335
      End
      Begin VB.CommandButton ComMotAlt 
         Caption         =   "Alterar"
         Height          =   375
         Index           =   0
         Left            =   -67440
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton CmdMotRev 
         Caption         =   "Motivos"
         Height          =   375
         Left            =   -67440
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1095
      End
      Begin VB.CommandButton CmdNovaREV 
         Caption         =   "Nova Revisão"
         Height          =   255
         Left            =   -74760
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1695
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Indicador Revisão PF"
         Height          =   255
         Left            =   -67080
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CommandButton importar 
         Caption         =   "Importar"
         Height          =   255
         Index           =   0
         Left            =   9480
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Renumerar"
         Height          =   255
         Index           =   0
         Left            =   9480
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1095
      End
      Begin VB.CommandButton ComMotAlt 
         Caption         =   "Ver"
         Height          =   375
         Index           =   1
         Left            =   -66000
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   1140
         Width           =   555
      End
      Begin VB.CommandButton CmdClo 
         Caption         =   "Concluir Item"
         Height          =   255
         Index           =   0
         Left            =   9480
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton CmdClo 
         Caption         =   "Concluir Todos"
         Height          =   495
         Index           =   1
         Left            =   9480
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CommandButton CmdProx 
         Caption         =   "-->"
         Height          =   375
         Left            =   -70200
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   1980
         Width           =   495
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Nova Revisão Periodica"
         Height          =   255
         Left            =   -70920
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   3000
         Width           =   2175
      End
      Begin VB.CheckBox femeacrg 
         Caption         =   "Nao Requer Cronograma"
         Height          =   255
         Left            =   -67080
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1980
         Width           =   2295
      End
      Begin VB.CommandButton CmdEscPro 
         Caption         =   "Escolher"
         Height          =   375
         Left            =   -69840
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   1380
         Width           =   735
      End
      Begin VB.CommandButton CmdEscFal 
         Caption         =   "Escolher"
         Height          =   375
         Left            =   -69840
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   3060
         Width           =   735
      End
      Begin VB.CommandButton CmdIncPf 
         Caption         =   "-->"
         Height          =   375
         Left            =   -69840
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   2460
         Width           =   735
      End
      Begin VB.CommandButton CmdExcPFPad 
         Caption         =   "Excluir Falha"
         Height          =   375
         Index           =   0
         Left            =   -66120
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   1860
         Width           =   1455
      End
      Begin VB.CommandButton CmdClo 
         Caption         =   "Concluir Todos"
         Height          =   495
         Index           =   3
         Left            =   -65400
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   3060
         Width           =   1095
      End
      Begin VB.CommandButton CmdClo 
         Caption         =   "Concluir Item"
         Height          =   255
         Index           =   2
         Left            =   -65400
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   2820
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Renumerar"
         Height          =   255
         Index           =   1
         Left            =   -65400
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   2460
         Width           =   1095
      End
      Begin VB.CommandButton importar 
         Caption         =   "Importar"
         Height          =   255
         Index           =   1
         Left            =   -65400
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   2220
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Novo"
         Height          =   255
         Index           =   1
         Left            =   -65400
         TabIndex        =   63
         Top             =   1140
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Excluir"
         Height          =   255
         Index           =   1
         Left            =   -65400
         TabIndex        =   62
         Top             =   1860
         Width           =   1095
      End
      Begin VB.CommandButton EditSeq 
         Caption         =   "Editar"
         Height          =   255
         Index           =   1
         Left            =   -65400
         TabIndex        =   61
         Top             =   1500
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Cadastro Padrao"
         Height          =   255
         Left            =   -69960
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1695
      End
      Begin VB.CommandButton CmdPreinc 
         Caption         =   "Novo"
         Height          =   255
         Left            =   -66480
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1095
      End
      Begin VB.CommandButton CmddelPre 
         Caption         =   "Excluir"
         Height          =   255
         Left            =   -66480
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1620
         Width           =   1095
      End
      Begin VB.CommandButton cmdedipre 
         Caption         =   "Editar"
         Height          =   255
         Index           =   2
         Left            =   -66480
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1980
         Width           =   1095
      End
      Begin VB.CommandButton CmdEdiPad 
         Caption         =   "Editar"
         Height          =   375
         Index           =   0
         Left            =   -69240
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   1020
         Width           =   735
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Localizar"
         Height          =   375
         Left            =   -66480
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   2340
         Width           =   1095
      End
      Begin VB.CommandButton CmdEdiPad 
         Caption         =   "Excluir"
         Height          =   375
         Index           =   1
         Left            =   -68400
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   1020
         Width           =   735
      End
      Begin VB.CommandButton Command8 
         Caption         =   "-->>"
         Height          =   375
         Left            =   -69840
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1860
         Width           =   735
      End
      Begin VB.CommandButton CmdNome 
         Caption         =   "Renomear"
         Height          =   375
         Left            =   -69120
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1860
         Width           =   855
      End
      Begin VB.CommandButton Cmdimp 
         Caption         =   "Importar"
         Height          =   255
         Left            =   -68160
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   1020
         Width           =   735
      End
      Begin VB.CommandButton CmdEdiPad 
         Caption         =   "Efeitos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   -66480
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1500
         Width           =   1335
      End
      Begin VB.CommandButton CmdOrdPad 
         Caption         =   "Processo"
         Height          =   375
         Index           =   0
         Left            =   -73800
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1020
         Width           =   855
      End
      Begin VB.CommandButton CmdOrdPad 
         Caption         =   "Item"
         Height          =   375
         Index           =   1
         Left            =   -74280
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1020
         Width           =   495
      End
      Begin VB.CommandButton CmdOrdPad 
         Caption         =   "RPN<"
         Height          =   375
         Index           =   2
         Left            =   -72960
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1020
         Width           =   615
      End
      Begin VB.CommandButton CmdEdiPad 
         Caption         =   "Atribuir Item"
         Height          =   375
         Index           =   3
         Left            =   -72960
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1500
         Width           =   975
      End
      Begin VB.CommandButton maisano 
         Caption         =   "+"
         Height          =   240
         Left            =   -71280
         TabIndex        =   45
         TabStop         =   0   'False
         ToolTipText     =   "Aumenta um Ano"
         Top             =   1020
         Width           =   330
      End
      Begin VB.CommandButton menosano 
         Caption         =   "-"
         Height          =   240
         Left            =   -70920
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "Diminui um Ano"
         Top             =   1020
         Width           =   330
      End
      Begin VB.CommandButton maismes 
         Caption         =   "+"
         Height          =   240
         Left            =   -73440
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "Aumenta um mes"
         Top             =   1020
         Width           =   330
      End
      Begin VB.CommandButton menosmes 
         Caption         =   "-"
         Height          =   240
         Left            =   -73080
         TabIndex        =   42
         TabStop         =   0   'False
         ToolTipText     =   "Diminui um mes"
         Top             =   1020
         Width           =   330
      End
      Begin VB.TextBox text 
         Alignment       =   1  'Right Justify
         Height          =   330
         Index           =   2
         Left            =   -72240
         TabIndex        =   41
         Text            =   "2006"
         Top             =   1020
         Width           =   810
      End
      Begin VB.TextBox text 
         Alignment       =   1  'Right Justify
         Height          =   330
         Index           =   1
         Left            =   -74160
         TabIndex        =   40
         Text            =   "1"
         Top             =   1020
         Width           =   690
      End
      Begin VB.CommandButton CmdVerApu 
         Caption         =   "Ver Competencia"
         Height          =   375
         Left            =   -70440
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1335
      End
      Begin VB.CommandButton Cmdeditar 
         Caption         =   "Editar"
         Height          =   375
         Left            =   -65880
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1500
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Editar"
         Height          =   375
         Left            =   -65880
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   3300
         Width           =   1215
      End
      Begin VB.CommandButton CmdEdiPad 
         Caption         =   "Falhas"
         Height          =   375
         Index           =   4
         Left            =   -67440
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   1500
         Width           =   855
      End
      Begin VB.CommandButton CmdEdiPad 
         Caption         =   "Processos"
         Height          =   375
         Index           =   5
         Left            =   -66840
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   1020
         Width           =   975
      End
      Begin VB.CommandButton CMDSIG 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -71040
         Picture         =   "Frmfemea.frx":3764
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   1260
         Width           =   375
      End
      Begin VB.TextBox txtsigi 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -71520
         MaxLength       =   1
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1020
         Width           =   495
      End
      Begin VB.CommandButton CmdEdiPad 
         Caption         =   "Efeitos"
         Height          =   375
         Index           =   6
         Left            =   -70320
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Marca Siginificativa todos os Efeitos"
         Top             =   1620
         Width           =   855
      End
      Begin VB.CommandButton CmdEdiPad 
         Caption         =   "Falhas"
         Height          =   375
         Index           =   7
         Left            =   -71520
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Marca Siginificativa todos as Falhas"
         Top             =   1620
         Width           =   855
      End
      Begin VB.CommandButton CmdEdiPad 
         Caption         =   "Processo"
         Height          =   375
         Index           =   8
         Left            =   -70320
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Marca Siginificativa todos os Processos"
         Top             =   1260
         Width           =   855
      End
      Begin VB.CommandButton CmdEdiPad 
         Caption         =   "Renomear"
         Height          =   375
         Index           =   9
         Left            =   -73920
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1500
         Width           =   855
      End
      Begin VBFLXGRD18.VBFlexGrid GridSeq 
         Height          =   3255
         Index           =   0
         Left            =   120
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   1380
         Width           =   9255
         _ExtentX        =   16320
         _ExtentY        =   5736
         Cols            =   8
      End
      Begin VBFLXGRD18.VBFlexGrid gridrev 
         Height          =   1335
         Left            =   -74880
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   3360
         Width           =   7455
         _ExtentX        =   13145
         _ExtentY        =   2350
      End
      Begin VBFLXGRD18.VBFlexGrid gridrevi 
         Height          =   1095
         Index           =   0
         Left            =   -74880
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   4680
         Width           =   7455
         _ExtentX        =   13145
         _ExtentY        =   1926
      End
      Begin VBFLXGRD18.VBFlexGrid gridopr 
         Height          =   2955
         Left            =   -74880
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   1140
         Width           =   9075
         _ExtentX        =   16002
         _ExtentY        =   5207
      End
      Begin VBFLXGRD18.VBFlexGrid gridrevi 
         Height          =   2115
         Index           =   1
         Left            =   -74820
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   1140
         Width           =   8655
         _ExtentX        =   15261
         _ExtentY        =   3725
      End
      Begin VBFLXGRD18.VBFlexGrid gridmsrd 
         Height          =   2955
         Left            =   -74880
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   1200
         Width           =   8715
         _ExtentX        =   15367
         _ExtentY        =   5207
      End
      Begin VBFLXGRD18.VBFlexGrid GridPRO 
         Height          =   1155
         Left            =   -74760
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   1380
         Width           =   4875
         _ExtentX        =   8594
         _ExtentY        =   2032
      End
      Begin VBFLXGRD18.VBFlexGrid GRIDFAL 
         Height          =   1155
         Left            =   -74760
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   2820
         Width           =   4875
         _ExtentX        =   8594
         _ExtentY        =   2032
      End
      Begin VBFLXGRD18.VBFlexGrid GridPadPF 
         Height          =   2715
         Left            =   -68280
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   1380
         Width           =   2115
         _ExtentX        =   3725
         _ExtentY        =   4784
      End
      Begin VBFLXGRD18.VBFlexGrid GridSeq 
         Height          =   3375
         Index           =   1
         Left            =   -74880
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   1380
         Width           =   9375
         _ExtentX        =   16531
         _ExtentY        =   5948
      End
      Begin VBFLXGRD18.VBFlexGrid Gridfemprecad 
         Height          =   4095
         Left            =   -74880
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   1020
         Width           =   8175
         _ExtentX        =   14415
         _ExtentY        =   7218
      End
      Begin VBFLXGRD18.VBFlexGrid GridPad 
         Height          =   4095
         Left            =   -74880
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   1980
         Width           =   10575
         _ExtentX        =   18648
         _ExtentY        =   7218
      End
      Begin VBFLXGRD18.VBFlexGrid Grid10 
         Height          =   1695
         Left            =   -74880
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   1500
         Width           =   8835
         _ExtentX        =   15579
         _ExtentY        =   2985
      End
      Begin VBFLXGRD18.VBFlexGrid GridOpe 
         Height          =   1695
         Left            =   -74880
         TabIndex        =   104
         TabStop         =   0   'False
         Top             =   3420
         Width           =   8835
         _ExtentX        =   15579
         _ExtentY        =   2985
      End
      Begin VBFLXGRD18.VBFlexGrid GrdQsbLep 
         Height          =   5115
         Left            =   -74880
         TabIndex        =   155
         TabStop         =   0   'False
         Top             =   1020
         Width           =   8955
         _ExtentX        =   15790
         _ExtentY        =   9017
      End
      Begin VBFLXGRD18.VBFlexGrid gridrevcau 
         Height          =   1995
         Index           =   0
         Left            =   120
         TabIndex        =   168
         TabStop         =   0   'False
         Top             =   5040
         Width           =   9195
         _ExtentX        =   16214
         _ExtentY        =   3514
      End
      Begin VBFLXGRD18.VBFlexGrid gridrevcau 
         Height          =   1515
         Index           =   1
         Left            =   -74880
         TabIndex        =   169
         TabStop         =   0   'False
         Top             =   4740
         Width           =   9315
         _ExtentX        =   16425
         _ExtentY        =   2667
      End
      Begin VBFLXGRD18.VBFlexGrid GridSeq 
         Height          =   3255
         Index           =   2
         Left            =   -74880
         TabIndex        =   185
         TabStop         =   0   'False
         Top             =   1380
         Width           =   9255
         _ExtentX        =   16320
         _ExtentY        =   5736
      End
      Begin VBFLXGRD18.VBFlexGrid gridrevcau 
         Height          =   1635
         Index           =   2
         Left            =   -74880
         TabIndex        =   186
         TabStop         =   0   'False
         Top             =   4620
         Width           =   9315
         _ExtentX        =   16425
         _ExtentY        =   2879
      End
      Begin VBFLXGRD18.VBFlexGrid GridSeq 
         Height          =   3255
         Index           =   3
         Left            =   -74880
         TabIndex        =   198
         TabStop         =   0   'False
         Top             =   1380
         Width           =   9255
         _ExtentX        =   16320
         _ExtentY        =   5736
      End
      Begin VBFLXGRD18.VBFlexGrid gridrevcau 
         Height          =   1635
         Index           =   3
         Left            =   -74880
         TabIndex        =   199
         TabStop         =   0   'False
         Top             =   4620
         Width           =   9315
         _ExtentX        =   16425
         _ExtentY        =   2879
      End
      Begin VBFLXGRD18.VBFlexGrid GridSeq 
         Height          =   3255
         Index           =   4
         Left            =   -74760
         TabIndex        =   211
         TabStop         =   0   'False
         Top             =   1140
         Width           =   9255
         _ExtentX        =   16320
         _ExtentY        =   5736
      End
      Begin VBFLXGRD18.VBFlexGrid gridpfpadhist 
         Height          =   2715
         Left            =   -74640
         TabIndex        =   218
         TabStop         =   0   'False
         Top             =   1140
         Width           =   2115
         _ExtentX        =   3725
         _ExtentY        =   4784
      End
      Begin VBFLXGRD18.VBFlexGrid GridSeq 
         Height          =   3255
         Index           =   5
         Left            =   -74880
         TabIndex        =   233
         Top             =   1440
         Width           =   9255
         _ExtentX        =   16320
         _ExtentY        =   5736
      End
      Begin VBFLXGRD18.VBFlexGrid gridrevcau 
         Height          =   1515
         Index           =   5
         Left            =   -74880
         TabIndex        =   236
         Top             =   4800
         Width           =   9255
         _ExtentX        =   16320
         _ExtentY        =   2667
      End
      Begin VB.Label Label14 
         Caption         =   "Projeto:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   252
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Ano:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   251
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Data Pre Producao"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   237
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   2
         X1              =   -67560
         X2              =   -67560
         Y1              =   900
         Y2              =   1980
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   1
         X1              =   -69360
         X2              =   -69360
         Y1              =   900
         Y2              =   1980
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Index           =   0
         X1              =   -71640
         X2              =   -71640
         Y1              =   900
         Y2              =   1980
      End
      Begin VB.Label lblLabels 
         Caption         =   "Excluir todos"
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   6
         Left            =   -67440
         TabIndex        =   122
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Ordenar"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   121
         Top             =   1020
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Obs:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   120
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Copia Para"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   119
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Grupo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   118
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Revisão"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   117
         Top             =   1980
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Elaborador"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   116
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Liberada"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -68040
         TabIndex        =   115
         Top             =   1500
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Elaborada"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -68040
         TabIndex        =   114
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Proxima"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -69600
         TabIndex        =   113
         Top             =   1980
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Funcao do Processo Requisitos"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   112
         Top             =   1140
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Modo de Falha Potencial"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   111
         Top             =   2580
         Width           =   2055
      End
      Begin VB.Label dizpro 
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -72360
         TabIndex        =   110
         Top             =   1020
         Width           =   3255
      End
      Begin VB.Label dizfal 
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   -72480
         TabIndex        =   109
         Top             =   2460
         Width           =   3375
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ano:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   4
         Left            =   -72660
         TabIndex        =   108
         Top             =   1080
         Width           =   270
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mes:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   3
         Left            =   -74640
         TabIndex        =   107
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label dizcomapu 
         Height          =   375
         Left            =   -68880
         TabIndex        =   106
         Top             =   1020
         Width           =   1455
      End
      Begin VB.Label lblLabels 
         Caption         =   "Siginificativa"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   14
         Left            =   -70920
         TabIndex        =   105
         Top             =   1020
         Width           =   975
      End
   End
   Begin VBCCR18.CommandButtonW Encerrar 
      Height          =   612
      Left            =   10320
      TabIndex        =   255
      Top             =   240
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
      Picture         =   "Frmfemea.frx":3CEE
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW cmdClose 
      Height          =   612
      Left            =   11520
      TabIndex        =   256
      Top             =   240
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
      Picture         =   "Frmfemea.frx":4288
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VB.Label lblLabels 
      Caption         =   "Tipo"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   8
      Left            =   8160
      TabIndex        =   243
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo Final"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   7
      Left            =   8520
      TabIndex        =   206
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo Integrado"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   40
      Left            =   7200
      TabIndex        =   134
      Top             =   1320
      Width           =   1275
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cod.Fiscal"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   34
      Left            =   60
      TabIndex        =   26
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label11 
      Caption         =   "PF"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "Opção"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   19
      Left            =   1920
      TabIndex        =   20
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Caption         =   "Data "
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   11
      Left            =   7560
      TabIndex        =   19
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "Revisão"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   10
      Left            =   6720
      TabIndex        =   18
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nome:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   5
      Left            =   3720
      TabIndex        =   17
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "CLIENTE:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   4
      Left            =   2520
      TabIndex        =   16
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descrição"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   3
      Left            =   2880
      TabIndex        =   15
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "CODIGO:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   700
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "CPF:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   13
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmfemea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private aNUM As Variant
Private aDIZ As Variant
Private cARQFEMEA As String
Private cARQFEMEAPF As String
Private cARQPF As String
Private cARQPFP As String
Private cARQPFG As String
Private cORDFEMPAD As String
Private cORDFEMUSO As String
Private cORDFEMAVU As String
Private cORDFEMPRE As String
Private cORDFEMGP12 As String
Private cORDFEMADC As String
Private cORDREVI As String
Dim cSQL As String
Dim aVAL As Variant
Dim aFOR As Variant
Dim aCAM As Variant
Dim aPAD As Variant
Dim nCAMPOS As Integer
Dim iLOOP As Integer

Private Sub CmdApagaRPN_Click()
  Dim cSQL As String
  Dim nRPN As Long
  nRPN = FixInt(txtRPN)
  If nRPN > 0 Then
    If MDG("Apagar RPN =" & nRPN) Then
      'cSQL = "select * from FEMEA WHERE INDRIS=" & nRPN & " AND PF=" & nPF
      cSQL = "select * from FEMEAPAD WHERE INDRIS=" & nRPN & " AND PF=" & nPF
      ApagaSQLP cARQFEMEA, cSQL, "Apagar RPN =" & nRPN
    End If
  End If
End Sub

Private Sub CmdApagaRPN_KeyPress(KeyAscii As Integer)
  KeyAscii = ValiText(KeyAscii, "#NI")
End Sub

Private Sub CmdApaRev_Click()
  Dim nTMP As Integer
  nTMP = FixInt(TxtRevisao, 0)
  ApagaSQL cARQPF, "select * from rev WHERE PF=" & nPF & " AND REVISAO=" & nTMP & " AND TIPO='FE'"
  ApagaSQL cARQPF, "select * from revI WHERE PF=" & nPF & " AND REVISAO=" & nTMP & " AND TIPO='FE'"
  Filgridrev
End Sub

Private Sub Cmdcheckfx_Click(Index As Integer)
  Dim nROWS As Integer
  Dim x As Integer
  Dim nITEM As Integer
  Dim cSQL As String
  Dim nTMPSEQ As Integer
  Dim nTMPSSQ As Integer
  Dim nTMPITEM As Integer
  If Index = 1 Then
    If GridPadPF.Rows < 2 Then
      Alert "Este Femea nao utiliza Padrao"
      Exit Sub
    End If
  End If
  If Index = 0 Then                            'normal com complementar
    If GridSeq(1).Rows > 1 Then
      Alert "Femea com complemetar nao pode usar normal"
      Exit Sub
    End If
  End If
  If Index = 0 Then                            'normal com padrao
    If GridPadPF.Rows > 1 Then
      Alert "Femea com Padrao nao pode usar normal"
      Exit Sub
    End If
  End If



  nROWS = GrdQsbLep.Rows - 1
  For x = 1 To nROWS                           ' zero e o titulo do grid comeca 1
    nITEM = 0
    GrdQsbLep.Row = x
    GrdQsbLep.Col = 0
    nTMPSEQ = FixInt(GrdQsbLep)
    GrdQsbLep.Col = 1
    nTMPSSQ = FixInt(GrdQsbLep)
    GrdQsbLep.Col = 2
    nTMPITEM = FixInt(GrdQsbLep)
    GrdQsbLep.Col = 3
    DizerBarra "verificando sequencia" & StrZero(nTMPSEQ, 2) & "/" & StrZero(nTMPSSQ, 2) & "/" & StrZero(nTMPITEM, 2)
    Select Case Index
    Case 0
      cSQL = "SELECT ITEM FROM FEMEA WHERE PF=" & nPF & " AND FXSEQ=" & nTMPSEQ & " AND FXSSQ=" & nTMPSSQ & " AND FXITEM=" & nTMPITEM
      nITEM = PegUltSQL(cARQFEMEAPF, cSQL, "ITEM", 0)
      If nITEM = 0 Then
        Command1_Click 4
      End If
    Case 1
      cSQL = "SELECT ITEM FROM FEMAVU WHERE PF=" & nPF & " AND FXSEQ=" & nTMPSEQ & " AND FXSSQ=" & nTMPSSQ & " AND FXITEM=" & nTMPITEM
      nITEM = PegUltSQL(cARQFEMEA, cSQL, "ITEM", 0)
      If nITEM = 0 Then
        Command1_Click 5
      End If
    Case 2
      cSQL = "SELECT ITEM FROM FEMEA WHERE PF=" & nPF & " AND FXSEQ=" & nTMPSEQ & " AND FXSSQ=" & nTMPSSQ & " AND FXITEM=" & nTMPITEM
      nITEM = PegUltSQL(cARQPFP, cSQL, "ITEM", 0)
      If nITEM = 0 Then
        Command1_Click 10
      End If
    Case 3
      cSQL = "SELECT ITEM FROM FEMEA WHERE PF=" & nPF & " AND FXSEQ=" & nTMPSEQ & " AND FXSSQ=" & nTMPSSQ & " AND FXITEM=" & nTMPITEM
      nITEM = PegUltSQL(cARQPFG, cSQL, "ITEM", 0)
      If nITEM = 0 Then
        Command1_Click 11
      End If
    End Select
  Next x
  If Index = 0 Or Index = 1 Or Index = 2 Or Index = 3 Then
    FilRelat (Index)
  End If
  DizerBarra ""
End Sub

Private Sub CmdcheckOp_Click(Index As Integer)
  Dim nROWS As Integer
  Dim x As Integer
  Dim nITEM As Integer
  Dim cSQL As String
  Dim nTMPSEQ As Integer
  Dim nTMPSSQ As Integer
  nROWS = gridopr.Rows - 1
  For x = 1 To nROWS                           ' zero e o titulo do grid comeca 1
    nITEM = 0
    gridopr.Col = 1
    nTMPSEQ = FixInt(gridopr)
    gridopr.Col = 2
    nTMPSSQ = FixInt(gridopr)
    If Index = 0 Then
      cSQL = "SELECT ITEM FROM FEMEA WHERE PF=" & nPF & " AND FXSEQ=" & nTMPSEQ & " AND FXSSQ=" & nTMPSSQ
      nITEM = PegUltSQL(cARQFEMEAPF, cSQL, "ITEM", 0)
    End If
    If Index = 1 Then
      cSQL = "SELECT ITEM FROM FEMAVU WHERE PF=" & nPF & " AND FXSEQ=" & nTMPSEQ & " AND FXSSQ=" & nTMPSSQ
      nITEM = PegUltSQL(cARQFEMEA, cSQL, "ITEM", 0)
    End If
    If nITEM = 0 Then
      If Index = 0 Then
        Command1_Click 2
      Else
        Command1_Click 3
      End If
    End If
  Next x
End Sub

Private Sub cmdClose_Click()
  Dim dFEMEA As Date
  Dim dPF As Date
  On Error Resume Next
  If IsDate(TXTFIELDS(23)) And IsDate(TXTFIELDS(17)) Then
    dFEMEA = TXTFIELDS(23)
    If IsDate(TXTFIELDS(25)) Then
      dFEMEA = TXTFIELDS(25)
    End If
    dPF = TXTFIELDS(17)
    If dFEMEA < dPF Then
      Alert ("Data Revisao Femea Menor que Processo Fabricacao")
    End If
  End If
  chkfemeas

  For iLOOP = 0 To nCAMPOS - 8
    aVAL(iLOOP) = TXTFIELDS(iLOOP)
  Next iLOOP
  aVAL(27) = FixNumBol(Check1.Value)
  aVAL(28) = FixNumBol(Check2.Value)
  aVAL(29) = FixNumBol(femeacrg.Value)
  aVAL(30) = DTPicker2.Value
  aVAL(31) = DTPicker1.Value

  For iLOOP = 32 To nCAMPOS - 1
    aVAL(iLOOP) = TXTFIELDS(iLOOP)
  Next iLOOP


  GrvSQL cARQPF, cSQL, nCAMPOS, aCAM, aVAL, aFOR
  ATUPFPG

  Screen.MousePointer = vbDefault
  Unload Me

End Sub

Private Sub CmdClo_Click(Index As Integer)
  Dim dDate As Date
  Dim DB As New ADODB.Connection
  Dim RSTAB As New ADODB.Recordset
  Dim sSQL As String
  dDate = Date

  If Index = 0 And Not AcessaMenu("FemeaPro", 1) Then Exit Sub
  If Index = 1 And Not AcessaMenu("FemeaPro", 1) Then Exit Sub

  If Index = 6 And Not AcessaMenu("FemeaGP12", 1) Then Exit Sub
  If Index = 7 And Not AcessaMenu("FemeaGP12", 1) Then Exit Sub

  If Index = 4 Or Index = 5 Then
    If Not AcessaMenu("FemeaPre", 1) Then
      Exit Sub
    End If
    If eproducao Then
      Exit Sub
    End If
  End If

  If GridSeq(Index).Row = 0 Then
    Exit Sub
  End If

  DB.ConnectionTimeout = 120
  If Index = 0 Or Index = 1 Then
    DB.Open GeracArq(cARQFEMEAPF)            'CARQPF
  End If
  If Index = 2 Or Index = 3 Or Index = 8 Or Index = 9 Then
    DB.Open GeracArq(cARQFEMEA)
  End If
  If Index = 4 Or Index = 5 Then
    DB.Open GeracArq(cARQPFP)
  End If
  If Index = 6 Or Index = 7 Then
    DB.Open GeracArq(cARQPFG)
  End If

  Select Case Index
  Case 1, 5, 7
    sSQL = "select * from FEMEA WHERE PF=" & nPF & " ORDER BY ITEM"
  Case 3
    sSQL = "select * from FEMAVU WHERE PF=" & nPF & " ORDER BY ITEM"
  Case 9
    sSQL = "select * from FEMADC WHERE PF=" & nPF & " ORDER BY ITEM"
  Case 0, 4, 6
    GridSeq(0).Col = 1
    nSEQ = GridSeq(0)
    sSQL = "select * from FEMEA WHERE PF=" & nPF & " AND ITEM=" & nSEQ
  Case 2
    GridSeq(1).Col = 1
    nSEQ = GridSeq(1)
    sSQL = "select * from FEMAVU WHERE PF=" & nPF & " AND ITEM=" & nSEQ
  Case 8
    GridSeq(5).Col = 1
    nSEQ = GridSeq(5)
    sSQL = "select * from FEMADC WHERE PF=" & nPF & " AND ITEM=" & nSEQ
  End Select



  RSTAB.Open sSQL, DB, adOpenDynamic, adLockOptimistic
  While Not RSTAB.EOF
    If RSTAB("ACAO") Then
      If IsDate(RSTAB("RESDAT")) Then
        dDate = RSTAB("RESDAT")
      End If
      If IsDate(RSTAB("RESDAT2")) Then
        If dDate < RSTAB("RESDAT2") Then
          dDate = RSTAB("RESDAT2")
        End If
      End If
      If IsDate(RSTAB("RESDAT3")) Then
        If dDate < RSTAB("RESDAT3") Then
          dDate = RSTAB("RESDAT3")
        End If
      End If
      RSTAB("ACATOM") = "Concluida"
      RSTAB("ACADAT") = dDate + 2
      RSTAB.Update
    End If
    RSTAB.MoveNext
  Wend
  RSTAB.Close
  DB.Close

End Sub

Private Sub CmdDelDef_Click()
  Dim cARQTMP As String
  Dim sSQL As String
  cARQTMP = PegPath("PATH", "MANA5TGQ")
  cARQTMP = GeraConn(cARQTMP, "JETFOX")
  gridmsrd.Col = 0
  sSQL = "select NUMERO from MSRD WHERE MSRD=" & gridmsrd
  If ApagaSQLP(cARQTMP, sSQL) Then
    filgridmsrd
  End If

End Sub

Private Sub CmddelPre_Click()
  Dim nSEQ As Integer
  Dim cSQL As String
  If Gridfemprecad.Row > 0 Then
    Gridfemprecad.Col = 0
    nSEQ = Gridfemprecad
    cSQL = "select * from FEMPRE WHERE SEGGRA=" & nSEQ
    If ApagaSQLP(cARQFEMEA, cSQL) Then
      FILGRIdfemprecad
    End If
  End If
End Sub

Private Sub CmdEdiDef_Click()
  gridmsrd.Col = 0
  ePASS01 = gridmsrd
  FrmMsrd.Show vbModal, Me
End Sub

Private Sub CmdEdiPad_Click(Index As Integer)

  Dim aVAL As Variant
  Dim nNUM As Integer
  Dim nROW As Integer
  Dim cSIG As String
  Dim cDIZ As String

  If Index = 6 Or Index = 7 Or Index = 8 Then
    cSIG = txtsigi.Text
    If Len(cSIG) = 0 Then
      Alert ("Siginificativa Nao Escolhida")
      Exit Sub
    End If
  End If
  aVAL = Array(0, 0, 0, 0)
  If Index <> 10 Then
    GridPad.Col = 2
    aVAL(0) = GridPad
    GridPad.Col = 3
    aVAL(1) = GridPad
    GridPad.Col = 4
    aVAL(2) = GridPad
    GridPad.Col = 5
    aVAL(3) = GridPad
    ePASS01 = "select * from FEMEA WHERE PRONUM=" & FixNum(aVAL(0)) & " AND FALNUM=" & aVAL(1) & " AND EFENUM=" & aVAL(2) & " AND CAUNUM=" & aVAL(3) & " AND PF=" & nPF
  End If
  ePASS02 = cARQFEMEA
  nROW = GridPad.Row
  Select Case Index
  Case 0
    ePASS01 = "select * from FEMEAPAD WHERE PRONUM=" & FixNum(aVAL(0)) & " AND FALNUM=" & aVAL(1) & " AND EFENUM=" & aVAL(2) & " AND CAUNUM=" & aVAL(3) & " AND PF=" & nPF
    frmFEMEI.Show vbModal, Me
  Case 1
    If ApagaSQLP(cARQFEMEA, ePASS01) Then
      ePASS01 = "select * from FEMRPNT WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1) & " AND EFENUM=" & aVAL(2) & " AND CAUNUM=" & aVAL(3) & " AND PF=" & nPF
      ApagaSQL cARQFEMEA, ePASS01
    End If
  Case 2
    If MDG("Apagar Todos os Efeitos " & CStr(aVAL(2))) Then
      'ePASS01 = "select * from FEMEA WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1) & " AND EFENUM=" & aVAL(2) & " AND PF=" & nPF
      ePASS01 = "select * from FEMEAPAD WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1) & " AND EFENUM=" & aVAL(2) & " AND PF=" & nPF
      If ApagaSQLP(cARQFEMEA, ePASS01) Then
        ePASS01 = "select * from FEMRPNT WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1) & " AND EFENUM=" & aVAL(2) & " AND PF=" & nPF
        ApagaSQL cARQFEMEA, ePASS01
      End If
    End If
  Case 3
    nNUM = 0
    nNUM = FixInt(Val(Busca("Nº Item", "Atribuindo", CStr(nNUM), 8)))
    If nNUM = 0 Then
      Alert ("Nº ITEM em Branco")
      Exit Sub
    End If
    If MDG("Atribuir Item " & CStr(nNUM) & "Para Processo " & CStr(aVAL(0))) Then
      '                ePASS01 = "select * from FEMEA WHERE PRONUM=" & aVAL(0) & " AND PF=" & nPF
      ePASS01 = "select * from FEMEAPAD WHERE PRONUM=" & aVAL(0) & " AND PF=" & nPF
      GrvSQL cARQFEMEA, ePASS01, 1, Array("ITEM"), Array(nNUM), Array("NI")
    End If
  Case 4
    If MDG("Apagar Todos as Falhas " & CStr(aVAL(1))) Then
      '            ePASS01 = "select * from FEMEA WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1) & " AND PF=" & nPF
      ePASS01 = "select * from FEMEAPAD WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1) & " AND PF=" & nPF
      If ApagaSQLP(cARQFEMEA, ePASS01) Then
        ePASS01 = "select * from FEMPF WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1) & " AND PF=" & nPF
        ApagaSQL cARQFEMEA, ePASS01
        ePASS01 = "select * from FEMRPNT WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1) & " AND PF=" & nPF
        ApagaSQL cARQFEMEA, ePASS01
        filgridpadpf
      End If
    End If
  Case 5
    If MDG("Apagar Todos os Processos " & CStr(aVAL(0))) Then
      '                ePASS01 = "select * from FEMEA WHERE PRONUM=" & FixInt(aVAL(0)) & " AND PF=" & nPF
      ePASS01 = "select * from FEMEAPAD WHERE PRONUM=" & FixInt(aVAL(0)) & " AND PF=" & nPF
      If ApagaSQLP(cARQFEMEA, ePASS01) Then
        ePASS01 = "select * from FEMPF WHERE PRONUM=" & FixInt(aVAL(0)) & " AND PF=" & nPF
        ApagaSQL cARQFEMEA, ePASS01
        ePASS01 = "select * from FEMRPNO WHERE PRONUM=" & FixInt(aVAL(0)) & " AND PF=" & nPF
        ApagaSQL cARQFEMEA, ePASS01
        ePASS01 = "select * from FEMRPNT WHERE PRONUM=" & FixInt(aVAL(0)) & " AND PF=" & nPF
        ApagaSQL cARQFEMEA, ePASS01
        filgridpadpf
      End If
    End If
  Case 6
    If MDG("Atribuir Todos os Efeitos " & CStr(aVAL(2))) Then
      '                ePASS01 = "select * from FEMEA WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1) & " AND EFENUM=" & aVAL(2) & " AND PF=" & nPF
      ePASS01 = "select * from FEMEAPAD WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1) & " AND EFENUM=" & aVAL(2) & " AND PF=" & nPF
      GrvSQL cARQFEMEA, ePASS01, 1, Array("SIGI"), Array(cSIG), Array("C")
    End If
  Case 7
    If MDG("Atribuir Todos as Falhas " & CStr(aVAL(1))) Then
      '                ePASS01 = "select * from FEMEA WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1) & " AND PF=" & nPF
      ePASS01 = "select * from FEMEAPAD WHERE PRONUM=" & aVAL(0) & " AND FALNUM=" & aVAL(1) & " AND PF=" & nPF
      GrvSQL cARQFEMEA, ePASS01, 1, Array("SIGI"), Array(cSIG), Array("C")
    End If
  Case 8
    If MDG("Atribuir Todos os Processos " & CStr(aVAL(0))) Then
      '                ePASS01 = "select * from FEMEA WHERE PRONUM=" & aVAL(0) & " AND PF=" & nPF
      ePASS01 = "select * from FEMEAPAD WHERE PRONUM=" & aVAL(0) & " AND PF=" & nPF
      GrvSQL cARQFEMEA, ePASS01, 1, Array("SIGI"), Array(cSIG), Array("C")
    End If
  Case 9
    cDIZ = pegdizseqssq("", cARQPF)
    If Len(cDIZ) = 0 Then
      Alert ("Descricao Nao Especificada")
      Exit Sub
    End If
    If MDG("Renomear Processos " & CStr(aVAL(0))) Then
      '                ePASS01 = "select * from FEMEA WHERE PRONUM=" & aVAL(0) & " AND PF=" & nPF
      ePASS01 = "select * from FEMEAPAD WHERE PRONUM=" & aVAL(0) & " AND PF=" & nPF
      GrvSQL cARQFEMEA, ePASS01, 1, Array("PROCESSO"), Array(cDIZ), Array("C")
    End If
  Case 10
    If MDG("Apagar Todos Femea Padrao ") Then
      ePASS01 = "select * from FEMEAPAD WHERE PF=" & nPF
      '                ePASS01 = "select * from FEMEA WHERE PF=" & nPF
      If ApagaSQLP(cARQFEMEA, ePASS01) Then
        ePASS01 = "select * from FEMPF WHERE PF=" & nPF
        ApagaSQL cARQFEMEA, ePASS01
        ePASS01 = "select * from FEMRPNO WHERE PF=" & nPF
        ApagaSQL cARQFEMEA, ePASS01
        ePASS01 = "select * from FEMRPNT WHERE PF=" & nPF
        ApagaSQL cARQFEMEA, ePASS01
        filgridpadpf
      End If
    End If

  End Select
  FILGRIdPAD
  If nROW > 0 And nROW < GridPad.Rows - 1 Then
    GridPad.Row = nROW
    GridPad.TopRow = nROW
  End If

End Sub

Private Sub cmdedipre_Click(Index As Integer)

  Dim nSEQ As Integer
  If Gridfemprecad.Row > 0 Then
    Gridfemprecad.Col = 0
    nSEQ = Gridfemprecad
    ePASS01 = "select * from FEMPRE WHERE SEGGRA=" & nSEQ
    ePASS02 = cARQFEMEA
    frmFEMEI.Show vbModal, Me
    FILGRIdfemprecad
  End If
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

Private Sub CmdEscFal_Click()

  Dim nNUM As Integer
  Dim cDIZ As String
  GRIDFAL.Col = 1

  nNUM = FixInt(GRIDFAL)
  aNUM(1) = nNUM
  GRIDFAL.Col = 2
  cDIZ = FixStr(GRIDFAL)
  aDIZ(1) = cDIZ
  dizfal.Caption = CStr(nNUM) & " - " & cDIZ
  LocalizaGrid GRIDFAL, aNUM(1), 2, False, 0

End Sub

Private Sub CmdEscPro_Click()

  Dim nNUM As Integer
  Dim cDIZ As String
  GridPRO.Col = 0

  nNUM = FixInt(GridPRO)
  aNUM(0) = nNUM
  aNUM(1) = 0
  GridPRO.Col = 1
  cDIZ = FixStr(GridPRO)
  aDIZ(0) = cDIZ
  FILGRIdFAL
  dizpro.Caption = CStr(nNUM) & " - " & cDIZ
  LocalizaGrid GridPRO, aNUM(0), 1, False, 0

End Sub

Private Sub CmdExcPFPad_Click(Index As Integer)

  Dim aVAL As Variant
  Dim cSQL As String
  Dim lAPAGA As Boolean





  aVAL = Array(0, 0, 0)

  GridPadPF.Col = 1
  aVAL(1) = FixInt(GridPadPF)
  GridPadPF.Col = 2
  aVAL(2) = FixInt(GridPadPF)
  If Index = 1 Then                            ''Processo
    cSQL = "select * from FEMPF WHERE PRONUM=" & aVAL(1) & " AND PF=" & nPF
    If ApagaSQLP(cARQFEMEA, cSQL) Then
      cSQL = "select * from FEMEAPAD WHERE PRONUM=" & aVAL(1) & " AND PF=" & nPF  'FEMEAPAD
      ApagaSQL cARQFEMEA, cSQL
      cSQL = "select * from FEMRPNO WHERE PRONUM=" & aVAL(1) & " AND PF=" & nPF
      ApagaSQL cARQFEMEA, cSQL
      cSQL = "select * from FEMRPNT WHERE PRONUM=" & aVAL(1) & " And PF = " & nPF
      ApagaSQL cARQFEMEA, cSQL
    End If
  End If
  If Index = 0 Or Index = 2 Then               ''Efeito
    lAPAGA = False
    cSQL = "select * from FEMPF WHERE PRONUM=" & aVAL(1) & " AND FALNUM=" & aVAL(2) & " AND PF=" & nPF
    DizerBarra "Excluindo Femea Padrao-PF"
    If Index = 0 Then
      lAPAGA = ApagaSQLP(cARQFEMEA, cSQL)
    End If
    If Index = 2 Then
      ApagaSQL cARQFEMEA, cSQL
      lAPAGA = True
    End If
    If lAPAGA Then
      DizerBarra "Excluindo Femea Padrao_Itens"
      cSQL = "select * from FEMEAPAD WHERE PRONUM=" & aVAL(1) & " AND FALNUM=" & aVAL(2) & " AND PF=" & nPF
      ApagaSQL cARQFEMEA, cSQL
      DizerBarra "Excluindo Femea Padrao Apuracao RPN"
      cSQL = "select * from FEMRPNT WHERE PRONUM=" & aVAL(1) & " AND FALNUM=" & aVAL(2) & " AND PF=" & nPF
      ApagaSQL cARQFEMEA, cSQL
      If Index = 2 Then
        DizerBarra "Gravando Padrao-PF Historico"
        cSQL = "select * from FEMPFHIS WHERE PRONUM=" & aVAL(1) & " AND FALNUM=" & aVAL(2) & " AND PF=" & nPF
        IncluiSQL cARQFEMEA, cSQL, 3, Array("PF", "PRONUM", "FALNUM"), Array(nPF, aVAL(1), aVAL(2)), True, False
        Exit Sub
      End If
    End If
    DizerBarra ""
  End If
  filgridpadpf
  FILGRIdPAD

End Sub

Private Sub cmdimp_Click()
  Xcmdimp_Click 0
End Sub

Private Sub Xcmdimp_Click(nPFTMP)
  Dim cSQL As String
  Dim oDB As New ADODB.Connection
  Dim oRS As New ADODB.Recordset
  Dim oRSDES As New ADODB.Recordset
  Dim nFALTMP As Integer
  Dim nPROTMP As Integer
  Dim lGRAVA As Boolean
  lGRAVA = False
  On Error Resume Next
  If nPFTMP = 0 Then
    nPFTMP = FixInt(Val(Busca("Nº PF", "Importando", CStr(nPFTMP), 8)))
    If nPFTMP = 0 Then
      Alert ("Nº PF Nao Especificado")
      Exit Sub
    End If
    lGRAVA = True
  End If
  cSQL = "SELECT * FROM FEMPF WHERE PF=" & nPFTMP & " ORDER BY PRONUM,FALNUM"
  oDB.ConnectionTimeout = 120
  oDB.Open GeracArq(cARQFEMEA)
  oRS.Open cSQL, oDB, adOpenForwardOnly, adLockReadOnly
  Do While Not oRS.EOF
    nPROTMP = FixInt(oRS("PRONUM"))
    nFALTMP = FixInt(oRS("FALNUM"))
    If lGRAVA Then
      GERAFEMEAPAD nPROTMP, nFALTMP, FixInt(nPF)  ''pf e atual pftmp para sql
    Else
      cSQL = "select * from FEMPF WHERE PRONUM=" & oRS("PRONUM") & " AND FALNUM=" & oRS("FALNUM") & " AND PF=" & nPF
      oRSDES.Open cSQL, oDB, adOpenDynamic, adLockOptimistic
      If oRSDES.EOF Then
        oRSDES.AddNew
        oRSDES("PRONUM") = oRS("PRONUM")
        oRSDES("FALNUM") = oRS("FALNUM")
        oRSDES("PF") = nPF
        oRSDES.Update
      End If
      oRSDES.Close
    End If
    oRS.MoveNext
  Loop
  oRS.Close
  oDB.Close
  filgridpadpf
  FILGRIdPAD
End Sub

Private Sub CmdIncDef_Click()
  Dim nSAC As Long
  Dim cCAM As String
  Dim cTITULO As String
  Dim aCAMPOS As Variant
  Dim avalor As Variant

  cTITULO = FixStr(Busca("Digite Descricao", "Descricao", Space(50), 50), "")
  If Len(cTITULO) = 0 Then
    Alert ("Necessario Digitar Descricao")
    Exit Sub
  End If

  cCAM = PegPath("PATH", "MANA5TGQ")

  nSAC = PegMAXSQL(GeraConn(cCAM, "FOX"), "MSRD", "MSDR", 1)
  nSAC = nSAC + 1

  aCAMPOS = Array("MSRD", "DATA", "PF", "CODIGO", "CLIENTE", "DESC01")
  avalor = Array(nSAC, Format(Date, "DD/MM/YY"), FixNum(TXTFIELDS(0)), FixStr(TXTFIELDS(2)), FixNum(TXTFIELDS(4)), FixStr(cTITULO))

  IncluiSQL GeraConn(cCAM, "FOX"), "SELECT * FROM MSRD WHERE MSRD=" & nSAC, 6 _
          , aCAMPOS _
          , avalor, _
          , False, True


  filgridmsrd
End Sub

Private Sub CmdIncPf_Click()
  Dim cDIZ As String
  On Error Resume Next
  If aNUM(0) = 0 Then
    Alert ("Processo Nao Escolhido")
    Exit Sub
  End If
  If aNUM(1) = 0 Then
    Alert ("Falha Nao Escolhido")
    Exit Sub
  End If
  cDIZ = pegdizseqssq(dizpro, cARQPF)
  If Len(cDIZ) = 0 Then
    Alert ("Descricao Nao Especificada")
    Exit Sub
  End If
  If Not MDG(cDIZ, "Incluir") Then
    Exit Sub
  End If
  GERAFEMEAPAD aNUM(0), aNUM(1), FixInt(nPF), cDIZ
  filgridpadpf
  FILGRIdPAD

End Sub

Private Sub CmdLiberar_Click()
  Dim sSQL As String
  sSQL = "SELECT LIGADO FROM USUCAD WHERE FORM = 'escFEMEA' AND INDICE=3 AND IDUSUARIO=" & zUSERID
  'eRETU01 = PegUltSQL(Dbname, sSQL, "LIGADO", False)
  eRETU01 = PegCampoSQLADO(Dbname, sSQL, "", False)
  If eRETU01 Then
    TXTFIELDS(13) = zIDFOLHA
    TXTFIELDS(14) = zNOMEFOLHA
    TXTFIELDS(24) = Today()
    Encerrar.Visible = False
  Else
    Alert ("Sem direito")
  End If
End Sub

Private Sub CmdMarcaFemea_Click()
  Dim cSQLTMP As String
  If GridPFrev.Row = 0 Then
    Exit Sub
  End If
  GridPFrev.Col = 3
  If Len(GridPFrev) > 0 Then
    Alert ("Ja Verificado")
    Exit Sub
  End If
  GridPFrev.Col = 1
  cSQLTMP = "select * from PFrev WHERE PF=" & nPF & " AND REV=" & FixInt(GridPFrev)
  GrvSQL cARQPF, cSQLTMP, 2, Array("FEMEAD", "FEMEAN"), Array(Date, zNOMEFOLHA), Array("D", "C")
  FILGRIdpfrev
End Sub

Private Sub CmdMotRev_Click()
  If gridrev.Row = 0 Then
    Exit Sub
  End If

  gridrev.Col = 1
  nREV = gridrev
  ComMotAlt(0).Enabled = True
  ComMotExc.Enabled = True
  ComMotInc.Enabled = True
  Filgridrevi (0)

End Sub

Private Sub CmdNome_Click()
  Dim cDIZ As String
  Dim cSQL As String
  On Error Resume Next

  If aNUM(0) = 0 Then
    Alert ("Processo Nao Escolhido")
    Exit Sub
  End If
  cDIZ = pegdizseqssq(dizpro, cARQPF)
  If Len(cDIZ) = 0 Then
    Alert ("Descricao Nao Especificada")
    Exit Sub
  End If
  If Not MDG(cDIZ, "Trocar Descricao Processo") Then
    Exit Sub
  End If
  cSQL = "select PROCESSO from FEMeaPAD WHERE PRONUM=" & FixInt(aNUM(0)) & " AND PF=" & FixInt(nPF)
  GrvSQL cARQFEMEA, cSQL, 1, Array("PROCESSO"), Array(cDIZ), Array("C")

End Sub

Private Sub CmdNovaREV_Click()
  Dim dDATAREV As Date
  Encerrar.Visible = False
  dDATAREV = Date

  FilRelat 0
  If Not lRETU Then
    Alert ("Corriga antes de fazer a revisao")
    Exit Sub
  End If


  If Len(TXTFIELDS(23).Text) > 0 Then          'corrige erro importacao protheus data em branco
    If TXTFIELDS(23) > Today() Then
      Alert ("Data Revisao Anterior Maior que Data do Sistema")
      Exit Sub
    End If
  End If

  If MDG("gerar revisao com historico") Then
    If MDG("gerar Utilizando novos indicadores") Then
      CmdREvisarTodas_Click 0              ''CmdRevFemea_Click 16
    Else
      If MDG("gerar nova revisando mantendo os indicadores") Then
        CmdREvisarTodas_Click 16         ''CmdRevFemea_Click 16
      End If
    End If
  End If

  If IsNumeric(TXTFIELDS(12)) Then
    TXTFIELDS(12) = TXTFIELDS(12) + 1
  Else
    TXTFIELDS(12) = 1
  End If
  TXTFIELDS(20) = True
  TXTFIELDS(13) = 0
  TXTFIELDS(14) = ""
  TXTFIELDS(24) = ""
  TXTFIELDS(23) = dDATAREV
  Check1.Value = False
  nREV = FixInt(TXTFIELDS(12))
  IncluiSQL cARQPF, "select * from rev WHERE PF=" & nPF & " AND TIPO='FE' AND REVISAO=" & nREV, 4, Array("PF", "REVISAO", "TIPO", "DATA"), Array(FixInt(nPF), nREV, "FE", dDATAREV), True, False
  CmdProx_Click
  Filgridrev
  Filgridrevi (0)

  Command3_Click



  eRETU02 = "PF:" & TXTPF.Text & Chr(13) & Chr(10)
  eRETU02 = eRETU02 & " Desenho:" & TXTFIELDS(2) & Chr(13) & Chr(10)
  eRETU02 = eRETU02 & " Data:" & Fdata(dDATAREV) & Chr(13) & Chr(10)
  MAILENV "FEM00001", eRETU02

  cmdClose_Click

End Sub

Private Sub CmdOrdFem_Click(Index As Integer)
  Select Case Index
  Case 0
    cORDFEMUSO = "ITEM"
    FilRelat (0)
  Case 1
    cORDFEMUSO = "FXSEQ,FXSSQ,FXITEM,FXITEMS"
    FilRelat (0)
  Case 2
    cORDFEMAVU = "ITEM"
    FilRelat (1)
  Case 3
    cORDFEMAVU = "FXSEQ,FXSSQ,FXITEM,FXITEMS"
    FilRelat (1)
  Case 4
    cORDFEMPRE = "FXSEQ,FXSSQ,FXITEM,FXITEMS"
    FilRelat (2)
  Case 5
    cORDFEMPRE = "ITEM"
    FilRelat (2)
  Case 6
    cORDFEMGP12 = "FXSEQ,FXSSQ,FXITEM,FXITEMS"
    FilRelat (3)
  Case 7
    cORDFEMGP12 = "ITEM"
    FilRelat (3)
  Case 9
    cORDFEMADC = "FXSEQ,FXSSQ,FXITEM,FXITEMS"
    FilRelat (5)
  Case 8
    cORDFEMADC = "ITEM"
    FilRelat (5)
  Case 10
    cORDREVI = "ITEM,SEGGRA"
    FILGRIDrevCAU (0)
  Case 11
    cORDREVI = "FXSEQ,FXSSQ,FXITEM,FXITEMS,FEMEAREV"
    FILGRIDrevCAU (0)
  End Select
End Sub

Private Sub CmdOrdPad_Click(Index As Integer)

  Select Case Index
  Case 0
    cORDFEMPAD = "PRONUM,FALNUM,EFENUM,CAUNUM"
  Case 1
    cORDFEMPAD = "ITEM,PRONUM,FALNUM,EFENUM,CAUNUM"
  Case 2
    cORDFEMPAD = "INDRIS,PRONUM,FALNUM,EFENUM,CAUNUM DESC"
  Case 3
    cORDFEMPAD = "INDRIS,PRONUM,FALNUM,EFENUM,CAUNUM"

  End Select
  FILGRIdPAD

End Sub

Private Sub CmdPreinc_Click()

  Dim cTITULO As String
  cTITULO = FixStr(Busca("Digite o Titulo", "Titulo Pre-Cadastro Femea", Space(50), 50), "")

  If Len(cTITULO) = 0 Then
    Alert ("Titulo Necessario")
    Exit Sub
  End If
  IncluiSQL cARQFEMEA, "select * from FEMPRE where titulo='" & cTITULO & "'", 1, Array("TITULO"), Array(cTITULO), True
  FILGRIdfemprecad

End Sub

Private Sub CmdProx_Click()

  If femeacrg.Value = False Then
    If IsDate(TXTFIELDS(23)) Then
      DTPicker1.Value = CDate(TXTFIELDS(23)) + 730
    Else
      If IsDate(DTPicker1) Then
        DTPicker1 = DTPicker1 + 730
      Else
        DTPicker1 = Today() + 730
      End If
    End If
  End If

End Sub

Private Sub CmdRevFemea_Click(Index As Integer)
  RevisaoFemea (Index)
End Sub

Private Sub RevisaoFemea(ByVal Index As Integer)
  Dim nREVI As Long
  Dim nSEQ As Long
  Dim cSQL As String
  Dim cARQ As String
  Dim aVAL As Variant
  Dim aCAM As Variant
  Dim aFOR As Variant
  Dim aPAD As Variant
  Dim lVALORES As Boolean
  Dim nCAMPOSUSO As Integer
  Dim nINDEXINI As Integer


  nINDEXINI = Index

  '0-producao 1-Padraocomplementar 2 pre 3 gp12 5 adicional 14-complementar->historico 15-complementar->pre 16-Producao sem checagem valor
  lVALORES = True
  If Index = 16 Then
    lVALORES = False
    Index = 0
  End If



  If Index = 0 Then
    If Not AcessaMenu("FemeaPro", 1) Then
      Exit Sub
    End If
  End If

  If Index = 3 Then
    If Not AcessaMenu("FemeaGP12", 1) Then
      Exit Sub
    End If
  End If

  If Index = 2 Then
    If Not AcessaMenu("FemeaPre", 1) Then
      Exit Sub
    End If
    If eproducao Then
      Exit Sub
    End If

  End If


  'Sequencia do femea
  If Index < 6 Then
    If GridSeq(Index).Row = 0 Then
      Exit Sub
    End If
    GridSeq(Index).Col = 1
    nSEQ = GridSeq(Index)
  End If
  If Index = 14 Or Index = 15 Then
    If GridSeq(1).Row = 0 Then
      Exit Sub
    End If
    GridSeq(1).Col = 1
    nSEQ = GridSeq(1)
  End If
  If Index = 0 Then
    cARQ = cARQFEMEAPF                       ',cARQPF
  End If
  If Index = 1 Or Index = 14 Or Index = 5 Then
    cARQ = cARQFEMEA
  End If
  If Index = 2 Or Index = 15 Then
    cARQ = cARQPFP
  End If
  If Index = 3 Then
    cARQ = cARQPFG
  End If


  If Index = 0 Or Index = 1 Or Index = 2 Or Index = 3 Or Index = 5 Then
    cSQL = "FEMrevi WHERE  TIPOAPU='" & tipoapu(Index) & "' AND PF=" & nPF & " AND ITEM=" & nSEQ
  End If

  ''Numero da Proxima revisao conforme pf e sequencia P-padrao


  If Index = 14 Then
    cSQL = "FEMHIS WHERE PF=" & nPF
  End If
  If Index = 15 Then
    cSQL = "FEMEA WHERE PF=" & nPF
  End If

  If Index = 0 Or Index = 1 Or Index = 2 Or Index = 3 Or Index = 5 Then
    nREVI = FixInt(PegMAXSQL(cARQFEMEA, cSQL, "SEGGRA", 0))
  End If
  If Index = 14 Then
    nREVI = FixInt(PegMAXSQL(cARQFEMEA, cSQL, "ITEM", 0))
  End If
  If Index = 15 Then
    nREVI = FixInt(PegMAXSQL(cARQPFP, cSQL, "ITEM", 0))
  End If
  nREVI = nREVI + 1


  nCAMPOSUSO = 53
  aCAM = Array("PF", "ITEM", "PROCESSO", "FALTIP", "FALEFE", _
               "FALCAU", "CRTATU", "INDOCO", "INDSEV", "INDDET", _
               "INDRIS", "ACAREC", "RESCOD", "RESNOM", "RESCOD2", _
               "ACATOM", "RESCOD3", "RINDOCO", "RINDSER", "RINDDET", _
               "RINDRIS", "RESNOM2", "RESNOM3", "SITUACAO", "SIGI", _
               "CARAPREV", "RESDAT", "RESDAT2", "RESDAT3", "ACADAT", _
               "ACAO", "EXCRPN", "ALTMAN", "MUDPAD", "PSA", _
               "FXSEQ", "FXSSQ", "FXITEM", "PRONUM", "EFENUM", _
               "FALNUM", "CAUNUM", "SEGGRA", "TIPOAPU", "PRONUM", _
               "FALNUM", "EFENUM", "CAUNUM", "SIG2", "SIG3", _
               "SUBTIPO", "FXITEMS", "FEMEAREV")

  aFOR = Array("NI", "NI", "C", "C", "C", _
               "C", "C", "NI", "NI", "NI", _
               "NI", "C", "NI", "C", "NI", _
               "C", "NI", "NI", "NI", "NI", _
               "NI", "C", "C", "C", "C", "C", _
               "DN", "DN", "DN", "DN", "BN", _
               "BN", "BN", "BN", "C", "NI", "NI", "NI", "NI", "NI", "NI", "NI", "NI", "C", _
               "NI", "NI", "NI", "NI", "C", "C", "C", "", "NI")

  aPAD = Array(0, 0, "", "", "", _
               "", "", 0, 0, 0, _
               0, "", 0, "", 0, _
               "", 0, 0, 0, 0, 0, _
               "", "", "", "", "", _
               "", "", "", "", False, _
               False, False, False, "", 0, 0, 0, 0, 0, 0, 0, 0, " ", _
               0, 0, 0, 0, "C", "C", "C", 0, 0)



  If Index = 0 Or Index = 2 Or Index = 3 Then
    cSQL = "select * from FEMEA WHERE PF=" & nPF & " AND ITEM=" & nSEQ
  End If
  If Index = 1 Or Index = 14 Or Index = 15 Then
    cSQL = "select * from FEMAVU WHERE PF=" & nPF & " AND ITEM=" & nSEQ
  End If
  If Index = 5 Then
    cSQL = "select * from FEMADC WHERE PF=" & nPF & " AND ITEM=" & nSEQ
  End If


  'carrega o femea atual
  If Index = 15 Then                           'o destino e diferente da origem no index 15 carq=cARQPFP mas pega do femavu(cARQFEMEA)
    aVAL = PegSQL(cARQFEMEA, cSQL, nCAMPOSUSO, aCAM, aFOR, aPAD)
  Else
    aVAL = PegSQL(cARQ, cSQL, nCAMPOSUSO, aCAM, aFOR, aPAD)
  End If

  If Index = 0 Or Index = 1 Or Index = 2 Or Index = 3 Or Index = 5 Then
    'checa se foi feita a revisao
    If lVALORES Then
      If aVAL(17) = 0 Or aVAL(18) = 0 Or aVAL(19) = 0 Or aVAL(20) = 0 Then
        Alert ("Novos Indices Nao Preenchidos")
        Exit Sub
      End If
    End If

    'atualiza revisao e tipo
    aVAL(42) = nREVI


    'If Index = 0 Or Index = 1 Or Index = 2 Or Index = 3 Or Index = 5 Then 'ja no if acima
    aVAL(43) = tipoapu(Index)
    'End If

    If FixNum(aVAL(52)) = 0 Then             'revisao do femea checar em branco pois historio pode ter varias versoes
      aVAL(52) = FixNum(TXTFIELDS(12).Text)
    End If


  End If
  If Index = 14 Or Index = 15 Then             'a revisao e a passada no command10
    aVAL(1) = nREVI
  End If


  If Index = 0 Or Index = 1 Or Index = 2 Or Index = 3 Or Index = 5 Then
    cSQL = "select * from FEMrevi WHERE  TIPOAPU='" & tipoapu(Index) & "' AND PF=" & nPF & " AND ITEM=" & nSEQ & " AND SEGGRA=" & nREVI
  End If

  If Index = 14 Then
    cSQL = "select * from FEMHIS WHERE PF=" & nPF & " AND ITEM=" & nREVI
  End If
  If Index = 15 Then
    cSQL = "select * from FEMEA WHERE PF=" & nPF & " AND ITEM=" & nREVI
  End If


  If Index < 6 Or Index = 14 Then
    IncluiSQL cARQFEMEA, cSQL, nCAMPOSUSO, aCAM, aVAL, True, False
  End If

  If Index = 15 Then
    IncluiSQL cARQPFP, cSQL, nCAMPOSUSO, aCAM, aVAL, True, False
  End If

  If Index = 14 Then
    cSQL = "select * from FEMAVU WHERE PF=" & nPF & " AND ITEM=" & nSEQ
    ApagaSQL cARQFEMEA, cSQL
    cSQL = "select * FROM FEMrevi WHERE  TIPOAPU='C' AND PF=" & nPF & " AND ITEM=" & nSEQ
    GrvSQL cARQFEMEA, cSQL, 2, Array("ITEM", "TIPOAPU"), Array(nREVI, "1"), Array("NI", "C")
    Exit Sub
  End If

  If Index = 15 Then
    cSQL = "select * from FEMAVU WHERE PF=" & nPF & " AND ITEM=" & nSEQ
    ApagaSQL cARQFEMEA, cSQL
    cSQL = "select * FROM FEMrevi WHERE  TIPOAPU='C' AND PF=" & nPF & " AND ITEM=" & nSEQ
    GrvSQL cARQFEMEA, cSQL, 2, Array("ITEM", "TIPOAPU"), Array(nREVI, "R"), Array("NI", "C")
    Exit Sub
  End If

  If Index = 0 Or Index = 2 Or Index = 3 Then
    cSQL = "select * from FEMEA  WHERE PF=" & nPF & " AND ITEM=" & nSEQ
  End If
  If Index = 1 Then
    cSQL = "select * from FEMAVU WHERE PF=" & nPF & " AND ITEM=" & nSEQ
  End If
  If Index = 5 Then
    cSQL = "select * from FEMADC WHERE PF=" & nPF & " AND ITEM=" & nSEQ
  End If

  If lVALORES Then                             'so zera se trocar os valores
    GrvSQL cARQ, cSQL, 17, Array("ACAREC", "RESCOD", "RESNOM", "RESCOD2", "ACATOM", _
                                 "RESCOD3", "RINDOCO", "RINDSER", "RINDDET", "RINDRIS", _
                                 "RESNOM2", "RESNOM3", "INDOCO", "INDSEV", "INDDET", _
                                 "INDRIS", "ACAO"), _
                                 Array("", 0, "", 0, "", 0, 0, 0, 0, 0, "", "", aVAL(17), aVAL(18), aVAL(19), aVAL(20), False), _
                                 Array("C", "NI", "C", "NI", "C", "NI", "NI", "NI", "NI", "NI", "C", "C", "NI", "NI", "NI", "NI", "")
  End If
  If Index > 9 Then
    Index = 1
  End If

  Index = nINDEXINI                            'Retorna o index inicial pois index e byref

  If nINDEXINI <> 16 Then
    FILGRIDrevCAU (Index)
  End If


End Sub

Private Sub CmdREvi_Click(Index As Integer)
  Dim nREVI As Long
  Dim nSEQ As Long

  If Index = 0 Or Index = 1 Or Index = 2 Or Index = 3 Or Index = 5 Then
    If gridrevcau(Index).Row = 0 Then
      Exit Sub
    End If
    gridrevcau(Index).Col = 0
    nSEQ = gridrevcau(Index)
    gridrevcau(Index).Col = 1
    nREVI = gridrevcau(Index)
    If Index = 0 Or Index = 1 Or Index = 2 Or Index = 3 Or Index = 5 Then
      ePASS01 = "select * from FEMrevi WHERE  TIPOAPU='" & tipoapu(Index) & "' AND PF=" & nPF & " AND ITEM=" & nSEQ & " AND SEGGRA=" & nREVI
    End If
  End If
  If Index = 4 Then
    If GridSeq(4).Row = 0 Then
      Exit Sub
    End If
    GridSeq(4).Col = 1
    nSEQ = GridSeq(4)
    ePASS01 = "select * from FEMHIS WHERE PF=" & nPF & " AND ITEM=" & nSEQ
  End If
  ePASS02 = cARQFEMEA

  Load frmFEMEI
  frmFEMEI.TXTFIELDS(2).Enabled = False
  frmFEMEI.TXTFIELDS(3).Enabled = False
  frmFEMEI.TXTFIELDS(4).Enabled = False
  frmFEMEI.TXTFIELDS(5).Enabled = False
  frmFEMEI.cmdClose.Enabled = False
  frmFEMEI.CmdAltman.Enabled = False
  frmFEMEI.Show vbModal, Me



End Sub

Private Sub CmdREvisarTodas_Click(Index As Integer)
  Dim nROWS As Integer
  Dim x As Integer
  Dim nIndex As Integer
  If GridSeq(0).Rows = 0 Then
    Exit Sub
  End If
  If Index = 0 Then
    If Not MDG("Transfere indices e abre uma nova revisao todos") Then
      Exit Sub
    End If
  End If
  If Index = 16 Then
    If Not MDG("Nao Transfere indices e abre uma nova revisao todos ") Then
      Exit Sub
    End If
  End If

  nROWS = GridSeq(0).Rows
  nROWS = nROWS - 1
  For x = 1 To nROWS
    GridSeq(0).Row = x
    DizerBarra "Gerando Revisao femea " & x & "/" & nROWS
    RevisaoFemea nIndex
  Next x
  DizerBarra ""
  FilRelat 1
  If nIndex = 16 Then
    FILGRIDrevCAU 0
  End If
End Sub

Private Sub CMDSIG_Click(Index As Integer)

  lRETU = False
  eRETU01 = ""
  eRETU02 = ""
  If Index = 0 Then
    escFLX.Show vbModal, Me
  Else
    ePASS01 = "isoqsymbol"
    ePASS02 = True
    frmCharacters.Show vbModal, Me
  End If
  If lRETU Then
    txtsigi = eRETU01
  End If

End Sub

Private Sub CmdTotFlux_Click(Index As Integer)
  Dim nROWS As Integer
  Dim x As Integer
  If Index = 1 Then
    If GridPadPF.Rows < 2 Then
      Alert "Este Femea nao utiliza Padrao"
      Exit Sub
    End If
  End If
  If Index = 0 Then                            'normal com complementar
    If GridSeq(1).Rows > 1 Then
      Alert "Femea com complemetar nao pode usar normal"
      Exit Sub
    End If
  End If

  If Index = 0 Then                            'normal com padrao
    If GridPadPF.Rows > 1 Then
      Alert "Femea com Padrao nao pode usar normal"
      Exit Sub
    End If
  End If

  nROWS = GrdQsbLep.Rows - 1
  For x = 1 To nROWS                           ' zero e o titulo do grid comeca 1
    GrdQsbLep.Row = x
    Select Case Index
    Case 0
      Command1_Click 4
    Case 1
      Command1_Click 5
    Case 2
      Command1_Click 8
    Case 3
      Command1_Click 9
    End Select
  Next x

  Select Case Index
  Case 0                                       ' Command1_Click 4
    FilRelat (0)
  Case 1                                       'Command1_Click 5
    FilRelat (1)
  Case 2                                       'Command1_Click 8
    FilRelat (2)
  Case 3                                       'Command1_Click 9
    FilRelat (3)
  End Select



End Sub

Private Sub CmdTotPad_Click(Index As Integer)
  Dim nROWS As Integer
  Dim x As Integer
  nROWS = gridopr.Rows - 1
  For x = 1 To nROWS                           ' zero e o titulo do grid comeca 1
    gridopr.Row = x
    If Index = 0 Then
      Command1_Click 2
    Else
      Command1_Click 3
    End If
  Next x
End Sub

Private Sub CmdTrocaREv_Click()
  Dim nORI As Integer
  Dim nDES As Integer
  Dim cSQL As String
  nORI = FixInt(TxtRevisao, 0)
  nDES = FixInt(TxtRevNova, 0)
  cSQL = "select * from rev WHERE PF=" & nPF & " AND REVISAO=" & nORI & " AND TIPO='FE'"
  GrvSQL cARQPF, cSQL, 1, Array("REVISAO"), Array(nDES), Array("NI")
  cSQL = "select * from revI WHERE PF=" & nPF & " AND REVISAO=" & nORI & " AND TIPO='FE'"
  GrvSQL cARQPF, cSQL, 1, Array("REVISAO"), Array(nDES), Array("NI")
  TXTFIELDS(12).Enabled = True
  TXTFIELDS(23).Enabled = True
  TXTFIELDS(12).Locked = False
  TXTFIELDS(23).Locked = False
  Filgridrev
End Sub

Private Sub CmdVerApu_Click()
  Dim cSQL As String
  Dim nMES As Integer
  Dim nANO As Integer
  DizerBarra "carregando apuracoes"
  nMES = FixInt(Text(1))
  nANO = FixInt(Text(2))
  cSQL = "SELECT PF, ITEM, INDRIS, PRONUM, FALNUM, EFENUM, CAUNUM, ALTMAN, FALEFE, FALCAU FROM FEMRPNT WHERE ANO=" & nANO & " AND MES=" & nMES & " AND PF=" & nPF & " ORDER BY PF,INDRIS"
  MontaGridFast Grid10, 10, Array(400, 400, 400, 400, 400, 400, 400, 300, 2500, 2500), Array("PF", "ITEM", "RPN", "NºPro", "NºFal", "NºEfe", "NºCau", "Man", "Efeito", "Causa"), Array("PF", "ITEM", "INDRIS", "PRONUM", "FALNUM", "EFENUM", "CAUNUM", "ALTMAN", "FALEFE", "FALCAU"), cARQFEMEA, cSQL
  DizerBarra "carregando apuracoes RPN"
  cSQL = "SELECT PF, SEQ, PRONUM, TOTRPN, TOTCAU, TOTM40, MAIRPN, TOTFX01, TOTFX02, TOTFX03, TOTRPN01, TOTRPN02, TOTRPN03 FROM FEMRPNO WHERE ANO=" & nANO & " AND MES=" & nMES & " AND PF=" & nPF & " ORDER BY PF,PRONUM"
  MontaGridUltra GridOpe, 13, Array(500, 500, 500, 800, 500, 500, 500, 500, 500, 500, 800, 800, 800), Array("PF", "SEQ", "NºPro", "Tot RPN", "NºCau", ">40", "Maior", "0-40", "40-100", ">100", "Ant", "Ant", "Ant"), Array("PF", "SEQ", "PRONUM", "TOTRPN", "TOTCAU", "TOTM40", "MAIRPN", "TOTFX01", "TOTFX02", "TOTFX03", "TOTRPN01", "TOTRPN02", "TOTRPN03"), cARQFEMEA, cSQL
  dizcomapu.Caption = StrZero(nMES, 2) & "/" & StrZero(nANO, 4)
  DizerBarra ""
End Sub

Private Sub CmdVerDef_Click()
  filgridmsrd
End Sub

Private Sub Command1_Click(Index As Integer)
  Dim nNUM As Integer
  Dim cSQL As String
  Dim cDESC As String
  Dim nTMPSEQ As Integer
  Dim nTMPSSQ As Integer
  Dim nTMPITEM As Integer

  If Index = 7 Then
    If Not AcessaMenu("FemeaGP12", 1) Then
      Exit Sub
    End If
  End If
  If Index = 0 Then
    If Not AcessaMenu("FemeaPro", 1) Then
      Exit Sub
    End If
  End If

  If Index = 6 Then
    If Not AcessaMenu("FemeaPre", 1) Then
      Exit Sub
    End If
    If eproducao Then
      Exit Sub
    End If
  End If

  If Index = 8 Then
    If nREV = 0 Then
      Alert "Revisão Nao Escolhida"
      Exit Sub
    End If
    ePASS01 = "select * from rev WHERE PF=" & nPF & " AND TIPO='FE' AND REVISAO=" & nREV
    FrmRev.Show vbModal, Me
    Exit Sub
  End If

  If Index = 1 Or Index = 5 Then               'Complementar sem padrao
    If GridPadPF.Rows < 2 Then
      Alert "Este Femea nao utiliza Padrao"
      Exit Sub
    End If

  End If

  If Index = 0 Or Index = 4 Then               'normal com complementar
    If GridSeq(1).Rows > 1 Then
      Alert "Femea com complemetar nao pode usar normal"
      Exit Sub
    End If
  End If

  If Index = 0 Or Index = 4 Then               'normal com padrao
    If GridPadPF.Rows > 1 Then
      Alert "Femea com Padrao nao pode usar normal"
      Exit Sub
    End If
  End If


  If Index = 0 Or Index = 1 Or Index = 6 Or Index = 7 Or Index = 20 Then
    nTMPSEQ = 0
    nTMPSSQ = 0
    nTMPITEM = 0
    cDESC = ""
  End If
  If Index = 2 Or Index = 3 Then
    gridopr.Col = 1
    nTMPSEQ = FixInt(gridopr)
    gridopr.Col = 2
    nTMPSSQ = FixInt(gridopr)
    nTMPITEM = 0
    gridopr.Col = 4
    cDESC = FixStr(gridopr)
  End If

  If Index = 4 Or Index = 5 Or Index = 8 Or Index = 9 Or Index = 10 Or Index = 11 Or Index = 12 Or Index = 13 Then
    GrdQsbLep.Col = 0
    nTMPSEQ = FixInt(GrdQsbLep)
    GrdQsbLep.Col = 1
    nTMPSSQ = FixInt(GrdQsbLep)
    GrdQsbLep.Col = 2
    nTMPITEM = FixInt(GrdQsbLep)
    GrdQsbLep.Col = 3
    cDESC = FixStr(GrdQsbLep)
  End If


  Select Case Index
  Case 0, 2, 4
    nNUM = FixInt(PegMAXSQL(cARQFEMEAPF, "FEMEA WHERE PF=" & nPF, "ItEM", 0)) + 1  'CARQPF
  Case 1, 3, 5
    nNUM = FixInt(PegMAXSQL(cARQFEMEA, "FEMAVU WHERE PF=" & nPF, "ITEM", 0)) + 1
  Case 6, 8, 10, 12
    nNUM = FixInt(PegMAXSQL(cARQPFP, "FEMEA WHERE PF=" & nPF, "ItEM", 0)) + 1
  Case 7, 9, 11, 13
    nNUM = FixInt(PegMAXSQL(cARQPFG, "FEMEA WHERE PF=" & nPF, "ItEM", 0)) + 1
  Case 20                                      'ADCIONAL INCLUIR
    nNUM = FixInt(PegMAXSQL(cARQFEMEA, "FEMADC WHERE PF=" & nPF, "ITEM", 0)) + 1
  End Select

  If nNUM = 0 Then
    nNUM = FixInt(Val(Busca("Nº " & StrZero(nTMPSEQ, 2) & "/" & StrZero(nTMPSSQ, 2) & "/" & StrZero(nTMPITEM, 2) & "- " & cDESC, "Incluindo", CStr(nNUM), 8)))
  End If
  If nNUM = 0 Then
    Alert ("Nº em Branco")
    Exit Sub
  End If


  Select Case Index
  Case 0, 2, 4  '0 producao 2 padrao '4 fluxo 2
    cSQL = "select * from FEMEA WHERE PF=" & nPF & " AND ITEM=" & nNUM  'CARQPF
    IncluiSQL cARQFEMEAPF, cSQL, 6, Array("PF", "ITEM", "FXSEQ", "FXSSQ", "FXITEM", "PROCESSO"), Array(nPF, nNUM, nTMPSEQ, nTMPSSQ, nTMPITEM, cDESC), True, False

  Case 1, 3, 5  '3 complementar 1 padrao complemntar
    cSQL = "select * from FEMAVU WHERE PF=" & nPF & " AND ITEM=" & nNUM
    IncluiSQL cARQFEMEA, cSQL, 6, Array("PF", "ITEM", "FXSEQ", "FXSSQ", "FXITEM", "PROCESSO"), Array(nPF, nNUM, nTMPSEQ, nTMPSSQ, nTMPITEM, cDESC), True, False

  Case 6, 8, 10, 12  'preliminar
    cSQL = "select * from FEMEA WHERE PF=" & nPF & " AND ITEM=" & nNUM
    IncluiSQL cARQPFP, cSQL, 6, Array("PF", "ITEM", "FXSEQ", "FXSSQ", "FXITEM", "PROCESSO"), Array(nPF, nNUM, nTMPSEQ, nTMPSSQ, nTMPITEM, cDESC), True, False

  Case 7, 9, 11, 13  'gp12
    cSQL = "select * from FEMEA WHERE PF=" & nPF & " AND ITEM=" & nNUM
    IncluiSQL cARQPFG, cSQL, 6, Array("PF", "ITEM", "FXSEQ", "FXSSQ", "FXITEM", "PROCESSO"), Array(nPF, nNUM, nTMPSEQ, nTMPSSQ, nTMPITEM, cDESC), True, False

  Case 20         'Incluir adcional
    cSQL = "select * from FEMADC WHERE PF=" & nPF & " AND ITEM=" & nNUM
    IncluiSQL cARQFEMEA, cSQL, 6, Array("PF", "ITEM", "FXSEQ", "FXSSQ", "FXITEM", "PROCESSO"), Array(nPF, nNUM, nTMPSEQ, nTMPSSQ, nTMPITEM, cDESC), True, False
  End Select

  Select Case Index
  Case 0, 2
    FilRelat (0)
  Case 1, 3
    FilRelat (1)
  Case 6, 12
    FilRelat (2)
  Case 7, 13
    FilRelat (3)
  Case 20
    FilRelat (5)
  End Select
End Sub

Private Sub Command10_Click(Index As Integer)
  Dim nROWS As Integer
  Dim x
  If GridSeq(1).Rows = 0 Then
    Exit Sub
  End If
  nROWS = GridSeq(1).Rows
  nROWS = nROWS - 1
  For x = 1 To nROWS
    GridSeq(1).Row = x
    DizerBarra "Movendo femea " & x & "/" & nROWS
    CmdRevFemea_Click 14 + Index
  Next x
  DizerBarra ""
  FilRelat 1
  FilRelat 2
  FilRelat 4
  FilRelat 5
End Sub

Private Sub Command11_Click()
  CmdRevFemea_Click 14
  FilRelat 4
End Sub

Private Sub Command12_Click()
  CmdRevFemea_Click 15
  FilRelat 4
End Sub

Private Sub Command13_Click()
  Dim nROWS As Integer
  Dim x As Integer
  Dim cSQL As String
  If GridPadPF.Rows = 0 Then
    Exit Sub
  End If
  If Not MDG("Mover todos para historico") Then
    Exit Sub
  End If
  nROWS = GridPadPF.Rows
  nROWS = nROWS - 1
  For x = 1 To nROWS
    DizerBarra "Movendo femea padrao" & x & "/" & nROWS
    GridPadPF.Row = x
    CmdExcPFPad_Click 2
  Next x
  DizerBarra "Excluindo complementares"
  cSQL = "select * from FEMPF WHERE  PF=" & nPF
  ApagaSQL cARQFEMEA, cSQL
  cSQL = "select * from FEMEAPAD WHERE  PF=" & nPF
  ApagaSQL cARQFEMEA, cSQL
  cSQL = "select * from FEMRPNO WHERE PF=" & nPF
  ApagaSQL cARQFEMEA, cSQL
  cSQL = "select * from FEMRPNT WHERE  PF = " & nPF
  ApagaSQL cARQFEMEA, cSQL
  DizerBarra ""
  filgridpadpf
  FILGRIdPAD
End Sub

Private Sub Command14_Click()
  CmdExcPFPad_Click 2
  filgridpadpf
  FILGRIdPAD
End Sub

Private Sub Command15_Click()
  Dim cDIZ As String
  Encerrar.Visible = False
  cDIZ = "PF: " & nPF & "Revisao: " & TXTFIELDS(16) & " Nao necessita revisar femea"
  nREV = FixInt(TXTFIELDS(12))
  XComMotInc_Click (cDIZ)
  Check1.Value = False
  TXTFIELDS(25) = Today()
  cmdClose_Click
End Sub

Private Sub Command16_Click()
  DTPicker2.Value = Today()
End Sub

Private Sub Command2_Click(Index As Integer)
  Dim nSEQ As Integer
  Dim sSQL As String

  If Index = 10 Or Index = 15 Then             'todos os itens abaixo tratativa
  Else
    If GridSeq(Index).Row = 0 Then
      Alert "Linha Nao Selecionada"
      Exit Sub
    End If
  End If

  If Index = 3 Then
    If Not AcessaMenu("FemeaGP12", 1) Then
      Exit Sub
    End If
  End If
  If Index = 0 Then
    If Not AcessaMenu("FemeaPro", 1) Then
      Exit Sub
    End If
  End If
  If Index = 10 Then
    If Not AcessaMenu("FemeaPro", 1) Then
      Exit Sub
    End If
  End If

  If Index = 2 Then
    If Not AcessaMenu("FemeaPre", 1) Then
      Exit Sub
    End If
    If eproducao Then
      Exit Sub
    End If
  End If


  If Index = 10 Or Index = 15 Then             'todos os itens abaixo tratativa
  Else
    GridSeq(Index).Col = 1
    nSEQ = GridSeq(Index)
  End If


  If Index = 0 Then
    sSQL = "select * from FEMEA WHERE PF=" & nPF & " AND ITEM=" & nSEQ
    If ApagaSQLP(cARQFEMEAPF, sSQL) Then
      sSQL = "select * FEMrevi WHERE  TIPOAPU='N' AND PF=" & nPF & " AND ITEM=" & nSEQ
      ApagaSQL cARQFEMEA, sSQL
    End If
  End If
  If Index = 1 Then
    sSQL = "select * from FEMAVU WHERE PF=" & nPF & " AND ITEM=" & nSEQ
    If ApagaSQLP(cARQFEMEA, sSQL) Then
      sSQL = "select * FEMrevi WHERE  TIPOAPU='C' AND PF=" & nPF & " AND ITEM=" & nSEQ
      ApagaSQL cARQFEMEA, sSQL
    End If
  End If
  If Index = 2 Then
    sSQL = "select * from FEMEA WHERE PF=" & nPF & " AND ITEM=" & nSEQ
    If ApagaSQLP(cARQPFP, sSQL) Then
      sSQL = "select * FEMrevi WHERE  TIPOAPU='R' AND PF=" & nPF & " AND ITEM=" & nSEQ
      ApagaSQL cARQFEMEA, sSQL
    End If
  End If
  If Index = 3 Then
    sSQL = "select * from FEMEA WHERE PF=" & nPF & " AND ITEM=" & nSEQ
    If ApagaSQLP(cARQPFG, sSQL) Then
      sSQL = "select * FEMrevi WHERE  TIPOAPU='G' AND PF=" & nPF & " AND ITEM=" & nSEQ
      ApagaSQL cARQFEMEA, sSQL
    End If
  End If
  If Index = 5 Then                            'Adicionais
    sSQL = "select * from FEMADC WHERE PF=" & nPF & " AND ITEM=" & nSEQ
    If ApagaSQLP(cARQFEMEA, sSQL) Then
      sSQL = "select * FEMrevi WHERE  TIPOAPU='A' AND PF=" & nPF & " AND ITEM=" & nSEQ
      ApagaSQL cARQFEMEA, sSQL
    End If
  End If

  If Index = 10 Then                           'producao todos
    If MDG("Excluir Todos os Itens Producao") Then
      sSQL = "select * from FEMEA WHERE PF=" & nPF  '
      If ApagaSQLP(cARQFEMEAPF, sSQL) Then
        sSQL = "select * FEMrevi WHERE  TIPOAPU='N' AND PF=" & nPF
        ApagaSQL cARQFEMEA, sSQL
      End If
    End If
  End If
  If Index = 15 Then                           'adcionais todos
    If MDG("Excluir Todos os Itens Adicionais") Then
      sSQL = "select * from FEMADC WHERE PF=" & nPF  '// & " AND ITEM=" & nSEQ
      If ApagaSQLP(cARQPF, sSQL) Then
        sSQL = "select * FEMrevi WHERE  TIPOAPU='A' AND PF=" & nPF  ''& " AND ITEM=" & nSEQ
        ApagaSQL cARQFEMEA, sSQL
      End If
    End If
  End If
  If Index >= 10 Then
    Index = Index - 10
  End If
  FilRelat (Index)
End Sub

Private Sub Command3_Click()
  If MDG("Gravar Elaborador", "Confirme Gravação") Then
    frmfemea.TXTFIELDS(6) = zIDFOLHA
    frmfemea.TXTFIELDS(9) = zNOMEFOLHA
    TXTFIELDS(22) = Date
    If IsNull(TXTFIELDS(23)) Or TXTFIELDS(23) = "" Then
      TXTFIELDS(23) = Date
    End If
  End If
End Sub

Private Sub Command4_Click(Index As Integer)

  Dim nSEQ As Integer
  Dim nNUM As Integer
  Dim cSQL As String

  If Index = 0 And Not AcessaMenu("FemeaPro", 1) Then Exit Sub
  If Index = 3 And Not AcessaMenu("FemeaGP12", 1) Then Exit Sub

  If Index = 2 Then
    If Not AcessaMenu("FemeaPre", 1) Then
      Exit Sub
    End If
    If eproducao Then
      Exit Sub
    End If

  End If

  If GridSeq(Index).Row = 0 Then
    Exit Sub
  End If

  nNUM = 0

  nNUM = FixInt(Val(Busca("Nº", "Renumerando", CStr(nNUM), 8)))
  If nNUM = 0 Then
    Alert ("Nº em Branco")
    Exit Sub
  End If
  GridSeq(Index).Col = 1
  nSEQ = GridSeq(Index)
  If Index = 0 Then
    cSQL = "select * from FEMEA WHERE PF=" & nPF & " AND ITEM=" & nSEQ
    GrvSQL cARQFEMEAPF, cSQL, 1, Array("ITEM"), Array(nNUM), Array("NI")
  End If
  If Index = 1 Then
    cSQL = "select * from FEMAVU WHERE PF=" & nPF & " AND ITEM=" & nSEQ
    GrvSQL cARQFEMEA, cSQL, 1, Array("ITEM"), Array(nNUM), Array("NI")
  End If
  If Index = 2 Then
    cSQL = "select * from FEMEA WHERE PF=" & nPF & " AND ITEM=" & nSEQ
    GrvSQL cARQPFP, cSQL, 1, Array("ITEM"), Array(nNUM), Array("NI")
  End If
  If Index = 3 Then
    cSQL = "select * from FEMEA WHERE PF=" & nPF & " AND ITEM=" & nSEQ
    GrvSQL cARQPFG, cSQL, 1, Array("ITEM"), Array(nNUM), Array("NI")
  End If
  If Index = 5 Then
    cSQL = "select * from FEMADC WHERE PF=" & nPF & " AND ITEM=" & nSEQ
    GrvSQL cARQFEMEA, cSQL, 1, Array("ITEM"), Array(nNUM), Array("NI")
  End If
  FilRelat (Index)

End Sub

Private Sub Command5_Click()
  Encerrar.Visible = False
  CmdNovaREV_Click
  nREV = FixInt(TXTFIELDS(12))
  XComMotInc_Click ("Revisão Periodica")
  Filgridrevi (0)

  Command3_Click

  eRETU02 = "PF:" & nPF & Chr(13) & Chr(10)
  eRETU02 = eRETU02 & " Desenho:" & TXTFIELDS(2) & Chr(13) & Chr(10)
  eRETU02 = eRETU02 & " Data:" & Fdata(Date) & Chr(13) & Chr(10)
  MAILENV "FEM00001", eRETU02


  cmdClose_Click
End Sub

Private Sub Command6_Click()

  FrmFemeaPad.Show vbModal, Me
  FILGRIdPRO
  FILGRIdFAL
  FILGRIdPAD
  filgridpadpf
  ''zera pois o usuario pode ter mechido
  aNUM = Array(0, 0, 0, 0)
  aDIZ = Array("", "", "", "")
  dizpro.Caption = ""
  dizfal.Caption = ""

End Sub

Private Sub Command7_Click()
  LocalizaGri1 Gridfemprecad
End Sub

Private Sub Command8_Click()
  Dim cDIZ As String
  Dim nROWS As Integer
  Dim nTMPFAL As Integer
  Dim x As Integer
  On Error Resume Next

  If aNUM(0) = 0 Then
    Alert ("Processo Nao Escolhido")
    Exit Sub
  End If
  cDIZ = pegdizseqssq(dizpro, cARQPF)
  If Len(cDIZ) = 0 Then
    Alert ("Descricao Nao Especificada")
    Exit Sub
  End If
  If Not MDG(cDIZ, "Incluir") Then
    Exit Sub
  End If
  If Not MDG("Incluir Todas Falhas") Then
    Exit Sub
  End If
  nROWS = GRIDFAL.Rows - 1
  GRIDFAL.Col = 1
  For x = 0 To nROWS
    GRIDFAL.Row = x
    nTMPFAL = FixInt(GRIDFAL)
    If nTMPFAL > 0 Then
      GERAFEMEAPAD aNUM(0), nTMPFAL, FixInt(nPF), cDIZ
    End If
  Next x
  filgridpadpf
  FILGRIdPAD

End Sub

Private Sub Command9_Click()
  Dim nSEQTMP As Integer
  Dim nPROTMP As Long
  GridOpe.Col = 1
  nSEQTMP = FixInt(GridOpe)
  GridOpe.Col = 2
  nPROTMP = FixInt(GridOpe)
  ePASS01 = "SELECT * FROM FEMRPNO WHERE SEQ=" & nSEQTMP & " AND PF=" & nPF & " AND PRONUM=" & nPROTMP
  FrmFEMRPNO.Show vbModal, Me

End Sub

Private Sub ComMotAlt_Click(Index As Integer)
  On Error Resume Next

  If gridrevi(Index).Row = 0 Then
    Exit Sub
  End If

  gridrevi(Index).Col = 1
  nREV = FixInt(gridrevi(Index))
  gridrevi(Index).Col = 3
  nREVI = FixInt(gridrevi(Index))

  Select Case Index
  Case 0                                       'Femea
    lEDITAR = True
    ePASS01 = "select * from revi WHERE PF=" & nPF & " AND TIPO='FE' AND REVISAO=" & nREV & " AND ITEM=" & nREVI
  Case 1                                       'pf
    lEDITAR = False
    ePASS01 = "select * from revi WHERE PF=" & nPF & " AND TIPO='PF' AND REVISAO=" & nREV & " AND ITEM=" & nREVI
  End Select

  frmREVI.Show vbModal, Me
  Filgridrevi (Index)

End Sub

Private Sub ComMotExc_Click()
  Dim sSQL As String
  If gridrevi(0).Row = 0 Then
    Exit Sub
  End If
  gridrevi(0).Col = 3
  nREVI = FixInt(gridrevi(0))
  sSQL = "select * from revi WHERE PF=" & nPF & " AND TIPO='FE' AND REVISAO=" & nREV & " AND  ITEM=" & nREVI
  ApagaSQLP cARQPF, sSQL
  Filgridrevi (0)

End Sub

Private Sub ComMotInc_Click()

  XComMotInc_Click ("")

End Sub

Private Sub EditSeq_Click(Index As Integer)

  ePASS03 = Index                              'Para usar com o revisao abaixo 'atualizacao de versao
  ePASS04 = TXTFIELDS(12)                      ''REVISAO PF

  If Index = 3 And (Not AcessaMenu("FemeaGP12", 1)) Then Exit Sub
  If Index = 0 And (Not AcessaMenu("FemeaPro", 1)) Then Exit Sub

  If Index = 2 Then
    If Not AcessaMenu("FemeaPre", 1) Then
      Exit Sub
    End If
    If eproducao Then
      Exit Sub
    End If
  End If

  If GridSeq(Index).Row = 0 Then
    Alert "Escolha Linha diferente do cabecario"
    Exit Sub
  End If
  GridSeq(Index).Col = 1
  nSEQ = GridSeq(Index)
  If Index = 0 Then  'producao
    ePASS01 = "select * from FEMEA WHERE PF=" & nPF & " AND ITEM=" & nSEQ
    ePASS02 = cARQFEMEAPF
  End If
  If Index = 1 Then  'padrao complementar
    ePASS01 = "select * from FEMAVU WHERE PF=" & nPF & " AND ITEM=" & nSEQ
    ePASS02 = cARQFEMEA
  End If
  If Index = 2 Then  'pre
    ePASS01 = "select * from FEMEA WHERE PF=" & nPF & " AND ITEM=" & nSEQ
    ePASS02 = cARQPFP
  End If
  If Index = 3 Then  'gp12
    ePASS01 = "select * from FEMEA WHERE PF=" & nPF & " AND ITEM=" & nSEQ
    ePASS02 = cARQPFG
  End If
  If Index = 5 Then  'adcional
    ePASS01 = "select * from FEMADC WHERE PF=" & nPF & " AND ITEM=" & nSEQ
    ePASS02 = cARQFEMEA
  End If
  GridSeq(Index).Col = 2
  If GridSeq(Index).CellBackColor = vbRed Then
    Alert ("Sequencia ou sub sequencia inexistente no  processo")
  End If
  If GridSeq(Index).CellBackColor = vbYellow Then
    Alert ("Sequencia ou sub sequencia ou item em branco")
  End If


  frmFEMEI.Show vbModal, Me
  FilRelat (Index)

End Sub

Private Sub Filgridrev()

  Dim cSQL As String
  Dim cTIPO As String
  cTIPO = "FE"
  DizerBarra "Carregando revisoes"
  cSQL = "select PF,REVISAO,TIPO,DATA from rev WHERE PF=" & nPF & " AND TIPO='" & cTIPO & "' ORDER BY REVISAO"
  MontaGridUltra gridrev, 4, Array(600, 400, 400, 2000), Array("PF", "REV", "Tipo", "DATA"), Array("PF", "REVISAO", "TIPO", "L$DATA"), cARQPF, cSQL
  DizerBarra ""

End Sub

Private Sub Filgridrevi(nIndex)

  Dim cTIPO As String
  Dim cSQL As String

  If nIndex > 0 Or nREV > 0 Then
    Select Case nIndex
    Case 0
      cTIPO = "FE"
      cSQL = "select PF, REVISAO, TIPO, ITEM,MOTIVO from revI WHERE PF=" & nPF & " AND REVISAO=" & nREV & " AND TIPO='FE' ORDER BY ITEM"
    Case 1
      cTIPO = "PF"
      cSQL = "select PF, REVISAO, TIPO, ITEM,MOTIVO  from revI WHERE PF=" & nPF & " AND TIPO='" & cTIPO & "' ORDER BY ITEM"
    End Select
    DizerBarra "Carregando Itens Revisao" & nIndex
    MontaGridFast gridrevi(nIndex), 5, Array(400, 400, 400, 400, 5000), Array("PF", "Rev", "Tipo", "Item", "Motivo"), Array("PF", "REVISAO", "TIPO", "ITEM", "MOTIVO"), cARQPF, cSQL
    DizerBarra ""
  End If
End Sub

Private Sub Filopr()
  Dim cSQL As String
  cSQL = "select PF, SEQ, SSQ, CODINT, DESCRI from PFS WHERE PF=" & nPF & " ORDER BY SEQ,SSQ"
  DizerBarra "Carregando operacoes"
  MontaGridFast gridopr, 5, Array(400, 400, 400, 1600, 6000), Array("PF", "SEQ", "SSQ", "Cod Int", "Descriçao"), Array("PF", "SEQ", "SSQ", "L$CODINT", "DESCRI"), cARQPF, cSQL
  DizerBarra ""
End Sub

Private Sub FilRelat(Index As Integer)
  Dim cSQL As String
  Dim cARQ As String
  Dim i As Integer
  Dim nROWS  As Integer
  Dim lTEM As Boolean
  Dim nTMPSEQ   As Integer
  Dim nTMPSSQ   As Integer
  Dim nTMPITEM   As Integer
  Dim nROWSQ      As Integer
  Dim nTMPSEQQ     As Integer
  Dim nTMPSSQQ   As Integer
  Dim nTMPITEMQ   As Integer
  Dim J   As Integer

  lRETU = True

  Select Case Index
  Case 0
    cARQ = cARQFEMEAPF
    cSQL = "select PF,ITEM,PROCESSO,FXSEQ,FXSSQ,FXITEM,FXITEMS,FALCAU from FEMEA WHERE PF=" & nPF & " ORDER BY " & cORDFEMUSO
    DizerBarra "carregando femea padrao"
  Case 1
    cARQ = cARQFEMEA
    cSQL = "select PF,ITEM,PROCESSO,FXSEQ,FXSSQ,FXITEM,FXITEMS,FALCAU from FEMAVU WHERE PF=" & nPF & " ORDER BY " & cORDFEMAVU
    DizerBarra "carregando femea avulso"
  Case 2
    cARQ = cARQPFP
    cSQL = "select PF,ITEM,PROCESSO,FXSEQ,FXSSQ,FXITEM,FXITEMS,FALCAU from FEMEA WHERE PF=" & nPF & " ORDER BY " & cORDFEMPRE
    DizerBarra "carregando femea pre"
  Case 3
    cARQ = cARQPFG
    cSQL = "select PF,ITEM,PROCESSO,FXSEQ,FXSSQ,FXITEM,FXITEMS,FALCAU from FEMEA WHERE PF=" & nPF & " ORDER BY " & cORDFEMGP12
    DizerBarra "carregando femea gp12"
  Case 4
    cARQ = cARQFEMEA
    cSQL = "select PF,ITEM,PROCESSO,FXSEQ,FXSSQ,FXITEM,FXITEMS,FALCAU from FEMHIS WHERE PF=" & nPF & " ORDER BY ITEM"  '& cORDFEMAVU
    DizerBarra "carregando femea historico"
  Case 5
    cARQ = cARQFEMEA
    cSQL = "select PF,ITEM,PROCESSO,FXSEQ,FXSSQ,FXITEM,FXITEMS,FALCAU from FEMADC WHERE PF=" & nPF & " ORDER BY " & cORDFEMADC
    DizerBarra "carregando femea adicional"


  End Select
  MontaGridFast GridSeq(Index), 7, Array(600, 600, 400, 400, 400, 600, 5000), Array("PF", "REG", "SEQ", "SSQ", "IT", "ITS", "Descriçao"), Array("PF", "ITEM", "FXSEQ", "FXSSQ", "FXITEM", "FXITEMS", "FALCAU"), cARQ, cSQL
  DizerBarra "verificando sequencias"
  'nROWSQ = GrdQsbLep.Rows - 1
  nROWSQ = gridopr.Rows - 1
  nROWS = GridSeq(Index).Rows
  nROWS = nROWS - 1
  For i = 1 To nROWS                           'ROW 0 = Cabecario
    GridSeq(Index).Row = i
    GridSeq(Index).Col = 2
    nTMPSEQ = FixInt(GridSeq(Index))
    GridSeq(Index).Col = 3
    nTMPSSQ = FixInt(GridSeq(Index))
    GridSeq(Index).Col = 4
    nTMPITEM = FixInt(GridSeq(Index))
    lTEM = False

    For J = 1 To nROWSQ                      ' zero e o titulo do grid comeca 1

      gridopr.Row = J
      gridopr.Col = 1
      nTMPSEQQ = FixInt(gridopr)
      gridopr.Col = 2
      nTMPSSQQ = FixInt(gridopr)
      If nTMPSEQ = nTMPSEQQ And nTMPSSQ = nTMPSSQQ Then  'And nTMPITEM = nTMPITEMQ Then
        lTEM = True
      End If


    Next J
    If Not lTEM Then
      GridSeq(Index).Col = 2
      GridSeq(Index).CellBackColor = vbRed
      GridSeq(Index).Col = 3
      GridSeq(Index).CellBackColor = vbRed
      GridSeq(Index).Col = 4
      GridSeq(Index).CellBackColor = vbRed
      If Index = 0 Then
        lRETU = False
      End If
    End If
    GridSeq(Index).Col = 2
    If Val(GridSeq(Index)) = 0 Then
      GridSeq(Index).CellBackColor = vbYellow  '&HC0FFFF
      If Index = 0 Then
        lRETU = False
      End If
    End If
    GridSeq(Index).Col = 3
    If Val(GridSeq(Index)) = 0 Then
      GridSeq(Index).CellBackColor = vbYellow  '&HC0FFFF
      If Index = 0 Then
        lRETU = False
      End If
    End If
    GridSeq(Index).Col = 4
    If Val(GridSeq(Index)) = 0 Then
      GridSeq(Index).CellBackColor = vbYellow  '&HC0FFFF
      If Index = 0 Then
        lRETU = False
      End If
    End If
    GridSeq(Index).Col = 5
    If Val(GridSeq(Index)) = 0 Then
      GridSeq(Index).CellBackColor = vbYellow  '&HC0FFFF
      If Index = 0 Then
        lRETU = False
      End If
    End If
  Next i

  If Index = 0 And Not lRETU Then
    Alert ("Existes operacao ou itens nao preenchidos")
  End If

  DizerBarra ""
End Sub

Private Sub Encerrar_Click()
  If Not MDG("Sair sem gravar") Then
    chkfemeas
    Exit Sub
  End If
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub escidfolha_Click(Index As Integer)
  ePASS01 = ""
  If Index > 9 Then
    ePASS01 = "LOGIX"
  End If

  escMP04.Show vbModal, Me
  If lRETU Then
    Select Case Index
    Case 0, 10
      TXTFIELDS(7).Text = TXTFIELDS(7).Text & " , " & Trim(MMCase(eRETU02))
    Case 1, 11
      TXTFIELDS(8).Text = TXTFIELDS(8).Text & " , " & Trim(MMCase(eRETU02))
    End Select
  End If
End Sub

Private Sub ESCprofin_Click(Index As Integer)
  Dim cCHAVEBUS As String
  ePASS01 = "MANA5"
  If Index = 1 Then
    ePASS01 = "LOGIX"
  End If
  If Index = 2 Then
    ePASS01 = "MICRO"
  End If
  If Len(TXTFIELDS(61)) = 0 Then
    cCHAVEBUS = TXTFIELDS(60)
  Else
    cCHAVEBUS = TXTFIELDS(61)
  End If
  escms01.Show vbModal, Me
  If lRETU Then
    TXTFIELDS(61) = eRETU01
  End If
  cCHAVEBUS = ""
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  TXTPF = nPF
  aNUM = Array(0, 0, 0, 0)
  aDIZ = Array("", "", "", "")
  cARQFEMEA = PegPath("PATH", "FEMEA")
  cARQFEMEAPF = PegPath("PATH", "FEMEAPF")
  cARQPF = PegPath("PATH", "PF")
  cARQPFP = PegPath("PATH", "PFP")
  cARQPFG = PegPath("PATH", "PFG")

  CHECKPFPG                                    'inclui embalagem preliminar gp12

  nCAMPOS = 34
  cSQL = "select * from PF WHERE PF=" & nPF
  aCAM = Array("CODFINAL", "CPF", "CODIGO", "DESCR", "CLIENTE", _
               "CLINOME", "FEMEAF", "FEMEAG", "FEMEAC", "FEMEAEN", _
               "NOSREV", "NOSDAT", "FEMEAREV", "FEMEAF", "FEMEAR", _
               "CODIGOINT", "REVPRO", "REVDAT", "FEMEAOBS", "OPCAO", _
               "STFEPC", "CODFISCAL", "FEMEAED", "FEMEAREVD", "FEMEAD", _
               "FEMEAREVD2", "TIPO", "STPFFE", "EXCRPN", "FEMEACRG", _
               "FEMEAPREPRO", "FEMEAPRO", "femeaano", "femeaproj")
  aFOR = Array("C", "C", "C", "C", "NI", _
               "C", "C", "C", "C", "C", _
               "NI", "D", "NI", "NI", "C", _
               "C", "NI", "D", "C", "NI", _
               "B", "C", "D", "DZ", "D", _
               "DZ", "C", "B", "B", "B", _
               "DZ", "D", "C", "C")
  aPAD = Array("", "", "", "", 0, _
               "", "", "", "", "", _
               0, "", 0, 0, "", _
               "", 0, "", "", 0, _
               "", "", "", "", "", _
               "", "", "", "", "", _
               "", "", "", "")
  aVAL = PegSQL(cARQPF, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
  For iLOOP = 0 To nCAMPOS - 8
    TXTFIELDS(iLOOP) = aVAL(iLOOP)
  Next iLOOP


  Check1.Value = FixBolNum(aVAL(27))
  Check2.Value = FixBolNum(aVAL(28))
  femeacrg.Value = FixBolNum(aVAL(29))
  If IsDate(aVAL(30)) Then
    DTPicker2.Value = aVAL(30)
  End If
  If IsDate(aVAL(31)) Then
    DTPicker1.Value = aVAL(31)
  End If

  For iLOOP = 32 To nCAMPOS - 1
    TXTFIELDS(iLOOP) = aVAL(iLOOP)
  Next iLOOP


  cORDFEMUSO = "FXSEQ,FXSSQ,FXITEM,FXITEMS"
  cORDFEMPAD = "PRONUM,FALNUM,EFENUM,CAUNUM"
  cORDFEMAVU = "FXSEQ,FXSSQ,FXITEM"
  cORDFEMPRE = "FXSEQ,FXSSQ,FXITEM"
  cORDFEMGP12 = "FXSEQ,FXSSQ,FXITEM"
  cORDFEMADC = "FXSEQ,FXSSQ,FXITEM"
  cORDREVI = "FXSEQ,FXSSQ,FXITEM,FXITEMS,FEMEAREV"

  Filopr                                       'antes das relat
  Text(1) = CStr(Month(Date))
  Text(2) = CStr(Year(Date))
  FilQSBLEP
  FilRelat (0)
  FilRelat (1)
  FilRelat (2)
  FilRelat (3)
  FilRelat (4)
  Filgridrev
  'Filopr
  Filgridrevi (1)
  FILGRIdPRO
  FILGRIdFAL
  FILGRIdfemprecad
  filgridpadpf
  FILGRIdpfrev
  If zUSER = "ADMIN" Then
    TXTFIELDS(12).Enabled = True
    TXTFIELDS(12).Locked = False
    TXTFIELDS(23).Enabled = True
    TXTFIELDS(23).Locked = False
  Else
    CmdApaRev.Visible = False
    CmdTrocaREv.Visible = False
    TxtRevisao.Visible = False
    TxtRevNova.Visible = False
  End If
  FILGRIDrevCAU (0)
  FILGRIDrevCAU (1)
  FILGRIDrevCAU (2)
  FILGRIDrevCAU (3)
  FILGRIDrevCAU (5)
  chkfemeas


  txtsigi.Font = "isoqsymbol"


End Sub

Private Sub chkfemeas()
  If GridSeq(1).Rows > 1 Then
    If GridSeq(0).Rows > 1 Then
      Alert "Femea com complementar nao pode usar normal"
    End If
  End If
  If GridSeq(0).Rows > 1 Then
    If GridSeq(1).Rows > 1 Then
      Alert "Femea normal nao pode usar complementar"
    End If
  End If
  If GridPadPF.Rows > 1 Then
    If GridSeq(0).Rows > 1 Then
      Alert "Femea com Padrao nao pode usar normal"
    End If
  End If
  If GridSeq(0).Rows > 1 Then
    If GridPadPF.Rows > 1 Then
      Alert "Femea com normal nao pode usar padrao"
    End If
  End If
End Sub

Private Sub importar_Click(Index As Integer)

  If Index = 0 And Not AcessaMenu("FemeaPro", 1) Then Exit Sub
  If Index = 3 And Not AcessaMenu("FemeaGP12", 1) Then Exit Sub

  If Index = 1 Then
    If GridPadPF.Rows < 2 Then
      Alert "Este Femea nao utiliza Padrao"
      Exit Sub
    End If
  End If
  If Index = 0 Then                            'normal com complementar
    If GridSeq(1).Rows > 1 Then
      Alert "Femea com complemetar nao pode usar normal"
      Exit Sub
    End If
  End If
  If Index = 0 Then                            'normal com padrao
    If GridPadPF.Rows > 1 Then
      Alert "Femea com Padrao nao pode usar normal"
      Exit Sub
    End If
  End If

  If Index = 2 Then
    If Not AcessaMenu("FemeaPre", 1) Then
      Exit Sub
    End If
    If eproducao Then
      Exit Sub
    End If
  End If

  Select Case Index
  Case 0
    cARQIMP = "FEMEA"
  Case 1
    cARQIMP = "FEMAVU"
  Case 2
    cARQIMP = "FEMEAPRE"
  Case 3
    cARQIMP = "FEMEAGP12"
  Case 5
    cARQIMP = "FEMADC"
  Case 10
    cARQIMP = "FEMAX"
  End Select
  FrmImp.Show vbModal, Me
  If Index = 10 Then
    If lRETU Then
      Xcmdimp_Click eRETU01
    End If
  Else
    FilRelat (Index)
  End If
End Sub

Private Sub XComMotInc_Click(cDIZ)
  Dim nITEM As Long
  Dim cSQL As String
  If nREV > 0 Then
    nITEM = FixInt(PegMAXSQL(cARQPF, "revI WHERE PF=" & nPF & " AND TIPO='FE' AND REVISAO=" & nREV, "ITEM", 0)) + 1
    cSQL = "select * from revi WHERE PF=" & nPF & " AND TIPO='FE' AND REVISAO=" & nREV & " AND  ITEM=" & nITEM
    IncluiSQL cARQPF, cSQL, 5, Array("PF", "REVISAO", "TIPO", "ITEM", "MOTIVO"), Array(FixInt(nPF), FixInt(nREV), "FE", nITEM, cDIZ), True, False
    Filgridrevi (0)
  End If
End Sub

Private Sub filgridmsrd()
  Dim cARQ As String
  Dim cSQL As String
  DizerBarra "carregando registro de defeitos"
  cARQ = PegPath("PATH", "MANA5TGQ")
  cSQL = "SELECT MSRD,DATA,DESC01,INV01,ACA01 FROM MSRD WHERE PF=" & nPF & " ORDER BY MSRD"
  cARQ = GeraConn(cARQ, "JETFOX")
  MontaGridUltra gridmsrd, 5, Array(800, 1200, 2000, 2000, 2000), Array("Nº", "Data", "Desc", "Inv", "Acao"), Array("MSRD", "C$DATA", "L$DESC01", "L$INV01", "L$ACA01"), cARQ, cSQL
  DizerBarra ""
End Sub

Private Sub FILGRIdPRO()
  Dim cSQL As String
  DizerBarra "carregando femea processos"
  cSQL = "SELECT PRONUM,PROCESSO FROM FEMPRO ORDER BY PRONUM"
  MontaGridFast GridPRO, 2, Array(800, 3000), Array("Nº", "Descricao"), Array("PRONUM", "PROCESSO"), cARQFEMEA, cSQL
  DizerBarra ""
End Sub

Private Sub FILGRIdFAL()
  Dim cSQL As String
  DizerBarra "carregando femea falhas"
  cSQL = "SELECT PRONUM,FALNUM, FALTIP FROM FEMFAL WHERE PRONUM=" & FixInt(aNUM(0)) & " order by FALNUM"
  MontaGridFast GRIDFAL, 3, Array(800, 800, 3000), Array("NºPro", "Nº", "Descricao"), Array("PRONUM", "FALNUM", "FALTIP"), cARQFEMEA, cSQL
  DizerBarra ""
End Sub

Private Sub filgridpadpf()
  Dim cSQL As String
  DizerBarra "carregando femea padrao"
  cSQL = "SELECT PF,PRONUM,FALNUM FROM FEMPF WHERE PF=" & nPF & " ORDER BY PRONUM,FALNUM"
  MontaGridUltra GridPadPF, 3, Array(400, 400, 400), Array("PF", "NºPro", "NºFAL"), Array("PF", "PRONUM", "FALNUM"), cARQFEMEA, cSQL
  DizerBarra "carregando femea historico"
  cSQL = "SELECT PF,PRONUM,FALNUM FROM FEMPF WHERE PF=" & nPF & " ORDER BY PRONUM,FALNUM"
  MontaGridUltra gridpfpadhist, 3, Array(400, 400, 400), Array("PF", "NºPro", "NºFAL"), Array("PF", "PRONUM", "FALNUM"), cARQFEMEA, cSQL
  DizerBarra ""
End Sub

Private Sub FILGRIdpfrev()
  Dim cSQL As String
  DizerBarra "carregando pfrev"
  cSQL = "SELECT PF,REV,PFDATA,FEMEAD,FEMEAN FROM PFREV WHERE PF=" & nPF & " ORDER BY REV"
  MontaGridUltra GridPFrev, 5, Array(400, 400, 1200, 1200, 2000), _
                 Array("PF", "REV", "DATA", "FEMEA Em", "POR"), _
                 Array("PF", "REV", "PFDATA", "FEMEAD", "FEMEAN"), cARQPF, cSQL
  DizerBarra ""
End Sub

Private Sub FILGRIdfemprecad()
  Dim cSQL As String
  DizerBarra "carregando femea pre cadastro"
  cSQL = "SELECT SEGGRA,TITULO FROM FEMPRE ORDER BY TITULO"
  MontaGridUltra Gridfemprecad, 2, Array(600, 4000), Array("Código", "Descrição"), Array("SEGGRA", "L$TITULO"), cARQFEMEA, cSQL
  DizerBarra ""
End Sub

Private Sub FILGRIdPAD()
  Dim cSQL As String
  DizerBarra "carregando femea"
  cSQL = "SELECT ITEM,INDRIS, PRONUM, FALNUM, EFENUM, CAUNUM, ALTMAN, MUDPAD, PROCESSO, FALTIP, FALEFE, FALCAU FROM FEMEAPAD WHERE PF=" & nPF & " ORDER BY " & cORDFEMPAD
  'cSQL = "SELECT ITEM,INDRIS, PRONUM, FALNUM, EFENUM, CAUNUM, ALTMAN, MUDPAD, PROCESSO, FALTIP, FALEFE, FALCAU FROM FEMEA WHERE PF=" & nPF & " ORDER BY " & cORDFEMPAD
  MontaGrid GridPad, 12, Array(400, 400, 400, 400, 400, 400, 300, 300, 1000, 2000, 3000, 2000), Array("ITEM", "RPN", "NºPro", "NºFal", "NºEfe", "NºCau", "Man", "Mod", "Processo", "Falha", "Efeito", "Causa"), Array("ITEM", "INDRIS", "PRONUM", "FALNUM", "EFENUM", "CAUNUM", "ALTMAN", "MUDPAD", "PROCESSO", "FALTIP", "FALEFE", "FALCAU"), cARQFEMEA, cSQL
  DizerBarra ""
End Sub

Private Sub maisano_Click()
  Text(2) = CStr(INCDEC(Text(2), 1, 1900, 2300))
End Sub

Private Sub maismes_Click()
  Text(1) = CStr(INCDEC(Text(1), 1, 1, 12))
End Sub

Private Sub menosano_Click()
  Text(2) = CStr(INCDEC(Text(2), -1, 1900, 2300))
End Sub

Private Sub menosmes_Click()
  Text(1) = CStr(INCDEC(Text(1), -1, 1, 12))
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub Text_GotFocus(Index As Integer)
  FocusMe
End Sub

Private Sub tEXT_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = ValiText(KeyAscii, "#NI")
End Sub

Private Sub FilQSBLEP()
  Dim sSQL As String
  DizerBarra "carregando fluxo II"
  sSQL = "select SEQ,SSQ,ITEM,DESCRICAO from PFQSBLEP WHERE PF=" & nPF & " ORDER BY SEQ,SSQ,ITEM"
  MontaGridUltra GrdQsbLep, 4, Array(400, 400, 400, 7000), Array("Seq", "SSQ", "Ord", "descricao"), _
                 Array("SEQ", "SSQ", "Item", "descricao"), cARQPF, sSQL
  DizerBarra ""
End Sub

Private Sub FILGRIDrevCAU(Index As Integer)
  Dim cSQL As String
  If Index = 0 Or Index = 1 Or Index = 2 Or Index = 3 Or Index = 5 Then
    cSQL = "select ITEM,SEGGRA,FXSEQ,FXSSQ,FXITEM,FXITEMS,FEMEAREV,FALCAU from FEMREVI WHERE  TIPOAPU='" & tipoapu(Index) & "' AND PF=" & nPF & " ORDER BY " & cORDREVI  ''ITEM,SEGGRA"
    DizerBarra "carregando femea revisoes " & tipoapu(Index)
  End If
  MontaGridFast gridrevcau(Index), 8, Array(400, 400, 500, 500, 400, 400, 400, 7000), Array("REG", "REGREV", "SEQ", "SSQ", "IT", "ITS", "FEREV", "Descriçao"), Array("ITEM", "SEGGRA", "FXSEQ", "FXSSQ", "FXITEM", "FXITEMS", "FEMEAREV", "FALCAU"), cARQFEMEA, cSQL
  DizerBarra ""
End Sub

Private Function tipoapu(ByVal Index As Integer)
  tipoapu = ""                                 ''N-ormal C-omplementar P-adrao R-pre G-gp12
  If Index = 0 Then
    tipoapu = "N"
  End If
  If Index = 1 Then
    tipoapu = "C"
  End If
  If Index = 2 Then
    tipoapu = "R"
  End If
  If Index = 3 Then
    tipoapu = "G"
  End If
  If Index = 5 Then
    tipoapu = "A"
  End If
End Function

Private Function eproducao()
  eproducao = False
  If IsDate(DTPicker2.Value) Then
    Alert "Ja Passou para producao"
    eproducao = True
  Else
    If TXTFIELDS(26).Text = "P" Then
      Alert "O Processo de Fabricao ja e TIPO (P) producao"
      eproducao = True
    End If
  End If
End Function


