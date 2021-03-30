VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.0#0"; "vbccr17.ocx"
Begin VB.Form frmFEMEI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FMEA"
   ClientHeight    =   8130
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   12945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   405
      Index           =   46
      Left            =   8400
      Locked          =   -1  'True
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdpeg 
      Height          =   360
      Index           =   44
      Left            =   10560
      Picture         =   "frmFEMEI.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   405
      Index           =   44
      Left            =   9960
      Locked          =   -1  'True
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton grvitem 
      Caption         =   "GravarPre"
      Height          =   375
      Left            =   11040
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1275
   End
   Begin VB.CommandButton CmdRetornaAcao 
      Caption         =   "Retorna Revisao Reg"
      Enabled         =   0   'False
      Height          =   435
      Index           =   1
      Left            =   11040
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton CmdRetornaAcao 
      Caption         =   "Retorna Revisao Femea"
      Enabled         =   0   'False
      Height          =   435
      Index           =   0
      Left            =   11040
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1335
   End
   Begin VB.CommandButton CmdAltman 
      Caption         =   "Alteracao Manual"
      Height          =   495
      Left            =   11280
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CheckBox Check3 
      Height          =   375
      Left            =   11040
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   1200
      Width           =   195
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   405
      Index           =   45
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   405
      Index           =   37
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   405
      Index           =   36
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   405
      Index           =   35
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmddizsessq 
      Caption         =   "Seq"
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton cmddizsessq 
      Caption         =   "Fx II"
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmddizsessq 
      Caption         =   "Oper"
      Height          =   255
      Index           =   0
      Left            =   3240
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton pegitem 
      Caption         =   "UsarPre"
      Height          =   255
      Left            =   2520
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   120
      Width           =   795
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   405
      Index           =   1
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   405
      Index           =   0
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   120
      Width           =   735
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   6
      TabHeight       =   520
      TabCaption(0)   =   "Passo 2"
      TabPicture(0)   =   "frmFEMEI.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblLabels(23)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblLabels(24)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblLabels(25)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtFields(47)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtFields(48)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtFields(49)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Passo 3"
      TabPicture(1)   =   "frmFEMEI.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblLabels(26)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblLabels(27)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblLabels(28)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtFields(50)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "txtFields(51)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtFields(52)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Passo 4 5"
      TabPicture(2)   =   "frmFEMEI.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblLabels(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label1(0)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblLabels(4)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lblLabels(5)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "lblLabels(16)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "lblLabels(6)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label1(1)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "lblLabels(8)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "lblLabels(7)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "lblLabels(9)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "lblLabels(10)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "lblLabels(14)"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "lblLabels(21)"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "lblLabels(29)"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "lblLabels(33)"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "lblLabels(34)"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "txtFields(38)"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "cmdpegpro(0)"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "txtFields(2)"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "txtFields(39)"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "cmdpegpro(1)"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "txtFields(3)"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "txtFields(40)"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "cmdpegpro(2)"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "txtFields(4)"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "txtFields(41)"
      Tab(2).Control(25).Enabled=   0   'False
      Tab(2).Control(26)=   "cmdpegpro(3)"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "txtFields(5)"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).Control(28)=   "cmdpeg(25)"
      Tab(2).Control(28).Enabled=   0   'False
      Tab(2).Control(29)=   "txtFields(25)"
      Tab(2).Control(29).Enabled=   0   'False
      Tab(2).Control(30)=   "cmdpeg(6)"
      Tab(2).Control(30).Enabled=   0   'False
      Tab(2).Control(31)=   "txtFields(6)"
      Tab(2).Control(31).Enabled=   0   'False
      Tab(2).Control(32)=   "txtFields(8)"
      Tab(2).Control(32).Enabled=   0   'False
      Tab(2).Control(33)=   "cmdmais(8)"
      Tab(2).Control(33).Enabled=   0   'False
      Tab(2).Control(34)=   "cmdmenos(8)"
      Tab(2).Control(34).Enabled=   0   'False
      Tab(2).Control(35)=   "pfcescid(2)"
      Tab(2).Control(35).Enabled=   0   'False
      Tab(2).Control(36)=   "txtFields(7)"
      Tab(2).Control(36).Enabled=   0   'False
      Tab(2).Control(37)=   "cmdmais(7)"
      Tab(2).Control(37).Enabled=   0   'False
      Tab(2).Control(38)=   "cmdmenos(7)"
      Tab(2).Control(38).Enabled=   0   'False
      Tab(2).Control(39)=   "pfcescid(1)"
      Tab(2).Control(39).Enabled=   0   'False
      Tab(2).Control(40)=   "txtFields(9)"
      Tab(2).Control(40).Enabled=   0   'False
      Tab(2).Control(41)=   "cmdmais(9)"
      Tab(2).Control(41).Enabled=   0   'False
      Tab(2).Control(42)=   "cmdmenos(9)"
      Tab(2).Control(42).Enabled=   0   'False
      Tab(2).Control(43)=   "pfcescid(3)"
      Tab(2).Control(43).Enabled=   0   'False
      Tab(2).Control(44)=   "txtFields(10)"
      Tab(2).Control(44).Enabled=   0   'False
      Tab(2).Control(45)=   "txtFields(24)"
      Tab(2).Control(45).Enabled=   0   'False
      Tab(2).Control(46)=   "CMDSIG(0)"
      Tab(2).Control(46).Enabled=   0   'False
      Tab(2).Control(47)=   "CMDSIG(1)"
      Tab(2).Control(47).Enabled=   0   'False
      Tab(2).Control(48)=   "txtFields(42)"
      Tab(2).Control(48).Enabled=   0   'False
      Tab(2).Control(49)=   "CMDSIG(2)"
      Tab(2).Control(49).Enabled=   0   'False
      Tab(2).Control(50)=   "CMDSIG(3)"
      Tab(2).Control(50).Enabled=   0   'False
      Tab(2).Control(51)=   "txtFields(43)"
      Tab(2).Control(51).Enabled=   0   'False
      Tab(2).Control(52)=   "CMDSIG(4)"
      Tab(2).Control(52).Enabled=   0   'False
      Tab(2).Control(53)=   "CMDSIG(5)"
      Tab(2).Control(53).Enabled=   0   'False
      Tab(2).Control(54)=   "Check4"
      Tab(2).Control(54).Enabled=   0   'False
      Tab(2).Control(55)=   "Check2"
      Tab(2).Control(55).Enabled=   0   'False
      Tab(2).Control(56)=   "txtFields(34)"
      Tab(2).Control(56).Enabled=   0   'False
      Tab(2).Control(57)=   "txtFields(53)"
      Tab(2).Control(57).Enabled=   0   'False
      Tab(2).Control(58)=   "txtFields(54)"
      Tab(2).Control(58).Enabled=   0   'False
      Tab(2).Control(59)=   "txtFields(55)"
      Tab(2).Control(59).Enabled=   0   'False
      Tab(2).ControlCount=   60
      TabCaption(3)   =   "Tab 3"
      TabPicture(3)   =   "frmFEMEI.frx":05DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Tab 4"
      TabPicture(4)   =   "frmFEMEI.frx":05FA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "Passo 6"
      TabPicture(5)   =   "frmFEMEI.frx":0616
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "lblLabels(15)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "lblLabels(12)"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).Control(2)=   "lblLabels(13)"
      Tab(5).Control(2).Enabled=   0   'False
      Tab(5).Control(3)=   "Label1(2)"
      Tab(5).Control(3).Enabled=   0   'False
      Tab(5).Control(4)=   "lblLabels(18)"
      Tab(5).Control(4).Enabled=   0   'False
      Tab(5).Control(5)=   "lblLabels(17)"
      Tab(5).Control(5).Enabled=   0   'False
      Tab(5).Control(6)=   "lblLabels(19)"
      Tab(5).Control(6).Enabled=   0   'False
      Tab(5).Control(7)=   "lblLabels(20)"
      Tab(5).Control(7).Enabled=   0   'False
      Tab(5).Control(8)=   "lblLabels(11)"
      Tab(5).Control(8).Enabled=   0   'False
      Tab(5).Control(9)=   "lblLabels(30)"
      Tab(5).Control(9).Enabled=   0   'False
      Tab(5).Control(10)=   "lblLabels(31)"
      Tab(5).Control(10).Enabled=   0   'False
      Tab(5).Control(11)=   "lblLabels(32)"
      Tab(5).Control(11).Enabled=   0   'False
      Tab(5).Control(12)=   "Label1(3)"
      Tab(5).Control(12).Enabled=   0   'False
      Tab(5).Control(13)=   "Check1"
      Tab(5).Control(13).Enabled=   0   'False
      Tab(5).Control(14)=   "txtFields(11)"
      Tab(5).Control(14).Enabled=   0   'False
      Tab(5).Control(15)=   "Command2"
      Tab(5).Control(15).Enabled=   0   'False
      Tab(5).Control(16)=   "Command3"
      Tab(5).Control(16).Enabled=   0   'False
      Tab(5).Control(17)=   "txtFields(15)"
      Tab(5).Control(17).Enabled=   0   'False
      Tab(5).Control(18)=   "txtFields(12)"
      Tab(5).Control(18).Enabled=   0   'False
      Tab(5).Control(19)=   "txtFields(14)"
      Tab(5).Control(19).Enabled=   0   'False
      Tab(5).Control(20)=   "txtFields(16)"
      Tab(5).Control(20).Enabled=   0   'False
      Tab(5).Control(21)=   "Command1(0)"
      Tab(5).Control(21).Enabled=   0   'False
      Tab(5).Control(22)=   "Command1(1)"
      Tab(5).Control(22).Enabled=   0   'False
      Tab(5).Control(23)=   "Command1(2)"
      Tab(5).Control(23).Enabled=   0   'False
      Tab(5).Control(24)=   "Command1(10)"
      Tab(5).Control(24).Enabled=   0   'False
      Tab(5).Control(25)=   "Command1(11)"
      Tab(5).Control(25).Enabled=   0   'False
      Tab(5).Control(26)=   "Command1(12)"
      Tab(5).Control(26).Enabled=   0   'False
      Tab(5).Control(27)=   "Command4(0)"
      Tab(5).Control(27).Enabled=   0   'False
      Tab(5).Control(28)=   "Command4(1)"
      Tab(5).Control(28).Enabled=   0   'False
      Tab(5).Control(29)=   "Command4(2)"
      Tab(5).Control(29).Enabled=   0   'False
      Tab(5).Control(30)=   "Command1(3)"
      Tab(5).Control(30).Enabled=   0   'False
      Tab(5).Control(31)=   "Command1(4)"
      Tab(5).Control(31).Enabled=   0   'False
      Tab(5).Control(32)=   "Command1(5)"
      Tab(5).Control(32).Enabled=   0   'False
      Tab(5).Control(33)=   "txtFields(13)"
      Tab(5).Control(33).Enabled=   0   'False
      Tab(5).Control(34)=   "txtFields(21)"
      Tab(5).Control(34).Enabled=   0   'False
      Tab(5).Control(35)=   "txtFields(22)"
      Tab(5).Control(35).Enabled=   0   'False
      Tab(5).Control(36)=   "txtFields(18)"
      Tab(5).Control(36).Enabled=   0   'False
      Tab(5).Control(37)=   "cmdmais(18)"
      Tab(5).Control(37).Enabled=   0   'False
      Tab(5).Control(38)=   "cmdmenos(18)"
      Tab(5).Control(38).Enabled=   0   'False
      Tab(5).Control(39)=   "pfcescid(5)"
      Tab(5).Control(39).Enabled=   0   'False
      Tab(5).Control(40)=   "txtFields(17)"
      Tab(5).Control(40).Enabled=   0   'False
      Tab(5).Control(41)=   "cmdmais(17)"
      Tab(5).Control(41).Enabled=   0   'False
      Tab(5).Control(42)=   "cmdmenos(17)"
      Tab(5).Control(42).Enabled=   0   'False
      Tab(5).Control(43)=   "pfcescid(4)"
      Tab(5).Control(43).Enabled=   0   'False
      Tab(5).Control(44)=   "txtFields(19)"
      Tab(5).Control(44).Enabled=   0   'False
      Tab(5).Control(45)=   "cmdmais(19)"
      Tab(5).Control(45).Enabled=   0   'False
      Tab(5).Control(46)=   "cmdmenos(19)"
      Tab(5).Control(46).Enabled=   0   'False
      Tab(5).Control(47)=   "pfcescid(6)"
      Tab(5).Control(47).Enabled=   0   'False
      Tab(5).Control(48)=   "txtFields(20)"
      Tab(5).Control(48).Enabled=   0   'False
      Tab(5).Control(49)=   "txtFields(23)"
      Tab(5).Control(49).Enabled=   0   'False
      Tab(5).Control(50)=   "Cmdsitu(1)"
      Tab(5).Control(50).Enabled=   0   'False
      Tab(5).Control(51)=   "Cmdsitu(2)"
      Tab(5).Control(51).Enabled=   0   'False
      Tab(5).Control(52)=   "Cmdsitu(0)"
      Tab(5).Control(52).Enabled=   0   'False
      Tab(5).Control(53)=   "txtFields(58)"
      Tab(5).Control(53).Enabled=   0   'False
      Tab(5).Control(54)=   "txtFields(56)"
      Tab(5).Control(54).Enabled=   0   'False
      Tab(5).Control(55)=   "txtFields(57)"
      Tab(5).Control(55).Enabled=   0   'False
      Tab(5).Control(56)=   "txtFields(59)"
      Tab(5).Control(56).Enabled=   0   'False
      Tab(5).Control(57)=   "DTPicker1"
      Tab(5).Control(57).Enabled=   0   'False
      Tab(5).Control(58)=   "DTPicker2"
      Tab(5).Control(58).Enabled=   0   'False
      Tab(5).Control(59)=   "DTPicker3"
      Tab(5).Control(59).Enabled=   0   'False
      Tab(5).Control(60)=   "DTPicker4"
      Tab(5).Control(60).Enabled=   0   'False
      Tab(5).ControlCount=   61
      Begin VBCCR17.DTPicker DTPicker4 
         Height          =   375
         Left            =   -66360
         TabIndex        =   157
         Top             =   4680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Value           =   44285
         CheckBox        =   -1  'True
         AllowUserInput  =   -1  'True
      End
      Begin VBCCR17.DTPicker DTPicker3 
         Height          =   375
         Left            =   -66360
         TabIndex        =   156
         Top             =   4320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Value           =   44285
         CheckBox        =   -1  'True
         AllowUserInput  =   -1  'True
      End
      Begin VBCCR17.DTPicker DTPicker2 
         Height          =   375
         Left            =   -66360
         TabIndex        =   155
         Top             =   3960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Value           =   44285
         CheckBox        =   -1  'True
         AllowUserInput  =   -1  'True
      End
      Begin VBCCR17.DTPicker DTPicker1 
         Height          =   375
         Left            =   -66120
         TabIndex        =   154
         Top             =   3000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         Value           =   44285
         CheckBox        =   -1  'True
         AllowUserInput  =   -1  'True
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   55
         Left            =   -72720
         TabIndex        =   17
         Top             =   4920
         Width           =   735
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   54
         Left            =   -72720
         TabIndex        =   16
         Top             =   4560
         Width           =   735
      End
      Begin VB.TextBox txtFields 
         Height          =   495
         Index           =   59
         Left            =   -73680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   6240
         Width           =   8175
      End
      Begin VB.TextBox txtFields 
         Height          =   495
         Index           =   57
         Left            =   -73440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   1440
         Width           =   8175
      End
      Begin VB.TextBox txtFields 
         Height          =   495
         Index           =   56
         Left            =   -73440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   840
         Width           =   8175
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFFF&
         Height          =   405
         Index           =   58
         Left            =   -68640
         MaxLength       =   1
         TabIndex        =   148
         TabStop         =   0   'False
         Top             =   5280
         Width           =   495
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFFF&
         Height          =   405
         Index           =   53
         Left            =   -69360
         MaxLength       =   1
         TabIndex        =   146
         TabStop         =   0   'False
         Top             =   4080
         Width           =   495
      End
      Begin VB.TextBox txtFields 
         Height          =   405
         Index           =   52
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   3120
         Width           =   8175
      End
      Begin VB.TextBox txtFields 
         Height          =   405
         Index           =   51
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2400
         Width           =   8175
      End
      Begin VB.TextBox txtFields 
         Height          =   405
         Index           =   50
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1680
         Width           =   8175
      End
      Begin VB.TextBox txtFields 
         Height          =   405
         Index           =   49
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   2640
         Width           =   8175
      End
      Begin VB.TextBox txtFields 
         Height          =   405
         Index           =   48
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1920
         Width           =   8175
      End
      Begin VB.TextBox txtFields 
         Height          =   405
         Index           =   47
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   1200
         Width           =   8175
      End
      Begin VB.CommandButton Cmdsitu 
         Caption         =   "CEP"
         Height          =   255
         Index           =   0
         Left            =   -72000
         TabIndex        =   138
         TabStop         =   0   'False
         Top             =   5880
         Width           =   495
      End
      Begin VB.CommandButton Cmdsitu 
         Caption         =   "Esp"
         Height          =   255
         Index           =   2
         Left            =   -72600
         TabIndex        =   137
         TabStop         =   0   'False
         Top             =   5880
         Width           =   495
      End
      Begin VB.CommandButton Cmdsitu 
         Caption         =   "Nor"
         Height          =   255
         Index           =   1
         Left            =   -73200
         TabIndex        =   136
         TabStop         =   0   'False
         Top             =   5880
         Width           =   495
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   23
         Left            =   -73560
         Locked          =   -1  'True
         TabIndex        =   135
         TabStop         =   0   'False
         Top             =   5880
         Width           =   255
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   20
         Left            =   -69720
         Locked          =   -1  'True
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   5280
         Width           =   855
      End
      Begin VB.CommandButton pfcescid 
         Height          =   240
         Index           =   6
         Left            =   -70080
         Picture         =   "frmFEMEI.frx":0632
         Style           =   1  'Graphical
         TabIndex        =   131
         TabStop         =   0   'False
         Top             =   5280
         Width           =   255
      End
      Begin VB.CommandButton cmdmenos 
         Caption         =   "-"
         Height          =   240
         Index           =   19
         Left            =   -70320
         TabIndex        =   130
         TabStop         =   0   'False
         ToolTipText     =   "Diminui um mes"
         Top             =   5520
         Width           =   210
      End
      Begin VB.CommandButton cmdmais 
         Caption         =   "+"
         Height          =   240
         Index           =   19
         Left            =   -70320
         TabIndex        =   129
         TabStop         =   0   'False
         ToolTipText     =   "Aumenta um mes"
         Top             =   5280
         Width           =   210
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   19
         Left            =   -71160
         TabIndex        =   31
         Top             =   5400
         Width           =   855
      End
      Begin VB.CommandButton pfcescid 
         Height          =   240
         Index           =   4
         Left            =   -71520
         Picture         =   "frmFEMEI.frx":0BBC
         Style           =   1  'Graphical
         TabIndex        =   127
         TabStop         =   0   'False
         Top             =   5280
         Width           =   255
      End
      Begin VB.CommandButton cmdmenos 
         Caption         =   "-"
         Height          =   240
         Index           =   17
         Left            =   -71760
         TabIndex        =   126
         TabStop         =   0   'False
         ToolTipText     =   "Diminui um mes"
         Top             =   5520
         Width           =   210
      End
      Begin VB.CommandButton cmdmais 
         Caption         =   "+"
         Height          =   240
         Index           =   17
         Left            =   -71760
         TabIndex        =   125
         TabStop         =   0   'False
         ToolTipText     =   "Aumenta um mes"
         Top             =   5280
         Width           =   210
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   17
         Left            =   -72600
         TabIndex        =   30
         Top             =   5400
         Width           =   855
      End
      Begin VB.CommandButton pfcescid 
         Height          =   240
         Index           =   5
         Left            =   -72960
         Picture         =   "frmFEMEI.frx":1146
         Style           =   1  'Graphical
         TabIndex        =   123
         TabStop         =   0   'False
         Top             =   5280
         Width           =   255
      End
      Begin VB.CommandButton cmdmenos 
         Caption         =   "-"
         Height          =   240
         Index           =   18
         Left            =   -73200
         TabIndex        =   122
         TabStop         =   0   'False
         ToolTipText     =   "Diminui um mes"
         Top             =   5520
         Width           =   210
      End
      Begin VB.CommandButton cmdmais 
         Caption         =   "+"
         Height          =   240
         Index           =   18
         Left            =   -73200
         TabIndex        =   121
         TabStop         =   0   'False
         ToolTipText     =   "Aumenta um mes"
         Top             =   5280
         Width           =   210
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   18
         Left            =   -74040
         TabIndex        =   29
         Top             =   5400
         Width           =   855
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   22
         Left            =   -71760
         MaxLength       =   40
         TabIndex        =   28
         Top             =   4680
         Width           =   5235
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   21
         Left            =   -71760
         MaxLength       =   40
         TabIndex        =   26
         Top             =   4320
         Width           =   5235
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   13
         Left            =   -71760
         MaxLength       =   40
         TabIndex        =   24
         Top             =   3960
         Width           =   5235
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..A"
         Height          =   255
         Index           =   5
         Left            =   -72240
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   4680
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..A"
         Height          =   255
         Index           =   4
         Left            =   -72240
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..A"
         Height          =   255
         Index           =   3
         Left            =   -72240
         TabIndex        =   115
         TabStop         =   0   'False
         Top             =   3960
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Caption         =   "-->"
         Height          =   255
         Index           =   2
         Left            =   -72720
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   4680
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Caption         =   "-->"
         Height          =   255
         Index           =   1
         Left            =   -72720
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Caption         =   "-->"
         Height          =   255
         Index           =   0
         Left            =   -72720
         TabIndex        =   112
         TabStop         =   0   'False
         Top             =   3960
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "LX"
         Height          =   255
         Index           =   12
         Left            =   -73200
         Picture         =   "frmFEMEI.frx":16D0
         Style           =   1  'Graphical
         TabIndex        =   111
         TabStop         =   0   'False
         Top             =   4680
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "LX"
         Height          =   255
         Index           =   11
         Left            =   -73200
         Picture         =   "frmFEMEI.frx":1C5A
         Style           =   1  'Graphical
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "LX"
         Height          =   255
         Index           =   10
         Left            =   -73200
         Picture         =   "frmFEMEI.frx":21E4
         Style           =   1  'Graphical
         TabIndex        =   109
         TabStop         =   0   'False
         Top             =   3960
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   255
         Index           =   2
         Left            =   -73680
         Picture         =   "frmFEMEI.frx":276E
         Style           =   1  'Graphical
         TabIndex        =   108
         TabStop         =   0   'False
         Top             =   4680
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   255
         Index           =   1
         Left            =   -73680
         Picture         =   "frmFEMEI.frx":2CF8
         Style           =   1  'Graphical
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   4320
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   255
         Index           =   0
         Left            =   -73680
         Picture         =   "frmFEMEI.frx":3282
         Style           =   1  'Graphical
         TabIndex        =   106
         TabStop         =   0   'False
         Top             =   3960
         Width           =   375
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   16
         Left            =   -74760
         TabIndex        =   27
         Top             =   4680
         Width           =   975
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   14
         Left            =   -74760
         TabIndex        =   25
         Top             =   4320
         Width           =   975
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   12
         Left            =   -74760
         TabIndex        =   23
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   34
         Left            =   -68880
         TabIndex        =   18
         Top             =   5640
         Width           =   2295
      End
      Begin VB.TextBox txtFields 
         Height          =   765
         Index           =   15
         Left            =   -73440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   2880
         Width           =   7215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Não"
         Height          =   255
         Left            =   -74160
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   2280
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Sim"
         Height          =   255
         Left            =   -74760
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox txtFields 
         Height          =   735
         Index           =   11
         Left            =   -73440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   2040
         Width           =   8175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Ação Cor."
         Enabled         =   0   'False
         Height          =   195
         Left            =   -74760
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Desconsiderar Apuraçao RPN"
         Height          =   255
         Left            =   -72480
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   5640
         Width           =   2715
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Modificacao Padrao"
         Height          =   255
         Left            =   -74640
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   5640
         Width           =   1755
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   375
         Index           =   5
         Left            =   -65520
         Picture         =   "frmFEMEI.frx":380C
         Style           =   1  'Graphical
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   4200
         Width           =   375
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   375
         Index           =   4
         Left            =   -65520
         Picture         =   "frmFEMEI.frx":3D96
         Style           =   1  'Graphical
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   3840
         Width           =   375
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   43
         Left            =   -66120
         MaxLength       =   1
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   4200
         Width           =   495
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   375
         Index           =   3
         Left            =   -66600
         Picture         =   "frmFEMEI.frx":4320
         Style           =   1  'Graphical
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   4200
         Width           =   375
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   375
         Index           =   2
         Left            =   -66600
         Picture         =   "frmFEMEI.frx":48AA
         Style           =   1  'Graphical
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   3840
         Width           =   375
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   42
         Left            =   -67080
         MaxLength       =   1
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   4200
         Width           =   495
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   375
         Index           =   1
         Left            =   -67560
         Picture         =   "frmFEMEI.frx":4E34
         Style           =   1  'Graphical
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   4200
         Width           =   375
      End
      Begin VB.CommandButton CMDSIG 
         Height          =   375
         Index           =   0
         Left            =   -67560
         Picture         =   "frmFEMEI.frx":53BE
         Style           =   1  'Graphical
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   3840
         Width           =   375
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   24
         Left            =   -68160
         MaxLength       =   1
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   4200
         Width           =   495
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   -70320
         Locked          =   -1  'True
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   4080
         Width           =   855
      End
      Begin VB.CommandButton pfcescid 
         Height          =   240
         HelpContextID   =   3
         Index           =   3
         Left            =   -70680
         Picture         =   "frmFEMEI.frx":5948
         Style           =   1  'Graphical
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   3840
         Width           =   255
      End
      Begin VB.CommandButton cmdmenos 
         Caption         =   "-"
         Height          =   240
         Index           =   9
         Left            =   -70920
         TabIndex        =   79
         TabStop         =   0   'False
         ToolTipText     =   "Diminui um mes"
         Top             =   4080
         Width           =   210
      End
      Begin VB.CommandButton cmdmais 
         Caption         =   "+"
         Height          =   240
         Index           =   9
         Left            =   -70920
         TabIndex        =   78
         TabStop         =   0   'False
         ToolTipText     =   "Aumenta um mes"
         Top             =   3840
         Width           =   210
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   9
         Left            =   -71760
         TabIndex        =   15
         Top             =   4080
         Width           =   855
      End
      Begin VB.CommandButton pfcescid 
         Height          =   240
         Index           =   1
         Left            =   -72120
         Picture         =   "frmFEMEI.frx":5ED2
         Style           =   1  'Graphical
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   3840
         Width           =   255
      End
      Begin VB.CommandButton cmdmenos 
         Caption         =   "-"
         Height          =   240
         Index           =   7
         Left            =   -72360
         TabIndex        =   75
         TabStop         =   0   'False
         ToolTipText     =   "Diminui um mes"
         Top             =   4080
         Width           =   210
      End
      Begin VB.CommandButton cmdmais 
         Caption         =   "+"
         Height          =   240
         Index           =   7
         Left            =   -72360
         TabIndex        =   74
         TabStop         =   0   'False
         ToolTipText     =   "Aumenta um mes"
         Top             =   3840
         Width           =   210
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   7
         Left            =   -73200
         TabIndex        =   14
         Top             =   4080
         Width           =   855
      End
      Begin VB.CommandButton pfcescid 
         Height          =   240
         Index           =   2
         Left            =   -73560
         Picture         =   "frmFEMEI.frx":645C
         Style           =   1  'Graphical
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   3840
         Width           =   255
      End
      Begin VB.CommandButton cmdmenos 
         Caption         =   "-"
         Height          =   240
         Index           =   8
         Left            =   -73800
         TabIndex        =   71
         TabStop         =   0   'False
         ToolTipText     =   "Diminui um mes"
         Top             =   4080
         Width           =   210
      End
      Begin VB.CommandButton cmdmais 
         Caption         =   "+"
         Height          =   240
         Index           =   8
         Left            =   -73800
         TabIndex        =   70
         TabStop         =   0   'False
         ToolTipText     =   "Aumenta um mes"
         Top             =   3840
         Width           =   210
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   8
         Left            =   -74640
         TabIndex        =   13
         Top             =   4080
         Width           =   855
      End
      Begin VB.TextBox txtFields 
         Height          =   375
         Index           =   6
         Left            =   -73320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   3120
         Width           =   8175
      End
      Begin VB.CommandButton cmdpeg 
         Height          =   240
         Index           =   6
         Left            =   -73680
         Picture         =   "frmFEMEI.frx":69E6
         Style           =   1  'Graphical
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   3120
         Width           =   255
      End
      Begin VB.TextBox txtFields 
         Height          =   375
         Index           =   25
         Left            =   -73320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   2640
         Width           =   8175
      End
      Begin VB.CommandButton cmdpeg 
         Height          =   240
         Index           =   25
         Left            =   -73680
         Picture         =   "frmFEMEI.frx":6F70
         Style           =   1  'Graphical
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   2640
         Width           =   255
      End
      Begin VB.TextBox txtFields 
         Height          =   405
         Index           =   5
         Left            =   -73320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2160
         Width           =   8175
      End
      Begin VB.CommandButton cmdpegpro 
         Height          =   240
         Index           =   3
         Left            =   -73680
         Picture         =   "frmFEMEI.frx":74FA
         Style           =   1  'Graphical
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   2160
         Width           =   255
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   41
         Left            =   -74280
         Locked          =   -1  'True
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox txtFields 
         Height          =   375
         Index           =   4
         Left            =   -73320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   1680
         Width           =   8175
      End
      Begin VB.CommandButton cmdpegpro 
         Height          =   240
         Index           =   2
         Left            =   -73680
         Picture         =   "frmFEMEI.frx":7A84
         Style           =   1  'Graphical
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   1800
         Width           =   255
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   40
         Left            =   -74280
         Locked          =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1800
         Width           =   495
      End
      Begin VB.TextBox txtFields 
         Height          =   405
         Index           =   3
         Left            =   -73320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   1200
         Width           =   8175
      End
      Begin VB.CommandButton cmdpegpro 
         Height          =   240
         Index           =   1
         Left            =   -73680
         Picture         =   "frmFEMEI.frx":800E
         Style           =   1  'Graphical
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1320
         Width           =   255
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   39
         Left            =   -74280
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtFields 
         Height          =   405
         Index           =   2
         Left            =   -73320
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   720
         Width           =   8175
      End
      Begin VB.CommandButton cmdpegpro 
         Height          =   240
         Index           =   0
         Left            =   -73680
         Picture         =   "frmFEMEI.frx":8598
         Style           =   1  'Graphical
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   38
         Left            =   -74280
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblLabels 
         Caption         =   "Codigo do Filtro"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   34
         Left            =   -74640
         TabIndex        =   153
         Top             =   4920
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Caption         =   "Caracteristica Especial"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   33
         Left            =   -74640
         TabIndex        =   152
         Top             =   4560
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Observacao"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   3
         Left            =   -74760
         TabIndex        =   151
         Top             =   6240
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Ação Deteccao"
         ForeColor       =   &H00C00000&
         Height          =   495
         Index           =   32
         Left            =   -74760
         TabIndex        =   150
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Ação Prevenção"
         ForeColor       =   &H00C00000&
         Height          =   495
         Index           =   31
         Left            =   -74760
         TabIndex        =   149
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   "PA"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   30
         Left            =   -68640
         TabIndex        =   147
         Top             =   5040
         Width           =   495
      End
      Begin VB.Label lblLabels 
         Caption         =   "PA"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   29
         Left            =   -69360
         TabIndex        =   145
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label lblLabels 
         Caption         =   "Funcao Elemento do Trabalho"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   28
         Left            =   -74760
         TabIndex        =   144
         Top             =   2880
         Width           =   4095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Funcao Etapa Processo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   27
         Left            =   -74760
         TabIndex        =   143
         Top             =   2160
         Width           =   4095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Funcao Item Processo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   26
         Left            =   -74760
         TabIndex        =   142
         Top             =   1320
         Width           =   4095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Estacao do Trabalho"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   25
         Left            =   240
         TabIndex        =   141
         Top             =   2400
         Width           =   4095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Estacao Nome do Elemento"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   24
         Left            =   240
         TabIndex        =   140
         Top             =   1680
         Width           =   4095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Processo Elemento Sistema"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   23
         Left            =   240
         TabIndex        =   139
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Situacao"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   11
         Left            =   -74400
         TabIndex        =   134
         Top             =   5880
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Risco"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   20
         Left            =   -69720
         TabIndex        =   132
         Top             =   5040
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Deteçao"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   19
         Left            =   -71160
         TabIndex        =   128
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Ocorrencia"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   17
         Left            =   -72600
         TabIndex        =   124
         Top             =   5160
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Severidade"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   18
         Left            =   -74040
         TabIndex        =   120
         Top             =   5160
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Indices"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   119
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nome"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   13
         Left            =   -71760
         TabIndex        =   118
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Responsável"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   12
         Left            =   -74760
         TabIndex        =   105
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   "PSA"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   21
         Left            =   -69480
         TabIndex        =   104
         Top             =   5640
         Width           =   495
      End
      Begin VB.Label lblLabels 
         Caption         =   "Ação Tomada"
         ForeColor       =   &H00C00000&
         Height          =   495
         Index           =   15
         Left            =   -74760
         TabIndex        =   103
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Simbologia"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   14
         Left            =   -68520
         TabIndex        =   88
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Risco"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   10
         Left            =   -70320
         TabIndex        =   81
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Deteçao"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   9
         Left            =   -71760
         TabIndex        =   77
         Top             =   3840
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Ocorrencia"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   7
         Left            =   -73200
         TabIndex        =   73
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Severidade"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   8
         Left            =   -74640
         TabIndex        =   69
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Indices"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   68
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Controle Atual"
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   6
         Left            =   -74880
         TabIndex        =   65
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Crtl Aut Prev"
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   16
         Left            =   -74880
         TabIndex        =   64
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Causa"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   5
         Left            =   -74880
         TabIndex        =   61
         Top             =   2280
         Width           =   495
      End
      Begin VB.Label lblLabels 
         Caption         =   "Efeito"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   4
         Left            =   -74880
         TabIndex        =   58
         Top             =   1800
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Falhas"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   55
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblLabels 
         Caption         =   "Proc."
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   52
         Top             =   840
         Width           =   495
      End
   End
   Begin XPControls.XPButton cmdClose 
      Height          =   435
      Left            =   11040
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      Picture         =   "frmFEMEI.frx":8B22
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
      Left            =   11040
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      Picture         =   "frmFEMEI.frx":90BC
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
   Begin VB.Label lblLabels 
      Caption         =   "RevFemea"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   22
      Left            =   7440
      TabIndex        =   86
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "SubTipo"
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   3
      Left            =   9240
      TabIndex        =   83
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "Item:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   35
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Caption         =   "PF:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   33
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmFEMEI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' PF/PFP/PFG FEMEA FEMPRE
' FEMEA FEMEA FEMAVU FEMADC FEMREVI FEMCAU FEMCAUREV FEMHIS FEMMAQ FEMRPNT
'
Option Explicit
Private cARQUSO As String
Private cSQLUSO As String
Private cARQPF As String
Private cARQFEMEA As String
Private cARQFEMEAPF As String
Private aCAM
Private aFOR
Private aVAL
Private aPAD As Variant
Private nITEM
Private nCAMPOS As Long

Private Sub CmdAltman_Click()

    Check3.Value = Checked
    EnableControls Me, True
    acaocampos
    ''sEMPRE DESABILITADOS
    Check3.Enabled = False                       ''ALTMAN
    Check1.Enabled = False                       ''ACAO
    TXTFIELDS(23).Enabled = False                ''SITUACAO
    TXTFIELDS(10).Enabled = False                ''RISCO
    TXTFIELDS(20).Enabled = False                ''RISCO
    TXTFIELDS(0).Enabled = False                 ''PF
    TXTFIELDS(1).Enabled = False                 ''ITEM

End Sub

Private Sub cmdClose_Click()
    If TXTFIELDS(38) = 0 Or TXTFIELDS(38) = 0 Or TXTFIELDS(38) = 0 Or TXTFIELDS(38) = 0 Then
        Alert "Processo Falha Efeito Causa nao selecionados"
    End If

    If MDG("Gravar e Sair", "Gravando ") Then
        For nITEM = 0 To 25
            aVAL(nITEM) = TXTFIELDS(nITEM)
        Next nITEM
        aVAL(26) = DTPicker1.Value
        aVAL(27) = DTPicker2.Value
        aVAL(28) = DTPicker3.Value
        aVAL(29) = DTPicker4.Value
        If Len(aVAL(15)) = 0 Then                ''Sem Acao zera Data
            aVAL(29) = ""
        End If
        If FixInt(aVAL(12)) = 0 Then
            aVAL(26) = ""
        End If
        If FixInt(aVAL(14)) = 0 Then
            aVAL(27) = ""
        End If
        If FixInt(aVAL(16)) = 0 Then
            aVAL(28) = ""
        End If
        aVAL(30) = Check1.Value
        aVAL(31) = Check2.Value
        aVAL(32) = Check3.Value
        aVAL(33) = Check4.Value
        For nITEM = 34 To 59 '46
            aVAL(nITEM) = TXTFIELDS(nITEM)
        Next
        GrvSQL cARQUSO, cSQLUSO, nCAMPOS, aCAM, aVAL, aFOR
        GravaLog 0, 0, "Gravou "
    Else
        GravaLog 0, 0, "Nao Gravou "
    End If
    Screen.MousePointer = vbDefault
    Unload Me

End Sub

Private Sub cmddizsessq_Click(Index As Integer)

    Dim cDIZ
    Dim cARQ As String
    cARQ = PegPath("PATH", "PF")

    cDIZ = pegdizseqssq(TXTFIELDS(2), cARQ, FixInt(TXTFIELDS(35).tEXT), FixInt(TXTFIELDS(36).tEXT), FixInt(TXTFIELDS(37).tEXT), FixInt(TXTFIELDS(45).tEXT))
    
    If Index = 0 Then
        TXTFIELDS(2).tEXT = eRETU01(5)
    End If
    If Index = 1 Then
        TXTFIELDS(2).tEXT = eRETU01(4)
    End If

    
    TXTFIELDS(35).tEXT = eRETU01(0)
    TXTFIELDS(36).tEXT = eRETU01(1)
    TXTFIELDS(37).tEXT = eRETU01(2)
    TXTFIELDS(45).tEXT = eRETU01(3)
    

End Sub

Private Sub cmdmais_Click(Index As Integer)

    TXTFIELDS(Index) = CStr(IncDec(TXTFIELDS(Index), 1, 1, 10))

End Sub

Private Sub cmdmenos_Click(Index As Integer)

    TXTFIELDS(Index) = CStr(IncDec(TXTFIELDS(Index), -1, 1, 10))

End Sub

Private Sub cmdpeg_Click(Index As Integer)

    Select Case Index
    Case 25
        iMU01 = 304
    Case 6
        iMU01 = 305
    Case 44
        iMU01 = 306
    End Select
    eRETU02 = ""                                 'Evita Cancelamento Escolha
    escIED.Show vbModal, Me
    If lRETU Then
        If Index = 44 Then
            TXTFIELDS(Index) = Mid(eRETU02, 1, 1)
        Else
            TXTFIELDS(Index) = eRETU02
        End If
    End If

End Sub

Private Sub cmdpegpro_Click(Index As Integer)
    Dim sSQL
    Dim lATU
    If Index > 0 And Val(TXTFIELDS(38)) = 0 Then
        Alert ("Escolha um Processo")
        Exit Sub
    End If
    If Index > 1 And Val(TXTFIELDS(39)) = 0 Then
        Alert ("Escolha uma Falha")
        Exit Sub
    End If
    If Index > 2 And Val(TXTFIELDS(40)) = 0 Then
        Alert ("Escolha um Efeito")
        Exit Sub
    End If
    ePASS01 = "FEMEA"
    If Index = 0 Then                            'PROCESSO
        ePASS02 = "SELECT pronum AS CODIGO,processo AS NOME FROM FEMPRO"
        ePASS03 = Array("PRONUM", "PROCESSO")
    End If
    If Index = 1 Then                            'FALHA
        ePASS02 = "SELECT FALNUM AS CODIGO, FALTIP AS NOME FROM FEMFAL WHERE PRONUM=" & FixInt(TXTFIELDS(38))
        ePASS03 = Array("FALNUM", "FALTIP")
    End If
    If Index = 2 Then                            'EFEITO
        ePASS02 = "SELECT EFENUM AS CODIGO, FALEFE AS NOME FROM FEMEFE WHERE PRONUM=" & FixInt(TXTFIELDS(38)) & " AND FALNUM=" & FixInt(TXTFIELDS(39))
        ePASS03 = Array("EFENUM", "FALEFE")
    End If
    If Index = 3 Then                            'CAUSA
        ePASS02 = "SELECT cAUNUM as codigo,FALCAU as nome FROM FEMCAU WHERE PRONUM=" & FixInt(TXTFIELDS(38)) & " AND FALNUM=" & FixInt(TXTFIELDS(39)) & " AND EFEnum=" & FixInt(TXTFIELDS(40))
        ePASS03 = Array("CAUNUM", "FALCAU")
    End If
    EscCodNomSim.Show vbModal, Me
    If lRETU Then
        lATU = MDG("Atualizar descritivo")
        If Index = 0 Then
            TXTFIELDS(38) = eRETU01
            If Len(TXTFIELDS(2)) = 0 Or lATU Then
                TXTFIELDS(2) = eRETU02
            End If
            If MDG("Apagar Falha Efeito Causa") Then
                TXTFIELDS(39) = 0
                TXTFIELDS(3) = ""
                TXTFIELDS(40) = 0
                TXTFIELDS(4) = ""
                TXTFIELDS(41) = 0
                TXTFIELDS(5) = ""
            End If
        End If
        If Index = 1 Then
            TXTFIELDS(39) = eRETU01
            If Len(TXTFIELDS(3)) = 0 Or lATU Then
                TXTFIELDS(3) = eRETU02
            End If
            If MDG("Apagar Efeito Causa") Then
                TXTFIELDS(40) = 0
                TXTFIELDS(4) = ""
                TXTFIELDS(41) = 0
                TXTFIELDS(5) = ""
            End If
        End If
        If Index = 2 Then
            TXTFIELDS(40) = eRETU01
            If Len(TXTFIELDS(4)) = 0 Or lATU Then
                TXTFIELDS(4) = eRETU02
            End If
            If MDG("Apagar Causa") Then
                TXTFIELDS(41) = 0
                TXTFIELDS(5) = ""
            End If
        End If
        If Index = 3 Then
            TXTFIELDS(41) = eRETU01
            If Len(TXTFIELDS(5)) = 0 Or lATU Then
                TXTFIELDS(5) = eRETU02
            End If
            sSQL = "select * from FEMCAU WHERE PRONUM=" & TXTFIELDS(38) & " AND FALNUM=" & TXTFIELDS(39) & " AND EFENUM=" & TXTFIELDS(40) & " AND CAUNUM=" & TXTFIELDS(41)
            aVAL = PegSQL(cARQFEMEA, sSQL, nCAMPOS, aCAM, aFOR, aPAD)
            gravaAvalTxt (3)
            TXTFIELDS_Change (7)
            TXTFIELDS_Change (17)
        End If
    End If
End Sub

Private Sub CmdRetornaAcao_Click(Index As Integer)
    Dim aCAMUSO As Variant
    Dim aPADUSO As Variant
    Dim aFORUSO As Variant
    Dim aRETU   As Variant
    Dim sSQL    As String
    Dim cREVFEM As String
    Dim nCAMUSO
    Dim nTMPGRA
    nCAMUSO = 17
    If Index = 0 Then
        cREVFEM = TXTFIELDS(46).tEXT
        cREVFEM = Busca("Qual Revisao Femeao", "Retornar Revisao Femea", cREVFEM, 3)
        sSQL = "select * from FEMrevi WHERE  TIPOAPU='N' AND PF=" & nPF & " AND FEMEAREV=" & FixNum(cREVFEM)
        sSQL = sSQL & " AND FXSEQ=" & TXTFIELDS(35)
        sSQL = sSQL & " AND FXSSQ=" & TXTFIELDS(36)
        sSQL = sSQL & " AND FXITEM=" & TXTFIELDS(37)
        sSQL = sSQL & " AND FXITEMS=" & TXTFIELDS(45)
    End If
    If Index = 1 Then
        nTMPGRA = 0
        nTMPGRA = FixInt(Val(Busca("Qual sequencia", "Retornar Sequencia Revisao", CStr(nTMPGRA), 8)))
        sSQL = "select * from FEMrevi WHERE  TIPOAPU='N' AND PF=" & nPF & " AND ITEM=" & TXTFIELDS(1)
        sSQL = sSQL & " AND SEGGRA=" & nTMPGRA
    End If

    If Not MDG("Retornar versao") Then
        Exit Sub
    End If

    aCAMUSO = Array("ACAREC", "RESCOD", "RESNOM", "RESCOD2", "ACATOM", _
                    "RESCOD3", "RINDOCO", "RINDSER", "RINDDET", "RINDRIS", _
                    "RESNOM2", "RESNOM3", "INDOCO", "INDSEV", "INDDET", _
                    "INDRIS", "ACAO")
    aPADUSO = Array("", 0, "", 0, "", 0, 0, 0, 0, 0, "", "", 0, 0, 0, 0, False)
    aFORUSO = Array("C", "NI", "C", "NI", "C", "NI", "NI", "NI", "NI", "NI", "C", "C", "NI", "NI", "NI", "NI", "BN")
    aRETU = PegSQL(cARQFEMEA, sSQL, nCAMUSO, aCAMUSO, aFORUSO, aPADUSO)
    If Not lRETU Then
        Alert ("Versao anterior nao encontrada")
        Exit Sub
    End If

    TXTFIELDS(11) = aRETU(0)
    TXTFIELDS(12) = aRETU(1)
    TXTFIELDS(13) = aRETU(2)
    TXTFIELDS(14) = aRETU(3)
    TXTFIELDS(15) = aRETU(4)
    TXTFIELDS(16) = aRETU(5)
    TXTFIELDS(7) = aRETU(6)
    TXTFIELDS(8) = aRETU(7)
    TXTFIELDS(9) = aRETU(8)
    TXTFIELDS(10) = aRETU(9)
    TXTFIELDS(21) = aRETU(10)
    TXTFIELDS(22) = aRETU(11)
    TXTFIELDS(17) = aRETU(12)
    TXTFIELDS(18) = aRETU(13)
    TXTFIELDS(19) = aRETU(14)
    TXTFIELDS(20) = aRETU(15)
    Check1.Value = aRETU(16)
End Sub

Private Sub CMDSIG_Click(Index As Integer)
    lRETU = False
    eRETU01 = ""
    eRETU02 = ""
    If Index = 0 Or Index = 2 Or Index = 4 Then
        escFLX.Show vbModal, Me
    Else
        ePASS01 = "isoqsymbol"
        ePASS02 = True
        frmCharacters.Show vbModal, Me
    End If
    If lRETU Then
        If Index = 0 Or Index = 1 Then
            Me.TXTFIELDS(24) = eRETU01
        End If
        If Index = 2 Or Index = 3 Then
            Me.TXTFIELDS(42) = eRETU01
        End If
        If Index = 4 Or Index = 5 Then
            Me.TXTFIELDS(43) = eRETU01
        End If
    End If


End Sub

Private Sub Cmdsitu_Click(Index As Integer)

    Select Case Index
    Case 0
        TXTFIELDS(23) = "C"
    Case 1
        TXTFIELDS(23) = "N"
    Case 2
        TXTFIELDS(23) = "E"
    End Select

End Sub

Private Sub Command1_Click(Index As Integer)
    ePASS01 = ""
    If Index > 9 Then
        ePASS01 = "LOGIX"
    End If
    Select Case Index
    Case 0, 10
        escMP04.Show vbModal, Me
        If lRETU Then
            frmFEMEI.TXTFIELDS(12) = eRETU01
            frmFEMEI.TXTFIELDS(13) = eRETU02
        End If
    Case 1, 11
        escMP04.Show vbModal, Me
        If lRETU Then
            frmFEMEI.TXTFIELDS(14) = eRETU01
            frmFEMEI.TXTFIELDS(21) = eRETU02
        End If
    Case 2, 12
        escMP04.Show vbModal, Me
        If lRETU Then
            frmFEMEI.TXTFIELDS(16) = eRETU01
            frmFEMEI.TXTFIELDS(22) = eRETU02
        End If
    Case 3
        escMP05.Show vbModal, Me
        If lRETU Then
            frmFEMEI.TXTFIELDS(12) = 0
            frmFEMEI.TXTFIELDS(13) = eRETU01
        End If
    Case 4
        escMP05.Show vbModal, Me
        If lRETU Then
            frmFEMEI.TXTFIELDS(14) = 0
            frmFEMEI.TXTFIELDS(21) = eRETU01
        End If
    Case 5
        escMP05.Show vbModal, Me
        If lRETU Then
            frmFEMEI.TXTFIELDS(16) = 0
            frmFEMEI.TXTFIELDS(22) = eRETU01
        End If
    End Select

End Sub

Private Sub Command2_Click()

    Check1.Value = Checked
    acaocampos

End Sub

Private Sub Command3_Click()
    On Error Resume Next

    If MDG("Sem Ação Zera Dados? ", "Confirme Gravação") Then
        Check1.Value = Unchecked
        acaocampos
        TXTFIELDS(11) = ""
        TXTFIELDS(12) = 0
        TXTFIELDS(13) = ""
        TXTFIELDS(14) = 0
        TXTFIELDS(21) = ""
        TXTFIELDS(16) = 0
        TXTFIELDS(22) = ""
        TXTFIELDS(15) = ""
        TXTFIELDS(17) = 0
        TXTFIELDS(18) = 0
        TXTFIELDS(19) = 0
        TXTFIELDS(20) = 0
        DTPicker1 = NullDate()
        DTPicker2 = NullDate()
        DTPicker3 = NullDate()
        DTPicker4 = NullDate()
    End If

End Sub

Private Sub Command4_Click(Index As Integer)
    Dim cARQ As String
    Dim aRETU As Variant
    Dim sSQL             As String
    Dim nNUMERO          As Long

    Select Case Index
    Case 0
        nNUMERO = FixInt(frmFEMEI.TXTFIELDS(12), 0)
    Case 1
        nNUMERO = FixInt(frmFEMEI.TXTFIELDS(14), 0)
    Case 2
        nNUMERO = FixInt(frmFEMEI.TXTFIELDS(16), 0)
    End Select
    cARQ = PegPath("PATH", "LOGIXODBC")
    sSQL = "SELECT nom_completo as NOMTEC FROM funcionario WHERE cod_empresa='01' and num_matricula=" & nNUMERO
    aRETU = PegSQL(cARQ, sSQL, 1, Array("NOMTEC"), Array("C"), Array(""))
    If Not lRETU Then
        cARQ = PegPath("PATH", "CADMP04")
        cARQ = GeraConn(cARQ, "JETFOX")
        sSQL = "SELECT NOMTEC FROM MP04 WHERE TECNICO=" & nNUMERO
        aRETU = PegSQL(cARQ, sSQL, 1, Array("NOMTEC"), Array("C"), Array(""))
    End If
    
    If lRETU Then
        Select Case Index
        Case 0
            frmFEMEI.TXTFIELDS(13) = aRETU(0)
        Case 1
            frmFEMEI.TXTFIELDS(21) = aRETU(0)
        Case 2
            frmFEMEI.TXTFIELDS(22) = aRETU(0)
        End Select
    End If
End Sub

Private Sub Encerrar_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
    Encerrar.Enabled = True
    Center Me
    cARQPF = PegPath("PATH", "PF")
    cARQFEMEA = PegPath("PATH", "FEMEA")
    cARQFEMEAPF = PegPath("PATH", "FEMEAPF")
    cARQUSO = ePASS02
    cSQLUSO = ePASS01
    '    ePASS04 versao do femea
    '    epass03 index qual arquivo trabalha 0 producao
    '    utilizacao abaixo
    nCAMPOS = 60 '47
    aCAM = Array("PF", "ITEM", "PROCESSO", "FALTIP", "FALEFE", _
                 "FALCAU", "CRTATU", "INDOCO", "INDSEV", "INDDET", _
                 "INDRIS", "ACAREC", "RESCOD", "RESNOM", "RESCOD2", _
                 "ACATOM", "RESCOD3", "RINDOCO", "RINDSER", "RINDDET", _
                 "RINDRIS", "RESNOM2", "RESNOM3", "SITUACAO", "SIGI", _
                 "CARAPREV", "RESDAT", "RESDAT2", "RESDAT3", "ACADAT", _
                 "ACAO", "EXCRPN", "ALTMAN", "MUDPAD", "PSA", "FXSEQ", "FXSSQ", "FXITEM", _
                 "PRONUM", "EFENUM", "FALNUM", "CAUNUM", "SIG2", "SIG3", "SUBTIPO", "FXITEMS", "FEMEAREV", _
                 "elemento", "estacao", "trabalho", "funcaoitem", "funcaoetapa", "funcaoelemento", _
                 "pafemea", "caraespecial", "filtroespecial", "acaoprev", "acaodet", "rpafemea", "observacao")
    aFOR = Array("NI", "NI", "C", "C", "C", _
                 "C", "C", "NI", "NI", "NI", _
                 "NI", "C", "NI", "C", "NI", _
                 "C", "NI", "NI", "NI", "NI", _
                 "NI", "C", "C", "C", "C", "C", _
                 "DN", "DN", "DN", "DN", "BN", _
                 "BN", "BN", "BN", "C", "NI", "NI", "NI", "NI", "NI", "NI", "NI", "C", "C", "C", "NI", "NI", _
                 "C", "C", "C", "C", "C", "C", "C", "C", "C", "C", "C", "C", "C")
    aPAD = Array(0, 0, "", "", "", _
                 "", "", 0, 0, 0, _
                 0, "", 0, "", 0, _
                 "", 0, 0, 0, 0, 0, _
                 "", "", "", "", "", _
                 "", "", "", "", False, _
                 False, False, False, "", 0, 0, 0, 0, 0, 0, 0, "", "", "", 0, 0, _
                 "", "", "", "", "", "", "", "", "", "", "", "", "")
                 
    aVAL = PegSQL(cARQUSO, cSQLUSO, nCAMPOS, aCAM, aFOR, aPAD)
    gravaAvalTxt (0)
    
    If ePASS03 = 0 Then                          'so producao
        If FixNum(TXTFIELDS(46)) = 0 Then        'precisa checar se e em branco pois no historico e revisoes consta revisoes antigas
            TXTFIELDS(46) = ePASS04              ''Revisao femea
        Else
            If FixNum(TXTFIELDS(46)) < FixNum(ePASS04) Then 'houve revisao e esta menor
                TXTFIELDS(46) = ePASS04          ''Revisao femea
            End If
        End If
    End If
    
    EnableControls Me, False
    cmdClose.Enabled = False
    CmdAltman.Enabled = True
    Check3.Enabled = False
    Encerrar.Enabled = True
    
    
    TXTFIELDS(24).Font = "isoqsymbol"
    TXTFIELDS(42).Font = "isoqsymbol"
    TXTFIELDS(43).Font = "isoqsymbol"
    
    'Calcular indices e PA
    TXTFIELDS_Change 7
    TXTFIELDS_Change 17
    
End Sub

Private Sub gravaAvalTxt(nUSO)
    For nITEM = nUSO To 25
        TXTFIELDS(nITEM) = aVAL(nITEM)
    Next nITEM
    If Not IsNull(aVAL(26)) And aVAL(26) <> "" Then
        DTPicker1.Value = aVAL(26)
    End If
    If Not IsNull(aVAL(27)) And aVAL(27) <> "" Then
        DTPicker2.Value = aVAL(27)
    End If
    If Not IsNull(aVAL(28)) And aVAL(28) <> "" Then
        DTPicker3.Value = aVAL(28)
    End If
    If Not IsNull(aVAL(29)) And aVAL(29) <> "" Then
        DTPicker4.Value = aVAL(29)
    End If
    Check1.Value = aVAL(30)
    Check2.Value = aVAL(31)
    If nUSO = 0 Then
        Check3.Value = aVAL(32)
        Check4.Value = aVAL(33)
        For nITEM = 34 To 41                     'aVAL 34 35 36 37 38 39 40 41  pronum,falcum,efenum,caunum nao gravar
            TXTFIELDS(nITEM) = aVAL(nITEM)
        Next
    Else
        For nITEM = 34 To 37                     'aVAL 38 39 40 41  pronum,falcum,efenum,caunum nao gravar
            TXTFIELDS(nITEM) = aVAL(nITEM)
        Next
    End If
    For nITEM = 42 To 59 '46                         'campos textos finais
        TXTFIELDS(nITEM) = aVAL(nITEM)
    Next
 
End Sub

Private Sub acaocampos()

    Dim lENABLE As Boolean
    Dim x As Integer
    lENABLE = False

    If Check1.Value = Checked Then
        lENABLE = True
    End If
    For x = 11 To 22
        TXTFIELDS(x).Enabled = lENABLE
    Next x
    DTPicker1.Enabled = lENABLE
    DTPicker4.Enabled = lENABLE
    DTPicker2.Enabled = lENABLE
    DTPicker3.Enabled = lENABLE
    Command4(0).Enabled = lENABLE
    Command4(1).Enabled = lENABLE
    Command4(2).Enabled = lENABLE
    For x = 0 To 5
        Command1(x).Enabled = lENABLE
    Next x
    For x = 17 To 19
        cmdmais(x).Enabled = lENABLE
        cmdmenos(x).Enabled = lENABLE
    Next
    pfcescid(4).Enabled = lENABLE
    pfcescid(5).Enabled = lENABLE
    pfcescid(6).Enabled = lENABLE
    For x = 42 To 45
        TXTFIELDS(x).Enabled = lENABLE
    Next x
    ''46 Revisao sempre travado
    For x = 47 To 59
        TXTFIELDS(x).Enabled = lENABLE
    Next x
    CmdRetornaAcao(0).Enabled = lENABLE
    CmdRetornaAcao(1).Enabled = lENABLE
End Sub

Private Sub grvitem_Click()
    Dim cTITULO As String
    Dim aCAMTMP As Variant


    
    cTITULO = FixStr(Busca("Digite o Titulo", "Titulo Pre-Cadastro Femea", Space(50), 50), "")
    If Len(cTITULO) = 0 Then
        Alert ("Titulo Necessario")
        Exit Sub
    End If
    
    
    For nITEM = 0 To 25
        aVAL(nITEM) = TXTFIELDS(nITEM)
    Next nITEM
    aVAL(26) = NullDate()
    aVAL(27) = NullDate()
    aVAL(28) = NullDate()
    aVAL(29) = NullDate()
    aVAL(30) = Check1.Value
    aVAL(31) = Check2.Value
    aVAL(32) = Check3.Value
    aVAL(33) = Check4.Value
    aCAMTMP = aCAM
    aCAMTMP(2) = "TITULO"
    aVAL(2) = cTITULO
    For nITEM = 34 To 37
        aVAL(nITEM) = TXTFIELDS(nITEM)
    Next
    IncluiSQL cARQFEMEA, "SELECT * FROM FEMPRE WHERE TITULO'=" & cTITULO & "'", nCAMPOS, aCAMTMP, aVAL, False, False

End Sub

Private Sub pegitem_Click()
    Dim sSQL As String
    On Error Resume Next
    escfrmpre.Show vbModal, Me
    If lRETU And Val(eRETU01) > 0 Then
        sSQL = "SELECT * FROM FEMPRE WHERE SEGGRA=" & eRETU01
        aVAL = PegSQL(cARQFEMEA, sSQL, nCAMPOS, aCAM, aFOR, aPAD)
        gravaAvalTxt (3)
        TXTFIELDS_Change (7)
        TXTFIELDS_Change (17)
    End If
End Sub

Private Sub pfcescid_Click(Index As Integer)
    iMU01 = 300 + Index
    If Index > 3 Then iMU01 = iMU01 - 3
    eRETU03 = 0                                  'Evita Cancelamento Escolha
    escIED.Show vbModal, Me
    If lRETU Then
        Select Case Index
        Case 1
            TXTFIELDS(7) = eRETU03
        Case 2
            TXTFIELDS(8) = eRETU03
        Case 3
            TXTFIELDS(9) = eRETU03
        Case 4
            TXTFIELDS(17) = eRETU03
        Case 5
            TXTFIELDS(18) = eRETU03
        Case 6
            TXTFIELDS(19) = eRETU03
        End Select
    End If
End Sub



Private Sub TXTFIELDS_Change(Index As Integer)
If Index = 7 Or Index = 8 Or Index = 9 Or Index = 17 Or Index = 18 Or Index = 19 Then
    If FixNum(TXTFIELDS(Index)) > 10 Then
       TXTFIELDS(Index) = 10
    End If
    Select Case Index
        Case 7, 8, 9
            TXTFIELDS(10) = FixNum(TXTFIELDS(7)) * FixNum(TXTFIELDS(8)) * FixNum(TXTFIELDS(9))
            TXTFIELDS(53) = subcalcpa(FixNum(TXTFIELDS(8)), FixNum(TXTFIELDS(7)), FixNum(TXTFIELDS(9)))
        Case 17, 18, 19
            TXTFIELDS(20) = FixNum(TXTFIELDS(17)) * FixNum(TXTFIELDS(18)) * FixNum(TXTFIELDS(19))
            TXTFIELDS(58) = subcalcpa(FixNum(TXTFIELDS(18)), FixNum(TXTFIELDS(17)), FixNum(TXTFIELDS(19)))
    End Select
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

Private Function subcalcpa(nSEV As Integer, nOCO As Integer, nDET)
subcalcpa = ""
Select Case nSEV
    Case 9, 10
        Select Case nOCO
             Case 1
                subcalcpa = "B"
             Case 2, 3
                   Select Case nDET
                       Case 7, 8, 9, 10
                            subcalcpa = "A"
                       Case 5, 6
                            subcalcpa = "M"
                       Case 1, 2, 3, 4
                            subcalcpa = "B"
                  End Select
             Case 4, 5
                  Select Case nDET
                       Case 2, 3, 4, 5, 6, 7, 8, 9, 10
                            subcalcpa = "A"
                       Case 1
                            subcalcpa = "M"
                  End Select
             Case 6, 7, 8, 9, 10
                subcalcpa = "A"
        End Select
    
    Case 7, 8
        Select Case nOCO
            Case 1
                subcalcpa = "B"
            Case 2, 3
                Select Case nDET
                       Case 1, 2, 3, 4
                            subcalcpa = "B"
                       Case 5, 6, 7, 8, 9, 10
                            subcalcpa = "M"
                End Select
            Case 4, 5
                Select Case nDET
                       Case 1, 2, 3, 4, 5, 6
                            subcalcpa = "M"
                       Case 7, 8, 9, 10
                            subcalcpa = "A"
                End Select
            Case 6, 7
                Select Case nDET
                       Case 1
                            subcalcpa = "M"
                       Case 2, 3, 4, 5, 6, 7, 8, 9, 10
                            subcalcpa = "A"
                End Select
            Case 8, 9, 10
                subcalcpa = "A"
        End Select
    Case 4, 5, 6
       Select Case nOCO
           Case 1, 2, 3
               subcalcpa = "B"
           Case 4, 5
                Select Case nDET
                       Case 1, 2, 3, 4, 5, 6
                            subcalcpa = "B"
                       Case 7, 8, 9, 10
                            subcalcpa = "M"
                End Select
           Case 6, 7
                Select Case nDET
                       Case 1
                            subcalcpa = "B"
                       Case 2, 3, 4, 5, 6, 7, 8, 9, 10
                            subcalcpa = "M"
                End Select

           Case 8, 10
                Select Case nDET
                       Case 1, 2, 3, 4
                            subcalcpa = "M"
                       Case 5, 6, 7, 8, 9, 10
                            subcalcpa = "A"
                End Select

       End Select
    Case 2, 3
       Select Case nOCO
              Case 1, 2, 3, 4, 5, 6, 7
                   subcalcpa = "B"
              Case 8, 9, 10
                   Select Case nDET
                        Case 1, 2, 3, 4
                            subcalcpa = "B"
                        Case 5, 6, 7, 8, 9, 10
                            subcalcpa = "M"
                   End Select
       End Select
    Case 1
     subcalcpa = "B"
End Select
End Function

