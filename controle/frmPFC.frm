VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form frmPFC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PFC"
   ClientHeight    =   6360
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   9600
   Icon            =   "frmPFC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox TxtCodComp 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   372
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox Txtssq 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   372
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   94
      TabStop         =   0   'False
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox Txtseq 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   372
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtPF 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   372
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   92
      TabStop         =   0   'False
      Top             =   960
      Width           =   855
   End
   Begin TabExCtl.SSTabEx SSTab1 
      Height          =   4695
      Left            =   120
      TabIndex        =   16
      Top             =   1500
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   8281
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Caracteristicas"
      TabPicture(0)   =   "frmPFC.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblLabels(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblLabels(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblLabels(15)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblLabels(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblLabels(12)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblLabels(9)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblLabels(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblLabels(13)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblLabels(7)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblLabels(10)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "CmdPegRel"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "pfcescid(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "pfcescid(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "pfcescid(8)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtFields(2)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtFields(1)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtFields(15)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtFields(5)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "pfcescid(3)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "pfcescid(4)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cmdescpfd"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtFields(6)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "pfcescid(5)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtFields(13)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "pfcescid(7)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtFields(7)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Check1"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "pfcescid(6)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtFields(11)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtFields(4)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtFields(10)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).ControlCount=   31
      TabCaption(1)   =   "Simbologia Instrumentos"
      TabPicture(1)   =   "frmPFC.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtFields(20)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "EscTipIns(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command4(1)"
      Tab(1).Control(3)=   "txtFields(9)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "EscTipIns(0)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command4(0)"
      Tab(1).Control(6)=   "txtFields(19)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command2(1)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "CmdEditMe04(1)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Command3(1)"
      Tab(1).Control(10)=   "ESCMS01A(1)"
      Tab(1).Control(11)=   "txtFields(8)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtFields(17)"
      Tab(1).Control(13)=   "Command2(0)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "CmdEditMe04(0)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Command3(0)"
      Tab(1).Control(16)=   "ESCMS01A(0)"
      Tab(1).Control(17)=   "txtFields(16)"
      Tab(1).Control(18)=   "txtFields(22)"
      Tab(1).Control(19)=   "Command1(5)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Command1(4)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "txtFields(21)"
      Tab(1).Control(22)=   "Command1(3)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Command1(2)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Command1(1)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Command1(0)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "txtFields(3)"
      Tab(1).Control(27)=   "cmdFOTO(0)"
      Tab(1).Control(28)=   "cmdFOTO(1)"
      Tab(1).Control(29)=   "lblLabels(16)"
      Tab(1).Control(30)=   "lblLabels(8)"
      Tab(1).Control(31)=   "lblLabels(5)"
      Tab(1).ControlCount=   32
      TabCaption(2)   =   "figura"
      TabPicture(2)   =   "frmPFC.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture1"
      Tab(2).Control(1)=   "Picture2"
      Tab(2).Control(2)=   "cmdimprimir"
      Tab(2).Control(3)=   "CmdPaste"
      Tab(2).Control(4)=   "CmdConfImp"
      Tab(2).Control(5)=   "Incluirimagem"
      Tab(2).Control(6)=   "DelImg"
      Tab(2).Control(7)=   "VerImg"
      Tab(2).Control(8)=   "CmdSalvaImagem"
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "Ensaios"
      TabPicture(3)   =   "frmPFC.frx":05DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      Begin VB.PictureBox Picture1 
         Height          =   495
         Left            =   -74880
         ScaleHeight     =   435
         ScaleWidth      =   195
         TabIndex        =   78
         Top             =   3960
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   3375
         Left            =   -74880
         ScaleHeight     =   3315
         ScaleWidth      =   3555
         TabIndex        =   77
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   20
         Left            =   -71280
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   3000
         Width           =   3375
      End
      Begin VB.CommandButton EscTipIns 
         Height          =   315
         Index           =   1
         Left            =   -71880
         Picture         =   "frmPFC.frx":05FA
         Style           =   1  'Graphical
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   2520
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Caption         =   "MC"
         Height          =   315
         Index           =   1
         Left            =   -72360
         TabIndex        =   74
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   9
         Left            =   -71280
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   1920
         Width           =   3375
      End
      Begin VB.CommandButton EscTipIns 
         Height          =   315
         Index           =   0
         Left            =   -71880
         Picture         =   "frmPFC.frx":0B84
         Style           =   1  'Graphical
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Caption         =   "MC"
         Height          =   315
         Index           =   0
         Left            =   -72360
         TabIndex        =   71
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   19
         Left            =   -72360
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   3000
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Height          =   375
         Index           =   1
         Left            =   -73200
         Picture         =   "frmPFC.frx":110E
         Style           =   1  'Graphical
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   2520
         Width           =   375
      End
      Begin VB.CommandButton CmdEditMe04 
         Caption         =   "E"
         Height          =   375
         Index           =   1
         Left            =   -73680
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   2520
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Caption         =   "MC"
         Height          =   375
         Index           =   1
         Left            =   -74160
         TabIndex        =   67
         Top             =   2520
         Width           =   375
      End
      Begin VB.CommandButton ESCMS01A 
         Caption         =   "LX"
         Height          =   375
         Index           =   1
         Left            =   -74760
         TabIndex        =   66
         Top             =   2520
         Width           =   435
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   8
         Left            =   -72360
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   17
         Left            =   -74880
         MaxLength       =   20
         TabIndex        =   64
         Top             =   3000
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         Height          =   375
         Index           =   0
         Left            =   -73200
         Picture         =   "frmPFC.frx":1698
         Style           =   1  'Graphical
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton CmdEditMe04 
         Caption         =   "E"
         Height          =   375
         Index           =   0
         Left            =   -73680
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Caption         =   "MC"
         Height          =   375
         Index           =   0
         Left            =   -74160
         TabIndex        =   61
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton ESCMS01A 
         Caption         =   "LX"
         Height          =   375
         Index           =   0
         Left            =   -74760
         TabIndex        =   60
         Top             =   1440
         Width           =   435
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   16
         Left            =   -74880
         MaxLength       =   20
         TabIndex        =   58
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   15.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   22
         Left            =   -71400
         MaxLength       =   1
         TabIndex        =   56
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Index           =   5
         Left            =   -71880
         Picture         =   "frmPFC.frx":1C22
         Style           =   1  'Graphical
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Index           =   4
         Left            =   -71880
         Picture         =   "frmPFC.frx":21AC
         Style           =   1  'Graphical
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   15.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   21
         Left            =   -72480
         MaxLength       =   1
         TabIndex        =   53
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Index           =   3
         Left            =   -72960
         Picture         =   "frmPFC.frx":2736
         Style           =   1  'Graphical
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Index           =   2
         Left            =   -72960
         Picture         =   "frmPFC.frx":2CC0
         Style           =   1  'Graphical
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Index           =   1
         Left            =   -74040
         Picture         =   "frmPFC.frx":324A
         Style           =   1  'Graphical
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   315
         Index           =   0
         Left            =   -74040
         Picture         =   "frmPFC.frx":37D4
         Style           =   1  'Graphical
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   15.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   3
         Left            =   -73560
         MaxLength       =   1
         TabIndex        =   48
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   10
         Left            =   2280
         TabIndex        =   46
         Top             =   3000
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   4
         Left            =   1920
         TabIndex        =   45
         Top             =   2040
         Width           =   6495
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   11
         Left            =   2280
         TabIndex        =   44
         Top             =   3960
         Width           =   3375
      End
      Begin VB.CommandButton pfcescid 
         Height          =   315
         Index           =   6
         Left            =   1800
         Picture         =   "frmPFC.frx":3D5E
         Style           =   1  'Graphical
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   3960
         Width           =   375
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Certificado Fornecedor"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   6120
         TabIndex        =   41
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   7
         Left            =   2280
         TabIndex        =   40
         Top             =   3480
         Width           =   3375
      End
      Begin VB.CommandButton pfcescid 
         Height          =   315
         Index           =   7
         Left            =   1800
         Picture         =   "frmPFC.frx":42E8
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   3480
         Width           =   375
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   13
         Left            =   6720
         TabIndex        =   37
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton pfcescid 
         Height          =   315
         Index           =   5
         Left            =   6120
         Picture         =   "frmPFC.frx":4872
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   6
         Left            =   3600
         TabIndex        =   34
         Top             =   2520
         Width           =   1875
      End
      Begin VB.CommandButton cmdescpfd 
         Height          =   315
         Left            =   3120
         Picture         =   "frmPFC.frx":4DFC
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   2520
         Width           =   375
      End
      Begin VB.CommandButton pfcescid 
         Height          =   315
         Index           =   4
         Left            =   1800
         Picture         =   "frmPFC.frx":5386
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   3000
         Width           =   375
      End
      Begin VB.CommandButton pfcescid 
         Height          =   315
         Index           =   3
         Left            =   960
         Picture         =   "frmPFC.frx":5910
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   5
         Left            =   1440
         TabIndex        =   28
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtFields 
         Height          =   405
         Index           =   15
         Left            =   1920
         TabIndex        =   26
         Top             =   1560
         Width           =   6495
      End
      Begin VB.TextBox txtFields 
         Height          =   525
         Index           =   1
         Left            =   1920
         ScrollBars      =   3  'Both
         TabIndex        =   25
         Top             =   960
         Width           =   6495
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   24
         Top             =   480
         Width           =   6495
      End
      Begin VB.CommandButton pfcescid 
         Height          =   315
         Index           =   8
         Left            =   1320
         Picture         =   "frmPFC.frx":5E9A
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2040
         Width           =   375
      End
      Begin VB.CommandButton pfcescid 
         Height          =   315
         Index           =   1
         Left            =   1320
         Picture         =   "frmPFC.frx":6424
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton pfcescid 
         Height          =   315
         Index           =   2
         Left            =   1320
         Picture         =   "frmPFC.frx":69AE
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   480
         Width           =   375
      End
      Begin XPControls.XPButton cmdFOTO 
         Height          =   435
         Index           =   0
         Left            =   -71160
         TabIndex        =   81
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
         Picture         =   "frmPFC.frx":6F38
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
         Height          =   435
         Index           =   1
         Left            =   -71160
         TabIndex        =   82
         Top             =   2520
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   767
         Picture         =   "frmPFC.frx":74D2
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
         Left            =   -74400
         TabIndex        =   83
         Top             =   3960
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
         Picture         =   "frmPFC.frx":7A6C
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
         Left            =   -72840
         TabIndex        =   84
         Top             =   3960
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   767
         Picture         =   "frmPFC.frx":8006
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
      Begin XPControls.XPButton CmdConfImp 
         Height          =   435
         Left            =   -70320
         TabIndex        =   85
         Top             =   3960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFC.frx":85A0
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
      Begin XPControls.XPButton Incluirimagem 
         Height          =   435
         Left            =   -71160
         TabIndex        =   86
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFC.frx":8B3A
         Caption         =   "Incluir Image"
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
         Left            =   -71160
         TabIndex        =   87
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFC.frx":8FD4
         Caption         =   "Excluir Image"
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
         Left            =   -71160
         TabIndex        =   88
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmPFC.frx":946E
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
      Begin XPControls.XPButton CmdSalvaImagem 
         Height          =   435
         Left            =   -71160
         TabIndex        =   90
         Top             =   2040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFC.frx":9908
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
      Begin XPControls.XPButton CmdPegRel 
         Height          =   375
         Left            =   1320
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   1440
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Picture         =   "frmPFC.frx":9DA2
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
         Caption         =   "Tipo Instr."
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   16
         Left            =   -72120
         TabIndex        =   59
         Top             =   1200
         Width           =   915
      End
      Begin VB.Label lblLabels 
         Caption         =   "Instrumento"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   8
         Left            =   -74880
         TabIndex        =   57
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Simbologia"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   5
         Left            =   -74880
         TabIndex        =   47
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Plano de Reação"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   42
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Caption         =   "Frequência"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   38
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Carta"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   13
         Left            =   5640
         TabIndex        =   35
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Qtde:"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   6
         Left            =   2640
         TabIndex        =   32
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label lblLabels 
         Caption         =   "Capacidade/Precisão"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   30
         Top             =   3000
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tolerancia"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   27
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Especificação"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Carac.Processo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   21
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         Caption         =   "Descrição"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Característica"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Layout(RI)"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6000
      TabIndex        =   15
      Top             =   480
      Width           =   1155
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Processo"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6000
      TabIndex        =   14
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   18
      Left            =   5160
      TabIndex        =   4
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton Setor 
      Caption         =   "Recebimento"
      Height          =   315
      Index           =   3
      Left            =   3960
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton Setor 
      Caption         =   "Laboratorio"
      Height          =   315
      Index           =   2
      Left            =   3960
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton Setor 
      Caption         =   "Operador"
      Height          =   315
      Index           =   1
      Left            =   2640
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Setor 
      Caption         =   "Produção"
      Height          =   315
      Index           =   0
      Left            =   2640
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   14
      Left            =   720
      MaxLength       =   3
      TabIndex        =   1
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   12
      Left            =   2160
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   375
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   2
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtItem 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   495
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   7680
      TabIndex        =   79
      Top             =   480
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmPFC.frx":A33C
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
      Left            =   7680
      TabIndex        =   80
      Top             =   0
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmPFC.frx":A8D6
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
   Begin VB.Label Label3 
      Caption         =   "Componente"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3000
      TabIndex        =   98
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "SSQ"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2040
      TabIndex        =   97
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Seq"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1080
      TabIndex        =   96
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label11 
      Caption         =   "PF"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   91
      Top             =   720
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Fabrica"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   18
      Left            =   5160
      TabIndex        =   13
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "Item"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   14
      Left            =   840
      TabIndex        =   8
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Setores"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   11
      Left            =   2160
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Caption         =   "OP:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Caption         =   "ORD:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmPFC"
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
Dim lTROCOU As Boolean

Private Sub cmdClose_Click()
    On Error Resume Next
    
    If Not CHECKME04 Then
        If MDG("Continuar Mesmo Assim") Then
            If Not MDG("Gravar Mesmo com Erro no Instrumento") Then
                Exit Sub
            End If
        End If
    End If
    
    If MDG("Gravar alteraçôes") Then
        For iLOOP = 0 To nCAMPOS - 4
            aVAL(iLOOP) = TXTFIELDS(iLOOP)
        Next iLOOP
        aVAL(23) = FixNumBol(Check1.Value)
        aVAL(24) = FixNumBol(Check2.Value)
        aVAL(25) = FixNumBol(Check3.Value)
        GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR
        
        If lTROCOU Then
            ADOGrvBlob cARQ, cSQL, Picture1, "IMAGEM"
        End If
       
        'Select Case nARQPCS
        '     Case 0, 1
        '         eRETU02 = "PF:" & nPF & " " & " SEQ:" & nSEQ & " SSQ:" & nSSQ & " ITEM:" & nORD
        '     Case 2
        '         eRETU02 = "PF:" & nPF & " Componente:" & Ccodcomp & " Item:" & nORD
        ' End Select
        ' eRETU02 = eRETU02 & " Desenho:" & frmPCX.TXTPF(2) & Chr(13) & Chr(10)
        ' eRETU02 = eRETU02 & " Data:" & Fdata(Date) & Chr(13) & Chr(10)
        ' MAILENV "PI000001", eRETU02
    End If
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub CmdConfImp_Click()
    FrmPrintSetup.Show vbModal, Me
End Sub

Private Sub cmdescpfd_Click()

    If Me.TXTFIELDS(12) = "" Then

        Alert "Setor em Branco"
        Exit Sub

    End If

    Zsetor = Me.TXTFIELDS(12)
    lRETU = False
    eRETU01 = ""
    eRETU02 = ""
    escPFD.Show vbModal, Me

    If lRETU Then

        Me.TXTFIELDS(6) = eRETU01
        Me.TXTFIELDS(7) = eRETU02

    End If

End Sub

Private Sub cmdFOTO_Click(Index As Integer)
    Dim cSQL As String
    Dim cCODIGO As String

   
    If Index = 0 Then
        cCODIGO = FixStr(TXTFIELDS(16))
    Else
        cCODIGO = FixStr(TXTFIELDS(17))
    End If
    If Len(cCODIGO) = 0 Then
        Alert "Codigo nao preenchido"
        Exit Sub
    End If
    iImage = 5
    cARQRTF = PegPath("PATH", "IMGME04")
    cSQL = "select * from IMAGENS WHERE CODIGO='" & cCODIGO & "'"
    IncluiSQL cARQRTF, cSQL, 1, Array("CODIGO"), Array(cCODIGO), True, False
    zgrp = cCODIGO                               ''zGRP Interno frm imagens
    Load frmIMAGENS
    frmIMAGENS.TXTFIELDS(0).Enabled = False
    frmIMAGENS.Escolher(0).Visible = False
    frmIMAGENS.Show vbModal, Me
End Sub

Private Sub CmdImprimir_Click()
     If Picture1.Height = 0 Then
        Alert ("Sem Imagem")
    Else
        PrinterEx.ShowPrintPreview Me, "MyPrintingRoutine"
    End If
End Sub

Private Sub CmdPaste_Click()
    If Picture1.Height = 0 Then
        Alert ("Sem Imagem")
    Else
        Clipboard.Clear
        Clipboard.SetData Picture1.Image, (vbCFBitmap)

        ' CopyEntirePicture Picture1
    End If

End Sub

Private Sub CmdPegRel_Click()
    Dim cARQ As String
    Dim aRETU As Variant
    Dim sSQL As String
    cARQ = PegPath("PATH", "PF")
    sSQL = "select regular from PFS WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ
    aRETU = PegSQL(cARQ, sSQL, 1, Array("REGULAR"), Array(""), Array(""))
    TXTFIELDS(15) = aRETU(0)
End Sub

Private Sub CmdSalvaImagem_Click()
    salvarpict Me, Picture1, "imagem"
End Sub

Private Sub Command1_Click(Index As Integer)

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
            Me.TXTFIELDS(3) = eRETU01
        End If
        If Index = 2 Or Index = 3 Then
            Me.TXTFIELDS(21) = eRETU01
        End If
        If Index = 4 Or Index = 5 Then
            Me.TXTFIELDS(22) = eRETU01
        End If
        
    End If

End Sub

Private Sub Command2_Click(Index As Integer)
    lRETU = False
    If Index = 0 Then
        eRETU01 = TXTFIELDS(8)
    Else
        eRETU01 = TXTFIELDS(19)
    End If
    eRETU02 = ""
    eRETU03 = ""
    ePASS01 = ""
    escME04.Show vbModal, Me
    
    If lRETU Then
        If Index = 0 Then
            TXTFIELDS(16) = eRETU01
            TXTFIELDS(8) = eRETU02
            TXTFIELDS(9) = eRETU03
        Else
            TXTFIELDS(17) = eRETU01
            TXTFIELDS(19) = eRETU02
            TXTFIELDS(20) = eRETU03
        End If
    End If
    
    CHECKME04
  
End Sub

Private Sub Command3_Click(Index As Integer)
    Dim cARQTMP As String
    Dim cSQLTMP As String
    ePASS01 = "MICROSIGAODBC"
    ePASS02 = "SELECT qm2_instr as codigo,qm2_descr as nome from qm2010"
    ePASS03 = Array("qm2_instr", "qm2_descr")
    EscCodNomSim.Show vbModal, Me
    If lRETU Then
        If Index = 0 Then
            TXTFIELDS(16) = eRETU01
            TXTFIELDS(9) = eRETU02
        Else
            TXTFIELDS(17) = eRETU01
            TXTFIELDS(20) = eRETU02
        End If

        cARQTMP = PegPath("PATH", "MICROSIGAODBC")
        cSQLTMP = "SELECT qm2_tipo from qm2010 where qm2_instr='" & eRETU01 & "'"
   
        eRETU01 = PegSQL(cARQTMP, cSQLTMP, 1, Array("qm2_tipo"), Array("C"), Array(""))
        If lRETU Then
            If Index = 0 Then
                TXTFIELDS(8) = Mid(eRETU01(0), 1, 3)
            Else
                TXTFIELDS(19) = Mid(eRETU01(0), 1, 3)
            End If
        End If
    End If


End Sub

Private Sub Command4_Click(Index As Integer)
    ePASS01 = "MICROSIGAODBC"
    ePASS02 = "SELECT qm1_tipo as codigo,qm1_descr as nome from qm1010"
    ePASS03 = Array("qm1_tipo", "qm1_descr")
    EscCodNomSim.Show vbModal, Me
    If lRETU Then
        If lRETU Then
            If Index = 0 Then
                TXTFIELDS(8) = eRETU01
                TXTFIELDS(9) = eRETU02
            Else
                TXTFIELDS(19) = eRETU01
                TXTFIELDS(20) = eRETU02
            End If
        End If
    End If
End Sub

Private Sub DelImg_Click()
    Set Picture1.Picture = Nothing
    Set Picture2.Picture = Nothing
    lTROCOU = True

End Sub

Private Sub Encerrar_Click()
    If Not MDG("Sair sem gravar") Then
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub ESCMS01A_Click(Index As Integer)
    iMU01 = 12
    escmp.Show vbModal
    If lRETU Then
        Select Case Index
        Case 0
            TXTFIELDS(16) = eRETU01
        Case 1
            TXTFIELDS(17) = eRETU01
        End Select
    End If

End Sub

Private Sub EscTipIns_Click(Index As Integer)
    iMD02 = 1
    escMD02.Show vbModal, Me
    If lRETU Then
        If Index = 0 Then
            TXTFIELDS(8) = eRETU01
            TXTFIELDS(9) = eRETU02
        Else
            TXTFIELDS(19) = eRETU01
            TXTFIELDS(20) = eRETU02
        End If
    End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode
End Sub

Private Sub Form_Load()

    '  EscTipIns(0).Enabled = False
    '  EscTipIns(1).Enabled = False

    CenterFormToScreen Me
    txtItem = nORD
   
    If nPF > 0 Then
        TXTPF.tEXT = nPF
    Else
        TXTPF.Visible = False
        Label1.Visible = False
    End If
    If nSEQ > 0 Then
        TXTSEQ.tEXT = nSEQ
    Else
        TXTSEQ.Visible = False
        Label1.Visible = False
    End If
    If nSSQ > 0 Then
        TXTSSQ.tEXT = nSSQ
    Else
        TXTSSQ.Visible = False
        Label2.Visible = False
    End If
    If nARQPCS = 2 Then
        TxtCodComp.tEXT = Ccodcomp
    Else
        TxtCodComp.Visible = False
        Label3.Visible = False
    End If
   
   
   
 
    aCAM = Array("COP", "DEscr", "CARAC", "SIGI", "ESPE", "TOL", "cqtde", "freq", _
                 "TIPINS", "InStr", "CAPA", "REACAO", "Setor", "CARTA", "CITEM", "PROCESSO", _
                 "CODME04", "CODME04B", "FILIAL", "TIPINSB", "INSTRB", "SIG2", "SIG3", "CERTFOR", "SAIPROC", "SAIRI")
      
    aPAD = Array("", "", "", "", "", "", "", "", _
                 "", "", "", "", "", "", "", "", _
                 "", "", "", "", "", "", "", False, False, False)
                    
    aFOR = Array("C", "C", "C", "C", "C", "C", "C", "C", _
                 "C", "C", "C", "C", "C", "C", "C", "C", _
                 "C", "C", "C", "C", "C", "C", "C", "BN", "BN", "BN")

    cARQ = Sdb

    cSQL = "select COP,DEscr,CARAC,SIGI,ESPE,TOL,cqtde,freq, "
    cSQL = cSQL & "TIPINS,InStr,CAPA,REACAO,Setor,CARTA,CITEM,PROCESSO,"
    cSQL = cSQL & "CODME04,CODME04B,FILIAL,TIPINSB,INSTRB,SIG2,SIG3,CERTFOR,SAIPROC,SAIRI from "
    
    Select Case nARQPCS
    Case 0
        cSQL = "PFC WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ & " AND ITEM=" & nORD
    Case 1
        cSQL = "PFCO WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ & " AND ITEM=" & nORD
    Case 2
        cSQL = "PFCMS03 WHERE PF=" & nPF & " AND CODCOMP='" & Ccodcomp & "' AND ITEM=" & nORD
    End Select

    
    nCAMPOS = 26
    aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
    For iLOOP = 0 To nCAMPOS - 4
        TXTFIELDS(iLOOP) = aVAL(iLOOP)
    Next iLOOP
    Check1.Value = aVAL(23)
    Check2.Value = aVAL(24)
    Check3.Value = aVAL(25)

    TXTFIELDS(3).Font = "isoqsymbol"
    TXTFIELDS(21).Font = "isoqsymbol"
    TXTFIELDS(22).Font = "isoqsymbol"

    If ADOPegBlob(cARQ, cSQL, Picture1, "IMAGEM") Then
        StretchSourcePictureFromPicture Picture1, Picture2
        If FixNum(eRETU01) > 500000 Then
            Alert ("Imagem Muito Grande,Ajuste o tamanho")
            salvarpict Me, Picture1, "PFC_" & StrZero(nPF, 6) & "_SEQ_" & StrZero(nSEQ, 3) & "_SSQ_" & StrZero(nSSQ, 3) & "_item_" & StrZero(nORD, 3)
            Set Picture1.Picture = Nothing
            Set Picture2.Picture = Nothing
            lTROCOU = True
        End If
       
    Else
        Set Picture1.Picture = Nothing
        Set Picture2.Picture = Nothing
    End If


End Sub

Private Sub Incluirimagem_Click()
    Dim STMPFILE
    STMPFILE = OpenArqExt(Me, "", "JPG", "JPEG *.JPG")
    If lerarquivoimagem(STMPFILE, Picture1, Picture2) Then
        lTROCOU = True
    End If
End Sub

Private Sub pfcescid_Click(Index As Integer)

    iMU01 = 100 + Index
    escIED.Show vbModal, Me

    If lRETU Then

        Select Case iMU01

        Case 101
            TXTFIELDS(1) = eRETU02

        Case 102
            TXTFIELDS(2) = eRETU02

        Case 103
            TXTFIELDS(5) = eRETU02

        Case 104
            TXTFIELDS(10) = eRETU02

        Case 105
            TXTFIELDS(13) = eRETU02

        Case 106
                
            TXTFIELDS(11) = eRETU02
                
        Case 107
                
            TXTFIELDS(7) = eRETU02
        Case 108
                
            TXTFIELDS(4) = eRETU02
                

        End Select

    End If

End Sub

Private Sub Setor_Click(Index As Integer)

    Select Case Index

    Case 0
        Me.TXTFIELDS(12) = "P"

    Case 1
        Me.TXTFIELDS(12) = "O"

    Case 2
        Me.TXTFIELDS(12) = "L"

    Case 3
        Me.TXTFIELDS(12) = "R"

    End Select

End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)

    Select Case Index

    Case 1
        
        '' KeyAscii = ValiText(KeyAscii, "#NI")

           

    Case 5
        'KeyAscii = ValiText(KeyAscii, "SN", , "#CU")
        ''keypressed = ValiText(KeyAscii, "SN", True)

    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdEditMe04_Click(Index As Integer)
    Dim cARQ As String
    Dim cCODIGO As String
    Dim cTIPO As String
    Dim cCODTIP As String
    Dim cTMP As String
    Dim cAPLICACAO As String
    Dim cCAPACIDADE As String
    Dim cARQUSO As String
    Dim cARQCDX As String
    Dim cALIAS As String
    Dim nHANDLE
    Dim lTEM As Boolean
    On Error GoTo errhandler
    If Index = 0 Then
        cCODIGO = FixStr(TXTFIELDS(16))
        cCODTIP = FixStr(TXTFIELDS(8))
        cTIPO = FixStr(TXTFIELDS(9))
    Else
        cCODIGO = FixStr(TXTFIELDS(17))
        cCODTIP = FixStr(TXTFIELDS(19))
        cTIPO = FixStr(TXTFIELDS(20))
    End If
    If Len(cCODIGO) = 0 Then
        Exit Sub
    End If

    cAPLICACAO = FixStr(frmPCX.TXTPF(12))
    cCAPACIDADE = FixStr(TXTFIELDS(4)) & FixStr(TXTFIELDS(5))

    cARQ = PegPath("PATH", "MANA5INS")
    sx_SetEpoch (Year(Date) - 30)
    sx_SetDateFormat BRITISH
    sx_SetDeleted True
    sx_SetSoftSeek True

    cARQUSO = cARQ & "ME04.DBF"
    cARQCDX = cARQ & "ME04.CDX"
    cALIAS = "ME04"
    If Not FileExist(cARQUSO, True) Then Exit Sub
    If Not FileExist(cARQCDX, True) Then Exit Sub
    nHANDLE = sx_Use(cARQUSO, cALIAS, READWRITE, SDEFOX)
    sx_SetDeleted True
    sx_SetSoftSeek True
    sx_SetOrder (1)
    sx_GoTop
    lTEM = False
    If sx_Seek(cCODIGO) Then
        If Trim(sx_EvalString(sx_IndexKey())) <> Trim(cCODIGO) Then
            sx_AppendBlank
        Else
            lTEM = True
        End If
        If sx_Rlock(sx_RecNo()) Then
            If Not lTEM Then
                sx_PutVariant "CODIGO", cCODIGO
            End If
            If Len(cTIPO) > 0 Then
                cTMP = Trim(sx_GetString("TIPO"))
                If Len(cTMP) = 0 Then
                    sx_PutVariant "TIPO", cTIPO
                End If
            End If
      
            If Len(cCODTIP) > 0 Then
                cTMP = Trim(sx_GetString("codTIPO"))
                If Len(cTMP) = 0 Then
                    sx_PutVariant "codtipo", cCODTIP
                End If
            End If
      
            If Len(cAPLICACAO) > 0 Then
                cTMP = Trim(sx_GetString("APLICACAO"))
                If Len(cTMP) = 0 Then
                    sx_PutVariant "APLICACAO", cAPLICACAO
                End If
            End If
      
            If Len(cCAPACIDADE) > 0 Then
                cTMP = Trim(sx_GetString("CAPACIDADE"))
                If Len(cTMP) = 0 Then
                    sx_PutVariant "CAPACIDADE", cCAPACIDADE
                End If
            End If

       
      
      
      
            sx_Unlock sx_RecNo()
        End If
    End If
    sx_Close

    ePASS01 = cCODIGO
    frmMe04.Show vbModal, Me
    Exit Sub
errhandler:
    Select Case Err.Number
    Case Else
        SayErro "frmPFC ME04 Editar" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cSQL
        Exit Sub
    End Select
End Sub

Function CHECKME04()
    Dim cARQ As String
    Dim cSQL As String
    Dim X As Integer
    Dim aRETU As Variant
    Dim dDATA As Date
    Dim cCODIGO As String
    CHECKME04 = True
    For X = 1 To 2
        If X = 1 Then
            cCODIGO = FixStr(TXTFIELDS(16))
        Else
            cCODIGO = FixStr(TXTFIELDS(17))
        End If
        cCODIGO = Trim(cCODIGO)
        If Len(cCODIGO) > 0 Then
            cARQ = PegPath("PATH", "MANA5INS")
            cARQ = GeraConn(cARQ, "JETFOX")
            cSQL = "select * from ME04 WHERE CODIGO='" & cCODIGO & "'"
            aRETU = PegSQL(cARQ, cSQL, 5, Array("SITUACAO", "DATAUSO", "DATAFIM", "CALPRO", "DATAEXT"), _
                           Array("C", "DC", "DC", "DC", "DC"), _
                           Array("", "", "", "", ""))
                                    
            If lRETU Then
                If aRETU(0) <> "A" Then
                    Alert ("Instrumento nao esta Ativo")
                    CHECKME04 = False
                End If
                If Not IsDate(aRETU(1)) Then
                    Alert ("data de uso nao esta preenchida")
                    CHECKME04 = False
                End If
                If IsDate(aRETU(2)) Then
                    Alert ("data fim esta preenchida")
                    CHECKME04 = False
                End If
                If IsDate(aRETU(3)) Then
                    dDATA = aRETU(3)
                End If
                If IsDate(aRETU(4)) Then
                    If aRETU(4) > dDATA Then
                        dDATA = aRETU(4)
                    End If
                End If
                If dDATA < Today() Then
                    Alert ("Data Calibracao ou Extensao Vencida")
                    CHECKME04 = False
                End If
            End If
        End If
        If Not CHECKME04 And Len(cCODIGO) > 0 Then
            If MDG("Editar Instrumento " & cCODIGO) Then
                CmdEditMe04_Click (X - 1)
            End If
        End If
    Next X
End Function
Private Sub VerImg_Click()
    frmPicViewer.Show vbModal, Me
    If lRETU Then
        If lerarquivoimagem(eRETU01, Picture1, Picture2) Then
            lTROCOU = True
        End If
    End If
End Sub
Public Sub MyPrintingRoutine()
    On Error Resume Next
    Printer.Print
    Printer.PaintPicture Picture1, 0, 0
    Printer.EndDoc
End Sub
Public Property Get Printer() As Printer
    Set Printer = vbExtra.Printer2
End Property
Public Property Set Printer(nPrinter As Printer)
    Set vbExtra.Printer2 = nPrinter
End Property

