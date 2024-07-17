VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{EA478B61-D9EC-47F6-BB21-95A533AF2251}#1.3#0"; "TabExt01.OCX"
Object = "{451B73A5-1563-45D5-A6AC-7B2B7D30B778}#1.1#0"; "BSPrin10.ocx"
Begin VB.Form frmIE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IE"
   ClientHeight    =   9045
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   11445
   Icon            =   "frmIE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   11445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin BSPrinter.PrintPreview PrintPreview1 
      Left            =   7920
      Top             =   960
      _ExtentX        =   1191
      _ExtentY        =   1191
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   375
      Index           =   25
      Left            =   4920
      MaxLength       =   24
      TabIndex        =   63
      Top             =   120
      Width           =   2715
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   0
      Left            =   2160
      TabIndex        =   9
      Top             =   240
      Width           =   1395
   End
   Begin TabExCtl.SSTabEx SSTab1 
      Height          =   7455
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   13150
      Tabs            =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabsPerRow      =   4
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Dizeres"
      Tab(0).ControlCount=   19
      Tab(0).Control(0)=   "cmdescIED(5)"
      Tab(0).Control(1)=   "cmdescIED(4)"
      Tab(0).Control(2)=   "cmdescIED(3)"
      Tab(0).Control(3)=   "cmdescIED(2)"
      Tab(0).Control(4)=   "cmdescIED(1)"
      Tab(0).Control(5)=   "TXTFIELDS(15)"
      Tab(0).Control(6)=   "TXTFIELDS(6)"
      Tab(0).Control(7)=   "TXTFIELDS(5)"
      Tab(0).Control(8)=   "TXTFIELDS(3)"
      Tab(0).Control(9)=   "TXTFIELDS(2)"
      Tab(0).Control(10)=   "TXTFIELDS(18)"
      Tab(0).Control(11)=   "TXTFIELDS(1)"
      Tab(0).Control(12)=   "lblLabels(15)"
      Tab(0).Control(13)=   "lblLabels(6)"
      Tab(0).Control(14)=   "lblLabels(5)"
      Tab(0).Control(15)=   "lblLabels(3)"
      Tab(0).Control(16)=   "lblLabels(2)"
      Tab(0).Control(17)=   "lblLabels(20)"
      Tab(0).Control(18)=   "lblLabels(1)"
      TabCaption(1)   =   "Revisao"
      Tab(1).ControlCount=   12
      Tab(1).Control(0)=   "Command1"
      Tab(1).Control(1)=   "TXTFIELDS(26)"
      Tab(1).Control(2)=   "TXTFIELDS(21)"
      Tab(1).Control(3)=   "NovaRev"
      Tab(1).Control(4)=   "CmdMotRev"
      Tab(1).Control(5)=   "TXTFIELDS(17)"
      Tab(1).Control(6)=   "gridrev"
      Tab(1).Control(7)=   "gridrevi"
      Tab(1).Control(8)=   "pegelaborador"
      Tab(1).Control(9)=   "ComMotInc"
      Tab(1).Control(10)=   "ComMotAlt"
      Tab(1).Control(11)=   "ComMotExc"
      TabCaption(2)   =   "Imagens"
      Tab(2).ControlCount=   27
      Tab(2).Control(0)=   "Picture1(2)"
      Tab(2).Control(1)=   "Picture2(2)"
      Tab(2).Control(2)=   "Picture1(1)"
      Tab(2).Control(3)=   "Picture2(1)"
      Tab(2).Control(4)=   "Picture1(0)"
      Tab(2).Control(5)=   "Picture2(0)"
      Tab(2).Control(6)=   "TXTFIELDS(23)"
      Tab(2).Control(7)=   "TXTFIELDS(22)"
      Tab(2).Control(8)=   "cmdimprimir(0)"
      Tab(2).Control(9)=   "cmdimprimir(1)"
      Tab(2).Control(10)=   "cmdimprimir(2)"
      Tab(2).Control(11)=   "CmdPaste(0)"
      Tab(2).Control(12)=   "CmdPaste(1)"
      Tab(2).Control(13)=   "CmdPaste(2)"
      Tab(2).Control(14)=   "CmdConfImp"
      Tab(2).Control(15)=   "Incluirimagem(0)"
      Tab(2).Control(16)=   "Incluirimagem(1)"
      Tab(2).Control(17)=   "Incluirimagem(2)"
      Tab(2).Control(18)=   "DelImg(0)"
      Tab(2).Control(19)=   "DelImg(1)"
      Tab(2).Control(20)=   "DelImg(2)"
      Tab(2).Control(21)=   "imgsave(0)"
      Tab(2).Control(22)=   "imgsave(1)"
      Tab(2).Control(23)=   "imgsave(2)"
      Tab(2).Control(24)=   "VerImg(0)"
      Tab(2).Control(25)=   "VerImg(1)"
      Tab(2).Control(26)=   "VerImg(2)"
      TabCaption(3)   =   "Qtde/Obs"
      Tab(3).ControlCount=   20
      Tab(3).Control(0)=   "TXTFIELDS(14)"
      Tab(3).Control(1)=   "TXTFIELDS(13)"
      Tab(3).Control(2)=   "TXTFIELDS(12)"
      Tab(3).Control(3)=   "TXTFIELDS(11)"
      Tab(3).Control(4)=   "TXTFIELDS(9)"
      Tab(3).Control(5)=   "TXTFIELDS(8)"
      Tab(3).Control(6)=   "TXTFIELDS(7)"
      Tab(3).Control(7)=   "TXTFIELDS(10)"
      Tab(3).Control(8)=   "TXTFIELDS(16)"
      Tab(3).Control(9)=   "TXTFIELDS(24)"
      Tab(3).Control(10)=   "lblLabels(14)"
      Tab(3).Control(11)=   "lblLabels(13)"
      Tab(3).Control(12)=   "lblLabels(12)"
      Tab(3).Control(13)=   "lblLabels(11)"
      Tab(3).Control(14)=   "lblLabels(10)"
      Tab(3).Control(15)=   "lblLabels(9)"
      Tab(3).Control(16)=   "lblLabels(8)"
      Tab(3).Control(17)=   "lblLabels(7)"
      Tab(3).Control(18)=   "lblLabels(4)"
      Tab(3).Control(19)=   "lblLabels(16)"
      Begin XPControls.XPButton Command1 
         Height          =   495
         Left            =   -67080
         TabIndex        =   67
         Top             =   1680
         Width           =   1335
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Distribuicao"
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
      Begin XPControls.XPButton cmdescIED 
         Height          =   195
         Index           =   5
         Left            =   -73320
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   6120
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "..."
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
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   375
         Index           =   26
         Left            =   -70440
         Locked          =   -1  'True
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   600
         Width           =   3675
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Index           =   2
         Left            =   8040
         ScaleHeight     =   435
         ScaleWidth      =   315
         TabIndex        =   61
         Top             =   3960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         Height          =   2535
         Index           =   2
         Left            =   6240
         ScaleHeight     =   2475
         ScaleWidth      =   2835
         TabIndex        =   60
         Top             =   1320
         Width           =   2895
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Index           =   1
         Left            =   5040
         ScaleHeight     =   435
         ScaleWidth      =   315
         TabIndex        =   59
         Top             =   3960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         Height          =   2535
         Index           =   1
         Left            =   3240
         ScaleHeight     =   2475
         ScaleWidth      =   2835
         TabIndex        =   58
         Top             =   1320
         Width           =   2895
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Index           =   0
         Left            =   2040
         ScaleHeight     =   435
         ScaleWidth      =   315
         TabIndex        =   57
         Top             =   3960
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         Height          =   2535
         Index           =   0
         Left            =   240
         ScaleHeight     =   2475
         ScaleWidth      =   2835
         TabIndex        =   56
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox TXTFIELDS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox TXTFIELDS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   13
         Left            =   -73800
         TabIndex        =   44
         Text            =   "0"
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox TXTFIELDS 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   -72840
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1575
         Width           =   855
      End
      Begin VB.TextBox TXTFIELDS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   11
         Left            =   -74760
         TabIndex        =   42
         Text            =   "0"
         Top             =   1575
         Width           =   855
      End
      Begin VB.TextBox TXTFIELDS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   -73800
         TabIndex        =   41
         Text            =   "0"
         Top             =   975
         Width           =   855
      End
      Begin VB.TextBox TXTFIELDS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   -74760
         TabIndex        =   40
         Text            =   "0"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox TXTFIELDS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   -74760
         TabIndex        =   39
         Text            =   "0"
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox TXTFIELDS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   10
         Left            =   -73800
         TabIndex        =   38
         Text            =   "0"
         Top             =   1575
         Width           =   855
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   615
         Index           =   16
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   37
         Top             =   2760
         Width           =   9975
      End
      Begin VB.TextBox TXTFIELDS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   24
         Left            =   -72840
         TabIndex        =   36
         Text            =   "0"
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   23
         Left            =   6240
         MaxLength       =   24
         TabIndex        =   35
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   22
         Left            =   3240
         MaxLength       =   24
         TabIndex        =   34
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   375
         Index           =   21
         Left            =   -72360
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   600
         Width           =   1275
      End
      Begin XPControls.XPButton NovaRev 
         Height          =   375
         Left            =   -74400
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   600
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Nova Revisão"
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
      Begin XPControls.XPButton CmdMotRev 
         Height          =   375
         Left            =   -67080
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1335
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Motivos"
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
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   375
         Index           =   17
         Left            =   -72960
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   600
         Width           =   375
      End
      Begin XPControls.XPButton cmdescIED 
         Height          =   195
         Index           =   4
         Left            =   -73200
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   4200
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "..."
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
      Begin XPControls.XPButton cmdescIED 
         Height          =   195
         Index           =   3
         Left            =   -73200
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "..."
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
      Begin XPControls.XPButton cmdescIED 
         Height          =   195
         Index           =   2
         Left            =   -73200
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "..."
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
      Begin XPControls.XPButton cmdescIED 
         Height          =   195
         Index           =   1
         Left            =   -73200
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1095
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "..."
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
      Begin VB.TextBox TXTFIELDS 
         Height          =   615
         Index           =   15
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   5400
         Width           =   9975
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   615
         Index           =   6
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   4440
         Width           =   9975
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   615
         Index           =   5
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   3480
         Width           =   9975
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   615
         Index           =   3
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   2520
         Width           =   9975
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   615
         Index           =   2
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   1560
         Width           =   9975
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   675
         Index           =   18
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   6360
         Width           =   10035
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   615
         Index           =   1
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   600
         Width           =   9975
      End
      Begin MSFlexGridLib.MSFlexGrid gridrev 
         Height          =   1815
         Left            =   -74640
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1200
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   3201
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrevi 
         Height          =   1575
         Left            =   -74640
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   3120
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   2778
         _Version        =   393216
      End
      Begin XPControls.XPButton cmdimprimir 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   70
         Top             =   5400
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         Picture         =   "frmIE.frx":058A
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
         Height          =   375
         Index           =   1
         Left            =   3240
         TabIndex        =   71
         Top             =   5400
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         Picture         =   "frmIE.frx":0B24
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
         Height          =   375
         Index           =   2
         Left            =   6360
         TabIndex        =   72
         Top             =   5400
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         Picture         =   "frmIE.frx":10BE
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
         Height          =   375
         Index           =   0
         Left            =   1680
         TabIndex        =   73
         Top             =   5400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Picture         =   "frmIE.frx":1658
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
         Height          =   375
         Index           =   1
         Left            =   4680
         TabIndex        =   74
         Top             =   5400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Picture         =   "frmIE.frx":1BF2
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
         Height          =   375
         Index           =   2
         Left            =   7800
         TabIndex        =   75
         Top             =   5400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Picture         =   "frmIE.frx":218C
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
      Begin XPControls.XPButton CmdConfImp 
         Height          =   435
         Left            =   9000
         TabIndex        =   76
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmIE.frx":2726
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
      Begin XPControls.XPButton pegelaborador 
         Height          =   375
         Left            =   -70920
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Picture         =   "frmIE.frx":2CC0
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
      Begin XPControls.XPButton ComMotInc 
         Height          =   435
         Left            =   -66960
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmIE.frx":325A
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
         Left            =   -66960
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmIE.frx":37F4
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
         Left            =   -66960
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmIE.frx":3D8E
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
      Begin XPControls.XPButton Incluirimagem 
         Height          =   435
         Index           =   0
         Left            =   240
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmIE.frx":4328
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
      Begin XPControls.XPButton Incluirimagem 
         Height          =   435
         Index           =   1
         Left            =   3240
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmIE.frx":47C2
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
      Begin XPControls.XPButton Incluirimagem 
         Height          =   435
         Index           =   2
         Left            =   6360
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmIE.frx":4C5C
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
         Left            =   240
         TabIndex        =   93
         TabStop         =   0   'False
         Top             =   4440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmIE.frx":50F6
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
         Left            =   3240
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   4440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmIE.frx":5590
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
         Left            =   6360
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   4440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmIE.frx":5A2A
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
      Begin XPControls.XPButton imgsave 
         Height          =   435
         Index           =   0
         Left            =   240
         TabIndex        =   96
         TabStop         =   0   'False
         Top             =   4920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmIE.frx":5EC4
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
      Begin XPControls.XPButton imgsave 
         Height          =   435
         Index           =   1
         Left            =   3240
         TabIndex        =   97
         TabStop         =   0   'False
         Top             =   4920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmIE.frx":635E
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
      Begin XPControls.XPButton imgsave 
         Height          =   435
         Index           =   2
         Left            =   6360
         TabIndex        =   98
         TabStop         =   0   'False
         Top             =   4920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmIE.frx":67F8
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
         Index           =   0
         Left            =   240
         TabIndex        =   99
         TabStop         =   0   'False
         Top             =   5880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmIE.frx":6C92
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
         Left            =   3240
         TabIndex        =   100
         TabStop         =   0   'False
         Top             =   5880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmIE.frx":712C
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
         Left            =   6360
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   5880
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmIE.frx":75C6
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
         Caption         =   "Bruto"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   14
         Left            =   -72840
         TabIndex        =   55
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tara"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   13
         Left            =   -73800
         TabIndex        =   54
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Total Liq."
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   12
         Left            =   -72840
         TabIndex        =   53
         Top             =   1335
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Unitario"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   11
         Left            =   -74760
         TabIndex        =   52
         Top             =   1335
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Caption         =   "Altura"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   10
         Left            =   -72840
         TabIndex        =   51
         Top             =   735
         Width           =   495
      End
      Begin VB.Label lblLabels 
         Caption         =   "Largura"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   9
         Left            =   -73680
         TabIndex        =   50
         Top             =   735
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Comprimento"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   8
         Left            =   -74760
         TabIndex        =   49
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Qtde Modulo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   7
         Left            =   -74760
         TabIndex        =   48
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "N Pec.Emb."
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   4
         Left            =   -73800
         TabIndex        =   47
         Top             =   1335
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Observação"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   16
         Left            =   -74760
         TabIndex        =   46
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Metodo"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   15
         Left            =   -74880
         TabIndex        =   27
         Top             =   5160
         Width           =   975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Proteção"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   6
         Left            =   -74880
         TabIndex        =   26
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Identificação"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   5
         Left            =   -74880
         TabIndex        =   25
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Fechamento"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   3
         Left            =   -74880
         TabIndex        =   24
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tipo de Embalagens"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   23
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Caption         =   "Local"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   20
         Left            =   -74880
         TabIndex        =   22
         Top             =   6120
         Width           =   495
      End
      Begin VB.Label lblLabels 
         Caption         =   "Detalhes"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox TXTPF 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   375
      Index           =   20
      Left            =   3960
      TabIndex        =   2
      Top             =   600
      Width           =   6015
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   375
      Index           =   19
      Left            =   720
      TabIndex        =   1
      Top             =   600
      Width           =   1755
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   375
      Index           =   4
      Left            =   1560
      MaxLength       =   24
      TabIndex        =   0
      Top             =   1080
      Width           =   1875
   End
   Begin XPControls.XPButton cmdClose 
      Height          =   435
      Left            =   10080
      TabIndex        =   68
      Top             =   120
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   767
      Picture         =   "frmIE.frx":7A60
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
      Left            =   10080
      TabIndex        =   69
      Top             =   720
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   767
      Picture         =   "frmIE.frx":7FFA
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
   Begin XPControls.XPButton Cmdescmr 
      Height          =   375
      Left            =   2520
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "frmIE.frx":8594
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
   Begin XPControls.XPButton Cmdbusmr 
      Height          =   375
      Left            =   2880
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "frmIE.frx":8B2E
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
   Begin XPControls.XPButton ESCpro 
      Height          =   375
      Index           =   0
      Left            =   7800
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "frmIE.frx":90C8
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
   Begin XPControls.XPButton ESCpro 
      Height          =   375
      Index           =   6
      Left            =   8880
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Picture         =   "frmIE.frx":9662
      Caption         =   "MC"
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
   Begin XPControls.XPButton ESCpro 
      Height          =   375
      Index           =   2
      Left            =   8160
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   120
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Picture         =   "frmIE.frx":9BFC
      Caption         =   "LX"
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
   Begin XPControls.XPButton ESCpro 
      Height          =   375
      Index           =   4
      Left            =   3480
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "frmIE.frx":A196
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
   Begin XPControls.XPButton ESCpro 
      Height          =   375
      Index           =   3
      Left            =   3840
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   1080
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Picture         =   "frmIE.frx":A730
      Caption         =   "LX"
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
   Begin XPControls.XPButton ESCpro 
      Height          =   375
      Index           =   7
      Left            =   4560
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   1080
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Picture         =   "frmIE.frx":ACCA
      Caption         =   "MC"
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
   Begin XPControls.XPButton ESCpro 
      Height          =   375
      Index           =   1
      Left            =   3240
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   600
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Picture         =   "frmIE.frx":B264
      Caption         =   "LX"
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
      Caption         =   "Codigo:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   21
      Left            =   120
      TabIndex        =   66
      Top             =   600
      Width           =   675
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo Cliente"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   17
      Left            =   3720
      TabIndex        =   62
      Top             =   240
      Width           =   1155
   End
   Begin VB.Label lblLabels 
      Caption         =   "PF:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   255
   End
   Begin VB.Label lblLabels 
      Caption         =   "Data"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   18
      Left            =   1560
      TabIndex        =   6
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "dd/mm/aaaa"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2280
      TabIndex        =   5
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo Integrado"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   40
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   1275
   End
End
Attribute VB_Name = "frmIE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cARQPF As String
Dim cARQIE As String
Dim cSQL As String
Dim aVAL As Variant
Dim aFOR As Variant
Dim aCAM As Variant
Dim aPAD As Variant
Dim nCAMPOS As Integer
Dim iLOOP As Integer
Dim lTROCOU As Variant
Dim aCAMIMG As Variant

Private Sub Cmdbusmr_Click()
  Dim aRETU As Variant
  Dim cARQ, cSQL, cCODIGO As String
  cCODIGO = FixStr(TXTFIELDS(19))
  If Len(cCODIGO) > 0 Then
    cARQ = GeraConn(zMANA5EMP, "JETFOX")
    cSQL = "SELECT NOME,NOM2 FROM mR01 WHERE CODIGO='" & cCODIGO & "'"
    aRETU = PegSQL(cARQ, cSQL, 2, Array("NOME", "NOM2"), Array("C", "C"), Array("", ""))
    If lRETU Then
      TXTFIELDS(20) = aRETU(0) + " " + aRETU(1)
    End If
  End If
End Sub

Private Sub cmdClose_Click()
  On Error Resume Next
  ''Data Revisao em Branco
  If IsNull(TXTFIELDS(21)) Or Len(TXTFIELDS(21)) = 0 Then
    If gridrev.Rows > 1 Then
      gridrev.Row = gridrev.Rows - 1
      gridrev.Col = 3
      TXTFIELDS(21).tEXT = gridrev
    Else
      TXTFIELDS(21) = TXTFIELDS(24)
    End If
  End If

  'Corrige revisao 1 qdo devia ser zero
  If gridrev.Rows = 1 Then
    If FixNum(TXTFIELDS(17)) > 0 Then
      TXTFIELDS(17) = 0
    End If
  End If

  'Qdto rev=0 data da revisao = elaboracao
  If FixNum(TXTFIELDS(17)) = 0 Then
    TXTFIELDS(21) = Fdata(TXTFIELDS(24))
  End If

  '    If MDG("Gravar alteraçôes") Then 'sempre gravar por causa revisao
  For iLOOP = 0 To nCAMPOS - 1
    aVAL(iLOOP) = TXTFIELDS(iLOOP)
  Next iLOOP
  GrvSQL cARQIE, cSQL, nCAMPOS, aCAM, aVAL, aFOR

  For iLOOP = 0 To 2
    If lTROCOU(iLOOP) Then
      ADOGrvBlob cARQIE, cSQL, Picture1(iLOOP), aCAMIMG(iLOOP)
    End If
  Next iLOOP



  '    End If
  Screen.MousePointer = vbDefault

  If MDG("deseja IMprimir") Then
    cTIPO = "R"
    zgrp = "PF"
    ZGRPSUB = "IE"
    eLOCALIZA = "ITA00148"
    escRPT.Show vbModal, Me
  End If
  Unload Me


End Sub

Private Sub CmdConfImp_Click()
  FrmPrintSetup.Show vbModal, Me
End Sub

Private Sub cmdescIED_Click(Index As Integer)
  iMU01 = Index
  escIED.Show vbModal, Me
  If lRETU Then
    Select Case Index
    Case 1
      TXTFIELDS(2) = TXTFIELDS(2) & " " & eRETU02
    Case 2
      TXTFIELDS(3) = TXTFIELDS(3) & " " & eRETU02
    Case 3
      TXTFIELDS(5) = TXTFIELDS(5) & " " & eRETU02
    Case 4
      TXTFIELDS(6) = TXTFIELDS(6) & " " & eRETU02
    Case 5
      TXTFIELDS(18) = TXTFIELDS(18) & " " & eRETU02
    End Select
  End If
End Sub

Private Sub Cmdescmr_Click()
  iMU01 = 4
  escmu01.Show vbModal, Me
  If lRETU Then
    TXTFIELDS(19) = eRETU01
    TXTFIELDS(20) = eRETU02
  End If
End Sub
Private Sub PrintPreview1_PrepareReport(Cancel As Boolean)
  Printer.PaintPicture Picture1(ePASS02), 0, 0
End Sub
Private Sub CmdImprimir_Click(Index As Integer)
  If Picture1(Index).Height = 0 Then
    Alert ("Sem Imagem")
  Else
    ePASS02 = Index
    PrintPreview1.ShowPreview
  End If

End Sub

Private Sub CmdMotRev_Click()
  gridrev.Col = 1
  nREV = FixInt(gridrev)
  ComMotAlt.Enabled = True
  ComMotExc.Enabled = True
  ComMotInc.Enabled = True
  Filgridrevi
End Sub

Private Sub CmdPaste_Click(Index As Integer)
  If Picture1(Index).Height = 0 Then
    Alert ("Sem Imagem")
  Else
    Clipboard.Clear
    Clipboard.SetData Picture1(Index).Image, (vbCFBitmap)


  End If

End Sub

Private Sub Command1_Click()
  If nREV = 0 Then
    Alert "Revisão Nao Escolhida"
    Exit Sub
  End If
  ePASS01 = "select * from rev WHERE PF=" & nPF & " AND TIPO='IE' AND REVISAO=" & nREV
  FrmRev.Show vbModal, Me
End Sub

Private Sub ComMotAlt_Click()
  If nREV = 0 Then
    Alert "Revisão Nao Escolhida"
    Exit Sub
  End If
  If gridrevi.Row > 0 Then
    gridrevi.Col = 3
    nREVI = FixInt(gridrevi)
    lEDITAR = True
    ePASS01 = "select * from revi WHERE PF=" & nPF & " AND TIPO='IE' AND REVISAO=" & nREV & " AND ITEM=" & nREVI
    frmREVI.Show vbModal, Me
    Filgridrevi
  End If
End Sub

Private Sub ComMotExc_Click()
  Dim cSQL As String
  If nREV = 0 Then
    Alert "Revisão Nao Escolhida"
    Exit Sub
  End If
  If gridrevi.Row > 0 Then
    gridrevi.Col = 3
    nREVI = FixInt(gridrevi)
    cSQL = "select * from revi WHERE PF=" & nPF & " AND TIPO='IE' AND REVISAO=" & nREV & " AND  ITEM=" & nREVI
    If ApagaSQLP(cARQPF, cSQL) Then
      Filgridrevi
    End If
  End If
End Sub

Private Sub ComMotInc_Click()
  Dim nITEM As Long
  If nREV = 0 Then
    Alert "Revisão Nao Escolhida"
    Exit Sub
  End If
  nITEM = 1
  If gridrevi.Rows > 1 Then
    gridrevi.Row = gridrevi.Rows - 1         ''//Posicao comeca com zero
    gridrevi.Col = 3
    nITEM = FixInt(gridrevi) + 1
  End If
  IncluiSQL cARQPF, "SELECT * FROM REVI WHERE PF=" & nPF & " AND TIPO='IE' AND REVisao=" & nREV & " AND ITEM=" & nITEM, 4, Array("PF", "REVISAO", "TIPO", "ITEM"), _
            Array(nPF, nREV, "IE", nITEM), True, False
  Filgridrevi
End Sub

Private Sub Filgridrev()
  Dim cSQL As String
  cSQL = "select PF,REVISAO,TIPO,DATA from rev WHERE PF=" & nPF & " AND TIPO='IE' ORDER BY REVISAO"
  MontaGridFast gridrev, 4, Array(600, 400, 400, 1200), Array("PF", "REV", "Tipo", "DATA"), _
                Array("PF", "REVISAO", "TIPO", "C$DATA"), cARQPF, cSQL
End Sub

Private Sub Filgridrevi()
  Dim cSQL As String
  cSQL = "select * from revI WHERE PF=" & nPF & " AND REVISAO=" & FixInt(nREV) & " AND TIPO='IE' ORDER BY ITEM"
  MontaGridFast gridrevi, 5, Array(600, 400, 400, 400, 2000), Array("PF", "REV", "Tipo", "ITEM", "Motivo"), _
                Array("PF", "REVISAO", "TIPO", "ITEM", "Motivo"), cARQPF, cSQL
End Sub

Private Sub DelImg_Click(Index As Integer)
  Set Picture1(Index).Picture = Nothing
  Set Picture2(Index).Picture = Nothing
  lTROCOU(Index) = True
End Sub

Private Sub Encerrar_Click()
  If Not MDG("Sair sem gravar") Then
    Exit Sub
  End If
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub ESCpro_Click(Index As Integer)
  ePASS01 = "MANA5"
  If Index = 1 Then
    ePASS01 = "LXEMB"
  End If
  If Index = 2 Or Index = 3 Then
    ePASS01 = "LOGIE"
  End If
  If Index = 6 Or Index = 7 Then
    ePASS01 = "MICRO"
  End If
  escms01.Show vbModal, Me
  If lRETU Then
    If Index = 0 Or Index = 6 Then
      frmIE.TXTFIELDS(25) = eRETU01
      frmIE.TXTFIELDS(20) = eRETU02
      frmIE.TXTFIELDS(24) = eRETU03
    End If
    If Index = 2 Or Index = 3 Then
      frmIE.TXTFIELDS(25) = eRETU01
      frmIE.TXTFIELDS(20) = eRETU02
      frmIE.TXTFIELDS(4) = eRETU03
    End If
    If Index = 4 Or Index = 7 Then
      frmIE.TXTFIELDS(4) = eRETU03
    End If
    If Index = 1 Then
      frmIE.TXTFIELDS(19) = eRETU01
      frmIE.TXTFIELDS(20) = eRETU02
    End If
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub Form_Load()

  CenterFormToScreen Me
  TXTPF.tEXT = nPF
  lTROCOU = Array(False, False, False)
  aCAMIMG = Array("IMAGEM", "FOTOESQ", "FOTODIR")

  cARQIE = PegPath("PATH", "IE")
  cARQPF = PegPath("PATH", "PF")

  CHECKPFPG                                    'inclui embalagem preliminar gp12

  cSQL = "select * from IE WHERE PF=" & nPF
  nCAMPOS = 27

  aCAM = Array("DATA", "DETALHE", "TIPO", "FECHA", "CODIGOINT" _
       , "IDENT", "PROTE", "QML", "COM", "LAR" _
       , "NPED", "UND", "TOTL", "TARA", "BRUTO" _
       , "METODO", "OBS", "REVISAO", "LOCAL", "CODMR01" _
       , "NOMMR01", "DATAREV", "CODESQ", "CODDIR", "ALT", "CODIGOCLI", "ELABORADOR")
  aFOR = Array("D", "C", "C", "C", "C", _
               "C", "C", "NI", "NI", "NI", _
               "NI", "N", "N", "N", "N", _
               "C", "C", "C", "C", "C", _
               "C", "D", "C", "C", "NI", "C", "C")
  aPAD = Array("", "", "", "", "", _
               "", "", 0, 0, 0, _
               0, 0, 0, 0, 0, _
               "", "", "", "", "", _
               "", "", "", "", 0, "", "")

  aVAL = PegSQL(cARQIE, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
  For iLOOP = 0 To nCAMPOS - 1
    TXTFIELDS(iLOOP) = aVAL(iLOOP)           '
  Next iLOOP
  TXTFIELDS_Change 10                          'atualiza calculo


  For iLOOP = 0 To 2
    If ADOPegBlob(cARQIE, cSQL, Picture1(iLOOP), aCAMIMG(iLOOP)) Then
      StretchSourcePictureFromPicture Picture1(iLOOP), Picture2(iLOOP)
      If FixNum(eRETU01) > 500000 Then
        Alert ("Imagem Muito Grande,Ajuste o tamanho")
        salvarpict Me, Picture1(iLOOP), aCAMIMG(iLOOP) & "_IE_" & StrZero(nPF, 6)
        Set Picture1(iLOOP).Picture = Nothing
        Set Picture2(iLOOP).Picture = Nothing
        lTROCOU(iLOOP) = True
      End If
    Else
      Set Picture1(iLOOP).Picture = Nothing
      Set Picture2(iLOOP).Picture = Nothing
    End If
  Next iLOOP

  Filgridrev


  If gridrev.Rows > 1 Then
    gridrev.Row = gridrev.Rows - 1
    CmdMotRev_Click
  End If
  PrintPreview1.AuxiliaryButtonVisible = PrintPreview1.PrinterExists("Microsoft Print to PDF")
  PrintPreview1.AuxiliaryButtonToolTipText = "Salvar como PDF"

End Sub
Public Sub PrintPreview1_AuxiliaryButtonClick(UpdateReport As Boolean)
  PrintPreview1.ShowSaveToFile "Microsoft Print to PDF", "*.pdf"
  UpdateReport = False  ' we don't need to update the report in the Print preview window after this action (the default value of UpdateReport parameter is True)
End Sub
Private Sub imgsave_Click(Index As Integer)
  Dim cARQ As String
  Select Case Index
  Case 0
    cARQ = "IE" & StrZero(FixInt(TXTFIELDS(0)), 6)
  Case 1
    cARQ = TiraOut(TXTFIELDS(22))
  Case 2
    cARQ = TiraOut(TXTFIELDS(23))
  End Select
  salvarpict Me, Picture1(Index), cARQ

End Sub

Private Sub Incluirimagem_Click(Index As Integer)
  Dim STMPFILE
  STMPFILE = OpenArqExt(Me, "", "JPG", "JPEG *.JPG")
  If lerarquivoimagem(STMPFILE, Picture1(Index), Picture2(Index)) Then
    lTROCOU(Index) = True
  End If
End Sub

Private Sub NovaRev_Click()
  Dim nREV As Integer
  If TXTFIELDS(21) > Today() Then
    Alert ("Data Revisao Anterior Maior que Data do Sistema")
    Exit Sub
  End If
  If MDG("Fazer Nova Revisao") Then
    nREV = FixNum(TXTFIELDS(17)) + 1
    IncluiSQL cARQPF, "SELECT * FROM REV WHERE PF=" & nPF & " and tipo='IE' and revisao=" & nREV, 4, Array("PF", "REVISAO", "TIPO", "DATA"), _
              Array(nPF, nREV, "IE", Today()), True, False
    TXTFIELDS(17) = nREV
    TXTFIELDS(21) = Date
    TXTFIELDS(26) = zNOMEFOLHA
    cmdClose_Click
    'Filgridrev
  End If
End Sub

Private Sub pegelaborador_Click()
  If TXTFIELDS(26) = "" Then
    TXTFIELDS(26) = zNOMEFOLHA
  End If
End Sub

Private Sub TXTFIELDS_Change(Index As Integer)
  Select Case Index
  Case 10, 11, 12, 13
    'Peso Liquido=Peso Unitario * N Peca por Embalagem
    TXTFIELDS(12).tEXT = CDbl(FixNum(TXTFIELDS(10))) * CDbl(FixNum(TXTFIELDS(11)))
    'Peso Bruto Peso Liquido + Tara
    TXTFIELDS(14).tEXT = CDbl(FixNum(TXTFIELDS(12))) + CDbl(FixNum(TXTFIELDS(13)))
  End Select
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case Index
  Case 0
    KeyAscii = ValidaDados(KeyAscii, "DL")   ''dd/mm/yyyy
  Case 8, 7, 10, 9, 24
    KeyAscii = ValiText(KeyAscii, "#NI")
  Case 11, 13
    KeyAscii = ValiText(KeyAscii, "#N")
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

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub


