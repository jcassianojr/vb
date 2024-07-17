VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{EA478B61-D9EC-47F6-BB21-95A533AF2251}#1.3#0"; "TabExt01.OCX"
Object = "{451B73A5-1563-45D5-A6AC-7B2B7D30B778}#1.1#0"; "BSPrin10.ocx"
Begin VB.Form frmPFI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PFI"
   ClientHeight    =   7035
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   10170
   Icon            =   "frmPFI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin BSPrinter.PrintPreview PrintPreview1 
      Left            =   6960
      Top             =   120
      _ExtentX        =   1191
      _ExtentY        =   1191
   End
   Begin VB.TextBox TXTSSQ 
      BackColor       =   &H00C0FFFF&
      Height          =   405
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox TXTSEQ 
      BackColor       =   &H00C0FFFF&
      Height          =   405
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox TXTPF 
      BackColor       =   &H00C0FFFF&
      Height          =   405
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   120
      Width           =   855
   End
   Begin TabExCtl.SSTabEx SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   21
      Top             =   1320
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   9763
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
      TabCaption(0)   =   "Item"
      Tab(0).ControlCount=   34
      Tab(0).Control(0)=   "TXTFIELDS(18)"
      Tab(0).Control(1)=   "TXTFIELDS(17)"
      Tab(0).Control(2)=   "TXTFIELDS(12)"
      Tab(0).Control(3)=   "TXTFIELDS(11)"
      Tab(0).Control(4)=   "TXTFIELDS(10)"
      Tab(0).Control(5)=   "TXTFIELDS(7)"
      Tab(0).Control(6)=   "TXTFIELDS(16)"
      Tab(0).Control(7)=   "TXTFIELDS(15)"
      Tab(0).Control(8)=   "TXTFIELDS(14)"
      Tab(0).Control(9)=   "TXTFIELDS(13)"
      Tab(0).Control(10)=   "TXTFIELDS(9)"
      Tab(0).Control(11)=   "TXTFIELDS(6)"
      Tab(0).Control(12)=   "TXTFIELDS(8)"
      Tab(0).Control(13)=   "TXTFIELDS(5)"
      Tab(0).Control(14)=   "TXTFIELDS(4)"
      Tab(0).Control(15)=   "TXTFIELDS(2)"
      Tab(0).Control(16)=   "TXTFIELDS(3)"
      Tab(0).Control(17)=   "cmdpeg(13)"
      Tab(0).Control(18)=   "cmdpeg(15)"
      Tab(0).Control(19)=   "cmdpeg(14)"
      Tab(0).Control(20)=   "cmdpeg(16)"
      Tab(0).Control(21)=   "cmdpeg(17)"
      Tab(0).Control(22)=   "cmdpeg(18)"
      Tab(0).Control(23)=   "lblLabels(12)"
      Tab(0).Control(24)=   "lblLabels(10)"
      Tab(0).Control(25)=   "lblLabels(9)"
      Tab(0).Control(26)=   "lblLabels(7)"
      Tab(0).Control(27)=   "lblLabels(5)"
      Tab(0).Control(28)=   "lblLabels(11)"
      Tab(0).Control(29)=   "lblLabels(6)"
      Tab(0).Control(30)=   "lblLabels(8)"
      Tab(0).Control(31)=   "lblLabels(4)"
      Tab(0).Control(32)=   "lblLabels(2)"
      Tab(0).Control(33)=   "lblLabels(3)"
      TabCaption(1)   =   "Figura"
      Tab(1).ControlCount=   9
      Tab(1).Control(0)=   "Picture1"
      Tab(1).Control(1)=   "Picture2"
      Tab(1).Control(2)=   "cmdimprimir"
      Tab(1).Control(3)=   "CmdPaste"
      Tab(1).Control(4)=   "CmdConfImp"
      Tab(1).Control(5)=   "Incluirimagem"
      Tab(1).Control(6)=   "DelImg"
      Tab(1).Control(7)=   "VerImg"
      Tab(1).Control(8)=   "Command4"
      Begin VB.TextBox TXTFIELDS 
         Height          =   375
         Index           =   18
         Left            =   5880
         MaxLength       =   1
         TabIndex        =   18
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   375
         Index           =   17
         Left            =   5880
         MaxLength       =   1
         TabIndex        =   17
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Left            =   -74760
         ScaleHeight     =   435
         ScaleWidth      =   195
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   3960
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   3375
         Left            =   -74760
         ScaleHeight     =   3315
         ScaleWidth      =   3555
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   480
         Width           =   3615
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   765
         Index           =   12
         Left            =   1200
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   4320
         Width           =   6855
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   765
         Index           =   11
         Left            =   1200
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   3480
         Width           =   6855
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   765
         Index           =   10
         Left            =   1200
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2640
         Width           =   6855
      End
      Begin VB.TextBox TXTFIELDS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   2040
         TabIndex        =   9
         Text            =   "0"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   375
         Index           =   16
         Left            =   5040
         MaxLength       =   1
         TabIndex        =   16
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   375
         Index           =   15
         Left            =   4200
         MaxLength       =   1
         TabIndex        =   14
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   375
         Index           =   14
         Left            =   5040
         MaxLength       =   1
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   375
         Index           =   13
         Left            =   4200
         MaxLength       =   1
         TabIndex        =   13
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   9
         Left            =   3120
         TabIndex        =   8
         Text            =   "0"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox TXTFIELDS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   6
         Left            =   1440
         TabIndex        =   7
         Text            =   "0"
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox TXTFIELDS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   3120
         TabIndex        =   6
         Text            =   "0"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox TXTFIELDS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   1440
         TabIndex        =   5
         Text            =   "0"
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   4
         Left            =   1440
         MaxLength       =   78
         TabIndex        =   4
         Top             =   1200
         Width           =   5655
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   2
         Left            =   1440
         MaxLength       =   20
         TabIndex        =   2
         Top             =   480
         Width           =   3855
      End
      Begin VB.TextBox TXTFIELDS 
         Height          =   285
         Index           =   3
         Left            =   1440
         MaxLength       =   78
         TabIndex        =   3
         Top             =   840
         Width           =   5655
      End
      Begin XPControls.XPButton cmdimprimir 
         Height          =   435
         Left            =   -74280
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
         Picture         =   "frmPFI.frx":058A
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
         Left            =   -72720
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   3960
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   767
         Picture         =   "frmPFI.frx":0B24
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
         Left            =   -70200
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFI.frx":10BE
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
      Begin XPControls.XPButton cmdpeg 
         Height          =   375
         Index           =   13
         Left            =   4680
         TabIndex        =   40
         Top             =   1560
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Picture         =   "frmPFI.frx":1658
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
      Begin XPControls.XPButton cmdpeg 
         Height          =   375
         Index           =   15
         Left            =   4680
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   1920
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Picture         =   "frmPFI.frx":1BF2
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
      Begin XPControls.XPButton cmdpeg 
         Height          =   375
         Index           =   14
         Left            =   5520
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1560
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Picture         =   "frmPFI.frx":218C
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
      Begin XPControls.XPButton cmdpeg 
         Height          =   375
         Index           =   16
         Left            =   5520
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   1920
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Picture         =   "frmPFI.frx":2726
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
      Begin XPControls.XPButton Incluirimagem 
         Height          =   435
         Left            =   -70920
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFI.frx":2CC0
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
         Left            =   -70920
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFI.frx":315A
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
         Left            =   -70920
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmPFI.frx":35F4
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
      Begin XPControls.XPButton Command4 
         Height          =   435
         Left            =   -70920
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPFI.frx":3A8E
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
      Begin XPControls.XPButton cmdpeg 
         Height          =   375
         Index           =   17
         Left            =   6360
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   1560
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Picture         =   "frmPFI.frx":3F28
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
      Begin XPControls.XPButton cmdpeg 
         Height          =   375
         Index           =   18
         Left            =   6360
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   1920
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Picture         =   "frmPFI.frx":44C2
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
         Caption         =   "PorQue"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   12
         Left            =   240
         TabIndex        =   32
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Como"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   31
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Oque"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   9
         Left            =   240
         TabIndex        =   30
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tempo Movimentacao"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   29
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "taktat"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   5
         Left            =   2400
         TabIndex        =   28
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tempo Maquina"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   11
         Left            =   240
         TabIndex        =   27
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "takt"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   6
         Left            =   2400
         TabIndex        =   26
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tempo Manual"
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   8
         Left            =   240
         TabIndex        =   25
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Razao"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   24
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Norma/Detalhe"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   23
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "PontoChave"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   555
      Index           =   1
      Left            =   1080
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Width           =   7335
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   405
      Index           =   0
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   615
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   8520
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmPFI.frx":4A5C
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
      Left            =   8520
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmPFI.frx":4FF6
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
   Begin VB.Label lblLabels 
      Caption         =   "SSQ:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   15
      Left            =   3000
      TabIndex        =   52
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Caption         =   "SEQ:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   14
      Left            =   1680
      TabIndex        =   50
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "PF:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   48
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descrição:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   20
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "ORD:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   19
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmPFI"
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
  If MDG("Gravar alteraçôes") Then
    For iLOOP = 0 To nCAMPOS - 1
      aVAL(iLOOP) = TXTFIELDS(iLOOP)
    Next iLOOP
    GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR

    If lTROCOU Then
      ADOGrvBlob cARQ, cSQL, Picture1, "IMAGEM"
    End If


  End If
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub CmdConfImp_Click()
  FrmPrintSetup.Show vbModal, Me
End Sub

Private Sub CmdImprimir_Click()
  If Picture1.Height = 0 Then
    Alert ("Sem Imagem")
  Else
    PrintPreview1.ShowPreview
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

Private Sub cmdpeg_Click(Index As Integer)
  iMU01 = 404                                  ' 5S
  eRETU02 = ""                                 'Evita Cancelamento Escolha
  escIED.Show vbModal, Me
  If lRETU Then
    TXTFIELDS(Index) = Mid(eRETU02, 1, 1)
  End If
End Sub

Private Sub Command4_Click()
  salvarpict Me, Picture1
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

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  cARQ = Sdb                                   'agora sdb pois pode ser pf pfp pfg

  TXTPF.tEXT = nPF
  TXTSEQ.tEXT = nSEQ
  TXTSSQ.tEXT = nSSQ

  cSQL = "select * from PFI WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ & " AND ITEM=" & nORD

  nCAMPOS = 19
  aCAM = Array("ITEM", "DESCR", "NORMA", "PTOCHAVE", "RAZAO", "TEMPOMAN", "TEMPOMAQ", "TEMPOMOV", "TAKT", "TAKTAT", _
               "5SOQUE", "5SCOMO", "5SPORQUE", "5SSIMB01", "5SSIMB02", "5SSIMB03", "5SSIMB04", "5SSIMB05", "5SSIMB06")
  aFOR = Array("NI", "C", "C", "C", "C", "N", "N", "N", "N", "N", "C", "C", "C", "C", "C", "C", "C", "C", "C")
  aPAD = Array(0, "", "", "", "", 0, 0, 0, 0, 0, "", "", "", "", "", "", "", "", "")
  aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
  For iLOOP = 0 To nCAMPOS - 1
    TXTFIELDS(iLOOP) = aVAL(iLOOP)
  Next iLOOP


  If ADOPegBlob(cARQ, cSQL, Picture1, "IMAGEM") Then
    StretchSourcePictureFromPicture Picture1, Picture2
    If FixNum(eRETU01) > 500000 Then
      Alert ("Imagem Muito Grande,Ajuste o tamanho")
      salvarpict Me, Picture1, "PFI_" & StrZero(nPF, 6) & "_SEQ_" & StrZero(nSEQ, 3) & "_SSQ_" & StrZero(nSSQ, 3) & "_item_" & StrZero(nORD, 3)
      Set Picture1.Picture = Nothing
      Set Picture2.Picture = Nothing
      lTROCOU = True
    End If
  Else
    Set Picture1.Picture = Nothing
    Set Picture2.Picture = Nothing
  End If
  PrintPreview1.AuxiliaryButtonVisible = PrintPreview1.PrinterExists("Microsoft Print to PDF")
  PrintPreview1.AuxiliaryButtonToolTipText = "Salvar como PDF"
End Sub
Public Sub PrintPreview1_AuxiliaryButtonClick(UpdateReport As Boolean)
  PrintPreview1.ShowSaveToFile "Microsoft Print to PDF", "*.pdf"
  UpdateReport = False  ' we don't need to update the report in the Print preview window after this action (the default value of UpdateReport parameter is True)
End Sub
Private Sub Incluirimagem_Click()
  Dim STMPFILE
  STMPFILE = OpenArqExt(Me, "", "JPG", "JPEG *.JPG")
  If lerarquivoimagem(STMPFILE, Picture1, Picture2) Then
    lTROCOU = True
  End If
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case Index
  Case 0
    KeyAscii = ValiText(KeyAscii, "#NI")
  Case 5, 6, 7, 8, 9
    KeyAscii = ValiText(KeyAscii, "#N")
  End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub VerImg_Click()
  frmPicViewer.Show vbModal, Me
  If lRETU Then
    If lerarquivoimagem(eRETU01, Picture1, Picture2) Then
      lTROCOU = True
    End If
  End If
End Sub
Public Sub PrintPreview1_PrepareReport(Cancel As Boolean)
  On Error Resume Next
  Printer.Print
  Printer.PaintPicture Picture1, 0, 0
End Sub

