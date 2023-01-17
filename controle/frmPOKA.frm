VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{EA478B61-D9EC-47F6-BB21-95A533AF2251}#1.0#0"; "TabExC01.OCX"
Object = "{451B73A5-1563-45D5-A6AC-7B2B7D30B778}#1.0#0"; "BSPrin10.ocx"
Begin VB.Form frmPOKA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MENU"
   ClientHeight    =   9000
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   11760
   Icon            =   "frmPOKA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin BSPrinter.PrintPreview PrintPreview1 
      Left            =   8640
      Top             =   720
      _ExtentX        =   1191
      _ExtentY        =   1191
      LcK1b           =   "yefT59bnyufI583n1ufV59HnyefN58nn0+fK58nnzefR58rn1+fW5w=="
      LcK2b           =   "reeI58zn+ue059jn1Of656fni+eN543nleeM54HnlueR5w=="
      AmbientBb       =   $"frmPOKA.frx":058A
   End
   Begin VB.CommandButton CmdTrocaSequencia 
      Caption         =   "TrocarSequencia"
      Height          =   375
      Left            =   7920
      TabIndex        =   55
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox item 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   54
      Text            =   "0"
      Top             =   120
      Width           =   735
   End
   Begin TabExCtl.SSTabEx SSTab1 
      Height          =   7095
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   12515
      Tabs            =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tab             =   5
      TabHeight       =   520
      TabCaption(0)   =   "Colaborador Elabarador"
      Tab(0).ControlCount=   19
      Tab(0).Control(0)=   "Command2(2)"
      Tab(0).Control(1)=   "Command2(1)"
      Tab(0).Control(2)=   "Command2(0)"
      Tab(0).Control(3)=   "TXT(11)"
      Tab(0).Control(4)=   "TXT(6)"
      Tab(0).Control(5)=   "TXT(5)"
      Tab(0).Control(6)=   "TXT(4)"
      Tab(0).Control(7)=   "TXT(3)"
      Tab(0).Control(8)=   "TXT(9)"
      Tab(0).Control(9)=   "Command7X"
      Tab(0).Control(10)=   "TXT(8)"
      Tab(0).Control(11)=   "TXT(7)"
      Tab(0).Control(12)=   "Label8"
      Tab(0).Control(13)=   "Label5"
      Tab(0).Control(14)=   "Label7"
      Tab(0).Control(15)=   "Label6"
      Tab(0).Control(16)=   "Label4"
      Tab(0).Control(17)=   "Label12"
      Tab(0).Control(18)=   "Label10"
      TabCaption(1)   =   "passo 01"
      Tab(1).ControlCount=   9
      Tab(1).Control(0)=   "TXT(10)"
      Tab(1).Control(1)=   "Picture1(0)"
      Tab(1).Control(2)=   "Picture2(0)"
      Tab(1).Control(3)=   "Incluirimagem(0)"
      Tab(1).Control(4)=   "DelImg(0)"
      Tab(1).Control(5)=   "ximgsave(0)"
      Tab(1).Control(6)=   "cmdimprimir(0)"
      Tab(1).Control(7)=   "CmdPaste(0)"
      Tab(1).Control(8)=   "VerImg(0)"
      TabCaption(2)   =   "passo 02"
      Tab(2).ControlCount=   9
      Tab(2).Control(0)=   "Picture1(1)"
      Tab(2).Control(1)=   "Picture2(1)"
      Tab(2).Control(2)=   "TXT(12)"
      Tab(2).Control(3)=   "Incluirimagem(1)"
      Tab(2).Control(4)=   "DelImg(1)"
      Tab(2).Control(5)=   "ximgsave(1)"
      Tab(2).Control(6)=   "cmdimprimir(1)"
      Tab(2).Control(7)=   "CmdPaste(1)"
      Tab(2).Control(8)=   "VerImg(1)"
      TabCaption(3)   =   "passo 03"
      Tab(3).ControlCount=   9
      Tab(3).Control(0)=   "Picture1(0)"
      Tab(3).Control(1)=   "Picture2(1)"
      Tab(3).Control(2)=   "TXT(10)"
      Tab(3).Control(3)=   "Incluirimagem(2)"
      Tab(3).Control(4)=   "DelImg(2)"
      Tab(3).Control(5)=   "ximgsave(2)"
      Tab(3).Control(6)=   "cmdimprimir(2)"
      Tab(3).Control(7)=   "CmdPaste(2)"
      Tab(3).Control(8)=   "VerImg(2)"
      TabCaption(4)   =   "passo 04"
      Tab(4).ControlCount=   9
      Tab(4).Control(0)=   "Picture1(1)"
      Tab(4).Control(1)=   "Picture2(3)"
      Tab(4).Control(2)=   "TXT(16)"
      Tab(4).Control(3)=   "Incluirimagem(3)"
      Tab(4).Control(4)=   "DelImg(3)"
      Tab(4).Control(5)=   "ximgsave(3)"
      Tab(4).Control(6)=   "cmdimprimir(3)"
      Tab(4).Control(7)=   "CmdPaste(3)"
      Tab(4).Control(8)=   "VerImg(3)"
      TabCaption(5)   =   "Revisao"
      Tab(5).ControlCount=   14
      Tab(5).Control(0)=   "ComMotExc"
      Tab(5).Control(1)=   "ComMotInc"
      Tab(5).Control(2)=   "ComMotAlt"
      Tab(5).Control(3)=   "Command6"
      Tab(5).Control(4)=   "CmdMotRev"
      Tab(5).Control(5)=   "TXT(15)"
      Tab(5).Control(6)=   "TXT(13)"
      Tab(5).Control(7)=   "NovaRev"
      Tab(5).Control(8)=   "Picture1(3)"
      Tab(5).Control(9)=   "Picture1(2)"
      Tab(5).Control(10)=   "Picture2(2)"
      Tab(5).Control(11)=   "TXT(14)"
      Tab(5).Control(12)=   "gridrev"
      Tab(5).Control(13)=   "gridrevi"
      Begin VB.CommandButton Command2 
         Caption         =   "FPC"
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   52
         Top             =   5400
         Width           =   495
      End
      Begin VB.CommandButton ComMotExc 
         Caption         =   "Excluir"
         Height          =   375
         Left            =   8040
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   4320
         Width           =   1335
      End
      Begin VB.CommandButton ComMotInc 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   8040
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   3840
         Width           =   1335
      End
      Begin VB.CommandButton ComMotAlt 
         Caption         =   "Alterar"
         Height          =   375
         Left            =   8040
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   3360
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Distribuicao"
         Height          =   495
         Left            =   8040
         TabIndex        =   48
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CommandButton CmdMotRev 
         Caption         =   "Motivos"
         Height          =   375
         Left            =   8040
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox TXT 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   15
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox TXT 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   13
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton NovaRev 
         Caption         =   "Nova Revisão"
         Height          =   375
         Left            =   240
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "FPP"
         Height          =   255
         Index           =   1
         Left            =   -74280
         TabIndex        =   41
         Top             =   5040
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "F"
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   40
         Top             =   5040
         Width           =   375
      End
      Begin VB.TextBox TXT 
         Height          =   645
         Index           =   11
         Left            =   -73680
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   38
         Top             =   4800
         Width           =   9615
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Index           =   3
         Left            =   3480
         ScaleHeight     =   435
         ScaleWidth      =   195
         TabIndex        =   36
         Top             =   6000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Index           =   2
         Left            =   3360
         ScaleHeight     =   435
         ScaleWidth      =   195
         TabIndex        =   35
         Top             =   6000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Index           =   1
         Left            =   -71520
         ScaleHeight     =   435
         ScaleWidth      =   195
         TabIndex        =   34
         Top             =   6000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   3975
         Index           =   3
         Left            =   -74880
         ScaleHeight     =   3915
         ScaleWidth      =   4155
         TabIndex        =   33
         Top             =   720
         Width           =   4215
      End
      Begin VB.PictureBox Picture2 
         Height          =   3975
         Index           =   2
         Left            =   120
         ScaleHeight     =   3915
         ScaleWidth      =   4155
         TabIndex        =   32
         Top             =   840
         Width           =   4215
      End
      Begin VB.PictureBox Picture2 
         Height          =   3975
         Index           =   1
         Left            =   -74880
         ScaleHeight     =   3915
         ScaleWidth      =   4155
         TabIndex        =   31
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox TXT 
         Height          =   1245
         Index           =   16
         Left            =   -70560
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   30
         Top             =   1320
         Width           =   6735
      End
      Begin VB.TextBox TXT 
         Height          =   1245
         Index           =   14
         Left            =   4440
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   1200
         Width           =   6735
      End
      Begin VB.TextBox TXT 
         Height          =   1245
         Index           =   12
         Left            =   -70560
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   1080
         Width           =   6735
      End
      Begin VB.TextBox TXT 
         Height          =   1245
         Index           =   10
         Left            =   -70560
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   1200
         Width           =   6735
      End
      Begin VB.TextBox TXT 
         Height          =   765
         Index           =   6
         Left            =   -73680
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   3960
         Width           =   9615
      End
      Begin VB.TextBox TXT 
         Height          =   765
         Index           =   5
         Left            =   -73680
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   22
         Top             =   3120
         Width           =   9615
      End
      Begin VB.TextBox TXT 
         Height          =   765
         Index           =   4
         Left            =   -73680
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   2280
         Width           =   9615
      End
      Begin VB.TextBox TXT 
         Height          =   765
         Index           =   3
         Left            =   -73680
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   1440
         Width           =   9615
      End
      Begin VB.TextBox TXT 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   9
         Left            =   -67320
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   900
         Width           =   1095
      End
      Begin VB.CommandButton Command7X 
         Caption         =   "..."
         Height          =   375
         Left            =   -72480
         TabIndex        =   18
         Top             =   900
         Width           =   375
      End
      Begin VB.TextBox TXT 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   8
         Left            =   -72000
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   900
         Width           =   4095
      End
      Begin VB.TextBox TXT 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   7
         Left            =   -73680
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   900
         Width           =   1095
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Index           =   0
         Left            =   -71640
         ScaleHeight     =   435
         ScaleWidth      =   195
         TabIndex        =   13
         Top             =   6000
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   3975
         Index           =   0
         Left            =   -74880
         ScaleHeight     =   3915
         ScaleWidth      =   4155
         TabIndex        =   12
         Top             =   780
         Width           =   4215
      End
      Begin MSFlexGridLib.MSFlexGrid gridrev 
         Height          =   1815
         Left            =   240
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   1440
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   3201
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid gridrevi 
         Height          =   1575
         Left            =   240
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   3360
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   2778
         _Version        =   393216
      End
      Begin XPControls.XPButton Incluirimagem 
         Height          =   435
         Index           =   0
         Left            =   -70440
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPOKA.frx":062A
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
         Left            =   -70440
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   2520
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPOKA.frx":0AC4
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
         Left            =   -70560
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPOKA.frx":0F5E
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
         Index           =   3
         Left            =   -70440
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPOKA.frx":13F8
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
         Left            =   -70440
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPOKA.frx":1892
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
         Left            =   -70440
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPOKA.frx":1D2C
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
         Left            =   -70560
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPOKA.frx":21C6
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
         Left            =   -70440
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPOKA.frx":2660
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
      Begin XPControls.XPButton ximgsave 
         Height          =   435
         Index           =   0
         Left            =   -70440
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPOKA.frx":2AFA
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
      Begin XPControls.XPButton ximgsave 
         Height          =   435
         Index           =   1
         Left            =   -70440
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   3480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPOKA.frx":2F94
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
      Begin XPControls.XPButton ximgsave 
         Height          =   435
         Index           =   2
         Left            =   -70560
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPOKA.frx":342E
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
      Begin XPControls.XPButton ximgsave 
         Height          =   435
         Index           =   3
         Left            =   -70440
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   3600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPOKA.frx":38C8
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
         Left            =   -70440
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPOKA.frx":3D62
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
         Left            =   -70440
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPOKA.frx":42FC
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
         Left            =   -70560
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPOKA.frx":4896
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
         Left            =   -70440
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPOKA.frx":4E30
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
         Left            =   -70440
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   4560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPOKA.frx":53CA
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
         Left            =   -70440
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   4440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPOKA.frx":5964
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
         Left            =   -70560
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   4560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPOKA.frx":5EFE
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
         Left            =   -70440
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   4560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPOKA.frx":6498
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
         Left            =   -70440
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   5040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmPOKA.frx":6A32
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
         Left            =   -70440
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   4920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmPOKA.frx":6ECC
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
         Left            =   -70560
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   5040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmPOKA.frx":7366
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
         Left            =   -70440
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   5040
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmPOKA.frx":7800
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
      Begin VB.Label Label8 
         Caption         =   "Obs"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   39
         Top             =   4800
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Quem?"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   37
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Onde?"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   26
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Por que?"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   25
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "O Que?"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   24
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Em"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -67800
         TabIndex        =   17
         Top             =   900
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "Elaborador"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74640
         TabIndex        =   14
         Top             =   1020
         Width           =   855
      End
   End
   Begin VB.TextBox SSQ 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "0"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox SEQ 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox PF 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   372
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   732
   End
   Begin VB.TextBox TXT 
      Height          =   330
      Index           =   2
      Left            =   840
      TabIndex        =   4
      Top             =   1080
      Width           =   4350
   End
   Begin VB.TextBox TXT 
      Height          =   315
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   3390
   End
   Begin VB.TextBox TXT 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   330
      Index           =   0
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1095
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   9960
      TabIndex        =   56
      Top             =   600
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmPOKA.frx":7C9A
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
      TabIndex        =   57
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmPOKA.frx":8234
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
      Left            =   9840
      TabIndex        =   58
      Top             =   1080
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmPOKA.frx":87CE
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
   Begin XPControls.XPButton Command5 
      Height          =   375
      Left            =   5400
      TabIndex        =   59
      Top             =   720
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "frmPOKA.frx":8D68
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
   Begin XPControls.XPButton ESCMS01A 
      Height          =   375
      Index           =   0
      Left            =   4320
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   720
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Picture         =   "frmPOKA.frx":9302
      Caption         =   "mana5"
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
      Height          =   375
      Left            =   5880
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   720
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Picture         =   "frmPOKA.frx":989C
      Caption         =   "PF"
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
   Begin XPControls.XPButton ESCMS01A 
      Height          =   375
      Index           =   1
      Left            =   6720
      TabIndex        =   86
      TabStop         =   0   'False
      Top             =   720
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Picture         =   "frmPOKA.frx":9E36
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
   Begin XPControls.XPButton ESCMS01A 
      Height          =   375
      Index           =   2
      Left            =   7560
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   720
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Picture         =   "frmPOKA.frx":A3D0
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
   Begin VB.Label Label9 
      Caption         =   "Item"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6360
      TabIndex        =   53
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "SSQ"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Seq"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "PF"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Numero:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   630
   End
End
Attribute VB_Name = "frmPOKA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cARQ As String
Dim cARQPF As String
Dim cSQL As String
Dim aVAL As Variant
Dim aFOR As Variant
Dim aCAM As Variant
Dim aPAD As Variant
Dim nCAMPOS As Integer
Dim iLOOP As Integer
Dim lTROCOU As Variant

Private Sub cmdClose_Click()
    On Error Resume Next
    Dim i As Integer

    If MDG("Gravar alteraçôes") Then
        For iLOOP = 0 To nCAMPOS - 1
            aVAL(iLOOP) = TXT(iLOOP)
        Next iLOOP
        GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR
       
        For i = 0 To 3
            If lTROCOU(i) Then
                ADOGrvBlob cARQ, cSQL, Picture1(i), "FOTOPASSO0" + StrZero(i + 1, 1)
            End If
        Next i
       
    End If
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub CmdConfImp_Click()
    FrmPrintSetup.Show vbModal, Me
End Sub
Private Sub PrintPreview1_PrepareReport(Cancel As Boolean)
   Printer.PaintPicture Picture1(ePASS02), 0, 0
   Printer.EndDoc
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

Private Sub CmdTrocaSequencia_Click()
    iMU01 = 4
    NewSeq.Show vbModal
    SEQ.Text = nSEQ
    SSQ.Text = nSSQ
    item.Text = nORD
End Sub

Private Sub Command1_Click()
    escpffim.Show vbModal, Me
    If lRETU Then
        TXT(1) = eRETU02
        TXT(2) = eRETU03
    End If
End Sub

Private Sub Command2_Click(Index As Integer)
    Select Case Index
    Case 0
        TXT(11).Text = "SIMULAR AS FALHAS COM PEÇAS"
    Case 1
        TXT(11).Text = "SIMULAR AS FALHAS COM PEÇAS DA PRODUÇÃO"
    Case 2
        TXT(11).Text = "SIMULAR AS FALHAS COM PEÇAS COELHO"
    End Select
End Sub

Private Sub Command4_Click(Index As Integer)
    salvarpict Me, Picture1(Index), StrZero(TXT(0), 8)
End Sub

Private Sub Command5_Click()
    Dim cARQ As String
    Dim aRETU As Variant
    Dim sSQL             As String
    Dim cCODIGO          As String
    cCODIGO = FixStr(TXT(1), "", "TRIM")
    cARQ = GeraConn(zMANA5EMP, "JETFOX")
    sSQL = "SELECT NOME FROM MS01 WHERE CODIGO='" & cCODIGO & "'"
    aRETU = PegSQL(cARQ, sSQL, 1, Array("NOME"), Array("C"), Array(""))
    If lRETU Then
        TXT(2) = aRETU(0)
    End If
End Sub

Private Sub Command6_Click()
    If nREV = 0 Then
        Alert "Revisão Nao Escolhida"
        Exit Sub
    End If
    ePASS01 = "select * from rev WHERE PF=" & nPPAP & " AND TIPO='POK' AND REVISAO=" & nREV
    FrmRev.Show vbModal, Me
End Sub

Private Sub Command7_Click()
    TXT(7) = zIDFOLHA
    TXT(8) = zNOMEFOLHA
    TXT(9) = Date
End Sub

Private Sub PegCodigoDescricaoPf()
    Dim sSQL, aRETU, sARQ
    sARQ = PegPath("PATH", "PF")
    sSQL = "SELECT CODIGO,DESCR FROM PF WHERE PF=" & nPF
    aRETU = PegSQL(sARQ, sSQL, 2, Array("CODIGO", "DESCR"), Array("C", "C"), Array("", ""))
    If lRETU Then
        TXT(1) = aRETU(0)
        TXT(2) = aRETU(1)
        TXT(1).Enabled = False
        TXT(2).Enabled = False
        TXT(1).Locked = True
        TXT(2).Locked = True
        ESCMS01A(0).Enabled = False
        ESCMS01A(0).Visible = False
        ESCMS01A(1).Enabled = False
        ESCMS01A(1).Visible = False
        ESCMS01A(2).Enabled = False
        ESCMS01A(2).Visible = False
        Command5.Enabled = False
        Command5.Visible = False
        Command1.Enabled = False
        Command1.Visible = False
        TXT(1).BackColor = &HC0FFFF
        TXT(2).BackColor = &HC0FFFF
    End If
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
        ePASS01 = "select * from revi WHERE PF=" & nPPAP & " AND TIPO='POK' AND REVISAO=" & nREV & " AND ITEM=" & nREVI
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
        cSQL = "select * from revi WHERE PF=" & nPPAP & " AND TIPO='POK' AND REVISAO=" & nREV & " AND  ITEM=" & nREVI
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
    IncluiSQL cARQPF, "SELECT * FROM REVI WHERE PF=" & nPPAP & " AND TIPO='POK' AND REVISAO=" & nREV & " AND ITEM=" & nITEM, 4, Array("PF", "REVISAO", "TIPO", "ITEM"), _
        Array(nPPAP, nREV, "POK", nITEM), True, False
    Filgridrevi

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

Private Sub ESCMS01A_Click(Index As Integer)
    Dim cCHAVEBUS As String
    cCHAVEBUS = TXT(1)
    ePASS01 = "MANA5"
    If Index = 1 Then
        ePASS01 = "LOGIX"
    End If
    If Index = 2 Then
        ePASS01 = "MICRO"
    End If
    escms01.Show vbModal, Me
    If lRETU Then
        TXT(1) = eRETU01
        TXT(2) = eRETU02
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
    Dim i As Integer

    CenterFormToScreen Me

    PF.Text = nPF
    SEQ.Text = nSEQ
    SSQ.Text = nSSQ
    item.Text = nORD
    
    lTROCOU = Array(False, False, False, False)
    
    
    cARQPF = PegPath("PATH", "PF")
    cARQ = PegPath("PATH", "POKA")
    cSQL = "select * from poKa WHERE numero=" & nPPAP
    nCAMPOS = 17
    aCAM = Array("NUMERO", "CODIGO", "NOME", "5SOQUE", "5SPORQUE", "5SOQUEM", "5SONDE", "ELANUM", "ELANOM", "ELADAT", _
                 "PASSO01", "OBSPASSO01", "PASSO02", "REVISAO", "PASSO03", "DATAREV", "PASSO04")
    aFOR = Array("NI", "C", "C", "C", "C", "C", "C", "NI", "C", "DZ", "C", "C", "C", "NI", "C", "DZ", "C")
    aPAD = Array(0, "", "", "", "", "", "", 0, "", "", "", "", "", 0, "", "", "")
    aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
    For iLOOP = 0 To nCAMPOS - 1
        TXT(iLOOP) = aVAL(iLOOP)
    Next iLOOP
    For i = 0 To 3
        If ADOPegBlob(cARQ, cSQL, Picture1(i), "FOTOPASSO0" + StrZero(i + 1, 1)) Then
            StretchSourcePictureFromPicture Picture1(i), Picture2(i)
           
             
            If FixNum(eRETU01) > 500000 Then
                Alert ("Imagem Muito Grande,Ajuste o tamanho")
               
                salvarpict Me, Picture1(i), "FOTOPASSO0" & StrZero(i + 1, 1) & "_PF_" & StrZero(nPF, 6) & "_SEQ_" & StrZero(nSEQ, 3) & "_SSQ_" & StrZero(nSSQ, 3)
               
               
                Set Picture1(i).Picture = Nothing
                Set Picture2(i).Picture = Nothing
                lTROCOU(i) = True
            End If
           
        Else
            Set Picture1(i).Picture = Nothing
            Set Picture2(i).Picture = Nothing
        End If
    Next i
    If nPF > 0 Then
        PegCodigoDescricaoPf
    End If
    

    
    Filgridrev
    
    
    If gridrev.Rows > 1 Then
        gridrev.Row = gridrev.Rows - 1
        CmdMotRev_Click
    End If
        
    
End Sub

Private Sub Incluirimagem_Click(Index As Integer)
    Dim STMPFILE
    STMPFILE = OpenArqExt(Me, "", "JPG", "JPEG *.JPG")
    If lerarquivoimagem(STMPFILE, Picture1(Index), Picture2(Index)) Then
        lTROCOU(Index) = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

Private Sub NovaRev_Click()
    Dim nREV As Integer
    If Len(TXT(15)) > 0 Then
        If TXT(15) > Today() Then
            Alert ("Data Revisao Anterior Maior que Data do Sistema")
            Exit Sub
        End If
    End If
    If MDG("Fazer Nova Revisao") Then
        nREV = FixNum(TXT(13)) + 1
        IncluiSQL cARQPF, "SELECT * FROM REV WHERE PF=" & nPF & "AND TIPO='POK' AND REVISAO=" & nREV, 4, Array("PF", "REVISAO", "TIPO", "DATA"), _
        Array(nPPAP, nREV, "POK", Today()), True, False
        TXT(13) = nREV
        TXT(15) = Date
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

Private Sub Filgridrev()
    Dim cSQL As String
    cSQL = "select PF,REVISAO,TIPO,DATA from rev WHERE PF=" & nPPAP & " AND TIPO='POK' ORDER BY REVISAO"
    MontaGridFast gridrev, 4, Array(600, 400, 400, 1200), Array("PF", "REV", "Tipo", "DATA"), _
        Array("PF", "REVISAO", "TIPO", "C$DATA"), cARQPF, cSQL
End Sub

Private Sub Filgridrevi()
    Dim cSQL As String
    cSQL = "select * from revI WHERE PF=" & nPPAP & " AND REVISAO=" & nREV & " AND TIPO='POK' ORDER BY ITEM"
    MontaGridFast gridrevi, 5, Array(600, 400, 400, 400, 2000), Array("PF", "REV", "Tipo", "ITEM", "Motivo"), _
        Array("PF", "REVISAO", "TIPO", "ITEM", "Motivo"), cARQPF, cSQL
End Sub

Private Sub ximgsave_Click(Index As Integer)
    salvarpict Me, Picture1(Index), StrZero(TXT(0), 8)
End Sub


