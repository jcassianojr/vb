VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{EA478B61-D9EC-47F6-BB21-95A533AF2251}#1.3#0"; "TabExt01.OCX"
Object = "{451B73A5-1563-45D5-A6AC-7B2B7D30B778}#1.1#0"; "BSPrin10.ocx"
Begin VB.Form frmPCS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PFS"
   ClientHeight    =   7920
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   10920
   Icon            =   "FrmPcs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin BSPrinter.PrintPreview PrintPreview1 
      Left            =   9840
      Top             =   1200
      _ExtentX        =   1191
      _ExtentY        =   1191
   End
   Begin VB.TextBox TXTSSQ 
      BackColor       =   &H00C0FFFF&
      Height          =   405
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox TXTSEQ 
      BackColor       =   &H00C0FFFF&
      Height          =   405
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox TXTPF 
      BackColor       =   &H00C0FFFF&
      Height          =   405
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox TxtCodComp 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   372
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   405
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2295
   End
   Begin VB.CommandButton cmdMarPlr 
      Caption         =   "Marcar Planta"
      Height          =   315
      Index           =   2
      Left            =   2520
      TabIndex        =   5
      Top             =   1200
      Width           =   1515
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   525
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   70
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Width           =   6495
   End
   Begin TabExCtl.SSTabEx SSTab1 
      Height          =   5655
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   9975
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
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Plano Controle"
      Tab(0).ControlCount=   15
      Tab(0).Control(0)=   "TxtItem"
      Tab(0).Control(1)=   "CmdControleDispo"
      Tab(0).Control(2)=   "Command6"
      Tab(0).Control(3)=   "Command5"
      Tab(0).Control(4)=   "Command1"
      Tab(0).Control(5)=   "Command3"
      Tab(0).Control(6)=   "Command2"
      Tab(0).Control(7)=   "Duplicar(10)"
      Tab(0).Control(8)=   "Duplicar(0)"
      Tab(0).Control(9)=   "Command4(0)"
      Tab(0).Control(10)=   "Command9(0)"
      Tab(0).Control(11)=   "Command8(0)"
      Tab(0).Control(12)=   "txtFields(2)"
      Tab(0).Control(13)=   "txtFields(1)"
      Tab(0).Control(14)=   "GridPla(0)"
      TabCaption(1)   =   "Croqui"
      Tab(1).ControlCount=   8
      Tab(1).Control(0)=   "Picture1(0)"
      Tab(1).Control(1)=   "Picture2(0)"
      Tab(1).Control(2)=   "cmdimprimir(0)"
      Tab(1).Control(3)=   "CmdPaste(0)"
      Tab(1).Control(4)=   "Incluirimagem(0)"
      Tab(1).Control(5)=   "DelImg(0)"
      Tab(1).Control(6)=   "imgsave(0)"
      Tab(1).Control(7)=   "VerImg(0)"
      TabCaption(2)   =   "Controle Dispositivo"
      Tab(2).ControlCount=   6
      Tab(2).Control(0)=   "Duplicar(13)"
      Tab(2).Control(1)=   "Duplicar(3)"
      Tab(2).Control(2)=   "Command4(3)"
      Tab(2).Control(3)=   "Command9(3)"
      Tab(2).Control(4)=   "Command8(3)"
      Tab(2).Control(5)=   "GridPla(3)"
      TabCaption(3)   =   "On the Job"
      Tab(3).ControlCount=   6
      Tab(3).Control(0)=   "Command4(2)"
      Tab(3).Control(1)=   "Command8(2)"
      Tab(3).Control(2)=   "Duplicar(12)"
      Tab(3).Control(3)=   "Duplicar(2)"
      Tab(3).Control(4)=   "Command4(1)"
      Tab(3).Control(5)=   "GridPla(2)"
      TabCaption(4)   =   "Checklist"
      Tab(4).ControlCount=   5
      Tab(4).Control(0)=   "Duplicar(11)"
      Tab(4).Control(1)=   "Duplicar(1)"
      Tab(4).Control(2)=   "Command9(1)"
      Tab(4).Control(3)=   "Command8(1)"
      Tab(4).Control(4)=   "GridPla(1)"
      TabCaption(5)   =   "Alerta"
      Tab(5).ControlCount=   8
      Tab(5).Control(0)=   "Picture1(1)"
      Tab(5).Control(1)=   "Picture2(1)"
      Tab(5).Control(2)=   "cmdimprimir(1)"
      Tab(5).Control(3)=   "CmdPaste(1)"
      Tab(5).Control(4)=   "Incluirimagem(1)"
      Tab(5).Control(5)=   "DelImg(1)"
      Tab(5).Control(6)=   "imgsave(1)"
      Tab(5).Control(7)=   "VerImg(1)"
      Begin VB.CommandButton Command4 
         Caption         =   "Editar"
         Height          =   375
         Index           =   2
         Left            =   -66960
         TabIndex        =   62
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Novo"
         Height          =   375
         Index           =   2
         Left            =   -66960
         TabIndex        =   61
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox TxtItem 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   495
         Left            =   4680
         TabIndex        =   0
         Text            =   "1"
         Top             =   4680
         Width           =   495
      End
      Begin VB.CommandButton CmdControleDispo 
         Caption         =   "ControleDispositivo"
         Height          =   495
         Left            =   5400
         TabIndex        =   1
         Top             =   4680
         Width           =   1935
      End
      Begin VB.CommandButton Duplicar 
         Caption         =   "Reordenar"
         Height          =   375
         Index           =   13
         Left            =   -66840
         TabIndex        =   2
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton Duplicar 
         Caption         =   "Reordenar"
         Height          =   375
         Index           =   12
         Left            =   -66960
         TabIndex        =   3
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CommandButton Duplicar 
         Caption         =   "Reordenar"
         Height          =   375
         Index           =   11
         Left            =   -66960
         TabIndex        =   6
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton Duplicar 
         Caption         =   "Duplicar"
         Height          =   375
         Index           =   3
         Left            =   -66840
         TabIndex        =   37
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton Duplicar 
         Caption         =   "Duplicar"
         Height          =   375
         Index           =   2
         Left            =   -66960
         TabIndex        =   36
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton Duplicar 
         Caption         =   "Duplicar"
         Height          =   375
         Index           =   1
         Left            =   -66960
         TabIndex        =   35
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Editar"
         Height          =   375
         Index           =   3
         Left            =   -66840
         TabIndex        =   34
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Editar"
         Height          =   375
         Index           =   1
         Left            =   -66960
         TabIndex        =   33
         Top             =   1800
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Apaga"
         Height          =   375
         Index           =   3
         Left            =   -66840
         TabIndex        =   32
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Apaga"
         Height          =   375
         Index           =   1
         Left            =   -66960
         TabIndex        =   31
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Novo"
         Height          =   375
         Index           =   3
         Left            =   -66840
         TabIndex        =   27
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Novo"
         Height          =   375
         Index           =   1
         Left            =   -66960
         TabIndex        =   26
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Poka Yoka"
         Height          =   495
         Left            =   2520
         TabIndex        =   25
         Top             =   4680
         Width           =   1935
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Olhos Abertos"
         Height          =   495
         Left            =   120
         TabIndex        =   24
         Top             =   4680
         Width           =   2175
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Index           =   1
         Left            =   -71880
         ScaleHeight     =   435
         ScaleWidth      =   315
         TabIndex        =   23
         Top             =   3480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Index           =   0
         Left            =   -71520
         ScaleHeight     =   435
         ScaleWidth      =   315
         TabIndex        =   22
         Top             =   2520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox Picture2 
         Height          =   2535
         Index           =   1
         Left            =   -74880
         ScaleHeight     =   2475
         ScaleWidth      =   2835
         TabIndex        =   21
         Top             =   840
         Width           =   2895
      End
      Begin VB.PictureBox Picture2 
         Height          =   2535
         Index           =   0
         Left            =   -74760
         ScaleHeight     =   2475
         ScaleWidth      =   2835
         TabIndex        =   20
         Top             =   840
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         Caption         =   "TIPO"
         Height          =   495
         Left            =   2640
         TabIndex        =   19
         Top             =   3960
         Width           =   735
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Importar Componente"
         Height          =   555
         Left            =   8520
         TabIndex        =   18
         Top             =   3360
         Width           =   1035
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Importar SEQ/SSQ"
         Height          =   552
         Left            =   8520
         TabIndex        =   17
         Top             =   2760
         Width           =   1035
      End
      Begin VB.CommandButton Duplicar 
         Caption         =   "Reordenar"
         Height          =   375
         Index           =   10
         Left            =   8520
         TabIndex        =   16
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton Duplicar 
         Caption         =   "Duplicar"
         Height          =   375
         Index           =   0
         Left            =   8520
         TabIndex        =   15
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Editar"
         Height          =   375
         Index           =   0
         Left            =   8520
         TabIndex        =   14
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Apaga"
         Height          =   375
         Index           =   0
         Left            =   8520
         TabIndex        =   13
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Novo"
         Height          =   375
         Index           =   0
         Left            =   8520
         TabIndex        =   12
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   2
         Left            =   3720
         TabIndex        =   11
         Top             =   4080
         Width           =   5175
      End
      Begin VB.TextBox txtFields 
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   4080
         Width           =   2415
      End
      Begin MSFlexGridLib.MSFlexGrid GridPla 
         Height          =   3135
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   5530
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridPla 
         Height          =   2895
         Index           =   1
         Left            =   -74760
         TabIndex        =   28
         Top             =   840
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5106
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridPla 
         Height          =   2895
         Index           =   2
         Left            =   -74880
         TabIndex        =   29
         Top             =   720
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5106
         _Version        =   393216
      End
      Begin MSFlexGridLib.MSFlexGrid GridPla 
         Height          =   2895
         Index           =   3
         Left            =   -74880
         TabIndex        =   30
         Top             =   720
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   5106
         _Version        =   393216
      End
      Begin XPControls.XPButton cmdimprimir 
         Height          =   375
         Index           =   0
         Left            =   -74640
         TabIndex        =   40
         Top             =   3600
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         Picture         =   "FrmPcs.frx":058A
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
         Left            =   -74760
         TabIndex        =   41
         Top             =   3480
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   661
         Picture         =   "FrmPcs.frx":0B24
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
         Index           =   1
         Left            =   -73320
         TabIndex        =   42
         Top             =   3480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Picture         =   "FrmPcs.frx":10BE
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
         Index           =   0
         Left            =   -73200
         TabIndex        =   43
         Top             =   3600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         Picture         =   "FrmPcs.frx":1658
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
      Begin XPControls.XPButton Incluirimagem 
         Height          =   435
         Index           =   0
         Left            =   -71640
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "FrmPcs.frx":1BF2
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
         Left            =   -71760
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "FrmPcs.frx":208C
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
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "FrmPcs.frx":2526
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
         Left            =   -71760
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "FrmPcs.frx":29C0
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
         Left            =   -71640
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "FrmPcs.frx":2E5A
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
         Left            =   -71760
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "FrmPcs.frx":32F4
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
         Left            =   -71640
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   3360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "FrmPcs.frx":378E
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
         Left            =   -71760
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "FrmPcs.frx":3C28
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
   End
   Begin XPControls.XPButton cmdClose 
      Height          =   435
      Left            =   9240
      TabIndex        =   38
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FrmPcs.frx":40C2
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
      Left            =   9240
      TabIndex        =   39
      Top             =   720
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FrmPcs.frx":465C
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
   Begin XPControls.XPButton CmdConfImp 
      Height          =   435
      Left            =   7440
      TabIndex        =   44
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "FrmPcs.frx":4BF6
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
   Begin VB.Label lblLabels 
      Caption         =   "SSQ:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   15
      Left            =   2880
      TabIndex        =   59
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Caption         =   "SEQ:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   14
      Left            =   1560
      TabIndex        =   57
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "PF:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   13
      Left            =   120
      TabIndex        =   55
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label4 
      Caption         =   "Componente"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   4320
      TabIndex        =   53
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "FRMPCS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aVAL As Variant
Dim aFOR As Variant
Dim aCAM As Variant
Dim aPAD As Variant
Dim nCAMPOS As Integer
Dim iLOOP As Integer
Dim cARQPF As String
Dim cARQPFP As String
Dim cARQPFG As String
Dim cARQUSO As String
Dim cSQL As String
Dim lTROCOU As Variant

Private Sub cmdClose_Click()
  Dim cSQLIMG As String
  Dim cCAMIMG As String
  On Error Resume Next
  txtFields(1).tEXT = Replace(txtFields(1).tEXT, "/", "-")  'chave apresentando erro de gravacao

  If MDG("Gravar alteraçôes") Then
    'campos
    For iLOOP = 0 To nCAMPOS - 1
      aVAL(iLOOP) = txtFields(iLOOP)
    Next iLOOP
    GrvSQL cARQUSO, cSQL, nCAMPOS, aCAM, aVAL, aFOR
    'imagens
    For iLOOP = 0 To 1
      If lTROCOU(iLOOP) Then
        Select Case iLOOP
        Case 0
          cCAMIMG = "IMGCROQUI"        'croqui
        Case 1
          cCAMIMG = "IMGALERTA"        'alerta
        End Select
        'cSQLIMG = "select PF,SEQ,SSQ," & cCAMIMG & " from PFS WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ
        ADOGrvBlob cARQUSO, "PFS", Picture1(iLOOP), cCAMIMG, "PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ
      End If
    Next iLOOP
  End If
  Screen.MousePointer = vbDefault
  Unload Me
End Sub
Private Sub CHECKDISPO()
  Dim cARQDISPO
  cARQDISPO = PegPath("PATH", "DISPO")
  nPPAP = PegUltSQL(cARQDISPO, "select numero from DISPO WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ, "NUMERO", 0)
  If nPPAP = 0 Then
    nPPAP = FixInt(PegMAXSQL(cARQDISPO, "DISPO", "NUMERO", 0)) + 1
    IncluiSQL cARQDISPO, "SELECT * FROM DISPO WHERE NUMERO=" & nPPAP, 4, Array("NUMERO", "PF", "SEQ", "SSQ"), Array(nPPAP, nPF, nSEQ, nSSQ), True, False
  End If
End Sub

Private Sub CmdConfImp_Click()
 ' FrmPrintSetup.Show vbModal, Me
End Sub

Private Sub CmdControleDispo_Click()
  frmDISPO.Show vbModal
End Sub

Private Sub CmdImprimir_Click(Index As Integer)
  If Picture1(Index).Height = 0 Then
    Alert ("Sem Imagem")
  Else
    ePASS01 = Index
    PrintPreview1.ShowPreview
  End If
End Sub

Private Sub cmdMarPlr_Click(Index As Integer)
  Dim cFILIAL As String
  Dim sSQL As String
  cFILIAL = InputBox("Digite a Filial", "Inclusão Todas Seq", "1")
  cFILIAL = FixStr(cFILIAL, "1", "TRIM", 5)
  Select Case nARQPCS
  Case 0
    sSQL = "select PF,SEQ,SSQ,item,FILIAL from PFC WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ & " ORDER BY ITEM"
  Case 1
    sSQL = "select PF,SEQ,SSQ,item,FILIAL from PFCO WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ & " ORDER BY ITEM"
  Case 2
    sSQL = "select PF,SEQ,SSQ,item,FILIAL from PFCMS03 WHERE PF=" & nPF & " AND CODCOMP='" & Ccodcomp & "' ORDER BY ITEM"
  End Select
  GrvSQL Sdb, sSQL, 1, Array("FILIAL"), Array(cFILIAL), Array("")
  FilRelat (Index)
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
  lRETU = False
  eRETU01 = ""
  eRETU02 = ""
  ESCPCT.Show vbModal, Me
  If lRETU Then
    FRMPCS.txtFields(1) = eRETU02
  End If
End Sub

Private Sub Command2_Click()
  Select Case nARQPCS
  Case 0
    cARQIMP = "PFC"
  Case 1
    cARQIMP = "PFCO"
  Case 2
    cARQIMP = "PFCMS03"
  End Select
  FrmImp3.Show vbModal, Me
  FilRelat 0
End Sub

Private Sub Command3_Click()
  Select Case nARQPCS
  Case 0
    cARQIMP = "PFC"
  Case 1
    cARQIMP = "PFCO"
  Case 2
    cARQIMP = "PFCMS03"
  End Select
  frmIMP4.Show vbModal, Me
  FilRelat 0
End Sub

Private Sub Command4_Click(Index As Integer)
  Dim sSQL, cDESCRICAO, aVALTMP As Variant
  If GridPla(Index).Row > 0 Then
    GridPla(Index).Col = 0
    nORD = GridPla(Index)
    Select Case Index
    Case 0
      frmPFC.Show vbModal, Me
    Case 1, 2
      sSQL = SUBGERASQL(Index)
      aVALTMP = PegSQL(Sdb, sSQL, 1, Array("DESCRICAO"), Array("C"), Array(""))
      cDESCRICAO = aVALTMP(0)
      cDESCRICAO = InputBox("Digite a Descricao", "Descricao", cDESCRICAO)
      cDESCRICAO = FixStr(cDESCRICAO, "1", "TRIM", 255)
      GrvSQL Sdb, sSQL, 1, Array("DESCRICAO"), Array(cDESCRICAO), Array("C")
    Case 3
      frmDISPOI.Show vbModal, Me
    End Select
    FilRelat (Index)
  End If
End Sub

Private Sub Command5_Click()
  Dim cARQPOA
  cARQPOA = PegPath("PATH", "POA")
  nPPAP = PegUltSQL(cARQPOA, "select numero from POA WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ, "NUMERO", 0)
  If nPPAP = 0 Then
    nPPAP = FixInt(PegMAXSQL(cARQPOA, "POA", "NUMERO", 0)) + 1
    IncluiSQL cARQPOA, "SELECT * FROM POA WHERE NUMERO=" & nPPAP, 4, Array("NUMERO", "PF", "SEQ", "SSQ"), Array(nPPAP, nPF, nSEQ, nSSQ)
  End If
  frmPOA.Show vbModal
End Sub

Private Sub Command6_Click()
  Dim cARQPOKA

  If Not AcessaMenu("escPOKA", 2) Then
    Exit Sub
  End If

  cARQPOKA = PegPath("PATH", "POKA")
  nORD = TxtItem.tEXT
  nPPAP = PegUltSQL(cARQPOKA, "select numero from POKA WHERE PF=" & FixInt(nPF) & " AND SEQ=" & FixInt(nSEQ) & " AND SSQ=" & FixInt(nSSQ) & " AND ITEM=" & FixInt(nORD), "NUMERO", 0)
  If nPPAP = 0 Then
    nPPAP = FixInt(PegMAXSQL(cARQPOKA, "POKA", "NUMERO", 0)) + 1
    IncluiSQL cARQPOKA, "SELECT * FROM POKA WHERE NUMERO=" & nPPAP, 5, Array("NUMERO", "PF", "SEQ", "SSQ", "item"), Array(nPPAP, nPF, nSEQ, nSSQ, nORD)
  End If
  frmPOKA.Show vbModal
End Sub

Private Sub subgerimu(Index As Integer)
  Select Case Index
  Case 0, 10
    Select Case nARQPCS
    Case 0
      iMU01 = 3
    Case 1
      iMU01 = 4
    Case 2
      iMU01 = 5
    End Select
  Case 1, 11
    iMU01 = 7
  Case 2, 12
    iMU01 = 8
  Case 3, 13
    iMU01 = 9
  End Select
End Sub

Private Function SUBGERASQL(Index As Integer)
  Dim sSQL As String
  Select Case Index
  Case 0, 10
    Select Case nARQPCS
    Case 0
      sSQL = "select * from PFC WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ & " AND ITEM=" & nORD
    Case 1
      sSQL = "select * from PFCO WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ & " AND ITEM=" & nORD
    Case 2
      sSQL = "select * from PFCMS03 WHERE PF=" & nPF & " AND CODCOMP='" & Ccodcomp & "' AND ITEM=" & nORD
    End Select
  Case 1, 11
    sSQL = "select * from PFSCHECKLIST WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ & " AND ORDEM=" & nORD
  Case 2, 12
    sSQL = "select * from PFSONTHEJOB WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ & " AND ORDEM=" & nORD
  Case 3, 13
    sSQL = "select * from DISPOI WHERE NUMERO=" & nPPAP & " AND ITEM=" & nORD
  End Select
  SUBGERASQL = sSQL
End Function

Private Sub Command8_Click(Index As Integer)
  Dim nGRIDINDEX
  nNOVAPOS = FixInt(PegUltGrid(GridPla(Index), 0), 0) + 1
  subgerimu Index
  NewOrd.Show vbModal, Me
  FilRelat (Index)
End Sub

Private Sub Command9_Click(Index As Integer)
  Dim sSQL As String
  Dim cSQL As String
  If GridPla(Index).Row > 0 Then
    GridPla(Index).Col = 0
    nORD = FixInt(GridPla(Index))
    sSQL = SUBGERASQL(Index)
    cSQL = Sdb
    If Index = 3 Then
      cSQL = PegPath("PATH", "DISPO")
    End If
    If ApagaSQLP(cSQL, sSQL) Then
      FilRelat (Index)
    End If
  End If
End Sub

Private Sub DelImg_Click(Index As Integer)
  Set Picture1(Index).Picture = Nothing
  Set Picture2(Index).Picture = Nothing
  lTROCOU(Index) = True
End Sub

Private Sub Duplicar_Click(Index As Integer)
  Dim aVALTMP As Variant
  Dim aCAMTMP As Variant
  Dim aFORTMP As Variant
  Dim aPADTMP As Variant
  Dim sSQL As String
  Dim nCAMPOS As Integer
  Dim cARQ As String
  Dim NINDEXREF As Integer


  cARQ = Sdb
  If Index = 3 Or Index = 13 Then
    cARQ = PegPath("PATH", "DISPO")
  End If
  NINDEXREF = Index
  If NINDEXREF > 10 Then
    NINDEXREF = Index - 10
  End If


  GridPla(NINDEXREF).Col = 0
  nORD = FixInt(GridPla(NINDEXREF))
  If nORD = 0 Then
    Alert ("Clique no Item de Origem Antes")
    Exit Sub
  End If

  nNOVAPOS = FixInt(PegUltGrid(GridPla(NINDEXREF), 0), 0) + 1
  nNOVAPOS = FixInt(Val(Busca("No", "Renumerando", CStr(nNOVAPOS), 8)))
  If nNOVAPOS = 0 Then
    Alert ("No em Branco")
    Exit Sub
  End If
  LocalizaGrid GridPla(NINDEXREF), nNOVAPOS, , False
  If lRETU Then
    Alert "sequencia ja existe"
    Exit Sub
  End If

  sSQL = SUBGERASQL(Index)
  Select Case Index
  Case 0, 10                                   ' plano de controle
    nCAMPOS = 24
    aCAMTMP = Array("COP", "DEscr", "CARAC", "SIGI", "ESPE", "TOL", "cqtde", "freq", _
                    "TIPINS", "InStr", "CAPA", "REACAO", "Setor", "CARTA", "CITEM", "PROCESSO", _
                    "CODME04", "CODME04B", "FILIAL", "TIPINSB", "INSTRB", "CERTFOR", "SAIPROC", "SAIRI")

    aPADTMP = Array("", "", "", "", "", "", "", "", _
                    "", "", "", "", "", "", "", "", _
                    "", "", "", "", "", False, False, False)

    aFORTMP = Array("C", "C", "C", "C", "C", "C", "C", "C", _
                    "C", "C", "C", "C", "C", "C", "C", "C", _
                    "C", "C", "C", "C", "C", "B", "B", "B")
  Case 1, 11, 2, 12, 13                        'checklist on the job
    nCAMPOS = 1
    aCAMTMP = Array("DESCRICAO")
    aPADTMP = Array("")
    aFORTMP = Array("C")
  End Select

  aVALTMP = PegSQL(cARQ, sSQL, nCAMPOS, aCAMTMP, aFORTMP, aPADTMP)

  eRETU01 = 0
  subgerimu Index
  NewOrd.Show vbModal, Me


  If Index >= 10 Then                          'reordenacao apaga
    ApagaSQL cARQ, sSQL
  End If

  nORD = eRETU01
  sSQL = SUBGERASQL(Index)

  GrvSQL cARQ, sSQL, nCAMPOS, aCAMTMP, aVALTMP, aFORTMP
  FilRelat (Index)

End Sub

Private Sub FilRelat(Index As Integer)
  On Error Resume Next
  Dim cSQL As String
  Select Case Index
  Case 0                                       'item
    Select Case nARQPCS
    Case 0
      cSQL = "select ITEM,CITEM, FILIAL, SETOR, DESCR from PFC WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ & " ORDER BY ITEM"
    Case 1
      cSQL = "select ITEM,CITEM, FILIAL, SETOR, DESCR from PFCO WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ & " ORDER BY ITEM"
    Case 2
      cSQL = "select ITEM,CITEM, FILIAL, SETOR, DESCR from PFCMS03 WHERE PF=" & nPF & " AND CODCOMP='" & Ccodcomp & "' ORDER BY ITEM"
    End Select
    MontaGridFast GridPla(Index), 5, Array(400, 600, 600, 400, 4000), Array("Ord", "Item", "Filial", "Setor", "Descricao"), _
                  Array("L$ITEM", "L$CITEM", "L$FILIAL", "L$SETOR", "L$DESCR"), Sdb, cSQL
  Case 1                                       'checklist
    cSQL = "select ORDEM,DESCRICAO from PFSCHECKLIST WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ & " ORDER BY ORDEM"
    MontaGridFast GridPla(Index), 2, Array(400, 4000), Array("Ord", "Descricao"), _
                  Array("L$ORDEM", "L$DESCRICAO"), Sdb, cSQL
  Case 2                                       'on the job
    cSQL = "select ORDEM,DESCRICAO from PFSONTHEJOB  WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ & " ORDER BY ORDEM"
    MontaGridFast GridPla(Index), 2, Array(400, 4000), Array("Ord", "Descricao"), _
                  Array("L$ORDEM", "L$DESCRICAO"), Sdb, cSQL
  Case 3                                       'dispositivo
    cSQL = "select ITEM,DESCRICAO from DISPOI  WHERE NUMERO=" & nPPAP & " ORDER BY ITEM"
    MontaGridFast GridPla(Index), 2, Array(400, 4000), Array("Ord", "Descricao"), _
                  Array("L$ITEM", "L$DESCRICAO"), PegPath("PATH", "DISPO"), cSQL

  End Select
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
  Dim cCAMIMG As String
  Dim cSQLIMG As String
  CenterFormToScreen Me

  TXTPF.tEXT = nPF
  TXTSEQ.tEXT = nSEQ
  TXTSSQ.tEXT = nSSQ


  lTROCOU = Array(False, False, False, False, False)
  cARQPF = PegPath("PATH", "PF")
  cARQPFG = PegPath("PATH", "PFG")
  cARQPFP = PegPath("PATH", "PFP")
  cARQUSO = cARQPF
  If nSEQ = 99 And nSSQ = 99 Then
    If InStr(Sdb, "PFG.MDB") Then
      cARQUSO = cARQPFG
    End If
    If InStr(Sdb, "PFP.MDB") Then
      cARQUSO = cARQPFP
    End If
  End If
  If nARQPCS = 2 Then
    TxtCodComp.tEXT = Ccodcomp
  Else
    TxtCodComp.Visible = False
    Label4.Visible = False
  End If

  Select Case nARQPCS
  Case 0
    cSQL = "select PF,SEQ,SSQ,DESCRI,PCTIPO,PCOBS,CODINT from PFS WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ
  Case 1
    cSQL = "select PF,SEQ,SSQ,DESCRI,PCTIPO,PCOBS,CODINT from PFSO WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ
  Case 2
    cSQL = "select PF,codcomp,DESCRI,PCTIPO,PCOBS,CODINT from PFMS03 WHERE PF=" & nPF & " AND CODCOMP='" & Ccodcomp & "'"
  End Select
  nCAMPOS = 4
  aCAM = Array("DESCRI", "PCTIPO", "PCOBS", "CODINT")
  aFOR = Array("C", "C", "C", "C")
  aPAD = Array("", "", "", "")
  aVAL = PegSQL(cARQUSO, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
  For iLOOP = 0 To nCAMPOS - 1
    txtFields(iLOOP) = aVAL(iLOOP)
  Next iLOOP
  For iLOOP = 0 To 1
    Select Case iLOOP
    Case 0
      cCAMIMG = "IMGCROQUI"                'croqui
    Case 1
      cCAMIMG = "IMGALERTA"                'alerta
    End Select
    cSQLIMG = "select " & cCAMIMG & " from PFS WHERE PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ
    
    If ADOPegBlob(Picture1(iLOOP), cARQUSO, "PFS", "PF=" & nPF & " AND SEQ=" & nSEQ & " AND SSQ=" & nSSQ, cCAMIMG) Then 'ADOPegBlob(cARQUSO, cSQLIMG, Picture1(iLOOP), cCAMIMG) Then
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
  CHECKDISPO
  FilRelat (0)                                 'item
  FilRelat (1)                                 'checklist
  FilRelat (2)                                 'on the job
  FilRelat (3)                                 'dispositivo

  txtFields(1).tEXT = Replace(txtFields(1).tEXT, "/", "-")  'chave apresentando erro de gravacao
  PrintPreview1.AuxiliaryButtonVisible = PrintPreview1.PrinterExists("Microsoft Print to PDF")
  PrintPreview1.AuxiliaryButtonToolTipText = "Salvar como PDF"

End Sub

Public Sub PrintPreview1_AuxiliaryButtonClick(UpdateReport As Boolean)
  PrintPreview1.ShowSaveToFile "Microsoft Print to PDF", "*.pdf"
  UpdateReport = False  ' we don't need to update the report in the Print preview window after this action (the default value of UpdateReport parameter is True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub VerImg_Click(Index As Integer)
  frmPicViewer.Show vbModal, Me
  If lRETU Then
    If lerarquivoimagem(eRETU01, Picture1(Index), Picture2(Index)) Then
      lTROCOU(Index) = True
    End If
  End If
End Sub
Private Sub imgsave_Click(Index As Integer)
  Dim cARQ As String
  Select Case Index
  Case 0
    cARQ = "croqui"
  Case 1
    cARQ = "alerta"
  End Select
  cARQ = cARQ & "_PF" & StrZero(nPF, 6) & "_SEQ_" & StrZero(nSEQ, 3) & "_SSQ_" & StrZero(nSSQ, 3)
  salvarpict Me, Picture1(Index), cARQ
End Sub
Private Sub Incluirimagem_Click(Index As Integer)
  Dim STMPFILE
  STMPFILE = OpenArqExt(Me, "", "JPG", "JPEG *.JPG")
  If lerarquivoimagem(STMPFILE, Picture1(Index), Picture2(Index)) Then
    lTROCOU(Index) = True
  End If
End Sub
Public Sub PrintPreview1_PrepareReport(Cancel As Boolean)
  On Error Resume Next
  Printer.Print
  Printer.PaintPicture Picture1(ePASS01), 0, 0
End Sub

