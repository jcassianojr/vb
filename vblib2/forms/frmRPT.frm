VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{66E63055-5A66-4C79-9327-4BC077858695}#9.0#0"; "newtab01.OCX"
Begin VB.Form FrmRpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relatórios /Documentos"
   ClientHeight    =   4590
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   11325
   HelpContextID   =   7
   Icon            =   "frmRPT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox text 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   3
      Left            =   9840
      TabIndex        =   30
      ToolTipText     =   "Data de Inclusão do Relatorio"
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox text 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "Quantidade de Vezes que Foi Utilizado"
      Top             =   2640
      Width           =   1095
   End
   Begin VB.TextBox text 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Data da Ultima Utilizacao do Relatorio/Documento"
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox text 
      Height          =   285
      Index           =   16
      Left            =   6240
      TabIndex        =   9
      ToolTipText     =   "Codigo Interno do Relatorio"
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox text 
      Height          =   285
      Index           =   17
      Left            =   1080
      TabIndex        =   7
      ToolTipText     =   "Descriçao do Relatorio"
      Top             =   480
      Width           =   6255
   End
   Begin VB.TextBox text 
      Height          =   285
      Index           =   15
      Left            =   4200
      TabIndex        =   2
      ToolTipText     =   "Sub Grupo do Relatorio"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox text 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Codigo do Relatorio (Sistema)"
      Top             =   120
      Width           =   1815
   End
   Begin NewTabCtl.NewTab SSTab1 
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   900
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   6165
      ControlJustAdded=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabOrientation  =   1
      TabHeight       =   732
      ControlVersion  =   9
      TabPic16(0)     =   "frmRPT.frx":058A
      TabCaption(0)   =   "Configuracao"
      Tab(0).ControlCount=   19
      Tab(0).Control(0)=   "Command2"
      Tab(0).Control(1)=   "text(7)"
      Tab(0).Control(2)=   "text(6)"
      Tab(0).Control(3)=   "text(8)"
      Tab(0).Control(4)=   "CmdAbrirCom(1)"
      Tab(0).Control(5)=   "CmdAbrirCom(0)"
      Tab(0).Control(6)=   "text(5)"
      Tab(0).Control(7)=   "chkFields(26)"
      Tab(0).Control(8)=   "text(4)"
      Tab(0).Control(9)=   "chkFields(25)"
      Tab(0).Control(10)=   "chkFields(24)"
      Tab(0).Control(11)=   "escolherrpt"
      Tab(0).Control(12)=   "CmdEDIT(0)"
      Tab(0).Control(13)=   "CmdEDIT(1)"
      Tab(0).Control(14)=   "lblLabels(10)"
      Tab(0).Control(15)=   "lblLabels(12)"
      Tab(0).Control(16)=   "lblLabels(18)"
      Tab(0).Control(17)=   "lblLabels(3)"
      Tab(0).Control(18)=   "lblLabels(5)"
      TabPic16(1)     =   "frmRPT.frx":08DC
      TabCaption(1)   =   "Arquivos"
      Tab(1).ControlCount=   21
      Tab(1).Control(0)=   "text(14)"
      Tab(1).Control(1)=   "text(13)"
      Tab(1).Control(2)=   "text(12)"
      Tab(1).Control(3)=   "text(11)"
      Tab(1).Control(4)=   "text(9)"
      Tab(1).Control(5)=   "text(10)"
      Tab(1).Control(6)=   "text(21)"
      Tab(1).Control(7)=   "text(20)"
      Tab(1).Control(8)=   "text(19)"
      Tab(1).Control(9)=   "text(18)"
      Tab(1).Control(10)=   "escarq(0)"
      Tab(1).Control(11)=   "escarq(1)"
      Tab(1).Control(12)=   "escarq(2)"
      Tab(1).Control(13)=   "escarq(3)"
      Tab(1).Control(14)=   "lblLabels(17)"
      Tab(1).Control(15)=   "lblLabels(16)"
      Tab(1).Control(16)=   "lblLabels(15)"
      Tab(1).Control(17)=   "lblLabels(14)"
      Tab(1).Control(18)=   "lblLabels(13)"
      Tab(1).Control(19)=   "lblLabels(19)"
      Tab(1).Control(20)=   "lblLabels(11)"
      TabPic16(2)     =   "frmRPT.frx":0C2E
      TabCaption(2)   =   "Filtro Sql"
      Tab(2).ControlCount=   4
      Tab(2).Control(0)=   "text(23)"
      Tab(2).Control(1)=   "text(22)"
      Tab(2).Control(2)=   "lblLabels(21)"
      Tab(2).Control(3)=   "lblLabels(0)"
      Begin XPControls.XPButton Command2 
         Height          =   375
         Left            =   3600
         TabIndex        =   52
         Top             =   600
         Width           =   615
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Limpar"
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
      Begin VB.TextBox text 
         Height          =   795
         Index           =   23
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   51
         ToolTipText     =   "Obervarções Sobre o Relatorio"
         Top             =   1440
         Width           =   8535
      End
      Begin VB.TextBox text 
         Height          =   555
         Index           =   7
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   45
         ToolTipText     =   "Obervarções Sobre o Relatorio"
         Top             =   1440
         Width           =   7575
      End
      Begin VB.TextBox text 
         Height          =   285
         Index           =   6
         Left            =   1080
         TabIndex        =   44
         ToolTipText     =   "Titulo para o Relatorio"
         Top             =   1080
         Width           =   6015
      End
      Begin VB.TextBox text 
         Height          =   405
         Index           =   8
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         ToolTipText     =   "Mesagem a Ser Exibida Usuario"
         Top             =   2160
         Width           =   7575
      End
      Begin XPControls.XPButton CmdAbrirCom 
         Height          =   375
         Index           =   1
         Left            =   3120
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "Abrir Com Todas Extensoes"
         Top             =   600
         Width           =   375
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "*..."
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
      Begin XPControls.XPButton CmdAbrirCom 
         Height          =   375
         Index           =   0
         Left            =   2520
         TabIndex        =   40
         TabStop         =   0   'False
         ToolTipText     =   "Abrir com Somente da Extencao"
         Top             =   600
         Width           =   495
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Ext..."
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
      Begin VB.TextBox text 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "Nome do Manipulador"
         Top             =   600
         Width           =   1455
      End
      Begin VB.CheckBox chkFields 
         Caption         =   "Filtro"
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   26
         Left            =   6240
         TabIndex        =   38
         ToolTipText     =   "Marque se o relatorio necessitar Filtro"
         Top             =   600
         Width           =   675
      End
      Begin VB.TextBox text 
         Height          =   285
         Index           =   4
         Left            =   1080
         TabIndex        =   36
         ToolTipText     =   "Nome do Arquivo do Relatorio/Documento"
         Top             =   240
         Width           =   7575
      End
      Begin VB.TextBox text 
         Height          =   795
         Index           =   22
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   34
         ToolTipText     =   "Obervarções Sobre o Relatorio"
         Top             =   360
         Width           =   8535
      End
      Begin VB.TextBox text 
         Height          =   285
         Index           =   14
         Left            =   -68040
         TabIndex        =   18
         ToolTipText     =   "Sub Grupo do Relatorio"
         Top             =   2145
         Width           =   855
      End
      Begin VB.TextBox text 
         Height          =   285
         Index           =   13
         Left            =   -70080
         TabIndex        =   17
         ToolTipText     =   "Sub Grupo do Relatorio"
         Top             =   2145
         Width           =   870
      End
      Begin VB.TextBox text 
         Height          =   285
         Index           =   12
         Left            =   -72000
         TabIndex        =   16
         ToolTipText     =   "Sub Grupo do Relatorio"
         Top             =   2145
         Width           =   945
      End
      Begin VB.TextBox text 
         Height          =   285
         Index           =   11
         Left            =   -73800
         TabIndex        =   15
         ToolTipText     =   "Sub Grupo do Relatorio"
         Top             =   2145
         Width           =   975
      End
      Begin VB.TextBox text 
         Height          =   285
         Index           =   9
         Left            =   -73800
         TabIndex        =   20
         ToolTipText     =   "Tabela Nome Referencia"
         Top             =   1665
         Width           =   2805
      End
      Begin VB.TextBox text 
         Height          =   285
         Index           =   10
         Left            =   -70080
         TabIndex        =   19
         ToolTipText     =   "Tabela Nome Referencia"
         Top             =   1665
         Width           =   2505
      End
      Begin VB.TextBox text 
         Height          =   285
         Index           =   21
         Left            =   -73560
         TabIndex        =   14
         ToolTipText     =   "Nome da 4ª Base de Dados"
         Top             =   1305
         Width           =   6015
      End
      Begin VB.TextBox text 
         Height          =   285
         Index           =   20
         Left            =   -73560
         TabIndex        =   13
         ToolTipText     =   "Nome da 3ª Base de Dados"
         Top             =   945
         Width           =   6015
      End
      Begin VB.TextBox text 
         Height          =   285
         Index           =   19
         Left            =   -73560
         TabIndex        =   12
         ToolTipText     =   "Nome da 2ª Base de Dados"
         Top             =   585
         Width           =   6015
      End
      Begin VB.TextBox text 
         Height          =   285
         Index           =   18
         Left            =   -73560
         TabIndex        =   11
         ToolTipText     =   "Nome da 1ª Base de Dados"
         Top             =   195
         Width           =   6015
      End
      Begin VB.CheckBox chkFields 
         Caption         =   "Busca"
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   25
         Left            =   5280
         TabIndex        =   6
         ToolTipText     =   "Marque se o relatorio Possuir Busca"
         Top             =   600
         Width           =   795
      End
      Begin VB.CheckBox chkFields 
         Caption         =   "Árvore"
         ForeColor       =   &H00C00000&
         Height          =   285
         Index           =   24
         Left            =   4440
         TabIndex        =   5
         ToolTipText     =   "Marque Para Exibir a Arvore"
         Top             =   600
         Width           =   855
      End
      Begin XPControls.XPButton escolherrpt 
         Height          =   315
         Left            =   8760
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         Picture         =   "frmRPT.frx":0F80
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
      Begin XPControls.XPButton CmdEDIT 
         Height          =   315
         Index           =   0
         Left            =   8760
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   600
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         Picture         =   "frmRPT.frx":151A
         Caption         =   "E.."
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
      Begin XPControls.XPButton CmdEDIT 
         Height          =   315
         Index           =   1
         Left            =   8760
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   960
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   556
         Picture         =   "frmRPT.frx":1AB4
         Caption         =   "E.."
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
      Begin XPControls.XPButton escarq 
         Height          =   315
         Index           =   0
         Left            =   -67440
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   225
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         Picture         =   "frmRPT.frx":204E
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
      Begin XPControls.XPButton escarq 
         Height          =   315
         Index           =   1
         Left            =   -67440
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   585
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         Picture         =   "frmRPT.frx":25E8
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
      Begin XPControls.XPButton escarq 
         Height          =   315
         Index           =   2
         Left            =   -67440
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   945
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         Picture         =   "frmRPT.frx":2B82
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
      Begin XPControls.XPButton escarq 
         Height          =   315
         Index           =   3
         Left            =   -67440
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   1305
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   556
         Picture         =   "frmRPT.frx":311C
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
         BackStyle       =   0  'Transparent
         Caption         =   "SQL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   21
         Left            =   -74880
         TabIndex        =   50
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Obs:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   10
         Left            =   240
         TabIndex        =   48
         Top             =   1560
         Width           =   525
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Título"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   47
         Top             =   1080
         Width           =   585
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Mensagem"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   18
         Left            =   120
         TabIndex        =   46
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Abrir Com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   42
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Arquivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   705
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Prefiltro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   0
         Left            =   -74760
         TabIndex        =   35
         Top             =   120
         Width           =   855
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Acumulado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   17
         Left            =   -69120
         TabIndex        =   25
         Top             =   2145
         Width           =   945
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Fechado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   16
         Left            =   -70980
         TabIndex        =   24
         Top             =   2145
         Width           =   735
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Baixa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   15
         Left            =   -72720
         TabIndex        =   23
         Top             =   2145
         Width           =   495
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Atual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   14
         Left            =   -74520
         TabIndex        =   22
         Top             =   2145
         Width           =   555
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Tabela"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   13
         Left            =   -74520
         TabIndex        =   27
         Top             =   1665
         Width           =   585
      End
      Begin VB.Label lblLabels 
         BackStyle       =   0  'Transparent
         Caption         =   "Alias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   19
         Left            =   -70920
         TabIndex        =   26
         Top             =   1665
         Width           =   450
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Arquivo Caminhos Parâmetros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   975
         Index           =   11
         Left            =   -74760
         TabIndex        =   21
         Top             =   195
         Width           =   1035
      End
   End
   Begin XPControls.XPButton cmdClose 
      Height          =   435
      Left            =   9480
      TabIndex        =   53
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmRPT.frx":36B6
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
   Begin XPControls.XPButton Command1 
      Height          =   375
      Left            =   9840
      TabIndex        =   58
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Picture         =   "frmRPT.frx":3C50
      Caption         =   "Liberar"
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
      Caption         =   "Prefiltro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   20
      Left            =   360
      TabIndex        =   49
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "Inclusão Alteracao"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   495
      Index           =   8
      Left            =   9840
      TabIndex        =   33
      Top             =   3000
      Width           =   1065
   End
   Begin VB.Label lblLabels 
      Caption         =   "Qtde"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   7
      Left            =   9840
      TabIndex        =   32
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "Ultilizado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   4
      Left            =   9840
      TabIndex        =   31
      Top             =   1800
      Width           =   915
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cód.Interno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   6
      Left            =   5040
      TabIndex        =   10
      Top             =   120
      Width           =   1035
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descrição"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "Sub-Grupo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   9
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   945
   End
   Begin VB.Label lblLabels 
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "FrmRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const nFORMID = 1066
Const cFORMID = "Relatorios /Documentos"
Dim aCAM As Variant
Dim aFOR As Variant
Dim aVAL As Variant
Dim aPAD As Variant
Dim cARQ As String
Dim cSQL As String
Dim nITEM As Long
Dim nCAMPOS As Long
Option Explicit

Private Sub CmdAbrirCom_Click(Index As Integer)

  eLOCALIZA = text(5).text
  ePASS01 = ""
  If Index = 0 Then
    ePASS01 = EXTENSAO(FixStr(text(4).text))
  End If
  EscRptExec.Show vbModal, Me
  If lRETU Then
    text(5).text = eRETU02
  End If

End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub CmdEDIT_Click(Index As Integer)
Dim cEXTENSAO As String
  eLOCALIZA = text(5).text
  If Index = 0 Then
    cARQRTF = text(4).text
    cEXTENSAO = EXTENSAO(cARQRTF)
    Select Case cEXTENSAO
      Case "RPT"
        If Index = 0 Then
          ShellEx cARQRTF, essSW_SHOWDEFAULT, , , , CLng(Me.hWnd)
        Else
          ShellEx cARQRTF, essSW_SHOWDEFAULT, , , , CLng(Me.hWnd)
        End If
    End Select
  End If
End Sub

Private Sub Command1_Click()
  escrptusr.Show vbModal
End Sub

Private Sub Command2_Click()
  text(5).text = ""
End Sub

Private Sub escarq_Click(Index As Integer)
  Index = Index + 18
  ePASS01 = text(Index)
 ' FrmPegdb.Show vbModal, Me
  If lRETU Then
    text(Index) = eRETU01
  End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If MDG("Gravar e Sair", "Gravando ") Then
    For nITEM = 0 To 23
      aVAL(nITEM) = text(nITEM)
    Next nITEM
    For nITEM = 24 To 26
      aVAL(nITEM) = FixNumBol(chkFields(nITEM).Value)
    Next nITEM
    GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR, 4 'comeca no 1 0=RPT chave 1 utilizacao 2 qtde 3 inclusao
    GravaLog nFORMID, 0, "Gravou "
  Else
    GravaLog nFORMID, 0, "Nao Gravou "
  End If
  Screen.MousePointer = vbDefault
End Sub

Private Sub escolherrpt_Click()
  Dim cFILTER As String
  Dim sPath As String

  cFILTER = "Crystal Reports (*.RPT)" & vbNullChar & "*.RPT" & vbNullChar

  ' cFILTER = cFILTER & "Report Manager (*.REP)" & vbNullChar & "*.REP" & vbNullChar 'Removidos rep com pouco uso
  cFILTER = cFILTER & "CA-RET Reports VORETRUN (*.RET)" & vbNullChar & "*.RET" & vbNullChar  ''usar voretrun pois caretrun e 16 bits
  cFILTER = cFILTER & "ZPL Etiquetas preview(*.ZPL)" & vbNullChar & "*.RET" & vbNullChar

  'cFILTER = cFILTER & "QuickReport MMK-Report (*.QRP)" & vbNullChar & "*.QRP" & vbNullChar

  cFILTER = cFILTER & "Rich-Text Format (*.RTF)" & vbNullChar & "*.RTF" & vbNullChar
  cFILTER = cFILTER & "Acrobat Reader (*.PDF)" & vbNullChar & "*.PDF" & vbNullChar
  cFILTER = cFILTER & "Arquivos de Textos (*.TXT)" & vbNullChar & "*.TXT" & vbNullChar
  cFILTER = cFILTER & "Documentos do Word  (*.DOC)" & vbNullChar & "*.DOC" & vbNullChar
  cFILTER = cFILTER & "Planilhas do Excel  (*.XLS)" & vbNullChar & "*.XLS" & vbNullChar
  cFILTER = cFILTER & "Apresentacao Power-Point (*.PPS)(*.PPT)" & vbNullChar & "*.PPS;*.PPT" & vbNullChar

  cFILTER = cFILTER & "Ajuda Windows (*.HLP)" & vbNullChar & "*.HLP" & vbNullChar
  cFILTER = cFILTER & "Ajuda Windows Compactadas (*.CHM)" & vbNullChar & "*.CHM" & vbNullChar

  cFILTER = cFILTER & "Internet (*.HTM;*.HTML) " & vbNullChar & "*.HTM;*.HTML" & vbNullChar

  cFILTER = cFILTER & "Windows Control Panel (*.CPL)" & vbNullChar & "*.CPL" & vbNullChar
  cFILTER = cFILTER & "Video - MPEG (*.mpeg;*.MPG;*.M1V;*.MP2;*.MPA;*.MPE) " & vbNullChar & "*.mpeg;*.MPG;*.M1V;*.MP2;*.MPA;*.MPE" & vbNullChar
  cFILTER = cFILTER & "Videos (*.avi;*.wmf) " & vbNullChar & "*.avi;*.wmf" & vbNullChar
  cFILTER = cFILTER & "Audio (*.wav;*.snd;*.au;*.aif;*.aifc;*.aiff;*.wma;*.mp3) " & vbNullChar & "*.wav;*.snd;*.au;*.aif;*.aifc;*.aiff;*.wma;*.mp3" & vbNullChar
  cFILTER = cFILTER & "MIDI (*.mid;*.rmi;*.midi) " & vbNullChar & "*.mid;*.rmi;*.midi" & vbNullChar
  cFILTER = cFILTER & "CD (*.cda) " & vbNullChar & "*.cda" & vbNullChar
  cFILTER = cFILTER & "Windows Media (*.asf;*.wm;*.wma;*.wmv) " & vbNullChar & "*.asf;*.wm;*.wma;*.wmv" & vbNullChar
  cFILTER = cFILTER & "Rep Midia (*.asx;*.wax;*.m3u;*.mvx;*.wmx) " & vbNullChar & "*.asx;*.wax;*.m3u;*.mvx;*.wmx" & vbNullChar
  cFILTER = cFILTER & "Imagens-1 (*.BMP;*.ICO;*.WMF;*.EMF;*.JPG;*.GIF) " & vbNullChar & "*.BMP;*.ICO;*.WMF;*.EMF;*.JPG;*.GIF" & vbNullChar
  cFILTER = cFILTER & "Imagens-2 (*.DIB;*.JPEG;*.PCX;*.TGA;*.TIF;*.TIFF;*.PNG;*.PIC) " & vbNullChar & "*.DIB;*.JPEG;*.PCX;*.TGA;*.TIF;*.TIFF;*.PNG;*.PIC" & vbNullChar
  cFILTER = cFILTER & "Desenhos(*.DWF;*.DWG;*.DXF;*.RML;*.IPT;*.IAM;*.IDW)" & vbNullChar & "*.DWF;*.DWG;*.DXF;*.RML;*.IPT;*.IAM;*.IDW" & vbNullChar
  cFILTER = cFILTER & "Todos os Arquivos (*.*) " & vbNullChar & "*.*"
  sPath = App.Path
  If Len(zRPTGRPCAM) > 0 Then
    sPath = Caminex(zRPTGRPCAM)
  Else
    If Len(zRPTCAM) > 0 Then
      sPath = Caminex(zRPTCAM)
    End If
  End If
  text(4).text = FileOpen(Me, cFILTER, 1, "", "*", sPath, "Escolher Arquivo")

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub Form_Load()

  CenterFormToScreen Me
  ''Configura Help
  Me.Caption = cFORMID
  HelpContextID = nFORMID

  ''Arquivo Compatibilizar Versao Antiga
  cARQ = ePASS02
  cSQL = ePASS01
  ''cARQ = pegpath("PATH", "SYSRPT")

  If Not FileConnExist(cARQ, True, , cSQL) Then
    End
  End If

  'Variaveis sql

  nCAMPOS = 27
  aCAM = Array("RPT", "DATAIMP", "UTILIZADO", "DATACRI", "ARQUIVO", _
               "ABRIRCOM", "TITULO", "MENSAGEM", "COBSMSG", "TABNAME", _
               "TABALIAS", "CARQUSO", "CARQBAI", "CARQFEC", "CARQACU", _
               "SUBGRP", "COGNOME", "NOME", "CAMINHO", "CAMINH2", _
               "CAMINH3", "CAMINH4", "PREFILTRO", "SQLUSO", "ARVORE", "BUSCA", "LFILTRO")
  aFOR = Array("C", "DH", "N", "DH", "C", _
               "C", "C", "C", "C", "C", _
               "C", "C", "C", "C", "C", _
               "C", "C", "C", "C", "C", _
               "C", "C", "C", "C", "BN", "BN", "BN")
  aPAD = Array("", Now, 0, Now, "", _
               "", "", "", "", "", _
               "", "", "", "", "", _
               "", "", "", "", "", _
               "", "", "", "", False, False, False)
  aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
  For nITEM = 0 To 23
    text(nITEM) = aVAL(nITEM)
  Next nITEM
  For nITEM = 24 To 26
    chkFields(nITEM) = aVAL(nITEM)
  Next nITEM
End Sub


