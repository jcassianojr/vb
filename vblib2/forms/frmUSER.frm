VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.1#0"; "vbccr18.ocx"
Begin VB.Form frmUSER 
   Caption         =   "Cadastro de Usuário"
   ClientHeight    =   7500
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   10515
   HelpContextID   =   12
   Icon            =   "frmUSER.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7500
   ScaleWidth      =   10515
   WindowState     =   2  'Maximized
   Begin vbExtra.SizeGrip SizeGrip1 
      Height          =   285
      Left            =   10230
      Top             =   7215
      Width           =   285
      _ExtentX        =   503
      _ExtentY        =   503
   End
   Begin XPControls.XPButton CmdLibGrp 
      Height          =   375
      Index           =   11
      Left            =   7800
      TabIndex        =   46
      Top             =   3600
      Width           =   1155
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Integrados"
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
   Begin XPControls.XPButton CmdLibGrp 
      Height          =   375
      Index           =   10
      Left            =   7800
      TabIndex        =   45
      Top             =   3240
      Width           =   1155
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Folha"
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
   Begin XPControls.XPButton CmdLibGrp 
      Height          =   375
      Index           =   9
      Left            =   7800
      TabIndex        =   44
      Top             =   2880
      Width           =   1155
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Documentos"
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
   Begin VB.Frame Frame1 
      Height          =   7005
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   10275
      Begin VB.TextBox tEXT 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   61
         TabStop         =   0   'False
         ToolTipText     =   "Senha"
         Top             =   6360
         Width           =   9615
      End
      Begin VBCCR18.ProgressBar Barra 
         Height          =   375
         Left            =   5520
         Top             =   4680
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
      End
      Begin VBCCR18.DTPicker DTPicker2 
         Height          =   375
         Left            =   240
         TabIndex        =   55
         Top             =   2520
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Value           =   44285
      End
      Begin VBCCR18.DTPicker DTPicker1 
         Height          =   375
         Left            =   1440
         TabIndex        =   54
         Top             =   960
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         Value           =   44285
         CheckBox        =   -1  'True
         AllowUserInput  =   -1  'True
      End
      Begin XPControls.XPButton cmdZeraSenha 
         Height          =   495
         Left            =   2280
         TabIndex        =   53
         Top             =   3120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Picture         =   "frmUSER.frx":038A
         Caption         =   "zerar senha"
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
      Begin XPControls.XPButton cmdTroca 
         Height          =   495
         Left            =   2280
         TabIndex        =   52
         Top             =   2520
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         Picture         =   "frmUSER.frx":0924
         Caption         =   "trocar senha"
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
      Begin XPControls.XPButton CmdApaAll 
         Height          =   375
         Left            =   7560
         TabIndex        =   47
         Top             =   3960
         Width           =   1215
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "ApagarTudo"
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
      Begin XPControls.XPButton CmdLibGrp 
         Height          =   375
         Index           =   8
         Left            =   7680
         TabIndex        =   43
         Top             =   2400
         Width           =   1155
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Relatorios"
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
      Begin XPControls.XPButton CmdLibGrp 
         Height          =   375
         Index           =   7
         Left            =   6360
         TabIndex        =   41
         Top             =   3480
         Width           =   1155
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Integrados"
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
      Begin XPControls.XPButton CmdLibGrp 
         Height          =   375
         Index           =   6
         Left            =   5040
         TabIndex        =   40
         Top             =   3480
         Width           =   1155
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Integrados"
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
      Begin XPControls.XPButton CmdLibGrp 
         Height          =   375
         Index           =   5
         Left            =   6360
         TabIndex        =   39
         Top             =   3120
         Width           =   1155
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Folha"
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
      Begin XPControls.XPButton CmdLibGrp 
         Height          =   375
         Index           =   4
         Left            =   5040
         TabIndex        =   38
         Top             =   3120
         Width           =   1155
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Folha"
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
      Begin XPControls.XPButton Cmdimpwrpt 
         Height          =   375
         Index           =   3
         Left            =   3720
         TabIndex        =   34
         ToolTipText     =   "Importar Direito Documentos Outro Usuario"
         Top             =   3480
         Width           =   1155
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Integrados"
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
      Begin XPControls.XPButton Cmdimpwrpt 
         Height          =   375
         Index           =   2
         Left            =   3720
         TabIndex        =   33
         ToolTipText     =   "Importar Direito Documentos Outro Usuario"
         Top             =   3120
         Width           =   1155
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Folha"
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
      Begin VB.CheckBox chkAtivo 
         Caption         =   "Ativo"
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
         Height          =   525
         Left            =   7500
         TabIndex        =   32
         ToolTipText     =   "Ativo - Acessa/Bloqueia Sistema"
         Top             =   300
         Width           =   915
      End
      Begin XPControls.XPButton Cmdimpwrpt 
         Height          =   375
         Index           =   0
         Left            =   3720
         TabIndex        =   31
         ToolTipText     =   "Importar Direitos Relatorio Outro Usuario"
         Top             =   2400
         Width           =   1155
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Relatorios"
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
      Begin XPControls.XPButton Cmdimpwrpt 
         Height          =   375
         Index           =   1
         Left            =   3720
         TabIndex        =   30
         ToolTipText     =   "Importar Direito Documentos Outro Usuario"
         Top             =   2760
         Width           =   1155
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Documentos"
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
      Begin XPControls.XPButton Cmdimpmenu 
         Height          =   375
         Left            =   1800
         TabIndex        =   29
         ToolTipText     =   "Importar Direito Menus Outro Usuario"
         Top             =   3840
         Width           =   1815
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Importar Direito Menus"
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
      Begin XPControls.XPButton CMDIMPBTN 
         Height          =   375
         Left            =   1800
         TabIndex        =   28
         ToolTipText     =   "Importar Direito Documentos Outro Usuario"
         Top             =   4260
         Width           =   1815
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Importar Direito Botoes"
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
      Begin XPControls.XPButton CmdLibGrp 
         Height          =   375
         Index           =   0
         Left            =   5040
         TabIndex        =   27
         Top             =   2400
         Width           =   1155
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Relatorios"
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
      Begin XPControls.XPButton CmdLibGrp 
         Height          =   375
         Index           =   1
         Left            =   5040
         TabIndex        =   26
         Top             =   2760
         Width           =   1155
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Documentos"
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
      Begin XPControls.XPButton CmdLibGrp 
         Height          =   375
         Index           =   2
         Left            =   6360
         TabIndex        =   25
         Top             =   2760
         Width           =   1155
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Documentos"
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
      Begin XPControls.XPButton CmdLibGrp 
         Height          =   375
         Index           =   3
         Left            =   6360
         TabIndex        =   24
         Top             =   2400
         Width           =   1155
         _ExtentX        =   0
         _ExtentY        =   0
         Caption         =   "Relatorios"
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
      Begin VB.TextBox DIZAPU 
         Enabled         =   0   'False
         ForeColor       =   &H8000000D&
         Height          =   435
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   4680
         Width           =   2295
      End
      Begin VB.Frame Frame12 
         Caption         =   "Buscar Id Envolvidos"
         ForeColor       =   &H00C00000&
         Height          =   1215
         Left            =   120
         TabIndex        =   19
         Top             =   3420
         Visible         =   0   'False
         Width           =   1575
         Begin VB.TextBox IdForm 
            Height          =   285
            Left            =   240
            TabIndex        =   20
            Top             =   300
            Width           =   1035
         End
         Begin VB.Label Label3 
            Caption         =   "do Form"
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   240
            TabIndex        =   21
            Top             =   660
            Width           =   975
         End
      End
      Begin VB.TextBox tEXT 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   7020
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Equivalente"
         Top             =   420
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox tEXT 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   1080
         TabIndex        =   3
         Text            =   "0"
         ToolTipText     =   "Nº Folha Pagamento"
         Top             =   1740
         Width           =   735
      End
      Begin VB.TextBox tEXT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   3840
         MaxLength       =   50
         TabIndex        =   4
         ToolTipText     =   "Nome do Usuario"
         Top             =   1680
         Width           =   3855
      End
      Begin VB.ComboBox cboEQUIVALENTE 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5460
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Equivalente"
         Top             =   420
         Width           =   1455
      End
      Begin VB.TextBox tEXT 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Senha"
         Top             =   3000
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         Caption         =   "Horarios Permitidos de Uso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   795
         Left            =   3360
         TabIndex        =   13
         Top             =   840
         Width           =   6855
         Begin VB.CheckBox chkweekend 
            Caption         =   "Acesso no fim-de-semana"
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
            Height          =   525
            Left            =   5040
            TabIndex        =   22
            ToolTipText     =   "Se o Usuario acesso final de semana"
            Top             =   120
            Width           =   1635
         End
         Begin VBCCR18.DTPicker DTPicker3 
            Height          =   375
            Left            =   1320
            TabIndex        =   56
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Value           =   44285
            Format          =   2
            CheckBox        =   -1  'True
            AllowUserInput  =   -1  'True
         End
         Begin VBCCR18.DTPicker DTPicker4 
            Height          =   375
            Left            =   3600
            TabIndex        =   57
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            Value           =   44285
            Format          =   2
            CheckBox        =   -1  'True
            AllowUserInput  =   -1  'True
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Hora inicial :"
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
            TabIndex        =   15
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label lblLabels 
            Alignment       =   1  'Right Justify
            Caption         =   "Hora final :"
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
            Left            =   2520
            TabIndex        =   14
            Top             =   240
            Width           =   1125
         End
      End
      Begin VB.TextBox tEXT 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Nº Usuario"
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox tEXT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   1
         ToolTipText     =   "Nome Usuario"
         Top             =   420
         Width           =   2295
      End
      Begin VB.TextBox tEXT 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   3840
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Ultimo Acesso"
         Top             =   435
         Width           =   1365
      End
      Begin XPControls.XPButton CmdClose 
         Height          =   435
         Left            =   8520
         TabIndex        =   48
         Top             =   360
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
         Picture         =   "frmUSER.frx":0DBE
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
      Begin XPControls.XPButton escidfolha 
         Height          =   375
         Index           =   0
         Left            =   1920
         TabIndex        =   49
         Top             =   1680
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Picture         =   "frmUSER.frx":1358
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
      Begin XPControls.XPButton escidfolha 
         Height          =   375
         Index           =   1
         Left            =   2400
         TabIndex        =   50
         Top             =   1680
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Picture         =   "frmUSER.frx":18F2
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
      Begin XPControls.XPButton Command1 
         Height          =   375
         Left            =   3360
         TabIndex        =   51
         Top             =   1680
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Picture         =   "frmUSER.frx":1E8C
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
      Begin VB.Label Lblchaveh 
         BackColor       =   &H00C0FFFF&
         Caption         =   "chaveh"
         Height          =   255
         Left            =   240
         TabIndex        =   60
         Top             =   6000
         Width           =   9615
      End
      Begin VB.Label LblpostelaB 
         BackColor       =   &H00C0FFFF&
         Caption         =   "postelab"
         Height          =   255
         Left            =   3360
         TabIndex        =   59
         Top             =   5640
         Width           =   2655
      End
      Begin VB.Label Lblpostelaa 
         BackColor       =   &H00C0FFFF&
         Caption         =   "postelaa"
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   5640
         Width           =   3015
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Apagar Grupo"
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   7680
         TabIndex        =   42
         Top             =   2160
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Liberar Todos"
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   6360
         TabIndex        =   37
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Liberar Grupo"
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   5040
         TabIndex        =   36
         Top             =   2160
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Importar Direitos "
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3780
         TabIndex        =   35
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Id Folha"
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
         Left            =   300
         TabIndex        =   18
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblLabels 
         Alignment       =   1  'Right Justify
         Caption         =   "Usuário Equivalente"
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
         Left            =   5340
         TabIndex        =   17
         Top             =   180
         Width           =   1845
      End
      Begin VB.Label lblLabels 
         Caption         =   "Data Para Troca Senha"
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
         Height          =   315
         Index           =   6
         Left            =   240
         TabIndex        =   16
         Top             =   2280
         Width           =   2385
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Data Limite Para Acessar o Sistema"
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
         Height          =   615
         Index           =   4
         Left            =   300
         TabIndex        =   12
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nrº Usuário"
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
         Left            =   120
         TabIndex        =   11
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label lblLabels 
         Caption         =   "Cognome do Usuário"
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
         Left            =   1320
         TabIndex        =   10
         Top             =   180
         Width           =   2265
      End
      Begin VB.Label lblLabels 
         Alignment       =   2  'Center
         Caption         =   "Último Acesso em :"
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
         Left            =   3540
         TabIndex        =   9
         Top             =   120
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmUSER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const nFORMID = 1012
Const cFORMID = "Cadastro de Usuários"
Dim aCAM, aFOR, aVAL, aPAD As Variant
Attribute aFOR.VB_VarUserMemId = 1073938432
Attribute aVAL.VB_VarUserMemId = 1073938432
Attribute aPAD.VB_VarUserMemId = 1073938432
Dim cARQ, cSQL As String
Attribute cARQ.VB_VarUserMemId = 1073938436
Attribute cSQL.VB_VarUserMemId = 1073938436
Dim nITEM, nCAMPOS As Long
Attribute nITEM.VB_VarUserMemId = 1073938438
Attribute nCAMPOS.VB_VarUserMemId = 1073938438

Private Sub cboEQUIVALENTE_Change()
'ComboChange  cboEQUIVALENTE
End Sub

Private Sub cboEQUIVALENTE_GotFocus()
  FocusMe
End Sub

Private Sub cboEQUIVALENTE_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Or KeyCode = 40 Then
    Visual
    SendKeys Chr(9)
  End If
  If KeyCode = 38 Then
    Visual
    SendKeys "+" + Chr(9)
  End If
End Sub

Private Sub cboEQUIVALENTE_LostFocus()
'ComboLostFocus cboEQUIVALENTE
End Sub

Private Sub chkAtivo_Click()
  Visual
End Sub

Private Sub chkAtivo_GotFocus()
  FocusMe
End Sub

Private Sub chkAtivo_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Or KeyCode = 40 Then
    Visual
    SendKeys Chr(9)
  End If
  If KeyCode = 38 Then
    Visual
    SendKeys "+" + Chr(9)
  End If

End Sub

Private Sub chkweekend_Click()
  Visual
End Sub

Private Sub chkweekend_GotFocus()
  FocusMe
End Sub

Private Sub chkweekend_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Or KeyCode = 40 Then
    Visual
    SendKeys Chr(9)
  End If
  If KeyCode = 38 Then
    Visual
    SendKeys "+" + Chr(9)
  End If

End Sub

Private Sub CmdApaAll_Click()
''RTFUSR
  sSQL = "select * from RTFUSR WHERE IDUSUARIO=" & tEXT(0)
  ApagaSQL DBWRPT, sSQL
  'RPTUSR
  sSQL = "select * from RPTUSR WHERE IDUSUARIO=" & tEXT(0)
  ApagaSQL DBWRPT, sSQL
  'RPTFOLUSR
  sSQL = "select * from RPTFOLUSR WHERE IDUSUARIO=" & tEXT(0)
  ApagaSQL DBWRPT, sSQL
  'RPTINTUSR
  sSQL = "select * from RPTINTUSR WHERE IDUSUARIO=" & tEXT(0)
  ApagaSQL DBWRPT, sSQL
  Alert "exclusao concluida"
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub CMDIMPBTN_Click()
  Dim nORIGEM, nDESTINO As Long
  Dim oDBORI As ADODB.Connection
  Dim oRSORI As New ADODB.Recordset
  Dim oRSDES As New ADODB.Recordset
  Dim cARQORI As String
  Dim cSQLORI, cSQLDES As String
  Dim aVAL, aCAM As Variant

  ''On Error Resume Next

  nORIGEM = Val(Busca("Digite o Numero de Origem", "Numero de Origem", "0", 8))
  nDESTINO = Val(Busca("Digite O Numero de destino", "Numero de Destino", CStr(zIDTEMP), 8))
  If Not MDG("Importar " & CStr(nORIGEM) & " para " & CStr(nDESTINO), "Confirme Importaçao") Then
    Exit Sub
  End If

  Screen.MousePointer = vbHourglass
  cSQLORI = "select * from USUCAD WHERE IDUSUARIO=" & nORIGEM
  cARQORI = GeracArq(Dbname)
  Set oDBORI = New ADODB.Connection
  Set oRSORI = New ADODB.Recordset

  oDBORI.ConnectionTimeout = 120
  oDBORI.Open cARQORI
  oRSORI.Open cSQLORI, oDBORI, adOpenForwardOnly, adLockReadOnly

  If Not oRSORI.EOF Then
    Barra.Value = 0
    nBARPOS = 0
    While Not oRSORI.EOF
      cMENU = oRSORI("FORM")
      cTIPO = oRSORI("CONTROLE")
      iMENU = oRSORI("INDICE")
      DIZAPU = cMENU & "/" & cTIPO & "/" & CStr(iMENU)
      BARPOS
      cSQLDES = "select * from USUCAD WHERE FORM='" & cMENU & "' AND CONTROLE='" & cTIPO & "' AND INDICE=" & iMENU & " AND IDUSUARIO=" & nDESTINO
      'aCAM = Array("IDUSUARIO", "FORM", "CONTROLE", "INDICE", "ATUALIZADO", "LIGADO")
      'aVAL = Array(nDESTINO, cMENU, cTIPO, iMENU, True, True)
      'IncluiSQL Dbname, cSQLDES, 6, aCAM, aVAL, True, False
      oRSDES.Open cSQLDES, oDBORI, adOpenDynamic, adLockOptimistic
      If oRSDES.EOF Then
        oRSDES.AddNew
        oRSDES("IDUSUARIO") = nDESTINO
        oRSDES("FORM") = cMENU
        oRSDES("CONTROLE") = cTIPO
        oRSDES("INDICE") = iMENU
        oRSDES("ATUALIZADO") = True
        oRSDES("LIGADO") = True
        oRSDES.Update
      End If
      oRSDES.Close
      oRSORI.MoveNext
    Wend
  End If
  oRSORI.Close
  oDBORI.Close
  Set oRSORI = Nothing
  Set oDBORI = Nothing
  Screen.MousePointer = vbDefault

End Sub

Private Sub Cmdimpmenu_Click()
  Dim nORIGEM, nDESTINO As Long
  Dim oDBORI As ADODB.Connection
  Dim oRSORI As New ADODB.Recordset
  Dim oRSDES As New ADODB.Recordset
  Dim cSQLORI, cSQLDES As String
  Dim aVAL, aCAM As Variant

  On Error Resume Next

  nORIGEM = Val(Busca("Digite o Numero de Origem", "Numero de Origem", "0", 8))
  nDESTINO = Val(Busca("Digite O Numero de destino", "Numero de Destino", CStr(zIDTEMP), 8))
  If Not MDG("Importar " & CStr(nORIGEM) & " para " & CStr(nDESTINO), "Confirme Importaçao") Then
    Exit Sub
  End If

  Screen.MousePointer = vbHourglass

  Set oDBORI = New ADODB.Connection
  oDBORI.CursorLocation = adUseClient
  oDBORI.ConnectionTimeout = 120
  oDBORI.Open GeracArq(Dbname)


  cSQLORI = "select * from MENUUSU WHERE IDUSUARIO=" & nORIGEM
  Set oRSORI = New ADODB.Recordset
  oRSORI.Open cSQLORI, oDBORI, adOpenForwardOnly, adLockReadOnly

  If Not oRSORI.EOF Then
    Barra.Value = 0
    nBARPOS = 0
    oRSORI.MoveLast
    nROWREC = oRSORI.RecordCount
    oRSORI.MoveFirst
    While Not oRSORI.EOF
      cMENU = oRSORI("MENU")
      iMENU = oRSORI("INDICE")
      DIZAPU = cMENU & "/" & CStr(iMENU)
      BARPOS
      cSQLDES = "select * from MENUUSU WHERE MENU='" & cMENU & "' AND INDICE=" & iMENU & " AND IDUSUARIO=" & nDESTINO
      'aCAM = Array("IDUSUARIO", "MENU", "INDICE", "ATUALIZADO", "LIGADO")
      'aVAL = Array(nDESTINO, cMENU, iMENU, True, True)
      oRSDES.Open cSQLDES, oDBORI, adOpenDynamic, adLockOptimistic
      If oRSDES.EOF Then
        oRSDES.AddNew
        oRSDES("IDUSUARIO") = nDESTINO
        oRSDES("MENU") = cMENU
        oRSDES("INDICE") = iMENU
        oRSDES("ATUALIZADO") = True
        oRSDES("LIGADO") = True
        oRSDES.Update
      End If
      oRSDES.Close
      ''IncluiSQL Dbname, cSQLDES, 5, aCAM, aVAL, True, False ''tabela ja aberta
      oRSORI.MoveNext
    Wend
  End If
  oRSORI.Close
  oDBORI.Close
  Set oRSORI = Nothing
  Set oDBORI = Nothing
  Screen.MousePointer = vbDefault

End Sub

Private Sub Cmdimpwrpt_Click(Index As Integer)
  Dim nORIGEM, nDESTINO As Long
  Dim oDBORI As ADODB.Connection
  Dim oRSORI As ADODB.Recordset
  Dim oRSDES As ADODB.Recordset

  Dim cSQLORI, cSQLDES, cGRP, cRPT, cTabela As String

  On Error Resume Next

  cTabela = "RPTUSR"
  If Index = 1 Then cTabela = "RTFUSR"
  If Index = 2 Then cTabela = "RPTFOLUSR"
  If Index = 3 Then cTabela = "RPTINTUSR"

  nORIGEM = Val(Busca("Digite o Numero de Origem", "Numero de Origem", "0", 8))
  nDESTINO = Val(Busca("Digite O Numero de destino", "Numero de Destino", CStr(zIDTEMP), 8))
  If Not MDG("Importar " & CStr(nORIGEM) & " para " & CStr(nDESTINO), "Confirme Importaçao") Then
    Exit Sub
  End If

  Screen.MousePointer = vbHourglass
  cSQLORI = "select * from " & cTabela & " WHERE IDUSUARIO=" & nORIGEM
  Set oDBORI = New ADODB.Connection
  oDBORI.CursorLocation = adUseClient
  oDBORI.ConnectionTimeout = 120
  oDBORI.Open GeracArq(DBWRPT)

  Set oRSORI = New ADODB.Recordset
  Set oRSDES = New ADODB.Recordset
  oRSORI.Open cSQLORI, oDBORI, adOpenForwardOnly, adLockReadOnly



  If Not oRSORI.EOF Then
    Barra.Value = 0
    nBARPOS = 0
    oRSORI.MoveLast
    nROWREC = oRSORI.RecordCount
    oRSORI.MoveFirst
    While Not oRSORI.EOF
      cGRP = oRSORI("GRP")
      cRPT = oRSORI("RPT")
      DIZAPU = cGRP & "/" & cRPT
      BARPOS
      cSQLDES = "select * from " & cTabela & " WHERE GRP='" & cGRP & "' AND RPT='" & cRPT & "' AND IDUSUARIO=" & nDESTINO
      oRSDES.Open cSQLDES, oDBORI, adOpenKeyset, adLockOptimistic  ' adOpenStatic
      If oRSDES.EOF Then
        oRSDES.AddNew
        oRSDES("IDUSUARIO") = nDESTINO
        oRSDES("GRP") = cGRP
        oRSDES("RPT") = cRPT
        oRSDES("IMPRIME") = oRSORI("IMPRIME")
        oRSDES("EXPORTA") = oRSORI("EXPORTA")
        oRSDES("VISUALIZA") = oRSORI("VISUALIZA")
        oRSDES("SALVARTF") = oRSORI("SALVARTF")
        oRSDES("SALVATXT") = oRSORI("SALVATXT")
        oRSDES("NOVO") = oRSORI("NOVO")
        oRSDES("ABRIR") = oRSORI("ABRIR")
        oRSDES("EDITAR") = oRSORI("EDITAR")
        oRSDES("ABRIRCOM") = oRSORI("ABRIRCOM")
        oRSDES.Update
      End If
      oRSDES.Close
      oRSORI.MoveNext
    Wend
  End If
  oRSORI.Close
  oDBORI.Close
  Set oRSORI = Nothing
  Set oRSDES = Nothing
  Set oDBORI = Nothing
  Screen.MousePointer = vbDefault

End Sub

Private Sub CmdLibGrp_Click(Index As Integer)
  Dim nDESTINO As Long
  Dim oDBORI As ADODB.Connection
  Dim oDBDESI As ADODB.Connection
  Dim oRSORI As ADODB.Recordset
  Dim oRSDES As ADODB.Recordset
  Dim cORIGEM As String
  Dim cSQLORI As String
  Dim cSQLDES As String
  Dim cGRP As String
  Dim cRPT As String
  Dim cTabela As String

  '   On Error Resume Next

  cTabela = "RPTUSR"
  zRPTARQ = PegPath("PATH", "RPT") + "RPT.MDB"
  If Index = 1 Or Index = 2 Or Index = 9 Then
    zRPTARQ = PegPath("PATH", "RTF") + "RTF.MDB"
    cTabela = "RTFUSR"
  End If


  If Index = 4 Or Index = 5 Or Index = 10 Then
    zRPTARQ = PegPath("PATH", "FOL") + "RPTFOL.MDB"
    cTabela = "RPTFOLUSR"
  End If


  If Index = 6 Or Index = 7 Or Index = 11 Then
    zRPTARQ = PegPath("PATH", "INT") + "RPTINT.MDB"
    cTabela = "RPTINTUSR"
  End If


  If Index = 0 Or Index = 1 Or Index = 4 Or Index = 6 Or Index >= 8 Then
    cORIGEM = UCase(Busca("Digite o Grupo", "Codigo do Grupo", "", 8))
  Else
    cORIGEM = "Todos"
  End If
  nDESTINO = FixInt(Busca("Digite O Numero de destino", "Numero de Destino", tEXT(0), 8))
  If nDESTINO = 0 Or (Index < 2 And Len(cORIGEM) = 0) Then
    Alert ("Necessario Preencher Origem e Destino")
    Exit Sub
  End If
  If Index >= 8 Then
    If Not MDG("Apagar " & cORIGEM & " para " & CStr(nDESTINO), "Confirmar Exclusao") Then
      Exit Sub
    End If
    cSQLDES = "select * from " & cTabela & " WHERE GRP='" & cORIGEM & "' AND IDUSUARIO=" & nDESTINO
    ApagaSQL DBWRPT, cSQLDES
    Exit Sub
  Else
    If Not MDG("Liberar " & cORIGEM & " para " & CStr(nDESTINO), "Confirme Liberacao") Then
      Exit Sub
    End If
  End If


  Screen.MousePointer = vbHourglass
  If Index = 0 Or Index = 1 Or Index = 4 Or Index = 6 Then
    cSQLORI = "select * from RPT WHERE GRP='" & cORIGEM & "'"
  Else
    cSQLORI = "select * from RPT"
  End If

  Set oDBORI = New ADODB.Connection
  oDBORI.CursorLocation = adUseClient
  oDBORI.ConnectionTimeout = 120
  oDBORI.Open GeracArq(zRPTARQ)

  Set oDBDES = New ADODB.Connection
  oDBDES.CursorLocation = adUseClient
  oDBDES.ConnectionTimeout = 120
  oDBDES.Open GeracArq(DBWRPT)

  Set oRSORI = New ADODB.Recordset
  Set oRSDES = New ADODB.Recordset


  oRSORI.Open cSQLORI, oDBORI, adOpenForwardOnly, adLockReadOnly

  If Not oRSORI.EOF Then
    Barra.Value = 0
    nBARPOS = 0
    oRSORI.MoveLast
    nROWREC = oRSORI.RecordCount
    oRSORI.MoveFirst
    While Not oRSORI.EOF
      cGRP = oRSORI("GRP")
      cRPT = oRSORI("RPT")
      DIZAPU = cGRP & "/" & cRPT
      BARPOS
      cSQLDES = "select * from " & cTabela & " WHERE GRP='" & cGRP & "' AND RPT='" & cRPT & "' AND IDUSUARIO=" & nDESTINO

      oRSDES.Open cSQLDES, oDBDES, adOpenKeyset, adLockOptimistic  'adOpenStatic
      If oRSDES.EOF Then
        oRSDES.AddNew
        oRSDES("IDUSUARIO") = nDESTINO
        oRSDES("GRP") = cGRP
        oRSDES("RPT") = cRPT
        oRSDES("IMPRIME") = True
        oRSDES("EXPORTA") = True
        oRSDES("VISUALIZA") = True
        oRSDES("SALVARTF") = True
        oRSDES("SALVATXT") = True
        oRSDES("NOVO") = True
        oRSDES("ABRIR") = True
        oRSDES("EDITAR") = True
        oRSDES("ABRIRCOM") = ""
        oRSDES.Update
      End If
      oRSDES.Close
      oRSORI.MoveNext
    Wend
  End If
  oRSORI.Close
  oDBORI.Close
  oDBDES.Close
  Set oRSORI = Nothing
  Set oRSDES = Nothing
  Set oDBORI = Nothing
  Set oDBDES = Nothing
  Screen.MousePointer = vbDefault

End Sub

Private Sub DTPicker1_GotFocus()
  FocusMe
End Sub

Private Sub DTPicker1_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Or KeyCode = 40 Then
    Visual
    SendKeys Chr(9)
  End If
  If KeyCode = 38 Then
    Visual
    SendKeys "+" + Chr(9)
  End If

End Sub

Private Sub DTPicker2_GotFocus()
  FocusMe
End Sub

Private Sub DTPicker2_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Or KeyCode = 40 Then
    Visual
    SendKeys Chr(9)
  End If
  If KeyCode = 38 Then
    Visual
    SendKeys "+" + Chr(9)
  End If

End Sub

Private Sub DTPicker3_GotFocus()
  FocusMe
End Sub

Private Sub DTPicker3_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Or KeyCode = 40 Then
    Visual
    SendKeys Chr(9)
  End If
  If KeyCode = 38 Then
    Visual
    SendKeys "+" + Chr(9)
  End If

End Sub

Private Sub DTPicker4_GotFocus()
  FocusMe
End Sub

Private Sub DTPicker4_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Or KeyCode = 40 Then
    Visual
    SendKeys Chr(9)
  End If
  If KeyCode = 38 Then
    Visual
    SendKeys "+" + Chr(9)
  End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If FixInt(tEXT(5)) = 0 Then
    If MDG("Colocar o numero do seu Funcionario na folha") Then
      Cancel = 1
      Exit Sub
    End If
  End If

  If Len(Trim(tEXT(2))) = 0 Then
    If MDG("Senha nao Colocada deseja Colocar ? ", "Gravar Senha") Then
      zIDTEMP = tEXT(0)
      frmUSUSENHA.Show vbModal, Me
      If lRETU Then
        tEXT(2) = eRETU01
        DTPicker2.Value = Date + 60
      End If
    End If
  End If

  ' CHAVEH POSTELAA POSTELAB nao sao gravados por isso nao sao atribuidos abaixo
  If MDG("Gravar e Sair", "Gravando Usuarios") Then
    For nITEM = 0 To 6                       'Array comeca 0
      aVAL(nITEM) = tEXT(nITEM)
    Next nITEM
    aVAL(9) = FixNumBol(chkAtivo)
    aVAL(10) = FixNumBol(chkweekend)
    aVAL(11) = DTPicker1
    aVAL(12) = DTPicker2
    aVAL(7) = DTPicker3.Hour & ":" & DTPicker3.Minute
    aVAL(8) = DTPicker4.Hour & ":" & DTPicker4.Minute
    aVAL(13) = tEXT(13)  ' CHAVEV
    'volta para 14 para nao gravar postelaa postelab chaveh estao apenas para exibir no label
    nCAMPOS = 14
    GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR
    GravaLog nFORMID, 0, "Gravou Usuario"
  Else
    GravaLog nFORMID, 0, "Nao Gravou Usuario"
  End If
  Screen.MousePointer = vbDefault
End Sub

Private Sub Command1_Click()
  Dim cARQ As String
  Dim aRETU As Variant
  Dim sSQL As String
  Dim nNUMERO As Long
  nNUMERO = FixInt(tEXT(5), 0)
  If demitido(nNUMERO) Then
    If lRETU Then
      If MDG("bloquear acesso e Zerar Senha") Then
        chkAtivo.Value = 0
        chkweekend.Value = 0
        DTPicker1.Value = eRETU01
        DTPicker2.Value = eRETU01
        tEXT(2) = ""                     ''zERA SeNHA
      End If
      If MDG("zerar matricula") Then
        tEXT(5) = 0
      End If
    End If
    Exit Sub
  End If

  lRETU = False
  If zusalx <> "NAO" Then
    cARQ = PegPath("PATH", "LOGIXODBC")
    sSQL = "SELECT nom_completo as nomtec,dat_demis as demitido FROM funcionario WHERE cod_empresa='01' and num_matricula=" & nNUMERO
    aRETU = PegSQL(cARQ, sSQL, 2, Array("NOMTEC", "DEMITIDO"), Array("C", "DN"), Array("", Today()))
  End If
  If Not lRETU Then
    cARQ = PegPath("PATH", "CADMP04")
    cARQ = GeraConn(cARQ, "JETFOX")
    sSQL = "SELECT NOMTEC,DEMITIDO FROM MP04 WHERE TECNICO=" & nNUMERO
    aRETU = PegSQL(cARQ, sSQL, 2, Array("NOMTEC", "DEMITIDO"), Array("C", "DN"), Array("", Today()))
  End If
  If lRETU Then
    tEXT(6) = aRETU(0)
  End If

End Sub

Private Sub cmdTroca_Click()
  zIDTEMP = tEXT(0)
  frmUSUSENHA.Show vbModal, Me
  If lRETU Then
    tEXT(2) = eRETU01
    DTPicker2.Value = Date + 60
    tEXT(13) = UCase(CreateSHA256HashString(UCase(Trim(tEXT(1))) + UCase(Trim(eRETU02))))
  End If

End Sub

Private Sub cmdZeraSenha_Click()
  tEXT(2) = " "
  tEXT(13) = " "
  cmdClose_Click
End Sub

Private Sub escidfolha_Click(Index As Integer)
  If Index = 0 Then
    ePASS01 = ""
    escMP04.Show vbModal, Me
  Else
    ePASS01 = "LOGIX"
    escMP04.Show vbModal, Me
  End If

  If lRETU Then

    frmUSER.tEXT(5) = eRETU01
    frmUSER.tEXT(6) = eRETU02
    Command1_Click

  End If


End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  If Left(UCase(CStr(App.EXEName)), 4) <> "WRPT" Then
    Cmdimpwrpt(0).Enabled = False
    Cmdimpwrpt(1).Enabled = False
    Cmdimpwrpt(2).Enabled = False
    Cmdimpwrpt(3).Enabled = False
    CmdLibGrp(0).Enabled = False
    CmdLibGrp(1).Enabled = False
    CmdLibGrp(2).Enabled = False
    CmdLibGrp(3).Enabled = False
    CmdLibGrp(4).Enabled = False
    CmdLibGrp(5).Enabled = False
    CmdLibGrp(6).Enabled = False
    CmdLibGrp(7).Enabled = False
    CmdLibGrp(8).Enabled = False
    CmdLibGrp(9).Enabled = False
    CmdLibGrp(10).Enabled = False
    CmdLibGrp(11).Enabled = False
  End If
  'Configura Help
  Me.Caption = cFORMID
  HelpContextID = nFORMID

  If zusalx = "NAO" Then
    escidfolha(1).Visible = False
    escidfolha(1).TabStop = False

  End If


  ''Arquivo
  cARQ = PegPath("PATH", "SYSUSER")
  If Not FileExist(cARQ, True) Then
    End
  End If

  'Variaveis sql
  cSQL = ePASS01
  IdForm = ePASS02
  zIDTEMP = ePASS02

  nCAMPOS = 17
  aCAM = Array("IDUSUARIO", "USUARIO", "SENHA", "EQUIVALENTE", _
               "DATAULT", "IDFOLHA", "NOMEFOLHA", "HORAINI", "HORAFIM", _
               "ATIVO", "WEEKEND", "DATAVAL", "TROCAR", "chavev", "POSTELAA", "POSTELAB", "CHAVEH")
  aFOR = Array("N", "C", "C", "C", "D", "N", "C", "", "", "BN", "BN", "DC", "DC", "C", "C", "C", "C")
  aPAD = Array(0, "", "", "", "", 0, "", Now, Now, False, False, Today() + 30, Today() + 60, "", "", "", "")
  aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
  For nITEM = 0 To 6                           '' array comeca 0
    tEXT(nITEM) = aVAL(nITEM)
  Next nITEM
  If aVAL(7) <> "" Then DTPicker3 = Date + aVAL(7)  ''Adciona date pois o datapicker nao aceita vazio
  If aVAL(8) <> "" Then DTPicker4 = Date + aVAL(8)  ''na mascara datapicker fica so  a hora

  chkAtivo = FixBolNum(aVAL(9))
  chkweekend = FixBolNum(aVAL(10))
  If IsDate(aVAL(11)) Then
    DTPicker1 = aVAL(11)
  End If
  If IsDate(aVAL(12)) Then
    DTPicker2 = aVAL(12)
  End If
  tEXT(13) = aVAL(13)
  Lblpostelaa.Caption = aVAL(14)
  LblpostelaB.Caption = aVAL(15)
  Lblchaveh.Caption = aVAL(16)
  senhapos (LblpostelaB.Caption)
  Visual
End Sub
Private Sub Text_GotFocus(Index As Integer)
  FocusMe
End Sub
Private Sub tEXT_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index = 5 Then
    KeyAscii = ValiText(KeyAscii, "#NI")
  End If
End Sub

Private Sub Text_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Or KeyCode = 40 Then
    Visual
    SendKeys Chr(9)
  End If
  If KeyCode = 38 Then
    Visual
    SendKeys "+" + Chr(9)
  End If
End Sub

Public Sub Visual()
  If chkAtivo.Value = 1 Then
    chkAtivo.Caption = "Usuário Ativo"
  Else
    chkAtivo.Caption = "Usuário Desativado"
  End If
  If chkweekend.Value = 1 Then
    chkweekend.Caption = "Pode Acessar no fim-de-semana"
  Else
    chkweekend.Caption = "Não Pode Acessar no fim-de-semana"
  End If
End Sub

Private Sub BARPOS()
  Dim nPOSATU As Long
  nPOSATU = 0
  nBARPOS = nBARPOS + 1
  If nROWREC > 0 Then
    nPOSATU = Int(nBARPOS / nROWREC * 100)
  End If
  If nPOSATU > 0 And nPOSATU < 101 Then
    Debug.Print nPOSATU
    Barra.Value = nPOSATU
  End If
  DIZAPU.Refresh
End Sub


