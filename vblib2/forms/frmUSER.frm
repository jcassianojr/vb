VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form frmUSER 
   Caption         =   "Cadastro de Usuário"
   ClientHeight    =   5535
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   10515
   HelpContextID   =   12
   Icon            =   "frmUSER.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   10515
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdLibGrp 
      Caption         =   "Integrados"
      Height          =   375
      Index           =   11
      Left            =   7800
      TabIndex        =   51
      Top             =   3600
      Width           =   1155
   End
   Begin VB.CommandButton CmdLibGrp 
      Caption         =   "Folha"
      Height          =   375
      Index           =   10
      Left            =   7800
      TabIndex        =   50
      Top             =   3240
      Width           =   1155
   End
   Begin VB.CommandButton CmdLibGrp 
      Caption         =   "Documentos"
      Height          =   375
      Index           =   9
      Left            =   7800
      TabIndex        =   49
      Top             =   2880
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   5085
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   10275
      Begin XPControls.XPButton cmdZeraSenha 
         Height          =   495
         Left            =   2280
         TabIndex        =   58
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
         TabIndex        =   57
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
      Begin VB.CommandButton CmdApaAll 
         Caption         =   "ApagarTudo"
         Height          =   375
         Left            =   7560
         TabIndex        =   52
         Top             =   3960
         Width           =   1215
      End
      Begin VB.CommandButton CmdLibGrp 
         Caption         =   "Relatorios"
         Height          =   375
         Index           =   8
         Left            =   7680
         TabIndex        =   48
         Top             =   2400
         Width           =   1155
      End
      Begin VB.CommandButton CmdLibGrp 
         Caption         =   "Integrados"
         Height          =   375
         Index           =   7
         Left            =   6360
         TabIndex        =   46
         Top             =   3480
         Width           =   1155
      End
      Begin VB.CommandButton CmdLibGrp 
         Caption         =   "Integrados"
         Height          =   375
         Index           =   6
         Left            =   5040
         TabIndex        =   45
         Top             =   3480
         Width           =   1155
      End
      Begin VB.CommandButton CmdLibGrp 
         Caption         =   "Folha"
         Height          =   375
         Index           =   5
         Left            =   6360
         TabIndex        =   44
         Top             =   3120
         Width           =   1155
      End
      Begin VB.CommandButton CmdLibGrp 
         Caption         =   "Folha"
         Height          =   375
         Index           =   4
         Left            =   5040
         TabIndex        =   43
         Top             =   3120
         Width           =   1155
      End
      Begin VB.CommandButton Cmdimpwrpt 
         Caption         =   "Integrados"
         Height          =   375
         Index           =   3
         Left            =   3720
         TabIndex        =   39
         ToolTipText     =   "Importar Direito Documentos Outro Usuario"
         Top             =   3480
         Width           =   1155
      End
      Begin VB.CommandButton Cmdimpwrpt 
         Caption         =   "Folha"
         Height          =   375
         Index           =   2
         Left            =   3720
         TabIndex        =   38
         ToolTipText     =   "Importar Direito Documentos Outro Usuario"
         Top             =   3120
         Width           =   1155
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
         TabIndex        =   37
         ToolTipText     =   "Ativo - Acessa/Bloqueia Sistema"
         Top             =   300
         Width           =   915
      End
      Begin VB.CommandButton Cmdimpwrpt 
         Caption         =   "Relatorios"
         Height          =   375
         Index           =   0
         Left            =   3720
         TabIndex        =   35
         ToolTipText     =   "Importar Direitos Relatorio Outro Usuario"
         Top             =   2400
         Width           =   1155
      End
      Begin VB.CommandButton Cmdimpwrpt 
         Caption         =   "Documentos"
         Height          =   375
         Index           =   1
         Left            =   3720
         TabIndex        =   34
         ToolTipText     =   "Importar Direito Documentos Outro Usuario"
         Top             =   2760
         Width           =   1155
      End
      Begin VB.CommandButton Cmdimpmenu 
         Caption         =   "Importar Direito Menus"
         Height          =   375
         Left            =   1740
         TabIndex        =   33
         ToolTipText     =   "Importar Direito Menus Outro Usuario"
         Top             =   3840
         Width           =   1815
      End
      Begin VB.CommandButton CMDIMPBTN 
         Caption         =   "Importar Direito Botoes"
         Height          =   375
         Left            =   1740
         TabIndex        =   32
         ToolTipText     =   "Importar Direito Documentos Outro Usuario"
         Top             =   4260
         Width           =   1815
      End
      Begin VB.CommandButton CmdLibGrp 
         Caption         =   "Relatorios"
         Height          =   375
         Index           =   0
         Left            =   5040
         TabIndex        =   31
         Top             =   2400
         Width           =   1155
      End
      Begin VB.CommandButton CmdLibGrp 
         Caption         =   "Documentos"
         Height          =   375
         Index           =   1
         Left            =   5040
         TabIndex        =   30
         Top             =   2760
         Width           =   1155
      End
      Begin VB.CommandButton CmdLibGrp 
         Caption         =   "Documentos"
         Height          =   375
         Index           =   2
         Left            =   6360
         TabIndex        =   29
         Top             =   2760
         Width           =   1155
      End
      Begin VB.CommandButton CmdLibGrp 
         Caption         =   "Relatorios"
         Height          =   375
         Index           =   3
         Left            =   6360
         TabIndex        =   28
         Top             =   2400
         Width           =   1155
      End
      Begin VB.TextBox DIZAPU 
         Enabled         =   0   'False
         Height          =   315
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Frame Frame12 
         Caption         =   "Buscar Id Envolvidos"
         Height          =   1215
         Left            =   120
         TabIndex        =   23
         Top             =   3420
         Visible         =   0   'False
         Width           =   1335
         Begin VB.TextBox IdForm 
            Height          =   285
            Left            =   90
            TabIndex        =   24
            Top             =   300
            Width           =   1035
         End
         Begin VB.Label Label3 
            Caption         =   "do Form"
            Height          =   315
            Left            =   120
            TabIndex        =   25
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
         TabIndex        =   9
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
         TabIndex        =   6
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
         TabIndex        =   7
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
         TabIndex        =   8
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
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Senha"
         Top             =   3060
         Visible         =   0   'False
         Width           =   975
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
         Left            =   3480
         TabIndex        =   17
         Top             =   840
         Width           =   6735
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
            Left            =   4500
            TabIndex        =   26
            ToolTipText     =   "Se o Usuario acesso final de semana"
            Top             =   120
            Width           =   2130
         End
         Begin MSComCtl2.DTPicker DTPicker3 
            Height          =   330
            Left            =   1380
            TabIndex        =   4
            Top             =   255
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "00:00"
            Format          =   121110530
            CurrentDate     =   38397
         End
         Begin MSComCtl2.DTPicker DTPicker4 
            Height          =   330
            Left            =   3480
            TabIndex        =   5
            Top             =   240
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "00:00"
            Format          =   121110530
            CurrentDate     =   0.999988425925926
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
            TabIndex        =   19
            Top             =   300
            Width           =   1215
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
            Left            =   2280
            TabIndex        =   18
            Top             =   300
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
         Height          =   285
         Index           =   0
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "Nº Usuario"
         Top             =   420
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
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1500
         TabIndex        =   3
         Tag             =   "Expiraçao"
         ToolTipText     =   "Expiraçao"
         Top             =   840
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   121110529
         CurrentDate     =   36516
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   300
         TabIndex        =   10
         ToolTipText     =   "Data Troca Senha"
         Top             =   2580
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   121110529
         CurrentDate     =   36516
      End
      Begin MSComctlLib.ProgressBar Barra 
         Height          =   375
         Left            =   7200
         TabIndex        =   36
         Top             =   4560
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   1
      End
      Begin XPControls.XPButton CmdClose 
         Height          =   435
         Left            =   8520
         TabIndex        =   53
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
         TabIndex        =   54
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
         TabIndex        =   55
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
         TabIndex        =   56
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
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Apagar Grupo"
         Height          =   315
         Left            =   7680
         TabIndex        =   47
         Top             =   2160
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Liberar Todos"
         Height          =   315
         Left            =   6360
         TabIndex        =   42
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Liberar Grupo"
         Height          =   315
         Left            =   5040
         TabIndex        =   41
         Top             =   2160
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Importar Direitos "
         Height          =   315
         Left            =   3780
         TabIndex        =   40
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
         TabIndex        =   22
         Top             =   1740
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
Dim cARQ, cSQL As String
Dim nITEM, nCAMPOS As Long

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
            oRSDES.Open cSQLDES, oDBORI, adOpenStatic, adLockOptimistic
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

            oRSDES.Open cSQLDES, oDBDES, adOpenStatic, adLockOptimistic
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
                DTPicker2.Value = eRETU02
            End If
        End If
    End If

       
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
    Dim sSQL             As String
    Dim nNUMERO          As Long
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
        DTPicker2.Value = eRETU02
    End If
    
End Sub

Private Sub cmdZeraSenha_Click()
    tEXT(2) = " "
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
    Center Me
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

    nCAMPOS = 13
    aCAM = Array("IDUSUARIO", "USUARIO", "SENHA", "EQUIVALENTE", _
                 "DATAULT", "IDFOLHA", "NOMEFOLHA", "HORAINI", "HORAFIM", _
                 "ATIVO", "WEEKEND", "DATAVAL", "TROCAR")
    aFOR = Array("N", "C", "C", "C", "D", "N", "C", "", "", "BN", "BN", "DC", "DC")
    aPAD = Array(0, "", "", "", "", 0, "", "07:00", "18:00", False, False, Today() + 30, Today() + 60)
    aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
    For nITEM = 0 To 6                           '' array comeca 0
        tEXT(nITEM) = aVAL(nITEM)
    Next nITEM
    If aVAL(7) <> "" Then DTPicker3 = aVAL(7)
    If aVAL(8) <> "" Then DTPicker4 = aVAL(8)
     
    chkAtivo = FixBolNum(aVAL(9))
    chkweekend = FixBolNum(aVAL(10))
    If IsDate(aVAL(11)) Then
        DTPicker1 = aVAL(11)
    End If
    If IsDate(aVAL(12)) Then
        DTPicker2 = aVAL(12)
    End If
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


