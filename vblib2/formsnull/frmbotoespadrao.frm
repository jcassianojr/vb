VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form frmbotoespadrao 
   Caption         =   "Form1"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   8190
   StartUpPosition =   3  'Windows Default
   Begin XPControls.XPButton XPButton6 
      Height          =   435
      Left            =   5760
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":0000
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
   Begin XPControls.XPButton XPButton3 
      Height          =   435
      Left            =   3840
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":059A
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
   Begin XPControls.XPButton cmdFOTO 
      Height          =   435
      Left            =   1680
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":0B34
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
   Begin XPControls.XPButton XPButton4 
      Height          =   435
      Left            =   240
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":10CE
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
      Left            =   240
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3840
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":1668
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
   Begin XPControls.XPButton CmdSendMail 
      Height          =   435
      Left            =   2760
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":1C02
      Caption         =   "EnviarEmail"
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
   Begin XPControls.XPButton btnOpen 
      Height          =   435
      Index           =   0
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":219C
      Caption         =   "Abrir"
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
      Left            =   1920
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":2736
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
   Begin XPControls.XPButton CmdTransfer 
      Height          =   435
      Left            =   1920
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":2CD0
      Caption         =   "Executar"
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
   Begin XPControls.XPButton CmdEscPfim 
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   720
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "frmbotoespadrao.frx":326A
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
   Begin XPControls.XPButton CmdVisua 
      Height          =   435
      Left            =   1920
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":3804
      Caption         =   "Visualizar"
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
      Left            =   1920
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":3D9E
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
   Begin XPControls.XPButton cmdClose 
      Height          =   435
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":4338
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
   Begin XPControls.XPButton ESCCLI 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   720
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Picture         =   "frmbotoespadrao.frx":48D2
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
   Begin XPControls.XPButton cancelar 
      Height          =   435
      Left            =   3960
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":4E6C
      Caption         =   "Cancelar"
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
   Begin XPControls.XPButton XPButton1 
      Height          =   435
      Left            =   3960
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":5406
      Caption         =   "Confirmar"
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
   Begin XPControls.XPButton XPButton2 
      Height          =   375
      Left            =   2400
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   720
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "frmbotoespadrao.frx":59A0
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
   Begin XPControls.XPButton CmdFiltro 
      Height          =   435
      Left            =   600
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   240
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":5F3A
      Caption         =   "filtrar"
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
   Begin XPControls.XPButton CmdOrdem 
      Height          =   435
      Left            =   3960
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":64D4
      Caption         =   "ordernar"
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
   Begin XPControls.XPButton Salvar 
      Height          =   435
      Index           =   1
      Left            =   240
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":6A6E
      Caption         =   "Salvar &TXT"
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
   Begin XPControls.XPButton Salvar 
      Height          =   435
      Index           =   0
      Left            =   1920
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":6E08
      Caption         =   "Salvar &PDF"
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
      Left            =   2280
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":73A2
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
   Begin XPControls.XPButton cmdexportar 
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Picture         =   "frmbotoespadrao.frx":793C
      Caption         =   "xls(TAB)"
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
   Begin XPControls.XPButton cmdexportar 
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   5280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Picture         =   "frmbotoespadrao.frx":7ED6
      Caption         =   "xml"
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
   Begin XPControls.XPButton cmdexportar 
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   5280
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Picture         =   "frmbotoespadrao.frx":8470
      Caption         =   "Html"
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
   Begin XPControls.XPButton cmdexportar 
      Height          =   375
      Index           =   3
      Left            =   3240
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   5280
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      Picture         =   "frmbotoespadrao.frx":8A0A
      Caption         =   "TXT"
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
   Begin XPControls.XPButton cmdexportar 
      Height          =   375
      Index           =   5
      Left            =   4080
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5280
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      Picture         =   "frmbotoespadrao.frx":8DA4
      Caption         =   "RTF"
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
   Begin XPControls.XPButton cmdexportar 
      Height          =   375
      Index           =   4
      Left            =   4920
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5280
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      Picture         =   "frmbotoespadrao.frx":913E
      Caption         =   "PDF"
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
   Begin XPControls.XPButton XPButton7 
      Height          =   375
      Left            =   3600
      TabIndex        =   28
      Top             =   4800
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   661
      Picture         =   "frmbotoespadrao.frx":94D8
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
      PictureAlignment=   1
   End
   Begin XPControls.XPButton cmdexportar 
      Height          =   375
      Index           =   6
      Left            =   6600
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   5280
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      Picture         =   "frmbotoespadrao.frx":9872
      Caption         =   "doc"
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
   Begin XPControls.XPButton cmdexportar 
      Height          =   375
      Index           =   7
      Left            =   5760
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5280
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      Picture         =   "frmbotoespadrao.frx":9E0C
      Caption         =   "xls"
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
      Left            =   5760
      TabIndex        =   31
      Top             =   4560
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      Picture         =   "frmbotoespadrao.frx":A3A6
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
   Begin XPControls.XPButton Incluirimagem 
      Height          =   435
      Left            =   6000
      TabIndex        =   32
      Top             =   1440
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":A840
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
      Left            =   6120
      TabIndex        =   33
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":ACDA
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
   Begin XPControls.XPButton zxCommand2 
      Height          =   435
      Left            =   6000
      TabIndex        =   34
      Top             =   2520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":B174
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
   Begin XPControls.XPButton CmdEditar 
      Height          =   435
      Left            =   5760
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":B60E
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
   Begin XPControls.XPButton XPButton8 
      Height          =   375
      Left            =   3960
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Picture         =   "frmbotoespadrao.frx":BBA8
      Caption         =   "Localizar"
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
      Height          =   435
      Left            =   5040
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":C142
      Caption         =   "Abrir Com"
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
   Begin XPControls.XPButton XPButton5 
      Height          =   435
      Left            =   4200
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":C6DC
      Caption         =   "Buscar Pasta"
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
   Begin XPControls.XPButton XPButton9 
      Height          =   375
      Left            =   3120
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   720
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "frmbotoespadrao.frx":CC76
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
   Begin XPControls.XPButton cmdexportar 
      Height          =   375
      Index           =   8
      Left            =   600
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   5760
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      Picture         =   "frmbotoespadrao.frx":D210
      Caption         =   "zip"
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
   Begin XPControls.XPButton XPButton10 
      Height          =   375
      Left            =   1800
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   720
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "frmbotoespadrao.frx":D6AA
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
   Begin XPControls.XPButton cmdexportar 
      Height          =   615
      Index           =   9
      Left            =   1680
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1085
      Picture         =   "frmbotoespadrao.frx":DB44
      Caption         =   "zip"
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
      Left            =   6720
      TabIndex        =   43
      Top             =   3480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Picture         =   "frmbotoespadrao.frx":DE5E
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
   Begin XPControls.XPButton xxx 
      Height          =   615
      Left            =   3000
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   1085
      Picture         =   "frmbotoespadrao.frx":E3F8
      Caption         =   "Competência "
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
   Begin XPControls.XPButton XPButton11 
      Height          =   435
      Left            =   6240
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   4080
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":E792
      Caption         =   "Apagar Filtro"
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
   Begin XPControls.XPButton CmdShell 
      Height          =   315
      Left            =   120
      TabIndex        =   46
      Top             =   240
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      Caption         =   "S"
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
   Begin XPControls.XPButton XPButton12 
      Height          =   435
      Left            =   5640
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   240
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmbotoespadrao.frx":EB2C
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
End
Attribute VB_Name = "frmbotoespadrao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub
