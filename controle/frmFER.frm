VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{EA478B61-D9EC-47F6-BB21-95A533AF2251}#1.3#0"; "TabExt01.OCX"
Begin VB.Form frmFER 
   Caption         =   "Cadastro Ferramenta"
   ClientHeight    =   6405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9900
   Icon            =   "frmFER.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   9900
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtNumero 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1335
   End
   Begin TabExCtl.SSTabEx SSTab1 
      Height          =   4215
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   7435
      Tabs            =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabsPerRow      =   5
      TabHeight       =   794
      TabCaption(0)   =   "Basica"
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Preventiva Vida Util"
      Tab(1).ControlCount=   8
      Tab(1).Control(0)=   "TXTFIELDS(3)"
      Tab(1).Control(1)=   "TXTFIELDS(2)"
      Tab(1).Control(2)=   "TXTFIELDS(1)"
      Tab(1).Control(3)=   "TXTFIELDS(0)"
      Tab(1).Control(4)=   "Label(5)"
      Tab(1).Control(5)=   "Label(4)"
      Tab(1).Control(6)=   "Label(3)"
      Tab(1).Control(7)=   "Label(2)"
      TabCaption(2)   =   "Dimensoes Elaborador"
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Tecnicas I"
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "Tecnicas II"
      Tab(4).ControlCount=   82
      Tab(4).Control(0)=   "TXTFIELDS(10)"
      Tab(4).Control(1)=   "TXTFIELDS(9)"
      Tab(4).Control(2)=   "TXTFIELDS(8)"
      Tab(4).Control(3)=   "TXTFIELDS(7)"
      Tab(4).Control(4)=   "TXTFIELDS(6)"
      Tab(4).Control(5)=   "Command20"
      Tab(4).Control(6)=   "TXTFIELDS(11)"
      Tab(4).Control(7)=   "TXTFIELDS(12)"
      Tab(4).Control(8)=   "TXTFIELDS(13)"
      Tab(4).Control(9)=   "TXTFIELDS(14)"
      Tab(4).Control(10)=   "TXTFIELDS(15)"
      Tab(4).Control(11)=   "TXTFIELDS(16)"
      Tab(4).Control(12)=   "TXTFIELDS(17)"
      Tab(4).Control(13)=   "TXTFIELDS(18)"
      Tab(4).Control(14)=   "esc1(1)"
      Tab(4).Control(15)=   "esc1(2)"
      Tab(4).Control(16)=   "esc1(3)"
      Tab(4).Control(17)=   "esc1(4)"
      Tab(4).Control(18)=   "esc1(5)"
      Tab(4).Control(19)=   "esc1(6)"
      Tab(4).Control(20)=   "esc1(7)"
      Tab(4).Control(21)=   "esc1(8)"
      Tab(4).Control(22)=   "esc1(9)"
      Tab(4).Control(23)=   "esc1(10)"
      Tab(4).Control(24)=   "esc1(11)"
      Tab(4).Control(25)=   "esc1(12)"
      Tab(4).Control(26)=   "esc1(13)"
      Tab(4).Control(27)=   "esc1(14)"
      Tab(4).Control(28)=   "esc1(15)"
      Tab(4).Control(29)=   "esc1(16)"
      Tab(4).Control(30)=   "esc1(17)"
      Tab(4).Control(31)=   "esc1(18)"
      Tab(4).Control(32)=   "esc1(19)"
      Tab(4).Control(33)=   "esc1(20)"
      Tab(4).Control(34)=   "esc1(21)"
      Tab(4).Control(35)=   "esc1(22)"
      Tab(4).Control(36)=   "esc1(23)"
      Tab(4).Control(37)=   "esc1(24)"
      Tab(4).Control(38)=   "esc1(25)"
      Tab(4).Control(39)=   "esc1(26)"
      Tab(4).Control(40)=   "esc1(27)"
      Tab(4).Control(41)=   "esc1(28)"
      Tab(4).Control(42)=   "esc1(29)"
      Tab(4).Control(43)=   "esc1(30)"
      Tab(4).Control(44)=   "esc1(31)"
      Tab(4).Control(45)=   "esc1(32)"
      Tab(4).Control(46)=   "esc1(33)"
      Tab(4).Control(47)=   "esc1(34)"
      Tab(4).Control(48)=   "esc1(35)"
      Tab(4).Control(49)=   "esc1(36)"
      Tab(4).Control(50)=   "esc1(37)"
      Tab(4).Control(51)=   "esc1(38)"
      Tab(4).Control(52)=   "esc1(39)"
      Tab(4).Control(53)=   "esc1(40)"
      Tab(4).Control(54)=   "esc1(41)"
      Tab(4).Control(55)=   "esc1(42)"
      Tab(4).Control(56)=   "esc1(43)"
      Tab(4).Control(57)=   "esc1(44)"
      Tab(4).Control(58)=   "esc1(45)"
      Tab(4).Control(59)=   "esc1(46)"
      Tab(4).Control(60)=   "esc1(47)"
      Tab(4).Control(61)=   "esc1(48)"
      Tab(4).Control(62)=   "esc1(49)"
      Tab(4).Control(63)=   "esc1(50)"
      Tab(4).Control(64)=   "esc1(51)"
      Tab(4).Control(65)=   "esc1(52)"
      Tab(4).Control(66)=   "esc1(53)"
      Tab(4).Control(67)=   "esc1(54)"
      Tab(4).Control(68)=   "esc1(55)"
      Tab(4).Control(69)=   "esc1(56)"
      Tab(4).Control(70)=   "esc1(57)"
      Tab(4).Control(71)=   "esc1(58)"
      Tab(4).Control(72)=   "esc1(59)"
      Tab(4).Control(73)=   "esc1(60)"
      Tab(4).Control(74)=   "esc1(61)"
      Tab(4).Control(75)=   "esc1(62)"
      Tab(4).Control(76)=   "esc1(63)"
      Tab(4).Control(77)=   "esc1(64)"
      Tab(4).Control(78)=   "Label4"
      Tab(4).Control(79)=   "Label3"
      Tab(4).Control(80)=   "Label2"
      Tab(4).Control(81)=   "Label1"
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         Left            =   -74400
         MaxLength       =   24
         TabIndex        =   98
         Top             =   2640
         Width           =   6675
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   -69960
         MaxLength       =   24
         TabIndex        =   96
         Top             =   1800
         Width           =   1995
      End
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   8
         Left            =   -70680
         Locked          =   -1  'True
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   1440
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   7
         Left            =   -70680
         Locked          =   -1  'True
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   1080
         Width           =   495
      End
      Begin VB.TextBox TXTFIELDS 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Index           =   6
         Left            =   -70680
         Locked          =   -1  'True
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   600
         Width           =   495
      End
      Begin VB.CommandButton Command20 
         Caption         =   "Tipo Almofados 4x6/8x8"
         Height          =   255
         Left            =   -69960
         TabIndex        =   89
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         Left            =   -74880
         MaxLength       =   24
         TabIndex        =   88
         Top             =   600
         Width           =   1995
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   12
         Left            =   -74880
         MaxLength       =   24
         TabIndex        =   87
         Top             =   840
         Width           =   1995
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   13
         Left            =   -74880
         MaxLength       =   24
         TabIndex        =   86
         Top             =   1080
         Width           =   1995
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   14
         Left            =   -74880
         MaxLength       =   24
         TabIndex        =   85
         Top             =   1320
         Width           =   1995
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   15
         Left            =   -74880
         MaxLength       =   24
         TabIndex        =   84
         Top             =   1560
         Width           =   1995
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   16
         Left            =   -74880
         MaxLength       =   24
         TabIndex        =   83
         Top             =   1800
         Width           =   1995
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   17
         Left            =   -74880
         MaxLength       =   24
         TabIndex        =   82
         Top             =   2040
         Width           =   1995
      End
      Begin VB.TextBox TXTFIELDS 
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   18
         Left            =   -74880
         MaxLength       =   24
         TabIndex        =   81
         Top             =   2280
         Width           =   1995
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -72840
         TabIndex        =   80
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -72600
         TabIndex        =   79
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -72360
         TabIndex        =   78
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   -72120
         TabIndex        =   77
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   -71880
         TabIndex        =   76
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   -71640
         TabIndex        =   75
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   -71400
         TabIndex        =   74
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   -71160
         TabIndex        =   73
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   -72840
         TabIndex        =   72
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   -72600
         TabIndex        =   71
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   -72360
         TabIndex        =   70
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   -72120
         TabIndex        =   69
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   -71880
         TabIndex        =   68
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   -71640
         TabIndex        =   67
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   -71400
         TabIndex        =   66
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   -71160
         TabIndex        =   65
         Top             =   840
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   -72840
         TabIndex        =   64
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   18
         Left            =   -72600
         TabIndex        =   63
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   19
         Left            =   -72360
         TabIndex        =   62
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   -72120
         TabIndex        =   61
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   21
         Left            =   -71880
         TabIndex        =   60
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   22
         Left            =   -71640
         TabIndex        =   59
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   23
         Left            =   -71400
         TabIndex        =   58
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   24
         Left            =   -71160
         TabIndex        =   57
         Top             =   1080
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   25
         Left            =   -72840
         TabIndex        =   56
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   26
         Left            =   -72600
         TabIndex        =   55
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   27
         Left            =   -72360
         TabIndex        =   54
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   28
         Left            =   -72120
         TabIndex        =   53
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   29
         Left            =   -71880
         TabIndex        =   52
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   30
         Left            =   -71640
         TabIndex        =   51
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   31
         Left            =   -71400
         TabIndex        =   50
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   32
         Left            =   -71160
         TabIndex        =   49
         Top             =   1320
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   33
         Left            =   -72840
         TabIndex        =   48
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   34
         Left            =   -72600
         TabIndex        =   47
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   35
         Left            =   -72360
         TabIndex        =   46
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   36
         Left            =   -72120
         TabIndex        =   45
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   37
         Left            =   -71880
         TabIndex        =   44
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   38
         Left            =   -71640
         TabIndex        =   43
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   39
         Left            =   -71400
         TabIndex        =   42
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   40
         Left            =   -71160
         TabIndex        =   41
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   41
         Left            =   -72840
         TabIndex        =   40
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   42
         Left            =   -72600
         TabIndex        =   39
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   43
         Left            =   -72360
         TabIndex        =   38
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   44
         Left            =   -72120
         TabIndex        =   37
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   45
         Left            =   -71880
         TabIndex        =   36
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   46
         Left            =   -71640
         TabIndex        =   35
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   47
         Left            =   -71400
         TabIndex        =   34
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   48
         Left            =   -71160
         TabIndex        =   33
         Top             =   1800
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   49
         Left            =   -72840
         TabIndex        =   32
         Top             =   2040
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   50
         Left            =   -72600
         TabIndex        =   31
         Top             =   2040
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   51
         Left            =   -72360
         TabIndex        =   30
         Top             =   2040
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   52
         Left            =   -72120
         TabIndex        =   29
         Top             =   2040
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   53
         Left            =   -71880
         TabIndex        =   28
         Top             =   2040
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   54
         Left            =   -71640
         TabIndex        =   27
         Top             =   2040
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   55
         Left            =   -71400
         TabIndex        =   26
         Top             =   2040
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   56
         Left            =   -71160
         TabIndex        =   25
         Top             =   2040
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   57
         Left            =   -72840
         TabIndex        =   24
         Top             =   2280
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   58
         Left            =   -72600
         TabIndex        =   23
         Top             =   2280
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   59
         Left            =   -72360
         TabIndex        =   22
         Top             =   2280
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   60
         Left            =   -72120
         TabIndex        =   21
         Top             =   2280
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   61
         Left            =   -71880
         TabIndex        =   20
         Top             =   2280
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   62
         Left            =   -71640
         TabIndex        =   19
         Top             =   2280
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   63
         Left            =   -71400
         TabIndex        =   18
         Top             =   2280
         Width           =   255
      End
      Begin VB.CommandButton esc1 
         Caption         =   "б"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   64
         Left            =   -71160
         TabIndex        =   17
         Top             =   2280
         Width           =   255
      End
      Begin VB.TextBox TXTFIELDS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   -72000
         TabIndex        =   12
         Text            =   "0"
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox TXTFIELDS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   -72000
         TabIndex        =   11
         Text            =   "0"
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox TXTFIELDS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   -72000
         TabIndex        =   10
         Text            =   "0"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox TXTFIELDS 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   -72000
         TabIndex        =   9
         Text            =   "0"
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "Obs:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   97
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Pratileira"
         Height          =   255
         Left            =   -70680
         TabIndex        =   95
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Usa Dispositivos"
         Height          =   255
         Left            =   -69960
         TabIndex        =   94
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Usa Delimitadores"
         Height          =   255
         Left            =   -69960
         TabIndex        =   93
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label 
         Caption         =   "Horas Vida Util"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   5
         Left            =   -74760
         TabIndex        =   16
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label 
         Caption         =   "Quantidade Vida Util"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   4
         Left            =   -74760
         TabIndex        =   15
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label Label 
         Caption         =   "Horas Manutenчуo Preventiva"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   3
         Left            =   -74760
         TabIndex        =   14
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label 
         Caption         =   "Quantidade Manutenчуo Preventiva"
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   13
         Top             =   720
         Width           =   2655
      End
   End
   Begin VB.TextBox TXTFIELDS 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   0
      Text            =   "0"
      Top             =   1080
      Width           =   1095
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   5
      Left            =   1560
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1080
      Width           =   4335
   End
   Begin VB.TextBox txtCodigo 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   240
      Width           =   2055
   End
   Begin VB.TextBox txtNome 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   4215
   End
   Begin XPControls.XPButton cmdClose 
      Height          =   435
      Left            =   8280
      TabIndex        =   101
      Top             =   240
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmFER.frx":058A
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
      Left            =   8280
      TabIndex        =   102
      Top             =   960
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmFER.frx":0B24
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
   Begin XPControls.XPButton cmdFOTO 
      Height          =   495
      Left            =   8280
      TabIndex        =   103
      Top             =   1560
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   873
      Picture         =   "frmFER.frx":10BE
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
   Begin XPControls.XPButton Command4 
      Height          =   375
      Left            =   3240
      TabIndex        =   104
      TabStop         =   0   'False
      Top             =   600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "frmFER.frx":1658
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
   Begin XPControls.XPButton ESCCLI 
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "frmFER.frx":1BF2
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
   Begin XPControls.XPButton ESCCLI 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   106
      TabStop         =   0   'False
      Top             =   600
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Picture         =   "frmFER.frx":218C
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
   Begin XPControls.XPButton ESCCLI 
      Height          =   375
      Index           =   2
      Left            =   2400
      TabIndex        =   107
      TabStop         =   0   'False
      Top             =   600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Picture         =   "frmFER.frx":2726
      Caption         =   "LX2"
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
   Begin VB.Label Label 
      Caption         =   "Numero"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   6
      Left            =   6600
      TabIndex        =   100
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "CLIENTE:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nome:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   5
      Left            =   3840
      TabIndex        =   6
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label 
      Caption         =   "Codigo"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label 
      Caption         =   "Nome"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmFER"
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

Private Sub cmdClose_Click()
  On Error Resume Next
  If MDG("Gravar alteraчЇes") Then
    For iLOOP = 0 To nCAMPOS - 1
      aVAL(iLOOP) = TXTFIELDS(iLOOP)
    Next iLOOP
    GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR
  End If
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub cmdFOTO_Click()
  zgrp = TxtCodigo.tEXT
  iImage = 3
  cARQRTF = PegPath("PATH", "IMGFER")
  Load frmIMAGENS
  frmIMAGENS.TXTFIELDS(0).Enabled = False
  frmIMAGENS.Escolher(0).Visible = False
  frmIMAGENS.Show vbModal, Me
End Sub

Private Sub Command20_Click()
  Dim x
  Dim cTEMP As String
  If Not MDG("Trocar Tipo Almofada") Then
    Exit Sub
  End If
  If TXTFIELDS(6).tEXT = "6" Then
    TXTFIELDS(6).tEXT = "8"
  Else
    TXTFIELDS(6).tEXT = "6"
  End If
  For x = 1 To 8
    If TXTFIELDS(6).tEXT = "6" Then
      cTEMP = String(6, "б") + Space(2)
      If x >= 5 Then
        cTEMP = Space(8)
      End If
    Else
      cTEMP = String(8, "б")
    End If
    TXTFIELDS(10 + x).tEXT = cTEMP
  Next
End Sub

Private Sub Command4_Click()
  Dim cARQ As String
  Dim aRETU As Variant
  Dim sSQL As String
  Dim nNUMERO As Long
  nNUMERO = FixInt(TXTFIELDS(7), 0)
  cARQ = GeraConn(zMANA5EMP, "JETFOX")
  sSQL = "SELECT NOME FROM MA01 WHERE NUMERO=" & nNUMERO
  aRETU = PegSQL(cARQ, sSQL, 1, Array("NOME"), Array("C"), Array(""))
  If lRETU Then
    TXTFIELDS(8) = aRETU(0)
  End If
End Sub

Private Sub Encerrar_Click()
  If Not MDG("Sair sem gravar") Then
    Exit Sub
  End If
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub esc1_Click(Index As Integer)
  Dim linha
  Dim coluna
  Dim cTEMP As String
  Dim cDIG As String
  Dim cGRV As String
  linha = (Int((Index - 1) / 8)) + 1
  coluna = (Index + 8) - (linha * 8)
  'Alert coluna
  'ALERT LINHA
  If TXTFIELDS(6).tEXT = "6" And linha > 4 Then
    Alert ("Linha Incorreta para Formato 4x6")
    Exit Sub
  End If
  If TXTFIELDS(6).tEXT = "6" And coluna > 6 Then
    Alert ("Coluna Incorreta para Formato 4x6")
    Exit Sub
  End If
  cTEMP = TXTFIELDS(10 + linha).tEXT
  cDIG = Mid(cTEMP, coluna, 1)
  If cDIG = "l" Then
    cDIG = "б"
  Else
    cDIG = "l"
  End If
  Select Case coluna
  Case 1
    cGRV = cDIG & Mid(cTEMP, 2, 7)
  Case 8
    cGRV = Mid(cTEMP, 1, 7) + cDIG
  Case Else
    cGRV = Mid(cTEMP, 1, coluna - 1) + cDIG + Mid(cTEMP, coluna + 1)
  End Select
  TXTFIELDS(10 + linha).tEXT = cGRV
  TXTFIELDS(10 + linha).Refresh
End Sub

Private Sub ESCCLI_Click(Index As Integer)

  lRETU = False
  eRETU01 = ""
  eRETU02 = ""
  eRETU03 = ""
  cARQESC = "MA01"
  Select Case Index
  Case 0
    cARQESC = "MA01"
  Case 1
    ePASS01 = "LOGIX"
  Case 2
    ePASS01 = "LOGIX"
    cARQESC = "2DIG"
  End Select
  escNUMNOM.Show vbModal, Me

  If lRETU Then

    TXTFIELDS(7) = eRETU01
    TXTFIELDS(6) = eRETU03

  End If

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  cARQ = GeraConn(PegPath("PATH", "MANA5FER"), "SDECDX")
  cSQL = "select "
  cSQL = cSQL & " FERRAM , QTDEBASE, HRBAS, VDBAS, VDHBAS, COGCLI, CLIENTE, ESQTIP, USADEMI, USADISP, PRATILE, OBST01,"
  cSQL = cSQL & " ESQL01 , ESQL02, ESQL03, ESQL04, ESQL05, ESQL06, ESQL07, ESQL08"
  cSQL = cSQL & " FROM FERRAM WHERE FERRAM='" & LTrim(RTrim(ePASS01)) & "'"
  TxtCodigo.tEXT = ePASS01
  TxtNumero.tEXT = ePASS02
  TxtNome.tEXT = ePASS03
  nCAMPOS = 18
  aCAM = Array("QTDEBASE", "HRBAS", "VDBAS", "VDHBAS", "COGCLI", "CLIENTE", "ESQTIP", "USADEMI", "USADISP", "PRATILE", "OBST01", _
               "ESQL01", "ESQL02", "ESQL03", "ESQL04", "ESQL05", "ESQL06", "ESQL07", "ESQL08")
  aFOR = Array("N", "N", "N", "N", "C", "NI", "C", "C", "C", "C", "C", _
               "C", "C", "C", "C", "C", "C", "C", "C")
  aPAD = Array(0, 0, 0, 0, "", 0, "", "", "", "", "", _
               "", "", "", "", "", "", "", "")
  aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
  For iLOOP = 0 To nCAMPOS - 1
    TXTFIELDS(iLOOP) = aVAL(iLOOP)
  Next iLOOP
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub txtFields_GotFocus(Index As Integer)
  FocusMe
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index <> 6 Then
    KeyAscii = ValiText(KeyAscii, "#NI")
  End If
End Sub


