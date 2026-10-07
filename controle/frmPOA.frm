VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{451B73A5-1563-45D5-A6AC-7B2B7D30B778}#3.0#0"; "BSPrin30.ocx"
Object = "{66E63055-5A66-4C79-9327-4BC077858695}#15.0#0"; "newtab01.OCX"
Begin VB.Form frmPOA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MENU"
   ClientHeight    =   9000
   ClientLeft      =   1092
   ClientTop       =   336
   ClientWidth     =   11760
   Icon            =   "frmPOA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin BSPrinter.PrintPreview PrintPreview1 
      Left            =   8160
      Top             =   360
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin NewTabCtl.NewTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   20
      Top             =   3240
      Width           =   11535
      _ExtentX        =   20341
      _ExtentY        =   9123
      ControlJustAdded=   0   'False
      Tabs            =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Tab             =   1
      TabHeight       =   520
      ControlVersion  =   15
      TabCaption(0)   =   "Colaborador Elabarador"
      Tab(0).ControlCount=   12
      Tab(0).Control(0)=   "txt(12)"
      Tab(0).Control(1)=   "txt(11)"
      Tab(0).Control(2)=   "txt(10)"
      Tab(0).Control(3)=   "txt(5)"
      Tab(0).Control(4)=   "txt(4)"
      Tab(0).Control(5)=   "escidfolha(0)"
      Tab(0).Control(6)=   "escidfolha(1)"
      Tab(0).Control(7)=   "Command6"
      Tab(0).Control(8)=   "Command7"
      Tab(0).Control(9)=   "Label12"
      Tab(0).Control(10)=   "Label10"
      Tab(0).Control(11)=   "lbl(4)"
      TabCaption(1)   =   "Imagem"
      Tab(1).ControlCount=   9
      Tab(1).Control(0)=   "Picture1"
      Tab(1).Control(1)=   "Picture2"
      Tab(1).Control(2)=   "CmdConfImp"
      Tab(1).Control(3)=   "CmdImprimir"
      Tab(1).Control(4)=   "CmdPaste"
      Tab(1).Control(5)=   "Incluirimagem"
      Tab(1).Control(6)=   "DelImg"
      Tab(1).Control(7)=   "VerImg"
      Tab(1).Control(8)=   "Command4"
      Begin VB.TextBox txt 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   375
         Index           =   12
         Left            =   -66720
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txt 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   375
         Index           =   11
         Left            =   -71280
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   1440
         Width           =   4095
      End
      Begin VB.TextBox txt 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   375
         Index           =   10
         Left            =   -73680
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Height          =   375
         Index           =   5
         Left            =   -70680
         TabIndex        =   25
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox txt 
         Height          =   375
         Index           =   4
         Left            =   -73680
         TabIndex        =   24
         Top             =   720
         Width           =   1110
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Left            =   4680
         ScaleHeight     =   444
         ScaleWidth      =   204
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   3975
         Left            =   120
         ScaleHeight     =   3924
         ScaleWidth      =   4164
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   480
         Width           =   4215
      End
      Begin XPControls.XPButton CmdConfImp 
         Height          =   435
         Left            =   1920
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   4560
         Width           =   1575
         _ExtentX        =   2773
         _ExtentY        =   762
         Picture         =   "frmPOA.frx":058A
         Caption         =   "Configurar Impressora"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPControls.XPButton CmdImprimir 
         Height          =   435
         Left            =   240
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   4560
         Width           =   1515
         _ExtentX        =   2667
         _ExtentY        =   762
         Picture         =   "frmPOA.frx":0B24
         Caption         =   "Imprimir"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         Left            =   -72480
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   720
         Width           =   375
         _ExtentX        =   656
         _ExtentY        =   656
         Picture         =   "frmPOA.frx":10BE
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         Left            =   -72000
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   720
         Width           =   735
         _ExtentX        =   1291
         _ExtentY        =   656
         Picture         =   "frmPOA.frx":1658
         Caption         =   "LX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPControls.XPButton Command6 
         Height          =   375
         Left            =   -71160
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   720
         Width           =   375
         _ExtentX        =   656
         _ExtentY        =   656
         Picture         =   "frmPOA.frx":1BF2
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPControls.XPButton Command7 
         Height          =   375
         Left            =   -72360
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1440
         Width           =   375
         _ExtentX        =   656
         _ExtentY        =   656
         Picture         =   "frmPOA.frx":218C
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPControls.XPButton CmdPaste 
         Height          =   435
         Left            =   3720
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   4560
         Width           =   2415
         _ExtentX        =   4255
         _ExtentY        =   762
         Picture         =   "frmPOA.frx":2726
         Caption         =   "Copia Area Transferencia"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPControls.XPButton Incluirimagem 
         Height          =   435
         Left            =   4560
         TabIndex        =   41
         Top             =   600
         Width           =   1575
         _ExtentX        =   2773
         _ExtentY        =   762
         Picture         =   "frmPOA.frx":2CC0
         Caption         =   "Incluir Image"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPControls.XPButton DelImg 
         Height          =   435
         Left            =   4560
         TabIndex        =   42
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2773
         _ExtentY        =   762
         Picture         =   "frmPOA.frx":315A
         Caption         =   "Excluir Image"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPControls.XPButton VerImg 
         Height          =   495
         Left            =   4560
         TabIndex        =   43
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2773
         _ExtentY        =   868
         Picture         =   "frmPOA.frx":35F4
         Caption         =   "Navegar Imagens"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPControls.XPButton Command4 
         Height          =   435
         Left            =   4560
         TabIndex        =   44
         Top             =   2520
         Width           =   1575
         _ExtentX        =   2773
         _ExtentY        =   762
         Picture         =   "frmPOA.frx":3A8E
         Caption         =   "Salvar Imagem"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label12 
         Caption         =   "Em"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -67080
         TabIndex        =   29
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "Elaborador"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   -74760
         TabIndex        =   26
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Colaborador"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   4
         Left            =   -74760
         TabIndex        =   23
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.TextBox SSQ 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
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
      TabIndex        =   17
      TabStop         =   0   'False
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   120
      Width           =   732
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   330
      Index           =   8
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1800
      Width           =   750
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Gera Sac"
      Height          =   375
      Left            =   8280
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1455
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   330
      Index           =   9
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   330
      Index           =   7
      Left            =   8520
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1800
      Width           =   750
   End
   Begin VB.TextBox txt 
      Height          =   690
      Index           =   6
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2400
      Width           =   6255
   End
   Begin VB.TextBox txt 
      Height          =   690
      Index           =   3
      Left            =   1080
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1560
      Width           =   6255
   End
   Begin VB.TextBox txt 
      Height          =   330
      Index           =   2
      Left            =   720
      TabIndex        =   4
      Top             =   1080
      Width           =   4575
   End
   Begin VB.TextBox txt 
      Height          =   315
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   3390
   End
   Begin VB.TextBox txt 
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
      Left            =   10080
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   600
      Width           =   1515
      _ExtentX        =   2667
      _ExtentY        =   762
      Picture         =   "frmPOA.frx":3F28
      Caption         =   "Retornar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton cmdClose 
      Height          =   435
      Left            =   10080
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   120
      Width           =   1515
      _ExtentX        =   2667
      _ExtentY        =   762
      Picture         =   "frmPOA.frx":44C2
      Caption         =   "Salvar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton Command5 
      Height          =   375
      Left            =   5160
      TabIndex        =   35
      Top             =   600
      Width           =   375
      _ExtentX        =   656
      _ExtentY        =   656
      Picture         =   "frmPOA.frx":4A5C
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      Left            =   4200
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   600
      Width           =   975
      _ExtentX        =   1715
      _ExtentY        =   656
      Picture         =   "frmPOA.frx":4FF6
      Caption         =   "mana5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton Command1 
      Height          =   375
      Left            =   5520
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   600
      Width           =   735
      _ExtentX        =   1291
      _ExtentY        =   656
      Picture         =   "frmPOA.frx":5590
      Caption         =   "PF"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      Left            =   6240
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   600
      Width           =   735
      _ExtentX        =   1291
      _ExtentY        =   656
      Picture         =   "frmPOA.frx":5B2A
      Caption         =   "LX"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      Left            =   6960
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   600
      Width           =   735
      _ExtentX        =   1291
      _ExtentY        =   656
      Picture         =   "frmPOA.frx":60C4
      Caption         =   "MC"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "SSQ"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   4920
      TabIndex        =   18
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Seq"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   3480
      TabIndex        =   16
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label3 
      Caption         =   "PF"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   2040
      TabIndex        =   14
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sac:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   7
      Left            =   7680
      TabIndex        =   10
      Top             =   1440
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Analise:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Problema:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   630
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
      Top             =   600
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
Attribute VB_Name = "frmPOA"
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
      aVAL(iLOOP) = txt(iLOOP)
    Next iLOOP
    GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR, 1 'pula 1 chave numero
    If lTROCOU Then
      'cSQL = "select * from poa WHERE numero=" & nPPAP
      ADOGrvBlob cARQ, "POA", Picture1, "FOTO", "numero=" & nPPAP
    End If
  End If
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub CmdConfImp_Click()
'  FrmPrintSetup.Show vbModal, Me
End Sub
Private Sub PrintPreview1_PrepareReport(Cancel As Boolean)
  Printer.PaintPicture Picture1, 0, 0
End Sub
Private Sub CmdImprimir_Click()
  If Picture1.Height = 0 Then
    Alert ("Sem Imagem")
  Else
    PrintPreview1.ShowPreview
  End If
End Sub

Private Sub CmdPaste_Click()
' Verifica se existe uma imagem válida carregada no Picture1
  If Picture1.Height = 0 Or Picture1.Picture = 0 Then
    Alert "Sem Imagem"
  Else
    ' Instancia a stdImage a partir da imagem atual e envia para a área de transferência
    Dim imgObj As Object
    Set imgObj = stdImage.CreateFromStdPicture(Picture1.Picture)
    imgObj.ToClipboard
  End If
End Sub

Private Sub Command1_Click()
  escpffim.Show vbModal, Me
  If lRETU Then
    txt(1) = eRETU02
    txt(2) = eRETU03
  End If
End Sub

Private Sub Command2_Click()
  Dim cARQ As String
  Dim aRETU As Variant
  Dim sSQL As String
  Dim nNUMERO As Long
  nNUMERO = FixInt(txt(4), 0)

  cARQ = PegPath("PATH", "LOGIXODBC")
  sSQL = "SELECT nom_completo as NOMTEC FROM funcionario WHERE cod_empresa='01' and num_matricula=" & nNUMERO
  aRETU = PegSQL(cARQ, sSQL, 1, Array("NOMTEC"), Array("C"), Array(""))
  If Not lRETU Then
    cARQ = PegPath("PATH", "CADMP04")
    cARQ = GeraConn(cARQ, "JETFOX")
    sSQL = "SELECT NOMTEC FROM MP04 WHERE TECNICO=" & nNUMERO
    aRETU = PegSQL(cARQ, sSQL, 1, Array("NOMTEC"), Array("C"), Array(""))
  End If
  If lRETU Then
    txt(5) = aRETU(0)
  End If
End Sub

Private Sub Command3_Click()
  Dim nSAC As Long
  Dim nLen As Long
  Dim cDESC01 As String
  Dim cDESC02 As String
  Dim cDESC03 As String
  Dim cDESC04 As String
  Dim cPROBLEMA As String
  Dim cCAM As String
  Dim cData As Variant
 ' Dim SACAREA As Variant
 ' Dim iRETVAL As Variant
  Dim cCAMJET As String
  Dim cCOMJET As String
  Dim aCAMJET As Variant
  Dim aVALJET As Variant

  If txt(8) = "N" Then
    Alert ("Ja respondida")
    Exit Sub
  End If

  If txt(8) = "S" And FixNum(txt(7)) > 0 Then
    Alert ("Ja respondida")
    Exit Sub
  End If

  If Not MDG("Gerar Sac") Then
    txt(8) = "N"
    txt(9) = Today()
    Exit Sub
  End If
  txt(8) = "S"
  txt(9) = Today()

  cPROBLEMA = FixStr(txt(3))
  nLen = Len(cPROBLEMA)

  cDESC01 = Mid(cPROBLEMA, 1, 100)
  If nLen > 100 Then
    cDESC02 = Mid(cPROBLEMA, 101, 100)
  End If
  If nLen > 200 Then
    cDESC03 = Mid(cPROBLEMA, 201, 100)
  End If
  If nLen > 300 Then
    cDESC04 = Mid(cPROBLEMA, 301)
  End If

  cCAMJET = PegPath("PATH", "MANA5TGQ")
  cCOMJET = GeraConn(cCAM, "FOX")
  cData = format(Date, "DD/MM/YY")
  nSAC = PegMAXSQL(cCOMJET, "SAC", "SAC", 1)
  nSAC = nSAC + 1
  
  
 
     aCAMJET = Array("SAC", "INCUSER", "POA", "CODIGO", "NOME", "DESC01", _
     "DESC02", "DESC03", "DESC04", "DOCUMENTO", "DATA", "INCDATA")
     aVALJET = Array(nSAC, zUSER, FixNum(txt(0)), FixStr(txt(1)), FixStr(cDESC01), _
                    FixStr(cDESC02), FixStr(cDESC03), FixStr(cDESC04), "Prog. Olhos Abertos", cData, cData)
    IncluiSQL cCOMJET, "SELECT * FROM SAC WHERE SAC=" & nSAC, 12 _
           , aCAMJET _
           , aVALJET, _
           , False, True
           
  txt(7) = nSAC


End Sub

Private Sub Command4_Click()
  salvarpict Me, Picture1, StrZero(txt(0), 8)
End Sub

Private Sub Command5_Click()
  Dim cARQ As String
  Dim aRETU As Variant
  Dim sSQL As String
  Dim cCODIGO As String
  cCODIGO = FixStr(txt(1), "", "TRIM")
  cARQ = GeraConn(zMANA5EMP, "JETFOX")
  sSQL = "SELECT NOME FROM MS01 WHERE CODIGO='" & cCODIGO & "'"
  aRETU = PegSQL(cARQ, sSQL, 1, Array("NOME"), Array("C"), Array(""))
  If lRETU Then
    txt(2) = aRETU(0)
  End If
End Sub

Private Sub Command6_Click()
  txt(4) = zIDFOLHA
  txt(5) = zNOMEFOLHA
End Sub

Private Sub Command7_Click()
  txt(10) = zIDFOLHA
  txt(11) = zNOMEFOLHA
  txt(12) = Date
End Sub

Private Sub PegCodigoDescricaoPf()
  Dim sSQL As String
  Dim aRETU As Variant
  Dim sARQ As String
  sARQ = PegPath("PATH", "PF")
  sSQL = "SELECT CODIGO,DESCR FROM PF WHERE PF=" & nPF
  aRETU = PegSQL(sARQ, sSQL, 2, Array("CODIGO", "DESCR"), Array("C", "C"), Array("", ""))
  If lRETU Then
    txt(1) = aRETU(0)
    txt(2) = aRETU(1)
    txt(1).Enabled = False
    txt(2).Enabled = False
    txt(1).Locked = True
    txt(2).Locked = True
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

Private Sub escidfolha_Click(Index As Integer)
  ePASS01 = ""
  If Index = 1 Then
    ePASS01 = "LOGIX"
  End If
  escMP04.Show vbModal, Me

  If lRETU Then
    txt(4) = eRETU01
    txt(5) = eRETU02
  End If
End Sub

Private Sub ESCMS01A_Click(Index As Integer)
  Dim cCHAVEBUS As String
  cCHAVEBUS = txt(1)
  ePASS01 = "MANA5"
  If Index = 1 Then
    ePASS01 = "LOGIX"
  End If
  If Index = 2 Then
    ePASS01 = "MICRO"
  End If
  escms01.Show vbModal, Me
  If lRETU Then
    txt(1) = eRETU01
    txt(2) = eRETU02
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
  PF.tEXT = nPF
  SEQ.tEXT = nSEQ
  SSQ.tEXT = nSSQ
  CenterFormToScreen Me

  lTROCOU = False
  cARQ = PegPath("PATH", "POA")
  cSQL = "select * from poa WHERE numero=" & nPPAP
  nCAMPOS = 13
  aCAM = Array("NUMERO", "CODIGO", "NOME", "PROBLEMA", "NUMFUN", "NOMFUN", "ANALISE", "SAC", "SACSN", "DATASAC", "ELANUM", "ELANOM", "ELADAT")
  aFOR = Array("NI", "C", "C", "C", "NI", "C", "C", "NI", "C", "DZ", "NI", "C", "DZ")
  aPAD = Array(0, "", "", "", 0, "", "", 0, "", "", 0, "", "")
  aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
  For iLOOP = 0 To nCAMPOS - 1
    txt(iLOOP) = aVAL(iLOOP)
  Next iLOOP
  If ADOPegBlob(Picture1, cARQ, "POA", "numero=" & nPPAP, "FOTO") Then  ' ADOPegBlob(cARQ, cSQL, Picture1, "FOTO") Then
    StretchSourcePictureFromPicture Picture1, Picture2
    If FixNum(eRETU01) > 500000 Then
      Alert ("Imagem Muito Grande,Ajuste o tamanho")
      salvarpict Me, Picture1, "POA_" & StrZero(nPPAP, 6)
      Set Picture1.Picture = Nothing
      Set Picture2.Picture = Nothing
      lTROCOU = True
    End If
  Else
    Set Picture1.Picture = Nothing
    Set Picture2.Picture = Nothing
  End If
  If nPF > 0 Then
    PegCodigoDescricaoPf
  End If

  PrintPreview1.AuxiliaryButtonVisible = PrintPreview1.PrinterExists("Microsoft Print to PDF")
  PrintPreview1.AuxiliaryButtonToolTipText = "Salvar como PDF"
End Sub
Public Sub PrintPreview1_AuxiliaryButtonClick(UpdateReport As Boolean)
  PrintPreview1.ShowSaveToFile "Microsoft Print to PDF", "*.pdf"
  UpdateReport = False  ' we don't need to update the report in the Print preview window after this action (the default value of UpdateReport parameter is True)
End Sub
Private Sub Incluirimagem_Click()
  Dim STMPFILE As String
  STMPFILE = OpenArqExt(Me, "", "JPG", "JPEG *.JPG")
  If lerarquivoimagem(STMPFILE, Picture1, Picture2) Then
    lTROCOU = True
  End If
End Sub

Private Sub VerImg_Click()
'  frmPicViewer.Show vbModal, Me
'  If lRETU Then
'    If lerarquivoimagem(eRETU01, Picture1, Picture2) Then
'      lTROCOU = True
'    End If
'  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub xCmdImprimir_Click()

End Sub

