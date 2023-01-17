VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{EA478B61-D9EC-47F6-BB21-95A533AF2251}#1.0#0"; "TabExC01.OCX"
Object = "{451B73A5-1563-45D5-A6AC-7B2B7D30B778}#1.0#0"; "BSPrin10.ocx"
Begin VB.Form frmPOA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MENU"
   ClientHeight    =   9000
   ClientLeft      =   1095
   ClientTop       =   330
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
      _ExtentX        =   1191
      _ExtentY        =   1191
      LcK1b           =   "yefT59bnyufI583n1ufV59HnyefN58nn0+fK58nnzefR58rn1+fW5w=="
      LcK2b           =   "reeI58zn+ue059jn1Of656fni+eN543nleeM54HnlueR5w=="
      AmbientBb       =   $"frmPOA.frx":058A
   End
   Begin TabExCtl.SSTabEx SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   20
      Top             =   3240
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9128
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
         Left            =   8280
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
         Left            =   3720
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
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txt 
         Height          =   375
         Index           =   5
         Left            =   4320
         TabIndex        =   25
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox txt 
         Height          =   375
         Index           =   4
         Left            =   1320
         TabIndex        =   24
         Top             =   720
         Width           =   1110
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Left            =   -70320
         ScaleHeight     =   435
         ScaleWidth      =   195
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         Height          =   3975
         Left            =   -74880
         ScaleHeight     =   3915
         ScaleWidth      =   4155
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   480
         Width           =   4215
      End
      Begin XPControls.XPButton CmdConfImp 
         Height          =   435
         Left            =   -73080
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   4560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPOA.frx":062A
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
      Begin XPControls.XPButton CmdImprimir 
         Height          =   435
         Left            =   -74760
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   4560
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
         Picture         =   "frmPOA.frx":0BC4
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
      Begin XPControls.XPButton escidfolha 
         Height          =   375
         Index           =   0
         Left            =   2520
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   720
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Picture         =   "frmPOA.frx":115E
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
         Left            =   3000
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   720
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         Picture         =   "frmPOA.frx":16F8
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
      Begin XPControls.XPButton Command6 
         Height          =   375
         Left            =   3840
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   720
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Picture         =   "frmPOA.frx":1C92
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
      Begin XPControls.XPButton Command7 
         Height          =   375
         Left            =   2640
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1440
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         Picture         =   "frmPOA.frx":222C
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
      Begin XPControls.XPButton CmdPaste 
         Height          =   435
         Left            =   -71280
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   4560
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   767
         Picture         =   "frmPOA.frx":27C6
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
      Begin XPControls.XPButton Incluirimagem 
         Height          =   435
         Left            =   -70440
         TabIndex        =   41
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPOA.frx":2D60
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
         Left            =   -70440
         TabIndex        =   42
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPOA.frx":31FA
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
         Left            =   -70440
         TabIndex        =   43
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   873
         Picture         =   "frmPOA.frx":3694
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
         Left            =   -70440
         TabIndex        =   44
         Top             =   2520
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "frmPOA.frx":3B2E
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
      Begin VB.Label Label12 
         Caption         =   "Em"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   7920
         TabIndex        =   29
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label10 
         Caption         =   "Elaborador"
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
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
         Left            =   240
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
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmPOA.frx":3FC8
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
      Left            =   10080
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmPOA.frx":4562
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
   Begin XPControls.XPButton Command5 
      Height          =   375
      Left            =   5160
      TabIndex        =   35
      Top             =   600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "frmPOA.frx":4AFC
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
      Left            =   4200
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Picture         =   "frmPOA.frx":5096
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
      Left            =   5520
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   600
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Picture         =   "frmPOA.frx":5630
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
      Left            =   6240
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   600
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Picture         =   "frmPOA.frx":5BCA
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
      Left            =   6960
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   600
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Picture         =   "frmPOA.frx":6164
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
        GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR
        If lTROCOU Then
            ADOGrvBlob cARQ, cSQL, Picture1, "FOTO"
        End If
    End If
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub CmdConfImp_Click()
    FrmPrintSetup.Show vbModal, Me
End Sub
Private Sub PrintPreview1_PrepareReport(Cancel As Boolean)
   Printer.PaintPicture Picture1, 0, 0
   Printer.EndDoc
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
    Dim sSQL             As String
    Dim nNUMERO          As Long
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
    Dim nSAC, nLEN As Long
    Dim cDESC01, cDESC02, cDESC03, cDESC04, cPROBLEMA, cCAM As String
    Dim cData As Variant
    Dim SACAREA As Variant
    Dim iRETVAL As Variant
   
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
    nLEN = Len(cPROBLEMA)

    cDESC01 = Mid(cPROBLEMA, 1, 100)
    If nLEN > 100 Then
        cDESC02 = Mid(cPROBLEMA, 101, 100)
    End If
    If nLEN > 200 Then
        cDESC03 = Mid(cPROBLEMA, 201, 100)
    End If
    If nLEN > 300 Then
        cDESC04 = Mid(cPROBLEMA, 301)
    End If

    cData = Format(Date, "DD/MM/YY")
    cCAM = PegPath("PATH", "MANA5TGQ")
    sx_SetEpoch (Year(Date) - 30)
    sx_SetDateFormat BRITISH
    sx_SetDeleted True
    SACAREA = sx_Use(cCAM & "sac.dbf", "sac", READWRITE, SDEFOX)
    iRETVAL = sx_SetOrder(1)
    sx_GoBottom
    nSAC = FixInt(sx_GetVariant("SAC"))
    nSAC = nSAC + 1
    sx_AppendBlank
    If sx_Rlock(sx_RecNo()) Then
        sx_PutVariant "SAC", nSAC
        sx_PutVariant "INCUSER", zUSER
        sx_PutVariant "POA", FixNum(txt(0))
        sx_PutVariant "CODIGO", FixStr(txt(1))
        sx_PutVariant "NOME", FixStr(txt(2))
        sx_PutVariant "DESC01", FixStr(cDESC01)
        sx_PutVariant "DESC02", FixStr(cDESC02)
        sx_PutVariant "DESC03", FixStr(cDESC03)
        sx_PutVariant "DESC04", FixStr(cDESC04)
        sx_PutVariant "DOCUMENTO", "Prog. Olhos Abertos"
        sx_PutVariant "DATA", cData
        sx_PutVariant "INCDATA", cData
        sx_Commit
    End If
    sx_Unlock sx_RecNo()
    sx_CloseAll
    txt(7) = nSAC


End Sub

Private Sub Command4_Click()
    salvarpict Me, Picture1, StrZero(txt(0), 8)
End Sub

Private Sub Command5_Click()
    Dim cARQ As String
    Dim aRETU As Variant
    Dim sSQL             As String
    Dim cCODIGO          As String
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
    Dim sSQL, aRETU, sARQ
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
    PF.Text = nPF
    SEQ.Text = nSEQ
    SSQ.Text = nSSQ
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
    If ADOPegBlob(cARQ, cSQL, Picture1, "FOTO") Then
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
End Sub

Private Sub Incluirimagem_Click()
    Dim STMPFILE
    STMPFILE = OpenArqExt(Me, "", "JPG", "JPEG *.JPG")
    If lerarquivoimagem(STMPFILE, Picture1, Picture2) Then
        lTROCOU = True
    End If
End Sub

Private Sub VerImg_Click()
    frmPicViewer.Show vbModal, Me
    If lRETU Then
        If lerarquivoimagem(eRETU01, Picture1, Picture2) Then
            lTROCOU = True
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

Private Sub xCmdImprimir_Click()

End Sub

