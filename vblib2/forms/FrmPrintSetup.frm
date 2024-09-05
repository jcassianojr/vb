VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form FrmPrintSetup 
   Caption         =   "Configurar Impressoras"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7125
   Icon            =   "FrmPrintSetup.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   7125
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDefault 
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
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   6615
   End
   Begin VB.ComboBox cboPrinters 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1560
      Width           =   6555
   End
   Begin XPControls.XPButton Command9 
      Height          =   435
      Left            =   3240
      TabIndex        =   5
      Top             =   2040
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FrmPrintSetup.frx":058A
      Caption         =   "Propriedades"
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
      Left            =   5160
      TabIndex        =   6
      Top             =   2040
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "FrmPrintSetup.frx":0B24
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
   Begin XPControls.XPButton Command6 
      Height          =   435
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   767
      Picture         =   "FrmPrintSetup.frx":10BE
      Caption         =   "Adcionar Nova Impressora"
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
      Left            =   3240
      TabIndex        =   8
      Top             =   2520
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FrmPrintSetup.frx":1558
      Caption         =   "configuração "
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
   Begin VB.Label Label1 
      Caption         =   "Escollher Outra Impressora como Padrao"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label lbl1 
      Caption         =   "Impressora Padrao Atual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   135
      TabIndex        =   3
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lista Impressoras e Define a Padrão"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   4800
   End
End
Attribute VB_Name = "FrmPrintSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const nFORMID = 1111
Const cFORMID = "frmprintsetup"
Option Explicit
Dim DLG As New clsCommonDialogs
Dim OldDefaultPrinter As String
Dim Loaded As Boolean
Private Sub Command6_Click()
  DLG.DialogConnectToPrinter
End Sub
Private Sub Command9_Click()
  DLG.DialogPrint

End Sub
Private Sub Encerrar_Click()
  Unload Me
End Sub
Private Sub cboPrinters_Click()
  Dim OUSO
  If Loaded Then
    Me.MousePointer = vbHourglass
    OUSO = DLG.GetPrinterFromDevice(cboPrinters.Text)
    '' If DLG.SetDefaultPrinter(Byref OUSO) Then
    ''     txtDefault.tEXT = cboPrinters.tEXT
    '' End If
    Me.MousePointer = vbNormal
  End If
End Sub
Private Sub Form_Load()
  CenterFormToScreen Me
  Me.Caption = cFORMID
  HelpContextID = nFORMID

  DLG.hWnd = Me.hWnd
  txtDefault.Text = DLG.GetDefaultPrinter
  OldDefaultPrinter = txtDefault.Text
  DLG.LoadPrinters cboPrinters, False
  Loaded = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub XPButton1_Click()
  DLG.DialogPrintSetup
End Sub
