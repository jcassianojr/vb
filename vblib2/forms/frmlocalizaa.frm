VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form frmLocalizaa 
   Caption         =   "Localizar"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6285
   Icon            =   "frmlocalizaa.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List 
      Height          =   1815
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   3975
   End
   Begin VB.TextBox cBUSCA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Preencha com A Chave de Busca"
      Top             =   600
      Width           =   4035
   End
   Begin XPControls.XPButton Command2 
      Height          =   435
      Left            =   4440
      TabIndex        =   3
      Top             =   1200
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmlocalizaa.frx":058A
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
   Begin XPControls.XPButton Command1 
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Picture         =   "frmlocalizaa.frx":0B24
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
   Begin VB.Label Label1 
      Caption         =   "Digite a Chave de Busca"
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
      Height          =   330
      Left            =   240
      TabIndex        =   1
      Top             =   210
      Width           =   2355
   End
End
Attribute VB_Name = "frmLocalizaa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const nFORMID = 1072
Const cFORMID = "Dialogo de Busca"
Option Explicit

Private Sub cBUSCA_GotFocus()
  FocusMe
End Sub

Private Sub Command1_Click()
  lRETU = True
  Unload Me
End Sub

Private Sub Command2_Click()
  lRETU = False
  Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
  Dim nLEN As Integer
  Dim X As Integer
  CenterFormToScreen Me
  Me.Caption = cFORMID
  HelpContextID = nFORMID

  'CentralizaJanela Me
  lRETU = False
  eRETU01 = ""
  eRETU02 = 1
  nLEN = UBound(ePASS01)
  ''nLEN = nLEN ''aRRAY comeca 0
  List.Clear
  For X = 0 To nLEN
    List.AddItem ePASS01(X)
  Next X
  List.ListIndex = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  eRETU01 = cBUSCA
  eRETU02 = List.ListIndex + 1                 ''matriz zero posional 1
End Sub

