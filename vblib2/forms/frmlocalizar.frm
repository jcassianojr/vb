VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.0#0"; "vbccr18.ocx"
Begin VB.Form frmLocalizar 
   Caption         =   "Localizar"
   ClientHeight    =   2025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6345
   Icon            =   "frmlocalizar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2025
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
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
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Preencha com A Chave de Busca"
      Top             =   480
      Width           =   4395
   End
   Begin XPControls.XPButton Command2 
      Height          =   435
      Left            =   4680
      TabIndex        =   3
      Top             =   600
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmlocalizar.frx":058A
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
   Begin vbccr18.SpinBox nIND 
      Height          =   615
      Left            =   2160
      TabIndex        =   4
      Top             =   1200
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   1
      Max             =   99
      Value           =   1
   End
   Begin XPControls.XPButton Command1 
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Picture         =   "frmlocalizar.frx":0B24
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
   Begin VB.Label Label2 
      Caption         =   "Digite a Coluna de Busca"
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
      Height          =   420
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1995
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2355
   End
End
Attribute VB_Name = "frmLocalizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const nFORMID = 1072
Const cFORMID = "Dialogo de Busca"

Private Sub cBUSCA_GotFocus()
    FocusMe
End Sub

Private Sub cBUSCA_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Or KeyCode = 40 Then
        SendKeys Chr(9)
    End If
    If KeyCode = 38 Then
        SendKeys "+" + Chr(9)
    End If
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
    CenterFormToScreen Me
    Me.Caption = cFORMID
    HelpContextID = nFORMID
    
    lRETU = False
    'frmLocalizar.Left = (Screen.Width - frmLocalizar.Width) / 2
    'frmLocalizar.Top = (Screen.Height - frmLocalizar.Height) / 2
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    eRETU01 = cBUSCA
    eRETU02 = nIND
End Sub

Private Sub mais_Click()
    nIND = CStr(IncDec(nIND, 1, 1, 99))
End Sub

Private Sub menos_Click()
    nIND = CStr(IncDec(nIND, -1, 1, 99))
End Sub

Private Sub nIND_GotFocus()
    FocusMe
End Sub

Private Sub nIND_KeyPress(KeyAscii As Integer)
    KeyAscii = ValiText(KeyAscii, "#NI")
End Sub

Private Sub nIND_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Or KeyCode = 40 Then
        SendKeys Chr(9)
    End If
    If KeyCode = 38 Then
        SendKeys "+" + Chr(9)
    End If
End Sub

