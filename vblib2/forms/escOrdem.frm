VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "xpcontrols.ocx"
Begin VB.Form escOrdem 
   Caption         =   "Escolha a Ordem"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5985
   Icon            =   "escOrdem.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   5985
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
   Begin XPControls.XPButton CmdCancelar 
      Height          =   435
      Left            =   4320
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "escOrdem.frx":058A
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
End
Attribute VB_Name = "escOrdem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const nFORMID = 1075
Const cFORMID = "Escolher Ordem"

Private Sub CmdCancelar_Click()
    lRETU = False
    Unload Me
End Sub

Private Sub Form_Load()
    Dim nLEN As Integer

    Center Me
    'CentralizaJanela Me

    lRETU = False
    eRETU01 = 0
    nLEN = UBound(ePASS01)
    ''nLEN = nLEN ''aRRAY comeca 0
    List.Clear
    For X = 0 To nLEN
        List.AddItem ePASS01(X)
    Next X
    HelpContextID = nFORMID
    Me.Caption = cFORMID

End Sub

Private Sub List_Click()
    lRETU = True
    eRETU01 = List.ListIndex
    Unload Me
    
End Sub

