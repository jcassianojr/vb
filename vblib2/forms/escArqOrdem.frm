VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form EscArqOrdem 
   Caption         =   "Escolher Ordem"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6525
   Icon            =   "escArqOrdem.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin XPControls.XPButton CmdEscolher 
      Height          =   495
      Left            =   5040
      TabIndex        =   3
      Top             =   120
      Width           =   1335
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Escolher Ordem"
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
   Begin XPControls.XPButton CmdZerar 
      Height          =   495
      Left            =   5040
      TabIndex        =   2
      Top             =   840
      Width           =   1335
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Apagar Ordem"
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
   Begin VB.TextBox TxtOrdem 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "TxtOrdem"
      Top             =   120
      Width           =   4815
   End
   Begin VB.ListBox List 
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4815
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   5040
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   767
      Picture         =   "escArqOrdem.frx":058A
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
End
Attribute VB_Name = "EscArqOrdem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const nFORMID = 1075
Const cFORMID = "Escolher Ordem"

Private Sub CmdEscolher_Click()
  lRETU = True
  eRETU01 = TxtOrdem
  Unload Me
End Sub

Private Sub CmdZerar_Click()
  TxtOrdem = ""
End Sub

Private Sub Encerrar_Click()
  lRETU = False
  eRETU01 = ""
  Unload Me
End Sub

Private Sub Form_Load()
'Dim nLEN As Integer
  Dim cARQ As String
  Dim cSQL As String
  Dim oRS As New ADODB.Recordset
  Dim oCONN As New ADODB.Connection

  '  CentralizaJanela Me
  CenterFormToScreen Me
  lRETU = False
  TxtOrdem = ""
  eRETU01 = ""
  cARQ = ePASS01
  cSQL = ePASS02
  cARQ = GeracArq(cARQ, , False)
  oCONN.ConnectionTimeout = 120
  oCONN.Open cARQ
  oRS.Open cSQL, oCONN, adOpenForwardOnly, adLockReadOnly
  List.Clear
  For x = 0 To oRS.Fields.Count - 1
    List.AddItem oRS.Fields(x).Name
  Next x
  HelpContextID = nFORMID
  Me.Caption = cFORMID
End Sub

Private Sub List_Click()
  If TxtOrdem = "" Then
    TxtOrdem = List
  Else
    TxtOrdem = TxtOrdem & "," & List
  End If
End Sub

