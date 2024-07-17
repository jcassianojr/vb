VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.1#0"; "vbccr18.ocx"
Begin VB.Form NewOrd 
   Caption         =   "Incluindo Nova Ordem"
   ClientHeight    =   1365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3120
   Icon            =   "NewOrd.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1365
   ScaleWidth      =   3120
   StartUpPosition =   2  'CenterScreen
   Begin XPControls.XPButton grava 
      Height          =   435
      Left            =   1320
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   480
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "NewOrd.frx":058A
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
   Begin VBCCR18.SpinBox ord 
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   1
      Max             =   999
      Value           =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Ordem No."
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "NewOrd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  ord.Value = nNOVAPOS
End Sub

Private Sub Grava_Click()
  Dim nORDEM As Long
  Dim sSQL As String
  nORDEM = FixInt(NewOrd.ord.Value)
  Select Case iMU01
  Case 1
    sSQL = "PFI"
  Case 3
    sSQL = "PFC"
  Case 4
    sSQL = "PFCO"
  Case 5
    sSQL = "PFCMS03"
  Case 6
    sSQL = "PFQSBLEP"
  Case 7
    sSQL = "PFSCHECKLIST"
  Case 8
    sSQL = "PFSONTHEJOB"
  Case 9
    sSQL = "DISPOI"
  End Select
  If nORDEM = 0 Then
    Alert "Ordem nao pode ser zero"
    Exit Sub
  End If
  Select Case iMU01
  Case 9
    sSQL = "select * from " & sSQL & " where numero=" & nPPAP & " and item=" & nORDEM
    IncluiSQL PegPath("PATH", "DISPO"), sSQL, 2, Array("NUMERO", "ITEM"), _
              Array(nPPAP, nORDEM)
  Case 7, 8
    sSQL = "select * from " & sSQL & " where pf=" & nPF & " and seq=" & nSEQ & " and ssq=" & nSSQ & " and ordem=" & nORDEM
    IncluiSQL Sdb, sSQL, 4, Array("PF", "SEQ", "SSQ", "ORDEM"), _
              Array(nPF, nSEQ, nSSQ, nORDEM)
  Case 5
    sSQL = "select * from " & sSQL & " where pf=" & nPF & " and seq=" & nSEQ & " and ssq=" & nSSQ & " and item=" & nORDEM & " and codcomp='" & Ccodcomp & "'"
    IncluiSQL Sdb, sSQL, 5, Array("PF", "SEQ", "SSQ", "ITEM", "CODCOMP"), _
              Array(nPF, nSEQ, nSSQ, nORDEM, Ccodcomp)
  Case Else
    If nSEQ = 0 And nSSQ = 0 Then
      Alert "Sequencia e subsequencia zeradas"
    Else
      sSQL = "select * from " & sSQL & " where pf=" & nPF & " and seq=" & nSEQ & " and ssq=" & nSSQ & " and item=" & nORDEM
      IncluiSQL Sdb, sSQL, 4, Array("PF", "SEQ", "SSQ", "ITEM"), _
                Array(nPF, nSEQ, nSSQ, nORDEM)
    End If
  End Select
  eRETU01 = nORDEM
  Unload Me
End Sub

Private Sub ORD_GotFocus()
  FocusMe
End Sub

Private Sub ORD_KeyPress(KeyAscii As Integer)
  KeyAscii = ValiText(KeyAscii, "#NI")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub


