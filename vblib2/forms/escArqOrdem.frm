VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "VBCCR18.OCX"
Begin VB.Form EscArqOrdem 
   Caption         =   "Escolher Ordem"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   6528
   Icon            =   "escArqOrdem.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3480
   ScaleWidth      =   6528
   StartUpPosition =   3  'Windows Default
   Begin XPControls.XPButton CmdZerar 
      Height          =   852
      Left            =   5160
      TabIndex        =   2
      Top             =   1680
      Width           =   1092
      _ExtentX        =   1926
      _ExtentY        =   1503
      Caption         =   "Apagar Ordem"
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
   Begin VB.TextBox TxtOrdem 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Text            =   "TxtOrdem"
      Top             =   120
      Width           =   4815
   End
   Begin VB.ListBox List 
      Height          =   2544
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4815
   End
   Begin VBCCR18.CommandButtonW encerrar 
      Height          =   612
      Left            =   5160
      TabIndex        =   3
      Top             =   840
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Retornar"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "escArqOrdem.frx":058A
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW CmdEscolher 
      Height          =   612
      Left            =   5160
      TabIndex        =   4
      Top             =   120
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Escolher"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "escArqOrdem.frx":0B24
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
End
Attribute VB_Name = "EscArqOrdem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const nFORMID = 1075
Const cFORMID = "Escolher Ordem"

Private Sub CmdEscolher_Click()
  lRETU = True
  eRETU01 = TxtOrdem
  Unload Me
End Sub

Private Sub CmdSair_Click()

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
  Dim cARQ As String
  Dim cSQL As String
  Dim oRS As New ADODB.Recordset
  Dim oCONN As New ADODB.Connection
  Dim x As Integer
  
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

