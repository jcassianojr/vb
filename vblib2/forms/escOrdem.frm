VERSION 5.00
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "VBCCR18.OCX"
Begin VB.Form escOrdem 
   Caption         =   "Escolha a Ordem"
   ClientHeight    =   3192
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   5988
   Icon            =   "escOrdem.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3192
   ScaleWidth      =   5988
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List 
      Height          =   2928
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4452
   End
   Begin VBCCR18.CommandButtonW CmdCancelar 
      Height          =   612
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Confimar"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "escOrdem.frx":058A
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
End
Attribute VB_Name = "escOrdem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const nFORMID = 1075
Const cFORMID = "Escolher Ordem"
Option Explicit

Private Sub CmdCancelar_Click()
  lRETU = False
  Unload Me
End Sub

Private Sub Form_Load()
  Dim nLEN As Integer
  Dim x As Integer


  CenterFormToScreen Me
  Me.Caption = cFORMID
  HelpContextID = nFORMID

  'CentralizaJanela Me

  lRETU = False
  eRETU01 = 0
  nLEN = UBound(ePASS01)
  ''nLEN = nLEN ''aRRAY comeca 0
  List.Clear
  For x = 0 To nLEN
    List.AddItem ePASS01(x)
  Next x
  HelpContextID = nFORMID
  Me.Caption = cFORMID

End Sub

Private Sub List_Click()
  lRETU = True
  eRETU01 = List.ListIndex
  Unload Me

End Sub

