VERSION 5.00
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "VBCCR18.OCX"
Begin VB.Form frmLocalizaa 
   Caption         =   "Localizar"
   ClientHeight    =   3168
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   5652
   Icon            =   "frmlocalizaa.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3168
   ScaleWidth      =   5652
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List 
      Height          =   1776
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   4092
   End
   Begin VB.TextBox cBUSCA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      Width           =   4032
   End
   Begin VBCCR18.CommandButtonW Command2 
      Height          =   612
      Left            =   4440
      TabIndex        =   3
      Top             =   1320
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Cancelar"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "frmlocalizaa.frx":058A
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW Command1 
      Height          =   612
      Left            =   4440
      TabIndex        =   4
      Top             =   480
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Localizar"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "frmlocalizaa.frx":0A24
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW CommandButtonW1 
      Height          =   252
      Left            =   4560
      TabIndex        =   5
      Top             =   2280
      Width           =   372
      _ExtentX        =   656
      _ExtentY        =   445
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "frmlocalizaa.frx":0FBE
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Digite a Chave de Busca"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
  Dim x As Integer
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
  For x = 0 To nLEN
    List.AddItem ePASS01(x)
  Next x
  List.ListIndex = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  eRETU01 = cBUSCA
  eRETU02 = List.ListIndex + 1                 ''matriz zero posional 1
End Sub

