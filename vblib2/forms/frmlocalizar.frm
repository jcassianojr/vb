VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "VBCCR18.OCX"
Begin VB.Form frmLocalizar 
   Caption         =   "Localizar"
   ClientHeight    =   2028
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   6348
   Icon            =   "frmlocalizar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2028
   ScaleWidth      =   6348
   StartUpPosition =   3  'Windows Default
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
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Preencha com A Chave de Busca"
      Top             =   480
      Width           =   4395
   End
   Begin VBCCR18.SpinBox nIND 
      Height          =   615
      Left            =   2160
      TabIndex        =   3
      Top             =   1200
      Width           =   735
      _ExtentX        =   1291
      _ExtentY        =   1080
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.8
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   656
      Picture         =   "frmlocalizar.frx":058A
      Caption         =   "Localizar"
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
   Begin VBCCR18.CommandButtonW Command2 
      Height          =   612
      Left            =   4680
      TabIndex        =   5
      Top             =   1080
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
      Picture         =   "frmlocalizar.frx":0B24
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VB.Label Label2 
      Caption         =   "Digite a Coluna de Busca"
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
         Size            =   7.8
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
Option Explicit
Private Sub cBUSCA_GotFocus()
  FocusMe
End Sub
Private Sub cBUSCA_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
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
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  eRETU01 = cBUSCA
  eRETU02 = nIND
End Sub
Private Sub nIND_GotFocus()
  FocusMe
End Sub
Private Sub nIND_KeyPress(KeyAscii As Integer)
  KeyAscii = ValiText(KeyAscii, "#NI")
End Sub
Private Sub nIND_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

