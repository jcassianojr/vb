VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form frmPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MDB Repair Tool"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3195
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   3195
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   0
      Picture         =   "frmPassword.frx":048A
      ScaleHeight     =   1815
      ScaleWidth      =   405
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   405
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   600
      MaxLength       =   14
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin XPControls.XPButton Command1 
      Height          =   435
      Left            =   1200
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmPassword.frx":21CE
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Database Password"
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
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  sDBpassword = txtPassword.Text
  Unload Me
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me

End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
  If KeyAscii = 32 Then  'don't allow spaces
    KeyAscii = 0
  End If
End Sub
