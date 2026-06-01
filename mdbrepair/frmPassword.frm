VERSION 5.00
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "vbccr18.ocx"
Begin VB.Form frmPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MDB Repair Tool"
   ClientHeight    =   1716
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3204
   ControlBox      =   0   'False
   DrawMode        =   14  'Copy Pen
   Icon            =   "frmPassword.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1716
   ScaleWidth      =   3204
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   0
      Picture         =   "frmPassword.frx":048A
      ScaleHeight     =   1812
      ScaleWidth      =   408
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
      Width           =   2292
   End
   Begin VBCCR18.CommandButtonW Command1 
      Height          =   492
      Left            =   600
      TabIndex        =   3
      Top             =   840
      Width           =   2292
      _ExtentX        =   4043
      _ExtentY        =   868
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Caption         =   "Confirmar"
      Picture         =   "frmPassword.frx":21CE
      PictureAndCaption=   -1  'True
      Style           =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Database Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
