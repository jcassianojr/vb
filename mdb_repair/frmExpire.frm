VERSION 5.00
Begin VB.Form frmExpire 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MDB Repair Tool has expired. Order Now!!!"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4260
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2520
      Picture         =   "frmExpire.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   495
      TabIndex        =   3
      ToolTipText     =   " Buy server version at $499.95 per server "
      Top             =   1560
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   1920
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1320
      Picture         =   "frmExpire.frx":01EE
      ScaleHeight     =   255
      ScaleWidth      =   495
      TabIndex        =   0
      ToolTipText     =   " Buy local version at $49.95 per machine "
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmExpire.frx":03DC
      ForeColor       =   &H000000C0&
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmExpire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Picture1_Click()
Call Buy(Me, 1)
End Sub

Private Sub Picture2_Click()
Call Buy(Me, 2)
End Sub
