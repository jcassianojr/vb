VERSION 5.00
Begin VB.Form frmAboutBox 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MDB Repair Tool"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   3825
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2280
      Picture         =   "frmAboutBox.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   495
      TabIndex        =   6
      ToolTipText     =   " Buy server version at $499.95 per server "
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   960
      Picture         =   "frmAboutBox.frx":01EE
      ScaleHeight     =   255
      ScaleWidth      =   495
      TabIndex        =   2
      ToolTipText     =   " Buy local version at $49.95 per machine "
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin XPControls.XPButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAboutBox.frx":03DC
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   380
      Width           =   3495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAboutBox.frx":0498
      ForeColor       =   &H000000C0&
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   40
      Width           =   3495
   End
End
Attribute VB_Name = "frmAboutBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Label1.Caption = App.LegalCopyright
Me.Caption = "  " & App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision & " (Registered)"
Me.Height = 2655
Command1.Top = 1500
End Sub

Private Sub Picture1_Click()
Call Buy(Me, 1)
End Sub

Private Sub Picture2_Click()
Call Buy(Me, 2)
End Sub
