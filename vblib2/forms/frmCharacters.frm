VERSION 5.00
Begin VB.Form frmCharacters 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Insert Character"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   7935
   Icon            =   "frmCharacters.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstList 
      Appearance      =   0  'Flat
      Columns         =   7
      BeginProperty Font 
         Name            =   "isoqsymbol"
         Size            =   36
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7545
   End
   Begin VB.Label Lblchr 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "isoqsymbol"
         Size            =   14.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1200
      TabIndex        =   2
      ToolTipText     =   "Character Description"
      Top             =   4560
      Width           =   600
   End
   Begin VB.Label lblChar 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Character Description"
      Top             =   4560
      Width           =   840
   End
End
Attribute VB_Name = "frmCharacters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - '
' Hyperwrite from NIXON                                  '
' Copyright (C) 2004-2008 NIXON Software Corporation.    '
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - '
' You may use this code freely in your own applications. '
' If you are distributing your code/application(s), it   '
' would be greatly appreciated if you credit NIXON in    '
' your About dialog. Please note that portions of this   '
' code belongs to other people. For more details, please '
' view the About dialog.                                 '
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - '
Option Explicit
''Public strOtherField As String
Public lCLOSE As Boolean

Private Sub Form_Load()
    Center Me
    On Error GoTo 5
    lRETU = False
    eRETU01 = ""
    eRETU02 = 0
    If Len(ePASS01) > 0 Then
        lstList.FontName = ePASS01
        Lblchr.FontName = ePASS01
    End If
    Lblchr.FontSize = 36                         '20
    lCLOSE = ePASS02
    Dim i As Integer
    For i = 1 To 255
        lstList.AddItem Chr$(i)
    Next
    lstList.ListIndex = 0
5:
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    'If Me.Width < 3090 Then Me.Width = 3090
    'If Me.Height < 1500 Then Me.Height = 1500
    'lstList.Width = Me.ScaleWidth - lstList.Left * 2
    'lstList.Columns = lstList.Width / 650
    'lstList.Height = lblChar.Top - lstList.Top - 90
    'lblChar.Width = Me.ScaleWidth - lblChar.Left * 2
End Sub

Private Sub lstList_Click()
    On Error GoTo 5
    lblChar.Caption = lstList.ListIndex + 1      ' & ": " & GetCharName(lstList.ListIndex + 1), vbProperCase)
    Lblchr.Caption = Chr(lstList.ListIndex + 1)
5:
End Sub

Private Sub lstList_DblClick()
    eRETU01 = Chr$(lstList.ListIndex + 1)
    eRETU02 = lstList.ListIndex + 1
    lRETU = True
    Unload Me
End Sub

Private Sub lstList_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 27 Then
        Unload Me
    End If
    If KeyAscii = 13 Then
        lstList_DblClick
    Else
        lstList.ListIndex = KeyAscii - 1
    End If
End Sub

