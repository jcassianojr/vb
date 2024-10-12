VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Object = "{2DA70529-3366-414A-B408-46083BCD481B}#1.8#0"; "VBFLXGRD18.OCX"
Begin VB.Form escmanreg 
   Caption         =   "Escolha o grupo de Relatorio"
   ClientHeight    =   5775
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8880
   Icon            =   "escmanreg.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VBFLXGRD18.VBFlexGrid Grid 
      Height          =   5175
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   9128
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   5775
      Left            =   7410
      TabIndex        =   0
      Top             =   0
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   10186
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin vbExtra.FlexFn FlexFn1 
      Height          =   405
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   714
      PageNumbersFormat=   ""
      PageNumbersFormatIndex=   0
      BeginProperty PageNumbersFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "escmanreg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdSair_Click()

  Screen.MousePointer = vbDefault
  Unload Me

End Sub

Private Sub Escolher_Click()

  Grid.Col = 0
  zgrp = Grid
  escMANREL.Show vbModal, Me
  FilRelat

End Sub

Private Sub FilRelat()

  Dim cARQ As String
  Dim cSQL As String

  cARQ = ZMANA5CON
  cSQL = "SELECT GRUPO,DESCRICAO FROM MANREG ORDER BY grupo"
  cARQ = GeraConn(cARQ, "JETFOX")
  MontaGridUltra Grid, 2, Array(400, 5000), Array("GRP", "Nome"), _
                 Array("L$Grupo", "Descricao"), cARQ, cSQL


End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  xmontatoolbar Me.Toolbar1, "escMANREG", True
  FilRelat
  If FixStr(eLOCALIZA) <> "" Then LocalizaGri1 Grid, eLOCALIZA, 1
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Escolher_Click
  End If
  If KeyAscii > 31 And KeyAscii < 123 Then
    LocalizaGrid Grid, Chr(KeyAscii), 1, False
  End If


End Sub

'Private Sub Grid_SelChange()
'    With Grid
'        If .Rows > 2 Then
'           .Col = .Cols - 1
'          .ColSel = 0
'         .TopRow = .Row
'    End If
'End With

'End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

  Dim sButton As String
  sButton = Button
  sButton = Left(UCase(Replace(sButton, "&", "")), 3)
  If Not AcessaBtnOld("escMANREG", Button.Index) Then
    Exit Sub
  End If
  ''TbrLog Me, Button

  Select Case sButton

  Case "ESC"
    Escolher_Click

  Case "LOC"
    LocalizaGri1 Grid

  Case "SAI"
    CmdSair_Click

  End Select

End Sub


