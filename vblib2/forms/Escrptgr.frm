VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Object = "{075212A8-C1CF-444E-939D-F6046CCDBC08}#1.0#0"; "VBFLXGRD18.OCX"
Begin VB.Form escRPTGRP 
   Caption         =   "Escolha o grupo de Relatorio"
   ClientHeight    =   6135
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9345
   Icon            =   "Escrptgr.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   9345
   StartUpPosition =   2  'CenterScreen
   Begin VBFLXGRD18.VBFlexGrid Grid 
      Height          =   5415
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9551
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6135
      Left            =   7875
      TabIndex        =   0
      Top             =   0
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   10821
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "escRPTGRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Apaga_Click()
  Dim cGRUPO As String
  Dim sSQL As String
  If Grid.Row > 0 Then  'And Grid.Row < Grid.Rows - 1 Then
    Grid.Col = 0
    cGRUPO = FixStr(Grid.tEXT)
    '******************************************************************
    'RPTGRP
    sSQL = "select * from RPTGRP WHERE GRP='" & cGRUPO & "'"
    If ApagaSQLP(zRPTARQ, sSQL) Then
      '******************************************************************
      'RPT
      sSQL = "select * from RPT WHERE GRP='" & cGRUPO & "'"
      '******************************************************************
      ApagaSQL zRPTARQ, sSQL
      'RPTUSR
      sSQL = "select * from " & ArqRPTUsr() & " WHERE GRP='" & cGRUPO & "'"
      ApagaSQL DBWRPT, sSQL
      FilRelat
    End If
  End If
End Sub

Private Sub CmdSair_Click()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub Edit_Click()
  If Grid.Row > 0 Then  ''And Grid.Row < Grid.Rows - 1 Then
    Grid.Col = 0
    zgrp = Grid.tEXT
    frmRPTGRP.Show vbModal
    FilRelat
  End If
End Sub

Private Sub Escolher_Click()
  If Grid.Row > 0 Then  'And Grid.Row < Grid.Rows - 1 Then
    Grid.Col = 0
    zgrp = Grid.tEXT
    escRPT.Show vbModal, Me
  End If
End Sub

Private Sub FilRelat()
  Dim cSQL As String
  cSQL = "SELECT GRP,NOME FROM rptgrp ORDER BY grp"
  MontaGridUltra Grid, 2, Array(800, 4000), Array("GRP", "Nome"), _
                 Array("grp", "nome"), zRPTARQ, cSQL
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  zgrp = ""
  ZGRPSUB = ""
  eLOCALIZA = ""
  xmontatoolbar Me.Toolbar1, "escRPTGRP", True
  FilRelat
  If FixStr(eLOCALIZA) <> "" Then LocalizaGri1 Grid, eLOCALIZA, 1
End Sub

Private Sub grid_Click()
  ' Grid.Col = 0
  Grid.ColSel = Grid.cols - 1
  'For X = 0 To Grid.Cols - 1
  '  Grid.ColSel = X
  'Next X
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
  If KeyAscii > 31 And KeyAscii < 123 Then
    LocalizaGrid Grid, Chr(KeyAscii), 1, False
  End If

End Sub
Private Sub Novo_Click()
  Dim cSQL As String
  zgrp = InputBox("Digite o Codigo", "Inclusão Relatorio", "____")
  zgrp = FixStr(zgrp, "", "TRIM", 4)
  cSQL = "select * from RPTGRP WHERE GRP='" & zgrp & "'"
  If IncluiSQL(zRPTARQ, cSQL, 1, Array("GRP"), Array(zgrp), True, True) Then
    frmRPTGRP.Show vbModal
    FilRelat
  End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim sButton As String
  sButton = Button
  sButton = Left(UCase(Replace(sButton, "&", "")), 3)
  If Not AcessaBtnOld("escRPTGRP", Button.Index) Then
    Exit Sub
  End If

  GravaLog 0, Button.Index, sButton, "escRPTGRP"

  Select Case sButton

  Case "NOV"
    Novo_Click

  Case "EDI"
    Edit_Click

  Case "EXC"
    Apaga_Click

  Case "ESC"
    Escolher_Click

  Case "LOC"
    LocalizaGri1 Grid

  Case "SAI"
    CmdSair_Click

  End Select

End Sub


