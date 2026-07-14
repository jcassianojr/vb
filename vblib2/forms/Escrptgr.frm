VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Object = "{451B73A5-1563-45D5-A6AC-7B2B7D30B778}#3.0#0"; "BSPrin30.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "vbccr18.ocx"
Object = "{075212A8-C1CF-444E-939D-F6046CCDBC08}#1.6#0"; "VBFLXGRD18.OCX"
Begin VB.Form escRPTGRP 
   Caption         =   "Escolha o grupo de Relatorio"
   ClientHeight    =   6132
   ClientLeft      =   168
   ClientTop       =   456
   ClientWidth     =   9348
   Icon            =   "Escrptgr.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6132
   ScaleWidth      =   9348
   StartUpPosition =   2  'CenterScreen
   Begin BSPrinter.PrintPreview PrintPreview1 
      Left            =   2640
      Top             =   0
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin VBFLXGRD18.VBFlexGrid Grid 
      Height          =   5415
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   7575
      _ExtentX        =   13356
      _ExtentY        =   9546
   End
   Begin vbExtra.FlexFn FlexFn1 
      Height          =   396
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   699
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
      Height          =   6132
      Left            =   7872
      TabIndex        =   0
      Top             =   0
      Width           =   1476
      _ExtentX        =   2604
      _ExtentY        =   10816
      ButtonWidth     =   614
      ButtonHeight    =   953
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin VBCCR18.CommandButtonW CmdPreviewGRID 
      Height          =   372
      Left            =   2040
      TabIndex        =   3
      Top             =   0
      Width           =   612
      _ExtentX        =   1080
      _ExtentY        =   656
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "Escrptgr.frx":058A
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW CommandButtonW7 
      Height          =   252
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   372
      _ExtentX        =   656
      _ExtentY        =   445
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "Escrptgr.frx":0B24
      WordWrap        =   0   'False
      Style           =   1
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

Private Sub CmdPreviewGRID_Click()
   PrintPreview1.ShowPreview
   'PrintPreview1.PrintGrid Grid
End Sub
Public Sub PrintPreview1_PrepareReport(Cancel As Boolean)
    ' Title
   ' Printer.FontSize = 18
   ' Printer.FontBold = True
   ' Printer.Print lblTitle.Caption
   ' Printer.FontSize = 12
   ' Printer.Print lblTitle2.Caption
   ' Printer.Print
   ' Printer.FontBold = False
    
    ' Print the grid
    PrintPreview1.PrintGrid Grid ' (this is the only code line needed for printing the grid)
    
   ' Printer.Print
   ' Printer.Print "Data published by United Nations"
   ' Printer.FontSize = 10
   ' Printer.Print "https://population.un.org/wpp/Download/Standard/Population/"
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


