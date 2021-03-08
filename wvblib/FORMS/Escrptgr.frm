VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form escRPTGRP 
   Caption         =   "Escolha o grupo de Relatorio"
   ClientHeight    =   5775
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin vbExtra.FlexFn FlexFn1 
      Height          =   405
      Left            =   120
      TabIndex        =   2
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
      Height          =   5775
      Left            =   7410
      TabIndex        =   1
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
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5235
      Left            =   60
      TabIndex        =   0
      Top             =   480
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   9234
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
    Dim cGRUPO, sSQL As String
    Grid.Col = 0
    cGRUPO = FixStr(Grid)
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
End Sub

Private Sub CmdSair_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Edit_Click()
    Grid.Col = 0
    zgrp = Grid
    frmRPTGRP.Show vbModal
    FilRelat
End Sub

Private Sub Escolher_Click()
    Grid.Col = 0
    zgrp = Grid
    escRPT.Show vbModal, Me
End Sub

Private Sub FilRelat()
    Dim cSQL As String
    cSQL = "SELECT GRP,NOME FROM rptgrp ORDER BY grp"
    MontaGridUltra Grid, 2, Array(800, 4000), Array("GRP", "Nome"), _
        Array("grp", "nome"), zRPTARQ, cSQL
End Sub

Private Sub Form_Load()
    Center Me
    zgrp = ""
    ZGRPSUB = ""
    eLOCALIZA = ""
    xmontatoolbar Me.Toolbar1, "escRPTGRP", True
    FilRelat
    If FixStr(eLOCALIZA) <> "" Then LocalizaGri1 Grid, eLOCALIZA, 1
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
    If KeyAscii > 31 And KeyAscii < 123 Then
        LocalizaGrid Grid, Chr(KeyAscii), 1, False
    End If

End Sub

Private Sub Grid_SelChange()
    With Grid
        If .Rows > 2 Then
            .Col = .Cols - 1
            .ColSel = 0
            .TopRow = .Row
        End If
    End With
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
    Dim sButton          As String
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


