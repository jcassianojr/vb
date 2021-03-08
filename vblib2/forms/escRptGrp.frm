VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form escRptGrp 
   Caption         =   "Selecione o Desejado"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9105
   HelpContextID   =   6
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   5775
      Left            =   7605
      TabIndex        =   1
      ToolTipText     =   "Botoes de Açoes"
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   10186
      ButtonWidth     =   714
      ButtonHeight    =   714
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Lista de Registros"
      Top             =   90
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   9763
      _Version        =   393216
   End
End
Attribute VB_Name = "escRptGrp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const nFORMID = 1059
Const cFORMID = "Grupos Relatorios / Documentos"
Private Sub Apagar()
Dim cARQ, cSQL, cGRUPO As String
    grid.Col = 1
    cGRUPO = grid
    cARQ = PegPath("PATH", "SYSRPT")
    cSQL = "select * from RPTGRP WHERE CODIGO='" & ZRPTCOD & "' AND GRP='" & cGRUPO & "'"
    lRETU = ApagaSQLP(cARQ, cSQL, "Apagar")
    If lRETU Then
       ''Apagando Relatorio
       cSQL = "select * from RPT WHERE CODIGO='" & ZRPTCOD & "' AND GRP='" & cGRUPO & "'"
       ApagaSQL cARQ, cSQL
       Filgrid
    End If
End Sub

Private Sub Editar()

Dim cGRUPO As String

    grid.Col = 1
    cGRUPO = grid
    ePASS01 = "select * from RPTGRP WHERE CODIGO='" & ZRPTCOD & "' AND GRP='" & cGRUPO & "'"
    FrmRptGrp.Show vbModal
    Filgrid
    LocalizaGrid grid, cGRUPO, 1, False

End Sub

Private Sub Escolher()
    lRETU = True
    grid.Col = 1
    zRPTGRP = grid
    grid.Col = 3
    zRPTGRPCAM = Caminex(grid)
    escRpt.Show vbModal, Me
End Sub

Private Sub Filgrid()
Dim cARQ As String
Dim cSQL As String
    cARQ = PegPath("PATH", "SYSRPT")
    cSQL = "select * from RPTGRP WHERE CODIGO='" & ZRPTCOD & "' order by GRP"
    MontaGridFast grid, 4, Array(1000, 1000, 4000, 1), Array("CODIGO", "Grupo", "Nome", "Caminho"), _
            Array("CODIGO", "GRP", "NOME", "CAMINHO"), cARQ, cSQL
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    ''If KeyAscii > 31 And KeyAscii < 123 Then
    ''    LocalizaGrid Grid, Chr(KeyAscii), 1, False
    ''End If
End Sub

Private Sub Form_Load()
    If AcessaForm(nFORMID) Then
        MontaToolbar Me.Toolbar1, nFORMID
        Filgrid
        If FixStr(eLOCALIZA) <> "" Then LocalizaGrid grid, eLOCALIZA, 1
    Else
        MontaToolbar Me.Toolbar1, 999
    End If
    ''Configura Help
    lRETU = False
    HelpContextID = nFORMID
    Me.Caption = cFORMID
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
If KeyAscii > 31 And KeyAscii < 123 Then
        LocalizaGrid grid, Chr(KeyAscii), 1, False
    End If
End Sub

Private Sub Grid_SelChange()
    With grid
        If .Rows > 2 Then
            .Col = .Cols - 1
            .ColSel = 0
            .TopRow = .Row
        End If
    End With
End Sub

Private Sub incluir()
Dim cARQ, cGRUPO As String
    cGRUPO = LCase(Busca("Codigo", "Incluindo " & cFORMID, "", 8))
    cGRUPO = Trim(UCase(cGRUPO))
    ePASS01 = "select * from RPTGRP WHERE CODIGO='" & ZRPTCOD & "' AND GRP='" & cGRUPO & "'"
    cARQ = PegPath("PATH", "SYSRPT")
    If Len(zRPTCAM) > 0 Then
       IncluiSQL cARQ, ePASS01, 3, Array("CODIGO", "GRP", "CAMINHO"), Array(ZRPTCOD, cGRUPO, zRPTCAM), True
    Else
       IncluiSQL cARQ, ePASS01, 2, Array("CODIGO", "GRP"), Array(ZRPTCOD, cGRUPO), True
    End If
    If lRETU Then
       FrmRptGrp.Show vbModal, Me
       Filgrid
    End If
    LocalizaGrid grid, cGRUPO, 1, False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim sButton As String
Dim nINDICE As Integer
    
    sButton = Button
    sButton = Left(UCase(Replace(sButton, "&", "")), 3)
    
    If sButton = "SAI" Then
        Unload Me
    End If
    
    
    nINDICE = FixNum(Button.Index)
    If Not AcessaBotao(nFORMID, nINDICE) Then
        Exit Sub
    End If
    GravaLog nFORMID, nINDICE, sButton
    
    Select Case sButton
        Case "NOV"
            incluir
        Case "VER"
            '' FrmImpOldWrpt.Show vbModal, Me
        Case "EDI"
            Editar
        Case "ESC", "IMP"
            Escolher
        Case "EXC"
            Apagar
        Case "LOC"
            LocalizaGrid grid
    End Select
End Sub
