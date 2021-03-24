VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form escPPAP 
   Caption         =   "Selecione o Processo  desejado:"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   9763
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6135
      Left            =   7065
      TabIndex        =   1
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
Attribute VB_Name = "escPPAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cARQPF As String
Dim aORDEM As Variant
Dim aORDES As Variant
Dim cORDEM As String
Dim cSUBWHERE As String

Private Sub CmdEditar()
    Grid.Col = 0
    nPPAP = Grid
    If iPPAP = 3 Then
        frmGp11.Show vbModal
    Else
        frmPPAP.Show vbModal
    End If
    FilRelat
End Sub

Private Sub CmdExcluir()
    'Dim eRETU As Variant
    'Dim cdb As String
    Dim sSQL As String
    Grid.Col = 0
    nPPAP = FixInt(Grid)
    If iPPAP = 11 Then Exit Sub                  ''Cronograma nao apaga
    Select Case iPPAP
    Case 1, 11
        sSQL = "select * from PPAP WHERE PPAP=" & nPPAP
    Case 2
        sSQL = "select * from PPAF WHERE PPAP=" & nPPAP
    Case 3
        sSQL = "select * from PPAG WHERE PPAP=" & nPPAP
    End Select
    If ApagaSQLP(cARQPF, sSQL) Then
        FilRelat
    End If
End Sub

Private Sub CmdNovo()
    Dim cSQL As String
    If iPPAP = 11 Then Exit Sub                  ''Cronograma nao inclui
    Select Case iPPAP
    Case 1, 11
        cSQL = "PPAP"
    Case 2
        cSQL = "PPAF"
    Case 3
        cSQL = "PPAG"
    End Select
    nPPAP = PegUltSQL(cARQPF, "select numero from " & cSQL & " WHERE NUMERO=" & nPPAP, "NUMERO", 0)
    If IncluiSQL(cARQPF, "select numero from " & cSQL & " WHERE NUMERO=" & nPPAP, 1, Array("NUMERO"), Array(nPPAP), True, False) Then
        FilRelat
    End If
End Sub

Private Sub cmdSair()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub FilRelat()
    Dim cSQL As String
    If Len(cSUBWHERE) = 0 Then
        Select Case iPPAP
        Case 1, 11
            cSQL = "SELECT PPAP,Codigo,NOME FROM PPAP ORDER BY " & cORDEM
        Case 2
            cSQL = "SELECT PPAP,Codigo,NOME FROM PPAF ORDER BY " & cORDEM
        Case 3
            cSQL = "SELECT PPAP,Codigo,NOME,SSMT FROM PPAG ORDER BY " & cORDEM
        End Select
    Else
        Select Case iPPAP
        Case 1, 11
            cSQL = "SELECT PPAP,Codigo,NOME FROM PPAP WHERE " & cSUBWHERE & " ORDER BY " & cORDEM
        Case 2
            cSQL = "SELECT PPAP,Codigo,NOME FROM PPAF WHERE " & cSUBWHERE & " ORDER BY " & cORDEM
        Case 3
            cSQL = "SELECT PPAP,Codigo,NOME,SSMT FROM PPAG WHERE " & cSUBWHERE & " ORDER BY " & cORDEM
        End Select
    End If
    If iPPAP = 3 Then
        MontaGridFast Grid, 4, Array(400, 1600, 4000, 2000), Array("N167", "Produto", "Descrição", "SSMT"), _
        Array("PPAP", "L$CODIGO", "NOME", "SSMT"), cARQPF, cSQL
  
    Else
        MontaGridFast Grid, 3, Array(400, 1600, 4000), Array("NÝ", "Produto", "Descrição"), _
        Array("PPAP", "L$CODIGO", "NOME"), cARQPF, cSQL
    End If
End Sub

Private Sub Form_Load()
    Center Me
    If iPPAP = 3 Then
        aORDES = Array("Nº", "Produto", "Descrição", "SSMT")
        aORDEM = Array("PPAP", "CODIGO", "NOME", "SSMT")
    Else
        aORDES = Array("Nº", "Produto", "Descrição")
        aORDEM = Array("PPAP", "CODIGO", "NOME")
    End If
    cORDEM = "PPAP"
    cSUBWHERE = ""
    cARQPF = PegPath("PATH", "PPAP")
    xmontatoolbar Me.Toolbar1, "escPPAP", True
    FilRelat
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)

    If KeyAscii > 31 And KeyAscii < 123 Then
        LocalizaGrid Grid, Chr(KeyAscii), 1, False
    End If


End Sub

'Private Sub Grid_SelChange()
'    With Grid
 '       If .Rows > 2 Then
  '          .Col = .Cols - 1
   '         .ColSel = 0
    '        .TopRow = .Row
'        End If
    'End With
'End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim sButton As String
    sButton = Button
    sButton = Left(UCase(Replace(sButton, "&", "")), 3)
    If Not AcessaBtnOld("escPPAP", Button.Index) Then
        Exit Sub
    End If
    
    GravaLog 0, Button.Index, sButton, "escPPAP"
    
    Select Case sButton
    Case "ORD"
        ePASS01 = aORDES
        escOrdem.Show vbModal, Me
        If lRETU Then
            cORDEM = aORDEM(eRETU01)
            FilRelat
        End If
    Case "FIL"
        cSUBWHERE = ""
        If MDG("Usar Filtro Avancado") Then
            aARQUIVOS = Array(cARQPF)
            ''Posicao 12 Nome da Tabela
            ''Posicao 13 Nome da Tabela
            Select Case iPPAP
            Case 1, 11
                aRELCFG = Array("", "", "", 0, False, _
                                False, "", "", "", "", "", _
                                False, "PPAP", "PPAP", "", "")

            Case 2
                aRELCFG = Array("", "", "", 0, False, _
                                False, "", "", "", "", "", _
                                False, "PPAF", "PPAF", "", "")
            Case 3
                aRELCFG = Array("", "", "", 0, False, _
                                False, "", "", "", "", "", _
                                False, "PPAG", "PPAG", "", "")
            End Select
            FrmFiltro.Show vbModal, Me
            If lRETU Then
                cSUBWHERE = Replace(Replace(eRETU01, "{", ""), "}", "")
            End If
        Else
            ePASS01 = aORDES
            frmLocalizaa.Show vbModal, Me
            If lRETU Then
                If iPPAP = 3 Then
                    cSUBWHERE = MontaFiltro(aORDEM, Array("=", "L%", "L%", "L%"), eRETU01, eRETU02)
                Else
                    cSUBWHERE = MontaFiltro(aORDEM, Array("=", "L%", "L%"), eRETU01, eRETU02)
                End If
            End If
        End If
        FilRelat


    Case "IMP"
        cTIPO = "R"
        zgrp = "PPAP"
        escRPT.Show vbModal, Me

    Case "NOV"
        CmdNovo

    Case "EDI"
        CmdEditar

    Case "EXC"
        CmdExcluir

    Case "LOC"
        ePASS01 = aORDES
        frmLocalizaa.Show vbModal, Me
        If lRETU Then
            LocalizaGrid Grid, eRETU01, eRETU02, , 1
        End If
    Case "SAI"
        cmdSair

    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault

End Sub


