VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form escPOA 
   Caption         =   "Selecione o Processo  desejado:"
   ClientHeight    =   7260
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "Escpoa.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7260
   ScaleWidth      =   11400
   StartUpPosition =   2  'CenterScreen
   Begin vbExtra.FlexFn FlexFn1 
      Height          =   405
      Left            =   360
      TabIndex        =   2
      Top             =   240
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
      Height          =   7260
      Left            =   9930
      TabIndex        =   1
      Top             =   0
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   12806
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   6315
      Left            =   60
      TabIndex        =   0
      Top             =   840
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   11139
      _Version        =   393216
   End
End
Attribute VB_Name = "escPOA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aORDEM As Variant
Dim aORDES As Variant
Dim cORDEM As String
Dim cSUBWHERE As String
Dim cARQPOA As String

Private Sub ApagaNF_Click()
    Dim sSQL As String
    Grid.Col = 0
    nPPAP = Grid
    sSQL = "select * from " & zPFTIPO & " WHERE numero=" & nPPAP
    If ApagaSQLP(cARQPOA, sSQL) Then
        FilRelat
    End If
End Sub

Private Sub CmdSair_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub EditPf_Click()
    Grid.Col = 0
    nPPAP = Grid
    Grid.Col = 1
    nPF = Grid
    Grid.Col = 2
    nSEQ = Grid
    Grid.Col = 3
    nSSQ = Grid
    Grid.Col = 4
    nORD = Grid
    
    If zPFTIPO = "POA" Then
        frmPOA.Show vbModal
    End If
    If zPFTIPO = "POKA" Then
        frmPOKA.Show vbModal
    End If
    If zPFTIPO = "DISPO" Then
        frmDISPO.Show vbModal
    End If
    
    FilRelat
End Sub

Private Sub FilRelat()
    Dim cSQL As String
    If Len(cSUBWHERE) = 0 Then
        cSQL = "SELECT NUMERO,PF,SEQ,SSQ,ITEM,codigo,NOME,resnum,eladat FROM " & zPFTIPO & " ORDER BY " & cORDEM
    Else
        cSQL = "SELECT NUMERO,PF,SEQ,SSQ,ITEM,codigo,NOME,resnum,eladat FROM " & zPFTIPO & " WHERE " & cSUBWHERE & " ORDER BY " & cORDEM
    End If
    MontaGridFast Grid, 9, Array(600, 600, 400, 400, 400, 1600, 1000, 4000, 600), Array("Nº", "PF", "SEQ", "SSQ", "Item", "Produto", "Elaboracao", "Descricao", "LibNum"), _
        Array("NUMERO", "PF", "SEQ", "SSQ", "ITEM", "L$CODIGO", "ELADAT", "NOME", "RESNUM"), cARQPOA, cSQL
End Sub

Private Sub Form_Load()
    CenterFormToScreen Me
    aORDEM = Array("PF,SEQ,SSQ", "NUMERO", "CODIGO", "NOME", "ELADAT")
    aORDES = Array("PF", "Nº", "Produto", "Descricao", "Elaboracao")
    cORDEM = "PF,SEQ,SSQ"
    cSUBWHERE = ""
    cARQPOA = PegPath("PATH", zPFTIPO)
    xmontatoolbar Me.Toolbar1, "Esc" & zPFTIPO, True
    FilRelat
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
    If KeyAscii > 31 And KeyAscii < 123 Then
        LocalizaGrid Grid, Chr(KeyAscii), 1, False
    End If
End Sub

'Private Sub Grid_SelChange()
 '   With Grid
  '      If .Rows > 2 Then
   '         .Col = .Cols - 1
    '        .ColSel = 0
     '       .TopRow = .Row
      '  End If
    'End With
'End Sub

Private Sub NovoPF_Click()
    iMU01 = 3
    nPF = 0
    NewSeq.Show vbModal
    'nPPAP = PegUltSQL(cARQPOA, "select numero from " & zPFTIPO & " WHERE PF=" & FixInt(nPF) & " AND SEQ=" & FixInt(nSEQ) & " AND SSQ=" & FixInt(nSSQ) & " AND ITEM=" & FixInt(nORD), "NUMERO", 0)
    
    nPPAP = PegCampoSQLADO(cARQPOA, "select numero from " & zPFTIPO & " WHERE PF=" & FixInt(nPF) & " AND SEQ=" & FixInt(nSEQ) & " AND SSQ=" & FixInt(nSSQ) & " AND ITEM=" & FixInt(nORD), "", 0)
      
    'checagem acima necessario pos checas as sequencias
    If nPPAP = 0 Then
        nPPAP = FixInt(PegMAXSQL(cARQPOA, zPFTIPO, "NUMERO", 0)) + 1
        IncluiSQL cARQPOA, "SELECT * FROM " & zPFTIPO & " WHERE NUMERO=" & nPPAP, 5, Array("NUMERO", "PF", "SEQ", "SSQ", "ITEM"), Array(nPPAP, nPF, nSEQ, nSSQ, nORD)
    End If
    If zPFTIPO = "POA" Then
        frmPOA.Show vbModal
    End If
    If zPFTIPO = "POKA" Then
        frmPOKA.Show vbModal
    End If
    If zPFTIPO = "DISPO" Then
        frmDISPO.Show vbModal
    End If
    
    FilRelat
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim sButton As String
    sButton = Button
    sButton = Left(UCase(Replace(sButton, "&", "")), 3)
    
    If Not AcessaBtnOld("esc" & zPFTIPO, Button.Index) Then
        Exit Sub
    End If
    GravaLog 0, Button.Index, sButton, "esc" & zPFTIPO
    
    Select Case sButton
    
    Case "LIB"
        CmdLiberar_Click
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
            aARQUIVOS = Array(cARQPOA)
            ''Posicao 12 Nome da Tabela
            ''Posicao 13 Nome da Tabela
            aRELCFG = Array("", "", "", 0, False, _
                            False, "", "", "", "", "", _
                            False, zPFTIPO, zPFTIPO, "", "")
            FrmFiltro.Show vbModal, Me
            If lRETU Then
                cSUBWHERE = Replace(Replace(eRETU01, "{", ""), "}", "")
            End If
        Else
            ePASS01 = aORDES
            frmLocalizaa.Show vbModal, Me
            If lRETU Then
                cSUBWHERE = MontaFiltro(aORDEM, Array("=", "L%", "L%", "L%"), eRETU01, eRETU02)
            End If
        End If
        FilRelat
    Case "IMP"
        cTIPO = "R"
        zgrp = "SAC"
        escRPT.Show vbModal, Me
    Case "NOV"
        NovoPF_Click
    Case "EDI"
        EditPf_Click
    Case "EXC"
        ApagaNF_Click
    Case "LOC"
        ePASS01 = aORDES
        frmLocalizaa.Show vbModal, Me
        If lRETU Then
            LocalizaGrid Grid, eRETU01, eRETU02, , 1
        End If
          
    Case "SAI"
        CmdSair_Click
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdLiberar_Click()
    Dim nTMPELA
  
    Grid.Col = 0
    nPPAP = Grid
    Grid.Col = 1
    nPF = Grid
    Grid.Col = 2
    nSEQ = Grid
    Grid.Col = 3
    nSSQ = Grid

    'nTMPELA = PegUltSQL(Sdb, "SELECT ELANUM FROM " & zPFTIPO & " WHERE NUMERO=" & nPPAP, "ELANUM", 0)
    nTMPELA = PegCampoSQLADO(Sdb, "SELECT ELANUM FROM " & zPFTIPO & " WHERE NUMERO=" & nPPAP, "", 0)
    
    
    If nTMPELA = zIDFOLHA Then
        Alert ("Voce ja e o Elaborador")
        Exit Sub
    End If
    ePASS01 = zPFTIPO
    frmPFLIB.Show vbModal, Me
    FilRelat
End Sub


