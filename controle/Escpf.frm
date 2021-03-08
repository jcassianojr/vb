VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form escpf 
   Caption         =   "Selecione o Processo  desejado:"
   ClientHeight    =   7920
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   Begin vbExtra.FlexFn FlexFn1 
      Height          =   405
      Left            =   240
      TabIndex        =   2
      Top             =   120
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
      Height          =   7920
      Left            =   10170
      TabIndex        =   1
      Top             =   0
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   13970
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   7275
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   12832
      _Version        =   393216
   End
End
Attribute VB_Name = "escpf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cARQPF As String
Dim cARQFEMEA As String
Dim cARQPFP As String
Dim cARQPFG As String
Dim aORDEM As Variant
Dim aORDES As Variant
Dim cORDEM As String
Dim cSUBWHERE As String

Private Sub ApagaNF_Click()
    Dim sSQL As String
    Dim cARQUSO As String
    Dim cARQIE As String
    Dim X

    cARQIE = PegPath("PATH", "IE")


    If Not MDG("Apagar Processo Fabricação", "Confirme Exclusão") Then
        Exit Sub
    End If
    If Not MDG("Deseja Realmente Apagar Processo Fabricação", "Confirme Exclusão") Then
        Exit Sub
    End If

    
    Grid.Col = 0
    nPF = Grid
    
    For X = 1 To 3
        Select Case X
        Case 1
            cARQUSO = cARQPF
        Case 2
            cARQUSO = cARQPFP
        Case 3
            cARQUSO = cARQPFG
        End Select
        'PF
        DizerBarra "Apagando PF"
        sSQL = "select * from PF WHERE PF=" & nPF
        ApagaSQL cARQUSO, sSQL
        DizerBarra "Apagando Sequencia Operacoes"
        'PFS
        sSQL = "select * from PFS WHERE PF=" & nPF
        ApagaSQL cARQUSO, sSQL
        DizerBarra "Apagando Outras Sequencias"
        'PFSO
        sSQL = "select * from PFSO WHERE PF=" & nPF
        ApagaSQL cARQUSO, sSQL
        DizerBarra "Apagando Composicao"
        'PFCO
        sSQL = "select * from PFCO WHERE PF=" & nPF
        ApagaSQL cARQUSO, sSQL
        DizerBarra "Apagando Instrucao de trabalho"
        'PFI
        sSQL = "select * from PFI WHERE PF=" & nPF
        ApagaSQL cARQUSO, sSQL
        DizerBarra "Apagando Composicao"
        'PFC
        sSQL = "select * from PFC WHERE PF=" & nPF
        ApagaSQL cARQUSO, sSQL
        'PFMS03
        DizerBarra "Apagando composicao"
        sSQL = "select * from PFMS03 WHERE PF=" & nPF
        ApagaSQL cARQUSO, sSQL
        'PFCMS03
        DizerBarra "Apagando Plano controle composicao"
        sSQL = "select * from PFCMS03 WHERE PF=" & nPF
        ApagaSQL cARQUSO, sSQL
        DizerBarra "Apagando revisoes"
        ''Revisoes
        sSQL = "select * from REV WHERE PF=" & nPF
        ApagaSQL cARQUSO, sSQL
        DizerBarra "Apagando itens revisoes"
        sSQL = "select * from REVI WHERE PF=" & nPF
        ApagaSQL cARQUSO, sSQL
        DizerBarra "Apagando femea"
        ''fEMEA PF
        sSQL = "select * from FEMEA WHERE PF=" & nPF
        ApagaSQL cARQUSO, sSQL
        DizerBarra "Apagando femea Padrao"
        'IE
        DizerBarra "Apagando Instrucao de embalagem"
        sSQL = "select * from IE WHERE PF=" & nPF
        ApagaSQL cARQIE, sSQL
    Next X
    
    ''fEMEAS padrao
    sSQL = "select * from FEMEA WHERE PF=" & nPF
    ApagaSQL cARQFEMEA, sSQL
    DizerBarra "Apagando femea PF"
    sSQL = "select * from FEMPF WHERE PF=" & nPF
    ApagaSQL cARQFEMEA, sSQL
    DizerBarra "Apagando femea RPNO"
    sSQL = "select * from FEMRPNO WHERE PF=" & nPF
    ApagaSQL cARQFEMEA, sSQL
    DizerBarra "Apagando femea RPNT"
    sSQL = "select * from FEMRPNT WHERE PF=" & nPF
    ApagaSQL cARQFEMEA, sSQL
    DizerBarra "Apagando femea RPNG"
    sSQL = "select * from FEMRPNG WHERE PF=" & nPF
    ApagaSQL cARQFEMEA, sSQL
    DizerBarra "Apagando femea Avulso"
    sSQL = "select * from FEMAVU WHERE PF=" & nPF
    ApagaSQL cARQFEMEA, sSQL
    DizerBarra "Apagando femea revi"
    sSQL = "select * from FEMREVI WHERE PF=" & nPF
    ApagaSQL cARQFEMEA, sSQL
    DizerBarra "Apagando femea historio"
    sSQL = "select * from FEMHIS WHERE PF=" & nPF
    ApagaSQL cARQFEMEA, sSQL


     

    
    
    FilRelat

End Sub

Private Sub CmdLiberar_Click(ByVal cTIPOLIB)
    Dim cSQL, cSQL2 As String
    Dim DB As New ADODB.Connection
    Dim RSTAB As New ADODB.Recordset
    Dim RSTA2 As New ADODB.Recordset
    Dim nTMPELA


    'Checa operacao a operacao para ver se tem sequencia e instrucao por isso o loop
    Grid.Col = 0
    nPF = Grid
    cSQL = "select PF,SEQ,SSQ from PFS WHERE PF=" & nPF & " ORDER BY SEQ,SSQ"
    DB.ConnectionTimeout = 120
    DB.Open GeracArq(cARQPF, , False)
    RSTAB.Open cSQL, DB, adOpenForwardOnly, adLockReadOnly
    If RSTAB.EOF Then
        RSTAB.Close
        DB.Close
        Alert ("Nenhuma Sequencia Operaçao Cadastrada PF:" & CStr(nPF))
        Exit Sub
    End If
    RSTAB.MoveFirst
    While Not RSTAB.EOF
        nSEQ = FixNum(RSTAB("SEQ"))
        nSSQ = FixNum(RSTAB("SSQ"))
        cSQL2 = "select PF,SEQ,SSQ from PFI WHERE PF=" & nPF & " AND SEQ=" & CStr(FixNum(RSTAB("SEQ"))) & " AND SSQ=" & CStr(FixNum(RSTAB("SSQ")))
        RSTA2.Open cSQL2, DB, adOpenForwardOnly, adLockReadOnly
        If RSTA2.EOF Then
            RSTAB.Close
            RSTA2.Close
            DB.Close
            Alert ("PF " & CStr(nPF) & "Seq " & CStr(nSEQ) & "/" & CStr(nSSQ) & " Sem Instruçao")
            Exit Sub
        End If
        RSTA2.Close
        RSTAB.MoveNext
    Wend
    RSTAB.Close
    DB.Close
    
    If cTIPOLIB = "PF" Then
        Grid.Col = 4
        If IsDate(Grid) Then
            If Not MDG("Ja Liberada - Liberar Novamente", "Confirme Liberacao") Then
                Exit Sub
            End If
        End If
    End If
    
    
    Grid.Col = 0
    nPF = Grid
    
    nTMPELA = PegUltSQL(Sdb, "SELECT ELANUM FROM PF WHERE PF=" & nPF, "ELANUM", 0)
    If nTMPELA = zIDFOLHA Then
        Alert ("Voce ja e o Elaborador")
        Exit Sub
    End If
    
    
    ePASS01 = cTIPOLIB
    frmPFLIB.Show vbModal, Me
    FilRelat

End Sub

Private Sub CmdSair_Click()

    Screen.MousePointer = vbDefault
    Unload Me

End Sub

Private Sub EditPf_Click()

    Grid.Col = 0
    nPF = Grid
    frmPF.Show vbModal
    FilRelat

End Sub
Private Sub FilRelat()
    Dim cSQL As String
    If Len(cSUBWHERE) > 0 Then
        cSQL = "SELECT PF,codigo,DEscr,resdat,resnum,codigoint FROM PF WHERE " & cSUBWHERE & " ORDER BY " & cORDEM
    Else
        cSQL = "SELECT PF,codigo,DEscr,resdat,resnum,codigoint FROM PF ORDER BY " & cORDEM
    End If
    MontaGrid Grid, 6, Array(800, 1600, 4000, 500, 1000, 1800), Array("PF", "Produto", "Descricao", "Lib", "Liberada", "Codigo inT"), _
        Array("PF", "L$CODIGO", "L$DESCR", "RESNUM", "C$NPD,RESNUM,RESDAT", "L$CODIGOINT"), cARQPF, cSQL
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    'fazcodint 'Atualizar codigo logix
    'fazcodint2 'atualiza codigo logix seq
    'fazcodcli
    Center Me
    cARQPF = PegPath("PATH", "PF")
    cARQFEMEA = PegPath("PATH", "FEMEA")
    cARQPFP = PegPath("PATH", "PFP")
    cARQPFG = PegPath("PATH", "PFG")
    Select Case zPFTIPO
    Case "PFPRO"
        xmontatoolbar Me.Toolbar1, "escPF", True
    Case "PFPRE"
        xmontatoolbar Me.Toolbar1, "escPFPRE", True
    Case "PFGP12"
        xmontatoolbar Me.Toolbar1, "escPFGP12", True
    End Select
    aORDEM = Array("PF", "CODIGO", "DESCR", "RESNUM", "RESDAT", "CODIGOINT")
    aORDES = Array("PF", "Codigo", "Descricao", "Nº Liberador", "Data Liberacao", "Codigo Int")
    cORDEM = "PF"
    cSUBWHERE = ""
    FilRelat
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

Private Sub NovoPF_Click()
    nPF = FixInt(PegMAXSQL(cARQPF, "PF", "PF", 0)) + 1
    If IncluiSQL(cARQPF, "select pf,revdat from PF where pf=" & nPF, 2, Array("PF", "REVDAT"), Array(nPF, Today()), True, False) Then
        CHECKPFPG                                'inclui embalagem preliminar gp12
        FilRelat
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim sButton As String
    Dim aORDLIB As Variant
    Dim aDESLIB As Variant

    sButton = Button
    sButton = Left(UCase(Replace(sButton, "&", "")), 3)
    Select Case zPFTIPO
    Case "PFPRO"
        If Not AcessaBtnOld("escPF", Button.Index) Then
            Exit Sub
        End If
    Case "PFPRE"
        If Not AcessaBtnOld("escPFPRE", Button.Index) Then
            Exit Sub
        End If
    Case "PFGP12"
        If Not AcessaBtnOld("escPFGP12", Button.Index) Then
            Exit Sub
        End If
    End Select
 
    lRETU = False
    GravaLog 0, Button.Index, sButton, "escPF"
    
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
            aRELCFG = Array("", "", "", 0, False, _
                            False, "", "", "", "", "", _
                            False, "PF", "PF", "", "")
            FrmFiltro.Show vbModal, Me
            If lRETU Then
                cSUBWHERE = Replace(Replace(eRETU01, "{", ""), "}", "")
               
            End If
        Else
            ePASS01 = aORDES
            frmLocalizaa.Show vbModal, Me
            If lRETU Then
                cSUBWHERE = MontaFiltro(aORDEM, Array("=", "L*", "L*", "=", "L*", "L*"), eRETU01, eRETU02)
            End If
        End If
        FilRelat
    Case "IMP"
        cTIPO = "R"
        zgrp = "PF"
        ZGRPSUB = ""
        eLOCALIZA = ""
        escRPT.Show vbModal, Me
    Case "ESC"
        lRETU = True
        Grid.Col = 0
        eRETU01 = Grid
    Case "NOV"
        NovoPF_Click
    Case "EDI"
        EditPf_Click
    Case "EXC"
        ApagaNF_Click
    Case "COP"
        cARQIMP = "PF"
        FrmImp2.Show vbModal, Me
    Case "LOC"
        ePASS01 = aORDES
        frmLocalizaa.Show vbModal, Me
        If lRETU Then
            LocalizaGrid Grid, eRETU01, eRETU02, , 1
        End If
    Case "LIB"
        CmdLiberar_Click "PF"
    Case "PRO"
        CmdLiberar_Click "PFPRO"
    Case "SEG"
        CmdLiberar_Click "PFSEG"
    Case "TRO"
        Alert "Faca esta alteracao pela edicao do pf"
        'Trocapf_Click
    Case "SAI"
        CmdSair_Click
    End Select

End Sub
Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub


