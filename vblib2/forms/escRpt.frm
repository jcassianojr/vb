VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form escRpt 
   Caption         =   "Selecione o Desejado"
   ClientHeight    =   7515
   ClientLeft      =   300
   ClientTop       =   450
   ClientWidth     =   11520
   HelpContextID   =   6
   Icon            =   "escRpt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   11520
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   7365
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Lista de Registros"
      Top             =   90
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   12991
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   7515
      Left            =   10050
      TabIndex        =   1
      Top             =   0
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   13256
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "escRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const nFORMID = 1065
Const cFORMID = "Relatorios / Documentos"
Private Sub Apagar()
Dim cARQ, cSQL, cRPT As String
    Grid.Col = 1
    cRPT = Grid
    cARQ = PegPath("PATH", "SYSRPT")
    cSQL = "select * from RPT WHERE CODIGO='" & ZRPTCOD & "' AND GRP='" & zRPTGRP & "' AND RPT='" & cRPT & "'"
    lRETU = ApagaSQLP(cARQ, cSQL, "Apagar")
    If lRETU Then
        FILGRID
    End If
End Sub
Private Sub Editar()
    Grid.Col = 1
    zRPT = Grid
    ePASS01 = "select * from RPT WHERE CODIGO='" & ZRPTCOD & "' AND GRP='" & zRPTGRP & "' AND RPT='" & zRPT & "'"
    ePASS02 = PegPath("PATH", "SYSRPT")
    FrmRpt.Show vbModal
    FILGRID
    LocalizaGrid Grid, zRPT, 2, False
End Sub
Private Sub Liberar()
    Grid.Col = 1
    zRPT = Grid
    escrptusr.Show vbModal
End Sub
Private Sub Escolher()
Dim cARQ, cSQL, cARQNOME, cFECDIZ As String
Dim cEXTENSAO
Dim cDIZBX, cCAMFEC, cANO, cMES As String
Dim nPOS, X As Integer
Dim zrelano, zrelmes, zrelemp As Integer
Dim aRETU As Variant
Dim lLIBGRP As Variant
Dim eRUN As String

On Error GoTo erro

    zrelano = 0
    zrelmes = 0
    zrelemp = 0

    
    Grid.Col = 1
    zRPT = Grid
    
    cARQ = PegPath("PATH", "SYSRPT")
    cSQL = "select CAMINHO,LIBERAR from RPTGRP WHERE CODIGO='" & ZRPTCOD & "' AND GRP='" & zRPTGRP & "'"
    aRETU = PegSQL(cARQ, cSQL, 2, Array("CAMINHO", "LIBERAR"), Array("C", "BF"), Array("", False))
    zRPTGRPCAM = aRETU(0)
    lLIBGRP = aRETU(1)

    
   'Temporariarmente todos os Direitos
   ' 0-Novo,1-Abrir,2-Salvar,3-Salvacomo
   ' 4-Editar,5-Visualizar,6-Imprimir,7-Exportar,8-Abrir com
    cARQ = PegPath("PATH", "SYSUSER")
    cSQL = "SELECT * FROM USURPT" & _
           " WHERE CODIGO='" & ZRPTCOD & "' AND GRP='" & zRPTGRP & "' AND RPT='" & zRPT & "' AND IDUSUARIO=" & zUSERID
    aDIREITOS = PegSQL(cARQ, cSQL, 9, _
            Array("NOVO", "ABRIR", "SALVARTF", "SALVATXT", "EDITAR", "VISUALIZA", "IMPRIME", "EXPORTA", "ABRIRCOM"), _
            Array("B", "B", "B", "B", "B", "B", "B", "B", "CU"), _
            Array(False, False, False, False, False, False, False, False, ""))
            
            
    If (Not lRETU) And (Not lLIBGRP) Then
       Alert ("Nao Liberado")
       Exit Sub
    End If
    
    
    cARQ = PegPath("PATH", "SYSRPT")
    cSQL = "select CAMINHO from RPTMAIN  WHERE CODIGO='" & ZRPTCOD & "'"
    aRETU = PegSQL(cARQ, cSQL, 1, Array("CAMINHO"), Array("C"), Array(""))
    zRPTCAM = aRETU(0)
    
        
    
    
    cSQL = "select * from RPT WHERE CODIGO='" & ZRPTCOD & "' AND GRP='" & zRPTGRP & "' AND RPT='" & zRPT & "'"
    aRELCFG = PegSQL(cARQ, cSQL, 20, _
              Array("ARQUIVO", "ABRIRCOM", "COBSMSG", "UTILIZADO", "ARVORE", _
              "BUSCA", "TITULO", "CAMINHO", "CAMINH2", "CAMINH3", "CAMINH4", _
              "LFILTRO", "TABALIAS", "TABNAME", "PREFILTRO", "SQLUSO", _
              "CARQUSO", "CARQBAI", "CARQFEC", "CARQACU"), _
              Array("CU", "CU", "C", "N", "B", _
                    "B", "C", "CU", "CU", "CU", "CU", _
                    "B", "CU", "CU", "CU", "CU", _
                    "CU", "CU", "Cu", "CU"), _
              Array("", "", "", 0, False, _
                     False, "", "", "", "", "", _
                     False, "", "", "", "", _
                     "", "", "", ""))
    ''aRELCFG
    ''0 Arquivo
    ''1 Abrir como
    ''2 Mensagem Antes Imprimir
    ''3 Nº Utilizacoes
    ''4 Arvore
    ''5 Busca
    ''6 Titulo do Relatorio
    ''7 Caminho
    ''8 Caminh2
    ''9 Caminh3
    ''10 Caminh4
    ''11 pergunta filtro
    ''12 alias da tabela
    ''13 nome da tabela
    ''14 prefiltro
    ''15 SQL
    ''16 carquso
    ''17 carqbai
    ''18 carqfec
    ''19 carqapu
                  
    If Len(aDIREITOS(8)) > 0 Then ''Abrir com do Usuario
       aRELCFG(1) = aDIREITOS(8)
    End If
              
    ''Troca se For Indicado
    aRELCFG(0) = Caminex(aRELCFG(0))
    
    ''Aarquivos
    aARQUIVOS = Array(aRELCFG(7), aRELCFG(8), aRELCFG(9), aRELCFG(10))
    
    For X = 0 To 3
        If Len(aARQUIVOS(X)) > 0 Then
           aARQUIVOS(X) = Caminex(aARQUIVOS(X), zrelano, zrelmes, zrelemp)
           zrelano = Val(eRETU01)
           zrelmes = Val(eRETU02)
           zrelemp = Val(eRETU03)
        End If
    Next X
    
    
    
    ''Tenta Caminho Informado
    cARQRTF = aRELCFG(0)
    
    If Not FileExist(cARQRTF) Then
       ''Tenta Caminho do grupo
       cARQRTF = Caminex(zRPTGRPCAM) & aRELCFG(0)
       If Not FileExist(cARQRTF) Then
          ''Tenta Caminho do Codigo
           cARQRTF = Caminex(zRPTCAM) & aRELCFG(0)
           If Not FileExist(cARQRTF) Then
              ''Tenta Caminho do Aplicativo
              cARQRTF = App.Path & aRELCFG(0)
              If Not FileExist(cARQRTF) Then
                 If aRELCFG(1) <> "SITE" And aRELCFG(1) <> "FORM" Then
                    Alert ("Arquivo Não Encontrado " & aRELCFG(0))
                    Exit Sub
                 Else
                    cARQRTF = aRELCFG(0)
                    aRELCFG(1) = "SHELL"
                 End If
              End If
           End If
       End If
    End If
    
        'Arquivos Fechamento
    aARQFEC = Array(aRELCFG(16), aRELCFG(17), aRELCFG(18), aRELCFG(19), "", "", "", "", aRELCFG(6))
            
    Mana5Fec
    
    cEXTENSAO = "XXX"
    nPOS = InStrRev(cARQRTF, ".")
    If nPOS > 0 Then
       cEXTENSAO = Mid(cARQRTF, nPOS + 1)
    End If
    
    If aRELCFG(1) = "" Or aRELCFG(1) = "_" Then    'Abrir Com Não Especificado
        Select Case cEXTENSAO
            Case "M5P" ''Mana5 Padrao /Mana5 Padrao
                 aRELCFG(1) = "IMPRELM5P" ''PADREL
            Case "M5M"  ''Mana5 Padrao /Mana5 Manrel
                 aRELCFG(1) = "IMPRELM5M" ''MANREL
            Case "REP"
                aRELCFG(1) = "FRMREP"
             Case "CPL"
                aRELCFG(1) = "EXECUTARDLL"
            Case "TXL"
                 aRELCFG(1) = "FRMTXL"
            Case "PDF", "HTM", "HTML", "CHM", "HLP"
                aRELCFG(1) = "SHELL"
            Case "DOC", "XLS", "PPS", "PPT"
                aRELCFG(1) = "SHELL"
            Case "EXE"
                aRELCFG(1) = "SHELL"
            Case "INI"
                aRELCFG(1) = "FRMINIEDITOR"
            Case "TXT", "MAN"
                 ePASS01 = Array("Editor Interno", "Imprimir Direto Impressora")
                 escOrdem.Show vbModal, Me
                 eRETU01 = FixInt(eRETU01, 0)
                 Select Case eRETU01
                        Case 0
                            aRELCFG(1) = "FRMRTF"
                        Case 1
                            aRELCFG(1) = "FRMTXL"
                 End Select
            Case "RTF"
                aRELCFG(1) = "FRMRTF"
            Case "QRP"
                aRELCFG(1) = "qrpview.EXE"
            Case "RPT"
                aRELCFG(1) = "FRMCRW"
                   ''aRELCFG(1) = "FRMCRWCRX" ''Versao 10,11
            Case "RET"
                 aRELCFG(1) = "CARETRUN.EXE"
            Case "DWF"
                aRELCFG(1) = "FRMDWF"
            Case "DWG", "DXF", "RML", "IPT", "IAM", "IDW"
                aRELCFG(1) = "FRMDWG"
            Case "WMF", "EMF", "BMP", "ICO", "JPG", "JPEG", "JNG", "KOALA", "LBM", "IFF", "MNG", "PBM", "PBMRAW", _
                 "PCD", "PCX", "PGM", "PGMRAW", "PNG", "PPM", "PPMRAW", "RAS", "TARGA", _
                 "DIB", "TGA", "PIC", "TIF", "TIFF", "WBMP", "PSD", "CUT", "XBM", "XPM", "DDS", "GIF", "HDR"
                 aRELCFG(1) = "FRMIMG"
            'Case "BMP", "JPG", "JPEG", "ICO", "WMF", "EMF", "GIF"
            ''    aRELCFG(1) = "FRMIMG"
            ''Case "DIB", "PIC", "PCX", "TGA", "PNG", "TIFF", "TIF"
            ''    aRELCFG(1) = "FABVIEW.EXE"
            Case "MPEG", "MPG", "M1V", "MP2", "MPE", "AVI", "WMF", "WAV", "SND", "AU" _
                    , "AIF", "AIFC", "AIFF", "WMA", "WMA", "MP3", "MID", "RMI" _
                    , "MIDI", "CDA", "WMA", "ASF", "WM", "WMV", "ASX" _
                    , "WAX", "M3U", "MVX", "WMX"
                aRELCFG(1) = "FRMVID"
            Case Else
                Alert ("Abrir Com Não Configurado")
                Exit Sub
        End Select
    End If
  
    
    ''Grava Utilizaçao
    aRELCFG(3) = aRELCFG(3) + 1
    GrvSQL cARQ, cSQL, 2, Array("DATAIMP", "UTILIZADO"), Array(Now, aRELCFG(3)), Array("", "N")
    
    ''Exibe Mensagem Configurada
    If Len(aRELCFG(2)) > 0 Then Alert aRELCFG(2), "Atençao"
    
    ''Ajusta Nome Arquivos
    If Len(aARQUIVOS(0)) > 0 Then cARQNOME = NomeArq(aARQUIVOS(0))
    If Len(cARQNOME) = 0 Then cARQNOME = NomeArq(cARQRTF)
    
   'Ajusta Titulo
   If InStr(aRELCFG(6), "[ZFEC]") > 0 Then
        aRELCFG(6) = Replace(aRELCFG(6), "[ZFEC]", cFECDIZ)
    End If
    If InStr(aRELCFG(6), "[ZARQ]") > 0 Then
        aRELCFG(6) = Replace(aRELCFG(6), "[ZARQ]", cARQNOME)
    End If
    If InStr(aRELCFG(6), "[ZRPT]") > 0 Then
        aRELCFG(6) = Replace(aRELCFG(6), "[ZRPT]", NomeArq(cARQRTF))
    End If
    aRELCFG(6) = Caminex(aRELCFG(6), zrelano, zrelmes, zrelemp)
    If InStr(aRELCFG(6), "[ZCOMP]") > 0 Then
        aRELCFG(6) = Replace(aRELCFG(6), "[ZCOMP]", CStr(zrelmes) & "/" & CStr(zrelano))
    End If
    If InStr(aRELCFG(6), "[ZEMP]") > 0 Then
        aRELCFG(6) = Replace(aRELCFG(6), "[ZEMP]", CStr(zrelemp))
    End If
    
    
    
    '' Executavel
    If InStr(aRELCFG(1), ".EXE") > 0 Then
       ''Direitos
        If aRELCFG(1) = "VORETRUN.EXE" Then
           If aRELCFG(11) Then
              FrmFiltro.Show vbModal, Me
              If lRETU Then
                 aRELCFG(15) = TrocaSqlWhere(aRELCFG(15), Replace(Replace(eRETU01, "{", ""), "}", ""))
              End If
           End If
           cARQRTF = cARQRTF & MONTARSN(, "") & "#" & aARQUIVOS(0) & "#" & aRELCFG(15)
        End If
        ''mkrview parametros
        If aRELCFG(1) = "MKRVIEW.EXE" Then
           cARQRTF = Chr(34) & aARQUIVOS(0) & Chr(34) & " " & Chr(34) & "SYSDBA" & Chr(34) & " " & Chr(34) & "masterkey" & Chr(34) & " " & Chr(34) & " " & Chr(34) & " " & Chr(34) & "none" & Chr(34) & " " & Chr(34) & cARQRTF & Chr(34)
        End If
        If aRELCFG(1) = "VORTF.EXE" Then cARQRTF = cARQRTF & MONTARSN()
        aRELCFG(1) = Trim(Replace(aRELCFG(1), ".EXE", " "))
''       Shell aRELCFG(1) & cARQRTF, vbNormalFocus
        Select Case aRELCFG(1)
            Case "CARETRUN" ''16 Bits nao funciona shellex
                Shell "caretrun " & cARQRTF, vbNormalFocus
            Case "VORETRUN"
                Shell "VORETRUN " & cARQRTF, vbNormalFocus
            Case "MKRVIEW"
                ShellEx aRELCFG(1), essSW_SHOWDEFAULT, cARQRTF, , , Me.Hwnd
            Case Else
                ShellEx aRELCFG(1), essSW_SHOWDEFAULT, cARQRTF, , , Me.Hwnd
        End Select
    Else
       ''Imprime conforme Abrir Com
       Select Case aRELCFG(1)
             '' Case "M5P", "M5M" ''Mana5 Padrao /Mana5 mANREL
              ''cTIPO = "P"   ARQREL="PADREL"' ELSE ARQREL="MANREL"
              Case "EXECUTARDLL"
                   ExecutarDLL cARQRTF
              Case "IMPRELM5P"
                   ''eGRUPO , eCODIGO, zUSER, cTIPO, eCAMINHO
                   eRUN = PegPath("PATH", "ZMANA5IMP") & "IMPREL " & zRPTGRP & " " & zRPT & " " & zUSER & " P " & PegPath("PATH", "ZMANA5IMP")
                   Shell eRUN, vbNormalFocus
              Case "IMPRELM5M"
                   ''eGRUPO , eCODIGO, zUSER, cTIPO, eCAMINHO
                   eRUN = PegPath("PATH", "ZMANA5IMP") & "IMPREL " & zRPTGRP & " " & zRPT & " " & zUSER & " M " & PegPath("PATH", "ZMANA5IMP")
                   Shell eRUN, vbNormalFocus
              Case "SHELL"
                   ShellEx cARQRTF, essSW_SHOWDEFAULT, , , , Me.Hwnd
              Case "FRMRTF"
                   FrmRTf.Show vbModal, Me
              Case "FRMINIEDITOR"
                   frmIniEditor.Show vbModal, Me
              Case "FRMSENDMAIL"
                   ePASS01 = Array("", "", "", "", "", cARQRTF, "", False)
                   FrmSendMail.Show vbModal, Me
              Case "FRMREP"
                   FrmRep.Show vbModal, Me
              Case "FRMTXL"
                   FrmTxl.Show vbModal, Me
              Case "FRMCRW"
                   FrmCrw.Show vbModal, Me
              Case "FRMCRWENG"
                   FrmCrwEng.Show vbModal, Me
              Case "FRMCRWCRX"
                  '' FrmCrwCrx.Show vbModal, Me                   '
              Case "FRMIMG"
                   Frmimg.Show vbModal, Me
              Case "FRMDWF"
                   ''FrmDWF.Show vbModal, Me
              Case "FRMDWG"
                   ''FrmDWG.Show vbModal, Me
              Case "FRMVID"
                   FrmVid.Show vbModal, Me
              Case Else ''Encerra Para Evitar Nova Abertura Sistema Novo
                 rptimpsys aRELCFG(1)
       End Select
    End If
    Exit Sub
erro:
Select Case Err.Number
     Case Else
         SayErro "EscRPT-Imprimir"
End Select
End Sub

Private Sub FILGRID()
Dim cARQ As String
Dim cSQL As String
    cARQ = PegPath("PATH", "SYSRPT")
    cSQL = "select * from RPT WHERE CODIGO='" & ZRPTCOD & "' AND GRP='" & zRPTGRP & "' ORDER BY SUBGRP,COGNOME,RPT"
    MontaGrid Grid, 5, Array(800, 1000, 1400, 6000, 800), Array("Sub-Grupo", "Nome", "Cognome", "Descrição", "Ext"), _
            Array("L$SUBGRP", "L$RPT", "L$COGNOME", "L$NOME", "EXT,ARQUIVO,ARQUIVO"), cARQ, cSQL
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    ''If KeyAscii > 31 And KeyAscii < 123 Then
    ''    LocalizaGrid Grid, Chr(KeyAscii), 1, False
    ''End If
End Sub

Private Sub Form_Load()
    If AcessaForm(nFORMID) Then
        MontaToolbar Me.Toolbar1, nFORMID
        FILGRID
        If FixStr(eLOCALIZA) <> "" Then LocalizaGrid Grid, eLOCALIZA, 1
    Else
        MontaToolbar Me.Toolbar1, 999
    End If
    ''Configura Help
    lRETU = False
    HelpContextID = nFORMID
    Me.Caption = cFORMID
End Sub

Private Sub Grid_DblClick()

''Editar
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

Private Sub incluir()
Dim cARQ, cRPT As String
    cRPT = LCase(Busca("Codigo", "Incluindo " & cFORMID, "", 8))
    cRPT = Trim(UCase(cRPT))
    ePASS01 = "select * from RPT WHERE CODIGO='" & ZRPTCOD & "' AND GRP='" & zRPTGRP & "' AND RPT='" & cRPT & "'"
    cARQ = PegPath("PATH", "SYSRPT")
    IncluiSQL cARQ, ePASS01, 3, Array("CODIGO", "GRP", "RPT"), Array(ZRPTCOD, zRPTGRP, cRPT), True, True
    ePASS02 = PegPath("PATH", "SYSRPT")
    If lRETU Then
       FrmRpt.Show vbModal, Me
       FILGRID
    End If
    LocalizaGrid Grid, cRPT, 2, False
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
        Case "EDI"
            Editar
        Case "LIB"
            Liberar
        Case "ESC", "IMP"
            Escolher
        Case "EXC"
            Apagar
        Case "LOC"
            LocalizaGrid Grid
    End Select
End Sub
Function MONTARSN(Optional ByVal cDELI As String = "$", Optional ByVal cSEP As String = ",") As String
Dim X As Integer
    MONTARSN = cDELI
    For X = 0 To 7
        MONTARSN = MONTARSN & IIf(aDIREITOS(X), "S", "N") & cSEP
    Next X
    MONTARSN = MONTARSN & cDELI
End Function

