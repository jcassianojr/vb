VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Object = "{451B73A5-1563-45D5-A6AC-7B2B7D30B778}#1.1#0"; "BSPrin10.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.1#0"; "vbccr18.ocx"
Begin VB.Form escRPT 
   Caption         =   "Escolha o grupo de Relatorio"
   ClientHeight    =   6675
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10785
   Icon            =   "escrpt.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   10785
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   6840
      ScaleHeight     =   315
      ScaleWidth      =   1875
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin BSPrinter.PrintPreview PrintPreview1 
      Left            =   3840
      Top             =   0
      _ExtentX        =   1191
      _ExtentY        =   1191
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6675
      Left            =   9315
      TabIndex        =   1
      Top             =   0
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   11774
      ButtonWidth     =   1138
      ButtonHeight    =   582
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5895
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   10398
      _Version        =   393216
   End
   Begin vbExtra.FlexFn FlexFn1 
      Height          =   405
      Left            =   0
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
   Begin VBCCR18.RichTextBox RichTextBox1 
      Height          =   255
      Left            =   5160
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      MultiLine       =   -1  'True
      TextRTF         =   "escrpt.frx":058A
   End
End
Attribute VB_Name = "escRPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aORDEM As Variant
Dim aORDES As Variant
Dim cORDEM As String

Private Sub Apaga_Click()
  Dim sSQL As String
  If Grid.Row > 0 Then  ''And Grid.Row < Grid.Rows - 1 Then
    Grid.Col = 2
    zRPT = Grid.Text
    sSQL = "select * from RPT WHERE GRP='" & zgrp & "' AND RPT='" & zRPT & "'"
    If ApagaSQLP(zRPTARQ, sSQL) Then
      FilRelat
    End If
  End If
End Sub

Private Sub CmdSair_Click()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub Edit_Click()
  If Grid.Row > 0 Then  'And Grid.Row <= Grid.Rows Then
    Grid.Col = 2
    zRPT = Grid.Text
    ePASS02 = zRPTARQ
    ePASS01 = "select * from RPT WHERE GRP='" & zgrp & "' AND RPT='" & zRPT & "'"
    FrmRpt.Show vbModal
    FilRelat
  End If
End Sub

Private Sub FilRelat()
  Dim cSQL As String
  If ZGRPSUB = "" Then
    cSQL = "SELECT GRP,SUBGRP,rPT,COGNOME,NOME,ARQUIVO FROM rpt WHERE GRP='" & zgrp & "' order by " & cORDEM
  Else
    cSQL = "SELECT GRP,SUBGRP,rPT,COGNOME,NOME,ARQUIVO FROM rpt WHERE GRP='" & zgrp & "' and SUBGRP='" & ZGRPSUB & "' order by " & cORDEM
  End If
  MontaGrid Grid, 6, Array(800, 800, 1000, 1200, 5000, 800), Array("GRP", "SUB", "Codigo", "Interno", "Descricao", "EXT"), _
            Array("L$GRP", "L$SUBGRP", "L$RPT", "L$COGNOME", "L$NOME", "EXT,ARQUIVO,ARQUIVO"), zRPTARQ, cSQL
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  If InStr(UCase(cARQRTF), ".TXT") Then
     If FileExist(cARQRTF, True) Then
        ePASS03 = 1
        PrintPreview1.ShowPreview
     End If
     End
     'Unload Me
  End If
  If InStr(UCase(cARQRTF), ".RTF") Then
     If FileExist(cARQRTF, True) Then
        RichTextBox1.LoadFile cARQRTF, RtfLoadSaveFormatRTF
        ePASS03 = 2
        PrintPreview1.ShowPreview
     End If
     End
     'Unload Me
  End If
  
  If InStr(UCase(cARQRTF), ".JPG") Then
     If FileExist(cARQRTF, True) Then
        Picture1.Picture = LoadPicture(cARQRTF)
        ePASS03 = 4
        PrintPreview1.ShowPreview
     End If
     End
     'Unload Me
  End If
  
  
  aORDEM = Array("RPT,COGNOME", "COGNOME,RPT", "NOME", "SUBGRP,RPT,COGNOME", "SUBGRP,COGNOME,RPT")
  aORDES = Array("Codigo", "Cod.Int.", "Descricao", "Sub-Codigo", "Sub-Cod.Int.")
  cORDEM = "RPT,COGNOME"
  xmontatoolbar Me.Toolbar1, "escRPT", True
  FilRelat
  If FixStr(eLOCALIZA) <> "" Then
    LocalizaGri1 Grid, eLOCALIZA, 3, False 'posicao 3 codigo rpt
    eLOCALIZA = ""
    If lRETU Then
      imprima_click
      ''Alert ("achei")
    End If
  End If
End Sub
Private Sub Grid_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    imprima_click
  Else
    If KeyAscii > 31 And KeyAscii < 123 Then
      LocalizaGrid Grid, Chr(KeyAscii), 1, False
    End If
  End If
End Sub

'Private Sub Grid_SelChange()
'    With Grid
'        If .Rows > 2 Then
'            .Col = .Cols - 1
'            .ColSel = 0
'            .TopRow = .Row
'        End If
'    End With
'End Sub

Private Sub imprima_click()
  Dim nPOS, x As Integer
  Dim zrelano, zrelmes, zrelemp, zrelcog, zrelqui
  Dim cARQ, cSQL, cEXTENSAO, cARQNOME As String
  Dim lLIBGRP As Boolean
  Dim aRETU As Variant
  Dim eRUN As String
  Dim fileFile As Integer
  Dim STRBUFFER As String
  Dim cLINHA As String


  On Error Resume Next

  If Grid.Row < 1 Then  'nao entre se a selecao for cabecario row 0
    Exit Sub
  End If
  ''Pega Nome Relatorio
  Grid.Col = 2
  zRPT = Grid.Text

  cARQ = zRPTARQ
  cSQL = "select CAMINHO,LIBERAR from RPTGRP WHERE GRP='" & zgrp & "'"
  aRETU = PegSQL(cARQ, cSQL, 2, Array("CAMINHO", "LIBERAR"), Array("C", "BF"), Array("", False))
  zRPTGRPCAM = aRETU(0)
  lLIBGRP = aRETU(1)

  ''Arquivo sql Direitos
  cARQ = DBWRPT
  cSQL = "select * from " & ArqRPTUsr() & " WHERE GRP='" & zgrp & "' AND RPT='" & zRPT & "'" & "AND IDUSUARIO=" & zWRPTID

  ' 0-Novo,1-Abrir,2-Salvar,3-Salvacomo
  ' 4-Editar,5-Visualizar,6-Imprimir,7-Exportar 8-Abrir Com
  aDIREITOS = PegSQL(cARQ, cSQL, 9, _
                     Array("NOVO", "ABRIR", "SALVARTF", "SALVATXT", "EDITAR", "VISUALIZA", "IMPRIME", "EXPORTA", "ABRIRCOM"), _
                     Array("B", "B", "B", "B", "B", "B", "B", "B", "CU"), _
                     Array(False, False, False, False, False, False, False, False, ""))

  If Not aDIREITOS(5) And Not aDIREITOS(6) And (Not lLIBGRP) Then
    Alert "Não Liberado Contate o Administrador"
    Exit Sub
  End If


  cSQL = "select * from RPT WHERE GRP='" & zgrp & "' AND RPT='" & zRPT & "'"
  cARQ = zRPTARQ

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

  aARQFEC = Array(aRELCFG(16), aRELCFG(17), aRELCFG(18), aRELCFG(19), "", "", "", "", aRELCFG(6))
  ''     carquso      carqbai      carqfec      carqacu   fecdiz arqnome ccamfec arquso titulo original




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
  ''15 sqluso
  ''16 carquso
  ''17 carqbai
  ''18 carqfec
  ''19 carqacu


  ''Tenta Caminho Informado
  cARQRTF = aRELCFG(0)
  If InStr(cARQRTF, "{") > 0 Then
    cARQRTF = Caminex(cARQRTF)
  End If

  If Not FileExist(cARQRTF) Then
    ''Tenta Caminho do grupo
    cARQRTF = Caminex(zRPTGRPCAM & aRELCFG(0))
    If Not FileExist(cARQRTF) Then
      ''Tenta Caminho indicado
      cARQRTF = Caminex(ZRELAT & aRELCFG(0))
      If Not FileExist(cARQRTF) Then
        ''Tenta Caminho do Aplicativo
        cARQRTF = Caminex(App.Path & aRELCFG(0))
        If Not FileExist(cARQRTF) Then
          If aRELCFG(1) <> "SITE" Then
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


  ''Grava Utilizaçao
  aRELCFG(3) = aRELCFG(3) + 1
  GrvSQL cARQ, cSQL, 2, Array("DATAIMP", "UTILIZADO"), Array(Now, aRELCFG(3)), Array("", "N")



  ''Aarquivos
  aARQUIVOS = Array(aRELCFG(7), aRELCFG(8), aRELCFG(9), aRELCFG(10))

  'Competencia
  zrelano = 0
  zrelmes = 0
  zrelemp = 0
  zrelcog = ""
  For x = 0 To 3
    If Len(aARQUIVOS(x)) > 0 Then
      aARQUIVOS(x) = Caminex(aARQUIVOS(x), zrelano, zrelmes, zrelemp)
      zrelano = Val(eRETU01)
      zrelmes = Val(eRETU02)
      zrelemp = Val(eRETU03)
    End If
  Next x

  ''Verifica fechamento Mana5
  Mana5Fec

  cEXTENSAO = "XXX"
  nPOS = InStrRev(cARQRTF, ".")
  If nPOS > 0 Then
    cEXTENSAO = Mid(cARQRTF, nPOS + 1)
  End If
  If aRELCFG(1) = "" Or aRELCFG(1) = "_" Then  'Abrir Com Não Especificado
    Select Case cEXTENSAO
      'extensoes mais usadas na frente acelerar
    Case "RPT"
      aRELCFG(1) = "FRMCRW"
    Case "REP"
      aRELCFG(1) = "FRMREP"
    Case "RET"
      aRELCFG(1) = "VORETRUN"  ''"CARETRUN.EXE"
    Case "TXL"
      aRELCFG(1) = "FRMTXL"
    Case "SQL"
      aRELCFG(1) = "FRMSQL"
    Case "CPL"
      aRELCFG(1) = "EXECUTARDLL"
    Case "M5P"                               ''Mana5 Padrao /Mana5 Padrao
      aRELCFG(1) = "IMPRELM5P"             ''PADREL
    Case "M5M"                               ''Mana5 Padrao /Mana5 Manrel
      aRELCFG(1) = "IMPRELM5M"             ''MANREL
    Case "HTM", "HTML", "ZPL"
      aRELCFG(1) = "FRMPREVIEW"
    Case "PDF", "CHM", "HLP"
      aRELCFG(1) = "SHELL"
    Case "DOC", "XLS", "PPS", "PPT"
      aRELCFG(1) = "SHELL"
    Case "EXE"
      aRELCFG(1) = "SHELL"
    Case "INI"
      aRELCFG(1) = "FRMINIEDITOR"
    Case "TXT", "MAN"
      ePASS01 = Array("Editor Interno", "Imprimir Direto Impressora", "Preview Interno")
      escOrdem.Show vbModal, Me
      eRETU01 = FixInt(eRETU01, 0)
      Select Case eRETU01
      Case 0
        aRELCFG(1) = "FRMRTF"
      Case 1
        aRELCFG(1) = "FRMTXL"
      Case 2
        ePASS03 = 1
        PrintPreview1.ShowPreview
        Exit Sub
      End Select
      'Case "QRP"
      '   aRELCFG(1) = "qrpVIEW.EXE"
    Case "RTF"
      ePASS01 = Array("Editor Interno", "Preview Interno")
      escOrdem.Show vbModal, Me
      eRETU01 = FixInt(eRETU01, 0)
      Select Case eRETU01
      Case 0
        aRELCFG(1) = "FRMRTF"
      Case 2
        ePASS03 = 2
        PrintPreview1.ShowPreview
        Exit Sub
      End Select
      'Case "DWF"
      '    aRELCFG(1) = "FRMDWF"
      'Case "DWG", "DXF", "RML", "IPT", "IAM", "IDW"
      '    aRELCFG(1) = "FRMDWG"
    Case "WMF", "EMF", "BMP", "ICO", "JPG", "JPEG", "JNG", "KOALA", "LBM", "IFF", "MNG", "PBM", "PBMRAW", _
         "PCD", "PCX", "PGM", "PGMRAW", "PNG", "PPM", "PPMRAW", "RAS", "TARGA", _
         "DIB", "TGA", "PIC", "TIF", "TIFF", "WBMP", "PSD", "CUT", "XBM", "XPM", "DDS", "GIF", "HDR"
      aRELCFG(1) = "FRMIMG"
    Case "MPEG", "MPG", "M1V", "MP2", "MPE", "AVI", "WMF", "WAV", "SND", "AU" _
       , "AIF", "AIFC", "AIFF", "WMA", "WMA", "MP3", "MID", "RMI" _
       , "MIDI", "CDA", "WMA", "ASF", "WM", "WMV", "ASX" _
       , "WAX", "M3U", "MVX", "WMX"
      aRELCFG(1) = "FRMVID"
    Case Else
      aRELCFG(1) = "ABRIRCOM"
      ''Alert ("Abrir Com Não Configurado")
      Exit Sub
    End Select
  End If


  ''Exibe Mensagem Configurada
  If Len(aRELCFG(2)) > 0 Then Alert aRELCFG(2), "Atençao"

  ''Ajusta Nome Arquivos
  If Len(aARQUIVOS(0)) > 0 Then cARQNOME = NomeArq(aARQUIVOS(0))
  If Len(cARQNOME) = 0 Then cARQNOME = NomeArq(cARQRTF)


  If InStr(aRELCFG(6), "[ZFEC]") > 0 Then
    aRELCFG(6) = Replace(aRELCFG(6), "[ZFEC]", aARQFEC(4))
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
  If InStr(aRELCFG(6), "[ZCOG]") > 0 Then
    zrelcog = PegUltSQL(DBWRPT, "SELECT COGNOME FROM EMPRESA WHERE NUMERO=" & zrelemp, "COGNOME", "")
    aRELCFG(6) = Replace(aRELCFG(6), "[ZCOG]", zrelcog)
  End If



  ''Direitos agora no shellex abaixo
  'If aRELCFG(1) = "VORETRUN.EXE" Then cARQRTF = cARQRTF & MONTARSN(, "") & "#" & aARQUIVOS(0) & "#" & aRELCFG(15)
  'If aRELCFG(1) = "VORTF.EXE" Then cARQRTF = cARQRTF & MONTARSN()


  ''Imprime conforme Abrir Com
  Select Case aRELCFG(1)
  Case "FRMCRW"
    If MDG("usar novo visualizador") Then
      FrmCrwENG.Show vbModal, Me
    Else
      FrmCrw.Show vbModal, Me
    End If
  Case "FRMSQL"
    FrmSql.Show vbModal, Me
    'Case "FRMREP"
    '   FrmRep.Show vbModal, Me
  Case "EXECUTARDLL"
    ExecutarDLL cARQRTF
  Case "IMPRELM5P"
    'eRUN = ZMANA5IMP & "IMPREL " & zgrp & " " & zRPT & " " & zUSER & " P " & ZMANA5IMP
    'Shell eRUN, vbNormalFocus
    ShellEx "IMPREL", essSW_SHOWDEFAULT, zgrp & " " & zRPT & " " & zUSER & " P ", ZMANA5IMP, , Me.hWnd

  Case "IMPRELM5M"
    'eRUN = ZMANA5IMP & "IMPREL " & zgrp & " " & zRPT & " " & zUSER & " M " & ZMANA5IMP
    'Shell eRUN, vbNormalFocus
    ShellEx "IMPREL", essSW_SHOWDEFAULT, zgrp & " " & zRPT & " " & zUSER & " M ", ZMANA5IMP, , Me.hWnd
  Case "ABRIRCOM"
    Call OpenWith(cARQRTF, OAIF_ALLOW_REGISTRATION Or OAIF_EXEC Or OAIF_FORCE_REGISTRATION, Me.hWnd)
  Case "SHELL"
    ShellEx cARQRTF, essSW_SHOWDEFAULT, , , , Me.hWnd
  Case "FRMRTF"
    FrmRTf.Show vbModal, Me
  Case "FRMINIEDITOR"
    frmIniEditor.Show vbModal, Me
    'Case "FRWCRWVIEW"
    '     FrmCrwView.Show vbModal, Me
  Case "FRMCRWENG"
    FrmCrwENG.Show vbModal, Me
  Case "FRMIMG"
      If FileExist(cARQRTF, True) Then
        Picture1.Picture = LoadPicture(cARQRTF)
        ePASS03 = 4
        PrintPreview1.ShowPreview
     End If
  Case "FRMPREVIEW"
    If cEXTENSAO = "ZPL" Then
      fileFile = FreeFile
      Open cARQRTF For Input As #fileFile
      Do While Not EOF(fileFile)
        'read line
        Input #fileFile, STRBUFFER
        cLINHA = cLINHA + STRBUFFER
      Loop
      Close fileFile
      cLINHA = Replace(cLINHA, Chr(13), "")
      cLINHA = Replace(cLINHA, Chr(10), "")
      ePASS02 = "http://api.labelary.com/v1/printers/8dpmm/labels/4x6/0/"
      ePASS02 = ePASS02 + Chr(34) + cLINHA + Chr(34)
      OpenUrl (ePASS02)
      Exit Sub
    End If

    ePASS01 = Array("Navegador", "Preview Interno")
    escOrdem.Show vbModal, Me
    eRETU01 = FixInt(eRETU01, 0)
    Select Case eRETU01
    Case 0
      OpenUrl (cARQRTF)
    Case 1
      ePASS03 = 3
      PrintPreview1.ShowPreview
      Exit Sub
    End Select

  Case "FRMVID"
    FrmVid.Show vbModal, Me
    ''Encerra Para Evitar Nova Abertura Sistema Novo
    Exit Sub
  End Select

  '' Executavel
  If InStr(aRELCFG(1), ".EXE") Then
    ' If aRELCFG(1) = "MKRVIEW.EXE" Then
    '     cARQRTF = Chr(34) & aARQUIVOS(0) & Chr(34) & " " & Chr(34) & "SYSDBA" & Chr(34) & " " & Chr(34) & "masterkey" & Chr(34) & " " & Chr(34) & " " & Chr(34) & " " & Chr(34) & "none" & Chr(34) & " " & Chr(34) & cARQRTF & Chr(34)
    ' End If
    aRELCFG(1) = Trim(Replace(aRELCFG(1), ".EXE", " "))
    Select Case aRELCFG(1)
    Case "CARETRUN"
      'Shell "caretrun  " & cARQRTF, vbNormalFocus
      Alert ("configur para voretrun")
      Exit Sub
    Case "VORETRUN"
      'Shell "VORETRUN " & cARQRTF, vbNormalFocus
      ShellEx "VORETRUN", essSW_SHOWDEFAULT, cARQRTF & MONTARSN(, "") & "#" & aARQUIVOS(0) & "#" & aRELCFG(15), PegPath("PATH", "VORETRUN"), , Me.hWnd
      'Case "MKRVIEW"
      '    ShellEx aRELCFG(1), essSW_SHOWDEFAULT, cARQRTF, , , Me.hWnd
    Case Else
      ShellEx aRELCFG(1), essSW_SHOWDEFAULT, cARQRTF, , , Me.hWnd
    End Select


  End If


End Sub

Private Sub liberar_click()
  If Grid.Row > 0 Then  'And Grid.Row < Grid.Rows - 1 Then
    Grid.Col = 2
    zRPT = Grid.Text
    escrptusr.Show vbModal
  End If
End Sub

Function MONTARSN(Optional ByVal cDELI As String = "$", Optional ByVal cSEP As String = ",") As String
  Dim x As Integer
  MONTARSN = cDELI
  For x = 0 To 7
    MONTARSN = MONTARSN & IIf(aDIREITOS(x), "S", "N") & cSEP
  Next x
  MONTARSN = MONTARSN & cDELI
End Function

Private Sub Novo_Click()
  Dim cSQL As String
  zRPT = InputBox("Digite o Codigo", "Inclusão Relatorio", "__")
  zRPT = FixStr(zRPT, "", "TRIM", 8)
  If Len(zRPT) > 0 Then
    cSQL = "select * from RPT WHERE GRP='" & zgrp & "' AND RPT='" & zRPT & "'"
    IncluiSQL zRPTARQ, cSQL, 4, Array("GRP", "ARQUIVO", "RPT", "DATACRI"), _
              Array(zgrp, zRPT & ".", zRPT, Today()), True, True
    If lRETU Then
      FilRelat
    End If
  End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim sButton As String
  sButton = Button
  sButton = Left(UCase(Replace(sButton, "&", "")), 3)
  If Not AcessaBtnOld("escRPT", Button.Index) Then
    Exit Sub
  End If

  GravaLog 0, Button.Index, sButton, "escRPT"

  Select Case sButton
  Case "ORD"
    ePASS01 = aORDES
    escOrdem.Show vbModal, Me
    If lRETU Then
      cORDEM = aORDEM(eRETU01)
      FilRelat
    End If
  Case "NOV"
    Novo_Click
  Case "EDI"
    Edit_Click
  Case "IMP"
    imprima_click
  Case "LIB"
    liberar_click
  Case "EXC"
    Apaga_Click
  Case "LOC"
    LocalizaGri1 Grid, , 2
  Case "CON"
    FrmPrintSetup.Show vbModal, Me
  Case "SAI"
    CmdSair_Click
  End Select

End Sub


Public Sub PrintPreview1_PrepareReport(Cancel As Boolean)
  If ePASS03 = 1 Then  'TXT
    MyPrintingTXT
  End If
  If ePASS03 = 2 Then  'rtf
    MyPrintingRTF
  End If
  If ePASS03 = 3 Then  'HTML
    MyPrintinghtml
  End If
  If ePASS03 = 4 Then 'JPG
      ePASS01 = Array("Esquerda Superior", "Direita Superior", "Esquera Inferior", "Direita Inferior", "Centralizado")
      escOrdem.Show vbModal, Me
      ePASS02 = FixInt(eRETU01, 0)
      MyPrintingJPG
  End If
End Sub

Public Sub MyPrintinghtml()
  Dim cTEXTO As String
  Dim cLINHA As String
  Dim LINES() As String
  Dim i As Integer

  ' If Not FileExist(cARQRTF, True) Then 'ja checado cmdvisualclick
  '     Exit Sub
  ' End If
  cTEXTO = FileText(cARQRTF)
  cTEXTO = HtmlToText(cTEXTO)

  LINES = Split(cLINHA, vbCrLf)

  For i = 0 To UBound(LINES)
    Printer.Print LINES(i)
  Next
End Sub
Public Sub MyPrintingRTF()
  PrinterEx.PrintRichTextBox RichTextBox1
End Sub
Public Sub MyPrintingJPG()
  Select Case ePASS02
   Case 1
       Printer.PaintPicture Picture1.Picture, 0, 0
   Case 1
       Printer.PaintPicture Picture1.Picture, Printer.ScaleWidth - Picture1.ScaleX(Picture1.Picture.Width, vbHimetric, Printer.ScaleMode), 0
  Case 2
    Printer.PaintPicture Picture1.Picture, 0, Printer.ScaleHeight - Picture1.ScaleY(Picture1.Picture.Height, vbHimetric, Printer.ScaleMode)
  Case 3
    Printer.PaintPicture Picture1.Picture, Printer.ScaleWidth - Picture1.ScaleX(Picture1.Picture.Width, vbHimetric, Printer.ScaleMode), Printer.ScaleHeight - Picture1.ScaleY(Picture1.Picture.Height, vbHimetric, Printer.ScaleMode)

  Case 4
    Printer.PaintPicture Picture1.Picture, (Printer.ScaleWidth - Picture1.ScaleX(Picture1.Picture.Width, vbHimetric, Printer.ScaleMode)) / 2, (Printer.ScaleHeight - Picture1.ScaleY(Picture1.Picture.Height, vbHimetric, Printer.ScaleMode)) / 2

  
  Case Else
  Printer.PaintPicture Picture1.Picture, 0, 0
       
 End Select
End Sub

Public Sub MyPrintingTXT()
  Dim fileFile As Integer
  Dim STRBUFFER As String
  'If Not FileExist(cARQRTF, True) Then 'ja checado na cmdvisual click
  '    Exit Sub
  ' End If
  fileFile = FreeFile
  Open cARQRTF For Input As #fileFile
  Do While Not EOF(fileFile)
    'read line
    Input #fileFile, STRBUFFER
    Printer.Print STRBUFFER
  Loop
  Close fileFile
End Sub


