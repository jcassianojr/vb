VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Criatxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarnCOL As Integer                      'local copy
Private mvarnROW As Integer                      'local copy
Private mvarArquivo As String                    'local copy
Private mvaraLINHAS As Variant                   'local copy
Private mvarDESTINO As String                    'local copy
Private mvarcaminho As String                    'local copy

' VARIÁVEIS MIGRADAS DO MOTOR FSO
Private objFSO As Object                         ' Objeto principal FSO
Private objStreamOut As Object                   ' Objeto de fluxo para gravação (Stream)

Public Property Let caminho(ByVal vData As String)
  mvarcaminho = vData
End Property

Public Property Get caminho() As String
  caminho = mvarcaminho
End Property

Public Property Let Destino(ByVal vData As String)
  On Error Resume Next                         ''Algumas impressora nao Aceitam Orientacao
  mvarDESTINO = vData
  If mvarDESTINO = "PRINTER" Then
    Printer.FontSize = 12
    Printer.orientation = PrinterObjectConstants.vbPRORPortrait
    Printer.PaperSize = PrinterObjectConstants.vbPRPSA4
    Printer.ColorMode = PrinterObjectConstants.vbPRCMMonochrome
  End If
End Property

Public Property Get Destino() As String
  Destino = mvarDESTINO
End Property

Public Sub gravalincon(ByVal cTEXTO As String, ByVal nROW As Integer, ByVal nCOL As Integer)
  Dim cLINHA As String
  Dim nLEN As Integer
  nLEN = Len(cTEXTO)
  If nLEN = 0 Then
    ''Texto em Branco
    Exit Sub
  End If
  If nROW > mvarnROW Or nROW < 1 Then
    Alert ("Cria txt Linha Invalida")
    ''Linhas Acima das dimensionadas
    Exit Sub
  End If
  If nCOL > mvarnCOL Or nCOL < 1 Then
    ''Colunas acima das Dimensionadas
    Alert ("Cria txt coluna Invalida")
    Exit Sub
  End If
  nROW = nROW - 1                              ''Pegar Linha Correta da Matriz que Comeca com 0
  cLINHA = mvaraLINHAS(nROW)
  If nCOL > 1 Then
    cLINHA = Mid(cLINHA, 1, nCOL - 1) & cTEXTO & Mid(cLINHA, nCOL + nLEN)
  Else
    cLINHA = cTEXTO & Mid(cLINHA, nLEN + 1)
  End If

  mvaraLINHAS(nROW) = cLINHA
End Sub

'Imprime configuracao incial da impressora (escape codes outros)
Public Sub impsetup(cSETUP As String)
  'LPT1,LPT2.... COM1,COM2... IMPNET ou arquivo físico via FSO
  If Mid(mvarDESTINO, 1, 3) = "LPT" Or Mid(mvarDESTINO, 1, 3) = "COM" Or mvarDESTINO = "IMPNET" Or mvarDESTINO = "ARQ" Then
     If Not objStreamOut Is Nothing Then
        objStreamOut.WriteLine cSETUP
     End If
  End If
  If mvarDESTINO = "PRINTER" Then
     Printer.Print cSETUP
  End If
End Sub

Public Sub implinha(cLINHA As String)
  impsetup (cLINHA)
End Sub

Public Sub destini(ByVal cGRUPO As String, ByVal cCHAVE As String)
  Dim cDESTINO As String
  cDESTINO = PegPath(cGRUPO, cCHAVE)
'LPT1,LPT2.... COM1,COM2...
If Mid(cDESTINO, 1, 3) = "LPT" Or Mid(cDESTINO, 1, 3) = "COM" Or mvarDESTINO = "PRINTER" Then
   Me.Destino = cDESTINO
End If
If cDESTINO = "IMPNET" Then
    Me.Destino = "IMPNET"
    Me.caminho = cDESTINO
End If
End Sub

Public Sub gravafolha()
  Dim i As Integer
  For i = 0 To mvarnROW - 1
  
    If mvarDESTINO = "PRINTER" Then
      Printer.Print mvaraLINHAS(i)
    End If
    'COM1,COM2... LPT1,LPT2... IMPNET ou ARQ (unificados sob o motor FSO)
    If Mid(mvarDESTINO, 1, 3) = "LPT" Or Mid(mvarDESTINO, 1, 3) = "COM" Or mvarDESTINO = "IMPNET" Or mvarDESTINO = "ARQ" Then
      If Not objStreamOut Is Nothing Then
         objStreamOut.WriteLine mvaraLINHAS(i)
      End If
    End If
  Next
End Sub

' MANTIDO POR COMPATIBILIDADE LEGADA (Caso alguma rotina chame a propriedade)
Public Property Let Handle(ByVal vData As Integer)
  ' Não atribui mais canal numérico interno do VB6
End Property

Public Property Get Handle() As Integer
  Handle = 1 ' Retorna um valor fictício fixo apenas para não quebrar chamadas antigas
End Property

Public Sub fechar()
  ' Fecha o stream FSO se ele estiver ativo
  If Not objStreamOut Is Nothing Then
     objStreamOut.Close
     Set objStreamOut = Nothing
  End If
  
  If mvarDESTINO = "PRINTER" Then
    ''Printer.EndDoc
  End If
End Sub

Public Sub arqview()
  cARQRTF = mvarArquivo
  If FileConnExist(cARQRTF, True) Then
    PrinterEx.ShowPrintPreview Me, "MyPrintingTXT"
  End If
End Sub

Public Sub ArqEdit()
  cARQRTF = mvarArquivo
  ' 0-Novo,1-Abrir,2-Salvar,3-Salvacomo
  ' 4-Editar,5-Visualizar,6-Imprimir,7-Exportar
  aDIREITOS = Array(False, False, False, False, False, True, True, False)
  FrmRtfView.Show vbModal
End Sub

' ADICIONADO: Nova assinatura e novas funções de exportação integradas das interações anteriores
Public Sub GravaWord(Optional cDESTINO As String = "", Optional cTITULO As String = "", Optional cAUTOR As String = "")
  If cDESTINO = "" Then
    cDESTINO = TrocaExt(mvarArquivo, "doc")
  End If
  txttodoc mvarArquivo, cDESTINO, cTITULO, cAUTOR
End Sub

Public Sub GravaExcel(Optional cDESTINO As String = "", Optional cTITULO As String = "", Optional cAUTOR As String = "", Optional cDELIMITADOR As String = ";")
  If cDESTINO = "" Then
    cDESTINO = TrocaExt(mvarArquivo, "xls")
  End If
  txttoxls mvarArquivo, cDESTINO, cTITULO, cAUTOR, cDELIMITADOR
End Sub

Public Sub GravaHtml(Optional cDESTINO As String = "", Optional cTITULO As String = "", Optional cAUTOR As String = "")
  If cDESTINO = "" Then
    cDESTINO = TrocaExt(mvarArquivo, "html")
  End If
  txttohtml mvarArquivo, cDESTINO, cTITULO, cAUTOR
End Sub

Public Sub GravaRTF(Optional cDESTINO As String = "")
  If cDESTINO = "" Then
    cDESTINO = TrocaExt(mvarArquivo, "rtf")
  End If
  txttortf mvarArquivo, cDESTINO
End Sub

Public Sub GravaPDF(Optional cDESTINO As String = "")
  ePASS01 = mvarArquivo
  formConvertToPDF.Show
End Sub

Public Sub arqimpdos(Optional ByVal cDESTINO As String = ">PRN")
  If mvarDESTINO = "ARQ" Then
    CopyFileWindowsWay mvarArquivo, cDESTINO
  End If
End Sub

Public Sub ABRIR()
  Dim cVAROPEN As String
  On Error GoTo TrataErro
  
  ' Inicializa o FileSystemObject localmente
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  
  ' Determina qual rota abrir baseado no mvarDESTINO usando FSO
  If mvarDESTINO = "ARQ" Then
    ' 2 = ForWriting, True = Create
    Set objStreamOut = objFSO.OpenTextFile(mvarArquivo, 2, True)
  
  ElseIf Mid(mvarDESTINO, 1, 3) = "LPT" Then
    cVAROPEN = mvarDESTINO & ":"
    Set objStreamOut = objFSO.OpenTextFile(cVAROPEN, 2, True)
    
  ElseIf Mid(mvarDESTINO, 1, 3) = "COM" Then
    cVAROPEN = mvarDESTINO & ":"
    Set objStreamOut = objFSO.OpenTextFile(cVAROPEN, 2, True)
    
  ElseIf mvarDESTINO = "IMPNET" Then
    Set objStreamOut = objFSO.OpenTextFile(mvarcaminho, 2, True)
  End If

  Exit Sub
TrataErro:
  Alert "Erro ao inicializar fluxo de escrita (FSO): " & Err.Description
End Sub

Public Sub novafolha()
  Dim x As Integer
  ReDim mvaraLINHAS(mvarnROW)
  For x = 0 To mvarnROW - 1
    mvaraLINHAS(x) = Space(mvarnCOL)
  Next x
End Sub

Public Property Let aLINHAS(ByVal vData As Variant)
  mvaraLINHAS = vData
End Property

Public Property Set aLINHAS(ByVal vData As Variant)
  Set mvaraLINHAS = vData
End Property

Public Property Get aLINHAS() As Variant
  If IsObject(mvaraLINHAS) Then
    Set aLINHAS = mvaraLINHAS
  Else
    aLINHAS = mvaraLINHAS
  End If
End Property

Public Property Let arquivo(ByVal vData As String)
  mvarArquivo = vData
End Property

Public Property Get arquivo() As String
  arquivo = mvarArquivo
End Property

Public Property Let nROW(ByVal vData As Integer)
  mvarnROW = vData
End Property

Public Property Get nROW() As Integer
  nROW = mvarnROW
End Property

Private Sub Class_Initialize()
  mvarnROW = 66
  mvarnCOL = 80
  mvarArquivo = App.Path & "\" & zUSER & Format(Now, "_DDMMYYYY_HHMMSS") & ".TXT"
  mvarDESTINO = "ARQ"
  mvarcaminho = "LPT1"
End Sub

Private Sub Class_Terminate()
  ' Elimina referências de memória de objetos instanciados
  Call fechar
  Set objFSO = Nothing
End Sub

Public Property Let nCOL(ByVal vData As Integer)
  mvarnCOL = vData
End Property

Public Property Get nCOL() As Integer
  nCOL = mvarnCOL
End Property

Public Sub MyPrintingTXT()
  Dim objStreamIn As Object
  Dim STRBUFFER As String
  
  ' Substitui o antigo Open For Input nativo por FSO
  If Not objFSO Is Nothing Then
     If objFSO.FileExists(cARQRTF) Then
        ' 1 = ForReading
        Set objStreamIn = objFSO.OpenTextFile(cARQRTF, 1, False)
        Do While Not objStreamIn.AtEndOfStream
           STRBUFFER = objStreamIn.ReadLine
           Printer.Print STRBUFFER
        Loop
        objStreamIn.Close
        Set objStreamIn = Nothing
        Printer.EndDoc
     End If
  End If
End Sub
