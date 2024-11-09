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
Private mvarhandle As Integer                    'local copy
Private mvarDESTINO As String                    'local copy
Private mvarcaminho As String                    'local copy

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
    Printer.Orientation = PrinterObjectConstants.vbPRORPortrait
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


  ''If nCOL > 1 Then
  ''   cLINHA = Mid(mvaraLINHAS(nROW), 1, nCOL - 1)
  ''   cLINHA = cLINHA + cTexto
  ''   cLINHA = cLINHA + Mid(mvaraLINHAS(nROW), nCOL + nLEN)
  ''Else
  ''   cLINHA = cTexto + Mid(mvaraLINHAS(nROW), nLEN + 1)
  ''End If
  mvaraLINHAS(nROW) = cLINHA
End Sub
'Imprime configuracao incial da impressora (escape codes outros)
Public Sub impsetup(cSETUP As String)
Dim cDESTINO As String
'LPT1,LPT2.... COM1,COM2...
If Mid(mvarDESTINO, 1, 3) = "LPT" Or Mid(mvarDESTINO, 1, 3) = "COM" Or mvarDESTINO = "IMPNET" Then
   Print #mvarhandle, cSETUP
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
    'COM1,COM2... LPT1,LPT2...
    If Mid(mvarDESTINO, 1, 3) = "LPT" Or Mid(mvarDESTINO, 1, 3) = "COM" Or mvarDESTINO = "IMPNET" Or mvarDESTINO = "ARQ" Then
      Print #mvarhandle, mvaraLINHAS(i)
    End If
  Next
  If mvarDESTINO = "PRINTER" Then
    ''Printer.EndDoc
    ''Printer.NewPage
  End If
End Sub

Public Property Let Handle(ByVal vData As Integer)
  mvarhandle = vData
End Property

Public Property Get Handle() As Integer
  Handle = mvarhandle
End Property

Public Sub fechar()
  If mvarDESTINO = "ARQ" Then
    Close (mvarhandle)
  End If
  'COM1,COM2... LPT1,LPT2...
  If Mid(mvarDESTINO, 1, 3) = "LPT" Or Mid(mvarDESTINO, 1, 3) = "COM" Or mvarDESTINO = "IMPNET" Then
    Close #mvarhandle
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
  txttoRTF mvarArquivo, cDESTINO
End Sub
Public Sub GravaPDF(Optional cDESTINO As String = "")
  ePASS01 = mvarArquivo
  formConvertToPDF.Show
End Sub


Public Sub arqimpdos(Optional ByVal cDESTINO As String = ">PRN")
''Caminho Rede "\\Server\HP"
''">LPT1"
''Mapeamento impressora rede e usb
''net use lpt1 \\computer\printer password /persistent:yes
  If mvarDESTINO = "ARQ" Then
    'Shell "copy " & mvarArquivo & " " & cDESTINO
    CopyFileWindowsWay mvarArquivo, cDESTINO
  End If
End Sub

Public Sub ABRIR()
Dim cVAROPEN As String
  If mvarDESTINO = "ARQ" Then
    Open mvarArquivo For Output As #mvarhandle
  End If
  If Mid(mvarDESTINO, 1, 3) = "LPT" Then
    'LPT1,LPT2,LPT3...
    cVAROPEN = mvarDESTINO + ":"
    'Open "LPT1:" For Output As #mvarhandle
    Open cVAROPEN For Output As #mvarhandle
  End If
  
  If Mid(mvarDESTINO, 1, 3) = "COM" Then
    'COM1,COM2,COM3...
    cVAROPEN = mvarDESTINO + ":"
    'Open "COM1:" For Output As #mvarhandle
    Open cVAROPEN For Output As #mvarhandle
  End If
  If mvarDESTINO = "IMPNET" Then
    Open mvarcaminho For Output As #mvarhandle
  End If
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
  mvarhandle = FreeFile
  mvarDESTINO = "ARQ"
  mvarcaminho = "LPT1"
End Sub

Public Property Let nCOL(ByVal vData As Integer)
  mvarnCOL = vData
End Property

Public Property Get nCOL() As Integer
  nCOL = mvarnCOL
End Property
'usando bsprint no lugar vbextra preview
'definicoes em mprintpreview.bas
'Private Property Get Printer() As Printer
'  Set Printer = vbExtra.Printer2
'End Property
'Private Property Set Printer(nPrinter As Printer)
'  Set vbExtra.Printer2 = nPrinter
'End Property
Public Sub MyPrintingTXT()
  Dim fileFile As Integer
  Dim STRBUFFER As String
  fileFile = FreeFile
  Open cARQRTF For Input As #fileFile
  Do While Not EOF(fileFile)
    'read line
    Input #fileFile, STRBUFFER
    Printer.Print STRBUFFER
  Loop
  Close fileFile
  Printer.EndDoc
End Sub

