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

Private mvarnCOL As Integer
Private mvarnROW As Integer
Private mvarArquivo As String
Private mvaraLINHAS As Variant
Private mvarDESTINO As String
Private mvarcaminho As String

Private objFSO As Object
Private objStreamOut As Object

Public Property Let caminho(ByVal vData As String)
  mvarcaminho = vData
End Property

Public Property Get caminho() As String
  caminho = mvarcaminho
End Property

Public Property Let Destino(ByVal vData As String)
  On Error Resume Next
  mvarDESTINO = vData
  If mvarDESTINO = "PRINTER" Then
    Printer.FontSize = 10
    Printer.orientation = PrinterObjectConstants.vbPRORPortrait
    Printer.PaperSize = PrinterObjectConstants.vbPRPSA4
    Printer.ColorMode = PrinterObjectConstants.vbPRCMMonochrome
  End If
End Property

Public Property Get Destino() As String
  Destino = mvarDESTINO
End Property

' =================================================================================
' CALCULA O TAMANHO VISÍVEL DA STRING (Ignora as tags para não desalinhar colunas)
' =================================================================================
Private Function GetLenVisible(ByVal cTEXTO As String) As Integer
    Dim cTemp As String
    cTemp = cTEXTO
    ' 1. Converte escapes antigos se houver
    cTemp = ParseEscapeToBBCode(cTemp)
    ' 2. Limpa todas as tags BBCode para contar só o texto útil
    cTemp = Replace(cTemp, "[B]", "", , , vbTextCompare)
    cTemp = Replace(cTemp, "[/B]", "", , , vbTextCompare)
    cTemp = Replace(cTemp, "[I]", "", , , vbTextCompare)
    cTemp = Replace(cTemp, "[/I]", "", , , vbTextCompare)
    cTemp = Replace(cTemp, "[U]", "", , , vbTextCompare)
    cTemp = Replace(cTemp, "[/U]", "", , , vbTextCompare)
    cTemp = Replace(cTemp, "[PAGE]", "", , , vbTextCompare)
    cTemp = Replace(cTemp, "[SIZE=8]", "", , , vbTextCompare)
    cTemp = Replace(cTemp, "[SIZE=12]", "", , , vbTextCompare)
    cTemp = Replace(cTemp, "[SIZE=14]", "", , , vbTextCompare)
    cTemp = Replace(cTemp, "[/SIZE]", "", , , vbTextCompare)
    ' 3. Remove binários
    cTemp = TiraControlChars(cTemp)
    GetLenVisible = Len(cTemp)
End Function

Public Sub gravalincon(ByVal cTEXTO As String, ByVal nROW As Integer, ByVal nCOL As Integer)
  Dim cLINHA As String
  Dim nLEN As Integer
  Dim nVisLen As Integer
  
  nLEN = Len(cTEXTO)
  If nLEN = 0 Then Exit Sub
  
  If nROW > mvarnROW Or nROW < 1 Then
    Alert ("Cria txt Linha Invalida")
    Exit Sub
  End If
  If nCOL > mvarnCOL Or nCOL < 1 Then
    Alert ("Cria txt coluna Invalida")
    Exit Sub
  End If
  
  nROW = nROW - 1
  cLINHA = mvaraLINHAS(nROW)
  
  ' Calcula quantos espaços visuais a string consome
  nVisLen = GetLenVisible(cTEXTO)
  
  If nCOL > 1 Then
    cLINHA = Mid(cLINHA, 1, nCOL - 1) & cTEXTO & Mid(cLINHA, nCOL + nVisLen)
  Else
    cLINHA = cTEXTO & Mid(cLINHA, nVisLen + 1)
  End If

  mvaraLINHAS(nROW) = cLINHA
End Sub

' =================================================================================
' MOTOR DE IMPRESSÃO NATIVA VB6 (Interpretação do BBCode em tempo real)
' =================================================================================
Private Sub PrintToVBPrinter(ByVal cLINHA As String)
    Dim nPosIni As Long, nPosFim As Long
    Dim cText As String, cCmd As String, cTag As String, cVal As String
    
    cLINHA = ParseEscapeToBBCode(cLINHA)
    cLINHA = TiraControlChars(cLINHA)
    
    Do While Len(cLINHA) > 0
        nPosIni = InStr(1, cLINHA, "[")
        If nPosIni > 0 Then
            ' Imprime o fragmento antes da tag SEM pular linha (Ponto e Vírgula)
            If nPosIni > 1 Then
                cText = Left(cLINHA, nPosIni - 1)
                Printer.Print cText;
            End If
            
            nPosFim = InStr(nPosIni, cLINHA, "]")
            If nPosFim > 0 Then
                cTag = Mid(cLINHA, nPosIni + 1, nPosFim - nPosIni - 1)
                cCmd = UCase(cTag)
                
                If Left(cCmd, 5) = "SIZE=" Then
                    cVal = Mid(cCmd, 6)
                    If cVal = "8" Then
                        Printer.FontSize = 8
                    ElseIf cVal = "12" Then
                        Printer.FontSize = 10
                    ElseIf cVal = "14" Then
                        Printer.FontSize = 12
                    Else
                        Printer.FontSize = Val(cVal)
                    End If
                    cCmd = "SIZE"
                End If
                
                ' Liga e Desliga as propriedades da fonte instantaneamente
                Select Case cCmd
                    Case "B": Printer.FontBold = True
                    Case "/B": Printer.FontBold = False
                    Case "I": Printer.FontItalic = True
                    Case "/I": Printer.FontItalic = False
                    Case "U": Printer.FontUnderline = True
                    Case "/U": Printer.FontUnderline = False
                    Case "PAGE": Printer.NewPage
                End Select
                
                cLINHA = Mid(cLINHA, nPosFim + 1)
            Else
                Printer.Print cLINHA;
                cLINHA = ""
            End If
        Else
            Printer.Print cLINHA;
            cLINHA = ""
        End If
    Loop
    ' Ao final da instrução completa, executa a quebra de linha normal
    Printer.Print ""
End Sub

Public Sub impsetup(cSETUP As String)
  If Mid(mvarDESTINO, 1, 3) = "LPT" Or Mid(mvarDESTINO, 1, 3) = "COM" Or mvarDESTINO = "IMPNET" Or mvarDESTINO = "ARQ" Then
     If Not objStreamOut Is Nothing Then
        objStreamOut.WriteLine cSETUP
     End If
  End If
  If mvarDESTINO = "PRINTER" Then
     PrintToVBPrinter cSETUP
  End If
End Sub

Public Sub implinha(cLINHA As String)
  impsetup (cLINHA)
End Sub

Public Sub destini(ByVal cGRUPO As String, ByVal cCHAVE As String)
  Dim cDESTINO As String
  cDESTINO = PegPath(cGRUPO, cCHAVE)
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
      PrintToVBPrinter mvaraLINHAS(i)
    End If
    If Mid(mvarDESTINO, 1, 3) = "LPT" Or Mid(mvarDESTINO, 1, 3) = "COM" Or mvarDESTINO = "IMPNET" Or mvarDESTINO = "ARQ" Then
      If Not objStreamOut Is Nothing Then
         objStreamOut.WriteLine mvaraLINHAS(i)
      End If
    End If
  Next
End Sub

Public Property Let Handle(ByVal vData As Integer)
End Property

Public Property Get Handle() As Integer
  Handle = 1
End Property

Public Sub fechar()
  If Not objStreamOut Is Nothing Then
     objStreamOut.Close
     Set objStreamOut = Nothing
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
  aDIREITOS = Array(False, False, False, False, False, True, True, False)
  FrmRtfView.Show vbModal
End Sub

Public Sub GravaWord(Optional cDESTINO As String = "", Optional cTITULO As String = "", Optional cAUTOR As String = "")
  If cDESTINO = "" Then cDESTINO = TrocaExt(mvarArquivo, "doc")
  txttodoc mvarArquivo, cDESTINO, cTITULO, cAUTOR
End Sub

Public Sub GravaExcel(Optional cDESTINO As String = "", Optional cTITULO As String = "", Optional cAUTOR As String = "", Optional cDELIMITADOR As String = ";")
  If cDESTINO = "" Then cDESTINO = TrocaExt(mvarArquivo, "xls")
  txttoxls mvarArquivo, cDESTINO, cTITULO, cAUTOR, cDELIMITADOR
End Sub

Public Sub GravaHtml(Optional cDESTINO As String = "", Optional cTITULO As String = "", Optional cAUTOR As String = "")
  If cDESTINO = "" Then cDESTINO = TrocaExt(mvarArquivo, "html")
  txttohtml mvarArquivo, cDESTINO, cTITULO, cAUTOR
End Sub

Public Sub GravaRTF(Optional cDESTINO As String = "")
  If cDESTINO = "" Then cDESTINO = TrocaExt(mvarArquivo, "rtf")
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
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  If mvarDESTINO = "ARQ" Then
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
  
  If Not objFSO Is Nothing Then
     If objFSO.FileExists(cARQRTF) Then
        Set objStreamIn = objFSO.OpenTextFile(cARQRTF, 1, False)
        Do While Not objStreamIn.AtEndOfStream
           STRBUFFER = objStreamIn.ReadLine
           ' Passa pelo motor de parser em vez de jogar cru no Printer.Print
           PrintToVBPrinter STRBUFFER
        Loop
        objStreamIn.Close
        Set objStreamIn = Nothing
        Printer.EndDoc
     End If
  End If
End Sub
