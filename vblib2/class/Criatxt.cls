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

Public Sub gravalincon(ByVal cTexto As String, ByVal nROW As Integer, ByVal nCOL As Integer)
    Dim CLINHA As String
    Dim nLEN As Integer
    nLEN = Len(cTexto)
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
    CLINHA = mvaraLINHAS(nROW)
    If nCOL > 1 Then
        CLINHA = Mid(CLINHA, 1, nCOL - 1) & cTexto & Mid(CLINHA, nCOL + nLEN)
    Else
        CLINHA = cTexto & Mid(CLINHA, nLEN + 1)
    End If


    ''If nCOL > 1 Then
    ''   cLINHA = Mid(mvaraLINHAS(nROW), 1, nCOL - 1)
    ''   cLINHA = cLINHA + cTexto
    ''   cLINHA = cLINHA + Mid(mvaraLINHAS(nROW), nCOL + nLEN)
    ''Else
    ''   cLINHA = cTexto + Mid(mvaraLINHAS(nROW), nLEN + 1)
    ''End If
    mvaraLINHAS(nROW) = CLINHA
End Sub

Public Sub impsetup(cSETUP As String)
    Select Case mvarDESTINO
    Case "PRINTER"
        Printer.Print cSETUP
    Case "LPT1", "LPT2", "LPT3", "IMPNET", "COM1", "COM2", "ARQ"
        Print #mvarhandle, cSETUP
    End Select
End Sub

Public Sub implinha(CLINHA As String)
    impsetup (CLINHA)
End Sub

Public Sub destini(ByVal cGRUPO As String, ByVal cCHAVE As String)
    Dim cDESTINO As String
    cDESTINO = PegPath(cGRUPO, cCHAVE)
    Select Case cDESTINO
    Case "LPT1", "LPT2", "LPT3", "IMPNET", "COM1", "COM2", "ARQ", "PRINTER"
        Me.Destino = cDESTINO
    Case "IMPNET"
        Me.Destino = "IMPNET"
        Me.caminho = cDESTINO
    End Select
End Sub

Public Sub gravafolha()
    Dim i As Integer
    For i = 0 To mvarnROW - 1
        Select Case mvarDESTINO
        Case "PRINTER"
            Printer.Print mvaraLINHAS(i)
        Case "LPT1", "LPT2", "LPT3", "IMPNET", "COM1", "COM2", "ARQ"
            Print #mvarhandle, mvaraLINHAS(i)
        End Select
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
    Select Case mvarDESTINO
    Case "ARQ"
        Close (mvarhandle)
    Case "LPT1", "LPT2", "LPT3", "IMPNET", "COM1", "COM2"
        Close #mvarhandle
    Case "PRINTER"
        ''Printer.EndDoc
    End Select
End Sub

Public Sub arqview()
    cARQRTF = mvarArquivo
' 0-Novo,1-Abrir,2-Salvar,3-Salvacomo
            ' 4-Editar,5-Visualizar,6-Imprimir,7-Exportar
           aDIREITOS = Array(False, False, False, False, False, True, True, False)
    FrmRTf.Show vbModal
End Sub

Public Sub gravahtml(Optional cDESTINO As String = "", Optional cTITULO As String = "", Optional cAUTOR As String = "")
    If cDESTINO = "" Then
       cDESTINO = TrocaExt(mvarArquivo, "html")
    End If
    txttohtml mvarArquivo, cDESTINO, cTITULO, cAUTOR
End Sub

Public Sub gravartf(Optional cDESTINO As String = "")
    If cDESTINO = "" Then
       cDESTINO = TrocaExt(mvarArquivo, "rtf")
    End If
    txttoRTF mvarArquivo, cDESTINO
End Sub
Public Sub gravapdf(Optional cDESTINO As String = "")
   'If cDESTINO = "" Then
   '    cDESTINO = TrocaExt(mvarArquivo, "pdf") 'no formConvertToPDF ja tem trocaext
   ' End If
    'ConvertToPDF(FileName As String, outputfile As String, _
    '                    Optional TextAuthor As String, Optional TextCreator As String, Optional TextKeywords As String, _
    '                    Optional TextSubject As String, Optional TextTitle As String, _
    '                    Optional FontName As String = "Courier", Optional FontSize As Integer = 10, Optional Rotation As Integer, _
    '                    Optional pwidth As Single = 8.5, Optional pheight As Single = 11, _
    '                    Optional lPAISAGEM As Boolean = False)
                        
    
    'ConvertToPDF mvarArquivo, cDESTINO
    
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
    If mvarDESTINO = "ARQ" Then
        Open mvarArquivo For Output As #mvarhandle
    End If
    If mvarDESTINO = "LPT1" Then
        Open "LPT1:" For Output As #mvarhandle
    End If
    If mvarDESTINO = "LPT2" Then
        Open "LPT2:" For Output As #mvarhandle
    End If
    If mvarDESTINO = "LPT3" Then
        Open "LPT3:" For Output As #mvarhandle
    End If
    If mvarDESTINO = "COM1" Then
        Open "COM1:" For Output As #mvarhandle
    End If
    If mvarDESTINO = "COM2" Then
        Open "COM2:" For Output As #mvarhandle
    End If
    If mvarDESTINO = "IMPNET" Then
        Open mvarcaminho For Output As #mvarhandle
    End If
End Sub

Public Sub novafolha()
    Dim X As Integer
    ReDim mvaraLINHAS(mvarnROW)
    For X = 0 To mvarnROW - 1
        mvaraLINHAS(X) = Space(mvarnCOL)
    Next X
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
    'Dim X As Integer
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

