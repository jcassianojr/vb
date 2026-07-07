VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form FrmPreview 
   Caption         =   "FrmPreview - cWebView2"
   ClientHeight    =   5124
   ClientLeft      =   60
   ClientTop       =   408
   ClientWidth     =   14316
   LinkTopic       =   "Form1"
   ScaleHeight     =   5124
   ScaleWidth      =   14316
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picHost 
      BorderStyle     =   0  'None
      HasDC           =   0   'False
      Height          =   4692
      Left            =   360
      ScaleHeight     =   4692
      ScaleWidth      =   10812
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   360
      Width           =   10812
   End
   Begin VB.CommandButton Chamamotor 
      Caption         =   "motor"
      Height          =   372
      Left            =   12480
      TabIndex        =   8
      Top             =   360
      Width           =   1212
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   432
      Left            =   12600
      TabIndex        =   0
      Top             =   960
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   762
      Caption         =   "Retornar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton cmdimp 
      Height          =   432
      Left            =   12600
      TabIndex        =   1
      Top             =   1560
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   762
      Caption         =   "Imprimir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton cmdSavehtml 
      Height          =   372
      Left            =   12600
      TabIndex        =   2
      Top             =   4200
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   656
      Caption         =   "Html"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton cmdSaveTXT 
      Height          =   372
      Left            =   12600
      TabIndex        =   3
      Top             =   3840
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   656
      Caption         =   "TXT"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton cmdSavePNG 
      Height          =   312
      Left            =   12600
      TabIndex        =   4
      Top             =   2280
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   550
      Caption         =   "PNG"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton cmdSavejpg 
      Height          =   312
      Left            =   12600
      TabIndex        =   5
      Top             =   2640
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   550
      Caption         =   "JPG"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton cmdsavedoc 
      Height          =   372
      Left            =   12600
      TabIndex        =   6
      Top             =   3000
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   656
      Caption         =   "DOC"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton CmdSavePDF 
      Height          =   372
      Left            =   12600
      TabIndex        =   7
      Top             =   3480
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   656
      Caption         =   "PDF"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mvarCaminhoArquivo As String
Dim bJaInicializado As Boolean

' 1) Instancia do novo motor Chromium cWebView2
Private WithEvents m_oWebView2 As cWebView2
Attribute m_oWebView2.VB_VarHelpID = -1

Private Sub Form_Load()
    Dim hResult
    bJaInicializado = False
    If cARQRTF = "" Then
        'Carqrtf = "c:\temp\uso.txt"
        'cARQRTF = "c:\temp\README.md"
        'Carqrtf = "c:\temp\totalcmd.chm"
    End If
    
    ' 2) Inicializa o cWebView2 e vincula ao PictureBox host (picHost)
    Set m_oWebView2 = New cWebView2
'    m_oWebView2.BindTo picHost.hwnd, , , , "--allow-run-as-system", , False, True, False
     hResult = m_oWebView2.BindTo(picHost.hwnd, , , , "--allow-run-as-system --allow-file-access-from-files --disable-web-security", , False, True, False)

End Sub

Private Sub Form_Activate()
    bJaInicializado = True
    chamamotor_click
End Sub

Private Sub Form_Unload(Cancel As Integer)
       Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim arTmp(4) As String
    arTmp(0) = App.Path & "\\~doc_engine.html"
    arTmp(1) = App.Path & "\\~xlsx_engine.html"
    arTmp(2) = App.Path & "\\~marked_engine.html"
    arTmp(3) = App.Path & "\\~zpl_view_engine.html"
    arTmp(4) = App.Path & "\\~rtf_view_engine.html"
    Dim i As Integer
    For i = 0 To UBound(arTmp)
        If fso.FileExists(arTmp(i)) Then
            On Error Resume Next
            fso.DeleteFile arTmp(i), True
        End If
    Next i
    Set fso = Nothing

    If Not m_oWebView2 Is Nothing Then
        m_oWebView2.Shutdown
        Set m_oWebView2 = Nothing
    End If
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If bJaInicializado Then
        ' Redimensiona o PictureBox base
        picHost.Height = Me.ScaleHeight - picHost.Top - 200
        picHost.Width = Me.ScaleWidth - picHost.Left - 1800
        
        ' 4) Sincroniza o tamanho do cWebView2 com o PictureBox
        If Not m_oWebView2 Is Nothing Then
            m_oWebView2.SyncSizeToHostWindow
        End If
    End If
End Sub

Private Sub cmdimp_Click()
    ' Comando nativo via JS para chamar a janela de impressao
    If Not m_oWebView2 Is Nothing Then
        m_oWebView2.ExecuteScript "window.print();"
    End If
End Sub

Private Sub cmdSavehtml_Click()
    Dim sFileName As String
    Dim sFILTER As String
    sFILTER = "Arquivos de Textos (*.HTML)" & vbNullChar & "*.HTML" & vbNullChar & "Todos Arquivos" & vbNullChar & "*.*"
    sFileName = FileSave(Me, sFILTER, 1, "HTML", , , "Salvar HTML Como")
    
    If sFileName <> "" Then
        FileWrite sFileName, m_oWebView2.jsProp("document.documentElement.outerHTML")
    End If
End Sub

Private Sub cmdSaveTXT_Click()
    Dim sFileName As String
    Dim sFILTER As String
    sFILTER = "Arquivos de Textos (*.TXT)" & vbNullChar & "*.TXT" & vbNullChar & "Todos Arquivos" & vbNullChar & "*.*"
    sFileName = FileSave(Me, sFILTER, 1, "TXT", , , "Salvar TXT Como")
    
    If sFileName <> "" Then
        FileWrite sFileName, m_oWebView2.jsProp("document.body.innerText")
    End If
End Sub

Private Sub cmdSavejpg_Click()
    Dim sFILTER As String
    Dim sFileName As String
    Dim sSugestaoNome As String
    sSugestaoNome = NomeArq(mvarCaminhoArquivo, True)
    sFILTER = "Imagem JPG (*.jpg;*.jpeg)|*.jpg;*.jpeg|Todos Arquivos (*.*)|*.*"
    sFileName = FileSave(Me, sFILTER, 1, "JPG", sSugestaoNome, , "Salvar Etiqueta Como Imagem JPG")
    
    If sFileName <> "" Then
       ' Call SalvarPreviewComoImagem(sFileName)
    End If
End Sub

Private Sub cmdSavePNG_Click()
    Dim sFILTER As String
    Dim sFileName As String
    Dim sSugestaoNome As String
    sSugestaoNome = NomeArq(mvarCaminhoArquivo, True)
    sFILTER = "Imagem PNG (*.png)|*.png|Todos Arquivos (*.*)|*.*"
    sFileName = FileSave(Me, sFILTER, 1, "PNG", sSugestaoNome, , "Salvar Etiqueta Como Imagem PNG")
    
    If sFileName <> "" Then
      '  Call SalvarPreviewComoImagem(sFileName)
    End If
End Sub

Private Sub CmdSavePDF_Click()
    Dim sFILTER As String
    Dim sFileName As String
    Dim sSugestaoNome As String
    
    On Error GoTo ErroSavePDF
    sSugestaoNome = NomeArq(mvarCaminhoArquivo, True)
    sFILTER = "Documento PDF (*.pdf)|*.pdf|Todos Arquivos (*.*)|*.*"
    sFileName = FileSave(Me, sFILTER, 1, "PDF", sSugestaoNome, , "Salvar Como PDF")
    
    If sFileName = "" Then Exit Sub
    
    ' Dispara a impressao do Chromium configurada para salvar em PDF
    If Not m_oWebView2 Is Nothing Then
        m_oWebView2.ExecuteScript "window.print();"
    End If
    Exit Sub
ErroSavePDF:
    MsgBox "Erro ao exportar o arquivo para PDF: " & Err.Description, vbCritical, "Erro de Exportacao"
End Sub

Private Sub cmdsavedoc_Click()
    Dim sFILTER As String
    Dim sFileName As String
    Dim sSugestaoNome As String
    sSugestaoNome = NomeArq(mvarCaminhoArquivo, True)
    sFILTER = "Documento do Word (*.doc)|*.doc|Documento PDF (*.pdf)|*.pdf|Todos Arquivos (*.*)|*.*"
    sFileName = FileSave(Me, sFILTER, 1, "DOC", sSugestaoNome, , "Salvar Documento Como")
    If sFileName = "" Then Exit Sub
    
    If InStr(1, sFileName, ".pdf", vbTextCompare) > 0 Then
        If Not m_oWebView2 Is Nothing Then
            m_oWebView2.ExecuteScript "window.print();"
        End If
    Else
        ' LOGICA DO DOC (Mantenha igual ao seu original caso exista logica complexa)
    End If
End Sub

Private Sub Encerrar_Click()
    Unload Me
End Sub

Private Sub chamamotor_click()
   On Error GoTo TrataErro
    Dim cEXT As String
    
    If m_oWebView2 Is Nothing Then
        MsgBox "Erro: O componente cWebView2 não foi inicializado!", vbCritical
        Exit Sub
    End If
    
    Me.SetFocus
    DoEvents
    
    mvarCaminhoArquivo = Trim(CStr(cARQRTF))
    cEXT = parsefile(mvarCaminhoArquivo, "E")
    
    If cEXT = "zpl" Then
        cmdSavehtml.Visible = False
        cmdSaveTXT.Visible = False
        cmdSavePNG.Visible = True
        cmdSavejpg.Visible = True
        cmdSavePNG.Top = cmdSavehtml.Top
        cmdSavejpg.Top = cmdSavePNG.Top + cmdSavePNG.Height + 120
        Call RenderizarMotorZplLocal
            
    ElseIf cEXT = "chm" Then
        cmdSavehtml.Visible = False
        cmdSaveTXT.Visible = False
        cmdSavePNG.Visible = False
        cmdSavejpg.Visible = False
        Me.Caption = "Manual de Ajuda - " & NomeArq(mvarCaminhoArquivo, False)
        m_oWebView2.Navigate "its:" & mvarCaminhoArquivo & "::/"
        
    ElseIf cEXT = "hlp" Then
        cmdSavehtml.Visible = False
        cmdSaveTXT.Visible = False
        cmdSavePNG.Visible = False
        cmdSavejpg.Visible = False
        If IsArquivoChmDisfarcado(mvarCaminhoArquivo) Then
            Me.Caption = "Manual de Ajuda - " & NomeArq(mvarCaminhoArquivo, False)
            m_oWebView2.Navigate "its:" & mvarCaminhoArquivo & "::/"
        Else
            Call RenderizarMotorHlpLocal
        End If
        
    ElseIf cEXT = "xlsx" Or cEXT = "xls" Or cEXT = "ods" Then
        cmdSavehtml.Visible = False
        cmdSaveTXT.Visible = False
        cmdSavePNG.Visible = False
        cmdSavejpg.Visible = False
        Me.Caption = "Visualizador de Planilhas - " & NomeArq(mvarCaminhoArquivo, False)
        Call RenderizarMotorPlanilhaLocal
        
    ElseIf cEXT = "csv" Then
        cmdSavehtml.Visible = False
        cmdSaveTXT.Visible = False
        cmdSavePNG.Visible = False
        cmdSavejpg.Visible = False
        Me.Caption = "Visualizador de Dados - " & NomeArq(mvarCaminhoArquivo, False)
        Call RenderizarMotorDelimitadoLocal(mvarCaminhoArquivo)
        
    ElseIf cEXT = "rtf" Then
        cmdSavehtml.Visible = False
        cmdSaveTXT.Visible = False
        cmdSavePNG.Visible = False
        cmdSavejpg.Visible = False
        Me.Caption = "Visualizador RTF - " & NomeArq(mvarCaminhoArquivo, False)
        Call RenderizarMotorRtfLocal
        
    ElseIf cEXT = "docx" Or cEXT = "doc" Or cEXT = "odt" Then
        cmdSavehtml.Visible = False
        cmdSaveTXT.Visible = False
        cmdSavePNG.Visible = False
        cmdSavejpg.Visible = False
        cmdsavedoc.Visible = False
        Me.Caption = "Visualizador de Documentos - " & NomeArq(mvarCaminhoArquivo, False)
        Call RenderizarMotorDocLocal
        
    ElseIf cEXT = "md" Or cEXT = "markdown" Then
        cmdSavehtml.Visible = False
        cmdSaveTXT.Visible = False
        cmdSavePNG.Visible = False
        cmdSavejpg.Visible = False
        Me.Caption = "Visualizador de Documentos - " & NomeArq(mvarCaminhoArquivo, False)
        Call RenderizarMotorMarked
        
    ElseIf cEXT = "pdf" Then
        cmdSavehtml.Visible = False
        cmdSaveTXT.Visible = False
        cmdSavePNG.Visible = False
        cmdSavejpg.Visible = False
        CmdSavePDF.Visible = False
        Me.Caption = "Leitor de PDF - " & NomeArq(mvarCaminhoArquivo, False)
        m_oWebView2.Navigate "file:///" & Replace(mvarCaminhoArquivo, "\\", "/")
         
    Else
        Me.Caption = "Visualizador - " & NomeArq(mvarCaminhoArquivo, False)
        cmdSavehtml.Visible = True
        cmdSaveTXT.Visible = True
        cmdSavePNG.Visible = False
        cmdSavejpg.Visible = False
        m_oWebView2.Navigate "file:///" & Replace(mvarCaminhoArquivo, "\\", "/")
    End If
    Exit Sub
TrataErro:
    MsgBox "Erro crítico ao inicializar os motores: " & Err.Description, vbCritical, "Erro de Interface"
End Sub
    

' ==================================================================
' ROTINA DE VALIDAÇÃO DE CONECTIVIDADE DA URL (TIMEOUT ULTRA-CURTO 800ms)
' ==================================================================
Private Function TestarLinkMotorJS(ByVal sURL As String) As Boolean
    On Error GoTo ErroLink
    Dim oHttp As Object
    
    Set oHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    oHttp.setTimeouts 800, 800, 800, 800 ' Evita travamento da tela se o cliente estiver sem rede
    
    oHttp.Open "HEAD", sURL, False
    oHttp.Send
    
    If oHttp.Status = 200 Then TestarLinkMotorJS = True Else TestarLinkMotorJS = False
    Set oHttp = Nothing
    Exit Function

ErroLink:
    TestarLinkMotorJS = False
    Set oHttp = Nothing
End Function

Private Sub RenderizarMotorMarked()
    Dim fso As Object, streamOut As Object
    Dim cHtmlTempPath As String, cJsLocal As String
    Dim cConteudoMD As String
    Dim cConteudoTratado As String
    
    cJsLocal = App.Path & "\WebResources\marked.min.js"
    cHtmlTempPath = App.Path & "\~marked_engine.html"
    
    ' 1. Lê o conteúdo do arquivo
    cConteudoMD = LerArquivoTexto(mvarCaminhoArquivo)
    
    ' 2. TRATAMENTO DO CONTEÚDO (CRÍTICO)
    ' Substituímos os caracteres que quebrariam a string do JavaScript
    cConteudoTratado = Replace(cConteudoMD, "\", "\\")      ' Escapa barras invertidas
    cConteudoTratado = Replace(cConteudoTratado, "`", "\`")  ' Escapa crases (fim da template literal)
    cConteudoTratado = Replace(cConteudoTratado, "'", "\'")  ' Escapa aspas simples
    cConteudoTratado = Replace(cConteudoTratado, """", "\""") ' Escapa aspas duplas
    cConteudoTratado = Replace(cConteudoTratado, vbCrLf, "\n") ' Normaliza quebras de linha
    cConteudoTratado = Replace(cConteudoTratado, vbCr, "\n")
    cConteudoTratado = Replace(cConteudoTratado, vbLf, "\n")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set streamOut = fso.OpenTextFile(cHtmlTempPath, 2, True)
    
    ' 3. Geração do HTML temporário
    streamOut.WriteLine "<html><head><meta charset='utf-8'>"
    streamOut.WriteLine "<script src='file:///" & Replace(cJsLocal, "\", "/") & "'></script>"
    streamOut.WriteLine "<style>body{font-family:sans-serif; padding:20px; word-wrap:break-word;}</style></head>"
    streamOut.WriteLine "<body><div id='content'></div>"
    streamOut.WriteLine "<script>"
    
    ' Aqui injetamos o conteúdo já tratado
    streamOut.WriteLine "document.getElementById('content').innerHTML = marked.parse(`" & cConteudoTratado & "`);"
    
    streamOut.WriteLine "</script></body></html>"
    streamOut.Close
    
    ' 4. Navegação segura
    m_oWebView2.Navigate "file:///" & Replace(cHtmlTempPath, "\", "/")
    
    Set fso = Nothing
End Sub

' Função auxiliar para leitura simples
Private Function LerArquivoTexto(ByVal cCaminho As String) As String
    Dim fNum As Integer
    fNum = FreeFile
    Open cCaminho For Input As #fNum
    LerArquivoTexto = Input$(LOF(fNum), fNum)
    Close #fNum
End Function

Private Sub RenderizarMotorZplLocal()
    Dim fso As Object, streamOut As Object
    Dim cHtmlTempPath As String, cJsLocal As String
    Dim cConteudoZpl As String
    
    'On Error GoTo ErroHandler
    On Error Resume Next
    
    cJsLocal = App.Path & "\WebResources\bwip-js-min.js"
    cHtmlTempPath = App.Path & "\~zpl_view_engine.html"
    
    ' 1. LER O CONTEÚDO DO ARQUIVO ZPL
    cConteudoZpl = LerArquivoTexto(mvarCaminhoArquivo)
    
    ' Tratamento para evitar quebras de string no JS
    cConteudoZpl = Replace(cConteudoZpl, "\", "\\")
    cConteudoZpl = Replace(cConteudoZpl, vbCrLf, "\n")
    cConteudoZpl = Replace(cConteudoZpl, "'", "\'")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set streamOut = fso.OpenTextFile(cHtmlTempPath, 2, True)
    
    ' 2. ESTRUTURA HTML COM INJEÇÃO DO ZPL
    streamOut.WriteLine "<html><head><meta charset='utf-8'></head>"
    streamOut.WriteLine "<body>"
    streamOut.WriteLine "<canvas id='zplCanvas'></canvas>"
    streamOut.WriteLine "<script src='file:///" & Replace(cJsLocal, "\", "/") & "'></script>"
    streamOut.WriteLine "<script>"
    ' Injeção do conteúdo e comando de renderização
    streamOut.WriteLine "try {"
    streamOut.WriteLine "  bwipjs.toCanvas('zplCanvas', {"
    streamOut.WriteLine "    bcid: 'code128', "
    streamOut.WriteLine "    text: '" & cConteudoZpl & "'"
    streamOut.WriteLine "  });"
    streamOut.WriteLine "} catch(e) { document.body.innerHTML = 'Erro JS: ' + e; }"
    streamOut.WriteLine "</script>"
    streamOut.WriteLine "</body></html>"
    streamOut.Close
    
    ' 3. NAVEGAÇÃO
    m_oWebView2.Navigate "file:///" & Replace(cHtmlTempPath, "\", "/")
    
    Set fso = Nothing
    Exit Sub

'ErroHandler:
'    MsgBox "Erro ao renderizar motor: " & Err.Description, vbCritical
End Sub

Private Sub RenderizarMotorPlanilhaLocal()
    Dim fso As Object, streamOut As Object
    Dim cHtmlTempPath As String, cJsLocal As String
    Dim cCaminhoPlanilhaFormatado As String
    
    cJsLocal = App.Path & "\\WebResources\\xlsx.full.min.js"
    cHtmlTempPath = App.Path & "\\~xlsx_engine.html"
    cCaminhoPlanilhaFormatado = Replace(mvarCaminhoArquivo, "\\", "/")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set streamOut = fso.OpenTextFile(cHtmlTempPath, 2, True)
    
    streamOut.WriteLine "<!DOCTYPE html><html><head><meta charset='utf-8'>"
    streamOut.WriteLine "<style>body{font-family:'Segoe UI',Arial,sans-serif;margin:20px;background:#f4f4f9;}"
    streamOut.WriteLine "#sheet-container{background:#fff;padding:20px;box-shadow:0 4px 15px rgba(0,0,0,0.1);border-radius:4px;overflow-x:auto;}"
    streamOut.WriteLine "table{border-collapse:collapse;width:100%;margin:10px 0;font-size:14px;}"
    streamOut.WriteLine "th{background-color:#107c41;color:white;font-weight:bold;padding:10px;border:1px solid #ddd;}"
    streamOut.WriteLine "td{padding:8px;border:1px solid #ddd;text-align:left;}tr:nth-child(even){background-color:#f9f9f9;}</style>"
    streamOut.WriteLine "<script src='file:///" & Replace(cJsLocal, "\\", "/") & "'></script></head>"
    streamOut.WriteLine "<body><div id='sheet-container'>Processando dados da planilha local...</div>"
    streamOut.WriteLine "<script>"
    streamOut.WriteLine "fetch('file:///" & cCaminhoPlanilhaFormatado & "')"
    streamOut.WriteLine "  .then(res => res.arrayBuffer())"
    streamOut.WriteLine "  .then(ab => {" 'Office
    streamOut.WriteLine "    let workbook = XLSX.read(new Uint8Array(ab), {type: 'array'});"
    streamOut.WriteLine "    let firstSheetName = workbook.SheetNames[0];"
    streamOut.WriteLine "    let worksheet = workbook.Sheets[firstSheetName];"
    streamOut.WriteLine "    document.getElementById('sheet-container').innerHTML = XLSX.utils.sheet_to_html(worksheet);"
    streamOut.WriteLine "  }).catch(err => { document.getElementById('sheet-container').innerHTML = '<h2>Erro ao processar a planilha.</h2><p>' + err.message + '</p>'; });"
    streamOut.WriteLine "</script></body></html>"
    streamOut.Close
    
    m_oWebView2.Navigate "file:///" & Replace(cHtmlTempPath, "\\", "/")
    Set fso = Nothing
End Sub
Private Sub RenderizarMotorRtfLocal()
    Dim fso As Object, streamOut As Object
    Dim cHtmlTempPath As String
    Dim cCaminhoRtfFormatado As String
    
    cHtmlTempPath = App.Path & "\~rtf_view_engine.html"
    cCaminhoRtfFormatado = Replace(mvarCaminhoArquivo, "\", "/")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set streamOut = fso.OpenTextFile(cHtmlTempPath, 2, True)
    
    streamOut.WriteLine "<!DOCTYPE html><html><head><meta charset='utf-8'>"
    streamOut.WriteLine "<style>body{font-family:'Segoe UI',Arial,sans-serif;margin:30px;background-color:#f4f4f9;color:#333;}"
    streamOut.WriteLine "#rtf-container{max-width:800px;margin:0 auto;background:#fff;padding:40px;box-shadow:0 4px 15px rgba(0,0,0,0.1);border-radius:4px;min-height:600px;white-space:pre-wrap;}</style></head>"
    streamOut.WriteLine "<body><div id='rtf-container'>Processando documento...</div>"
    streamOut.WriteLine "<script>"
    streamOut.WriteLine "fetch('file:///" & cCaminhoRtfFormatado & "').then(res => res.text()).then(rtfText => {"
    ' Regex para limpeza básica de RTF: remove chaves, comandos e formatação técnica
    streamOut.WriteLine "  let clean = rtfText.replace(/\\{[^}]*\\}/g, '').replace(/\\[a-z0-9]+/ig, '').replace(/[\r\n]+/g, '<br>');"
    streamOut.WriteLine "  document.getElementById('rtf-container').innerHTML = clean;"
    streamOut.WriteLine "}).catch(err => { document.getElementById('rtf-container').innerHTML = '<h2>Erro ao processar RTF.</h2><p>' + err.message + '</p>'; });"
    streamOut.WriteLine "</script></body></html>"
    streamOut.Close
    
    m_oWebView2.Navigate "file:///" & Replace(cHtmlTempPath, "\", "/")
    Set fso = Nothing
End Sub
Private Sub RenderizarMotorRtfLocalold()
    Dim fso As Object, streamOut As Object
    Dim cHtmlTempPath As String, cJsLocal As String
    Dim cCaminhoRtfFormatado As String
    
    cJsLocal = App.Path & "\\WebResources\\rtf-parser.min.js"
    cHtmlTempPath = App.Path & "\\~rtf_view_engine.html"
    cCaminhoRtfFormatado = Replace(mvarCaminhoArquivo, "\\", "/")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set streamOut = fso.OpenTextFile(cHtmlTempPath, 2, True)
    
    streamOut.WriteLine "<!DOCTYPE html><html><head><meta charset='utf-8'>"
    streamOut.WriteLine "<style>body{font-family:'Segoe UI',Arial,sans-serif;margin:30px;background-color:#f4f4f9;color:#333;}"
    streamOut.WriteLine "#rtf-container{max-width:800px;margin:0 auto;background:#fff;padding:40px;box-shadow:0 4px 15px rgba(0,0,0,0.1);border-radius:4px;min-height:600px;white-space:pre-wrap;}</style>"
    streamOut.WriteLine "<script src='file:///" & Replace(cJsLocal, "\\", "/") & "'></script></head>"
    streamOut.WriteLine "<body><div id='rtf-container'>Processando documento formatado RTF...</div>"
    streamOut.WriteLine "<script>"
    streamOut.WriteLine "fetch('file:///" & cCaminhoRtfFormatado & "')"
    streamOut.WriteLine "  .then(res => res.text())"
    streamOut.WriteLine "  .then(rtfText => {"
    streamOut.WriteLine "    let htmlResult = rtfText.replace(/\\\\par/g, '<br>').replace(/\\\\b\\s+(.*?)\\\\b0/g, '<b>$1</b>').replace(/\\\\i\\s+(.*?)\\\\i0/g, '<i>$1</i>').replace(/\\\\{[^{}]*\\\\}/g, '');" 'office
    streamOut.WriteLine "    htmlResult = htmlResult.replace(/\\\\[a-z0-9-]+/g, '').replace(/[{}]/g, '');"
    streamOut.WriteLine "    document.getElementById('rtf-container').innerHTML = htmlResult.trim();"
    streamOut.WriteLine "  }).catch(err => { document.getElementById('rtf-container').innerHTML = '<h2>Erro ao processar RTF.</h2><p>' + err.message + '</p>'; });"
    streamOut.WriteLine "</script></body></html>"
    streamOut.Close
    
    m_oWebView2.Navigate "file:///" & Replace(cHtmlTempPath, "\\", "/")
    Set fso = Nothing
End Sub

Private Sub RenderizarMotorDelimitadoLocal(ByVal cCaminhoTxt As String)
    Dim fso As Object, streamOut As Object
    Dim cHtmlTempPath As String, cJsLocal As String
    Dim sLinhas() As String, sPrimeiraLinha As String, sDelimitador As String
    Dim lContador As Long, lMaxRepeticoes As Long, nMaxColunas As Long, i As Long, J As Long
    Dim cConteudoBruto As String, cConteudoMD As String, cLinhaSeparador As String
    
    cJsLocal = App.Path & "\\WebResources\\marked.min.js"
    cHtmlTempPath = App.Path & "\\~marked_engine.html"
    
    cConteudoBruto = LerArquivoTexto(cCaminhoTxt)
    sLinhas = Split(cConteudoBruto, vbCrLf)
    If UBound(sLinhas) < 0 Then Exit Sub
    sPrimeiraLinha = sLinhas(0)
    
    sDelimitador = "|"
    lMaxRepeticoes = Len(sPrimeiraLinha) - Len(Replace(sPrimeiraLinha, "|", ""))
    
    lContador = Len(sPrimeiraLinha) - Len(Replace(sPrimeiraLinha, ";", ""))
    If lContador > lMaxRepeticoes Then lMaxRepeticoes = lContador: sDelimitador = ";"
    lContador = Len(sPrimeiraLinha) - Len(Replace(sPrimeiraLinha, ",", ""))
    If lContador > lMaxRepeticoes Then lMaxRepeticoes = lContador: sDelimitador = ","
    lContador = Len(sPrimeiraLinha) - Len(Replace(sPrimeiraLinha, vbTab, ""))
    If lContador > lMaxRepeticoes Then lMaxRepeticoes = lContador: sDelimitador = vbTab

    Dim vColunas As Variant
    vColunas = Split(sPrimeiraLinha, sDelimitador)
    nMaxColunas = UBound(vColunas)
    
    cConteudoMD = "|"
    For J = 0 To nMaxColunas
        cConteudoMD = cConteudoMD & " Coluna " & (J + 1) & " |"
    Next J
    cConteudoMD = cConteudoMD & vbCrLf
    
    cLinhaSeparador = "|"
    For J = 0 To nMaxColunas
        cLinhaSeparador = cLinhaSeparador & "---| "
    Next J
    cConteudoMD = cConteudoMD & cLinhaSeparador & vbCrLf
    
    For i = 0 To UBound(sLinhas)
        If Trim(sLinhas(i)) <> "" Then
            vColunas = Split(sLinhas(i), sDelimitador)
            cConteudoMD = cConteudoMD & "|"
            For J = 0 To nMaxColunas
                If J <= UBound(vColunas) Then
                    Dim sDadoLimpo As String
                    sDadoLimpo = Trim(vColunas(J))
                    If Left(sDadoLimpo, 1) = """" And Right(sDadoLimpo, 1) = """" Then
                        sDadoLimpo = Mid(sDadoLimpo, 2, Len(sDadoLimpo) - 2)
                    End If
                    cConteudoMD = cConteudoMD & " " & sDadoLimpo & " |"
                Else
                    cConteudoMD = cConteudoMD & "  |"
                End If
            Next J
            cConteudoMD = cConteudoMD & vbCrLf
        End If
    Next i
    
    cConteudoMD = Replace(cConteudoMD, "\\", "\\\\")
    cConteudoMD = Replace(cConteudoMD, "`", "\\`")
    cConteudoMD = Replace(cConteudoMD, "'", "\\'")
    cConteudoMD = Replace(cConteudoMD, vbCrLf, "\\n")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set streamOut = fso.OpenTextFile(cHtmlTempPath, 2, True)
    
    streamOut.WriteLine "<html><head><meta charset='utf-8'>"
    streamOut.WriteLine "<script src='file:///" & Replace(cJsLocal, "\\", "/") & "'></script>"
    streamOut.WriteLine "<style>body{font-family:sans-serif;padding:35px;background:#fff;} table{border-collapse:collapse;width:100%;} th,td{border:1px solid #d0d7de;padding:8px 13px;} th{background:#f6f8fa;font-weight:600;}</style></head>"
    streamOut.WriteLine "<body><div id='content'></div><script>"
    streamOut.WriteLine "document.getElementById('content').innerHTML = marked.parse(`" & cConteudoMD & "`);"
    streamOut.WriteLine "</script></body></html>"
    streamOut.Close
    
    m_oWebView2.Navigate "file:///" & Replace(cHtmlTempPath, "\\", "/")
    Set fso = Nothing
End Sub

Private Sub RenderizarMotorDocLocal()
    Dim fso As Object, streamOut As Object
    Dim cHtmlTempPath As String, cJsLocal As String
    Dim cCaminhoDocFormatado As String
    
    cJsLocal = App.Path & "\\WebResources\\mammoth.browser.min.js"
    cHtmlTempPath = App.Path & "\\~doc_engine.html"
    cCaminhoDocFormatado = Replace(mvarCaminhoArquivo, "\\", "/")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set streamOut = fso.OpenTextFile(cHtmlTempPath, 2, True)
    
    streamOut.WriteLine "<!DOCTYPE html><html><head><meta charset='utf-8'>"
    streamOut.WriteLine "<style>body{font-family:'Segoe UI',Arial,sans-serif;margin:30px;background:#f4f4f9;}"
    streamOut.WriteLine "#document-container{max-width:800px;margin:0 auto;background:#fff;padding:40px;box-shadow:0 4px 15px rgba(0,0,0,0.1);border-radius:4px;}"
    streamOut.WriteLine "table{border-collapse:collapse;width:100%;}th,td{border:1px solid #ddd;padding:8px;}img{max-width:100%;height:auto;}</style>"
    streamOut.WriteLine "<script src='file:///" & Replace(cJsLocal, "\\", "/") & "'></script></head>"
    streamOut.WriteLine "<body><div id='document-container'>Carregando documento local...</div>"
    streamOut.WriteLine "<script>"
    streamOut.WriteLine "fetch('file:///" & cCaminhoDocFormatado & "')"
    streamOut.WriteLine "  .then(res => res.arrayBuffer())"
    streamOut.WriteLine "  .then(ab => mammoth.convertToHtml({arrayBuffer: ab}))"
    streamOut.WriteLine "  .then(result => { document.getElementById('document-container').innerHTML = result.value; })"
    streamOut.WriteLine "  .catch(err => { document.getElementById('document-container').innerHTML = '<h2>Erro ao carregar o documento.</h2><p>' + err.message + '</p>'; });"
    streamOut.WriteLine "</script></body></html>"
    streamOut.Close
    
    m_oWebView2.Navigate "file:///" & Replace(cHtmlTempPath, "\\", "/")
    Set fso = Nothing
End Sub

Private Sub RenderizarMotorHlpLocal()
    Dim cHTMLAviso As String
    cHTMLAviso = "<html><head><meta charset='utf-8'>" & _
                 "<style>body{font-family:sans-serif;padding:50px;text-align:center;background:#fafafa;}" & _
                 ".card{background:#fff;padding:30px;border-radius:8px;box-shadow:0 4px 12px rgba(0,0,0,0.1);max-width:500px;margin:0 auto;border-top:4px solid #e67e22;}" & _
                 "h3{color:#d35400;margin-top:0;} p{font-size:14px;line-height:1.5;color:#555;}" & _
                 ".file-name{font-family:monospace;background:#f1f2f6;padding:4px 8px;border-radius:4px;color:#c0392b;}</style></head><body>" & _
                 "  <div class='card'>" & _
                 "    <h3>Formato de Ajuda Incompatível</h3>" & _
                 "    <p>O ficheiro <span class='file-name'>" & NomeArq(mvarCaminhoArquivo, True) & "</span> está num formato legado (.hlp) cujo motor de renderização foi descontinuado pela Microsoft no Windows 10/11.</p>" & _
                 "    <p>Recomenda-se a conversão deste manual para formatos suportados como <b>.CHM</b> ou <b>.PDF</b> para exibição integrada.</p>" & _
                 "  </div>" & _
                 "</body></html>"
                 
    m_oWebView2.NavigateToString cHTMLAviso
End Sub

Private Sub SalvarPreviewComoImagem(ByVal cCaminhoDestino As String)
    Dim cNomeArquivo As String, cEXTENSAO As String, cMimeType As String, cScript As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    cNomeArquivo = fso.GetFileName(cCaminhoDestino)
    cEXTENSAO = LCase(fso.GetExtensionName(cCaminhoDestino))
    Set fso = Nothing
    
    If cEXTENSAO = "jpg" Or cEXTENSAO = "jpeg" Then
        cMimeType = "image/jpeg"
    Else
        cMimeType = "image/png"
    End If
    
    cScript = "(function() {" & _
              "  var canvas = document.getElementById('zplCanvas');" & _
              "  if (canvas) {" & _
              "    var finalDataUrl;" & _
              "    if ('" & cMimeType & "' === 'image/jpeg') {" & _
              "      var tempCanvas = document.createElement('canvas');" & _
              "      tempCanvas.width = canvas.width;" & _
              "      tempCanvas.height = canvas.height;" & _
              "      var ctx = tempCanvas.getContext('2d');" & _
              "      ctx.fillStyle = '#FFFFFF';" & _
              "      ctx.fillRect(0, 0, tempCanvas.width, tempCanvas.height);" & _
              "      ctx.drawImage(canvas, 0, 0);" & _
              "      finalDataUrl = tempCanvas.toDataURL('image/jpeg', 0.9);" & _
              "    } else {" & _
              "      finalDataUrl = canvas.toDataURL('image/png');" & _
              "    }" & _
              "    var lnk = document.createElement('a');" & _
              "    lnk.download = '" & cNomeArquivo & "';" & _
              "    lnk.href = finalDataUrl;" & _
              "    document.body.appendChild(lnk);" & _
              "    lnk.click();" & _
              "    document.body.removeChild(lnk);" & _
              "  }" & _
              "})();"
              
    m_oWebView2.ExecuteScript cScript
End Sub

Private Function IsArquivoChmDisfarcado(ByVal cCaminho As String) As Boolean
    On Error GoTo Sair
    Dim fNum As Integer
    Dim sHeader As String * 4
    If Dir(cCaminho) <> "" Then
        fNum = FreeFile
        Open cCaminho For Binary As #fNum
            Get #fNum, 1, sHeader
        Close #fNum
        If sHeader = "ITSF" Then IsArquivoChmDisfarcado = True
    End If
    Exit Function
Sair:
    If fNum > 0 Then Close #fNum
End Function

