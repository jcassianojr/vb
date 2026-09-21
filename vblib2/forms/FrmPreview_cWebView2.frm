VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form FrmPreview 
   Caption         =   "FrmPreview - cWebView2"
   ClientHeight    =   5568
   ClientLeft      =   60
   ClientTop       =   408
   ClientWidth     =   14316
   LinkTopic       =   "Form1"
   ScaleHeight     =   5568
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
      Top             =   3960
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
      Top             =   3600
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
      Top             =   2160
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
      Top             =   2520
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
      Top             =   2880
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
      Top             =   3240
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
   Begin XPControls.XPButton CmdSavejson 
      Height          =   372
      Left            =   12600
      TabIndex        =   10
      Top             =   4320
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   656
      Caption         =   "json"
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
   Begin XPControls.XPButton CmdSavemd 
      Height          =   372
      Left            =   12600
      TabIndex        =   11
      Top             =   4680
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   656
      Caption         =   "md"
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
'cWEBIEW2 https://github.com/wqweto/cWebView2

Private mvarCaminhoArquivo As String
Private m_bWebViewInicializado As Boolean
Private m_bPreviewCarregado As Boolean
Private m_bEncerrando As Boolean
Private m_cArquivosTemporarios As Collection
Private m_sPastaTemporaria As String
Private m_sPastaChm As String

' 1) Instancia do novo motor Chromium cWebView2
Private WithEvents m_oWebView2 As cWebView2
Attribute m_oWebView2.VB_VarHelpID = -1

Private Sub Form_Load()
    Dim lResultado As Long
    m_bWebViewInicializado = False
    m_bPreviewCarregado = False
    m_bEncerrando = False
    Set m_cArquivosTemporarios = New Collection
    
    Set m_oWebView2 = New cWebView2
    lResultado = m_oWebView2.BindTo(picHost.hWnd, , , , vbNullString, , False, True, False)
    If lResultado = 0 Then
        MsgBox "Nao foi possivel inicializar o WebView2. Verifique se o Runtime do WebView2 esta instalado.", vbCritical, "Preview"
        Exit Sub
    End If
    m_bWebViewInicializado = True
End Sub

Private Sub Form_Activate()
    If m_bWebViewInicializado And Not m_bPreviewCarregado Then
        CarregarPreview
    End If
End Sub

Private Sub m_oWebView2_InitComplete()
    If m_bWebViewInicializado And Not m_bPreviewCarregado Then
        CarregarPreview
    End If
End Sub

Private Sub m_oWebView2_NavigationCompleted(ByVal IsSuccess As Boolean, ByVal WebErrorStatus As Long)
    If Not IsSuccess And Not m_bEncerrando Then
        MsgBox "Falha ao carregar o preview. Codigo WebView2: " & CStr(WebErrorStatus), vbExclamation, "Preview"
    End If
End Sub

Private Sub m_oWebView2_ProcessFailed()
    If Not m_bEncerrando Then
        MsgBox "O processo do WebView2 foi encerrado inesperadamente.", vbCritical, "Preview"
    End If
End Sub

Private Sub m_oWebView2_NewWindowRequested(ByVal IsUserInitiated As Boolean, IsHandled As Boolean, ByVal URI As String, NewWindowFeatures As Collection)
    IsHandled = True
End Sub
Private Sub cmdSavejson_Click()
    Dim sFileName As String
    Dim sFilter As String
    Dim sJson As String
    sFilter = "Arquivos de Textos (*.TXT)" & vbNullChar & "*.TXT" & vbNullChar & "Todos Arquivos" & vbNullChar & "*.*"
    sFileName = FileSave(Me, sFilter, 1, "json", , , "Salvar json Como")
    
    If sFileName <> "" Then
        sJson = m_oWebView2.jsProp("JSON.stringify(Array.from(document.querySelectorAll('tr')).map(r => Array.from(r.querySelectorAll('th,td')).map(c => c.innerText)))")
        FileWrite sFileName, sJson
    End If
End Sub
Private Sub cmdSaveMD_Click()
    Dim sFileName As String
    Dim sFilter As String
    Dim sMarkdown As String
    sFilter = "Arquivos de Textos (*.TXT)" & vbNullChar & "*.TXT" & vbNullChar & "Todos Arquivos" & vbNullChar & "*.*"
    sFileName = FileSave(Me, sFilter, 1, "md", , , "Salvar md Como")
    If sFileName <> "" Then
        sMarkdown = m_oWebView2.jsProp("window.turndownService ? window.turndownService.turndown(document.body.innerHTML) : ''")
    End If
    If sFileName <> "" And Len(sMarkdown) > 0 Then
        FileWrite sFileName, sMarkdown
    ElseIf sFileName <> "" Then
        MsgBox "A exportacao Markdown requer o servico Turndown no WebView2.", vbExclamation, "Preview"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_bEncerrando = True

    If Not m_oWebView2 Is Nothing Then
        m_oWebView2.Shutdown
        Set m_oWebView2 = Nothing
    End If
    LimparArquivosTemporarios
    
End Sub

Private Sub Form_Resize()
    On Error GoTo TrataErro
    If m_bWebViewInicializado Then
        ' Redimensiona o PictureBox base
        picHost.Height = Me.ScaleHeight - picHost.Top - 200
        picHost.Width = Me.ScaleWidth - picHost.Left - 1800
        
        ' 4) Sincroniza o tamanho do cWebView2 com o PictureBox
        If Not m_oWebView2 Is Nothing Then
            m_oWebView2.SyncSizeToHostWindow
        End If
    End If
    Exit Sub
TrataErro:
    Debug.Print "FrmPreview.Form_Resize: " & Err.Description
End Sub

Private Sub cmdimp_Click()
    ' Comando nativo via JS para chamar a janela de impressao
    If Not m_oWebView2 Is Nothing Then
        m_oWebView2.ExecuteScript "window.print();"
    End If
End Sub

Private Sub cmdSavehtml_Click()
    Dim sFileName As String
    Dim sFilter As String
    sFilter = "Arquivos de Textos (*.HTML)" & vbNullChar & "*.HTML" & vbNullChar & "Todos Arquivos" & vbNullChar & "*.*"
    sFileName = FileSave(Me, sFilter, 1, "HTML", , , "Salvar HTML Como")
    
    If sFileName <> "" Then
        FileWrite sFileName, m_oWebView2.jsProp("document.documentElement.outerHTML")
    End If
End Sub

Private Sub cmdSaveTXT_Click()
    Dim sFileName As String
    Dim sFilter As String
    sFilter = "Arquivos de Textos (*.TXT)" & vbNullChar & "*.TXT" & vbNullChar & "Todos Arquivos" & vbNullChar & "*.*"
    sFileName = FileSave(Me, sFilter, 1, "TXT", , , "Salvar TXT Como")
    
    If sFileName <> "" Then
        FileWrite sFileName, m_oWebView2.jsProp("document.body.innerText")
    End If
End Sub

Private Sub cmdSavejpg_Click()
    Dim sFilter As String
    Dim sFileName As String
    Dim sSugestaoNome As String
    Dim sBase64JPG As String
    
    sSugestaoNome = NomeArq(mvarCaminhoArquivo, True)
    sFilter = "Imagem JPG (*.jpg;*.jpeg)|*.jpg;*.jpeg|Todos Arquivos (*.*)|*.*"
    sFileName = FileSave(Me, sFilter, 1, "JPG", sSugestaoNome, , "Salvar Como Imagem JPG")
    
    If sFileName <> "" Then
        sBase64JPG = m_oWebView2.jsProp("document.getElementById('zplCanvas').toDataURL('image/jpeg', 0.9)")
        Call SalvarBase64ComoImagem(sBase64JPG, sFileName)
    End If
End Sub

Private Sub cmdSavePNG_Click()
    Dim sFilter As String
    Dim sFileName As String
    Dim sSugestaoNome As String
    Dim sBase64PNG As String
    
    sSugestaoNome = NomeArq(mvarCaminhoArquivo, True)
    sFilter = "Imagem PNG (*.png)|*.png|Todos Arquivos (*.*)|*.*"
    sFileName = FileSave(Me, sFilter, 1, "PNG", sSugestaoNome, , "Salvar Como Imagem PNG")
    
    If sFileName <> "" Then
        sBase64PNG = m_oWebView2.jsProp("document.getElementById('zplCanvas').toDataURL('image/png')")
        Call SalvarBase64ComoImagem(sBase64PNG, sFileName)
    End If
End Sub

Private Sub CmdSavePDF_Click()
    Dim sFilter As String
    Dim sFileName As String
    Dim sSugestaoNome As String
    Dim vResult         As Variant

    
    On Error GoTo ErroSavePDF
    sSugestaoNome = NomeArq(mvarCaminhoArquivo, True)
    sFilter = "Documento PDF (*.pdf)|*.pdf|Todos Arquivos (*.*)|*.*"
    sFileName = FileSave(Me, sFilter, 1, "PDF", sSugestaoNome, , "Salvar Como PDF")
    
    If sFileName = "" Then Exit Sub
    
    vResult = m_oWebView2.PrintToPdf(sFileName)
    If vResult = True Then
     ' lblStatus.Caption = "Saved " & sFileName & " (" & FileLen(sFileName) & " bytes)"
    Else
      '  lblStatus.Caption = "PrintToPdf failed"
    End If
    Exit Sub


ErroSavePDF:
    MsgBox "Erro ao exportar o arquivo para PDF: " & Err.Description, vbCritical, "Erro de Exportacao"
End Sub

Private Sub cmdsavedoc_Click()
    Dim sFilter As String
    Dim sFileName As String
    Dim sSugestaoNome As String
    sSugestaoNome = NomeArq(mvarCaminhoArquivo, True)
    sFilter = "Documento do Word (*.doc)|*.doc|Documento PDF (*.pdf)|*.pdf|Todos Arquivos (*.*)|*.*"
    sFileName = FileSave(Me, sFilter, 1, "DOC", sSugestaoNome, , "Salvar Documento Como")
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

Private Sub Chamamotor_Click()
    If m_bWebViewInicializado Then CarregarPreview
End Sub

Private Sub CarregarPreview()
   On Error GoTo TrataErro
    Dim cEXT As String
    
    If m_oWebView2 Is Nothing Then
        MsgBox "O componente cWebView2 nao foi inicializado.", vbCritical, "Preview"
        Exit Sub
    End If
    
    Me.SetFocus
    DoEvents
    
    mvarCaminhoArquivo = Trim(CStr(cARQRTF))
    If Len(mvarCaminhoArquivo) = 0 Or Dir$(mvarCaminhoArquivo) = vbNullString Then
        MsgBox "O arquivo de preview nao foi encontrado.", vbExclamation, "Preview"
        Exit Sub
    End If
    cEXT = LCase(parsefile(mvarCaminhoArquivo, "E"))
    ConfigurarBotoesPreview cEXT
    m_bPreviewCarregado = True
    
    If cEXT = "zpl" Then
        cmdSavePNG.Visible = True
        cmdSavejpg.Visible = True
        cmdSavePNG.Top = cmdSavehtml.Top
        cmdSavejpg.Top = cmdSavePNG.Top + cmdSavePNG.Height + 120
        Call RenderizarMotorZplLocal
            
      ElseIf cEXT = "chm" Then
        Me.Caption = "Manual de Ajuda - " & NomeArq(mvarCaminhoArquivo, False)
        Call RenderizarMotorChmLocal(mvarCaminhoArquivo)
        'm_oWebView2.Navigate "its:" & mvarCaminhoArquivo & "::/"
        
    ElseIf cEXT = "hlp" Then
       Me.Caption = "Manual de Ajuda - " & NomeArq(mvarCaminhoArquivo, False)
        If IsArquivoChmDisfarcado(mvarCaminhoArquivo) Then
           Call RenderizarMotorChmLocal(mvarCaminhoArquivo)
        Else
           Call RenderizarMotorHlpLocal
        End If
      
    ElseIf cEXT = "xlsx" Or cEXT = "xls" Or cEXT = "ods" Then
        Me.Caption = "Visualizador de Planilhas - " & NomeArq(mvarCaminhoArquivo, False)
        Call RenderizarMotorPlanilhaLocal
        
    ElseIf cEXT = "csv" Then
        Me.Caption = "Visualizador de Dados - " & NomeArq(mvarCaminhoArquivo, False)
        Call RenderizarMotorDelimitadoLocal(mvarCaminhoArquivo)
        
    ElseIf cEXT = "rtf" Then
        Me.Caption = "Visualizador RTF - " & NomeArq(mvarCaminhoArquivo, False)
        Call RenderizarMotorRtfLocal
        
    ElseIf cEXT = "docx" Then
         Me.Caption = "Visualizador de Documentos - " & NomeArq(mvarCaminhoArquivo, False)
        Call RenderizarMotorDocLocal
        
    ElseIf cEXT = "md" Or cEXT = "markdown" Then
        Me.Caption = "Visualizador de Documentos - " & NomeArq(mvarCaminhoArquivo, False)
        Call RenderizarMotorMarked
        
    ElseIf cEXT = "pdf" Then
        Me.Caption = "Leitor de PDF - " & NomeArq(mvarCaminhoArquivo, False)
        m_oWebView2.Navigate "file:///" & Replace(mvarCaminhoArquivo, "\", "/")
         
    Else
        Me.Caption = "Visualizador - " & NomeArq(mvarCaminhoArquivo, False)
        m_oWebView2.Navigate "file:///" & Replace(mvarCaminhoArquivo, "\", "/")
    End If
    Exit Sub
TrataErro:
    m_bPreviewCarregado = False
    MsgBox "Erro ao inicializar o preview: " & Err.Description, vbCritical, "Preview"
End Sub

Private Sub ConfigurarBotoesPreview(ByVal cEXT As String)
    Dim lEhZpl As Boolean
    Dim lEhDocumento As Boolean

    lEhZpl = (cEXT = "zpl")
    lEhDocumento = (cEXT = "rtf" Or cEXT = "docx" Or cEXT = "md" Or cEXT = "markdown")

    cmdSavePNG.Visible = lEhZpl
    cmdSavejpg.Visible = lEhZpl
    cmdSavehtml.Visible = Not (cEXT = "zpl" Or cEXT = "chm" Or cEXT = "hlp" Or cEXT = "xlsx" Or cEXT = "xls" Or cEXT = "ods" Or cEXT = "csv")
    cmdSaveTXT.Visible = lEhDocumento Or cmdSavehtml.Visible
    CmdSavePDF.Visible = Not (cEXT = "pdf" Or cEXT = "chm" Or cEXT = "hlp")
    CmdSavejson.Visible = (cEXT = "xlsx" Or cEXT = "xls" Or cEXT = "ods" Or cEXT = "csv")
    CmdSavemd.Visible = False
    cmdsavedoc.Visible = False
End Sub
    

' ==================================================================
' ROTINA DE VALIDAï¿½ï¿½O DE CONECTIVIDADE DA URL (TIMEOUT ULTRA-CURTO 800ms)
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

Private Function NovoArquivoHtmlTemporario(ByVal cNome As String) As String
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(m_sPastaTemporaria) = 0 Then
        m_sPastaTemporaria = Environ$("TEMP") & "\FrmPreview_" & _
                             CStr(Hex$(CLng(Timer * 1000))) & "_" & CStr(App.hInstance)
        If Not fso.FolderExists(m_sPastaTemporaria) Then
            fso.CreateFolder m_sPastaTemporaria
        End If
    End If

    NovoArquivoHtmlTemporario = m_sPastaTemporaria & "\" & cNome
    If Not m_cArquivosTemporarios Is Nothing Then
        m_cArquivosTemporarios.Add NovoArquivoHtmlTemporario
    End If
End Function

Private Sub LimparArquivosTemporarios()
    Dim fso As Object
    Dim vArquivo As Variant

    On Error GoTo Saida
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not m_cArquivosTemporarios Is Nothing Then
        For Each vArquivo In m_cArquivosTemporarios
            If fso.FileExists(CStr(vArquivo)) Then fso.DeleteFile CStr(vArquivo), True
        Next vArquivo
    End If
    If Len(m_sPastaTemporaria) > 0 Then
        If fso.FolderExists(m_sPastaTemporaria) Then fso.DeleteFolder m_sPastaTemporaria, True
    End If
    If Len(m_sPastaChm) > 0 Then
        If fso.FolderExists(m_sPastaChm) Then fso.DeleteFolder m_sPastaChm, True
    End If
Saida:
    Set fso = Nothing
End Sub

Private Sub RenderizarMotorMarked()
    Dim fso As Object, streamOut As Object
    Dim cHtmlTempPath As String, cJsLocal As String
    Dim cConteudoMD As String
    Dim cConteudoTratado As String
    
    cJsLocal = App.Path & "\WebResources\marked.min.js"
    cHtmlTempPath = NovoArquivoHtmlTemporario("marked_engine.html")
    
    ' 1. Lï¿½ o conteï¿½do do arquivo
    cConteudoMD = LerArquivoTexto(mvarCaminhoArquivo)
    
    ' 2. TRATAMENTO DO CONTEï¿½DO (CRï¿½TICO)
    ' Substituï¿½mos os caracteres que quebrariam a string do JavaScript
    cConteudoTratado = Replace(cConteudoMD, "\", "\\")      ' Escapa barras invertidas
    cConteudoTratado = Replace(cConteudoTratado, "`", "\`")  ' Escapa crases (fim da template literal)
    cConteudoTratado = Replace(cConteudoTratado, "'", "\'")  ' Escapa aspas simples
    cConteudoTratado = Replace(cConteudoTratado, """", "\""") ' Escapa aspas duplas
    cConteudoTratado = Replace(cConteudoTratado, vbCrLf, "\n") ' Normaliza quebras de linha
    cConteudoTratado = Replace(cConteudoTratado, vbCr, "\n")
    cConteudoTratado = Replace(cConteudoTratado, vbLf, "\n")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set streamOut = fso.OpenTextFile(cHtmlTempPath, 2, True)
    
    ' 3. Geraï¿½ï¿½o do HTML temporï¿½rio
    streamOut.WriteLine "<html><head><meta charset='utf-8'>"
    streamOut.WriteLine "<script src='file:///" & Replace(cJsLocal, "\", "/") & "'></script>"
    streamOut.WriteLine "<style>body{font-family:sans-serif; padding:20px; word-wrap:break-word;}</style></head>"
    streamOut.WriteLine "<body><div id='content'></div>"
    streamOut.WriteLine "<script>"
    
    ' Aqui injetamos o conteï¿½do jï¿½ tratado
    streamOut.WriteLine "document.getElementById('content').innerHTML = marked.parse(`" & cConteudoTratado & "`);"
    
    streamOut.WriteLine "</script></body></html>"
    streamOut.Close
    
    ' 4. Navegaï¿½ï¿½o segura
    m_oWebView2.Navigate "file:///" & Replace(cHtmlTempPath, "\", "/")
    
    Set fso = Nothing
End Sub

' Funï¿½ï¿½o auxiliar para leitura simples
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
    
    On Error GoTo TrataErro
    
    cJsLocal = App.Path & "\WebResources\bwip-js-min.js"
    cHtmlTempPath = NovoArquivoHtmlTemporario("zpl_view_engine.html")
    
    ' 1. LER O CONTEï¿½DO DO ARQUIVO ZPL
    cConteudoZpl = LerArquivoTexto(mvarCaminhoArquivo)
    
    ' Tratamento para evitar quebras de string no JS
    cConteudoZpl = Replace(cConteudoZpl, "\", "\\")
    cConteudoZpl = Replace(cConteudoZpl, vbCrLf, "\n")
    cConteudoZpl = Replace(cConteudoZpl, "'", "\'")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set streamOut = fso.OpenTextFile(cHtmlTempPath, 2, True)
    
    ' 2. ESTRUTURA HTML COM INJEï¿½ï¿½O DO ZPL
    streamOut.WriteLine "<html><head><meta charset='utf-8'></head>"
    streamOut.WriteLine "<body>"
    streamOut.WriteLine "<canvas id='zplCanvas'></canvas>"
    streamOut.WriteLine "<script src='file:///" & Replace(cJsLocal, "\", "/") & "'></script>"
    streamOut.WriteLine "<script>"
    ' Injeï¿½ï¿½o do conteï¿½do e comando de renderizaï¿½ï¿½o
    streamOut.WriteLine "try {"
    streamOut.WriteLine "  bwipjs.toCanvas('zplCanvas', {"
    streamOut.WriteLine "    bcid: 'code128', "
    streamOut.WriteLine "    text: '" & cConteudoZpl & "'"
    streamOut.WriteLine "  });"
    streamOut.WriteLine "} catch(e) { document.body.innerHTML = 'Erro JS: ' + e; }"
    streamOut.WriteLine "</script>"
    streamOut.WriteLine "</body></html>"
    streamOut.Close
    
    ' 3. NAVEGAï¿½ï¿½O
    m_oWebView2.Navigate "file:///" & Replace(cHtmlTempPath, "\", "/")
    
    Set fso = Nothing
    Exit Sub

TrataErro:
    If Not streamOut Is Nothing Then streamOut.Close
    Set streamOut = Nothing
    Set fso = Nothing
    Err.Raise Err.Number, "RenderizarMotorZplLocal", Err.Description
End Sub

Private Sub RenderizarMotorPlanilhaLocal()
    Dim fso As Object, streamOut As Object
    Dim cHtmlTempPath As String, cJsLocal As String
    Dim cCaminhoPlanilhaFormatado As String
    
    cJsLocal = App.Path & "\WebResources\xlsx.full.min.js"
    cHtmlTempPath = NovoArquivoHtmlTemporario("xlsx_engine.html")
    cCaminhoPlanilhaFormatado = Replace(mvarCaminhoArquivo, "\", "/")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set streamOut = fso.OpenTextFile(cHtmlTempPath, 2, True)
    
    streamOut.WriteLine "<!DOCTYPE html><html><head><meta charset='utf-8'>"
    streamOut.WriteLine "<style>body{font-family:'Segoe UI',Arial,sans-serif;margin:20px;background:#f4f4f9;}"
    streamOut.WriteLine "#sheet-container{background:#fff;padding:20px;box-shadow:0 4px 15px rgba(0,0,0,0.1);border-radius:4px;overflow-x:auto;}"
    streamOut.WriteLine "table{border-collapse:collapse;width:100%;margin:10px 0;font-size:14px;}"
    streamOut.WriteLine "th{background-color:#107c41;color:white;font-weight:bold;padding:10px;border:1px solid #ddd;}"
    streamOut.WriteLine "td{padding:8px;border:1px solid #ddd;text-align:left;}tr:nth-child(even){background-color:#f9f9f9;}</style>"
    streamOut.WriteLine "<script src='file:///" & Replace(cJsLocal, "\", "/") & "'></script></head>"
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
    
    m_oWebView2.Navigate "file:///" & Replace(cHtmlTempPath, "\", "/")
    Set fso = Nothing
End Sub
Private Sub RenderizarMotorRtfLocal()
    Dim fso As Object, streamOut As Object
    Dim cHtmlTempPath As String
    Dim cCaminhoRtfFormatado As String
    
    cHtmlTempPath = NovoArquivoHtmlTemporario("rtf_view_engine.html")
    cCaminhoRtfFormatado = Replace(mvarCaminhoArquivo, "\", "/")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set streamOut = fso.OpenTextFile(cHtmlTempPath, 2, True)
    
    streamOut.WriteLine "<!DOCTYPE html><html><head><meta charset='utf-8'>"
    streamOut.WriteLine "<style>body{font-family:'Segoe UI',Arial,sans-serif;margin:30px;background-color:#f4f4f9;color:#333;}"
    streamOut.WriteLine "#rtf-container{max-width:800px;margin:0 auto;background:#fff;padding:40px;box-shadow:0 4px 15px rgba(0,0,0,0.1);border-radius:4px;min-height:600px;white-space:pre-wrap;}</style></head>"
    streamOut.WriteLine "<body><div id='rtf-container'>Processando documento...</div>"
    streamOut.WriteLine "<script>"
    streamOut.WriteLine "fetch('file:///" & cCaminhoRtfFormatado & "').then(res => res.text()).then(rtfText => {"
    ' Regex para limpeza bï¿½sica de RTF: remove chaves, comandos e formataï¿½ï¿½o tï¿½cnica
    streamOut.WriteLine "  let clean = rtfText.replace(/\\{[^}]*\\}/g, '').replace(/\\[a-z0-9]+/ig, '').replace(/[\r\n]+/g, '\n');"
    streamOut.WriteLine "  document.getElementById('rtf-container').textContent = clean;"
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
    
    cJsLocal = App.Path & "\WebResources\rtf-parser.min.js"
    cHtmlTempPath = NovoArquivoHtmlTemporario("rtf_view_engine_old.html")
    cCaminhoRtfFormatado = Replace(mvarCaminhoArquivo, "\", "/")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set streamOut = fso.OpenTextFile(cHtmlTempPath, 2, True)
    
    streamOut.WriteLine "<!DOCTYPE html><html><head><meta charset='utf-8'>"
    streamOut.WriteLine "<style>body{font-family:'Segoe UI',Arial,sans-serif;margin:30px;background-color:#f4f4f9;color:#333;}"
    streamOut.WriteLine "#rtf-container{max-width:800px;margin:0 auto;background:#fff;padding:40px;box-shadow:0 4px 15px rgba(0,0,0,0.1);border-radius:4px;min-height:600px;white-space:pre-wrap;}</style>"
    streamOut.WriteLine "<script src='file:///" & Replace(cJsLocal, "\", "/") & "'></script></head>"
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
    
    m_oWebView2.Navigate "file:///" & Replace(cHtmlTempPath, "\", "/")
    Set fso = Nothing
End Sub
Private Sub RenderizarMotorDelimitadoLocal(ByVal cCaminhoTxt As String)
    Dim fso As Object, streamOut As Object
    Dim cHtmlTempPath As String
    Dim cConteudoBruto As String, cConteudoTratado As String
    
    cHtmlTempPath = NovoArquivoHtmlTemporario("csv_engine.html")
    cConteudoBruto = LerArquivoTexto(cCaminhoTxt)
    
    ' Escapa caracteres para injeï¿½ï¿½o no Template Literal do JS
    cConteudoTratado = Replace(cConteudoBruto, "\", "\\")
    cConteudoTratado = Replace(cConteudoTratado, "`", "\`")
    cConteudoTratado = Replace(cConteudoTratado, "$", "\$")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set streamOut = fso.OpenTextFile(cHtmlTempPath, 2, True)
    
    streamOut.WriteLine "<!DOCTYPE html><html><head><meta charset='utf-8'>"
    streamOut.WriteLine "<style>body{font-family:sans-serif;padding:20px;} table{border-collapse:collapse;width:100%;} th,td{border:1px solid #ddd;padding:8px;} th{background:#eee;}</style></head>"
    streamOut.WriteLine "<body><div id='content'>Gerando tabela...</div>"
    streamOut.WriteLine "<script>"
    streamOut.WriteLine "const csvData = `" & cConteudoTratado & "`;"
    ' Parser JS nativo para respeitar aspas duplas no CSV
    streamOut.WriteLine "function parseCSV(text) { let p='', row=[''], ret=[row], i=0, r=0, s=!0, l; for (l of text) { if ('\" ' === l) { if (s && l === p) row[i] += l; s = !s; } else if (',' === l && s) l = row[++i] = ''; else if ('\\n' === l && s) { if ('\\r' === p) row[i] = row[i].slice(0, -1); row = ret[++r] = [l = '']; i = 0; } else row[i] += l; p = l; } return ret; }"
    streamOut.WriteLine "const rows = parseCSV(csvData);"
    streamOut.WriteLine "const esc = value => String(value).replace(/[&<>]/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[ch]));"
    streamOut.WriteLine "let html = '<table><thead><tr>';"
    streamOut.WriteLine "if(rows.length > 0) { rows[0].forEach(h => html += '<th>' + esc(h) + '</th>'); html += '</tr></thead><tbody>'; }"
    streamOut.WriteLine "for(let j=1; j<rows.length; j++) { if(rows[j].join('').trim() !== '') { html += '<tr>'; rows[j].forEach(c => html += '<td>' + esc(c) + '</td>'); html += '</tr>'; } }"
    streamOut.WriteLine "html += '</tbody></table>';"
    streamOut.WriteLine "document.getElementById('content').innerHTML = html;"
    streamOut.WriteLine "</script></body></html>"
    streamOut.Close
    
    m_oWebView2.Navigate "file:///" & Replace(cHtmlTempPath, "\", "/")
    Set fso = Nothing
End Sub
Private Sub RenderizarMotorDocLocal()
    Dim fso As Object, streamOut As Object
    Dim cHtmlTempPath As String, cJsLocal As String
    Dim cCaminhoDocFormatado As String
    
    cJsLocal = App.Path & "\WebResources\mammoth.browser.min.js"
    cHtmlTempPath = NovoArquivoHtmlTemporario("doc_engine.html")
    cCaminhoDocFormatado = Replace(mvarCaminhoArquivo, "\", "/")
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set streamOut = fso.OpenTextFile(cHtmlTempPath, 2, True)
    
    streamOut.WriteLine "<!DOCTYPE html><html><head><meta charset='utf-8'>"
    streamOut.WriteLine "<style>body{font-family:'Segoe UI',Arial,sans-serif;margin:30px;background:#f4f4f9;}"
    streamOut.WriteLine "#document-container{max-width:800px;margin:0 auto;background:#fff;padding:40px;box-shadow:0 4px 15px rgba(0,0,0,0.1);border-radius:4px;}"
    streamOut.WriteLine "table{border-collapse:collapse;width:100%;}th,td{border:1px solid #ddd;padding:8px;}img{max-width:100%;height:auto;}</style>"
    streamOut.WriteLine "<script src='file:///" & Replace(cJsLocal, "\", "/") & "'></script></head>"
    streamOut.WriteLine "<body><div id='document-container'>Carregando documento local...</div>"
    streamOut.WriteLine "<script>"
    streamOut.WriteLine "fetch('file:///" & cCaminhoDocFormatado & "')"
    streamOut.WriteLine "  .then(res => res.arrayBuffer())"
    streamOut.WriteLine "  .then(ab => mammoth.convertToHtml({arrayBuffer: ab}))"
    streamOut.WriteLine "  .then(result => { document.getElementById('document-container').innerHTML = result.value; })"
    streamOut.WriteLine "  .catch(err => { document.getElementById('document-container').innerHTML = '<h2>Erro ao carregar o documento.</h2><p>' + err.message + '</p>'; });"
    streamOut.WriteLine "</script></body></html>"
    streamOut.Close
    
    m_oWebView2.Navigate "file:///" & Replace(cHtmlTempPath, "\", "/")
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
                 "    <h3>Formato de Ajuda Incompatï¿½vel</h3>" & _
                 "    <p>O ficheiro <span class='file-name'>" & NomeArq(mvarCaminhoArquivo, True) & "</span> estï¿½ num formato legado (.hlp) cujo motor de renderizaï¿½ï¿½o foi descontinuado pela Microsoft no Windows 10/11.</p>" & _
                 "    <p>Recomenda-se a conversï¿½o deste manual para formatos suportados como <b>.CHM</b> ou <b>.PDF</b> para exibiï¿½ï¿½o integrada.</p>" & _
                 "  </div>" & _
                 "</body></html>"
                 
    m_oWebView2.NavigateToString cHTMLAviso
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

Private Sub RenderizarMotorChmLocal(ByVal cCaminhoCHM As String)
    Dim sTempPath As String, sArquivoPrincipal As String
    Dim objShell As Object, objFSO As Object, f As Object
    
    sTempPath = Environ$("TEMP") & "\FrmPreview_" & _
                CStr(Hex$(CLng(Timer * 1000))) & "_" & CStr(App.hInstance) & "_chm"
    m_sPastaChm = sTempPath
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Not objFSO.FolderExists(sTempPath) Then objFSO.CreateFolder sTempPath
    
    Set objShell = CreateObject("Shell.Application")
    objShell.NameSpace(sTempPath).CopyHere objShell.NameSpace(cCaminhoCHM).Items
    If Not AguardarExtracaoChm(objFSO, sTempPath, 10) Then
        m_oWebView2.NavigateToString "<html><body><h3>Falha ao extrair o arquivo CHM.</h3></body></html>"
        Exit Sub
    End If
    
    ' Busca os arquivos padrï¿½es de entrada quebrando a linha apï¿½s o "Then"
    If objFSO.FileExists(sTempPath & "\index.html") Then
        sArquivoPrincipal = "\index.html"
    ElseIf objFSO.FileExists(sTempPath & "\default.html") Then
        sArquivoPrincipal = "\default.html"
    ElseIf objFSO.FileExists(sTempPath & "\default.htm") Then
        sArquivoPrincipal = "\default.htm"
    ElseIf objFSO.FileExists(sTempPath & "\index.htm") Then
        sArquivoPrincipal = "\index.htm"
    Else
        ' Fallback genï¿½rico: pega o primeiro HTML que encontrar na raiz
        For Each f In objFSO.GetFolder(sTempPath).Files
            If LCase(objFSO.GetExtensionName(f.Name)) = "html" Or LCase(objFSO.GetExtensionName(f.Name)) = "htm" Then
                sArquivoPrincipal = "\" & f.Name
                Exit For
            End If
        Next f
    End If
    
    If sArquivoPrincipal <> "" Then
        m_oWebView2.Navigate "file:///" & Replace(sTempPath & sArquivoPrincipal, "\", "/")
    Else
        m_oWebView2.NavigateToString "<html><body><h3>Falha ao localizar ï¿½ndice do CHM.</h3></body></html>"
    End If
    
    Set objShell = Nothing
    Set objFSO = Nothing
End Sub

Private Function AguardarExtracaoChm(ByVal objFSO As Object, ByVal cPasta As String, ByVal nSegundos As Long) As Boolean
    Dim dInicio As Double
    dInicio = Timer
    Do
        If objFSO.GetFolder(cPasta).Files.Count > 0 Then
            AguardarExtracaoChm = True
            Exit Function
        End If
        DoEvents
        If Timer - dInicio >= nSegundos Then Exit Do
    Loop
End Function
