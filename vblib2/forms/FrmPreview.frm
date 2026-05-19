VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{F74FE4B4-2AF4-4F6A-B323-05690FDB6730}#1.3#0"; "OrdoWebView2.OCX"
Begin VB.Form FrmPreview 
   Caption         =   "FrmPreview"
   ClientHeight    =   5124
   ClientLeft      =   60
   ClientTop       =   408
   ClientWidth     =   13188
   LinkTopic       =   "Form1"
   ScaleHeight     =   5124
   ScaleWidth      =   13188
   StartUpPosition =   3  'Windows Default
   Begin OrdoWebView2.OrdoWebView OrdoWebView1 
      Height          =   4695
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   10815
      _ExtentX        =   19071
      _ExtentY        =   8276
      Search_URL      =   ""
      IsScriptEnabled =   0   'False
      UseEdgeFixedVersion=   -1  'True
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   11520
      TabIndex        =   1
      Top             =   720
      Width           =   1215
      _ExtentX        =   2138
      _ExtentY        =   762
      Picture         =   "FrmPreview.frx":0000
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
      Height          =   435
      Left            =   11520
      TabIndex        =   2
      Top             =   240
      Width           =   1215
      _ExtentX        =   2138
      _ExtentY        =   762
      Picture         =   "FrmPreview.frx":059A
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
      Height          =   375
      Left            =   11520
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
      _ExtentX        =   2138
      _ExtentY        =   656
      Picture         =   "FrmPreview.frx":0B34
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
      Height          =   375
      Left            =   11520
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
      _ExtentX        =   2138
      _ExtentY        =   656
      Picture         =   "FrmPreview.frx":10CE
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
      Left            =   11520
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   550
      Picture         =   "FrmPreview.frx":1468
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
      Left            =   11520
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2880
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   550
      Picture         =   "FrmPreview.frx":1902
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
      Height          =   375
      Left            =   11520
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2138
      _ExtentY        =   656
      Picture         =   "FrmPreview.frx":1D9C
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
      Height          =   375
      Left            =   11520
      TabIndex        =   8
      Top             =   3720
      Width           =   1215
      _ExtentX        =   2138
      _ExtentY        =   656
      Picture         =   "FrmPreview.frx":2136
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
Private Sub cmdimp_Click()
OrdoWebView1.showPrintDialog
End Sub

Private Sub cmdSavehtml_Click()
Dim sFileName As String
Dim sFILTER As String
  sFILTER = "Arquivos de Textos (*.HTML)" & vbNullChar & "*.HTML" & vbNullChar & "Todos Arquivo" & vbNullChar & "*.*"
  sFileName = FileSave(Me, sFILTER, 1, "HTML", , , "Salvar HTML Como")
  FileWrite sFileName, OrdoWebView1.GetInnerHTML
End Sub

Private Sub cmdSavejpg_Click()
Dim sFILTER As String
    Dim sFileName As String
    Dim sSugestaoNome As String
    
    ' Extrai o nome do arquivo atual para sugerir na caixa de diálogo (ex: "etiqueta")
    sSugestaoNome = NomeArq(mvarCaminhoArquivo, True)
    
    ' Filtro formatado usando "|" conforme exigido pela sua função mFileAPI
    sFILTER = "Imagem JPG (*.jpg;*.jpeg)|*.jpg;*.jpeg|Todos Arquivos (*.*)|*.*"
    
    sFileName = FileSave(Me, sFILTER, 1, "JPG", sSugestaoNome, , "Salvar Etiqueta Como Imagem JPG")
    
    If sFileName <> "" Then
        Call SalvarPreviewComoImagem(sFileName)
    End If
End Sub



Private Sub cmdSavePNG_Click()
Dim sFILTER As String
    Dim sFileName As String
    Dim sSugestaoNome As String
    
    ' Extrai o nome do arquivo atual para sugerir na caixa de diálogo
    sSugestaoNome = NomeArq(mvarCaminhoArquivo, True)
    
    ' Filtro formatado usando "|" conforme exigido pela sua função mFileAPI
    sFILTER = "Imagem PNG (*.png)|*.png|Todos Arquivos (*.*)|*.*"
    
    sFileName = FileSave(Me, sFILTER, 1, "PNG", sSugestaoNome, , "Salvar Etiqueta Como Imagem PNG")
    
    If sFileName <> "" Then
        Call SalvarPreviewComoImagem(sFileName)
    End If
End Sub

Private Sub cmdSaveTXT_Click()
Dim sFileName As String
Dim sFILTER As String
  sFILTER = "Arquivos de Textos (*.TXT)" & vbNullChar & "*.TXT" & vbNullChar & "Todos Arquivo" & vbNullChar & "*.*"
  sFileName = FileSave(Me, sFILTER, 1, "HTML", , , "Salvar TXT Como")
  FileWrite sFileName, OrdoWebView1.GetInnerText
End Sub

Private Sub Encerrar_Click()
Unload Me
End Sub

Private Sub Form_Load()
    ' 1. Captura o destino passado pela variável global de transporte
    mvarCaminhoArquivo = Trim(CStr(cARQRTF))
    
    ' 2. CONTROLE DE INTERFACE DINÂMICO
    If InStr(1, mvarCaminhoArquivo, ".zpl", vbTextCompare) > 0 Then
        ' --- MODO ETIQUETA ZPL ---
        cmdSavehtml.Visible = False
        cmdSaveTXT.Visible = False
        cmdSavePNG.Visible = True
        cmdSavejpg.Visible = True
        
        cmdSavePNG.Top = cmdSavehtml.Top
        cmdSavejpg.Top = cmdSavePNG.Top + cmdSavePNG.Height + 120
        
        Call RenderizarMotorZplLocal
    ElseIf InStr(1, mvarCaminhoArquivo, ".xlsx", vbTextCompare) > 0 Or _
           InStr(1, mvarCaminhoArquivo, ".xls", vbTextCompare) > 0 Or _
           InStr(1, mvarCaminhoArquivo, ".ods", vbTextCompare) > 0 Then
        ' --- MODO LEITOR DE PLANILHAS (EXCEL / CALC) ---
        cmdSavehtml.Visible = False
        cmdSaveTXT.Visible = False
        cmdSavePNG.Visible = False
        cmdSavejpg.Visible = False
        
        Me.Caption = "Visualizador de Planilhas - " & NomeArq(mvarCaminhoArquivo, False)
        
        ' Dispara o motor de renderização de planilhas local
        Call RenderizarMotorPlanilhaLocal
   ElseIf InStr(1, mvarCaminhoArquivo, ".csv", vbTextCompare) > 0 Then
      ' Tanto arquivos Markdown quanto textos puros ganham o tratamento visual moderno
       RenderizarMotorDelimitadoLocal mvarCaminhoArquivo
   ElseIf InStr(1, mvarCaminhoArquivo, ".rtf", vbTextCompare) > 0 Then
        ' --- MODO LEITOR DE RTF MODERNO ---
        cmdSavehtml.Visible = False
        cmdSaveTXT.Visible = False
        cmdSavePNG.Visible = False
        cmdSavejpg.Visible = False
        
        Me.Caption = "Visualizador RTF - " & NomeArq(mvarCaminhoArquivo, False)
        
        ' Dispara a renderização e conversão do RTF para HTML local
        Call RenderizarMotorRtfLocal
        
   ElseIf InStr(1, mvarCaminhoArquivo, ".docx", vbTextCompare) > 0 Or _
           InStr(1, mvarCaminhoArquivo, ".doc", vbTextCompare) > 0 Or _
           InStr(1, mvarCaminhoArquivo, ".odt", vbTextCompare) > 0 Then
        ' --- MODO LEITOR DE DOCUMENTOS (WORD / ODT) ---
        ' Escondemos os botões de gravação de texto padrão do VB6 por enquanto
        cmdSavehtml.Visible = False
        cmdSaveTXT.Visible = False
        cmdSavePNG.Visible = False
        cmdSavejpg.Visible = False
        
        Me.Caption = "Visualizador de Documentos - " & NomeArq(mvarCaminhoArquivo, False)
        
        ' Dispara a renderização e conversão do Word para HTML local
        Call RenderizarMotorDocLocal
        
    ElseIf InStr(1, mvarCaminhoArquivo, ".pdf", vbTextCompare) > 0 Then
        ' --- MODO LEITOR DE PDF ---
        ' No modo PDF, os botões de salvar texto/HTML do VB6 não fazem sentido,
        ' pois o próprio leitor do Chromium já tem botões internos de download/print.
        cmdSavehtml.Visible = False
        cmdSaveTXT.Visible = False
        cmdSavePNG.Visible = False
        cmdSavejpg.Visible = False
        
        ' Altera o título da tela com o nome do PDF
        Me.Caption = "Leitor de PDF - " & NomeArq(mvarCaminhoArquivo, False)
        
        ' Ativa o motor e navega para o arquivo PDF local usando o protocolo correto
        OrdoWebView1.IsScriptEnabled = True
        OrdoWebView1.Navigate "file:///" & Replace(mvarCaminhoArquivo, "\", "/")
        
    Else
        ' --- MODO RELATÓRIO NORMAL (HTML, TXT, URL) ---
        cmdSavehtml.Visible = True
        cmdSaveTXT.Visible = True
        cmdSavePNG.Visible = False
        cmdSavejpg.Visible = False
        
        Me.Caption = "Visualizador - " & mvarCaminhoArquivo
        
        If InStr(1, mvarCaminhoArquivo, "://", vbTextCompare) = 0 Then
            OrdoWebView1.Navigate "file:///" & Replace(mvarCaminhoArquivo, "\", "/")
        Else
            OrdoWebView1.Navigate mvarCaminhoArquivo
        End If
    End If
    
    DoEvents
End Sub
Private Sub Form_Resize()
On Error Resume Next
    ' Redimensionamento dinâmico responsivo: o preview acompanha o tamanho da janela do utilizador
    OrdoWebView1.Height = Me.ScaleHeight - OrdoWebView1.Top - 200
    OrdoWebView1.Width = Me.ScaleWidth - OrdoWebView1.Left - 1800 ' Reserva o espaço dos botões à direita
End Sub

'==============================================================================
' FUNÇÃO ISOLADA: Só é acionada se a string contiver a marcação ".zpl"
' ==============================================================================
Private Sub RenderizarMotorZplLocal()
    Dim fso As Object
    Dim streamIn As Object
    Dim streamOut As Object
    Dim cZplConteudo As String
    Dim cHtmlTempPath As String
    
    Me.Caption = "Preview de Etiqueta ZPL (Modo Offline)"
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Se for um arquivo físico local, lê os comandos ZPL via Stream FSO
    If fso.FileExists(mvarCaminhoArquivo) Then
        Set streamIn = fso.OpenTextFile(mvarCaminhoArquivo, 1, False)
        cZplConteudo = streamIn.ReadAll
        streamIn.Close
        Set streamIn = Nothing
    Else
        ' Se não for um arquivo no HD (ex: se for uma URL ZPL), passamos uma string de aviso ou tratamos
        cZplConteudo = "^XA^FO50,50^A0N,36,36^FDFALHA: ARQUIVO LOCAL NAO ENCONTRADO^FS^XZ"
    End If
    
    ' Trata caracteres de escape e quebras de linha para o JavaScript não quebrar
    cZplConteudo = Replace(cZplConteudo, "\", "\\")
    cZplConteudo = Replace(cZplConteudo, "'", "\'")
    cZplConteudo = Replace(cZplConteudo, vbCrLf, "\n")
    cZplConteudo = Replace(cZplConteudo, vbLf, "\n")
    
    ' Define o caminho do arquivo HTML temporário do motor gráfico
    cHtmlTempPath = App.Path & "\~zpl_view_engine.html"
    
    ' Cria a página injetando o renderizador matemático de Canvas (zpl-image)
    Set streamOut = fso.OpenTextFile(cHtmlTempPath, 2, True)
    streamOut.WriteLine "<html>"
    streamOut.WriteLine "<head>"
    streamOut.WriteLine "  <meta charset='utf-8'>"
    streamOut.WriteLine "  <style>"
    streamOut.WriteLine "    body { margin:0; padding:10px; background:#E0E0E0; display:flex; justify-content:center; align-items:center; height:100vh; font-family:Arial; }"
    streamOut.WriteLine "    #canvas { box-shadow: 0px 5px 15px rgba(0,0,0,0.3); background:white; max-width:100%; max-height:95vh; object-fit:contain; }"
    streamOut.WriteLine "  </style>"
    streamOut.WriteLine "  <script src='https://cdn.jsdelivr.net/npm/zpl-image@0.1.3/zpl-image.min.js'></script>"
    streamOut.WriteLine "</head>"
    streamOut.WriteLine "<body>"
    streamOut.WriteLine "  <canvas id='canvas'></canvas>"
    streamOut.WriteLine "  <script>"
    streamOut.WriteLine "    try {"
    streamOut.WriteLine "      var code = '" & cZplConteudo & "';"
    streamOut.WriteLine "      var cvs = document.getElementById('canvas');"
    streamOut.WriteLine "      zplImage.render(code, cvs);"
    streamOut.WriteLine "    } catch(err) {"
    streamOut.WriteLine "      document.body.innerHTML = '<b style=""color:red"">Erro: </b>' + err.message;"
    streamOut.WriteLine "    }"
    streamOut.WriteLine "  </script>"
    streamOut.WriteLine "</body>"
    streamOut.WriteLine "</html>"
    streamOut.Close
    
    OrdoWebView1.IsScriptEnabled = True
    ' Executa a navegação estável para o motor gerado localmente
    OrdoWebView1.Navigate "file:///" & Replace(cHtmlTempPath, "\", "/")
    
    Set streamOut = Nothing
    Set fso = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
' 1. IF SIMPLES NA UNLOAD: Só executa a verificação e limpeza se o arquivo aberto for ZPL
    If InStr(1, mvarCaminhoArquivo, ".zpl", vbTextCompare) > 0 Then
        Dim fso As Object
        Dim cHtmlTempPath As String
        
        On Error Resume Next
        Set fso = CreateObject("Scripting.FileSystemObject")
        cHtmlTempPath = App.Path & "\~zpl_view_engine.html"
        
        ' Garante que o arquivo temporário existe antes de apagar
        If fso.FileExists(cHtmlTempPath) Then
            fso.DeleteFile cHtmlTempPath, True
        End If
        
        Set fso = Nothing
    End If
End Sub

' ==============================================================================
' FUNÇÃO: Força o navegador a extrair o Canvas (PNG ou JPG) e dispara o download
' ==============================================================================
Public Sub SalvarPreviewComoImagem(ByVal cCaminhoDestino As String)
    On Error GoTo ErroSalvar
    Dim fso As Object
    Dim cNomeArquivo As String
    Dim cEXTENSAO As String
    Dim cMimeType As String
    Dim cScript As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    cNomeArquivo = fso.GetFileName(cCaminhoDestino)
    cEXTENSAO = LCase(fso.GetExtensionName(cCaminhoDestino))
    Set fso = Nothing
    
    ' Define o MimeType correto com base na escolha do utilizador
    If cEXTENSAO = "jpg" Or cEXTENSAO = "jpeg" Then
        cMimeType = "image/jpeg"
    Else
        cMimeType = "image/png"
    End If
    
    ' Garante que o motor aceita a execução de scripts antes de rodar
    OrdoWebView1.IsScriptEnabled = True
    
    ' Monta o script JavaScript adaptado. Se for JPEG, ele cria um canvas temporário
    ' com fundo branco para evitar que a transparência por trás da etiqueta fique preta.
    cScript = "(function() {" & _
              "  var canvas = document.getElementById('canvas');" & _
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
              
    ' Executa o script de download nativo do Chromium sem travar o VB6
    OrdoWebView1.ExecuteScript cScript
    
    Exit Sub

ErroSalvar:
    MsgBox "Falha ao exportar imagem da etiqueta: " & Err.Description, vbCritical, "Erro de Salvamento"
End Sub
Public Sub RenderizarMotorMDLocal(ByVal cCaminhoArquivo As String)
    On Error GoTo ErroRenderMD
    
    Dim fNum            As Integer
    Dim cConteudoMD     As String
    Dim cHTMLTemplate   As String
    
    ' 1. Validação de segurança do arquivo físico
    If Dir(cCaminhoArquivo) = "" Then
        MsgBox "O arquivo especificado não foi encontrado: " & cCaminhoArquivo, vbExclamation, "Aviso"
        Exit Sub
    End If
    
    ' 2. Captura o conteúdo do arquivo de texto bruto
    fNum = FreeFile
    Open cCaminhoArquivo For Input As #fNum
        cConteudoMD = Input(LOF(fNum), #fNum)
    Close #fNum
    
    ' 3. Sanitização estrita para não quebrar o JavaScript do WebView2
    cConteudoMD = Replace(cConteudoMD, vbCrLf, "\n")
    cConteudoMD = Replace(cConteudoMD, vbCr, "\n")
    cConteudoMD = Replace(cConteudoMD, vbLf, "\n")
    cConteudoMD = Replace(cConteudoMD, "'", "\'")
    
    ' 4. Montagem estruturada do HTML (Acumulando em blocos para evitar o erro de limite de linhas)
    cHTMLTemplate = "<html><head><meta charset='utf-8'>"
    cHTMLTemplate = cHTMLTemplate & "<script src='https://cdn.jsdelivr.net/npm/marked/marked.min.js'></script>"
    
    ' Injeção de Estilos (CSS padrão GitHub)
    cHTMLTemplate = cHTMLTemplate & "<style>"
    cHTMLTemplate = cHTMLTemplate & "  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Helvetica, Arial, sans-serif; padding: 35px; background-color: #ffffff; color: #24292f; line-height: 1.6; }"
    cHTMLTemplate = cHTMLTemplate & "  code { background-color: rgba(175,184,193,0.2); padding: 2px 5px; border-radius: 6px; font-family: monospace; font-size: 85%; }"
    cHTMLTemplate = cHTMLTemplate & "  pre { background-color: #f6f8fa; padding: 16px; border-radius: 6px; overflow: auto; }"
    cHTMLTemplate = cHTMLTemplate & "  table { border-collapse: collapse; width: 100%; margin-top: 0; margin-bottom: 16px; font-size: 14px; }"
    cHTMLTemplate = cHTMLTemplate & "  th, td { border: 1px solid #d0d7de; padding: 6px 13px; }"
    cHTMLTemplate = cHTMLTemplate & "  th { background-color: #f6f8fa; font-weight: 600; }"
    cHTMLTemplate = cHTMLTemplate & "  tr:nth-child(even) { background-color: #f6f8fa; }"
    cHTMLTemplate = cHTMLTemplate & "  h1, h2, h3 { border-bottom: 1px solid #d8dee4; padding-bottom: .3em; color: #1f2328; }"
    cHTMLTemplate = cHTMLTemplate & "</style></head>"
    
    ' Injeção do Corpo e do Motor JavaScript (Marked.js)
    cHTMLTemplate = cHTMLTemplate & "<body>"
    cHTMLTemplate = cHTMLTemplate & "  <div id='content'></div>"
    cHTMLTemplate = cHTMLTemplate & "  <script>"
    cHTMLTemplate = cHTMLTemplate & "    document.getElementById('content').innerHTML = marked.parse('" & cConteudoMD & "');"
    cHTMLTemplate = cHTMLTemplate & "  </script>"
    cHTMLTemplate = cHTMLTemplate & "</body></html>"
                 
    ' 5. Alimenta o componente Chromium diretamente através da string consolidada
    OrdoWebView1.NavigateToString cHTMLTemplate
    Exit Sub

ErroRenderMD:
    MsgBox "Falha crítica no motor RenderizarMotorMDLocal: " & Err.Description, vbCritical, "Erro de Engine"
    If fNum > 0 Then Close #fNum
End Sub
Private Sub RenderizarMotorDocLocal()
    On Error GoTo ErroDoc
    Dim cHtmlTempPath As String
    Dim cHtmlContent As String
    Dim cCaminhoDocFormatado As String
    
    ' Define o caminho do arquivo temporário do motor do Word
    cHtmlTempPath = App.Path & "\~doc_view_engine.html"
    
    ' Guarda o caminho do documento real formatando as barras para o JavaScript
    cCaminhoDocFormatado = Replace(mvarCaminhoArquivo, "\", "/")
    
    ' --- CONSTRUÇÃO DA STRING CONCATENADA (Evita o erro 'Too many line continuations') ---
    cHtmlContent = "<!DOCTYPE html>" & vbCrLf
    cHtmlContent = cHtmlContent & "<html>" & vbCrLf
    cHtmlContent = cHtmlContent & "<head>" & vbCrLf
    cHtmlContent = cHtmlContent & "  <meta charset='utf-8'>" & vbCrLf
    cHtmlContent = cHtmlContent & "  <style>" & vbCrLf
    cHtmlContent = cHtmlContent & "    body { font-family: 'Segoe UI', Arial, sans-serif; margin: 30px; background-color: #f4f4f9; color: #333; }" & vbCrLf
    cHtmlContent = cHtmlContent & "    #document-container { max-width: 800px; margin: 0 auto; background: #fff; padding: 40px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); border-radius: 4px; min-height: 1000px; }" & vbCrLf
    cHtmlContent = cHtmlContent & "    table { border-collapse: collapse; width: 100%; margin: 20px 0; }" & vbCrLf
    cHtmlContent = cHtmlContent & "    th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }" & vbCrLf
    cHtmlContent = cHtmlContent & "    img { max-width: 100%; height: auto; }" & vbCrLf
    cHtmlContent = cHtmlContent & "  </style>" & vbCrLf
    cHtmlContent = cHtmlContent & "  " & vbCrLf
    cHtmlContent = cHtmlContent & "  <script src='https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js'></script>" & vbCrLf
    cHtmlContent = cHtmlContent & "</head>" & vbCrLf
    cHtmlContent = cHtmlContent & "<body>" & vbCrLf
    cHtmlContent = cHtmlContent & "  <div id='document-container'>Carregando e processando documento...</div>" & vbCrLf
    cHtmlContent = cHtmlContent & "  <script>" & vbCrLf
    cHtmlContent = cHtmlContent & "    // Faz a requisição local do binário mapeado em formato ArrayBuffer para o parser" & vbCrLf
    cHtmlContent = cHtmlContent & "    fetch('file:///" & cCaminhoDocFormatado & "')" & vbCrLf
    cHtmlContent = cHtmlContent & "      .then(response => response.arrayBuffer())" & vbCrLf
    cHtmlContent = cHtmlContent & "      .then(arrayBuffer => {" & vbCrLf
    cHtmlContent = cHtmlContent & "        // Converte o binário (.docx/.doc) diretamente em HTML sem depender do Office" & vbCrLf
    cHtmlContent = cHtmlContent & "        return mammoth.convertToHtml({arrayBuffer: arrayBuffer});" & vbCrLf
    cHtmlContent = cHtmlContent & "      })" & vbCrLf
    cHtmlContent = cHtmlContent & "      .then(result => {" & vbCrLf
    cHtmlContent = cHtmlContent & "        document.getElementById('document-container').innerHTML = result.value;" & vbCrLf
    cHtmlContent = cHtmlContent & "      })" & vbCrLf
    cHtmlContent = cHtmlContent & "      .catch(err => {" & vbCrLf
    cHtmlContent = cHtmlContent & "        document.getElementById('document-container').innerHTML = '<h2>Erro ao carregar o documento.</h2><p>' + err.message + '</p>';" & vbCrLf
    cHtmlContent = cHtmlContent & "      });" & vbCrLf
    cHtmlContent = cHtmlContent & "  </script>" & vbCrLf
    cHtmlContent = cHtmlContent & "</body>" & vbCrLf
    cHtmlContent = cHtmlContent & "</html>"

    ' Grava fisicamente a casca do motor usando a sua FileWrite do FSO
    If FileWrite(cHtmlTempPath, cHtmlContent) Then
        OrdoWebView1.IsScriptEnabled = True
        OrdoWebView1.Navigate "file:///" & Replace(cHtmlTempPath, "\", "/")
    End If
    Exit Sub

ErroDoc:
    MsgBox "Falha ao instanciar motor de leitura de documentos: " & Err.Description, vbCritical, "Erro"
End Sub
Private Sub RenderizarMotorRtfLocal()
    On Error GoTo ErroRtf
    Dim cHtmlTempPath As String
    Dim cHtmlContent As String
    Dim cCaminhoRtfFormatado As String
    
    cHtmlTempPath = App.Path & "\~rtf_view_engine.html"
    cCaminhoRtfFormatado = Replace(mvarCaminhoArquivo, "\", "/")
    
    ' Constrói a string de forma acumulativa para garantir estabilidade no compilador VB6
    cHtmlContent = "<!DOCTYPE html>" & vbCrLf
    cHtmlContent = cHtmlContent & "<html>" & vbCrLf
    cHtmlContent = cHtmlContent & "<head>" & vbCrLf
    cHtmlContent = cHtmlContent & "  <meta charset='utf-8'>" & vbCrLf
    cHtmlContent = cHtmlContent & "  <style>" & vbCrLf
    cHtmlContent = cHtmlContent & "    body { font-family: 'Segoe UI', Arial, sans-serif; margin: 30px; background-color: #f4f4f9; color: #333; }" & vbCrLf
    cHtmlContent = cHtmlContent & "    #rtf-container { max-width: 800px; margin: 0 auto; background: #fff; padding: 40px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); border-radius: 4px; min-height: 800px; white-space: pre-wrap; }" & vbCrLf
    cHtmlContent = cHtmlContent & "  </style>" & vbCrLf
    cHtmlContent = cHtmlContent & "  " & vbCrLf
    cHtmlContent = cHtmlContent & "  <script src='https://cdn.jsdelivr.net/npm/rtf-parser@1.3.0/dist/rtf-parser.min.js'></script>" & vbCrLf
    cHtmlContent = cHtmlContent & "</head>" & vbCrLf
    cHtmlContent = cHtmlContent & "<body>" & vbCrLf
    cHtmlContent = cHtmlContent & "  <div id='rtf-container'>Processando documento formatado RTF...</div>" & vbCrLf
    cHtmlContent = cHtmlContent & "  <script>" & vbCrLf
    cHtmlContent = cHtmlContent & "    // Lê o conteúdo bruto do RTF do disco rígido" & vbCrLf
    cHtmlContent = cHtmlContent & "    fetch('file:///" & cCaminhoRtfFormatado & "')" & vbCrLf
    cHtmlContent = cHtmlContent & "      .then(response => response.text())" & vbCrLf
    cHtmlContent = cHtmlContent & "      .then(rtfText => {" & vbCrLf
    cHtmlContent = cHtmlContent & "        // Função simplificada de fallback para limpar metadados de controle RTF" & vbCrLf
    cHtmlContent = cHtmlContent & "        // e converter quebras antigas \par em tags web funcionais" & vbCrLf
    cHtmlContent = cHtmlContent & "        let htmlResult = rtfText" & vbCrLf
    cHtmlContent = cHtmlContent & "          .replace(/\\par/g, '<br>')" & vbCrLf
    cHtmlContent = cHtmlContent & "          .replace(/\\b\s+(.*?)\\b0/g, '<b>$1</b>')" & vbCrLf
    cHtmlContent = cHtmlContent & "          .replace(/\\i\s+(.*?)\\i0/g, '<i>$1</i>')" & vbCrLf
    cHtmlContent = cHtmlContent & "          .replace(/\\{[^{}]*\\}/g, ''); // Remove blocos de cabeçalho brutos" & vbCrLf
    cHtmlContent = cHtmlContent & "        " & vbCrLf
    cHtmlContent = cHtmlContent & "        // Remove quaisquer tags residuais de formatação interna do RTF" & vbCrLf
    cHtmlContent = cHtmlContent & "        htmlResult = htmlResult.replace(/\\[a-z0-9-]+/g, '');" & vbCrLf
    cHtmlContent = cHtmlContent & "        htmlResult = htmlResult.replace(/[{}]/g, '');" & vbCrLf
    cHtmlContent = cHtmlContent & "        " & vbCrLf
    cHtmlContent = cHtmlContent & "        document.getElementById('rtf-container').innerHTML = htmlResult.trim();" & vbCrLf
    cHtmlContent = cHtmlContent & "      })" & vbCrLf
    cHtmlContent = cHtmlContent & "      .catch(err => {" & vbCrLf
    cHtmlContent = cHtmlContent & "        document.getElementById('rtf-container').innerHTML = '<h2>Erro ao processar o RTF.</h2><p>' + err.message + '</p>';" & vbCrLf
    cHtmlContent = cHtmlContent & "      });" & vbCrLf
    cHtmlContent = cHtmlContent & "  </script>" & vbCrLf
    cHtmlContent = cHtmlContent & "</body>" & vbCrLf
    cHtmlContent = cHtmlContent & "</html>"

    ' Grava o arquivo temporário usando a sua FileWrite do FSO
    If FileWrite(cHtmlTempPath, cHtmlContent) Then
        OrdoWebView1.IsScriptEnabled = True
        OrdoWebView1.Navigate "file:///" & Replace(cHtmlTempPath, "\", "/")
    End If
    Exit Sub

ErroRtf:
    MsgBox "Falha ao instanciar leitor de arquivos RTF: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub RenderizarMotorPlanilhaLocal()
    On Error GoTo ErroPlanilha
    Dim cHtmlTempPath As String
    Dim cHtmlContent As String
    Dim cCaminhoPlanilhaFormatado As String
    
    cHtmlTempPath = App.Path & "\~sheet_view_engine.html"
    cCaminhoPlanilhaFormatado = Replace(mvarCaminhoArquivo, "\", "/")
    
    ' Constrói a estrutura do motor HTML de forma acumulativa
    cHtmlContent = "<!DOCTYPE html>" & vbCrLf
    cHtmlContent = cHtmlContent & "<html>" & vbCrLf
    cHtmlContent = cHtmlContent & "<head>" & vbCrLf
    cHtmlContent = cHtmlContent & "  <meta charset='utf-8'>" & vbCrLf
    cHtmlContent = cHtmlContent & "  <style>" & vbCrLf
    cHtmlContent = cHtmlContent & "    body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background-color: #f4f4f9; color: #333; }" & vbCrLf
    cHtmlContent = cHtmlContent & "    #sheet-container { background: #fff; padding: 20px; box-shadow: 0 4px 15px rgba(0,0,0,0.1); border-radius: 4px; overflow-x: auto; }" & vbCrLf
    cHtmlContent = cHtmlContent & "    table { border-collapse: collapse; width: 100%; margin: 10px 0; font-size: 14px; }" & vbCrLf
    cHtmlContent = cHtmlContent & "    th { background-color: #107c41; color: white; font-weight: bold; padding: 10px; border: 1px solid #ddd; }" & vbCrLf
    cHtmlContent = cHtmlContent & "    td { padding: 8px; border: 1px solid #ddd; text-align: left; }" & vbCrLf
    cHtmlContent = cHtmlContent & "    tr:nth-child(even) { background-color: #f9f9f9; }" & vbCrLf
    cHtmlContent = cHtmlContent & "    tr:hover { background-color: #f1f1f1; }" & vbCrLf
    cHtmlContent = cHtmlContent & "  </style>" & vbCrLf
    cHtmlContent = cHtmlContent & "  " & vbCrLf
    cHtmlContent = cHtmlContent & "  <script src='https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js'></script>" & vbCrLf
    cHtmlContent = cHtmlContent & "</head>" & vbCrLf
    cHtmlContent = cHtmlContent & "<body>" & vbCrLf
    cHtmlContent = cHtmlContent & "  <div id='sheet-container'>Processando dados da planilha...</div>" & vbCrLf
    cHtmlContent = cHtmlContent & "  <script>" & vbCrLf
    cHtmlContent = cHtmlContent & "    // Carrega o arquivo binário direto do HD" & vbCrLf
    cHtmlContent = cHtmlContent & "    fetch('file:///" & cCaminhoPlanilhaFormatado & "')" & vbCrLf
    cHtmlContent = cHtmlContent & "      .then(response => response.arrayBuffer())" & vbCrLf
    cHtmlContent = cHtmlContent & "      .then(arrayBuffer => {" & vbCrLf
    cHtmlContent = cHtmlContent & "        // Lê o array de bytes da planilha (detecta automaticamente se é XLS, XLSX ou ODS)" & vbCrLf
    cHtmlContent = cHtmlContent & "        let workbook = XLSX.read(new Uint8Array(arrayBuffer), {type: 'array'});" & vbCrLf
    cHtmlContent = cHtmlContent & "        " & vbCrLf
    cHtmlContent = cHtmlContent & "        // Pega o nome da primeira aba ativa" & vbCrLf
    cHtmlContent = cHtmlContent & "        let firstSheetName = workbook.SheetNames[0];" & vbCrLf
    cHtmlContent = cHtmlContent & "        let worksheet = workbook.Sheets[firstSheetName];" & vbCrLf
    cHtmlContent = cHtmlContent & "        " & vbCrLf
    cHtmlContent = cHtmlContent & "        // Converte a aba interna em uma tabela HTML pura" & vbCrLf
    cHtmlContent = cHtmlContent & "        let htmlTable = XLSX.utils.sheet_to_html(worksheet);" & vbCrLf
    cHtmlContent = cHtmlContent & "        " & vbCrLf
    cHtmlContent = cHtmlContent & "        document.getElementById('sheet-container').innerHTML = htmlTable;" & vbCrLf
    cHtmlContent = cHtmlContent & "      })" & vbCrLf
    cHtmlContent = cHtmlContent & "      .catch(err => {" & vbCrLf
    cHtmlContent = cHtmlContent & "        document.getElementById('sheet-container').innerHTML = '<h2>Erro ao processar a planilha.</h2><p>' + err.message + '</p>';" & vbCrLf
    cHtmlContent = cHtmlContent & "      });" & vbCrLf
    cHtmlContent = cHtmlContent & "  </script>" & vbCrLf
    cHtmlContent = cHtmlContent & "</body>" & vbCrLf
    cHtmlContent = cHtmlContent & "</html>"

    ' Grava o arquivo temporário do motor usando a sua FileWrite do FSO
    If FileWrite(cHtmlTempPath, cHtmlContent) Then
        OrdoWebView1.IsScriptEnabled = True
        OrdoWebView1.Navigate "file:///" & Replace(cHtmlTempPath, "\", "/")
    End If
    Exit Sub

ErroPlanilha:
    MsgBox "Falha ao instanciar motor de leitura de planilhas: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub cmdSaveDoc_Click()
    On Error GoTo ErroSaveDoc
    Dim sFILTER As String
    Dim sFileName As String
    Dim sSugestaoNome As String
    
    ' 1. Extrai o nome atual do ficheiro para sugerir na caixa de diálogo
    sSugestaoNome = NomeArq(mvarCaminhoArquivo, True)
    
    ' 2. Define os filtros de salvamento (Permite salvar como Documento Word ou PDF)
    ' Usamos a barra vertical "|" que a sua função FileSave exige para converter para o Windows
    sFILTER = "Documento do Word (*.doc)|*.doc|Documento PDF (*.pdf)|*.pdf|Todos Arquivos (*.*)|*.*"
    
    ' 3. Abre a caixa de diálogo nativa da sua mFileAPI
    sFileName = FileSave(Me, sFILTER, 1, "DOC", sSugestaoNome, , "Salvar Documento Como")
    
    ' Se o utilizador cancelou a janela, sai da sub-rotina
    If sFileName = "" Then Exit Sub
    
    ' 4. Identifica a extensão que o utilizador escolheu para aplicar a lógica correta
    If InStr(1, sFileName, ".pdf", vbTextCompare) > 0 Then
        ' --- EXPORTAR COMO PDF ---
        ' Dispara o comando de impressão nativo do Edge Chromium.
        ' O utilizador poderá escolher "Salvar como PDF" com fidelidade gráfica total.
        OrdoWebView1.ExecuteScript "window.print();"
        
    Else
        ' --- EXPORTAR COMO DOCUMENTO (.DOC) ---
        ' Solicitamos ao WebView2 que nos devolva o HTML exato que está dentro do container do documento
        ' Usamos uma técnica de callback assíncrona ou captura direta se o seu componente suportar,
        ' ou injetamos um script para iniciar o download do blob modificado de texto:
        
        Dim cScript As String
        cScript = "(function() {" & _
                  "  var conteudo = document.getElementById('document-container').innerHTML;" & _
                  "  var blob = new Blob(['<html><head><meta charset='utf-8'></head><body>' + conteudo + '</body></html>'], {type: 'application/msword'});" & _
                  "  var lnk = document.createElement('a');" & _
                  "  lnk.download = '" & NomeArq(sFileName, False) & "';" & _
                  "  lnk.href = window.URL.createObjectURL(blob);" & _
                  "  document.body.appendChild(lnk);" & _
                  "  lnk.click();" & _
                  "  document.body.removeChild(lnk);" & _
                  "})();"
                  
        OrdoWebView1.ExecuteScript cScript
    End If
    
    Exit Sub

ErroSaveDoc:
    MsgBox "Erro ao exportar o documento: " & Err.Description, vbCritical, "Erro de Exportação"
End Sub

Private Sub cmdSavePDF_Click()
    On Error GoTo ErroSavePDF
    Dim sFILTER As String
    Dim sFileName As String
    Dim sSugestaoNome As String
    
    ' 1. Extrai o nome atual do arquivo para sugerir na caixa de diálogo
    sSugestaoNome = NomeArq(mvarCaminhoArquivo, True)
    
    ' 2. Define o filtro específico para salvar em PDF
    ' Usamos a barra vertical "|" que a sua função FileSave exige para converter internamente
    sFILTER = "Documento PDF (*.pdf)|*.pdf|Todos Arquivos (*.*)|*.*"
    
    ' 3. Abre a caixa de diálogo nativa do Windows através da sua mFileAPI
    sFileName = FileSave(Me, sFILTER, 1, "PDF", sSugestaoNome, , "Salvar Como PDF")
    
    ' Se o usuário cancelou a janela de salvamento, aborta a operação
    If sFileName = "" Then Exit Sub
    
    ' 4. Dispara a conversão nativa do Chromium
    ' Isso abre a janela moderna de impressão do Edge configurada para "Salvar como PDF"
    ' garantindo fidelidade total de cores, tabelas e fontes.
    OrdoWebView1.ExecuteScript "window.print();"
    
    Exit Sub

ErroSavePDF:
    MsgBox "Erro ao exportar o arquivo para PDF: " & Err.Description, vbCritical, "Erro de Exportação"
End Sub

Public Sub RenderizarMotorDelimitadoLocal(ByVal cCaminhoArquivo As String)
    On Error GoTo ErroRenderDelimitado
    
    Dim fNum            As Integer
    Dim cConteudoBruto  As String
    Dim cConteudoMD     As String
    Dim cHTMLTemplate   As String
    Dim sDelimitador    As String
    Dim vLinhas         As Variant
    Dim vColunas        As Variant
    Dim i               As Long
    Dim j               As Long
    Dim nMaxColunas     As Long
    Dim cLinhaSeparador As String
    
    ' Variáveis do motor de identificação automática
    Dim sPrimeiraLinha  As String
    Dim lContador       As Long
    Dim lMaxRepeticoes  As Long
    
    ' 1. Validação de segurança do arquivo físico
    If Dir(cCaminhoArquivo) = "" Then
        MsgBox "O arquivo especificado não foi encontrado: " & cCaminhoArquivo, vbExclamation, "Aviso"
        Exit Sub
    End If
    
    ' 2. Captura o conteúdo do arquivo de texto bruto
    fNum = FreeFile
    Open cCaminhoArquivo For Input As #fNum
        cConteudoBruto = Input(LOF(fNum), #fNum)
    Close #fNum
    
    vLinhas = Split(cConteudoBruto, vbCrLf)
    If UBound(vLinhas) < 0 Then Exit Sub
    sPrimeiraLinha = vLinhas(0)
    
    ' 3. MOTOR DE INTELIGÊNCIA: Identifica qual delimitador venceu na amostragem
    ' Definimos o seu favorito (Pipe) como o Fallback padrão inicial
    sDelimitador = "|"
    lMaxRepeticoes = 0
    
    ' Teste 1: Pipe (|) - Seu delimitador principal e mais seguro
    lContador = Len(sPrimeiraLinha) - Len(Replace(sPrimeiraLinha, "|", ""))
    If lContador > lMaxRepeticoes Then lMaxRepeticoes = lContador: sDelimitador = "|"
    
    ' Teste 2: Ponto e Vírgula (;)
    lContador = Len(sPrimeiraLinha) - Len(Replace(sPrimeiraLinha, ";", ""))
    If lContador > lMaxRepeticoes Then lMaxRepeticoes = lContador: sDelimitador = ";"
    
    ' Teste 3: Vírgula (,)
    lContador = Len(sPrimeiraLinha) - Len(Replace(sPrimeiraLinha, ",", ""))
    If lContador > lMaxRepeticoes Then lMaxRepeticoes = lContador: sDelimitador = ","
    
    ' Teste 4: Cerquilha (#)
    lContador = Len(sPrimeiraLinha) - Len(Replace(sPrimeiraLinha, "#", ""))
    If lContador > lMaxRepeticoes Then lMaxRepeticoes = lContador: sDelimitador = "#"
    
    ' Teste 5: Til (~)
    lContador = Len(sPrimeiraLinha) - Len(Replace(sPrimeiraLinha, "~", ""))
    If lContador > lMaxRepeticoes Then lMaxRepeticoes = lContador: sDelimitador = "~"
    
    ' Teste 6: Tabulação (vbTab)
    lContador = Len(sPrimeiraLinha) - Len(Replace(sPrimeiraLinha, vbTab, ""))
    If lContador > lMaxRepeticoes Then lMaxRepeticoes = lContador: sDelimitador = vbTab

    ' 4. Processamento dos vetores (A primeira linha já contém dados válidos)
    vColunas = Split(sPrimeiraLinha, sDelimitador)
    nMaxColunas = UBound(vColunas)
    
    ' Como não há cabeçalho textual enviado, geramos a linha técnica de títulos: | Coluna 1 | Coluna 2 |
    cConteudoMD = "|"
    For j = 0 To nMaxColunas
        cConteudoMD = cConteudoMD & " Coluna " & (j + 1) & " |"
    Next j
    cConteudoMD = cConteudoMD & vbCrLf
    
    ' Linha de marcação estrutural exigida pelo interpretador Markdown
    cLinhaSeparador = "|"
    For j = 0 To nMaxColunas
        cLinhaSeparador = cLinhaSeparador & "---| "
    Next j
    cConteudoMD = cConteudoMD & cLinhaSeparador & vbCrLf
    
    ' 5. Varre todas as linhas incluindo o índice 0 (garante que o primeiro registro apareça na tabela)
    For i = 0 To UBound(vLinhas)
        If Trim(vLinhas(i)) <> "" Then
            vColunas = Split(vLinhas(i), sDelimitador)
            cConteudoMD = cConteudoMD & "|"
            For j = 0 To nMaxColunas
                If j <= UBound(vColunas) Then
                    ' Limpa possíveis aspas residuais (" ou ') que tenham vindo no dado bruto
                    Dim sDadoLimpo As String
                    sDadoLimpo = Trim(vColunas(j))
                    If Left(sDadoLimpo, 1) = """" And Right(sDadoLimpo, 1) = """" Then
                        sDadoLimpo = Mid(sDadoLimpo, 2, Len(sDadoLimpo) - 2)
                    ElseIf Left(sDadoLimpo, 1) = "'" And Right(sDadoLimpo, 1) = "'" Then
                        sDadoLimpo = Mid(sDadoLimpo, 2, Len(sDadoLimpo) - 2)
                    End If
                    
                    cConteudoMD = cConteudoMD & " " & sDadoLimpo & " |"
                Else
                    cConteudoMD = cConteudoMD & "  |" ' Previne falhas se faltar coluna em linhas incompletas
                End If
            Next j
            cConteudoMD = cConteudoMD & vbCrLf
        End If
    Next i
    
    ' 6. Sanitização de escape para segurança de injeção JavaScript no WebView2
    cConteudoMD = Replace(cConteudoMD, vbCrLf, "\n")
    cConteudoMD = Replace(cConteudoMD, vbCr, "\n")
    cConteudoMD = Replace(cConteudoMD, vbLf, "\n")
    cConteudoMD = Replace(cConteudoMD, "'", "\'")
    
    ' 7. Montagem do HTML em blocos (Evita o limite oculto de linhas do VB6)
    cHTMLTemplate = "<html><head><meta charset='utf-8'>"
    cHTMLTemplate = cHTMLTemplate & "<script src='https://cdn.jsdelivr.net/npm/marked/marked.min.js'></script>"
    cHTMLTemplate = cHTMLTemplate & "<style>"
    cHTMLTemplate = cHTMLTemplate & "  body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Helvetica, Arial, sans-serif; padding: 35px; background-color: #ffffff; color: #24292f; line-height: 1.6; }"
    cHTMLTemplate = cHTMLTemplate & "  table { border-collapse: collapse; width: 100%; margin-top: 0; margin-bottom: 16px; font-size: 14px; }"
    cHTMLTemplate = cHTMLTemplate & "  th, td { border: 1px solid #d0d7de; padding: 8px 13px; }"
    cHTMLTemplate = cHTMLTemplate & "  th { background-color: #f6f8fa; font-weight: 600; color: #57606a; }"
    cHTMLTemplate = cHTMLTemplate & "  tr:nth-child(even) { background-color: #f6f8fa; }"
    cHTMLTemplate = cHTMLTemplate & "</style></head>"
    cHTMLTemplate = cHTMLTemplate & "<body>"
    cHTMLTemplate = cHTMLTemplate & "  <div id='content'></div>"
    cHTMLTemplate = cHTMLTemplate & "  <script>"
    cHTMLTemplate = cHTMLTemplate & "    document.getElementById('content').innerHTML = marked.parse('" & cConteudoMD & "');"
    cHTMLTemplate = cHTMLTemplate & "  </script>"
    cHTMLTemplate = cHTMLTemplate & "</body></html>"
                 
    ' 8. Renderiza a tabela limpa e formatada na janela do Chromium
    OrdoWebView1.NavigateToString cHTMLTemplate
    Exit Sub

ErroRenderDelimitado:
    MsgBox "Falha no motor RenderizarMotorDelimitadoLocal: " & Err.Description, vbCritical, "Erro de Engine"
    If fNum > 0 Then Close #fNum
End Sub
