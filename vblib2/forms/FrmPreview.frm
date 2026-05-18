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
Dim SFILTER As String
  SFILTER = "Arquivos de Textos (*.HTML)" & vbNullChar & "*.HTML" & vbNullChar & "Todos Arquivo" & vbNullChar & "*.*"
  sFileName = FileSave(Me, SFILTER, 1, "HTML", , , "Salvar HTML Como")
  FileWrite sFileName, OrdoWebView1.GetInnerHTML
End Sub

Private Sub cmdSavejpg_Click()
Dim SFILTER As String
    Dim sFileName As String
    Dim sSugestaoNome As String
    
    ' Extrai o nome do arquivo atual para sugerir na caixa de diálogo (ex: "etiqueta")
    sSugestaoNome = NomeArq(mvarCaminhoArquivo, True)
    
    ' Filtro formatado usando "|" conforme exigido pela sua função mFileAPI
    SFILTER = "Imagem JPG (*.jpg;*.jpeg)|*.jpg;*.jpeg|Todos Arquivos (*.*)|*.*"
    
    sFileName = FileSave(Me, SFILTER, 1, "JPG", sSugestaoNome, , "Salvar Etiqueta Como Imagem JPG")
    
    If sFileName <> "" Then
        Call SalvarPreviewComoImagem(sFileName)
    End If
End Sub

Private Sub cmdSavePNG_Click()
Dim SFILTER As String
    Dim sFileName As String
    Dim sSugestaoNome As String
    
    ' Extrai o nome do arquivo atual para sugerir na caixa de diálogo
    sSugestaoNome = NomeArq(mvarCaminhoArquivo, True)
    
    ' Filtro formatado usando "|" conforme exigido pela sua função mFileAPI
    SFILTER = "Imagem PNG (*.png)|*.png|Todos Arquivos (*.*)|*.*"
    
    sFileName = FileSave(Me, SFILTER, 1, "PNG", sSugestaoNome, , "Salvar Etiqueta Como Imagem PNG")
    
    If sFileName <> "" Then
        Call SalvarPreviewComoImagem(sFileName)
    End If
End Sub

Private Sub cmdSaveTXT_Click()
Dim sFileName As String
Dim SFILTER As String
  SFILTER = "Arquivos de Textos (*.TXT)" & vbNullChar & "*.TXT" & vbNullChar & "Todos Arquivo" & vbNullChar & "*.*"
  sFileName = FileSave(Me, SFILTER, 1, "HTML", , , "Salvar TXT Como")
  FileWrite sFileName, OrdoWebView1.GetInnerText
End Sub

Private Sub Encerrar_Click()
Unload Me
End Sub

Private Sub Form_Load()
 ' 1. Captura o destino passado pela variável global (pode ser "C:\arq.txt" ou "http://...")
    mvarCaminhoArquivo = Trim(CStr(cARQRTF)) '//Trim(CStr(ePASS01))
    
    ' 2. IF SIMPLES POR SUBSTRING:
    ' Se encontrar ".zpl" (seja no nome de um arquivo local ou em uma URL de etiqueta),
    ' desvia para o motor de renderização local offline.
    If InStr(1, mvarCaminhoArquivo, ".zpl", vbTextCompare) > 0 Then
        cmdSavehtml.Visible = False
        cmdSaveTXT.Visible = False
        Call RenderizarMotorZplLocal
    Else
        ' ROTA NORMAL LIVRE: Se for .html, .txt, ou qualquer URL externa, o WebView2 abre direto
        Me.Caption = "Visualizador - " & mvarCaminhoArquivo
        
        ' Se for um arquivo local, adicionamos o protocolo file:///, se for URL o Navigate resolve direto
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
    Dim cExtensao As String
    Dim cMimeType As String
    Dim cScript As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    cNomeArquivo = fso.GetFileName(cCaminhoDestino)
    cExtensao = LCase(fso.GetExtensionName(cCaminhoDestino))
    Set fso = Nothing
    
    ' Define o MimeType correto com base na escolha do utilizador
    If cExtensao = "jpg" Or cExtensao = "jpeg" Then
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

