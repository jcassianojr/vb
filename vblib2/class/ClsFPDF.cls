VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsFPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =========================================================================
' PROPRIEDADES PRIVADAS DA CLASSE (Escopo Global interno)
' =========================================================================
Private page As Long              ' Número da página ativa
Private n As Long                 ' Contador de objetos internos do PDF
Private buffer As String          ' Buffer de comandos da página atual
Private pages() As String         ' Array estrutural contendo o texto de cada página
Private w As Single               ' Largura útil da folha
Private h As Single               ' Altura útil da folha
Private k As Single               ' Fator de conversão (escala mm -> pontos gráficos)
Private font_size As Single       ' Guarda o tamanho da fonte atual da classe
Private font_family As String     ' Guarda o nome da fonte ativa da classe

' =========================================================================
' INICIALIZAÇÃO DO MOTOR DO PDF
' =========================================================================
Public Sub CreatePDF(Optional ByVal orientation As String = "P", Optional ByVal unit As String = "mm")
    page = 0
    n = 2 ' Objetos 1 e 2 são reservados para o Catálogo de Páginas estrutural
    buffer = ""
    
    ' Define o fator multiplicador matemático baseado na unidade
    If unit = "mm" Then
        k = 72 / 25.4
    Else
        k = 1
    End If
    
    ' Configuração padrão para o formato internacional de folha A4 (210mm x 297mm)
    If UCase(orientation) = "P" Then
        w = 210
        h = 297
    Else
        w = 297
        h = 210
    End If
    
    ' Inicializa o array dinâmico de páginas com um tamanho seguro base
    ReDim pages(1 To 1)
End Sub

' =========================================================================
' GERENCIAMENTO DE PÁGINAS
' =========================================================================
Public Sub AddPage()
    ' Se já houver uma página aberta anteriormente, descarrega o buffer nela
    If page > 0 Then
        pages(page) = buffer
    End If
    
    ' Incrementa o ponteiro e redimensiona o array preservando os dados anteriores
    page = page + 1
    ReDim Preserve pages(1 To page)
    buffer = ""
    
    ' Força a inicialização com a fonte padrão necessária a cada página nova
    SetFont "Helvetica", 10
End Sub

' =========================================================================
' CONFIGURAÇÃO TIPOGRÁFICA (Blindada e Opcional)
' =========================================================================
Public Sub SetFont(Optional ByVal family As String = "", Optional ByVal size As Single = 0)
    ' Declaração explícita para respeitar estritamente o Option Explicit
    Dim family_pdf As String
    
    ' Se o argumento foi enviado, atualiza o estado interno da classe
    If Len(family) > 0 Then font_family = family
    If size > 0 Then font_size = size
    
    ' Tratamento de segurança caso o método seja disparado sem nenhuma definição prévia
    If Len(font_family) = 0 Then font_family = "Helvetica"
    If font_size = 0 Then font_size = 10
    
    ' Mapeamento: O leitor PDF reconhece Helvetica como o equivalente nativo do Arial
    family_pdf = font_family
    If UCase(family_pdf) = "ARIAL" Then family_pdf = "Helvetica"
    
    ' Cospe o comando de troca de fonte para o fluxo interno do arquivo
    ' O formato PDF exige pontuação americana para decimais (Ex: 10.0), por isso o Replace
    Out "BT /F1 " & Replace(format(font_size, "0.0"), ",", ".") & " Tf ET"
End Sub

' =========================================================================
' DESENHO E IMPRESSÃO DE TEXTO
' =========================================================================
Public Sub Cell(ByVal txt As String, ByVal x_pos As Single, ByVal y_pos As Single)
    Dim y_pdf As Single
    
    ' Escapa caracteres reservados pelo motor de renderização de strings do PDF
    txt = Replace(txt, "(", "\(")
    txt = Replace(txt, ")", "\)")
    
    ' Conversão do plano cartesiano: No VB as coordenadas descem, no PDF elas sobem
    y_pdf = (h - y_pos) * k
    
    ' Insere os dados de posicionamento vetorial (Td) e o texto bruto (Tj) no fluxo
    Out "BT /F1 " & Replace(format(font_size, "0.0"), ",", ".") & " Tf " & _
        Replace(format(x_pos * k, "0.0"), ",", ".") & " " & _
        Replace(format(y_pdf, "0.0"), ",", ".") & " Td (" & txt & ") Tj ET"
End Sub

' =========================================================================
' COMPILAÇÃO E CRIAÇÃO FÍSICA DO ARQUIVO BINÁRIO
' =========================================================================
Public Sub Output(ByVal filename As String)
    ' Salva os dados da última página processada antes do encerramento
    If page > 0 Then pages(page) = buffer
    
    Dim i As Long
    Dim total_buffer As String
    Dim page_references As String
    Dim start_xref As Long
    Dim fNum As Integer
    Dim current_stream As String
    Dim object_positions() As Long
    
    ' Dimensiona a tabela de posições dinamicamente baseado no volume de páginas
    ReDim object_positions(1 To (page * 3) + 10)
    
    ' Assinatura obrigatória do cabeçalho PDF
    total_buffer = "%PDF-1.4" & vbCrLf
    
    ' 1. Criação do Objeto Raiz (Catalog)
    n = 1
    object_positions(n) = Len(total_buffer)
    total_buffer = total_buffer & n & " 0 obj" & vbCrLf & "<< /Type /Catalog /Pages 2 0 R >>" & vbCrLf & "endobj" & vbCrLf
    
    ' Compila os identificadores sequenciais de cada página cadastrada
    page_references = ""
    For i = 1 To page
        page_references = page_references & ((i * 3) + 0) & " 0 R "
    Next i
    
    ' 2. Criação do Objeto Gerenciador de Árvore (Pages)
    n = 2
    object_positions(n) = Len(total_buffer)
    total_buffer = total_buffer & n & " 0 obj" & vbCrLf & "<< /Type /Pages /Kids [" & Trim(page_references) & "] /Count " & page & " >>" & vbCrLf & "endobj" & vbCrLf
    
    ' 3. Montagem individual da estrutura física tridimensional de cada página
    For i = 1 To page
        ' Definição da Folha e apontamento de recursos
        n = (i * 3) + 0
        object_positions(n) = Len(total_buffer)
        total_buffer = total_buffer & n & " 0 obj" & vbCrLf & _
                      "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 " & Replace(format(w * k, "0.0"), ",", ".") & " " & Replace(format(h * k, "0.0"), ",", ".") & "] /Contents " & (n + 1) & " 0 R /Resources " & (n + 2) & " 0 R >>" & vbCrLf & "endobj" & vbCrLf
        
        ' Injeção do Fluxo de Texto Puro (Stream)
        n = (i * 3) + 1
        object_positions(n) = Len(total_buffer)
        current_stream = pages(i)
        total_buffer = total_buffer & n & " 0 obj" & vbCrLf & "<< /Length " & Len(current_stream) & " >>" & vbCrLf & "stream" & vbCrLf & current_stream & "endstream" & vbCrLf & "endobj" & vbCrLf
        
        ' Dicionário de Recursos com Codificação ANSI para compatibilidade de caracteres
        n = (i * 3) + 2
        object_positions(n) = Len(total_buffer)
        total_buffer = total_buffer & n & " 0 obj" & vbCrLf & "<< /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /" & font_family & " /Encoding /WinAnsiEncoding >> >> >>" & vbCrLf & "endobj" & vbCrLf
    Next i
    
    ' Armazena o byte exato de início da tabela XREF (essencial para leitores PDF rápidos)
    start_xref = Len(total_buffer)
    total_buffer = total_buffer & "xref" & vbCrLf & "0 " & (n + 1) & vbCrLf & "0000000000 65535 f " & vbCrLf
    
    ' Preenche os offsets de bytes de cada sub-objeto gerado no documento
    For i = 1 To n
        total_buffer = total_buffer & format(object_positions(i), "0000000000") & " 00000 n " & vbCrLf
    Next i
    
    ' Trailer e Finalizador de Arquivo Padrão
    total_buffer = total_buffer & "trailer" & vbCrLf & "<< /Size " & (n + 1) & " /Root 1 0 R >>" & vbCrLf
    total_buffer = total_buffer & "startxref" & vbCrLf & start_xref & vbCrLf & "%%EOF"
    
    ' Gravação física de alta performance despejada diretamente em disco
    fNum = FreeFile
    Open filename For Output As #fNum
    Print #fNum, total_buffer
    Close #fNum
End Sub

' =========================================================================
' ESCRITA NO BUFFER INTERNO
' =========================================================================
Private Sub Out(ByVal s As String)
    buffer = buffer & s & vbCrLf
End Sub

