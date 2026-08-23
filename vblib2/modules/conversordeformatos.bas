Attribute VB_Name = "conversordeformatos"
' +--------------------------------------------------------------------
' + Parsers Universais de BBCode (Semelhantes ao Harbour)
' +--------------------------------------------------------------------
Public Function ParseEscapeToBBCode(ByVal cLINHA As String) As String
    ' 1. Epson / Genéricas Matriciais
    cLINHA = Replace(cLINHA, Chr(27) & Chr(69), "[B]")
    cLINHA = Replace(cLINHA, Chr(27) & Chr(70), "[/B]")
    cLINHA = Replace(cLINHA, Chr(27) & Chr(45) & Chr(0), "[U]")
    cLINHA = Replace(cLINHA, Chr(27) & Chr(45) & Chr(1), "[/U]")
    cLINHA = Replace(cLINHA, Chr(27) & Chr(52), "[I]")
    cLINHA = Replace(cLINHA, Chr(27) & Chr(53), "[/I]")
    
    ' Tamanhos
    cLINHA = Replace(cLINHA, Chr(14) & Chr(15), "[SIZE=12][B]")
    cLINHA = Replace(cLINHA, Chr(15), "[SIZE=8]")
    cLINHA = Replace(cLINHA, Chr(18), "[SIZE=12]")
    cLINHA = Replace(cLINHA, Chr(14), "[SIZE=14]")
    cLINHA = Replace(cLINHA, Chr(20), "[SIZE=12]")
    
    ' 2. HP Deskjet / Laser
    cLINHA = Replace(cLINHA, Chr(27) & Chr(40) & Chr(115) & "16" & Chr(72), "[SIZE=12]")
    cLINHA = Replace(cLINHA, Chr(27) & Chr(40) & Chr(115) & "12" & Chr(72), "[SIZE=8]")
    cLINHA = Replace(cLINHA, Chr(27) & Chr(40) & Chr(115) & "23" & Chr(72), "[SIZE=14]")
    cLINHA = Replace(cLINHA, Chr(27) & Chr(40) & Chr(115) & Chr(51) & Chr(66), "[B]")
    cLINHA = Replace(cLINHA, Chr(27) & Chr(40) & Chr(115) & Chr(45) & Chr(51) & Chr(66), "[/B]")
    cLINHA = Replace(cLINHA, Chr(27) & Chr(38) & Chr(107) & Chr(50) & Chr(83), "[SIZE=12]")
    cLINHA = Replace(cLINHA, Chr(27) & Chr(38) & Chr(107) & Chr(48) & Chr(83), "[SIZE=8]")
    cLINHA = Replace(cLINHA, Chr(27) & Chr(40) & Chr(115) & Chr(48) & Chr(66), "[/B]")

    ' 3. Lexmark
    cLINHA = Replace(cLINHA, Chr(27) & Chr(38) & Chr(107) & Chr(50) & Chr(83) & Chr(27) & Chr(38) & Chr(108) & "8" & Chr(68) & Chr(27) & Chr(38) & Chr(108) & "90" & Chr(80), "[SIZE=12]")
    cLINHA = Replace(cLINHA, Chr(27) & Chr(38) & Chr(107) & Chr(52) & Chr(83) & Chr(27) & Chr(38) & Chr(108) & "5" & Chr(68) & Chr(27) & Chr(38) & Chr(108) & "66" & Chr(80), "[SIZE=8]")
    cLINHA = Replace(cLINHA, Chr(27) & Chr(40) & Chr(115) & "23" & Chr(72) & Chr(27) & Chr(38) & Chr(108) & "10" & Chr(68) & Chr(27) & Chr(38) & Chr(108) & "90" & Chr(80), "[SIZE=14]")

    ' 4. Comandos de Estrutura
    cLINHA = Replace(cLINHA, Chr(12), "[PAGE]")
    cLINHA = Replace(cLINHA, "##page##", "[PAGE]", , , vbTextCompare)
    
    ParseEscapeToBBCode = cLINHA
End Function

Public Function ParseBBCodeToHTML(ByVal cLINHA As String) As String
    Dim cResult As String
    cResult = cLINHA
    
    ' Tags diretas
    cResult = Replace(cResult, "[B]", "<b>", , , vbTextCompare)
    cResult = Replace(cResult, "[/B]", "</b>", , , vbTextCompare)
    cResult = Replace(cResult, "[I]", "<i>", , , vbTextCompare)
    cResult = Replace(cResult, "[/I]", "</i>", , , vbTextCompare)
    cResult = Replace(cResult, "[U]", "<u>", , , vbTextCompare)
    cResult = Replace(cResult, "[/U]", "</u>", , , vbTextCompare)
    cResult = Replace(cResult, "[PAGE]", "<div style=""page-break-after: always;""></div>", , , vbTextCompare)
    
    ' Regex ou Replace Dinamico para Tamanhos (Simulando o Parser)
    cResult = Replace(cResult, "[SIZE=8]", "<font size=""1"">", , , vbTextCompare)
    cResult = Replace(cResult, "[SIZE=12]", "<font size=""2"">", , , vbTextCompare)
    cResult = Replace(cResult, "[SIZE=14]", "<font size=""4"">", , , vbTextCompare)
    cResult = Replace(cResult, "[/SIZE]", "</font>", , , vbTextCompare)
    
    ParseBBCodeToHTML = cResult
End Function

Public Function ParseBBCodeToRTF(ByVal cLINHA As String) As String
    Dim cResult As String
    cResult = cLINHA
    cResult = Replace(cResult, "[B]", "\b ", , , vbTextCompare)
    cResult = Replace(cResult, "[/B]", "\b0 ", , , vbTextCompare)
    cResult = Replace(cResult, "[I]", "\i ", , , vbTextCompare)
    cResult = Replace(cResult, "[/I]", "\i0 ", , , vbTextCompare)
    cResult = Replace(cResult, "[U]", "\ul ", , , vbTextCompare)
    cResult = Replace(cResult, "[/U]", "\ulnone ", , , vbTextCompare)
    cResult = Replace(cResult, "[PAGE]", "\page ", , , vbTextCompare)
    
    ' O RTF calcula fontes em meio-ponto
    cResult = Replace(cResult, "[SIZE=8]", "\fs16 ", , , vbTextCompare)
    cResult = Replace(cResult, "[SIZE=12]", "\fs24 ", , , vbTextCompare)
    cResult = Replace(cResult, "[SIZE=14]", "\fs28 ", , , vbTextCompare)
    cResult = Replace(cResult, "[/SIZE]", "\fs20 ", , , vbTextCompare) ' volta ao padrao
    
    ParseBBCodeToRTF = cResult
End Function

' Um limpador exclusivo para TXT. Tira lixo, mas NÃO DESTRÓI os colchetes do BBCode!
Public Function TiraControlChars(ByVal texto As String) As String
    Dim x As Integer
    For x = 0 To 31
        If x <> 10 And x <> 13 Then
            texto = Replace(texto, Chr(x), "")
        End If
    Next
    TiraControlChars = texto
End Function


Public Function txttopdf(ByVal cOrigem As String, Optional ByVal cDestino As String = "", Optional ByVal cTITULO As String = "", Optional ByVal cAUTOR As String = "") As Boolean
    Dim fso As Object, streamIn As Object, pdf As ClsFPDF
    Dim cLINHA As String
    Dim nMargemEsquerda As Single, nLinhaAtual As Single
    Dim nAlturaLinha As Single, nLimiteInferior As Single
    
    If Not FileConnExist(cOrigem, True) Then Exit Function
    If Len(cDestino) = 0 Then cDestino = TrocaExt(cOrigem, "PDF")
    If FileConnExist(cDestino, False) Then
        Alert ("Arquivo Destino Ja existe")
        Exit Function
    End If
    
    On Error GoTo TrataErro
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set streamIn = fso.OpenTextFile(cOrigem, 1, False)
    
    Set pdf = New ClsFPDF
    pdf.CreatePDF "P", "mm"
    pdf.AddPage
    
    nMargemEsquerda = 15: nLinhaAtual = 15
    nAlturaLinha = 5: nLimiteInferior = 275
    
    Do While Not streamIn.AtEndOfStream
        cLINHA = streamIn.ReadLine
        
        ' 1. Converte lixo HP/Epson em [B], [PAGE], etc.
        cLINHA = ParseEscapeToBBCode(cLINHA)
        ' 2. Limpa resto dos binarios preservando a formatação BBCode recém criada!
        cLINHA = TiraControlChars(cLINHA)
        
        ' 3. Aplica logica de salto de página
        If InStr(1, cLINHA, "[PAGE]", vbTextCompare) > 0 Then
            pdf.AddPage
            nLinhaAtual = 15
        Else
            ' Chama o novo metodo de parsing direto no motor do PDF
            pdf.PrintBBCodeLine cLINHA, nMargemEsquerda, nLinhaAtual
            
            nLinhaAtual = nLinhaAtual + nAlturaLinha
            If nLinhaAtual > nLimiteInferior Then
                pdf.AddPage
                nLinhaAtual = 15
            End If
        End If
    Loop
    
    pdf.Output cDestino
    txttopdf = True
Fim:
    If Not streamIn Is Nothing Then streamIn.Close
    Set streamIn = Nothing: Set fso = Nothing: Set pdf = Nothing
    Exit Function
TrataErro:
    Alert "Erro na geração do PDF: " & Err.Description
    Resume Fim
End Function

Public Function txttohtml(ByVal cOrigem As String, Optional ByVal cDestino As String = "", Optional ByVal cTITULO As String = "", Optional ByVal cAUTOR As String = "") As Boolean
    Dim fso As Object, streamIn As Object, streamOut As Object
    Dim cLINHA As String
    If Not FileConnExist(cOrigem, True) Then Exit Function
    If Len(cDestino) = 0 Then cDestino = TrocaExt(cOrigem, "HTML")
    If Len(cAUTOR) = 0 Then cAUTOR = "Sistema"
    If Len(cTITULO) = 0 Then cTITULO = NomeArq(cOrigem, True)
    
    On Error GoTo TrataErro
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set streamIn = fso.OpenTextFile(cOrigem, 1, False)
    Set streamOut = fso.CreateTextFile(cDestino, True)
    
    streamOut.Write "<html><head><title>" & cTITULO & "</title></head><body><pre>" & vbCrLf
    
    Do While Not streamIn.AtEndOfStream
        cLINHA = streamIn.ReadLine
        cLINHA = ParseEscapeToBBCode(cLINHA)
        cLINHA = TiraControlChars(cLINHA)
        cLINHA = str2html(cLINHA)
        cLINHA = ParseBBCodeToHTML(cLINHA) ' Aplica estilos HTML <b>
        
        streamOut.Write cLINHA & vbCrLf
    Loop
    streamOut.Write "</pre></body></html>"
    txttohtml = True
Fim:
    If Not streamIn Is Nothing Then streamIn.Close
    If Not streamOut Is Nothing Then streamOut.Close
    Set streamIn = Nothing: Set streamOut = Nothing: Set fso = Nothing
    Exit Function
TrataErro:
    Resume Fim
End Function

Public Function txttortf(ByVal cOrigem As String, Optional ByVal cDestino As String = "", Optional ByVal cTITULO As String = "", Optional ByVal cAUTOR As String = "") As Boolean
    Dim fso As Object, streamIn As Object, streamOut As Object
    Dim cLINHA As String
    If Not FileConnExist(cOrigem, True) Then Exit Function
    If Len(cDestino) = 0 Then cDestino = TrocaExt(cOrigem, "RTF")
    
    On Error GoTo TrataErro
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set streamIn = fso.OpenTextFile(cOrigem, 1, False)
    Set streamOut = fso.CreateTextFile(cDestino, True)
    
    streamOut.Write "{\rtf1\ansi\ansicpg1252\deff0\deflang1031{\fonttbl{\f0\fmodern\fprq1 Courier New;}}\viewkind4\uc1\pard\f0\fs20 "
    
    Do While Not streamIn.AtEndOfStream
        cLINHA = streamIn.ReadLine
        cLINHA = ParseEscapeToBBCode(cLINHA)
        cLINHA = TiraControlChars(cLINHA)
        
        ' Escapes estruturais do RTF
        If InStr(cLINHA, "\") > 0 Then cLINHA = Replace(cLINHA, "\", "\\")
        If InStr(cLINHA, "{") > 0 Then cLINHA = Replace(cLINHA, "{", "\{")
        If InStr(cLINHA, "}") > 0 Then cLINHA = Replace(cLINHA, "}", "\}")
        
        cLINHA = ParseBBCodeToRTF(cLINHA) ' Aplica estilos RTF \b
        
        streamOut.Write cLINHA & "\par " & vbCrLf
    Loop
    streamOut.Write "\par}"
    txttortf = True
Fim:
    If Not streamIn Is Nothing Then streamIn.Close
    If Not streamOut Is Nothing Then streamOut.Close
    Set streamIn = Nothing: Set streamOut = Nothing: Set fso = Nothing
    Exit Function
TrataErro:
    Resume Fim
End Function

Public Function txttodoc(ByVal cOrigem As String, Optional ByVal cDestino As String = "", Optional ByVal cTITULO As String = "", Optional ByVal cAUTOR As String = "") As Boolean
    Dim fso As Object, streamIn As Object, streamOut As Object
    Dim cLINHA As String, strOutputFile As String
    If Dir(cOrigem) = "" Then Exit Function
    If Trim(cDestino) = "" Then strOutputFile = Replace(LCase(cOrigem), ".txt", ".doc") Else strOutputFile = cDestino
    
    On Error GoTo TrataErro
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set streamIn = fso.OpenTextFile(cOrigem, 1, False)
    Set streamOut = fso.OpenTextFile(strOutputFile, 2, True)
    
    streamOut.WriteLine "<html><head><style>body { font-family: 'Courier New', Courier, monospace; font-size: 10pt; line-height: 1.2; } p { margin: 0; padding: 0; white-space: pre; }</style></head><body>"
    
    Do While Not streamIn.AtEndOfStream
        cLINHA = streamIn.ReadLine
        cLINHA = ParseEscapeToBBCode(cLINHA)
        cLINHA = TiraControlChars(cLINHA)
        
        cLINHA = Replace(cLINHA, "&", "&amp;")
        cLINHA = Replace(cLINHA, "<", "&lt;")
        cLINHA = Replace(cLINHA, ">", "&gt;")
        
        cLINHA = ParseBBCodeToHTML(cLINHA)
        
        If Len(Trim(cLINHA)) = 0 Then streamOut.WriteLine "<p>&nbsp;</p>" Else streamOut.WriteLine "<p>" & cLINHA & "</p>"
    Loop
    streamOut.WriteLine "</body></html>"
    txttodoc = True
Fim:
    If Not streamIn Is Nothing Then streamIn.Close
    If Not streamOut Is Nothing Then streamOut.Close
    Set streamIn = Nothing: Set streamOut = Nothing: Set fso = Nothing
    Exit Function
TrataErro:
    Resume Fim
End Function
