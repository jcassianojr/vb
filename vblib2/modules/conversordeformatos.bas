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
Public Function txttoodt(ByVal cOrigem As String, Optional ByVal cDestino As String = "") As Boolean
    Dim fso As Object, streamIn As Object
    Dim cLINHA As String
    Dim cTempDir As String, cContent As String, cManifest As String
    Dim oZip As cZipArchive
    
    If Dir(cOrigem) = "" Then Exit Function
    If Trim(cDestino) = "" Then cDestino = Replace(LCase(cOrigem), ".txt", ".odt")
    
    On Error GoTo TrataErro
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Cria pasta temporária para montar o pacote ODT
    Randomize
    cTempDir = fso.GetSpecialFolder(2) & "\OdtVB_" & Format(Int((9999 - 1000 + 1) * Rnd + 1000), "0000")
    If fso.FolderExists(cTempDir) Then fso.DeleteFolder cTempDir, True
    fso.CreateFolder cTempDir
    fso.CreateFolder cTempDir & "\META-INF"
    
    ' 1. mimetype (Sem quebra de linha)
    Dim oStream As Object
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Type = 2: oStream.Charset = "UTF-8": oStream.Open
    oStream.WriteText "application/vnd.oasis.opendocument.text"
    oStream.SaveToFile cTempDir & "\mimetype", 2
    oStream.Close
    
    ' 2. manifest.xml
    cManifest = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
                "<manifest:manifest xmlns:manifest=""urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"" manifest:version=""1.2"">" & vbCrLf & _
                " <manifest:file-entry manifest:full-path=""/"" manifest:media-type=""application/vnd.oasis.opendocument.text""/>" & vbCrLf & _
                " <manifest:file-entry manifest:full-path=""content.xml"" manifest:media-type=""text/xml""/>" & vbCrLf & _
                "</manifest:manifest>"
    fso.CreateTextFile(cTempDir & "\META-INF\manifest.xml", True).Write cManifest
    
    ' 3. content.xml com suporte a BBCode e Estilos
    cContent = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbCrLf & _
               "<office:document-content xmlns:office=""urn:oasis:names:tc:opendocument:xmlns:office:1.0"" " & _
               "xmlns:text=""urn:oasis:names:tc:opendocument:xmlns:text:1.0"" " & _
               "xmlns:style=""urn:oasis:names:tc:opendocument:xmlns:style:1.0"" " & _
               "xmlns:fo=""urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"" office:version=""1.2"">" & vbCrLf & _
               " <office:automatic-styles>" & vbCrLf & _
               "  <style:style style:name=""T_B"" style:family=""text""><style:text-properties fo:font-weight=""bold""/></style:style>" & vbCrLf & _
               "  <style:style style:name=""T_I"" style:family=""text""><style:text-properties fo:font-style=""italic""/></style:style>" & vbCrLf & _
               "  <style:style style:name=""T_U"" style:family=""text""><style:text-properties style:text-underline-style=""solid""/></style:style>" & vbCrLf & _
               "  <style:style style:name=""P_PageBreak"" style:family=""paragraph""><style:paragraph-properties fo:break-before=""page""/></style:style>" & vbCrLf & _
               " </office:automatic-styles>" & vbCrLf & _
               " <office:body><office:text>" & vbCrLf
               
    Set streamIn = fso.OpenTextFile(cOrigem, 1, False)
    Do While Not streamIn.AtEndOfStream
        cLINHA = streamIn.ReadLine
        cLINHA = ParseEscapeToBBCode(cLINHA)
        cLINHA = TiraControlChars(cLINHA)
        
        ' Limpeza e mapeamento básico XML
        cLINHA = Replace(cLINHA, "&", "&amp;")
        cLINHA = Replace(cLINHA, "<", "&lt;")
        cLINHA = Replace(cLINHA, ">", "&gt;")
        
        Dim cParStyle As String
        cParStyle = ""
        If InStr(1, cLINHA, "[PAGE]", vbTextCompare) > 0 Then
            cParStyle = " text:style-name=""P_PageBreak"""
            cLINHA = Replace(cLINHA, "[PAGE]", "", , , vbTextCompare)
        End If
        
        ' Substituições rápidas de BBCode para ODT Span
        cLINHA = Replace(cLINHA, "[B]", "<text:span text:style-name=""T_B"">", , , vbTextCompare)
        cLINHA = Replace(cLINHA, "[/B]", "</text:span>", , , vbTextCompare)
        cLINHA = Replace(cLINHA, "[I]", "<text:span text:style-name=""T_I"">", , , vbTextCompare)
        cLINHA = Replace(cLINHA, "[/I]", "</text:span>", , , vbTextCompare)
        
        cContent = cContent & "   <text:p" & cParStyle & ">" & cLINHA & "</text:p>" & vbCrLf
    Loop
    streamIn.Close
    
    cContent = cContent & "  </office:text></office:body></office:document-content>"
    
    ' Grava o content.xml sem BOM
    oStream.Open
    oStream.WriteText cContent
    oStream.Position = 0: oStream.Type = 1: oStream.Position = 3
    Dim oStreamNoBOM As Object
    Set oStreamNoBOM = CreateObject("ADODB.Stream")
    oStreamNoBOM.Type = 1: oStreamNoBOM.Open
    oStream.CopyTo oStreamNoBOM
    oStreamNoBOM.SaveToFile cTempDir & "\content.xml", 2
    oStream.Close: oStreamNoBOM.Close
    
    ' 4. Compacta usando a cZipArchive
    If fso.FileExists(cDestino) Then fso.DeleteFile cDestino, True
    Set oZip = New cZipArchive
    oZip.AddFile cTempDir & "\mimetype", "mimetype"
    oZip.AddFile cTempDir & "\content.xml", "content.xml"
    oZip.AddFile cTempDir & "\META-INF\manifest.xml", "META-INF/manifest.xml"
    oZip.CompressArchive cDestino
    
    fso.DeleteFolder cTempDir, True
    txttoodt = True
    Exit Function
TrataErro:
    txttoodt = False
End Function
Public Function txttodocx(ByVal cOrigem As String, Optional ByVal cDestino As String = "", Optional ByVal cTITULO As String = "", Optional ByVal cAUTOR As String = "") As Boolean
    Dim fso As Object, streamIn As Object
    Dim cLINHA As String
    Dim cTempDir As String, cContent As String
    Dim oZip As cZipArchive
    
    If Dir(cOrigem) = "" Then Exit Function
    If Trim(cDestino) = "" Then cDestino = Replace(LCase(cOrigem), ".txt", ".docx")
    
    On Error GoTo TrataErro
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Cria estrutura de pastas do DOCX
    Randomize
    cTempDir = fso.GetSpecialFolder(2) & "\DocxVB_" & Format(Int((9999 - 1000 + 1) * Rnd + 1000), "0000")
    If fso.FolderExists(cTempDir) Then fso.DeleteFolder cTempDir, True
    fso.CreateFolder cTempDir
    fso.CreateFolder cTempDir & "\_rels"
    fso.CreateFolder cTempDir & "\docProps"
    fso.CreateFolder cTempDir & "\word"
    fso.CreateFolder cTempDir & "\word\_rels"
    
    ' 1. [Content_Types].xml (Incluindo o mapeamento do core.xml para o autor/titulo)
    Dim cTypes As String
    cTypes = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf & _
             "<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">" & vbCrLf & _
             " <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>" & vbCrLf & _
             " <Default Extension=""xml"" ContentType=""application/xml""/>" & vbCrLf & _
             " <Override PartName=""/word/document.xml"" ContentType=""application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml""/>" & vbCrLf & _
             " <Override PartName=""/docProps/core.xml"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>" & vbCrLf & _
             "</Types>"
    fso.CreateTextFile(cTempDir & "\[Content_Types].xml", True).Write cTypes
    
    ' 2. _rels/.rels (Apontando para document.xml e core.xml)
    Dim cRels As String
    cRels = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf & _
            "<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">" & vbCrLf & _
            " <Relationship Id=""rId1"" Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"" Target=""word/document.xml""/>" & vbCrLf & _
            " <Relationship Id=""rId2"" Type=""http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"" Target=""docProps/core.xml""/>" & vbCrLf & _
            "</Relationships>"
    fso.CreateTextFile(cTempDir & "\_rels\.rels", True).Write cRels
    
    ' 3. docProps/core.xml (Metadados de Título e Autor)
    Dim cCoreProps As String
    cCoreProps = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf & _
                 "<cp:coreProperties xmlns:cp=""http://schemas.openxmlformats.org/package/2006/metadata/core-properties"" " & _
                 "xmlns:dc=""http://purl.org/dc/elements/1.1/"" xmlns:dcterms=""http://purl.org/dc/terms/"" " & _
                 "xmlns:dcmitype=""http://purl.org/dc/dcmitype/"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"">" & vbCrLf & _
                 " <dc:title>" & cTITULO & "</dc:title>" & vbCrLf & _
                 " <dc:creator>" & cAUTOR & "</dc:creator>" & vbCrLf & _
                 " <cp:lastModifiedBy>" & cAUTOR & "</cp:lastModifiedBy>" & vbCrLf & _
                 "</cp:coreProperties>"
    fso.CreateTextFile(cTempDir & "\docProps\core.xml", True).Write cCoreProps
    
    ' 4. word/document.xml (Com fonte Courier New forçada para manter alinhadas as colunas do TXT)
    cContent = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf & _
               "<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">" & vbCrLf & _
               " <w:body>" & vbCrLf
               
    Set streamIn = fso.OpenTextFile(cOrigem, 1, False)
    Do While Not streamIn.AtEndOfStream
        cLINHA = streamIn.ReadLine
        cLINHA = ParseEscapeToBBCode(cLINHA)
        cLINHA = TiraControlChars(cLINHA)
        
        cLINHA = Replace(cLINHA, "&", "&amp;")
        cLINHA = Replace(cLINHA, "<", "&lt;")
        cLINHA = Replace(cLINHA, ">", "&gt;")
        
        Dim cRunProps As String
        cRunProps = "<w:rPr><w:rFonts w:ascii=""Courier New"" w:hAnsi=""Courier New""/></w:rPr>"
        
        If InStr(1, cLINHA, "[B]", vbTextCompare) > 0 Then
            cRunProps = "<w:rPr><w:rFonts w:ascii=""Courier New"" w:hAnsi=""Courier New""/><w:b/></w:rPr>"
            cLINHA = Replace(cLINHA, "[B]", "", , , vbTextCompare)
            cLINHA = Replace(cLINHA, "[/B]", "", , , vbTextCompare)
        End If
        
        If Len(Trim(cLINHA)) = 0 Then cLINHA = " "
        
        cContent = cContent & "  <w:p><w:r>" & cRunProps & "<w:t xml:space=""preserve"">" & cLINHA & "</w:t></w:r></w:p>" & vbCrLf
    Loop
    streamIn.Close
    
    cContent = cContent & " </w:body></w:document>"
    
    Dim oStream As Object
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Type = 2: oStream.Charset = "UTF-8": oStream.Open
    oStream.WriteText cContent
    oStream.Position = 0: oStream.Type = 1: oStream.Position = 3
    Dim oStreamNoBOM As Object
    Set oStreamNoBOM = CreateObject("ADODB.Stream")
    oStreamNoBOM.Type = 1: oStreamNoBOM.Open
    oStream.CopyTo oStreamNoBOM
    oStreamNoBOM.SaveToFile cTempDir & "\word\document.xml", 2
    oStream.Close: oStreamNoBOM.Close
    
    ' 5. Compacta o DOCX usando a cZipArchive
    If fso.FileExists(cDestino) Then fso.DeleteFile cDestino, True
    Set oZip = New cZipArchive
    oZip.AddFile cTempDir & "\[Content_Types].xml", "[Content_Types].xml"
    oZip.AddFile cTempDir & "\_rels\.rels", "_rels/.rels"
    oZip.AddFile cTempDir & "\docProps\core.xml", "docProps/core.xml"
    oZip.AddFile cTempDir & "\word\document.xml", "word/document.xml"
    oZip.CompressArchive cDestino
    
    fso.DeleteFolder cTempDir, True
    txttodocx = True
    Exit Function
TrataErro:
    txttodocx = False
End Function

