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

Private page As Long
Private n As Long
Private buffer As String
Private pages() As String
Private w As Single
Private h As Single
Private k As Single
Private font_size As Single

Public Sub CreatePDF(Optional ByVal orientation As String = "P", Optional ByVal unit As String = "mm")
    page = 0
    n = 2
    buffer = ""
    If unit = "mm" Then k = 72 / 25.4 Else k = 1
    
    If UCase(orientation) = "P" Then
        w = 210: h = 297
    Else
        w = 297: h = 210
    End If
    ReDim pages(1 To 1)
End Sub

Public Sub AddPage()
    If page > 0 Then pages(page) = buffer
    page = page + 1
    ReDim Preserve pages(1 To page)
    buffer = ""
    font_size = 10
End Sub

Public Sub SetFont(Optional ByVal family As String = "", Optional ByVal size As Single = 0)
    If size > 0 Then font_size = size
End Sub

' Compatibilidade antiga mantida (chama o write normal em fonte padrao)
Public Sub Cell(ByVal txt As String, ByVal x_pos As Single, ByVal y_pos As Single)
    WriteFragment txt, x_pos, y_pos, False, False, font_size
End Sub

' =========================================================================
' O MOTOR DE BBCODE DO PDF (Move o X horizontalmente conforme imprime!)
' =========================================================================
Public Sub PrintBBCodeLine(ByVal cLinha As String, ByVal x_pos As Single, ByVal y_pos As Single)
    Dim nPosIni As Long, nPosFim As Long
    Dim cText As String, cCmd As String, cTag As String, cVal As String
    Dim lBold As Boolean, lItalic As Boolean
    Dim currentX As Single
    Dim nCurrentSize As Single
    
    currentX = x_pos
    nCurrentSize = font_size
    If nCurrentSize = 0 Then nCurrentSize = 10
    
    Do While Len(cLinha) > 0
        nPosIni = InStr(1, cLinha, "[")
        If nPosIni > 0 Then
            If nPosIni > 1 Then
                cText = Left(cLinha, nPosIni - 1)
                currentX = WriteFragment(cText, currentX, y_pos, lBold, lItalic, nCurrentSize)
            End If
            
            nPosFim = InStr(nPosIni, cLinha, "]")
            If nPosFim > 0 Then
                cTag = Mid(cLinha, nPosIni + 1, nPosFim - nPosIni - 1)
                cCmd = UCase(cTag)
                
                If Left(cCmd, 5) = "SIZE=" Then
                    cVal = Mid(cCmd, 6)
                    
                    ' CORRECAO: Bloco If estruturado multiplas linhas
                    If cVal = "8" Then
                        nCurrentSize = 7
                    ElseIf cVal = "12" Then
                        nCurrentSize = 10
                    ElseIf cVal = "14" Then
                        nCurrentSize = 12
                    Else
                        nCurrentSize = Val(cVal)
                    End If
                    
                    cCmd = "SIZE"
                End If
                
                Select Case cCmd
                    Case "B": lBold = True
                    Case "/B": lBold = False
                    Case "I": lItalic = True
                    Case "/I": lItalic = False
                    Case "/SIZE": nCurrentSize = font_size
                End Select
                
                cLinha = Mid(cLinha, nPosFim + 1)
            Else
                currentX = WriteFragment(cLinha, currentX, y_pos, lBold, lItalic, nCurrentSize)
                cLinha = ""
            End If
        Else
            currentX = WriteFragment(cLinha, currentX, y_pos, lBold, lItalic, nCurrentSize)
            cLinha = ""
        End If
    Loop
End Sub

' Metodo Interno que lida com o Avanco do Cursor (Exclusivo Courier)
Private Function WriteFragment(ByVal txt As String, ByVal x_pos As Single, ByVal y_pos As Single, lBold As Boolean, lItalic As Boolean, nSize As Single) As Single
    If txt = "" Then
        WriteFragment = x_pos
        Exit Function
    End If
    
    Dim fObj As String
    
    ' CORRECAO: Bloco If estruturado multiplas linhas
    If lBold And lItalic Then
        fObj = "/F4"
    ElseIf lBold Then
        fObj = "/F2"
    ElseIf lItalic Then
        fObj = "/F3"
    Else
        fObj = "/F1"
    End If
    
    txt = Replace(txt, "(", "\(")
    txt = Replace(txt, ")", "\)")
    
    Dim y_pdf As Single
    y_pdf = (h - y_pos) * k
    
    Out "BT " & fObj & " " & Replace(Format(nSize, "0.0"), ",", ".") & " Tf " & _
        Replace(Format(x_pos * k, "0.0"), ",", ".") & " " & _
        Replace(Format(y_pdf, "0.0"), ",", ".") & " Td (" & txt & ") Tj ET"
        
    ' A fonte Courier é monoespaçada! Na norma PDF, cada caractere ocupa exatamente 0.6 do próprio tamanho.
    WriteFragment = x_pos + ((Len(txt) * nSize * 0.6) / k)
End Function
Public Sub Output(ByVal filename As String)
    If page > 0 Then pages(page) = buffer
    
    Dim i As Long, fNum As Integer
    Dim total_buffer As String, page_references As String, current_stream As String
    Dim start_xref As Long
    Dim object_positions() As Long
    ReDim object_positions(1 To (page * 3) + 10)
    
    total_buffer = "%PDF-1.4" & vbCrLf
    n = 1
    object_positions(n) = Len(total_buffer)
    total_buffer = total_buffer & n & " 0 obj" & vbCrLf & "<< /Type /Catalog /Pages 2 0 R >>" & vbCrLf & "endobj" & vbCrLf
    
    For i = 1 To page: page_references = page_references & ((i * 3) + 0) & " 0 R ": Next i
    
    n = 2
    object_positions(n) = Len(total_buffer)
    total_buffer = total_buffer & n & " 0 obj" & vbCrLf & "<< /Type /Pages /Kids [" & Trim(page_references) & "] /Count " & page & " >>" & vbCrLf & "endobj" & vbCrLf
    
    For i = 1 To page
        n = (i * 3) + 0
        object_positions(n) = Len(total_buffer)
        total_buffer = total_buffer & n & " 0 obj" & vbCrLf & _
                      "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 " & Replace(Format(w * k, "0.0"), ",", ".") & " " & Replace(Format(h * k, "0.0"), ",", ".") & "] /Contents " & (n + 1) & " 0 R /Resources " & (n + 2) & " 0 R >>" & vbCrLf & "endobj" & vbCrLf
        
        n = (i * 3) + 1
        object_positions(n) = Len(total_buffer)
        current_stream = pages(i)
        total_buffer = total_buffer & n & " 0 obj" & vbCrLf & "<< /Length " & Len(current_stream) & " >>" & vbCrLf & "stream" & vbCrLf & current_stream & "endstream" & vbCrLf & "endobj" & vbCrLf
        
        n = (i * 3) + 2
        object_positions(n) = Len(total_buffer)
        ' >>> AQUI FOI INJETADO O SUPORTE GLOBAL ÀS 4 VARIAÇÕES DA COURIER <<<
        total_buffer = total_buffer & n & " 0 obj" & vbCrLf & _
            "<< /Font << " & _
            "/F1 << /Type /Font /Subtype /Type1 /BaseFont /Courier /Encoding /WinAnsiEncoding >> " & _
            "/F2 << /Type /Font /Subtype /Type1 /BaseFont /Courier-Bold /Encoding /WinAnsiEncoding >> " & _
            "/F3 << /Type /Font /Subtype /Type1 /BaseFont /Courier-Oblique /Encoding /WinAnsiEncoding >> " & _
            "/F4 << /Type /Font /Subtype /Type1 /BaseFont /Courier-BoldOblique /Encoding /WinAnsiEncoding >> " & _
            ">> >>" & vbCrLf & "endobj" & vbCrLf
    Next i
    
    start_xref = Len(total_buffer)
    total_buffer = total_buffer & "xref" & vbCrLf & "0 " & (n + 1) & vbCrLf & "0000000000 65535 f " & vbCrLf
    For i = 1 To n: total_buffer = total_buffer & Format(object_positions(i), "0000000000") & " 00000 n " & vbCrLf: Next i
    
    total_buffer = total_buffer & "trailer" & vbCrLf & "<< /Size " & (n + 1) & " /Root 1 0 R >>" & vbCrLf
    total_buffer = total_buffer & "startxref" & vbCrLf & start_xref & vbCrLf & "%%EOF"
    
    fNum = FreeFile
    Open filename For Output As #fNum: Print #fNum, total_buffer: Close #fNum
End Sub

Private Sub Out(ByVal s As String)
    buffer = buffer & s & vbCrLf
End Sub

