Attribute VB_Name = "SqlComum"
Option Explicit

' =================================================================
' FUNÃ‡Ã•ES COMPARTILHADAS (Comuns a RC6, TC6, SQLite4VB e VBSQLite)
' =================================================================

' Substitui: ValorPadraoRC6, ValorArrayOuEscalarRC6, ValorArrayOuEscalar4vb, ValorDoArrayOuPadrao, ValorPadraoTC6
Public Function ObterValorArrayOuEscalar(ByVal vValor As Variant, ByVal nIndex As Long) As Variant
    If IsArray(vValor) Then
        ObterValorArrayOuEscalar = vValor(nIndex)
    Else
        ObterValorArrayOuEscalar = vValor
    End If
End Function

' CentralizaÃ§Ã£o do LimpaTag
Public Function LimpaTag(ByVal cCON As String) As String
    LimpaTag = Replace(Replace(Replace(Replace(cCON, "[VBSQLITE]", ""), _
        "[SQLITERC6]", ""), "[TC6SQLITE]", ""), "[SQLITE4VB]", "")
End Function

' Substitui: ValorParaGravar, ValorParaGravar4vb
Public Function PrepararValorGravar(ByVal vValue As Variant, ByVal cFormat As Variant) As Variant
    If IsNull(vValue) Or IsEmpty(vValue) Then
        PrepararValorGravar = Null
    ElseIf Len(CStr(cFormat)) > 0 And UCase$(Left$(CStr(cFormat), 1)) = "D" Then
        If IsDate(vValue) Then
            PrepararValorGravar = Format$(CDate(vValue), "yyyy-mm-dd hh:nn:ss")
        Else
            PrepararValorGravar = Null
        End If
    Else
        Select Case VarType(vValue)
            Case vbBoolean
                PrepararValorGravar = Abs(CLng(vValue))
            Case vbSingle
                PrepararValorGravar = CDbl(vValue)
            Case vbDate
                PrepararValorGravar = Format$(CDate(vValue), "yyyy-mm-dd hh:nn:ss")
            Case Else
                PrepararValorGravar = vValue
        End Select
    End If
End Function

' Substitui: ExtrairTabela4vb, ExtraiTabela, ExtrairNomeTabela
Public Function SqlExtrairTabela(ByVal cSQL As String) As String
    Dim text As String, pFrom As Long, pWhere As Long
    text = Replace(Replace(Trim$(cSQL), vbCr, " "), vbLf, " ")
    pFrom = InStr(1, UCase$(text), " FROM ", vbTextCompare)
    If pFrom = 0 Then Exit Function
    pFrom = pFrom + 6
    pWhere = InStr(pFrom, UCase$(text), " WHERE ", vbTextCompare)
    If pWhere = 0 Then
        SqlExtrairTabela = Trim$(Mid$(text, pFrom))
    Else
        SqlExtrairTabela = Trim$(Mid$(text, pFrom, pWhere - pFrom))
    End If
End Function

' Substitui: ExtrairUpdate4vb, ExtrairUpdatePartes
Public Function SqlExtrairUpdatePartes(ByVal cSQL As String, ByRef cTabela As String, ByRef cWhere As String) As Boolean
    Dim text As String, pFrom As Long, pWhere As Long
    text = Replace(Replace(Trim$(cSQL), vbCr, " "), vbLf, " ")
    pFrom = InStr(1, UCase$(text), " FROM ", vbTextCompare)
    pWhere = InStr(1, UCase$(text), " WHERE ", vbTextCompare)
    If pFrom = 0 Or pWhere = 0 Or pWhere <= pFrom Then Exit Function
    cTabela = Trim$(Mid$(text, pFrom + 6, pWhere - pFrom - 6))
    cWhere = Mid$(text, pWhere)
    SqlExtrairUpdatePartes = (Len(cTabela) > 0)
End Function

' Substitui: CamposArray4vb, CamposUpdate4vb
Public Function SqlConstruirCampos(ByVal vFields As Variant, ByVal nFirst As Long, ByVal nLast As Long) As String
    Dim i As Long, result As String
    For i = nFirst To nLast - 1
        If i > nFirst Then result = result & ", "
        result = result & CStr(vFields(i))
    Next i
    SqlConstruirCampos = result
End Function