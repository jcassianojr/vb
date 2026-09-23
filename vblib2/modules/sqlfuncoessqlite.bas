Attribute VB_Name = "SqlComum"
Option Explicit

' =================================================================
' FUNÇÕES COMPARTILHADAS (Comuns a RC6, TC6, SQLite4VB e VBSQLite)
' =================================================================

' Substitui: ValorPadraoRC6, ValorArrayOuEscalarRC6, ValorArrayOuEscalar4vb, ValorDoArrayOuPadrao, ValorPadraoTC6
Public Function ObterValorArrayOuEscalar(ByVal vValor As Variant, ByVal nIndex As Long) As Variant
    If IsArray(vValor) Then
        ObterValorArrayOuEscalar = vValor(nIndex)
    Else
        ObterValorArrayOuEscalar = vValor
    End If
End Function

' Centralização do LimpaTag
Public Function LimpaTag(ByVal cCON As String) As String
    LimpaTag = Replace(Replace(Replace(Replace(cCON, "[VBSQLITE]", ""), _
        "[SQLITERC6]", ""), "[TC6SQLITE]", ""), "[SQLITE4VB]", "")
End Function

Public Function PrepararValorSQLite(ByVal vValue As Variant, Optional ByVal cFormat As String = "") As String
    Dim vType As VbVarType
    
    ' Obtém o tipo real da variável
    vType = VarType(vValue)
    
    ' 1. Tratar Nulos e Vazios (vbNull = 1, vbEmpty = 0)
    If vType = vbNull Or vType = vbEmpty Then
        PrepararValorSQLite = "NULL"
        Exit Function
    End If
    
    ' 2. Verifica se o parâmetro cFormat exige data com hora (herdado da função antiga)
    Dim bDateWithTime As Boolean
    If Len(cFormat) > 0 And UCase$(Left$(cFormat, 1)) = "D" Then
        bDateWithTime = True
    End If

    ' 3. Tratar os demais tipos de dados usando VbVarType
    Select Case vType
        Case vbInteger, vbLong, vbByte
            ' Números inteiros não precisam de conversão de ponto flutuante
            PrepararValorSQLite = CStr(vValue)
            
        Case vbSingle, vbDouble, vbCurrency, vbDecimal
            ' Números decimais: forçar o uso de PONTO em vez de VÍRGULA
            PrepararValorSQLite = Replace(CStr(vValue), ",", ".")
            
        Case vbBoolean
            ' Booleano no SQLite é gravado como INTEGER (0 para False, 1 para True)
            ' No VBA, True é -1, então usamos Abs() para converter -1 em 1
            PrepararValorSQLite = CStr(Abs(CInt(vValue)))
            
        Case vbDate
            ' Tratamento de Datas
            If bDateWithTime Then
                PrepararValorSQLite = "'" & Format$(vValue, "yyyy-mm-dd hh:nn:ss") & "'"
            Else
                ' Inteligência extra: verifica se a data possui fração (ou seja, se contém horário)
                If CDbl(vValue) <> Int(CDbl(vValue)) Then
                    PrepararValorSQLite = "'" & Format$(vValue, "yyyy-mm-dd hh:nn:ss") & "'"
                Else
                    PrepararValorSQLite = "'" & Format$(vValue, "yyyy-mm-dd") & "'"
                End If
            End If
            
        Case vbString
            ' Textos: Dobrar aspas simples para fazer o escape correto no SQL e evitar injeção
            PrepararValorSQLite = "'" & Replace(vValue, "'", "''") & "'"
            
        Case vbObject, vbError, vbDataObject
            ' Tipos de objetos VBA não podem ser passados via String SQL diretamente
            PrepararValorSQLite = "NULL"
            
        Case Else
            ' Captura genérica para vbVariant (se contiver array) ou tipos definidos pelo usuário
            If IsArray(vValue) Then
                ' Arrays (como arrays de Bytes para BLOBs) exigem gravação via ADODB.Parameters
                ' Como string inline, definimos como NULL por segurança.
                PrepararValorSQLite = "NULL"
            ElseIf IsNumeric(vValue) Then
                PrepararValorSQLite = Replace(CStr(vValue), ",", ".")
            Else
                PrepararValorSQLite = "'" & Replace(CStr(vValue), "'", "''") & "'"
            End If
    End Select
End Function

' Substitui: ExtrairTabela4vb, ExtraiTabela, ExtrairNomeTabela
'Public Function SqlExtrairTabela(ByVal cSQL As String) As String
''    Dim text As String, pFrom As Long, pWhere As Long
''    text = Replace(Replace(Trim$(cSQL), vbCr, " "), vbLf, " ")
''    pFrom = InStr(1, UCase$(text), " FROM ", vbTextCompare)
''    If pFrom = 0 Then Exit Function
''    pFrom = pFrom + 6
''    pWhere = InStr(pFrom, UCase$(text), " WHERE ", vbTextCompare)
''    If pWhere = 0 Then
''        SqlExtrairTabela = Trim$(Mid$(text, pFrom))
''    Else
''        SqlExtrairTabela = Trim$(Mid$(text, pFrom, pWhere - pFrom))
''    End If
''End Function

' Substitui: ExtrairUpdate4vb, ExtrairUpdatePartes
Public Function SqlExtrairUpdatePartes(ByVal cSQL As String, ByRef cTabela As String, ByRef cWhere As String) As Boolean
    Dim tEXT As String, pFrom As Long, pWhere As Long
    tEXT = Replace(Replace(Trim$(cSQL), vbCr, " "), vbLf, " ")
    pFrom = InStr(1, UCase$(tEXT), " FROM ", vbTextCompare)
    pWhere = InStr(1, UCase$(tEXT), " WHERE ", vbTextCompare)
    If pFrom = 0 Or pWhere = 0 Or pWhere <= pFrom Then Exit Function
    cTabela = Trim$(Mid$(tEXT, pFrom + 6, pWhere - pFrom - 6))
    cWhere = Mid$(tEXT, pWhere)
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
