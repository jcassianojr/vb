Attribute VB_Name = "sqlsqlite4vb"
Option Explicit

' Adaptador encapsulado para cSQLite (sqlite4vb).
' Os valores sao ligados por parametros; nomes de tabela/campo continuam
' sendo derivados do SQL legado e devem ser controlados pelo chamador.

Public Function PegUltSQLite4vb(ByVal cCON As String, ByVal cSQL As String, _
                                ByVal cCAMPO As String, _
                                ByVal eDEFAULT As Variant) As Variant
    Dim DB As cSQLite
    Dim rs As cSQLiteResults
    Dim Value As Variant

    On Error GoTo TrataErro
    Set DB = AbrirSQLite4vb(cCON, True)
    Set rs = DB.Query(cSQL)
    Do While rs.MoveNext()
        Value = rs(cCAMPO)
    Loop
    If IsNull(Value) Or IsEmpty(Value) Then Value = eDEFAULT
    PegUltSQLite4vb = Value
Saida:
    FecharSQLite4vb DB, rs
    Exit Function
TrataErro:
    PegUltSQLite4vb = eDEFAULT
    Resume Saida
End Function

Public Function PegMinSQLite4vb(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                                ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    PegMinSQLite4vb = PegOperSQLite4vb(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "MIN")
End Function

Public Function PegMaxSQLite4vb(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                                ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    PegMaxSQLite4vb = PegOperSQLite4vb(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "MAX")
End Function

Public Function PegSumSQLite4vb(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                                ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    PegSumSQLite4vb = PegOperSQLite4vb(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "SUM")
End Function

Public Function PegCountSQLite4vb(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                                  ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    PegCountSQLite4vb = PegOperSQLite4vb(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "COUNT")
End Function

Public Function PegCampoSQLite4vb(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                                  ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    PegCampoSQLite4vb = PegOperSQLite4vb(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "CAMPO")
End Function

Public Function PegOperSQLite4vb(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                                 ByVal cCAMPO As String, ByVal eDEFAULT As Variant, _
                                 ByVal coper As String) As Variant
    Dim DB As cSQLite
    Dim sql As String
    Dim Value As Variant
    Dim op As String

    On Error GoTo TrataErro
    op = UCase$(Trim$(coper))
    Select Case op
        Case "CAMPO"
            If Len(cCAMPO) = 0 Then
                sql = AliasCampoSQLite4vb(cTABLEWHERE)
            Else
                sql = "SELECT " & cCAMPO & " AS CAMPO FROM " & cTABLEWHERE
            End If
        Case "MIN", "MAX", "SUM", "COUNT"
            sql = "SELECT " & op & "(" & cCAMPO & ") AS CAMPO FROM " & cTABLEWHERE
        Case Else
            Exit Function
    End Select

    Set DB = AbrirSQLite4vb(cCON, True)
    Value = DB.Scalar(sql)
    If IsNull(Value) Or IsEmpty(Value) Then Value = eDEFAULT
    PegOperSQLite4vb = Value
Saida:
    FecharSQLite4vb DB
    Exit Function
TrataErro:
    PegOperSQLite4vb = eDEFAULT
    Resume Saida
End Function

Public Function ComandoSqlite4VB(ByVal cCON As String, ByVal cSQL As String) As Boolean
    Dim DB As cSQLite
    On Error GoTo TrataErro
    Set DB = AbrirSQLite4vb(cCON, False)
    DB.Execute DialetoSQLite4vb(cSQL)
    ComandoSqlite4VB = True
Saida:
    FecharSQLite4vb DB
    Exit Function
TrataErro:
    ComandoSqlite4VB = False
    Resume Saida
End Function

Public Function PegSQLite4vb(ByVal cCON As String, ByVal cSQL As String, _
                             ByVal nITEM As Long, ByVal aCAM As Variant, _
                             ByVal aFOR As Variant, ByVal aPAD As Variant) As Variant
    Dim DB As cSQLite
    Dim rs As cSQLiteResults
    Dim result() As Variant
    Dim i As Long
    Dim Value As Variant

    On Error GoTo TrataErro
    If nITEM <= 0 Then
        PegSQLite4vb = aPAD
        Exit Function
    End If

    ReDim result(0 To nITEM - 1)
    For i = 0 To nITEM - 1
        result(i) = ObterValorArrayOuEscalar(aPAD, i)
    Next i

    Set DB = AbrirSQLite4vb(cCON, True)
    Set rs = DB.Query(cSQL)
    If rs.MoveNext() Then
        For i = 0 To nITEM - 1
            Value = ValorCampo4vb(rs, aCAM(i))
            If IsNull(Value) Or IsEmpty(Value) Then
                result(i) = ObterValorArrayOuEscalar(aPAD, i)
            Else
                result(i) = FVar(Value, ObterValorArrayOuEscalar(aFOR, i), _
                                 ObterValorArrayOuEscalar(aPAD, i))
            End If
        Next i
    End If
    PegSQLite4vb = result
Saida:
    FecharSQLite4vb DB, rs
    Exit Function
TrataErro:
    PegSQLite4vb = aPAD
    Resume Saida
End Function

Public Function PegSQLiteDeli4vb(ByVal cCON As String, ByVal cSQL As String, _
                                 ByVal aCAM As Variant, _
                                 Optional ByVal cDELI As String = ",", _
                                 Optional ByVal aPAD As Variant = "", _
                                 Optional ByVal aFOR As Variant = "") As Variant
    Dim DB As cSQLite
    Dim rs As cSQLiteResults
    Dim result() As String
    Dim i As Long
    Dim Value As Variant
    Dim hasRow As Boolean
    Dim rowCount As Long

    On Error GoTo TrataErro
    ReDim result(LBound(aCAM) To UBound(aCAM))
    Set DB = AbrirSQLite4vb(cCON, True)
    Set rs = DB.Query(cSQL)
    Do While rs.MoveNext()
        hasRow = True
        rowCount = rowCount + 1
        For i = LBound(aCAM) To UBound(aCAM)
            Value = ValorCampo4vb(rs, aCAM(i))
            If IsNull(Value) Or IsEmpty(Value) Then
                Value = ObterValorArrayOuEscalar(aPAD, i)
            ElseIf IsArray(aFOR) Then
                Value = FVar(Value, aFOR(i), ObterValorArrayOuEscalar(aPAD, i))
            End If
            If rowCount > 1 Then
                result(i) = result(i) & cDELI
            End If
            result(i) = result(i) & FixStr(Value)
        Next i
    Loop
    If Not hasRow Then
        For i = LBound(aCAM) To UBound(aCAM)
            result(i) = CStr(ObterValorArrayOuEscalar(aPAD, i))
        Next i
    End If
    PegSQLiteDeli4vb = result
Saida:
    FecharSQLite4vb DB, rs
    Exit Function
TrataErro:
    PegSQLiteDeli4vb = result
    Resume Saida
End Function

Public Function SomaSQLite4vb(ByVal cCON As String, ByVal cSQL As String, _
                               ByVal aCAM As Variant) As Variant
    Dim DB As cSQLite
    Dim rs As cSQLiteResults
    Dim result() As Variant
    Dim i As Long
    Dim Value As Variant
    Dim nFields As Long

    On Error GoTo TrataErro
    nFields = UBound(aCAM) - LBound(aCAM) + 1
    ReDim result(LBound(aCAM) To UBound(aCAM))
    For i = LBound(aCAM) To UBound(aCAM)
        result(i) = 0
    Next i

    Set DB = AbrirSQLite4vb(cCON, True)
    Set rs = DB.Query(cSQL)
    Do While rs.MoveNext()
        For i = LBound(aCAM) To UBound(aCAM)
            Value = ValorCampo4vb(rs, aCAM(i))
            If Not IsNull(Value) And Not IsEmpty(Value) Then
                result(i) = result(i) + Value
            End If
        Next i
    Loop
    SomaSQLite4vb = result
Saida:
    FecharSQLite4vb DB, rs
    Exit Function
TrataErro:
    SomaSQLite4vb = result
    Resume Saida
End Function

Public Function GrvSQLite4vb(ByVal cARQ As String, ByVal cSQL_SELECT As String, _
                             ByVal nITEM As Long, ByVal aCAM As Variant, _
                             ByVal aVAL As Variant, ByVal aFOR As Variant, _
                             Optional ByVal nStartItem As Long = 0) As Boolean
    Dim DB As cSQLite
    Dim tableName As String, whereSql As String, sql As String
    Dim values() As Variant
    Dim i As Long, nValues As Long

    On Error GoTo TrataErro
    If nStartItem < 0 Then nStartItem = 0
    If nStartItem >= nITEM Then Exit Function
    If Not SqlExtrairUpdatePartes(cSQL_SELECT, tableName, whereSql) Then Exit Function

    nValues = nITEM - nStartItem
    ReDim values(0 To nValues - 1)
    For i = 0 To nValues - 1
        values(i) = PrepararValorSQLite(ObterValorArrayOuEscalar(aVAL, nStartItem + i), _
                                        ObterValorArrayOuEscalar(aFOR, nStartItem + i))
    Next i

    Set DB = AbrirSQLite4vb(cARQ, False)
    sql = MontarUpdate4vb(tableName, whereSql, aCAM, nStartItem, nITEM)
    DB.ExecUpdateArr tableName, whereSql, CamposUpdate4vb(aCAM, nStartItem, nITEM), values
    GrvSQLite4vb = True
Saida:
    FecharSQLite4vb DB
    Exit Function
TrataErro:
    GrvSQLite4vb = False
    Resume Saida
End Function

Public Function ExecuteScalarSQLite4vb(ByVal cARQ As String, ByVal cSQL As String) As Variant
    Dim DB As cSQLite
    On Error GoTo TrataErro
    Set DB = AbrirSQLite4vb(cARQ, True)
    ExecuteScalarSQLite4vb = DB.Scalar(cSQL)
Saida:
    FecharSQLite4vb DB
    Exit Function
TrataErro:
    ExecuteScalarSQLite4vb = Null
    Resume Saida
End Function

Public Function IncluiSQLite4vb(ByVal cARQ As String, ByVal cSQL_SELECT As String, _
                                ByVal nITEM As Long, ByVal aCAM As Variant, _
                                ByVal aVAL As Variant, ByVal lCHECK As Boolean, _
                                ByVal lMES As Boolean, ByVal aIDDES As Variant) As Boolean
    Dim DB As cSQLite
    Dim rs As cSQLiteResults
    Dim tableName As String, fields As String
    Dim values() As Variant
    Dim i As Long
    Dim existed As Boolean
    Dim insertId As Currency

    On Error GoTo TrataErro
    tableName = NomeTableSql(cSQL_SELECT)
    If Len(tableName) = 0 Then Exit Function

    Set DB = AbrirSQLite4vb(cARQ, False)
    If lCHECK Then
        Set rs = DB.Query(cSQL_SELECT)
        existed = rs.MoveNext()
        Set rs = Nothing
        If existed Then
            If lMES Then Alert "Item ja Cadastrado Com Esta Chave"
            Exit Function
        End If
    End If

    If nITEM <= 0 Then Exit Function
    ReDim values(0 To nITEM - 1)
    For i = 0 To nITEM - 1
        values(i) = PrepararValorSQLite(ObterValorArrayOuEscalar(aVAL, i), "C")
    Next i
    fields = SqlConstruirCampos(aCAM, 0, nITEM)
    insertId = DB.ExecInsertArr(tableName, fields, values)
    PreencherIDs4vb DB, tableName, aIDDES, insertId
    IncluiSQLite4vb = True
Saida:
    FecharSQLite4vb DB, rs
    Exit Function
TrataErro:
    If lMES Then SayErro "Inclui SQLite4VB: " & Err.Description
    IncluiSQLite4vb = False
    Resume Saida
End Function

Public Function PegLastidbsqLite4vb(ByVal cCON As String, ByVal cTabela As String) As Long
    Dim DB As cSQLite
    Dim Value As Variant
    On Error GoTo TrataErro
    Set DB = AbrirSQLite4vb(cCON, True)
    Value = DB.Scalar("SELECT COALESCE(MAX(rowid), 0) FROM " & cTabela)
    If Not IsNull(Value) And Not IsEmpty(Value) Then PegLastidbsqLite4vb = CLng(Value)
Saida:
    FecharSQLite4vb DB
    Exit Function
TrataErro:
    PegLastidbsqLite4vb = 0
    Resume Saida
End Function

Public Function ApagaSQLite4vb(ByVal cCON As String, ByVal cSQL As String) As Boolean
    Dim sql As String
    Dim p As Long
    Dim upperSql As String

    upperSql = UCase$(Trim$(cSQL))
    If InStr(upperSql, "WHERE") = 0 Then Exit Function
    p = InStr(upperSql, "FROM")
    If p = 0 Then Exit Function
    sql = "DELETE " & Mid$(Trim$(cSQL), p)
    ApagaSQLite4vb = ComandoSqlite4VB(cCON, sql)
End Function

Public Function SQLMoveRegSQLite4vb(ByVal cARQORI As String, _
                                    ByVal cSQLORI As String, _
                                    Optional ByVal cOPEORI As String = "", _
                                    Optional ByVal aCAMORI As Variant = 0, _
                                    Optional ByVal aOUTORI As Variant = 0, _
                                    Optional ByVal cARQDES As String = "", _
                                    Optional ByVal cSQLDES As String = "", _
                                    Optional ByVal cOPEDES As String = "", _
                                    Optional ByVal aCAMDES As Variant = 0, _
                                    Optional ByVal aOUTDES As Variant = 0, _
                                    Optional ByVal aIDDES As Variant = 0) As Boolean
    Dim Source As cSQLite, target As cSQLite
    Dim rs As cSQLiteResults, targetCheck As cSQLiteResults
    Dim tableName As String, fields As String
    Dim values() As Variant, extraValues() As Variant
    Dim i As Long, nFields As Long, nExtra As Long
    Dim insertedId As Currency
    Dim operation As String

    On Error GoTo TrataErro
    operation = UCase$(Trim$(cOPEDES))
    If InStr(1, operation, "INC", vbTextCompare) = 0 Then Exit Function
    If InStr(1, UCase$(cOPEORI), "DEL", vbTextCompare) > 0 Then
        ' A consulta legada nao fornece necessariamente a chave fisica da
        ' linha; nao apagar a origem sem uma clausula segura e explicita.
        Exit Function
    End If

    tableName = NomeTableSql(cSQLDES)
    If Len(tableName) = 0 Or Not IsArray(aCAMDES) Then Exit Function
    nFields = UBound(aCAMDES) - LBound(aCAMDES) + 1
    If IsArray(aCAMORI) Then
        If UBound(aCAMORI) - LBound(aCAMORI) + 1 <> nFields Then Exit Function
    Else
        Exit Function
    End If

    Set Source = AbrirSQLite4vb(cARQORI, True)
    Set target = AbrirSQLite4vb(cARQDES, False)
    Set rs = Source.Query(cSQLORI)
    If rs.MoveNext() Then
        ReDim values(0 To nFields - 1)
        For i = 0 To nFields - 1
            values(i) = ValorCampo4vb(rs, aCAMORI(LBound(aCAMORI) + i))
        Next i
        fields = SqlConstruirCampos(aCAMDES, 0, nFields)

        If IsArray(aOUTDES) Then
            ReDim extraValues(0 To UBound(aOUTDES) - LBound(aOUTDES))
            For i = LBound(extraValues) To UBound(extraValues)
                extraValues(i) = ObterValorArrayOuEscalar(aOUTORI, i)
            Next i
            nExtra = UBound(extraValues) - LBound(extraValues) + 1
            fields = fields & ", " & SqlConstruirCampos(aOUTDES, LBound(aOUTDES), UBound(aOUTDES) + 1)
            ReDim values(0 To nFields + nExtra - 1)
            For i = 0 To nFields - 1
                values(i) = ValorCampo4vb(rs, aCAMORI(LBound(aCAMORI) + i))
            Next i
            For i = 0 To nExtra - 1
                values(nFields + i) = extraValues(LBound(extraValues) + i)
            Next i
        End If

        insertedId = target.ExecInsertArr(tableName, fields, values)
        PreencherIDs4vb target, tableName, aIDDES, insertedId
        SQLMoveRegSQLite4vb = True
    End If
Saida:
    FecharSQLite4vb Source, rs
    FecharSQLite4vb target, targetCheck
    Exit Function
TrataErro:
    SQLMoveRegSQLite4vb = False
    Resume Saida
End Function

Public Function VBSQLiteSetValues4vb(ByVal DB As cSQLite) As Boolean
    On Error GoTo TrataErro
    If DB Is Nothing Then Exit Function
    DB.Execute "PRAGMA cache_size = 2000"
    DB.Execute "PRAGMA synchronous = NORMAL"
    DB.Execute "PRAGMA temp_store = MEMORY"
    VBSQLiteSetValues4vb = True
    Exit Function
TrataErro:
    VBSQLiteSetValues4vb = False
End Function

Private Function AbrirSQLite4vb(ByVal connection As String, ByVal readOnly As Boolean) As cSQLite
    Dim DB As cSQLite
    Set DB = New cSQLite
    DB.OpenDB LimpaTag(connection), readOnly, Not readOnly
    If Not readOnly Then
        If Not VBSQLiteSetValues4vb(DB) Then
            Err.Raise vbObjectError + 30400, "AbrirSQLite4vb", _
                      "Falha ao configurar a conexao SQLite."
        End If
    End If
    Set AbrirSQLite4vb = DB
End Function

Private Sub FecharSQLite4vb(ByRef DB As cSQLite, Optional ByVal rs As cSQLiteResults)
    On Error Resume Next
    Set rs = Nothing
    If Not DB Is Nothing Then DB.CloseDB
    Set DB = Nothing
End Sub

Private Function ValorCampo4vb(ByVal rs As cSQLiteResults, ByVal field As Variant) As Variant
    ValorCampo4vb = rs(field)
End Function


Private Function CamposUpdate4vb(ByVal fields As Variant, ByVal first As Long, _
                                 ByVal last As Long) As String
    Dim i As Long, result As String
    For i = first To last - 1
        If i > first Then result = result & ", "
        result = result & CStr(fields(i))
    Next i
    CamposUpdate4vb = result
End Function

Private Function MontarUpdate4vb(ByVal tableName As String, ByVal whereSql As String, _
                                 ByVal fields As Variant, ByVal first As Long, _
                                 ByVal last As Long) As String
    MontarUpdate4vb = "UPDATE " & tableName & " SET " & _
                      CamposUpdate4vb(fields, first, last) & " " & whereSql
End Function



Private Function AliasCampoSQLite4vb(ByVal sourceSql As String) As String
    Dim tEXT As String, p As Long
    tEXT = UCase$(sourceSql)
    p = InStr(1, tEXT, " FROM ", vbTextCompare)
    If p = 0 Then
        AliasCampoSQLite4vb = sourceSql
    Else
        AliasCampoSQLite4vb = Left$(sourceSql, p - 1) & " AS CAMPO" & Mid$(sourceSql, p)
    End If
End Function

Private Function DialetoSQLite4vb(ByVal sourceSql As String) As String
    DialetoSQLite4vb = Replace(sourceSql, "CURRENTDATETIME", "current_timestamp", 1, -1, vbTextCompare)
End Function

Private Sub PreencherIDs4vb(ByVal DB As cSQLite, ByVal tableName As String, _
                            ByVal ids As Variant, ByVal insertId As Currency)
    Dim rs As cSQLiteResults
    Dim result() As Variant
    Dim i As Long

    If Not IsArray(ids) Then
        eRETU01 = insertId
        Exit Sub
    End If

    Set rs = DB.Query("SELECT " & SqlConstruirCampos(ids, LBound(ids), UBound(ids) + 1) & " FROM " & tableName & " WHERE rowid = ?", insertId)
    If rs.MoveNext() Then
        ReDim result(LBound(ids) To UBound(ids))
        For i = LBound(ids) To UBound(ids)
            result(i) = rs(ids(i))
        Next i
        eRETU01 = result
    Else
        eRETU01 = insertId
    End If
    Set rs = Nothing
End Sub
