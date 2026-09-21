Attribute VB_Name = "sqlsqlite4vb"
Option Explicit

' Adaptador encapsulado para cSQLite (sqlite4vb).
' Os valores sao ligados por parametros; nomes de tabela/campo continuam
' sendo derivados do SQL legado e devem ser controlados pelo chamador.

Public Function PegUltSQLite4vb(ByVal cCON As String, ByVal cSQL As String, _
                                ByVal cCAMPO As String, _
                                ByVal eDEFAULT As Variant) As Variant
    Dim db As cSQLite
    Dim rs As cSQLiteResults
    Dim value As Variant

    On Error GoTo TrataErro
    Set db = AbrirSQLite4vb(cCON, True)
    Set rs = db.Query(cSQL)
    Do While rs.MoveNext()
        value = rs(cCAMPO)
    Loop
    If IsNull(value) Or IsEmpty(value) Then value = eDEFAULT
    PegUltSQLite4vb = value
Saida:
    FecharSQLite4vb db, rs
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
    Dim db As cSQLite
    Dim sql As String
    Dim value As Variant
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

    Set db = AbrirSQLite4vb(cCON, True)
    value = db.Scalar(sql)
    If IsNull(value) Or IsEmpty(value) Then value = eDEFAULT
    PegOperSQLite4vb = value
Saida:
    FecharSQLite4vb db
    Exit Function
TrataErro:
    PegOperSQLite4vb = eDEFAULT
    Resume Saida
End Function

Public Function ComandoSqlite4VB(ByVal cCON As String, ByVal cSQL As String) As Boolean
    Dim db As cSQLite
    On Error GoTo TrataErro
    Set db = AbrirSQLite4vb(cCON, False)
    db.Execute DialetoSQLite4vb(cSQL)
    ComandoSqlite4VB = True
Saida:
    FecharSQLite4vb db
    Exit Function
TrataErro:
    ComandoSqlite4VB = False
    Resume Saida
End Function

Public Function PegSQLite4vb(ByVal cCON As String, ByVal cSQL As String, _
                             ByVal nITEM As Long, ByVal aCAM As Variant, _
                             ByVal aFOR As Variant, ByVal aPAD As Variant) As Variant
    Dim db As cSQLite
    Dim rs As cSQLiteResults
    Dim result() As Variant
    Dim i As Long
    Dim value As Variant

    On Error GoTo TrataErro
    If nITEM <= 0 Then
        PegSQLite4vb = aPAD
        Exit Function
    End If

    ReDim result(0 To nITEM - 1)
    For i = 0 To nITEM - 1
        result(i) = ValorArrayOuEscalar4vb(aPAD, i)
    Next i

    Set db = AbrirSQLite4vb(cCON, True)
    Set rs = db.Query(cSQL)
    If rs.MoveNext() Then
        For i = 0 To nITEM - 1
            value = ValorCampo4vb(rs, aCAM(i))
            If IsNull(value) Or IsEmpty(value) Then
                result(i) = ValorArrayOuEscalar4vb(aPAD, i)
            Else
                result(i) = FVar(value, ValorArrayOuEscalar4vb(aFOR, i), _
                                 ValorArrayOuEscalar4vb(aPAD, i))
            End If
        Next i
    End If
    PegSQLite4vb = result
Saida:
    FecharSQLite4vb db, rs
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
    Dim db As cSQLite
    Dim rs As cSQLiteResults
    Dim result() As String
    Dim i As Long
    Dim value As Variant
    Dim hasRow As Boolean
    Dim rowCount As Long

    On Error GoTo TrataErro
    ReDim result(LBound(aCAM) To UBound(aCAM))
    Set db = AbrirSQLite4vb(cCON, True)
    Set rs = db.Query(cSQL)
    Do While rs.MoveNext()
        hasRow = True
        rowCount = rowCount + 1
        For i = LBound(aCAM) To UBound(aCAM)
            value = ValorCampo4vb(rs, aCAM(i))
            If IsNull(value) Or IsEmpty(value) Then
                value = ValorArrayOuEscalar4vb(aPAD, i)
            ElseIf IsArray(aFOR) Then
                value = FVar(value, aFOR(i), ValorArrayOuEscalar4vb(aPAD, i))
            End If
            If rowCount > 1 Then
                result(i) = result(i) & cDELI
            End If
            result(i) = result(i) & FixStr(value)
        Next i
    Loop
    If Not hasRow Then
        For i = LBound(aCAM) To UBound(aCAM)
            result(i) = CStr(ValorArrayOuEscalar4vb(aPAD, i))
        Next i
    End If
    PegSQLiteDeli4vb = result
Saida:
    FecharSQLite4vb db, rs
    Exit Function
TrataErro:
    PegSQLiteDeli4vb = result
    Resume Saida
End Function

Public Function SomaSQLite4vb(ByVal cCON As String, ByVal cSQL As String, _
                               ByVal aCAM As Variant) As Variant
    Dim db As cSQLite
    Dim rs As cSQLiteResults
    Dim result() As Variant
    Dim i As Long
    Dim value As Variant
    Dim nFields As Long

    On Error GoTo TrataErro
    nFields = UBound(aCAM) - LBound(aCAM) + 1
    ReDim result(LBound(aCAM) To UBound(aCAM))
    For i = LBound(aCAM) To UBound(aCAM)
        result(i) = 0
    Next i

    Set db = AbrirSQLite4vb(cCON, True)
    Set rs = db.Query(cSQL)
    Do While rs.MoveNext()
        For i = LBound(aCAM) To UBound(aCAM)
            value = ValorCampo4vb(rs, aCAM(i))
            If Not IsNull(value) And Not IsEmpty(value) Then
                result(i) = result(i) + value
            End If
        Next i
    Loop
    SomaSQLite4vb = result
Saida:
    FecharSQLite4vb db, rs
    Exit Function
TrataErro:
    SomaSQLite4vb = result
    Resume Saida
End Function

Public Function GrvSQLite4vb(ByVal cARQ As String, ByVal cSQL_SELECT As String, _
                             ByVal nITEM As Long, ByVal aCAM As Variant, _
                             ByVal aVAL As Variant, ByVal aFOR As Variant, _
                             Optional ByVal nStartItem As Long = 0) As Boolean
    Dim db As cSQLite
    Dim tableName As String, whereSql As String, sql As String
    Dim values() As Variant
    Dim i As Long, nValues As Long

    On Error GoTo TrataErro
    If nStartItem < 0 Then nStartItem = 0
    If nStartItem >= nITEM Then Exit Function
    If Not ExtrairUpdate4vb(cSQL_SELECT, tableName, whereSql) Then Exit Function

    nValues = nITEM - nStartItem
    ReDim values(0 To nValues - 1)
    For i = 0 To nValues - 1
        values(i) = ValorParaGravar4vb(ValorArrayOuEscalar4vb(aVAL, nStartItem + i), _
                                        ValorArrayOuEscalar4vb(aFOR, nStartItem + i))
    Next i

    Set db = AbrirSQLite4vb(cARQ, False)
    sql = MontarUpdate4vb(tableName, whereSql, aCAM, nStartItem, nITEM)
    db.ExecUpdateArr tableName, whereSql, CamposUpdate4vb(aCAM, nStartItem, nITEM), values
    GrvSQLite4vb = True
Saida:
    FecharSQLite4vb db
    Exit Function
TrataErro:
    GrvSQLite4vb = False
    Resume Saida
End Function

Public Function ExecuteScalarSQLite4vb(ByVal cARQ As String, ByVal cSQL As String) As Variant
    Dim db As cSQLite
    On Error GoTo TrataErro
    Set db = AbrirSQLite4vb(cARQ, True)
    ExecuteScalarSQLite4vb = db.Scalar(cSQL)
Saida:
    FecharSQLite4vb db
    Exit Function
TrataErro:
    ExecuteScalarSQLite4vb = Null
    Resume Saida
End Function

Public Function IncluiSQLite4vb(ByVal cARQ As String, ByVal cSQL_SELECT As String, _
                                ByVal nITEM As Long, ByVal aCAM As Variant, _
                                ByVal aVAL As Variant, ByVal lCHECK As Boolean, _
                                ByVal lMES As Boolean, ByVal aIDDES As Variant) As Boolean
    Dim db As cSQLite
    Dim rs As cSQLiteResults
    Dim tableName As String, fields As String
    Dim values() As Variant
    Dim i As Long
    Dim existed As Boolean
    Dim insertId As Currency

    On Error GoTo TrataErro
    tableName = ExtrairTabela4vb(cSQL_SELECT)
    If Len(tableName) = 0 Then Exit Function

    Set db = AbrirSQLite4vb(cARQ, False)
    If lCHECK Then
        Set rs = db.Query(cSQL_SELECT)
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
        values(i) = ValorParaGravar4vb(ValorArrayOuEscalar4vb(aVAL, i), "C")
    Next i
    fields = CamposArray4vb(aCAM, nITEM)
    insertId = db.ExecInsertArr(tableName, fields, values)
    PreencherIDs4vb db, tableName, aIDDES, insertId
    IncluiSQLite4vb = True
Saida:
    FecharSQLite4vb db, rs
    Exit Function
TrataErro:
    If lMES Then SayErro "Inclui SQLite4VB: " & Err.Description
    IncluiSQLite4vb = False
    Resume Saida
End Function

Public Function PegLastidbsqLite4vb(ByVal cCON As String, ByVal cTABELA As String) As Long
    Dim db As cSQLite
    Dim value As Variant
    On Error GoTo TrataErro
    Set db = AbrirSQLite4vb(cCON, True)
    value = db.Scalar("SELECT COALESCE(MAX(rowid), 0) FROM " & cTABELA)
    If Not IsNull(value) And Not IsEmpty(value) Then PegLastidbsqLite4vb = CLng(value)
Saida:
    FecharSQLite4vb db
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
    Dim source As cSQLite, target As cSQLite
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

    tableName = ExtrairTabela4vb(cSQLDES)
    If Len(tableName) = 0 Or Not IsArray(aCAMDES) Then Exit Function
    nFields = UBound(aCAMDES) - LBound(aCAMDES) + 1
    If IsArray(aCAMORI) Then
        If UBound(aCAMORI) - LBound(aCAMORI) + 1 <> nFields Then Exit Function
    Else
        Exit Function
    End If

    Set source = AbrirSQLite4vb(cARQORI, True)
    Set target = AbrirSQLite4vb(cARQDES, False)
    Set rs = source.Query(cSQLORI)
    If rs.MoveNext() Then
        ReDim values(0 To nFields - 1)
        For i = 0 To nFields - 1
            values(i) = ValorCampo4vb(rs, aCAMORI(LBound(aCAMORI) + i))
        Next i
        fields = CamposArray4vb(aCAMDES, nFields)

        If IsArray(aOUTDES) Then
            ReDim extraValues(0 To UBound(aOUTDES) - LBound(aOUTDES))
            For i = LBound(extraValues) To UBound(extraValues)
                extraValues(i) = ValorArrayOuEscalar4vb(aOUTORI, i)
            Next i
            nExtra = UBound(extraValues) - LBound(extraValues) + 1
            fields = fields & ", " & CamposArray4vb(aOUTDES, UBound(aOUTDES) - LBound(aOUTDES) + 1)
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
    FecharSQLite4vb source, rs
    FecharSQLite4vb target, targetCheck
    Exit Function
TrataErro:
    SQLMoveRegSQLite4vb = False
    Resume Saida
End Function

Public Function VBSQLiteSetValues4vb(ByVal db As cSQLite) As Boolean
    On Error GoTo TrataErro
    If db Is Nothing Then Exit Function
    db.Execute "PRAGMA cache_size = 2000"
    db.Execute "PRAGMA synchronous = NORMAL"
    db.Execute "PRAGMA temp_store = MEMORY"
    VBSQLiteSetValues4vb = True
    Exit Function
TrataErro:
    VBSQLiteSetValues4vb = False
End Function

Private Function AbrirSQLite4vb(ByVal connection As String, ByVal readOnly As Boolean) As cSQLite
    Dim db As cSQLite
    Set db = New cSQLite
    db.OpenDB LimpaTag(connection), readOnly, Not readOnly
    If Not readOnly Then
        If Not VBSQLiteSetValues4vb(db) Then
            Err.Raise vbObjectError + 30400, "AbrirSQLite4vb", _
                      "Falha ao configurar a conexao SQLite."
        End If
    End If
    Set AbrirSQLite4vb = db
End Function

Private Sub FecharSQLite4vb(ByRef db As cSQLite, Optional ByVal rs As cSQLiteResults)
    On Error Resume Next
    Set rs = Nothing
    If Not db Is Nothing Then db.CloseDB
    Set db = Nothing
End Sub

Private Function ValorCampo4vb(ByVal rs As cSQLiteResults, ByVal field As Variant) As Variant
    ValorCampo4vb = rs(field)
End Function

Private Function ValorArrayOuEscalar4vb(ByVal value As Variant, ByVal index As Long) As Variant
    If IsArray(value) Then
        ValorArrayOuEscalar4vb = value(index)
    Else
        ValorArrayOuEscalar4vb = value
    End If
End Function

Private Function ValorParaGravar4vb(ByVal value As Variant, ByVal cFormat As Variant) As Variant
    If IsNull(value) Or IsEmpty(value) Then
        ValorParaGravar4vb = Null
    ElseIf Left$(UCase$(CStr(cFormat)), 1) = "D" Then
        If IsDate(value) Then
            ValorParaGravar4vb = Format$(CDate(value), "yyyy-mm-dd hh:nn:ss")
        Else
            ValorParaGravar4vb = Null
        End If
    Else
        ValorParaGravar4vb = value
    End If
End Function

Private Function CamposArray4vb(ByVal fields As Variant, ByVal count As Long) As String
    Dim i As Long, result As String
    For i = 0 To count - 1
        If i > 0 Then result = result & ", "
        result = result & CStr(fields(i))
    Next i
    CamposArray4vb = result
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

Private Function ExtrairTabela4vb(ByVal sourceSql As String) As String
    Dim text As String, pFrom As Long, pWhere As Long
    text = Replace(Replace(Trim$(sourceSql), vbCr, " "), vbLf, " ")
    pFrom = InStr(1, UCase$(text), " FROM ", vbTextCompare)
    If pFrom = 0 Then Exit Function
    pFrom = pFrom + 6
    pWhere = InStr(pFrom, UCase$(text), " WHERE ", vbTextCompare)
    If pWhere = 0 Then
        ExtrairTabela4vb = Trim$(Mid$(text, pFrom))
    Else
        ExtrairTabela4vb = Trim$(Mid$(text, pFrom, pWhere - pFrom))
    End If
End Function

Private Function ExtrairUpdate4vb(ByVal sourceSql As String, ByRef tableName As String, _
                                  ByRef whereSql As String) As Boolean
    Dim text As String, pFrom As Long, pWhere As Long
    text = Replace(Replace(Trim$(sourceSql), vbCr, " "), vbLf, " ")
    pFrom = InStr(1, UCase$(text), " FROM ", vbTextCompare)
    pWhere = InStr(1, UCase$(text), " WHERE ", vbTextCompare)
    If pFrom = 0 Or pWhere = 0 Or pWhere <= pFrom Then Exit Function
    tableName = Trim$(Mid$(text, pFrom + 6, pWhere - pFrom - 6))
    whereSql = Mid$(text, pWhere)
    ExtrairUpdate4vb = (Len(tableName) > 0)
End Function

Private Function AliasCampoSQLite4vb(ByVal sourceSql As String) As String
    Dim text As String, p As Long
    text = UCase$(sourceSql)
    p = InStr(1, text, " FROM ", vbTextCompare)
    If p = 0 Then
        AliasCampoSQLite4vb = sourceSql
    Else
        AliasCampoSQLite4vb = Left$(sourceSql, p - 1) & " AS CAMPO" & Mid$(sourceSql, p)
    End If
End Function

Private Function DialetoSQLite4vb(ByVal sourceSql As String) As String
    DialetoSQLite4vb = Replace(sourceSql, "CURRENTDATETIME", "current_timestamp", 1, -1, vbTextCompare)
End Function

Private Sub PreencherIDs4vb(ByVal db As cSQLite, ByVal tableName As String, _
                            ByVal ids As Variant, ByVal insertId As Currency)
    Dim rs As cSQLiteResults
    Dim result() As Variant
    Dim i As Long

    If Not IsArray(ids) Then
        eRETU01 = insertId
        Exit Sub
    End If

    Set rs = db.Query("SELECT " & CamposArray4vb(ids, UBound(ids) - LBound(ids) + 1) & _
                      " FROM " & tableName & " WHERE rowid = ?", insertId)
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
