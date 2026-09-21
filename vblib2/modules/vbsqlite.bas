Attribute VB_Name = "sqlvbsqlite"
Option Explicit

' Adaptador SQLite usado por sqlfuncoes.bas.
' As assinaturas publicas sao mantidas compativeis com SqlFuncoesAdo.bas.

Public Function PegUltSQLite(ByVal cCON As String, ByVal cSQL As String, _
                             ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    Dim db As SQLiteConnection
    Dim RS As SQLiteDataSet
    Dim v As Variant

    On Error GoTo TrataErro
    PegUltSQLite = eDEFAULT
    Set db = AbrirSQLite(cCON, True)
    Set RS = db.OpenDataSet(sqldialeto(cSQL, "SQLITE"))
    If Not RS.EOF Then
        RS.MoveLast
        v = RS(cCAMPO)
        If Not IsNull(v) Then PegUltSQLite = v
    End If
Saida:
    Set RS = Nothing
    FecharSQLite db
    Exit Function
TrataErro:
    PegUltSQLite = eDEFAULT
    Resume Saida
End Function

Public Function PegMinSQLite(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                             ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    PegMinSQLite = PegOperSQLite(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "MIN")
End Function

Public Function PegMaxSQLite(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                             ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    PegMaxSQLite = PegOperSQLite(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "MAX")
End Function

Public Function PegSumSQLite(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                             ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    PegSumSQLite = PegOperSQLite(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "SUM")
End Function

Public Function PegCountSQLite(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                               ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    PegCountSQLite = PegOperSQLite(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "COUNT")
End Function

Public Function PegCampoSQLite(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                               ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    PegCampoSQLite = PegOperSQLite(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "CAMPO")
End Function

Public Function PegOperSQLite(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                              ByVal cCAMPO As String, ByVal eDEFAULT As Variant, _
                              ByVal coper As String) As Variant
    Dim db As SQLiteConnection
    Dim RS As SQLiteDataSet
    Dim sql As String
    Dim v As Variant
    Dim op As String

    On Error GoTo TrataErro
    PegOperSQLite = eDEFAULT
    op = UCase$(Trim$(coper))
    Select Case op
        Case "MIN", "MAX", "SUM", "COUNT"
            sql = "SELECT " & op & "(" & cCAMPO & ") AS CAMPO FROM " & cTABLEWHERE
        Case "CAMPO"
            If Len(Trim$(cCAMPO)) = 0 Then
                sql = Replace(UCase$(cTABLEWHERE), " FROM ", " AS CAMPO FROM ")
            Else
                sql = "SELECT " & cCAMPO & " AS CAMPO FROM " & cTABLEWHERE
            End If
        Case Else
            Err.Raise 5, "PegOperSQLite", "Operacao SQLite invalida: " & coper
    End Select

    Set db = AbrirSQLite(cCON, True)
    Set RS = db.OpenDataSet(sqldialeto(sql, "SQLITE"))
    If Not RS.EOF Then
        v = RS("CAMPO")
        If Not IsNull(v) Then PegOperSQLite = v
    End If
Saida:
    Set RS = Nothing
    FecharSQLite db
    Exit Function
TrataErro:
    PegOperSQLite = eDEFAULT
    Resume Saida
End Function

Public Function ComandoSqlite(ByVal cCON As String, ByVal cSQL As String) As Boolean
    Dim db As SQLiteConnection
    On Error GoTo TrataErro
    Set db = AbrirSQLite(cCON, False)
    db.Execute sqldialeto(cSQL, "SQLITE")
    ComandoSqlite = True
Saida:
    FecharSQLite db
    Exit Function
TrataErro:
    ComandoSqlite = False
    Resume Saida
End Function

Public Function PegSQLite(ByVal cCON As String, ByVal cSQL As String, _
                          ByVal nITEM As Long, ByVal aCAM As Variant, _
                          ByVal aFOR As Variant, ByVal aPAD As Variant) As Variant
    Dim db As SQLiteConnection
    Dim RS As SQLiteDataSet
    Dim ret() As Variant
    Dim i As Long
    Dim v As Variant

    On Error GoTo TrataErro
    If nITEM <= 0 Then
        PegSQLite = aPAD
        Exit Function
    End If
    ReDim ret(0 To nITEM - 1)
    For i = 0 To nITEM - 1
        ret(i) = aPAD(i)
    Next i

    Set db = AbrirSQLite(cCON, True)
    Set RS = db.OpenDataSet(sqldialeto(cSQL, "SQLITE"))
    If Not RS.EOF Then
        For i = 0 To nITEM - 1
            v = ValorDeCampo(RS, aCAM(i))
            If IsNull(v) Or IsEmpty(v) Then
                ret(i) = aPAD(i)
            Else
                ret(i) = FVar(v, aFOR(i), aPAD(i))
            End If
        Next i
        PegSQLite = ret
    Else
        PegSQLite = aPAD
    End If
Saida:
    Set RS = Nothing
    FecharSQLite db
    Exit Function
TrataErro:
    PegSQLite = aPAD
    Resume Saida
End Function

Public Function GrvSQLite(ByVal cARQ As String, ByVal cSQL_SELECT As String, _
                          ByVal nITEM As Long, ByVal aCAM As Variant, _
                          ByVal aVAL As Variant, ByVal aFOR As Variant, _
                          Optional ByVal nStartItem As Long = 0) As Boolean
    Dim db As SQLiteConnection
    Dim RS As SQLiteDataSet
    Dim cmd As SQLiteCommand
    Dim tableName As String, whereSql As String, sql As String
    Dim i As Long, firstItem As Long
    Dim Value As Variant

    On Error GoTo TrataErro
    If nStartItem < 0 Then nStartItem = 0
    If nStartItem >= nITEM Then Exit Function
    If Not ExtrairUpdatePartes(cSQL_SELECT, tableName, whereSql) Then Exit Function

    sql = "UPDATE " & tableName & " SET "
    For i = nStartItem To nITEM - 1
        If i > nStartItem Then sql = sql & ", "
        sql = sql & CStr(aCAM(i)) & " = ?"
    Next i
    sql = sql & " " & whereSql

    Set db = AbrirSQLite(cARQ, False)
    Set cmd = db.CreateCommand(sql)
    For i = nStartItem To nITEM - 1
        Value = ValorParaGravar(aVAL(i), aFOR(i))
        cmd.SetParameterValue i - nStartItem + 1, Value
    Next i
    cmd.Execute
    GrvSQLite = True
Saida:
    Set cmd = Nothing
    Set RS = Nothing
    FecharSQLite db
    Exit Function
TrataErro:
    GrvSQLite = False
    Resume Saida
End Function

Public Function IncluiSQLite(ByVal cARQ As String, ByVal cSQL_SELECT As String, _
                             ByVal nITEM As Long, Optional ByVal aCAM As Variant = 0, _
                             Optional ByVal aVAL As Variant = 0, _
                             Optional ByVal lCHECK As Boolean = False, _
                             Optional ByVal lMES As Boolean = True, _
                             Optional ByVal aIDDES As Variant = 0) As Boolean
    Dim db As SQLiteConnection
    Dim RS As SQLiteDataSet
    Dim cmd As SQLiteCommand
    Dim tableName As String, sql As String
    Dim fields As String, marks As String
    Dim i As Long, Value As Variant
    Dim alreadyExists As Boolean

    On Error GoTo TrataErro
    tableName = ExtrairNomeTabela(cSQL_SELECT)
    If Len(tableName) = 0 Or nITEM <= 0 Then Exit Function

    Set db = AbrirSQLite(cARQ, False)
    If lCHECK Then
        Set RS = db.OpenDataSet(sqldialeto(cSQL_SELECT, "SQLITE"))
        alreadyExists = Not RS.EOF
        Set RS = Nothing
        If alreadyExists Then
            If lMES Then MsgBox "Item ja cadastrado com esta chave.", vbInformation
            Exit Function
        End If
    End If

    For i = 0 To nITEM - 1
        If i > 0 Then
            fields = fields & ", "
            marks = marks & ", "
        End If
        fields = fields & CStr(aCAM(i))
        marks = marks & "?"
    Next i

    sql = "INSERT INTO " & tableName & " (" & fields & ") VALUES (" & marks & ")"
    Set cmd = db.CreateCommand(sql)
    For i = 0 To nITEM - 1
        Value = ValorParaGravar(aVAL(i), "C")
        cmd.SetParameterValue i + 1, Value
    Next i
    cmd.Execute

    ' O valor deve ser consumido na mesma conexao que executou o INSERT.
    If IsArray(aIDDES) Then
        PreencherIDs db, tableName, aIDDES
    End If
    IncluiSQLite = True
Saida:
    Set cmd = Nothing
    Set RS = Nothing
    FecharSQLite db
    Exit Function
TrataErro:
    If lMES Then MsgBox "Erro ao incluir: " & Err.Description, vbExclamation
    IncluiSQLite = False
    Resume Saida
End Function

Public Function PegSQLiteDeli(ByVal cCON As String, ByVal cSQL As String, _
                              ByVal aCAM As Variant, Optional ByVal cDELI As String = ",", _
                              Optional ByVal aPAD As Variant = "", _
                              Optional ByVal aFOR As Variant = "") As Variant
    Dim db As SQLiteConnection
    Dim RS As SQLiteDataSet
    Dim ret() As String
    Dim op As Variant, Value As Variant
    Dim x As Long, count As Long

    On Error GoTo TrataErro
    count = UBound(aCAM)
    ReDim ret(0 To count)
    Set db = AbrirSQLite(cCON, True)
    Set RS = db.OpenDataSet(sqldialeto(cSQL, "SQLITE"))

    Do While Not RS.EOF
        For x = 0 To count
            op = SepSqlOpe(aCAM(x))
            If IsArray(op) And op(0) <> "" And op(1) <> "" And op(2) <> "" Then
                Value = MathOper(RS(op(1)), RS(op(2)), op(0))
            Else
                Value = ValorDeCampo(RS, aCAM(x))
            End If
            If IsNull(Value) Then Value = ValorDoArrayOuPadrao(aPAD, x)
            If IsArray(aFOR) Then Value = FVar(Value, aFOR(x), ValorDoArrayOuPadrao(aPAD, x))
            If Len(ret(x)) > 0 Then ret(x) = ret(x) & cDELI
            ret(x) = ret(x) & FixStr(Value)
        Next x
        RS.MoveNext
    Loop
    PegSQLiteDeli = ret
Saida:
    Set RS = Nothing
    FecharSQLite db
    Exit Function
TrataErro:
    PegSQLiteDeli = ret
    Resume Saida
End Function

Public Function PegLastidbsqLite(ByVal cCON As String, ByVal cTABELA As String) As Long
    Dim db As SQLiteConnection
    On Error GoTo TrataErro
    Set db = AbrirSQLite(cCON, False)
    PegLastidbsqLite = CLng(db.LastInsertRowID)
Saida:
    FecharSQLite db
    Exit Function
TrataErro:
    PegLastidbsqLite = 0
    Resume Saida
End Function

Public Function SomaSQLite(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                           ByVal cCAMPO As String, Optional ByVal eDEFAULT As Variant, _
                           Optional ByVal nDEC As Integer = 2) As Variant
    Dim db As SQLiteConnection
    Dim RS As SQLiteDataSet
    Dim ops As Variant, lineValues() As String
    Dim Total As Double, found As Boolean
    Dim x As Long, Value As Variant

    On Error GoTo TrataErro
    ops = SepSqlOpe(cCAMPO)
    Set db = AbrirSQLite(cCON, True)
    Set RS = db.OpenDataSet("SELECT * FROM " & cTABLEWHERE)
    Do While Not RS.EOF
        found = True
        ReDim lineValues(0 To UBound(ops))
        For x = 0 To UBound(ops)
            If InStr(1, "+-*/()", CStr(ops(x)), vbBinaryCompare) = 0 And ops(x) <> "" Then
                Value = ValorDeCampo(RS, ops(x))
                If IsNull(Value) Then Value = 0
                lineValues(x) = CStr(Value)
            Else
                lineValues(x) = CStr(ops(x))
            End If
        Next x
        Total = Total + CDbl(Val(MathOper(lineValues, nDEC)))
        RS.MoveNext
    Loop
    If found Then SomaSQLite = Total Else SomaSQLite = eDEFAULT
Saida:
    Set RS = Nothing
    FecharSQLite db
    Exit Function
TrataErro:
    SomaSQLite = eDEFAULT
    Resume Saida
End Function

Public Function ApagaSQLite(ByVal cCON As String, ByVal cSQL As String) As Boolean
    Dim tEXT As String, p As Long
    tEXT = Trim$(cSQL)
    If UCase$(Left$(tEXT, 6)) <> "DELETE" Then
        p = InStr(1, UCase$(tEXT), "FROM", vbTextCompare)
        If p > 0 Then tEXT = "DELETE " & Mid$(tEXT, p)
    End If
    If UCase$(Left$(Trim$(tEXT), 6)) <> "DELETE" Then Exit Function
    ApagaSQLite = ComandoSqlite(cCON, tEXT)
End Function

Public Function SQLMoveRegSQLite(ByVal cCONORI As String, ByVal cSQLORI As String, _
   Optional ByVal cOPEORI As String = "", Optional ByVal aCAMORI As Variant = 0, _
   Optional ByVal aOUTORI As Variant = 0, Optional ByVal cCONDES As String = "", _
   Optional ByVal cSQLDES As String = "", Optional ByVal cOPEDES As String = "", _
   Optional ByVal aCAMDES As Variant = 0, Optional ByVal aOUTDES As Variant = 0, _
   Optional ByVal aIDDES As Variant = 0) As Boolean
    Dim Source As SQLiteConnection, target As SQLiteConnection
    Dim RS As SQLiteDataSet, checkRs As SQLiteDataSet
    Dim values() As Variant, ops As Variant, tableName As String, whereSql As String
    Dim sql As String, fields As String, marks As String, sets As String
    Dim x As Long, Value As Variant
    Dim cmd As SQLiteCommand

    On Error GoTo TrataErro
    Set Source = AbrirSQLite(cCONORI, True)
    Set target = AbrirSQLite(cCONDES, False)
    Set RS = Source.OpenDataSet(sqldialeto(cSQLORI, "SQLITE"))
    If RS.EOF Then GoTo Saida

    If IsArray(aCAMORI) Then
        ReDim values(0 To UBound(aCAMORI))
        For x = 0 To UBound(aCAMORI)
            ops = SepSqlOpe(aCAMORI(x))
            If IsArray(ops) And ops(0) <> "" Then
                values(x) = MathOper(RS(ops(1)), RS(ops(2)), ops(0))
            Else
                values(x) = ValorDeCampo(RS, aCAMORI(x))
            End If
        Next x
    End If

    tableName = ExtraiTabela(cSQLDES)
    whereSql = ExtraiWhere(cSQLDES)
    Set checkRs = target.OpenDataSet("SELECT count(*) AS N FROM " & tableName & " " & whereSql)
    If checkRs(1) > 0 Then
        For x = 0 To UBound(aCAMDES)
            If Len(sets) > 0 Then sets = sets & ", "
            sets = sets & CStr(aCAMDES(x)) & " = ?"
        Next x
        If IsArray(aOUTDES) Then
            For x = 0 To UBound(aOUTDES)
                sets = sets & ", " & CStr(aOUTDES(x)) & " = ?"
            Next x
        End If
        sql = "UPDATE " & tableName & " SET " & sets & " " & whereSql
    Else
        For x = 0 To UBound(aCAMDES)
            If x > 0 Then fields = fields & ", ": marks = marks & ", "
            fields = fields & CStr(aCAMDES(x)): marks = marks & "?"
        Next x
        If IsArray(aOUTDES) Then
            For x = 0 To UBound(aOUTDES)
                fields = fields & ", " & CStr(aOUTDES(x))
                marks = marks & ", ?"
            Next x
        End If
        sql = "INSERT INTO " & tableName & " (" & fields & ") VALUES (" & marks & ")"
    End If
    Set cmd = target.CreateCommand(sql)
    For x = 0 To UBound(aCAMDES)
        Value = values(x)
        cmd.SetParameterValue x + 1, Value
    Next x
    If IsArray(aOUTDES) Then
        For x = 0 To UBound(aOUTDES)
            Value = ValorDoArrayOuPadrao(aOUTORI, x)
            cmd.SetParameterValue UBound(aCAMDES) + x + 2, Value
        Next x
    End If
    cmd.Execute
    SQLMoveRegSQLite = True
Saida:
    Set cmd = Nothing
    Set RS = Nothing
    Set checkRs = Nothing
    FecharSQLite Source
    FecharSQLite target
    Exit Function
TrataErro:
    SQLMoveRegSQLite = False
    Resume Saida
End Function

Public Function VBSQLiteSetValues(ByRef oCONN As SQLiteConnection) As Boolean
    On Error GoTo TrataErro
    If oCONN Is Nothing Then Exit Function
    oCONN.Execute "PRAGMA temp_store = MEMORY"
    oCONN.Execute "PRAGMA cache_size = 2000"
    oCONN.Execute "PRAGMA journal_mode = WAL"
    oCONN.Execute "PRAGMA synchronous = NORMAL"
    VBSQLiteSetValues = True
    Exit Function
TrataErro:
    VBSQLiteSetValues = False
End Function

Public Function LimpaTag(ByVal cCON As String) As String
    LimpaTag = Replace(Replace(Replace(Replace(cCON, "[VBSQLITE]", ""), _
        "[SQLITERC6]", ""), "[TC6SQLITE]", ""), "[SQLITE4VB]", "")
End Function

Public Function TratarValorParaSQL(ByVal vValor As Variant, ByVal cTIPO As String, _
                                   ByVal cDIALETO As String) As String
    If IsNull(vValor) Or IsEmpty(vValor) Or (VarType(vValor) = vbString And vValor = "") Then
        TratarValorParaSQL = "NULL"
    ElseIf UCase$(cDIALETO) = "SQLITE" Then
        Select Case UCase$(Left$(cTIPO, 1))
            Case "D": TratarValorParaSQL = "'" & format$(CDate(vValor), "yyyy-mm-dd hh:nn:ss") & "'"
            Case "N": TratarValorParaSQL = Replace$(Trim$(CStr(vValor)), ",", ".")
            Case Else: TratarValorParaSQL = "'" & Replace$(CStr(vValor), "'", "''") & "'"
        End Select
    Else
        TratarValorParaSQL = "'" & Replace$(CStr(vValor), "'", "''") & "'"
    End If
End Function

Private Function AbrirSQLite(ByVal cCON As String, ByVal lReadOnly As Boolean) As SQLiteConnection
    Dim db As New SQLiteConnection
    If lReadOnly Then
        db.OpenDB LimpaTag(cCON), SQLiteReadOnly
    Else
        db.OpenDB LimpaTag(cCON), SQLiteReadWrite
        If Not VBSQLiteSetValues(db) Then Err.Raise 5, "AbrirSQLite", "Falha ao configurar a conexao SQLite."
    End If
    Set AbrirSQLite = db
End Function

Private Sub FecharSQLite(ByRef db As SQLiteConnection)
    On Error Resume Next
    If Not db Is Nothing Then db.CloseDB
    Set db = Nothing
    On Error GoTo 0
End Sub

Private Function ValorDeCampo(ByVal RS As SQLiteDataSet, ByVal campo As Variant) As Variant
    If IsNumeric(campo) And CLng(campo) = 0 Then
        ValorDeCampo = RS(1)
    Else
        ValorDeCampo = RS(campo)
    End If
End Function

Private Function ValorParaGravar(ByVal vValue As Variant, _
                                 ByVal cFormat As Variant) As Variant
    If IsNull(vValue) Or IsEmpty(vValue) Then
        ValorParaGravar = Null
    ElseIf Len(CStr(cFormat)) > 0 And _
           UCase$(Left$(CStr(cFormat), 1)) = "D" Then

        If IsDate(vValue) Then
            ValorParaGravar = format$(CDate(vValue), _
                                      "yyyy-mm-dd hh:nn:ss")
        Else
            ValorParaGravar = Null
        End If
    Else
        Select Case VarType(vValue)
            Case vbBoolean
                ValorParaGravar = Abs(CLng(vValue))
            Case vbSingle
                ValorParaGravar = CDbl(vValue)
            Case vbDate
                ValorParaGravar = format$(CDate(vValue), _
                                          "yyyy-mm-dd hh:nn:ss")
            Case Else
                ValorParaGravar = vValue
        End Select
    End If
End Function

Private Function ValorDoArrayOuPadrao(ByVal values As Variant, ByVal Index As Long) As Variant
    If IsArray(values) Then ValorDoArrayOuPadrao = values(Index) Else ValorDoArrayOuPadrao = values
End Function

Private Function ExtrairUpdatePartes(ByVal sourceSql As String, ByRef tableName As String, ByRef whereSql As String) As Boolean
    Dim pFrom As Long, pWhere As Long, upperSql As String
    upperSql = UCase$(sourceSql)
    pFrom = InStr(1, upperSql, " FROM ", vbBinaryCompare)
    pWhere = InStr(1, upperSql, " WHERE ", vbBinaryCompare)
    If pFrom = 0 Or pWhere = 0 Or pWhere <= pFrom Then Exit Function
    tableName = Trim$(Mid$(sourceSql, pFrom + 6, pWhere - pFrom - 6))
    whereSql = Mid$(sourceSql, pWhere)
    ExtrairUpdatePartes = (Len(tableName) > 0)
End Function

Private Sub PreencherIDs(ByVal db As SQLiteConnection, ByVal tableName As String, ByVal ids As Variant)
    Dim x As Long
    Dim RS As SQLiteDataSet
    Dim sql As String
    Dim result() As Variant

    sql = "SELECT "
    For x = LBound(ids) To UBound(ids)
        If x > LBound(ids) Then sql = sql & ", "
        sql = sql & CStr(ids(x))
    Next x
    sql = sql & " FROM " & tableName & " WHERE rowid = last_insert_rowid()"
    Set RS = db.OpenDataSet(sql)
    If Not RS.EOF Then
        ReDim result(LBound(ids) To UBound(ids))
        For x = LBound(ids) To UBound(ids)
            result(x) = RS(x - LBound(ids) + 1)
        Next x
        eRETU01 = result
    Else
        eRETU01 = db.LastInsertRowID
    End If
End Sub
