Attribute VB_Name = "sqlsqlite4vb"
Option Explicit

' Substitua as referÃªncias da biblioteca antiga pela da sqlite4vb

Public Function PegUltSQLite4vb(ByVal cCON As String, ByVal cSQL As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    Dim db As New cSQLite, rs As cSQLiteResults, vUltimo As Variant
    On Error GoTo Erro
    vUltimo = eDEFAULT
    db.OpenDB LimpaTag(cCON)
    Set rs = db.Query(cSQL)
    Do While rs.MoveNext()
        vUltimo = rs(cCAMPO)
    Loop
    PegUltSQLite4vb = IIf(IsNull(vUltimo), eDEFAULT, vUltimo)
    db.CloseDB: Set rs = Nothing: Set db = Nothing
    Exit Function
Erro: PegUltSQLite4vb = eDEFAULT
End Function

Public Function PegMinSQLite4vb(ByVal cCON As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    PegMinSQLite4vb = PegOperSQLite4vb(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "MIN")
End Function

Public Function PegMaxSQLite4vb(ByVal cCON As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    PegMaxSQLite4vb = PegOperSQLite4vb(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "MAX")
End Function

Public Function PegSumSQLite4vb(ByVal cCON As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    PegSumSQLite4vb = PegOperSQLite4vb(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "SUM")
End Function

Public Function PegCountSQLite4vb(ByVal cCON As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    PegCountSQLite4vb = PegOperSQLite4vb(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "COUNT")
End Function

Public Function PegOperSQLite4vb(ByVal cCON As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant, ByVal coper As String) As Variant
    Dim db As New cSQLite, eValue As Variant
    On Error GoTo Erro
    db.OpenDB LimpaTag(cCON)
    eValue = db.Scalar("SELECT " & coper & "(" & cCAMPO & ") FROM " & cTABLEWHERE)
    PegOperSQLite4vb = IIf(IsNull(eValue) Or IsEmpty(eValue), eDEFAULT, eValue)
    db.CloseDB: Set db = Nothing
    Exit Function
Erro: PegOperSQLite4vb = eDEFAULT
End Function

Public Function SQLiteComando4vb(ByVal cCON As String, ByVal cSQL As String) As Boolean
    Dim db As New cSQLite
    On Error GoTo Erro
    db.OpenDB LimpaTag(cCON)
    db.Execute cSQL
    SQLiteComando4vb = True
    db.CloseDB: Set db = Nothing
    Exit Function
Erro: SQLiteComando4vb = False
End Function

Public Function PegSQLite4vb(ByVal cCON As String, ByVal cSQL As String, ByVal nITEM As Long, ByVal aCAM As Variant, ByVal aFOR As Variant, ByVal aPAD As Variant) As Variant
    Dim db As New cSQLite, rs As cSQLiteResults, i As Integer, aVAL() As Variant
    On Error GoTo Erro
    ReDim aVAL(nITEM - 1)
    db.OpenDB LimpaTag(cCON)
    Set rs = db.Query(cSQL)
    If rs.MoveNext() Then
        For i = 0 To nITEM - 1
            aVAL(i) = IIf(IsNull(rs(aCAM(i))), aPAD(i), rs(aCAM(i)))
        Next i
        PegSQLite4vb = aVAL
    Else
        PegSQLite4vb = aPAD
    End If
    db.CloseDB: Set rs = Nothing: Set db = Nothing
    Exit Function
Erro: PegSQLite4vb = aPAD
End Function

Public Function GrvSQLite4vb(ByVal cARQ As String, ByVal cSQL_SELECT As String, ByVal nITEM As Long, ByVal aCAM As Variant, ByVal aVAL As Variant, ByVal aFOR As Variant, Optional ByVal nStartItem As Long = 0) As Boolean
    Dim db As New cSQLite, i As Integer, cSql As String
    On Error GoTo Erro
    cSql = "UPDATE " & ExtrairTabela(cSQL_SELECT) & " SET "
    For i = nStartItem To (nITEM - 1)
        cSql = cSql & aCAM(i) & " = ? " & IIf(i < (nITEM - 1), ", ", " ")
    Next i
    cSql = cSql & ExtrairWhere(cSQL_SELECT)
    db.OpenDB LimpaTag(cARQ)
    Dim stmt As cSQLiteStatement: Set stmt = db.Prepare(cSql)
    For i = nStartItem To (nITEM - 1): stmt.Bind i + 1 - nStartItem, aVAL(i): Next i
    stmt.Execute: GrvSQLite4vb = True
    db.CloseDB: Set db = Nothing
    Exit Function
Erro: GrvSQLite4vb = False
End Function

Public Function IncluiSQLite4vb(ByVal cARQ As String, ByVal cSQL_SELECT As String, ByVal nITEM As Long, ByVal aCAM As Variant, ByVal aVAL As Variant, ByVal lCHECK As Boolean, ByVal lMES As Boolean, ByVal aIDDES As Variant) As Boolean
    Dim db As New cSQLite, i As Long, cC As String, cV As String
    On Error GoTo Erro
    For i = 0 To nITEM - 1
        cC = cC & IIf(i = 0, "", ", ") & aCAM(i)
        cV = cV & IIf(i = 0, "", ", ") & "?"
    Next i
    db.OpenDB LimpaTag(cARQ)
    Dim stmt As cSQLiteStatement: Set stmt = db.Prepare("INSERT INTO " & ExtrairTabela(cSQL_SELECT) & " (" & cC & ") VALUES (" & cV & ")")
    For i = 0 To nITEM - 1: stmt.Bind i + 1, aVAL(i): Next i
    stmt.Execute: IncluiSQLite4vb = True
    db.CloseDB: Set db = Nothing
    Exit Function
Erro: If lMES Then MsgBox Err.Description: IncluiSQLite4vb = False
End Function

Public Function PegCampoSQLite4vb(ByVal cCON As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    PegCampoSQLite4vb = PegOperSQLite4vb(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "")
End Function

Public Function PegLastidbsqLite4vb(ByVal cCON As String, ByVal cTABELA As String) As Long
    Dim db As New cSQLite
    db.OpenDB LimpaTag(cCON)
    PegLastidbsqLite4vb = CLng(db.LastInsertRowID)
    db.CloseDB
End Function

' --- FunÃ§Ãµes Auxiliares Internas ---
Private Function LimpaTag(ByVal cCON As String) As String: LimpaTag = Replace(cCON, "[VBSQLITE]", ""): End Function
Private Function ExtrairTabela(ByVal s As String) As String: ExtrairTabela = Trim(Mid(s, InStr(1, UCase(s), "FROM ") + 5, InStr(1, UCase(s) & " ", " ") - 1)): End Function
Private Function ExtrairWhere(ByVal s As String) As String: Dim p As Long: p = InStr(1, UCase(s), " WHERE "): If p > 0 Then ExtrairWhere = Mid(s, p): End Function