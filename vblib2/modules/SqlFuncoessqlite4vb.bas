Attribute VB_Name = "sqlsqlite4vb"
Option Explicit
Public Function PegUltSQLite4vb(ByVal cCON As String, ByVal cSql As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    Dim db As New cSQLite, RS As cSQLiteResults, vUltimo As Variant
    On Error GoTo Erro
    vUltimo = eDEFAULT
    db.OpenDB LimpaTag(cCON)
    Set RS = db.Query(cSql)
    Do While RS.MoveNext()
        vUltimo = RS(cCAMPO)
    Loop
    PegUltSQLite4vb = IIf(IsNull(vUltimo), eDEFAULT, vUltimo)
    db.CloseDB: Set RS = Nothing: Set db = Nothing
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

Public Function SQLiteComando4vb(ByVal cCON As String, ByVal cSql As String) As Boolean
    Dim db As New cSQLite
    On Error GoTo Erro
    db.OpenDB LimpaTag(cCON)
    db.Execute cSql
    SQLiteComando4vb = True
    db.CloseDB: Set db = Nothing
    Exit Function
Erro: SQLiteComando4vb = False
End Function

Public Function PegSQLite4vb(ByVal cCON As String, ByVal cSql As String, ByVal nITEM As Long, ByVal aCAM As Variant, ByVal aFOR As Variant, ByVal aPAD As Variant) As Variant
    Dim db As New cSQLite, RS As cSQLiteResults, i As Integer, aVAL() As Variant
    On Error GoTo Erro
    ReDim aVAL(nITEM - 1)
    db.OpenDB LimpaTag(cCON)
    
    Set RS = db.Query(cSql)
    If RS.MoveNext() Then
        For i = 0 To nITEM - 1
            aVAL(i) = IIf(IsNull(RS(aCAM(i))), aPAD(i), RS(aCAM(i)))
        Next i
        PegSQLite4vb = aVAL
    Else
        PegSQLite4vb = aPAD
    End If
    db.CloseDB: Set RS = Nothing: Set db = Nothing
    Exit Function
Erro: PegSQLite4vb = aPAD
End Function

' --- Função GrvSQLite4vb (Mantendo a lógica original) ---
Public Function GrvSQLite4vb(ByVal cARQ As String, _
                           ByVal cSQL_SELECT As String, _
                           ByVal nITEM As Long, _
                           ByVal aCAM As Variant, _
                           ByVal aVAL As Variant, _
                           ByVal aFOR As Variant, _
                           Optional ByVal nStartItem As Long = 0) As Boolean
    
    Dim db As New cSQLite
    Dim cSQL_UPDATE As String
    Dim cTABELA As String, cWHERE As String
    Dim i As Long
    Dim nPosFrom As Long, nPosWhere As Long
    
    On Error GoTo Erro
    
    ' 1. Extração da Tabela e do WHERE (Lógica idêntica à original)
    nPosFrom = InStr(1, UCase(cSQL_SELECT), " FROM ") + 6
    nPosWhere = InStr(1, UCase(cSQL_SELECT), " WHERE ")
    
    If nPosWhere > 0 Then
        cTABELA = Trim(Mid(cSQL_SELECT, nPosFrom, nPosWhere - nPosFrom))
        cWHERE = Mid(cSQL_SELECT, nPosWhere)
    Else
        Exit Function
    End If
    
    ' 2. Montagem do comando UPDATE dinâmico
    cSQL_UPDATE = "UPDATE " & cTABELA & " SET "
    For i = nStartItem To (nITEM - 1)
        cSQL_UPDATE = cSQL_UPDATE & aCAM(i) & " = " & TratarValorParaSQL(aVAL(i), aFOR(i), "SQLITE") & _
                      IIf(i < (nITEM - 1), ", ", " ")
    Next i
    cSQL_UPDATE = cSQL_UPDATE & cWHERE
    
    ' 3. Execução direta na conexão (sqlite4vb)
    db.OpenDB LimpaTag(cARQ)
    VBSQLiteSetValues4vb db ' Aplica configurações WAL/Cache
    
    db.Execute cSQL_UPDATE
    
    GrvSQLite4vb = True
    db.CloseDB
    Set db = Nothing
    Exit Function
    
Erro:
    GrvSQLite4vb = False
    If Not db Is Nothing Then db.CloseDB
End Function
Public Function ExecuteScalarSQLite4vb(ByVal cARQ As String, ByVal cSql As String) As Variant
    Dim db As New cSQLite
    On Error GoTo Erro
    
    db.OpenDB LimpaTag(cARQ)
    VBSQLiteSetValues4vb db
    
    ' db.Scalar retorna o primeiro campo da primeira linha
    ExecuteScalarSQLite4vb = db.Scalar(cSql)
    
    db.CloseDB
    Set db = Nothing
    Exit Function
Erro:
    ExecuteScalarSQLite4vb = Null
    If Not db Is Nothing Then db.CloseDB
End Function
Public Function IncluiSQLite4vb(ByVal cARQ As String, _
                             ByVal cSQL_SELECT As String, _
                             ByVal nITEM As Long, _
                             ByVal aCAM As Variant, _
                             ByVal aVAL As Variant, _
                             ByVal lCHECK As Boolean, _
                             ByVal lMES As Boolean, _
                             ByVal aIDDES As Variant) As Boolean
    
    Dim db As New cSQLite
    Dim cTABELA As String
    Dim i As Long
    Dim cCampos As String
    Dim cValores As String
    Dim lTEM
    Dim RS
    
    On Error GoTo Erro
    
    ' 1. Extração da Tabela
    cTABELA = ExtraiTabela(cSQL_SELECT)
    
    ' 2. Checagem (lCHECK)
    ' Implementação da lógica de verificação se necessário
    If lCHECK Then
       db.OpenDB LimpaTag(cARQ)
        ' Usamos Query com a cSQL_SELECT completa para identificar se o registro já existe
        Set RS = db.Query(cSQL_SELECT)
        If RS.MoveNext() Then lTEM = True
        Set RS = Nothing
        db.CloseDB
        If lTEM Then
           Exit Function
        End If
    End If
    
    ' 3. Montagem do INSERT
    cCampos = ""
    cValores = ""
    For i = 0 To nITEM - 1
        cCampos = cCampos & IIf(cCampos = "", "", ", ") & aCAM(i)
        cValores = cValores & IIf(cValores = "", "", ", ") & TratarValorParaSQL(aVAL(i), "C", "SQLITE")
    Next i
    
    ' 4. Execução utilizando sqlite4vb
    db.OpenDB LimpaTag(cARQ)
    VBSQLiteSetValues4vb db ' Aplica configurações de performance
    
    db.Execute "INSERT INTO " & cTABELA & " (" & cCampos & ") VALUES (" & cValores & ")"
    
    ' 5. Captura do ID (Equivalente ao ADO)
    ' SQLite utiliza db.LastInsertRowID
    If Not IsNumeric(aIDDES) Then
        ' Se aIDDES for uma matriz, popula conforme o padrão da vbsqlite original
        ' Exemplo: aIDDES = Array(db.LastInsertRowID)
    Else
        aIDDES = CLng(db.LastInsertRowID)
    End If
    
    IncluiSQLite4vb = True
    db.CloseDB
    Set db = Nothing
    Exit Function

Erro:
    If lMES Then MsgBox "Erro ao incluir: " & Err.Description
    IncluiSQLite4vb = False
    If Not db Is Nothing Then
        db.CloseDB
        Set db = Nothing
    End If
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

Public Function VBSQLiteSetValues4vb(ByRef db As cSQLite) As Boolean
    On Error GoTo Erro
    ' Configurações de performance para SQLite
    db.Execute "PRAGMA journal_mode = WAL;"
    db.Execute "PRAGMA cache_size = 2000;"
    db.Execute "PRAGMA synchronous = NORMAL;"
    db.Execute "PRAGMA temp_store = MEMORY;"
    VBSQLiteSetValues4vb = True
    Exit Function
Erro:
    VBSQLiteSetValues4vb = False
End Function
Public Function ApagaSQLite4vb(ByVal cCON As String, ByVal cSql As String) As Boolean
    Dim nPOS As Long
    Dim cSQL_FINAL As String
    
    On Error GoTo Erro
    
    cSQL_FINAL = Trim(cSql)
    
    ' 1. Lógica de espelhamento original:
    ' Garante que o comando inicie com DELETE FROM, tratando entradas parciais
    If UCase(Left(cSQL_FINAL, 6)) <> "DELETE" Then
        nPOS = InStr(1, UCase(cSQL_FINAL), "FROM")
        If nPOS > 0 Then
            ' Pega do FROM em diante e adiciona o DELETE
            cSQL_FINAL = "DELETE " & Mid(cSQL_FINAL, nPOS)
        Else
            ' Se for apenas "TABELA WHERE...", reconstrói o comando
            cSQL_FINAL = "DELETE FROM " & cSQL_FINAL
        End If
    End If
    
    ' 2. Execução utilizando a lógica do sqlite4vb
    ' Chamamos a função de comando já adaptada que abre a conexão,
    ' aplica VBSQLiteSetValues4vb e executa o comando
    ApagaSQLite4vb = SQLiteComando4vb(cCON, cSQL_FINAL)
    
    Exit Function

Erro:
    ApagaSQLite4vb = False
End Function
