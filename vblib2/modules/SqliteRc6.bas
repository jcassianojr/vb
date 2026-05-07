Attribute VB_Name = "SqliteRc6"
Option Explicit

'---------------------------------------------------------------------------------------
' Biblioteca: sqliterc6 (Baseada em vbRichClient6 - New_c)
' Objetivo: Espelho funcional completo da SqlFuncoesAdo para SQLite Nativo
'---------------------------------------------------------------------------------------
' Exemplo para RC6
Public Function PegMinSQLiteRC6(cCON, cTW, cCP, eDF)
    PegMinSQLiteRC6 = PegOperSQLiteRC6(cCON, cTW, cCP, eDF, "MIN")
End Function

Public Function PegMaxSQLiteRC6(cCON, cTW, cCP, eDF)
    PegMaxSQLiteRC6 = PegOperSQLiteRC6(cCON, cTW, cCP, eDF, "MAX")
End Function

Public Function PegSumSQLiteRC6(cCON, cTW, cCP, eDF)
    PegSumSQLiteRC6 = PegOperSQLiteRC6(cCON, cTW, cCP, eDF, "SUM")
End Function
'---------------------------------------------------------------------------------------
' EQUIVALENTE A: PegSQLAdo
'---------------------------------------------------------------------------------------
Public Function PegSQLiteRC6(ByVal cCON As String, ByVal cSQL As String, _
                            ByVal nITEM As Long, ByVal aCAM As Variant, _
                            ByVal aFOR As Variant, ByVal aPAD As Variant) As Variant
    Dim loConn As RC6.cConnection
    Dim loRS As RC6.cRecordset
    Dim i As Long
    Dim aVAL() As Variant
    
    On Error GoTo Erro
    ReDim aVAL(nITEM)
    
    ' Abre a conexão e o recordset via RC6
    Set loConn = New_c.Connection(LimpaTagRC6(cCON), DBCreateInMemory)
    Set loRS = loConn.OpenRecordset(SQLDialeto(cSQL, "SQLITE"))
    
    If loRS.RecordCount > 0 Then
        For i = 0 To nITEM
            Dim vValor As Variant
            vValor = loRS.Fields(i).Value
            'vValor = loRS.Value(i)
            
            ' 1. Tratamento de Padrões (aPAD) se Nulo ou Vazio
            If IsNull(vValor) Or vValor = "" Then vValor = aPAD(i)
            
            ' 2. Tratamento de Formatos (aFOR) para tipagem correta no VB6
            Select Case UCase(aFOR(i))
                Case "C", "S": aVAL(i) = CStr(vValor)
                Case "N", "I", "NI", "R": aVAL(i) = Val(Replace(CStr(vValor), ",", "."))
                Case "D": If IsDate(vValor) Then aVAL(i) = CDate(vValor) Else aVAL(i) = aPAD(i)
                Case Else: aVAL(i) = vValor
            End Select
        Next
        PegSQLiteRC6 = aVAL
    Else
        ' Retorna o array de padrões caso não encontre registros
        PegSQLiteRC6 = aPAD
    End If
    
    Set loRS = Nothing
    Set loConn = Nothing
    Exit Function
Erro:
    PegSQLiteRC6 = aPAD
    Set loConn = Nothing
End Function

'---------------------------------------------------------------------------------------
' EQUIVALENTE A: SomaSQLAdo
'---------------------------------------------------------------------------------------
Public Function SomaSQLiteRC6(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                              ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    Dim vRet As Variant
    
    ' Replica a lógica de agregação do RC6 chamando a PegOperSQLiteRC6
    vRet = PegOperSQLiteRC6(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "SUM")
    
    If IsNull(vRet) Or vRet = "" Then
        SomaSQLiteRC6 = eDEFAULT
    Else
        SomaSQLiteRC6 = vRet
    End If
End Function

'---------------------------------------------------------------------------------------
' EQUIVALENTE A: PegCampoSQLADO / PegCampoSQL
'---------------------------------------------------------------------------------------
Public Function PegCampoSQLiteRC6(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                                  ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    
    ' Chama a função mestre de operação passando a string vazia para o operador
    ' para que ela execute um SELECT [cCAMPO] FROM [cTABLEWHERE]
    PegCampoSQLiteRC6 = PegOperSQLiteRC6(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "")
    
End Function

Public Function PegCountSQLiteRC6(ByVal cCON As String, ByVal cSQL As String) As Long
    PegCountSQLiteRC6 = CLng(PegOperSQLiteRC6(cCON, cSQL))
End Function





'---------------------------------------------------------------------------------------
' MOTOR DE OPERAÇÃO SCALAR (Equivalente a PegOperSQLADO)
'---------------------------------------------------------------------------------------
Public Function PegOperSQLiteRC6(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                                ByVal cCAMPO As String, ByVal eDEFAULT As Variant, _
                                ByVal coper As String) As Variant
    Dim loConn As RC6.cConnection
    Dim loRS As RC6.cRecordset
    Dim cSQL As String
    
    On Error GoTo Erro
    
    ' Monta a query dinamicamente como no ADO
    ' Ex: SELECT MIN(Preco) FROM Produtos WHERE Categoria = 1
    cSQL = "SELECT " & coper & "(" & cCAMPO & ") FROM " & cTABLEWHERE
    cSQL = SQLDialeto(cSQL, "SQLITE")
    
    Set loConn = New_c.Connection(LimpaTagRC6(cCON))
    Set loRS = loConn.OpenRecordset(cSQL)
    
    If Not loRS.EOF Then
        If IsNull(loRS.Fields(0).Value) Then
            PegOperSQLiteRC6 = eDEFAULT
        Else
            PegOperSQLiteRC6 = loRS.Fields(0).Value
        End If
    Else
        PegOperSQLiteRC6 = eDEFAULT
    End If
    
    Set loRS = Nothing: Set loConn = Nothing
    Exit Function

Erro:
    PegOperSQLiteRC6 = eDEFAULT
    Set loConn = Nothing
End Function

'---------------------------------------------------------------------------------------
' EQUIVALENTE A: GrvSQLado (Escrita via SQL UPDATE dinâmico)
'---------------------------------------------------------------------------------------
Public Function GrvSQLiteRC6(ByVal cCON As String, ByVal cSQL_SELECT As String, _
                            ByVal aCAM As Variant, ByVal aVAL As Variant) As Boolean
    Dim loConn As RC6.cConnection
    Dim cSQL_UP As String, cTAB As String, cWHE As String
    Dim i As Long, nPosF As Long, nPosW As Long
    
    On Error GoTo Erro
    
    ' Extrai Tabela e Cláusula WHERE do SELECT original
    nPosF = InStr(1, UCase(cSQL_SELECT), " FROM ") + 6
    nPosW = InStr(1, UCase(cSQL_SELECT), " WHERE ")
    
    If nPosW > 0 Then
        cTAB = Trim(Mid(cSQL_SELECT, nPosF, nPosW - nPosF))
        cWHE = Mid(cSQL_SELECT, nPosW)
    Else
        Exit Function
    End If
    
    ' Montagem do comando UPDATE dinâmico
    cSQL_UP = "UPDATE " & cTAB & " SET "
    For i = LBound(aCAM) To UBound(aCAM)
        cSQL_UP = cSQL_UP & aCAM(i) & " = " & FormataParaSQLRC6(aVAL(i)) & _
                 IIf(i < UBound(aCAM), ", ", " ")
    Next
    cSQL_UP = cSQL_UP & cWHE
    
    Set loConn = New_c.Connection(LimpaTagRC6(cCON))
    loConn.Execute SQLDialeto(cSQL_UP, "SQLITE")
    
    GrvSQLiteRC6 = True
    Set loConn = Nothing
    Exit Function
Erro:
    GrvSQLiteRC6 = False
    Set loConn = Nothing
End Function

'---------------------------------------------------------------------------------------
' EQUIVALENTE A: IncluiSQLAdo
'---------------------------------------------------------------------------------------
Public Function IncluiSQLiteRC6(ByVal cCON As String, ByVal cTABELA As String, _
                               ByVal aCAMPOS As Variant, ByVal aVALORES As Variant, _
                               Optional ByRef nID As Long) As Boolean
    Dim loConn As RC6.cConnection
    Dim loRS As RC6.cRecordset
    Dim cSQL As String, i As Long
    
    On Error GoTo Erro
    cSQL = "INSERT INTO " & cTABELA & " ("
    For i = LBound(aCAMPOS) To UBound(aCAMPOS)
        cSQL = cSQL & aCAMPOS(i) & IIf(i < UBound(aCAMPOS), ",", "")
    Next
    cSQL = cSQL & ") VALUES ("
    For i = LBound(aVALORES) To UBound(aVALORES)
        cSQL = cSQL & FormataParaSQLRC6(aVALORES(i)) & IIf(i < UBound(aVALORES), ",", "")
    Next
    cSQL = cSQL & ")"
    
    Set loConn = New_c.Connection(LimpaTagRC6(cCON))
    loConn.Execute cSQL
    
    Set loRS = loConn.OpenRecordset("SELECT last_insert_rowid()")
    
    If Not loRS.EOF Then
        ' Usamos Val e concatenamos com "" para evitar qualquer erro de tipo (Null ou Decimal)
        nID = Val(loRS.Fields(0).Value & "")
    Else
        nID = 0
    End If
    Set loRS = Nothing
    IncluiSQLiteRC6 = True
    Set loConn = Nothing
    Exit Function
Erro:
    IncluiSQLiteRC6 = False
    Set loConn = Nothing
End Function
'---------------------------------------------------------------------------------------
' FUNÇÕES AUXILIARES
'---------------------------------------------------------------------------------------

Private Function FormataParaSQLRC6(ByVal vValor As Variant) As String
    If IsNull(vValor) Or vValor = "" Then
        FormataParaSQLRC6 = "NULL"
    ElseIf IsDate(vValor) Then
        FormataParaSQLRC6 = "'" & Format(vValor, "YYYY-MM-DD HH:MM:SS") & "'"
    ElseIf IsNumeric(vValor) Then
        FormataParaSQLRC6 = Replace(CStr(vValor), ",", ".")
    Else
        FormataParaSQLRC6 = "'" & Replace(vValor, "'", "''") & "'"
    End If
End Function

Private Function LimpaTagRC6(ByVal cCON As String) As String
    LimpaTagRC6 = Replace(cCON, "[SQLITERC6]", "")
End Function

' EQUIVALENTE A: APAGASQLADO [cite: 7]
Public Function ApagaSQLiteRC6(ByVal cCON As String, ByVal cSQL As String) As Boolean
    Dim nPOS As Integer
    ' Converte SELECT ou strings parciais em DELETE funcional como na ADO [cite: 8]
    cSQL = UCase(cSQL)
    nPOS = InStr(cSQL, "FROM")
    If nPOS > 0 Then
        cSQL = "DELETE FROM " & Mid(cSQL, nPOS + 5)
        ApagaSQLiteRC6 = SQLiteComandoRC6(cCON, cSQL)
    End If
End Function

'---------------------------------------------------------------------------------------
' EQUIVALENTE A: PegUltSQLAdo
'---------------------------------------------------------------------------------------
Public Function PegUltSQLiteRC6(ByVal cCON As String, ByVal cTABELA As String) As Long
    Dim vRet As Variant
    
    ' Chama a PegOperSQLiteRC6 sem operador para executar:
    ' SELECT last_insert_rowid() FROM [cTABELA] LIMIT 1
    ' Nota: No SQLite o last_insert_rowid é global por conexão.
    
    vRet = PegOperSQLiteRC6(cCON, cTABELA, "last_insert_rowid()", 0, "")
    
    If IsNumeric(vRet) Then
        PegUltSQLiteRC6 = CLng(vRet)
    Else
        PegUltSQLiteRC6 = 0
    End If
End Function








