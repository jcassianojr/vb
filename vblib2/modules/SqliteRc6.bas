Attribute VB_Name = "SqliteRc6"
Option Explicit
Public Function pegultsliterc6(ByVal cCON As String, ByVal cSQL As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    Dim loRs As Object ' cRecordset
    On Error GoTo Erro
    
    ' O valor 1 é o equivalente interno a SQLiteConn no RC6
    ' Usamos o valor direto para evitar erro de "variável não encontrada"
    Set loRs = New_c.Connection(cCON, 1).OpenRecordset(cSQL)
    
    If loRs.RecordCount > 0 Then
        loRs.MoveLast
        pegultsliterc6 = IIf(IsNull(loRs(cCAMPO).Value), eDEFAULT, loRs(cCAMPO).Value)
    Else
        pegultsliterc6 = eDEFAULT
    End If
    
    Set loRs = Nothing
    Exit Function
Erro:
    pegultsliterc6 = eDEFAULT
End Function
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

'---------------------------------------------------------------------------------------
' EQUIVALENTE A: PegSUMSQLADO
'---------------------------------------------------------------------------------------
Public Function PegSumSQLiteRC6(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                               ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    
    ' Utiliza a PegOperSQLiteRC6 com o operador SUM.
    ' Garante que o motor do RC6 execute a soma nativa.
    PegSumSQLiteRC6 = PegOperSQLiteRC6(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "SUM")
    
End Function
'---------------------------------------------------------------------------------------
' EQUIVALENTE A: PegSQLAdo
'---------------------------------------------------------------------------------------
Public Function PegSQLiteRC6(ByVal cCON As String, ByVal cSQL As String, _
                            ByVal nITEM As Long, ByVal aCAM As Variant, _
                            ByVal aFOR As Variant, ByVal aPAD As Variant) As Variant
    Dim loConn As RC6.cConnection
    Dim loRs As RC6.cRecordset
    Dim i As Long
    Dim aVAL() As Variant
    
    On Error GoTo Erro
    ReDim aVAL(nITEM)
    
    ' Abre a conexão e o recordset via RC6
    Set loConn = New_c.Connection(LimpaTagRC6(cCON), DBCreateInMemory)
    Set loRs = loConn.OpenRecordset(SQLDialeto(cSQL, "SQLITE"))
    
    If loRs.RecordCount > 0 Then
        For i = 0 To nITEM
            Dim vValor As Variant
            vValor = loRs.Fields(i).Value
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
    
    Set loRs = Nothing
    Set loConn = Nothing
    Exit Function
Erro:
    PegSQLiteRC6 = aPAD
    Set loConn = Nothing
End Function

'---------------------------------------------------------------------------------------
' EQUIVALENTE FIEL A: SomaSQLAdo
'---------------------------------------------------------------------------------------
Public Function SomaSQLiteRC6(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                              ByVal cCAMPO As String, ByVal eDEFAULT As Variant, _
                              Optional ByVal nDEC As Integer = 2) As Variant
    Dim loConn As RC6.cConnection
    Dim loRs As RC6.cRecordset
    Dim nSoma As Double
    Dim aOPER As Variant
    Dim aVALORES_LINHA() As String
    Dim x As Integer
    
    On Error GoTo Erro
    
    aOPER = SepSqlOpe(cCAMPO)
    
    Set loConn = New_c.Connection(LimpaTagRC6(cCON))
    Set loRs = loConn.OpenRecordset("SELECT * FROM " & cTABLEWHERE)
    
    nSoma = 0
    Do While Not loRs.EOF
        ReDim aVALORES_LINHA(UBound(aOPER))
        For x = 0 To UBound(aOPER)
            If InStr("+-*/()", aOPER(x)) = 0 And aOPER(x) <> "" Then
                aVALORES_LINHA(x) = loRs.Fields(aOPER(x)).Value & ""
            Else
                aVALORES_LINHA(x) = aOPER(x)
            End If
        Next x
        
        ' REPASSA nDEC PARA A LOGICA DE CALCULO DA MYFUNCTIONS
        nSoma = nSoma + Val(MathOper(aVALORES_LINHA, nDEC))
        
        loRs.MoveNext
    Loop
    
    SomaSQLiteRC6 = IIf(nSoma = 0, eDEFAULT, nSoma)
    
    Set loRs = Nothing: Set loConn = Nothing
    Exit Function
Erro:
    SomaSQLiteRC6 = eDEFAULT
    Set loConn = Nothing
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

Public Function PegCountSQLiteRC6(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                                 ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    
    ' Repassa para a Oper com o comando "COUNT"
    ' cCAMPO normalmente será "*" nesta chamada
    PegCountSQLiteRC6 = PegOperSQLiteRC6(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "COUNT")
    
End Function





'---------------------------------------------------------------------------------------
' MOTOR DE OPERAÇÃO SCALAR (Equivalente a PegOperSQLADO)
'---------------------------------------------------------------------------------------
Public Function PegOperSQLiteRC6(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                                ByVal cCAMPO As String, ByVal eDEFAULT As Variant, _
                                ByVal coper As String) As Variant
    Dim loConn As RC6.cConnection
    Dim loRs As RC6.cRecordset
    Dim cSQL As String
    
    On Error GoTo Erro
    
    ' Monta a query dinamicamente como no ADO
    ' Ex: SELECT MIN(Preco) FROM Produtos WHERE Categoria = 1
    cSQL = "SELECT " & coper & "(" & cCAMPO & ") FROM " & cTABLEWHERE
    cSQL = SQLDialeto(cSQL, "SQLITE")
    
    Set loConn = New_c.Connection(LimpaTagRC6(cCON))
    Set loRs = loConn.OpenRecordset(cSQL)
    
    If Not loRs.EOF Then
        If IsNull(loRs.Fields(0).Value) Then
            PegOperSQLiteRC6 = eDEFAULT
        Else
            PegOperSQLiteRC6 = loRs.Fields(0).Value
        End If
    Else
        PegOperSQLiteRC6 = eDEFAULT
    End If
    
    Set loRs = Nothing: Set loConn = Nothing
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
    Dim loRs As RC6.cRecordset
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
    
    Set loRs = loConn.OpenRecordset("SELECT last_insert_rowid()")
    
    If Not loRs.EOF Then
        ' Usamos Val e concatenamos com "" para evitar qualquer erro de tipo (Null ou Decimal)
        nID = Val(loRs.Fields(0).Value & "")
    Else
        nID = 0
    End If
    Set loRs = Nothing
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

'---------------------------------------------------------------------------------------
' EQUIVALENTE A: ADOComando / SQLiteComando
' Objetivo: Executar comandos SQL (Action Queries) via RC6 sem retorno de recordset
'---------------------------------------------------------------------------------------
Public Function SQLiteComandoRC6(ByVal cCON As String, ByVal cSQL As String) As Boolean
    Dim loConn As RC6.cConnection
    
    On Error GoTo Erro
    
    ' 1. Limpa a Tag da conexão (Ex: tira o "SQLITE:")
    cCON = LimpaTagRC6(cCON)
    
    ' 2. Traduz o dialeto (converte sintaxe Access/VFP para SQLite se necessário)
    cSQL = SQLDialeto(cSQL, "SQLITE")
    
    ' 3. Abre a conexão e executa o comando
    Set loConn = New_c.Connection(cCON)
    loConn.Execute cSQL
    
    SQLiteComandoRC6 = True
    
    ' Limpeza
    Set loConn = Nothing
    Exit Function

Erro:
    ' Em caso de erro (ex: base de dados bloqueada ou erro de sintaxe SQL)
    SQLiteComandoRC6 = False
    Set loConn = Nothing
End Function


'---------------------------------------------------------------------------------------
' EQUIVALENTE A: PegUltSQLAdo
'---------------------------------------------------------------------------------------
Public Function PegLastIDSQLiteRC6(ByVal cCON As String, ByVal cTABELA As String) As Long
    Dim vRet As Variant
    
    ' Chama a PegOperSQLiteRC6 sem operador para executar:
    ' SELECT last_insert_rowid() FROM [cTABELA] LIMIT 1
    ' Nota: No SQLite o last_insert_rowid é global por conexão.
    
    vRet = PegOperSQLiteRC6(cCON, cTABELA, "last_insert_rowid()", 0, "")
    
    If IsNumeric(vRet) Then
        PegLastIDSQLiteRC6 = CLng(vRet)
    Else
        PegLastIDSQLiteRC6 = 0
    End If
End Function


'---------------------------------------------------------------------------------------
' EQUIVALENTE FIEL A: ApagaSQLpAdo / APAGASQLADO
'---------------------------------------------------------------------------------------
Public Function ApagaSQLiteRC6(ByVal cCON As String, ByVal cSQL As String) As Boolean
    Dim nPOS As Long
    Dim cSQL_FINAL As String
    
    On Error GoTo Erro
    
    cSQL_FINAL = Trim(cSQL)
    
    ' Checagem From/Where idêntica à lógica ADO
    If UCase(Left(cSQL_FINAL, 6)) <> "DELETE" Then
        nPOS = InStr(1, UCase(cSQL_FINAL), "FROM")
        If nPOS > 0 Then
            cSQL_FINAL = "DELETE " & Mid(cSQL_FINAL, nPOS)
        Else
            cSQL_FINAL = "DELETE FROM " & cSQL_FINAL
        End If
    End If
    
    ' Utiliza a SQLiteComandoRC6 que criamos anteriormente
    ApagaSQLiteRC6 = SQLiteComandoRC6(cCON, cSQL_FINAL)
    
    Exit Function
Erro:
    ApagaSQLiteRC6 = False
End Function


Public Function PegSQLiteDeliRC6(ByVal cCON As String, ByVal cSQL As String, _
                                ByVal aCAM As Variant, Optional ByVal cDELI As String = ",", _
                                Optional ByVal aPAD As Variant = "", Optional ByVal aFOR As Variant = "") As Variant
    Dim loRs As Object ' cRecordset
    Dim x As Long, nCAMPOS As Integer
    Dim aRETU As Variant, aOPE As Variant, eVAL As Variant

    On Error GoTo Erro

    nCAMPOS = UBound(aCAM) + 1
    ReDim aRETU(nCAMPOS - 1)
    For x = 0 To nCAMPOS - 1: aRETU(x) = "": Next x

    Set loRs = New_c.Connection(cCON, 1).OpenRecordset(cSQL)

    If loRs.RecordCount > 0 Then
        loRs.MoveFirst
        Do Until loRs.EOF
            For x = 0 To nCAMPOS - 1
                ' Lógica de Operações
                aOPE = SepSqlOpe(aCAM(x))
                If aOPE(0) = "" Or aOPE(1) = "" Or aOPE(2) = "" Then
                    eVAL = loRs(aCAM(x)).Value
                Else
                    eVAL = MathOper(loRs(aOPE(1)).Value, loRs(aOPE(2)).Value, aOPE(0))
                End If

                ' Tratamento de Nulo
                If IsNull(eVAL) Then
                    If IsArray(aPAD) Then eVAL = aPAD(x) Else eVAL = aPAD
                End If

                ' Formatação com FVar
                If IsArray(aFOR) Then
                    If IsArray(aPAD) Then
                        eVAL = FVar(eVAL, aFOR(x), aPAD(x))
                    Else
                        eVAL = FVar(eVAL, aFOR(x))
                    End If
                End If

                aRETU(x) = aRETU(x) & FixStr(eVAL)
            Next x

            loRs.MoveNext
            If Not loRs.EOF Then
                For x = 0 To nCAMPOS - 1: aRETU(x) = aRETU(x) & cDELI: Next x
            End If
        Loop
    End If

    PegSQLiteDeliRC6 = aRETU
    Exit Function
Erro:
    PegSQLiteDeliRC6 = aRETU
End Function


Public Function SQLMoveRegSQLiteRC6(ByVal cCONORI As String, ByVal cSQLORI As String, _
   Optional ByVal cOPEORI As String = "", Optional ByVal aCAMORI As Variant = 0, _
   Optional ByVal aOUTORI As Variant = 0, Optional ByVal cCONDES As String = "", _
   Optional ByVal cSQLDES As String = "", Optional ByVal cOPEDES As String = "", _
   Optional ByVal aCAMDES As Variant = 0, Optional ByVal aOUTDES As Variant = 0, _
   Optional ByVal aIDDES As Variant = 0) As Boolean

    Dim loRsOri As Object, loRsDes As Object ' cRecordset
    Dim x As Long, nCAMPOS As Long
    Dim aVALORI As Variant, aRETUID As Variant, aOPE As Variant, eVAL As Variant

    On Error GoTo Erro
    SQLMoveRegSQLiteRC6 = False

    ' 1. Abre Origem e Destino (Literal 1 para SQLiteConn)
    Set loRsOri = New_c.Connection(cCONORI, 1).OpenRecordset(cSQLORI)
    Set loRsDes = New_c.Connection(cCONDES, 1).OpenRecordset(cSQLDES)

    If loRsOri.RecordCount > 0 Then
        ' 2. Extração de Valores da Origem (Passo Matriz)
        nCAMPOS = UBound(aCAMORI)
        ReDim aVALORI(nCAMPOS)
        For x = 0 To nCAMPOS
            aOPE = SepSqlOpe(aCAMORI(x))
            If aOPE(0) = "" Or aOPE(1) = "" Or aOPE(2) = "" Then
                aVALORI(x) = loRsOri(aCAMORI(x)).Value
            Else
                aVALORI(x) = MathOper(loRsOri(aOPE(1)).Value, loRsOri(aOPE(2)).Value, aOPE(0))
            End If
        Next x

        ' 3. Gravação no Destino
        If loRsDes.RecordCount = 0 Then
            loRsDes.AddNew
        End If

        ' Campos Mapeados
        If IsArray(aCAMDES) Then
            For x = 0 To UBound(aCAMDES)
                loRsDes(aCAMDES(x)).Value = aVALORI(x)
            Next x
        End If

        ' Campos Adicionais
        If IsArray(aOUTDES) Then
            For x = 0 To UBound(aOUTDES)
                loRsDes(aOUTDES(x)).Value = aOUTORI(x)
            Next x
        End If

        loRsDes.UpdateBatch

        ' 4. Captura de IDs para eRETU01
        If IsArray(aIDDES) Then
            ReDim aRETUID(UBound(aIDDES))
            For x = 0 To UBound(aIDDES)
                aRETUID(x) = loRsDes(aIDDES(x)).Value
            Next x
            eRETU01 = aRETUID
        End If
        SQLMoveRegSQLiteRC6 = True
    End If
    Exit Function
Erro:
    SQLMoveRegSQLiteRC6 = False
End Function


