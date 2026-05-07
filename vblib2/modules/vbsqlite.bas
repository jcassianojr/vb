Attribute VB_Name = "vbsqlite"
Option Explicit
' Reference=*\G{7CC1A5F1-A0FF-4546-A0F1-FBFE744A4522}#1.1#0#..\..\..\..\..\..\WINDOWS\system32\VBSQLite12.DLL#VB SQLite Library 1.2
' Notas de Implementação:
' 1. Requer as classes cSQLiteConnection e cSQLiteCursor do projeto VBSQLite.
' 2. Mantém compatibilidade com a função SQLDialeto do adolib.bas.

'---------------------------------------------------------------------------------------
' WRAPPERS DE OPERAÇÃO - ESPELHOS DA SQLFUNCOESADO
'---------------------------------------------------------------------------------------

' EQUIVALENTE A: PegMINSQLADO
Public Function PegMinSQLite(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                            ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    PegMinSQLite = PegOperSQLite(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "MIN")
End Function

' EQUIVALENTE A: PegMAXSQLADO
Public Function PegMaxSQLite(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                            ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    PegMaxSQLite = PegOperSQLite(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "MAX")
End Function

' EQUIVALENTE A: PegSUMSQLADO
Public Function PegSumSQLite(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                            ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    PegSumSQLite = PegOperSQLite(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "SUM")
End Function

' EQUIVALENTE A: PegCountSQLADO
Public Function PegCountSQLite(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                              ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    ' No Count, se o campo for *, ele conta todos os registos
    PegCountSQLite = PegOperSQLite(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "COUNT")
End Function

'---------------------------------------------------------------------------------------
' FUNÇÃO CORE (Ajustada para aceitar os parâmetros do ADO)
'---------------------------------------------------------------------------------------
Public Function PegOperSQLite(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                             ByVal cCAMPO As String, ByVal eDEFAULT As Variant, _
                             ByVal coper As String) As Variant
    Dim loConn As New SQLiteConnection
    Dim loCursor As SQLiteCursor
    Dim cSQL As String
    
    On Error GoTo Erro
    
    ' Montagem da query idêntica ao processo ADO
    ' Ex: SELECT SUM(VALOR) FROM MOVI WHERE CODIGO = 10
    cSQL = "SELECT " & coper & "(" & cCAMPO & ") FROM " & cTABLEWHERE
    cSQL = SQLDialeto(cSQL, "SQLITE")
    
    loConn.OpenDB LimpaTag(cCON)
    Set loCursor = loConn.CreateCursor(cSQL)
    
    If Not loCursor.EOF Then
        If IsNull(loCursor.Value(0)) Then
            PegOperSQLite = eDEFAULT
        Else
            PegOperSQLite = loCursor.Value(0)
        End If
    Else
        PegOperSQLite = eDEFAULT
    End If
    
    loConn.CloseDB
    Set loCursor = Nothing: Set loConn = Nothing
    Exit Function

Erro:
    PegOperSQLite = eDEFAULT
    If Not loConn Is Nothing Then loConn.CloseDB
End Function
'---------------------------------------------------------------------------------------
' EQUIVALENTE A: ADOComando
'---------------------------------------------------------------------------------------
Public Function SQLiteComando(ByVal cCON As String, ByVal cSQL As String) As Boolean
    Dim loConn As New SQLiteConnection
    On Error GoTo Erro
    
    ' Limpeza da Tag e tradução do dialeto
    cCON = LimpaTag(cCON)
    cSQL = SQLDialeto(cSQL, "SQLITE")
    
    loConn.OpenDB cCON
    loConn.Execute cSQL
    
    SQLiteComando = True
    loConn.CloseDB
    Exit Function
Erro:
    SQLiteComando = False
End Function

'---------------------------------------------------------------------------------------
' EQUIVALENTE A: PegSQLAdo (Espelho Fiel)
'---------------------------------------------------------------------------------------
Public Function PegSQLite(ByVal cCON As String, ByVal cSQL As String, _
                         ByVal nITEM As Long, ByVal aCAM As Variant, _
                         ByVal aFOR As Variant, ByVal aPAD As Variant) As Variant
    Dim loConn As New SQLiteConnection
    Dim loCursor As SQLiteCursor
    Dim i As Integer
    Dim aVAL() As Variant
    
    On Error GoTo Erro
    
    ' 1. Inicializa o array de retorno com o tamanho de nITEM (ou baseado no aCAM)
    ReDim aVAL(nITEM)
    
    cCON = LimpaTag(cCON)
    loConn.OpenDB cCON
    
    ' 2. Abre o cursor (o SQL deve solicitar as colunas na ordem do aCAM)
    Set loCursor = loConn.CreateCursor(SQLDialeto(cSQL, "SQLITE"))
    
    If Not loCursor.EOF Then
        ' 3. Loop de processamento de campos, formatos e padrões
        For i = 0 To nITEM
            Dim vValor As Variant
            vValor = loCursor.Value(i)
            
            ' Tratamento de NULL ou Vazio usando o valor padrão (aPAD)
            If IsNull(vValor) Or vValor = "" Then
                vValor = aPAD(i)
            End If
            
            ' Tipagem baseada no array de formatos (aFOR)
            ' C = Caracter, N = Numérico, D = Data, etc.
            Select Case UCase(aFOR(i))
                Case "C", "S" ' String / Caracter
                    aVAL(i) = CStr(vValor)
                Case "N", "I", "NI", "R" ' Numérico (Inteiro ou Real)
                    aVAL(i) = Val(Replace(CStr(vValor), ",", "."))
                Case "D" ' Data
                    If IsDate(vValor) Then aVAL(i) = CDate(vValor) Else aVAL(i) = aPAD(i)
                Case Else
                    aVAL(i) = FVar(vValor)
            End Select
        Next
        PegSQLite = aVAL ' Retorna o array preenchido
    Else
        ' Se não encontrar registro, retorna o array preenchido com os padrões (aPAD)
        PegSQLite = aPAD
    End If
    
    ' Limpeza de memória e fechamento
    loConn.CloseDB
    Set loCursor = Nothing
    Set loConn = Nothing
    Exit Function

Erro:
    PegSQLite = aPAD ' Em caso de erro, retorna os valores padrão por segurança
    If Not loConn Is Nothing Then
        On Error Resume Next
        loConn.CloseDB
        Set loConn = Nothing
    End If
End Function

'---------------------------------------------------------------------------------------
' EQUIVALENTE A: GrvSQLado (Revisado para suportar arrays de campos e valores)
'---------------------------------------------------------------------------------------
Public Function GrvSQLite(ByVal cCON As String, ByVal cSQL_SELECT As String, _
                         ByVal aCAM As Variant, ByVal aVAL As Variant) As Boolean
    Dim loConn As New SQLiteConnection
    Dim cSQL_UPDATE As String
    Dim cTABELA As String
    Dim cWHERE As String
    Dim i As Integer
    Dim nPosFrom As Long, nPosWhere As Long
    
    On Error GoTo Erro
    
    ' 1. Extração da Tabela e do WHERE a partir do SELECT enviado
    nPosFrom = InStr(1, UCase(cSQL_SELECT), " FROM ") + 6
    nPosWhere = InStr(1, UCase(cSQL_SELECT), " WHERE ")
    
    If nPosWhere > 0 Then
        cTABELA = Trim(Mid(cSQL_SELECT, nPosFrom, nPosWhere - nPosFrom))
        cWHERE = Mid(cSQL_SELECT, nPosWhere) ' Captura o "WHERE ..."
    Else
        ' Segurança: não permite update sem WHERE para não afetar a tabela toda
        Exit Function
    End If
    
    ' 2. Montagem do comando UPDATE dinâmico
    cSQL_UPDATE = "UPDATE " & cTABELA & " SET "
    For i = LBound(aCAM) To UBound(aCAM)
        ' Usa a sua função FormataParaSQL para tratar Aspas, Datas e Decimais
        cSQL_UPDATE = cSQL_UPDATE & aCAM(i) & " = " & FormataParaSQL(aVAL(i)) & _
                      IIf(i < UBound(aCAM), ", ", " ")
    Next
    cSQL_UPDATE = cSQL_UPDATE & cWHERE
    
    ' 3. Execução direta na conexão
    cCON = LimpaTag(cCON)
    loConn.OpenDB cCON
    loConn.Execute SQLDialeto(cSQL_UPDATE, "SQLITE")
    
    GrvSQLite = True
    
    ' 4. Fechamento correto: Primeiro Close, depois Nothing
    loConn.CloseDB
    Set loConn = Nothing
    Exit Function
    
Erro:
    GrvSQLite = False
    If Not loConn Is Nothing Then
        On Error Resume Next
        loConn.CloseDB
        Set loConn = Nothing
    End If
End Function
'---------------------------------------------------------------------------------------
' EQUIVALENTE A: IncluiSQLAdo[cite: 2]
'---------------------------------------------------------------------------------------
Public Function IncluiSQLite(ByVal cCON As String, ByVal cTABELA As String, ByVal aCAMPOS As Variant, ByVal aVALORES As Variant, Optional ByRef nID As Long) As Boolean
    Dim loConn As New SQLiteConnection
    Dim cSQL As String
    Dim i As Integer
    
    On Error GoTo Erro
    
    cCON = LimpaTag(cCON)
    
    ' Montagem do INSERT manual (mais seguro para SQLite Nativo)
    cSQL = "INSERT INTO " & cTABELA & " ("
    For i = LBound(aCAMPOS) To UBound(aCAMPOS)
        cSQL = cSQL & aCAMPOS(i) & IIf(i < UBound(aCAMPOS), ",", "")
    Next
    cSQL = cSQL & ") VALUES ("
    For i = LBound(aVALORES) To UBound(aVALORES)
        cSQL = cSQL & FormataParaSQL(aVALORES(i)) & IIf(i < UBound(aVALORES), ",", "")
    Next
    cSQL = cSQL & ")"
    
    loConn.OpenDB cCON
    loConn.Execute cSQL
    
    ' Captura o ID gerado automaticamente
    nID = loConn.LastInsertRowID
    
    IncluiSQLite = True
    loConn.CloseDB
    
    Set loConn = Nothing   ' Adicione isso
    Exit Function
Erro:
    IncluiSQLite = False
End Function

'---------------------------------------------------------------------------------------
' FUNÇÕES AUXILIARES DE SUPORTE
'---------------------------------------------------------------------------------------

Private Function LimpaTag(ByVal cCON As String) As String
    ' Remove a tag customizada para obter o path puro do arquivo
    LimpaTag = Replace(cCON, "[SQLITEDIRETO]", "")
End Function

Private Function FormataParaSQL(ByVal vValor As Variant) As String
    ' Garante que o dado vá no formato correto para o SQLite[cite: 1, 2]
    If IsNull(vValor) Or vValor = "" Then
        FormataParaSQL = "NULL"
    ElseIf IsDate(vValor) Then
        FormataParaSQL = "'" & Format(vValor, "YYYY-MM-DD HH:MM:SS") & "'"
    ElseIf IsNumeric(vValor) Then
        FormataParaSQL = Replace(CStr(vValor), ",", ".")
    Else
        ' Trata aspas simples dentro de strings[cite: 1]
        FormataParaSQL = "'" & Replace(vValor, "'", "''") & "'"
    End If
End Function

'---------------------------------------------------------------------------------------
' EQUIVALENTES A: SomaSQLAdo, PegCountSQLADO, PegMINSQLADO, PegMAXSQLADO, PegSUMSQLADO
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' EQUIVALENTE A: SomaSQLAdo
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' EQUIVALENTE FIEL A: SomaSQLAdo (Com lógica SepSqlOpe e MathOper)
'---------------------------------------------------------------------------------------
Public Function SomaSQLite(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                           ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    Dim loConn As New SQLiteConnection
    Dim loCursor As SQLiteCursor
    Dim cSQL As String
    Dim aOPER() As String
    Dim cCAMPO_PROC As String
    
    On Error GoTo Erro
    
    ' REPLICAÇÃO DA LÓGICA ORIGINAL:
    ' SepSqlOpe divide a string se houver operadores (ex: "PRECO*QTD" -> ["PRECO", "QTD", "*"])
    aOPER = SepSqlOpe(cCAMPO)
    
    ' MathOper reconstrói a expressão de forma que o banco de dados entenda a operação
    ' ou trata prioridades de cálculo conforme a lógica da sua biblioteca original.
    cCAMPO_PROC = MathOper(aOPER)
    
    ' Monta o SQL com o campo processado
    cSQL = "SELECT SUM(" & cCAMPO_PROC & ") FROM " & cTABLEWHERE
    cSQL = SQLDialeto(cSQL, "SQLITE")
    
    loConn.OpenDB LimpaTag(cCON)
    Set loCursor = loConn.CreateCursor(cSQL)
    
    If Not loCursor.EOF Then
        If IsNull(loCursor.Value(0)) Then
            SomaSQLite = eDEFAULT
        Else
            SomaSQLite = loCursor.Value(0)
        End If
    Else
        SomaSQLite = eDEFAULT
    End If
    
    loConn.CloseDB
    Set loCursor = Nothing
    Exit Function

Erro:
    SomaSQLite = eDEFAULT
    If Not loConn Is Nothing Then loConn.CloseDB
End Function





'---------------------------------------------------------------------------------------
' EQUIVALENTE A: PegCampoSQLADO / PegCampoSQL
'---------------------------------------------------------------------------------------
Public Function PegCampoSQLite(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                               ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    
    ' O PegCampo no ADO nada mais é do que uma operação de busca simples.
    ' Usamos a lógica de "MAX" ou apenas o campo puro se a função PegOper aceitar.
    ' Para garantir compatibilidade total, chamamos a PegOper sem função agregadora
    ' ou simulamos via MAX para pegar o valor único.
    
    PegCampoSQLite = PegOperSQLite(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "")
    
End Function
Public Function PegSQLiteDeli(ByVal cCON As String, ByVal cSQL As String, _
                             ByVal aCAM As Variant, ByVal aFOR As Variant, _
                             ByVal aPAD As Variant, Optional ByVal cDELI As String = ",") As Variant
    Dim loConn As New SQLiteConnection
    Dim loCursor As SQLiteCursor
    Dim i As Integer, cRET As String
    
    On Error GoTo Erro
    cCON = LimpaTag(cCON)
    loConn.OpenDB cCON
    Set loCursor = loConn.CreateCursor(SQLDialeto(cSQL, "SQLITE"))
    
    If Not loCursor.EOF Then
        For i = LBound(aCAM) To UBound(aCAM)
            Dim vVal As Variant
            vVal = loCursor.Value(i)
            
            ' Aplica Padrão se Null
            If IsNull(vVal) Or vVal = "" Then vVal = aPAD(i)
            
            ' Formatação simplificada para string delimitada
            If UCase(aFOR(i)) = "N" Then
                vVal = Replace(CStr(vVal), ",", ".") ' Garante ponto em números
            End If
            
            cRET = cRET & CStr(vVal) & IIf(i < UBound(aCAM), cDELI, "")
        Next
        PegSQLiteDeli = cRET
    Else
        PegSQLiteDeli = ""
    End If
    
    loConn.CloseDB
    Set loCursor = Nothing: Set loConn = Nothing
    Exit Function
Erro:
    PegSQLiteDeli = ""
End Function

'---------------------------------------------------------------------------------------
' EQUIVALENTE A: PegUltSQLAdo
'---------------------------------------------------------------------------------------
Public Function PegUltSQLite(ByVal cCON As String, ByVal cTABELA As String) As Long
    Dim vRet As Variant
    
    ' No SQLite, last_insert_rowid() retorna o último ID da conexão atual.
    ' O parâmetro cTABELA é mantido para manter a assinatura idêntica ao ADO,
    ' embora o SQLite não exija o nome da tabela para esta função específica.
    
    vRet = PegOperSQLite(cCON, cTABELA, "last_insert_rowid()", 0, "")
    
    If IsNumeric(vRet) Then
        PegUltSQLite = CLng(vRet)
    Else
        PegUltSQLite = 0
    End If
End Function


