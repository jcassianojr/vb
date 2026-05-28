Attribute VB_Name = "vbsqlite"
Option Explicit
' Reference=*\G{7CC1A5F1-A0FF-4546-A0F1-FBFE744A4522}#1.1#0#..\..\..\..\..\..\WINDOWS\system32\VBSQLite12.DLL#VB SQLite Library 1.2
' Notas de Implementação:
' 1. Requer as classes cSQLiteConnection e cSQLiteCursor do projeto VBSQLite.
' 2. Mantém compatibilidade com a função SQLDialeto do adolib.bas.

'---------------------------------------------------------------------------------------
' WRAPPERS DE OPERAÇÃO - ESPELHOS DA SQLFUNCOESADO
'---------------------------------------------------------------------------------------
Public Function PegUltSQLite(ByVal cCON As String, ByVal cSQL As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    Dim loConn As New SQLiteConnection
    Dim loCursor As SQLiteCursor
    Dim vUltimo As Variant
    
    On Error GoTo Erro
    vUltimo = eDEFAULT
    
    loConn.OpenDB cCON
    
    VBSQLiteSetValues loConn
    ' O método correto é CreateCursor
    Set loCursor = loConn.CreateCursor(cSQL)
    
    If Not loCursor.EOF Then
        ' Como cursores SQLite são geralmente forward-only,
        ' percorremos até o final para pegar o último registro da consulta
        While Not loCursor.EOF
            vUltimo = loCursor.Value(cCAMPO)
            loCursor.MoveNext
        Wend
    End If
    
    ' Trata nulo se o campo existir mas estiver vazio
    PegUltSQLite = IIf(IsNull(vUltimo), eDEFAULT, vUltimo)
    
    loConn.CloseDB
    Set loCursor = Nothing
    Set loConn = Nothing
    Exit Function
    
Erro:
    PegUltSQLite = eDEFAULT
    If Not loConn Is Nothing Then loConn.CloseDB
End Function ' EQUIVALENTE A: PegMINSQLADO
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
    cSQL = sqldialeto(cSQL, "SQLITE")
    
    loConn.OpenDB LimpaTag(cCON)
    
    VBSQLiteSetValues loConn
    
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
    cSQL = sqldialeto(cSQL, "SQLITE")
    
    loConn.OpenDB cCON
    VBSQLiteSetValues loConn
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
    
    VBSQLiteSetValues loConn
    
    ' 2. Abre o cursor (o SQL deve solicitar as colunas na ordem do aCAM)
    Set loCursor = loConn.CreateCursor(sqldialeto(cSQL, "SQLITE"))
    
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
Public Function GrvSQLite(ByVal cARQ As String, _
                         ByVal cSQL_SELECT As String, _
                         ByVal nITEM As Long, _
                         ByVal aCAM As Variant, _
                         ByVal aVAL As Variant, _
                         ByVal aFOR As Variant, _
                         Optional ByVal nStartItem As Long = 0) As Boolean
    
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
        cWHERE = Mid(cSQL_SELECT, nPosWhere)
    Else
        Exit Function
    End If
    
    ' 2. Montagem do comando UPDATE dinâmico
    ' Note que agora usamos os parâmetros aCAM e aVAL que chegaram na função
    cSQL_UPDATE = "UPDATE " & cTABELA & " SET "
    
    ' Usamos nStartItem para que o loop comece de onde o sistema espera
    For i = nStartItem To (nITEM - 1)
        ' Aqui chamamos a função que trata o valor conforme o dialeto
        cSQL_UPDATE = cSQL_UPDATE & aCAM(i) & " = " & TratarValorParaSQL(aVAL(i), aFOR(i), "SQLITE") & _
                      IIf(i < (nITEM - 1), ", ", " ")
    Next
    
    cSQL_UPDATE = cSQL_UPDATE & cWHERE
    
    ' 3. Execução direta na conexão
    ' (Mantendo sua lógica de limpeza de tag)
    cARQ = LimpaTag(cARQ)
    loConn.OpenDB cARQ
    VBSQLiteSetValues loConn
    
    loConn.Execute cSQL_UPDATE
    
    GrvSQLite = True
    
    ' 4. Fechamento correto
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
Public Function IncluiSQLite(ByVal cARQ As String, _
                             ByVal cSQL_SELECT As String, _
                             ByVal nITEM As Long, _
                             ByVal aCAM As Variant, _
                             ByVal aVAL As Variant, _
                             ByVal lCHECK As Boolean, _
                             ByVal lMES As Boolean, _
                             ByVal aIDDES As Variant) As Boolean
    
    Dim loConn As New SQLiteConnection
    Dim loCursor As SQLiteCursor
    Dim cTABELA As String
    Dim i As Integer
    Dim cCampos As String, cValores As String
    
    On Error GoTo Erro
    
    ' 1. Extração da Tabela (O ADO provavelmente usa um SELECT fake como "SELECT * FROM Tabela WHERE 1=0")
    ' Vamos extrair o nome da tabela do seu cSQL
    cTABELA = ExtrairNomeTabela(cSQL_SELECT)
    
    ' 2. Checagem (lCHECK)
    If lCHECK Then
        ' Aqui você implementaria a lógica de verificar se o registro já existe
        ' antes de prosseguir com o INSERT.
    End If
    
    ' 3. Montagem do INSERT
    ' O padrão é: "INSERT INTO Tabela (Campos) VALUES (Valores)"
    cCampos = ""
    cValores = ""
    For i = 0 To nITEM - 1
        cCampos = cCampos & IIf(cCampos = "", "", ", ") & aCAM(i)
        cValores = cValores & IIf(cValores = "", "", ", ") & TratarValorParaSQL(aVAL(i), "C", "SQLITE")
    Next i
    
    ' 4. Execução
    loConn.OpenDB cARQ
    VBSQLiteSetValues loConn
    loConn.Execute "INSERT INTO " & cTABELA & " (" & cCampos & ") VALUES (" & cValores & ")"
    
    ' 5. Captura do ID (aIDDES) - Equivalente ao que o seu ADO faz
    ' SQLite tem o comando last_insert_rowid()
    If Not IsNumeric(aIDDES) Then
        ' Se aIDDES for uma matriz de campos de ID, buscamos o último valor inserido
        ' ... lógica para popular o retorno conforme seu padrão ...
    End If
    
    IncluiSQLite = True
    loConn.CloseDB
    Set loConn = Nothing
    Exit Function

Erro:
    If lMES Then MsgBox "Erro ao incluir: " & Err.Description
    IncluiSQLite = False
    ' ... fechar conexões ...
End Function
'---------------------------------------------------------------------------------------
' FUNÇÕES AUXILIARES DE SUPORTE
'---------------------------------------------------------------------------------------

Private Function LimpaTag(ByVal cCON As String) As String
    ' Remove a tag customizada para obter o path puro do arquivo
    LimpaTag = Replace(cCON, "[VBSQLITE]", "")
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
                             ByVal aCAM As Variant, Optional ByVal cDELI As String = ",", _
                             Optional ByVal aPAD As Variant = "", Optional ByVal aFOR As Variant = "") As Variant
    Dim loConn As New SQLiteConnection, loCursor As SQLiteCursor
    Dim x As Long, nCAMPOS As Integer
    Dim aRETU As Variant, aOPE As Variant, eVAL As Variant

    On Error GoTo Erro

    nCAMPOS = UBound(aCAM) + 1
    ReDim aRETU(nCAMPOS - 1)
    For x = 0 To nCAMPOS - 1: aRETU(x) = "": Next x

    loConn.OpenDB cCON
    VBSQLiteSetValues loConn
    Set loCursor = loConn.CreateCursor(cSQL)

    If Not loCursor.EOF Then
        While Not loCursor.EOF
            For x = 0 To nCAMPOS - 1
                ' Lógica de Operações (SepSqlOpe / MathOper)
                aOPE = SepSqlOpe(aCAM(x))
                If aOPE(0) = "" Or aOPE(1) = "" Or aOPE(2) = "" Then
                    eVAL = loCursor.Value(aCAM(x))
                Else
                    eVAL = MathOper(loCursor.Value(aOPE(1)), loCursor.Value(aOPE(2)), aOPE(0))
                End If

                ' Tratamento de Nulo
                If IsNull(eVAL) Then
                    If IsArray(aPAD) Then eVAL = aPAD(x) Else eVAL = aPAD
                End If

                ' Formatação com FVar (MyFunctions)
                If IsArray(aFOR) Then
                    If IsArray(aPAD) Then
                        eVAL = FVar(eVAL, aFOR(x), aPAD(x))
                    Else
                        eVAL = FVar(eVAL, aFOR(x))
                    End If
                End If

                aRETU(x) = aRETU(x) & FixStr(eVAL)
            Next x

            loCursor.MoveNext
            If Not loCursor.EOF Then
                For x = 0 To nCAMPOS - 1: aRETU(x) = aRETU(x) & cDELI: Next x
            End If
        Wend
    End If

    loConn.CloseDB
    PegSQLiteDeli = aRETU
    Exit Function
Erro:
    PegSQLiteDeli = aRETU
End Function

'---------------------------------------------------------------------------------------
' EQUIVALENTE A: PegUltSQLAdo
'---------------------------------------------------------------------------------------
Public Function PegLastidbsqLite(ByVal cCON As String, ByVal cTABELA As String) As Long
    Dim vRet As Variant
    
    ' No SQLite, last_insert_rowid() retorna o último ID da conexão atual.
    ' O parâmetro cTABELA é mantido para manter a assinatura idêntica ao ADO,
    ' embora o SQLite não exija o nome da tabela para esta função específica.
    
    vRet = PegOperSQLite(cCON, cTABELA, "last_insert_rowid()", 0, "")
    
    If IsNumeric(vRet) Then
        PegLastidbsqLite = CLng(vRet)
    Else
        PegLastidbsqLite = 0
    End If
End Function

'---------------------------------------------------------------------------------------
' EQUIVALENTE FIEL A: SomaSQLAdo
'---------------------------------------------------------------------------------------
Public Function SomaSQLite(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                           ByVal cCAMPO As String, Optional ByVal eDEFAULT As Variant, _
                           Optional ByVal nDEC As Integer = 2) As Variant
    Dim loConn As New SQLiteConnection
    Dim loCursor As SQLiteCursor
    Dim nSoma As Double
    Dim aOPER As Variant
    Dim aVALORES_LINHA() As String
    Dim x As Integer
    
    On Error GoTo Erro
    
    ' SepSqlOpe prepara o array de campos/operadores
    aOPER = SepSqlOpe(cCAMPO)
    
    loConn.OpenDB LimpaTag(cCON)
    VBSQLiteSetValues loConn
    ' Busca todos os registros para processamento linha a linha (como a original)
    Set loCursor = loConn.CreateCursor("SELECT * FROM " & cTABLEWHERE)
    
    nSoma = 0
    While Not loCursor.EOF
        ReDim aVALORES_LINHA(UBound(aOPER))
        For x = 0 To UBound(aOPER)
            ' Se não for operador, extrai o valor do campo do cursor
            If InStr("+-*/()", aOPER(x)) = 0 And aOPER(x) <> "" Then
                aVALORES_LINHA(x) = loCursor.Value(aOPER(x)) & ""
            Else
                aVALORES_LINHA(x) = aOPER(x)
            End If
        Next x
        
        ' CHAMA MATHOPER COM OS DOIS PARAMETROS: o array e as decimais
        nSoma = nSoma + Val(MathOper(aVALORES_LINHA, nDEC))
        
        loCursor.MoveNext
    Wend
    
    SomaSQLite = IIf(nSoma = 0, eDEFAULT, nSoma)
    
    loConn.CloseDB
    Set loCursor = Nothing
    Exit Function
Erro:
    SomaSQLite = eDEFAULT
    If Not loConn Is Nothing Then loConn.CloseDB
End Function

'---------------------------------------------------------------------------------------
' EQUIVALENTE FIEL A: ApagaSQLpAdo / APAGASQLADO
' Conceito: Checa se existe DELETE, se não, monta a partir do FROM/WHERE
'---------------------------------------------------------------------------------------
Public Function ApagaSQLite(ByVal cCON As String, ByVal cSQL As String) As Boolean
    Dim nPOS As Long
    Dim cSQL_FINAL As String
    
    On Error GoTo Erro
    
    cSQL_FINAL = Trim(cSQL)
    
    ' Lógica de espelhamento: Se não começa com DELETE, mas tem FROM, reconstrói
    If UCase(Left(cSQL_FINAL, 6)) <> "DELETE" Then
        nPOS = InStr(1, UCase(cSQL_FINAL), "FROM")
        If nPOS > 0 Then
            ' Pega do FROM em diante e adiciona o DELETE
            cSQL_FINAL = "DELETE " & Mid(cSQL_FINAL, nPOS)
        Else
            ' Se for apenas "TABELA WHERE...", tenta montar
            cSQL_FINAL = "DELETE FROM " & cSQL_FINAL
        End If
    End If
    
    ' Executa através do comando padrão que já trata Dialeto e Conexão
    ApagaSQLite = SQLiteComando(cCON, cSQL_FINAL)
    
    Exit Function
Erro:
    ApagaSQLite = False
End Function

Public Function SQLMoveRegSQLite(ByVal cCONORI As String, ByVal cSQLORI As String, _
   Optional ByVal cOPEORI As String = "", Optional ByVal aCAMORI As Variant = 0, _
   Optional ByVal aOUTORI As Variant = 0, Optional ByVal cCONDES As String = "", _
   Optional ByVal cSQLDES As String = "", Optional ByVal cOPEDES As String = "", _
   Optional ByVal aCAMDES As Variant = 0, Optional ByVal aOUTDES As Variant = 0, _
   Optional ByVal aIDDES As Variant = 0) As Boolean

    Dim loConnOri As New SQLiteConnection, loConnDes As New SQLiteConnection
    Dim loCurOri As SQLiteCursor, loCurDes As SQLiteCursor
    Dim x As Long, nCAMPOS As Long, nRegs As Long
    Dim aVALORI As Variant, aRETUID As Variant, aOPE As Variant
    Dim cTABELA As String, cWHERE As String, cSET As String, cINS_C As String, cINS_V As String

    On Error GoTo Erro
    SQLMoveRegSQLite = False

    loConnOri.OpenDB cCONORI
    VBSQLiteSetValues loConnOri
    loConnDes.OpenDB cCONDES
    VBSQLiteSetValues loConnDes

    ' 1. Coleta dados da Origem (Passo Matriz)
    Set loCurOri = loConnOri.CreateCursor(cSQLORI)
    If Not loCurOri.EOF Then
        nCAMPOS = UBound(aCAMORI)
        ReDim aVALORI(nCAMPOS)
        For x = 0 To nCAMPOS
            aOPE = SepSqlOpe(aCAMORI(x))
            If aOPE(0) = "" Then
                aVALORI(x) = loCurOri.Value(aCAMORI(x))
            Else
                aVALORI(x) = MathOper(loCurOri.Value(aOPE(1)), loCurOri.Value(aOPE(2)), aOPE(0))
            End If
        Next x

        ' 2. Metadados do Destino
        cTABELA = ExtraiTabela(cSQLDES)
        cWHERE = ExtraiWhere(cSQLDES)

        ' 3. Lógica de Gravação via Execute
        Set loCurDes = loConnDes.CreateCursor("SELECT count(*) FROM " & cTABELA & " " & cWHERE)
        nRegs = loCurDes.Value(0)
        
        If nRegs > 0 Then
            ' --- UPDATE ---
            cSET = ""
            For x = 0 To UBound(aCAMDES)
                cSET = cSET & IIf(cSET = "", "", ", ") & aCAMDES(x) & " = " & PrepararValorSQL(aVALORI(x))
            Next x
            
            If IsArray(aOUTDES) Then
                For x = 0 To UBound(aOUTDES)
                    cSET = cSET & ", " & aOUTDES(x) & " = " & PrepararValorSQL(aOUTORI(x))
                Next x
            End If
            
            loConnDes.Execute "UPDATE " & cTABELA & " SET " & cSET & " " & cWHERE
        Else
            ' --- INSERT ---
            cINS_C = "": cINS_V = ""
            For x = 0 To UBound(aCAMDES)
                cINS_C = cINS_C & IIf(cINS_C = "", "", ", ") & aCAMDES(x)
                cINS_V = cINS_V & IIf(cINS_V = "", "", ", ") & PrepararValorSQL(aVALORI(x))
            Next x
            
            If IsArray(aOUTDES) Then
                For x = 0 To UBound(aOUTDES)
                    cINS_C = cINS_C & ", " & aOUTDES(x)
                    cINS_V = cINS_V & ", " & PrepararValorSQL(aOUTORI(x))
                Next x
            End If
            
            loConnDes.Execute "INSERT INTO " & cTABELA & " (" & cINS_C & ") VALUES (" & cINS_V & ")"
        End If

        ' 4. Captura de IDs para eRETU01
        If IsArray(aIDDES) Then
            Set loCurDes = loConnDes.CreateCursor(cSQLDES)
            ReDim aRETUID(UBound(aIDDES))
            For x = 0 To UBound(aIDDES)
                aRETUID(x) = loCurDes.Value(aIDDES(x))
            Next x
            eRETU01 = aRETUID
        End If
        SQLMoveRegSQLite = True
    End If

    loConnOri.CloseDB: loConnDes.CloseDB
    Exit Function
Erro:
    SQLMoveRegSQLite = False
End Function

Public Function VBSQLiteSetValues(ByRef oCONN As SQLiteConnection) As Boolean
    On Error GoTo Erro
    
    ' A VBSQLite permite aplicar configurações diretamente via Execute
    ' Algumas destas configurações são específicas para esta biblioteca
    
    ' 1. Aumenta a performance de escrita
    oCONN.Execute "PRAGMA journal_mode = WAL;"
    
    ' 2. Ajusta o tamanho da cache (em páginas)
    oCONN.Execute "PRAGMA cache_size = 2000;"
    
    ' 3. Melhora a segurança de escrita com performance otimizada
    oCONN.Execute "PRAGMA synchronous = NORMAL;"
    
    ' 4. Define o modo de armazenamento temporário
    oCONN.Execute "PRAGMA temp_store = MEMORY;"
    
    VBSQLiteSetValues = True
    Exit Function
    
Erro:
    VBSQLiteSetValues = False
End Function

Public Function TratarValorParaSQL(ByVal vValor As Variant, ByVal cTipo As String, ByVal cDialeto As String) As String
    ' cTipo: "C" para Caractere, "N" para Número, "D" para Data
    ' cDialeto: "SQLITE", "PGSQL", "VFP"
    
    If IsNull(vValor) Or vValor = "" Then
        TratarValorParaSQL = "NULL"
        Exit Function
    End If
    
    Select Case UCase(cDialeto)
        Case "SQLITE"
            If UCase(cTipo) = "D" Then
                ' SQLite armazena datas como string YYYY-MM-DD
                TratarValorParaSQL = "'" & Format(vValor, "yyyy-mm-dd") & "'"
            ElseIf UCase(cTipo) = "N" Then
                TratarValorParaSQL = Replace(vValor, ",", ".") ' Garante ponto decimal
            Else
                TratarValorParaSQL = "'" & Replace(vValor, "'", "''") & "'" ' Escapa aspas simples
            End If
            
        Case "PGSQL", "VFP"
            ' Aqui você mantém a lógica que já funcionava para os outros bancos
            TratarValorParaSQL = "'" & vValor & "'"
    End Select
End Function
