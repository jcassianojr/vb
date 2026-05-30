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
    
15:     On Error GoTo Erro
16:     vUltimo = eDEFAULT
    
18:     loConn.OpenDB cCON
    
20:     VBSQLiteSetValues loConn
    ' O método correto é CreateCursor
22:     Set loCursor = loConn.CreateCursor(cSQL)
    
24:     If Not loCursor.EOF Then
        ' Como cursores SQLite são geralmente forward-only,
        ' percorremos até o final para pegar o último registro da consulta
27:         While Not loCursor.EOF
28:             vUltimo = loCursor.Value(cCAMPO)
29:             loCursor.MoveNext
30:         Wend
31:     End If
    
    ' Trata nulo se o campo existir mas estiver vazio
34:     PegUltSQLite = IIf(IsNull(vUltimo), eDEFAULT, vUltimo)
    
36:     loConn.CloseDB
37:     Set loCursor = Nothing
38:     Set loConn = Nothing
39:     Exit Function
    
Erro:
42:     PegUltSQLite = eDEFAULT
43:     If Not loConn Is Nothing Then loConn.CloseDB
End Function ' EQUIVALENTE A: PegMINSQLADO
Public Function PegMinSQLite(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                            ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
47:     PegMinSQLite = PegOperSQLite(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "MIN")
End Function

' EQUIVALENTE A: PegMAXSQLADO
Public Function PegMaxSQLite(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                            ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
53:     PegMaxSQLite = PegOperSQLite(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "MAX")
End Function

' EQUIVALENTE A: PegSUMSQLADO
Public Function PegSumSQLite(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                            ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
59:     PegSumSQLite = PegOperSQLite(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "SUM")
End Function

' EQUIVALENTE A: PegCountSQLADO
Public Function PegCountSQLite(ByVal cCON As String, ByVal cTABLEWHERE As String, _
                              ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    ' No Count, se o campo for *, ele conta todos os registos
66:     PegCountSQLite = PegOperSQLite(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "COUNT")
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
    
79:     On Error GoTo Erro
    
    ' Montagem da query idêntica ao processo ADO
    ' Ex: SELECT SUM(VALOR) FROM MOVI WHERE CODIGO = 10
83:     cSQL = "SELECT " & coper & "(" & cCAMPO & ") FROM " & cTABLEWHERE
84:     cSQL = sqldialeto(cSQL, "SQLITE")
    
86:     loConn.OpenDB LimpaTag(cCON)
    
88:     VBSQLiteSetValues loConn
    
90:     Set loCursor = loConn.CreateCursor(cSQL)
    
92:     If Not loCursor.EOF Then
93:         If IsNull(loCursor.Value(0)) Then
94:             PegOperSQLite = eDEFAULT
95:         Else
96:             PegOperSQLite = loCursor.Value(0)
97:         End If
98:     Else
99:         PegOperSQLite = eDEFAULT
100:     End If
    
102:     loConn.CloseDB
103:     Set loCursor = Nothing: Set loConn = Nothing
104:     Exit Function

Erro:
107:     PegOperSQLite = eDEFAULT
108:     If Not loConn Is Nothing Then loConn.CloseDB
End Function
'---------------------------------------------------------------------------------------
' EQUIVALENTE A: ADOComando
'---------------------------------------------------------------------------------------
Public Function SQLiteComando(ByVal cCON As String, ByVal cSQL As String) As Boolean
    Dim loConn As New SQLiteConnection
115:     On Error GoTo Erro
    
    ' Limpeza da Tag e tradução do dialeto
118:     cCON = LimpaTag(cCON)
119:     cSQL = sqldialeto(cSQL, "SQLITE")
    
121:     loConn.OpenDB cCON
122:     VBSQLiteSetValues loConn
123:     loConn.Execute cSQL
    
125:     SQLiteComando = True
126:     loConn.CloseDB
127:     Exit Function
Erro:
129:     SQLiteComando = False
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
    
143:     On Error GoTo Erro
    
    ' 1. Inicializa o array de retorno com o tamanho de nITEM (ou baseado no aCAM)
    ReDim aVAL(nITEM)
    
148:     cCON = LimpaTag(cCON)
149:     loConn.OpenDB cCON
    
151:     VBSQLiteSetValues loConn
    
    ' 2. Abre o cursor (o SQL deve solicitar as colunas na ordem do aCAM)
154:     Set loCursor = loConn.CreateCursor(sqldialeto(cSQL, "SQLITE"))
    
156:     If Not loCursor.EOF Then
        ' 3. Loop de processamento de campos, formatos e padrões
158:         For i = 0 To nITEM
            Dim vValor As Variant
160:             vValor = loCursor.Value(i)
            
            ' Tratamento de NULL ou Vazio usando o valor padrão (aPAD)
163:             If IsNull(vValor) Or vValor = "" Then
164:                 vValor = aPAD(i)
165:             End If
            
            ' Tipagem baseada no array de formatos (aFOR)
            ' C = Caracter, N = Numérico, D = Data, etc.
169:             Select Case UCase(aFOR(i))
                Case "C", "S" ' String / Caracter
171:                     aVAL(i) = CStr(vValor)
                Case "N", "I", "NI", "R" ' Numérico (Inteiro ou Real)
173:                     aVAL(i) = Val(Replace(CStr(vValor), ",", "."))
                Case "D" ' Data
175:                     If IsDate(vValor) Then aVAL(i) = CDate(vValor) Else aVAL(i) = aPAD(i)
                Case Else
177:                     aVAL(i) = FVar(vValor)
178:             End Select
179:         Next
180:         PegSQLite = aVAL ' Retorna o array preenchido
181:     Else
        ' Se não encontrar registro, retorna o array preenchido com os padrões (aPAD)
183:         PegSQLite = aPAD
184:     End If
    
    ' Limpeza de memória e fechamento
187:     loConn.CloseDB
188:     Set loCursor = Nothing
189:     Set loConn = Nothing
190:     Exit Function

Erro:
193:     PegSQLite = aPAD ' Em caso de erro, retorna os valores padrão por segurança
194:     If Not loConn Is Nothing Then
195:         On Error Resume Next
196:         loConn.CloseDB
197:         Set loConn = Nothing
198:     End If
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
    
219:     On Error GoTo Erro
    
    ' 1. Extração da Tabela e do WHERE a partir do SELECT enviado
222:     nPosFrom = InStr(1, UCase(cSQL_SELECT), " FROM ") + 6
223:     nPosWhere = InStr(1, UCase(cSQL_SELECT), " WHERE ")
    
225:     If nPosWhere > 0 Then
226:         cTABELA = Trim(Mid(cSQL_SELECT, nPosFrom, nPosWhere - nPosFrom))
227:         cWHERE = Mid(cSQL_SELECT, nPosWhere)
228:     Else
229:         Exit Function
230:     End If
    
    ' 2. Montagem do comando UPDATE dinâmico
    ' Note que agora usamos os parâmetros aCAM e aVAL que chegaram na função
234:     cSQL_UPDATE = "UPDATE " & cTABELA & " SET "
    
    ' Usamos nStartItem para que o loop comece de onde o sistema espera
237:     For i = nStartItem To (nITEM - 1)
        ' Aqui chamamos a função que trata o valor conforme o dialeto
        cSQL_UPDATE = cSQL_UPDATE & aCAM(i) & " = " & TratarValorParaSQL(aVAL(i), aFOR(i), "SQLITE") & _
                      IIf(i < (nITEM - 1), ", ", " ")
241:     Next
    
243:     cSQL_UPDATE = cSQL_UPDATE & cWHERE
    
    ' 3. Execução direta na conexão
    ' (Mantendo sua lógica de limpeza de tag)
247:     cARQ = LimpaTag(cARQ)
248:     loConn.OpenDB cARQ
249:     VBSQLiteSetValues loConn
    
251:     loConn.Execute cSQL_UPDATE
    
253:     GrvSQLite = True
    
    ' 4. Fechamento correto
256:     loConn.CloseDB
257:     Set loConn = Nothing
258:     Exit Function
    
Erro:
261:     GrvSQLite = False
262:     If Not loConn Is Nothing Then
263:         On Error Resume Next
264:         loConn.CloseDB
265:         Set loConn = Nothing
266:     End If
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
    
284:     On Error GoTo Erro
    
    ' 1. Extração da Tabela (O ADO provavelmente usa um SELECT fake como "SELECT * FROM Tabela WHERE 1=0")
    ' Vamos extrair o nome da tabela do seu cSQL
288:     cTABELA = ExtrairNomeTabela(cSQL_SELECT)
    
    ' 2. Checagem (lCHECK)
291:     If lCHECK Then
        ' Aqui você implementaria a lógica de verificar se o registro já existe
        ' antes de prosseguir com o INSERT.
294:     End If
    
    ' 3. Montagem do INSERT
    ' O padrão é: "INSERT INTO Tabela (Campos) VALUES (Valores)"
298:     cCampos = ""
299:     cValores = ""
300:     For i = 0 To nITEM - 1
301:         cCampos = cCampos & IIf(cCampos = "", "", ", ") & aCAM(i)
302:         cValores = cValores & IIf(cValores = "", "", ", ") & TratarValorParaSQL(aVAL(i), "C", "SQLITE")
303:     Next i
    
    ' 4. Execução
306:     loConn.OpenDB cARQ
307:     VBSQLiteSetValues loConn
308:     loConn.Execute "INSERT INTO " & cTABELA & " (" & cCampos & ") VALUES (" & cValores & ")"
    
    ' 5. Captura do ID (aIDDES) - Equivalente ao que o seu ADO faz
    ' SQLite tem o comando last_insert_rowid()
312:     If Not IsNumeric(aIDDES) Then
        ' Se aIDDES for uma matriz de campos de ID, buscamos o último valor inserido
        ' ... lógica para popular o retorno conforme seu padrão ...
315:     End If
    
317:     IncluiSQLite = True
318:     loConn.CloseDB
319:     Set loConn = Nothing
320:     Exit Function

Erro:
323:     If lMES Then MsgBox "Erro ao incluir: " & Err.Description
324:     IncluiSQLite = False
    ' ... fechar conexões ...
End Function
'---------------------------------------------------------------------------------------
' FUNÇÕES AUXILIARES DE SUPORTE
'---------------------------------------------------------------------------------------

Private Function LimpaTag(ByVal cCON As String) As String
    ' Remove a tag customizada para obter o path puro do arquivo
333:     LimpaTag = Replace(cCON, "[VBSQLITE]", "")
End Function

Private Function FormataParaSQL(ByVal vValor As Variant) As String
    ' Garante que o dado vá no formato correto para o SQLite[cite: 1, 2]
338:     If IsNull(vValor) Or vValor = "" Then
339:         FormataParaSQL = "NULL"
340:     ElseIf IsDate(vValor) Then
341:         FormataParaSQL = "'" & Format(vValor, "YYYY-MM-DD HH:MM:SS") & "'"
342:     ElseIf IsNumeric(vValor) Then
343:         FormataParaSQL = Replace(CStr(vValor), ",", ".")
344:     Else
        ' Trata aspas simples dentro de strings[cite: 1]
346:         FormataParaSQL = "'" & Replace(vValor, "'", "''") & "'"
347:     End If
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
    
360:     PegCampoSQLite = PegOperSQLite(cCON, cTABLEWHERE, cCAMPO, eDEFAULT, "")
    
End Function
Public Function PegSQLiteDeli(ByVal cCON As String, ByVal cSQL As String, _
                             ByVal aCAM As Variant, Optional ByVal cDELI As String = ",", _
                             Optional ByVal aPAD As Variant = "", Optional ByVal aFOR As Variant = "") As Variant
    Dim loConn As New SQLiteConnection, loCursor As SQLiteCursor
    Dim x As Long, nCAMPOS As Integer
    Dim aRETU As Variant, aOPE As Variant, eVAL As Variant

370:     On Error GoTo Erro

372:     nCAMPOS = UBound(aCAM) + 1
    ReDim aRETU(nCAMPOS - 1)
374:     For x = 0 To nCAMPOS - 1: aRETU(x) = "": Next x

376:     loConn.OpenDB cCON
377:     VBSQLiteSetValues loConn
378:     Set loCursor = loConn.CreateCursor(cSQL)

380:     If Not loCursor.EOF Then
381:         While Not loCursor.EOF
382:             For x = 0 To nCAMPOS - 1
                ' Lógica de Operações (SepSqlOpe / MathOper)
384:                 aOPE = SepSqlOpe(aCAM(x))
385:                 If aOPE(0) = "" Or aOPE(1) = "" Or aOPE(2) = "" Then
386:                     eVAL = loCursor.Value(aCAM(x))
387:                 Else
388:                     eVAL = MathOper(loCursor.Value(aOPE(1)), loCursor.Value(aOPE(2)), aOPE(0))
389:                 End If

                ' Tratamento de Nulo
392:                 If IsNull(eVAL) Then
393:                     If IsArray(aPAD) Then eVAL = aPAD(x) Else eVAL = aPAD
394:                 End If

                ' Formatação com FVar (MyFunctions)
397:                 If IsArray(aFOR) Then
398:                     If IsArray(aPAD) Then
399:                         eVAL = FVar(eVAL, aFOR(x), aPAD(x))
400:                     Else
401:                         eVAL = FVar(eVAL, aFOR(x))
402:                     End If
403:                 End If

405:                 aRETU(x) = aRETU(x) & FixStr(eVAL)
406:             Next x

408:             loCursor.MoveNext
409:             If Not loCursor.EOF Then
410:                 For x = 0 To nCAMPOS - 1: aRETU(x) = aRETU(x) & cDELI: Next x
411:             End If
412:         Wend
413:     End If

415:     loConn.CloseDB
416:     PegSQLiteDeli = aRETU
417:     Exit Function
Erro:
419:     PegSQLiteDeli = aRETU
End Function

'---------------------------------------------------------------------------------------
' EQUIVALENTE A: PegUltSQLAdo
'---------------------------------------------------------------------------------------
Public Function PegLastidbsqLite(ByVal cCON As String, ByVal cTABELA As String) As Long
    Dim vRet As Variant
    
    ' No SQLite, last_insert_rowid() retorna o último ID da conexão atual.
    ' O parâmetro cTABELA é mantido para manter a assinatura idêntica ao ADO,
    ' embora o SQLite não exija o nome da tabela para esta função específica.
    
432:     vRet = PegOperSQLite(cCON, cTABELA, "last_insert_rowid()", 0, "")
    
434:     If IsNumeric(vRet) Then
435:         PegLastidbsqLite = CLng(vRet)
436:     Else
437:         PegLastidbsqLite = 0
438:     End If
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
    
454:     On Error GoTo Erro
    
    ' SepSqlOpe prepara o array de campos/operadores
457:     aOPER = SepSqlOpe(cCAMPO)
    
459:     loConn.OpenDB LimpaTag(cCON)
460:     VBSQLiteSetValues loConn
    ' Busca todos os registros para processamento linha a linha (como a original)
462:     Set loCursor = loConn.CreateCursor("SELECT * FROM " & cTABLEWHERE)
    
464:     nSoma = 0
465:     While Not loCursor.EOF
        ReDim aVALORES_LINHA(UBound(aOPER))
467:         For x = 0 To UBound(aOPER)
            ' Se não for operador, extrai o valor do campo do cursor
469:             If InStr("+-*/()", aOPER(x)) = 0 And aOPER(x) <> "" Then
470:                 aVALORES_LINHA(x) = loCursor.Value(aOPER(x)) & ""
471:             Else
472:                 aVALORES_LINHA(x) = aOPER(x)
473:             End If
474:         Next x
        
        ' CHAMA MATHOPER COM OS DOIS PARAMETROS: o array e as decimais
477:         nSoma = nSoma + Val(MathOper(aVALORES_LINHA, nDEC))
        
479:         loCursor.MoveNext
480:     Wend
    
482:     SomaSQLite = IIf(nSoma = 0, eDEFAULT, nSoma)
    
484:     loConn.CloseDB
485:     Set loCursor = Nothing
486:     Exit Function
Erro:
488:     SomaSQLite = eDEFAULT
489:     If Not loConn Is Nothing Then loConn.CloseDB
End Function

'---------------------------------------------------------------------------------------
' EQUIVALENTE FIEL A: ApagaSQLpAdo / APAGASQLADO
' Conceito: Checa se existe DELETE, se não, monta a partir do FROM/WHERE
'---------------------------------------------------------------------------------------
Public Function ApagaSQLite(ByVal cCON As String, ByVal cSQL As String) As Boolean
    Dim nPOS As Long
    Dim cSQL_FINAL As String
    
500:     On Error GoTo Erro
    
502:     cSQL_FINAL = Trim(cSQL)
    
    ' Lógica de espelhamento: Se não começa com DELETE, mas tem FROM, reconstrói
505:     If UCase(Left(cSQL_FINAL, 6)) <> "DELETE" Then
506:         nPOS = InStr(1, UCase(cSQL_FINAL), "FROM")
507:         If nPOS > 0 Then
            ' Pega do FROM em diante e adiciona o DELETE
509:             cSQL_FINAL = "DELETE " & Mid(cSQL_FINAL, nPOS)
510:         Else
            ' Se for apenas "TABELA WHERE...", tenta montar
512:             cSQL_FINAL = "DELETE FROM " & cSQL_FINAL
513:         End If
514:     End If
    
    ' Executa através do comando padrão que já trata Dialeto e Conexão
517:     ApagaSQLite = SQLiteComando(cCON, cSQL_FINAL)
    
519:     Exit Function
Erro:
521:     ApagaSQLite = False
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

537:     On Error GoTo Erro
538:     SQLMoveRegSQLite = False

540:     loConnOri.OpenDB cCONORI
541:     VBSQLiteSetValues loConnOri
542:     loConnDes.OpenDB cCONDES
543:     VBSQLiteSetValues loConnDes

    ' 1. Coleta dados da Origem (Passo Matriz)
546:     Set loCurOri = loConnOri.CreateCursor(cSQLORI)
547:     If Not loCurOri.EOF Then
548:         nCAMPOS = UBound(aCAMORI)
        ReDim aVALORI(nCAMPOS)
550:         For x = 0 To nCAMPOS
551:             aOPE = SepSqlOpe(aCAMORI(x))
552:             If aOPE(0) = "" Then
553:                 aVALORI(x) = loCurOri.Value(aCAMORI(x))
554:             Else
555:                 aVALORI(x) = MathOper(loCurOri.Value(aOPE(1)), loCurOri.Value(aOPE(2)), aOPE(0))
556:             End If
557:         Next x

        ' 2. Metadados do Destino
560:         cTABELA = ExtraiTabela(cSQLDES)
561:         cWHERE = ExtraiWhere(cSQLDES)

        ' 3. Lógica de Gravação via Execute
564:         Set loCurDes = loConnDes.CreateCursor("SELECT count(*) FROM " & cTABELA & " " & cWHERE)
565:         nRegs = loCurDes.Value(0)
        
567:         If nRegs > 0 Then
            ' --- UPDATE ---
569:             cSET = ""
570:             For x = 0 To UBound(aCAMDES)
571:                 cSET = cSET & IIf(cSET = "", "", ", ") & aCAMDES(x) & " = " & PrepararValorSQL(aVALORI(x))
572:             Next x
            
574:             If IsArray(aOUTDES) Then
575:                 For x = 0 To UBound(aOUTDES)
576:                     cSET = cSET & ", " & aOUTDES(x) & " = " & PrepararValorSQL(aOUTORI(x))
577:                 Next x
578:             End If
            
580:             loConnDes.Execute "UPDATE " & cTABELA & " SET " & cSET & " " & cWHERE
581:         Else
            ' --- INSERT ---
583:             cINS_C = "": cINS_V = ""
584:             For x = 0 To UBound(aCAMDES)
585:                 cINS_C = cINS_C & IIf(cINS_C = "", "", ", ") & aCAMDES(x)
586:                 cINS_V = cINS_V & IIf(cINS_V = "", "", ", ") & PrepararValorSQL(aVALORI(x))
587:             Next x
            
589:             If IsArray(aOUTDES) Then
590:                 For x = 0 To UBound(aOUTDES)
591:                     cINS_C = cINS_C & ", " & aOUTDES(x)
592:                     cINS_V = cINS_V & ", " & PrepararValorSQL(aOUTORI(x))
593:                 Next x
594:             End If
            
596:             loConnDes.Execute "INSERT INTO " & cTABELA & " (" & cINS_C & ") VALUES (" & cINS_V & ")"
597:         End If

        ' 4. Captura de IDs para eRETU01
600:         If IsArray(aIDDES) Then
601:             Set loCurDes = loConnDes.CreateCursor(cSQLDES)
            ReDim aRETUID(UBound(aIDDES))
603:             For x = 0 To UBound(aIDDES)
604:                 aRETUID(x) = loCurDes.Value(aIDDES(x))
605:             Next x
606:             eRETU01 = aRETUID
607:         End If
608:         SQLMoveRegSQLite = True
609:     End If

611:     loConnOri.CloseDB: loConnDes.CloseDB
612:     Exit Function
Erro:
614:     SQLMoveRegSQLite = False
End Function

Public Function VBSQLiteSetValues(ByRef oCONN As SQLiteConnection) As Boolean
618:     On Error GoTo Erro
    
    ' A VBSQLite permite aplicar configurações diretamente via Execute
    ' Algumas destas configurações são específicas para esta biblioteca
    
    ' 1. Aumenta a performance de escrita
624:     oCONN.Execute "PRAGMA journal_mode = WAL;"
    
    ' 2. Ajusta o tamanho da cache (em páginas)
627:     oCONN.Execute "PRAGMA cache_size = 2000;"
    
    ' 3. Melhora a segurança de escrita com performance otimizada
630:     oCONN.Execute "PRAGMA synchronous = NORMAL;"
    
    ' 4. Define o modo de armazenamento temporário
633:     oCONN.Execute "PRAGMA temp_store = MEMORY;"
    
635:     VBSQLiteSetValues = True
636:     Exit Function
    
Erro:
639:     VBSQLiteSetValues = False
End Function

Public Function TratarValorParaSQL(ByVal vValor As Variant, ByVal cTIPO As String, ByVal cDIALETO As String) As String
    ' cTipo: "C" para Caractere, "N" para Número, "D" para Data
    ' cDialeto: "SQLITE", "PGSQL", "VFP"
    
646:     If IsNull(vValor) Or vValor = "" Then
647:         TratarValorParaSQL = "NULL"
648:         Exit Function
649:     End If
    
651:     Select Case UCase(cDIALETO)
        Case "SQLITE"
653:             If UCase(cTIPO) = "D" Then
                ' SQLite armazena datas como string YYYY-MM-DD
655:                 TratarValorParaSQL = "'" & Format(vValor, "yyyy-mm-dd") & "'"
656:             ElseIf UCase(cTIPO) = "N" Then
657:                 TratarValorParaSQL = Replace(vValor, ",", ".") ' Garante ponto decimal
658:             Else
659:                 TratarValorParaSQL = "'" & Replace(vValor, "'", "''") & "'" ' Escapa aspas simples
660:             End If
            
        Case "PGSQL", "VFP"
            ' Aqui você mantém a lógica que já funcionava para os outros bancos
664:             TratarValorParaSQL = "'" & vValor & "'"
665:     End Select
End Function
