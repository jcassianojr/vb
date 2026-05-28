Attribute VB_Name = "SqlFuncoesAdo"
Option Explicit

Public Function SQLiteSetValuesADO(ByRef oCON As Object) As Boolean
    Dim oCOMANDO As Object ' ADODB.Command
    
6:     On Error GoTo ErroSQLite
7:     SQLiteSetValuesADO = False
    
    ' 1. Verifica se a conexão é válida
10:     If oCON Is Nothing Then Exit Function
    
    ' 2. Cria objeto de comando ADO
13:     Set oCOMANDO = CreateObject("ADODB.Command")
14:     Set oCOMANDO.ActiveConnection = oCON
15:     oCOMANDO.CommandType = 1 ' adCmdText
    
    ' 3. Aplica as configurações de alta performance do SQLite
    ' O SQLite permite executar múltiplos comandos separados por ponto e vírgula
    oCOMANDO.CommandText = "PRAGMA temp_store = MEMORY;" & _
                           "PRAGMA cache_size = 2000;" & _
                           "PRAGMA journal_mode = WAL;" & _
                           "PRAGMA synchronous = NORMAL;" & _
                           "PRAGMA auto_vacuum = INCREMENTAL;"
    
25:     oCOMANDO.Execute
    
27:     SQLiteSetValuesADO = True
28:     Set oCOMANDO = Nothing
29:     Exit Function

ErroSQLite:
32:     Set oCOMANDO = Nothing
33:     SQLiteSetValuesADO = False
End Function

' ==============================================================================
' FUNÇÃO: ConfigurarConexaoADO
' Objetivo: Orquestrador central que aplica os SETs de performance e ambiente
'           baseado na string de conexão fornecida, usando tecnologia ADO.
' ==============================================================================
Public Function ConfigurarConexaoADO(ByRef oCON As Object, ByVal cStringConexao As String) As Boolean
    Dim cBUSCA As String
43:     cBUSCA = UCase(cStringConexao)
    
45:     On Error GoTo ErroConfig
    
    ' --- 1. VFP / FoxPro (Configura SET DELETED ON / NULL OFF) ---
48:     If InStr(cBUSCA, "VFPOLEDB") > 0 Then
49:         VFPSetValues oCON
    
    ' --- 2. PostgreSQL (Configura search_path / encoding) ---
52:     ElseIf InStr(cBUSCA, "PGSQL") > 0 Or InStr(cBUSCA, "POSTGRESQL") > 0 Then
53:         pgSetValues oCON
        
    ' --- 3. SQLite (Configura WAL, Cache e Performance via ADO) ---
56:     ElseIf InStr(cBUSCA, ".DB") > 0 Or InStr(cBUSCA, "SQLITE") > 0 Then
57:         SQLiteSetValuesADO oCON
        
59:     End If
    
61:     ConfigurarConexaoADO = True
62:     Exit Function

ErroConfig:
65:     ConfigurarConexaoADO = False
End Function
Public Function pgSetValues(ByRef oCON As Object) As Boolean
    Dim cCOM As String
    Dim oCOMANDO As Object ' Usando Object para manter o padrão flexível do seu projeto
    
71:     On Error GoTo ErroPostgres
72:     pgSetValues = False
    
    ' Se a conexão não estiver ativa, sai para evitar travamentos
75:     If oCON Is Nothing Then Exit Function
    
    ' Cria o objeto de comando ADO dinamicamente
78:     Set oCOMANDO = CreateObject("ADODB.Command")
79:     oCOMANDO.CommandType = 1 ' adCmdText = 1
80:     Set oCOMANDO.ActiveConnection = oCON

    ' INJEÇÃO DO SCHEMA E ENCODING ANSI (Porto Seguro)
    ' Configura o search_path para o seu esquema e garante tráfego em WIN1252
84:     cCOM = "SET search_path TO myschema, public; SET client_encoding TO 'WIN1252';"
85:     oCOMANDO.CommandText = cCOM
86:     oCOMANDO.Execute

88:     pgSetValues = True
89:     Set oCOMANDO = Nothing
90:     Exit Function

ErroPostgres:
    ' Se falhar (ex: banco temporariamente fora ou string malformada), limpa a memória
94:     Set oCOMANDO = Nothing
95:     pgSetValues = False
End Function
Public Function VFPSetValues(ByRef oCON As Object) As Boolean  ''como objeto nao gerou erro como byref ADODB.Connection as object
'exemplos copia e cola
'VFPSetValues DB
'VFPSetValues ODB
'VFPSetValues oCONN

  Dim cCOM As String
  Dim oCOMANDO As ADODB.Command

  'inicializa as variaveis
107:   VFPSetValues = False
108:   cCOM = ""
109:   Set oCOMANDO = New ADODB.Command

  'atribui o comando a conecao passada como referencia
112:   oCOMANDO.CommandType = adCmdText
113:   oCOMANDO.ActiveConnection = oCON

  'set deleted on pega o registro correto evitando gravar em registro apagados
116:   cCOM = "set deleted on"
117:   oCOMANDO.CommandText = cCOM
118:   oCOMANDO.Execute

  'set null permite gravar parcialmente (permite gravar um registro novo sem passar gravar todos os campos)
121:   cCOM = "set null off"
122:   oCOMANDO.CommandText = cCOM
123:   oCOMANDO.Execute
124:   VFPSetValues = True

End Function
Public Function AdoComandodbf(ByVal cARQ As String, ByVal cTable As String, ByVal CCOMANDO As String) As Boolean
  Dim cCOM As String
  Dim oCON As ADODB.Connection
  Dim oCOMANDO As ADODB.Command
131:   On Error GoTo TrataErro

  'ZAP PACK
  'delete from mail where numero=1
  'insert into mail (numero,erro,data,hora,de,destino,assunto,texto) values (recno(),"erro",ctod(date()),"12:00","de","destino","assunto","texto")

  'Passado "" pega o da string
138:   If Len(cTable) = 0 Then
139:     cTable = NomeTableSql(CCOMANDO)
140:   End If
  'atribui os valores iniciais
142:   AdoComandodbf = False
143:   cCOM = ""
144:   Set oCON = New ADODB.Connection
145:   Set oCOMANDO = New ADODB.Command

  'abre a conecao
148:   oCON.ConnectionTimeout = 120
149:   oCON.Open cARQ

  'atribui o comando a conecao
152:   oCOMANDO.CommandType = adCmdText
153:   oCOMANDO.ActiveConnection = oCON

  'para pack e zap considera todos necessario setar aqui pois na string de conecao esta marcado delete on
156:   If CCOMANDO = "ZAP" Or CCOMANDO = "PACK" Then
157:     cCOM = "set deleted off"
158:     oCOMANDO.CommandText = cCOM
159:     oCOMANDO.Execute
160:   Else
161:     cCOM = "set deleted on"
162:     oCOMANDO.CommandText = cCOM
163:     oCOMANDO.Execute
164:   End If
  'set null permite gravar parcialmente (permite gravar um registro novo sem passar gravar todos os campos)
166:   cCOM = "set null off"
167:   oCOMANDO.CommandText = cCOM
168:   oCOMANDO.Execute

  'abre a tabela
171:   If Len(cTable) > 0 Then
172:     cCOM = "use " & cTable
173:     oCOMANDO.CommandText = cCOM
174:     oCOMANDO.ActiveConnection = oCON
175:     oCOMANDO.Execute
176:   End If

  'executa o comando
179:   If CCOMANDO = "ZAP" Then
180:     cCOM = "DELETE FROM " & cTable & " WHERE 1=1"
181:     oCOMANDO.CommandText = cCOM
182:     oCOMANDO.Execute
183:   End If
184:   If CCOMANDO = "ZAP" Or CCOMANDO = "PACK" Then
185:     cCOM = "PACK"
186:     oCOMANDO.CommandText = cCOM
187:     oCOMANDO.Execute
188:   Else
189:     cCOM = CCOMANDO
190:     oCOMANDO.CommandText = cCOM
191:     oCOMANDO.Execute
192:   End If

  'fecha a coneccao
195:   oCON.Close
196:   AdoComandodbf = True

198:   Exit Function
TrataErro:
200:   On Error Resume Next
201:   Select Case Err.Number
  Case Else
203:     SayErro "ADO Comando DBF:" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cTable & Chr(13) & Chr(10) & cCOM & Chr(13) & Chr(10)
204:     Exit Function
205:   End Select
End Function
Public Function ADOComando(ByVal cARQ As String, ByVal cSQL As String) As Boolean
  Dim aRETU As Variant
  Dim oDB As ADODB.Connection
  Dim oCOM As ADODB.Command
  Dim cCONN As String

213:   On Error GoTo TrataErro
214:   ADOComando = False
215:   cARQ = GeraConn(cARQ)
216:   aRETU = TipoConn(cARQ)  ''Gera string conneccao
  
  '
  ' nao executa se nao for ado
  '
221:   If aRETU(0) <> "ADO" Then
222:      Exit Function
223:   End If
    
  '
  ' se for foxpro necessita rotina adocomandodbf
  '
228:     If InStr(UCase(cARQ), "VFPOLEDB") Then
229:       ADOComando = AdoComandodbf(cARQ, "", cSQL)
230:       Exit Function
231:     End If
    
233:     If InStr(cSQL, "CURRENTDATETIME") > 0 Then
234:        Select Case aRETU(2)
           Case "SQLITE"
236:                 cSQL = Replace(cSQL, "CURRENTDATETIME", " current_timestamp ")
           Case "MDB"
238:                 cSQL = Replace(cSQL, "CURRENTDATETIME", " now ")
           Case Else
240:                 cSQL = Replace(cSQL, "CURRENTDATETIME", "'" & Format(Date, "dd/mm/yyyy") & "'")
241:        End Select
       
243:     End If
    
    
246:     cCONN = aRETU(1)
247:     Set oDB = New ADODB.Connection
248:     Set oCOM = New ADODB.Command
    
250:     oDB.ConnectionTimeout = 120
251:     oDB.Open cCONN

253:     oCOM.ActiveConnection = oDB
254:     oCOM.CommandText = cSQL
255:     oCOM.Execute

257:     ADOComando = True
258:     oDB.Close
259:     Set oCOM = Nothing
260:     Set oDB = Nothing
  
262:   Exit Function
TrataErro:
264:   Select Case Err.Number
  Case Else
266:     SayErro "SQL ADO Comando:" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cSQL
267:     Exit Function
268:   End Select
End Function
Public Function APAGASQLADO(ByVal cARQ As String, ByVal cSQL As String) As Boolean
  Dim nPOS As Integer
  Dim cNOME As String
  'select * from tabela where campo=valorcampo ....
  'delete * from tabela where campo=valorcampo ....
  'delete from tabela where campo=valorcampo ....
  'from tabela where campo=valorcampo ....
  'todos os csql viram delete from tabela where campo=valorcampo ....
  'requer from e where no csql


281:   APAGASQLADO = False
  'Muda para maiscula para o instr usar em maiscula
283:   cNOME = UCase(cSQL)
284:   cNOME = Replace(cNOME, Chr(13), " ")
285:   cNOME = Replace(cNOME, Chr(10), " ")

287:   If InStr(cNOME, "WHERE") = 0 Then
288:     Alert ("where nao preenchido apagasql")
289:     Exit Function
290:   End If
291:   If InStr(cNOME, "FROM") = 0 Then
292:     Alert ("from nao preenchido apagasql")
293:     Exit Function
294:   End If

296:   nPOS = InStr(cNOME, "FROM")
  'efetua substituicao se estiver com select * from  ou outros casos --> delete from usando mid
  'todos os csql viram delete from tabela where campo=valorcampo .... opcoes de passagem csql acima
299:   If nPOS > 0 Then
300:     cSQL = "DELETE FROM " & Mid(cSQL, nPOS + 5)
301:     APAGASQLADO = ADOComando(cARQ, cSQL)
302:   Else
303:     If InStr(cNOME, "DELETE") > 0 Then  'so executa se tiver delete
304:       APAGASQLADO = ADOComando(cARQ, cSQL)
305:     End If
306:   End If
End Function

Public Function SomaSQLAdo(ByVal cARQ As String, ByVal cSQL As String, ByVal aCAM As Variant) As Variant
'esta faz soma com loop
'use pegsumsql se o loop nao for necessario
  Dim oDB As ADODB.Connection
  Dim oRS As ADODB.Recordset
  Dim lOPEN As Boolean
  Dim lRSOP As Boolean
  Dim x As Integer
  Dim nCAMPOS As Integer
  Dim aRETU As Variant
  Dim aARQ As Variant
  Dim aOPE As Variant
  Dim eVAL As Variant
  Dim cERRO As String
  Dim carqcon As String

325:   On Error GoTo errhandler

327:   lOPEN = False
328:   lRSOP = False
329:   nCAMPOS = UBound(aCAM) + 1
  ReDim aRETU(nCAMPOS)
331:   For x = 0 To nCAMPOS - 1
332:     aRETU(x) = 0
333:   Next x

335:   If Not FileConnExist(SomaExt(cARQ), True) Then
336:     SomaSQLAdo = aRETU
337:     Exit Function
338:   End If

340:   aARQ = TipoConn(cARQ, , , False)
341:   carqcon = aARQ(1) 'GeracArq(cARQ, , False)

343:   Set oDB = New ADODB.Connection
344:   oDB.CursorLocation = adUseClient
345:   oDB.ConnectionTimeout = 120
346:   oDB.Open carqcon
347:   lOPEN = True


350:   ConfigurarConexaoADO oDB, cARQ
  'na string de conecao delete =yes
  'If InStr(UCase(cARQ), "VFPOLEDB") > 0 Then  ''delete on para foxpro nao usar registro deletados
  '  VFPSetValues oDB
  'End If
  'If InStr(UCase(cARQ), "PGSQL") > 0 Or InStr(UCase(cARQ), "POSTGRESQL") > 0 Then
  '  pgSetValues oDB
  'End If

359:   Set oRS = New ADODB.Recordset
360:   If aARQ(2) = "SQLITE" Then
361:      oRS.Open cSQL, oDB, adOpenStatic, adLockReadOnly
362:   Else
363:      oRS.Open cSQL, oDB, adOpenForwardOnly, adLockReadOnly
364:   End If
365:   lRSOP = True



369:   If Not oRS.EOF Then
370:     oRS.MoveFirst
371:     While Not oRS.EOF
372:       For x = 0 To nCAMPOS - 1
373:         aOPE = SepSqlOpe(aCAM(x))
374:         If aOPE(0) = "" Or aOPE(1) = "" Or aOPE(2) = "" Then  ''Sem Operacoes
375:           eVAL = oRS(aCAM(x))
376:         Else
377:           eVAL = MathOper(oRS(aOPE(1)), oRS(aOPE(2)), aOPE(0))
378:         End If
379:         aRETU(x) = aRETU(x) + eVAL
380:       Next x
381:       oRS.MoveNext
382:     Wend
383:   End If
384:   oRS.Close
385:   oDB.Close
386:   Set oRS = Nothing
387:   Set oDB = Nothing
388:   SomaSQLAdo = aRETU
389:   Exit Function

errhandler:
392:   cERRO = "Soma SQL ADO :" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cSQL & Chr(13) & Chr(10)
393:   If lRSOP Then
394:     cERRO = cERRO & ADORsStatus(oRS.Status) & Chr(13) & Chr(10)
395:   End If
396:   Select Case Err.Number
  Case Else
398:     If lOPEN Then
399:       ADOErro oRS.ActiveConnection.Errors, cERRO
400:     Else
401:       ADOErro oDB.Errors, cERRO
402:     End If
403:     Exit Function
404:   End Select
End Function

Public Function PegSQLDeliAdo(ByVal cARQ As String, ByVal cSQL As String, _
                              ByVal aCAM As Variant, Optional ByVal cDELI As String = ",", Optional ByVal aPAD As Variant = "", Optional ByVal aFOR As Variant = "") As Variant

  Dim oDB As ADODB.Connection
  Dim oRS As ADODB.Recordset
  Dim x, nCAMPOS As Integer
  Dim aRETU As Variant
  Dim aARQ As Variant
  Dim aOPE As Variant
  Dim eVAL As Variant
  Dim lOPEN As Boolean
  Dim lRSOP As Boolean
  Dim cERRO As String
  Dim carqcon As String

422:   On Error GoTo errhandler

424:   lOPEN = False
425:   lRSOP = False
426:   nCAMPOS = UBound(aCAM) + 1
  ReDim aRETU(nCAMPOS)
428:   For x = 0 To nCAMPOS - 1
429:     aRETU(x) = ""
430:   Next x

432:   If Not FileConnExist(SomaExt(cARQ), True) Then
433:     PegSQLDeliAdo = aRETU
434:     Exit Function
435:   End If

437:   aARQ = TipoConn(cARQ, , , False)
  'ARQ = GeracArq(cARQ, , False)
439:   carqcon = aARQ(1)

441:   Set oDB = New ADODB.Connection
442:   oDB.CursorLocation = adUseClient
443:   oDB.ConnectionTimeout = 120
444:   oDB.Open carqcon

  'na string de conecao delete =yes
447:   ConfigurarConexaoADO oDB, cARQ
  'If InStr(cARQ, "VFPOLEDB") > 0 Then  ''delete on para foxpro nao usar registro deletados
  '  VFPSetValues oDB
  'End If
'If InStr(UCase(cARQ), "PGSQL") > 0 Or InStr(UCase(cARQ), "POSTGRESQL") > 0 Then
'        pgSetValues oDB
'  End If

455:   lOPEN = True

457:   Set oRS = New ADODB.Recordset

459:   oRS.Open cSQL, oDB, adOpenForwardOnly, adLockReadOnly

461:   lRSOP = True
462:   If Not oRS.EOF Then
463:     oRS.MoveFirst
464:     While Not oRS.EOF
465:       For x = 0 To nCAMPOS - 1
466:         aOPE = SepSqlOpe(aCAM(x))
467:         If aOPE(0) = "" Or aOPE(1) = "" Or aOPE(2) = "" Then  ''Sem Operacoes
468:           eVAL = oRS(aCAM(x))
469:         Else
470:           eVAL = MathOper(oRS(aOPE(1)), oRS(aOPE(2)), aOPE(0))
471:         End If
        
        
474:             If IsNull(eVAL) Then
475:                 If IsArray(aPAD) Then eVAL = aPAD(x) Else eVAL = aPAD
                
477:             End If
            
            ' Tratamento de Formato com aFOR (C=Texto, N=Numero, D=Data)
480:             If IsArray(aFOR) And IsArray(aPAD) Then
481:                 eVAL = FVar(eVAL, aFOR(x), aPAD(x))
482:             End If
            
484:              If IsArray(aFOR) And Not IsArray(aPAD) Then
485:                 eVAL = FVar(eVAL, aFOR(x))
486:             End If
            
        
489:         aRETU(x) = aRETU(x) & FixStr(eVAL)
490:       Next x
491:       oRS.MoveNext
492:       If Not oRS.EOF Then
493:         For x = 0 To nCAMPOS - 1
494:           aRETU(x) = aRETU(x) & cDELI
495:         Next x
496:       End If

498:     Wend
499:   End If
500:   oRS.Close
501:   oDB.Close
502:   Set oRS = Nothing
503:   Set oDB = Nothing
504:   PegSQLDeliAdo = aRETU
505:   Exit Function

errhandler:
508:   cERRO = "Peg SQL DELI ADO" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cSQL & Chr(13) & Chr(10)
509:   If lRSOP Then
510:     cERRO = cERRO & ADORsStatus(oRS.Status) & Chr(13) & Chr(10)
511:   End If
512:   Select Case Err.Number
  Case Else
514:     If lOPEN Then
515:       ADOErro oRS.ActiveConnection.Errors, cERRO
516:     Else
517:       ADOErro oDB.Errors, cERRO
518:     End If
519:     Exit Function
520:   End Select
End Function

Public Function ApagaSQLpAdo(ByVal cARQ As String, ByVal cSQL As String, Optional ByVal cTEXTO As String = "Confirme Exclusão") As Boolean
524:   ApagaSQLpAdo = False
525:   If MDG(cTEXTO, "Exclusão Registro") Then
526:     ApagaSQLpAdo = APAGASQLADO(cARQ, cSQL)
527:   End If
End Function

Public Function GrvSQLado(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, ByVal aCAM As Variant, _
                          ByVal aVAL As Variant, ByVal aFOR As Variant, Optional ByVal nSTARITEM = 0) As Boolean
  Dim oDB As ADODB.Connection
  Dim oRS As ADODB.Recordset
  Dim eVAL, aOPE As Variant
  Dim x As Long
  Dim nLENVAR As Long
  Dim nLENCAM As Long
  Dim lGRAVA As Boolean
  Dim cTABELA As String
  Dim eVAZIO As Variant
  Dim cTIPO As String
  Dim aRETU As Variant
  Dim aARQ As Variant
  Dim lOPEN As Boolean
  Dim lRSOP As Boolean
  Dim cERRO As String
  Dim oFIELD As ADODB.Field
  Dim carqcon As String

550:   On Error GoTo errhandler
551:   lOPEN = False
552:   lRSOP = False
553:   GrvSQLado = False
554:   cTABELA = NomeTableSql(cSQL)

556:   aARQ = TipoConn(cARQ)
557:   carqcon = aARQ(1)

559:   If aARQ(2) = "PGSQL" Then
560:      cSQL = SQLPGSQLDOUBLEQUOTES(cSQL)
561:   End If
  
563:   Set oDB = New ADODB.Connection
564:   oDB.CursorLocation = adUseClient
565:   oDB.ConnectionTimeout = 120
566:   oDB.Open carqcon
 
568:   ConfigurarConexaoADO oDB, cARQ
  'If InStr(cARQ, "VFPOLEDB") Then
  '  VFPSetValues oDB
  'End If
  'If InStr(UCase(cARQ), "PGSQL") > 0 Or InStr(UCase(cARQ), "POSTGRESQL") > 0 Then
  '      pgSetValues oDB
  'End If

576:   lOPEN = True
577:   Set oRS = New ADODB.Recordset
578:   If aARQ(2) = "SQLITE" Then
579:      oRS.Open cSQL, oDB, adOpenStatic, adLockOptimistic
580:   Else
581:      oRS.Open cSQL, oDB, adOpenKeyset, adLockOptimistic
582:   End If
583:   lRSOP = True

585:   If oRS.EOF Then
586:     oRS.Close
587:     oDB.Close
588:     Set oRS = Nothing
589:     Set oDB = Nothing
590:     Exit Function
591:   End If

593:   While Not oRS.EOF
594:     For x = nSTARITEM To nITEM - 1     ''Matriz Vb comeca 0 x = 0 To nITEM - 1
                                       'permite nao gravar campos iniciais evitando grava campo chave primaria
596:       Set oFIELD = oRS(aCAM(x))
597:       cTIPO = TipoDado2(oFIELD.Type)
598:       lGRAVA = True
      ''Evita Gravar Campos nullos correcao data nula abaixo
600:       If IsNull(aVAL(x)) Then
601:         lGRAVA = False
602:         eVAL = aVAL(x)
603:       Else
604:         aOPE = SepSqlOpe(aCAM(x))
605:         If aOPE(0) = "" Or aOPE(1) = "" Or aOPE(2) = "" Then  ''Operacao com Campos
606:           eVAL = aVAL(x)               ''Grava o Valor Default
607:         Else
608:           eVAL = MathOper(oRS(aOPE(1)), oRS(aOPE(2)), aOPE(0))
609:         End If
610:       End If
      ''Evitar Gravar String Vazias Campos Textos
612:       If VarType(eVAL) = vbString And cTIPO = "S" Then
613:         If eVAL = "" Then
          ''EVAL = " " ''nao apagava o campo ja preenchido antes
          ''lGRAVA = False
616:         Else
617:           nLENCAM = oFIELD.DefinedSize
618:           If nLENCAM > 0 Then          ''Evita Retorno Campos -1
619:             nLENVAR = Len(eVAL)
620:             If nLENVAR > nLENCAM Then
621:               eVAL = Mid(eVAL, 1, nLENCAM)
622:             End If
623:           End If
624:         End If
625:         If oFIELD.Type = adLongVarWChar And eVAL = "" Then  ''Nao saida '' ou null
626:           eVAL = "_"
627:         End If
628:       End If
      
           
      ''Evitar Gravar String Vazias Campos DAta
632:       If cTIPO = "D" Then
633:         If DataBranco(eVAL) Then
634:           If aARQ(2) = "MDB" Or aARQ(2) = "MYSQL" Or aARQ(2) = "MARIADB" Then
635:             eVAL = NullDate(aARQ(2))
636:             lGRAVA = True
637:             aFOR(x) = ""
638:           Else
639:             lGRAVA = False
640:           End If
641:         End If
642:       End If
     
644:       If cTIPO = "B" Then
645:         aFOR(x) = ""
646:       End If
      
      ''Efetua a Gravaçao
649:       If lGRAVA Then
650:          If aARQ(2) = "SQLITE" And Mid(aFOR(x), 1, 1) = "D" Then '(aFOR(x) = "DH" Or aFOR(x) = "D" Or aFOR(x) = "DC") Then
651:              oRS(aCAM(x)) = CStr(Format(eVAL, "yyyy-mm-dd hh:ss:mm"))
'             oRS(aCAM(x)) = eVAL 'FVar(eVAL, aFOR(x), eVAZIO)
'            If aFOR(x) = "DH" Then
'              oRS(aCAM(x)) = CStr(Now)
'            End If
'            If aFOR(x) = "D" Or aFOR(x) = "DC" Then
'              oRS(aCAM(x)) = CStr(Format(eVAL, "yyyy-mm-dd hh:ss:mm"))
'            End If
'            If aFOR(x) = "T" Then
'              oRS(aCAM(x)) = CStr(Format(eVAL, "yyyy-mm-dd hh:ss:mm"))
'            End If

663:          Else
664:             oRS(aCAM(x)) = FVar(eVAL, aFOR(x), eVAZIO)
665:          End If
666:       End If
667:     Next x
668:     Select Case aARQ(2) ' alguns nao aceitam update
           Case Else
670:                oRS.Update
671:     End Select
    
673:     oRS.MoveNext
674:   Wend
675:   oRS.Close
676:   oDB.Close
677:   Set oRS = Nothing
678:   Set oDB = Nothing
679:   GrvSQLado = True
680:   Exit Function

errhandler:
683:   cERRO = "Grava SQL ADO" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cSQL & Chr(13) & Chr(10)
684:   If lRSOP Then
685:     cERRO = cERRO & ADORsStatus(oRS.Status) & Chr(13) & Chr(10)
686:   End If
687:   Select Case Err.Number
  Case 3265, 3421, 94
689:     cERRO = cERRO & "Campo:" & aCAM(x) & Chr(13) & Chr(10)
690:     ADOErro oRS.ActiveConnection.Errors, cERRO
691:     Resume Next
  Case -2147217887
693:     GrvSQLado = True
694:     Exit Function
  Case -2147217864
696:     GrvSQLado = True
697:     Exit Function
  Case Else
699:     If lOPEN Then
700:       ADOErro oRS.ActiveConnection.Errors, cERRO
701:     Else
702:       ADOErro oDB.Errors, cERRO
703:     End If
704:   End Select

End Function

Public Function IncluiSQLAdo(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, _
                             Optional ByVal aCAM As Variant = 0, Optional ByVal aVAL As Variant = 0, _
                             Optional ByVal lCHECK As Boolean = False, _
                             Optional ByVal lMES As Boolean = True, _
                             Optional ByVal aIDDES As Variant = 0) As Boolean
  Dim oDB As ADODB.Connection
  Dim oRS As ADODB.Recordset
  Dim x As Long
  Dim nCAMPOS As Long
  Dim lTEM As Boolean
  Dim lGRAVA As Boolean
  Dim aRETUID As Variant
  Dim cARQ1 As String
  Dim aRETU As Variant
  Dim lOPEN As Boolean
  Dim lRSOP As Boolean
  Dim cERRO As String

726:   IncluiSQLAdo = False
727:   On Error GoTo errhandler

729:   lTEM = False
730:   lRETU = False
731:   aRETU = TipoConn(cARQ)

733:   cARQ1 = aRETU(1)
734:   lOPEN = False
735:   lRSOP = False
  
737:    If aRETU(2) = "PGSQL" Then
738:      cSQL = SQLPGSQLDOUBLEQUOTES(cSQL)
739:   End If

741:   Set oDB = New ADODB.Connection
742:   oDB.CursorLocation = adUseClient
743:   oDB.ConnectionTimeout = 120
744:   oDB.Open cARQ1

746:   lOPEN = True

  'Set null off permitido deixar campos em branco set deleted on ja esta na string de conecao
749:   ConfigurarConexaoADO oDB, cARQ1
  'If InStr(cARQ1, "VFPOLEDB") Then
  '  VFPSetValues oDB
  'End If
  'If InStr(UCase(cARQ1), "PGSQL") > 0 Or InStr(UCase(cARQ1), "POSTGRESQL") > 0 Then
  '      pgSetValues oDB
  'End If

757:   Set oRS = New ADODB.Recordset
758:   If aRETU(1) = "SQLITE" Then
759:   End If
760:   oRS.Open cSQL, oDB, adOpenKeyset, adLockOptimistic

762:   lRSOP = True

764:   If lCHECK Then
765:     If Not oRS.EOF Then
766:       lTEM = True
767:     End If
768:   End If

770:   If Not lTEM Then

772:     oRS.AddNew
773:     For x = 0 To nITEM - 1                   ''Matriz Vb comeca 0
774:       lGRAVA = True
      ''Evita Gravar Campos nullos
776:       If IsNull(aVAL(x)) Then
777:         lGRAVA = False
778:       End If
779:       If lGRAVA Then
780:         oRS(aCAM(x)) = aVAL(x)
781:       End If
782:     Next

    ' PEGAR O ID
785:     Select Case aRETU(2) 'opcao conn que nao aceitam update
      Case Else
787:          oRS.Update
788:     End Select
789:     If Not IsNumeric(aIDDES) Then            ''Se for numerico nao e matriz
790:       nCAMPOS = UBound(aIDDES)
      ReDim aRETUID(nCAMPOS + 1)
792:       For x = 0 To nCAMPOS
793:         aRETUID(x) = oRS(aIDDES(x))
794:       Next x
795:       eRETU01 = aRETUID
796:     End If
797:     Select Case aRETU(2) 'alguns nao aceitam update
        Case Else
799:           oRS.Update
800:     End Select

802:     lRETU = True
803:   End If
804:   oRS.Close
805:   oDB.Close
806:   Set oRS = Nothing
807:   Set oDB = Nothing

809:   If lCHECK And lTEM Then
810:     If lMES Then
811:       Alert ("Item ja Cadastrado Com Esta Chave")
812:     End If
813:   End If
814:   IncluiSQLAdo = lRETU
815:   Exit Function
errhandler:
817:   cERRO = "Inclui SQL ADO" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cSQL & Chr(13) & Chr(10)
818:   If lRSOP Then
819:     cERRO = cERRO & ADORsStatus(oRS.Status) & Chr(13) & Chr(10)
820:   End If
821:   Select Case Err.Number
  Case 3265                                    ''Campo Diferente
823:     cERRO = cERRO & "Campo Diferente:" & aCAM(x) & Chr(13) & Chr(10)
824:     ADOErro oRS.ActiveConnection.Errors, cERRO
825:     Resume Next
  Case 3315, 94                                'Campo null
827:     Resume Next
  Case 424                                     ''Campo Inexistente
829:     cERRO = cERRO & "Campo Inexistente:" & aCAM(x) & Chr(13) & Chr(10)
830:     ADOErro oRS.ActiveConnection.Errors, cERRO
831:     Resume Next
  Case 3163                                    ''Valor Maior que o Campo
833:     cERRO = cERRO & "Valor Maior que o Campo:" & aCAM(x) & Chr(13) & Chr(10)
834:     ADOErro oRS.ActiveConnection.Errors, cERRO
835:     Resume Next
  Case 3219                                    ''close jetfox
837:     Resume Next
  Case -2147217887                             ''Update jetfox
839:     Resume Next
  Case Else
841:     If lOPEN Then
842:       ADOErro oRS.ActiveConnection.Errors, cERRO
843:     Else
844:       ADOErro oDB.Errors, cERRO
845:     End If
846:     Exit Function
    ''        PegSQL = aPAD
848:   End Select
End Function

Public Function PegSQLAdo(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, ByVal aCAM As Variant, ByVal aFOR As Variant, ByVal aPAD As Variant) As Variant
  Dim oDB As ADODB.Connection
  Dim oRS As ADODB.Recordset
  Dim x As Long
  Dim aRETU As Variant
  Dim lOPEN As Boolean
  Dim lRSOP As Boolean
  Dim cERRO As String
  Dim aOPE As Variant
  Dim eVAL As Variant
  Dim aCON As Variant

863:   On Error GoTo errhandler
  ReDim aRETU(nITEM)
865:   lOPEN = False
866:   lRSOP = False
  
868:   aCON = TipoConn(cARQ, , , False)
  
870:    If aCON(2) = "PGSQL" Then
871:      cSQL = SQLPGSQLDOUBLEQUOTES(cSQL)
872:   End If

874:   Set oDB = New ADODB.Connection
875:   oDB.CursorLocation = adUseClient
876:   oDB.ConnectionTimeout = 120
877:   oDB.Open cARQ

  'na string de conecao delete =yes
880:   ConfigurarConexaoADO oDB, cARQ
  'If InStr(cARQ, "VFPOLEDB") > 0 Then  ''delete on para foxpro nao usar registro deletados
  '  VFPSetValues oDB
  'End If
  'If InStr(UCase(cARQ), "PGSQL") > 0 Or InStr(UCase(cARQ), "POSTGRESQL") > 0 Then
  '      pgSetValues oDB
  'End If

888:   lOPEN = True
889:   Set oRS = New ADODB.Recordset
890:   oRS.Open cSQL, oDB, adOpenForwardOnly, adLockReadOnly

892:   lRSOP = True


895:   If Not oRS.EOF Then
896:     lRETU = True
897:     For x = 0 To nITEM - 1                   ''Matriz Vb comeca 0
898:       aOPE = SepSqlOpe(aCAM(x))
899:       If aOPE(0) = "" Or aOPE(1) = "" Or aOPE(2) = "" Then  ''Sem Operacoes
900:         eVAL = oRS(aCAM(x))
901:       Else
902:         eVAL = MathOper(oRS(aOPE(1)), oRS(aOPE(2)), aOPE(0))
903:       End If

905:       aRETU(x) = FVar(eVAL, aFOR(x), aPAD(x))
      
907:     Next
908:   Else
909:     lRETU = False
910:     aRETU = aPAD
911:   End If
912:   oRS.Close
913:   oDB.Close
914:   Set oRS = Nothing
915:   Set oDB = Nothing
916:   PegSQLAdo = aRETU
917:   Exit Function

errhandler:
920:   cERRO = "Peg SQL ADO" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cSQL & Chr(13) & Chr(10)
921:   If lRSOP Then
922:     cERRO = cERRO & ADORsStatus(oRS.Status)
923:   End If
924:   Select Case Err.Number
  Case 3265, 5
926:     aRETU(x) = aPAD(x)
927:     cERRO = cERRO & "Campo:" & aCAM(x) & Chr(13) & Chr(10)
928:     ADOErro oRS.ActiveConnection.Errors, cERRO
929:     Resume Next
  Case 94                                      ''Campo Esta com Null Pega Padrao
    ''SayErro (" PEGSQL Campo:" & aCAM(x))
932:     aRETU(x) = aPAD(x)
933:     Resume Next
  Case Else
935:     If lOPEN Then
936:       ADOErro oRS.ActiveConnection.Errors, cERRO
937:     Else
938:       ADOErro oDB.Errors, cERRO
939:     End If
940:     PegSQLAdo = aPAD
941:   End Select
End Function
Public Function PegCountSQLADO(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
944:   PegCountSQLADO = PegOperSQLADO(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, "COUNT")
End Function
Public Function PegMINSQLADO(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
947:   PegMINSQLADO = PegOperSQLADO(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, "MIN")
End Function
Public Function PegMAXSQLADO(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
950:   PegMAXSQLADO = PegOperSQLADO(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, "MAX")
End Function
Public Function PegSUMSQLADO(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
953:   PegSUMSQLADO = PegOperSQLADO(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, "SUM")
End Function
Public Function PegCampoSQLADO(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
956:   PegCampoSQLADO = PegOperSQLADO(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, "CAMPO")
End Function
Public Function PegOperSQLADO(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant, ByVal coper As String) As Variant
  Dim aRETU As Variant
  Dim cSQL As String
961:   PegOperSQLADO = eDEFAULT


964:   If coper = "CAMPO" Then
965:     If Len(cCAMPO) = 0 Then  'Passado a string ja com o campo select numero from tabela where numero=999999
966:       cTABLEWHERE = UCase(cTABLEWHERE)
967:       cSQL = Replace(cTABLEWHERE, " FROM ", " AS CAMPO FROM ")  'inclui as campo variavel padrao no pegsql abaixo
968:     Else
969:       cSQL = "SELECT " & cCAMPO & " AS CAMPO FROM " & cTABLEWHERE
970:     End If
971:   Else
972:     cSQL = "SELECT " & coper & "(" & cCAMPO & ") AS CAMPO FROM " & cTABLEWHERE
973:   End If
974:   If coper = "SUM" Or coper = "COUNT" Or coper = "MAX" Or coper = "MIN" Or IsNumeric(eDEFAULT) Then
975:     aRETU = PegSQLAdo(cARQ, cSQL, 1, Array("CAMPO"), Array("N"), Array(eDEFAULT))  ''array retorno tipo N numerico
976:   Else
977:     aRETU = PegSQLAdo(cARQ, cSQL, 1, Array("CAMPO"), Array(""), Array(eDEFAULT))
978:   End If
979:   If lRETU Then
980:     PegOperSQLADO = aRETU(0)
981:   End If
End Function
Public Function PegUltSQLAdo(ByVal cARQ As String, ByVal cSQL As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
' Esta usa loop
' use a opcao pegsqlmax para a maioria dos casos
' ou quando o campo nao e numerica
  Dim oDB As ADODB.Connection
  Dim oRS As ADODB.Recordset
  Dim lOPEN As Boolean
  Dim lRSOP As Boolean
  Dim cERRO As String

993:   On Error GoTo errhandler

995:   PegUltSQLAdo = eDEFAULT

997:   lOPEN = False
998:   lRSOP = False
999:   Set oDB = New ADODB.Connection
1000:   oDB.CursorLocation = adUseClient
1001:   oDB.ConnectionTimeout = 120
1002:   oDB.Open cARQ


1005:   lOPEN = True
1006:   Set oRS = New ADODB.Recordset

1008:   oRS.Open cSQL, oDB, adOpenForwardOnly, adLockReadOnly
1009:   lRSOP = True


1012:   If Not oRS.EOF Then
1013:     oRS.MoveLast
1014:     If Not IsNull(oRS(cCAMPO)) Then
1015:       PegUltSQLAdo = oRS(cCAMPO)
1016:     End If
1017:   End If
1018:   oRS.Close
1019:   oDB.Close
1020:   Set oRS = Nothing
1021:   Set oDB = Nothing
1022:   Exit Function
errhandler:
1024:   cERRO = "Peg Ult SQL ADO" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cSQL & Chr(13) & Chr(10)
1025:   If lRSOP Then
1026:     cERRO = cERRO & ADORsStatus(oRS.Status) & Chr(13) & Chr(10)
1027:   End If
1028:   Select Case Err.Number
  Case Else
1030:     If lOPEN Then
1031:       ADOErro oRS.ActiveConnection.Errors, cERRO
1032:     Else
1033:       ADOErro oDB.Errors, cERRO
1034:     End If
1035:     Exit Function
1036:   End Select

End Function

Public Function SQLMoveRegADO(ByVal cARQORI As String, _
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

Dim oDB  As ADODB.Connection
Dim oDBDES As ADODB.Connection
Dim oRS As ADODB.Recordset
Dim oRSDES As ADODB.Recordset
Dim cARQORI1 As String
Dim cARQDES1 As String
Dim aRETU As Variant
Dim x As Integer
Dim nCAMPOS As Integer
Dim aVALORI As Variant
Dim aRETUID As Variant
Dim lOPEN As Boolean
Dim lRSOP As Boolean
Dim cERRO As String

1067: On Error GoTo errhandler

1069: SQLMoveRegADO = False
1070: lOPEN = False
1071: lRSOP = False

1073: aRETU = TipoConn(cARQORI)

1075: cARQORI1 = aRETU(1)

1077: aRETU = TipoConn(cARQDES)

1079: cARQDES1 = aRETU(1)

1081: Set oDB = New ADODB.Connection
1082: oDB.CursorLocation = adUseClient
1083: oDB.ConnectionTimeout = 120
1084: oDB.Open cARQORI1

1086: Set oDBDES = New ADODB.Connection
1087: oDBDES.CursorLocation = adUseClient
1088: oDBDES.ConnectionTimeout = 120
1089: oDBDES.Open cARQDES1

'Set null off permitido deixar campos em branco set deleted on ja esta na string de conecao
1092: ConfigurarConexaoADO oDB, cARQORI1
1093: ConfigurarConexaoADO oDBDES, cARQDES1
'If InStr(cARQDES1, "VFPOLEDB") Then
'   VFPSetValues oDBDES
'End If
'If InStr(UCase(cARQDES1), "PGSQL") > 0 Or InStr(UCase(cARQDES1), "POSTGRESQL") > 0 Then
'        pgSetValues oDB
'End If
  
1101: lOPEN = True

1103: Set oRS = New ADODB.Recordset
1104: oRS.Open cSQLORI, oDB, adOpenKeyset, adLockOptimistic

1106: Set oRSDES = New ADODB.Recordset
1107: oRSDES.Open cSQLDES, oDBDES, adOpenKeyset, adLockOptimistic

1109: lRSOP = True

1111: If Not oRS.EOF Then
1112:    oRS.MoveFirst
1113:    While Not oRS.EOF
1114:       If Not IsNumeric(aCAMORI) Then       ''Se for numerico nao e matriz
1115:          nCAMPOS = UBound(aCAMORI)
         ReDim aVALORI(nCAMPOS + 1)
1117:          For x = 0 To nCAMPOS
1118:             aVALORI(x) = oRS(aCAMORI(x))
1119:          Next x
1120:       End If
1121:       If InStr(cOPEORI, "DEL") > 0 Then
1122:          oRS.Delete
1123:       End If
1124:       If InStr(cOPEDES, "INC") > 0 Then
1125:          oRSDES.AddNew
1126:       Else
         ''               oRSDES.Edit
1128:       End If

1130:       If Not IsNumeric(aCAMDES) Then       ''Se for numerico nao e matriz
1131:          nCAMPOS = UBound(aCAMDES)
1132:          For x = 0 To nCAMPOS
1133:             oRSDES(aCAMDES(x)) = aVALORI(x)
1134:          Next x
1135:       End If
1136:       If Not IsNumeric(aOUTDES) Then       ''Se for numerico nao e matriz
1137:          nCAMPOS = UBound(aOUTDES)
1138:          For x = 0 To nCAMPOS
1139:             oRSDES(aOUTDES(x)) = aOUTORI(x)
1140:          Next x
1141:       End If
1142:       Select Case aRETU(2) 'alguns nao aceitam updade
      Case Else
1144:          oRSDES.Update
1145:       End Select

      ' PEGAR O ID
1148:       If Not IsNumeric(aIDDES) Then        ''Se for numerico nao e matriz
1149:          nCAMPOS = UBound(aIDDES)
         ReDim aRETUID(nCAMPOS + 1)
1151:          For x = 0 To nCAMPOS
1152:             aRETUID(x) = oRSDES(aIDDES(x))
1153:          Next x
1154:          eRETU01 = aRETUID
1155:       End If
1156:       Select Case aRETU(2) 'alguns nao aceita updade
      Case Else
1158:          oRSDES.Update
1159:       End Select
1160:       oRS.MoveNext
1161:    Wend
1162: End If
1163: oRS.Close
1164: oDB.Close

1166: Set oRS = Nothing
1167: Set oDB = Nothing
1168: SQLMoveRegADO = True
1169: Exit Function
errhandler:
1171: cERRO = "Soma SQL ADO" & Chr(13) & Chr(10) & cARQORI & Chr(13) & Chr(10) & cSQLORI & Chr(13) & Chr(10)
1172: If lRSOP Then
1173:    cERRO = cERRO & ADORsStatus(oRS.Status) & Chr(13) & Chr(10)
1174: End If
1175: Select Case Err.Number
Case Else
1177:    If lOPEN Then
1178:       ADOErro oRS.ActiveConnection.Errors, cERRO
1179:    Else
1180:       ADOErro oDB.Errors, cERRO
1181:    End If
1182:    Exit Function
1183: End Select
End Function


