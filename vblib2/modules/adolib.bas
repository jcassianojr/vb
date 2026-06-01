Attribute VB_Name = "AdoLib"
Option Explicit
'
'There is no provider named 'Microsoft.ACE.OLEDB.14.0' even though it's Access 2010 (aka version number 14) the provider
'that should be used
'still is named with version number 12.
'So change from 'Microsoft.ACE.OLEDB.14.0' to 'Microsoft.ACE.OLEDB.12.0' and it will probably work better!
'
Public Const cJetA12 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
Public Const cJetADV = "Provider=Advantage.OLEDB.1;Data Source="
Public Const cJetExt = ";Extended Properties="
Public Const cJetPro = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="

  'OLEDB:Engine Type=5
  'Unknown                      0
  'Microsoft Jet 1.0            1
  'Microsoft Jet 1.1            2
  'Microsoft Jet 2.0            3
  'Microsoft Jet 3.x(97)        4
  'Microsoft Jet 4.x(2000)      5
  'JetEngineType_Ace12 =        6         accdb
  
' --- Suas existentes (Corretas) ---
Public Const JET_ENGINETYPE_UNKNOWN = 0
Public Const JET_ENGINETYPE_JET10 = 1
Public Const JET_ENGINETYPE_JET11 = 2
Public Const JET_ENGINETYPE_JET2X = 3
Public Const JET_ENGINETYPE_JET3X = 4
Public Const JET_ENGINETYPE_JET4X = 5
Public Const Jet_EngineType_Ace12 = 6 ' .accdb

' --- DBase ---
Public Const JET_ENGINETYPE_DBASE3 = 10
Public Const JET_ENGINETYPE_DBASE4 = 11
Public Const JET_ENGINETYPE_DBASE5 = 12

' --- Excel ---
Public Const JET_ENGINETYPE_EXCEL30 = 20
Public Const JET_ENGINETYPE_EXCEL40 = 21
Public Const JET_ENGINETYPE_EXCEL50 = 22
Public Const JET_ENGINETYPE_EXCEL80 = 23
Public Const JET_ENGINETYPE_EXCEL90 = 24

' --- Complementos Importantes (Faltantes) ---
Public Const JET_ENGINETYPE_EXCEL12 = 25  ' Excel 2007-2016 (.xlsx)
Public Const JET_ENGINETYPE_EXCEL12BIN = 26 ' Excel 2007-2016 (.xlsb)

' --- Exchange/Lotus ---
Public Const JET_ENGINETYPE_EXCHANGE4 = 30
Public Const JET_ENGINETYPE_LOTUSWK1 = 40
Public Const JET_ENGINETYPE_LOTUSWK3 = 41
Public Const JET_ENGINETYPE_LOTUSWK4 = 42

' --- Paradox ---
Public Const JET_ENGINETYPE_PARADOX3X = 50
Public Const JET_ENGINETYPE_PARADOX4X = 51
Public Const JET_ENGINETYPE_PARADOX5X = 52
Public Const JET_ENGINETYPE_PARADOX7X = 53

' --- Texto e HTML ---
Public Const JET_ENGINETYPE_TEXT1X = 60
Public Const JET_ENGINETYPE_HTML1X = 70

' --- Adicionais para Cobrir Lacunas ---
Public Const JET_ENGINETYPE_HTMLIMPORT = 71
Public Const JET_ENGINETYPE_XBASE = 12 ' Comumente usado como alias para DBase 5

'' rs.Open "SELECT * FROM [Salary$A1:B2] ", cn, adOpenDynamic, adLockOptimistic excel select
''                         pasta$faixa

''Geracarq Gera Um Conecao conforme tipo dbf->foxpro mdb->oledbacess ...
'' convertendo [JETFOX],[JETMDB]...
''Geraconn gera um conecao conforme tipo dbf->[FOXPRO] mdb->[JETMDB]oledbacess ...
''Tipo con gera a matrix com o tipo conecao,string concecao,tipo interno conecao
''Tipodado2 retorna o tipo C aracter D ata ...
''Tipodado retorno o tipo conforme o padrao ado
Public Function GeracArq(ByVal cARQ As String, Optional ByVal cTIPO As String = "", Optional ByVal lWRITE As Boolean = True) As String
  Dim aRETU As Variant
78:   cARQ = GeraConn(cARQ, cTIPO)
79:   aRETU = TipoConn(cARQ, , , lWRITE)
80:   GeracArq = aRETU(1)
End Function


Public Function GeraConn(ByVal cARQ As String, Optional cTIPO As String = "") As String
  Dim nPOS As Long
  Dim cARQTMP As String
  
88:   GeraConn = cARQ
89:   cARQTMP = UCase(cARQ) 'usado no instr maiscula mas sempre atribui na carq conecao e case sensitive
90:   If InStr(cARQTMP, "[") > 0 Then                 ''ja e uma connecao
91:     Exit Function
92:   End If
93:   If cTIPO = "JETMDB" Or cTIPO = "MDB" Or InStr(cARQTMP, ".MDB") > 0 Then
94:     GeraConn = "[JETMDB]" & cARQ
95:     Exit Function
96:   End If
97:   If cTIPO = "JETFOX" Or cTIPO = "FOX" Or cTIPO = "DBF" Or InStr(cARQTMP, ".DBF") > 0 Then
98:     GeraConn = "[JETFOX]" & cARQ
99:     Exit Function
100:   End If
  
'
' foxpro .dbf 'vir antes do sqlite para nao confundir com .db tambem paradox.db
'
  
106:   If InStr(cARQTMP, ".DBF") > 0 Then  'DBF
107:     nPOS = InStrRev(cARQ, "\")               ''retira no nome do arquivo
108:     cARQ = Mid(cARQ, 1, nPOS)
109:     GeraConn = "[JETFOX]" & cARQ
110:     Exit Function
111:   End If
  
  If cTIPO = "SQLITE" Or InStr(LCase(cARQ), ".sqlite") > 0 Or InStr(LCase(cARQ), ".sqlite3") > 0 _
                 Or InStr(LCase(cARQ), ".fossil") > 0 Or InStr(LCase(cARQ), ".db3") > 0 _
                 Or (InStr(LCase(cARQ), ".db") > 0 And InStr(cARQTMP, ".DBF") = 0) Then
116:     GeraConn = "[SQLITE]" & cARQ
117:     Exit Function
118:   End If

120:   If LCase(Right(cARQ, 6)) = ".accdb" Then  'InStr(LCase(cARQ), ".accdb")>0
121:     GeraConn = "[ACCDB]" & cARQ
122:     Exit Function
123:   End If
  
125:   If InStr(LCase(cARQ), ".mariadb") > 0 Then
126:     GeraConn = "[MARIADB]" & cARQ
127:     Exit Function
128:   End If
  
130:   If InStr(LCase(cARQ), ".oracle") > 0 Then
131:     GeraConn = "[ORACLE]" & cARQ
132:     Exit Function
133:   End If
  
  
  
137:   If InStr(LCase(cARQ), ".pgsql") > 0 Or InStr(LCase(cARQ), ".postgresql") > 0 Then
138:     GeraConn = "[PGSQL]" & cARQ
139:     Exit Function
140:   End If
  
  
143:   If InStr(LCase(cARQ), ".mysql") > 0 Then
144:     GeraConn = "[MYSQL]" & cARQ
145:     Exit Function
146:   End If
      
148:   If InStr(LCase(cARQ), ".oracle") > 0 Then
149:     GeraConn = "[ORACLE]" & cARQ
150:     Exit Function
151:   End If
    

  
155:   If InStr(cARQTMP, ".PD") > 0 Then   ' paradox
156:     nPOS = InStrRev(cARQ, "\")               ''retira no nome do arquivo
157:     cARQ = Mid(cARQ, 1, nPOS)
158:     GeraConn = "[JETPDX5]" & cARQ
159:     Exit Function
160:   End If
  
  If Mid(cTIPO, 1, 5) = "ACCDB" Then ' ACCDBDBFIII  ACCDBMDB PX3 PX4 PX5 XLS XLXS XLXB, CSV... _
                                   cada versao jet suporta extensoes em outros formatos de arquivos _
                                   posicao 2,3 versao 12,13,14,115 _
                                   posicao 4 MDB ACCDB formato extensao
166:      GeraConn = "[ACCDB" & Mid(cTIPO, 6) & "]"
167:      Exit Function
168:   End If
 
  
171: If Len(cTIPO) > 0 Then
172:   Select Case cTIPO
      Case "XLS"
174:         GeraConn = "[XLS]" & cARQ
      Case "XLSX"
176:         GeraConn = "[XLSX]" & cARQ
      Case "XLSDRV"
178:         GeraConn = "[XLSDRV]" & cARQ
      Case "SQLITE"
180:         GeraConn = "[SQLITE]" & cARQ
      Case "JETTXT"
182:         GeraConn = "[JETTXT]" & cARQ
      Case "DBFIII"
184:         GeraConn = "[JETDBFIII]" & cARQ
      Case "ADSDX"
186:         GeraConn = "[ADSCDX]" & cARQ
      Case "ADSNTX"
188:         GeraConn = "[ADSNTX]" & cARQ
      Case "ADSADT"
190:         GeraConn = "[ADSADT]" & cARQ
      Case "ADSVFP"
192:         GeraConn = "[ADSVFP]" & cARQ
      Case "PDX3"
194:         GeraConn = "[JETPDX3]" & cARQ
      Case "PDX4"
196:         GeraConn = "[JETPDX4]" & cARQ
      Case "PDX5"
198:         GeraConn = "[JETPDX5]" & cARQ
199:   End Select
200: End If

202: If InStr(LCase(cARQ), ".fdb") > 0 Or InStr(LCase(cARQ), ".gdb") > 0 Then
203:     GeraConn = "[FIREBIRD]" & cARQ
204:     Exit Function
205: End If
End Function

Public Function TipoConn(ByVal cARQ As String, Optional ByVal cUSER As String = "", _
                         Optional ByVal cPASS As String = "", Optional ByVal lWRITE As Boolean = True, _
                         Optional ByVal cDATABASE As String = "", Optional ByVal cowner As String = "", Optional cPADTIPOCON As String = "N") As Variant
  Dim cARQTMP As String
  Dim cJETUSO As String
  Dim lTEMMDB As Boolean
  Dim lTEMSQLITE As Boolean
  Dim lTEMMARIADB As Boolean
  Dim lTEMPG As Boolean
  Dim lTEMMYSQL As Boolean
  Dim lTEMORACLE As Boolean
  Dim lTEMFIREBIRD As Boolean
  
  Dim cADSTIP As String
  Dim cADSNOM As String
  Dim cXLSVER As String '
  Dim aCONN As Variant
  Dim cSecaoCofre As String
  Dim cAuth As String
  
  'usa boolean para agilizar if,case... no lugar de comparacao com string
229:   lTEMMDB = False
230:   lTEMSQLITE = False
231:   lTEMMARIADB = False
232:   lTEMMYSQL = False
233:   lTEMPG = False
234:   lTEMORACLE = False
235:   lTEMFIREBIRD = False
  
  
  'inicial valores padrao
239:   TipoConn = Array("ADO", cARQ, "???")
240:   cARQTMP = UCase(cARQ)
  
  
243:   If cDATABASE = "" Then
244:           cDATABASE = TratarParametrosCofre(cDATABASE)
245:   End If
  
247:   If cDATABASE <> "" Then
  ' -----------------------------------------------------------------
    ' INTEGRAÇÃO COM O COFRE (modSeguranca.bas):
    ' Se os parâmetros opcionais vierem vazios, busca no config.dat
    ' -----------------------------------------------------------------
252:     cSecaoCofre = UCase(Trim(cDATABASE))
253:     If Trim(cUSER) = "" Then
254:         cUSER = LerDoCofre(cSecaoCofre, "User")
255:     End If

257:     If Trim(cPASS) = "" Then
258:         cPASS = LerDoCofre(cSecaoCofre, "Password")
259:     End If
260:          If Trim(cowner) = "" Then
261:             cowner = LerDoCofre(cSecaoCofre, "Owner")
262:          End If


265:   End If

267:      If InStr(cARQTMP, ".MDB") > 0 Then
268:         lTEMMDB = True
269:         TipoConn = Array("ADO", cARQ, "MDB")
270:    End If

  'sqlite
  '
274:   If InStr(cARQTMP, ".SQLITE") > 0 Then
275:     lTEMSQLITE = True
276:     TipoConn = Array("ADO", cARQ, "SQLITE")
277:   End If
  '
  'mariadb
  '
281:   If InStr(cARQTMP, ".MARIADB") > 0 Or InStr(cARQTMP, "{MARIADB") > 0 Then
282:     lTEMMARIADB = True
283:     TipoConn = Array("ADO", cARQ, "MARIADB")
284:   End If
  
  '
  'mariadb
  '
289:   If InStr(cARQTMP, ".ORACLE") > 0 Or InStr(cARQTMP, "{ORACLE") > 0 Then
290:     lTEMORACLE = True
291:     TipoConn = Array("ADO", cARQ, "ORACLE")
292:   End If
   
  '
  'mysql
  '
297:   If InStr(cARQTMP, ".MYSQL") > 0 Or InStr(cARQTMP, "{MYSQL") > 0 Then
298:     lTEMMYSQL = True
299:     TipoConn = Array("ADO", cARQ, "MYSQL")
300:   End If
  
    '
  'postgresql pgsql
  '
305:   If InStr(cARQTMP, ".PGSQL") > 0 Or InStr(cARQTMP, ".POSTGRESQL") > 0 Or InStr(cARQTMP, "{POSTGRESQL") > 0 Then
306:     lTEMPG = True
307:     TipoConn = Array("ADO", cARQ, "PGSQL")
308:   End If
    '
 
' 2. No bloco de detecção de tipo (após a checagem do SQLite):
312:    If InStr(cARQTMP, ".FIREBIRD") > 0 Or InStr(cARQTMP, "{FIREBIRD") > 0 Or InStr(cARQTMP, "[FIREBIRD") > 0 Then
313:        lTEMFIREBIRD = True
314:      TipoConn = Array("ADO", cARQ, "FIREBIRD")
315:    End If

  '
  'checando provider,driver connecao
  '
320:   If InStr(cARQ, "[") = 0 Then 'inclui se nao tiver [ evita duplicar [USO][USO]
321:     If InStr(cARQTMP, "PROVIDER") > 0 Then
322:       Exit Function
323:     End If
324:     If lTEMMDB Then
325:       cARQ = "[JETMDB]" & cARQ
326:       cARQTMP = UCase(cARQ)
327:     End If
328:     If lTEMSQLITE Then
329:       cARQ = "[SQLITE]" & cARQ
330:       cARQTMP = UCase(cARQ)
331:     End If
332:     If lTEMMARIADB Then
333:       cARQ = "[MARIADB]" & cARQ
334:       cARQTMP = UCase(cARQ)
335:     End If
336:    If lTEMORACLE Then
337:       cARQ = "[ORACLE]" & cARQ
338:       cARQTMP = UCase(cARQ)
339:     End If
340:     If lTEMMYSQL Then
341:       cARQ = "[MYSQL]" & cARQ
342:       cARQTMP = UCase(cARQ)
343:     End If
344:     If lTEMPG Then
345:       cARQ = "[PGSQL]" & cARQ
346:       cARQTMP = UCase(cARQ)
347:     End If
    
349:   End If
'
'access
'
353:   If lTEMMDB Then
354:     cARQ = Replace(cARQ, "[JETMDB]", "")
355:     If InStr("PROVIDER", cARQTMP) = 0 Then
356:        cARQ = cJetPro & cARQ
357:     End If
358:     If Len(cUSER) > 0 Then
359:       cARQ = cARQ & "; User Id=" & cUSER & "; Password=" & cPASS
360:     End If
361:     If Not lWRITE Then
362:       cARQ = cARQ & ";Mode=Share Deny Write"  '";Mode=Read"
363:     Else
364:       cARQ = cARQ & ";Mode=Share Deny None"  '";Mode=Read"
365:     End If
366:     TipoConn = Array("ADO", cARQ, "MDB")
367:     Exit Function                            '' encerra agilizar
368:   End If
  
  '
  ' SQLITE
  '
373:      If lTEMSQLITE Then  'c:\Program Files (x86)\SQLite ODBC Driver\readme.txt http://www.ch-werner.de/sqliteodbc/sqliteodbc.exe
374:         cARQ = Replace(cARQ, "[SQLITE]", "")
        
        
377:       cAuth = ""
378:       If Trim(cUSER) <> "" Then cAuth = ";UID=" & cUSER
379:       If Trim(cPASS) <> "" Then cAuth = cAuth & ";PWD=" & cPASS

381:       Select Case UCase(cPADTIPOCON)
         Case "D", "N"
            ' Driver ODBC geralmente usa UID e PWD
384:             cARQ = "Driver={SQLite3 ODBC Driver};Database=" & cARQ & cAuth & ";"
            
         Case "P" ' Case Else tratado aqui para simplificar
            ' Provider OLEDB pode variar, mas segue a lógica comum de user/password
388:             cARQ = "Provider=SQLite3OLEDB.1;Data Source=" & cARQ & IIf(cAuth <> "", ";" & Mid(cAuth, 2), "") & ";"
389:       End Select


392:     TipoConn = Array("ADO", cARQ, "SQLITE")
393:     Exit Function
394:   End If
  
396:   If lTEMMARIADB Then
397:     cARQ = Replace(cARQ, "[MARIADB]", "")
398:     If InStr(cARQTMP, "MARIADB ODBC") = 0 Then 'geracom se nao passado
399:        aCONN = Split(cARQ, ".") 'localhost.port.mariadb.banco 'localhost.3306.mariadb.citacao
             'inclui  a padrao caso seja um banco de teste mas pega as corretas pelo cofre acima
401:             If cPASS = "" Then
402:                cPASS = "admin"
403:             End If
404:             If cUSER = "" Then
405:                cUSER = "root"
406:             End If
  

409:             Select Case UCase(cPADTIPOCON)
      '       cARQ = "DRIVER={MariaDB ODBC 3.2 Driver};DATABASE=" + aCONN(3) + ";SERVER=" + aCONN(0) + ";UID=" + cUSER + ";PASSWORD=" + cPASS + ";"
      
            
                    Case "D", "N": cARQ = "Driver={MariaDB ODBC 3.2 Driver};Server=" & aCONN(0) & ";Database=" & aCONN(3) & ";User=" & cUSER & ";Password=" & cPASS & ";Option=3;"
                    Case "P": cARQ = "Provider=MSDASQL;Data Source=" & aCONN(3) & ";User Id=" & cUSER & ";Password=" & cPASS & ";"
415:             End Select
416:          End If
417:      TipoConn = Array("ADO", cARQ, "MARIADB")
418:      Exit Function
419:    End If
  
  
  
423:  If lTEMORACLE Then
'    "Provider=OraOLEDB.Oracle;data source=" & _
 '   sWorld & ".World;User id=" & sUID & ";password=" & sPWD & ";"

427:      cARQ = Replace(cARQ, "[ORACLE]", "")
428:     If InStr(cARQTMP, "ODBC For Oracle") = 0 Then 'geracom se nao passado
429:         aCONN = Split(cARQ, ".") 'localhost.port.mariadb.banco 'localhost.3306.mariadb.citacao
            Dim cDefSchema As String
431:             cDefSchema = IIf(Trim(cowner) <> "", ";DefaultSchema=" & cowner, "")
               'inclui  a padrao caso seja um banco de teste mas pega as corretas pelo cofre acima
433:             If cPASS = "" Then
434:                cPASS = "admin"
435:             End If
436:             If cUSER = "" Then
437:                cUSER = "root"
438:             End If
    
 '           cARQ = "DRIVER={Microsoft ODBC For Oracle};DATABASE=" + aCONN(3) + ";SERVER=" + aCONN(0) + ";UID=" + cUSER + ";PWD=" + cPASS + ";"
 
442:             Select Case UCase(cPADTIPOCON)
                    Case "N": cARQ = "DRIVER={Microsoft ODBC For Oracle};DATABASE=" + aCONN(3) + ";SERVER=" + aCONN(0) + ";UID=" + cUSER + ";PWD=" + cPASS + ";"
                    Case "D": cARQ = "Driver={Oracle in OraClient11g_home1};Dbq=" & aCONN(0) & ";Uid=" & cUSER & ";Pwd=" & cPASS & cDefSchema
                    Case "P": cARQ = "Provider=OraOLEDB.Oracle;Data Source=" & aCONN(0) & ";User Id=" & cUSER & ";Password=" & cPASS & cDefSchema
446:             End Select


            'If Len(cUSER) > 0 Then
            '  cARQ = "DRIVER={Microsoft ODBC For Oracle};DATABASE=" + aCONN(3) + ";SERVER=" + aCONN(0) + ";UID=" + cUSER + ";PWD=" + cPASS + ";"
            'Else
            '  cARQ = "DRIVER={Microsoft ODBC For Oracle};DATABASE=" + aCONN(3) + ";SERVER=" + aCONN(0) + ";UID=root;PASSWORD=admin;"
            'End If
            'If Trim(cowner) <> "" Then
            '  cARQ = cARQ & ";DefaultSchema=" & cowner
            'End If
457:     End If
458:     TipoConn = Array("ADO", cARQ, "ORACLE")
459:     Exit Function
460:   End If
  
  
  
  

466:    If lTEMMYSQL Then
467:      cARQ = Replace(cARQ, "[MYSQL]", "")
468:       If InStr(cARQTMP, "MYSQL ODBC") = 0 Then 'geracom se nao passado
469:           aCONN = Split(cARQ, ".") 'localhost.port.mysql.banco 'localhost.5432.mysql.citacao
470:             If cPASS = "" Then
471:                cPASS = "admin"
472:             End If
473:             If cUSER = "" Then
474:                cUSER = "root"
475:             End If

    '  cARQ = "Driver={MySQL ODBC 8.0 ANSI Driver};Server=" + aCONN(0) + ";Database=" + aCONN(3) + ";Uid=" + cUSER + ";Pwd=" + cPASS + ";"     '32 driver versao 8
478:             Select Case UCase(cPADTIPOCON)
                    Case "D", "N": cARQ = "Driver={MySQL ODBC 8.0 ANSI Driver};Server=" & aCONN(0) & ";Database=" & aCONN(3) & ";User=" & cUSER & ";Password=" & cPASS & ";Option=3;"
                    Case "P": cARQ = "Provider=MySQLProv;Data Source=" & aCONN(0) & ";User Id=" & cUSER & ";Password=" & cPASS & ";"
481:                 End Select

483:          End If
484:     TipoConn = Array("ADO", cARQ, "MYSQL")
485:     Exit Function
486:   End If

488:    If lTEMPG Then
489:     cARQ = Replace(cARQ, "[PGSQL]", "")
490:     cARQ = Replace(cARQ, "[POSTGRESQL]", "")
491:     If InStr(cARQTMP, "POSTGRESQL ANSI") = 0 Then 'geracom se nao passado
492:        aCONN = Split(cARQ, ".") 'localhost.port.postgresql.banco 'localhost.5432.postgresql.citacao
            'inclui  a padrao caso seja um banco de teste mas pega as corretas pelo cofre acima
494:             If cPASS = "" Then
495:                cPASS = "admin"
496:             End If
497:             If cUSER = "" Then
498:                cUSER = "postgres"
499:             End If
            Dim cSchema As String
501:                 cSchema = IIf(Trim(cowner) <> "", ";SearchPath=" & cowner, "")
        '        cARQ = "Driver={PostgreSQL ANSI};Server=" + aCONN(0) + ";Database=" + aCONN(3) + ";Uid=" + cUSER + ";Pwd=" + cPASS + ";"
503:                 Select Case UCase(cPADTIPOCON)
                    Case "D", "N": cARQ = "Driver={PostgreSQL ANSI};Server=" & aCONN(0) & ";Database=" & aCONN(3) & ";UID=" & cUSER & ";PWD=" & cPASS & cSchema
                    Case "P": cARQ = "Provider=PostgreSQLOLEDBSample;Data Source=" & aCONN(0) & ";Initial Catalog=" & cARQ & ";User ID=" & cUSER & ";Password=" & cPASS & cSchema
506:                 End Select

      '    cARQ = "Driver={PostgreSQL ANSI};Server=" + aCONN(0) + ";Database=" + aCONN(3) + ";Uid=postgres;Pwd=admin;"
509:     End If
510:       TipoConn = Array("ADO", cARQ, "PGSQL")
511:      Exit Function
512:     End If

' 3. No bloco que aplica a string de conexão (após a checagem do SQLite/MariaDB):
515: If lTEMFIREBIRD Then
516:     cARQ = Replace(cARQ, "[FIREBIRD]", "")
    ' Ajuste o Provider conforme a instalação do driver Firebird no seu ambiente
    'cARQ = "Provider=MSDASQL;Driver={Firebird/InterBase driver};DbName=" + cARQ + ";UID=" + cUSER + ";PWD=" + cPASS + ";"
519:     TipoConn = Array("ADO", cARQ, "FIREBIRD")
    
521:     Select Case UCase(cPADTIPOCON)
        Case "D", "N" ' Driver
523:             cARQ = "Driver={Firebird/InterBase driver};DbName=" & cARQ & ";UID=" & cUSER & ";PWD=" & cPASS & ";"
        Case "P" ' Provider
525:             cARQ = "Provider=LCPI.IBProvider;Location=" & cARQ & ";User ID=" & cUSER & ";Password=" & cPASS & ";"
526:     End Select
    
    
529:     Exit Function
530: End If
  '
  'jetfox
  '
534:    If InStr(cARQTMP, "[JETFOX]") > 0 Then
535:      cARQ = Replace(cARQ, "[JETFOX]", "")

    ' PROVIDER=VFPOLEDB.1;Data Source=caminho
    ' ;SourceType=dbf;Deleted=Yes;Mode=ReadWrite|Share Deny None;Mode=Share Deny None
    ' ;DELETED=True;CODEPAGE=1252;MVCOUNT=16384;ENGINEBEHAVIOR=90
    ' ;TABLEVALIDATE=0;REFRESH=5;VARCHARMAPPING=False
    ' ;ANSI=True;REPROCESS=5;OLE DB Services = 0;
    '   MVCOUNT - sets the maximum number of variables that Visual FoxPro can maintain.
    '   ENGINEBEHAVIOR - specifies the SQL Engine compatibility level.
    '   TABLEVALIDATE - specifies the level of table validation to perform.
    '   REFRESH - specifies how often local memory buffers are refreshed with changes from other users on the network.
    '   DELETED=true/false
    '   ANSI=true/false
    '   REPROCESS=nnn
    '   CODEPAGE=nnnn
    '   VARCHARMAPPING=true/false
    '   Collating Sequence=cSequenceName
552:     If lWRITE Then  'deleted=true tambem no where pode incluir  and deleted()
553:       cARQ = "Provider=VFPOLEDB.1;Data Source=" & cARQ & ";Mode=ReadWrite|Share Deny None;Persist Security Info=False;Collating Sequence=MACHINE;DELETED=True;"  'NULL=NO"
554:     Else
555:       cARQ = "Provider=VFPOLEDB.1;Data Source=" & cARQ & ";Mode=Read|Share Deny None;Persist Security Info=False;Collating Sequence=MACHINE;DELETED=True;"  'NULL=NO"
556:     End If
557:     TipoConn = Array("ADO", cARQ, "DBF")
558:     Exit Function
559:   End If
560:   If InStr(cARQTMP, "[XLSDRV]") > 0 Then
561:     cARQ = Replace(cARQ, "[XLSDRV]", "")
562:          cARQ = "DRIVER=Microsoft Excel Driver (*.xls);" & "DBQ=" & cARQ
563:         TipoConn = Array("ADO", cARQ, "XLSDRV")
564:       Exit Function
565:      End If
566:     If InStr(cARQTMP, "[CONN]") > 0 Then
567:      cARQ = Replace(cARQ, "[CONN]", "")
568:      If lTEMMDB Then
569:         TipoConn = Array("ADO", cARQ, "MDB")
570:       Else
571:         TipoConn = Array("ADO", cARQ, "CONN")
572:        End If
573:       Exit Function
574:     End If
575:      If InStr(cARQTMP, "[SQLSERVER]") > 0 Or InStr(cARQTMP, "[MSSQL]") > 0 Then
576:         cARQ = Replace(cARQ, "[SQLSERVER]", "")
577:         cARQ = Replace(cARQ, "[MSSQL]", "")
578:           aCONN = Split(cARQ, ".") 'localhost.port.mysdql.banco 'localhost.5432.mssql.citacao

580:             cAuth = IIf(Trim(cUSER) <> "", ";UID=" & cUSER, "") & IIf(Trim(cPASS) <> "", ";PWD=" & cPASS, "")
        
582:             Select Case UCase(cPADTIPOCON)
                Case "D", "N"
584:                     cARQ = "Driver={" & GetBestMSSQL("D") & "};Server=" & aCONN(0) & ";Database=" & cARQ & cAuth & ";"
                Case "P"
586:                     cARQ = "Provider=" & GetBestMSSQL("P") & ";Data Source=" & aCONN(0) & ";Initial Catalog=" & aCONN(3) & cAuth & ";"
                Case Else
                    Dim cMelhorP As String
589:                     cMelhorP = GetBestMSSQL("P")
590:                     If cMelhorP <> "" Then
591:                         cARQ = "Provider=" & cMelhorP & ";Data Source=" & aCONN(0) & ";Initial Catalog=" & aCONN(3) & cAuth & ";"
592:                     Else
593:                         cARQ = "Driver={" & GetBestMSSQL("D") & "};Server=" & aCONN(0) & ";Database=" & aCONN(3) & cAuth & ";"
594:                     End If
595:             End Select

       
598:     TipoConn = Array("ADO", cARQ, "SQLSERVER")
599:     Exit Function
600:   End If

602:     If InStr(cARQTMP, "[ACCDB") > 0 Then
603:       cJETUSO = cJetA12
604:       cARQ = Replace(cARQ, "[ACCDB", "")
605:       cARQ = Replace(cARQ, "MDB]", "")
606:      cARQ = Replace(cARQ, "]", "")
607:     If Len(cUSER) > 0 Then
608:       cARQ = cJETUSO & cARQ & "; User Id=" & cUSER & "; Password=" & cPASS
609:     Else
610:       cARQ = cJETUSO & cARQ                ''& ";Persist Security Info=False"
611:     End If
612:     If Not lWRITE Then
613:       cARQ = cARQ & ";Mode=Share Deny Write"  '";Mode=Read"
614:     Else
615:       cARQ = cARQ & ";Mode=Share Deny None"  '";Mode=Read"
616:     End If
617:     TipoConn = Array("ADO", cARQ, "ACEOLEDB")
618:     Exit Function
619:   End If

621:   If InStr(cARQTMP, "[INFORMIX]") > 0 Then
622:     cARQ = Replace(cARQ, "[INFORMIX]", "")
623:     TipoConn = Array("ADO", cARQ, "INFORMIX")
624:     Exit Function
625:   End If
626:   If InStr(cARQTMP, "[FIREBIRD]") > 0 Then
627:     cARQ = Replace(cARQ, "[FIREBIRD]", "")
628:     TipoConn = Array("ADO", cARQ, "FIREBIRD")
629:     Exit Function
630:   End If
631:   If InStr(cARQTMP, "[INTERBASE]") > 0 Then
632:     cARQ = Replace(cARQ, "[INTERBASE]", "")
633:     TipoConn = Array("ADO", cARQ, "INTERBASE")
634:     Exit Function
635:   End If
636:   If InStr(cARQTMP, "[POSTGRESQL]") > 0 Then
637:     cARQ = Replace(cARQ, "[POSTGRESQL]", "")
638:     TipoConn = Array("ADO", cARQ, "POSTGRESQL")
639:     Exit Function
640:   End If
641:   If InStr(cARQTMP, "[UDL]") > 0 Then
642:     cARQ = Replace(cARQ, "[UDL]", "")
643:     cARQ = "File Name=" & cARQ & ";"
644:     TipoConn = Array("ADO", cARQ, "UDL")
645:     Exit Function
646:   End If
  'ADS
648:   If InStr(cARQTMP, "[ADS") > 0 Then
649:     If InStr(cARQTMP, "[ADSCDX]") > 0 Then
650:       cADSTIP = "ADS_CDX"
651:       cADSNOM = "ADSCDX"
652:     End If
653:     If InStr(cARQTMP, "[ADSNTX]") > 0 Then
654:       cADSTIP = "ADS_NTX"
655:       cADSNOM = "ADSCDX"
656:     End If
657:     If InStr(cARQTMP, "[ADSADT]") > 0 Then
658:       cADSTIP = "ADS_ADT"
659:       cADSNOM = "ADSADT"
660:     End If
661:     If InStr(cARQTMP, "[ADSVFP]") > 0 Then
662:       cADSTIP = "ADS_VFP"
663:       cADSNOM = "ADSVFP"
664:     End If

666:     cARQ = Replace(cARQ, "[" + cADSNOM + "]", "")
    cARQ = "Provider=Advantage.OLEDB.1;" & _
           "Mode=Share Deny None;" & _
           "Show Deleted Records in DBF Tables with Advantage=False;" & _
           "Data Source=" & cARQ & ";Advantage Server Type=ADS_Local_Server;" & _
           "TableType=" + cADSTIP + ";Security Mode=ADS_IGNORERIGHTS;" & _
           "Lock Mode=Compatible;" & _
           "Use NULL values in DBF Tables with Advantage=True;" & _
           "Exclusive=No;Deleted=No;" & _
           "Persist Security Info=False;Advantage Character Data Type=ADS_OEM;"
676:     TipoConn = Array("ADO", cARQ, cADSNOM)
677:     Exit Function
678:   End If
  'tratamentos JET
680:   cJETUSO = cJetPro
681:   cXLSVER = "8.0"
682:   If InStr(cARQTMP, "12") > 0 Then
683:     cJETUSO = cJetA12
684:     cXLSVER = "12.0"
685:   End If
686:   If InStr(cARQTMP, "14") > 0 Then
687:     cJETUSO = cJetA12
688:     cXLSVER = "14.0"
689:   End If
690:   If InStr(cARQTMP, "15") > 0 Then
691:     cJETUSO = cJetA12
692:     cXLSVER = "15.0"
693:   End If
694:   If InStr(cARQTMP, "16") > 0 Then
695:     cJETUSO = cJetA12
696:     cXLSVER = "16.0"
697:   End If
  
  
  
701:   If InStr(cARQTMP, "[XLS") > 0 Then
    Dim cExtProp As String
703:     cExtProp = "Excel " & cXLSVER & ";HDR=Yes;IMEX=1"
    
    ' Limpeza de todas as tags de Excel possíveis
706:     cARQ = Replace(Replace(Replace(Replace(Replace(cARQ, "[XLSX]", ""), "[XLSM]", ""), "[XLSB]", ""), "[XLS]", ""), "[XLSB]", "")
707:     cARQ = cJETUSO & cARQ & cJetExt & Chr(34) & cExtProp & Chr(34) & ";"
    
709:     TipoConn = Array("ADO", cARQ, "XLS")
710:     Exit Function
711:   End If

  ' 2. Tratamento de Texto/CSV (JETTXT, JETTXTPIPE, JETTXTCSV, CSV, etc)
714:   If InStr(cARQTMP, "[JETTXT") > 0 Or InStr(cARQTMP, "[CSV]") > 0 Then
    Dim cFmt As String
    
    ' Define o formato baseado na tag específica
718:     If InStr(cARQTMP, "[JETTXTPIPE]") > 0 Then
719:         cFmt = "FMT=Delimited(|)"
720:     ElseIf InStr(cARQTMP, "[JETTXTTAB]") > 0 Then
721:         cFmt = "FMT=TabDelimited"
722:     Else
        ' Cobre: [JETTXT], [JETTXTCSV] e a nova tag [CSV]
724:         cFmt = "FMT=Delimited"
725:     End If
    
    ' Limpeza de todas as tags de texto/csv possíveis
728:     cARQ = Replace(Replace(Replace(Replace(Replace(cARQ, "[JETTXTPIPE]", ""), "[JETTXTTAB]", ""), "[JETTXTCSV]", ""), "[JETTXT]", ""), "[CSV]", "")
    
730:     cARQ = cJETUSO & cARQ & cJetExt & Chr(34) & "text;HDR=Yes;" & cFmt & ";IMEX=1" & Chr(34) & ";"
    
732:     TipoConn = Array("ADO", cARQ, "JETXT")
733:     Exit Function
734:   End If
  
  
  
 ' 3. Tratamento de DBF (dBase III)
739:   If InStr(cARQTMP, "[JETDBFIII]") > 0 Then
    ' Limpa a tag do nome do arquivo
741:     cARQ = Replace(cARQ, "[JETDBFIII]", "")
    
    ' Para DBF, não se usa delimitadores ou HDR=Yes,
    ' a estrutura já é interpretada pelo motor dBase IV/III
745:     cARQ = cJETUSO & cARQ & cJetExt & "DBASE III;"
    
747:     TipoConn = Array("ADO", cARQ, "DBF")
748:     Exit Function
749:   End If
  
  'paradox com jet a12 ou a16
752:   If InStr(cARQTMP, "[JETPD") > 0 Then
753:     If InStr(cARQTMP, "[JETPDX3]") > 0 Then
754:       cARQ = Replace(cARQ, "[JETPDX3]", "")
755:       cARQ = cJETUSO & cARQ & cJetExt & "Paradox 3.X" & ";"
756:       TipoConn = Array("ADO", cARQ, "DB")
757:       Exit Function
758:     End If
759:     If InStr(cARQTMP, "[JETPDX4]") > 0 Then
760:       cARQ = Replace(cARQ, "[JETPDX4]", "")
761:       cARQ = cJETUSO & cARQ & cJetExt & "Paradox 4.X" & ";"
762:       TipoConn = Array("ADO", cARQ, "DB")
763:       Exit Function
764:     End If
765:     If InStr(cARQTMP, "[JETPDX5]") > 0 Then
766:       cARQ = Replace(cARQ, "[JETPDX5]", "")
767:       cARQ = cJETUSO & cARQ & cJetExt & "Paradox 5.X" & ";"
768:       TipoConn = Array("ADO", cARQ, "DB")
769:       Exit Function
770:     End If
771:   End If
End Function
Public Function DriverExisteOdbc(cNOME As String, bE_DriverODBC As Boolean) As Boolean
774:     On Error GoTo TrataErro
    
776:     If bE_DriverODBC Then
        ' Verifica Driver ODBC via Registro
        Dim oShell As Object
779:         Set oShell = CreateObject("WScript.Shell")
        ' Tenta ler a subchave padrão de Drivers ODBC
        Dim sReg As String
782:         sReg = "HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBCINST.INI\" & cNOME & "\Driver"
783:         oShell.RegRead sReg
784:         DriverExisteOdbc = True
785:         Set oShell = Nothing
786:     Else
        ' Verifica Provider OLEDB via criação de objeto COM
        Dim oCONN As Object
789:         Set oCONN = CreateObject(cNOME)
790:         DriverExisteOdbc = True
791:         Set oCONN = Nothing
792:     End If
793:     Exit Function

TrataErro:
796:     DriverExisteOdbc = False
797:     Resume Proximo
Proximo:
End Function
Public Function TipoDado2(ByVal intType As Integer) As String
801:   Select Case intType
  Case adSmallInt, adInteger, adSingle, adDouble, adCurrency, adBigInt, adBinary, _
       adDecimal, adNumeric, adTinyInt, adUnsignedBigInt, adUnsignedInt, _
       adUnsignedSmallInt, adUnsignedTinyInt
805:     TipoDado2 = "N"                          ''E um Numero
  Case adDate, adDBDate, adDBTimeStamp
807:     TipoDado2 = "D"                          ''Tipos Data
  Case adBoolean
809:     TipoDado2 = "B"                          ''Tipos Boolean
  Case adVarChar, adLongVarChar, adChar, adLongVarWChar, adVarWChar, adWChar
811:     TipoDado2 = "S"                          ''Tipos String
  Case adDBTime
813:     TipoDado2 = "T"                          ''Tipos Time
  Case adBSTR, adChapter, adFileTime, adGUID, adIDispatch, adUserDefined
815:     TipoDado2 = "X"                          ''Outros Tipos
  Case adEmpty
817:     TipoDado2 = "U"                          ''uNDEFINIDOS VAZIOS
  Case adError, adIUnknown
819:     TipoDado2 = "E"                          ''Erros
  Case adLongVarBinary
821:     TipoDado2 = "O"                          ''Ole
  Case adPropVariant, adVarBinary, adVariant   ''Variantes
823:     TipoDado2 = "V"
  Case Else
825:     TipoDado2 = "U"
826:   End Select
End Function

Public Function TipoDados(ByVal intType As Integer) As String
''EMPTY 0
''DECIMAL 14
''DBTIMESTAMP 135
''TINYINT 16
''NUMERIC 131
''BSTR 8
''SMALLINT 2 BOOLEAN 11 CHAR 129
''INTEGER 3 ERROR 10 VARCHAR 200
''BIGINT 20 USERDEFINED 132 LONGVARCHAR 201
''UNSIGNEDTINYINT 17 VARIANT 12 WCHAR 130
''UNSIGNEDSMALLINT 18 IDISPATCH 9 VARWCHAR 202
''UNSIGNEDINT 19 IUNKNOWN 13 LONGVARWCHAR 203
''UNSIGNEDBIGINT 21 GUID 72 BINARY 128
''SINGLE 4 DATE 7 VARBINARY 204
''DOUBLE 5 DBDATE 133 LONGVARBINARY 205
''CURRENCY 6 DBTIME 134 CHAPTER 136

847:   Select Case intType
  Case adSmallInt
849:     TipoDados = "SmallInt"
  Case adInteger
851:     TipoDados = "Integer"
  Case adSingle
853:     TipoDados = "Single"
  Case adDouble
855:     TipoDados = "Double"
  Case adDate
857:     TipoDados = "Date"
  Case adBoolean
859:     TipoDados = "Boolean"
  Case adVarChar
861:     TipoDados = "VarChar"
  Case adLongVarChar
863:     TipoDados = "LongVarChar"
  Case adBigInt
865:     TipoDados = "BigInt"
  Case adChar
867:     TipoDados = "Char"
  Case adDBDate
869:     TipoDados = "DBDate"
  Case adDBTime
871:     TipoDados = "DBTime"
  Case adDBTimeStamp
873:     TipoDados = "DBTimeStamp"
  Case adDecimal
875:     TipoDados = "Decimal"
  Case adFileTime
877:     TipoDados = "FileTime"
  Case adNumeric
879:     TipoDados = "Numeric"
  Case adPropVariant
881:     TipoDados = "PropVariant"
  Case adTinyInt
883:     TipoDados = "TinyInt"
  Case adUnsignedBigInt
885:     TipoDados = "UnsignedBigInt"
  Case adUnsignedInt
887:     TipoDados = "UnsignedInt"
  Case adUnsignedSmallInt
889:     TipoDados = "UnsignedSmallInt"
  Case adUnsignedTinyInt
891:     TipoDados = "UnsignedTinyInt"
  Case adLongVarBinary
893:     TipoDados = "LongVarBinary(Ole)"
  Case adLongVarWChar
895:     TipoDados = "LongVarWChar(Memo)"
  Case adVarNumeric
897:     TipoDados = "VarNumeric"
  Case adVarWChar
899:     TipoDados = "VarWChar"
  Case adWChar
901:     TipoDados = "WChar"
  Case adCurrency
903:     TipoDados = "Currency"
  Case adBinary
905:     TipoDados = "Binary"
  Case adBSTR
907:     TipoDados = "BSTR"
  Case adChapter
909:     TipoDados = "Chapter"
  Case adBinary
911:     TipoDados = "Binary"
  Case adBSTR
913:     TipoDados = "BSTR"
  Case adChapter
915:     TipoDados = "Chapter"
  Case adGUID
917:     TipoDados = "GUID"
  Case adIDispatch
919:     TipoDados = "IDispatch"
  Case adIUnknown
921:     TipoDados = "IUnknown"
  Case adGUID
923:     TipoDados = "GUID"
  Case adIDispatch
925:     TipoDados = "IDispatch"
  Case adIUnknown
927:     TipoDados = "IUnknown"

  Case Else
930:     TipoDados = "Tipo de dados não reconhecido"
931:   End Select
End Function

Public Function ADORsStatus(ByRef eSTATUS) As String
935:   ADORsStatus = "Status Nao Especificado"
936:   Select Case eSTATUS
  Case adRecOK
938:     ADORsStatus = "Registro Gravado Corretamente"
  Case adRecNew
940:     ADORsStatus = "Novo Registro"
  Case adRecModified
942:     ADORsStatus = "Registro Modificado"
  Case adRecDeleted
944:     ADORsStatus = "Registro Deletado"
  Case adRecUnmodified
946:     ADORsStatus = "Registro nao modificado"
  Case adRecInvalid
948:     ADORsStatus = "Nao Salvo - Invalido BookMark"
  Case adRecMultipleChanges
950:     ADORsStatus = "Nao Salvo - Afetou muitos Registro"
  Case adRecPendingChanges
952:     ADORsStatus = "Nao Salvo - Insercao Pendente"
  Case adRecCanceled
954:     ADORsStatus = "Nao Salvo - Operacao Cancelada"
  Case adRecCantRelease
956:     ADORsStatus = "Nao Salvo - Existe registro travado"
  Case adRecConcurrencyViolation
958:     ADORsStatus = "Nao Salvo - Cursor Optimista em Uso"
  Case adRecIntegrityViolation
960:     ADORsStatus = "Nao Salvo - Violacao de Integridade"
  Case adRecMaxChangesExceeded
962:     ADORsStatus = "Nao Salvo - Muitas Pendencias"
  Case adRecObjectOpen
964:     ADORsStatus = "Nao Salvo - Falha no meio de Armazenamento"
  Case adRecOutOfMemory
966:     ADORsStatus = "Nao Salvo - Erro Memoria Computador"
  Case adRecPermissionDenied
968:     ADORsStatus = "Nao Salvo - Falta permissao"
  Case adRecSchemaViolation
970:     ADORsStatus = "Nao Salvo - Violacao da Estrutura do Arquivo"
  Case adRecDBDeleted
972:     ADORsStatus = "Registro Ja Deletado"
  Case Else
974:     ADORsStatus = "Status Nao Especificado"
975:   End Select

End Function

Function ADOErro(ByRef oErro As Variant, Optional ByVal cERRO As String = "")
  Dim errorObject As ADODB.Error
981:   For Each errorObject In oErro
982:     cERRO = cERRO & " Ado Erro Numero: " & errorObject.Number & vbCrLf
983:     cERRO = cERRO & " Ado Descricao  : " & errorObject.Description & vbCrLf
984:     cERRO = cERRO & " Ado Fonte      : " & errorObject.Source & vbCrLf
985:     cERRO = cERRO & " Ado SQL        : " & errorObject.SQLState & vbCrLf
986:     cERRO = cERRO & " Ado Erro Nativo: " & errorObject.NativeError & vbCrLf
987:   Next
988:   SayErro cERRO, True
End Function

Public Function ADO_IsRecordsetEmpty(ByRef oRS As ADODB.Recordset) As Boolean
  Dim cERRO As String
993:   On Error GoTo TrataErro
994:   ADO_IsRecordsetEmpty = False
995:   If Not ADO_IsOpen(oRS) Then
996:     ADO_IsRecordsetEmpty = True
997:     Exit Function
998:   End If
999:   If oRS.EOF And oRS.BOF Then
1000:     ADO_IsRecordsetEmpty = True
1001:     Exit Function
1002:   End If
1003:   If oRS.RecordCount = 0 Then
1004:     ADO_IsRecordsetEmpty = True
1005:     Exit Function
1006:   End If
1007:   Exit Function
TrataErro:
1009:   cERRO = "ADO_IsRecordsetEmpty"
1010:   cERRO = cERRO & ADORsStatus(oRS.Status) & Chr(13) & Chr(10)
1011:   Select Case Err.Number
  Case Else
1013:     ADOErro oRS.ActiveConnection.Errors, cERRO
1014:   End Select
End Function

Public Function ADO_FieldValueToString(ByRef FLD As ADODB.Field, Optional ByVal sNullRepresentation As String = "(null)") As String
'Call TraceEnters(MODULE_NAME & "::ADO_FieldValueToString")
'TraceDetail = "To convert the value of a field into a string"
1020:   On Error GoTo TrataErro

1022:   ADO_FieldValueToString = ""
1023:   Select Case (FLD.Type)
    '~~~~~~
  Case adBSTR, adChar, adVarChar, adVarWChar, adWChar, adLongVarChar, adLongVarWChar, adGUID:
1026:     ADO_FieldValueToString = Nz(FLD.Value, sNullRepresentation)
    '~~~~~~
  Case adBigInt, adCurrency, adDecimal, adDouble, adInteger, adNumeric, adSingle, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt, adBoolean:
1029:     ADO_FieldValueToString = Format(Nz(FLD.Value, sNullRepresentation))
    '~~~~~~
  Case adVariant:
1032:     ADO_FieldValueToString = Format(Nz(FLD.Value, sNullRepresentation))
    '~~~~~~
  Case adDate, adDBDate, adDBTime, adDBTimeStamp:
1035:     ADO_FieldValueToString = Format(Nz(FLD.Value, sNullRepresentation))
    '~~~~~~
  Case adBinary, adVarBinary, adLongVarBinary:
1038:     ADO_FieldValueToString = IIf(IsNull(FLD.Value), sNullRepresentation, BytesToHexString(FLD.Value))
    '~~~~~~
  Case Else:
    'Err.Raise 13, MODULE_NAME & "::ADO_FieldValueToString", "Sorry, unable to convert fields of type " & fld.Type & ", (" & ADO_DataTypeEnumToEnglish(fld.Type) & ") to string."
1042:   End Select
1043:   Exit Function
TrataErro:
1045:   SayErro "ADO_FieldValueToString"
1046:   Exit Function
End Function

Public Sub ADO_FreeRecordset(ByRef rs As ADODB.Recordset)
1050:   On Error Resume Next
1051:   If rs.State = adStateOpen Then
1052:     rs.Close
1053:   End If
1054:   Set rs = Nothing
1055:   On Error GoTo 0
End Sub

Public Function ADO_IsOpen(ByRef oADOObject As Object) As Boolean

1060:     On Error GoTo TrataErro  'resume Next 'Evita quebra se o objeto estiver em estado inválido
1061:         ADO_IsOpen = False
1062:     If oADOObject Is Nothing Then Exit Function
    
    'Verifica se o estado tem o bit de "Open" ligado
1065:     ADO_IsOpen = ((oADOObject.State And adStateOpen) = adStateOpen)
        
1067:   Exit Function

TrataErro:
1070:   SayErro "ADO_ISOPEN"

End Function

Public Function TemTabelaADO(ByVal cARQ As String, ByVal cTABELA As String, Optional ByVal lMES As Boolean = True) As Boolean
  Dim oCat As ADOX.Catalog
  Dim oTabela As ADOX.Table
1077:   On Error GoTo TrataErro

1079:   TemTabelaADO = False
1080:   Set oCat = New ADOX.Catalog
1081:   oCat.ActiveConnection = cARQ

1083:   For Each oTabela In oCat.Tables
1084:     If UCase(oTabela.Name) = UCase(cTABELA) Then
1085:       TemTabelaADO = True
1086:       Exit For
1087:     End If
1088:   Next

1090:   If lMES And Not TemTabelaADO Then
1091:     Alert ("Tabela nao Encontrada" & cTABELA & Chr(13) & Chr(10) & cARQ)
1092:   End If
TrataErro:
1094:   Select Case Err.Number
  Case Else
1096:     SayErro "Tem Tabela Ado :" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cTABELA & Chr(13) & Chr(10)
1097:     Exit Function
1098:   End Select

End Function

Public Function MSSqlOdbcDriver() As String
    
    'Customize this list to include the drivers you want to support
    Dim SupportedDrivers As Variant
    SupportedDrivers = Array( _
        "ODBC Driver 17 for SQL Server", _
        "ODBC Driver 13 for SQL Server", _
        "SQL Server Native Client 11.0", _
        "SQL Server")   'the "SQL Server" driver is the legacy driver that is included with Windows
    
    'We're using the registry to check for Drivers, so we'll just create the registry object
    '   once and use it each time through the loop below
    Dim oReg As Object
1115:     Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    Const RegPath As String = "SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers"
    Const HKEY_LOCAL_MACHINE = &H80000002
    
    
    Dim i As Long
1121:     For i = 0 To UBound(SupportedDrivers)
        Dim driverName As String: driverName = SupportedDrivers(i)
        
        Dim KeyExists As Boolean, RegValue As String
1125:         KeyExists = (oReg.GetStringValue(HKEY_LOCAL_MACHINE, RegPath, driverName, RegValue) = 0)
        
        Dim DriverIsInstalled
1128:         DriverIsInstalled = (KeyExists And (RegValue = "Installed"))
        
1130:         If DriverIsInstalled Then
            'Return the first Driver from the list that is installed on the computer
1132:             MSSqlOdbcDriver = driverName
1133:             Set oReg = Nothing
1134:             Exit Function
1135:         End If
1136:     Next i
1137:     Set oReg = Nothing
    
    'If we get here it means there were no matches found;
    '   until we work out all the kinks, we may be better served by returning the default Driver
    '   here rather than throwing an error
1142:     MsgBox "No ODBC Drivers found (not even the default that ships with Windows!)"
End Function

Private Function DriverExiste(cNOME As String) As Boolean
1146:     On Error Resume Next
    Dim oCONN As Object
1148:     Set oCONN = CreateObject(cNOME)
1149:     DriverExiste = (Err.Number = 0)
1150:     Set oCONN = Nothing
End Function
' ----------------------------------------------------------------
' Procedure : GetLatestOledbProvider --MSSqlOledbProvider
' Date      : 12/14/2022
' Author    : Mike Wolfe
' Source    : https://nolongerset.com/getlatestoledbprovider/
' Purpose   : Iterates through a custom list of OLEDB providers
'               and returns the first installed match.
' ----------------------------------------------------------------
Public Function MSSqlOledbProvider(Optional ByVal nTIPO As Integer = 1) As String
    Dim SupportedProviders As Variant
    SupportedProviders = Array( _
        "MSOLEDBSQL19", "Microsoft OLE DB Driver 19 for SQL Server", "EE5DE99A-4453-4C96-861C-F8832A7F59FE", _
        "MSOLEDBSQL", "Microsoft OLE DB Driver for SQL Server", "5A23DE84-1D7B-4A16-8DED-B29C09CB648D", _
        "SQLNCLI11", "SQL Server Native Client 11.0", "397C2819-8272-4532-AD3A-FB5E43BEAA39", _
        "SQLNCLI10", "SQL Server Native Client 10.0", "8F4A6B68-4F36-4e3c-BE81-BC7CA4E9C45C", _
        "SQLOLEDB", "Microsoft OLE DB Provider for SQL Server", "0C7FF16C-38E3-11d0-97AB-00C04FC2AD98")
    
    'We're using the registry to check for providers, so we'll just create the registry object
    '   once and use it each time through the loop below
    Dim oReg As Object
1172:     Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    Const HKEY_CLASSES_ROOT = &H80000000
    
    Dim i As Long
1176:     For i = 0 To UBound(SupportedProviders) Step 3
        Dim ProviderName As String: ProviderName = SupportedProviders(i)
        Dim ProviderDesc As String: ProviderDesc = SupportedProviders(i + 1)
        Dim ProviderUID As String: ProviderUID = SupportedProviders(i + 2)
        
        Dim ProviderIsInstalled As Boolean, SubkeyPath As String
1182:         SubkeyPath = "CLSID\{" & ProviderUID & "}\"
1183:         ProviderIsInstalled = (oReg.EnumKey(HKEY_CLASSES_ROOT, SubkeyPath, "", "") = 0)
        
1185:         If ProviderIsInstalled Then
            'Return the first provider from the list that is installed on the computer
1187:             Select Case nTIPO
                   Case 1
1189:                        MSSqlOledbProvider = ProviderName
                   Case 2
1191:                        MSSqlOledbProvider = ProviderDesc
                   Case 3
1193:                        MSSqlOledbProvider = ProviderUID
1194:             End Select
1195:             Set oReg = Nothing
1196:             Exit Function
1197:         End If
1198:     Next i
1199:     Set oReg = Nothing
    
    'If we get here it means there were no matches found
1202:     MsgBox "No OLE DB providers found (not even the default that ships with Windows!)"
End Function


Public Function ado_GetCursorType(CursorType As Integer) As String

    ' *******************
    ' CursorTypeEnum
    '
    ' adOpenDynamic = 2
    ' adOpenForwardOnly = 0
    ' adOpenKeyset = 1
    ' adOpenStatic = 3
    ' adOpenUnspecified = -1
    '*******************
    
1218:     Select Case CursorType
        Case 2
1220:             ado_GetCursorType = "Dynamic"
        Case 0
1222:             ado_GetCursorType = "ForwardOnly"
       Case 1
1224:             ado_GetCursorType = "Keyset"
        Case 3
1226:             ado_GetCursorType = "Static"
        Case -1
1228:             ado_GetCursorType = "Unspecified"
1229:     End Select
End Function

Function ado_GetLockType(LockType As Integer) As String
    
    '*******************
    ' LockTypeEnum
    '
    ' adLockBatchOptimistic = 4
    ' adLockOptimistic = 3
    ' adLockPessimistic = 2
    ' adLockReadOnly = 1
    ' adLockUnspecified = -1
    '*******************
    
1244:     Select Case LockType
        Case 4
1246:             ado_GetLockType = "BatchOptimistic"
        Case 3
1248:             ado_GetLockType = "Optimistic"
       Case 2
1250:             ado_GetLockType = "Pessimistic"
        Case 1
1252:             ado_GetLockType = "ReadOnly "
        Case -1
1254:             ado_GetLockType = "Unspecified"
1255:     End Select

End Function
Public Function ExtraiWhere(ByVal cSQL As String) As String
    Dim nPOS As Long
    Dim sUpperSQL As String
    
1262:     sUpperSQL = UCase(cSQL)
1263:     nPOS = InStr(sUpperSQL, "WHERE ")
    
1265:     If nPOS > 0 Then
        ' Retorna do "WHERE" até ao fim da string
1267:         ExtraiWhere = Trim(Mid(cSQL, nPOS))
1268:     Else
        ' Se não houver WHERE, retorna vazio (CUIDADO: isso afetaria a tabela toda)
1270:         ExtraiWhere = ""
1271:     End If
End Function


Public Function ExtraiTabela(ByVal cSQL As String) As String
    ' Busca o nome da tabela após o "FROM"
    Dim nPOS As Long
1278:     cSQL = UCase(cSQL)
1279:     nPOS = InStr(cSQL, "FROM ")
1280:     If nPOS > 0 Then
1281:         ExtraiTabela = Trim(Mid(cSQL, nPOS + 5))
        ' Remove o resto se houver WHERE ou ORDER
1283:         If InStr(ExtraiTabela, " ") > 0 Then
1284:             ExtraiTabela = Left(ExtraiTabela, InStr(ExtraiTabela, " ") - 1)
1285:         End If
1286:     End If
End Function


Public Function PrepararValorSQL(ByVal v As Variant) As String
1291:     If IsNull(v) Then
1292:         PrepararValorSQL = "NULL"
1293:     ElseIf IsNumeric(v) Then
        ' No SQLite, decimais devem usar PONTO, não vírgula
1295:         PrepararValorSQL = Replace(CStr(v), ",", ".")
1296:     ElseIf IsDate(v) Then
1297:         PrepararValorSQL = "'" & Format(v, "yyyy-mm-dd") & "'"
1298:     Else
        ' Texto: dobra as aspas simples para não quebrar o SQL
1300:         PrepararValorSQL = "'" & Replace(v, "'", "''") & "'"
1301:     End If
End Function

Public Function TratarParametrosCofre(ByVal cARQ As String) As String
    Dim cBancoPuro As String
    Dim cUpperARQ As String
    Dim aCONN() As String
    Dim nPosStart As Long
    Dim nPosEnd As Long
    
1311:     cBancoPuro = ""
1312:      cUpperARQ = UCase(Trim(cARQ))
    
    ' -------------------------------------------------------------------------
    ' CASO 1: Arquivos locais com caminho (\pasta\arquivo.sqlite, .mdb, .accdb)
    ' -------------------------------------------------------------------------
    If InStr(cUpperARQ, ".SQLITE") > 0 Or InStr(cUpperARQ, ".MDB") > 0 _
             Or InStr(cUpperARQ, ".FDB") > 0 Or InStr(cUpperARQ, ".GDB") > 0 _
                                        Or InStr(cUpperARQ, ".ACCDB") > 0 Then
        Dim nPosBarra As Long
        Dim nPosPonto As Long
        Dim cNomeComExtensao As String
        
        ' Acha a última barra do caminho
1325:         nPosBarra = InStrRev(cARQ, "\")
1326:         If nPosBarra = 0 Then nPosBarra = InStrRev(cARQ, "/")
        
1328:         cNomeComExtensao = Mid(cARQ, nPosBarra + 1)
        
        ' Remove a extensão para pegar só o nome limpo do banco
1331:         nPosPonto = InStrRev(cNomeComExtensao, ".")
1332:         If nPosPonto > 0 Then
1333:             cBancoPuro = Left(cNomeComExtensao, nPosPonto - 1)
1334:         Else
1335:             cBancoPuro = cNomeComExtensao
1336:         End If

    ' -------------------------------------------------------------------------
    ' CASO 2: String delimitada por pontos (ex: localhost.3306.mariadb.citacao)
    ' -------------------------------------------------------------------------
1341:     ElseIf InStr(cARQ, ".") > 0 And InStr(cUpperARQ, "DRIVER=") = 0 Then
1342:         aCONN = Split(cARQ, ".")
1343:         If UBound(aCONN) >= 3 Then
1344:             cBancoPuro = aCONN(3) ' Pega o quarto elemento (o nome do banco)
            
            ' Aproveita e captura o Servidor se ele estiver vazio
            'If Trim(cSERVER) = "" Then cSERVER = aCONN(0)
1348:         End If

    ' -------------------------------------------------------------------------
    ' CASO 3: String de conexão ODBC Padrão (Ex: DRIVER={...};DATABASE=nomebanco)
    ' Trata variações normais de mercado como DATABASE= ou DB=
    ' -------------------------------------------------------------------------
1354:     ElseIf InStr(cUpperARQ, "DATABASE=") > 0 Or InStr(cUpperARQ, "DB=") > 0 Then
        
        ' Identifica qual marcador foi utilizado na string
1357:         If InStr(cUpperARQ, "DATABASE=") > 0 Then
1358:             nPosStart = InStr(cUpperARQ, "DATABASE=") + 9
1359:         Else
1360:             nPosStart = InStr(cUpperARQ, "DB=") + 3
1361:         End If
        
        ' Procura o próximo ponto e vírgula que encerra o nome do banco
1364:         nPosEnd = InStr(nPosStart, cUpperARQ, ";")
        
1366:         If nPosEnd > 0 Then
1367:             cBancoPuro = Mid(cARQ, nPosStart, nPosEnd - nPosStart)
1368:         Else
1369:             cBancoPuro = Mid(cARQ, nPosStart)
1370:         End If
        
        ' Limpa aspas ou caracteres residuais que algumas strings ODBC podem conter
1373:         cBancoPuro = Replace(cBancoPuro, """", "")
1374:         cBancoPuro = Replace(cBancoPuro, "'", "")
1375:     End If
    
    ' Limpa espaços em branco e garante caixa alta para casar com o Harbour
1378:     cBancoPuro = UCase(Trim(cBancoPuro))
    
    ' -------------------------------------------------------------------------
    ' CONSULTA AO COFRE (modSeguranca.bas)
    ' Se as variáveis opcionais vierem vazias, busca usando o banco extraído
    ' -------------------------------------------------------------------------
    'If cBancoPuro <> "" Then
    '    If Trim(cUSER) = "" Then
    '        cUSER = LerDoCofre(cBancoPuro, "User")
    '    End If
        
    '    If Trim(cPASS) = "" Then
    '        cPASS = LerDoCofre(cBancoPuro, "Password")
    '    End If
        
    '    If Trim(cSERVER) = "" Then
    '        cSERVER = LerDoCofre(cBancoPuro, "Server")
    '    End If
    'End If

    ' Retorna o nome do banco identificado (caso precise usar na tipoconn)
1399:     TratarParametrosCofre = cBancoPuro
End Function

Public Function BytesToHexString(vaBytes As Variant) As String
' Purpose: To translate a Byte() Array into human readable Format
' Example/Note: BytesToHexString
' !! Assumes/Pre: Nothing
' Parameters:
'   vaBytes - Variant byte array
' Returns: String
'       Success- String with leading "0x" to denote hexadecimal format.
'       Failure- Raises error on failure
' Dependencies: None
'        mod->Sub
' Revision history:
'   Michael Johnson     2000/Aug/01 13:51     Initial creation
1415:   On Error GoTo TrataErro

  Dim sAccumulator As String
  Dim lCtr As Long
  Dim sHex As String
  Dim sFormatted As String

1422:   BytesToHexString = ""
1423:   If Not TypeName(vaBytes) = "Byte()" Then
1424:     BytesToHexString = ""
1425:     Exit Function
1426:   End If

1428:   For lCtr = 0 To UBound(vaBytes)
1429:     sHex = Hex(vaBytes(lCtr))
1430:     sFormatted = Format(sHex, "@@")
1431:     sAccumulator = sAccumulator & Replace(sFormatted, " ", "0")
1432:   Next
1433:   sAccumulator = "0x" & sAccumulator
1434:   BytesToHexString = sAccumulator
1435:   Exit Function
TrataErro:
1437:   BytesToHexString = ""
1438:   Exit Function
End Function

' Exemplo de função auxiliar para padronizar
Private Function GetJetExtendedProperties(cFormato As String, Optional lIMEX As Boolean = False) As String
    Dim sProps As String
1444:     sProps = cFormato & ";HDR=Yes"
1445:     If lIMEX Then sProps = sProps & ";IMEX=1"
1446:     GetJetExtendedProperties = ";Extended Properties='" & sProps & "';"
End Function


Public Function GetBestMSSQL(TIPO As String) As String
    Dim SupportedDrivers, SupportedProviders
    Dim item As Variant
    
1454:     SupportedDrivers = Array("ODBC Driver 17 for SQL Server", "ODBC Driver 13 for SQL Server", "SQL Server Native Client 11.0", "SQL Server")
1455:     SupportedProviders = Array("MSOLEDBSQL19", "MSOLEDBSQL", "SQLNCLI11", "SQLOLEDB")

1457:     If TIPO = "D" Then
1458:         For Each item In SupportedDrivers
            ' AQUI: Você deve passar o nome (item) e True (pois é driver ODBC)
1460:             If DriverExisteOdbc(CStr(item), True) Then
1461:                 GetBestMSSQL = CStr(item)
1462:                 Exit Function
1463:             End If
1464:         Next
1465:     Else
1466:         For Each item In SupportedProviders
            ' AQUI: Você deve passar o nome (item) e False (pois é Provider OLEDB)
1468:             If DriverExisteOdbc(CStr(item), False) Then
1469:                 GetBestMSSQL = CStr(item)
1470:                 Exit Function
1471:             End If
1472:         Next
1473:     End If
End Function

Public Function ExtrairNomeTabela(ByVal cSQL As String) As String
    Dim nPosFrom As Long
    Dim nPosWhere As Long
    Dim nPosOrder As Long
    Dim cTemp As String
    
    ' Converte para maiúsculas para facilitar a busca
1483:     cTemp = " " & UCase(cSQL) & " "
    
    ' Localiza a posição do " FROM "
1486:     nPosFrom = InStr(1, cTemp, " FROM ")
1487:     If nPosFrom = 0 Then
1488:         ExtrairNomeTabela = ""
1489:         Exit Function
1490:     End If
    
1492:     nPosFrom = nPosFrom + 6 ' Tamanho da string " FROM "
    
    ' Localiza o primeiro limitador após o FROM (WHERE, ORDER BY, GROUP BY)
1495:     nPosWhere = InStr(nPosFrom, cTemp, " WHERE ")
1496:     nPosOrder = InStr(nPosFrom, cTemp, " ORDER BY ")
    
    ' Define o fim do nome da tabela (o menor valor entre os delimitadores)
    Dim nFIM As Long
1500:     nFIM = Len(cTemp)
    
1502:     If nPosWhere > 0 And nPosWhere < nFIM Then nFIM = nPosWhere
1503:     If nPosOrder > 0 And nPosOrder < nFIM Then nFIM = nPosOrder
    
    ' Extrai e limpa o nome da tabela
1506:     ExtrairNomeTabela = Trim(Mid(cSQL, nPosFrom - 1, nFIM - nPosFrom + 1))
End Function
