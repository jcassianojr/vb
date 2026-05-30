Attribute VB_Name = "sqlfuncoes"
Option Explicit
Public Function TrocaSQLOrder(ByVal cSQL As String, ByVal cSUBORDER As String) As String
  Dim nPOS As Long
  Dim cSELECT As String
5:   cSELECT = cSQL
6:   nPOS = InStr(UCase(cSQL), "ORDER BY")
7:   If nPOS > 0 Then
8:     cSELECT = Mid(cSQL, 1, nPOS - 1)
9:   End If
10:   TrocaSQLOrder = cSELECT & " Order By " & cSUBORDER
End Function

Public Function TrocaSqlWhere(ByVal cSQL As String, ByVal cSUBWHERE As String) As String
  Dim nPOS As Long
  Dim nPOS2 As Long
  Dim cSELECT As String
  Dim cORDER As String
18:   cSELECT = ""
19:   cORDER = ""
20:   nPOS = InStr(UCase(cSQL), "WHERE")
21:   nPOS2 = InStr(UCase(cSQL), "ORDER BY")
22:   If nPOS > 0 Then
23:     cSELECT = Mid(cSQL, 1, nPOS - 1)
24:   End If
25:   If nPOS2 > 0 Then
26:     cORDER = Mid(cSQL, nPOS2 - 1)
27:     If nPOS = 0 Then
28:       cSELECT = Mid(cSQL, 1, nPOS2 - 1)
29:     End If
30:   End If
31:   If nPOS = 0 And nPOS2 = 0 Then
32:     cSELECT = cSQL
33:   End If
34:   If Len(cSUBWHERE) > 0 Then
35:     TrocaSqlWhere = Trim(cSELECT) & " WHERE " & cSUBWHERE & cORDER
36:   Else
37:     TrocaSqlWhere = Trim(cSELECT) & " " & cORDER
38:   End If
End Function

Public Function NomeTableSql(ByVal cSQL As String, Optional ByVal cEXTENSAO As String = "") As String
  Dim nPOS As Integer
  Dim cNOME As String
44:   NomeTableSql = ""
45:   cSQL = UCase(cSQL)
46:   cSQL = Replace(cSQL, Chr(13), " ")
47:   cSQL = Replace(cSQL, Chr(10), " ")
48:   nPOS = InStr(cSQL, "FROM")
49:   If nPOS > 0 Then
50:     cNOME = Mid(cSQL, nPOS + 5)
51:     nPOS = InStr(cNOME, " ")
52:     If nPOS > 0 Then
53:       cNOME = Mid(cNOME, 1, nPOS - 1)
54:     End If
55:     cNOME = cNOME + cEXTENSAO
56:     NomeTableSql = cNOME
57:   Else
58:     If cEXTENSAO = ".DBF" Then               ''passado no sql o nome da tabela direto exemplo bcofgts
59:       NomeTableSql = cSQL + cEXTENSAO
60:     End If
61:   End If
End Function

Public Function MontaFiltro(ByVal aCAM As Variant, ByVal aFOR As Variant, ByVal eBUSCA As Variant, ByVal nIndex As Integer)
65:   nIndex = nIndex - 1
66:   MontaFiltro = ""
67:   Select Case aFOR(nIndex)
  Case "="
69:     MontaFiltro = " (" & aCAM(nIndex) & " =" & eBUSCA & ") "
  Case "B"
71:     MontaFiltro = " (" & aCAM(nIndex) & ") "
  Case "L%"
73:     MontaFiltro = " (" & aCAM(nIndex) & " LIKE '%" & eBUSCA & "%'" & ") "
  Case "L*"
75:     MontaFiltro = " (" & aCAM(nIndex) & " LIKE '*" & eBUSCA & "*'" & ") "
76:   End Select
End Function

Public Function GrvSQL(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, ByVal aCAM As Variant, _
                       ByVal aVAL As Variant, ByVal aFOR As Variant, Optional ByVal nStartItem = 0)
  Dim aRETU As Variant
82: If IsQueryDestructive(cSQL) Then
83:      Exit Function
84:   End If
85:   aRETU = TipoConn(cARQ)
86:   cARQ = aRETU(1)
87:   Select Case aRETU(0)
  Case "ADO"
89:     GrvSQL = GrvSQLado(cARQ, cSQL, nITEM, aCAM, aVAL, aFOR, nStartItem)
  Case "VBSQLITE"
91:       GrvSQL = GrvSQLite(cARQ, cSQL, nITEM, aCAM, aVAL, aFOR, nStartItem)
  Case "SQLITERC6", "ADORC6"
93:       GrvSQL = GrvSQLrc6(cARQ, cSQL, nITEM, aCAM, aVAL, aFOR, nStartItem)
94:   End Select
End Function

Public Function IncluiSQL(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, _
                          Optional ByVal aCAM As Variant = 0, Optional ByVal aVAL As Variant = 0, _
                          Optional ByVal lCHECK As Boolean = False, _
                          Optional ByVal lMES As Boolean = True, Optional ByVal aIDDES As Variant = 0)
  Dim aRETU As Variant
102:   If IsQueryDestructive(cSQL) Then
103:      Exit Function
104:   End If
105:   aRETU = TipoConn(cARQ)
106:   cARQ = aRETU(1)
107:   Select Case aRETU(0)
  Case "ADO"
109:     IncluiSQL = IncluiSQLAdo(cARQ, cSQL, nITEM, aCAM, aVAL, lCHECK, lMES, aIDDES)
  Case "VBSQLITE"
111:        IncluiSQL = IncluiSQLite(cARQ, cSQL, nITEM, aCAM, aVAL, lCHECK, lMES, aIDDES)
  Case "SQLITERC6", "ADORC6"
113:         IncluiSQL = IncluiSQLrc6(cARQ, cSQL, nITEM, aCAM, aVAL, lCHECK, lMES, aIDDES)
    
115:   End Select
End Function

Public Function PegSQL(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, ByVal aCAM As Variant, ByVal aFOR As Variant, ByVal aPAD As Variant, Optional ByVal lWRITE As Boolean = False) As Variant
  Dim aRETU As Variant
120:   If Not IsQuerySafe(cSQL) Then
121:      Exit Function
122:   End If
123:   aRETU = TipoConn(cARQ, , , lWRITE)
124:   cARQ = aRETU(1)
125:   Select Case aRETU(0)
       Case "ADO"
127:            PegSQL = PegSQLAdo(cARQ, cSQL, nITEM, aCAM, aFOR, aPAD)
        Case "VBSQLITE"
129:             PegSQL = PegSQLite(cARQ, cSQL, nITEM, aCAM, aFOR, aPAD)
        Case "SQLITERC6", "ADORC6"
131:             PegSQL = PegSQLrc6(cARQ, cSQL, nITEM, aCAM, aFOR, aPAD)
    
133:   End Select
End Function

Public Function PegUltSQL(ByVal cARQ As String, ByVal cSQL As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant, Optional ByVal lWRITE As Boolean = False) As Variant
  Dim aRETU As Variant
138:    If Not IsQuerySafe(cSQL) Then
139:      Exit Function
140:   End If
141:   aRETU = TipoConn(cARQ, , , lWRITE)
142:   cARQ = aRETU(1)
143:   Select Case aRETU(0)
  Case "ADO"
145:     PegUltSQL = PegUltSQLAdo(cARQ, cSQL, cCAMPO, eDEFAULT)
    Case "VBSQLITE"
147:          PegUltSQL = PegUltSQLite(cARQ, cSQL, cCAMPO, eDEFAULT)
    Case "SQLITERC6", "ADORC6"
149:          PegUltSQL = PegUltSQLrc6(cARQ, cSQL, cCAMPO, eDEFAULT)
150:   End Select
End Function

Public Function PegMAXSQL(ByVal cARQ As String, ByVal cTable As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant, Optional ByVal lWRITE As Boolean = False) As Variant
  Dim aRETU As Variant
 'If Not IsQuerySafe(cSQL) Then
 '    Exit Function
 ' End If
158:   aRETU = TipoConn(cARQ, , , lWRITE)
159:   cARQ = aRETU(1)
160:   Select Case aRETU(0)
  Case "ADO"
162:     PegMAXSQL = PegMAXSQLADO(cARQ, cTable, cCAMPO, eDEFAULT)
    Case "VBSQLITE"
164:          PegMAXSQL = PegMaxSQLite(cARQ, cTable, cCAMPO, eDEFAULT)
    Case "SQLITERC6", "ADORC6"
166:          PegMAXSQL = PegMAXSQLrc6(cARQ, cTable, cCAMPO, eDEFAULT)
    
168:   End Select
End Function

Public Function PegCountSQL(ByVal cARQ As String, ByVal cTable As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant, Optional ByVal lWRITE As Boolean = False) As Variant
  Dim aRETU As Variant
' If Not IsQuerySafe(cSQL) Then
'     Exit Function
'  End If
176:   aRETU = TipoConn(cARQ, , , lWRITE)
177:   cARQ = aRETU(1)
178:   Select Case aRETU(0)
  Case "ADO"
180:     PegCountSQL = PegCountSQLADO(cARQ, cTable, cCAMPO, eDEFAULT)
 Case "VBSQLITE"
182:      PegCountSQL = PegCountSQLite(cARQ, cTable, cCAMPO, eDEFAULT)
       Case "SQLITERC6", "ADORC6"
184:         PegCountSQL = PegCountSQLrc6(cARQ, cTable, cCAMPO, eDEFAULT)
185:   End Select
End Function
Public Function ComandoSQL(ByVal cARQ As String, ByVal cSQL As String) As Boolean
  Dim aRETU As Variant
189:   aRETU = TipoConn(cARQ)
190:   cARQ = aRETU(1)
191:   If InStr(UCase(cARQ), "VFPOLEDB") Then
192:       ComandoSQL = AdoComandodbf(cARQ, "", cSQL)
193:   End If
194:   Select Case aRETU(0)
  Case "ADO"
196:        ComandoSQL = ADOComando(cARQ, cSQL)
  Case "VBSQLITE"
198:        ComandoSQL = SQLiteComando(cARQ, cSQL)
   Case "SQLITERC6", "ADORC6"
200:         ComandoSQL = ADOComandorc6(cARQ, cSQL)
201:   End Select
End Function

Public Function ApagaSQL(ByVal cARQ As String, ByVal cSQL As String) As Boolean
  Dim aRETU As Variant
206:   If IsQueryDestructive(cSQL) Then
207:      Exit Function
208:   End If
209:   aRETU = TipoConn(cARQ)
210:   cARQ = aRETU(1)
211:   Select Case aRETU(0)
  Case "ADO"
213:        ApagaSQL = APAGASQLADO(cARQ, cSQL)
  Case "VBSQLITE"
215:        ApagaSQL = ApagaSQLite(cARQ, cSQL)
   Case "SQLITERC6", "ADORC6"
217:         ApagaSQL = APAGASQLrc6(cARQ, cSQL)
218:   End Select
End Function

Public Function SomaSQL(ByVal cARQ As String, ByVal cSQL As String, ByVal aCAM As Variant, Optional ByVal lWRITE As Boolean = False) As Variant
  Dim aRETU As Variant
223:  If Not IsQuerySafe(cSQL) Then
224:      Exit Function
225:   End If
226:   aRETU = TipoConn(cARQ, , , lWRITE)
227:   Select Case aRETU(0)
  Case "ADO"
229:        SomaSQL = SomaSQLAdo(cARQ, cSQL, aCAM)
   Case "VBSQLITE"
            ' Note que o retorno agora é um Array, espelhando o ADO
232:          SomaSQL = SomaSQLite(cARQ, cSQL, aCAM)
          
        Case "SQLITERC6", "ADORC6"
235:             SomaSQL = SomaSQLrc6(cARQ, cSQL, aCAM)
    
237:   End Select
End Function
Public Function PegSQLDeli(ByVal cARQ As String, ByVal cSQL As String, _
                           ByVal aCAM As Variant, Optional ByVal cDELI As String = ",", _
                           Optional ByVal aPAD As Variant = "", Optional ByVal aFOR As Variant = "") As Variant
    
    Dim aRETU As Variant
244:     If Not IsQuerySafe(cSQL) Then
245:      Exit Function
246:   End If
247:   aRETU = TipoConn(cARQ)
248:   Select Case aRETU(0)
    
        Case "ADO"
251:              PegSQLDeli = PegSQLDeliAdo(cARQ, cSQL, aCAM, cDELI, aPAD, aFOR)
        Case "VBSQLITE"
            ' Chama a versão VBSQLite12
254:             PegSQLDeli = PegSQLiteDeli(cARQ, cSQL, aCAM, cDELI, aPAD, aFOR)
            
        Case "SQLITERC6", "ADORC6"
            ' Chama a versão vbRichClient6
258:             PegSQLDeli = PegSQLDelirc6(cARQ, cSQL, aCAM, cDELI, aPAD, aFOR)
            
      
261:     End Select
End Function

Public Function ApagaSQLP(ByVal cARQ As String, ByVal cSQL As String, Optional ByVal cTEXTO As String = "Confirme Exclusão")
  Dim aRETU As Variant
266:   aRETU = TipoConn(cARQ)
267:   cARQ = aRETU(1)
268:   Select Case aRETU(0)
  Case "ADO"
270:     ApagaSQLP = ApagaSQLpAdo(cARQ, cSQL, cTEXTO)
271:   End Select
End Function

Public Function SqlMoveReg(ByVal cARQORI As String, _
                           ByVal cSQLORI As String, _
                           Optional ByVal cOPEORI As String = "", _
                           Optional ByVal aCAMORI As Variant = 0, _
                           Optional ByVal aOUTORI As Variant = 0, _
                           Optional ByVal cARQDES As String = "", _
                           Optional ByVal cSQLDES As String = "", _
                           Optional ByVal cOPEDES As String = "", _
                           Optional ByVal aCAMDES As Variant = 0, _
                           Optional ByVal aOUTDES As Variant = 0, _
                           Optional ByVal aIDDES As Variant = 0)

286:  If IsQueryDestructive(cSQLORI) Then
287:      Exit Function
288:   End If

290:  If IsQueryDestructive(cSQLDES) Then
291:      Exit Function
292:   End If

''cARQORI arquivo origem
''csqlori sql origem
''copeori operacoes na origem
''aCAMORI array campos origem
''aOUTORI ARRAY VARIAVEIS ORIGEM

''cARQDES arquivo origem
''csqlDES sql origem
''copeDES operacoes na origem
''aCAMDES array campos origem
''aOUTDES ARRAY VARIAVEIS ORIGEM

''aIDDES  id que acabou de criar

  Dim aRETU As Variant

  
  
312:   aRETU = TipoConn(cARQORI)
  ''cARQ = aRETU(1)
314:   Select Case aRETU(0)
  Case "ADO"
    SqlMoveReg = SQLMoveRegADO(cARQORI, cSQLORI, cOPEORI, aCAMORI, aOUTORI, _
                  cARQDES, cSQLDES, cOPEDES, aCAMDES, aOUTDES, _
                  aIDDES)
   Case "VBSQLITE"
            SqlMoveReg = SQLMoveRegSQLite(cARQORI, cSQLORI, cOPEORI, aCAMORI, aOUTORI, _
                                          cARQDES, cSQLDES, cOPEDES, aCAMDES, aOUTDES, aIDDES)
        Case "SQLITERC6", "ADORC6"
            SqlMoveReg = SQLMoveRegrc6(cARQORI, cSQLORI, cOPEORI, aCAMORI, aOUTORI, _
                                             cARQDES, cSQLDES, cOPEDES, aCAMDES, aOUTDES, aIDDES)
                  
326:   End Select
End Function

Public Function Mana5Fec() As Boolean
  Dim cDIZBX As String
  Dim cARQUSO As String
  Dim cARQNOME As String
  Dim cMes As String
  Dim cAno As String
  Dim lSEGUNDO As Boolean
  Dim aSEGUNDO As Variant
  Dim cSQLSEG As String
  Dim cARQUS2 As String
  Dim cCAMFEC As String
  Dim cFECDIZ As String
341:   lSEGUNDO = False
342:   If Len(aRELCFG(8)) > 0 Then
343:     lSEGUNDO = True
344:     cSQLSEG = Replace("SELECT * FROM MANFEC WHERE ARQORI='" & Mid(aRELCFG(8), InStrRev(aRELCFG(8), "}") + 1) & "'", ".DBF", "")
    aSEGUNDO = PegSQL(GeraConn(PegPath("PATH", "MANA5CON"), "JETFOX"), cSQLSEG, 4, Array("ARQORI", "STRBAI", "STRDES", "STRANO"), _
                      Array("CU", "CU", "CU", "CU"), Array("", "", "", ""))
347:   End If
348:   Mana5Fec = False
349:   If Len(aARQFEC(0)) > 0 Or Len(aARQFEC(1)) > 0 Or Len(aARQFEC(2)) > 0 Or Len(aARQFEC(3)) > 0 Then
350:     cDIZBX = ""
351:     aARQFEC(0) = Trim(aARQFEC(0))
352:     aARQFEC(1) = Trim(aARQFEC(1))
353:     aARQFEC(2) = Trim(aARQFEC(2))
354:     aARQFEC(3) = Trim(aARQFEC(3))
355:     If Len(aARQFEC(0)) > 0 Then
356:       cDIZBX = cDIZBX & "(A)tual"
357:     End If
358:     If Len(aARQFEC(1)) > 0 Then
359:       cDIZBX = cDIZBX & "(B)aixado"
360:     End If
361:     If Len(aARQFEC(2)) > 0 Then
362:       cDIZBX = cDIZBX & "(F)echado"
363:     End If
364:     If Len(aARQFEC(3)) > 0 Then
365:       cDIZBX = cDIZBX & "a(C)umulado"
366:     End If
367:     cCAMFEC = UCase(Busca(cDIZBX, "Utilizar", "A", 1))
368:     If Len(aARQFEC(1)) = 0 And cCAMFEC = "B" Then
369:       Alert "Não Possue Baixa"
370:       Exit Function
371:     End If
372:     If Len(aARQFEC(2)) = 0 And cCAMFEC = "F" Then
373:       Alert "Não Possue Fechamento"
374:       Exit Function
375:     End If
376:     If Len(aARQFEC(3)) = 0 And cCAMFEC = "C" Then
377:       Alert "Não Possue Acumulado"
378:       Exit Function
379:     End If
380:     If cCAMFEC = "F" Then
381:       cAno = Busca("Digite o Ano AAAA", "Competencia Ano", CStr(Format(Date, "yyyy")), 4)
382:       cMes = Busca("Digite o Mes MM", "Competencia Mes", "01", 2)
383:       cARQUSO = PegCamini("{MANA5}") & "E" & cAno & "\" & aARQFEC(2) & Right(cAno, 2) & cMes & ".DBF"
384:       If lSEGUNDO Then
385:         cARQUS2 = PegCamini("{MANA5}") & "E" & cAno & "\" & Trim(aSEGUNDO(2)) & Right(cAno, 2) & cMes & ".DBF"
386:       End If
387:       cARQNOME = aARQFEC(2)
388:       cFECDIZ = Right(cAno, 2) & cMes & "-" & cMes & "/" & cAno
389:     Else
390:       Select Case cCAMFEC
      Case "A"
392:         cARQUSO = PegCamini("{MANA5EMP}") & aARQFEC(0) & ".DBF"
393:         If lSEGUNDO Then
394:           cARQUS2 = PegCamini("{MANA5EMP}") & Trim(aSEGUNDO(0)) & ".DBF"
395:         End If
396:         cARQNOME = aARQFEC(0)
397:         cFECDIZ = "Atual"
      Case "B"
399:         cARQUSO = PegCamini("{MANA5EMP}") & aARQFEC(1) & ".DBF"
400:         If lSEGUNDO Then
401:           cARQUS2 = PegCamini("{MANA5EMP}") & Trim(aSEGUNDO(1)) & ".DBF"
402:         End If
403:         cARQNOME = aARQFEC(1)
404:         cFECDIZ = "Baixado"
      Case "C"
406:         cARQUSO = PegCamini("{MANA5EMP}") & aARQFEC(3) & ".DBF"
407:         If lSEGUNDO Then
408:           cARQUS2 = PegCamini("{MANA5EMP}") & Trim(aSEGUNDO(3)) & ".DBF"
409:         End If
410:         cARQNOME = aARQFEC(3)
411:         cFECDIZ = "Acumulado"
412:       End Select
413:     End If
414:     If cCAMFEC = "C" Then
415:       If Not MDG("Usar Acumulo Anterior", "Confirmaçao") Then
        ''Shell "somaano $" & UCase(aARQFEC(0)), vbNormalFocus
417:         ShellEx "SOMAANO", essSW_SHOWDEFAULT, "$" & UCase(aARQFEC(0)), PegPath("PATH", "SOMAANO")
418:         Alert "Clique em Ok, Após o Termino do Acumulo", "Aguarde.."
419:       End If
420:     End If
421:     aARQFEC(4) = cFECDIZ
422:     aARQFEC(5) = cARQNOME
423:     aARQFEC(6) = cARQUSO
424:     aARQUIVOS(0) = aARQFEC(6)
425:     If lSEGUNDO Then
426:       aARQUIVOS(1) = cARQUS2
427:     End If
428:     aRELCFG(6) = aARQFEC(8)                  'Retorna Mensagem Original
429:     If InStr(aRELCFG(6), "[ZFEC]") > 0 Then
430:       aRELCFG(6) = Replace(aRELCFG(6), "[ZFEC]", aARQFEC(4))
431:     End If
432:     Mana5Fec = True
433:   End If
End Function


Public Function PegOperSQL(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant, ByVal coper As String) As Variant
  Dim aRETU As Variant
 'If Not IsQuerySafe(cSQL) Then
 '    Exit Function
 ' End If
442:   aRETU = TipoConn(cARQ)
443:   Select Case aRETU(0)
  Case "ADO"
445:     PegOperSQL = PegOperSQLADO(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, coper)
 Case "SQLITERC6", "ADORC6"
447:       PegOperSQL = PegOperSQLrc6(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, coper)
        Case "VBSQLITE"
449:       PegOperSQL = PegOperSQLite(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, coper)
    
451:   End Select

End Function
Public Function IsQuerySafe(ByVal sql As String) As Boolean
    Dim cleanSQL As String
    Dim i As Integer
    Dim charCode As Integer
    Dim currentChar As String
    Dim forbiddenWords As Variant
    
    ' 1. Normaliza para maiúsculas
462:     cleanSQL = UCase(sql)
    
    ' 2. Substitui caracteres não-alfabéticos por espaços
    ' Isso separa as palavras que o atacante tentou colar
466:     For i = 1 To Len(cleanSQL)
467:         currentChar = Mid(cleanSQL, i, 1)
468:         charCode = Asc(currentChar)
        ' Mantém A-Z e números
470:         If Not ((charCode >= 65 And charCode <= 90) Or (charCode >= 48 And charCode <= 57)) Then
471:             Mid(cleanSQL, i, 1) = " "
472:         End If
473:     Next i
    
    ' 3. Remove excesso de espaços para facilitar a análise
476:     cleanSQL = Trim(Replace(Replace(cleanSQL, "  ", " "), "  ", " "))
    
    ' 4. Verifica se a primeira palavra é SELECT
    ' Dividimos para pegar apenas a primeira parte
480:     If Split(cleanSQL, " ")(0) <> "SELECT" Then
481:         IsQuerySafe = False
482:         Exit Function
483:     End If
    
    ' 5. Lista de palavras proibidas (sem espaços nas pontas)
    forbiddenWords = Array("DELETE", "DROP", "UPDATE", "INSERT", _
                           "CREATE", "ALTER", "TRUNCATE", "EXEC", _
                           "EXECUTE", "GRANT", "REVOKE", "INTO")
    
    ' 6. Verifica se alguma palavra proibida existe na string processada
    ' Como limpamos a string no passo 2, o InStr encontrará a palavra
    ' mesmo que ela estivesse grudada originalmente
493:     For i = LBound(forbiddenWords) To UBound(forbiddenWords)
494:         If InStr(cleanSQL, forbiddenWords(i)) > 0 Then
495:             IsQuerySafe = False
496:             RegistrarLogSeguranca (sql)
497:             Exit Function
498:         End If
499:     Next i
    
501:     IsQuerySafe = True
End Function

Public Function IsQueryDestructive(ByVal sql As String) As Boolean
    Dim cleanSQL As String
    Dim i As Integer, charCode As Integer
    Dim destrutivos As Variant
    
    ' 1. Normaliza para maiúsculas
510:     cleanSQL = UCase(sql)
    
    ' 2. Limpeza profunda: substitui tudo que não é letra ou número por espaço
    ' Isso garante que "DELETE;" ou "DROP--" virem "DELETE " ou "DROP "
514:     For i = 1 To Len(cleanSQL)
515:         charCode = Asc(Mid(cleanSQL, i, 1))
516:         If Not ((charCode >= 65 And charCode <= 90) Or (charCode >= 48 And charCode <= 57)) Then
517:             Mid(cleanSQL, i, 1) = " "
518:         End If
519:     Next i
    
    ' 3. Remove espaços extras para deixar a string "limpa"
522:     cleanSQL = Trim(Replace(Replace(cleanSQL, "  ", " "), "  ", " "))
    
    ' 4. Lista do que é proibido
    ' Como removemos os símbolos, basta buscar a palavra pura
526:     destrutivos = Array("DROP", "ALTER", "CREATE", "TRUNCATE", "GRANT", "REVOKE")
    
    ' 5. Verifica se alguma palavra proibida existe dentro da string limpa
    ' O InStr encontrará "DROP" mesmo se estiver no meio da frase ou sem espaços
530:     For i = LBound(destrutivos) To UBound(destrutivos)
531:         If InStr(cleanSQL, destrutivos(i)) > 0 Then
532:             IsQueryDestructive = True ' É destrutivo!
533:             RegistrarLogSeguranca (sql)
534:             Exit Function
535:         End If
536:     Next i
    
538:     IsQueryDestructive = False ' É seguro para prosseguir
End Function

Public Sub RegistrarLogSeguranca(ByVal cSQL As String)
    Dim fso As Object
    Dim ts As Object
    Dim cLogPath As String
    
546:     cLogPath = App.Path & "\log_seguranca.txt"
    
    ' Cria o objeto FSO
549:     Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' O parâmetro 8 indica "ForAppending" (Adicionar ao final)
    ' O True indica que ele deve criar o arquivo caso ele não exista
553:     Set ts = fso.OpenTextFile(cLogPath, 8, True)
    
    ' Escreve as linhas de log
556:     ts.WriteLine "DATA: " & Now
557:     ts.WriteLine "SQL: " & cSQL
558:     ts.WriteLine "--------------------------------------------------------"
    
    ' Fecha e limpa
561:     ts.Close
562:     Set ts = Nothing
563:     Set fso = Nothing
End Sub
