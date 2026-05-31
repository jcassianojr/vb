Attribute VB_Name = "GridFuncoes"
Option Explicit

'mudar a cor da coluna
'    Dim i As Integer
'    For i = 1 To .rows - 1
'   .Row = i
'   .Col = 2:   .CellBackColor = &HC0FFFFD
'    Next
Public Function LocalizaGri1(ByRef oGRID As Variant, Optional ByVal eBUSCA As Variant = "", _
                             Optional ByVal iBUSCA As Integer = 1, _
                             Optional ByVal lMES As Boolean = True, _
                             Optional ByVal iPos As Integer = 0) As Double
13:   LocalizaGrid oGRID, eBUSCA, iBUSCA, lMES, iPos
End Function

Public Function SomaGrid(ByRef oGRID As Variant, Optional ByVal nCOL As Integer = 0, Optional ByVal nROWINI As Integer = 0)
  Dim nROWS  As Integer
  Dim nROW As Integer
  Dim x As Integer
  Dim nVAL  As Variant
  Dim nTMPVAL As Variant
22:   nROWS = oGRID.Rows - 1                       ''gRID ROW Indice inicia 0
23:   oGRID.Col = nCOL
24:   For x = nROWINI To nROWS
25:     oGRID.Row = x
26:     nTMPVAL = oGRID
27:     nVAL = nVAL + FixNum(nTMPVAL)
28:   Next x
29:   SomaGrid = nVAL
End Function

Public Function PegUltGrid(ByRef oGRID As Variant, Optional ByVal nCOL As Integer = 0)
  Dim nROWS  As Integer
  Dim nROW As Integer
35:   nROWS = oGRID.Rows - 1                       ''gRID ROW Indice inicia 0
36:   oGRID.Col = nCOL
37:   oGRID.Row = nROWS
38:   PegUltGrid = oGRID
End Function

Public Function LocalizaGrid(ByRef oGRID As Variant, Optional ByVal eBUSCA As Variant = "", _
                             Optional ByVal iBUSCA As Integer = 1, _
                             Optional ByVal lMES As Boolean = True, _
                             Optional ByVal iPos As Integer = 1) As Double
'Ogrid grid
'eBUSCA Expressao para Busca
'iBUSCA Indice
'lMes exibe mensagem padrao sim/nao
'iPOS Incremento posicional pois geralmente a col0 e o id e nao é buscada
  Dim nROWS As Integer
  Dim nROW As Integer
  Dim x As Integer
  Dim nTAM As Integer
  Dim lFOUND As Boolean
  Dim cCOMPARE As String
56:   lRETU = True
57:   LocalizaGrid = oGRID.Row
58:   nROWS = oGRID.Rows - 1                       ''gRID ROW Indice inicia 0
59:   nROW = oGRID.Row
60:   eBUSCA = FixStr(eBUSCA, "")
61:   iBUSCA = FixNum(iBUSCA)
62:   If eBUSCA = "" Or iBUSCA = 0 Then
63:     frmLocalizar.Show vbModal
64:     eBUSCA = FixStr(eRETU01)
65:     iBUSCA = FixNum(eRETU02) + iPos          '´´Mais Um Pois 0 e sempre id
66:   End If
67:   If iBUSCA > 0 Then iBUSCA = iBUSCA - 1       ''gRID COW Inicia 0
68:   oGRID.Col = iBUSCA
69:   nTAM = Len(eBUSCA)
70:   eBUSCA = UCase(eBUSCA)
71:   For x = 0 To nROWS
72:     oGRID.Row = x
73:     cCOMPARE = UCase(Left(FixStr(oGRID.Text), nTAM))
74:     If cCOMPARE = eBUSCA Then
75:       lFOUND = True
76:       oGRID.Col = oGRID.cols - 1
77:       oGRID.ColSel = 0
78:       If x > 0 Then
79:         oGRID.TopRow = x
80:       End If
81:       Exit For
82:     End If
83:   Next
84:   If Not lFOUND Then                           ''Volta a linha que estava
85:     lRETU = False
86:     oGRID.Row = nROW
87:   End If
88:   If lMES And Not lFOUND Then
89:     Alert "Não Localizado"
90:   End If
91:   eLOCALIZA = ""
End Function

Public Sub MontaGrid(ByRef oGRID As Variant, _
                     ByVal nITEM As Integer, _
                     ByVal aTAM As Variant, _
                     ByVal aDIZ As Variant, _
                     ByVal aCAM As Variant, _
                     ByVal cARQ As String, _
                     ByVal cSQL As String, _
                     Optional ByVal aFOR As Variant = 0)

103:   On Error GoTo errhandler
  Dim x As Integer
  Dim cDIZ As Variant
  Dim oDB As New ADODB.Connection
  Dim oRS As New ADODB.Recordset
  Dim aRETU As Variant
  Dim cARQTMP As String
  Dim eVAR As Variant
  Dim aOPE As Variant
  Dim l3265 As Boolean
  Dim cERRO As String


116:   l3265 = True


119:   If Not FileConnExist(cARQ, True, , cSQL) Then
120:     Exit Sub
121:   End If

123:   aRETU = TipoConn(cARQ, , , False)

125:   cARQTMP = aRETU(1)

127:   If aRETU(2) <> "MDB" Then
128:     cSQL = Replace(cSQL, "[", "")
129:     cSQL = Replace(cSQL, "]", "")
130:   End If

132:   If aRETU(2) = "PGSQL" Then
133:      cSQL = SQLPGSQLDOUBLEQUOTES(cSQL)
134:   End If


137:   Select Case aRETU(0)
  Case "ADO"
139:     Set oDB = New ADODB.Connection
140:     oDB.ConnectionTimeout = 120
141:     oDB.Open cARQTMP


   'sqlite nao tem cursor forwardonly openstatic ou adOpenUnspecified
145:     Set oRS = New ADODB.Recordset
146:     If aRETU(2) = "SQLITE" Then
147:       oRS.Open cSQL, oDB, adOpenStatic, adLockReadOnly
148:     Else
149:       oRS.Open cSQL, oDB, adOpenForwardOnly, adLockReadOnly
150:     End If
151:   End Select

  
154:   oGRID.Visible = False
155:   With oGRID
156:     .cols = nITEM
157:     .FixedCols = 0
158:     .Rows = 1
159:     .Row = 0
160:     For x = 0 To nITEM - 1
161:       .Col = x
162:       .Text = aDIZ(x)
163:       .ColWidth(x) = aTAM(x)
164:       If InStr(aCAM(x), "$") > 0 Then
165:         Select Case Mid(aCAM(x), 1, 1)  ' FlexAlignmentCENTERCenter FlexAlignmentRIGHTCenter FlexAlignmentLeftCenter
        Case "L"
167:           .ColAlignment(x) = FlexAlignmentLeftCenter 'flexAlignLeftCenter
        Case "R"
169:           .ColAlignment(x) = FlexAlignmentRightCenter 'flexAlignRightCenter
        Case "C"
171:           .ColAlignment(x) = FlexAlignmentCenterCenter 'flexAlignCenterCenter
172:         End Select
173:         aCAM(x) = Mid(aCAM(x), InStr(aCAM(x), "$") + 1)
174:       End If
175:     Next

177:     While Not oRS.EOF
178:       cDIZ = ""
179:       For x = 0 To nITEM - 1
180:         aOPE = SepSqlOpe(aCAM(x))
181:         If aOPE(0) = "" Or aOPE(1) = "" Or aOPE(2) = "" Then  ''Operacao com Campos
182:           eVAR = oRS(aCAM(x))          ''Grava o Valor Default
183:         Else
184:           eVAR = MathOper(oRS(aOPE(1)), oRS(aOPE(2)), aOPE(0))
185:         End If
186:         Select Case VarType(eVAR)
        Case vbDate
188:           If IsNumeric(aFOR) Then
189:             eVAR = Fdata(eVAR, "DF")  ''Nao tem matriz formatacao
190:           Else
191:             If aFOR(x) = "" Then
192:               eVAR = Fdata(eVAR, "DF")  ''Nao tem formatacao
193:             End If
194:           End If
        Case vbBoolean
196:           If eVAR Then
197:             eVAR = " * "
198:           Else
199:             eVAR = "   "
200:           End If
        Case vbString
202:           If Len(eVAR) > 100 Then
203:             eVAR = Mid(eVAR, 1, 100)
204:           End If
205:           eVAR = Replace(eVAR, Chr(13) & Chr(10), " ")
206:           If aRETU(2) = "SQLITE" Then
207:              If Mid(eVAR, 5, 1) = "-" And Mid(eVAR, 8, 1) = "-" Then 'datetime no sqlite vem no texto yyyy-mm-dd hh:mm:ss
208:                 eVAR = Mid(eVAR, 9, 2) + "/" + Mid(eVAR, 6, 2) + "/" + Mid(eVAR, 1, 4)
209:              End If
210:           End If
          
212:         End Select
213:         If Not IsNumeric(aFOR) Then      'E Matriz
214:           If aFOR(x) <> "" Then
215:             eVAR = Format(eVAR, aFOR(x))
216:           End If
217:         End If

219:         cDIZ = cDIZ & eVAR & vbTab
220:       Next

222:       .AddItem cDIZ
223:       oRS.MoveNext

225:     Wend

227:   End With

229:   oGRID.Visible = True
230:   oRS.Close
231:   oDB.Close

233:   Set oRS = Nothing
234:   Set oDB = Nothing

236:   Exit Sub

errhandler:

240:   cERRO = "Montagrid " & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cSQL & Chr(13) & Chr(10)
241:   Select Case Err.Number
  Case 424, 3265                               ''Campo Inexistente
243:     If l3265 Then                            ''So uma vez o erro
244:       cERRO = cERRO & "Campo:" & aCAM(x) & Chr(13) & Chr(10)
245:       SayErro cERRO
246:       l3265 = False
247:     End If
248:     cDIZ = cDIZ & " ? " & vbTab
249:     Resume Next
  Case 3051, 3262
251:     oGRID.Visible = True
252:     Alert "Arquivo em Uso Exclusivo Outro Usuario"
253:     SayErro cERRO
254:     Exit Sub
  Case Else
256:     oGRID.Visible = True
257:     SayErro cERRO
258:     Exit Sub
259:   End Select

End Sub

Public Sub MontaGridFast(ByRef oGRID As Variant, _
                         ByVal nITEM As Integer, _
                         ByVal aTAM As Variant, _
                         ByVal aDIZ As Variant, _
                         ByVal aCAM As Variant, _
                         ByVal cARQ As String, _
                         ByVal cSQL As String, _
                         Optional ByVal aFOR As Variant = 0)

272:   On Error GoTo errhandler
  Dim x As Integer
  Dim cDIZ As Variant
  Dim oDB As New ADODB.Connection
  Dim oRS As New ADODB.Recordset
  Dim oCAM00 As ADODB.Field
  Dim oCAM01 As ADODB.Field
  Dim oCAM02 As ADODB.Field
  Dim oCAM03 As ADODB.Field
  Dim oCAM04 As ADODB.Field
  Dim oCAM05 As ADODB.Field
  Dim oCAM06 As ADODB.Field
  Dim oCAM07 As ADODB.Field
  Dim oCAM08 As ADODB.Field
  Dim oCAM09 As ADODB.Field
  Dim cARQTMP As String
  Dim eVAR As Variant
  Dim l3265 As Boolean
  Dim cERRO As String
  Dim lOPEN As Boolean
  Dim aRETU As Variant
  

295:   lOPEN = False
296:   l3265 = True

298:   If Not FileConnExist(cARQ, True, , cSQL) Then
299:     Exit Sub
300:   End If

302:   If nITEM > 10 Then
303:     Alert ("MontaGridFast até 10 itens")
304:     Exit Sub
305:   End If

307:   aRETU = TipoConn(cARQ, , , False)

309:   cARQTMP = aRETU(1)

311:   If aRETU(2) = "PGSQL" Then
312:      cSQL = SQLPGSQLDOUBLEQUOTES(cSQL)
313:   End If


'  cARQTMP = GeracArq(cARQ, , False)

318:   oDB.ConnectionTimeout = 120
319:   oDB.Open cARQTMP



323:   lOPEN = True

325:   If aRETU(2) = "SQLITE" Then 'sqlite nao tem curos forwardonly usar static ou adOpenUnspecified
326:      oRS.Open cSQL, oDB, adOpenStatic, adLockReadOnly
327:   Else
328:      oRS.Open cSQL, oDB, adOpenForwardOnly, adLockReadOnly
329:   End If

331:   oGRID.Visible = False
332:   With oGRID
333:     .cols = nITEM
334:     .FixedCols = 0
335:     .Rows = 1
336:     .Row = 0
337:     For x = 0 To nITEM - 1
338:       .Col = x
339:       .Text = aDIZ(x)
340:       .ColWidth(x) = aTAM(x)
341:       If InStr(aCAM(x), "$") > 0 Then
342:         Select Case Mid(aCAM(x), 1, 1)  ' FlexAlignmentCENTERCenter FlexAlignmentRIGHTCenter FlexAlignmentLeftCenter
        Case "L"
344:           .ColAlignment(x) = FlexAlignmentLeftCenter 'flexAlignLeftCenter
        Case "R"
346:           .ColAlignment(x) = FlexAlignmentRightCenter 'flexAlignRightCenter
        Case "C"
348:           .ColAlignment(x) = FlexAlignmentCenterCenter 'flexAlignCenterCenter
349:         End Select
350:         aCAM(x) = Mid(aCAM(x), InStr(aCAM(x), "$") + 1)
351:       End If
352:     Next


355:     If Not oRS.EOF Then
356:       For x = 0 To nITEM - 1
357:         Select Case x
        Case 0
359:           Set oCAM00 = oRS(aCAM(0))
        Case 1
361:           Set oCAM01 = oRS(aCAM(1))
        Case 2
363:           Set oCAM02 = oRS(aCAM(2))
        Case 3
365:           Set oCAM03 = oRS(aCAM(3))
        Case 4
367:           Set oCAM04 = oRS(aCAM(4))
        Case 5
369:           Set oCAM05 = oRS(aCAM(5))
        Case 6
371:           Set oCAM06 = oRS(aCAM(6))
        Case 7
373:           Set oCAM07 = oRS(aCAM(7))
        Case 8
375:           Set oCAM08 = oRS(aCAM(8))
        Case 9
377:           Set oCAM09 = oRS(aCAM(9))
378:         End Select
379:       Next x
380:     End If



384:     While Not oRS.EOF
385:       cDIZ = ""
386:       For x = 0 To nITEM - 1
387:         Select Case x
        Case 0
389:           eVAR = oCAM00
        Case 1
391:           eVAR = oCAM01
        Case 2
393:           eVAR = oCAM02
        Case 3
395:           eVAR = oCAM03
        Case 4
397:           eVAR = oCAM04
        Case 5
399:           eVAR = oCAM05
        Case 6
401:           eVAR = oCAM06
        Case 7
403:           eVAR = oCAM07
        Case 8
405:           eVAR = oCAM08
        Case 9
407:           eVAR = oCAM09
408:         End Select



412:         Select Case VarType(eVAR)
        Case vbDate
414:           If IsNumeric(aFOR) Then
415:             eVAR = Fdata(eVAR, "DF")  ''Nao tem matriz formatacao
416:           Else
417:             If aFOR(x) = "" Then
418:               eVAR = Fdata(eVAR, "DF")  ''Nao tem formatacao
419:             End If
420:           End If
        Case vbBoolean
422:           If eVAR Then
423:             eVAR = " * "
424:           Else
425:             eVAR = "   "
426:           End If
        Case vbString
428:           If Len(eVAR) > 100 Then
429:             eVAR = Mid(eVAR, 1, 100)
430:           End If
431:           eVAR = Replace(eVAR, Chr(13) & Chr(10), " ")
432:           If aRETU(2) = "SQLITE" Then
433:              If Mid(eVAR, 5, 1) = "-" And Mid(eVAR, 8, 1) = "-" Then 'datetime no sqlite vem no texto yyyy-mm-dd hh:mm:ss
434:                 eVAR = Mid(eVAR, 9, 2) + "/" + Mid(eVAR, 6, 2) + "/" + Mid(eVAR, 1, 4)
435:              End If
436:           End If
437:         End Select
438:         If Not IsNumeric(aFOR) Then      'E Matriz
439:           If aFOR(x) <> "" Then
440:             eVAR = Format(eVAR, aFOR(x))
441:           End If
442:         End If

444:         cDIZ = cDIZ & eVAR & vbTab
445:       Next

447:       .AddItem cDIZ
448:       oRS.MoveNext

450:     Wend

452:   End With

454:   oGRID.Visible = True
455:   oRS.Close
456:   oDB.Close

458:   Set oRS = Nothing
459:   Set oDB = Nothing

461:   Exit Sub

errhandler:

465:   cERRO = "Montagrid " & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cSQL & Chr(13) & Chr(10)
466:   Select Case Err.Number
  Case 6
468:     Resume Next
  Case 424, 3265                               ''Campo Inexistente
470:     If l3265 Then                            ''So uma vez o erro
471:       cERRO = cERRO & "Campo:" & aCAM(x) & Chr(13) & Chr(10)
472:       SayErro cERRO
473:       l3265 = False
474:     End If
475:     cDIZ = cDIZ & " ? " & vbTab
476:     Resume Next
  Case 3051, 3262
478:     oGRID.Visible = True
479:     Alert "Arquivo em Uso Exclusivo Outro Usuario"
480:     If lOPEN Then
481:       ADOErro oRS.ActiveConnection.Errors, cERRO
482:     Else
483:       ADOErro oDB.Errors, cERRO
484:     End If
485:     Exit Sub
  Case Else
487:     oGRID.Visible = True
488:     If lOPEN Then
489:       ADOErro oRS.ActiveConnection.Errors, cERRO
490:     Else
491:       ADOErro oDB.Errors, cERRO
492:     End If
493:     Exit Sub
494:   End Select

End Sub

Public Sub MontaGridUltra(ByRef oGRID As Variant, _
                          ByVal nITEM As Integer, _
                          ByVal aTAM As Variant, _
                          ByVal aDIZ As Variant, _
                          ByVal aCAM As Variant, _
                          ByVal cARQ As String, _
                          ByVal cSQL As String, _
                          Optional ByVal aFOR As Variant = 0)

507:   On Error GoTo errhandler
  Dim x As Integer
  Dim cDIZ As Variant
  Dim oDB As New ADODB.Connection
  Dim oRS As New ADODB.Recordset
  Dim cARQTMP As String
  Dim l3265 As Boolean
  Dim cERRO As String
  Dim lOPEN As Boolean
  Dim aRETU As Variant
  

519:   lOPEN = False
520:   l3265 = True



524:   If Not FileConnExist(cARQ, True, , cSQL) Then
525:     Exit Sub
526:   End If

528:   aRETU = TipoConn(cARQ, , , False)

530:   cARQTMP = aRETU(1)
  
  
533:   If aRETU(2) = "PGSQL" Then
534:      cSQL = SQLPGSQLDOUBLEQUOTES(cSQL)
535:   End If
  
'  cARQTMP = GeracArq(cARQ, , False)



541:   oDB.ConnectionTimeout = 120
542:   oDB.Open cARQTMP



546:   lOPEN = True

  ' sqlite nao tem adOpenForwardOnly mudado para static ou adOpenUnspecified
549:   If aRETU(2) = "SQLITE" Then
550:     oRS.Open cSQL, oDB, adOpenStatic, adLockReadOnly
551:   Else
552:     oRS.Open cSQL, oDB, adOpenForwardOnly, adLockReadOnly
553:   End If
554:   If Not oRS.BOF Then
555:     oGRID.Visible = False
556:     With oGRID
557:       .FixedCols = 0
558:       .Row = 0
559:       oGRID.Rows = oRS.RecordCount + 1
560:       oGRID.cols = nITEM
561:       For x = 0 To nITEM - 1
562:         .Col = x
563:         .ColWidth(x) = aTAM(x)
564:         If InStr(aCAM(x), "$") > 0 Then
565:           Select Case Mid(aCAM(x), 1, 1)  ' FlexAlignmentCENTERCenter FlexAlignmentRIGHTCenter FlexAlignmentLeftCenter
          Case "L"
567:             .ColAlignment(x) = FlexAlignmentLeftCenter 'flexAlignLeftCenter
          Case "R"
569:             .ColAlignment(x) = FlexAlignmentRightCenter 'flexAlignRightCenter
          Case "C"
571:             .ColAlignment(x) = FlexAlignmentCenterCenter 'flexAlignCenterCenter
572:           End Select
573:           aCAM(x) = Mid(aCAM(x), InStr(aCAM(x), "$") + 1)
574:         End If
575:       Next
576:     End With
577: End If

579: If Not oRS.BOF Then
580:   oRS.MoveFirst

  'define o numero de linhas e colunas e configura o grid
583:   oGRID.Row = 1
584:   oGRID.Col = 0
585:   oGRID.RowSel = oGRID.Rows - 1
586:   oGRID.ColSel = oGRID.cols - 1

  'estamos usando a propriedade Clip e o método GetString para selecionar uma região do grid
589:   oGRID.clip = oRS.GetString(adClipString, -1, Chr(9), Chr(13), vbNullString)
590:   oGRID.Row = 0

592:   With oGRID
593:     For x = 0 To nITEM - 1
594:       .Col = x
595:       .Text = aDIZ(x)
596:     Next x
597:   End With
598:   oGRID.Visible = True

600: End If


603:   oRS.Close
604:   oDB.Close

606:   Set oRS = Nothing
607:   Set oDB = Nothing

609:   Exit Sub

errhandler:

613:   cERRO = "Montagrid " & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cSQL & Chr(13) & Chr(10)
614:   Select Case Err.Number
  Case 6
616:     Resume Next
  Case 424, 3265                               ''Campo Inexistente
618:     If l3265 Then                            ''So uma vez o erro
619:       cERRO = cERRO & "Campo:" & aCAM(x) & Chr(13) & Chr(10)
620:       SayErro cERRO
621:       l3265 = False
622:     End If
623:     cDIZ = cDIZ & " ? " & vbTab
624:     Resume Next
  Case 3051, 3262
626:     oGRID.Visible = True
627:     Alert "Arquivo em Uso Exclusivo Outro Usuario"
628:     If lOPEN Then
629:       ADOErro oRS.ActiveConnection.Errors, cERRO
630:     Else
631:       ADOErro oDB.Errors, cERRO
632:     End If
633:     Exit Sub
  Case 3021
    'eof bof
636:     oGRID.Visible = True
637:     Exit Sub
  Case Else
639:     oGRID.Visible = True
640:     If lOPEN Then
641:       ADOErro oRS.ActiveConnection.Errors, cERRO
642:     Else
643:       ADOErro oDB.Errors, cERRO
644:     End If
645:     Exit Sub
646:   End Select

End Sub

Public Function CloneGrid(ByRef oGRIDORI As Variant, ByRef oGRIDDES As Variant)
  Dim nROWS As Long
  Dim nCOLS As Long
  Dim x As Long
  Dim y As Long
  Dim cLINHA As String
656:   oGRIDORI.Visible = False
657:   oGRIDDES.Visible = False
658:   nROWS = oGRIDORI.Rows - 1
659:   nCOLS = oGRIDORI.cols - 1
660:   oGRIDORI.Row = 0
661:   With oGRIDDES
662:     .cols = oGRIDORI.cols
663:     .FixedCols = 0
664:     .Rows = 1
665:     .Row = 0
666:     For x = 0 To nCOLS
667:       oGRIDORI.Col = x
668:       .Col = x
669:       .Text = oGRIDORI
670:       .ColWidth(x) = oGRIDORI.ColWidth(x)
671:       .ColAlignment(x) = oGRIDORI.ColAlignment(x)
672:     Next
673:     For y = 1 To nROWS                       ''0 cabecario
674:       cLINHA = ""
675:       oGRIDORI.Row = y
676:       For x = 0 To nCOLS
677:         oGRIDORI.Col = x
678:         cLINHA = cLINHA & oGRIDORI & vbTab
679:       Next x
680:       .AddItem cLINHA
681:     Next y
682:   End With
683:   oGRIDORI.Visible = True
684:   oGRIDDES.Visible = True
End Function


