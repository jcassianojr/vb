Attribute VB_Name = "sqlfuncoes"
Option Explicit
Public Function TrocaSQLOrder(ByVal cSQL As String, ByVal cSUBORDER As String) As String
  Dim nPOS As Long
  Dim cSELECT As String
  cSELECT = cSQL
  nPOS = InStr(UCase(cSQL), "ORDER BY")
  If nPOS > 0 Then
    cSELECT = Mid(cSQL, 1, nPOS - 1)
  End If
  TrocaSQLOrder = cSELECT & " Order By " & cSUBORDER
End Function

Public Function TrocaSqlWhere(ByVal cSQL As String, ByVal cSUBWHERE As String) As String
  Dim nPOS As Long
  Dim nPOS2 As Long
  Dim cSELECT As String
  Dim cORDER As String
  cSELECT = ""
  cORDER = ""
  nPOS = InStr(UCase(cSQL), "WHERE")
  nPOS2 = InStr(UCase(cSQL), "ORDER BY")
  If nPOS > 0 Then
    cSELECT = Mid(cSQL, 1, nPOS - 1)
  End If
  If nPOS2 > 0 Then
    cORDER = Mid(cSQL, nPOS2 - 1)
    If nPOS = 0 Then
      cSELECT = Mid(cSQL, 1, nPOS2 - 1)
    End If
  End If
  If nPOS = 0 And nPOS2 = 0 Then
    cSELECT = cSQL
  End If
  If Len(cSUBWHERE) > 0 Then
    TrocaSqlWhere = Trim(cSELECT) & " WHERE " & cSUBWHERE & cORDER
  Else
    TrocaSqlWhere = Trim(cSELECT) & " " & cORDER
  End If
End Function

Public Function NomeTableSql(ByVal cSQL As String, Optional ByVal cEXTENSAO As String = "") As String
  Dim nPOS As Integer
  Dim cNOME As String
  NomeTableSql = ""
  cSQL = UCase(cSQL)
  cSQL = Replace(cSQL, Chr(13), " ")
  cSQL = Replace(cSQL, Chr(10), " ")
  nPOS = InStr(cSQL, "FROM")
  If nPOS > 0 Then
    cNOME = Mid(cSQL, nPOS + 5)
    nPOS = InStr(cNOME, " ")
    If nPOS > 0 Then
      cNOME = Mid(cNOME, 1, nPOS - 1)
    End If
    cNOME = cNOME + cEXTENSAO
    NomeTableSql = cNOME
  Else
    If cEXTENSAO = ".DBF" Then               ''passado no sql o nome da tabela direto exemplo bcofgts
      NomeTableSql = cSQL + cEXTENSAO
    End If
  End If
End Function

Public Function MontaFiltro(ByVal aCAM As Variant, ByVal aFOR As Variant, ByVal eBUSCA As Variant, ByVal nIndex As Integer)
  nIndex = nIndex - 1
  MontaFiltro = ""
  Select Case aFOR(nIndex)
  Case "="
    MontaFiltro = " (" & aCAM(nIndex) & " =" & eBUSCA & ") "
  Case "B"
    MontaFiltro = " (" & aCAM(nIndex) & ") "
  Case "L%"
    MontaFiltro = " (" & aCAM(nIndex) & " LIKE '%" & eBUSCA & "%'" & ") "
  Case "L*"
    MontaFiltro = " (" & aCAM(nIndex) & " LIKE '*" & eBUSCA & "*'" & ") "
  End Select
End Function

Public Function GrvSQL(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, ByVal aCAM As Variant, _
                       ByVal aVAL As Variant, ByVal aFOR As Variant, Optional ByVal nStartItem = 0)
  Dim aRETU As Variant
If IsQueryDestructive(cSQL) Then
     Exit Function
  End If
  aRETU = TipoConn(cARQ)
  'cARQ = aRETU(1)'mantem com as [ as funcoes tratam
  Select Case aRETU(0)
  Case "ADO"
    GrvSQL = GrvSQLado(cARQ, cSQL, nITEM, aCAM, aVAL, aFOR, nStartItem)
  Case "VBSQLITE"
      GrvSQL = GrvSQLite(cARQ, cSQL, nITEM, aCAM, aVAL, aFOR, nStartItem)
  Case "ADORC6"
      GrvSQL = GrvSQLRC6(cARQ, cSQL, nITEM, aCAM, aVAL, aFOR, nStartItem)
  Case "SQLITERC6"
      GrvSQL = GrvSQLITERC6(cARQ, cSQL, nITEM, aCAM, aVAL, aFOR, nStartItem)
  Case "TC6SQLITE"
      GrvSQL = GrvSQLTC6(cARQ, cSQL, nITEM, aCAM, aVAL, aFOR, nStartItem)
 Case "SQLITE4VB"
      GrvSQL = GrvSQLite4vb(cARQ, cSQL, nITEM, aCAM, aVAL, aFOR, nStartItem) ' Nova integração
  End Select
End Function

Public Function IncluiSQL(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, _
                          Optional ByVal aCAM As Variant = 0, Optional ByVal aVAL As Variant = 0, _
                          Optional ByVal lCHECK As Boolean = False, _
                          Optional ByVal lMES As Boolean = True, Optional ByVal aIDDES As Variant = 0)
  Dim aRETU As Variant
  If IsQueryDestructive(cSQL) Then
     Exit Function
  End If
  aRETU = TipoConn(cARQ)
  'cARQ = aRETU(1)'mantem com as [ as funcoes tratam
  Select Case aRETU(0)
  Case "ADO"
    IncluiSQL = IncluiSQLAdo(cARQ, cSQL, nITEM, aCAM, aVAL, lCHECK, lMES, aIDDES)
  Case "VBSQLITE"
       IncluiSQL = IncluiSQLite(cARQ, cSQL, nITEM, aCAM, aVAL, lCHECK, lMES, aIDDES)
  Case "ADORC6"
        IncluiSQL = IncluiSQLRC6(cARQ, cSQL, nITEM, aCAM, aVAL, lCHECK, lMES, aIDDES)
   Case "SQLITERC6"
        IncluiSQL = IncluiSQLITERC6(cARQ, cSQL, nITEM, aCAM, aVAL, lCHECK, lMES, aIDDES)
   Case "TC6SQLITE"
       IncluiSQL = IncluiSQLTC6(cARQ, cSQL, nITEM, aCAM, aVAL, lCHECK, lMES, aIDDES)
   Case "SQLITE4VB"
       IncluiSQL = IncluiSQLite4vb(cARQ, cSQL, nITEM, aCAM, aVAL, lCHECK, lMES, aIDDES) ' Nova integração
  End Select
End Function

Public Function PegSQL(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, ByVal aCAM As Variant, ByVal aFOR As Variant, ByVal aPAD As Variant, Optional ByVal lWRITE As Boolean = False) As Variant
  Dim aRETU As Variant
  If Not IsQuerySafe(cSQL) Then
     Exit Function
  End If
  aRETU = TipoConn(cARQ, , , lWRITE)
  'cARQ = aRETU(1)'mantem com as [ as funcoes tratam
  Select Case aRETU(0)
       Case "ADO"
           PegSQL = PegSQLAdo(cARQ, cSQL, nITEM, aCAM, aFOR, aPAD)
        Case "VBSQLITE"
            PegSQL = PegSQLite(cARQ, cSQL, nITEM, aCAM, aFOR, aPAD)
        Case "ADORC6"
            PegSQL = PegSQLRC6(cARQ, cSQL, nITEM, aCAM, aFOR, aPAD)
            Case "SQLITERC6"
            PegSQL = PegSQLITERC6(cARQ, cSQL, nITEM, aCAM, aFOR, aPAD)
    Case "TC6SQLITE"
            PegSQL = PegSQLTC6(cARQ, cSQL, nITEM, aCAM, aFOR, aPAD)
     Case "SQLITE4VB"
          PegSQL = PegSQLite4vb(cARQ, cSQL, nITEM, aCAM, aFOR, aPAD) ' Nova integração
  End Select
End Function

Public Function PegUltSQL(ByVal cARQ As String, ByVal cSQL As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant, Optional ByVal lWRITE As Boolean = False) As Variant
  Dim aRETU As Variant
   If Not IsQuerySafe(cSQL) Then
     Exit Function
  End If
  aRETU = TipoConn(cARQ, , , lWRITE)
Select Case aRETU(0)
  Case "ADO"
       PegUltSQL = PegUltSQLAdo(cARQ, cSQL, cCAMPO, eDEFAULT)
  Case "VBSQLITE"
         PegUltSQL = PegUltSQLite(cARQ, cSQL, cCAMPO, eDEFAULT)
  Case "ADORC6"
         PegUltSQL = PegUltSQLRC6(cARQ, cSQL, cCAMPO, eDEFAULT)
   Case "SQLITERC6"
         PegUltSQL = PegUltSQLITERC6(cARQ, cSQL, cCAMPO, eDEFAULT)
  Case "TC6SQLITE"
        PegUltSQL = PegUltSQLTC6(cARQ, cSQL, cCAMPO, eDEFAULT)
  Case "SQLITE4VB"
        PegUltSQL = PegUltSQLite4vb(cARQ, cSQL, cCAMPO, eDEFAULT) '
End Select
End Function

Public Function PegMAXSQL(ByVal cARQ As String, ByVal cTable As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant, Optional ByVal lWRITE As Boolean = False) As Variant
  Dim aRETU As Variant
 'If Not IsQuerySafe(cSQL) Then
 '    Exit Function
 ' End If
  aRETU = TipoConn(cARQ, , , lWRITE)
  Select Case aRETU(0)
  Case "ADO"
    PegMAXSQL = PegMAXSQLADO(cARQ, cTable, cCAMPO, eDEFAULT)
    Case "VBSQLITE"
         PegMAXSQL = PegMaxSQLite(cARQ, cTable, cCAMPO, eDEFAULT)
    Case "ADORC6"
         PegMAXSQL = PegMAXSQLRC6(cARQ, cTable, cCAMPO, eDEFAULT)
    Case "SQLITERC6"
         PegMAXSQL = PegMAXSQLITERC6(cARQ, cTable, cCAMPO, eDEFAULT)
    Case "TC6SQLITE"
         PegMAXSQL = PegMAXSQLTC6(cARQ, cTable, cCAMPO, eDEFAULT)
    Case "SQLITE4VB"
        PegMAXSQL = PegMaxSQLite4vb(cARQ, cTable, cCAMPO, eDEFAULT)
  End Select
End Function

Public Function PegCountSQL(ByVal cARQ As String, ByVal cTable As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant, Optional ByVal lWRITE As Boolean = False) As Variant
  Dim aRETU As Variant
' If Not IsQuerySafe(cSQL) Then
'     Exit Function
'  End If
  aRETU = TipoConn(cARQ, , , lWRITE)
  'cARQ = aRETU(1) 'mantem com as [ as funcoes tratam
Select Case aRETU(0)
  Case "ADO"
    PegCountSQL = PegCountSQLADO(cARQ, cTable, cCAMPO, eDEFAULT)
  Case "VBSQLITE"
     PegCountSQL = PegCountSQLite(cARQ, cTable, cCAMPO, eDEFAULT)
  Case "ADORC6"
        PegCountSQL = PegCountSQLRC6(cARQ, cTable, cCAMPO, eDEFAULT)
   Case "SQLITERC6"
        PegCountSQL = PegCountSQLITERC6(cARQ, cTable, cCAMPO, eDEFAULT)
 Case "TC6SQLITE"
       PegCountSQL = PegCountSQLTC6(cARQ, cTable, cCAMPO, eDEFAULT)
 Case "SQLITE4VB"
      PegCountSQL = PegCountSQLite4vb(cARQ, cTable, cCAMPO, eDEFAULT) ' Nova integração
End Select
End Function
Public Function ComandoSQL(ByVal cARQ As String, ByVal cSQL As String) As Boolean
  Dim aRETU As Variant
  aRETU = TipoConn(cARQ)
  'cARQ = aRETU(1) 'mantem com as [ as funcoes tratam
  If InStr(UCase(cARQ), "VFPOLEDB") Then
      ComandoSQL = AdoComandodbf(cARQ, "", cSQL)
  End If
Select Case aRETU(0)
  Case "ADO"
       ComandoSQL = ADOComando(cARQ, cSQL)
  Case "VBSQLITE"
       ComandoSQL = ComandoSQLite(cARQ, cSQL)
   Case "ADORC6"
        ComandoSQL = ComandoRC6(cARQ, cSQL)
    Case "SQLITERC6"
        ComandoSQL = ComandoSQLITERC6(cARQ, cSQL)
 Case "TC6SQLITE"
     ComandoSQL = ComandoTC6(cARQ, cSQL)
 Case "SQLITE4VB"
    ComandoSQL = ComandoSQLite4vb(cARQ, cSQL) ' Nova integração
End Select
End Function

Public Function ApagaSQL(ByVal cARQ As String, ByVal cSQL As String) As Boolean
  Dim aRETU As Variant
  If IsQueryDestructive(cSQL) Then
     Exit Function
  End If
  aRETU = TipoConn(cARQ)
  'cARQ = aRETU(1) 'mantem com as [ as funcoes tratam
Select Case aRETU(0)
  Case "ADO"
       ApagaSQL = APAGASQLADO(cARQ, cSQL)
  Case "VBSQLITE"
       ApagaSQL = ApagaSQLite(cARQ, cSQL)
   Case "ADORC6"
        ApagaSQL = APAGASQLRC6(cARQ, cSQL)
  Case "SQLITERC6"
        ApagaSQL = APAGASQLITERC6(cARQ, cSQL)
  Case "TC6SQLITE"
      ApagaSQL = APAGASQLTC6(cARQ, cSQL)
  Case "SQLITE4VB"
     ApagaSQL = ApagaSQLite4vb(cARQ, cSQL)
 End Select
End Function

Public Function SomaSQL(ByVal cARQ As String, ByVal cSQL As String, ByVal aCAM As Variant, Optional ByVal lWRITE As Boolean = False) As Variant
  Dim aRETU As Variant
 If Not IsQuerySafe(cSQL) Then
     Exit Function
  End If
  aRETU = TipoConn(cARQ, , , lWRITE)
  Select Case aRETU(0)
      Case "ADO"
         SomaSQL = SomaSQLAdo(cARQ, cSQL, aCAM)
      Case "VBSQLITE"
         SomaSQL = SomaSQLite(cARQ, cSQL, aCAM)
      Case "ADORC6"
            SomaSQL = SomaSQLRC6(cARQ, cSQL, aCAM)
      Case "SQLITERC6"
            SomaSQL = SomaSQLITERC6(cARQ, cSQL, aCAM)
            
      Case "TC6SQLITE"
            SomaSQL = SomaSQLTC6(cARQ, cSQL, aCAM)
  End Select
End Function
Public Function PegSQLDeli(ByVal cARQ As String, ByVal cSQL As String, _
                           ByVal aCAM As Variant, Optional ByVal cDELI As String = ",", _
                           Optional ByVal aPAD As Variant = "", Optional ByVal aFOR As Variant = "") As Variant
    
    Dim aRETU As Variant
    If Not IsQuerySafe(cSQL) Then
     Exit Function
  End If
  aRETU = TipoConn(cARQ)
  Select Case aRETU(0)
        Case "ADO"
             PegSQLDeli = PegSQLDeliAdo(cARQ, cSQL, aCAM, cDELI, aPAD, aFOR)
        Case "VBSQLITE"
            PegSQLDeli = PegSQLiteDeli(cARQ, cSQL, aCAM, cDELI, aPAD, aFOR)
        Case "ADORC6"
            PegSQLDeli = PegSQLDeliRC6(cARQ, cSQL, aCAM, cDELI, aPAD, aFOR)
        Case "SQLITERC6"
            PegSQLDeli = PegSQLDeliSQLITERC6(cARQ, cSQL, aCAM, cDELI, aPAD, aFOR)
        Case "TC6SQLITE"
            PegSQLDeli = PegSQLDeliTC6(cARQ, cSQL, aCAM, cDELI, aPAD, aFOR)
    End Select
End Function

Public Function ApagaSQLP(ByVal cARQ As String, ByVal cSQL As String, Optional ByVal cTEXTO As String = "Confirme Exclusão")
   If MDG(cTEXTO, "Exclusão Registro") Then
      ApagaSQLP = ApagaSQL(cARQ, cSQL)
   End If
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

 If IsQueryDestructive(cSQLORI) Then
     Exit Function
  End If

 If IsQueryDestructive(cSQLDES) Then
     Exit Function
  End If

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

  
  
  aRETU = TipoConn(cARQORI)
  ''cARQ = aRETU(1)
  Select Case aRETU(0)
  Case "ADO"
    SqlMoveReg = SQLMoveRegADO(cARQORI, cSQLORI, cOPEORI, aCAMORI, aOUTORI, _
                  cARQDES, cSQLDES, cOPEDES, aCAMDES, aOUTDES, _
                  aIDDES)
   Case "VBSQLITE"
            SqlMoveReg = SQLMoveRegSQLite(cARQORI, cSQLORI, cOPEORI, aCAMORI, aOUTORI, _
                                          cARQDES, cSQLDES, cOPEDES, aCAMDES, aOUTDES, aIDDES)
        Case "ADORC6"
            SqlMoveReg = SQLMoveRegRC6(cARQORI, cSQLORI, cOPEORI, aCAMORI, aOUTORI, _
                                             cARQDES, cSQLDES, cOPEDES, aCAMDES, aOUTDES, aIDDES)
         Case "SQLITERC6"
            SqlMoveReg = SQLMoveRegSQLITERC6(cARQORI, cSQLORI, cOPEORI, aCAMORI, aOUTORI, _
                                             cARQDES, cSQLDES, cOPEDES, aCAMDES, aOUTDES, aIDDES)
      Case "TC6SQLITE"
       SqlMoveReg = SQLMoveRegTC6(cARQORI, cSQLORI, cOPEORI, aCAMORI, aOUTORI, _
                                             cARQDES, cSQLDES, cOPEDES, aCAMDES, aOUTDES, aIDDES)
  End Select
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
  lSEGUNDO = False
  If Len(aRELCFG(8)) > 0 Then
    lSEGUNDO = True
    cSQLSEG = Replace("SELECT * FROM MANFEC WHERE ARQORI='" & Mid(aRELCFG(8), InStrRev(aRELCFG(8), "}") + 1) & "'", ".DBF", "")
    aSEGUNDO = PegSQL(GeraConn(PegPath("PATH", "MANA5CON"), "JETFOX"), cSQLSEG, 4, Array("ARQORI", "STRBAI", "STRDES", "STRANO"), _
                      Array("CU", "CU", "CU", "CU"), Array("", "", "", ""))
  End If
  Mana5Fec = False
  If Len(aARQFEC(0)) > 0 Or Len(aARQFEC(1)) > 0 Or Len(aARQFEC(2)) > 0 Or Len(aARQFEC(3)) > 0 Then
    cDIZBX = ""
    aARQFEC(0) = Trim(aARQFEC(0))
    aARQFEC(1) = Trim(aARQFEC(1))
    aARQFEC(2) = Trim(aARQFEC(2))
    aARQFEC(3) = Trim(aARQFEC(3))
    If Len(aARQFEC(0)) > 0 Then
      cDIZBX = cDIZBX & "(A)tual"
    End If
    If Len(aARQFEC(1)) > 0 Then
      cDIZBX = cDIZBX & "(B)aixado"
    End If
    If Len(aARQFEC(2)) > 0 Then
      cDIZBX = cDIZBX & "(F)echado"
    End If
    If Len(aARQFEC(3)) > 0 Then
      cDIZBX = cDIZBX & "a(C)umulado"
    End If
    cCAMFEC = UCase(Busca(cDIZBX, "Utilizar", "A", 1))
    If Len(aARQFEC(1)) = 0 And cCAMFEC = "B" Then
      Alert "Não Possue Baixa"
      Exit Function
    End If
    If Len(aARQFEC(2)) = 0 And cCAMFEC = "F" Then
      Alert "Não Possue Fechamento"
      Exit Function
    End If
    If Len(aARQFEC(3)) = 0 And cCAMFEC = "C" Then
      Alert "Não Possue Acumulado"
      Exit Function
    End If
    If cCAMFEC = "F" Then
      cAno = Busca("Digite o Ano AAAA", "Competencia Ano", CStr(Format(Date, "yyyy")), 4)
      cMes = Busca("Digite o Mes MM", "Competencia Mes", "01", 2)
      cARQUSO = PegCamini("{MANA5}") & "E" & cAno & "\" & aARQFEC(2) & Right(cAno, 2) & cMes & ".DBF"
      If lSEGUNDO Then
        cARQUS2 = PegCamini("{MANA5}") & "E" & cAno & "\" & Trim(aSEGUNDO(2)) & Right(cAno, 2) & cMes & ".DBF"
      End If
      cARQNOME = aARQFEC(2)
      cFECDIZ = Right(cAno, 2) & cMes & "-" & cMes & "/" & cAno
    Else
      Select Case cCAMFEC
      Case "A"
        cARQUSO = PegCamini("{MANA5EMP}") & aARQFEC(0) & ".DBF"
        If lSEGUNDO Then
          cARQUS2 = PegCamini("{MANA5EMP}") & Trim(aSEGUNDO(0)) & ".DBF"
        End If
        cARQNOME = aARQFEC(0)
        cFECDIZ = "Atual"
      Case "B"
        cARQUSO = PegCamini("{MANA5EMP}") & aARQFEC(1) & ".DBF"
        If lSEGUNDO Then
          cARQUS2 = PegCamini("{MANA5EMP}") & Trim(aSEGUNDO(1)) & ".DBF"
        End If
        cARQNOME = aARQFEC(1)
        cFECDIZ = "Baixado"
      Case "C"
        cARQUSO = PegCamini("{MANA5EMP}") & aARQFEC(3) & ".DBF"
        If lSEGUNDO Then
          cARQUS2 = PegCamini("{MANA5EMP}") & Trim(aSEGUNDO(3)) & ".DBF"
        End If
        cARQNOME = aARQFEC(3)
        cFECDIZ = "Acumulado"
      End Select
    End If
    If cCAMFEC = "C" Then
      If Not MDG("Usar Acumulo Anterior", "Confirmaçao") Then
        ''Shell "somaano $" & UCase(aARQFEC(0)), vbNormalFocus
        ShellEx "SOMAANO", essSW_SHOWDEFAULT, "$" & UCase(aARQFEC(0)), PegPath("PATH", "SOMAANO")
        Alert "Clique em Ok, Após o Termino do Acumulo", "Aguarde.."
      End If
    End If
    aARQFEC(4) = cFECDIZ
    aARQFEC(5) = cARQNOME
    aARQFEC(6) = cARQUSO
    aARQUIVOS(0) = aARQFEC(6)
    If lSEGUNDO Then
      aARQUIVOS(1) = cARQUS2
    End If
    aRELCFG(6) = aARQFEC(8)                  'Retorna Mensagem Original
    If InStr(aRELCFG(6), "[ZFEC]") > 0 Then
      aRELCFG(6) = Replace(aRELCFG(6), "[ZFEC]", aARQFEC(4))
    End If
    Mana5Fec = True
  End If
End Function


Public Function PegOperSQL(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant, ByVal coper As String) As Variant
  Dim aRETU As Variant
 'If Not IsQuerySafe(cSQL) Then
 '    Exit Function
 ' End If
  aRETU = TipoConn(cARQ)
  Select Case aRETU(0)
  Case "ADO"
    PegOperSQL = PegOperSQLADO(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, coper)
 Case "ADORC6"
      PegOperSQL = PegOperSQLRC6(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, coper)
  Case "SQLITERC6"
      PegOperSQL = PegOperSQLITERC6(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, coper)
  Case "VBSQLITE"
      PegOperSQL = PegOperSQLite(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, coper)
  Case "TC6SQLITE"
        PegOperSQL = PegOperSQLTC6(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, coper)
  End Select

End Function
Public Function IsQuerySafe(ByVal sql As String) As Boolean
    Dim cleanSQL As String
    Dim i As Integer
    Dim charCode As Integer
    Dim currentChar As String
    Dim forbiddenWords As Variant
    
    ' 1. Normaliza para maiúsculas
    cleanSQL = UCase(sql)
    
    ' 2. Substitui caracteres não-alfabéticos por espaços
    ' Isso separa as palavras que o atacante tentou colar
    For i = 1 To Len(cleanSQL)
        currentChar = Mid(cleanSQL, i, 1)
        charCode = Asc(currentChar)
        ' Mantém A-Z e números
        If Not ((charCode >= 65 And charCode <= 90) Or (charCode >= 48 And charCode <= 57)) Then
            Mid(cleanSQL, i, 1) = " "
        End If
    Next i
    
    ' 3. Remove excesso de espaços para facilitar a análise
    cleanSQL = Trim(Replace(Replace(cleanSQL, "  ", " "), "  ", " "))
    
    ' 4. Verifica se a primeira palavra é SELECT
    ' Dividimos para pegar apenas a primeira parte
    If Split(cleanSQL, " ")(0) <> "SELECT" Then
        IsQuerySafe = False
        Exit Function
    End If
    
    ' 5. Lista de palavras proibidas (sem espaços nas pontas)
    forbiddenWords = Array("DELETE", "DROP", "UPDATE", "INSERT", _
                           "CREATE", "ALTER", "TRUNCATE", "EXEC", _
                           "EXECUTE", "GRANT", "REVOKE", "INTO")
    
    ' 6. Verifica se alguma palavra proibida existe na string processada
    ' Como limpamos a string no passo 2, o InStr encontrará a palavra
    ' mesmo que ela estivesse grudada originalmente
    For i = LBound(forbiddenWords) To UBound(forbiddenWords)
        If InStr(cleanSQL, forbiddenWords(i)) > 0 Then
            IsQuerySafe = False
            RegistrarLogSeguranca (sql)
            Exit Function
        End If
    Next i
    
    IsQuerySafe = True
End Function

Public Function IsQueryDestructive(ByVal sql As String) As Boolean
    Dim cleanSQL As String
    Dim i As Integer, charCode As Integer
    Dim destrutivos As Variant
    
    ' 1. Normaliza para maiúsculas
    cleanSQL = UCase(sql)
    
    ' 2. Limpeza profunda: substitui tudo que não é letra ou número por espaço
    ' Isso garante que "DELETE;" ou "DROP--" virem "DELETE " ou "DROP "
    For i = 1 To Len(cleanSQL)
        charCode = Asc(Mid(cleanSQL, i, 1))
        If Not ((charCode >= 65 And charCode <= 90) Or (charCode >= 48 And charCode <= 57)) Then
            Mid(cleanSQL, i, 1) = " "
        End If
    Next i
    
    ' 3. Remove espaços extras para deixar a string "limpa"
    cleanSQL = Trim(Replace(Replace(cleanSQL, "  ", " "), "  ", " "))
    
    ' 4. Lista do que é proibido
    ' Como removemos os símbolos, basta buscar a palavra pura
    destrutivos = Array("DROP", "ALTER", "CREATE", "TRUNCATE", "GRANT", "REVOKE")
    
    ' 5. Verifica se alguma palavra proibida existe dentro da string limpa
    ' O InStr encontrará "DROP" mesmo se estiver no meio da frase ou sem espaços
    For i = LBound(destrutivos) To UBound(destrutivos)
        If InStr(cleanSQL, destrutivos(i)) > 0 Then
            IsQueryDestructive = True ' É destrutivo!
            RegistrarLogSeguranca (sql)
            Exit Function
        End If
    Next i
    
    IsQueryDestructive = False ' É seguro para prosseguir
End Function

Public Sub RegistrarLogSeguranca(ByVal cSQL As String)
    Dim fso As Object
    Dim ts As Object
    Dim cLogPath As String
    
    cLogPath = App.Path & "\log_seguranca.txt"
    
    ' Cria o objeto FSO
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' O parâmetro 8 indica "ForAppending" (Adicionar ao final)
    ' O True indica que ele deve criar o arquivo caso ele não exista
    Set ts = fso.OpenTextFile(cLogPath, 8, True)
    
    ' Escreve as linhas de log
    ts.WriteLine "DATA: " & Now
    ts.WriteLine "SQL: " & cSQL
    ts.WriteLine "--------------------------------------------------------"
    
    ' Fecha e limpa
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
End Sub
