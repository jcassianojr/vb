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
  aRETU = TipoConn(cARQ)
  cARQ = aRETU(1)
  Select Case aRETU(0)
  Case "ADO"
    GrvSQL = GrvSQLado(cARQ, cSQL, nITEM, aCAM, aVAL, aFOR, nStartItem)
  End Select
End Function

Public Function IncluiSQL(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, _
                          Optional ByVal aCAM As Variant = 0, Optional ByVal aVAL As Variant = 0, _
                          Optional ByVal lCHECK As Boolean = False, _
                          Optional ByVal lMES As Boolean = True, Optional ByVal aIDDES As Variant = 0)
  Dim aRETU As Variant
  aRETU = TipoConn(cARQ)
  cARQ = aRETU(1)
  Select Case aRETU(0)
  Case "ADO"
    IncluiSQL = IncluiSQLAdo(cARQ, cSQL, nITEM, aCAM, aVAL, lCHECK, lMES, aIDDES)
  End Select
End Function

Public Function PegSQL(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, ByVal aCAM As Variant, ByVal aFOR As Variant, ByVal aPAD As Variant, Optional ByVal lWRITE As Boolean = False) As Variant
  Dim aRETU As Variant
  aRETU = TipoConn(cARQ, , , lWRITE)
  cARQ = aRETU(1)
  Select Case aRETU(0)
  Case "ADO"
    PegSQL = PegSQLAdo(cARQ, cSQL, nITEM, aCAM, aFOR, aPAD)
  End Select
End Function

Public Function PegUltSQL(ByVal cARQ As String, ByVal cSQL As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant, Optional ByVal lWRITE As Boolean = False) As Variant
  Dim aRETU As Variant
  aRETU = TipoConn(cARQ, , , lWRITE)
  cARQ = aRETU(1)
  Select Case aRETU(0)
  Case "ADO"
    PegUltSQL = PegUltSQLAdo(cARQ, cSQL, cCAMPO, eDEFAULT)
  End Select
End Function

Public Function PegMAXSQL(ByVal cARQ As String, ByVal cTable As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant, Optional ByVal lWRITE As Boolean = False) As Variant
  Dim aRETU As Variant
  aRETU = TipoConn(cARQ, , , lWRITE)
  cARQ = aRETU(1)
  Select Case aRETU(0)
  Case "ADO"
    PegMAXSQL = PegMAXSQLADO(cARQ, cTable, cCAMPO, eDEFAULT)
  End Select
End Function

Public Function PegCountSQL(ByVal cARQ As String, ByVal cTable As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant, Optional ByVal lWRITE As Boolean = False) As Variant
  Dim aRETU As Variant
  aRETU = TipoConn(cARQ, , , lWRITE)
  cARQ = aRETU(1)
  Select Case aRETU(0)
  Case "ADO"
    PegCountSQL = PegCountSQLADO(cARQ, cTable, cCAMPO, eDEFAULT)
  End Select
End Function

Public Function ApagaSQL(ByVal cARQ As String, ByVal cSQL As String) As Boolean
  Dim aRETU As Variant
  aRETU = TipoConn(cARQ)
  cARQ = aRETU(1)
  Select Case aRETU(0)
  Case "ADO"
    ApagaSQL = APAGASQLADO(cARQ, cSQL)
  End Select
End Function

Public Function SomaSQL(ByVal cARQ As String, ByVal cSQL As String, ByVal aCAM As Variant, Optional ByVal lWRITE As Boolean = False) As Variant
  Dim aRETU As Variant
  aRETU = TipoConn(cARQ, , , lWRITE)
  Select Case aRETU(0)
  Case "ADO"
    SomaSQL = SomaSQLAdo(cARQ, cSQL, aCAM)
  End Select
End Function

Public Function PegSQLDeli(ByVal cARQ As String, ByVal cSQL As String, _
                           ByVal aCAM As Variant, Optional ByVal cDELI As String = ",", Optional ByVal lWRITE As Boolean = False) As Variant
  Dim aRETU As Variant
  aRETU = TipoConn(cARQ, , , lWRITE)
  Select Case aRETU(0)
  Case "ADO"
    PegSQLDeli = PegSQLDeliAdo(cARQ, cSQL, aCAM, cDELI)
  End Select
End Function

Public Function ApagaSQLP(ByVal cARQ As String, ByVal cSQL As String, Optional ByVal cTEXTO As String = "Confirme Exclusão")
  Dim aRETU As Variant
  aRETU = TipoConn(cARQ)
  cARQ = aRETU(1)
  Select Case aRETU(0)
  Case "ADO"
    ApagaSQLP = ApagaSQLpAdo(cARQ, cSQL, cTEXTO)
  End Select
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
    SQLMoveRegADO cARQORI, cSQLORI, cOPEORI, aCAMORI, aOUTORI, _
                  cARQDES, cSQLDES, cOPEDES, aCAMDES, aOUTDES, _
                  aIDDES
  End Select
End Function

Public Function Mana5Fec() As Boolean
  Dim cDIZBX As String
  Dim cARQUSO As String
  Dim cARQNOME As String
  Dim cMES As String
  Dim cANO As String
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
      cANO = Busca("Digite o Ano AAAA", "Competencia Ano", CStr(Format(Date, "yyyy")), 4)
      cMES = Busca("Digite o Mes MM", "Competencia Mes", "01", 2)
      cARQUSO = PegCamini("{MANA5}") & "E" & cANO & "\" & aARQFEC(2) & Right(cANO, 2) & cMES & ".DBF"
      If lSEGUNDO Then
        cARQUS2 = PegCamini("{MANA5}") & "E" & cANO & "\" & Trim(aSEGUNDO(2)) & Right(cANO, 2) & cMES & ".DBF"
      End If
      cARQNOME = aARQFEC(2)
      cFECDIZ = Right(cANO, 2) & cMES & "-" & cMES & "/" & cANO
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


