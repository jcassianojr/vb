Attribute VB_Name = "sqlfuncoes"
Option Explicit

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
      GrvSQL = GrvSQLite4vb(cARQ, cSQL, nITEM, aCAM, aVAL, aFOR, nStartItem) ' Nova integraï¿½ï¿½o
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
       IncluiSQL = IncluiSQLite4vb(cARQ, cSQL, nITEM, aCAM, aVAL, lCHECK, lMES, aIDDES) ' Nova integraï¿½ï¿½o
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
          PegSQL = PegSQLite4vb(cARQ, cSQL, nITEM, aCAM, aFOR, aPAD) ' Nova integraï¿½ï¿½o
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
      PegCountSQL = PegCountSQLite4vb(cARQ, cTable, cCAMPO, eDEFAULT) ' Nova integraï¿½ï¿½o
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
    ComandoSQL = ComandoSQLite4vb(cARQ, cSQL) ' Nova integraï¿½ï¿½o
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
      Case "SQLITE4VB"
            SomaSQL = SomaSQLite4vb(cARQ, cSQL, aCAM)
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
        Case "SQLITE4VB"
            PegSQLDeli = PegSQLiteDeli4vb(cARQ, cSQL, aCAM, cDELI, aPAD, aFOR)
    End Select
End Function

Public Function ApagaSQLP(ByVal cARQ As String, ByVal cSQL As String, Optional ByVal cTEXTO As String = "Confirme Exclusï¿½o")
   If MDG(cTEXTO, "Exclusï¿½o Registro") Then
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
      Case "SQLITE4VB"
       SqlMoveReg = SQLMoveRegSQLite4vb(cARQORI, cSQLORI, cOPEORI, aCAMORI, aOUTORI, _
                                             cARQDES, cSQLDES, cOPEDES, aCAMDES, aOUTDES, aIDDES)
  End Select
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
