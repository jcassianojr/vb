Attribute VB_Name = "sqlfuncoes"
Option Explicit

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

Public Function GrvSQL(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, ByVal aCAM As Variant, ByVal aVAL As Variant, ByVal aFOR As Variant)
    Dim aRETU As Variant
    aRETU = TipoConn(cARQ)
    cARQ = aRETU(1)
    Select Case aRETU(0)
    Case "ADO"
        GrvSQL = GrvSQLado(cARQ, cSQL, nITEM, aCAM, aVAL, aFOR)
    Case "SDE"
        GrvSQL = GrvSQLSDE(cARQ, cSQL, nITEM, aCAM, aVAL, aFOR)
        'Case Else
        '   GrvSQL = GrvSQLdao(cARQ, cSQL, nITEM, aCAM, aVAL, aFOR)
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
    Case "SDE"
        IncluiSQL = IncluiSQLSDE(cARQ, cSQL, nITEM, aCAM, aVAL, lCHECK, lMES, aIDDES)
        ' Case Else
        '   IncluiSQL = IncluiSQLDao(cARQ, cSQL, nITEM, aCAM, aVAL, lCHECK, lMES, aIDDES)
    End Select
End Function

Public Function PegSQL(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, ByVal aCAM As Variant, ByVal aFOR As Variant, ByVal aPAD As Variant, Optional ByVal lWRITE As Boolean = False) As Variant
    Dim aRETU As Variant
    aRETU = TipoConn(cARQ, , , lWRITE)
    cARQ = aRETU(1)
    Select Case aRETU(0)
    Case "ADO"
        PegSQL = PegSQLAdo(cARQ, cSQL, nITEM, aCAM, aFOR, aPAD)
    Case "SDE"
        PegSQL = PegSQLSDE(cARQ, cSQL, nITEM, aCAM, aFOR, aPAD)
        'Case Else
        '     PegSQL = PegSQLDao(cARQ, cSQL, nITEM, aCAM, aFOR, aPAD, "DAOMDB")
    End Select
End Function

Public Function PegUltSQL(ByVal cARQ As String, ByVal cSQL As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant, Optional ByVal lWRITE As Boolean = False) As Variant
    Dim aRETU As Variant
    aRETU = TipoConn(cARQ, , , lWRITE)
    cARQ = aRETU(1)
    Select Case aRETU(0)
    Case "ADO"
        PegUltSQL = PegUltSQLAdo(cARQ, cSQL, cCAMPO, eDEFAULT)
    Case "SDE"
        PegUltSQL = PegUltSQLSDE(cARQ, cSQL, cCAMPO, eDEFAULT)
        'Case Else
        '    PegUltSQL = PegUltSQLDao(cARQ, cSQL, cCAMPO, eDEFAULT)
    End Select
End Function

Public Function PegMAXSQL(ByVal cARQ As String, ByVal cTABLE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant, Optional ByVal lWRITE As Boolean = False) As Variant
    Dim aRETU As Variant
    aRETU = TipoConn(cARQ, , , lWRITE)
    cARQ = aRETU(1)
    Select Case aRETU(0)
    Case "ADO"
        PegMAXSQL = PegMAXSQLADO(cARQ, cTABLE, cCAMPO, eDEFAULT)
    Case "SDE"
        PegMAXSQL = PegMAXSQLSDE(cARQ, cTABLE, cCAMPO, eDEFAULT)
        'Case Else
        'PegMAXSQL = PegMAXSQLDAO(cARQ, cTABLE, cCAMPO, eDEFAULT)
    End Select
End Function

Public Function PegCountSQL(ByVal cARQ As String, ByVal cTABLE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant, Optional ByVal lWRITE As Boolean = False) As Variant
    Dim aRETU As Variant
    aRETU = TipoConn(cARQ, , , lWRITE)
    cARQ = aRETU(1)
    Select Case aRETU(0)
    Case "ADO"
        PegCountSQL = PegCountSQLADO(cARQ, cTABLE, cCAMPO, eDEFAULT)
    Case "SDE"
        PegCountSQL = PegCountSQLSDE(cARQ, cTABLE, cCAMPO, eDEFAULT)
        'Case Else
        'PegCountSQL = PegCountSQLDAO(cARQ, cTABLE, cCAMPO, eDEFAULT)
    End Select
End Function

Public Function ApagaSQL(ByVal cARQ As String, ByVal cSQL As String) As Boolean
    Dim aRETU As Variant
    aRETU = TipoConn(cARQ)
    cARQ = aRETU(1)
    Select Case aRETU(0)
    Case "ADO"
        ApagaSQL = APAGASQLADO(cARQ, cSQL)
    Case "SDE"
        ApagaSQL = ApagaSQLSDE(cARQ, cSQL)
        'Case Else
        '   ApagaSQL = ApagaSQLDao(cARQ, cSQL)
    End Select
End Function

Public Function SomaSQL(ByVal cARQ As String, ByVal cSQL As String, ByVal aCAM As Variant, Optional ByVal lWRITE As Boolean = False) As Variant
    Dim aRETU As Variant
    aRETU = TipoConn(cARQ, , , lWRITE)
    ''cARQ = aRETU(1)
    Select Case aRETU(0)
    Case "ADO"
        SomaSQL = SomaSQLAdo(cARQ, cSQL, aCAM)
    Case "SDE"
        SomaSQL = SomaSQLSDE(cARQ, cSQL, aCAM)
        'Case Else
        '   SomaSQL = SomaSQLDao(cARQ, cSQL, aCAM)
    End Select
End Function

Public Function PegSQLDeli(ByVal cARQ As String, ByVal cSQL As String, _
                           ByVal aCAM As Variant, Optional ByVal cDELI As String = ",", Optional ByVal lWRITE As Boolean = False) As Variant
    Dim aRETU As Variant
    aRETU = TipoConn(cARQ, , , lWRITE)
    ''cARQ = aRETU(1)
    Select Case aRETU(0)
    Case "ADO"
        PegSQLDeli = PegSQLDeliAdo(cARQ, cSQL, aCAM, cDELI)
    Case "SDE"
        PegSQLDeli = PegSQLDeliSDE(cARQ, cSQL, aCAM, cDELI)
        'Case Else
        '   PegSQLDeli = PegSQLDeliDao(cARQ, cSQL, aCAM, cDELI)
    End Select
End Function

Public Function ApagaSQLP(ByVal cARQ As String, ByVal cSQL As String, Optional ByVal cTexto As String = "Confirme Exclusão")
    Dim aRETU As Variant
    aRETU = TipoConn(cARQ)
    cARQ = aRETU(1)
    Select Case aRETU(0)
    Case "ADO"
        ApagaSQLP = ApagaSQLpAdo(cARQ, cSQL, cTexto)
    Case "SDE"
        ApagaSQLP = ApagaSQLPSDE(cARQ, cSQL, cTexto)
        'Case Else
        '   ApagaSQLP = ApagaSQLPDao(cARQ, cSQL, cTexto)
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
    Case "SDE"
        SQLMoveRegSDE cARQORI, cSQLORI, cOPEORI, aCAMORI, aOUTORI, _
                      cARQDES, cSQLDES, cOPEDES, aCAMDES, aOUTDES, _
                      aIDDES
                 
    Case Else
        'SQLMoveRegDAO cARQORI, cSQLORI, cOPEORI, aCAMORI, aOUTORI, _
        '              cARQDES, cSQLDES, cOPEDES, aCAMDES, aOUTDES, _
        '              aIDDES
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


