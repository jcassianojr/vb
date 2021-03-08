Attribute VB_Name = "sqlSde"
Public Function PegMAXSQLSDE(ByVal cARQ As String, ByVal cSQL As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    cARQ = GeraConn(cARQ, "JETFOX")
    PegMAXSQLSDE = PegMAXSQLADO(cARQ, cSQL, cCAMPO, eDEFAULT)
End Function

Public Function PegCountSQLSDE(ByVal cARQ As String, ByVal cSQL As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    cARQ = GeraConn(cARQ, "JETFOX")
    PegCountSQLSDE = PegCountSQLADO(cARQ, cSQL, cCAMPO, eDEFAULT)
End Function

Public Function GrvSQLsde(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, ByVal aCAM As Variant, ByVal aVAL As Variant, ByVal aFOR As Variant)
    cARQ = GeraConn(cARQ, "JETFOX")
    GrvSQLsde = GrvSQLado(cARQ, cSQL, nITEM, aCAM, aVAL, aFOR)
End Function

Public Function IncluiSQLSDE(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long _
                                                                        , ByVal aCAM As Variant, ByVal aVAL As Variant _
                                                                                                , Optional ByVal lCHECK As Boolean = False _
                                                                                                 , Optional ByVal lMES As Boolean = True, Optional ByVal aIDDES As Variant = 0)
    cARQ = GeraConn(cARQ, "JETFOX")
    IncluiSQLSDE = IncluiSQLAdo(cARQ, cSQL, nITEM, aCAM, aVAL, lCHECK, lMES)
End Function

Public Function PegSQLSDE(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, ByVal aCAM As Variant, ByVal aFOR As Variant, ByVal aPAD As Variant) As Variant
    cARQ = GeraConn(cARQ, "JETFOX")
    PegSQLSDE = PegSQLAdo(cARQ, cSQL, nITEM, aCAM, aFOR, aPAD)
End Function

Public Function PegUltSQLSDE(ByVal cARQ As String, ByVal cSQL As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    cARQ = GeraConn(cARQ, "JETFOX")
    PegUltSQLSDE = PegUltSQLAdo(cARQ, cSQL, cCAMPO, eDEFAULT)
End Function

Public Function ApagaSQLSDE(ByVal cARQ As String, ByVal cSQL As String) As Boolean
    cARQ = GeraConn(cARQ, "JETFOX")
    ApagaSQLSDE = APAGASQLADO(cARQ, cSQL)
End Function

Public Function somaSQLSDE(ByVal cARQ As String, ByVal cSQL As String, ByVal aCAM As Variant) As Variant
    cARQ = GeraConn(cARQ, "JETFOX")
    somaSQLSDE = SomaSQLAdo(cARQ, cSQL, aCAM)
End Function

Public Function pegSQLdeliSDE(ByVal cARQ As String, ByVal cSQL As String, _
                              ByVal aCAM As Variant, Optional ByVal cDELI As String = ",") As Variant
    cARQ = GeraConn(cARQ, "JETFOX")
    pegSQLdeliSDE = PegSQLDeliAdo(cARQ, cSQL, aCAM, cDELI)
End Function

Public Function ApagaSQLPSDE(ByVal cARQ As String, ByVal cSQL As String, Optional ByVal cTexto As String = "Confirme Exclusão")
    cARQ = GeraConn(cARQ, "JETFOX")
    ApagaSQLPSDE = ApagaSQLpAdo(cARQ, cSQL, cTexto)
End Function

Public Function sqlmoveregSDE(ByVal cARQORI As String, _
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
                    
    cARQ = GeraConn(cARQ, "JETFOX")
    SQLMoveRegADO cARQORI, _
                  cSQLORI, _
                  cOPEORI, _
                  aCAMORI, _
                  aOUTORI, _
                  cARQDES, _
                  cSQLDES, _
                  cOPEDES, _
                  aCAMDES, _
                  aOUTDES, _
                  aIDDES
                 
End Function


