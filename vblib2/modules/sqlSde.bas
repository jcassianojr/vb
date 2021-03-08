Attribute VB_Name = "sqlSde"
Option Explicit

Public Function TipoSDE(ByVal cARQ As String)
    'Dim aRETU As Variant
    '' SDENTX 1
    '' SDEFOX 2
    '' SDENSX 3
    '' SDENSX_DBT 4 (Clipper DBT using NSX indexes
    TipoSDE = Array(cARQ, SDEFOX, ".CDX")
    If InStr(cARQ, "[SDECDX]") > 0 Then
        cARQ = Replace(cARQ, "[SDECDX]", "")
        TipoSDE = Array(cARQ, SDEFOX, ".CDX")
    End If
    If InStr(cARQ, "[SDENTX]") > 0 Then
        cARQ = Replace(cARQ, "[SDENTX]", "")
        TipoSDE = Array(cARQ, SDENTX, ".NTX")    '
    End If
    If InStr(cARQ, "[SDENSX]") > 0 Then
        cARQ = Replace(cARQ, "[SDENSX]", "")
        TipoSDE = Array(cARQ, SDENSX, ".NSX")
    End If
End Function

Public Function PegCountSQLSDE(ByVal cARQ As String, ByVal cSQL As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    cARQ = TiraSDE(cARQ)
    cARQ = GeracArq(cARQ, "JETFOX")
    PegCountSQLSDE = PegCountSQLADO(cARQ, cSQL, cCAMPO, eDEFAULT)
End Function

Public Function GrvSQLSDE(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, ByVal aCAM As Variant, ByVal aVAL As Variant, ByVal aFOR As Variant)
    Dim cARQUSO, cARQCDX, cALIAS, cBUSCA As String
    ''Dim lTEM As Boolean
    Dim lGRAVA As Boolean
    Dim nPOS As Integer
    Dim x As Integer
    Dim nHANDLE As Long
    Dim eVAZIO As Variant
    Dim eVAL As Variant
    Dim aOPE As Variant
    Dim aRETU As Variant
    On Error GoTo errhandler

    cSQL = Replace(cSQL, Chr(39), "")
    aRETU = TipoSDE(cARQ)
    cARQ = aRETU(0)
    'GrvSQLSDE = False
    sx_SetEpoch (Year(Date) - 30)
    sx_SetDateFormat BRITISH
    sx_SetDeleted True
    cARQUSO = cARQ & NomeTableSql(cSQL, ".DBF")
    cARQCDX = cARQ & NomeTableSql(cSQL, aRETU(2))
    cALIAS = NomeTableSql(cSQL)
    nPOS = InStr(cSQL, "=")
    cBUSCA = Mid(cSQL, nPOS + 1)
    If Not FileExist(cARQUSO, True) Then Exit Function
    If Not FileExist(cARQCDX, True) Then Exit Function
    nHANDLE = sx_Use(cARQUSO, cALIAS, READWRITE, aRETU(1))
    sx_SetSoftSeek True
    sx_SetOrder (1)
    sx_GoTop
    sx_Seek (cBUSCA)
    'If sx_Seek(cBUSCA) Then
    If Trim(sx_EvalString(sx_IndexKey())) = Trim(cBUSCA) Then
        If sx_Rlock(sx_RecNo()) Then
            For x = 0 To nITEM - 1               ''Matriz Vb comeca 0
                lGRAVA = True
                eVAL = aVAL(x)                   ''Grava o Valor Default
                ''Evita Gravar Campos nullos
                If IsNull(aVAL(x)) Then
                    lGRAVA = False
                Else
                    aOPE = SepSqlOpe(aCAM(x))
                    If aOPE(0) = "" Or aOPE(1) = "" Or aOPE(2) = "" Then ''Operacao com Campos
                    Else
                        eVAL = MathOper(sx_GetVariant(aOPE(1)), sx_GetVariant(aOPE(2)), aOPE(0))
                    End If
                End If
                If lGRAVA Then
                    eVAL = FVar(eVAL, aFOR(x), eVAZIO)
                    sx_PutVariant aCAM(x), eVAL
                End If
            Next
            sx_Commit
            sx_Unlock sx_RecNo()
            GrvSQLSDE = True
        End If
    End If
    'End If
    sx_Close
    Exit Function
errhandler:
    Select Case Err.Number
    Case Else
        SayErro "GRVSQL SQLSDE" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cSQL
        Exit Function
    End Select
End Function

Public Function IncluiSQLSDE(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long _
                                                                        , ByVal aCAM As Variant, ByVal aVAL As Variant _
                                                                                                , Optional ByVal lCHECK As Boolean = False _
                                                                                                 , Optional ByVal lMES As Boolean = True, Optional ByVal aIDDES As Variant = 0)
    Dim cARQUSO, cARQCDX, cALIAS, cBUSCA As String
    Dim lTEM, lGRAVA As Boolean
    Dim nPOS, x, nCAMPOS As Integer
    Dim nHANDLE As Long
    Dim aRETU As Variant

    On Error GoTo errhandler
    IncluiSQLSDE = False
    cSQL = Replace(cSQL, Chr(39), "")
    aRETU = TipoSDE(cARQ)
    cARQ = aRETU(0)
    sx_SetEpoch (Year(Date) - 30)
    sx_SetDateFormat BRITISH
    sx_SetDeleted True
    cARQUSO = cARQ & NomeTableSql(cSQL, ".DBF")
    cARQCDX = cARQ & NomeTableSql(cSQL, aRETU(2))
    cALIAS = NomeTableSql(cSQL)
    nPOS = InStr(cSQL, "=")
    cBUSCA = Mid(cSQL, nPOS + 1)
    If Not FileExist(cARQUSO, True) Then Exit Function
    If Not FileExist(cARQCDX, True) Then Exit Function
    nHANDLE = sx_Use(cARQUSO, cALIAS, READWRITE, aRETU(1))
    sx_SetDeleted True
    sx_SetSoftSeek True
    sx_SetOrder (1)
    sx_GoTop
    lTEM = False
    If lCHECK Then
        sx_Seek (cBUSCA)                         '//IF Then
        If Trim(sx_EvalString(sx_IndexKey())) = Trim(cBUSCA) Then
            lTEM = True
        End If
        'ENDIF
    End If
    If Not lTEM Then
        sx_AppendBlank
        If sx_Rlock(sx_RecNo()) Then
            For x = 0 To nITEM - 1               ''Matriz Vb comeca 0
                lGRAVA = True
                If IsNull(aVAL(x)) Then          ''Evita Gravar Campos nullos
                    lGRAVA = False
                End If
                If lGRAVA Then
                    sx_PutVariant aCAM(x), aVAL(x)
                End If
            Next
            sx_Commit
            ' PEGAR O ID
            If Not IsNumeric(aIDDES) Then        ''Se for numerico nao e matriz
                nCAMPOS = UBound(aIDDES)
                ReDim aRETUID(nCAMPOS + 1)
                For x = 0 To nCAMPOS
                    aRETUID(x) = sx_GetVariant(aCAM(x))
                Next x
                eRETU01 = aRETUID
            End If
            sx_Unlock sx_RecNo()
        End If
    End If
    sx_Close
    lRETU = Not lTEM
    IncluiSQLSDE = lRETU
    If lCHECK And lTEM Then
        If lMES Then
            Alert ("Item ja Cadastrado Com Esta Chave")
        End If
    End If
    Exit Function
errhandler:
    Select Case Err.Number
    Case Else
        SayErro "Inclui SQLSDE" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cSQL
        Exit Function
    End Select
End Function

Public Function TiraSDE(cARQ)
    TiraSDE = Replace(cARQ, "[SDECDX]", "")
    TiraSDE = Replace(TiraSDE, "[SDENTX]", "")
    TiraSDE = Replace(TiraSDE, "[SDENSX]", "")
End Function

Public Function PegSQLSDE(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, ByVal aCAM As Variant, ByVal aFOR As Variant, ByVal aPAD As Variant) As Variant
    cARQ = TiraSDE(cARQ)
    cARQ = GeracArq(cARQ, "JETFOX")
    PegSQLSDE = PegSQLAdo(cARQ, cSQL, nITEM, aCAM, aFOR, aPAD)
End Function

Public Function PegUltSQLSDE(ByVal cARQ As String, ByVal cSQL As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    cARQ = TiraSDE(cARQ)
    cARQ = GeracArq(cARQ, "JETFOX")
    PegUltSQLSDE = PegUltSQLAdo(cARQ, cSQL, cCAMPO, eDEFAULT)
End Function

Public Function PegMAXSQLSDE(ByVal cARQ As String, ByVal cSQL As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    cARQ = TiraSDE(cARQ)
    cARQ = GeracArq(cARQ, "JETFOX")
    PegMAXSQLSDE = PegMAXSQLADO(cARQ, cSQL, cCAMPO, eDEFAULT)
End Function

Public Function ApagaSQLSDE(ByVal cARQ As String, ByVal cSQL As String) As Boolean
    ApagaSQLSDE = False
    Dim cARQUSO, cARQCDX, cALIAS, cBUSCA As String
    'Dim lTEM As Boolean
    Dim nPOS As Integer
    Dim nHANDLE As Long
    Dim aRETU As Variant
    On Error GoTo errhandler
    cSQL = Replace(cSQL, Chr(39), "")
    aRETU = TipoSDE(cARQ)
    cARQ = aRETU(0)
    ApagaSQLSDE = False
    sx_SetEpoch (Year(Date) - 30)
    sx_SetDateFormat BRITISH
    sx_SetDeleted True
    cARQUSO = cARQ & NomeTableSql(cSQL, ".DBF")
    cARQCDX = cARQ & NomeTableSql(cSQL, aRETU(2))
    cALIAS = NomeTableSql(cSQL)
    nPOS = InStr(cSQL, "=")
    cBUSCA = Mid(cSQL, nPOS + 1)
    If Not FileExist(cARQUSO, True) Then Exit Function
    If Not FileExist(cARQCDX, True) Then Exit Function
    nHANDLE = sx_Use(cARQUSO, cALIAS, READWRITE, aRETU(1))
    sx_SetOrder (1)
    sx_GoTop
    sx_SetSoftSeek True
    'If sx_Seek(cBUSCA) Then
    sx_Seek (cBUSCA)
    If Trim(FixStr(sx_EvalString(sx_IndexKey()))) = Trim(FixStr(cBUSCA)) Then
        ''sx_Rlock (sx_RecNo())
        sx_Delete
        ApagaSQLSDE = True
        ''sx_Unlock sx_RecNo()
    End If
    'End If
    sx_Close
    Exit Function
errhandler:
    Select Case Err.Number
    Case Else
        SayErro "Apaga SQL SDE" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cSQL
        Exit Function
    End Select
End Function

Public Function SomaSQLSDE(ByVal cARQ As String, ByVal cSQL As String, ByVal aCAM As Variant) As Variant
    cARQ = TiraSDE(cARQ)
    cARQ = GeracArq(cARQ, "JETFOX")
    SomaSQLSDE = SomaSQLAdo(cARQ, cSQL, aCAM)
End Function

Public Function PegSQLDeliSDE(ByVal cARQ As String, ByVal cSQL As String, _
                              ByVal aCAM As Variant, Optional ByVal cDELI As String = ",") As Variant
    cARQ = TiraSDE(cARQ)
    cARQ = GeracArq(cARQ, "JETFOX")
    PegSQLDeliSDE = PegSQLDeliAdo(cARQ, cSQL, aCAM, cDELI)
End Function

Public Function ApagaSQLPSDE(ByVal cARQ As String, ByVal cSQL As String, Optional ByVal cTexto As String = "Confirme Exclusão")
    If MDG(cTexto) Then
        ApagaSQLPSDE = ApagaSQLSDE(cARQ, cSQL)
    End If
End Function

Public Function SQLMoveRegSDE(ByVal cARQORI As String, _
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
                    
    cARQORI = TiraSDE(cARQORI)
    cARQDES = TiraSDE(cARQDES)
    cARQORI = GeracArq(cARQORI, "JETFOX")
    cARQDES = GeracArq(cARQDES, "JETFOX")
    SQLMoveRegADO cARQORI, cSQLORI, cOPEORI, aCAMORI, aOUTORI, _
                  cARQDES, cSQLDES, cOPEDES, aCAMDES, aOUTDES, _
                  aIDDES
                 
End Function


