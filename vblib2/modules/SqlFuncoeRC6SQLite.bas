Attribute VB_Name = "SqlRC6SQLite"
Option Explicit
'vbrichclient RC6 https://www.vbrichclient.com/en/Downloads.htm

Public Function ComandoSQLITERC6(ByVal cARQ As String, ByVal cSQL As String) As Boolean
    Dim oDB As New RC6.cConnection

    On Error GoTo TrataErro
    ComandoSQLITERC6 = False
    
    ' Ajustes de dialeto (Mantido conforme original)
    If InStr(cSQL, "CURRENTDATETIME") > 0 Then
          cSQL = Replace(cSQL, "CURRENTDATETIME", " current_timestamp ")
    End If
    
    cARQ = LimpaTag(cARQ)
    
    ' Inicialização simplificada e com tratamento de erro
    If Not oDB.OpenDB(cARQ) Then Err.Raise vbObjectError + 1000, "ComandoSQLITERC6", _
        "Nao foi possivel abrir o banco: " & cARQ
    
    ' O método Execute da conexão substitui a necessidade do objeto Command
    oDB.Execute cSQL
    
    ' Fechamento simplificado (apenas limpa o objeto)
    Set oDB = Nothing
    
    ComandoSQLITERC6 = True
    Exit Function

TrataErro:
    SayErro "SQL RC6 Comando:" & vbCrLf & cARQ & vbCrLf & cSQL & vbCrLf & Err.Description
    If Not oDB Is Nothing Then Set oDB = Nothing
    ComandoSQLITERC6 = False
End Function

Public Function APAGASQLITERC6(ByVal cARQ As String, ByVal cSQL As String) As Boolean
  Dim nPOS As Integer
  Dim cNOME As String

  APAGASQLITERC6 = False
  'Muda para maiscula para o instr usar em maiscula
  cNOME = UCase(cSQL)
  cNOME = Replace(cNOME, Chr(13), " ")
  cNOME = Replace(cNOME, Chr(10), " ")

  If InStr(cNOME, "WHERE") = 0 Then
    Alert ("where nao preenchido apagasql")
    Exit Function
  End If
  If InStr(cNOME, "FROM") = 0 Then
    Alert ("from nao preenchido apagasql")
    Exit Function
  End If

  nPOS = InStr(cNOME, "FROM")
  'efetua substituicao se estiver com select * from ou outros casos --> delete from usando mid
  If nPOS > 0 Then
    cSQL = "DELETE FROM " & Mid(cSQL, nPOS + 5)
    APAGASQLITERC6 = ComandoSQLITERC6(cARQ, cSQL)
  Else
    If InStr(cNOME, "DELETE") > 0 Then  'so executa se tiver delete
      APAGASQLITERC6 = ComandoSQLITERC6(cARQ, cSQL)
    End If
  End If
End Function

Public Function SomaSQLITERC6(ByVal cARQ As String, ByVal cSQL As String, ByVal aCAM As Variant) As Variant
    Dim oDB As New RC6.cConnection
    Dim oRS As RC6.cRecordset
    Dim lOPEN As Boolean
    Dim lRSOP As Boolean
    Dim x As Integer
    Dim nCAMPOS As Integer
    Dim aRETU As Variant
    Dim aOPE As Variant
    Dim eVAL As Variant

    On Error GoTo errhandler
    lRETU = False

    lOPEN = False
    lRSOP = False
    nCAMPOS = UBound(aCAM) + 1
    ReDim aRETU(nCAMPOS - 1) ' Ajustado para base 0
    For x = 0 To nCAMPOS - 1
        aRETU(x) = 0
    Next x

    If Not FileConnExist(SomaExt(cARQ), True) Then
        SomaSQLITERC6 = aRETU
        Exit Function
    End If

    cARQ = LimpaTag(cARQ)
    If Not oDB.OpenDB(cARQ) Then Err.Raise vbObjectError + 1000, "SomaSQLITERC6", _
        "Nao foi possivel abrir o banco: " & cARQ
    
    lOPEN = True
  
    If Not SqliteSetValuesRC6(oDB) Then Err.Raise vbObjectError + 1003, _
        "SomaSQLITERC6", "Falha ao configurar a conexao SQLite."

    ' Abertura do Recordset no RC6
    Set oRS = oDB.OpenRecordset(cSQL)
    lRSOP = True

    If Not oRS.EOF Then
        While Not oRS.EOF
            For x = 0 To nCAMPOS - 1
                aOPE = SepSqlOpe(aCAM(x))
                If aOPE(0) = "" Or aOPE(1) = "" Or aOPE(2) = "" Then
                    eVAL = oRS(aCAM(x)).Value
                Else
                    eVAL = MathOper(oRS(aOPE(1)).Value, oRS(aOPE(2)).Value, aOPE(0))
                End If
                If IsNull(eVAL) Or IsEmpty(eVAL) Then eVAL = 0
                aRETU(x) = aRETU(x) + eVAL
            Next x
            oRS.MoveNext
        Wend
    End If
    
    Set oRS = Nothing
    Set oDB = Nothing
    
    SomaSQLITERC6 = aRETU
    Exit Function

errhandler:
    SayErro "Erro SomaSQLITERC6:" & vbCrLf & cARQ & vbCrLf & cSQL & vbCrLf & Err.Description
    If Not oRS Is Nothing Then Set oRS = Nothing
    If Not oDB Is Nothing Then Set oDB = Nothing
    SomaSQLITERC6 = aRETU
End Function

Public Function PegSQLDeliSQLITERC6(ByVal cARQ As String, ByVal cSQL As String, _
                              ByVal aCAM As Variant, Optional ByVal cDELI As String = ",", _
                              Optional ByVal aPAD As Variant = "", Optional ByVal aFOR As Variant = "") As Variant

    Dim oDB As New RC6.cConnection
    Dim oRS As RC6.cRecordset
    Dim x As Integer
    Dim nCAMPOS As Integer
    Dim aRETU As Variant
    Dim aOPE As Variant
    Dim eVAL As Variant
    Dim lOPEN As Boolean
    Dim lRSOP As Boolean

    On Error GoTo errhandler

    lOPEN = False
    lRSOP = False
    nCAMPOS = UBound(aCAM) + 1
    ReDim aRETU(nCAMPOS - 1)
    For x = 0 To nCAMPOS - 1
        aRETU(x) = ""
    Next x

    If Not FileConnExist(SomaExt(cARQ), True) Then
        PegSQLDeliSQLITERC6 = aRETU
        Exit Function
    End If

    cARQ = LimpaTag(cARQ)
    If Not oDB.OpenDB(cARQ) Then Err.Raise vbObjectError + 1000, "PegSQLDeliSQLITERC6", _
        "Nao foi possivel abrir o banco: " & cARQ
    
    lOPEN = True

    If Not SqliteSetValuesRC6(oDB) Then Err.Raise vbObjectError + 1003, _
        "PegSQLDeliSQLITERC6", "Falha ao configurar a conexao SQLite."

    Set oRS = oDB.OpenRecordset(cSQL)
    lRSOP = True

    If Not oRS.EOF Then
        While Not oRS.EOF
            For x = 0 To nCAMPOS - 1
                aOPE = SepSqlOpe(aCAM(x))
                
                If aOPE(0) = "" Or aOPE(1) = "" Or aOPE(2) = "" Then
                    eVAL = oRS(aCAM(x)).Value
                Else
                    eVAL = MathOper(oRS(aOPE(1)).Value, oRS(aOPE(2)).Value, aOPE(0))
                End If
                
                If IsNull(eVAL) Then
                    If IsArray(aPAD) Then eVAL = aPAD(x) Else eVAL = aPAD
                End If
                
                If IsArray(aFOR) Then
                    If IsArray(aPAD) Then eVAL = FVar(eVAL, aFOR(x), aPAD(x)) Else eVAL = FVar(eVAL, aFOR(x))
                End If
                
                aRETU(x) = aRETU(x) & FixStr(eVAL)
            Next x
            
            oRS.MoveNext
            
            If Not oRS.EOF Then
                For x = 0 To nCAMPOS - 1
                    aRETU(x) = aRETU(x) & cDELI
                Next x
            End If
        Wend
    End If

    Set oRS = Nothing
    Set oDB = Nothing
    
    PegSQLDeliSQLITERC6 = aRETU
    Exit Function

errhandler:
    SayErro "Peg SQL DELI RC6:" & vbCrLf & cARQ & vbCrLf & cSQL & vbCrLf & Err.Description
    If Not oRS Is Nothing Then Set oRS = Nothing
    If Not oDB Is Nothing Then Set oDB = Nothing
    PegSQLDeliSQLITERC6 = aRETU
End Function

Public Function GrvSQLITERC6(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, ByVal aCAM As Variant, _
                          ByVal aVAL As Variant, ByVal aFOR As Variant, Optional ByVal nSTARITEM As Long = 0) As Boolean
    Dim oDB As New RC6.cConnection
    Dim oRS As RC6.cRecordset
    Dim x As Long
    Dim eVAL As Variant
    Dim aOPE As Variant
    
    On Error GoTo errhandler
    GrvSQLITERC6 = False
    
    cARQ = LimpaTag(cARQ)
    If Not oDB.OpenDB(cARQ) Then Err.Raise vbObjectError + 1000, "GrvSQLITERC6", _
        "Nao foi possivel abrir o banco: " & cARQ
    
    If Not SqliteSetValuesRC6(oDB) Then Err.Raise vbObjectError + 1003, _
        "GrvSQLITERC6", "Falha ao configurar a conexao SQLite."

    Set oRS = oDB.OpenRecordset(cSQL)

    If oRS.EOF Then
        Set oRS = Nothing: Set oDB = Nothing
        Exit Function
    End If
    If Not oRS.Updatable Then Err.Raise vbObjectError + 1001, "GrvSQLITERC6", _
        "O recordset nao e atualizavel."

    While Not oRS.EOF
        For x = nSTARITEM To nITEM - 1
            If IsNull(aVAL(x)) Then
                oRS(aCAM(x)).Value = Null
            Else
                aOPE = SepSqlOpe(aCAM(x))
                If aOPE(0) = "" Then
                    eVAL = aVAL(x)
                Else
                    eVAL = MathOper(oRS(aOPE(1)).Value, oRS(aOPE(2)).Value, aOPE(0))
                End If

                If Mid$(CStr(aFOR(x)), 1, 1) = "D" Then
                    If IsDate(eVAL) Then
                        oRS(aCAM(x)).Value = CDate(eVAL)
                    Else
                        oRS(aCAM(x)).Value = Null
                    End If
                Else
                    oRS(aCAM(x)).Value = FVar(eVAL, aFOR(x), "")
                End If
            End If
        Next x
        
        oRS.UpdateBatch
        oRS.MoveNext
    Wend

    Set oRS = Nothing
    Set oDB = Nothing
    GrvSQLITERC6 = True
    Exit Function

errhandler:
    SayErro "Erro na Gravacao (RC6):" & vbCrLf & cARQ & vbCrLf & cSQL & vbCrLf & Err.Description
    If Not oRS Is Nothing Then Set oRS = Nothing
    If Not oDB Is Nothing Then Set oDB = Nothing
    GrvSQLITERC6 = False
End Function

Public Function ApagaSQLpRC6(ByVal cARQ As String, ByVal cSQL As String, Optional ByVal cTEXTO As String = "Confirme Exclusão") As Boolean
  ApagaSQLpRC6 = False
  If MDG(cTEXTO, "Exclusão Registro") Then
    ApagaSQLpRC6 = APAGASQLITERC6(cARQ, cSQL)
  End If
End Function

Public Function IncluiSQLITERC6(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, _
                             Optional ByVal aCAM As Variant = 0, Optional ByVal aVAL As Variant = 0, _
                             Optional ByVal lCHECK As Boolean = False, _
                             Optional ByVal lMES As Boolean = True, _
                             Optional ByVal aIDDES As Variant = 0) As Boolean
    Dim oDB As New RC6.cConnection
    Dim oRS As RC6.cRecordset
    Dim x As Long
    Dim aRETU As Variant
    Dim lTEM As Boolean
    
    IncluiSQLITERC6 = False
    On Error GoTo errhandler

    lTEM = False
    lRETU = False
    
    cARQ = LimpaTag(cARQ)
    If Not oDB.OpenDB(cARQ) Then Err.Raise vbObjectError + 1000, "IncluiSQLITERC6", _
        "Nao foi possivel abrir o banco: " & cARQ
    
    If Not SqliteSetValuesRC6(oDB) Then Err.Raise vbObjectError + 1003, _
        "IncluiSQLITERC6", "Falha ao configurar a conexao SQLite."
    
    Set oRS = oDB.OpenRecordset(cSQL)

    If lCHECK And Not oRS.EOF Then
        lTEM = True
    End If

    If Not lTEM Then
        If Not oRS.Updatable Then Err.Raise vbObjectError + 1001, "IncluiSQLITERC6", _
            "O recordset nao e atualizavel."
        oRS.AddNew
        For x = 0 To nITEM - 1
            If Not IsNull(aVAL(x)) Then
                If Mid(aCAM(x), 1, 1) = "D" Then 
                    If IsDate(aVAL(x)) Then
                        oRS(aCAM(x)).Value = CDate(aVAL(x))
                    Else
                        oRS(aCAM(x)).Value = Null
                    End If
                Else
                    oRS(aCAM(x)).Value = aVAL(x)
                End If
            End If
        Next x

        oRS.UpdateBatch
        
        If IsArray(aIDDES) Then
            Dim nCAMPOS As Long
            nCAMPOS = UBound(aIDDES)
            ReDim eRETU01(nCAMPOS)
            For x = 0 To nCAMPOS
                eRETU01(x) = oRS(aIDDES(x)).Value
            Next x
        Else
            eRETU01 = oDB.LastInsertAutoID
        End If
        
        lRETU = True
    End If

    Set oRS = Nothing
    Set oDB = Nothing

    If lCHECK And lTEM And lMES Then
        Alert "Item ja Cadastrado Com Esta Chave"
    End If

    IncluiSQLITERC6 = lRETU
    Exit Function

errhandler:
    SayErro "Erro na Inclusao (RC6):" & vbCrLf & cARQ & vbCrLf & cSQL & vbCrLf & Err.Description
    If Not oRS Is Nothing Then Set oRS = Nothing
    If Not oDB Is Nothing Then Set oDB = Nothing
    IncluiSQLITERC6 = False
End Function

Public Function PegSQLITERC6(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, ByVal aCAM As Variant, ByVal aFOR As Variant, ByVal aPAD As Variant) As Variant
    Dim oDB As New RC6.cConnection
    Dim oRS As RC6.cRecordset
    Dim x As Long
    Dim aRETU As Variant
    Dim aOPE As Variant
    Dim eVAL As Variant

    On Error GoTo errhandler
    lRETU = False
    
    If nITEM <= 0 Then
        PegSQLITERC6 = aPAD
        Exit Function
    End If
    ReDim aRETU(0 To nITEM - 1)
    For x = 0 To nITEM - 1
        aRETU(x) = ValorPadraoRC6(aPAD, x)
    Next x
    
    cARQ = LimpaTag(cARQ)
    If Not oDB.OpenDB(cARQ) Then Err.Raise vbObjectError + 1000, "PegSQLITERC6", _
        "Nao foi possivel abrir o banco: " & cARQ

    If Not SqliteSetValuesRC6(oDB) Then Err.Raise vbObjectError + 1003, _
        "PegSQLITERC6", "Falha ao configurar a conexao SQLite."

    Set oRS = oDB.OpenRecordset(cSQL)

    If Not oRS.EOF Then
        lRETU = True
        For x = 0 To nITEM - 1
            aOPE = SepSqlOpe(aCAM(x))
            
            If aOPE(0) = "" Or aOPE(1) = "" Or aOPE(2) = "" Then
                eVAL = oRS(aCAM(x)).Value
            Else
                eVAL = MathOper(oRS(aOPE(1)).Value, oRS(aOPE(2)).Value, aOPE(0))
            End If

            If IsNull(eVAL) Then
                aRETU(x) = ValorPadraoRC6(aPAD, x)
            Else
                aRETU(x) = FVar(eVAL, aFOR(x), ValorPadraoRC6(aPAD, x))
            End If
        Next x
    End If

    Set oRS = Nothing
    Set oDB = Nothing
    
    PegSQLITERC6 = aRETU
    Exit Function

errhandler:
    Dim cERRO As String
    cERRO = "Peg SQL RC6:" & vbCrLf & cARQ & vbCrLf & cSQL & vbCrLf & "Erro: " & Err.Description
    
    SayErro cERRO
    If Not oRS Is Nothing Then Set oRS = Nothing
    If Not oDB Is Nothing Then Set oDB = Nothing
    PegSQLITERC6 = aRETU
End Function

Public Function PegCountSQLITERC6(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
  PegCountSQLITERC6 = PegOperSQLITERC6(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, "COUNT")
End Function

Public Function PegMINSQLITERC6(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
  PegMINSQLITERC6 = PegOperSQLITERC6(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, "MIN")
End Function

Public Function PegMAXSQLITERC6(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
  PegMAXSQLITERC6 = PegOperSQLITERC6(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, "MAX")
End Function

Public Function PegSUMSQLITERC6(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
  PegSUMSQLITERC6 = PegOperSQLITERC6(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, "SUM")
End Function

Public Function PegCampoSQLITERC6(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
  PegCampoSQLITERC6 = PegOperSQLITERC6(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, "CAMPO")
End Function

Public Function PegOperSQLITERC6(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant, ByVal coper As String) As Variant
  Dim aRETU As Variant
  Dim cSQL As String
  PegOperSQLITERC6 = eDEFAULT
  lRETU = False

  If coper = "CAMPO" Then
    If Len(cCAMPO) = 0 Then
      cTABLEWHERE = UCase(cTABLEWHERE)
      cSQL = Replace(cTABLEWHERE, " FROM ", " AS CAMPO FROM ")
    Else
      cSQL = "SELECT " & cCAMPO & " AS CAMPO FROM " & cTABLEWHERE
    End If
  Else
    cSQL = "SELECT " & coper & "(" & cCAMPO & ") AS CAMPO FROM " & cTABLEWHERE
  End If
  If coper = "SUM" Or coper = "COUNT" Or coper = "MAX" Or coper = "MIN" Or IsNumeric(eDEFAULT) Then
    aRETU = PegSQLITERC6(cARQ, cSQL, 1, Array("CAMPO"), Array("N"), Array(eDEFAULT))
  Else
    aRETU = PegSQLITERC6(cARQ, cSQL, 1, Array("CAMPO"), Array(""), Array(eDEFAULT))
  End If
  If lRETU Then
    PegOperSQLITERC6 = aRETU(0)
  End If
End Function

Public Function PegUltSQLITERC6(ByVal cARQ As String, ByVal cSQL As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    Dim oDB As New RC6.cConnection
    Dim oRS As RC6.cRecordset
    
    On Error GoTo errhandler
    
    PegUltSQLITERC6 = eDEFAULT
    
    cARQ = LimpaTag(cARQ)
    If Not oDB.OpenDB(cARQ) Then Err.Raise vbObjectError + 1000, "PegUltSQLITERC6", _
        "Nao foi possivel abrir o banco: " & cARQ
    
    Set oRS = oDB.OpenRecordset(cSQL)
    
    If Not oRS.EOF Then
        oRS.MoveLast
        If Not IsNull(oRS(cCAMPO).Value) Then
            PegUltSQLITERC6 = oRS(cCAMPO).Value
        End If
    End If
    
    Set oRS = Nothing
    Set oDB = Nothing
    
    Exit Function

errhandler:
    SayErro "PegUltSQLITERC6 - Erro: " & Err.Description & vbCrLf & _
            "Arquivo: " & cARQ & vbCrLf & _
            "SQL: " & cSQL
            
    If Not oRS Is Nothing Then Set oRS = Nothing
    If Not oDB Is Nothing Then Set oDB = Nothing
    
    PegUltSQLITERC6 = eDEFAULT
End Function

Public Function SQLMoveRegSQLITERC6(ByVal cARQORI As String, _
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

    Dim oDB As New RC6.cConnection
    Dim oDBDES As New RC6.cConnection
    Dim oRS As RC6.cRecordset
    Dim oRSDES As RC6.cRecordset
    Dim x As Integer
    Dim nCAMPOS As Integer
    Dim aVALORI As Variant
    Dim aRETUID As Variant

    On Error GoTo errhandler
    SQLMoveRegSQLITERC6 = False

    cARQORI = LimpaTag(cARQORI)
    cARQDES = LimpaTag(cARQDES)
    If Not oDB.OpenDB(cARQORI) Then Err.Raise vbObjectError + 1000, "SQLMoveRegSQLITERC6", _
        "Nao foi possivel abrir o banco de origem: " & cARQORI
    If Not oDBDES.OpenDB(cARQDES) Then Err.Raise vbObjectError + 1000, "SQLMoveRegSQLITERC6", _
        "Nao foi possivel abrir o banco de destino: " & cARQDES

    If Not SqliteSetValuesRC6(oDB) Then Err.Raise vbObjectError + 1003, _
        "SQLMoveRegSQLITERC6", "Falha ao configurar a conexao de origem."
    If Not SqliteSetValuesRC6(oDBDES) Then Err.Raise vbObjectError + 1003, _
        "SQLMoveRegSQLITERC6", "Falha ao configurar a conexao de destino."

    Set oRS = oDB.OpenRecordset(cSQLORI)
    Set oRSDES = oDBDES.OpenRecordset(cSQLDES)
    
    If Not oRS.Updatable And InStr(1, UCase$(cOPEORI), "DEL", vbTextCompare) > 0 Then _
        Err.Raise vbObjectError + 1001, "SQLMoveRegSQLITERC6", "O recordset de origem nao e atualizavel."
    If InStr(1, UCase$(cOPEDES), "INC", vbTextCompare) = 0 Then _
        Err.Raise vbObjectError + 1002, "SQLMoveRegSQLITERC6", "A operacao de destino deve conter INC."
    If Not oRSDES.Updatable Then _
        Err.Raise vbObjectError + 1001, "SQLMoveRegSQLITERC6", "O recordset de destino nao e atualizavel."

    If Not oRS.EOF Then
        While Not oRS.EOF
            If Not IsNumeric(aCAMORI) Then
                nCAMPOS = UBound(aCAMORI)
                ReDim aVALORI(nCAMPOS)
                For x = 0 To nCAMPOS
                    aVALORI(x) = oRS(aCAMORI(x)).Value
                Next x
            End If

            If InStr(cOPEORI, "DEL") > 0 Then oRS.Delete

            If InStr(cOPEDES, "INC") > 0 Then
                oRSDES.AddNew
            End If

            If IsArray(aCAMDES) And IsArray(aCAMORI) Then
                For x = 0 To UBound(aCAMDES)
                    oRSDES(aCAMDES(x)).Value = aVALORI(x)
                Next x
            End If
            
            If IsArray(aOUTDES) Then
                For x = 0 To UBound(aOUTDES)
                    oRSDES(aOUTDES(x)).Value = ValorArrayOuEscalarRC6(aOUTORI, x)
                Next x
            End If

            oRSDES.UpdateBatch

            If IsArray(aIDDES) Then
                nCAMPOS = UBound(aIDDES)
                ReDim aRETUID(nCAMPOS)
                For x = 0 To nCAMPOS
                    aRETUID(x) = oRSDES(aIDDES(x)).Value
                Next x
                eRETU01 = aRETUID
            Else
                eRETU01 = oDBDES.LastInsertAutoID
            End If

            oRS.MoveNext
        Wend
    End If

    If InStr(1, UCase$(cOPEORI), "DEL", vbTextCompare) > 0 Then
        oRS.UpdateBatch
    End If

    Set oRS = Nothing
    Set oRSDES = Nothing
    Set oDB = Nothing
    Set oDBDES = Nothing

    SQLMoveRegSQLITERC6 = True
    Exit Function

errhandler:
    SayErro "Erro em SQLMoveRegRC6:" & vbCrLf & Err.Description
    If Not oRS Is Nothing Then Set oRS = Nothing
    If Not oRSDES Is Nothing Then Set oRSDES = Nothing
    Set oDB = Nothing: Set oDBDES = Nothing
    SQLMoveRegSQLITERC6 = False
End Function

Public Function SqliteSetValuesRC6(ByVal oCON As RC6.cConnection) As Boolean
    On Error GoTo ErroSQLite
    SqliteSetValuesRC6 = False
    
    If oCON Is Nothing Then Exit Function
    
    oCON.Execute "PRAGMA temp_store = MEMORY;"
    oCON.Execute "PRAGMA cache_size = 2000;"
    oCON.Execute "PRAGMA journal_mode = WAL;"
    oCON.Execute "PRAGMA synchronous = NORMAL;"
    
    SqliteSetValuesRC6 = True
    Exit Function

ErroSQLite:
    SqliteSetValuesRC6 = False
End Function

' --- Funcoes auxiliares herdadas da versao aprimorada ---
Private Function ValorPadraoRC6(ByVal aPAD As Variant, ByVal index As Long) As Variant
    If IsArray(aPAD) Then
        ValorPadraoRC6 = aPAD(index)
    Else
        ValorPadraoRC6 = aPAD
    End If
End Function

Private Function ValorArrayOuEscalarRC6(ByVal values As Variant, ByVal index As Long) As Variant
    If IsArray(values) Then
        ValorArrayOuEscalarRC6 = values(index)
    Else
        ValorArrayOuEscalarRC6 = values
    End If
End Function