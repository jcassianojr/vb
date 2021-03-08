Attribute VB_Name = "SqlFuncoesAdo"
Option Explicit

Public Function AdoZeradbf(ByVal cARQ As String, ByVal cTABLE As String) As Boolean
    Dim cCOM As String
    Dim oCON As ADODB.Connection
    Dim oCOMANDO As ADODB.Command
    On Error GoTo trataerro
    AdoZeradbf = False
    cCOM = ""
    Set oCON = New ADODB.Connection
    Set oCOMANDO = New ADODB.Command
    oCON.ConnectionTimeout = 120
    oCON.Open cARQ
    oCOMANDO.CommandType = adCmdText
    cCOM = "use " & cTABLE
    oCOMANDO.CommandText = cCOM
    oCOMANDO.ActiveConnection = oCON
    oCOMANDO.Execute
    cCOM = "DELETE FROM " & cTABLE & " WHERE 1=1"
    oCOMANDO.CommandText = cCOM
    oCOMANDO.Execute
    cCOM = "PACK"
    oCOMANDO.CommandText = cCOM
    oCOMANDO.Execute
    oCON.Close
    AdoZeradbf = True
    Exit Function
trataerro:
    Select Case Err.Number
    Case Else
        SayErro "ADO Zera DBF:" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cTABLE & Chr(13) & Chr(10) & cCOM & Chr(13) & Chr(10)
        Exit Function
    End Select
End Function

Public Function TemTabelaADO(ByVal cARQ As String, ByVal cTabela As String, Optional ByVal lMES As Boolean = True) As Boolean
    Dim oCat As ADOX.Catalog
    Dim oTabela As ADOX.Table
    On Error GoTo trataerro

    TemTabelaADO = False
    Set oCat = New ADOX.Catalog
    oCat.ActiveConnection = cARQ

    For Each oTabela In oCat.Tables
        If UCase(oTabela.Name) = UCase(cTabela) Then
            TemTabelaADO = True
            Exit For
        End If
    Next

    If lMES And Not TemTabelaADO Then
        Alert ("Tabela nao Encontrada" & cTabela & Chr(13) & Chr(10) & cARQ)
    End If
trataerro:
    Select Case Err.Number
    Case Else
        SayErro "Tem Tabela Ado :" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cTabela & Chr(13) & Chr(10)
        Exit Function
    End Select

End Function

Public Function ADOComando(ByVal cARQ As String, ByVal cSQL As String) As Boolean
    Dim aRETU As Variant
    Dim oDB As ADODB.Connection
    Dim oCOM As ADODB.Command
    Dim cCONN As String
    On Error GoTo trataerro
    ADOComando = False
    cARQ = GeraConn(cARQ)
    aRETU = TipoConn(cARQ)                       ''Gera string conneccao
    If aRETU(0) = "ADO" Then
        cCONN = aRETU(1)
        Set oDB = New ADODB.Connection
        oDB.ConnectionTimeout = 120
        oDB.Open cCONN
        Set oCOM = New ADODB.Command
        oCOM.ActiveConnection = oDB
        oCOM.CommandText = cSQL
        oCOM.Execute
        ADOComando = True
        oDB.Close
        Set oCOM = Nothing
        Set oDB = Nothing
    End If
    Exit Function
trataerro:
    Select Case Err.Number
    Case Else
        SayErro "SQL ADO Comando:" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cSQL
        Exit Function
    End Select
End Function

Public Function AdoNewTable(ByVal cARQORI As String, Optional ByVal lCRIA As Boolean = False, _
                            Optional ByVal Ntipo As Integer = 5) As Boolean
    Dim cat As New ADOX.Catalog
    On Error GoTo trataerro
    AdoNewTable = False
    If Not FileExist(cARQORI, True) Then
        If lCRIA Or MDG("Criar Arquivo " & cARQORI) Then
            cat.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & cARQORI & _
                       ";Jet OLEDB:Engine Type=" & Ntipo & ";"
            AdoNewTable = True
        End If
    End If
    Exit Function
trataerro:
    Select Case Err.Number
    Case Else
        SayErro "ADO Novo Arquivo Access/MDB:" & Chr(13) & Chr(10) & cARQORI
        Exit Function
    End Select
End Function

Public Function APAGASQLADO(ByVal cARQ As String, ByVal cSQL As String) As Boolean
    Dim nPOS As Integer
    Dim cNOME As String
    APAGASQLADO = False
    cNOME = UCase(cSQL)
    cNOME = Replace(cNOME, Chr(13), " ")
    cNOME = Replace(cNOME, Chr(10), " ")
    nPOS = InStr(cNOME, "FROM")
    If nPOS > 0 Then
        cSQL = "DELETE FROM " & Mid(cSQL, nPOS + 5)
        APAGASQLADO = ADOComando(cARQ, cSQL)
    Else
        If InStr(cNOME, "DELETE") Then
            APAGASQLADO = ADOComando(cARQ, cSQL)
        End If
    End If
End Function

Public Function SomaSQLAdo(ByVal cARQ As String, ByVal cSQL As String, ByVal aCAM As Variant) As Variant

    Dim oDB              As ADODB.Connection
    Dim oRS              As ADODB.Recordset
    Dim lOPEN As Boolean
    Dim lRSOP As Boolean
    Dim X, nCAMPOS As Integer
    Dim aRETU, aOPE, eVAL As Variant
    Dim cERRO As String
    
    On Error GoTo errhandler
    
    lOPEN = False
    lRSOP = False
    nCAMPOS = UBound(aCAM) + 1
    ReDim aRETU(nCAMPOS)
    For X = 0 To nCAMPOS - 1
        aRETU(X) = 0
    Next X
    
    If Not FileExist(SomaExt(cARQ), True) Then
        SomaSQLAdo = aRETU
        Exit Function
    End If
    
    
    cARQ = GeracArq(cARQ, , False)

    Set oDB = New ADODB.Connection
    oDB.CursorLocation = adUseClient
    oDB.ConnectionTimeout = 120
    oDB.Open cARQ
    lOPEN = True
    Set oRS = New ADODB.Recordset
    oRS.Open cSQL, oDB, adOpenForwardOnly, adLockReadOnly
    lRSOP = True
    
    If Not oRS.EOF Then
        oRS.MoveFirst
        While Not oRS.EOF
            For X = 0 To nCAMPOS - 1
                aOPE = SepSqlOpe(aCAM(X))
                If aOPE(0) = "" Or aOPE(1) = "" Or aOPE(2) = "" Then ''Sem Operacoes
                    eVAL = oRS(aCAM(X))
                Else
                    eVAL = MathOper(oRS(aOPE(1)), oRS(aOPE(2)), aOPE(0))
                End If
                aRETU(X) = aRETU(X) + eVAL
            Next X
            oRS.MoveNext
        Wend
    End If
    oRS.Close
    oDB.Close
    Set oRS = Nothing
    Set oDB = Nothing
    SomaSQLAdo = aRETU
    Exit Function
    
errhandler:
    cERRO = "Soma SQL ADO :" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cSQL & Chr(13) & Chr(10)
    If lRSOP Then
        cERRO = cERRO & ADORsStatus(oRS.Status) & Chr(13) & Chr(10)
    End If
    Select Case Err.Number
    Case Else
        If lOPEN Then
            ADOErro oRS.ActiveConnection.Errors, cERRO
        Else
            ADOErro oDB.Errors, cERRO
        End If
        Exit Function
    End Select
End Function

Public Function PegSQLDeliAdo(ByVal cARQ As String, ByVal cSQL As String, _
                              ByVal aCAM As Variant, Optional ByVal cDELI As String = ",") As Variant

    Dim oDB              As ADODB.Connection
    Dim oRS              As ADODB.Recordset
    Dim X, nCAMPOS As Integer
    Dim aRETU, aOPE, eVAL As Variant
    Dim lOPEN As Boolean
    Dim lRSOP As Boolean
    Dim cERRO As String
    
    On Error GoTo errhandler
    
    lOPEN = False
    lRSOP = False
    nCAMPOS = UBound(aCAM) + 1
    ReDim aRETU(nCAMPOS)
    For X = 0 To nCAMPOS - 1
        aRETU(X) = ""
    Next X
    
    If Not FileExist(SomaExt(cARQ), True) Then
        PegSQLDeliAdo = aRETU
        Exit Function
    End If
    
    
    cARQ = GeracArq(cARQ, , False)

    Set oDB = New ADODB.Connection
    oDB.CursorLocation = adUseClient
    oDB.ConnectionTimeout = 120
    oDB.Open cARQ

    lOPEN = True
    
    Set oRS = New ADODB.Recordset
    
    oRS.Open cSQL, oDB, adOpenForwardOnly, adLockReadOnly

    lRSOP = True
    If Not oRS.EOF Then
        oRS.MoveFirst
        While Not oRS.EOF
            For X = 0 To nCAMPOS - 1
                aOPE = SepSqlOpe(aCAM(X))
                If aOPE(0) = "" Or aOPE(1) = "" Or aOPE(2) = "" Then ''Sem Operacoes
                    eVAL = oRS(aCAM(X))
                Else
                    eVAL = MathOper(oRS(aOPE(1)), oRS(aOPE(2)), aOPE(0))
                End If
                aRETU(X) = aRETU(X) & FixStr(eVAL)
            Next X
            oRS.MoveNext
            If Not oRS.EOF Then
                For X = 0 To nCAMPOS - 1
                    aRETU(X) = aRETU(X) & cDELI
                Next X
            End If

        Wend
    End If
    oRS.Close
    oDB.Close
    Set oRS = Nothing
    Set oDB = Nothing
    PegSQLDeliAdo = aRETU
    Exit Function
    
errhandler:
    cERRO = "Peg SQL DELI ADO" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cSQL & Chr(13) & Chr(10)
    If lRSOP Then
        cERRO = cERRO & ADORsStatus(oRS.Status) & Chr(13) & Chr(10)
    End If
    Select Case Err.Number
    Case Else
        If lOPEN Then
            ADOErro oRS.ActiveConnection.Errors, cERRO
        Else
            ADOErro oDB.Errors, cERRO
        End If
        Exit Function
    End Select
End Function

Public Function ApagaSQLpAdo(ByVal cARQ As String, ByVal cSQL As String, Optional ByVal cTexto As String = "Confirme Exclusão") As Boolean
    ApagaSQLpAdo = False
    If MDG(cTexto, "Exclusão Registro") Then
        ApagaSQLpAdo = APAGASQLADO(cARQ, cSQL)
    End If
End Function

Public Function GrvSQLado(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, ByVal aCAM As Variant, ByVal aVAL As Variant, ByVal aFOR As Variant) As Boolean
    Dim oDB              As ADODB.Connection
    Dim oRS              As ADODB.Recordset
    Dim eVAL, aOPE  As Variant
    Dim X, nLENVAR, nLENCAM As Long
    'Dim nPOS As Long
    Dim lGRAVA           As Boolean
    Dim cTabela          As String
    Dim eVAZIO As Variant
    Dim cTIPO As String
    Dim aRETU As Variant
    Dim lOPEN As Boolean
    Dim lRSOP As Boolean
    Dim cERRO As String
    Dim oFIELD As ADODB.Field

    On Error GoTo errhandler
    lOPEN = False
    lRSOP = False
    GrvSQLado = False
    cTabela = NomeTableSql(cSQL)

    aRETU = TipoConn(cARQ)

    Set oDB = New ADODB.Connection
    oDB.CursorLocation = adUseClient
    oDB.ConnectionTimeout = 120
    oDB.Open cARQ

    lOPEN = True
    Set oRS = New ADODB.Recordset
    oRS.Open cSQL, oDB, adOpenStatic, adLockOptimistic

    lRSOP = True
    
    If oRS.EOF Then
        oRS.Close
        oDB.Close
        Set oRS = Nothing
        Set oDB = Nothing
        Exit Function
    End If
    


    While Not oRS.EOF

        For X = 0 To nITEM - 1                   ''Matriz Vb comeca 0
       
            Set oFIELD = oRS(aCAM(X))
            ''nLENCAM = oRS(aCAM(x)).DefinedSize
            ''cTIPO = TipoDado2(oRS(aCAM(x)).type)
        
        
            cTIPO = TipoDado2(oFIELD.Type)
            lGRAVA = True
       
        
            ''Evita Gravar Campos nullos correcao data nula abaixo
            If IsNull(aVAL(X)) Then
                lGRAVA = False
                eVAL = aVAL(X)
            Else
                aOPE = SepSqlOpe(aCAM(X))
                If aOPE(0) = "" Or aOPE(1) = "" Or aOPE(2) = "" Then ''Operacao com Campos
                    eVAL = aVAL(X)               ''Grava o Valor Default
                Else
                    eVAL = MathOper(oRS(aOPE(1)), oRS(aOPE(2)), aOPE(0))
                End If
            End If
            '
        
      
            ''Evitar Gravar String Vazias Campos Textos
            If VarType(eVAL) = vbString And cTIPO = "S" Then
                If eVAL = "" Then
                    ''EVAL = " " ''nao apagava o campo ja preenchido antes
                    ''lGRAVA = False
                Else
                    nLENCAM = oFIELD.DefinedSize
                    If nLENCAM > 0 Then          ''Evita Retorno Campos -1
                        nLENVAR = Len(eVAL)
                        If nLENVAR > nLENCAM Then
                            eVAL = Mid(eVAL, 1, nLENCAM)
                        End If
                    End If
                End If
                
                If oFIELD.Type = adLongVarWChar And eVAL = "" Then ''Nao saida '' ou null
                   eVAL = "_"
                End If
                
                
            End If
      
      
            ''Evitar Gravar String Vazias Campos DAta
            If cTIPO = "D" Then
                If DataBranco(eVAL) Then
                    If aRETU(2) = "MDB" Or aRETU(2) = "MYSQL" Then
                        eVAL = NullDate(aRETU(2))
                        lGRAVA = True
                        aFOR(X) = ""
                    Else
                        lGRAVA = False
                    End If
                End If
            End If
        
            If cTIPO = "B" Then
                aFOR(X) = ""
            End If
            
      
            ''Efetua a Gravaçao
            If lGRAVA Then
                oRS(aCAM(X)) = FVar(eVAL, aFOR(X), eVAZIO)
                'oFIELD = FVar(eVAL, aFOR(X), eVAZIO)
            End If
 
        Next X
        oRS.Update
        oRS.MoveNext
      
    Wend
    
    oRS.Close
    oDB.Close
    Set oRS = Nothing
    Set oDB = Nothing
    GrvSQLado = True
    Exit Function
    
errhandler:
    cERRO = "Grava SQL ADO" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cSQL & Chr(13) & Chr(10)
    If lRSOP Then
        cERRO = cERRO & ADORsStatus(oRS.Status) & Chr(13) & Chr(10)
    End If
    Select Case Err.Number
    Case 3265, 3421, 94
        cERRO = cERRO & "Campo:" & aCAM(X) & Chr(13) & Chr(10)
        ADOErro oRS.ActiveConnection.Errors, cERRO
        Resume Next
    Case -2147217864 '', -2147467259                 'Update Mysql/informix
        ''      oRS.Close
        ''     oDB.Close
        ''     Set oRS = Nothing
        ''     Set oDB = Nothing
        GrvSQLado = True
        Exit Function
        ''             Resume Next
    Case Else
        If lOPEN Then
            ADOErro oRS.ActiveConnection.Errors, cERRO
        Else
            ADOErro oDB.Errors, cERRO
        End If
    End Select
    
End Function

Public Function IncluiSQLAdo(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, _
                             Optional ByVal aCAM As Variant = 0, Optional ByVal aVAL As Variant = 0, _
                             Optional ByVal lCHECK As Boolean = False, _
                             Optional ByVal lMES As Boolean = True, _
                             Optional ByVal aIDDES As Variant = 0) As Boolean
    Dim oDB              As ADODB.Connection
    Dim oRS              As ADODB.Recordset
    Dim X As Long
    Dim nCAMPOS               As Long
    Dim lTEM As Boolean
    Dim lGRAVA            As Boolean
    Dim aRETUID As Variant
    Dim cARQ1 As String
    Dim aRETU As Variant
    Dim lOPEN As Boolean
    Dim lRSOP As Boolean
    Dim cERRO As String
    
    IncluiSQLAdo = False
    On Error GoTo errhandler

    lTEM = False
    lRETU = False
    aRETU = TipoConn(cARQ)

    cARQ1 = aRETU(1)
    lOPEN = False
    lRSOP = False
    
    Set oDB = New ADODB.Connection
    oDB.CursorLocation = adUseClient
    oDB.ConnectionTimeout = 120
    oDB.Open cARQ1
    
    lOPEN = True

    Set oRS = New ADODB.Recordset
    oRS.Open cSQL, oDB, adOpenStatic, adLockOptimistic

    lRSOP = True

    If lCHECK Then
        If Not oRS.EOF Then
            lTEM = True
        End If
    End If

    If Not lTEM Then
        
        oRS.AddNew
        For X = 0 To nITEM - 1                   ''Matriz Vb comeca 0
            lGRAVA = True
            ''Evita Gravar Campos nullos
            If IsNull(aVAL(X)) Then
                lGRAVA = False
            End If
            If lGRAVA Then
                oRS(aCAM(X)) = aVAL(X)
            End If
        Next
        
        ' PEGAR O ID
        Select Case aRETU(2)
        Case "SQLSERVER", "MDB", "MYSQL"
            oRS.Update
        End Select
        If Not IsNumeric(aIDDES) Then            ''Se for numerico nao e matriz
            nCAMPOS = UBound(aIDDES)
            ReDim aRETUID(nCAMPOS + 1)
            For X = 0 To nCAMPOS
                aRETUID(X) = oRS(aIDDES(X))
            Next X
            eRETU01 = aRETUID
        End If
        Select Case aRETU(2)
        Case "SQLSERVER", "MDB", "MYSQL"
              
        Case Else
            oRS.Update
        End Select

        lRETU = True
    End If
    oRS.Close
    oDB.Close
    Set oRS = Nothing
    Set oDB = Nothing
    
    If lCHECK And lTEM Then
        If lMES Then
            Alert ("Item ja Cadastrado Com Esta Chave")
        End If
    End If
    IncluiSQLAdo = lRETU
    Exit Function
errhandler:
    cERRO = "Inclui SQL ADO" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cSQL & Chr(13) & Chr(10)
    If lRSOP Then
        cERRO = cERRO & ADORsStatus(oRS.Status) & Chr(13) & Chr(10)
    End If
    Select Case Err.Number
    Case 3265                                    ''Campo Diferente
        cERRO = cERRO & "Campo Diferente:" & aCAM(X) & Chr(13) & Chr(10)
        ADOErro oRS.ActiveConnection.Errors, cERRO
        Resume Next
    Case 3315, 94                                'Campo null
        Resume Next
    Case 424                                     ''Campo Inexistente
        cERRO = cERRO & "Campo Inexistente:" & aCAM(X) & Chr(13) & Chr(10)
        ADOErro oRS.ActiveConnection.Errors, cERRO
        Resume Next
    Case 3163                                    ''Valor Maior que o Campo
        cERRO = cERRO & "Valor Maior que o Campo:" & aCAM(X) & Chr(13) & Chr(10)
        ADOErro oRS.ActiveConnection.Errors, cERRO
        Resume Next
    Case 3219                                    ''close jetfox
        Resume Next
    Case -2147217887                             ''Update jetfox
        Resume Next
    Case Else
        If lOPEN Then
            ADOErro oRS.ActiveConnection.Errors, cERRO
        Else
            ADOErro oDB.Errors, cERRO
        End If
        Exit Function
        ''        PegSQL = aPAD
    End Select
End Function

Public Function PegSQLAdo(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, ByVal aCAM As Variant, ByVal aFOR As Variant, ByVal aPAD As Variant) As Variant
    Dim oDB         As ADODB.Connection
    Dim oRS         As ADODB.Recordset
    Dim X           As Long
    Dim aRETU       As Variant
    Dim lOPEN As Boolean
    Dim lRSOP As Boolean
    Dim cERRO As String
    Dim aOPE As Variant
    Dim eVAL As Variant

    On Error GoTo errhandler
    ReDim aRETU(nITEM)
    lOPEN = False
    lRSOP = False

    Set oDB = New ADODB.Connection
    oDB.CursorLocation = adUseClient
    oDB.ConnectionTimeout = 120
    oDB.Open cARQ

    lOPEN = True
    Set oRS = New ADODB.Recordset
    oRS.Open cSQL, oDB, adOpenForwardOnly, adLockReadOnly
    
    lRSOP = True


    If Not oRS.EOF Then
        lRETU = True
        For X = 0 To nITEM - 1                   ''Matriz Vb comeca 0
            aOPE = SepSqlOpe(aCAM(X))
            If aOPE(0) = "" Or aOPE(1) = "" Or aOPE(2) = "" Then ''Sem Operacoes
                eVAL = oRS(aCAM(X))
            Else
                eVAL = MathOper(oRS(aOPE(1)), oRS(aOPE(2)), aOPE(0))
            End If
              
            aRETU(X) = FVar(eVAL, aFOR(X), aPAD(X))
        Next
    Else
        lRETU = False
        aRETU = aPAD
    End If
    oRS.Close
    oDB.Close
    Set oRS = Nothing
    Set oDB = Nothing
    PegSQLAdo = aRETU
    Exit Function
    
errhandler:
    cERRO = "Peg SQL ADO" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cSQL & Chr(13) & Chr(10)
    If lRSOP Then
        cERRO = cERRO & ADORsStatus(oRS.Status)
    End If
    Select Case Err.Number
    Case 3265, 5
        aRETU(X) = aPAD(X)
        cERRO = cERRO & "Campo:" & aCAM(X) & Chr(13) & Chr(10)
        ADOErro oRS.ActiveConnection.Errors, cERRO
        Resume Next
    Case 94                                      ''Campo Esta com Null Pega Padrao
        ''SayErro (" PEGSQL Campo:" & aCAM(x))
        aRETU(X) = aPAD(X)
        Resume Next
    Case Else
        If lOPEN Then
            ADOErro oRS.ActiveConnection.Errors, cERRO
        Else
            ADOErro oDB.Errors, cERRO
        End If
        PegSQLAdo = aPAD
    End Select
End Function

Public Function PegCountSQLADO(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    Dim aRETU As Variant
    Dim cSQL As String
    PegCountSQLADO = eDEFAULT
    cSQL = "SELECT count(" & cCAMPO & ") AS CAMPO FROM " & cTABLEWHERE
    aRETU = PegSQLAdo(cARQ, cSQL, 1, Array("CAMPO"), Array(""), Array(eDEFAULT))
    If lRETU Then
        PegCountSQLADO = aRETU(0)
    End If
End Function

Public Function PegMAXSQLADO(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    Dim aRETU As Variant
    Dim cSQL As String
    PegMAXSQLADO = eDEFAULT
    cSQL = "SELECT MAX(" & cCAMPO & ") AS CAMPO FROM " & cTABLEWHERE
    aRETU = PegSQLAdo(cARQ, cSQL, 1, Array("CAMPO"), Array(""), Array(eDEFAULT))
    If lRETU Then
        PegMAXSQLADO = aRETU(0)
    End If
End Function

Public Function PegUltSQLAdo(ByVal cARQ As String, ByVal cSQL As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    Dim oDB              As ADODB.Connection
    Dim oRS              As ADODB.Recordset
    Dim lOPEN As Boolean
    Dim lRSOP As Boolean
    Dim cERRO As String

    On Error GoTo errhandler

    PegUltSQLAdo = eDEFAULT

    lOPEN = False
    lRSOP = False
    Set oDB = New ADODB.Connection
    oDB.CursorLocation = adUseClient
    oDB.ConnectionTimeout = 120
    oDB.Open cARQ

    lOPEN = True
    Set oRS = New ADODB.Recordset
    
    oRS.Open cSQL, oDB, adOpenForwardOnly, adLockReadOnly
    lRSOP = True


    If Not oRS.EOF Then
        oRS.MoveLast
        If Not IsNull(oRS(cCAMPO)) Then
            PegUltSQLAdo = oRS(cCAMPO)
        End If
    End If
    oRS.Close
    oDB.Close
    Set oRS = Nothing
    Set oDB = Nothing
    Exit Function
errhandler:
    cERRO = "Peg Ult SQL ADO" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cSQL & Chr(13) & Chr(10)
    If lRSOP Then
        cERRO = cERRO & ADORsStatus(oRS.Status) & Chr(13) & Chr(10)
    End If
    Select Case Err.Number
    Case Else
        If lOPEN Then
            ADOErro oRS.ActiveConnection.Errors, cERRO
        Else
            ADOErro oDB.Errors, cERRO
        End If
        Exit Function
    End Select
    
End Function

Public Function SQLMoveRegADO(ByVal cARQORI As String, _
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
                    
    Dim oDB, oDBDES  As ADODB.Connection
    Dim oRS, oRSDES  As ADODB.Recordset
    Dim cARQORI1 As String
    Dim cARQDES1 As String
    Dim aRETU As Variant
    Dim X As Integer
    Dim nCAMPOS As Integer
    Dim aVALORI As Variant
    Dim aRETUID As Variant
    Dim lOPEN As Boolean
    Dim lRSOP As Boolean
    Dim cERRO As String
    
    On Error GoTo errhandler

    SQLMoveRegADO = False
    lOPEN = False
    lRSOP = False
    
    aRETU = TipoConn(cARQORI)

    cARQORI1 = aRETU(1)
    
    aRETU = TipoConn(cARQDES)

    cARQDES1 = aRETU(1)

    
    Set oDB = New ADODB.Connection
    oDB.CursorLocation = adUseClient
    oDB.ConnectionTimeout = 120
    oDB.Open cARQORI1
    
    Set oDBDES = New ADODB.Connection
    oDBDES.CursorLocation = adUseClient
    oDBDES.ConnectionTimeout = 120
    oDBDES.Open cARQDES1
    
    lOPEN = True
  
    Set oRS = New ADODB.Recordset
    oRS.Open cSQLORI, oDB, adOpenStatic, adLockOptimistic
    
  
    Set oRSDES = New ADODB.Recordset
    oRSDES.Open cSQLDES, oDBDES, adOpenStatic, adLockOptimistic
  
    lRSOP = True
    
    If Not oRS.EOF Then
        oRS.MoveFirst
        While Not oRS.EOF
            If Not IsNumeric(aCAMORI) Then       ''Se for numerico nao e matriz
                nCAMPOS = UBound(aCAMORI)
                ReDim aVALORI(nCAMPOS + 1)
                For X = 0 To nCAMPOS
                    aVALORI(X) = oRS(aCAMORI(X))
                Next X
            End If
            If InStr(cOPEORI, "DEL") > 0 Then
                oRS.Delete
            End If
            If InStr(cOPEDES, "INC") > 0 Then
                oRSDES.AddNew
            Else
                ''               oRSDES.Edit
            End If
            
            If Not IsNumeric(aCAMDES) Then       ''Se for numerico nao e matriz
                nCAMPOS = UBound(aCAMDES)
                For X = 0 To nCAMPOS
                    oRSDES(aCAMDES(X)) = aVALORI(X)
                Next X
            End If
            If Not IsNumeric(aOUTDES) Then       ''Se for numerico nao e matriz
                nCAMPOS = UBound(aOUTDES)
                For X = 0 To nCAMPOS
                    oRSDES(aOUTDES(X)) = aOUTORI(X)
                Next X
            End If
            Select Case aRETU(2)
            Case "SQLSERVER", "MDB", "MYSQL"
                oRSDES.Update
            End Select
            
            ' PEGAR O ID
            If Not IsNumeric(aIDDES) Then        ''Se for numerico nao e matriz
                nCAMPOS = UBound(aIDDES)
                ReDim aRETUID(nCAMPOS + 1)
                For X = 0 To nCAMPOS
                    aRETUID(X) = oRSDES(aIDDES(X))
                Next X
                eRETU01 = aRETUID
            End If
            Select Case aRETU(2)
            Case "SQLSERVER", "MDB", "MYSQL"
                  
            Case Else
                oRSDES.Update
            End Select
            oRS.MoveNext
        Wend
    End If
    oRS.Close
    oDB.Close
    
    Set oRS = Nothing
    Set oDB = Nothing
    SQLMoveRegADO = True
    Exit Function
errhandler:
    cERRO = "Soma SQL ADO" & Chr(13) & Chr(10) & cARQORI & Chr(13) & Chr(10) & cSQLORI & Chr(13) & Chr(10)
    If lRSOP Then
        cERRO = cERRO & ADORsStatus(oRS.Status) & Chr(13) & Chr(10)
    End If
    Select Case Err.Number
    Case Else
        If lOPEN Then
            ADOErro oRS.ActiveConnection.Errors, cERRO
        Else
            ADOErro oDB.Errors, cERRO
        End If
        Exit Function
    End Select
End Function

Function ADOErro(ByRef oErro As Variant, Optional ByVal cERRO As String = "")
    Dim errorObject As ADODB.error
    For Each errorObject In oErro
        cERRO = cERRO & " Ado Erro Numero: " & errorObject.Number & vbCrLf
        cERRO = cERRO & " Ado Descricao  : " & errorObject.Description & vbCrLf
        cERRO = cERRO & " Ado Fonte      : " & errorObject.Source & vbCrLf
        cERRO = cERRO & " Ado SQL        : " & errorObject.SQLState & vbCrLf
        cERRO = cERRO & " Ado Erro Nativo: " & errorObject.NativeError & vbCrLf
    Next
    SayErro cERRO, True
End Function

Public Function ADO_IsRecordsetEmpty(ByRef oRS As ADODB.Recordset) As Boolean
    Dim cERRO As String
    On Error GoTo trataerro
    If Not ADO_IsOpen(oRS) Then
        ADO_IsRecordsetEmpty = True
        Exit Function
    End If
    If oRS.EOF And oRS.BOF Then
        ADO_IsRecordsetEmpty = True
        Exit Function
    End If
  
    If oRS.RecordCount = 0 Then
        ADO_IsRecordsetEmpty = True
        Exit Function
    End If
    Exit Function
trataerro:
    cERRO = "ADO_IsRecordsetEmpty"
    cERRO = cERRO & ADORsStatus(oRS.Status) & Chr(13) & Chr(10)
    Select Case Err.Number
    Case Else
        ADOErro oRS.ActiveConnection.Errors, cERRO
    End Select
End Function

Public Function ADO_FieldValueToString(ByRef FLD As ADODB.Field, Optional ByVal sNullRepresentation As String = "(null)") As String
    'Call TraceEnters(MODULE_NAME & "::ADO_FieldValueToString")
    'TraceDetail = "To convert the value of a field into a string"
    On Error GoTo trataerro
  
    Select Case (FLD.Type)
        '~~~~~~
    Case adBSTR, adChar, adVarChar, adVarWChar, adWChar, adLongVarChar, adLongVarWChar, adGUID:
        ADO_FieldValueToString = nZ(FLD.Value, sNullRepresentation)
        '~~~~~~
    Case adBigInt, adCurrency, adDecimal, adDouble, adInteger, adNumeric, adSingle, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt, adBoolean:
        ADO_FieldValueToString = Format(nZ(FLD.Value, sNullRepresentation))
        '~~~~~~
    Case adVariant:
        ADO_FieldValueToString = Format(nZ(FLD.Value, sNullRepresentation))
        '~~~~~~
    Case adDate, adDBDate, adDBTime, adDBTimeStamp:
        ADO_FieldValueToString = Format(nZ(FLD.Value, sNullRepresentation))
        '~~~~~~
    Case adBinary, adVarBinary, adLongVarBinary:
        ADO_FieldValueToString = IIf(IsNull(FLD.Value), sNullRepresentation, BytesToHexString(FLD.Value))
        '~~~~~~
    Case Else:
        'Err.Raise 13, MODULE_NAME & "::ADO_FieldValueToString", "Sorry, unable to convert fields of type " & fld.Type & ", (" & ADO_DataTypeEnumToEnglish(fld.Type) & ") to string."
    End Select
    Exit Function
trataerro:
    SayErro "ADO_FieldValueToString"
    Exit Function
End Function

Public Sub ADO_FreeRecordset(ByRef rs As ADODB.Recordset)
    On Error Resume Next
    If rs.State = adStateOpen Then
        rs.Close
    End If
    Set rs = Nothing
    On Error GoTo 0
End Sub

Public Function ADO_IsOpen(ByRef oADOObject As Object) As Boolean
    ' Purpose: To determine if a connection or a recordset is open
    ' !! Assumes/Pre: Nothing
    ' Parameters:
    '  oADOObject as Object  -
    ' Returns: Boolean
    '       Success-
    '       Failure- Raises error on failure
    ' Revision history:
    '   2004-Feb-20 10:47 [Michael Johnson] Initial creation
    'Call TraceEnters(MODULE_NAME & "::ADO_IsOpen")
    'TraceDetail = "To determine if a connection or a recordset is open"
    On Error GoTo trataerro
  
    If oADOObject Is Nothing Then
        ' ADO_IsOpen = False ', already the default.
        Exit Function
    End If
  
    If oADOObject.State = adStateOpen Then
        ADO_IsOpen = True
        Exit Function
    End If
    Exit Function

trataerro:
    SayErro "ADO_ISOPEN"
  
End Function

Public Function nZ( _
       vValue As Variant, _
       Optional vReplacementIfNull As Variant = 0 _
       ) As Variant
    ' Purpose: To replace a NULL with another value, if the value is Null.
    ' Example/Note:     sResult = Nz(rs.Fields(sFieldName), "") ' See MS Access VBA for Nz documentation
    ' !! Assumes/Pre: Nothing
    ' Parameters:
    '   vValue- Value to evaluate if null
    '   vReplacementIfNull - What should replace a Null value
    ' Returns: Variant
    '       Success- If not null, returns the supplied value, else returns the replacement
    '       Failure- Raises error on failure
    ' Revision history:
    '   Michael Johnson     2002-Mar-12 1243     Initial creation
    '   2003-Aug-22 10:17 [Michael B. Johnson] Abreviated and changed to variants
    '   2004-Feb-18 16:36 [Michael B. Johnson] Changed from using TypeName() to IsNull()
    'Call TraceEnters(MODULE_NAME & "::Nz")
    'TraceDetail = "To replace a NULL with a string, if the value is Null."
  
    If IsNull(vValue) Then
        nZ = vReplacementIfNull
    Else
        nZ = vValue
    End If
  
    Exit Function
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
    On Error GoTo trataerro
    
    Dim sAccumulator As String, lCtr As Long
    Dim sHex As String, sFormatted As String
    
    If Not TypeName(vaBytes) = "Byte()" Then
        BytesToHexString = ""
        Exit Function
    End If
    
    For lCtr = 0 To UBound(vaBytes)
        sHex = Hex(vaBytes(lCtr))
        sFormatted = Format(sHex, "@@")
        sAccumulator = sAccumulator & Replace(sFormatted, " ", "0")
    Next
    sAccumulator = "0x" & sAccumulator
    BytesToHexString = sAccumulator
    Exit Function
trataerro:
    BytesToHexString = ""
    Exit Function
End Function


