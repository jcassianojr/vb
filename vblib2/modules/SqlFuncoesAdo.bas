Attribute VB_Name = "SqlFuncoesAdo"
Option Explicit
Public Function VFPSetValues(ByRef oCON As Object) As Boolean ''as ADODB.Connection as object
 'VFPSetValues DB
 'VFPSetValues ODB
 'VFPSetValues oCONN
   
    Dim cCOM As String
    Dim oCOMANDO As ADODB.Command
    
    'inicializa as variaveis
    VFPSetValues = False
    cCOM = ""
    Set oCOMANDO = New ADODB.Command
    
   'atribui o comando a conecao passada como referencia
    oCOMANDO.CommandType = adCmdText
    oCOMANDO.ActiveConnection = oCON
    
     'set deleted on pega o registro correto evitando gravar em registro apagados
    cCOM = "set deleted on"
    oCOMANDO.CommandText = cCOM
    oCOMANDO.Execute
        
     'set null permite gravar parcialmente (permite gravar um registro novo sem passar gravar todos os campos)
     cCOM = "set null off"
     oCOMANDO.CommandText = cCOM
     oCOMANDO.Execute
    VFPSetValues = True

End Function
Public Function AdoComandodbf(ByVal cARQ As String, ByVal cTABLE As String, ByVal cCOMANDO As String) As Boolean
    Dim cCOM As String
    Dim oCON As ADODB.Connection
    Dim oCOMANDO As ADODB.Command
    On Error GoTo trataerro
    
    'ZAP PACK
    'delete from mail where numero=1
    'insert into mail (numero,erro,data,hora,de,destino,assunto,texto) values (recno(),"erro",ctod(date()),"12:00","de","destino","assunto","texto")
    
    'Passado "" pega o da string
    If Len(cTABLE) = 0 Then
       cTABLE = NomeTableSql(cCOMANDO)
    End If
    'atribui os valores iniciais
    AdoComandodbf = False
    cCOM = ""
    Set oCON = New ADODB.Connection
    Set oCOMANDO = New ADODB.Command
    
    'abre a conecao
    oCON.ConnectionTimeout = 120
    oCON.Open cARQ
    
    'atribui o comando a conecao
    oCOMANDO.CommandType = adCmdText
    oCOMANDO.ActiveConnection = oCON
    
    'para pack e zap considera todos necessario setar aqui pois na string de conecao esta marcado delete on
    If cCOMANDO = "ZAP" Or cCOMANDO = "PACK" Then
       cCOM = "set deleted off"
       oCOMANDO.CommandText = cCOM
       oCOMANDO.Execute
    Else
       cCOM = "set deleted on"
       oCOMANDO.CommandText = cCOM
       oCOMANDO.Execute
    End If
    'set null permite gravar parcialmente (permite gravar um registro novo sem passar gravar todos os campos)
     cCOM = "set null off"
     oCOMANDO.CommandText = cCOM
     oCOMANDO.Execute
    
    'abre a tabela
    cCOM = "use " & cTABLE
    oCOMANDO.CommandText = cCOM
    oCOMANDO.ActiveConnection = oCON
    oCOMANDO.Execute
    
    'executa o comando
    If cCOMANDO = "ZAP" Then
        cCOM = "DELETE FROM " & cTABLE & " WHERE 1=1"
        oCOMANDO.CommandText = cCOM
        oCOMANDO.Execute
    End If
    If cCOMANDO = "ZAP" Or cCOMANDO = "PACK" Then
        cCOM = "PACK"
        oCOMANDO.CommandText = cCOM
        oCOMANDO.Execute
    Else
        cCOM = cCOMANDO
        oCOMANDO.CommandText = cCOM
        oCOMANDO.Execute
    End If
    
    'fecha a coneccao
    oCON.Close
    AdoComandodbf = True
    
    Exit Function
trataerro:
    Select Case Err.Number
    Case Else
        SayErro "ADO Comando DBF:" & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cTABLE & Chr(13) & Chr(10) & cCOM & Chr(13) & Chr(10)
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
    aRETU = TipoConn(cARQ) ''Gera string conneccao
    If aRETU(0) = "ADO" Then
        cCONN = aRETU(1)
        If InStr(UCase(cARQ), "VFPOLEDB") Then
           ADOComando = AdoComandodbf(cARQ, "", cSQL)
           Exit Function
        End If
        Set oDB = New ADODB.Connection
        oDB.ConnectionTimeout = 120
        oDB.Open cCONN
        
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
Public Function APAGASQLADO(ByVal cARQ As String, ByVal cSQL As String) As Boolean
    Dim nPOS As Integer
    Dim cNOME As String
    'select * from tabela where campo=valorcampo ....
    'delete * from tabela where campo=valorcampo ....
    'delete from tabela where campo=valorcampo ....
    'from tabela where campo=valorcampo ....
    'todos os csql viram delete from tabela where campo=valorcampo ....
    'requer from e where no csql
    
    
    APAGASQLADO = False
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
    'efetua substituicao se estiver com select * from  ou outros casos --> delete from usando mid
    'todos os csql viram delete from tabela where campo=valorcampo .... opcoes de passagem csql acima
    If nPOS > 0 Then
        cSQL = "DELETE FROM " & Mid(cSQL, nPOS + 5)
        APAGASQLADO = ADOComando(cARQ, cSQL)
    Else
        If InStr(cNOME, "DELETE") > 0 Then 'so executa se tiver delete
            APAGASQLADO = ADOComando(cARQ, cSQL)
        End If
    End If
End Function

Public Function SomaSQLAdo(ByVal cARQ As String, ByVal cSQL As String, ByVal aCAM As Variant) As Variant
'esta faz soma com loop
'use pegsumsql se o loop nao for necessario
    Dim oDB              As ADODB.Connection
    Dim oRS              As ADODB.Recordset
    Dim lOPEN As Boolean
    Dim lRSOP As Boolean
    Dim X, nCAMPOS As Integer
    Dim aRETU, aOPE, eVAL As Variant
    Dim cERRO As String
     Dim oCOMANDO As ADODB.Command
    Dim cCOM As String
    
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
    
    'na string de conecao delete =yes
    'If  InStr(ucase(cARQ), "VFPOLEDB")>0 Then ''delete on para foxpro nao usar registro deletados
    '     vfpsetvalues ODB
    'End If
    
    
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
    Dim oCOMANDO As ADODB.Command
    Dim cCOM As String
    
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
    
    'na string de conecao delete =yes
    'If InStr(cARQ, "VFPOLEDB")>0 Then ''delete on para foxpro nao usar registro deletados
    '  vfpsetvalues ODB
    'End If
  

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
     Dim oCOMANDO As ADODB.Command
    Dim cCOM As String

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
    
 'Set null off permitido deixar campos em branco set deleted on ja esta na string de conecao
    If InStr(cARQ, "VFPOLEDB") Then
    'vfpsetvalues ODB 'Testar funcao para precisar das linhas abaixo
        Set oCOMANDO = New ADODB.Command
        oCOMANDO.ActiveConnection = oDB
        oCOMANDO.CommandType = adCmdText
        cCOM = "Set null off"
        oCOMANDO.CommandText = cCOM
        oCOMANDO.Execute
    End If

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
    Dim oCOMANDO As ADODB.Command
    Dim cCOM As String
    
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
    
    'Set null off permitido deixar campos em branco set deleted on ja esta na string de conecao
    If InStr(cARQ1, "VFPOLEDB") Then
       'vfpsetvalues odb 'testar funcao para nao precisar das linhas abaixo
        Set oCOMANDO = New ADODB.Command
        oCOMANDO.ActiveConnection = oDB
        oCOMANDO.CommandType = adCmdText
        cCOM = "Set null off"
        oCOMANDO.CommandText = cCOM
        oCOMANDO.Execute
    End If
 
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
    Dim oCOMANDO As ADODB.Command
    Dim cCOM As String
    

    On Error GoTo errhandler
    ReDim aRETU(nITEM)
    lOPEN = False
    lRSOP = False

    Set oDB = New ADODB.Connection
    oDB.CursorLocation = adUseClient
    oDB.ConnectionTimeout = 120
    oDB.Open cARQ

'na string de conecao delete =yes
'If  InStr(cARQ, "VFPOLEDB")>0 Then ''delete on para foxpro nao usar registro deletados
'   vfpsetvalues odb
'End If

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
    PegCountSQLADO = PegOperSQLADO(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, "COUNT")
End Function
Public Function PegMINSQLADO(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    PegMINSQLADO = PegOperSQLADO(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, "MIN")
End Function
Public Function PegMAXSQLADO(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    PegMAXSQLADO = PegOperSQLADO(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, "MAX")
End Function
Public Function PegSUMSQLADO(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
  PegSUMSQLADO = PegOperSQLADO(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, "SUM")
End Function
Public Function PegCampoSQLADO(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
  PegCampoSQLADO = PegOperSQLADO(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, "CAMPO")
End Function
Public Function PegOperSQLADO(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant, ByVal coper As String) As Variant
    Dim aRETU As Variant
    Dim cSQL As String
    PegOperSQLADO = eDEFAULT
    
    
    If coper = "CAMPO" Then
       If Len(cCAMPO) = 0 Then 'Passado a string ja com o campo select numero from tabela where numero=999999
          cTABLEWHERE = UCase(cTABLEWHERE)
          cSQL = Replace(cTABLEWHERE, " FROM ", " AS CAMPO FROM ") 'inclui as campo variavel padrao no pegsql abaixo
       Else
          cSQL = "SELECT " & cCAMPO & " AS CAMPO FROM " & cTABLEWHERE
       End If
    Else
       cSQL = "SELECT " & coper & "(" & cCAMPO & ") AS CAMPO FROM " & cTABLEWHERE
    End If
    If coper = "SUM" Or coper = "COUNT" Or coper = "MAX" Or coper = "MIN" Or IsNumeric(eDEFAULT) Then
       aRETU = PegSQLAdo(cARQ, cSQL, 1, Array("CAMPO"), Array("N"), Array(eDEFAULT)) ''array retorno tipo N numerico
    Else
       aRETU = PegSQLAdo(cARQ, cSQL, 1, Array("CAMPO"), Array(""), Array(eDEFAULT))
    End If
    If lRETU Then
        PegOperSQLADO = aRETU(0)
    End If
End Function
Public Function PegUltSQLAdo(ByVal cARQ As String, ByVal cSQL As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
' Esta usa loop
' use a opcao pegsqlmax para a maioria dos casos
' ou quando o campo nao e numerica
    Dim oDB              As ADODB.Connection
    Dim oRS              As ADODB.Recordset
    Dim lOPEN As Boolean
    Dim lRSOP As Boolean
    Dim cERRO As String
   Dim oCOMANDO As ADODB.Command
    Dim cCOM As String

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
     Dim oCOMANDO As ADODB.Command
     Dim OCOMANDO2 As ADODB.Command
     
     
    Dim cCOM As String
    
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
    

   'Set null off permitido deixar campos em branco set deleted on ja esta na string de conecao
    If InStr(cARQDES1, "VFPOLEDB") Then
        'vfpsetvalues odbdes 'testar funcao para na precisar das linhas abaixo
        Set OCOMANDO2 = New ADODB.Command
        OCOMANDO2.ActiveConnection = oDBDES
        OCOMANDO2.CommandType = adCmdText
        cCOM = "Set null off"
        OCOMANDO2.CommandText = cCOM
        OCOMANDO2.Execute
    End If
    
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

