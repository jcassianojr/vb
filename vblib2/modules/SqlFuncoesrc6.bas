Attribute VB_Name = "SqlFuncoesrc6"
Option Explicit
' CORREÇÃO NO VFPSetValuesrc6

Public Function pgSetValuesrc6(ByRef oCON As Object) As Boolean
    ' Não precisamos declarar oCOMANDO como cCommand, pois usaremos
    ' a execução direta na conexão, que é o padrão nativo e veloz do RC6.
    
    On Error GoTo ErroPostgres
    pgSetValuesrc6 = False
    
    ' 1. Verifica se a conexão é válida
    If oCON Is Nothing Then Exit Function
    
    ' 2. Execução Direta: Substitui o CreateCommand + CommandText + Execute
    ' O método Execute da conexão é o equivalente direto para comandos
    ' sem parâmetros no RC6.
    
    ' Nota: O SQLite/RC6 pode não reconhecer "SET search_path" ou "client_encoding"
    ' dependendo do driver. Se for para configurar o ambiente do Postgres via RC6:
    oCON.Execute "SET search_path TO myschema, public;"
    oCON.Execute "SET client_encoding TO 'WIN1252';"
    
    pgSetValuesrc6 = True
    Exit Function

ErroPostgres:
    ' Como não instanciamos objeto de comando, não há necessidade de limpeza de Set = Nothing
    pgSetValuesrc6 = False
End Function
Public Function VFPSetValuesrc6(ByRef oCON As Object) As Boolean
    ' Não precisamos mais declarar o oCOMANDO como RC6.cCommand
    ' pois não usaremos a propriedade CommandText.
    
    On Error GoTo ErroVFP
    
    ' Inicializa a função
    VFPSetValuesrc6 = False
    
    ' Verifica se a conexão é válida
    If oCON Is Nothing Then Exit Function
    
    ' O RC6 permite executar comandos SQL diretamente através do objeto Connection.
    ' Isso substitui perfeitamente a necessidade de um objeto Command para
    ' configurações simples como "SET ...".
    
    ' 1. Executa o SET DELETED
    oCON.Execute "set deleted on"
    
    ' 2. Executa o SET NULL
    oCON.Execute "set null off"
    
    ' Se chegou até aqui, sucesso
    VFPSetValuesrc6 = True
    
    Exit Function

ErroVFP:
    ' O tratamento de erro continua simples, pois não há objetos de comando para limpar
    VFPSetValuesrc6 = False
End Function


Public Function ADOComandoDBFrc6(ByVal cARQ As String, ByVal cTable As String, ByVal CCOMANDO As String) As Boolean
    Dim cCOM As String
    Dim oCON As RC6.cConnection
    
    On Error GoTo TrataErro

    ' Se cTable estiver vazio, tenta extrair da string de comando
    If Len(cTable) = 0 Then
        cTable = NomeTableSql(CCOMANDO)
    End If
    
    ADOComandoDBFrc6 = False
    cCOM = ""
    
    ' Inicializa Conexão RC6
    Set oCON = New_c.Connection(cARQ)
    
    ' --- SUBSTITUIÇÃO: Execução direta na conexão ---
    ' Eliminamos o objeto cCommand para evitar erros de CommandText e melhorar a performance.

    ' Configurações de ambiente (SETs)
    If CCOMANDO = "ZAP" Or CCOMANDO = "PACK" Then
        oCON.Execute "set deleted off"
    Else
        oCON.Execute "set deleted on"
    End If

    oCON.Execute "set null off"

    ' "USE" (Emulação de ambiente DBF no SQLite)
    If Len(cTable) > 0 Then
        oCON.Execute "use " & cTable
    End If

    ' Execução do comando principal
    If CCOMANDO = "ZAP" Then
        oCON.Execute "DELETE FROM " & cTable & " WHERE 1=1"
    End If
    
    If CCOMANDO = "ZAP" Or CCOMANDO = "PACK" Then
        oCON.Execute "PACK"
    Else
        cCOM = CCOMANDO
        oCON.Execute cCOM
    End If

    Set oCON = Nothing
    ADOComandoDBFrc6 = True
    Exit Function

TrataErro:
    ' Captura o erro
    Dim sMsg As String
    sMsg = "ADO Comando DBF (RC6):" & vbCrLf & cARQ & vbCrLf & cTable & vbCrLf & cCOM & vbCrLf & Err.Description
    SayErro sMsg
    
    ' Limpeza em caso de erro
    Set oCON = Nothing
    ADOComandoDBFrc6 = False
End Function
Public Function ADOComandorc6(ByVal cARQ As String, ByVal cSQL As String) As Boolean
    Dim aRETU As Variant
    Dim oDB As RC6.cConnection
    Dim cCONN As String

    On Error GoTo TrataErro
    ADOComandorc6 = False
    
    cARQ = GeraConn(cARQ)
    aRETU = TipoConn(cARQ)
    
    ' Se a lógica ditar que não é para rodar via RC6, sai
    ' (Ajuste a condição conforme o seu novo identificador de tag, ex: "ADORC6")
    If aRETU(0) <> "ADORC6" Then
        Exit Function
    End If
    
    ' Mantém a compatibilidade com VFP se necessário
    If InStr(UCase(cARQ), "VFPOLEDB") Then
        ADOComandorc6 = AdoComandodbf(cARQ, "", cSQL)
        Exit Function
    End If
    
    ' Ajustes de dialeto (Mantido conforme original)
    If InStr(cSQL, "CURRENTDATETIME") > 0 Then
       Select Case aRETU(2)
          Case "SQLITE": cSQL = Replace(cSQL, "CURRENTDATETIME", " current_timestamp ")
          Case "MDB":    cSQL = Replace(cSQL, "CURRENTDATETIME", " now ")
          Case Else:     cSQL = Replace(cSQL, "CURRENTDATETIME", "'" & Format(Date, "dd/mm/yyyy") & "'")
       End Select
    End If
    
    cCONN = aRETU(1)
    
    ' Inicialização simplificada no RC6
    Set oDB = New_c.Connection(cCONN)
    
    ' O método Execute da conexão substitui a necessidade do objeto Command
    ' para comandos simples (sem parâmetros).
    oDB.Execute cSQL
    
    ' Fechamento simplificado (apenas limpa o objeto)
    Set oDB = Nothing
    
    ADOComandorc6 = True
    Exit Function

TrataErro:
    SayErro "SQL RC6 Comando:" & vbCrLf & cARQ & vbCrLf & cSQL & vbCrLf & Err.Description
    If Not oDB Is Nothing Then Set oDB = Nothing
    ADOComandorc6 = False
End Function

Public Function APAGASQLrc6(ByVal cARQ As String, ByVal cSQL As String) As Boolean
  Dim nPOS As Integer
  Dim cNOME As String
  'select * from tabela where campo=valorcampo ....
  'delete * from tabela where campo=valorcampo ....
  'delete from tabela where campo=valorcampo ....
  'from tabela where campo=valorcampo ....
  'todos os csql viram delete from tabela where campo=valorcampo ....
  'requer from e where no csql


  APAGASQLrc6 = False
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
    APAGASQLrc6 = ADOComandorc6(cARQ, cSQL)
  Else
    If InStr(cNOME, "DELETE") > 0 Then  'so executa se tiver delete
      APAGASQLrc6 = ADOComandorc6(cARQ, cSQL)
    End If
  End If
End Function

Public Function SomaSQLrc6(ByVal cARQ As String, ByVal cSQL As String, ByVal aCAM As Variant) As Variant
    Dim oDB As RC6.cConnection
    Dim oRS As RC6.cRecordset
    Dim lOPEN As Boolean
    Dim lRSOP As Boolean
    Dim x As Integer
    Dim nCAMPOS As Integer
    Dim aRETU As Variant
    Dim aARQ As Variant
    Dim aOPE As Variant
    Dim eVAL As Variant
    Dim carqcon As String

    On Error GoTo errhandler

    lOPEN = False
    lRSOP = False
    nCAMPOS = UBound(aCAM) + 1
    ReDim aRETU(nCAMPOS - 1) ' Ajustado para base 0
    For x = 0 To nCAMPOS - 1
        aRETU(x) = 0
    Next x

    If Not FileConnExist(SomaExt(cARQ), True) Then
        SomaSQLrc6 = aRETU
        Exit Function
    End If

    aARQ = TipoConn(cARQ, , , False)
    carqcon = aARQ(1)

    ' Inicialização simplificada no RC6
    Set oDB = New_c.Connection(carqcon)
    lOPEN = True
  
  
    ConfigurarConexaoRC6 oDB, cARQ

    
    ' Ajustes de ambiente
    'If InStr(UCase(cARQ), "VFPOLEDB") > 0 Then VFPSetValuesrc6 oDB
    'If InStr(UCase(cARQ), "PGSQL") > 0 Or InStr(UCase(cARQ), "POSTGRESQL") > 0 Then pgSetValuesrc6 oDB

    ' Abertura do Recordset no RC6:
    ' No RC6, o método é oDB.OpenRecordset(cSQL)
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
                aRETU(x) = aRETU(x) + eVAL
            Next x
            oRS.MoveNext
        Wend
    End If
    
    ' Limpeza (RC6 não usa .Close para conexão, apenas para Recordset)
    
    Set oRS = Nothing
    Set oDB = Nothing
    
    SomaSQLrc6 = aRETU
    Exit Function

errhandler:
    ' Ajustado para Logar erros do RC6
    SayErro "Erro SomaSQLrc6:" & vbCrLf & cARQ & vbCrLf & cSQL & vbCrLf & Err.Description
    If Not oRS Is Nothing Then Set oRS = Nothing
    If Not oDB Is Nothing Then Set oDB = Nothing
    SomaSQLrc6 = aRETU
End Function
Public Function PegSQLDelirc6(ByVal cARQ As String, ByVal cSQL As String, _
                              ByVal aCAM As Variant, Optional ByVal cDELI As String = ",", _
                              Optional ByVal aPAD As Variant = "", Optional ByVal aFOR As Variant = "") As Variant

    Dim oDB As RC6.cConnection
    Dim oRS As RC6.cRecordset
    Dim x As Integer
    Dim nCAMPOS As Integer
    Dim aRETU As Variant
    Dim aARQ As Variant
    Dim aOPE As Variant
    Dim eVAL As Variant
    Dim lOPEN As Boolean
    Dim lRSOP As Boolean
    Dim carqcon As String

    On Error GoTo errhandler

    lOPEN = False
    lRSOP = False
    nCAMPOS = UBound(aCAM) + 1
    ReDim aRETU(nCAMPOS - 1)
    For x = 0 To nCAMPOS - 1
        aRETU(x) = ""
    Next x

    If Not FileConnExist(SomaExt(cARQ), True) Then
        PegSQLDelirc6 = aRETU
        Exit Function
    End If

    aARQ = TipoConn(cARQ, , , False)
    carqcon = aARQ(1)

    ' Inicialização do RC6 (Substitui o .Open e propriedades ADO)
    Set oDB = New_c.Connection(carqcon)
    lOPEN = True

    ' Ajustes de ambiente
    ConfigurarConexaoRC6 oDB, cARQ
    'If InStr(cARQ, "VFPOLEDB") > 0 Then VFPSetValuesrc6 oDB
    'If InStr(UCase(cARQ), "PGSQL") > 0 Or InStr(UCase(cARQ), "POSTGRESQL") > 0 Then pgSetValuesrc6 oDB

    ' Abertura do Recordset (RC6 usa apenas o SQL e a Conexão)
    ' O método OpenRecordset é o padrão do RC6
    Set oRS = oDB.OpenRecordset(cSQL)
    lRSOP = True

    If Not oRS.EOF Then
        While Not oRS.EOF
            For x = 0 To nCAMPOS - 1
                aOPE = SepSqlOpe(aCAM(x))
                
                ' Extração do valor
                If aOPE(0) = "" Or aOPE(1) = "" Or aOPE(2) = "" Then
                    eVAL = oRS(aCAM(x)).Value
                Else
                    eVAL = MathOper(oRS(aOPE(1)).Value, oRS(aOPE(2)).Value, aOPE(0))
                End If
                
                ' Tratamentos (Pad, For, FVar)
                If IsNull(eVAL) Then
                    If IsArray(aPAD) Then eVAL = aPAD(x) Else eVAL = aPAD
                End If
                
                If IsArray(aFOR) Then
                    If IsArray(aPAD) Then eVAL = FVar(eVAL, aFOR(x), aPAD(x)) Else eVAL = FVar(eVAL, aFOR(x))
                End If
                
                aRETU(x) = aRETU(x) & FixStr(eVAL)
            Next x
            
            oRS.MoveNext
            
            ' Adiciona o delimitador se não for o último
            If Not oRS.EOF Then
                For x = 0 To nCAMPOS - 1
                    aRETU(x) = aRETU(x) & cDELI
                Next x
            End If
        Wend
    End If

    ' Finalização (Lembre-se: oDB não tem .Close)
    
    Set oRS = Nothing
    Set oDB = Nothing
    
    PegSQLDelirc6 = aRETU
    Exit Function

errhandler:
    SayErro "Peg SQL DELI RC6:" & vbCrLf & cARQ & vbCrLf & cSQL & vbCrLf & Err.Description
    If Not oRS Is Nothing Then Set oRS = Nothing
    If Not oDB Is Nothing Then Set oDB = Nothing
    PegSQLDelirc6 = aRETU
End Function
Public Function GrvSQLrc6(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, ByVal aCAM As Variant, _
                          ByVal aVAL As Variant, ByVal aFOR As Variant, Optional ByVal nSTARITEM As Long = 0) As Boolean
    Dim oDB As RC6.cConnection
    Dim oRS As RC6.cRecordset
    Dim x As Long
    Dim eVAL As Variant
    Dim aOPE As Variant
    Dim aARQ As Variant
    Dim carqcon As String
    
    On Error GoTo errhandler
    GrvSQLrc6 = False

    ' 1. Identifica o dialeto/conexão
    aARQ = TipoConn(cARQ)
    carqcon = aARQ(1)
    Set oDB = New_c.Connection(carqcon)

    ' 2. Ajustes de ambiente
    ConfigurarConexaoRC6 oDB, cARQ
    'If InStr(cARQ, "VFPOLEDB") > 0 Then VFPSetValuesrc6 oDB
    'If InStr(UCase(cARQ), "PGSQL") > 0 Or InStr(UCase(cARQ), "POSTGRESQL") > 0 Then pgSetValuesrc6 oDB

    ' 3. Abertura do Recordset RC6
    Set oRS = oDB.OpenRecordset(cSQL)

    If oRS.EOF Then
        Set oRS = Nothing: Set oDB = Nothing
        Exit Function
    End If

    While Not oRS.EOF
        For x = nSTARITEM To nITEM - 1
            ' Resolve o valor
            If IsNull(aVAL(x)) Then
                oRS(aCAM(x)).Value = Null
            Else
                aOPE = SepSqlOpe(aCAM(x))
                If aOPE(0) = "" Then
                    eVAL = aVAL(x)
                Else
                    eVAL = MathOper(oRS(aOPE(1)).Value, oRS(aOPE(2)).Value, aOPE(0))
                End If

                ' --- INTEGRAÇÃO COM DIALETO DE DATA ---
                If Mid(aFOR(x), 1, 1) = "D" Then
                    ' Usa a função dialetoDataParaSQL do seu sqldialeto.bas
                    ' aARQ(2) contém o identificador do tipo de banco
                    oRS(aCAM(x)).Value = dialetoDataParaSQL(eVAL, aARQ(2))
                Else
                    ' Tratamento padrão (Texto/Número)
                    oRS(aCAM(x)).Value = FVar(eVAL, aFOR(x), "")
                End If
            End If
        Next x
        
        ' Gravação em lote (nativa RC6)
        oRS.UpdateBatch
        oRS.MoveNext
    Wend

    Set oRS = Nothing
    Set oDB = Nothing
    GrvSQLrc6 = True
    Exit Function

errhandler:
    SayErro "Erro na Gravação (RC6):" & vbCrLf & cARQ & vbCrLf & cSQL & vbCrLf & Err.Description
    If Not oRS Is Nothing Then Set oRS = Nothing
    If Not oDB Is Nothing Then Set oDB = Nothing
    GrvSQLrc6 = False
End Function

Public Function ApagaSQLprc6(ByVal cARQ As String, ByVal cSQL As String, Optional ByVal cTEXTO As String = "Confirme Exclusão") As Boolean
  ApagaSQLprc6 = False
  If MDG(cTEXTO, "Exclusão Registro") Then
    ApagaSQLprc6 = APAGASQLrc6(cARQ, cSQL)
  End If
End Function


Public Function IncluiSQLrc6(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, _
                             Optional ByVal aCAM As Variant = 0, Optional ByVal aVAL As Variant = 0, _
                             Optional ByVal lCHECK As Boolean = False, _
                             Optional ByVal lMES As Boolean = True, _
                             Optional ByVal aIDDES As Variant = 0) As Boolean
    Dim oDB As RC6.cConnection
    Dim oRS As RC6.cRecordset
    Dim x As Long
    Dim aRETU As Variant
    Dim lTEM As Boolean
    Dim lRETU As Boolean
    Dim carqcon As String
    
    IncluiSQLrc6 = False
    On Error GoTo errhandler

    lTEM = False
    lRETU = False
    aRETU = TipoConn(cARQ) ' Retorna Array (Tag, StringConn, TipoBanco)
    carqcon = aRETU(1)
    ' Inicializa Conexão RC6
    Set oDB = New_c.Connection(carqcon)

    ' Ajustes de ambiente
    ConfigurarConexaoRC6 oDB, carqcon
    'If InStr(aRETU(1), "VFPOLEDB") > 0 Then VFPSetValuesrc6 oDB
    'If InStr(UCase(aRETU(1)), "PGSQL") > 0 Or InStr(UCase(aRETU(1)), "POSTGRESQL") > 0 Then pgSetValuesrc6 oDB

    ' Abertura do Recordset (O RC6 abre como ForwardOnly/ReadOnly se não especificado,
    ' mas como vamos dar AddNew, ele precisa ser editável)
    Set oRS = oDB.OpenRecordset(cSQL)

    ' Verifica se já existe (Check)
    If lCHECK And Not oRS.EOF Then
        lTEM = True
    End If

    If Not lTEM Then
        oRS.AddNew
        For x = 0 To nITEM - 1
            If Not IsNull(aVAL(x)) Then
                ' Tratamento de Datas para o RC6/Dialetos
                If Mid(aCAM(x), 1, 1) = "D" Then ' Ou a lógica de aFOR que você usa
                    oRS(aCAM(x)).Value = dialetoDataParaSQL(aVAL(x), aRETU(2))
                Else
                    oRS(aCAM(x)).Value = aVAL(x)
                End If
            End If
        Next x

        ' Gravação
        oRS.UpdateBatch
        
        ' Recupera IDs se necessário
        If Not IsNumeric(aIDDES) Then
            ' No RC6, o registro inserido está posicionado
            Dim nCAMPOS As Long
            nCAMPOS = UBound(aIDDES)
            ReDim eRETU01(nCAMPOS)
            For x = 0 To nCAMPOS
                eRETU01(x) = oRS(aIDDES(x)).Value
            Next x
        End If
        
        lRETU = True
    End If

    ' Limpeza (RC6 não usa .Close na conexão, apenas no Recordset)
    Set oRS = Nothing
    Set oDB = Nothing

    If lCHECK And lTEM And lMES Then
        Alert "Item já Cadastrado Com Esta Chave"
    End If

    IncluiSQLrc6 = lRETU
    Exit Function

errhandler:
    SayErro "Erro na Inclusão (RC6):" & vbCrLf & cARQ & vbCrLf & cSQL & vbCrLf & Err.Description
    If Not oRS Is Nothing Then Set oRS = Nothing
    If Not oDB Is Nothing Then Set oDB = Nothing
    IncluiSQLrc6 = False
End Function

Public Function PegSQLrc6(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, ByVal aCAM As Variant, ByVal aFOR As Variant, ByVal aPAD As Variant) As Variant
    Dim oDB As RC6.cConnection
    Dim oRS As RC6.cRecordset
    Dim x As Long
    Dim aRETU As Variant
    Dim aCON As Variant
    Dim aOPE As Variant
    Dim eVAL As Variant
    Dim carqcon As String

    On Error GoTo errhandler
    
    ' Inicializa o array de retorno
    ReDim aRETU(0 To nITEM - 1)
    
    ' Identifica a conexão
    aCON = TipoConn(cARQ, , , False)
    
    ' Ajuste de dialeto PGSQL (se necessário)
    If aCON(2) = "PGSQL" Then
        cSQL = SQLPGSQLDOUBLEQUOTES(cSQL)
    End If

    carqcon = aCON(1)
    ' Inicialização do RC6 (cConnection não precisa de Open com timeout)
    Set oDB = New_c.Connection(carqcon)

    ' Ajustes de ambiente
    ConfigurarConexaoRC6 oDB, carqcon
   ' If InStr(cARQ, "VFPOLEDB") > 0 Then VFPSetValuesrc6 oDB
   ' If InStr(UCase(cARQ), "PGSQL") > 0 Or InStr(UCase(cARQ), "POSTGRESQL") > 0 Then pgSetValuesrc6 oDB

    ' Abertura do Recordset no RC6
    Set oRS = oDB.OpenRecordset(cSQL)

    If Not oRS.EOF Then
        For x = 0 To nITEM - 1
            aOPE = SepSqlOpe(aCAM(x))
            
            ' Extração do valor
            If aOPE(0) = "" Or aOPE(1) = "" Or aOPE(2) = "" Then
                eVAL = oRS(aCAM(x)).Value
            Else
                eVAL = MathOper(oRS(aOPE(1)).Value, oRS(aOPE(2)).Value, aOPE(0))
            End If

            ' Formatação (FVar) e tratamento de Nulo
            If IsNull(eVAL) Then
                aRETU(x) = aPAD(x)
            Else
                aRETU(x) = FVar(eVAL, aFOR(x), aPAD(x))
            End If
        Next x
    Else
        ' Se não encontrar registros, retorna o padrão
        aRETU = aPAD
    End If

    ' Fechamento e Limpeza
    Set oRS = Nothing
    Set oDB = Nothing
    
    ' Retorno da função
    PegSQLrc6 = aRETU
    Exit Function

errhandler:
    ' Captura de erro simplificada para RC6
    Dim cERRO As String
    cERRO = "Peg SQL RC6:" & vbCrLf & cARQ & vbCrLf & cSQL & vbCrLf & "Erro: " & Err.Description
    
    ' Se ocorrer erro em um campo, tenta usar o valor padrão
    If Err.Number = 3265 Or Err.Number = 5 Or Err.Number = 94 Then
        If x >= 0 And x < nITEM Then aRETU(x) = aPAD(x)
        Resume Next
    End If
    
    SayErro cERRO
    If Not oRS Is Nothing Then Set oRS = Nothing
    If Not oDB Is Nothing Then Set oDB = Nothing
    PegSQLrc6 = aPAD
End Function
Public Function PegCountSQLrc6(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
  PegCountSQLrc6 = PegOperSQLrc6(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, "COUNT")
End Function
Public Function PegMINSQLrc6(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
  PegMINSQLrc6 = PegOperSQLrc6(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, "MIN")
End Function
Public Function PegMAXSQLrc6(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
  PegMAXSQLrc6 = PegOperSQLrc6(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, "MAX")
End Function
Public Function PegSUMSQLrc6(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
  PegSUMSQLrc6 = PegOperSQLrc6(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, "SUM")
End Function
Public Function PegCampoSQLrc6(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
  PegCampoSQLrc6 = PegOperSQLrc6(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, "CAMPO")
End Function
Public Function PegOperSQLrc6(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant, ByVal coper As String) As Variant
  Dim aRETU As Variant
  Dim cSQL As String
  PegOperSQLrc6 = eDEFAULT


  If coper = "CAMPO" Then
    If Len(cCAMPO) = 0 Then  'Passado a string ja com o campo select numero from tabela where numero=999999
      cTABLEWHERE = UCase(cTABLEWHERE)
      cSQL = Replace(cTABLEWHERE, " FROM ", " AS CAMPO FROM ")  'inclui as campo variavel padrao no pegsql abaixo
    Else
      cSQL = "SELECT " & cCAMPO & " AS CAMPO FROM " & cTABLEWHERE
    End If
  Else
    cSQL = "SELECT " & coper & "(" & cCAMPO & ") AS CAMPO FROM " & cTABLEWHERE
  End If
  If coper = "SUM" Or coper = "COUNT" Or coper = "MAX" Or coper = "MIN" Or IsNumeric(eDEFAULT) Then
    aRETU = PegSQLrc6(cARQ, cSQL, 1, Array("CAMPO"), Array("N"), Array(eDEFAULT))  ''array retorno tipo N numerico
  Else
    aRETU = PegSQLrc6(cARQ, cSQL, 1, Array("CAMPO"), Array(""), Array(eDEFAULT))
  End If
  If lRETU Then
    PegOperSQLrc6 = aRETU(0)
  End If
End Function
Public Function PegUltSQLrc6(ByVal cARQ As String, ByVal cSQL As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    Dim oDB As RC6.cConnection
    Dim oRS As RC6.cRecordset
    
    On Error GoTo errhandler
    
    ' Valor padrão inicial
    PegUltSQLrc6 = eDEFAULT
    
    ' Inicializa a conexão (O RC6 abre a conexão ao instanciar com o caminho)
    Set oDB = New_c.Connection(cARQ)
    
    ' Abre o Recordset nativo do RC6
    ' O RC6 gerencia o cursor automaticamente conforme o driver (SQLite/Outros)
    Set oRS = oDB.OpenRecordset(cSQL)
    
    If Not oRS.EOF Then
        ' Move para o último registro para pegar o "último" valor
        oRS.MoveLast
        
        ' Verifica se o campo não é nulo antes de atribuir
        If Not IsNull(oRS(cCAMPO).Value) Then
            PegUltSQLrc6 = oRS(cCAMPO).Value
        End If
    End If
    
    ' Limpeza (RC6 não tem .Close em conexão)
    Set oRS = Nothing
    Set oDB = Nothing
    
    Exit Function

errhandler:
    ' Registro de erro centralizado
    SayErro "PegUltSQLrc6 - Erro: " & Err.Description & vbCrLf & _
            "Arquivo: " & cARQ & vbCrLf & _
            "SQL: " & cSQL
            
    If Not oRS Is Nothing Then Set oRS = Nothing
    If Not oDB Is Nothing Then Set oDB = Nothing
    
    ' Retorna o padrão em caso de erro
    PegUltSQLrc6 = eDEFAULT
End Function
Public Function SQLMoveRegrc6(ByVal cARQORI As String, _
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

    Dim oDB As RC6.cConnection
    Dim oDBDES As RC6.cConnection
    Dim oRS As RC6.cRecordset
    Dim oRSDES As RC6.cRecordset
    Dim aRetuOri As Variant
    Dim aRetuDes As Variant
    Dim x As Integer
    Dim nCAMPOS As Integer
    Dim aVALORI As Variant
    Dim aRETUID As Variant
    Dim cCONORI As String
    Dim cCONDES As String

    On Error GoTo errhandler
    SQLMoveRegrc6 = False

    ' 1. Identifica Conexões
    aRetuOri = TipoConn(cARQORI)
    aRetuDes = TipoConn(cARQDES)
    
   cCONORI = aRetuOri(1)
   cCONDES = aRetuDes(1)
   
   
    ' 2. Inicializa Conexões RC6 (New_c.Connection já abre a conexão)
    Set oDB = New_c.Connection(cCONORI)
    Set oDBDES = New_c.Connection(cCONDES)

    ' Ajustes de ambiente
    ConfigurarConexaoRC6 oDB, cCONORI
    ConfigurarConexaoRC6 oDBDES, cCONDES
    'If InStr(aRetuDes(1), "VFPOLEDB") > 0 Then VFPSetValuesrc6 oDBDES
    'If InStr(UCase(aRetuDes(1)), "PGSQL") > 0 Or InStr(UCase(aRetuDes(1)), "POSTGRESQL") > 0 Then pgSetValuesrc6 oDBDES

    ' 3. Abre Recordsets
    Set oRS = oDB.OpenRecordset(cSQLORI)
    Set oRSDES = oDBDES.OpenRecordset(cSQLDES)

    ' 4. Processamento
    If Not oRS.EOF Then
        While Not oRS.EOF
            ' Copia valores da origem para array
            If Not IsNumeric(aCAMORI) Then
                nCAMPOS = UBound(aCAMORI)
                ReDim aVALORI(nCAMPOS)
                For x = 0 To nCAMPOS
                    aVALORI(x) = oRS(aCAMORI(x)).Value
                Next x
            End If

            ' Operação Origem (Delete)
            If InStr(cOPEORI, "DEL") > 0 Then oRS.Delete

            ' Operação Destino (Append)
            If InStr(cOPEDES, "INC") > 0 Then
                oRSDES.AddNew
            End If

            ' Preenche campos do destino
            If Not IsNumeric(aCAMDES) Then
                For x = 0 To UBound(aCAMDES)
                    oRSDES(aCAMDES(x)).Value = aVALORI(x)
                Next x
            End If
            
            ' Preenche campos externos (Outros)
            If Not IsNumeric(aOUTDES) Then
                For x = 0 To UBound(aOUTDES)
                    oRSDES(aOUTDES(x)).Value = aOUTORI(x)
                Next x
            End If

            ' Gravação em Lote
            oRSDES.UpdateBatch

            ' Captura IDs inseridos
            If Not IsNumeric(aIDDES) Then
                nCAMPOS = UBound(aIDDES)
                ReDim aRETUID(nCAMPOS)
                For x = 0 To nCAMPOS
                    aRETUID(x) = oRSDES(aIDDES(x)).Value
                Next x
                eRETU01 = aRETUID
            End If

            oRS.MoveNext
        Wend
    End If

    ' Limpeza
     Set oRS = Nothing
     Set oRSDES = Nothing
    Set oDB = Nothing
    Set oDBDES = Nothing

    SQLMoveRegrc6 = True
    Exit Function

errhandler:
    SayErro "Erro em SQLMoveRegrc6:" & vbCrLf & Err.Description
    If Not oRS Is Nothing Then Set oRS = Nothing
    If Not oRSDES Is Nothing Then Set oRSDES = Nothing
    Set oDB = Nothing: Set oDBDES = Nothing
    SQLMoveRegrc6 = False
End Function

'Public Function CriaMdbAccessRC6(ByVal cCaminho As String) As Boolean
    ' Usamos Late Binding para não precisar da referência ao ADOX no projeto RC6
 '   Dim cat As Object
  '  On Error GoTo Erro
    
   ' Set cat = CreateObject("ADOX.Catalog")
    ' Cria o arquivo .accdb (o Provider ACE.OLEDB.12.0 cria o arquivo fisicamente)
    'cat.Create "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & cCaminho
    
    'CriaMdbAccessRC6 = True
    'Exit Function
'Erro:
 '   CriaMdbAccessRC6 = False
'End Function

'Public Function CompactarMdbAccessRC6(ByVal cOrigem As String, ByVal cDestino As String) As Boolean
'    Dim jro As Object
'    On Error GoTo Erro
    
    ' O JRO.JetEngine faz o trabalho pesado de compactação do Access
'    Set jro = CreateObject("JRO.JetEngine")
    
    ' Compacta a origem para um arquivo temporário (cDestino)
 '   jro.CompactDatabase "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & cOrigem, _
                        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & cDestino & ";Jet OLEDB:Engine Type=6"
    
  '  CompactarMdbAccessRC6 = True
   ' Exit Function
'Erro:
 '   CompactarMdbAccessRC6 = False
'End Function

' ==============================================================================
' FUNÇÃO: ConfigurarConexaoRC6
' Objetivo: Orquestrador central que aplica os SETs de performance e ambiente
'           baseado na string de conexão fornecida.
' ==============================================================================
Public Function ConfigurarConexaoRC6(ByRef oCON As Object, ByVal cStringConexao As String) As Boolean
    Dim cBUSCA As String
    cBUSCA = UCase(cStringConexao)
    
    On Error GoTo ErroConfig
    
    ' --- 1. VFP / FoxPro ---
    If InStr(cBUSCA, "VFPOLEDB") > 0 Then
        VFPSetValuesrc6 oCON
    
    ' --- 2. PostgreSQL ---
    ElseIf InStr(cBUSCA, "PGSQL") > 0 Or InStr(cBUSCA, "POSTGRESQL") > 0 Then
        pgSetValuesrc6 oCON
        
    ' --- 3. SQLite (Nativo RC6) ---
    ElseIf InStr(cBUSCA, "SQLITE") > 0 Then
        ' Assumindo que oCON seja RC6.cConnection ou compatível
        SQLiteSetValuesrc6 oCON
    End If
    
    ConfigurarConexaoRC6 = True
    Exit Function

ErroConfig:
    ConfigurarConexaoRC6 = False
End Function

Public Function SQLiteSetValuesrc6(ByRef oCON As Object) As Boolean
    On Error GoTo ErroSQLite
    SQLiteSetValuesrc6 = False
    
    If oCON Is Nothing Then Exit Function
    
    ' 1. Armazena temporários na memória
    oCON.Execute "PRAGMA temp_store = MEMORY;"
    
    ' 2. Cache de 2000 páginas (aprox. 8MB se a página for 4KB)
    oCON.Execute "PRAGMA cache_size = 2000;"
    
    ' 3. Modo WAL: Essencial para performance em concorrência
    oCON.Execute "PRAGMA journal_mode = WAL;"
    
    ' 4. Sincronização NORMAL (seguro com WAL e muito mais rápido)
    oCON.Execute "PRAGMA synchronous = NORMAL;"
    
    ' 5. Auto Vacuum Incremental
    oCON.Execute "PRAGMA auto_vacuum = INCREMENTAL;"
    
    SQLiteSetValuesrc6 = True
    Exit Function

ErroSQLite:
    ' Aqui você pode registrar o erro se desejar
    SQLiteSetValuesrc6 = False
End Function
'Public Function RepararMdbAccessRC6(ByVal cCaminho As String) As Boolean
    ' A estratégia padrão para reparar MDB/ACCDB é tentar compactá-lo.
    ' Se houver corrupção física, a função de compactação retornará erro.
 '   Dim cTemp As String
  '  cTemp = cCaminho & ".tmp"
    
   ' If CompactarMdbAccessRC6(cCaminho, cTemp) Then
        ' Se deu certo, substitui o corrompido pelo reparado
    '    Kill cCaminho
     '   Name cTemp As cCaminho
     '   RepararMdbAccessRC6 = True
    'Else
     '   RepararMdbAccessRC6 = False
    'End If
'End Function
