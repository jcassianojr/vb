Attribute VB_Name = "SqlTC6SQLite"
Option Explicit
'TC6 https://github.com/wqweto/TC6
Public Function ComandoTC6(ByVal cARQ As String, ByVal cSQL As String) As Boolean
    Dim oDB As New TC6SQLITE.cConnection

    On Error GoTo TrataErro
    ComandoTC6 = False
    
    ' Ajustes de dialeto (Mantido conforme original)
    If InStr(cSQL, "CURRENTDATETIME") > 0 Then
          cSQL = Replace(cSQL, "CURRENTDATETIME", " current_timestamp ")
    End If
    
    cARQ = LimpaTag(cARQ)
    ' Inicialização simplificada no TC6
    oDB.OpenDB cARQ  'New_c.Connection(cCONN)
    
    ' O método Execute da conexão substitui a necessidade do objeto Command
    ' para comandos simples (sem parâmetros).
    oDB.Execute cSQL
    
    ' Fechamento simplificado (apenas limpa o objeto)
    Set oDB = Nothing
    
    ComandoTC6 = True
    Exit Function

TrataErro:
    SayErro "SQL TC6 Comando:" & vbCrLf & cARQ & vbCrLf & cSQL & vbCrLf & Err.Description
    If Not oDB Is Nothing Then Set oDB = Nothing
    ComandoTC6 = False
End Function

Public Function APAGASQLTC6(ByVal cARQ As String, ByVal cSQL As String) As Boolean
  Dim nPOS As Integer
  Dim cNOME As String
  'select * from tabela where campo=valorcampo ....
  'delete * from tabela where campo=valorcampo ....
  'delete from tabela where campo=valorcampo ....
  'from tabela where campo=valorcampo ....
  'todos os csql viram delete from tabela where campo=valorcampo ....
  'requer from e where no csql


  APAGASQLTC6 = False
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
    APAGASQLTC6 = ComandoTC6(cARQ, cSQL)
  Else
    If InStr(cNOME, "DELETE") > 0 Then  'so executa se tiver delete
      APAGASQLTC6 = ComandoTC6(cARQ, cSQL)
    End If
  End If
End Function

Public Function SomaSQLTC6(ByVal cARQ As String, ByVal cSQL As String, ByVal aCAM As Variant) As Variant
    Dim oDB As New TC6SQLITE.cConnection
    Dim oRS As TC6SQLITE.cRecordset
    Dim lOPEN As Boolean
    Dim lRSOP As Boolean
    Dim x As Integer
    Dim nCAMPOS As Integer
    Dim aRETU As Variant
    Dim aOPE As Variant
    Dim eVAL As Variant

    On Error GoTo errhandler

    lOPEN = False
    lRSOP = False
    nCAMPOS = UBound(aCAM) + 1
    ReDim aRETU(nCAMPOS - 1) ' Ajustado para base 0
    For x = 0 To nCAMPOS - 1
        aRETU(x) = 0
    Next x

    If Not FileConnExist(SomaExt(cARQ), True) Then
        SomaSQLTC6 = aRETU
        Exit Function
    End If

    cARQ = LimpaTag(cARQ)
    oDB.OpenDB cARQ
    
    lOPEN = True
  
    SqliteSetValuesTC6 oDB

    ' Abertura do Recordset no TC6:
    ' No TC6, o método é oDB.OpenRecordset(cSQL)
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
    
    ' Limpeza (TC6 não usa .Close para conexão, apenas para Recordset)
    
    Set oRS = Nothing
    Set oDB = Nothing
    
    SomaSQLTC6 = aRETU
    Exit Function

errhandler:
    ' Ajustado para Logar erros do TC6
    SayErro "Erro SomaSQLTC6:" & vbCrLf & cARQ & vbCrLf & cSQL & vbCrLf & Err.Description
    If Not oRS Is Nothing Then Set oRS = Nothing
    If Not oDB Is Nothing Then Set oDB = Nothing
    SomaSQLTC6 = aRETU
End Function
Public Function PegSQLDeliTC6(ByVal cARQ As String, ByVal cSQL As String, _
                              ByVal aCAM As Variant, Optional ByVal cDELI As String = ",", _
                              Optional ByVal aPAD As Variant = "", Optional ByVal aFOR As Variant = "") As Variant

    Dim oDB As New TC6SQLITE.cConnection
    Dim oRS As TC6SQLITE.cRecordset
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
        PegSQLDeliTC6 = aRETU
        Exit Function
    End If

    cARQ = LimpaTag(cARQ)
    oDB.OpenDB cARQ
    
    lOPEN = True

    ' Ajustes de ambiente
    SqliteSetValuesTC6 oDB

    ' Abertura do Recordset (TC6 usa apenas o SQL e a Conexão)
    ' O método OpenRecordset é o padrão do TC6
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
    
    PegSQLDeliTC6 = aRETU
    Exit Function

errhandler:
    SayErro "Peg SQL DELI TC6:" & vbCrLf & cARQ & vbCrLf & cSQL & vbCrLf & Err.Description
    If Not oRS Is Nothing Then Set oRS = Nothing
    If Not oDB Is Nothing Then Set oDB = Nothing
    PegSQLDeliTC6 = aRETU
End Function
Public Function GrvSQLTC6(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, ByVal aCAM As Variant, _
                          ByVal aVAL As Variant, ByVal aFOR As Variant, Optional ByVal nSTARITEM As Long = 0) As Boolean
    Dim oDB As New TC6SQLITE.cConnection
    Dim oRS As TC6SQLITE.cRecordset
    Dim x As Long
    Dim eVAL As Variant
    Dim aOPE As Variant
    
    On Error GoTo errhandler
    GrvSQLTC6 = False

    
    cARQ = LimpaTag(cARQ)
    oDB.OpenDB cARQ
    
    ' 2. Ajustes de ambiente
    SqliteSetValuesTC6 oDB

    ' 3. Abertura do Recordset TC6
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
                    oRS(aCAM(x)).Value = dialetoDataParaSQL(eVAL, "SQLITE")
                Else
                    ' Tratamento padrão (Texto/Número)
                    oRS(aCAM(x)).Value = FVar(eVAL, aFOR(x), "")
                End If
            End If
        Next x
        
        ' Gravação em lote (nativa TC6)
        oRS.UpdateBatch
        oRS.MoveNext
    Wend

    Set oRS = Nothing
    Set oDB = Nothing
    GrvSQLTC6 = True
    Exit Function

errhandler:
    SayErro "Erro na Gravação (TC6):" & vbCrLf & cARQ & vbCrLf & cSQL & vbCrLf & Err.Description
    If Not oRS Is Nothing Then Set oRS = Nothing
    If Not oDB Is Nothing Then Set oDB = Nothing
    GrvSQLTC6 = False
End Function

Public Function ApagaSQLpTC6(ByVal cARQ As String, ByVal cSQL As String, Optional ByVal cTEXTO As String = "Confirme Exclusão") As Boolean
  ApagaSQLpTC6 = False
  If MDG(cTEXTO, "Exclusão Registro") Then
    ApagaSQLpTC6 = APAGASQLTC6(cARQ, cSQL)
  End If
End Function


Public Function IncluiSQLTC6(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, _
                             Optional ByVal aCAM As Variant = 0, Optional ByVal aVAL As Variant = 0, _
                             Optional ByVal lCHECK As Boolean = False, _
                             Optional ByVal lMES As Boolean = True, _
                             Optional ByVal aIDDES As Variant = 0) As Boolean
    Dim oDB As New TC6SQLITE.cConnection
    Dim oRS As TC6SQLITE.cRecordset
    Dim x As Long
    Dim aRETU As Variant
    Dim lTEM As Boolean
    
    IncluiSQLTC6 = False
    On Error GoTo errhandler

    lTEM = False
    lRETU = False
    
    cARQ = LimpaTag(cARQ)
    oDB.OpenDB cARQ
    
    ' Ajustes de ambiente
    SqliteSetValuesTC6 oDB
    
    
    Set oRS = oDB.OpenRecordset(cSQL)

    ' Verifica se já existe (Check)
    If lCHECK And Not oRS.EOF Then
        lTEM = True
    End If

    If Not lTEM Then
        oRS.AddNew
        For x = 0 To nITEM - 1
            If Not IsNull(aVAL(x)) Then
                ' Tratamento de Datas para o TC6/Dialetos
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
            ' No TC6, o registro inserido está posicionado
            Dim nCAMPOS As Long
            nCAMPOS = UBound(aIDDES)
            ReDim eRETU01(nCAMPOS)
            For x = 0 To nCAMPOS
                eRETU01(x) = oRS(aIDDES(x)).Value
            Next x
        End If
        
        lRETU = True
    End If

    ' Limpeza (TC6 não usa .Close na conexão, apenas no Recordset)
    Set oRS = Nothing
    Set oDB = Nothing

    If lCHECK And lTEM And lMES Then
        Alert "Item já Cadastrado Com Esta Chave"
    End If

    IncluiSQLTC6 = lRETU
    Exit Function

errhandler:
    SayErro "Erro na Inclusão (TC6):" & vbCrLf & cARQ & vbCrLf & cSQL & vbCrLf & Err.Description
    If Not oRS Is Nothing Then Set oRS = Nothing
    If Not oDB Is Nothing Then Set oDB = Nothing
    IncluiSQLTC6 = False
End Function

Public Function PegSQLTC6(ByVal cARQ As String, ByVal cSQL As String, ByVal nITEM As Long, ByVal aCAM As Variant, ByVal aFOR As Variant, ByVal aPAD As Variant) As Variant
    Dim oDB As New TC6SQLITE.cConnection
    Dim oRS As TC6SQLITE.cRecordset
    Dim x As Long
    Dim aRETU As Variant
    Dim aOPE As Variant
    Dim eVAL As Variant

    On Error GoTo errhandler
    
    ' Inicializa o array de retorno
    ReDim aRETU(0 To nITEM - 1)
    
    cARQ = LimpaTag(cARQ)
   oDB.OpenDB cARQ
   

    ' Ajustes de ambiente
    SqliteSetValuesTC6 oDB

    ' Abertura do Recordset no TC6
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
    PegSQLTC6 = aRETU
    Exit Function

errhandler:
    ' Captura de erro simplificada para TC6
    Dim cERRO As String
    cERRO = "Peg SQL TC6:" & vbCrLf & cARQ & vbCrLf & cSQL & vbCrLf & "Erro: " & Err.Description
    
    ' Se ocorrer erro em um campo, tenta usar o valor padrão
    If Err.Number = 3265 Or Err.Number = 5 Or Err.Number = 94 Then
        If x >= 0 And x < nITEM Then aRETU(x) = aPAD(x)
        Resume Next
    End If
    
    SayErro cERRO
    If Not oRS Is Nothing Then Set oRS = Nothing
    If Not oDB Is Nothing Then Set oDB = Nothing
    PegSQLTC6 = aPAD
End Function
Public Function PegCountSQLTC6(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
  PegCountSQLTC6 = PegOperSQLTC6(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, "COUNT")
End Function
Public Function PegMINSQLTC6(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
  PegMINSQLTC6 = PegOperSQLTC6(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, "MIN")
End Function
Public Function PegMAXSQLTC6(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
  PegMAXSQLTC6 = PegOperSQLTC6(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, "MAX")
End Function
Public Function PegSUMSQLTC6(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
  PegSUMSQLTC6 = PegOperSQLTC6(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, "SUM")
End Function
Public Function PegCampoSQLTC6(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
  PegCampoSQLTC6 = PegOperSQLTC6(cARQ, cTABLEWHERE, cCAMPO, eDEFAULT, "CAMPO")
End Function
Public Function PegOperSQLTC6(ByVal cARQ As String, ByVal cTABLEWHERE As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant, ByVal coper As String) As Variant
  Dim aRETU As Variant
  Dim cSQL As String
  PegOperSQLTC6 = eDEFAULT


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
    aRETU = PegSQLTC6(cARQ, cSQL, 1, Array("CAMPO"), Array("N"), Array(eDEFAULT))  ''array retorno tipo N numerico
  Else
    aRETU = PegSQLTC6(cARQ, cSQL, 1, Array("CAMPO"), Array(""), Array(eDEFAULT))
  End If
  If lRETU Then
    PegOperSQLTC6 = aRETU(0)
  End If
End Function
Public Function PegUltSQLTC6(ByVal cARQ As String, ByVal cSQL As String, ByVal cCAMPO As String, ByVal eDEFAULT As Variant) As Variant
    Dim oDB As New TC6SQLITE.cConnection
    Dim oRS As TC6SQLITE.cRecordset
    
    On Error GoTo errhandler
    
    ' Valor padrão inicial
    PegUltSQLTC6 = eDEFAULT
    
    ' Inicializa a conexão (O TC6 abre a conexão ao instanciar com o caminho)
    cARQ = LimpaTag(cARQ)
    oDB.OpenDB cARQ
    
    ' Abre o Recordset nativo do TC6
    ' O TC6 gerencia o cursor automaticamente conforme o driver (SQLite/Outros)
    Set oRS = oDB.OpenRecordset(cSQL)
    
    If Not oRS.EOF Then
        ' Move para o último registro para pegar o "último" valor
        oRS.MoveLast
        
        ' Verifica se o campo não é nulo antes de atribuir
        If Not IsNull(oRS(cCAMPO).Value) Then
            PegUltSQLTC6 = oRS(cCAMPO).Value
        End If
    End If
    
    ' Limpeza (TC6 não tem .Close em conexão)
    Set oRS = Nothing
    Set oDB = Nothing
    
    Exit Function

errhandler:
    ' Registro de erro centralizado
    SayErro "PegUltSQLTC6 - Erro: " & Err.Description & vbCrLf & _
            "Arquivo: " & cARQ & vbCrLf & _
            "SQL: " & cSQL
            
    If Not oRS Is Nothing Then Set oRS = Nothing
    If Not oDB Is Nothing Then Set oDB = Nothing
    
    ' Retorna o padrão em caso de erro
    PegUltSQLTC6 = eDEFAULT
End Function
Public Function SQLMoveRegTC6(ByVal cARQORI As String, _
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

    Dim oDB As New TC6SQLITE.cConnection
    Dim oDBDES As New TC6SQLITE.cConnection
    Dim oRS As TC6SQLITE.cRecordset
    Dim oRSDES As TC6SQLITE.cRecordset
    Dim x As Integer
    Dim nCAMPOS As Integer
    Dim aVALORI As Variant
    Dim aRETUID As Variant

    On Error GoTo errhandler
    SQLMoveRegTC6 = False

    cARQORI = LimpaTag(cARQORI)
    cARQDES = LimpaTag(cARQDES)
    oDB.OpenDB cARQORI
    oDBDES.OpenDB cARQDES

    ' Ajustes de ambiente
    SqliteSetValuesTC6 oDB
    SqliteSetValuesTC6 oDBDES

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

    SQLMoveRegTC6 = True
    Exit Function

errhandler:
    SayErro "Erro em SQLMoveRegTC6:" & vbCrLf & Err.Description
    If Not oRS Is Nothing Then Set oRS = Nothing
    If Not oRSDES Is Nothing Then Set oRSDES = Nothing
    Set oDB = Nothing: Set oDBDES = Nothing
    SQLMoveRegTC6 = False
End Function
Public Function SqliteSetValuesTC6(ByRef oCON As Object) As Boolean
    On Error GoTo ErroSQLite
    SqliteSetValuesTC6 = False
    
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
    
    SqliteSetValuesTC6 = True
    Exit Function

ErroSQLite:
    ' Aqui você pode registrar o erro se desejar
    SqliteSetValuesTC6 = False
End Function
