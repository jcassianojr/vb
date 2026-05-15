Attribute VB_Name = "modMigracaoAccess"
Option Explicit

' ===========================================================================
' Módulo: modMigracaoAccess
' Finalidade: Comparar estrutura entre MDB e ACCDB e gerar script de correção
' Desenvolvido para: VB6 (Visual Basic 6.0)
' ===========================================================================

Public Sub GerarScriptCorrecao(ByVal sPathMDB As String, ByVal sPathACCDB As String, ByVal sPathSQL As String, ByVal bExecutarAoFinal As Boolean)
    Dim dbOrigem As DAO.Database
    Dim dbDestino As DAO.Database
    Dim tdfOrigem As DAO.TableDef
    Dim tdfDestino As DAO.TableDef
    Dim fld As DAO.Field
    Dim idx As DAO.Index
    Dim fFile As Integer
    Dim fso As Object
    
    Dim iContagemIndices As Integer
    Set fso = CreateObject("Scripting.FileSystemObject")
    iContagemIndices = 0
    
    
    On Error GoTo Err_Handle
    
    ' --- NOVA LÓGICA: VERIFICAÇÃO DO DESTINO ---
    If Not fso.FileExists(sPathACCDB) Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox("O arquivo de destino não existe:" & vbCrLf & sPathACCDB & _
                      vbCrLf & vbCrLf & "Deseja criar um banco de dados novo em branco?", _
                      vbQuestion + vbYesNo, "Destino não encontrado")
        
        If resp = vbYes Then
            If Not CriarBancoAcessoEmBranco(sPathACCDB) Then Exit Sub
        Else
            Exit Sub ' Usuário cancelou
        End If
    End If
    
    
    
    ' Abre os bancos de dados
    ' Nota: O motor ACE (para .accdb) é aberto corretamente via DAO 12.0+
    Set dbOrigem = OpenDatabase(sPathMDB)
    Set dbDestino = OpenDatabase(sPathACCDB, False, False)
    
    fFile = FreeFile
    Open sPathSQL For Output As #fFile
    
    Print #fFile, "-- ======================================================"
    Print #fFile, "-- SCRIPT DE CORREÇÃO DE ESTRUTURA (MDB -> ACCDB)"
    Print #fFile, "-- Gerado em: " & Now
    Print #fFile, "-- ======================================================"
    Print #fFile, ""

    For Each tdfOrigem In dbOrigem.TableDefs
        ' Ignora tabelas ocultas do sistema Access
        If Left(tdfOrigem.Name, 4) <> "MSys" Then
        
           ' RESET DO CONTADOR PARA CADA TABELA
            iContagemIndices = 0
            
            Debug.Print "Analisando tabela: " & tdfOrigem.Name
            
            On Error Resume Next
            Set tdfDestino = dbDestino.TableDefs(tdfOrigem.Name)
            
            If Err.Number <> 0 Then
                ' Se a tabela nem existir no destino
                Print #fFile, "-- ATENÇÃO: Tabela [" & tdfOrigem.Name & "] não encontrada no destino."
                Err.Clear
                
                Dim sSqlCreate As String
                sSqlCreate = GerarSqlCriacaoTabela(tdfOrigem) ' Nova função
    
                Print #fFile, "-- Criando tabela inexistente: " & tdfOrigem.Name
                Print #fFile, sSqlCreate
                
                
            Else
                ' --- 1. VERIFICAÇÃO DE CAMPOS ---
                For Each fld In tdfOrigem.Fields
                    Dim fldDest As DAO.Field
                    Dim sDefault As String
                    Dim sDefaultOrigem As String
                    sDefault = ""
                    Set fldDest = tdfDestino.Fields(fld.Name)
                    
                    sDefaultOrigem = fld.DefaultValue
                    
                    If Err.Number <> 0 Then
                    
                        ' CENÁRIO A: O campo não existe no destino
                        Err.Clear
                        Dim sSqlAdd As String
                        sSqlAdd = "ALTER TABLE [" & tdfOrigem.Name & "] ADD COLUMN [" & fld.Name & "] " & GetSQLType(fld)
                        
                        If sDefaultOrigem <> "" Then
                            sSqlAdd = sSqlAdd & " DEFAULT " & sDefaultOrigem
                        End If
                        
                        Print #fFile, sSqlAdd & ";"
                    Else
                    ' CENÁRIO B: O campo existe, mas vamos comparar o DefaultValue
                        ' Se a origem tem default e o destino não (ou se forem diferentes)
                        If sDefaultOrigem <> fldDest.DefaultValue Then
                            ' Gera alteração apenas para o valor padrão
                            ' Nota: No Access SQL, usamos SET DEFAULT
                            Print #fFile, "ALTER TABLE [" & tdfOrigem.Name & "] ALTER COLUMN [" & fld.Name & "] SET DEFAULT " & IIf(sDefaultOrigem = "", "NULL", sDefaultOrigem) & ";"
                        End If
                    End If
                Next fld
                
                ' --- 2. VERIFICAÇÃO DE ÍNDICES ---
                For Each idx In tdfOrigem.Indexes
                    ' Ignora índices criados automaticamente pelo Access para campos simples (opcional)
                    ' Se quiser ser rigoroso, verificamos todos.
                    iContagemIndices = iContagemIndices + 1
                    
                    Dim idxDest As DAO.Index
                    Set idxDest = tdfDestino.Indexes(idx.Name)
                    
                    If Err.Number <> 0 Then
                        ' Índice não existe: Gera comando de criação
                        Dim sCampos As String
                        Dim fldIdx As DAO.Field
                        
                        sCampos = ""
                        For Each fldIdx In idx.Fields
                            sCampos = sCampos & "[" & fldIdx.Name & "], "
                        Next
                        sCampos = Left(sCampos, Len(sCampos) - 2)
                        
                        If idx.Primary Then
                            ' Se for Chave Primária
                            Print #fFile, "ALTER TABLE [" & tdfOrigem.Name & "] ADD CONSTRAINT [" & idx.Name & "] PRIMARY KEY (" & sCampos & ");"
                        Else
                            ' Índice Comum ou Unique
                            Dim sUnique As String
                            sUnique = IIf(idx.Unique, "UNIQUE ", "")
                            Print #fFile, "CREATE " & sUnique & "INDEX [" & idx.Name & "] ON [" & tdfOrigem.Name & "] (" & sCampos & ");"
                        End If
                        Err.Clear
                    End If
                Next idx
                
                ' Verifica se a contagem é zero antes de passar para a próxima tabela
                If iContagemIndices = 0 Then
                    Print #fFile, "-- OBSERVAÇÃO: A tabela [" & tdfOrigem.Name & "] não possui índices definidos na origem."
                End If
            End If
            On Error GoTo Err_Handle
        End If
    Next tdfOrigem

   
    Print #fFile, ""
    Print #fFile, "-- Fim do Script"
    Close #fFile
    
    ' CHAMADA DA NOVA FUNÇÃO:
    If bExecutarAoFinal Then
        Call ExecutarArquivoSQL(sPathSQL, dbDestino)
    End If
    
      
    MsgBox "Processo concluído!" & vbCrLf & "Script gerado em: " & sPathSQL, vbInformation, "Migração Access"
    
    'LIMPEZA OBRIGATÓRIA PARA EVITAR QUE O PROGRAMA FECHE NA PRÓXIMA EXECUÇÃO
    If Not dbOrigem Is Nothing Then dbOrigem.Close
    If Not dbDestino Is Nothing Then dbDestino.Close
    
    Set dbOrigem = Nothing
    Set dbDestino = Nothing
    Set fso = Nothing
    
       Exit Sub

Err_Handle:
    MsgBox "Erro: " & Err.Number & " - " & Err.Description, vbCritical, "Erro de Processamento"
    If fFile > 0 Then Close #fFile
End Sub

' Função para converter o tipo DAO para a sintaxe SQL do Access
Private Function GetSQLType(ByRef fld As DAO.Field) As String
    Select Case fld.Type
        Case dbBoolean: GetSQLType = "BIT"
        Case dbByte: GetSQLType = "BYTE"
        Case dbInteger: GetSQLType = "SHORT"
        Case dbLong: GetSQLType = "LONG"
        Case dbCurrency: GetSQLType = "CURRENCY"
        Case dbSingle: GetSQLType = "SINGLE"
        Case dbDouble: GetSQLType = "DOUBLE"
        Case dbDate: GetSQLType = "DATETIME"
        Case dbText: GetSQLType = "TEXT(" & fld.Size & ")"
        Case dbLongBinary: GetSQLType = "LONGBINARY" ' Ole Object
        Case dbMemo: GetSQLType = "MEMO"
        Case dbAttachment: GetSQLType = "ATTACHMENT" ' Novo no ACCDB
        Case Else: GetSQLType = "TEXT(255)"
    End Select
End Function


Public Sub ExecutarArquivoSQL(ByVal sPathSQL As String, ByRef dbDestino As DAO.Database)
    Dim fFile As Integer
    Dim sLinha As String
    Dim sComando As String
    
    fFile = FreeFile
    Open sPathSQL For Input As #fFile
    
    On Error Resume Next ' Para ignorar comentários e erros individuais de execução
    Do While Not EOF(fFile)
        Line Input #fFile, sLinha
        sLinha = Trim(sLinha)
        
        ' Ignora linhas vazias ou comentários do script
        If sLinha <> "" And Left(sLinha, 2) <> "--" Then
            ' Executa o comando SQL diretamente no banco
            dbDestino.Execute sLinha, dbFailOnError
            
            If Err.Number <> 0 Then
                Debug.Print "Erro ao executar: " & sLinha & " -> " & Err.Description
                Err.Clear
            End If
        End If
    Loop
    
    Close #fFile
End Sub


' Adicione esta função auxiliar para criar o banco vazio
Private Function CriarBancoAcessoEmBranco(ByVal sPath As String) As Boolean
    On Error GoTo Err_Criar
    Dim cat As Object ' ADOX.Catalog
    Set cat = CreateObject("ADOX.Catalog")
    
    ' String de conexão para o motor ACE (ACCDB)
    cat.Create "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sPath
    
    Set cat = Nothing
    CriarBancoAcessoEmBranco = True
    Exit Function

Err_Criar:
    MsgBox "Erro ao criar banco de destino: " & Err.Description, vbCritical
    CriarBancoAcessoEmBranco = False
End Function


Private Function GerarSqlCriacaoTabela(ByRef tdf As DAO.TableDef) As String
    Dim sSql As String
    Dim fld As DAO.Field
    
    sSql = "CREATE TABLE [" & tdf.Name & "] ("
    
    For Each fld In tdf.Fields
        sSql = sSql & "[" & fld.Name & "] " & GetSQLType(fld) & ", "
    Next
    
    ' Remove a última vírgula e fecha parênteses
    sSql = Left(sSql, Len(sSql) - 2) & ");"
    
    GerarSqlCriacaoTabela = sSql
End Function


' Nova rotina para processar uma pasta inteira
Public Sub ProcessarPastaCompleta(ByVal sFolder As String, ByVal bExecutarAoFinal As Boolean)
    Dim fso As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim sPathMDB As String
    Dim sPathACCDB As String
    Dim sPathSQL As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(sFolder) Then
        MsgBox "Diretório não encontrado!", vbCritical
        Exit Sub
    End If
    
    Set oFolder = fso.GetFolder(sFolder)
    
    ' Itera por todos os arquivos da pasta
    For Each oFile In oFolder.Files
        ' Verifica se é um arquivo .mdb
        If LCase(fso.GetExtensionName(oFile.Path)) = "mdb" Then
            sPathMDB = oFile.Path
            ' Define o destino como o mesmo nome mas com extensão .accdb
            sPathACCDB = fso.BuildPath(sFolder, fso.GetBaseName(oFile.Name) & ".accdb")
            ' Define o nome do script SQL
            sPathSQL = fso.BuildPath(sFolder, fso.GetBaseName(oFile.Name) & "_correcao.sql")
            
            ' Chama a lógica de comparação existente para cada par de arquivos
            ' Nota: A lógica interna já pergunta se deseja criar o .accdb se ele não existir
            Debug.Print "Processando: " & oFile.Name
            Call GerarScriptCorrecao(sPathMDB, sPathACCDB, sPathSQL, bExecutarAoFinal)
        End If
    Next oFile
    
    MsgBox "Processamento da pasta concluído!", vbInformation
    
    Set oFolder = Nothing
    Set fso = Nothing
End Sub
