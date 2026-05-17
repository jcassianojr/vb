Attribute VB_Name = "modMigracaoAccess"
Option Explicit


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


' ===========================================================================
' Nova Função Auxiliar: Gera a string SQL de todos os índices de uma tabela
' ===========================================================================
Private Function GerarSqlIndicesTabela(ByRef tdf As DAO.TableDef) As String
    Dim sSql As String
    Dim idx As DAO.Index
    Dim fldIdx As DAO.Field
    Dim sCampos As String
    Dim sUnique As String
    
    sSql = ""
    For Each idx In tdf.Indexes
        sCampos = ""
        For Each fldIdx In idx.Fields
            sCampos = sCampos & "[" & fldIdx.Name & "], "
        Next fldIdx
        
        If sCampos <> "" Then
            sCampos = Left(sCampos, Len(sCampos) - 2)
            
            If idx.Primary Then
                sSql = sSql & "ALTER TABLE [" & tdf.Name & "] ADD CONSTRAINT [" & idx.Name & "] PRIMARY KEY (" & sCampos & ");" & vbCrLf
            Else
                sUnique = IIf(idx.Unique, "UNIQUE ", "")
                sSql = sSql & "CREATE " & sUnique & "INDEX [" & idx.Name & "] ON [" & tdf.Name & "] (" & sCampos & ");" & vbCrLf
            End If
        End If
    Next idx
    
    GerarSqlIndicesTabela = sSql
End Function

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
    
    ' --- VERIFICAÇÃO DO DESTINO ---
    If Not fso.FileExists(sPathACCDB) Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox("O arquivo de destino não existe:" & vbCrLf & sPathACCDB & _
                      vbCrLf & vbCrLf & "Deseja criar um banco de dados novo em branco?", _
                      vbQuestion + vbYesNo, "Destino não encontrado")
        
        If resp = vbYes Then
            If Not CriarBancoAcessoEmBranco(sPathACCDB) Then Exit Sub
        Else
            Exit Sub
        End If
    End If
    
    ' Abre os bancos de dados
    Set dbOrigem = OpenDatabase(sPathMDB)
    Set dbDestino = OpenDatabase(sPathACCDB, False, False)
    
    fFile = FreeFile
    Open sPathSQL For Output As #fFile
    
    ' Nota: Removidos cabeçalhos de comentários antigos para evitar falhas no interpretador SQL

    For Each tdfOrigem In dbOrigem.TableDefs
        ' Ignora tabelas ocultas do sistema Access
        If Left(tdfOrigem.Name, 4) <> "MSys" Then
        
            iContagemIndices = 0
            Debug.Print "Analisando tabela: " & tdfOrigem.Name
            
            On Error Resume Next
            Set tdfDestino = dbDestino.TableDefs(tdfOrigem.Name)
     
            If Err.Number <> 0 Then
                ' ==============================================================================
                ' CENÁRIO: A tabela não existe no destino. Criamos a tabela e os índices.
                ' ==============================================================================
                Err.Clear
                
                Dim sSqlCreate As String
                sSqlCreate = GerarSqlCriacaoTabela(tdfOrigem)
    
                ' Grava estritamente a instrução SQL de criação com ponto e vírgula no arquivo
                Print #fFile, sSqlCreate & ";"
                
                ' Gera e grava os índices da nova tabela no arquivo (instruções puras)
                Dim sSqlIndicesNovos As String
                sSqlIndicesNovos = GerarSqlIndicesTabela(tdfOrigem)
                If sSqlIndicesNovos <> "" Then
                    Print #fFile, sSqlIndicesNovos
                End If
                
                ' Executa a criação IMEDIATA no banco para que as próximas checagens funcionem
                If bExecutarAoFinal Then
                    On Error Resume Next
                    dbDestino.Execute sSqlCreate, dbFailOnError
                    
                    ' Aplica os índices da nova tabela imediatamente no banco de dados
                    If sSqlIndicesNovos <> "" Then
                        Dim arrCmd() As String
                        Dim vCmd As Variant
                        arrCmd = Split(sSqlIndicesNovos, vbCrLf)
                        For Each vCmd In arrCmd
                            If Trim(vCmd) <> "" Then dbDestino.Execute vCmd, dbFailOnError
                        Next vCmd
                    End If
                    On Error GoTo Err_Handle
                End If
                ' ==============================================================================
                
            Else
                ' --- 1. VERIFICAÇÃO DE CAMPOS INEXISTENTES OU DIFERENTES ---
                For Each fld In tdfOrigem.Fields
                    Dim fldDest As DAO.Field
                    Dim sDefaultOrigem As String
                    
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
                        If sDefaultOrigem <> fldDest.DefaultValue Then
                            Print #fFile, "ALTER TABLE [" & tdfOrigem.Name & "] ALTER COLUMN [" & fld.Name & "] SET DEFAULT " & IIf(sDefaultOrigem = "", "NULL", sDefaultOrigem) & ";"
                        End If
                    End If
                Next fld
                
                ' --- 2. VERIFICAÇÃO DE ÍNDICES FALTANTES ---
                Dim idxDest As DAO.Index
                For Each idx In tdfOrigem.Indexes
                    On Error Resume Next
                    
                    iContagemIndices = iContagemIndices + 1
                    Set idxDest = tdfDestino.Indexes(idx.Name)
                    
                    If Err.Number <> 0 Then
                        ' Índice não existe no destino: Gera comando de criação puro
                        Dim sCampos As String
                        Dim fldIdx As DAO.Field
                        
                        sCampos = ""
                        For Each fldIdx In idx.Fields
                            sCampos = sCampos & "[" & fldIdx.Name & "], "
                        Next fldIdx
                        
                        If sCampos <> "" Then
                            sCampos = Left(sCampos, Len(sCampos) - 2)
                            
                            If idx.Primary Then
                                Print #fFile, "ALTER TABLE [" & tdfOrigem.Name & "] ADD CONSTRAINT [" & idx.Name & "] PRIMARY KEY (" & sCampos & ");"
                            Else
                                Dim sUnique As String
                                sUnique = IIf(idx.Unique, "UNIQUE ", "")
                                Print #fFile, "CREATE " & sUnique & "INDEX [" & idx.Name & "] ON [" & tdfOrigem.Name & "] (" & sCampos & ");"
                            End If
                        End If
                        Err.Clear
                    End If
                Next idx
            End If
            On Error GoTo Err_Handle
        End If
    Next tdfOrigem

    Close #fFile
    
    ' Força o sistema operacional a salvar e fechar o arquivo fisicamente antes da leitura
    DoEvents
    
    ' 1. EXECUTA AS MUDANÇAS NO DESTINO VIA SCRIPT (Lendo o arquivo limpo)
    If bExecutarAoFinal Then
        Call ExecutarArquivoSQL(sPathSQL, dbDestino)
    End If
 
    ' 2. GERA OS SCRIPTS INTEGRAIS ATUALIZADOS
    Dim sPathSQLOrigem As String
    Dim sPathSQLDestino As String
    
    sPathSQLOrigem = Left(sPathSQL, Len(sPathSQL) - 13) & "__MDB.sql"
    sPathSQLDestino = Left(sPathSQL, Len(sPathSQL) - 13) & "_ACCDB.sql"
    
    Call ExportarEstruturaCompleta(dbOrigem, sPathSQLOrigem, "ORIGEM (MDB)")
    Call ExportarEstruturaCompleta(dbDestino, sPathSQLDestino, "DESTINO (ACCDB)")
      
    ' 3. MENSAGEM FINAL E LIMPEZA
    MsgBox "Processo concluído com sucesso!" & vbCrLf & _
           "Script Diferencial: " & sPathSQL & vbCrLf & _
           "Script Total MDB: " & sPathSQLOrigem & vbCrLf & _
           "Script Total ACCDB: " & sPathSQLDestino, vbInformation, "Migração Access"
    
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
' ===========================================================================
' Nova Rotina Pública: Exporta a estrutura completa de um banco para SQL DDL
' ===========================================================================
Public Sub ExportarEstruturaCompleta(ByRef dbBanco As DAO.Database, ByVal sPathSQL As String, ByVal sTitulo As String)
    Dim tdf As DAO.TableDef
    Dim idx As DAO.Index
    Dim fldIdx As DAO.Field
    Dim fFile As Integer
    Dim sCampos As String
    Dim sUnique As String
    
    If dbBanco Is Nothing Then Exit Sub
    
    fFile = FreeFile
    Open sPathSQL For Output As #fFile
    
    Print #fFile, "-- ======================================================"
    Print #fFile, "-- SCRIPT DE ESTRUTURA COMPLETA: " & sTitulo
    Print #fFile, "-- Gerado em: " & Now
    Print #fFile, "-- ======================================================"
    Print #fFile, ""
    
    For Each tdf In dbBanco.TableDefs
        ' Ignora tabelas ocultas do sistema Access
        If Left(tdf.Name, 4) <> "MSys" Then
            Print #fFile, "-- --------------------------------------------------"
            Print #fFile, "-- Estrutura da Tabela: " & tdf.Name
            Print #fFile, "-- --------------------------------------------------"
            
            ' 1. Escreve o comando CREATE TABLE principal
            Print #fFile, GerarSqlCriacaoTabela(tdf)
            
            ' 2. Escreve todos os índices da tabela de forma individualizada
            For Each idx In tdf.Indexes
                sCampos = ""
                For Each fldIdx In idx.Fields
                    sCampos = sCampos & "[" & fldIdx.Name & "], "
                Next fldIdx
                
                If sCampos <> "" Then
                    sCampos = Left(sCampos, Len(sCampos) - 2)
                    
                    If idx.Primary Then
                        Print #fFile, "ALTER TABLE [" & tdf.Name & "] ADD CONSTRAINT [" & idx.Name & "] PRIMARY KEY (" & sCampos & ");"
                    Else
                        sUnique = IIf(idx.Unique, "UNIQUE ", "")
                        Print #fFile, "CREATE " & sUnique & "INDEX [" & idx.Name & "] ON [" & tdf.Name & "] (" & sCampos & ");"
                    End If
                End If
            Next idx
            
            Print #fFile, ""
        End If
    Next tdf
    
    Print #fFile, "-- Fim do Script de " & sTitulo
    Close #fFile
End Sub

