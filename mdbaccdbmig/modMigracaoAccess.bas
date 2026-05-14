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
    
    On Error GoTo Err_Handle
    
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
            
            Debug.Print "Analisando tabela: " & tdfOrigem.Name
            
            On Error Resume Next
            Set tdfDestino = dbDestino.TableDefs(tdfOrigem.Name)
            
            If Err.Number <> 0 Then
                ' Se a tabela nem existir no destino
                Print #fFile, "-- ATENÇÃO: Tabela [" & tdfOrigem.Name & "] não encontrada no destino."
                Err.Clear
            Else
                ' --- 1. VERIFICAÇÃO DE CAMPOS ---
                For Each fld In tdfOrigem.Fields
                    Dim fldDest As DAO.Field
                    Set fldDest = tdfDestino.Fields(fld.Name)
                    
                    If Err.Number <> 0 Then
                        ' Campo não existe: Gera ALTER TABLE
                        Print #fFile, "ALTER TABLE [" & tdfOrigem.Name & "] ADD COLUMN [" & fld.Name & "] " & GetSQLType(fld) & ";"
                        Err.Clear
                    End If
                Next fld
                
                ' --- 2. VERIFICAÇÃO DE ÍNDICES ---
                For Each idx In tdfOrigem.Indexes
                    ' Ignora índices criados automaticamente pelo Access para campos simples (opcional)
                    ' Se quiser ser rigoroso, verificamos todos.
                    
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
            End If
            On Error GoTo Err_Handle
        End If
    Next tdfOrigem

    Print #fFile, ""
    Print #fFile, "-- Fim do Script"
    Close #fFile
    
    dbOrigem.Close
    dbDestino.Close
    
    MsgBox "Processo concluído!" & vbCrLf & "Script gerado em: " & sPathSQL, vbInformation, "Migração Access"
    
    ' CHAMADA DA NOVA FUNÇÃO:
    If bExecutarAoFinal Then
        Call ExecutarArquivoSQL(sPathSQL, dbDestino)
    End If
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

Função para executar o conteúdo de um arquivo .sql no banco de destino
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
