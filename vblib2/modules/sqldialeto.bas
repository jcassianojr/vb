Attribute VB_Name = "sqldialetofuncoes"
' ==============================================================================
' FUNÇÃO: SQLPGSQLDOUBLEQUOTES
' DESCRIÇÃO: Converte identificadores (tabelas/campos) para o padrão PostgreSQL
'            utilizando aspas duplas, preservando literais em aspas simples.
'
' EXEMPLOS DE PARSE (TESTES):
'
' 1. SQL Simples:
'    Entrada: SELECT Codigo, Nome FROM Clientes
'    Saída:   SELECT "Codigo", "Nome" FROM "Clientes"
'
' 2. Proteção de Literais (Lookahead):
'    Entrada: SELECT Nome FROM Usuarios WHERE Perfil = 'ADMIN' AND Acao = 'SELECT'
'    Saída:   SELECT "Nome" FROM "Usuarios" WHERE "Perfil" = 'ADMIN' AND "Acao" = 'SELECT'
'    (Nota: O valor 'SELECT' não recebe aspas duplas por estar dentro de aspas simples)
'
' 3. Funções e Agregações:
'    Entrada: SELECT SUM(Total), Data FROM Vendas GROUP BY Data
'    Saída:   SELECT SUM("Total"), "Data" FROM "Vendas" GROUP BY "Data"
'
' 4. Operadores e Datas:
'    Entrada: SELECT * FROM Movimentos WHERE DtLanc = '2023-01-01'
'    Saída:   SELECT * FROM "Movimentos" WHERE "DtLanc" = '2023-01-01'
' ==============================================================================
Public Function SQLPGSQLDOUBLEQUOTES(ByVal cSQL As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    
    regEx.Global = True
    regEx.IgnoreCase = True
    
    ' Padrão: Identifica palavras que são seguidas por vírgula ou espaço,
    ' mas ignora palavras reservadas (simplificado para o exemplo)
    ' O padrão abaixo busca palavras que NÃO estão entre aspas simples.
    regEx.Pattern = "\b([a-z_][a-z0-9_]*)\b(?=(?:[^']*'[^']*')*[^']*$)"
    
    ' Aqui usamos o método do objeto regEx
    ' O "$1" refere-se ao que foi encontrado no primeiro parênteses do Pattern
    Dim sResult As String
    sResult = regEx.Replace(cSQL, """$1""")
    
    ' Limpeza de palavras reservadas que não devem ter aspas
    ' (Regex também pode fazer isso, mas aqui ilustro o ajuste pós-processo)
    sResult = Replace(sResult, """SELECT""", "SELECT")
    sResult = Replace(sResult, """FROM""", "FROM")
    sResult = Replace(sResult, """WHERE""", "WHERE")
    sResult = Replace(sResult, " ""SUM""(", " SUM(")
    sResult = Replace(sResult, " ""COUNT""(", " COUNT(")
    sResult = Replace(sResult, """SUM""", "SUM")
    sResult = Replace(sResult, """COUNT""", "COUNT")
    sResult = Replace(sResult, """GROUP""", "GROUP")
    sResult = Replace(sResult, """BY""", "BY")
    sResult = Replace(sResult, """AND""", "AND")
    sResult = Replace(sResult, """OR""", "OR")
    sResult = Replace(sResult, """ORDER""", "ORDER")
    
    SQLPGSQLDOUBLEQUOTES = sResult
End Function

Public Function SQLPGSQLDOUBLEQUOTES_antiga(ByVal cSQL As String) As String
cSQL = UCase(cSQL)
    cSQL = Replace(cSQL, "  ", " ")
    cSQL = Replace(cSQL, " ,", ",")
    cSQL = Replace(cSQL, ", ", ",")
    cSQL = Replace(cSQL, "SELECT ", "SELECT " + Chr(34))
    cSQL = Replace(cSQL, ",", Chr(34) + "," + Chr(34))
    cSQL = Replace(cSQL, " FROM ", Chr(34) + " FROM " + Chr(34))
    If InStr(cSQL, " WHERE ") > 0 Then
      cSQL = Replace(cSQL, " WHERE ", Chr(34) + " WHERE " + Chr(34))
      cSQL = Replace(cSQL, "=", Chr(34) + "=")
    End If
    If InStr(cSQL, " ORDER BY ") > 0 Then
      cSQL = Replace(cSQL, " ORDER BY ", Chr(34) + " ORDER BY " + Chr(34))
    End If
    cSQL = cSQL + Chr(34)
    cSQL = Replace(cSQL, Chr(39) + Chr(34), Chr(39))
    cSQL = Replace(cSQL, Chr(34) + Chr(34), Chr(34))
    cSQL = Replace(cSQL, Chr(34) + "*" + Chr(34), "*") 'astericos nao requer double quotes
    
    SQLPGSQLDOUBLEQUOTES_antiga = cSQL
End Function




' =========================================================================
' Módulo de Tradução de Dialetos SQL Unificado (VB6 - ADODB)
' Espelho fiel e expandido do arquivo dbudialeto.prg do Harbour
' =========================================================================

' -------------------------------------------------------------------------
' FUNÇÃO: Dialeto_begin
' Retorna o comando exato para iniciar uma transação conforme o banco
' -------------------------------------------------------------------------
Public Function Dialeto_begin(ByVal cTIPO As String) As String
    Dim CCOMANDO As String
    cTIPO = Trim(UCase(cTIPO))
    
    CCOMANDO = "BEGIN TRANSACTION"
    
    Select Case cTIPO
        Case "MSSQL", "SQLSERVER"
            CCOMANDO = "BEGIN TRANSACTION"
            
        Case "MYSQL", "MYSQL64", "MARIADB"
            CCOMANDO = "START TRANSACTION;"
            
        Case "FIREBIRD"
            CCOMANDO = "SET TRANSACTION"
            
        Case "PGSQL", "POSTGRESQL", "PGSQL64"
            CCOMANDO = "BEGIN;"
            
        Case "ORACLE", "OCI"
            CCOMANDO = "" ' Oracle inicia transações implicitamente no primeiro comando DML
            
        Case "SQLITE"
            CCOMANDO = "BEGIN TRANSACTION;"
            
        Case "ACCESS"
            CCOMANDO = "BEGIN TRANSACTION"
    End Select
    
    Dialeto_begin = CCOMANDO
End Function

' -------------------------------------------------------------------------
' FUNÇÃO: Dialeto_commit
' Retorna o comando exato para confirmar uma transação conforme o banco
' -------------------------------------------------------------------------
Public Function Dialeto_commit(ByVal cTIPO As String) As String
    Dim CCOMANDO As String
    cTIPO = Trim(UCase(cTIPO))
    
    CCOMANDO = "COMMIT"
    
    Select Case cTIPO
        Case "MSSQL", "SQLSERVER"
            CCOMANDO = "IF @@TRANCOUNT > 0 COMMIT"
            
        Case "MYSQL", "MYSQL64", "MARIADB"
            CCOMANDO = "COMMIT;"
            
        Case "PGSQL", "POSTGRESQL", "PGSQL64"
            CCOMANDO = "COMMIT;"
            
        Case "SQLITE"
            CCOMANDO = "end transaction"
            
        Case "FIREBIRD", "ORACLE", "OCI", "ACCESS"
            CCOMANDO = "COMMIT"
    End Select
    
    Dialeto_commit = CCOMANDO
End Function

' -------------------------------------------------------------------------
' FUNÇÃO: Dialeto_rollback
' Retorna o comando exato para anular uma transação conforme o banco
' -------------------------------------------------------------------------
Public Function Dialeto_rollback(ByVal cTIPO As String) As String
    Dim CCOMANDO As String
    cTIPO = Trim(UCase(cTIPO))
    
    CCOMANDO = "ROLLBACK"
    
    Select Case cTIPO
        Case "MSSQL", "SQLSERVER"
            CCOMANDO = "IF @@TRANCOUNT > 0 ROLLBACK"
            
        Case "MYSQL", "MYSQL64", "MARIADB"
            CCOMANDO = "ROLLBACK;"
            
        Case "PGSQL", "POSTGRESQL", "PGSQL64"
            CCOMANDO = "ROLLBACK;"
            
        Case "SQLITE"
            CCOMANDO = "ROLLBACK;"
            
        Case "FIREBIRD", "ORACLE", "OCI", "ACCESS"
            CCOMANDO = "ROLLBACK"
    End Select
    
    Dialeto_rollback = CCOMANDO
End Function

' -------------------------------------------------------------------------
' FUNÇÃO: Dialeto_GetIdentity
' Retorna a Query para recuperar o ID Auto-Incremento gerado na sessão
' -------------------------------------------------------------------------
Public Function Dialeto_GetIdentity(ByVal cTIPO As String) As String
    Dim CCOMANDO As String
    cTIPO = Trim(UCase(cTIPO))
    
    CCOMANDO = ""
    
    Select Case cTIPO
        Case "SQLITE"
            CCOMANDO = "SELECT last_insert_rowid();"
            
        Case "PGSQL", "POSTGRESQL", "PGSQL64"
            CCOMANDO = "SELECT lastval();"
            
        Case "MSSQL", "SQLSERVER"
            CCOMANDO = "SELECT @@IDENTITY;"
            
        Case "MYSQL", "MYSQL64", "MARIADB"
            CCOMANDO = "SELECT LAST_INSERT_ID();"
            
        Case "ACCESS"
            CCOMANDO = "SELECT @@IDENTITY;"
    End Select
    
    Dialeto_GetIdentity = CCOMANDO
End Function

' -------------------------------------------------------------------------
' FUNÇÃO PRINCIPAL: SQLDialeto (Revisada com a chamada correta da sua SQLDIAPAR)
' -------------------------------------------------------------------------
Public Function sqldialeto(ByVal cSQLCNV As String, ByVal cDIALETO As String) As String
    Dim sResult As String
    sResult = cSQLCNV
    
    sResult = Replace(sResult, "  ", " ")
    cDIALETO = Trim(UCase(cDIALETO))

    Select Case cDIALETO
        
        ' -----------------------------------------------------------------
        ' REGREAS POSTGRESQL (Mesclado com dbudialeto.prg)
        ' -----------------------------------------------------------------
        Case "PGSQL", "POSTGRESQL", "PGSQL64"
            ' Concatenação e datas padrão Access
            sResult = Replace(sResult, " & ", " || ")
            sResult = Replace(sResult, "#", "'")
            
            ' Literais Lógicos
            sResult = Replace(sResult, ".T.", "TRUE")
            sResult = Replace(sResult, ".F.", "FALSE")
            sResult = Replace(sResult, " = True", " = TRUE")
            sResult = Replace(sResult, " = False", " = FALSE")
            
            ' Tradução das funções cronológicas usando a sua SQLDIAPAR real
            sResult = SQLDIAPAR(sResult, "DAY(", ")", "DATE_PART('day', ", ")")
            sResult = SQLDIAPAR(sResult, "MONTH(", ")", "DATE_PART('month', ", ")")
            sResult = SQLDIAPAR(sResult, "YEAR(", ")", "DATE_PART('year', ", ")")
            sResult = SQLDIAPAR(sResult, "HOUR(", ")", "DATE_PART('hour', ", ")")
            sResult = SQLDIAPAR(sResult, "MINUTE(", ")", "DATE_PART('minute', ", ")")
            sResult = SQLDIAPAR(sResult, "SECOND(", ")", "DATE_PART('second', ", ")")
            sResult = SQLDIAPAR(sResult, "DTOS(", ")", "TO_CHAR(", ", 'YYYYMMDD')")
            
            ' Funções Estáticas e Strings com "(" colado para proteção
            sResult = Replace(sResult, "TODAY()", "CURRENT_DATE")
            sResult = Replace(sResult, "DATE()", "CURRENT_DATE")
            sResult = Replace(sResult, "NOW()", "NOW()")
            sResult = Replace(sResult, "CURRENTDATETIME", "NOW()")
            sResult = Replace(sResult, "SUBSTR(", "SUBSTRING(")
            sResult = Replace(sResult, "LEN(", "LENGTH(")
            sResult = Replace(sResult, "ALLTRIM(", "TRIM(")
            sResult = Replace(sResult, "ASC(", "ASCII(")
            sResult = Replace(sResult, "CHR(", "CHR(")
            
            ' Máscaras Jet/Access
            sResult = Replace(sResult, "IIF(", "CASE WHEN ")
            sResult = Replace(sResult, "IFNULL(", "COALESCE(")
            
            ' Purga de vazios padrão Clipper/FoxPro
            sResult = Replace(sResult, "'  /  /  '", "NULL")
            sResult = Replace(sResult, "'00/00/0000'", "NULL")

        ' -----------------------------------------------------------------
        ' REGRAS SQLITE (Ajustadas com o formato exato da sua SQLDIAPAR)
        ' -----------------------------------------------------------------
        Case "SQLITE"
            sResult = Replace(sResult, " & ", " || ")
            sResult = Replace(sResult, "#", "'")
            sResult = Replace(sResult, " = True", " = 1")
            sResult = Replace(sResult, " = False", " = 0")
            
            sResult = Replace(sResult, "CURRENTDATETIME", " current_timestamp ")
            sResult = Replace(sResult, "TODAY()", "CURRENT_DATE ")
            sResult = Replace(sResult, "DATE()", "CURRENT_DATE ")
            sResult = Replace(sResult, "CHR(", "CHAR(")
            sResult = Replace(sResult, "ASC(", "ASCII(")
            sResult = Replace(sResult, "TRIM(", "RTRIM(")
            sResult = Replace(sResult, "ALLTRIM(", "TRIM(")
            sResult = Replace(sResult, "LEN(", "LENGTH(")
            sResult = Replace(sResult, "IIF(", "CASE WHEN ")
            sResult = Replace(sResult, "IFNULL(", "COALESCE(")
            
            ' Chamadas ajustadas para casar perfeitamente com a sua lógica interna de loops
            sResult = SQLDIAPAR(sResult, "DTOS(", ")", "strftime('%Y%m%d', ", ")")
            sResult = SQLDIAPAR(sResult, "DAY(", ")", "cast(strftime('%d', ", ") as int)")
            sResult = SQLDIAPAR(sResult, "MONTH(", ")", "cast(strftime('%m', ", ") as int)")
            sResult = SQLDIAPAR(sResult, "YEAR(", ")", "cast(strftime('%Y', ", ") as int)")

        ' -----------------------------------------------------------------
        ' REGRAS MYSQL / MARIADB
        ' -----------------------------------------------------------------
        Case "MYSQL", "MARIADB", "MYSQL64"
            sResult = Replace(sResult, " & ", " || ")
            sResult = Replace(sResult, "#", "'")
            sResult = Replace(Replace(sResult, "[", "`"), "]", "`")
            
            sResult = Replace(sResult, "TODAY()", "SYSDATE()")
            sResult = Replace(sResult, "DATE()", "CURDATE()")
            sResult = Replace(sResult, "CURRENTDATETIME", "CURRENT_TIMESTAMP()")
            sResult = Replace(sResult, "CHR(", "CHAR(")
            sResult = Replace(sResult, "ASC(", "ASCII(")
            sResult = Replace(sResult, "TRIM(", "RTRIM(")
            sResult = Replace(sResult, "ALLTRIM(", "TRIM(")
            sResult = Replace(sResult, "REPL(", "REPEAT(")
            sResult = Replace(sResult, "SUBSTR(", "SUBSTRING(")
            sResult = Replace(sResult, "LEN(", "LENGTH(")
            
            sResult = SQLDIAPAR(sResult, "DTOS(", ")", "DATE_FORMAT(", ", '%Y%m%d')")

        ' -----------------------------------------------------------------
        ' REGRAS MSSQL / SQLSERVER
        ' -----------------------------------------------------------------
        Case "MSSQL", "SQLSERVER"
            sResult = Replace(sResult, "TODAY()", "GETDATE() ")
            sResult = Replace(sResult, "DATE()", "CAST(GETDATE() AS DATE)")
            sResult = Replace(sResult, "CURRENTDATETIME", "SYDATETIME()")
            sResult = Replace(sResult, "ASC(", "ASCII(")
            sResult = Replace(sResult, "TRIM(", "RTRIM(")
            sResult = Replace(sResult, "ALLTRIM(", "TRIM(")
            sResult = Replace(sResult, "REPL(", "REPLICATE(")
            sResult = Replace(sResult, "CHR(", "CHAR(")
            sResult = Replace(sResult, "SUBSTR(", "SUBSTRING(")
            sResult = Replace(sResult, "LEN(", "LENGTH(")
            
            sResult = SQLDIAPAR(sResult, "DTOS(", ")", "CONVERT(char(8), ", ", 11)")

    End Select
    
    sqldialeto = sResult
End Function

' -------------------------------------------------------------------------
' A SUA FUNÇÃO ORIGINAL COMPLETA (Preservada intacta)
' -------------------------------------------------------------------------
Public Function SQLDIAPAR(ByVal cSQLCNV As String, cDelimIni As String, cDelimFim As String, cREPINI As String, cREPFIM As String) As String
    Dim aRETU
    SQLDIAPAR = cSQLCNV
    aRETU = pegue2delimitado(cSQLCNV, cDelimIni, cDelimFim)
    If Len(aRETU(0)) > 0 Then
       SQLDIAPAR = Replace(aRETU(2), cDelimIni, "")
       SQLDIAPAR = SQLDIAPAR & cREPINI & aRETU(0) & cREPFIM & aRETU(1)
    End If
End Function

