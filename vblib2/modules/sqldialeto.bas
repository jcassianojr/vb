Attribute VB_Name = "sqldialetofuncoes"
' ==============================================================================
' Módulo: sqldialetofuncoes.bas (VB6)
' Descrição: Espelho completo, fiel e expandido do dbudialeto.prg (Harbour)
' ==============================================================================

' -------------------------------------------------------------------------
' 1. HIGIENIZAÇÃO DE ESTRUTURAS E CONVERSÕES AUXILIARES
' -------------------------------------------------------------------------
Public Function CleanDbStruct(ByVal aStruct As Variant) As Variant
    Dim i As Long, cType As String, nPOS As Long
    For i = LBound(aStruct, 1) To UBound(aStruct, 1)
        cType = aStruct(i, 2)
        nPOS = InStr(cType, ":")
        If nPOS > 0 Then
            aStruct(i, 2) = Left(cType, nPOS - 1)
        End If
    Next i
    CleanDbStruct = aStruct
End Function

' -------------------------------------------------------------------------
' 2. TRATAMENTO DE CAMPOS VAZIOS PARA SQL (EMPTY / NOT EMPTY)
' -------------------------------------------------------------------------
Public Function ConverterEmptyParaSQL(ByVal cSQL As String, Optional ByVal cTipoSql As String = "ANSI") As String
    Dim nPOS As Long, nInicio As Long, nFim As Long, cCAMPO As String, cSubst As String, lNot As Boolean, nTamRemover As Long
    
    Do While InStr(UCase(cSQL), "EMPTY(") > 0
        nPOS = InStr(UCase(cSQL), "EMPTY(")
        lNot = DetectarNegacao(cSQL, nPOS)
        
        nInicio = InStr(Mid(cSQL, nPOS), "(") + nPOS - 1
        nFim = InStr(Mid(cSQL, nInicio), ")")
        
        If nFim > 0 Then
            nFim = nFim + nInicio - 1
            cCAMPO = Trim(Mid(cSQL, nInicio + 1, nFim - nInicio - 1))
            
            If lNot Then
                nTamRemover = IIf(InStr(UCase(Mid(cSQL, IIf(nPOS - 4 > 0, nPOS - 4, 1), 4)), "NOT") > 0, 3, 1)
                cSQL = Left(cSQL, nPOS - nTamRemover - 1) + Mid(cSQL, nFim + 1)
                nPOS = nPOS - nTamRemover
            Else
                cSQL = Left(cSQL, nPOS - 1) + Mid(cSQL, nFim + 1)
            End If
            
            cSubst = GerarFragmentoSQL(cCAMPO, lNot, cTipoSql)
            cSQL = Left(cSQL, nPOS - 1) + cSubst + Mid(cSQL, nPOS)
        Else
            Exit Do
        End If
    Loop
    ConverterEmptyParaSQL = cSQL
End Function

Private Function DetectarNegacao(ByVal cSQL As String, ByVal nPOS As Long) As Boolean
    Dim cPrecedente As String
    If nPOS > 5 Then
        cPrecedente = Trim(UCase(Mid(cSQL, nPOS - 5, 5)))
    Else
        cPrecedente = Trim(UCase(Left(cSQL, nPOS - 1)))
    End If
    If Right(cPrecedente, 1) = "!" Or Right(cPrecedente, 3) = "NOT" Then
        DetectarNegacao = True
    Else
        DetectarNegacao = False
    End If
End Function

Private Function GerarFragmentoSQL(ByVal cCAMPO As String, ByVal lNot As Boolean, ByVal cTipoSql As String) As String
    Dim cRet As String
    cTipoSql = UCase(cTipoSql)
    If lNot Then
        Select Case cTipoSql
            Case "ORACLE"
                cRet = " ( " + cCAMPO + " IS NOT NULL ) "
            Case Else
                cRet = " ( " + cCAMPO + " IS NOT NULL AND " + cCAMPO + " <> '' ) "
        End Select
    Else
        Select Case cTipoSql
            Case "ORACLE"
                cRet = " ( " + cCAMPO + " IS NULL ) "
            Case Else
                cRet = " ( " + cCAMPO + " IS NULL OR " + cCAMPO + " = '' ) "
        End Select
    End If
    GerarFragmentoSQL = cRet
End Function

' -------------------------------------------------------------------------
' 3. TRANSAÇÕES (BEGIN, COMMIT, ROLLBACK)
' -------------------------------------------------------------------------
Public Function Dialeto_begin(ByVal cTIPO As String) As String
    Dim CCOMANDO As String
    cTIPO = Trim(UCase(cTIPO))
    Select Case cTIPO
        Case "MSSQL", "SQLSERVER"
            CCOMANDO = "BEGIN TRANSACTION"
        Case "MYSQL", "MYSQL64", "MARIADB"
            CCOMANDO = "START TRANSACTION;"
        Case "FIREBIRD"
            CCOMANDO = "SET TRANSACTION"
        Case "SQLITE"
            CCOMANDO = "BEGIN TRANSACTION;"
        Case "PGSQL", "POSTGRESQL", "PGSQL64"
            CCOMANDO = "BEGIN;"
        Case "ORACLE", "OCI"
            CCOMANDO = "SET TRANSACTION READ WRITE;"
        Case "DUCKDB", "DUCKLAKE"
            CCOMANDO = "BEGIN TRANSACTION;"
        Case Else
            CCOMANDO = "BEGIN TRANSACTION"
    End Select
    Dialeto_begin = CCOMANDO
End Function

Public Function Dialeto_commit(ByVal cTIPO As String) As String
    Dim CCOMANDO As String
    cTIPO = Trim(UCase(cTIPO))
    Select Case cTIPO
        Case "MSSQL", "SQLSERVER"
            CCOMANDO = "IF @@TRANCOUNT > 0 COMMIT"
        Case "MYSQL", "MYSQL64", "MARIADB"
            CCOMANDO = "COMMIT;"
        Case "FIREBIRD", "ORACLE", "OCI", "ACCESS"
            CCOMANDO = "COMMIT"
        Case "SQLITE"
            CCOMANDO = "end transaction"
        Case "PGSQL", "POSTGRESQL", "PGSQL64", "DUCKDB", "DUCKLAKE"
            CCOMANDO = "COMMIT;"
        Case Else
            CCOMANDO = "COMMIT TRANSACTION"
    End Select
    Dialeto_commit = CCOMANDO
End Function

Public Function Dialeto_rollback(ByVal cTIPO As String) As String
    Dim CCOMANDO As String
    cTIPO = Trim(UCase(cTIPO))
    Select Case cTIPO
        Case "MSSQL", "SQLSERVER"
            CCOMANDO = "IF @@TRANCOUNT > 0 ROLLBACK"
        Case "MYSQL", "MYSQL64", "MARIADB"
            CCOMANDO = "ROLLBACK;"
        Case "FIREBIRD", "ORACLE", "OCI", "ACCESS"
            CCOMANDO = "ROLLBACK"
        Case "SQLITE", "PGSQL", "POSTGRESQL", "PGSQL64", "DUCKDB", "DUCKLAKE"
            CCOMANDO = "ROLLBACK;"
        Case Else
            CCOMANDO = "ROLLBACK TRANSACTION"
    End Select
    Dialeto_rollback = CCOMANDO
End Function

' -------------------------------------------------------------------------
' 4. FUNÇÕES DE BANCO DE DADOS (DATAS, VERSÕES E IDENTITIES)
' -------------------------------------------------------------------------
Public Function Dialeto_DataBanco(ByVal cTIPO As String) As String
    Dim CCOMANDO As String
    cTIPO = Trim(UCase(cTIPO))
    Select Case cTIPO
        Case "MSSQL", "SQLSERVER"
            CCOMANDO = "GETDATE()"
        Case "MYSQL", "MYSQL64", "MARIADB"
            CCOMANDO = "SYSDATE()"
        Case "SQLITE", "PGSQL", "PGSQL64", "POSTGRESQL", "DUCKDB", "DUCKLAKE"
            CCOMANDO = "CURRENT_DATE"
        Case "ORACLE", "OCI"
            CCOMANDO = "SYSDATE"
        Case Else
            CCOMANDO = "CURRENT_DATE"
    End Select
    Dialeto_DataBanco = CCOMANDO
End Function

Public Function Dialeto_DataHoraBanco(ByVal cTIPO As String) As String
    Dim CCOMANDO As String
    cTIPO = Trim(UCase(cTIPO))
    Select Case cTIPO
        Case "MSSQL", "SQLSERVER"
            CCOMANDO = "SYSDATETIME()"
        Case "MYSQL", "MYSQL64", "MARIADB"
            CCOMANDO = "CURRENT_TIMESTAMP()"
        Case "SQLITE"
            CCOMANDO = "CURRENT_TIMESTAMP"
        Case "PGSQL", "PGSQL64", "POSTGRESQL"
            CCOMANDO = "NOW()"
        Case "ORACLE", "OCI"
            CCOMANDO = "SYSTIMESTAMP"
        Case "DUCKDB", "DUCKLAKE"
            CCOMANDO = "CURRENT_TIMESTAMP"
        Case Else
            CCOMANDO = "CURRENT_TIMESTAMP"
    End Select
    Dialeto_DataHoraBanco = CCOMANDO
End Function

Public Function Dialeto_DataVazia(ByVal cSQLDIALETO As String) As String
    Dim CCOMANDO As String
    cSQLDIALETO = Trim(UCase(cSQLDIALETO))
    Select Case cSQLDIALETO
        Case "SQLITE"
            CCOMANDO = ""
        Case Else
            CCOMANDO = "NULL"
    End Select
    Dialeto_DataVazia = CCOMANDO
End Function

Public Function Dialeto_GetIdentity(ByVal cTIPO As String) As String
    Dim CCOMANDO As String
    cTIPO = Trim(UCase(cTIPO))
    Select Case cTIPO
        Case "MSSQL", "SQLSERVER", "ACCESS"
            CCOMANDO = "SELECT @@IDENTITY AS LAST_ID;"
        Case "MYSQL", "MYSQL64", "MARIADB", "ORACLE", "OCI"
            CCOMANDO = "SELECT LAST_INSERT_ID() AS LAST_ID;"
        Case "SQLITE"
            CCOMANDO = "SELECT last_insert_rowid() AS LAST_ID;"
        Case "PGSQL", "PGSQL64", "POSTGRESQL"
            CCOMANDO = "SELECT lastval() AS LAST_ID;"
        Case "CUBRID"
            CCOMANDO = "SELECT LAST_INSERT_ID();"
        Case Else
            CCOMANDO = ""
    End Select
    Dialeto_GetIdentity = CCOMANDO
End Function

Public Function Dialeto_ShowDatabases(ByVal cTIPO As String) As String
    Dim CCOMANDO As String
    cTIPO = Trim(UCase(cTIPO))
    Select Case cTIPO
        Case "MYSQL", "MYSQL64", "MARIADB"
            CCOMANDO = "SELECT SCHEMA_NAME AS DB_NAME FROM INFORMATION_SCHEMA.SCHEMATA;"
        Case "PGSQL", "PGSQL64", "POSTGRESQL"
            CCOMANDO = "SELECT datname AS DB_NAME FROM pg_database WHERE datistemplate = false;"
        Case "MSSQL", "SQLSERVER"
            CCOMANDO = "SELECT name AS DB_NAME FROM master.dbo.sysdatabases WHERE name NOT IN ('master', 'tempdb', 'model', 'msdb');"
        Case "ORACLE", "OCI"
            CCOMANDO = "SELECT username AS DB_NAME FROM dba_users WHERE account_status = 'OPEN' ORDER BY username;"
        Case "SQLITE"
            CCOMANDO = "SELECT name AS DB_NAME FROM pragma_database_list;"
        Case "FIREBIRD", "FDB", "GDB", "IB"
            CCOMANDO = "SELECT RDB$GET_CONTEXT('SYSTEM', 'DB_NAME') AS DB_NAME FROM RDB$DATABASE;"
        Case "DUCKDB"
            CCOMANDO = "SELECT schema_name AS DB_NAME FROM information_schema.schemata;"
        Case "DUCKLAKE"
            CCOMANDO = "SELECT schema_name AS DB_NAME FROM ducklake_schema;"
    End Select
    Dialeto_ShowDatabases = CCOMANDO
End Function

Public Function Dialeto_Version(ByVal cTIPO As String) As String
    Dim CCOMANDO As String
    cTIPO = Trim(UCase(cTIPO))
    Select Case cTIPO
        Case "MSSQL", "SQLSERVER"
            CCOMANDO = "SELECT @@VERSION AS VER;"
        Case "MYSQL", "MYSQL64", "MARIADB"
            CCOMANDO = "SELECT VERSION() AS VER;"
        Case "FIREBIRD", "FDB", "GDB", "IB"
            CCOMANDO = "SELECT RDB$GET_CONTEXT('SYSTEM', 'ENGINE_VERSION') AS VER FROM RDB$DATABASE;"
        Case "SQLITE"
            CCOMANDO = "SELECT sqlite_version() AS VER;"
        Case "PGSQL", "PGSQL64", "POSTGRESQL"
            CCOMANDO = "SELECT VERSION() AS VER;"
        Case "ORACLE", "OCI"
            CCOMANDO = "SELECT BANNER AS VER FROM V$VERSION WHERE ROWNUM = 1;"
        Case "DUCKDB", "DUCKLAKE"
            CCOMANDO = "SELECT duckdb_version() AS VER;"
    End Select
    Dialeto_Version = CCOMANDO
End Function

Public Function Dialeto_GetRowCount(ByVal cTIPO As String) As String
    Dim CCOMANDO As String
    cTIPO = Trim(UCase(cTIPO))
    Select Case cTIPO
        Case "MSSQL", "SQLSERVER"
            CCOMANDO = "SELECT @@ROWCOUNT;"
        Case "MYSQL", "MYSQL64", "MARIADB"
            CCOMANDO = "SELECT ROW_COUNT();"
        Case "SQLITE"
            CCOMANDO = "SELECT changes();"
        Case "ORACLE", "OCI"
            CCOMANDO = "SELECT SQL%ROWCOUNT;"
        Case Else
            CCOMANDO = ""
    End Select
    Dialeto_GetRowCount = CCOMANDO
End Function

Public Function Dialeto_SetLimit(ByVal nQtd As Long, ByVal cTIPO As String) As String
    Dim CCOMANDO As String, cQtd As String
    cQtd = CStr(nQtd)
    cTIPO = Trim(UCase(cTIPO))
    Select Case cTIPO
        Case "MYSQL", "MYSQL64", "MARIADB", "PGSQL", "PGSQL64", "POSTGRESQL", "SQLITE", "CUBRID", "DUCKDB", "DUCKLAKE"
            CCOMANDO = " LIMIT " & cQtd
        Case "ORACLE", "OCI"
            CCOMANDO = " FETCH FIRST " & cQtd & " ROWS ONLY"
        Case "INFORMIX"
            CCOMANDO = " FIRST " & cQtd
        Case "MSSQL", "SQLSERVER", "SYBASE"
            CCOMANDO = " TOP " & cQtd & " "
    End Select
    Dialeto_SetLimit = CCOMANDO
End Function

' -------------------------------------------------------------------------
' 5. TRADUTORES E CONVERSORES DE CLÁUSULAS SQL
' -------------------------------------------------------------------------
Public Function Dialeto_condicionais(ByVal cSQLCNV As String, ByVal cTipoSql As String) As String
    cSQLCNV = Replace(cSQLCNV, ".NOT.", " NOT ")
    cSQLCNV = Replace(cSQLCNV, ".OR.", " OR ")
    cSQLCNV = Replace(cSQLCNV, ".AND.", " AND ")
    cTipoSql = UCase(cTipoSql)
    Select Case cTipoSql
        Case "MSSQL", "SQLSERVER"
            cSQLCNV = Replace(cSQLCNV, "!=", " <> ")
        Case "PGSQL", "PGSQL64", "POSTGRESQL"
            cSQLCNV = Replace(cSQLCNV, "<>", " != ")
    End Select
    Dialeto_condicionais = cSQLCNV
End Function

Public Function sqldialeto(ByVal cSQLCNV As String, ByVal cDIALETO As String) As String
    Dim sResult As String
    sResult = cSQLCNV
    sResult = Replace(sResult, "  ", " ")
    cDIALETO = Trim(UCase(cDIALETO))
    
    sResult = Dialeto_condicionais(sResult, cDIALETO)
    sResult = ConverterEmptyParaSQL(sResult, cDIALETO)

    Select Case cDIALETO
        Case "PGSQL", "POSTGRESQL", "PGSQL64"
            sResult = Replace(sResult, " & ", " || ")
            sResult = Replace(sResult, "#", "'")
            sResult = Replace(sResult, ".T.", "TRUE")
            sResult = Replace(sResult, ".F.", "FALSE")
            
            sResult = SQLDIAPAR(sResult, "DAY(", ")", "DATE_PART('day', ", ")")
            sResult = SQLDIAPAR(sResult, "MONTH(", ")", "DATE_PART('month', ", ")")
            sResult = SQLDIAPAR(sResult, "YEAR(", ")", "DATE_PART('year', ", ")")
            sResult = SQLDIAPAR(sResult, "HOUR(", ")", "DATE_PART('hour', ", ")")
            sResult = SQLDIAPAR(sResult, "MINUTE(", ")", "DATE_PART('minute', ", ")")
            sResult = SQLDIAPAR(sResult, "SECOND(", ")", "DATE_PART('second', ", ")")
            sResult = SQLDIAPAR(sResult, "DTOS(", ")", "TO_CHAR(", ", 'YYYYMMDD')")
            
            sResult = Replace(sResult, "TODAY()", "CURRENT_DATE")
            sResult = Replace(sResult, "DATE()", "CURRENT_DATE")
            sResult = Replace(sResult, "SUBSTR(", "SUBSTRING(")
            sResult = Replace(sResult, "LEN(", "LENGTH(")
            sResult = Replace(sResult, "ALLTRIM(", "TRIM(")
            sResult = Replace(sResult, "ASC(", "ASCII(")
            sResult = Replace(sResult, "IIF(", "CASE WHEN ")
            sResult = Replace(sResult, "IFNULL(", "COALESCE(")
            sResult = Replace(sResult, "'  /  /  '", "NULL")
            sResult = Replace(sResult, "'00/00/0000'", "NULL")

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
            
            sResult = SQLDIAPAR(sResult, "DTOS(", ")", "strftime('%Y%m%d', ", ")")
            sResult = SQLDIAPAR(sResult, "DAY(", ")", "cast(strftime('%d', ", ") as int)")
            sResult = SQLDIAPAR(sResult, "MONTH(", ")", "cast(strftime('%m', ", ") as int)")
            sResult = SQLDIAPAR(sResult, "YEAR(", ")", "cast(strftime('%Y', ", ") as int)")

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
            sResult = Replace(sResult, "AT(", "CHARINDEX(")
            sResult = SQLDIAPAR(sResult, "DTOS(", ")", "CONVERT(char(8), ", ", 11)")
    End Select
    
    sqldialeto = sResult
End Function

Public Function SQLDIAPAR(ByVal cSQLCNV As String, ByVal cDelimIni As String, ByVal cDelimFim As String, ByVal cREPINI As String, ByVal cREPFIM As String) As String
    Dim aRETU As Variant
    SQLDIAPAR = cSQLCNV
    aRETU = pegue2delimitado(cSQLCNV, cDelimIni, cDelimFim)
    If Len(aRETU(0)) > 0 Then
       SQLDIAPAR = Replace(aRETU(2), cDelimIni, "")
       SQLDIAPAR = SQLDIAPAR & cREPINI & aRETU(0) & cREPFIM & aRETU(1)
    End If
End Function

' Função auxiliar interna para suporte ao SQLDIAPAR
Private Function pegue2delimitado(ByVal cTEXTO As String, ByVal cIni As String, ByVal cFim As String) As Variant
    Dim nPos1 As Long, nPOS2 As Long
    Dim aRet(2) As String
    nPos1 = InStr(UCase(cTEXTO), UCase(cIni))
    If nPos1 > 0 Then
        nPOS2 = InStr(nPos1 + Len(cIni), UCase(cTEXTO), UCase(cFim))
        If nPOS2 > nPos1 Then
            aRet(0) = Mid(cTEXTO, nPos1 + Len(cIni), nPOS2 - (nPos1 + Len(cIni)))
            aRet(1) = Mid(cTEXTO, nPOS2 + Len(cFim))
            aRet(2) = Left(cTEXTO, nPos1 - 1)
        End If
    End If
    pegue2delimitado = aRet
End Function
' ==============================================================================
' FUNÇÃO: dialetoDataParaSQL
' Objetivo: Normaliza e formata datas para o formato de cada motor SQL,
'           respeitando as TAGS retornadas pela função TipoConn.
' ==============================================================================
Public Function dialetoDataParaSQL(ByVal vData As Variant, ByVal cTipoBanco As String) As Variant
    
    ' 1. Definição do que é considerado "Vazio" (Limpeza padrão)
    Dim bEhVazio As Boolean
    bEhVazio = IsNull(vData) Or DataBranco(vData) Or Trim(vData & "") = "" Or vData = "  /  /    "
    
    ' 2. Lógica específica por Banco de Dados (Dialetos)
    Select Case UCase(cTipoBanco)
        
        Case "SQLITE"
            If bEhVazio Then dialetoDataParaSQL = Null Else dialetoDataParaSQL = Format(vData, "yyyy-mm-dd hh:nn:ss")
            
        Case "MDB", "ACCDB"
            ' Access/Jet exige o delimitador #
            If bEhVazio Then dialetoDataParaSQL = Null Else dialetoDataParaSQL = "#" & Format(vData, "mm/dd/yyyy") & "#"
            
        Case "PGSQL", "POSTGRESQL"
            If bEhVazio Then dialetoDataParaSQL = Null Else dialetoDataParaSQL = "'" & Format(vData, "yyyy-mm-dd hh:nn:ss") & "'"
            
        Case "MYSQL", "MARIADB"
            If bEhVazio Then dialetoDataParaSQL = Null Else dialetoDataParaSQL = "'" & Format(vData, "yyyy-mm-dd hh:nn:ss") & "'"
            
        Case "FIREBIRD"
            If bEhVazio Then dialetoDataParaSQL = Null Else dialetoDataParaSQL = "'" & Format(vData, "yyyy-mm-dd hh:nn:ss") & "'"
            
        Case "ORACLE"
            ' Oracle exige o TO_DATE para garantir precisão e evitar erros de NLS
            If bEhVazio Then
                dialetoDataParaSQL = Null
            Else
                dialetoDataParaSQL = "TO_DATE('" & Format(vData, "yyyy-mm-dd hh:nn:ss") & "', 'YYYY-MM-DD HH24:MI:SS')"
            End If
            
        Case "VFPOLEDB", "FOXPRO"
            ' FoxPro é sensível a NULL, retornamos vazio para evitar erros de Constraint
            If bEhVazio Then
                dialetoDataParaSQL = ""
            Else
                dialetoDataParaSQL = "{^" & Format(vData, "yyyy-mm-dd hh:nn:ss") & "}"
            End If
            
        Case Else
            ' Comportamento genérico ANSI (Data simples)
            If bEhVazio Then dialetoDataParaSQL = Null Else dialetoDataParaSQL = "'" & Format(vData, "yyyy-mm-dd") & "'"
            
    End Select
End Function

Public Function SQLPGSQLDOUBLEQUOTES(ByVal cSQL As String) As String
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.IgnoreCase = True
    regEx.Pattern = "\b([a-z_][a-z0-9_]*)\b(?=(?:[^']*'[^']*')*[^']*$)"
    
    Dim sResult As String
    sResult = regEx.Replace(cSQL, """$1""")
    sResult = Replace(sResult, """SELECT""", "SELECT")
    sResult = Replace(sResult, """FROM""", "FROM")
    sResult = Replace(sResult, """WHERE""", "WHERE")
    SQLPGSQLDOUBLEQUOTES = sResult
End Function
Public Function FormataBlocoSql(ByVal cTextoBruto As String) As String
    Dim aLinhas() As String, i As Long, cTextoFormatado As String, cInstrucao As String, cLinhaLimpa As String
    Dim nPosAbre As Long, nPosFecha As Long, cHeader As String, cMiolo As String, aCampos() As String, j As Long, cCAMPO As String, nContCampo As Long
    
    aLinhas = Split(cTextoBruto, vbCrLf)
    cTextoFormatado = "": cInstrucao = ""
    
    For i = LBound(aLinhas) To UBound(aLinhas)
        cLinhaLimpa = Trim(aLinhas(i))
        If InStr(UCase(cLinhaLimpa), "GRANT ") > 0 Or InStr(UCase(cLinhaLimpa), "ENGINE=") > 0 Or InStr(UCase(cLinhaLimpa), "PRAGMA ") > 0 Then
            cTextoFormatado = cTextoFormatado & cLinhaLimpa & vbCrLf
            GoTo ProximaLinha
        End If
        
        If cLinhaLimpa = "" Or cLinhaLimpa = ";" Or (Left(cLinhaLimpa, 1) = "[" And Right(cLinhaLimpa, 1) = "]") Then
            If cLinhaLimpa <> "" And cLinhaLimpa <> ";" Then cTextoFormatado = cTextoFormatado & cLinhaLimpa & vbCrLf
            GoTo ProximaLinha
        End If
        
        If Right(cLinhaLimpa, 1) = ";" Then cLinhaLimpa = Trim(Left(cLinhaLimpa, Len(cLinhaLimpa) - 1))
        cInstrucao = cInstrucao & IIf(cInstrucao = "", "", " ") & cLinhaLimpa
        
        If InStr(UCase(cInstrucao), "CREATE TABLE") = 0 And InStr(UCase(cInstrucao), "INDEX") = 0 And InStr(UCase(cInstrucao), "ALTER") = 0 Then
            GoTo ProximaLinha
        End If
        
        If InStr(UCase(cInstrucao), "CREATE TABLE") > 0 Then
            If InStr(cLinhaLimpa, ")") = 0 Then GoTo ProximaLinha
            nPosAbre = InStr(cInstrucao, "(")
            nPosFecha = InStrRev(cInstrucao, ")")
            
            If nPosAbre > 0 Then
                cHeader = Left(cInstrucao, nPosAbre)
                cMiolo = Mid(cInstrucao, nPosAbre + 1, IIf(nPosFecha > 0, nPosFecha - nPosAbre - 1, Len(cInstrucao)))
                cTextoFormatado = cTextoFormatado & Trim(cHeader) & vbCrLf
                
                aCampos = Split(cMiolo, ",")
                nContCampo = 0
                For j = LBound(aCampos) To UBound(aCampos)
                    cCAMPO = Trim(aCampos(j))
                    If InStr(cCAMPO, "(") > 0 And InStr(cCAMPO, ")") = 0 And j < UBound(aCampos) Then
                        j = j + 1
                        cCAMPO = cCAMPO & "," & Trim(aCampos(j))
                    End If
                    If cCAMPO <> "" Then
                        nContCampo = nContCampo + 1
                        If nContCampo = 1 Then
                            cTextoFormatado = cTextoFormatado & "    " & cCAMPO & vbCrLf
                        Else
                            cTextoFormatado = cTextoFormatado & "    ," & cCAMPO & vbCrLf
                        End If
                    End If
                Next j
                cTextoFormatado = cTextoFormatado & ") ;" & vbCrLf
            Else
                cTextoFormatado = cTextoFormatado & cInstrucao & " ;" & vbCrLf
            End If
            cInstrucao = ""
        ElseIf InStr(UCase(cInstrucao), "CREATE") > 0 And InStr(UCase(cInstrucao), "INDEX") > 0 Then
            cInstrucao = Replace(cInstrucao, " ;", "")
            cTextoFormatado = cTextoFormatado & Trim(cInstrucao) & " ;" & vbCrLf
            cInstrucao = ""
        Else
            cTextoFormatado = cTextoFormatado & Trim(cInstrucao) & " ;" & vbCrLf
            cInstrucao = ""
        End If
ProximaLinha:
    Next i
    FormataBlocoSql = cTextoFormatado
End Function

' -------------------------------------------------------------------------
' 7. DETECÇÃO DE ALVO E METADADOS (ESTRUTURA E PK)
' -------------------------------------------------------------------------
Public Function DIALETO_DetectTargetDb(ByVal cTargetDB As String) As String
    Dim cSystemID As String
    cTargetDB = UCase(cTargetDB)
    cSystemID = ""
    
    If InStr(cTargetDB, "ORACLE") > 0 Then
        cSystemID = "ORACLE"
    ElseIf InStr(cTargetDB, "SQL SERVER") > 0 Or InStr(cTargetDB, "MSSQL") > 0 Then
        cSystemID = "MSSQL"
    ElseIf InStr(cTargetDB, "SQLITE") > 0 Then
        cSystemID = "SQLITE"
    ElseIf InStr(cTargetDB, "FOXPRO") > 0 Then
        cSystemID = "FOXPRO"
    ElseIf InStr(cTargetDB, "SYBASE") > 0 Then
        cSystemID = "SYBASE"
    ElseIf InStr(cTargetDB, "ACCESS") > 0 Then
        cSystemID = "ACCESS"
    ElseIf InStr(cTargetDB, "INGRES") > 0 Then
        cSystemID = "INGRES"
    ElseIf InStr(cTargetDB, "INFORMIX") > 0 Then
        cSystemID = "INFORMIX"
    ElseIf InStr(cTargetDB, "ADABAS") > 0 Then
        cSystemID = "ADABAS"
    ElseIf InStr(cTargetDB, "POSTGRESQL") > 0 Then
        cSystemID = "POSTGRESQL"
    ElseIf InStr(cTargetDB, "DB2") > 0 Then
        cSystemID = "DB2"
    ElseIf InStr(cTargetDB, "MYSQL") > 0 Then
        cSystemID = "MYSQL"
    ElseIf InStr(cTargetDB, "MARIADB") > 0 Then
        cSystemID = "MARIADB"
    ElseIf InStr(cTargetDB, "FIREBIRD") > 0 Or InStr(cTargetDB, "INTERBASE") > 0 Then
        cSystemID = "FIREBIRD"
    ElseIf InStr(cTargetDB, "CUBRID") > 0 Then
        cSystemID = "CUBRID"
    ElseIf InStr(cTargetDB, "DUCKDB") > 0 Then
        cSystemID = "DUCKDB"
    ElseIf InStr(cTargetDB, "DUCKLAKE") > 0 Then
        cSystemID = "DUCKLAKE"
    End If
    DIALETO_DetectTargetDb = cSystemID
End Function

Public Function DIALETO_Estrutura(ByVal cTargetDB As String, ByVal cTABELA As String, Optional ByVal cUsuario As String = "") As String
    Dim CCOMANDO As String, cSchemaSQL As String
    cTargetDB = UCase(cTargetDB)
    CCOMANDO = ""
    
    Select Case cTargetDB
        Case "SQLITE"
            CCOMANDO = "SELECT name AS FIELD_NAME, type AS DATA_TYPE, 0 AS FIELD_LEN, 0 AS FIELD_DEC FROM pragma_table_info('" & cTABELA & "');"
        Case "MYSQL", "MYSQL64", "MARIADB"
            If Trim(cUsuario) = "" Then
                CCOMANDO = "SELECT COLUMN_NAME AS FIELD_NAME, DATA_TYPE AS DATA_TYPE, COALESCE(CHARACTER_MAXIMUM_LENGTH, NUMERIC_PRECISION) AS FIELD_LEN, COALESCE(NUMERIC_SCALE, 0) AS FIELD_DEC FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" & cTABELA & "' AND TABLE_SCHEMA = DATABASE() ORDER BY ORDINAL_POSITION;"
            Else
                CCOMANDO = "SELECT COLUMN_NAME AS FIELD_NAME, DATA_TYPE AS DATA_TYPE, COALESCE(CHARACTER_MAXIMUM_LENGTH, NUMERIC_PRECISION) AS FIELD_LEN, COALESCE(NUMERIC_SCALE, 0) AS FIELD_DEC FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" & cTABELA & "' AND TABLE_SCHEMA = '" & cUsuario & "' ORDER BY ORDINAL_POSITION;"
            End If
        Case "PGSQL", "PGSQL64", "POSTGRESQL"
            cSchemaSQL = IIf(Trim(cUsuario) = "", "public", cUsuario)
            CCOMANDO = "SELECT column_name AS FIELD_NAME, udt_name AS DATA_TYPE, COALESCE(character_maximum_length, numeric_precision) AS FIELD_LEN, COALESCE(numeric_scale, 0) AS FIELD_DEC FROM information_schema.columns WHERE LOWER(table_name) = '" & LCase(cTABELA) & "' AND table_schema = '" & cSchemaSQL & "' ORDER BY ordinal_position;"
        Case "MSSQL", "SQLSERVER"
            cSchemaSQL = IIf(Trim(cUsuario) = "", "dbo", cUsuario)
            CCOMANDO = "SELECT COLUMN_NAME AS FIELD_NAME, DATA_TYPE AS DATA_TYPE, ISNULL(CHARACTER_MAXIMUM_LENGTH, NUMERIC_PRECISION) AS FIELD_LEN, ISNULL(NUMERIC_SCALE, 0) AS FIELD_DEC FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" & cTABELA & "' AND TABLE_SCHEMA = '" & cSchemaSQL & "' ORDER BY ORDINAL_POSITION;"
        Case "ORACLE", "OCI"
            Dim cUserOracle As String
            cUserOracle = IIf(Trim(cUsuario) = "", "USER_TAB_COLUMNS", "ALL_TAB_COLUMNS")
            CCOMANDO = "SELECT COLUMN_NAME AS FIELD_NAME, DATA_TYPE AS DATA_TYPE, DATA_LENGTH AS FIELD_LEN, COALESCE(DATA_SCALE, 0) AS FIELD_DEC FROM " & cUserOracle & " WHERE TABLE_NAME = '" & UCase(cTABELA) & "' " & IIf(Trim(cUsuario) <> "", "AND OWNER = '" & UCase(cUsuario) & "' ", "") & "ORDER BY COLUMN_ID;"
        Case "FIREBIRD"
            CCOMANDO = "SELECT TRIM(F.RDB$FIELD_NAME) AS FIELD_NAME, CASE T.RDB$FIELD_TYPE WHEN 7 THEN 'SMALLINT' WHEN 8 THEN 'INTEGER' WHEN 16 THEN 'BIGINT' WHEN 10 THEN 'FLOAT' WHEN 27 THEN 'DOUBLE PRECISION' WHEN 14 THEN 'CHAR' WHEN 37 THEN 'VARCHAR' WHEN 12 THEN 'DATE' WHEN 13 THEN 'TIME' WHEN 35 THEN 'TIMESTAMP' WHEN 261 THEN 'BLOB' END AS DATA_TYPE, T.RDB$FIELD_LENGTH AS FIELD_LEN, COALESCE(T.RDB$FIELD_SCALE, 0) * -1 AS FIELD_DEC FROM RDB$RELATION_FIELDS F JOIN RDB$FIELDS T ON F.RDB$FIELD_SOURCE = T.RDB$FIELD_NAME WHERE F.RDB$RELATION_NAME = '" & UCase(cTABELA) & "' ORDER BY F.RDB$FIELD_POSITION;"
        Case "DUCKDB", "DUCKLAKE"
            cSchemaSQL = IIf(Trim(cUsuario) = "", "main", LCase(cUsuario))
            CCOMANDO = "SELECT column_name AS FIELD_NAME, data_type AS DATA_TYPE, COALESCE(character_maximum_length, numeric_precision) AS FIELD_LEN, COALESCE(numeric_scale, 0) AS FIELD_DEC FROM information_schema.columns WHERE table_name = '" & cTABELA & "' AND table_schema = '" & cSchemaSQL & "' ORDER BY ordinal_position;"
    End Select
    DIALETO_Estrutura = CCOMANDO
End Function

Public Function Dialeto_PK(ByVal cTargetDB As String, ByVal cTABELA As String, Optional ByVal cUsuario As String = "") As String
    Dim CCOMANDO As String, cSchema As String, cSchemaSQL As String, cUserOracle As String, cColOracle As String
    cTargetDB = UCase(cTargetDB)
    CCOMANDO = ""
    
    Select Case cTargetDB
        Case "SQLITE"
            CCOMANDO = "SELECT name AS FIELD_NAME FROM pragma_table_info('" & cTABELA & "') WHERE pk > 0 ORDER BY pk;"
        Case "MYSQL", "MYSQL64", "MARIADB"
            cSchema = IIf(Trim(cUsuario) = "", "DATABASE()", "'" & cUsuario & "'")
            CCOMANDO = "SELECT kcu.COLUMN_NAME AS FIELD_NAME FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE kcu ON tc.CONSTRAINT_NAME = kcu.CONSTRAINT_NAME AND tc.TABLE_SCHEMA = kcu.TABLE_SCHEMA WHERE tc.CONSTRAINT_TYPE = 'PRIMARY KEY' AND tc.TABLE_NAME = '" & cTABELA & "' AND tc.TABLE_SCHEMA = " & cSchema & " ORDER BY kcu.ORDINAL_POSITION;"
        Case "PGSQL", "PGSQL64", "POSTGRESQL"
            cSchema = IIf(Trim(cUsuario) = "", "public", LCase(cUsuario))
            CCOMANDO = "SELECT c.attname AS FIELD_NAME FROM pg_class a, pg_class b, pg_attribute c, pg_index d, pg_namespace e WHERE a.oid = d.indrelid AND a.relname = '" & LCase(cTABELA) & "' AND b.oid = d.indexrelid AND c.attrelid = b.oid AND d.indisprimary AND e.oid = a.relnamespace AND e.nspname = '" & cSchema & "';"
        Case "MSSQL", "SQLSERVER"
            cSchemaSQL = IIf(Trim(cUsuario) = "", "dbo", cUsuario)
            CCOMANDO = "SELECT kcu.COLUMN_NAME AS FIELD_NAME FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS tc JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE kcu ON tc.CONSTRAINT_NAME = kcu.CONSTRAINT_NAME WHERE tc.CONSTRAINT_TYPE = 'PRIMARY KEY' AND tc.TABLE_NAME = '" & cTABELA & "' AND tc.TABLE_SCHEMA = '" & cSchemaSQL & "' ORDER BY kcu.ORDINAL_POSITION;"
        Case "ORACLE", "OCI"
            cUserOracle = IIf(Trim(cUsuario) = "", "USER_CONSTRAINTS", "ALL_CONSTRAINTS")
            cColOracle = IIf(Trim(cUsuario) = "", "USER_CONS_COLUMNS", "ALL_CONS_COLUMNS")
            CCOMANDO = "SELECT cols.column_name AS FIELD_NAME FROM " & cUserOracle & " cons JOIN " & cColOracle & " cols ON cons.constraint_name = cols.constraint_name " & IIf(Trim(cUsuario) <> "", "AND cons.owner = cols.owner ", "") & "WHERE cons.constraint_type = 'P' AND cons.table_name = '" & UCase(cTABELA) & "' " & IIf(Trim(cUsuario) <> "", "AND cons.owner = '" & UCase(cUsuario) & "' ", "") & "ORDER BY cols.position;"
        Case "FIREBIRD"
            CCOMANDO = "SELECT TRIM(S.RDB$FIELD_NAME) AS FIELD_NAME FROM RDB$INDEX_SEGMENTS S JOIN RDB$RELATION_CONSTRAINTS RC ON S.RDB$INDEX_NAME = RC.RDB$INDEX_NAME WHERE RC.RDB$CONSTRAINT_TYPE = 'PRIMARY KEY' AND RC.RDB$RELATION_NAME = '" & UCase(cTABELA) & "' ORDER BY S.RDB$FIELD_POSITION;"
        Case "DUCKDB", "DUCKLAKE"
            cSchemaSQL = IIf(Trim(cUsuario) = "", "main", LCase(cUsuario))
            CCOMANDO = "SELECT kcu.column_name AS FIELD_NAME FROM information_schema.table_constraints tc JOIN information_schema.key_column_usage kcu ON tc.constraint_name = kcu.constraint_name WHERE tc.constraint_type = 'PRIMARY KEY' AND tc.table_name = '" & cTABELA & "' AND tc.table_schema = '" & cSchemaSQL & "' ORDER BY kcu.ordinal_position;"
    End Select
    Dialeto_PK = CCOMANDO
End Function

