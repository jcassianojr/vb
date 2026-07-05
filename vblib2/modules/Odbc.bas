Attribute VB_Name = "Odbc"
Option Explicit

#If (VBA7 = 0) Then
Private Enum LongPtr
[_]
End Enum
#End If

#If Win64 Then
Private Const NULL_PTR As LongPtr = 0
Private Const PTR_SIZE As Long = 8
#Else
Private Const NULL_PTR As Long = 0
Private Const PTR_SIZE As Long = 4
#End If

' --- DECLARAÇÕES DE API COM SUPORTE A 32-BIT / 64-BIT (VBA7) ---
#If VBA7 Or Win64 Then
    Public Declare PtrSafe Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hWndParent As LongPtr, ByVal fRequest As LongPtr, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
    Private Declare PtrSafe Function SQLInstallerError Lib "odbccp32.dll" (ByVal iError As Integer, ByRef pfErrorCode As Long, ByVal lpszErrorMsg As String, ByVal cbErrorMsgMax As Integer, ByRef pcbErrorMsg As Integer) As Integer
    Private Declare PtrSafe Function SQLGetPrivateProfileString Lib "odbccp32.dll" Alias "SQLGetPrivateProfileString" (ByVal lpszSection As String, ByVal lpszEntry As String, ByVal lpszDefault As String, ByVal RetBuffer As String, ByVal cbRetBuffer As Long, ByVal lpszFilename As String) As Long
#Else
    Public Declare Function SQLConfigDataSource Lib "odbccp32.dll" (ByVal hWndParent As Long, ByVal fRequest As Long, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
    Private Declare Function SQLInstallerError Lib "odbccp32.dll" (ByVal iError As Integer, ByRef pfErrorCode As Long, ByVal lpszErrorMsg As String, ByVal cbErrorMsgMax As Integer, ByRef pcbErrorMsg As Integer) As Integer
    Private Declare Function SQLGetPrivateProfileString Lib "odbccp32.dll" (ByVal lpszSection As String, ByVal lpszEntry As String, ByVal lpszDefault As String, ByVal RetBuffer As String, ByVal cbRetBuffer As Long, ByVal lpszFileName As String) As Long
#End If

' Constantes para gerenciamento de DSN
Private Const ODBC_ADD_DSN = 1
Private Const ODBC_CONFIG_DSN = 2
Private Const ODBC_REMOVE_DSN = 3
Private Const ODBC_ADD_SYS_DSN = 4
Public Function DriverExisteOdbc(cNOME As String, bE_DriverODBC As Boolean) As Boolean
    On Error GoTo TrataErro
    
    If bE_DriverODBC Then
        ' Verifica Driver ODBC via Registro
        Dim oShell As Object
        Set oShell = CreateObject("WScript.Shell")
        ' Tenta ler a subchave padrão de Drivers ODBC
        Dim sReg As String
        sReg = "HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBCINST.INI\" & cNOME & "\Driver"
        oShell.RegRead sReg
        DriverExisteOdbc = True
        Set oShell = Nothing
    Else
        ' Verifica Provider OLEDB via criação de objeto COM
        Dim oCONN As Object
        Set oCONN = CreateObject(cNOME)
        DriverExisteOdbc = True
        Set oCONN = Nothing
    End If
    Exit Function

TrataErro:
    DriverExisteOdbc = False
    Resume Proximo
Proximo:
End Function
' --- FUNÇÃO REESCRITA E AMPLIADA ---
Public Function AddDSN(ByVal strDSN As String, _
                         ByVal strDescription As String, _
                         ByVal strDB As String, _
                         Optional ByVal strDriverType As String = "MDB", _
                         Optional ByVal lUSER As Boolean = False, _
                         Optional ByVal sHost As String = "", _
                         Optional ByVal sUser As String = "", _
                         Optional ByVal sPass As String = "", _
                         Optional ByVal sPort As String = "") As Boolean



    Dim strAttributes As String
    Dim StrDriver As String
    
    On Error GoTo Hell

    ' 1. Verifica se o DSN já existe
    If DSNExists(strDSN) Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox("O DSN '" & strDSN & "' já existe. Deseja substituí-lo?", vbYesNo + vbQuestion, "DSN Existente")
        If resp = vbNo Then
            AddDSN = False
            Exit Function
        End If
        
        ' Remove o DSN antigo antes de recriar
        Call RemoverDSN(strDSN, strDriverType)
    End If

    ' 2. Resolve o Driver correto e monta a string de atributos específicos
    Select Case UCase(strDriverType)
        Case "MDB", "ACCDB"
            StrDriver = "Microsoft Access Driver (*.mdb, *.accdb)"
            strAttributes = "DBQ=" & strDB & Chr(0) & "Exclusive=0" & Chr(0) & "ReadOnly=0" & Chr(0)
            
        Case "DBF"
            StrDriver = "Microsoft Visual FoxPro Driver"
            strAttributes = "SourceDB=" & strDB & Chr(0) & "SourceType=DBF" & Chr(0) & "Exclusive=0" & Chr(0)
            
        Case "SQLITE"
            StrDriver = "SQLite3 ODBC Driver"
            strAttributes = "Database=" & strDB & Chr(0)
            
        Case "SQLSERVER", "MSSQL"
            StrDriver = "SQL Server"
            strAttributes = "SERVER=" & sHost & Chr(0) & "DATABASE=" & strDB & Chr(0)
            
        Case "MYSQL"
            StrDriver = "MySQL ODBC 8.0 ANSI Driver"
            If sPort = "" Then sPort = "3306"
            strAttributes = "SERVER=" & sHost & Chr(0) & _
                            "DATABASE=" & strDB & Chr(0) & _
                            "USER=" & sUser & Chr(0) & _
                            "PASSWORD=" & sPass & Chr(0) & _
                            "PORT=" & sPort & Chr(0) & _
                            "OPTION=3" & Chr(0)

        Case "MARIADB"
            StrDriver = "MariaDB ODBC 3.0 Driver"
            If sPort = "" Then sPort = "3306"
            strAttributes = "SERVER=" & sHost & Chr(0) & _
                            "DATABASE=" & strDB & Chr(0) & _
                            "USER=" & sUser & Chr(0) & _
                            "PASSWORD=" & sPass & Chr(0) & _
                            "PORT=" & sPort & Chr(0)

        Case "POSTGRESQL", "PGSQL"
            StrDriver = "PostgreSQL ANSI"
            If sPort = "" Then sPort = "5432"
            strAttributes = "SERVER=" & sHost & Chr(0) & _
                            "DATABASE=" & strDB & Chr(0) & _
                            "UID=" & sUser & Chr(0) & _
                            "PWD=" & sPass & Chr(0) & _
                            "PORT=" & sPort & Chr(0)

        Case "ORACLE"
            StrDriver = "Oracle in OraClient19Home1"
            strAttributes = "DBQ=" & strDB & Chr(0) & _
                            "UID=" & sUser & Chr(0) & _
                            "PWD=" & sPass & Chr(0)
        Case Else
            MsgBox "Tipo de driver não suportado no assistente.", vbExclamation
            AddDSN = False
            Exit Function
    End Select

    ' 3. Adiciona os atributos comuns (DSN e Descrição) no início
    strAttributes = "DSN=" & strDSN & Chr(0) & _
                    "DESCRIPTION=" & strDescription & Chr(0) & _
                    strAttributes

    ' 4. Executa a criação (User DSN ou System DSN)
    Dim intRet As Long
    If lUSER Then
        intRet = SQLConfigDataSource(0&, ODBC_ADD_DSN, StrDriver, strAttributes)
    Else
        intRet = SQLConfigDataSource(0&, ODBC_ADD_SYS_DSN, StrDriver, strAttributes)
    End If

    If intRet <> 0 Then
        AddDSN = True
    Else
        AddDSN = False
        Call ShowODBCError
    End If

    Exit Function
Hell:
    AddDSN = False
    MsgBox "Erro inesperado: " & Err.Description, vbCritical
End Function

' --- FUNÇÕES AUXILIARES DA MESCLA ---

Public Function DSNExists(ByVal dsnName As String) As Boolean
    Dim sBuf As String * 256
    Dim nRet As Long
    nRet = SQLGetPrivateProfileString("ODBC Data Sources", dsnName, "", sBuf, 256, "ODBC.INI")
    DSNExists = (nRet > 0)
End Function

Private Sub RemoverDSN(ByVal dsnName As String, ByVal strDriverType As String)
    Dim StrDriver As String
    Select Case UCase(strDriverType)
        Case "MDB", "ACCDB": StrDriver = "Microsoft Access Driver (*.mdb, *.accdb)"
        Case "DBF":          StrDriver = "Microsoft Visual FoxPro Driver"
        Case "SQLITE":       StrDriver = "SQLite3 ODBC Driver"
        Case "SQLSERVER", "MSSQL": StrDriver = GetBestMSSQL("D") '"SQL Server"
        Case "MYSQL":        StrDriver = "MySQL ODBC 8.0 ANSI Driver"
        Case "MARIADB":      StrDriver = "MariaDB ODBC 3.0 Driver"
        Case "POSTGRESQL", "PGSQL": StrDriver = "PostgreSQL ANSI"
        Case "ORACLE":       StrDriver = "Oracle in OraClient19Home1"
        Case "FIREBIRD":     StrDriver = FirebirdODBC()
    End Select
    Call SQLConfigDataSource(0&, ODBC_REMOVE_DSN, StrDriver, "DSN=" & dsnName & Chr$(0) & Chr$(0))
End Sub
Public Function FirebirdODBC() As String
    Dim shell As Object
    Dim sRegKey As String
    Dim bAchouV3 As Boolean
    Dim bAchouV2 As Boolean
    
    Set shell = CreateObject("WScript.Shell")
    
    ' 1. TESTA SE A VERSÃO 3 ESTÁ INSTALADA
    On Error Resume Next
    sRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers\Firebird ODBC Driver"
    shell.RegRead sRegKey
    bAchouV3 = (Err.Number = 0)
    
    If Not bAchouV3 Then
        Err.Clear
        sRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\ODBC\ODBCINST.INI\ODBC Drivers\Firebird ODBC Driver"
        shell.RegRead sRegKey
        bAchouV3 = (Err.Number = 0)
    End If
    
    ' Se achou a Versão 3, retorna a string correspondente imediatamente
    If bAchouV3 Then
        FirebirdODBC = "Firebird ODBC Driver"
        Set shell = Nothing
        Exit Function
    End If

    ' 2. TESTA SE A VERSÃO 2 ESTÁ INSTALADA (CASO NÃO TENHA A V3)
    Err.Clear
    sRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers\Firebird/InterBase(r) driver"
    shell.RegRead sRegKey
    bAchouV2 = (Err.Number = 0)
    
    If Not bAchouV2 Then
        Err.Clear
        sRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\ODBC\ODBCINST.INI\ODBC Drivers\Firebird/InterBase(r) driver"
        shell.RegRead sRegKey
        bAchouV2 = (Err.Number = 0)
    End If
    
    ' Se achou a Versão 2, retorna a string dela
    If bAchouV2 Then
        FirebirdODBC = "Firebird/InterBase(r) driver"
    Else
        ' Se nenhum driver for localizado no Windows
        FirebirdODBC = ""
    End If
    
    Set shell = Nothing
End Function
Public Function GetBestMSSQL(TIPO As String) As String
    Dim SupportedDrivers, SupportedProviders
    Dim item As Variant
    
    SupportedDrivers = Array("ODBC Driver 17 for SQL Server", "ODBC Driver 13 for SQL Server", "SQL Server Native Client 11.0", "SQL Server")
    SupportedProviders = Array("MSOLEDBSQL19", "MSOLEDBSQL", "SQLNCLI11", "SQLOLEDB")

    If TIPO = "D" Then
        For Each item In SupportedDrivers
            ' AQUI: Você deve passar o nome (item) e True (pois é driver ODBC)
            If DriverExisteOdbc(CStr(item), True) Then
                GetBestMSSQL = CStr(item)
                Exit Function
            End If
        Next
    Else
        For Each item In SupportedProviders
            ' AQUI: Você deve passar o nome (item) e False (pois é Provider OLEDB)
            If DriverExisteOdbc(CStr(item), False) Then
                GetBestMSSQL = CStr(item)
                Exit Function
            End If
        Next
    End If
End Function
Public Sub ShowODBCError()
    Dim sMsg As String * 512
    Dim lErr As Long
    Dim iLen As Integer
    Dim r As Integer
    r = SQLInstallerError(1, lErr, sMsg, 512, iLen)
    If iLen > 0 Then
        MsgBox "Erro no Instalador ODBC " & lErr & ": " & Left$(sMsg, iLen), vbCritical, "Erro ODBC"
    Else
        MsgBox "Falha na operação ODBC, mas nenhuma informação detalhada foi retornada.", vbCritical, "Erro Odbc"
    End If
End Sub

Public Function IsDriverInstalled(ByVal sDriverName As String) As Boolean
    Dim shell As Object
    Dim sRegKey As String
    On Error Resume Next
    Set shell = CreateObject("WScript.Shell")
    sRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers\" & sDriverName
    shell.RegRead sRegKey
    If Err.Number = 0 Then
        IsDriverInstalled = True
    Else
        Err.Clear
        sRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\ODBC\ODBCINST.INI\ODBC Drivers\" & sDriverName
        shell.RegRead sRegKey
        IsDriverInstalled = (Err.Number = 0)
    End If
    Set shell = Nothing
End Function

