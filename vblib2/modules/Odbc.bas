Attribute VB_Name = "sqlOdbc"
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

' --- DECLARAÇÕES DE API ODBC COM SUPORTE A 32-BIT / 64-BIT ---
#If VBA7 Or Win64 Then
    Public Declare PtrSafe Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hWndParent As LongPtr, ByVal fRequest As Integer, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
    Private Declare PtrSafe Function SQLInstallerError Lib "odbccp32.dll" (ByVal iError As Integer, ByRef pfErrorCode As Long, ByVal lpszErrorMsg As String, ByVal cbErrorMsgMax As Integer, ByRef pcbErrorMsg As Integer) As Integer
    
    ' API do Registro do Windows (Substituindo WMI)
    Private Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As LongPtr, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As LongPtr) As Long
    Private Declare PtrSafe Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal lpReserved As LongPtr, lpType As Long, lpData As Any, lpcbData As Long) As Long
    Private Declare PtrSafe Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As LongPtr, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As LongPtr, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
    Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As LongPtr) As Long
#Else
    Public Declare Function SQLConfigDataSource Lib "odbccp32.dll" (ByVal hWndParent As Long, ByVal fRequest As Integer, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
    Private Declare Function SQLInstallerError Lib "odbccp32.dll" (ByVal iError As Integer, ByRef pfErrorCode As Long, ByVal lpszErrorMsg As String, ByVal cbErrorMsgMax As Integer, ByRef pcbErrorMsg As Integer) As Integer
    
    ' API do Registro do Windows (Substituindo WMI)
    Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
    Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
    Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
    Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
#End If

' Constantes ODBC
Private Const ODBC_ADD_DSN As Integer = 1
Private Const ODBC_CONFIG_DSN As Integer = 2
Private Const ODBC_REMOVE_DSN As Integer = 3
Private Const ODBC_ADD_SYS_DSN As Integer = 4
Private Const ODBC_CONFIG_SYS_DSN As Integer = 5
Private Const ODBC_REMOVE_SYS_DSN As Integer = 6

' Constantes do Registro
Private Const HKEY_CLASSES_ROOT As Long = &H80000000
Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const HKEY_LOCAL_MACHINE As Long = &H80000002
Private Const KEY_READ As Long = &H20019
Private Const REG_SZ As Long = 1

Public Function DriverExisteOdbc(cNOME As String, bE_DriverODBC As Boolean) As Boolean
    If bE_DriverODBC Then
        DriverExisteOdbc = IsDriverInstalled(cNOME)
    Else
        DriverExisteOdbc = OleDbProviderExists(cNOME)
    End If
End Function

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
    Dim strDriver As String
    Dim strType As String
    Dim intRet As Long
    Dim nRequest As Integer
    Dim resp As VbMsgBoxResult

    On Error GoTo TrataErro
    AddDSN = False
    strType = UCase$(Trim$(strDriverType))

    If Len(Trim$(strDSN)) = 0 Then
        MsgBox "O nome do DSN nao pode ficar vazio.", vbExclamation
        Exit Function
    End If

    If Not MontarAtributosDSN(strType, strDB, sHost, sUser, sPass, sPort, strDriver, strAttributes) Then
        MsgBox "Tipo de driver nao suportado ou driver indisponivel: " & strDriverType, vbExclamation
        Exit Function
    End If

    If DSNExistsEscopo(strDSN, lUSER) Then
        resp = MsgBox("O DSN '" & strDSN & "' ja existe. Deseja substitui-lo?", _
                      vbYesNo + vbQuestion, "DSN Existente")
        If resp <> vbYes Then Exit Function
        If lUSER Then
            nRequest = ODBC_CONFIG_DSN
        Else
            nRequest = ODBC_CONFIG_SYS_DSN
        End If
    ElseIf lUSER Then
        nRequest = ODBC_ADD_DSN
    Else
        nRequest = ODBC_ADD_SYS_DSN
    End If

    strAttributes = "DSN=" & strDSN & Chr$(0) & _
                    "DESCRIPTION=" & strDescription & Chr$(0) & strAttributes
    intRet = SQLConfigDataSource(NULL_PTR, nRequest, strDriver, strAttributes)

    If intRet <> 0 Then
        AddDSN = True
    Else
        ShowODBCError
    End If
    Exit Function

TrataErro:
    AddDSN = False
    MsgBox "Erro inesperado ao configurar o DSN: " & Err.Description, vbCritical
End Function

Private Function MontarAtributosDSN(ByVal strType As String, ByVal strDB As String, _
                                    ByVal sHost As String, ByVal sUser As String, _
                                    ByVal sPass As String, ByVal sPort As String, _
                                    ByRef strDriver As String, ByRef strAttributes As String) As Boolean
    Select Case strType
        Case "MDB", "ACCDB"
            strDriver = "Microsoft Access Driver (*.mdb, *.accdb)"
            strAttributes = "DBQ=" & strDB & Chr$(0) & "Exclusive=0" & Chr$(0) & "ReadOnly=0" & Chr$(0)
        Case "DBF"
            strDriver = "Microsoft Visual FoxPro Driver"
            strAttributes = "SourceDB=" & strDB & Chr$(0) & "SourceType=DBF" & Chr$(0) & "Exclusive=0" & Chr$(0)
        Case "SQLITE"
            strDriver = "SQLite3 ODBC Driver"
            strAttributes = "Database=" & strDB & Chr$(0)
        Case "SQLSERVER", "MSSQL"
            strDriver = GetBestMSSQL("D")
            If Len(strDriver) = 0 Then strDriver = "SQL Server"
            strAttributes = "SERVER=" & sHost & Chr$(0) & "DATABASE=" & strDB & Chr$(0)
            If Len(sUser) > 0 Then strAttributes = strAttributes & "UID=" & sUser & Chr$(0)
            If Len(sPass) > 0 Then strAttributes = strAttributes & "PWD=" & sPass & Chr$(0)
        Case "MYSQL"
            strDriver = "MySQL ODBC 8.0 ANSI Driver"
            If Len(sPort) = 0 Then sPort = "3306"
            strAttributes = "SERVER=" & sHost & Chr$(0) & "DATABASE=" & strDB & Chr$(0) & _
                            "UID=" & sUser & Chr$(0) & "PWD=" & sPass & Chr$(0) & _
                            "PORT=" & sPort & Chr$(0) & "OPTION=3" & Chr$(0)
        Case "MARIADB"
            strDriver = "MariaDB ODBC 3.0 Driver"
            If Len(sPort) = 0 Then sPort = "3306"
            strAttributes = "SERVER=" & sHost & Chr$(0) & "DATABASE=" & strDB & Chr$(0) & _
                            "UID=" & sUser & Chr$(0) & "PWD=" & sPass & Chr$(0) & _
                            "PORT=" & sPort & Chr$(0)
        Case "POSTGRESQL", "PGSQL"
            strDriver = "PostgreSQL ANSI"
            If Len(sPort) = 0 Then sPort = "5432"
            strAttributes = "SERVER=" & sHost & Chr$(0) & "DATABASE=" & strDB & Chr$(0) & _
                            "UID=" & sUser & Chr$(0) & "PWD=" & sPass & Chr$(0) & _
                            "PORT=" & sPort & Chr$(0)
        Case "FIREBIRD"
            strDriver = FirebirdODBC()
            If Len(strDriver) = 0 Then Exit Function
            strAttributes = "DBNAME=" & strDB & Chr$(0)
            If Len(sUser) > 0 Then strAttributes = strAttributes & "UID=" & sUser & Chr$(0)
            If Len(sPass) > 0 Then strAttributes = strAttributes & "PWD=" & sPass & Chr$(0)
        Case "ORACLE"
            strDriver = "Oracle in OraClient19Home1"
            strAttributes = "DBQ=" & strDB & Chr$(0) & _
                            "UID=" & sUser & Chr$(0) & _
                            "PWD=" & sPass & Chr$(0)
        Case Else
            Exit Function
    End Select
    MontarAtributosDSN = (Len(strDriver) > 0)
End Function

Public Function DSNExists(ByVal dsnName As String) As Boolean
    DSNExists = DSNExistsEscopo(dsnName, True) Or DSNExistsEscopo(dsnName, False)
End Function

Private Function DSNExistsEscopo(ByVal dsnName As String, ByVal lUSER As Boolean) As Boolean
    Dim nRoot As Long
    Dim sBase As String
    Dim sValue As String
    Dim sWow As String

    If lUSER Then
        nRoot = HKEY_CURRENT_USER
        sBase = "Software\ODBC\ODBC.INI\"
        sWow = "Software\WOW6432Node\ODBC\ODBC.INI\"
    Else
        nRoot = HKEY_LOCAL_MACHINE
        sBase = "Software\ODBC\ODBC.INI\"
        sWow = "Software\WOW6432Node\ODBC\ODBC.INI\"
    End If
    
    DSNExistsEscopo = RegistryReadString(nRoot, sBase & dsnName, "Driver", sValue)
    If Not DSNExistsEscopo Then
        DSNExistsEscopo = RegistryReadString(nRoot, sWow & dsnName, "Driver", sValue)
    End If
End Function

Private Function RemoverDSN(ByVal dsnName As String, ByVal strDriverType As String, _
                            ByVal lUSER As Boolean) As Boolean
    Dim strDriver As String
    Dim strAttributes As String
    Dim intRet As Long

    strDriver = GetDsnDriverName(dsnName, lUSER)
    If Len(strDriver) = 0 Then
        If Not MontarAtributosDSN(strDriverType, "", "", "", "", "", strDriver, strAttributes) Then Exit Function
    End If

    If lUSER Then
        intRet = SQLConfigDataSource(NULL_PTR, ODBC_REMOVE_DSN, strDriver, "DSN=" & dsnName & Chr$(0) & Chr$(0))
    Else
        intRet = SQLConfigDataSource(NULL_PTR, ODBC_REMOVE_SYS_DSN, strDriver, "DSN=" & dsnName & Chr$(0) & Chr$(0))
    End If
    RemoverDSN = (intRet <> 0)
End Function

Private Function GetDsnDriverName(ByVal dsnName As String, ByVal lUSER As Boolean) As String
    Dim nRoot As Long
    Dim sDsnBase As String
    Dim sDsnDriver As String
    
    If lUSER Then
        nRoot = HKEY_CURRENT_USER
        sDsnBase = "Software\ODBC\ODBC.INI\"
    Else
        nRoot = HKEY_LOCAL_MACHINE
        sDsnBase = "Software\ODBC\ODBC.INI\"
    End If
    
    If Not RegistryReadString(nRoot, sDsnBase & dsnName, "Driver", sDsnDriver) Then
        If Not RegistryReadString(nRoot, "Software\WOW6432Node\ODBC\ODBC.INI\" & dsnName, "Driver", sDsnDriver) Then Exit Function
    End If

    GetDsnDriverName = MatchOdbcDriverPath(sDsnDriver, "SOFTWARE\ODBC\ODBCINST.INI")
    If Len(GetDsnDriverName) = 0 Then
        GetDsnDriverName = MatchOdbcDriverPath(sDsnDriver, "SOFTWARE\WOW6432Node\ODBC\ODBCINST.INI")
    End If
End Function

Private Function MatchOdbcDriverPath(ByVal sDsnDriver As String, ByVal sDriverList As String) As String
    #If VBA7 Or Win64 Then
        Dim hKey As LongPtr
    #Else
        Dim hKey As Long
    #End If
    Dim lIndex As Long
    Dim lNameSize As Long
    Dim sSubKeyName As String
    Dim sRegisteredPath As String
    
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, sDriverList, 0, KEY_READ, hKey) = 0 Then
        lIndex = 0
        lNameSize = 255
        sSubKeyName = Space$(lNameSize)
        
        Do While RegEnumKeyEx(hKey, lIndex, sSubKeyName, lNameSize, 0, vbNullString, 0, ByVal 0&) = 0
            sSubKeyName = Left$(sSubKeyName, lNameSize)
            
            If RegistryReadString(HKEY_LOCAL_MACHINE, sDriverList & "\" & sSubKeyName, "Driver", sRegisteredPath) Then
                If StrComp(Replace$(sRegisteredPath, "/", "\"), Replace$(sDsnDriver, "/", "\"), vbTextCompare) = 0 Then
                    MatchOdbcDriverPath = sSubKeyName
                    RegCloseKey hKey
                    Exit Function
                End If
            End If
            
            lIndex = lIndex + 1
            lNameSize = 255
            sSubKeyName = Space$(lNameSize)
        Loop
        RegCloseKey hKey
    End If
End Function

Public Function FirebirdODBC() As String
    If IsDriverInstalled("Firebird ODBC Driver") Then
        FirebirdODBC = "Firebird ODBC Driver"
    ElseIf IsDriverInstalled("Firebird/InterBase(r) driver") Then
        FirebirdODBC = "Firebird/InterBase(r) driver"
    End If
End Function

Public Function GetBestMSSQL(TIPO As String) As String
    Dim vCandidates As Variant
    Dim vItem As Variant
    Dim bODBC As Boolean

    bODBC = (UCase$(Trim$(TIPO)) = "D")
    If bODBC Then
        vCandidates = Array("ODBC Driver 17 for SQL Server", "ODBC Driver 13 for SQL Server", _
                            "SQL Server Native Client 11.0", "SQL Server")
    Else
        GetBestMSSQL = MSSqlOledbProvider(1)
        Exit Function
    End If

    For Each vItem In vCandidates
        If DriverExisteOdbc(CStr(vItem), bODBC) Then
            GetBestMSSQL = CStr(vItem)
            Exit Function
        End If
    Next vItem
End Function

Public Sub ShowODBCError()
    Dim sMsg As String * 1024
    Dim sDetails As String
    Dim lErr As Long
    Dim iLen As Integer
    Dim iError As Integer
    Dim nResult As Integer

    For iError = 1 To 8
        iLen = 0
        nResult = SQLInstallerError(iError, lErr, sMsg, Len(sMsg), iLen)
        If nResult <> 0 And nResult <> 1 Then Exit For
        If iLen > 0 Then
            If Len(sDetails) > 0 Then sDetails = sDetails & vbCrLf
            sDetails = sDetails & "Erro " & CStr(lErr) & ": " & Left$(sMsg, iLen)
        End If
    Next iError
    If Len(sDetails) = 0 Then sDetails = "A API ODBC nao retornou detalhes do erro."
    MsgBox sDetails, vbCritical, "Erro ODBC"
End Sub

Public Function IsDriverInstalled(ByVal sDriverName As String) As Boolean
    Dim sValue As String
    Dim sPath As String

    sPath = "SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers"
    IsDriverInstalled = RegistryReadString(HKEY_LOCAL_MACHINE, sPath, sDriverName, sValue)
    If Not IsDriverInstalled Then
        sPath = "SOFTWARE\WOW6432Node\ODBC\ODBCINST.INI\ODBC Drivers"
        IsDriverInstalled = RegistryReadString(HKEY_LOCAL_MACHINE, sPath, sDriverName, sValue)
    End If
    If IsDriverInstalled Then IsDriverInstalled = (StrComp(sValue, "Installed", vbTextCompare) = 0)
End Function

' --- API NATIVA DO REGISTRO: Substituindo dependência de WMI ---
Private Function RegistryReadString(ByVal nRoot As Long, ByVal sSubKey As String, _
                                    ByVal sValueName As String, ByRef sValue As String) As Boolean
    #If VBA7 Or Win64 Then
        Dim hKey As LongPtr
    #Else
        Dim hKey As Long
    #End If
    Dim lRet As Long
    Dim lType As Long
    Dim lDataSize As Long
    Dim sData As String

    lRet = RegOpenKeyEx(nRoot, sSubKey, 0, KEY_READ, hKey)
    If lRet = 0 Then
        lRet = RegQueryValueEx(hKey, sValueName, 0, lType, ByVal 0&, lDataSize)
        If lRet = 0 And (lType = REG_SZ) Then
            sData = Space$(lDataSize)
            lRet = RegQueryValueEx(hKey, sValueName, 0, 0, ByVal sData, lDataSize)
            If lRet = 0 Then
                If InStr(sData, Chr$(0)) > 0 Then
                    sValue = Left$(sData, InStr(sData, Chr$(0)) - 1)
                Else
                    sValue = Trim$(sData)
                End If
                RegistryReadString = True
            End If
        End If
        RegCloseKey hKey
    End If
End Function

Private Function RegistryKeyExists(ByVal nRoot As Long, ByVal sSubKey As String) As Boolean
    #If VBA7 Or Win64 Then
        Dim hKey As LongPtr
    #Else
        Dim hKey As Long
    #End If
    
    If RegOpenKeyEx(nRoot, sSubKey, 0, KEY_READ, hKey) = 0 Then
        RegistryKeyExists = True
        RegCloseKey hKey
    End If
End Function

Private Function OleDbProviderExists(ByVal sProvider As String) As Boolean
    Dim oProvider As Object
    On Error GoTo ProviderNotAvailable
    Set oProvider = CreateObject(sProvider)
    OleDbProviderExists = True
    Set oProvider = Nothing
    Exit Function
ProviderNotAvailable:
    Set oProvider = Nothing
    OleDbProviderExists = False
End Function

Public Function MSSqlOledbProvider(Optional ByVal nTIPO As Integer = 1) As String
    Dim vProviders As Variant
    Dim i As Long
    Dim sProgID As String
    Dim sDescription As String
    Dim sCLSID As String
    Dim sKey As String

    vProviders = Array( _
        "MSOLEDBSQL19", "Microsoft OLE DB Driver 19 for SQL Server", "EE5DE99A-4453-4C96-861C-F8832A7F59FE", _
        "MSOLEDBSQL", "Microsoft OLE DB Driver for SQL Server", "5A23DE84-1D7B-4A16-8DED-B29C09CB648D", _
        "SQLNCLI11", "SQL Server Native Client 11.0", "397C2819-8272-4532-AD3A-FB5E43BEAA39", _
        "SQLNCLI10", "SQL Server Native Client 10.0", "8F4A6B68-4F36-4E3C-BE81-BC7CA4E9C45C", _
        "SQLOLEDB", "Microsoft OLE DB Provider for SQL Server", "0C7FF16C-38E3-11D0-97AB-00C04FC2AD98")

    For i = 0 To UBound(vProviders) Step 3
        sProgID = CStr(vProviders(i))
        sDescription = CStr(vProviders(i + 1))
        sCLSID = CStr(vProviders(i + 2))
        sKey = "CLSID\{" & sCLSID & "}"
        If RegistryKeyExists(HKEY_CLASSES_ROOT, sKey) Or _
           RegistryKeyExists(HKEY_CLASSES_ROOT, "WOW6432Node\" & sKey) Then
            Select Case nTIPO
                Case 1: MSSqlOledbProvider = sProgID
                Case 2: MSSqlOledbProvider = sDescription
                Case 3: MSSqlOledbProvider = sCLSID
            End Select
            Exit Function
        End If
    Next i
    MsgBox "Nenhum provider OLE DB do SQL Server foi encontrado.", vbExclamation
End Function

Public Function MSSqlOdbcDriver() As String
    MSSqlOdbcDriver = GetBestMSSQL("D")
End Function