Attribute VB_Name = "Odbc"
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

#If VBA7 Or Win64 Then
    ' --- VERSÃO 64-BIT / TWINBASIC / VBA7 ---
    Public Declare PtrSafe Function SQLAllocEnv Lib "odbc32.dll" (ByRef phenv As LongPtr) As Integer
    Public Declare PtrSafe Function SQLFreeEnv Lib "odbc32.dll" (ByVal hEnv As LongPtr) As Integer
    Public Declare PtrSafe Function SQLDataSources Lib "odbc32.dll" (ByVal hEnv As LongPtr, ByVal fDirection As Integer, ByVal szDSN As String, ByVal cbDSNMax As Integer, ByRef pcbDSN As Integer, ByVal szDescription As String, ByVal cbDescriptionMax As Integer, ByRef pcbDescription As Integer) As Integer
    Public Declare PtrSafe Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hWndParent As LongPtr, ByVal fRequest As LongPtr, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
    Public Declare PtrSafe Function SQLDrivers Lib "odbc32.dll" (ByVal hEnv As LongPtr, ByVal fDirection As Integer, ByVal szDriver As String, ByVal cbDriverMax As Integer, ByRef pcbDriver As Integer, ByVal szAttributes As String, ByVal cbAttrMax As Integer, ByRef pcbAttr As Integer) As Integer
#Else
    ' --- VERSÃO 32-BIT CLÁSSICA (VB6) ---
    Public Declare Function SQLAllocEnv Lib "odbc32.dll" (ByRef phenv As Long) As Integer
    Public Declare Function SQLFreeEnv Lib "odbc32.dll" (ByVal hEnv As Long) As Integer
    Public Declare Function SQLDataSources Lib "odbc32.dll" (ByVal hEnv As Long, ByVal fDirection As Integer, ByVal szDSN As String, ByVal cbDSNMax As Integer, ByRef pcbDSN As Integer, ByVal szDescription As String, ByVal cbDescriptionMax As Integer, ByRef pcbDescription As Integer) As Integer
    Public Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hWndParent As Long, ByVal fRequest As Long, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
    Public Declare Function SQLDrivers Lib "odbc32.dll" (ByVal hEnv As Long, ByVal fDirection As Integer, ByVal szDriver As String, ByVal cbDriverMax As Integer, ByRef pcbDriver As Integer, ByVal szAttributes As String, ByVal cbAttrMax As Integer, ByRef pcbAttr As Integer) As Integer
#End If
'Contants for adding DSN's
Private Const ODBC_ADD_DSN = 1
Private Const ODBC_ADD_SYS_DSN = 4

Public Function AddDSN(ByVal strDSN As String, ByVal strDescription As String, _
                       ByVal strDB As String, Optional ByVal StrDriver As String = "MDB", _
                       Optional ByVal lUSER As Boolean = False) As Boolean

  Dim strAttributes As String
  Dim cData As String
  '------------------------------------
  'Usage:
  ' AddDSN "MyDSN", "This is a test", "C:\test\myDB.mdb"
  '------------------------------------

  On Error GoTo Hell

  cData = strAttributes & "DATABASE=" & strDB & Chr(0)
  'Set the Driver Name
  Select Case StrDriver
  Case "MDB", "ACCDB"
    StrDriver = "Microsoft Access Driver (*.mdb, *.accdb)" '"Microsoft Access Driver (*.mdb)"
    cData = "DBQ=" & strDB & Chr(0)
  Case "DBF"
    StrDriver = "Microsoft Visual FoxPro Driver"
    cData = "SourceDB=" & strDB & Chr(0)
  Case "SQLITE"
    StrDriver = "SQLite3 ODBC Driver"
    cData = "Database=" & strDB & Chr(0)
  End Select
  
  'Select Case sTipo
  '  Case "DBF"
  '      StrDriver = "Microsoft Visual FoxPro Driver"
  '      strAttributes = "DSN=" & sDsnName & Chr(0) & "SourceType=DBF" & Chr(0) & _
  '                      "SourceDB=" & sPath & Chr(0) & "Exclusive=0" & Chr(0)

  '  Case "SQLITE"
  '      StrDriver = "SQLite3 ODBC Driver"
  '      strAttributes = "DSN=" & sDsnName & Chr(0) & "Database=" & sPath & Chr(0)

   ' Case "ACCDB"
        ' Novo Case para Access
   '     StrDriver = "Microsoft Access Driver (*.mdb, *.accdb)"
        ' No Access, DBQ deve ser o caminho completo do arquivo .accdb
    '    strAttributes = "DSN=" & sDsnName & Chr(0) & "DBQ=" & sPath & Chr(0) & _
                        "Exclusive=0" & Chr(0) & "ReadOnly=0" & Chr(0)
                        
    'Case Else
        ' Tratamento de erro ou driver padrão
'End Select

  'Build the attributes - Attributes must be Null separated

  strAttributes = strAttributes & "DESCRIPTION=" & strDescription & Chr(0)
  strAttributes = strAttributes & "DSN=" & strDSN & Chr(0)
  strAttributes = strAttributes & cData


  'Create DSN
  If lUSER Then
    AddDSN = SQLConfigDataSource(0&, ODBC_ADD_DSN, StrDriver, strAttributes)
  Else
    AddDSN = SQLConfigDataSource(0&, ODBC_ADD_SYS_DSN, StrDriver, strAttributes)
  End If

  Exit Function

Hell:

End Function

' Dentro da sua função AddDSN no Odbc.bas

'Select Case UCase(StrDriver)
    ' ... cases existentes (MDB, DBF, SQLITE)

 '   Case "MYSQL"
 '       ' Mantenha o padrão ANSI que você usa no AdoLib
 '       StrDriver = "MySQL ODBC 8.0 ANSI Driver"
        ' Atributos para o SQLConfigDataSource (separados por Chr(0))
'        strAttributes = "DSN=" & strDSN & Chr(0) & _
'                        "DESCRIPTION=" & strDescription & Chr(0) & _
'                        "SERVER=" & sHost & Chr(0) & _
'                        "DATABASE=" & strDB & Chr(0) & _
'                        "USER=" & sUser & Chr(0) & _
'                        "PASSWORD=" & sPass & Chr(0) & _
'                        "PORT=3306" & Chr(0) & _
'                        "OPTION=3" & Chr(0) ' OPTION=3 habilita flags de compatibilidade comuns

 '   Case "MARIADB"
 '       StrDriver = "MariaDB ODBC 3.0 Driver" ' Ou a versão que você usa
 '       strAttributes = "DSN=" & strDSN & Chr(0) & _
                        "DESCRIPTION=" & strDescription & Chr(0) & _
                        "SERVER=" & sHost & Chr(0) & _
                        "DATABASE=" & strDB & Chr(0) & _
                        "USER=" & sUser & Chr(0) & _
                        "PASSWORD=" & sPass & Chr(0) & _
                        "PORT=3306" & Chr(0)

'    Case "POSTGRESQL", "PGSQL"
        ' Mantenha o padrão ANSI que você usa no AdoLib
'        StrDriver = "PostgreSQL ANSI"
'        strAttributes = "DSN=" & strDSN & Chr(0) & _
'                        "DESCRIPTION=" & strDescription & Chr(0) & _
                        "SERVER=" & sHost & Chr(0) & _
                        "DATABASE=" & strDB & Chr(0) & _
                        "UID=" & sUser & Chr(0) & _
                        "PWD=" & sPass & Chr(0) & _
                        "PORT=5432" & Chr(0)

'    Case "ORACLE"
        ' Oracle usa TNSNAMES.ORA, então strDB deve ser o Alias do TNS
'        StrDriver = "Oracle in OraClient19Home1" ' Nome varia conforme instalação
'        strAttributes = "DSN=" & strDSN & Chr(0) & _
'                        "DESCRIPTION=" & strDescription & Chr(0) & _
'                        "DBQ=" & strDB & Chr(0) & _
'                        "UID=" & sUser & Chr(0) & _
'                        "PWD=" & sPass & Chr(0)
'End Select


' Função única para verificar qualquer driver ODBC pelo nome
'IsDriverInstalled("SQLite3 ODBC Driver")
'IsDriverInstalled("Microsoft Access Driver (*.mdb, *.accdb)")
'IsDriverInstalled("Microsoft Visual FoxPro Driver")
Public Function IsDriverInstalled(ByVal sDriverName As String) As Boolean
    Dim shell As Object
    Dim sRegKey As String
    
    On Error Resume Next
    
    Set shell = CreateObject("WScript.Shell")
    
    ' 1. Tenta buscar na chave padrão (32-bit ou 64-bit dependendo do contexto)
    sRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers\" & sDriverName
    shell.RegRead sRegKey
    
    If Err.Number = 0 Then
        IsDriverInstalled = True
    Else
        ' 2. Fallback: Se não achou, busca no Wow6432Node (específico para apps 32-bit em Windows 64-bit)
        Err.Clear
        sRegKey = "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\ODBC\ODBCINST.INI\ODBC Drivers\" & sDriverName
        shell.RegRead sRegKey
        
        If Err.Number = 0 Then
            IsDriverInstalled = True
        Else
            IsDriverInstalled = False
        End If
    End If
    
    Set shell = Nothing
End Function
