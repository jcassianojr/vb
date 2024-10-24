VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
  '**************************************
  'Windows API/Global Declarations for :cReadWriteEasyReg (Updated Again)
  '**************************************
  Option Explicit
  
  
  #If VBA7 Then
    Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As LongPtr
        bInheritHandle As Long
    End Type
#Else
    Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
    End Type
#End If
' RegCreateKeyEx creates the specified key. If the key
' already exists, the function opens it. The phkResult
' parameter receives the key handle.
#If VBA7 Then
    Private Declare PtrSafe Function RegCreateKeyEx _
                          Lib "advapi32.dll" Alias "RegCreateKeyExA" ( _
                              ByVal hKey As LongPtr, ByVal lpSubKey As String, _
                              ByVal Reserved As Long, ByVal lpClass As String, _
                              ByVal dwOptions As Long, ByVal samDesired As Long, _
                              lpSecurityAttributes As SECURITY_ATTRIBUTES, _
                              phkResult As LongPtr, lpdwDisposition As Long) As Long
#Else
    Private Declare Function RegCreateKeyEx _
                          Lib "advapi32.dll" Alias "RegCreateKeyExA" ( _
                              ByVal hKey As Long, ByVal lpSubKey As String, _
                              ByVal Reserved As Long, ByVal lpClass As String, _
                              ByVal dwOptions As Long, ByVal samDesired As Long, _
                              lpSecurityAttributes As SECURITY_ATTRIBUTES, _
                              phkResult As Long, lpdwDisposition As Long) As Long
#End If

'RegCloseKey releases a handle to the specified key.
'(Key handles should not be left open any longer than
'necessary.)
#If VBA7 Then
    Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" ( _
                                    ByVal hKey As LongPtr) As Long
#Else
    Private Declare Function RegCloseKey Lib "advapi32.dll" ( _
                                     ByVal hCurKey As Long) As Long
#End If
                                     
                                     
' RegQueryInfoKey retrieves information about the specified
'key, such as the number of subkeys and values, the length
'of the longest value and key name, and the size of the
'longest data component among the key's values.
#If VBA7 Then
    Private Declare PtrSafe Function RegQueryInfoKey _
                          Lib "advapi32.dll" Alias "RegQueryInfoKeyA" ( _
                             ByVal hKey As LongPtr, ByVal lpClass As String, _
                             lpcbClass As Long, ByVal lpReserved As LongPtr, _
                             lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, _
                             lpcbMaxClassLen As Long, lpcValues As Long, _
                             lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, _
                             lpcbSecurityDescriptor As Long, _
                             lpftLastWriteTime As Long) As Long
#Else
    Private Declare Function RegQueryInfoKey _
                          Lib "advapi32.dll" Alias "RegQueryInfoKeyA" ( _
                              ByVal hCurKey As Long, ByVal lpClass As String, _
                              lpcbClass As Long, ByVal lpReserved As Long, _
                              lpcSubKeys As Long, lpcbMaxSubKeyLen As Long, _
                              lpcbMaxClassLen As Long, lpcValues As Long, _
                              lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, _
                              lpcbSecurityDescriptor As Long, _
                              lpftLastWriteTime As Long) As Long
#End If

'RegEnumKeyEx enumerates subkeys of the specified open
'key. Retrieves the name (and its length) of each subkey.
#If VBA7 Then
    Private Declare PtrSafe Function RegEnumKeyEx Lib "advapi32.dll" _
                                     Alias "RegEnumKeyExA" (ByVal hKey As LongPtr, _
                                                            ByVal dwIndex As Long, ByVal lpName As String, _
                                                            lpcbName As Long, ByVal lpReserved As LongPtr, _
                                                            ByVal lpClass As String, lpcbClass As Long, _
                                                            lpftLastWriteTime As Long) As Long
#Else
    Private Declare Function RegEnumKeyEx Lib "advapi32.dll" _
                                      Alias "RegEnumKeyExA" (ByVal hCurKey As Long, _
                                                             ByVal dwIndex As Long, ByVal lpName As String, _
                                                             lpcbName As Long, ByVal lpReserved As Long, _
                                                             ByVal lpClass As String, lpcbClass As Long, _
                                                             lpftLastWriteTime As Long) As Long
#End If

'RegEnumValue enumerates the values for the specified open
'key. Retrieves the name (and its length) of each value,
'and the type, content and size of the data.
#If VBA7 Then
    Private Declare PtrSafe Function RegEnumValue Lib "advapi32.dll" _
                                      Alias "RegEnumValueA" (ByVal hKey As LongPtr, _
                                                             ByVal dwIndex As Long, ByVal lpValueName As String, _
                                                             lpcbValueName As Long, ByVal lpReserved As LongPtr, _
                                                             lpType As Long, lpData As Any, lpcbData As Long) As Long
#Else
    Private Declare Function RegEnumValue Lib "advapi32.dll" _
                                      Alias "RegEnumValueA" (ByVal hCurKey As Long, _
                                                             ByVal dwIndex As Long, ByVal lpValueName As String, _
                                                             lpcbValueName As Long, ByVal lpReserved As Long, _
                                                             lpType As Long, lpData As Any, lpcbData As Long) As Long
#End If

'RegQueryValueEx retrieves the type, content and data for
' a specified value name. Note that if you declare the
' lpData parameter as String, you must pass it By Value.
#If VBA7 Then
    Private Declare PtrSafe Function RegQueryValueEx _
                          Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
                              ByVal hKey As LongPtr, ByVal lpValueName As String, _
                              ByVal lpReserved As LongPtr, lpType As Long, _
                              lpData As Any, lpcbData As Long) As Long
#Else
    Private Declare Function RegQueryValueEx _
                          Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
                              ByVal hCurKey As Long, ByVal lpValueName As String, _
                              ByVal lpReserved As Long, lpType As Long, _
                              lpData As Any, lpcbData As Long) As Long
#End If

'RegSetValueEx sets the data and type of a specified
' value under a key.
' Note that if you declare the lpData parameter as String, you must pass it By Value.
#If VBA7 Then
    Private Declare PtrSafe Function RegSetValueEx Lib "advapi32.dll" _
                                       Alias "RegSetValueExA" (ByVal hKey As LongPtr, _
                                                               ByVal lpValueName As String, ByVal Reserved As Long, _
                                                               ByVal dwType As Long, lpData As Any, _
                                                               ByVal cbData As Long) As Long
#Else
    Private Declare Function RegSetValueEx Lib "advapi32.dll" _
                                       Alias "RegSetValueExA" (ByVal hCurKey As Long, _
                                                               ByVal lpValueName As String, ByVal Reserved As Long, _
                                                               ByVal dwType As Long, lpData As Any, _
                                                               ByVal cbData As Long) As Long
#End If

'RegDeleteValue removes a named value from specified key.
#If VBA7 Then
    Private Declare PtrSafe Function RegDeleteValue _
                          Lib "advapi32.dll" Alias "RegDeleteValueA" ( _
                              ByVal hKey As LongPtr, ByVal lpValueName As String) As Long
#Else
    Private Declare Function RegDeleteValue _
                          Lib "advapi32.dll" Alias "RegDeleteValueA" ( _
                              ByVal hCurKey As Long, ByVal lpValueName As String) As Long
#End If

'RegDeleteKey deletes a subkey. Under Win 95/98, also
'deletes all subkeys and values. Under Windows NT/2000,
'the subkey to be deleted must not have subkeys. The class
'attempts to use SHDeleteKey (see below) before using
'RegDeleteKey.
#If VBA7 Then
    Private Declare PtrSafe Function RegDeleteKey Lib "advapi32.dll" _
                                      Alias "RegDeleteKeyA" (ByVal hKey As LongPtr, _
                                                             ByVal lpSubKey As String) As Long
#Else
    Private Declare Function RegDeleteKey Lib "advapi32.dll" _
                                      Alias "RegDeleteKeyA" (ByVal hKey As Long, _
                                                             ByVal lpSubKey As String) As Long
#End If

'SHDeleteKey deletes a subkey and all its descendants.
'Under Windows NT 4.0, Internet Explorer 4.0 or later
'is required.
#If VBA7 Then ' Not Found in Win32API_PtrSafe.TXT
        Private Declare PtrSafe Function SHDeleteKey Lib "Shlwapi" _
                                     Alias "SHDeleteKeyA" (ByVal hKey As LongPtr, _
                                                           ByVal lpSubKey As String) As Long
#Else
    Private Declare Function SHDeleteKey Lib "Shlwapi" _
                                     Alias "SHDeleteKeyA" (ByVal hKey As Long, _
                                                           ByVal lpSubKey As String) As Long
#End If

#If VBA7 Then
    Private Declare PtrSafe Function LoadLibrary Lib "kernel32" _
                                     Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr
#Else
    Private Declare Function LoadLibrary Lib "kernel32" _
                                     Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
#End If

#If VBA7 Then
    Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
#Else
    Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
#End If

#If VBA7 Then
    Private Declare PtrSafe Function ExpandEnvStrings Lib "kernel32" _
                                          Alias "ExpandEnvironmentStringsA" ( _
                                          ByVal lpSrc As String, ByVal lpDst As String, _
                                          ByVal nSize As Long) As Long
#Else
    Private Declare Function ExpandEnvStrings Lib "kernel32" _
                                          Alias "ExpandEnvironmentStringsA" ( _
                                          ByVal lpSrc As String, ByVal lpDst As String, _
                                          ByVal nSize As Long) As Long
#End If

'faltou osversioninfo
'#If VBA7 Then
'    Private Declare PtrSafe Function GetVersionEx Lib "kernel32" _
'                                      Alias "GetVersionExA" ( _
'                                      lpVersionInformation As OSVERSIONINFO) As Long
'#Else
'    Private Declare Function GetVersionEx Lib "kernel32" _
'                                      Alias "GetVersionExA" ( _
'                                      lpVersionInformation As OSVERSIONINFO) As Long
'#End If
  
  
 #If VBA7 Then
    Private Declare prtsafe Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Longprt, ByVal lpSubKey As String, ByVal ulOptions As Longptr, ByVal samDesired As Longptr, phkResult As Longptr) As Long
 #Else
    Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
 #End If
  
  'Registry API's to use
''  Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
  'Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
  'Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, cbName As Long, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
''  Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
  'Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
  'Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpKeyName As String) As Long
''  Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
  'Private Declare Function ExpandEnvironmentStrings Lib "advapi32.dll" (lpSrc As String, lpDst As String, ByVal nSize As Long) As Long
''  Private Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
''  Private Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long



'Private Type SECURITY_ATTRIBUTES
'  nLength As Long
'  lpSecurityDescriptor As Variant
'  bInheritHandle As Long
'End Type

'Enum's for the OpenRegistry function


Public Enum HKeys
  HKEY_CLASSES_ROOT = &H80000000
  HKEY_CURRENT_USER = &H80000001
  HKEY_LOCAL_MACHINE = &H80000002
  HKEY_USERS = &H80000003
  HKEY_PERFORMANCE_DATA = &H80000004
  HKEY_CURRENT_CONFIG = &H80000005
  HKEY_DYN_DATA = &H80000006
End Enum
'Enum's for the DataTypes


Public Enum lDataType
  REG_NONE = 0
  REG_SZ = 1
  REG_EXPAND_SZ = 2
  REG_BINARY = 3
  REG_DWORD = 4
  REG_DWORD_LITTLE_ENDIAN = 4
  REG_DWORD_BIG_ENDIAN = 5
  REG_LINK = 6
  REG_MULTI_SZ = 7
  REG_RESOURCE_LIST = 8
  REG_FULL_RESOURCE_DESCRIPTOR = 9
  REG_RESOURCE_REQUIREMENTS_LIST = 10
End Enum
'Right's for the OpenRegistry
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_CREATE_LINK = &H20
Private Const SYNCHRONIZE = &H100000
Private Const KEY_READ = &H20009
Private Const KEY_WRITE = &H20006
Private Const KEY_READ_WRITE = ( _
        KEY_READ _
        Or _
        KEY_WRITE _
      )
Private Const KEY_ALL_ACCESS = ( _
      ( _
      STANDARD_RIGHTS_ALL Or _
        KEY_QUERY_VALUE Or _
        KEY_SET_VALUE Or _
        KEY_CREATE_SUB_KEY Or _
        KEY_ENUMERATE_SUB_KEYS Or _
        KEY_NOTIFY Or _
        KEY_CREATE_LINK _
      ) _
      And _
      ( _
      Not SYNCHRONIZE _
      ) _
    )
Private Const REG_OPTION_NON_VOLATILE = 0&
Private Const REG_OPTION_VOLATILE = &H1

'Local var's to keep track of things happening
Dim RootHKey As HKeys
Dim SubDir As String
#If VBA7 Then
  Dim hKey As LongPtr
#Else
  Dim hKey As Long
End If
Dim OpenRegOk As Boolean


Public Function CreateDirectory(ByVal sNewDirName As String) As Boolean
  Dim hNewKey As Long, lpdwDisposition As Long
  Dim lpSecurityAttributes As SECURITY_ATTRIBUTES
  Dim lReturn As Long

  If Not OpenRegOk Then Exit Function

  lReturn = RegCreateKeyEx(hKey, sNewDirName, 0&, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpSecurityAttributes, hNewKey, lpdwDisposition)


  If lReturn = 0 Then
    CreateDirectory = True
  Else
    CreateDirectory = False
  End If
End Function


Public Function CreateValue(ByVal sValueName As String, ByVal vWriteThis As Variant, ldValueDataType As lDataType, Optional Multi_SZ_AddtlStrings As Variant) As Boolean
  Dim lpData As String  'The pointer to the value written to the Registry key's value
  Dim cbData As Long  'The size of the data written to the Registry key's value, including termination characters If applicable
  Dim lReturn As Long  'The Error value returned by the Registry Function
  Dim Str As Variant

  If Not OpenRegOk Then Exit Function



  Select Case ldValueDataType
  Case REG_SZ, REG_EXPAND_SZ
    lpData = vWriteThis & Chr(0)
    cbData = Len(lpData)
    lReturn = RegSetValueEx(hKey, sValueName, 0&, ldValueDataType, lpData, cbData)


    If lReturn = 0 Then
      CreateValue = True
    Else
      CreateValue = False
    End If
  Case REG_MULTI_SZ
    lpData = vWriteThis & Chr(0)


    If Not IsMissing(Multi_SZ_AddtlStrings) Then


      If IsArray(Multi_SZ_AddtlStrings) Then


        For Each Str In Multi_SZ_AddtlStrings


          If Str <> "" And Str <> Chr(0) And Not IsNull(Str) Then
            lpData = lpData & Str & Chr(0)
          End If
        Next Str
      Else


        If Multi_SZ_AddtlStrings <> "" And Multi_SZ_AddtlStrings <> Chr(0) And Not IsNull(Multi_SZ_AddtlStrings) Then
          lpData = lpData & Multi_SZ_AddtlStrings & Chr(0)
        End If
      End If
    End If
    lpData = lpData & Chr(0)
    cbData = Len(lpData)
    lReturn = RegSetValueEx(hKey, sValueName, 0&, ldValueDataType, lpData, cbData)


    If lReturn = 0 Then
      CreateValue = True
    Else
      CreateValue = False
    End If
  Case REG_DWORD
    lpData = CLng(vWriteThis)
    cbData = 4
    lReturn = RegSetValueEx(hKey, sValueName, 0&, ldValueDataType, lpData, cbData)


    If lReturn = 0 Then
      CreateValue = True
    Else
      CreateValue = False
    End If
  Case Else
    MsgBox "Unable to process that Type of data."
    CreateValue = False
  End Select
End Function


Function GetValue(ByVal VarName As String, Optional ReturnBinStr As Boolean = False) As Variant
  On Error GoTo handelgetavalue
  Dim i As Integer
  Dim SubKey_Value As String, TempStr As String, ReturnArray() As Variant
  Dim Length As Long
  'Dim value_type As Long
  Dim RtnVal As Long, value_Type As lDataType

  If Not OpenRegOk Then Exit Function

  'Read the size of the value value
  RtnVal = RegQueryValueEx(hKey, VarName, 0&, value_Type, ByVal 0&, Length)


  Select Case RtnVal
  Case 0  'Ok so continue
  Case 2  'Not Found
    Exit Function
  Case 5  'Access Denied
    GetValue = "Access Denied"
    Exit Function
  Case Else  'What?
    GetValue = "RegQueryValueEx Returned : (" & RtnVal & ")"
    Exit Function
  End Select
  'declare the size of the value and read
  '     it


  SubKey_Value = Space$(Length)
  RtnVal = RegQueryValueEx(hKey, VarName, 0&, value_Type, ByVal SubKey_Value, Length)


  Select Case value_Type
  Case REG_NONE
    'Not defined


    SubKey_Value = "Not defined value_type=REG_NONE"
  Case REG_SZ  'A null-terminated String


    SubKey_Value = Left$(SubKey_Value, Length - 1)
  Case REG_EXPAND_SZ
    'A null-terminated string that contains
    '     unexpanded references to
    'environment variables (for example, "%P
    '     ATH%").
    'Use ExpandEnvironmentStrings to expand


    SubKey_Value = Left$(SubKey_Value, Length - 1)
  Case REG_BINARY  'Binary data in any form.


    SubKey_Value = Left$(SubKey_Value, Length)


    If Not ReturnBinStr Then
      TempStr = ""


      For i = 1 To Len(SubKey_Value)
        TempStr = TempStr & Right$("00" & Trim$(Hex(Asc(Mid$(SubKey_Value, i, 1)))), 2) & " "
      Next i


      SubKey_Value = TempStr
    End If
  Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN  'A 32-bit number.


    SubKey_Value = Left$(SubKey_Value, Length)


    If Not ReturnBinStr Then
      TempStr = ""


      For i = 1 To Len(SubKey_Value)
        TempStr = TempStr & Right$("00" & Trim$(Hex(Asc(Mid$(SubKey_Value, i, 1)))), 2) & " "
      Next i


      SubKey_Value = TempStr
    End If
  Case REG_DWORD_BIG_ENDIAN
    'A 32-bit number in big-endian format.
    'In big-endian format, a multi-byte valu
    '     e is stored in memory from
    'the highest byte (the "big end") to the
    '     lowest byte. For example,
    'the value 0x12345678 is stored as (0x12
    '     0x34 0x56 0x78) in big-endian
    'format.
  Case REG_LINK
    'A Unicode symbolic link. Used internall
    '     y; applications should not
    'use this type.


    SubKey_Value = "Not defined value_type=REG_LINK"
  Case REG_MULTI_SZ
    'Array of null-terminated string


    SubKey_Value = Left$(SubKey_Value, Length)
    i = 0


    While Len(SubKey_Value) > 0
      ReDim Preserve ReturnArray(i) As Variant
      ReturnArray(i) = Mid$(SubKey_Value, 1, InStr(1, SubKey_Value, Chr(0)) - 1)


      SubKey_Value = Mid$(SubKey_Value, InStr(1, SubKey_Value, Chr(0)) + 1)
      i = i + 1
    Wend
    GetValue = ReturnArray
    Exit Function
  Case REG_RESOURCE_LIST
    'Device driver resource list.


    SubKey_Value = "Not defined value_type=REG_RESOURCE_LIST"
  Case REG_FULL_RESOURCE_DESCRIPTOR
    'Device driver resource list.


    SubKey_Value = "Not defined value_type=REG_FULL_RESOURCE_DESCRIPTOR"
  Case REG_RESOURCE_REQUIREMENTS_LIST
    'Device driver resource list.


    SubKey_Value = "Not defined value_type=REG_RESOURCE_REQUIREMENTS_LIST"
  Case Else


    SubKey_Value = "value_type=" & value_Type
  End Select
  GetValue = SubKey_Value
  Exit Function
handelgetavalue:
  GetValue = ""
  Exit Function
End Function
'This property returns the current KeyVa
'     lue


Public Property Get RegistryRootKey() As HKeys
  RegistryRootKey = RootHKey
End Property
'This property returns the current 'Regi
'     stry Directory' your in


Public Property Get SubDirectory() As String


  SubDirectory = SubDir
End Property
'This function open's the registry at a
'     specific 'Registry Directory'
'eg.
' Dim MyVariant As Variant, MyReg As New
'     CReadWriteEasyReg, i As Integer
' If Not MyReg.OpenRegistry(HKEY_LOCAL_M
'     ACHINE, "") Then
' MsgBox "Couldn't open the registry"
' Exit Sub
' End If
' MyVariant = MyReg.GetAllSubDirectories
'
' For i = LBound(MyVariant) To UBound(My
'     Variant)
' Debug.Print MyVariant(i)
' Next i
' MyReg.CloseRegistry


Public Function OpenRegistry(ByVal RtHKey As HKeys, ByVal SbDr As String) As Integer
  On Error GoTo OpenReg
  Dim ReturnVal As Integer


  If RtHKey = 0 Then
    OpenRegistry = False
    OpenRegOk = False
    Exit Function
  End If
  RootHKey = RtHKey


  SubDir = SbDr


  If OpenRegOk Then
    CloseRegistry
    OpenRegOk = False
  End If
  ReturnVal = RegOpenKeyEx(RootHKey, SubDir, 0&, KEY_READ_WRITE, hKey)


  If ReturnVal <> 0 Then
    OpenRegistry = False
    Exit Function
  End If
  OpenRegOk = True
  OpenRegistry = True
  Exit Function
OpenReg:
  OpenRegOk = False
  OpenRegistry = False
  Exit Function
End Function


Public Function OneBackOnKey()


  SubDir = Mid$(SubDir, 1, FindLastBackSlash(SubDir) - 1)
  CloseRegistry
  OpenRegistry RootHKey, SubDir
End Function
'This function should be called after yo
'     u're done with the registry
'eg. (see other examples)


Public Function CloseRegistry() As Boolean
  On Error Resume Next


  If RegCloseKey(hKey) <> 0 Then
    CloseRegistry = False
    Exit Function
  End If
  CloseRegistry = True
  OpenRegOk = False
End Function


Private Sub Class_Initialize()
  RootHKey = &H0


  SubDir = ""
  hKey = 0
  OpenRegOk = False
End Sub


Private Sub Class_Terminate()
  On Error Resume Next


  If RegCloseKey(hKey) <> 0 Then
    Exit Sub
  End If
End Sub



Private Function FindLastBackSlash(VarValue As Variant) As Integer
  Dim i As Integer, iRtn As Integer
  iRtn = 0


  For i = Len(VarValue) To 1 Step -1


    If Mid$(VarValue, i, 1) = "\" Then
      iRtn = i
      Exit For
    End If
  Next i
  FindLastBackSlash = iRtn
End Function




