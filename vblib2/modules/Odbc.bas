Attribute VB_Name = "Odbc"
Public Declare Function SQLAllocEnv Lib "odbc32.dll" (phenv As Long) As Integer
Public Declare Function SQLFreeEnv Lib "odbc32.dll" (ByVal hEnv As Long) As Integer
Public Declare Function SQLDataSources Lib "odbc32.dll" _
                                       (ByVal hEnv As Long, ByVal fDirection As Integer, ByVal szDSN As String, _
                                        ByVal cbDSNMax As Integer, pcbDSN As Integer, _
                                        ByVal szDescription As String, ByVal cbDescriptionMax As Integer, _
                                        pcbDescription As Integer) As Integer
Public Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" _
                                            (ByVal hWndParent As Long, ByVal fRequest As Long, _
                                             ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
Public Declare Function SQLDrivers Lib "odbc32.dll" _
                                   (ByVal env As Long, ByVal Dir As Integer, ByVal descrip As String, ByVal bflen As Integer, descriplen As Integer, _
                                    ByVal attrib As String, ByVal bfattrlen As Integer, attriblen As Integer) As Long


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
  Case "MDB"
    StrDriver = "Microsoft Access Driver (*.mdb)"
    cData = "DBQ=" & strDB & Chr(0)
  Case "DBF"
    StrDriver = "Microsoft Visual FoxPro Driver"
    cData = "SourceDB=" & strDB & Chr(0)
  End Select

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


