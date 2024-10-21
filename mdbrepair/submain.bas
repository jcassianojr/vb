Attribute VB_Name = "submain"
Option Explicit

Public sDBpassword As String
Public lretu As Boolean


Type RelationType
  oRels As Object
  sRelation As String
  sTable As String
  sFTable As String
  lAttr As Long
  sField() As String
  sFField() As String
End Type
Public RelData() As RelationType

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

#If VBA7 Then
'for loading my web page
Private Declare  PtrSafe Function ShellExecute Lib "shell32.dll" Alias _
                                      "ShellExecuteA" (ByVal hWnd As LongPtr, ByVal lpOperation _
                                                                           As String, ByVal lpFile As String, ByVal lpParameters _
                                                                                                              As String, ByVal lpDirectory As String, ByVal nShowCmd _
                                                                                                                                                      As LongPtr) As Long

'for finding default browser
Private Declare  PtrSafe Function FindExecutable Lib "shell32.dll" Alias _
                                        "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As _
                                                                                   String, ByVal lpResult As String) As Long

#Else
'for loading my web page
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
                                      "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation _
                                                                           As String, ByVal lpFile As String, ByVal lpParameters _
                                                                                                              As String, ByVal lpDirectory As String, ByVal nShowCmd _
                                                                                                                                                      As Long) As Long

'for finding default browser
Private Declare Function FindExecutable Lib "shell32.dll" Alias _
                                        "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As _
                                                                                   String, ByVal lpResult As String) As Long
#End If

Sub Main()

  DAO.DBEngine.SetOption dbExclusiveAsyncDelay, 2000
  DAO.DBEngine.SetOption dbSharedAsyncDelay, 0
  DAO.DBEngine.SetOption dbFlushTransactionTimeout, 500
  '
  DAO.DBEngine.SetOption dbUserCommitSync, "yes"
  DAO.DBEngine.SetOption dbImplicitCommitSync, "yes"
  DAO.DBEngine.SetOption dbLockRetry, 20
  DAO.DBEngine.SetOption dbPageTimeout, 5000
  DAO.DBEngine.SetOption dbMaxLocksPerFile, 9500
  DAO.DBEngine.SetOption dbLockDelay, 100
  DAO.DBEngine.SetOption dbRecycleLVs, 0



  Form1.Show
  EnabledVisualStyles

End Sub

Public Function TrimNull(ByVal sTxt As String) As String

  Dim Arr() As String

  Arr() = Split(sTxt, Chr$(0))
  If UBound(Arr) >= 0 Then
    TrimNull = Arr(0)
  Else
    TrimNull = sTxt
  End If

End Function

Public Function FixStr(ByVal eVAR As Variant, _
                       Optional ByVal ePAD As Variant = "", _
                       Optional ByVal coper As String = "", _
                       Optional ByVal nLEN As Integer = 0) As Variant
  On Error GoTo errhandler
  If IsNull(eVAR) Then
    If ePAD <> "" Then
      eVAR = ePAD
    End If
  End If
  FixStr = CStr(eVAR)
  If UCase(eVAR) = "NULL" Then
    FixStr = ePAD
  End If
  If InStr(coper, "TRIM") > 0 Then
    eVAR = Trim(eVAR)
    FixStr = eVAR
  End If
  If nLEN > 0 And Len(eVAR) > nLEN Then
    eVAR = Mid(eVAR, 1, nLEN)
    FixStr = eVAR
  End If
  Exit Function
errhandler:
  FixStr = ""
  Resume Next
End Function

Public Function FileExists(ByVal PathName As String) As Boolean

On Error Resume Next
Dim Attributes As VbFileAttribute, ErrVal As Long
Attributes = GetAttr(PathName)
ErrVal = Err.Number
On Error GoTo 0
If (Attributes And (vbDirectory Or vbVolume)) = 0 And ErrVal = 0 Then FileExists = True
End Function
