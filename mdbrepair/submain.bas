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

