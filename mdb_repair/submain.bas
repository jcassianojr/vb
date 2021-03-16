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
        
'Public Sub Buy(frm As Form, i As Integer)
'Dim lngRet As Long
'Dim fileName As String
'Dim Dummy As String
'Dim BrowserExec As String * 255
'Dim FileNumber As Integer
'Dim strURL As String
'On Error Resume Next
'If i = 1 Then
'   strURL = "https://www.regnow.com/softsell/nph-softsell.cgi?item=4459-57"
'Else
'   strURL = "https://www.regnow.com/softsell/nph-softsell.cgi?item=4459-58"
'End If
'Create a known, temporary HTML file
'BrowserExec = Space(255)
'fileName = "C:\temphtm.HTM"
'FileNumber = FreeFile                    'Get unused file number
'Open fileName For Output As #FileNumber  'Create temp HTML file
'Write #FileNumber, "<HTML> <\HTML>"      'Output text
'Close #FileNumber                        'Close file
'Then find the application associated with it
'lngRet = FindExecutable(fileName, Dummy, BrowserExec)
'BrowserExec = Trim(BrowserExec)
'If an application is found, launch it!
'If lngRet <= 32 Or IsEmpty(BrowserExec) Then 'Error
'   MsgBox "Unable to find associated browser.", vbExclamation, App.Title
'Else
'   lngRet = ShellExecute(frm.hwnd, "open", BrowserExec, strURL, Dummy, 1)
'   If lngRet <= 32 Then        'Error
'      MsgBox "Unable to open web page.", vbExclamation, App.Title
'   End If
'End If
'Kill fileName                   'delete temp HTML file
'End Sub

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

    Dim arr() As String
   
    arr() = Split(sTxt, Chr$(0))
    If UBound(arr) >= 0 Then
        TrimNull = arr(0)
    Else
        TrimNull = sTxt
    End If
  
End Function

