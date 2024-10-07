Attribute VB_Name = "MdbRepComSimples"
Option Explicit

'Database type
'MS Access Version
'MSAccess2000 = *.mdb (version 2000 - 2003)
'MSAccess2007 = *.accdb (version 2007)
Public Enum enumDatabaseType
  [MSAccess2000]
  [MSAccess2007]
End Enum

'--------------------------------------------------------------------------------
' Project    :    prjCompactAccess
' Procedure  :    CompactAccess
' Description:    This procedure is use to compact & repair ms access database
'                 Compact & Repair function of MS Access will perform a routine
'                 maintenance on the database to ensure database integrity.
'
' Parameters :       xDatabaseLocation (String)
'                    xDatabasePassword (String)
'                    DatabaseType (enumDatabaseType = MSAccess2000)
'--------------------------------------------------------------------------------
Public Function CompactAccess(xDatabaseLocation As String, _
                              Optional xDatabasePassword As String, _
                              Optional DatabaseType As enumDatabaseType = MSAccess2000) As Boolean

'Microsoft Jet and Replecation 2.6
'<EhHeader>
  On Error GoTo CompactAccess_Err

  '</EhHeader>

10 Dim JRO As New JetEngine, SourceCon1 As String, SourceCon2 As String

12 Dim SourcePath As String, DesPath As String

14 SourcePath = xDatabaseLocation

16 If DatabaseType = MSAccess2000 Then
18  DesPath = App.Path & "\Compact.mdb"
20  SourceCon1 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & SourcePath & ";Jet OLEDB:Database Password=" & xDatabasePassword & ";"
22  SourceCon2 = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DesPath & ";Jet OLEDB:Database Password=" & xDatabasePassword & ";"
24 ElseIf DatabaseType = MSAccess2007 Then
26  DesPath = App.Path & "\Compact.accdb"
28  SourceCon1 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & SourcePath & ";Jet OLEDB:Database Password=" & xDatabasePassword & ";"
30  SourceCon2 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DesPath & ";Jet OLEDB:Database Password=" & xDatabasePassword & ";"
32 Else
34  MsgBox "Invalid Database or database is not yet supported with the code!", vbExclamation, "Not Supported"
36 End If

38 JRO.CompactDatabase SourceCon1, SourceCon2
40 DeleteFile SourcePath  '' kill SourcePath
42 CopyFileWindowsWay DesPath, SourcePath  'FileCopy DesPath, SourcePath
44 DeleteFile DesPath  '' Kill DesPath

46 Set JRO = Nothing
48 CompactAccess = True

  '<EhFooter>
  Exit Function

CompactAccess_Err:
  Call prompt_err(Err, "ModDatabase", "CompactAccess")
  CompactAccess = False

  '</EhFooter>

End Function

'Procedure used to promp unexpected errors
Public Sub prompt_err(ByVal sError As ErrObject, _
                      ByVal ModuleName As String, _
                      ByVal OccurIn As String)

'Dim cErr As clsErrorReport

  Dim s1, s2, s3, s4 As String
  Dim s5 As Long
''  Dim ff As Integer
  Dim strMsg As String

  s1 = ModuleName
  s2 = OccurIn
  s3 = sError.Number
  s4 = sError.Description

  '* Store our line # where error Occured (always 0 when line # is not use)
  s5 = Erl

  If s5 <> 0 Then
    strMsg = "An Unexpected Error has Occured in " & App.ProductName & " (Build " & App.Major & "." & App.Minor & "." & App.Revision & ")" & "." & vbCrLf & vbCrLf & "Please report this problem by contacting the author so it can be resolved." & vbCrLf & vbCrLf & "Technical Information:" & vbCrLf & "Error ID: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & "Module: " & ModuleName & vbCrLf & "Function: " & OccurIn & vbCrLf & vbCrLf & "Line Number: " & s5 & vbCrLf & vbCrLf & "Press Ctrl + C to copy this information to the clipboard"
  Else
    strMsg = "An Unexpected Error has Occured in " & App.ProductName & " (Build " & App.Major & "." & App.Minor & "." & App.Revision & ")" & "." & vbCrLf & vbCrLf & "Please report this problem by contacting the author so it can be resolved." & vbCrLf & vbCrLf & "Technical Information:" & vbCrLf & "Error ID: " & Err.Number & vbCrLf & "Description: " & Err.Description & vbCrLf & "Module: " & ModuleName & vbCrLf & "Function: " & OccurIn & vbCrLf & vbCrLf & "Press Ctrl + C to copy this information to the clipboard"
  End If

  MsgBox strMsg, vbExclamation, App.ProductName & ": Unexpected Error"

End Sub

