VERSION 5.00
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "vbccr18.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DBFix 1.0"
   ClientHeight    =   4704
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   10356
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.4
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4704
   ScaleWidth      =   10356
   StartUpPosition =   2  'CenterScreen
   Begin VBCCR18.CommandButtonW Command2 
      Height          =   372
      Left            =   6000
      TabIndex        =   6
      Top             =   960
      Width           =   3252
      _ExtentX        =   5736
      _ExtentY        =   656
      Caption         =   "&Fix Corrupt Database"
   End
   Begin VBCCR18.CommandButtonW Command1 
      Height          =   492
      Left            =   8520
      TabIndex        =   5
      Top             =   360
      Width           =   852
      _ExtentX        =   1503
      _ExtentY        =   868
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Alignment       =   0
      Picture         =   "Form1.frx":058A
      PictureAndCaption=   -1  'True
      Style           =   1
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   960
      Width           =   5652
   End
   Begin VB.CheckBox chkNoMods 
      Caption         =   "Do not modify database"
      Height          =   210
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   432
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   8100
   End
   Begin VBCCR18.ProgressBar PB1 
      Height          =   495
      Left            =   240
      Top             =   3120
      Width           =   8895
      _ExtentX        =   15685
      _ExtentY        =   868
      Step            =   10
      Text            =   "Form1.frx":0B24
   End
   Begin vbExtra.CommonDialogEx CDL1 
      Left            =   8640
      Top             =   2400
      _ExtentX        =   720
      _ExtentY        =   720
      MaxFileSize     =   255
      FontName        =   ""
   End
   Begin VBCCR18.CommandButtonW cmdCreate 
      Height          =   372
      Left            =   6000
      TabIndex        =   7
      Top             =   1440
      Width           =   3252
      _ExtentX        =   5736
      _ExtentY        =   656
      Caption         =   "&Create Master Schema Database"
   End
   Begin VBCCR18.CommandButtonW cmdCompare 
      Height          =   372
      Left            =   6000
      TabIndex        =   8
      Top             =   1920
      Width           =   3252
      _ExtentX        =   5736
      _ExtentY        =   656
      Caption         =   "Compare &Database with Schema"
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3720
      Width           =   8895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Database"
      Height          =   210
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   750
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'TODO
'1. Store relationships, forms, reports, macros, modules in schema database so it can be recreated

Dim dbTemp As Database
Dim dbx As Database
Dim bFix As Boolean
Dim sDir As String
Dim sAppPath As String

Dim oRel As Relation
Dim fld As Field
Dim flds As Fields
Dim bRel As Boolean

'Parameter variables
Dim sSchemaDB As String
Dim iMode As Integer
#If VBA7 Then
''for getting drive type
Private Declare  PtrSafe Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Declare  PtrSafe Function ShellExecute Lib "shell32.dll" Alias _
                                      "ShellExecuteA" (ByVal hWnd As LongPtr, ByVal lpOperation _
                                                                           As String, ByVal lpFile As String, ByVal lpParameters _
                                                                                                              As String, ByVal lpDirectory As String, ByVal nShowCmd _
                                                                                                                                                      As LongPtr) As Long
'Private Declare  PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
'Private Declare  PtrSafe Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
 Private Declare PtrSafe Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" _
        (ByVal lpDirectoryName As String, ByRef lpFreeBytesAvailableToMe As LongPtr, _
        ByRef lpTotalNumberOfBytes As LongPtr, ByRef lpTotalNumberOfFreeBytes As LongPtr) As Long


Private Declare  PtrSafe Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As LongPtr, lpBytesPerSector As LongPtr, lpNumberOfFreeClusters As LongPtr, lpTotalNumberOfClusters As LongPtr) As Long

#Else
'for getting drive type
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias _
                                      "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation _
                                                                           As String, ByVal lpFile As String, ByVal lpParameters _
                                                                                                              As String, ByVal lpDirectory As String, ByVal nShowCmd _
                                                                                                                                                      As Long) As Long
'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

 Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" _
        (ByVal lpDirectoryName As String, ByRef lpFreeBytesAvailableToMe As Long, _
        ByRef lpTotalNumberOfBytes As Long, ByRef lpTotalNumberOfFreeBytes As Long) As Long


'trocado currency por long acima
'Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
#End If

Private Sub cmdCreate_Click()

  If Text1.Text = "" Then
    MsgBox "Enter database name.", vbExclamation
    Exit Sub
  ElseIf FileExists(Text1.Text) = False Then
    If iMode = 1 Then
      WriteErrLog "Unable to create master schema file. " & Text1.Text & " does not exist"
      End
    Else
      lblStatus.Caption = "The database does not exist"
      Text1.Text = ""
      Exit Sub
    End If
  End If

  Dim S As String
  Dim sPath As String

  On Error Resume Next

  If iMode = 0 Then

    bFix = False

    S = Text1.Text
    S = Replace(S, ".mdb", "_Schema.mdb", 1, 1, vbTextCompare)

    sPath = Text1.Text
    sPath = Left(sPath, InStrRev(sPath, "\") - 1)

    'Show Save Dialog Screen
    CDL1.DefaultExt = "mdb"
    CDL1.InitDir = sPath
    CDL1.Filter = "Microsoft Access (*.mdb)|*.mdb|All Files (*.*)|*.*"
    CDL1.filename = S
    CDL1.CancelError = True
    CDL1.ShowSave
    If Err.Number Then  'Cancel selected
      Exit Sub
    End If




    sSchemaDB = CDL1.filename

    If LCase(Right(sSchemaDB, 11)) <> "_schema.mdb" Then
      MsgBox "The master schema database name must end with '_schema.mdb'.", vbExclamation
      Exit Sub
    End If

    If StrComp(Text1.Text, sSchemaDB, vbTextCompare) = 0 Then
      MsgBox "The master schema database cannot be the same name as the selected database.", vbExclamation
      Exit Sub
    End If

  ElseIf StrComp(Text1.Text, sSchemaDB, vbTextCompare) = 0 Then
    If iMode = 1 Then
      WriteErrLog "The master schema database cannot be the same name as the selected database."
      End
    Else
      lblStatus.Caption = "The master schema database cannot be the same name as the selected database."
      Exit Sub
    End If

  End If

  Dim rsloTb As Recordset
  Dim rsloQy As Recordset
  Dim rsloIx As Recordset

  'If MsgBox("You are going to override Master-Schema, are you sure ?", vbYesNo) = vbNo Then Exit Sub
  Text2.Text = ""

  CreateLocalTable sSchemaDB

  If Not OpenAllDB(sSchemaDB, False, False) Then
    lblStatus.Caption = "Unable to open databases"
    Text1.Text = ""
    Exit Sub
  End If

  Screen.MousePointer = vbHourglass

  DisableControls

  Set rsloTb = dbTemp.OpenRecordset("select * from myTABLE")
  Set rsloQy = dbTemp.OpenRecordset("select * from myQUERY")
  Set rsloIx = dbTemp.OpenRecordset("select * from myINDEX")

  On Error GoTo ErrProc

  putText "Creating TABLE and INDEX definitions..."
  Dim i%, J%, K%

  For i = 0 To dbx.TableDefs.Count - 1

    If dbx.TableDefs(i).Updatable Then  'Ignore linked tables

      'If UCase(Mid(dbx.TableDefs(i).Name, 1, 2)) = "TB" Or UCase(Mid(dbx.TableDefs(i).Name, 1, 2)) = "LO" Then
      If LCase(Left(dbx.TableDefs(i).Name, 4)) <> "msys" Then
        With dbx.TableDefs(i)
          For J = 0 To .Fields.Count - 1

            rsloTb.AddNew
            rsloTb("TableName") = .Name
            rsloTb("SeqNum") = .Fields(J).OrdinalPosition
            rsloTb("FieldName") = .Fields(J).Name
            rsloTb("FieldType") = .Fields(J).Type
            rsloTb("Attributes") = .Fields(J).Attributes
            rsloTb("Required") = .Fields(J).Required
            rsloTb("Size") = .Fields(J).Size
            rsloTb("AllowZeroLength") = .Fields(J).AllowZeroLength

            If .Fields(J).Type = 1 Or .Fields(J).Type = 4 Or .Fields(J).Type = 5 Then
              ' yes/no, number, currency
              rsloTb("DefaultValue") = 0
            ElseIf .Fields(J).Type = 10 Then
              ' Text
              rsloTb("DefaultValue") = .Fields(J).DefaultValue
            End If
            rsloTb.Update
          Next  'j
          For J = 0 To dbx.TableDefs(i).Indexes.Count - 1
            rsloIx.AddNew
            rsloIx("TableName") = .Name
            rsloIx("IndexName") = .Indexes(J).Name
            rsloIx("Fields") = .Indexes(J).Fields
            rsloIx("Primary") = .Indexes(J).Primary
            rsloIx("Unique") = .Indexes(J).Unique
            rsloIx.Update
Jump:
          Next  'j
        End With
      End If

    End If

  Next  'i

  putText "Finished."

  putText "Creating QUERY definitions..."
  For i = 0 To dbx.QueryDefs.Count - 1

    If dbx.QueryDefs(i).Updatable Then  'Ignore linked tables

      rsloQy.AddNew
      rsloQy!QueryName = dbx.QueryDefs(i).Name
      rsloQy!QueryDef = dbx.QueryDefs(i).SQL
      rsloQy.Update

    End If

  Next

  CloseAllDB

  putText "Finished."

  lblStatus.Caption = "Master schema database created"

  EnableControls

  If iMode = 1 Then
    End
  End If

  Screen.MousePointer = vbNormal

  Exit Sub
ErrProc:
  If iMode = 1 Then
    WriteErrLog Err.Number & " : " & Err.Description
  Else
    MsgBox Err.Description, vbExclamation, dbx.TableDefs(i).Name
  End If
  Resume Jump
End Sub



Private Sub Command1_Click()
  Dim sGood As String
  Dim sBad As String

  On Error Resume Next

  sDBpassword = ""


  Text1.Text = OpenArqExt(Me, "", "MDB", "Microsoft Access (*.mdb)")

  If IsNull(Text1.Text) Then
     Exit Sub
  End If

  sDir = Text1.Text
  sDir = Left(sDir, InStrRev(sDir, "\") - 1)
  SaveSetting App.Title, "Options", "Dir", sDir

  'attempt to open the database shared, read-only
  Set dbTemp = OpenDatabase(Text1.Text, False, True)
  If Err.Number = 3031 Then
    Err.Clear
    frmPassword.Show vbModal
    If sDBpassword <> "" Then
      Set dbTemp = OpenDatabase(Text1.Text, False, True, ";pwd=" & sDBpassword)
      If Err.Number > 0 Then
        lblStatus.Caption = "Unable to open database."
        lblStatus.Refresh
        Text1.Text = ""
        MsgBox Err.Number & " : " & Err.Description, vbExclamation
      Else
        dbTemp.Close
        Set dbTemp = Nothing
        lblStatus.Caption = "Database opened successfully."
        lblStatus.Refresh
      End If
    Else
      lblStatus.Caption = "Unable to open database - password not provided."
      lblStatus.Refresh
      Text1.Text = ""
    End If
  ElseIf Err.Number > 0 Then  '3343 Then 'corrupt, attempt to fix
    If MsgBox(Err.Number & " : " & Err.Description & vbCrLf & "The database is corrupted. Attempt to repair it?", vbYesNo + vbQuestion) = vbNo Then
      Text1.Text = ""
      lblStatus.Caption = ""
      Exit Sub
    End If
    Err.Clear
    sGood = sDir & "\good_" & Format(Date, "mmddyy") & ".mdb"
    If FileExists(sGood) Then
      DeleteFile sGood  'Kill sGood
    End If
    Screen.MousePointer = vbHourglass
    CompactDatabase Text1.Text, sGood
    Screen.MousePointer = vbDefault
    If Err.Number Then
      lblStatus.Caption = "Unable to repair database."
      lblStatus.Refresh
      Text1.Text = ""
      MsgBox Err.Number & " : " & Err.Description, vbExclamation
    Else
      If FileExists(sGood) Then
        sBad = sDir & "\bad_" & Format(Date, "mmddyy") & ".mdb"
        'Rename corrupt database
        Name Text1.Text As sBad
        'Rename fixed database
        Name sGood As Text1.Text
        MsgBox "The database was repaired." & vbCrLf & _
               "The corrupted database has been renamed as " & sBad & vbCrLf & _
               "If the database still has problems you should run 'Fix Corrupt Database'"
      Else
        lblStatus.Caption = "Unable to repair database."
        lblStatus.Refresh
        Text1.Text = ""
      End If
    End If
  Else
    dbTemp.Close
    Set dbTemp = Nothing
    lblStatus.Caption = "Database opened successfully."
    lblStatus.Refresh
  End If

  'End If
End Sub

Private Sub putText(teks As String)
  Text2.Text = Text2.Text & Chr(13) & Chr(10) & teks
End Sub

Private Function UpdNewStat(oldstat As String, At_Digit As Integer)
  UpdNewStat = Mid(oldstat, 1, At_Digit - 1) & "U" & Mid(oldstat, At_Digit + 1)
End Function

Private Sub CreateLocalTable(S As String)
  Dim NewTable As TableDef
  Dim NewField As Field
  Dim db As String

  On Error Resume Next

  db = Mid(S, InStrRev(S, "\") + 1)

  ' if dbFixTmp already exist, delete it
  If FileExists(S) Then 'StrComp(Dir(S), db, vbTextCompare) = 0 Then
    DeleteFile S  'Kill S
  End If

  Set dbTemp = CreateDatabase(S, dbLangGeneral)
  Set NewTable = Nothing
  Set NewTable = dbTemp.CreateTableDef("myTABLE")
  Set NewField = Nothing
  Set NewField = NewTable.CreateField("TableName")
  NewField.Type = 10
  NewField.Attributes = 2
  NewField.Size = 64  '64 is max length
  NewField.AllowZeroLength = True
  NewTable.Fields.Append NewField

  Set NewField = Nothing
  Set NewField = NewTable.CreateField("SeqNum")
  NewField.Type = 4
  NewField.Attributes = 1
  NewField.Size = 4
  NewTable.Fields.Append NewField

  Set NewField = Nothing
  Set NewField = NewTable.CreateField("FieldName")
  NewField.Type = 10
  NewField.Attributes = 2
  NewField.Size = 64  '64 is max length
  NewField.AllowZeroLength = True
  NewTable.Fields.Append NewField

  Set NewField = Nothing
  Set NewField = NewTable.CreateField("FieldType")
  NewField.Type = 4
  NewField.Attributes = 1
  NewField.Size = 4
  NewTable.Fields.Append NewField

  Set NewField = Nothing
  Set NewField = NewTable.CreateField("Attributes")
  NewField.Type = 4
  NewField.Attributes = 1
  NewField.Size = 4
  NewTable.Fields.Append NewField

  Set NewField = Nothing
  Set NewField = NewTable.CreateField("Required")
  NewField.Type = 1
  NewField.Attributes = 1
  NewField.Size = 1
  NewTable.Fields.Append NewField

  Set NewField = Nothing
  Set NewField = NewTable.CreateField("Size")
  NewField.Type = 4
  NewField.Attributes = 1
  NewField.Size = 4
  NewTable.Fields.Append NewField

  Set NewField = Nothing
  Set NewField = NewTable.CreateField("AllowZeroLength")
  NewField.Type = 1
  NewField.Attributes = 1
  NewField.Size = 1
  NewTable.Fields.Append NewField

  Set NewField = Nothing
  Set NewField = NewTable.CreateField("DefaultValue")
  NewField.Type = 10
  NewField.Attributes = 2
  NewField.Size = 255
  NewField.AllowZeroLength = True
  NewTable.Fields.Append NewField

  Set NewField = Nothing
  Set NewField = NewTable.CreateField("Status")
  NewField.Type = 10
  NewField.Attributes = 2
  NewField.Size = 20
  NewField.AllowZeroLength = True
  NewTable.Fields.Append NewField

  dbTemp.TableDefs.Append NewTable

  Set NewTable = Nothing
  Set NewTable = dbTemp.CreateTableDef("myINDEX")
  NewTable.Fields.Append NewTable.CreateField("TableName", dbText, 64)  '64 is max length
  NewTable.Fields.Append NewTable.CreateField("IndexName", dbText, 128)  '128 is max length
  NewTable.Fields.Append NewTable.CreateField("Fields", dbMemo, 0)  '128 max IndexName length X 10 max fields per index = 1,280 max length
  NewTable.Fields.Append NewTable.CreateField("Primary", dbBoolean, 1)
  NewTable.Fields.Append NewTable.CreateField("Unique", dbBoolean, 1)
  NewTable.Fields.Append NewTable.CreateField("Status", dbText, 20)
  dbTemp.TableDefs.Append NewTable

  Set NewTable = Nothing
  Set NewTable = dbTemp.CreateTableDef("myQUERY")
  NewTable.Fields.Append NewTable.CreateField("QueryName", dbText, 64)  '64 is max length
  NewTable.Fields.Append NewTable.CreateField("QueryDef", dbMemo, 0)
  NewTable.Fields.Append NewTable.CreateField("Status", dbText, 20)
  dbTemp.TableDefs.Append NewTable
  dbTemp.Close

End Sub

Private Function OpenAllDB(S As String, Optional bExclusive As Boolean, Optional bReadOnly As Boolean) As Boolean
  On Error Resume Next
'  dbx.Close
'  Set dbx = Nothing
'  dbTemp.Close
'  Set dbTemp = Nothing

  'On Error GoTo 0
  On Error GoTo Err
  If sDBpassword = "" Then
    Set dbx = OpenDatabase(Text1.Text, bExclusive, bReadOnly)
  Else
    Set dbx = OpenDatabase(Text1.Text, bExclusive, bReadOnly, ";pwd=" & sDBpassword)
  End If
  Set dbTemp = OpenDatabase(S, True, False)
  OpenAllDB = True

  Exit Function
Err:
  If iMode = 1 Or iMode = 2 Then
    WriteErrLog "Unable to open databases due to error " & Err.Number & " : " & Err.Description
    End
  Else
    Text1.Text = ""
    MsgBox "Error " & Err.Number & " : " & Err.Description, vbExclamation
  End If
  OpenAllDB = False
End Function

Private Function CloseAllDB() As Boolean
  On Error Resume Next
  dbx.Close
  Set dbx = Nothing
  dbTemp.Close
  Set dbTemp = Nothing
  CloseAllDB = True
End Function

Private Sub cmdCompare_Click()

Text2.Text = ""
If Text1.Text = "" Then
   MsgBox "Enter database name.", vbExclamation
   Exit Sub
ElseIf FileExists(Text1.Text) = False Then
   If iMode = 2 Then
      WriteErrLog "Unable to compare with master schema file. " & Text1.Text & " does not exist"
      End
   Else
      Text1.Text = ""
      lblStatus.Caption = "The database does not exist"
      Exit Sub
   End If
End If

If bFix = False Then
   If chkNoMods.Value = False Then
      If MsgBox("You are about to make changes to the database." & vbCrLf & "Do you want to continue?", vbYesNo + vbQuestion) = vbNo Then
         Exit Sub
      End If
   End If
End If

Dim S As String
Dim sPath As String

On Error Resume Next

If iMode = 0 Then

   If bFix = False Then
      S = Text1.Text
      S = Replace(S, ".mdb", "_Schema.mdb", 1, 1, vbTextCompare)
   End If

   bFix = False

   sPath = Text1.Text
   sPath = Left(sPath, InStrRev(sPath, "\") - 1)

   CDL1.CancelError = True
   CDL1.DefaultExt = "mdb"
   CDL1.InitDir = sPath
   CDL1.Filter = "Microsoft Access (*.mdb)|*.mdb|All Files (*.*)|*.*"
   CDL1.filename = S

   CDL1.ShowOpen

   If Err.Number Then  'Cancel selected
      Exit Sub
   End If

   If LCase(Right(CDL1.filename, 11)) <> "_schema.mdb" Then
      MsgBox "The master schema database name must end with '_schema.mdb'.", vbExclamation
      Exit Sub
   End If

   sSchemaDB = CDL1.filename

ElseIf FileExists(sSchemaDB) = False Then
   If iMode = 2 Then
      WriteErrLog "Master schema file " & sSchemaDB & " does not exist"
      End
   Else
      lblStatus.Caption = "The Master Schema Database does not exist"
      Exit Sub
   End If
End If

If StrComp(sSchemaDB, Text1.Text, vbTextCompare) = 0 Then
   If iMode = 2 Then
      WriteErrLog "The master schema database cannot be the same file as the selected database"
      End
   Else
      MsgBox "The master schema database cannot be the same file as the selected database.", vbExclamation
      Exit Sub
   End If
End If

'open database
If chkNoMods.Value Then
   'Open database to modify shared, not read-only
   If Not OpenAllDB(sSchemaDB, False, False) Then
      Exit Sub
   End If
Else
   'Open database to modify exclusively, not read-only
   If Not OpenAllDB(sSchemaDB, True, False) Then
      Exit Sub
   End If
End If

Dim rsloTb As Recordset
Dim rsloQy As Recordset
Dim rsloIx As Recordset
Dim rsClone As Recordset
Dim SQL As String

Dim NewTable As TableDef
Dim NewField As Field
Dim NewQuery As QueryDef
Dim NewIndex As Index
Dim NewStat As String

Dim isNewTable As Boolean
'Dim isNewField As Boolean
'Dim isNewIndex As Boolean
'Dim isNewQuery As Boolean
Dim blnNoMods As Boolean
Dim srIgnore As String
Dim bCorrupt As Boolean

On Error GoTo Err

Screen.MousePointer = vbHourglass

DisableControls

lblStatus.Caption = "Comparing..."
lblStatus.Refresh

blnNoMods = chkNoMods.Value

'Status set to NEW
dbTemp.Execute "update myTABLE set [status] = 'NEW'"
dbTemp.Execute "update myINDEX set [status] = 'NEW'"
dbTemp.Execute "update myQUERY set [status] = 'NEW'"

Set rsloTb = dbTemp.OpenRecordset("select * from myTABLE")
Set rsloQy = dbTemp.OpenRecordset("select * from myQUERY")
Set rsloIx = dbTemp.OpenRecordset("select * from myINDEX")

Dim i%, J%, K%, l%, mPB1%, mPB2%
DoEvents
putText "Checking..."

' Checking index and TableDef
For i = 0 To dbx.TableDefs.Count - 1

   PB1.Value = i / dbx.TableDefs.Count * 40
   DoEvents

   If dbx.TableDefs(i).Updatable = False Then  'this causes linked tables to be ignored
      putText dbx.TableDefs(i).Name & " is not updateable thus is ignored. May be a linked table."
      srIgnore = srIgnore & " " & dbx.TableDefs(i).Name
   Else

      'MsgBox dbx.TableDefs(i).Updatable, vbOKOnly, dbx.TableDefs(i).Name

      'Ignore system files
      If LCase(Mid(dbx.TableDefs(i).Name, 1, 4)) <> "msys" Then
         If rsloTb.RecordCount > 0 Then
            For J = 0 To dbx.TableDefs(i).Fields.Count - 1
               rsloTb.FindFirst _
                  " [TableName] = '" & dbx.TableDefs(i).Name & "' and " & _
                  " [FieldName] = '" & dbx.TableDefs(i).Fields(J).Name & "'"
               rsloTb.Edit
               If Not rsloTb.NoMatch Then
                  NewStat = "-----"
                  If rsloTb("FieldType") <> dbx.TableDefs(i).Fields(J).Type Then NewStat = UpdNewStat(NewStat, 1)
                  If rsloTb("Attributes") <> dbx.TableDefs(i).Fields(J).Attributes Then NewStat = UpdNewStat(NewStat, 2)
                  'If rsloTb("Required") <> dbx.TableDefs(i).Fields(j).Required Then NewStat = UpdNewStat(NewStat, 3)
                  If rsloTb("Size") <> dbx.TableDefs(i).Fields(J).Size Then NewStat = UpdNewStat(NewStat, 4)
                  If rsloTb("AllowZeroLength") <> dbx.TableDefs(i).Fields(J).AllowZeroLength Then NewStat = UpdNewStat(NewStat, 5)
                  rsloTb("Status") = NewStat
               End If
               rsloTb.Update
            Next  'j
         End If

         If rsloIx.RecordCount > 0 Then
            For J = 0 To dbx.TableDefs(i).Indexes.Count - 1
               rsloIx.FindFirst _
                  " [TableName] = '" & dbx.TableDefs(i).Name & "' and " & _
                  " [IndexName] = '" & dbx.TableDefs(i).Indexes(J).Name & "'"
               rsloIx.Edit
               'If dbx.TableDefs(i).Indexes(j).Name = "Doc Id" Then
               'Debug.Print "Sa"
               'End If
               If Not rsloIx.NoMatch Then
                  NewStat = "-----"
                  If UCase(rsloIx("Fields")) <> UCase(dbx.TableDefs(i).Indexes(J).Fields) Then NewStat = UpdNewStat(NewStat, 1)
                  If rsloIx("Primary") <> dbx.TableDefs(i).Indexes(J).Primary Then NewStat = UpdNewStat(NewStat, 2)
                  If rsloIx("Unique") <> dbx.TableDefs(i).Indexes(J).Unique Then NewStat = UpdNewStat(NewStat, 3)
                  rsloIx("Status") = NewStat
               End If
               rsloIx.Update
            Next  'j
         End If
      End If

   End If

Next i

' 10%
If rsloQy.RecordCount > 0 Then
   For i = 0 To dbx.QueryDefs.Count - 1
      PB1.Value = i / dbx.QueryDefs.Count * 10 + 40
      DoEvents
      rsloQy.FindFirst _
         " [QueryName] = '" & dbx.QueryDefs(i).Name & "'"
      rsloQy.Edit

      If Not rsloQy.NoMatch Then
         NewStat = "-----"
         If rsloQy!QueryDef <> dbx.QueryDefs(i).SQL Then NewStat = "U"
         rsloQy("Status") = NewStat
      End If
      rsloQy.Update
   Next
End If

'------------------------------------------------------------------
' Fix the database
' Progress Bar 30% untuk tabledef
putText "Comparing Table Definitions..."
Set rsClone = rsloTb.Clone
If rsloTb.RecordCount > 0 Then
   rsloTb.MoveFirst
   rsloTb.MoveLast
   mPB1 = rsloTb.RecordCount
   rsloTb.MoveFirst
   mPB2 = 0

   Do While Not rsloTb.EOF
      isNewTable = False
      mPB2 = mPB2 + 1
      PB1.Value = 50 + mPB2 / mPB1 * 40
      DoEvents
      If rsloTb("Status") = "NEW" Then
         ' New field or new Table
         isNewTable = True
         For i = 0 To dbx.TableDefs.Count - 1
            If dbx.TableDefs(i).Name = rsloTb!TableName Then isNewTable = False
         Next

         If isNewTable Then
            bCorrupt = True
            isNewTable = True
            Set NewTable = Nothing
            Set NewTable = dbx.CreateTableDef(rsloTb!TableName)
            putText " Create Table: " & rsloTb!TableName
         Else
            Set NewTable = Nothing
            Set NewTable = dbx.TableDefs(rsloTb!TableName)
         End If

         If NewTable.Updatable Then

            bCorrupt = True
            putText " Create Field: " & rsloTb!TableName & "." & rsloTb!FieldName

            Set NewField = Nothing
            Set NewField = NewTable.CreateField(rsloTb!FieldName)
            NewField.Type = rsloTb!FieldType
            NewField.Attributes = rsloTb!Attributes
            NewField.Size = rsloTb![Size]
            NewField.OrdinalPosition = rsloTb!SeqNum

            If blnNoMods = False Then

               If rsloTb!FieldType = 10 Then NewField.AllowZeroLength = rsloTb![AllowZeroLength]

               NewTable.Fields.Append NewField

               If isNewTable Then dbx.TableDefs.Append NewTable
               ' Assign new field
               If rsloTb!FieldType = 1 Or rsloTb!FieldType = 4 Or rsloTb!FieldType = 5 Then
                  SQL = "update [" & rsloTb!TableName & "] set [" & rsloTb!FieldName & "] = 0"
                  dbx.Execute SQL
               ElseIf rsloTb!FieldType = 10 Then
                  If IsNull(rsloTb!DefaultValue) Or rsloTb!DefaultValue = "" Then
                     SQL = "update [" & rsloTb!TableName & "] set [" & rsloTb!FieldName & "] = ''"
                  Else
                     SQL = "update [" & rsloTb!TableName & "] set [" & rsloTb!FieldName & "] = " & rsloTb!DefaultValue & ""
                  End If
                  dbx.Execute SQL
               End If

            End If

         End If

      ElseIf rsloTb!Status <> "-----" Then
         bCorrupt = True
         putText " Update Field: " & rsloTb!TableName & "." & rsloTb!FieldName

         ' Update field
         ' Cannot do update field properties,
         ' so the tricks is rename oldfield, create new field
         ' and move the content from oldfield to newfield then delete oldfield

         If blnNoMods = False Then

            Set NewTable = Nothing
            Set NewTable = dbx.TableDefs(rsloTb!TableName)

            If NewTable.Updatable Then

               ' Maybe indexes use that field, so delete indexe, later create again
               On Error Resume Next
               For l = 0 To dbx.TableDefs(rsloTb!TableName).Indexes.Count - 1
                  NewTable.Indexes.Delete dbx.TableDefs(rsloTb!TableName).Indexes(l).Name
               Next l

               'rename
               Set NewField = Nothing
               Set NewField = NewTable.Fields(rsloTb!FieldName)
               NewField.Name = rsloTb!FieldName & "_x"

               'create new
               Set NewField = Nothing
               Set NewField = NewTable.CreateField(rsloTb!FieldName)
               NewField.Type = rsloTb!FieldType
               NewField.Attributes = rsloTb!Attributes
               NewField.Size = rsloTb![Size]
               If rsloTb!FieldType = 10 Then NewField.AllowZeroLength = rsloTb![AllowZeroLength]
               NewField.OrdinalPosition = rsloTb!SeqNum
               NewTable.Fields.Append NewField

               ' copying content
               SQL = "update [" & rsloTb!TableName & "] set [" & rsloTb!FieldName & "] = [" & rsloTb!FieldName & "_x]"
               dbx.Execute SQL

               ' delete old field
               NewTable.Fields.Delete rsloTb!FieldName & "_x"

            End If

         End If

         On Error GoTo Err
      End If
      rsloTb.Edit
      rsloTb("Status") = rsloTb("Status") & "  DONE"
      rsloTb.Update
      rsloTb.MoveNext
   Loop
Else
   PB1.Value = 80
   DoEvents
End If

putText "Comparing Index Definitions..."
'open database
If chkNoMods.Value Then
   'Open database to modify shared, read-only
   If Not OpenAllDB(sSchemaDB, False, True) Then
      EnableControls
      lblStatus.Caption = "Unable to open databases"
      Exit Sub
   End If
Else
   'Open database to modify exclusively, not read-only
   If Not OpenAllDB(sSchemaDB, True, False) Then
      EnableControls
      lblStatus.Caption = "Unable to open databases"
      Exit Sub
   End If
End If

Set rsloTb = dbTemp.OpenRecordset("select * from myTABLE")
Set rsloQy = dbTemp.OpenRecordset("select * from myQUERY")
Set rsloIx = dbTemp.OpenRecordset("select * from myINDEX")

If rsloIx.RecordCount > 0 Then
   rsloIx.MoveFirst
   rsloIx.MoveLast
   mPB1 = rsloIx.RecordCount
   mPB2 = 0
   rsloIx.MoveFirst
   Do While Not rsloIx.EOF
      mPB2 = mPB2 + 1
      PB1.Value = 80 + mPB2 / mPB1 * 10
      DoEvents
      If rsloIx("Status") <> "-----" Then
         'New atau update
         'If update state just delete index and create new index
         'Set NewTable = Nothing
         Set NewTable = dbx.TableDefs(rsloIx!TableName)
         If InStr(1, srIgnore, NewTable.Name, vbTextCompare) Then
            'MsgBox "ignore"
         Else

            If rsloIx("Status") <> "NEW" Then
               bCorrupt = True
               putText " Update Index: " & rsloIx!TableName & "." & rsloIx!Indexname
               ' Delete old index, later create new index
               If blnNoMods = False Then
                  NewTable.Indexes.Delete rsloIx!Indexname
               End If
            Else
               bCorrupt = True
               putText " Create Index: " & rsloIx!TableName & "." & rsloIx!Indexname
            End If

            Set NewIndex = Nothing
            If blnNoMods = False Then
               Set NewIndex = NewTable.CreateIndex(rsloIx!Indexname)
               NewIndex.Fields = rsloIx("Fields")
               NewIndex.Primary = rsloIx("Primary")
               NewIndex.Unique = rsloIx("Unique")
               'set newindex
               dbx.TableDefs(rsloIx!TableName).Indexes.Append NewIndex
            End If

         End If

      End If
      rsloIx.MoveNext
   Loop
Else
   PB1.Value = 90
   DoEvents
End If

putText "Comparing Query Definitions..."
' Progress bar = 10%
If rsloQy.RecordCount > 0 Then
   rsloQy.MoveFirst
   rsloQy.MoveLast
   mPB1 = rsloQy.RecordCount
   mPB2 = 0
   rsloQy.MoveFirst
   Do While Not rsloQy.EOF
      mPB2 = mPB2 + 1
      PB1.Value = 90 + mPB2 / mPB1 * 10
      DoEvents
      If rsloQy("Status") <> "-----" Then
         If rsloQy("Status") <> "NEW" Then
            ' If query is different, just delete it
            bCorrupt = True
            putText " Update Query: " & rsloQy!QueryName
            If blnNoMods = False Then
               dbx.QueryDefs.Delete rsloQy!QueryName
            End If
         Else
            bCorrupt = True
            putText " Create Query: " & rsloQy!QueryName
         End If
         Set NewQuery = Nothing
         If blnNoMods = False Then
            Set NewQuery = dbx.CreateQueryDef(rsloQy!QueryName, rsloQy!QueryDef)
         End If
      End If
      rsloQy.MoveNext
   Loop
Else
   PB1.Value = 100
   DoEvents
End If
putText "Finished."

'On Error GoTo 0

'  db.Relations.Delete "CategoriesProducts"

'   Set rel = db.CreateRelation()
'   rel.Name = "CategoriesProducts"
'   rel.Table = "Categories"
'   rel.ForeignTable = "Products"
'
'   ' Specify cascading updates and deletes
'   rel.Attributes = dbRelationUpdateCascade Or _
'      dbRelationDeleteCascade
'
'   ' Create the field the tables are related on
'   Set fld = rel.CreateField("CategoryId")
'   ' Set ForeignName property of the field to the name of
'   ' the corresponding field in the primary table
'   fld.ForeignName = "CategoryId"
'
'   rel.Fields.Append fld
'
'   ' Append the relation to the collection
'   db.Relations.Append rel

SQL = Format(Now, "hhmmss")

If bRel = True And blnNoMods = False Then  'Add relationships
   'dbx.Execute "DELETE FROM MSysRelationships"

   'For i = 0 To UBound(oRels)
   For i = 0 To UBound(RelData)

      'Set oRel = dbx.CreateRelation()
      'oRel.Name = "CategoriesProducts"
      'oRel.Table = "Categories"
      'oRel.ForeignTable = "Products"

      Set oRel = dbx.CreateRelation(RelData(i).sRelation & SQL, RelData(i).sTable, RelData(i).sFTable, RelData(i).lAttr)
      'Set oRel = dbx.CreateRelation(sRelation(i) & Format(Now, "hhmmss"), sTable(i), sFTable(i), lAttr(i))

      'Get error "Index already exists" with code below
      'oRel.Attributes = dbRelationUpdateCascade Or dbRelationDeleteCascade
      'oRel.Attributes = dbRelationDontEnforce

      'append fields
      For J = 0 To UBound(RelData(i).sField)
         Set fld = oRel.CreateField(RelData(i).sField(J))
         fld.ForeignName = RelData(i).sFField(J)
         oRel.Fields.Append fld
      Next

      'Set fld = oRel.CreateField(sField(i))
      'Set fld = oRel.CreateField("CategoryId")

      'fld.ForeignName = sFField(i)
      'fld.ForeignName = "CategoryId"

      'oRel.Fields.Append fld
      dbx.Relations.Append oRel
   Next
   Set oRel = Nothing
   Set fld = Nothing
   'Erase oRels
   'Erase sRelation
   'Erase sTable
   'Erase sFTable
   'Erase lAttr
   'Erase sField
   Erase RelData
   bRel = False
End If

CloseAllDB

lblStatus.Caption = "Database compared to master schema database"

If bCorrupt Then
   i = FreeFile
   NewStat = sAppPath & "results.txt"
   Open NewStat For Output As #i
   Print #i, Format(Now, "mm/dd/yyyy hh:mm:ss am/pm")
   Print #i, "Database: " & Text1.Text
   Print #i, "Schema: " & sSchemaDB
   Print #i, Text2.Text
   Close #i
End If

EnableControls

Screen.MousePointer = vbNormal

If bCorrupt Then
   If FileExists(NewStat) Then
      Shell "notepad.exe " & NewStat, vbNormalFocus
   End If
End If

If iMode = 2 Then
   End
End If

Exit Sub

Err:
If iMode = 2 Then
   WriteErrLog Err.Number & " : " & Err.Description
ElseIf Err.Number = 3265 Or Err.Number = 3113 Then
   '3265: ignore item not found in this collection error
   '3113: Field not updatable
Else
   Text2.Text = Text2.Text & vbCrLf & "ERROR " & CStr(Err.Number) & " : " & Err.Description & vbCrLf
End If
Resume Next
'End If
End Sub

Private Sub Command2_Click()

  If Text1.Text = "" Then
    MsgBox "Enter database name.", vbExclamation
    Exit Sub
  End If

  If MsgBox("Before running this feature you should have a master schema database available for recreating indexes and queries." & vbCrLf & "Do you want to proceed?", vbYesNo + vbQuestion) = vbNo Then
    Exit Sub
  End If

  Dim i As Integer
  Dim K As Integer
  Dim FreeSpace As Currency
  Dim strNew As String
  Dim strrepaired As String
  Dim SQL As String
  Dim strTable As String
 ' Dim ws As Workspace
  Dim strPath As String
  Dim lItem

  On Error Resume Next

  bFix = False

  strNew = parsefile(Text1.Text, "C") & "_new.mdb"
  strrepaired = parsefile(Text1.Text, "C") & "_repaired.mdb"

  If FileExists(strNew) Then
    Kill strNew
    If Err.Number Then
      lblStatus.Caption = "Unable to delete " & strNew
      Exit Sub
    End If
  End If

  If FileExists(strrepaired) Then
    Kill strrepaired
    If Err.Number Then
      lblStatus.Caption = "Unable to delete " & strrepaired
      Exit Sub
    End If
  End If

  If FileExists(Text1.Text) = False Then
    Text1.Text = ""
    lblStatus.Caption = "Database does not exist"
    Exit Sub
  End If

  'attempt to open the database shared, read-only
  If sDBpassword = "" Then
    Set dbTemp = OpenDatabase(Text1.Text, False, True)
  Else
    Set dbTemp = OpenDatabase(Text1.Text, False, True, ";pwd=" & sDBpassword)
  End If
  If Err.Number Then
    'TODO: Fix database header if it is corrupted
    lblStatus.Caption = "Unable to open database."
    lblStatus.Refresh
    MsgBox Err.Number & " : " & Err.Description, vbExclamation
    Exit Sub
  Else
    dbTemp.Close
    Set dbTemp = Nothing
  End If

  FreeSpace = GetDiskSpaceFree(Left(strPath, 3))

  'MsgBox FreeSpace

  'If FreeSpace < (FileLen(Text1.Text) * 2) Then
  '   lblStatus.Caption = "Unable to repair due to insufficient disk space."
  '   Exit Sub
  'End If

  Screen.MousePointer = vbHourglass

  DisableControls

  Text2.Text = ""

  'make a copy of corrupt database
  lblStatus.Caption = "Making a copy of the database..."
  lblStatus.Refresh
  'FileCopy Text1.Text, strNew
  FileCopy Text1.Text, strNew
  If Err.Number Then
    Screen.MousePointer = vbDefault
    lblStatus.Caption = "Unable to copy database. If the database is open, close it and try again."
    lblStatus.Refresh
    MsgBox Err.Number & " : " & Err.Description, vbExclamation
    EnableControls
    Exit Sub
  End If

  lblStatus.Caption = "Making a copy of the database...done."
  lblStatus.Refresh

  'attempt to open the new database exclusively, non read-only
  If sDBpassword = "" Then
    Set dbTemp = OpenDatabase(strNew, True, False)
  Else
    Set dbTemp = OpenDatabase(strNew, True, False, ";pwd=" & sDBpassword)
    dbTemp.NewPassword sDBpassword, ""  'remove password
  End If

  If Err.Number Then
    Screen.MousePointer = vbDefault
    lblStatus.Caption = "Unable to open database copy."
    lblStatus.Refresh
    MsgBox Err.Number & " : " & Err.Description, vbExclamation
    EnableControls
    Exit Sub
  Else
    Load frmProgress
    frmProgress.ListView1.ColumnHeaders.Add , , "Table Name", 5000, LvwColumnHeaderAlignmentLeft
    'populate list
    For i = 0 To dbTemp.TableDefs.Count - 1
      strTable = dbTemp.TableDefs(i).Name
      If LCase(Left(strTable, 4)) <> "msys" Then
        'MsgBox strTable
        Set lItem = frmProgress.ListView1.ListItems.Add(, , strTable)
       ' frmProgress.ListView1.ListItems.Add(,,
        
        
        
      End If
    Next
    frmProgress.Show

    'delete relationships - this is necessary so that tables can be deleted

    'first store them in an array
    'For i = 0 To dbTemp.Relations.Count - 1

    'On Error GoTo 0

    If dbTemp.Relations.Count > 0 Then
      bRel = True
      i = 0
      'store relation data
      ReDim RelData(dbTemp.Relations.Count - 1)
      For Each oRel In dbTemp.Relations
        Set oRel = dbTemp.Relations(i)
        'ReDim Preserve RelData(i)
        Set RelData(i).oRels = oRel
        RelData(i).sRelation = oRel.Name
        RelData(i).sTable = oRel.Table
        RelData(i).sFTable = oRel.ForeignTable
        RelData(i).lAttr = oRel.Attributes
        'ReDim Preserve oRels(i)
        'ReDim Preserve sRelation(i)
        'ReDim Preserve sTable(i)
        'ReDim Preserve sFTable(i)
        'ReDim Preserve lAttr(i)
        'Set oRels(i) = oRel
        'sRelation(i) = oRel.Name
        'sTable(i) = oRel.Table
        'sFTable(i) = oRel.ForeignTable
        'lAttr(i) = oRel.Attributes
        Set flds = oRel.Fields
        'MsgBox flds.Count
        K = 0
        ReDim RelData(i).sField(flds.Count - 1)
        ReDim RelData(i).sFField(flds.Count - 1)
        For Each fld In flds
          'ReDim Preserve sField(i)
          'ReDim Preserve sFField(i)
          'ReDim Preserve RelData(i).sField(k)
          'ReDim Preserve RelData(i).sFField(k)
          'sField(i) = fld.Name
          'sFField(i) = fld.ForeignName
          RelData(i).sField(K) = fld.Name
          RelData(i).sFField(K) = fld.ForeignName
          K = K + 1
        Next
        i = i + 1
      Next

      'then delete them
      'For i = 0 To UBound(oRels)
      For i = 0 To UBound(RelData)
        'Set oRel = oRels(i)
        'MsgBox oRel.Name
        'oRel.Fields.Delete sField(i)
        dbTemp.Relations.Delete RelData(i).sRelation
        'dbTemp.Relations.Delete oRel.Name
      Next
      'dbTemp.Relations.Refresh
    End If

    ' For Each oRel In dbTemp.Relations
    'oRel.Attributes = dbRelationDeleteCascade + dbRelationUpdateCascade
    '     dbTemp.Relations.Delete oRel.Name
    'dbTemp.Relations.Refresh
    'Next

    'For i = 0 To dbTemp.Relations.Count - 1
    '    Set oRel = dbTemp.Relations(i)
    '    dbTemp.Relations.Delete oRel.Name
    'Next

    'delete all tables in new database
    K = 1
    frmProgress.Label1.Caption = "Deleting Tables in New MDB..."
    frmProgress.Label1.Refresh
    For i = 0 To dbTemp.TableDefs.Count - 1
      strTable = dbTemp.TableDefs(i).Name
      If LCase(Left(strTable, 4)) <> "msys" Then
        'MsgBox strTable
        Set lItem = frmProgress.ListView1.ListItems.Item(K)
        lItem.Checked = True
        lItem.EnsureVisible
        frmProgress.Refresh
        frmProgress.Label1.Caption = "Deleting Table " & CStr(K) & " of " & frmProgress.ListView1.ListItems.Count & " Tables..."
        frmProgress.Label1.Refresh
        frmProgress.ListView1.Refresh
        K = K + 1
        'Sleep 1000 'testing
        SQL = "DROP TABLE [" & strTable & "]"
        lblStatus.Caption = "Deleting table " & strTable & " in new database..."
        lblStatus.Refresh
        dbTemp.Execute SQL
      End If
    Next

    If sDBpassword = "" Then
      Set dbx = OpenDatabase(Text1.Text, False, True)
    Else
      Set dbx = OpenDatabase(Text1.Text, False, True, ";pwd=" & sDBpassword)
    End If

    On Error GoTo 0

    If Err.Number Then
      lblStatus.Caption = "Unable to open database."
      lblStatus.Refresh
      MsgBox Err.Number & " : " & Err.Description, vbExclamation
      Err.Clear
    Else
      'uncheck all items in ListView
      For K = 1 To frmProgress.ListView1.ListItems.Count
        Set lItem = frmProgress.ListView1.ListItems.Item(K)
        lItem.Checked = False
      Next
      'create tables in new database from corrupt database
      frmProgress.Label1.Caption = "Creating Tables in New MDB..."
      frmProgress.Label1.Refresh
      K = 1
      For i = 0 To dbx.TableDefs.Count - 1
        strTable = dbx.TableDefs(i).Name
        If LCase(Left(strTable, 4)) <> "msys" Then
          'MsgBox strTable
          Set lItem = frmProgress.ListView1.ListItems.Item(K)
          lItem.Checked = True
          lItem.EnsureVisible
          frmProgress.Refresh
          frmProgress.Label1.Caption = "Creating Table " & CStr(K) & " of " & frmProgress.ListView1.ListItems.Count & " Tables..."
          frmProgress.Label1.Refresh
          frmProgress.ListView1.Refresh
          K = K + 1
          'Sleep 1000 'testing
          SQL = "SELECT [" & strTable & "].* INTO [" & strTable & "] IN '" & strNew & _
                "' FROM [" & strTable & "]"
          lblStatus.Caption = "Creating table " & strTable & " in new database..."
          lblStatus.Refresh
          dbx.Execute SQL
        End If
      Next

      Unload frmProgress
      Form1.Refresh

    End If

    CloseAllDB

    'Compact and repair new database
    lblStatus.Caption = "Compacting and repairing new database..."
    lblStatus.Refresh
    If sDBpassword = "" Then
      CompactDatabase strNew, strrepaired
    Else
      CompactDatabase strNew, strrepaired, ";pwd=" & sDBpassword, , ";pwd=" & sDBpassword
    End If
    lblStatus.Caption = "Compacting and repairing new database...done."
    lblStatus.Refresh

    Kill strNew
    Text1.Enabled = True
    Text1.Text = strrepaired
    chkNoMods.Enabled = True
    chkNoMods.Value = 0

    'If bRel Then
    '   MsgBox "You will need to manually recreate " & CStr(UBound(oRels) + 1) & " relationships.", vbInformation
    'End If

    lblStatus.Caption = "Now click button 'Compare Database with Schema' to complete the process."
    lblStatus.Refresh
    bFix = True
  End If

  EnableControls

  Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()

'TODO: Accept command line values
  Dim i As Integer

  On Error Resume Next
  CenterFormToScreen Me


  Me.Caption = App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision
  Text1.Text = ""
  PB1.Value = 0
  sAppPath = App.Path
  If Right(sAppPath, 1) <> "\" Then
    sAppPath = sAppPath & "\"
  End If
  sDir = GetSetting(App.Title, "Options", "Dir", sAppPath)

  Dim lngCode As Long

  lngCode = GetDriveType(Left(sAppPath, 3))
  If lngCode = 2 Or lngCode = 3 Or lngCode = 5 Then
    'local drive
  Else
    'network drive
    #If Server = 0 Then
      MsgBox "Can't run from a network drive.", vbExclamation
      End
    #End If
  End If

  Dim S As String

  S = Command$

  If S = "/p" Or S = "" Then
    'MsgBox "interpretive mode"
  Else
    Dim a As Variant
    a = Split(S, "-")
    'For i = 0 To UBound(a)
    '    MsgBox a(i)
    'Next
    'a(1) = Target Database
    'a(2) = Master Schema Database
    'a(3) = Mode
    '1 = Create Master Schema Database then exit
    '2 = Compare Database with Schema then exit
    '3 = Create Master Schema Database and do not exit
    '4 = Compare Database with Schema and do not exit

    'Compare
    '"C:\Program Files\MDB Repair Tool\mdbrt.exe" -C:\Program Files\MDB Repair Tool\corrupt.mdb -C:\Program Files\MDB Repair Tool\northwind_schema.mdb -4

    'Create Master Schema
    '"C:\Program Files\MDB Repair Tool\mdbrt.exe" -C:\Program Files\MDB Repair Tool\northwind.mdb -C:\Program Files\MDB Repair Tool\northwind_schema.mdb -3

    i = UBound(a)

    If i = 3 Then  'no database password passed
      Text1.Text = Trim(a(1))
      sSchemaDB = Trim(a(2))
      iMode = CInt(a(3))
    ElseIf i = 4 Then  'database password passed
      Text1.Text = Trim(a(1))
      sSchemaDB = Trim(a(2))
      iMode = CInt(a(3))
      sDBpassword = Trim(a(4))
    Else  'Invalid number of values passed
      End
    End If

    If iMode = 2 Or iMode = 4 Then
      cmdCompare.Value = True  'trigger click event
    ElseIf iMode = 1 Or iMode = 3 Then
      cmdCreate.Value = True  'trigger click event
    End If
    iMode = 0
    'MsgBox "exe"
  End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  On Error Resume Next
  If dbx Is Nothing Then
  Else
   dbx.Close
    Set dbx = Nothing
    dbTemp.Close
    Set dbTemp = Nothing
  End If
End Sub

#If Win16 Then
Function DeleteDuplicateRecords(strTableName As String) As Boolean
' Deletes exact duplicates from the specified table.
' No user confirmation is required. Use with caution.
  Dim rst As DAO.Recordset
  'Dim rst2 As DAO.Recordset
  'Dim tdf As DAO.TableDef
  'Dim fld As DAO.Field
  Dim strSQL As String
  'Dim varBookmark As Variant
  Dim B As Boolean
  Dim S As String

  On Error GoTo ErrProc

  S = LCase(strTableName)

  Text2.Text = ""

  Select Case S
  Case "offense"
    strSQL = "SELECT Case_R_ID,SequenceNumber FROM Offense GROUP BY Case_R_ID,SequenceNumber HAVING COUNT(*) > 1;"
  Case "case"
    strSQL = "SELECT CaseCounty,LocalNumber,CaseCentury,CaseNumber,CaseType FROM [CASE] GROUP BY CaseCounty,LocalNumber,CaseCentury,CaseNumber,CaseType HAVING COUNT(*) > 1;"
  Case "event"
    strSQL = "SELECT R_ID,Case_R_ID,EventCode,Suffix FROM Event GROUP BY R_ID,Case_R_ID,EventCode,Suffix HAVING COUNT(*) > 1;"
  Case "participant"
    strSQL = "SELECT R_ID FROM Participant GROUP BY R_ID HAVING COUNT(*) > 1;"
  Case "person"
    strSQL = "SELECT R_ID FROM Person GROUP BY R_ID HAVING COUNT(*) > 1;"
  Case Else
    lblStatus.Caption = ""
    MsgBox "Valid tables are Offense, Case, Event, Participant, and Person", vbExclamation
    Exit Function
  End Select

  Set rst = dbTemp.OpenRecordset(strSQL, dbOpenDynaset, 0, dbPessimistic)

  If rst.RecordCount > 0 Then
    'Delete the duplicates!
    Select Case S
    Case "offense"
      strSQL = "SELECT * FROM Offense WHERE Case_R_ID = " & rst.Fields("Case_R_ID").Value & " AND SequenceNumber = '" & rst.Fields("SequenceNumber").Value & "' ORDER BY R_ID"
    Case "case"
      strSQL = "SELECT * FROM [CASE] WHERE CaseCounty = '" & rst.Fields("CaseCounty").Value & "' AND LocalNumber = '" & rst.Fields("LocalNumber").Value & "' AND CaseCentury = '" & rst.Fields("CaseCentury").Value & "' AND CaseNumber = '" & rst.Fields("CaseNumber").Value & "' AND CaseType = '" & rst.Fields("CaseType").Value & "' ORDER BY R_ID"
    Case "event"
      strSQL = "SELECT * FROM Event WHERE R_ID = " & rst.Fields("R_ID").Value & " AND Case_R_ID = " & rst.Fields("Case_R_ID").Value & " AND EventCode = '" & rst.Fields("EventCode").Value & "' AND Suffix = '" & rst.Fields("Suffix").Value & "' ORDER BY R_ID"
    Case "participant"
      strSQL = "SELECT * FROM Participant WHERE R_ID = " & rst.Fields("R_ID").Value & " ORDER BY R_ID"
    Case "person"
      strSQL = "SELECT * FROM Person WHERE R_ID = " & rst.Fields("R_ID").Value & " ORDER BY R_ID"
    End Select
    rst.Close

    If chkNoMods.Value Then
      Set rst = dbTemp.OpenRecordset(strSQL, dbOpenDynaset, 0, dbReadOnly)
    Else
      Set rst = dbTemp.OpenRecordset(strSQL, dbOpenDynaset, 0, dbPessimistic)
    End If

    If rst.RecordCount > 1 Then

      Select Case S
      Case "offense"
        Text2.Text = "Duplicates found in Offense Table" & vbCrLf
      Case "case"
        Text2.Text = "Duplicates found in Case Table" & vbCrLf
      Case "event"
        Text2.Text = "Duplicates found in Event Table" & vbCrLf
      Case "participant"
        Text2.Text = "Duplicates found in Participant Table" & vbCrLf
      Case "person"
        Text2.Text = "Duplicates found in Person Table" & vbCrLf
      End Select

      rst.MoveNext  'Important! Do not delete the first record!
      Do While Not rst.EOF
        Select Case S
        Case "offense"
          Text2.Text = Text2.Text & "R_ID = " & rst.Fields("R_ID").Value & ", Case_R_ID = " & rst.Fields("Case_R_ID").Value & ", SequenceNumber = '" & rst.Fields("SequenceNumber").Value & "'" & vbCrLf
        Case "case"
          Text2.Text = Text2.Text & "R_ID = " & rst.Fields("R_ID").Value & ", CaseCounty = '" & rst.Fields("CaseCounty").Value & "', LocalNumber = '" & rst.Fields("LocalNumber").Value & "', CaseCentury = '" & rst.Fields("CaseCentury").Value & "', CaseNumber = '" & rst.Fields("CaseNumber").Value & "', CaseType = '" & rst.Fields("CaseType").Value & "'" & vbCrLf
        Case "event"
          Text2.Text = Text2.Text & "R_ID = " & rst.Fields("R_ID").Value & ", Case_R_ID = " & rst.Fields("Case_R_ID").Value & ", EventCode = '" & rst.Fields("EventCode").Value & "', Suffix = '" & rst.Fields("Suffix").Value & "'" & vbCrLf
        Case "participant"
          Text2.Text = Text2.Text & "R_ID = " & rst.Fields("R_ID").Value & vbCrLf
        Case "person"
          Text2.Text = Text2.Text & "R_ID = " & rst.Fields("R_ID").Value & vbCrLf
        End Select
        If chkNoMods.Value = False Then
          rst.Delete
          B = True
        End If
        rst.MoveNext
      Loop
      If chkNoMods.Value Then
        Dim i As Integer
        Dim NewStat As String
        i = FreeFile
        NewStat = sAppPath & "results.txt"
        Open NewStat For Output As #i
        Print #i, Text2.Text
        Close #i
        Shell "notepad.exe " & NewStat, vbNormalFocus
      End If
    End If

  End If

  rst.Close
  Set rst = Nothing

  DeleteDuplicateRecords = B

  Exit Function

ErrProc:

  Text2.Text = Text2.Text & vbCrLf & "ERROR " & CStr(Err.Number) & " : " & Err.Description & vbCrLf

  DeleteDuplicateRecords = B

End Function
#End If
Function GetDiskSpaceFree(drive As String) As Currency

  Dim dl As Long
  Dim S$
  Dim SectersPerCluster&
  Dim BytesPerSector&
  Dim NumFreeClusters&
  Dim TotNumofClusters&
  Dim NumFreeBytes&
  Dim BytesPerCluster&

  Dim Status As Long
  Dim TotalBytes As Long
  Dim FreeBytes As Long
  Dim BytesAvailableToCaller As Long

  On Error GoTo GetDiskSpaceFreeExError

  S$ = drive

  Status = GetDiskFreeSpaceEx(S$, CLng(BytesAvailableToCaller), CLng(TotalBytes), CLng(FreeBytes))
  If Status <> 0 Then
    GetDiskSpaceFree = FreeBytes * 10000
  Else
    GetDiskSpaceFree = 0
  End If
  Exit Function

GetDiskSpaceFreeExError:
  If Err = 453 Then  ' specified dll not found (early versions of Windows 95)
    On Error GoTo GetDiskSpaceFreeError
    dl = GetDiskFreeSpace(S$, CLng(SectersPerCluster), CLng(BytesPerSector), CLng(NumFreeClusters), CLng(TotNumofClusters))
    BytesPerCluster = SectersPerCluster * BytesPerSector
    NumFreeBytes = BytesPerCluster * NumFreeClusters
    GetDiskSpaceFree = NumFreeBytes
    Exit Function
  End If

GetDiskSpaceFreeError:

  If Err = 6 Then  ' more than 2 gigabytes free
    GetDiskSpaceFree = 2000000000
  Else
    GetDiskSpaceFree = 0
  End If

  Exit Function

End Function

Private Sub DisableControls()
  Dim ctl As Control
  On Error Resume Next
  For Each ctl In Controls
   ' ctl.Enabled = False
    If Err.Number Then Err.Clear
  Next
  Text2.Enabled = True
  PB1.Enabled = True
  lblStatus.Enabled = True
End Sub

Private Sub EnableControls()
  Dim ctl As Control
  On Error Resume Next
  For Each ctl In Controls
  '  ctl.Enabled = True
    If Err.Number Then Err.Clear
  Next
End Sub

Private Sub WriteErrLog(S As String)
  Dim i As Integer
  On Error Resume Next
  i = FreeFile
  Open App.Path & "\errors.txt" For Append As #i
  Print #i, Format(Now, "mm/dd/yy hh:mm:ss am/pm") & vbTab & S
  Close #i
End Sub

