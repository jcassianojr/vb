VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_Convertion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Declare local variable
Private CN_source                    As ADODB.Connection
Private MyCN                         As ADODB.Connection
Private sMySQLTableType              As String

' added 2004-11-02 {to respon cancel event}
Public EndProgress                   As Boolean

' Public event
Public Event Progress(ProgressStatus As Long)
Public Event ExecuteInfo(StrSQL As String)

'**********************************************************
' Created date: 2004-10-29 08:46
' Purpose: Convert ADO column type to MySQL column type
'**********************************************************
Private Function ConvType(ByVal TypeVal As Long) As String

  Select Case TypeVal
  
      Case adBigInt:              ConvType = "BigInt"
      Case adBinary:              ConvType = "Blob"
      Case adBoolean:             ConvType = "TinyInt(1)"
      Case adBSTR:                ConvType = "Varchar(255)"
      Case adChar:                ConvType = "Char"
      Case adCurrency:            ConvType = "decimal(16,3)"
      Case adDate:                ConvType = "date"
      Case adDBDate:              ConvType = "Date"
      Case adDBTime:              ConvType = "Time"
      Case adDBTimeStamp:         ConvType = "TimeStamp(14)"
      Case adDecimal:             ConvType = "decimal(16,3)"
      Case adDouble:              ConvType = "Double"
      Case adInteger:             ConvType = "Int"
      Case adLongVarBinary:       ConvType = "LongBlob"
      Case adLongVarChar:         ConvType = "MediumText"
      Case adLongVarWChar:        ConvType = "MediumText"
      Case adNumeric:             ConvType = "decimal"
      Case adSingle:              ConvType = "Float"
      Case adSmallInt:            ConvType = "SmallInt"
      Case adTinyInt:             ConvType = "TinyInt"
      Case adUnsignedBigInt:      ConvType = "BigInt Unsigned"
      Case adUnsignedInt:         ConvType = "Int Unsigned"
      Case adUnsignedSmallInt:    ConvType = "SmallInt Unsigned"
      Case adUnsignedTinyInt:     ConvType = "TinyInt Unsigned"
      Case adVarBinary:           ConvType = "Blob"
      Case adVarChar:             ConvType = "Varchar"
      Case adVarWChar:            ConvType = "Varchar"
      Case adWChar:               ConvType = "Char"
      Case Else
           ConvType = "Varchar(255)"
           
   End Select
   
End Function

'*******************************************************************
' Created date: 2004-10-29 08:54
' Purpose: get MySQL column type from adodb field
'*******************************************************************
Public Function ConvFieldType(ByVal iField As ADODB.Field) As String
    Dim stmp As String
    
    ' get mysql column type
    stmp = LCase(ConvType(iField.Type))
    
    ' check column type
    If stmp = "varchar" Or stmp = "char" Then
        stmp = stmp & "(" & iField.DefinedSize & ")"
    End If
    
    ' Return field convertion
    ConvFieldType = stmp
    
    'clean up
    stmp = vbNullString
    
End Function

'*********************************************************************
' Created date: 2004-10-29 09:5
' Purpose: Convert Access Table to MySQL Table
'*********************************************************************
Public Sub ConvTable(ByVal sTableName As String, _
    Optional WithData As Boolean = False, _
    Optional DropTable As Boolean = True, _
    Optional DumpSQLonly As Boolean = False, Optional ByVal TableSource As String)
    On Error Resume Next
    
    ' declare variable
    Dim Rst As New ADODB.Recordset
    Dim RstSchema As ADODB.Recordset
     
    Dim i As Long
    Dim j As Long
    Dim StrTemp As String
    Dim StrAccess As String
    Dim StrMySQL  As String
    Dim X As Long
    Dim Y As Long
    Dim pStt As Long
    Dim lPost As Long
    
    Dim SQL As String
    Dim Ant_Primary As Boolean
    Dim Ant_Tabla As String
    Dim Indices As String
    Dim SQL_Template As String
    Dim Campitos As String
    
    ' create SQL statement to open access table
    StrAccess = "select * from [" & sTableName & "]"
    
    
    'StrTemp = "Create table `" & LCase(sTableName) & "`("
    
     StrTemp = "CREATE TABLE IF NOT EXISTS `" & LCase(sTableName) & "`("
    ' set cursor location
    Rst.CursorLocation = adUseClient
    ' open recordset
    Rst.Open StrAccess, CN_source, adOpenStatic, adLockOptimistic
    ' get max field index
    i = Rst.Fields.Count - 1

    
    'looping and get sql create table (DDL)
    For j = 0 To i
        ' get column info
        StrTemp = StrTemp & "`" & LCase(Rst.Fields(j).Name) & "` " & _
                  ConvFieldType(Rst.Fields(j))
                  
        ' add separator
        If j < i Then
            StrTemp = StrTemp & " , "
        Else
            StrTemp = StrTemp & " ) "
        End If
        
    Next j
    
    
    ' check drop table
    If DropTable Then
        StrMySQL = "Drop table if exists `" & LCase(sTableName) & "`;"
        If DumpSQLonly = False Then
            ExecuteSQL StrMySQL
        End If
                    

        RaiseEvent ExecuteInfo(StrMySQL)
        DoEvents
    End If
    
    StrMySQL = StrTemp & sMySQLTableType
    
    
    If DumpSQLonly = False Then
    ' Execute SQL (DDL) "Create Table into MySQL Database"
        ExecuteSQL StrMySQL
    End If
    
    StrMySQL = StrMySQL & " ;"
    RaiseEvent ExecuteInfo(StrMySQL)
    DoEvents
      
      
      
       'Para tomar los indices de las tablas
    Set RstSchema = CN_source.OpenSchema(adSchemaIndexes)
      
      'Cargamos los indices de la tabla
             Do Until RstSchema.EOF = True
                 If RstSchema.Fields!table_name = sTableName Then
                        Ant_Tabla = sTableName
                        
                        If RstSchema.Fields!PRIMARY_KEY = True Then
                            Ant_Primary = True
                            ' El índice es la clave principal de la tabla
                            'ALTER TABLE `trn01` ADD PRIMARY KEY (nrosocio)
                            
                            If Campitos = "" Then
                                Campitos = RstSchema.Fields!COLUMN_NAME
                            Else
                                Campitos = Campitos & "," & RstSchema.Fields!COLUMN_NAME
                            End If
                            
                            SQL_Template = "ALTER TABLE " & RstSchema.Fields!table_name & " ADD PRIMARY KEY (CAMPOS)"
                                  
                            SQL = 1
                             
                        ElseIf RstSchema.Fields!Unique = True Then
                            Ant_Primary = False
                            ' El índice es único
                            '
                            SQL = 1
                            
                            If Campitos = "" Then
                                Campitos = RstSchema.Fields!COLUMN_NAME
                            Else
                                Campitos = Campitos & "," & RstSchema.Fields!COLUMN_NAME
                            End If
                            
                            SQL_Template = "CREATE UNIQUE INDEX " & RstSchema.Fields!INDEX_NAME & _
                                  " ON " & RstSchema.Fields!table_name & _
                                  "(CAMPOS)"
                
                        Else
                            ' Cualquier otro índice existente en la tabla
                            '
                            Ant_Primary = False
                            
                            If Campitos = "" Then
                                Campitos = RstSchema.Fields!COLUMN_NAME
                            Else
                                Campitos = Campitos & "," & RstSchema.Fields!COLUMN_NAME
                            End If
                            
                            'CREATE INDEX IF NOT EXISTS
                            
                            'SQL_Template = "CREATE INDEX " & RstSchema.Fields!INDEX_NAME & _
                            '      " ON " & RstSchema.Fields!table_name & _
                            '      "(CAMPOS)"
                            
                            
                            SQL_Template = "CREATE INDEX IF NOT EXISTS " & RstSchema.Fields!INDEX_NAME & _
                                  " ON " & RstSchema.Fields!table_name & _
                                  "(CAMPOS)"
                            SQL = 1
                        End If
                    End If
                    
                    RstSchema.MoveNext
                    
                    If Not RstSchema.EOF Then
                        If SQL <> "" And Ant_Tabla <> RstSchema.Fields!table_name Then
                            'Armamos la consulta de SQL
                                    
                            SQL = Replace(SQL_Template, "CAMPOS", Campitos)
                            
                            'Creamos la sentencia otra vez
                            If DumpSQLonly = False Then
                               ExecuteSQL SQL
                            End If
                            SQL = SQL & " ;"
                            RaiseEvent ExecuteInfo(SQL)
                            Ant_Tabla = ""
                            Ant_Primary = False
                            Campitos = ""
                            SQL = ""
                            SQL_Template = ""
                            
                        End If
                    End If
            Loop
       
    
    'RaiseEvent ExecuteInfo(StrMySQL)
    'DoEvents
    
    'check with insert statement or not (include data or not)
    If WithData Then
        
        ' check recordset
        If Not (Rst.EOF And Rst.BOF) Then ' make sure recordset not empty
            'Generate Insert Statement (SQL:>DML)
            Y = Rst.RecordCount
            X = 0
            Do While ((Not Rst.EOF) And (Not EndProgress))
            
                ' get SQL insert statement
                StrMySQL = GenSQLInsert(Rst, LCase(sTableName))
                If DumpSQLonly = False Then
                    ' execute SQL (DML)
                    ExecuteSQL StrMySQL
                End If
                RaiseEvent ExecuteInfo(StrMySQL)
                
                ' calculate progress (in percantage)
                X = X + 1
                pStt = (X / Y) * 100
                
                If lPost <> pStt Then
                    lPost = pStt
                    ' return progress
                    RaiseEvent Progress(lPost)
                End If
                
                ' recordset move
                Rst.MoveNext
                
                ' enable for cancel method
                DoEvents
                
            Loop
            
        End If
        
    End If
    
    If DumpSQLonly = False Then
        ' flush table
        ExecuteSQL "FLUSH TABLES"
    End If
    
    ' clean up
    StrTemp = vbNullString
    StrAccess = vbNullString
    Rst.Close
    Set Rst = Nothing

End Sub

'***********************************************************************************
' Create date: 2004-1-29 10:19
' Purpose: Generate SQL insert statement (DML)
'***********************************************************************************
Private Function GenSQLInsert(r As ADODB.Recordset, sTableName As String) As String
     
    ' declare variable
    Dim k As Long
    Dim stmp As String
    
    ' set error handler
    On Error Resume Next
    
    stmp = "Insert into `" & sTableName & "` values(" '  '" deixando aspas para criar na get value string e data
    For k = 1 To r.Fields.Count
    
    
        
        ' get value from recordset
        If Not IsNull(r.Fields(k - 1).Value) Then
            stmp = stmp & GetValue(r.Fields(k - 1))
        Else
            If TipoDado2(r.Fields(k - 1).Type) = "S" Then
              stmp = stmp & "''"
            Else
              stmp = stmp & "NULL"
            End If
        End If
    
    
        ' add separator
        If k < r.Fields.Count Then
            stmp = stmp & ","  '  ',' deixando aspas para criar na get value string e data
        Else
            stmp = stmp & ");"  '  ') deixando aspas para criar na get value string e data
        End If
    
    Next k
    
    GenSQLInsert = stmp
    
    'clean up
    stmp = vbNullString
    
    
End Function

'************************************************************************
' Created date: 2004-10-29 11:03
' Purpose: Open MS Access Database (Source)
'************************************************************************
Public Function OpenSourceDB(ByVal sFileName As String, Optional pwd As String) As Boolean
On Error Resume Next
On Error GoTo err_exit_access
    ' check connection status {Close ite if is opened}
    If CN_source.State Then
        CN_source.Close
    End If
    
    ' Try to open database
    If pwd <> vbNullString Then
        CN_source.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sFileName & ";Persist Security Info=False;Jet OLEDB:Database Password=" & pwd
    Else
        CN_source.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sFileName
    End If
    
    ' return connection state
    OpenSourceDB = CN_source.State
Exit Function
err_exit_access:
    MsgBox Err.Description, vbCritical
End Function

'************************************************************************
' Created date: 2004-10-29 11:12
' Purpose: Open Destination Database (MySQL database)
'************************************************************************
Public Function OpenDestinationDB(Optional host As String = "localhost", _
       Optional user As String = "root", Optional pwd As String, Optional port As Long = 3306, Optional MARIADBDRIVER As Boolean = False) As Boolean
  ' declare local variable
  On Error GoTo err_exit_mycnconn
  Dim sConn As String
  
  ' create connection string
    sConn = "DRIVER={MySQL ODBC 8.0 ANSI Driver};" & _
         "SERVER=" & host & ";UID=" & user & _
         ";PWD=" & pwd & ";PORT=" & port
    If MARIADBDRIVER Then
       sConn = "DRIVER={MariaDB ODBC 3.2 Driver};" & _
         "SERVER=" & host & ";UID=" & user & _
         ";PWD=" & pwd & ";PORT=" & port
    End If

    ' check connection state
    If MyCN.State Then
        MyCN.Close
    End If
    
    MyCN.Open sConn
    
    ' try to open database and return connection state
    OpenDestinationDB = MyCN.State
    
Exit Function
err_exit_mycnconn:
    MsgBox Err.Description, vbCritical
    
End Function

'***********************************************************************
' Created date: 2004-10-29 11:21
' Purpose: Use Mysql Database (SelectDB)
'***********************************************************************.
Public Sub SelectDB(sDBName As String)

    ' using 'USE' statement
    MyCN.Execute "USE `" & sDBName & "`;"
    DoEvents
    
End Sub

'***********************************************************************
' Created date: 2004-10-29 11:23
' Purpose: Create Database on MySQL server
'***********************************************************************
Public Sub createDB(sDBName As String)

    ' create database
    MyCN.Execute "create database if not exists `" & sDBName & "`;"
    ' Use these database
    SelectDB sDBName
    DoEvents
    
End Sub

'***********************************************************************
' Created date: 2004-10-29 12:32
' Purpose: Change default table type (Destination table type)
'***********************************************************************
Public Sub SetTableType(isType As String)
    
   ' sMySQLTableType = "TYPE=" & isType
    sMySQLTableType = "ENGINE=" & isType

End Sub

'***********************************************************************
' Created date: 2004-10-29 12:35
' Purpose: Return connection of source database (ADODB.Connection)
'***********************************************************************
Public Function SourceDB() As ADODB.Connection
    
    Set SourceDB = CN_source

End Function

'************************************************************************
' Created date: 2004-10-29 12:37
' Purpose: Return connection of Destination Database
'************************************************************************
Public Function DestinationDB() As ADODB.Connection
    
    Set DestinationDB = MyCN

End Function

'************************************************************************
' Created date: 2004-10-29 13:03
' Purpose: Get field value from recordset
'************************************************************************
Private Function GetValue(IsField As ADODB.Field) As String

    If Not IsNull(IsField.Value) Then
        If TipoDado2(IsField.Type) = "B" Then  ' adBoolean
            GetValue = IIf(IsField.Value = True, 1, 0) ' get boolean value
        ElseIf TipoDado2(IsField.Type) = "D" Then 'adDate
            GetValue = MySQLDate(IsField.Value) ' get mysql date format
            GetValue = "'" & GetValue & "'"
        ElseIf TipoDado2(IsField.Type) = "O" Then 'adBinary adLongVarBinary
            ' sorry i can't get value from this column type
            ' if you have solution, please let me know
            GetValue = "Null"
        ElseIf TipoDado2(IsField.Type) = "N" Then
            GetValue = IsField.Value
        Else
            ' you must replace  "'" to "''", if not will retrun error on insert statement
            'GetValue = Replace(IsField.Value, "'", "''") ' get string value
            GetValue = IsField.Value ' pegar o valor string
            GetValue = Replace(GetValue, "'", " ") ' remover aspas simples
            GetValue = Replace(GetValue, Chr(13), " ") 'remover final de linha
            GetValue = Replace(GetValue, Chr(10), " ") 'remover final de linha
            GetValue = Replace(GetValue, "(", " ") 'remover parentes
            GetValue = Replace(GetValue, ")", " ") 'remover parentes
            GetValue = Replace(GetValue, ",", " ") 'remover virgula
            GetValue = Trim(GetValue)
            GetValue = "'" & GetValue & "'"
        End If
    Else
        
        GetValue = "NULL"
    
    End If

End Function

'************************************************************************
' Created date: 2004-10-29 13:07
' Purpose: return mysql data format (yyyy-mm-dd)
'************************************************************************
Private Function MySQLDate(isdate As Date) As String
    
    MySQLDate = Format(isdate, "yyyy") & "-" & Format(isdate, "mm") & "-" & Format(isdate, "dd")

End Function

'************************************************************************
' Class Methode
'************************************************************************
Private Sub Class_Initialize()

    ' initialize
    sMySQLTableType = "TYPE=MyISAM"
    Set CN_source = New ADODB.Connection
    Set MyCN = New ADODB.Connection
    
End Sub

Private Sub Class_Terminate()
    
    ' Clean up
    If CN_source.State Then
        CN_source.Close
    End If
    Set CN_source = Nothing
    
    If MyCN.State Then
        MyCN.Close
    End If
    Set MyCN = Nothing
    
End Sub

'************************************************************************
' Created date: 2004-10-29 14:24
' Final this class for version 0.1
' Elapsed time: 05:38:xx (5.633 hours) - 1 hour (rest) = 4.633 hours
'************************************************************************


'============================== Revision =================================

'*************************************************************************
' Created date: 2004-10-30 16:23
' Purpose: Execute SQL Scripts
'*************************************************************************
Public Sub ExecuteSQL(SQL As String)
    On Error Resume Next
    MyCN.Execute SQL
End Sub

'*************************************************************************
' Created date: 2004-11-01 09:21
' Purpose: Get MySQL databases
'*************************************************************************
Public Function GetMySQLDatabases(Dbname() As String) As Long
    On Error Resume Next
    Dim r As New ADODB.Recordset
    Dim i As Integer
    Dim j As Integer
    r.CursorLocation = adUseClient
    r.Open "show databases", MyCN, adOpenStatic, adLockOptimistic
    GetMySQLDatabases = r.RecordCount
    If Not (r.EOF And r.BOF) Then
        i = r.RecordCount
        j = 1
        ReDim Dbname(i) As String
        Do While Not r.EOF
            Dbname(j) = r.Fields(0).Value
            j = j + 1
            r.MoveNext
        Loop
    End If
    r.Close
    Set r = Nothing
End Function

'*************************************************************************
' Created date: 2004-11-01 10:40
' Purpose: Get tables from source database
'*************************************************************************
Public Function GetSourceTables(tblName() As String) As Long
On Error Resume Next
    Dim Rx As New ADODB.Recordset
    Dim i As Long
    Rx.CursorLocation = adUseClient
    Set Rx = CN_source.OpenSchema(adSchemaTables)
    i = 0
    Do While Not Rx.EOF
        If Rx!table_type = "TABLE" Then
            i = i + 1
            ReDim Preserve tblName(i) As String
            tblName(i) = Rx!table_name & ":" & GetRows(Rx!table_name)
        End If
        Rx.MoveNext
    Loop
    GetSourceTables = i
    Rx.Close
    Set Rx = Nothing
End Function


'*************************************************************************
' Created date: 2004-11-01 11:28
' Purpose: Get rows
'*************************************************************************
Private Function GetRows(TableName As String) As Long
On Error Resume Next
    Dim rr As New ADODB.Recordset
    rr.CursorLocation = adUseClient
    rr.Open "select * from [" & TableName & "]", CN_source, adOpenStatic, adLockOptimistic
    GetRows = rr.RecordCount
    rr.Close
    Set rr = Nothing
End Function

