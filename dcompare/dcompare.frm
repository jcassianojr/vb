VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.1#0"; "vbccr18.ocx"
Begin VB.Form dCompare 
   Caption         =   "Database Comparador/Corretor"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8955
   Icon            =   "dcompare.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   8955
   StartUpPosition =   3  'Windows Default
   Begin VBCCR18.CommandButtonW CmdTeste 
      Height          =   375
      Left            =   6240
      TabIndex        =   14
      Top             =   1200
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "TableInfo"
   End
   Begin XPControls.XPButton CmdExportarSqlite 
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Exportar Sqlite"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VBCCR18.TextBoxW Text3 
      Height          =   2895
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2280
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   5106
      Text            =   "dcompare.frx":058A
      MultiLine       =   -1  'True
      ScrollBars      =   2
   End
   Begin VBCCR18.TextBoxW Text2 
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   600
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   661
   End
   Begin VBCCR18.TextBoxW Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   661
   End
   Begin XPControls.XPButton Command3 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Caption         =   "Compare Arquivos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorScheme     =   1
      ColorBegin      =   8454143
      ColorEnd        =   8454143
   End
   Begin XPControls.XPButton Command2 
      Height          =   375
      Left            =   8520
      TabIndex        =   3
      Top             =   600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "dcompare.frx":05B0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton Command1 
      Height          =   375
      Left            =   8520
      TabIndex        =   2
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "dcompare.frx":0B4A
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton Command4 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1200
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Caption         =   "Comparar e Corrigir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorScheme     =   1
      ColorBegin      =   8454143
      ColorEnd        =   8454143
   End
   Begin XPControls.XPButton Command5 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Picture         =   "dcompare.frx":10E4
      Caption         =   "Criar Tabela de Destino"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorScheme     =   1
      ColorBegin      =   8454143
      ColorEnd        =   8454143
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   375
      Left            =   7560
      TabIndex        =   7
      Top             =   1080
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Picture         =   "dcompare.frx":167E
      Caption         =   "&Encerrar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton CmdCompactar 
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   1680
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      Picture         =   "dcompare.frx":1C18
      Caption         =   "Compactar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorScheme     =   1
      ColorBegin      =   8454143
      ColorEnd        =   8454143
   End
   Begin VB.Label lProgress 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4920
      TabIndex        =   13
      Top             =   1680
      Width           =   3360
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Origem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   825
   End
End
Attribute VB_Name = "dCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nERRO As Integer
Private WithEvents C As cConverter
Attribute C.VB_VarHelpID = -1
'Attribute C.VB_VarHelpID = -1

Private Sub CmdCompactar_Click()
  Dim lngBefore As Long
  Dim lngAfter As Long
  Dim strResult As String
  lngBefore = FileLen(Text1.Text)
  If CompactAccess(Text1.Text) Then

    lngAfter = FileLen(Text1.Text)
    strResult = "Resultados da Compactacao" & vbCrLf
    strResult = strResult & "-------------------------------------------" & vbCrLf
    strResult = strResult & "Tamanho Antes  da Compactacao: " & lngBefore & " bytes" & vbCrLf
    strResult = strResult & "Tamanho Depois da Compactacao: " & lngAfter & " bytes"
    MsgBox strResult, vbInformation
  Else
    MsgBox "Erro na compactacao do arquivo", vbExclamation
  End If
  
End Sub

Private Sub CmdExportarSqlite_Click()
   If MDG("Converter Reescreve Destino") Then
      Convert Text1.Text, Text2.Text
   End If
   
End Sub
Private Sub Convert(NWindMDBFileName$, SQLiteFileName$)

Dim aCnn As ADODB.Connection
Dim sCnn As cConnection
Dim cconn As String

  On Error Resume Next
  
  Set aCnn = New ADODB.Connection
  aCnn.CursorLocation = adUseClient
  cconn = GeracArq(NWindMDBFileName)
  aCnn.Open cconn
  
  If Err Then MsgBox Err.Description: Err.Clear: Exit Sub
 
  Err.Clear
  
  If InStr(UCase(SQLiteFileName), ".SQLITE") = 0 Then
     Alert "Destino nao e sql lite"
     Exit Sub
  End If
  
  Set sCnn = New_c.Connection
  sCnn.OpenDB SQLiteFileName
  
  If Err Then MsgBox Err.Description: Err.Clear: Exit Sub
  
  Set C = New cConverter
  C.ConvertDatabase aCnn, sCnn
  
  If Err Then MsgBox Err.Description: Err.Clear: Exit Sub
  lProgress.Caption = "Table-Schemas created, Table-Data transferred!"
  
  C.ConvertIndexes aCnn, sCnn
  
  If Err Then MsgBox Err.Description
  lProgress.Caption = "Index-Import finished!"
  
  Err.Clear
  
  If Err Then MsgBox Err.Description: Err.Clear

  Set C = Nothing
End Sub

Private Sub CmdTeste_Click()
listTables Text1.Text
Listindex Text1.Text, "RPT"
ListFields Text1.Text, "RPT"
End Sub

Private Sub Command1_Click()
  Dim sFileName As String
  Dim sPath As String
  Dim sRECENTFILE As String
  Dim sFILTER As String

  sPath = App.Path
  sFILTER = "Access mdb" & vbNullChar & "*.mdb" & vbNullChar & "DBF" & vbNullChar & "*.DBF"
  sFileName = FileOpen(Me, sFILTER, 1, sRECENTFILE, "ini", sPath, "Open Ini File")
  If Len(sFileName) = 0 Then
    Text1.Text = ""
    Exit Sub
  End If
  Text1.Text = sFileName
End Sub
Private Sub Command2_Click()
  Dim sFileName As String
  Dim sPath As String
  Dim sRECENTFILE As String
  Dim sFILTER As String


  sPath = App.Path
  sFILTER = "Access mdb" & vbNullChar & "*.mdb" & vbNullChar & "Sqlite" & vbNullChar & "*.sqlite"
  sFileName = FileOpen(Me, sFILTER, 1, sRECENTFILE, "ini", sPath, "Open Ini File")
  If Len(sFileName) = 0 Then
    Text2.Text = ""
    Exit Sub
  End If
  Text2.Text = sFileName
  
End Sub
Private Sub Command3_Click()
  If Len(Text1.Text) = 0 Or Len(Text2.Text) = 0 Then
    Alert "Origem ou Destino Nao Preenchido"
    Exit Sub
  End If
  corrige Text1.Text, Text2.Text, False
End Sub
Private Sub Command4_Click()
  If Len(Text1.Text) = 0 Or Len(Text2.Text) = 0 Then
    Alert "Origem ou Destino Nao Preenchido"
    Exit Sub
  End If
  corrige Text1.Text, Text2.Text, True
End Sub
Private Sub corrige(ByVal cORIGEM As String, ByVal cDESTINO As String, Optional ByVal lGRAVA As Boolean = False)
  Dim WrkSpace As DAO.Workspace
  Dim Baza1 As DAO.Database
  Dim Baza2 As DAO.Database
  Dim rec1 As DAO.Recordset
  Dim rec2 As DAO.Recordset
  Dim NewTable As DAO.TableDef
  Dim fldTemp As DAO.Field
  Dim N  As Integer
  Dim m As Integer
  Dim T_Ime As Variant
  Dim R_Ime As Variant
  Dim nekej As Variant
  Dim r_type As Variant
  Dim R_size As Variant
  Dim i As Integer
  Dim x As Integer

  On Error GoTo errhandler
  Set WrkSpace = DBEngine.CreateWorkspace("Compare", "Admin", "")
  If Not FileExists(Text1.Text) Then
     Alert ("Falta arquivo " + Text1.Text)
     Exit Sub
  End If
  If Not FileExists(Text2.Text) Then
     Alert ("Falta arquivo " + Text2.Text)
     Exit Sub
  End If
  Set Baza1 = WrkSpace.OpenDatabase(Text1.Text)
  Set Baza2 = WrkSpace.OpenDatabase(Text2.Text)
  ''On Error Resume Next
  Text3.Text = ""
  For N = 0 To Baza1.TableDefs.Count - 1
    If Baza1.TableDefs(N).Properties(5) = 0 Then
      nERRO = 0
      DoEvents
      T_Ime = Baza1.TableDefs(N).Name
      Set rec1 = Baza1.OpenRecordset(T_Ime)
      Set rec2 = Baza2.OpenRecordset(T_Ime)
      If nERRO = 3078 Then
        Text3.Text = Text3.Text & "Falta Tabela: " & T_Ime & vbNewLine
        If lGRAVA Then
          Set NewTable = Baza2.CreateTableDef(T_Ime)
          With NewTable
            For m = 0 To rec1.Fields.Count - 1
              R_Ime = rec1.Fields(m).Name
              r_type = rec1.Fields(m).Type
              R_size = rec1.Fields(m).Size
              Set fldTemp = NewTable.CreateField(R_Ime, r_type, R_size)
              If r_type = dbText Or r_type = dbMemo Then
                fldTemp.AllowZeroLength = True
              End If
              .Fields.Append fldTemp
            Next m
          End With
          Baza2.TableDefs.Append NewTable
        End If
      Else
        For m = 0 To rec1.Fields.Count - 1
          nERRO = 0
          R_Ime = rec1.Fields(m).Name
          nekej = rec2.Fields(R_Ime).Name
          If nERRO = 3265 Then
            Text3.Text = Text3.Text & "Falta Campo: " & T_Ime & "." & R_Ime & vbNewLine
            If lGRAVA Then
              Set NewTable = Baza2.TableDefs(T_Ime)
              rec2.Close
              With NewTable
                R_Ime = rec1.Fields(m).Name
                r_type = rec1.Fields(m).Type
                R_size = rec1.Fields(m).Size
                Set fldTemp = NewTable.CreateField(R_Ime, r_type, R_size)
                If r_type = dbText Or r_type = dbMemo Then
                  fldTemp.AllowZeroLength = True
                End If
                .Fields.Append fldTemp
              End With
              Set rec2 = Baza2.OpenRecordset(T_Ime)
            End If
          End If
        Next m
      End If
    End If
  Next N
  If lGRAVA Then
    rec2.Close
    rec1.Close
  End If
  Baza2.Close
  Baza1.Close
  WrkSpace.Close
  If Text3.Text = "" Then
    Text3.Text = "Arquivos sem diferencas"
  End If

  Exit Sub
errhandler:
  Select Case Err.Number
  Case 3422, 3078, 3265
    nERRO = FixInt(Err.Number)
    Resume Next
  Case 3219, 91
    Resume Next
  Case Else
    SayErro "Dcompare"
    Exit Sub
  End Select
End Sub

Private Sub Command5_Click()
  Dim sFILTER  As String
  Dim sFileName As String
  sFILTER = "Arquivos Access (*.MDB)" & vbNullChar & "*.MDB" & vbNullChar & "Todos Arquivo" & vbNullChar & "*.*"
  sFileName = FileSave(Me, sFILTER, 1, "MDB", "", "", "Criar Access Como")
  If Not Len(sFileName) = 0 Then
    If FileConnExist(sFileName) Then
      ''nAO GRAVA EM CIMA
    Else
      CriaMdbAccess sFileName, False, 5
      Text2.Text = sFileName
    End If
  End If
End Sub

Private Sub Encerrar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
'Center Me
  CenterFormToScreen Me
End Sub


Private Sub listTables(dbNameWithPath As String)
Dim dbConn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cTABLE As String
    Set dbConn = New ADODB.Connection
    
    dbConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" _
    & dbNameWithPath & ";" _
    & "Persist Security Info=False"
    dbConn.Open

   ' lstTables.Clear
   ' lstFields.Clear
    
    ' Use OpenSchema and get the table names.
    ' The final argument in the parameter array
    ' is "Table" to indicate we want a list of tables.

    Set rs = dbConn.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "Table"))
    Do Until rs.EOF
       ' Me.lstTables.AddItem (rs!Table_Name)
       cTABLE = rs!Table_Name
       Debug.Print cTABLE
       ListFields dbNameWithPath, cTABLE
       Listindexes dbNameWithPath, cTABLE
        rs.MoveNext
    Loop
    rs.Close
    dbConn.Close
End Sub
Private Sub ListFields(dbFileWithPath As String, dbTableName As String)
Dim dbConn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cTIPO As String
'Dim nLength

    Set dbConn = New ADODB.Connection
    
    dbConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" _
    & dbFileWithPath & ";" _
    & "Persist Security Info=False"
    dbConn.Open

    Set rs = dbConn.OpenSchema(adSchemaColumns, Array(Empty, Empty, dbTableName))
    
    Do Until rs.EOF
        Debug.Print
        Debug.Print "Table         :" & rs("table_name")
        Debug.Print "Campo         :" & rs("COLUMN_NAME")
        cTIPO = TipoDado2(rs("DATA_TYPE"))
        Debug.Print "Tipo          :"; TipoDado2(rs("DATA_TYPE")) & " -  " & rs("DATA_TYPE") & " - " & TipoDados(rs("DATA_TYPE"))
        If cTIPO = "S" Then
           Debug.Print "Tamanho Char  :" & FixInt(rs("CHARACTER_MAXIMUM_LENGTH"))
        End If
        If cTIPO = "N" Then
           Debug.Print "Tamanho Numero:" & FixInt(rs("NUMERIC_PRECISION"))
           Debug.Print "Scala Numero:" & FixInt(rs("NUMERIC_SCALE"))
        End If
        If cTIPO = "D" And FixInt(rs("DATETIME_PRECISION")) > 0 Then
          Debug.Print "Tamanho DATA:" & FixInt(rs("DATETIME_PRECISION"))
        End If
        rs.MoveNext
    Loop
    rs.Close
    dbConn.Close
End Sub


Private Sub Listindexes(dbFileWithPath As String, dbTableName As String)
Dim dbConn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cTIPO As String
'Dim nLength

    Set dbConn = New ADODB.Connection
    
    dbConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" _
    & dbFileWithPath & ";" _
    & "Persist Security Info=False"
    dbConn.Open

    Set rs = dbConn.OpenSchema(adSchemaIndexes, Array(Empty, Empty, Empty, Empty, dbTableName))
    
    Do Until rs.EOF
        Debug.Print
        '!TABLE_NAME, !INDEX_NAME, !PRIMARY_KEY
        Debug.Print "Table         :" & rs("table_name")
        Debug.Print "Indice        :" & rs("INDEX_NAME")
        Debug.Print "Primaria      :" & rs("PRIMARY_KEY")
        Debug.Print "Chave         :" & rs("COLUMN_NAME")
        rs.MoveNext
    Loop
    rs.Close
    dbConn.Close
End Sub


Private Sub Listindex(dbFileWithPath As String, dbTableName As String)
Dim dbConn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cINDEX   As String
Dim cCOLUNAS As String
    
    
    Set dbConn = New ADODB.Connection
    
    dbConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" _
    & dbFileWithPath & ";" _
    & "Persist Security Info=False"
    dbConn.Open

    Set rs = dbConn.OpenSchema(adSchemaIndexes) ', Array(Empty, Empty, dbTableName))
    
    cINDEX = ""
    cCOLUNAS = ""
    Do Until rs.EOF
       If UCase(rs.Fields!Table_Name) = UCase(dbTableName) Then
        Debug.Print
        Debug.Print "Table         :" & rs("table_name")
        Debug.Print "Indice        :" & rs("INDEX_NAME")
        If cINDEX = rs("INDEX_NAME") Then
           cCOLUNAS = cCOLUNAS & "," & rs("COLUMN_NAME")
        End If
        If cINDEX = "" Then
           cINDEX = rs("INDEX_NAME")
           cCOLUNAS = rs("COLUMN_NAME")
        End If
        If cINDEX <> rs("INDEX_NAME") Then
           Debug.Print cCOLUNAS
           cINDEX = rs("INDEX_NAME")
           cCOLUNAS = rs("COLUMN_NAME")
        End If
       End If
       rs.MoveNext
    Loop
    Debug.Print cCOLUNAS
    rs.Close
    dbConn.Close
End Sub



