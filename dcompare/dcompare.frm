VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.2#0"; "vbccr18.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form dCompare 
   Caption         =   "Database Comparador/Corretor"
   ClientHeight    =   5808
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   8952
   Icon            =   "dcompare.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5808
   ScaleWidth      =   8952
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8400
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VBCCR18.CommandButtonW CmdTeste 
      Height          =   375
      Left            =   6240
      TabIndex        =   14
      Top             =   1200
      Width           =   975
      _ExtentX        =   1715
      _ExtentY        =   656
      Caption         =   "TableInfo"
   End
   Begin XPControls.XPButton CmdExportarSqlite 
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
      _ExtentX        =   2138
      _ExtentY        =   656
      Caption         =   "Exportar Sqlite"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      _ExtentX        =   15050
      _ExtentY        =   5101
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
      _ExtentX        =   13145
      _ExtentY        =   656
   End
   Begin VBCCR18.TextBoxW Text1 
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   120
      Width           =   7455
      _ExtentX        =   13145
      _ExtentY        =   656
   End
   Begin XPControls.XPButton Command3 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2295
      _ExtentX        =   4043
      _ExtentY        =   656
      Caption         =   "Compare Arquivos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      _ExtentX        =   656
      _ExtentY        =   656
      Picture         =   "dcompare.frx":05B0
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      _ExtentX        =   656
      _ExtentY        =   656
      Picture         =   "dcompare.frx":0B4A
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      _ExtentX        =   4043
      _ExtentY        =   656
      Caption         =   "Comparar e Corrigir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      _ExtentX        =   4043
      _ExtentY        =   656
      Picture         =   "dcompare.frx":10E4
      Caption         =   "Criar Tabela de Destino"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      _ExtentX        =   2350
      _ExtentY        =   656
      Picture         =   "dcompare.frx":167E
      Caption         =   "&Encerrar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
      _ExtentX        =   4043
      _ExtentY        =   656
      Picture         =   "dcompare.frx":1C18
      Caption         =   "Compactar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
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
         Size            =   7.8
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
         Size            =   7.8
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
' Refatoração do método Convert para garantir fechamento de conexões [cite: 49, 50]
Private Sub Convert(NWindMDBFileName$, SQLiteFileName$)
    Dim aCnn As ADODB.Connection
    Dim sCnn As cConnection
    Dim cconn As String

    On Error GoTo errhandler ' Substitui On Error Resume Next [cite: 49]
    
    lProgress.Caption = "Iniciando conexão..."
    Set aCnn = New ADODB.Connection
    aCnn.CursorLocation = adUseClient
    cconn = GeracArq(NWindMDBFileName)
    aCnn.Open cconn
    
    If InStr(UCase(SQLiteFileName), ".SQLITE") = 0 Then
        MsgBox "Destino não é um arquivo SQLite válido.", vbExclamation
        GoTo Cleanup
    End If
    
    Set sCnn = New_c.Connection
    sCnn.OpenDB SQLiteFileName
    
    Set C = New cConverter
    lProgress.Caption = "Transferindo dados das tabelas..."
    C.ConvertDatabase aCnn, sCnn
    
    lProgress.Caption = "Criando índices..."
    C.ConvertIndexes aCnn, sCnn
    
    lProgress.Caption = "Conversão concluída com sucesso!"
    MsgBox "Processo finalizado!", vbInformation

Cleanup:
    ' Garante que os objetos sejam liberados mesmo em erro [cite: 51]
    If Not aCnn Is Nothing Then If aCnn.State = adStateOpen Then aCnn.Close
    Set aCnn = Nothing
    Set sCnn = Nothing
    Set C = Nothing
    Exit Sub

errhandler:
    MsgBox "Erro na conversão: " & Err.Description, vbCritical
    Resume Cleanup
End Sub

Private Sub CmdTeste_Click()
listTables Text1.Text
Listindex Text1.Text, "RPT"
ListFields Text1.Text, "RPT"
End Sub

Private Sub Command1_Click()
 On Error GoTo Erro_Cancel
    
    With CommonDialog1
        .CancelError = True
        ' Filtro para aceitar os dois tipos de banco Access
        .Filter = "Banco de Dados Access (*.mdb;*.accdb)|*.mdb;*.accdb|Todos os Arquivos (*.*)|*.*"
        .DialogTitle = "Selecione o Banco de Dados de ORIGEM"
        .InitDir = App.Path
        .ShowOpen
        
        ' Joga o caminho para o seu TextBox de origem
        Text1.Text = .FileName
    End With
    
    Exit Sub

Erro_Cancel:
    ' Se o usuário clicar em cancelar, não faz nada

End Sub
Private Sub Command2_Click()
On Error GoTo Erro_Cancel
    
    With CommonDialog1
        .CancelError = True
        .Filter = "Banco de Dados Access (*.mdb;*.accdb)|*.mdb;*.accdb"
        .DialogTitle = "Selecione o Banco de Dados de DESTINO"
        .ShowOpen
        
        ' Joga o caminho para o seu TextBox de destino
        Text2.Text = .FileName
    End With
    
    Exit Sub

Erro_Cancel:
  
End Sub
Private Sub Command3_Click()
 If Text1.Text = "" Or Text2.Text = "" Then
        MsgBox "Selecione os bancos de origem e destino.", vbExclamation
        Exit Sub
    End If

    ' Chama a rotina principal passando FALSE para não executar alterações
    migra_corrige Text1.Text, Text2.Text, False
    
    MsgBox "Comparação finalizada! Verifique o log de diferenças.", vbInformation
End Sub
Private Sub Command4_Click()
 Dim Resposta As VbMsgBoxResult
    
    If Text1.Text = "" Or Text2.Text = "" Then
        MsgBox "Selecione os bancos de origem e destino.", vbExclamation
        Exit Sub
    End If

    Resposta = MsgBox("Deseja aplicar as correções no banco de destino agora?", _
                      vbQuestion + vbYesNo, "Confirmar Alteração")
    
    If Resposta = vbYes Then
        ' Chama a rotina principal passando TRUE para executar as alterações
        migra_corrige Text1.Text, Text2.Text, True
        MsgBox "Correções aplicadas com sucesso!", vbInformation
    End If
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
  CenterFormToScreen Me
 
End Sub


Private Sub listTables(dbNameWithPath As String)
Dim dbConn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cTable As String
    Set dbConn = New ADODB.Connection
    
    dbConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" _
    & dbNameWithPath & ";" _
    & "Persist Security Info=False"
    dbConn.Open


    Set rs = dbConn.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "Table"))
    Do Until rs.EOF
       ' Me.lstTables.AddItem (rs!Table_Name)
       cTable = rs!Table_Name
       Debug.Print cTable
       ListFields dbNameWithPath, cTable
       Listindexes dbNameWithPath, cTable
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


Public Sub migra_corrige(ByVal sOrigem As String, ByVal sDestino As String, ByVal bExecutar As Boolean)
    Dim cnOri As New ADODB.Connection
    Dim cnDes As New ADODB.Connection
    Dim rsTabs As ADODB.Recordset
    Dim colSql As New Collection
    Dim sRelatorio As String
    Dim vCmd As Variant
    
    On Error GoTo Erro_Geral
    Screen.MousePointer = vbHourglass

    ' 1. Conexão Automática (MDB ou ACCDB)
    cnOri.Open GetStrConn(sOrigem)
    cnDes.Open GetStrConn(sDestino)

    ' 2. ANALISAR TABELAS
    Set rsTabs = cnOri.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))

    Do While Not rsTabs.EOF
        Dim sTab As String: sTab = rsTabs("TABLE_NAME")
        
        If Left(sTab, 4) <> "MSys" Then
            ' A. Checar Tabela
            If Not ExisteNoSchema(cnDes, adSchemaTables, sTab) Then
                colSql.Add GerarSqlNovaTabela(cnOri, sTab)
            End If

            ' B. Checar Campos
            Dim rsCols As ADODB.Recordset
            Set rsCols = cnOri.OpenSchema(adSchemaColumns, Array(Empty, Empty, sTab))
            Do While Not rsCols.EOF
                Dim sCol As String: sCol = rsCols("COLUMN_NAME")
                If Not ExisteNoSchema(cnDes, adSchemaColumns, sTab, sCol) Then
                    Dim sDef As String
                    sDef = TraduzirTipoADO(rsCols("DATA_TYPE"), rsCols("CHARACTER_MAXIMUM_LENGTH"))
                    colSql.Add "ALTER TABLE [" & sTab & "] ADD COLUMN [" & sCol & "] " & sDef
                End If
                rsCols.MoveNext
            Loop
            rsCols.Close

            ' C. Checar Índices (Via Schema de Índices - Mais estável)
            ChecarIndicesADO cnOri, cnDes, sTab, colSql
        End If
        rsTabs.MoveNext
    Loop

    cnOri.Close: Set cnOri = Nothing
    cnDes.Close: Set cnDes = Nothing

    ' 3. EXECUTAR
    If colSql.Count > 0 Then
        If bExecutar Then
            cnDes.Open GetStrConn(sDestino)
            For Each vCmd In colSql
                On Error Resume Next
                cnDes.Execute CStr(vCmd)
                sRelatorio = sRelatorio & IIf(Err.Number = 0, "OK: ", "ERRO: ") & vCmd & vbCrLf
                Err.Clear
            Next
            cnDes.Close
        Else
            For Each vCmd In colSql: sRelatorio = sRelatorio & "PENDENTE: " & vCmd & vbCrLf: Next
        End If
    End If

    Screen.MousePointer = vbDefault
    dCompare.Text3.Text = IIf(sRelatorio = "", "Bancos sincronizados.", sRelatorio)
    MsgBox "Concluído!", vbInformation
    Exit Sub

Erro_Geral:
    Screen.MousePointer = vbDefault
    MsgBox "Erro: " & Err.Description, vbCritical
End Sub

' --- AUXILIARES ---

Private Sub ChecarIndicesADO(cnO As ADODB.Connection, cnD As ADODB.Connection, sTab As String, ByRef col As Collection)
    Dim rsIdxO As ADODB.Recordset, rsIdxD As ADODB.Recordset
    On Error Resume Next
    
    ' Obtém índices da origem
    Set rsIdxO = cnO.OpenSchema(adSchemaIndexes, Array(Empty, Empty, Empty, Empty, sTab))
    
    Do While Not rsIdxO.EOF
        Dim sIdxNome As String: sIdxNome = rsIdxO("INDEX_NAME")
        
        ' Ignora índices automáticos de chaves estrangeiras
        If Left(sIdxNome, 2) <> "Reference" And Left(sIdxNome, 4) <> "MSys" Then
            ' Verifica se existe no destino
            Set rsIdxD = cnD.OpenSchema(adSchemaIndexes, Array(Empty, Empty, sIdxNome, Empty, sTab))
            
            If rsIdxD.EOF Then
                ' Monta o SQL (Simplificado: assume que o índice é sobre a coluna atual do loop do schema)
                Dim sCol As String: sCol = rsIdxO("COLUMN_NAME")
                Dim sUniq As String: sUniq = IIf(rsIdxO("UNIQUE") = True, "UNIQUE ", "")
                col.Add "CREATE " & sUniq & "INDEX [" & sIdxNome & "] ON [" & sTab & "] ([" & sCol & "])"
            End If
            rsIdxD.Close
        End If
        rsIdxO.MoveNext
    Loop
    rsIdxO.Close
End Sub

Private Function GetStrConn(ByVal sP As String) As String
    If LCase(Right(sP, 5)) = "accdb" Then
        GetStrConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & sP
    Else
        GetStrConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sP
    End If
End Function

Private Function ExisteNoSchema(cn As ADODB.Connection, nT As Long, ST As String, Optional sC As String = "") As Boolean
    Dim rs As ADODB.Recordset
    On Error Resume Next
    If sC = "" Then
        Set rs = cn.OpenSchema(nT, Array(Empty, Empty, ST, "TABLE"))
    Else
        Set rs = cn.OpenSchema(nT, Array(Empty, Empty, ST, sC))
    End If
    ExisteNoSchema = Not rs.EOF
    rs.Close
End Function

Private Function TraduzirTipoADO(nType As Long, nLEN As Variant) As String
    Select Case nType
        Case 11: TraduzirTipoADO = "BIT"
        Case 3:  TraduzirTipoADO = "LONG"
        Case 2:  TraduzirTipoADO = "SHORT"
        Case 6:  TraduzirTipoADO = "CURRENCY"
        Case 5:  TraduzirTipoADO = "DOUBLE"
        Case 7, 135: TraduzirTipoADO = "DATETIME"
        Case 201, 203: TraduzirTipoADO = "MEMO"
        Case Else: TraduzirTipoADO = "TEXT(" & IIf(IsNull(nLEN) Or nLEN <= 0, 255, nLEN) & ")"
    End Select
End Function

Private Function GerarSqlNovaTabela(cn As ADODB.Connection, sTab As String) As String
    Dim rs As ADODB.Recordset
    Set rs = cn.OpenSchema(adSchemaColumns, Array(Empty, Empty, sTab))
    If Not rs.EOF Then
        GerarSqlNovaTabela = "CREATE TABLE [" & sTab & "] ([" & rs("COLUMN_NAME") & "] " & _
                             TraduzirTipoADO(rs("DATA_TYPE"), rs("CHARACTER_MAXIMUM_LENGTH")) & ")"
    End If
    rs.Close
End Function
