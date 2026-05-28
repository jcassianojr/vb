VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form dCompare 
   Caption         =   "Database Comparador/Corretor"
   ClientHeight    =   5808
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   10356
   Icon            =   "dcompare.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5808
   ScaleWidth      =   10356
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1080
      TabIndex        =   14
      Top             =   600
      Width           =   7215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1080
      TabIndex        =   13
      Top             =   120
      Width           =   7215
   End
   Begin VB.TextBox Text3 
      Height          =   3285
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Top             =   2280
      Width           =   8055
   End
   Begin VB.CommandButton cmdCompactar 
      Caption         =   "Compactar"
      Height          =   255
      Left            =   2400
      TabIndex        =   11
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Criar Tabela de Destino"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   2055
   End
   Begin VB.CommandButton Encerrar 
      Caption         =   "Encerrar"
      Height          =   375
      Left            =   8520
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton CmdTeste 
      Caption         =   "TableInfo"
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton CmdExportarSqlite 
      Caption         =   "Exportar Sqlite"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Comparar e Corrigir"
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Compare Arquivos"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9240
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Destino"
      Height          =   255
      Left            =   8520
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Origem"
      Height          =   255
      Left            =   8520
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lProgress 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4920
      TabIndex        =   2
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
    cconn = GeraConexao(NWindMDBFileName)
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
    
    
'   2. Migra apenas as Tabelas, Campos e Tipos (sem os dados)
    'C.ConvertSchema aCnn, sCnn
    
    'ConvertQuery (Migração baseada em SQL)

'Este é um dos mais poderosos: ele permite que você execute uma consulta SQL na origem e converta o resultado dessa consulta em uma tabela no destino.

 '   Exemplo: Você pode fazer um SELECT * FROM Pedidos WHERE Data > '2025-01-01'
    
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
Private Sub Command5_Click()
Dim cARQUIVONOVO As String

On Error GoTo Erro_Cancel
    
    With CommonDialog1
        .CancelError = True
        ' Filtro para os formatos de salvamento do Access
        .Filter = "Banco de Dados Access (*.mdb)|*.mdb|Banco de Dados Access (*.accdb)|*.accdb"
        .DialogTitle = "Selecione o local para SALVAR o Banco de Dados"
        .InitDir = App.Path
        
        ' Avisa o usuário se o arquivo já existir antes de sobrescrever
        .Flags = cdlOFNOverwritePrompt
        
        .ShowSave
        
        ' Joga o caminho para o seu TextBox de destino
        cARQUIVONOVO = .FileName
        If Not ArquivoExiste(cARQUIVONOVO) Then
           CriaMdbAccess cARQUIVONOVO, False, 5
        End If
        
        
    End With
Exit Sub

Erro_Cancel:
    ' Entra aqui se o usuário clicar em "Cancelar"
    If Err.Number <> cdlCancel Then
        MsgBox "Erro ao selecionar destino: " & Err.Description, vbCritical
    End If

End Sub

Private Sub Encerrar_Click()
  Unload Me
End Sub

Private Sub Form_Load()
'Center Me
  CenterFormToScreen Me
  
  
  
   'Alert (infosistema)
  '
  'SayErro "teste", True
End Sub


Private Sub listTables(dbNameWithPath As String)
Dim dbConn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim cTable As String
    Set dbConn = New ADODB.Connection
    dbConn.ConnectionString = GeraConexao(dbNameWithPath)
    
   ' dbConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" _
   ' & dbNameWithPath & ";" _
   ' & "Persist Security Info=False"
    dbConn.Open

   ' lstTables.Clear
   ' lstFields.Clear
    
    ' Use OpenSchema and get the table names.
    ' The final argument in the parameter array
    ' is "Table" to indicate we want a list of tables.

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
  If Not IsNull(eVAR) Then
     FixStr = CStr(eVAR)
  End If
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
Public Function funNumeroPuro(ByVal pNumero) As String
  Dim i As Integer
  pNumero = FixStr(pNumero)
  funNumeroPuro = ""
  For i = 1 To Len(pNumero)
    If InStr("0123456789", Mid(pNumero, i, 1)) > 0 Then
      funNumeroPuro = funNumeroPuro & Mid(pNumero, i, 1)
    End If
  Next
End Function
Public Function FixInt(ByVal cUSO As Variant, Optional ByVal ePAD As Variant = 0)
  If IsNull(ePAD) Or ePAD = "" Then            'necessario pois pode ser passado outros tipos
    ePAD = 0
  End If
  If IsNull(cUSO) Or cUSO = "" Then
    FixInt = ePAD
    Exit Function
  End If
  If Not IsNumeric(cUSO) Then
    cUSO = funNumeroPuro(cUSO)
  End If
  If IsNull(cUSO) Or cUSO = "" Then            'necessario apos funnumeropuro
    FixInt = ePAD
    Exit Function
  End If
  FixInt = Int(cUSO)
End Function

' Função auxiliar para definir o Provider correto
Private Function MontaStringConexao(cPath As String) As String
    If InStr(LCase(cPath), ".accdb") > 0 Then
        MontaStringConexao = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & cPath
    Else
        MontaStringConexao = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & cPath
    End If
End Function


''' <summary>
''' Compara e sincroniza dois bancos de dados.
''' </summary>
''' <param name="cPathOri">Caminho do banco Modelo (Origem)</param>
''' <param name="cPathDest">Caminho do banco a ser Corrigido (Destino)</param>
''' <param name="lGrava">Se True, aplica as mudanças. Se False, apenas gera o log.</param>
Public Sub corrige(ByVal cPathOri As String, ByVal cPathDest As String, ByVal lGrava As Boolean)
    Dim catOri As Object
    Dim catDest As Object
    Dim tblOri As Object
    Dim tblDest As Object
    Dim colOri As Object
    Dim idxOri As Object
    Dim idxDest As Object
    Dim fldIdx As Object
    
    Dim strConnOri As String
    Dim strConnDest As String
    Dim existeTabela As Boolean
    Dim existeCampo As Boolean
    Dim existeIdx As Boolean

    On Error GoTo err_Geral

    ' Limpa o log e avisa o início
    Text3.Text = ""
    lProgress.Caption = "Analisando estruturas..."
    DoEvents

    ' Validação básica dos parâmetros recebidos
    If cPathOri = "" Or cPathDest = "" Then
        MsgBox "Caminhos de origem ou destino inválidos!", vbExclamation
        Exit Sub
    End If

    ' 1. Monta as strings de conexão dinamicamente
    strConnOri = MontaStringConexao(cPathOri)
    strConnDest = MontaStringConexao(cPathDest)

    Set catOri = CreateObject("ADOX.Catalog")
    Set catDest = CreateObject("ADOX.Catalog")

    ' Abre as conexões
    catOri.ActiveConnection = strConnOri
    catDest.ActiveConnection = strConnDest

    ' 2. Loop pelas tabelas da ORIGEM
    For Each tblOri In catOri.Tables
        ' Filtra apenas tabelas reais (ignora views, links e tabelas de sistema)
        If UCase(tblOri.Type) = "TABLE" Then
            lProgress.Caption = "Tabela: " & tblOri.Name
            DoEvents
            
            existeTabela = False
            On Error Resume Next
            Set tblDest = catDest.Tables(tblOri.Name)
            If Err.Number = 0 Then existeTabela = True
            On Error GoTo err_Geral

            If Not existeTabela Then
                Text3.Text = Text3.Text & " -> Tabela faltante: " & tblOri.Name & vbCrLf
                
                If lGrava Then
                    Set tblDest = CreateObject("ADOX.Table")
                    tblDest.Name = tblOri.Name
                    ' Copia colunas para a nova tabela
                    For Each colOri In tblOri.Columns
                        tblDest.Columns.Append colOri.Name, colOri.Type, colOri.DefinedSize
                    Next
                    catDest.Tables.Append tblDest
                    Text3.Text = Text3.Text & "    [OK] Tabela criada." & vbCrLf
                End If
            Else
                ' --- VERIFICAÇÃO DE CAMPOS ---
                For Each colOri In tblOri.Columns
                    existeCampo = False
                    On Error Resume Next
                    Dim vCol As Object
                    Set vCol = tblDest.Columns(colOri.Name)
                    If Err.Number = 0 Then existeCampo = True
                    On Error GoTo err_Geral

                    If Not existeCampo Then
                        Text3.Text = Text3.Text & " -> [" & tblOri.Name & "] Campo faltante: " & colOri.Name & vbCrLf
                        If lGrava Then
                            tblDest.Columns.Append colOri.Name, colOri.Type, colOri.DefinedSize
                            Text3.Text = Text3.Text & "    [OK] Campo adicionado." & vbCrLf
                        End If
                    End If
                Next

                ' --- VERIFICAÇÃO DE ÍNDICES ---
                For Each idxOri In tblOri.Indexes
                    existeIdx = False
                    On Error Resume Next
                    Set idxDest = tblDest.Indexes(idxOri.Name)
                    If Err.Number = 0 Then existeIdx = True
                    On Error GoTo err_Geral

                    If Not existeIdx Then
                        Text3.Text = Text3.Text & " -> [" & tblOri.Name & "] Índice faltante: " & idxOri.Name & vbCrLf
                        If lGrava Then
                            Set idxDest = CreateObject("ADOX.Index")
                            idxDest.Name = idxOri.Name
                            idxDest.PrimaryKey = idxOri.PrimaryKey
                            idxDest.Unique = idxOri.Unique
                            
                            For Each fldIdx In idxOri.Columns
                                idxDest.Columns.Append fldIdx.Name
                            Next
                            tblDest.Indexes.Append idxDest
                            Text3.Text = Text3.Text & "    [OK] Índice criado." & vbCrLf
                        End If
                    End If
                Next
            End If
        End If
    Next

    lProgress.Caption = "Finalizado."
    MsgBox "Processo concluído com sucesso!", vbInformation
    Exit Sub

err_Geral:
    MsgBox "Erro ao processar: " & Err.Description, vbCritical
    Set catOri = Nothing
    Set catDest = Nothing
End Sub
