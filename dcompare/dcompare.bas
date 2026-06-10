Attribute VB_Name = "dcompareMain"
Public lRETU As Boolean
' Bibliotecas necessárias no seu VBP:
' 1. Microsoft ADO Ext. 6.0 for DDL and Security (ADOX)
' 2. Microsoft ActiveX Data Objects 6.1 (ADO)
' 3. RC6 (vbRichClient6)
' Declaração para a Shell API no topo do módulo

Sub Main()
  dCompare.Show
End Sub

Public Function ArquivoExiste(ByVal FilePath As String) As Boolean
    ArquivoExiste = (Dir(FilePath, vbNormal + vbHidden + vbSystem) <> "")
End Function
Public Function PathFileExists(ByVal FilePath As String) As Boolean
    PathFileExists = (Dir(FilePath, vbNormal + vbHidden + vbSystem) <> "")
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
Public Function TrimNull(ByVal sTxt As String) As String
  Dim Arr() As String
  Arr() = Split(sTxt, Chr$(0))
  If UBound(Arr) >= 0 Then
    TrimNull = Arr(0)
  Else
    TrimNull = sTxt
  End If
End Function
Public Function GeraConexao(ByVal Caminho As String) As String
    Dim sProvider As String
    
    ' 1. Verifica se é SQLite (via sua função EArquivoSQLite)
    If EArquivoSQLite(Caminho) Then
        ' Conexão via Driver ODBC para SQLite
        GeraConexao = "Driver={SQLite3 ODBC Driver};Database=" & Caminho
        
    Else
        ' 2. Prioridade para Jet 4.0 para arquivos .mdb (Mais estável para ADOX/OpenSchema em .mdb)
        ' Se for .accdb, é obrigado a usar o ACE 12.0
        If LCase(Right(Caminho, 5)) = "accdb" Then
            sProvider = "Microsoft.ACE.OLEDB.12.0"
        Else
            ' Para .mdb, força o uso do Jet 4.0 que você confirmou ser mais estável
            sProvider = "Microsoft.Jet.OLEDB.4.0"
        End If
        
        GeraConexao = "Provider=" & sProvider & ";Data Source=" & Caminho & ";Persist Security Info=False;"
    End If
End Function

Public Function AbrirCatalog(ByVal sCaminho As String) As Object
    ' Retorna um objeto ADOX.Catalog para ler a estrutura do MDB
    Dim cat As Object
    Set cat = CreateObject("ADOX.Catalog")
    
    ' String de conexão genérica para MDB/ACCDB
    cat.ActiveConnection = GeraConexao(sCaminho)
    Set AbrirCatalog = cat
End Function

Public Sub AdicionarScript(ByVal ssql As String)
    If colScriptsPendentes Is Nothing Then Set colScriptsPendentes = New Collection
    colScriptsPendentes.Add ssql
    Debug.Print "Script Adicionado: " & ssql
End Sub

Public Function ValidarOuCriarDestino(ByVal cCaminho As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' 1. Checagem: O arquivo já existe?
    If fso.FileExists(cCaminho) Then
        ' Se existe, perguntamos se o usuário quer APENAS abrir para comparar
        ' ou se ele tem certeza que é o destino correto.
        Dim resposta As VbMsgBoxResult
        resposta = MsgBox("O arquivo de destino já existe." & vbCrLf & _
                          "Deseja utilizá-lo para sincronização?" & vbCrLf & vbCrLf & _
                          "Atenção: A estrutura pode ser alterada.", vbYesNo + vbInformation, "Confirmação")
        
        If resposta = vbYes Then
            ValidarOuCriarDestino = True
        Else
            ValidarOuCriarDestino = False
        End If
    Else
        ' 2. Se não existe, perguntamos se quer criar
        If MsgBox("O arquivo de destino não existe. Deseja criá-lo agora?", vbYesNo + vbQuestion) = vbYes Then
            ' Chama a rotina de criação que vimos anteriormente
            ValidarOuCriarDestino = CriarNovoDestino(cCaminho)
        Else
            ValidarOuCriarDestino = False
        End If
    End If
    
    Set fso = Nothing
End Function
' Adicione isto ao seu modDCompare.bas
Public Function CriarNovoDestino(ByVal cCaminho As String) As Boolean
    Dim lExiste As Boolean
    lExiste = ArquivoExiste(cCaminho)
    
    ' 1. Verificação de Segurança: Se existe e tem dados, abortar para não sobrescrever
    If lExiste Then
        If FileLen(cCaminho) > 0 Then
            MsgBox "Atenção: O arquivo já existe e contém dados. Não é possível sobrescrever.", vbCritical
            CriarNovoDestino = False
            Exit Function
        End If
    End If
    ' 2. Confirmação
    If MsgBox("O arquivo de destino será criado agora. Prosseguir?", vbYesNo + vbQuestion) = vbNo Then
        CriarNovoDestino = False
        Exit Function
    End If
    
    ' 3. Se existe (arquivo vazio de 0 bytes) ou não existe, vamos criar
    If lExiste Then Kill cCaminho
    
    Dim sEXT As String
    sEXT = LCase(cCaminho)
    
    If EArquivoSQLite(sEXT) Then
        ' Cria via motor RC6
        CriarNovoDestino = CriarSqlite(cCaminho)
    Else
        ' Cria via ADOX (MDB ou ACCDB)
        Dim nTIPO As Integer
        nTIPO = IIf(InStr(sEXT, ".accdb") > 0, 6, 5)
        CriarNovoDestino = CriaMdbAccess(cCaminho, True, nTIPO)
    End If
End Function
Public Function EArquivoSQLite(ByVal cCaminho As String) As Boolean
    
     cCaminho = LCase(cCaminho)
    
    ' Lista de extensões que o seu sistema aceita como SQLite
    If InStr(cCaminho, ".sqlite") > 0 Or _
       InStr(cCaminho, ".sqlite3") > 0 Or _
       InStr(cCaminho, ".db") > 0 Or _
       InStr(cCaminho, ".db3") > 0 Or _
       InStr(cCaminho, ".fossil") > 0 Then
       
       EArquivoSQLite = True
    Else
       EArquivoSQLite = False
    End If
End Function
Public Function MDG(ByVal cMEnSSAGEM As String, Optional cTITULO = "Confirme")
  Dim eRESP As Variant
  MDG = False
  eRESP = MsgBox(cMEnSSAGEM, vbYesNo + vbDefaultButton1, cTITULO)
  If eRESP = vbYes Then
    MDG = True
  End If
End Function
Public Function CriarSqlite(ByVal cCaminho As String) As Boolean
    On Error GoTo Erro
    Dim fso As Object
    Dim f As Object
    
    ' Cria o objeto FileSystemObject padrão
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Se o arquivo já existe, esta função apenas retornará sucesso ou erro.
    ' Se você quiser garantir que ele seja criado do zero, pode adicionar:
    ' If fso.FileExists(cCaminho) Then fso.DeleteFile cCaminho
    
    ' Cria o arquivo de texto (o parâmetro True sobrescreve se já existir)
    Set f = fso.CreateTextFile(cCaminho, True)
    
    ' Fecha o arquivo imediatamente
    f.Close
    
    Set f = Nothing
    Set fso = Nothing
    
    CriarSqlite = True
    Exit Function

Erro:
    MsgBox "Erro ao criar arquivo via FSO: " & Err.Description, vbCritical
    CriarSqlite = False
End Function
Public Function CompactarSQLite(ByVal sCaminho As String) As Boolean
    On Error GoTo Erro
    Dim oConn As New RC6.cConnection
    
    ' 1. Abre o banco de dados
    oConn.OpenDB sCaminho
    
    ' 2. O comando VACUUM reorganiza e compacta o ficheiro
    oConn.Execute "VACUUM"
    
    
    CompactarSQLite = True
    Exit Function
    
Erro:
    MsgBox "Erro ao compactar SQLite: " & Err.Description, vbCritical
    CompactarSQLite = False
End Function

Public Sub GerarInfoTabelasAccess(ByVal sCaminho As String, ByRef txtDestino As TextBox)
    Dim conn As Object
    Dim cat As Object ' ADOX.Catalog
    Dim tbl As Object ' ADOX.Table
    Dim idx As Object ' ADOX.Index
    Dim col As Object ' ADOX.Column
    Dim fld As Object ' ADODB.Field
    Dim rs As Object
    
    ' Conexão via ADO
    Set conn = CreateObject("ADODB.Connection")
    conn.Open GeraConexao(sCaminho)
    
    ' Usamos o ADOX para ler a estrutura (Catalog)
    Set cat = CreateObject("ADOX.Catalog")
    Set cat.ActiveConnection = conn
    
    txtDestino.Text = "Estrutura Access (ADOX): " & sCaminho & vbCrLf & String(60, "-") & vbCrLf
    
    For Each tbl In cat.Tables
        If tbl.Type = "TABLE" Then
            txtDestino.Text = txtDestino.Text & vbCrLf & "TABELA: " & tbl.Name & vbCrLf
            
            ' --- CAMPOS ---
            txtDestino.Text = txtDestino.Text & "  [Campos]:" & vbCrLf
            ' --- Lógica Refinada para Precisão e Escala ---
For Each col In tbl.Columns
    Dim sTipoFormatado As String
    Dim nPrec As Integer, nScale As Integer
    
    ' Obtém o tipo base
    sTipoFormatado = GetTipoNome(col.Type)
    
    ' Inicializa precisão e escala
    nPrec = 0
    nScale = 0
    
    ' Tenta capturar as propriedades de precisão (Numeric/Decimal)
    On Error Resume Next
    nPrec = col.Properties("NumericPrecision").Value
    nScale = col.Properties("NumericScale").Value
    On Error GoTo 0
    
    ' --- APLICANDO A LÓGICA DE EXIBIÇÃO (Igual ao seu Harbour) ---
    Dim sDetalhe As String
    
    ' Caso 1: Se for numérico (Precision > 0)
    If nPrec > 0 Then
        ' Se nScale for 0, mostra apenas (10). Se for > 0, mostra (10,2)
        If nScale > 0 Then
            sDetalhe = "(" & nPrec & "," & nScale & ")"
        Else
            sDetalhe = "(" & nPrec & ")"
        End If
        
    ' Caso 2: Se for Texto/Char (Tratamento conforme sua regra de 250)
    ElseIf col.Type = 129 Or col.Type = 200 Or col.Type = 202 Or col.Type = 203 Then
        Dim nLEN As Long
        nLEN = col.DefinedSize
        If nLEN <= 0 Or nLEN > 250 Then nLEN = 250
        sDetalhe = "(" & nLEN & ")"
        
    ' Caso 3: Campos Lógicos (L)
    ElseIf col.Type = 11 Then
        sDetalhe = "(1)"
        
    ' Caso 4: Outros casos
    Else
        sDetalhe = ""
    End If
    
    ' Saída Final no seu Relatório
    txtDestino.Text = txtDestino.Text & "    - " & col.Name & " [" & sTipoFormatado & "]" & sDetalhe & vbCrLf
Next col
            
            ' --- ÍNDICES E SEUS CAMPOS ---
            txtDestino.Text = txtDestino.Text & "  [Índices]:" & vbCrLf
            For Each idx In tbl.Indexes
                txtDestino.Text = txtDestino.Text & "    * Índice: " & idx.Name & " (Primary: " & idx.PrimaryKey & ")" & vbCrLf
                
                ' Listando os campos que compõem este índice
                For Each col In idx.Columns
                    txtDestino.Text = txtDestino.Text & "      -> Campo: " & col.Name & vbCrLf
                Next
            Next
            txtDestino.Text = txtDestino.Text & String(40, ".") & vbCrLf
        End If
    Next
    
    conn.Close
End Sub
' Função auxiliar para converter o número do tipo de dados em texto legível
Public Function GetTipoNome(ByVal nTIPO As Integer) As String
    Select Case nTIPO
        Case 0: GetTipoNome = "Empty"
        Case 2: GetTipoNome = "SmallInt"
        Case 3: GetTipoNome = "Integer"
        Case 4: GetTipoNome = "Single"
        Case 5: GetTipoNome = "Double"
        Case 6: GetTipoNome = "Currency"
        Case 7: GetTipoNome = "Date"
        Case 8: GetTipoNome = "BSTR"
        Case 9: GetTipoNome = "IDispatch"
        Case 10: GetTipoNome = "Error"
        Case 11: GetTipoNome = "Boolean"
        Case 14: GetTipoNome = "Decimal"
        Case 16: GetTipoNome = "TinyInt"
        Case 17: GetTipoNome = "UnsignedTinyInt"
        Case 18: GetTipoNome = "UnsignedSmallInt"
        Case 19: GetTipoNome = "UnsignedInt"
        Case 20: GetTipoNome = "BigInt"
        Case 21: GetTipoNome = "UnsignedBigInt"
        Case 72: GetTipoNome = "GUID"
        Case 128: GetTipoNome = "Binary"
        Case 129: GetTipoNome = "Char"
        Case 130: GetTipoNome = "WChar"
        Case 131: GetTipoNome = "Numeric"
        Case 132: GetTipoNome = "UserDefined"
        Case 133: GetTipoNome = "DBDate"
        Case 134: GetTipoNome = "DBTime"
        Case 135: GetTipoNome = "DBTimeStamp"
        Case 136: GetTipoNome = "Chapter"
        Case 200: GetTipoNome = "VarChar"
        Case 201: GetTipoNome = "LongVarChar"
        Case 202: GetTipoNome = "VarWChar" ' O famoso 202
        Case 203: GetTipoNome = "LongVarWChar(Memo)"
        Case 204: GetTipoNome = "VarBinary"
        Case 205: GetTipoNome = "LongVarBinary(Ole)"
        Case Else: GetTipoNome = "Tipo:" & nTIPO
    End Select
End Function

Public Sub GerarInfoTabelasSQLite(ByVal sCaminho As String, ByRef txtDestino As TextBox)
    Dim oConn As RC6.cConnection
    Dim oRS As RC6.cRecordset
    Dim oFldRS As RC6.cRecordset
    Dim oIdxRS As RC6.cRecordset
    Dim ssql As String
    Dim sTabela As String
    
    Set oConn = New_c.Connection(sCaminho)
    txtDestino.Text = "Estrutura Completa: " & sCaminho & vbCrLf & String(40, "-") & vbCrLf

    ' 1. Lista tabelas (Com a query que já validamos)
    ssql = "SELECT name FROM sqlite_master WHERE type = 'table' AND SUBSTR(name, 1, 7) <> 'sqlite_'"
    Set oRS = oConn.OpenRecordset(ssql)
    
    Do While Not oRS.EOF
        sTabela = oRS!Name
        txtDestino.Text = txtDestino.Text & vbCrLf & "TABELA: " & sTabela & vbCrLf
        
        ' 2. Campos com Tipos (Usando PRAGMA que agora deve funcionar)
        txtDestino.Text = txtDestino.Text & "  [Campos]:" & vbCrLf
        Set oFldRS = oConn.OpenRecordset("PRAGMA table_info([" & sTabela & "])")
        Do While Not oFldRS.EOF
            ' 'name' e 'type' são colunas padrão do PRAGMA table_info
            txtDestino.Text = txtDestino.Text & "    - " & oFldRS!Name & " (" & oFldRS!Type & ")" & vbCrLf
            oFldRS.MoveNext
        Loop
        
        ' 3. Índices (Usando PRAGMA index_list)
        'txtDestino.Text = txtDestino.Text & "  [Índices]:" & vbCrLf
        'Set oIdxRS = oConn.OpenRecordset("PRAGMA index_list([" & sTabela & "])")
        'Do While Not oIdxRS.EOF
        '    txtDestino.Text = txtDestino.Text & "    * " & oIdxRS!Name & vbCrLf
        '    oIdxRS.MoveNext
        'Loop
        
        ' 3. ÍNDICES
        txtDestino.Text = txtDestino.Text & "  [Índices]:" & vbCrLf
        Set oIdxRS = oConn.OpenRecordset("PRAGMA index_list([" & sTabela & "])")
        
        Do While Not oIdxRS.EOF
            Dim sNomeIndice As String
            sNomeIndice = oIdxRS!Name
            txtDestino.Text = txtDestino.Text & "    * Índice: " & sNomeIndice & vbCrLf
            
            ' Detalhes do Índice (Campos que compõem o índice)
            Dim oIdxInfoRS As RC6.cRecordset
            Set oIdxInfoRS = oConn.OpenRecordset("PRAGMA index_info([" & sNomeIndice & "])")
            
            Do While Not oIdxInfoRS.EOF
                ' oIdxInfoRS!name traz o nome da coluna que está no índice
                txtDestino.Text = txtDestino.Text & "      -> Campo: " & oIdxInfoRS!Name & vbCrLf
                oIdxInfoRS.MoveNext
            Loop
            
            oIdxRS.MoveNext
        Loop
        
        
        oRS.MoveNext
    Loop

    Set oRS = Nothing
    Set oConn = Nothing
End Sub
Public Function TratarValor(ByVal v As Variant) As String
    If IsNull(v) Then
        TratarValor = "NULL"
    ElseIf IsDate(v) Then
        ' Ajuste o formato da data conforme o padrão SQL que você deseja
        TratarValor = "'" & Format$(v, "yyyy-mm-dd hh:nn:ss") & "'"
    ElseIf IsNumeric(v) Then
        TratarValor = Replace(v, ",", ".") ' Garante separador decimal ponto
    Else
        ' Escapa aspas simples duplicando-as (padrão SQL)
        TratarValor = "'" & Replace(v, "'", "''") & "'"
    End If
End Function




Public Sub GerarScriptSQLite(ByVal sCaminho As String, ByVal bIncluirDados As Boolean)
    Dim oConn As New RC6.cConnection, oRsTabelas As RC6.cRecordset
    Dim sSqlPath As String, fNum As Integer
    
    sSqlPath = Left(sCaminho, InStrRev(sCaminho, ".")) & "sql"
    fNum = FreeFile
    Open sSqlPath For Output As #fNum
    
    oConn.OpenDB sCaminho
    
    Print #fNum, "BEGIN TRANSACTION;"
    
    ' 1. Seleciona apenas as tabelas definidas pelo usuário
    Set oRsTabelas = oConn.OpenRecordset("SELECT name, sql FROM sqlite_master WHERE type='table' AND SUBSTR(name, 1, 7) <> 'sqlite_'")
    
    Do While Not oRsTabelas.EOF
        Dim sTabela As String: sTabela = oRsTabelas!Name
        
        ' Grava o comando CREATE TABLE
        Print #fNum, oRsTabelas!SQL & ";"
        
        ' 2. Geração dos ÍNDICES desta tabela específica
        Dim oRsIndices As RC6.cRecordset
        Set oRsIndices = oConn.OpenRecordset("PRAGMA index_list([" & sTabela & "])")
        Do While Not oRsIndices.EOF
            ' Ignora índices automáticos de PRIMARY KEY (que já estão no CREATE TABLE)
            If oRsIndices!origin <> "pk" Then
                Dim oRsIdxSql As RC6.cRecordset
                Set oRsIdxSql = oConn.OpenRecordset("SELECT sql FROM sqlite_master WHERE name='" & oRsIndices!Name & "'")
                If Not oRsIdxSql.EOF Then
                    Print #fNum, oRsIdxSql!SQL & ";"
                End If
            End If
            oRsIndices.MoveNext
        Loop
        
        ' 3. INSERT INTO (Condicional)
        If bIncluirDados Then
            Dim oDados As RC6.cRecordset
            Set oDados = oConn.OpenRecordset("SELECT * FROM [" & sTabela & "]")
            Do While Not oDados.EOF
                Dim sVals As String, i As Long
                sVals = ""
                For i = 0 To oDados.Fields.Count - 1
                    sVals = sVals & TratarValor(oDados.Fields(i).Value) & ", "
                Next
                Print #fNum, "INSERT INTO [" & sTabela & "] VALUES (" & Left(sVals, Len(sVals) - 2) & ");"
                oDados.MoveNext
            Loop
        End If
        
        Print #fNum, "" ' Linha em branco para organização
        oRsTabelas.MoveNext
    Loop
    
    Print #fNum, "COMMIT TRANSACTION;"
    Close #fNum
End Sub

' Esta função traduz o tipo numérico do ADOX para o tipo SQL equivalente
Public Function GetTipoADOX(ByVal adType As Long) As String
    Select Case adType
        ' 3 = adInteger, 20 = adBigInt
        Case 3, 20:          GetTipoADOX = "LONG"
        ' 129 = adChar, 202 = adVarWChar, 200 = adVarChar, 130 = adWChar
        Case 129, 202, 200, 130: GetTipoADOX = "TEXT(255)"
        ' 7 = adDate, 133 = adDBDate, 135 = adDBTimeStamp
        Case 7, 133, 135:    GetTipoADOX = "DATETIME"
        ' 11 = adBoolean
        Case 11:             GetTipoADOX = "BIT"
        ' 6 = adCurrency
        Case 6:              GetTipoADOX = "CURRENCY"
        ' 4 = adSingle, 5 = adDouble
        Case 4, 5:           GetTipoADOX = "DOUBLE"
        ' 203 = adLongVarWChar (equivalente ao MEMO do Access)
        Case 203:            GetTipoADOX = "MEMO"
        ' 128 = adBinary, 204 = adVarBinary
        Case 128, 204:       GetTipoADOX = "LONGBINARY"
        ' Caso não encontre, assume texto
        Case Else:           GetTipoADOX = "TEXT(255)"
    End Select
End Function
Public Sub GerarScriptAccess(ByVal sCaminho As String, ByVal bIncluirDados As Boolean)
    Dim cat As Object, tbl As Object, col As Object, idx As Object, fld As Object
    Dim conn As Object, rs As Object
    Dim sSqlPath As String, fNum As Integer
    
    sSqlPath = Left(sCaminho, InStrRev(sCaminho, ".")) & "sql"
    fNum = FreeFile
    Open sSqlPath For Output As #fNum
    
    Set cat = CreateObject("ADOX.Catalog")
    cat.ActiveConnection = GeraConexao(sCaminho)
    Set conn = cat.ActiveConnection
    
    For Each tbl In cat.Tables
        ' Filtro para ignorar tabelas de sistema (Mesma lógica do seu Harbour)
        If tbl.Type = "TABLE" And Left(tbl.Name, 4) <> "MSys" Then
            
            ' 1. CREATE TABLE
            Print #fNum, "CREATE TABLE [" & tbl.Name & "] ("
            Dim sCols As String: sCols = ""
            For Each col In tbl.Columns
                sCols = sCols & "  [" & col.Name & "] " & GetTipoADOX(col.Type) & "," & vbCrLf
            Next
            Print #fNum, Left(sCols, Len(sCols) - 3) & vbCrLf & ");" & vbCrLf
            
            ' 2. ESTRUTURA DE ÍNDICES (Adaptada do seu gera_doc_base.prg)
            ' Usamos a mesma navegação por índices que você validou funcionar
            If tbl.Indexes.Count > 0 Then
                For Each idx In tbl.Indexes
                    Dim sIdxFields As String: sIdxFields = ""
                    For Each fld In idx.Columns
                        sIdxFields = sIdxFields & "[" & fld.Name & "], "
                    Next
                    sIdxFields = Left(sIdxFields, Len(sIdxFields) - 2)
                    
                    If idx.PrimaryKey Then
                        Print #fNum, "ALTER TABLE [" & tbl.Name & "] ADD CONSTRAINT [" & idx.Name & "] PRIMARY KEY (" & sIdxFields & ");"
                    Else
                        Print #fNum, "CREATE " & IIf(idx.Unique, "UNIQUE ", "") & "INDEX [" & idx.Name & "] ON [" & tbl.Name & "] (" & sIdxFields & ");"
                    End If
                Next
            End If
            Print #fNum, ""
            
            ' 3. INSERT INTO (Navegação via Recordset conforme seu PowerShell)
            If bIncluirDados Then
                Set rs = conn.Execute("SELECT * FROM [" & tbl.Name & "]")
                Do While Not rs.EOF
                    Dim sFields As String: sFields = ""
                    Dim sValues As String: sValues = ""
                    For Each fld In rs.Fields
                        sFields = sFields & "[" & fld.Name & "], "
                        sValues = sValues & TratarValor(fld.Value) & ", "
                    Next
                    Print #fNum, "INSERT INTO [" & tbl.Name & "] (" & Left(sFields, Len(sFields) - 2) & ") VALUES (" & Left(sValues, Len(sValues) - 2) & ");"
                    rs.MoveNext
                Loop
            End If
            Print #fNum, vbCrLf
        End If
    Next
    Close #fNum
    MsgBox "Arquivo gerado com sucesso!", vbInformation
End Sub

Public Sub Convert(ByVal NWindMDBFileName As String, ByVal SQLiteFileName As String)

Dim aCnn As ADODB.Connection
Dim sCnn As cConnection
Dim cconn As String

  On Error Resume Next
  
  Set aCnn = New ADODB.Connection
  aCnn.CursorLocation = adUseClient
  cconn = GeraConexao(NWindMDBFileName)
  aCnn.Open cconn
  
  If Err Then MsgBox Err.Description: Err.Clear: Exit Sub
 
  Err.Clear
  
  If Not EArquivoSQLite(SQLiteFileName) Then
     MsgBox "Destino nao e sql lite"
     Exit Sub
  End If
  
  ValidarOuCriarDestino (SQLiteFileName)
  
  Set sCnn = New_c.Connection
  sCnn.OpenDB SQLiteFileName
  
'  If Err Then MsgBox Err.Description: Err.Clear: Exit Sub
  
  Set C = New cConverter
  C.ConvertDatabase aCnn, sCnn
  
'  If Err Then MsgBox Err.Description: Err.Clear: Exit Sub
  lProgress.Caption = "Table-Schemas created, Table-Data transferred!"
  
  C.ConvertIndexes aCnn, sCnn
  
'  If Err Then MsgBox Err.Description
  lProgress.Caption = "Index-Import finished!"
  
 ' Err.Clear
  
  'If Err Then MsgBox Err.Description: Err.Clear

  Set C = Nothing
End Sub

Public Sub GerarScriptAlteracoes(ByVal sPathOrigem As String, ByVal sPathDestino As String, ByVal bExecutar As Boolean)
    Dim catOrigem As Object, catDestino As Object
    Dim tblOrigem As Object
    Dim connOrigem As New ADODB.Connection
    Dim connDestino As New ADODB.Connection
    Dim sPathSQL As String
    Dim fFile As Integer
    
    ' 1. Define o caminho do SQL (Destino + .sql)
    sPathSQL = sPathDestino & ".sql"
    
    ' 2. Abre o arquivo SQL para escrita (sobrescreve se já existir)
    fFile = FreeFile
    Open sPathSQL For Output As #fFile
    
    ' 3. Conexões
    connOrigem.Open GeraConexao(sPathOrigem)
    connDestino.Open GeraConexao(sPathDestino)
    
    Set catOrigem = CreateObject("ADOX.Catalog")
    Set catDestino = CreateObject("ADOX.Catalog")
    catOrigem.ActiveConnection = connOrigem
    catDestino.ActiveConnection = connDestino
    
    Debug.Print "--- Iniciando Comparação e Gerando Script em: " & sPathSQL & " ---"
    
    ' 4. Itera nas tabelas
    For Each tblOrigem In catOrigem.Tables
        If tblOrigem.Type = "TABLE" And Left(tblOrigem.Name, 4) <> "MSys" Then
            
            On Error Resume Next
            Dim objTemp As Object
            Set objTemp = catDestino.Tables(tblOrigem.Name)
            
            If Err.Number <> 0 Then
                ' Tabela não existe: Escreve no arquivo SQL
                Print #fFile, "-- Tabela ausente: " & tblOrigem.Name
                GerarCreateTabelaSQL tblOrigem, fFile
            Else
               ' PASSO DE SEGURANÇA:
                ' Garante que o catálogo destino está com a estrutura mais recente
                catDestino.Tables.Refresh
                
                ' Re-atribui a tabela para garantir a referência
                Dim tblDestino As Object
                Set tblDestino = catDestino.Tables(tblOrigem.Name)
                
                ' Agora chama a rotina de comparação
                GerarCamposAusentes tblOrigem, tblDestino, fFile
                
                Print #fFile, "-- Verificando índices da tabela: " & tblOrigem.Name
                GerarIndicesAusentes tblOrigem, tblDestino, fFile
                
            End If
            On Error GoTo 0
        End If
    Next
    
    ' 5. Limpeza e FECHAMENTO do arquivo
    Close #fFile
    
    Set catOrigem = Nothing
    Set catDestino = Nothing
    connOrigem.Close
    connDestino.Close
    
    If bExecutar Then
        Call ExecutarSQL(sPathDestino, sPathSQL)
    End If
    
    
    
    MsgBox "Script gerado com sucesso em: " & sPathSQL, vbInformation
End Sub

Public Sub GerarCreateTabelaSQL(ByRef tblOrigem As Object, ByVal fFile As Integer)
    Dim ssql As String
    Dim col As Object
    Dim bPrimeiro As Boolean
    
    ssql = "CREATE TABLE [" & tblOrigem.Name & "] ("
    bPrimeiro = True
    
    ' Percorre colunas
    For Each col In tblOrigem.Columns
        If Not bPrimeiro Then ssql = ssql & ", "
        
        ' Aqui você usa sua lógica de conversão de tipo (GetSQLType que já tem no módulo)
        ssql = ssql & "[" & col.Name & "] " & GetTipoADOX(col.Type)
        bPrimeiro = False
    Next
    
    ssql = ssql & ");"
    
    ' Grava no handle de gravação já aberto
    Print #fFile, ssql
    Print #fFile, "" ' Linha em branco para organização
End Sub

Public Sub GerarCamposAusentes(ByRef tblOrigem As Object, ByRef tblDestino As Object, ByVal fFile As Integer)
    Dim colOrigem As Object
    Dim colDestino As Object
    
    ' Percorre cada coluna da tabela de origem
    For Each colOrigem In tblOrigem.Columns
        On Error Resume Next
        ' Tenta acessar a coluna correspondente no destino
        Set colDestino = tblDestino.Columns(colOrigem.Name)
        
        ' Se Err.Number <> 0, significa que o campo não existe no destino
        If Err.Number <> 0 Then
            Err.Clear
            
            ' Gera o comando ALTER TABLE e escreve no arquivo SQL
            ' Nota: MapearTipoADOX é a função que você já deve ter configurada
            Dim ssql As String
            ssql = "ALTER TABLE [" & tblOrigem.Name & "] ADD COLUMN [" & colOrigem.Name & "] " & GetTipoADOX(colOrigem.Type) & ";"
            
            Print #fFile, ssql
            Debug.Print "Campo ausente encontrado: " & colOrigem.Name & " na tabela " & tblOrigem.Name
        End If
        On Error GoTo 0
    Next
End Sub
Public Sub GerarIndicesAusentes(ByRef tblOrigem As Object, ByRef tblDestino As Object, ByVal fFile As Integer)
    Dim idxOrigem As Object
    Dim idxDestino As Object
    Dim colIdx As Object
    Dim sCampos As String
    Dim ssql As String
    
    ' O Refresh é CRUCIAL para garantir que a coleção de índices está atualizada
    tblDestino.Indexes.Refresh
    
    For Each idxOrigem In tblOrigem.Indexes
        ' Se o seu MENU-1 for uma Chave Primária, esta verificação o está pulando
        ' Vamos remover o "If Not idxOrigem.PrimaryKey" temporariamente para testar
        
        On Error Resume Next
        Set idxDestino = tblDestino.Indexes(idxOrigem.Name)
        
        ' Se Err.Number <> 0, o índice realmente não existe no destino
        If Err.Number <> 0 Then
            Err.Clear
            
            ' Monta os campos
            sCampos = ""
            For Each colIdx In idxOrigem.Columns
                sCampos = sCampos & "[" & colIdx.Name & "], "
            Next
            If Len(sCampos) > 2 Then sCampos = Left(sCampos, Len(sCampos) - 2)
            
            ' Verifica se é Primário ou Único
            If idxOrigem.PrimaryKey Then
                ssql = "ALTER TABLE [" & tblOrigem.Name & "] ADD CONSTRAINT [" & idxOrigem.Name & "] PRIMARY KEY (" & sCampos & ");"
            Else
                Dim sUnique As String
                sUnique = IIf(idxOrigem.Unique, "UNIQUE ", "")
                ssql = "CREATE " & sUnique & "INDEX [" & idxOrigem.Name & "] ON [" & tblOrigem.Name & "] (" & sCampos & ");"
            End If
            
            Print #fFile, ssql
            Debug.Print "Gerando índice: " & idxOrigem.Name
        End If
        On Error GoTo 0
    Next
End Sub


Public Sub ExecutarSQL(ByVal sPathDestino As String, ByVal sPathSQL As String)
    Dim connDestino As New ADODB.Connection
    Dim fFile As Integer
    Dim sLinha As String
    Dim sComando As String
    
    ' Abre conexão com o destino
    connDestino.Open GeraConexao(sPathDestino)
    
    fFile = FreeFile
    Open sPathSQL For Input As #fFile
    
    On Error Resume Next ' Ignora erros caso um comando falhe (ex: índice já existe)
    Do While Not EOF(fFile)
        Line Input #fFile, sLinha
        
        ' Ignora comentários e linhas em branco
        If Trim(sLinha) <> "" And Left(sLinha, 2) <> "--" Then
            sComando = sComando & " " & sLinha
            
            ' Se a linha termina com ponto e vírgula, executa o comando acumulado
            If InStr(sLinha, ";") > 0 Then
                connDestino.Execute sComando
                sComando = "" ' Reseta para o próximo comando
            End If
        End If
    Loop
    
    Close #fFile
    connDestino.Close
    
    MsgBox "Script executado com sucesso no destino!", vbInformation
End Sub
Public Sub ExecutarSQLite(ByVal sPathDestino As String, ByVal sPathSQL As String)
    Dim conn As Object
    Dim fFile As Integer
    Dim sLinha As String, sComando As String
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open GeraConexao(sPathDestino)
    
    fFile = FreeFile
    Open sPathSQL For Input As #fFile
    
    sComando = ""
    
    Do While Not EOF(fFile)
        Line Input #fFile, sLinha
        sLinha = Trim(sLinha)
        
        ' 1. Ignora linhas vazias
        ' 2. Ignora linhas que começam com "--" (comentários padrão SQL)
        If sLinha <> "" And Left(sLinha, 2) <> "--" Then
            
            ' Adiciona a linha ao buffer de comando
            sComando = sComando & " " & sLinha
            
            ' Se encontrar o ponto e vírgula, o comando está completo
            If InStr(sComando, ";") > 0 Then
                ' Executa apenas a parte até o primeiro ";" encontrado
                ' Isso evita erros se houver lixo após o ";"
                Dim sExec As String
                sExec = Left(sComando, InStr(sComando, ";") - 1)
                
                On Error Resume Next
                conn.Execute sExec
                If Err.Number <> 0 Then
                    Debug.Print "Erro no comando: " & sExec & " | Detalhe: " & Err.Description
                    Err.Clear
                End If
                On Error GoTo 0
                
                ' Limpa o buffer para o próximo comando
                sComando = ""
            End If
        End If
    Loop
    
    Close #fFile
    conn.Close
    Set conn = Nothing
    
    MsgBox "Script executado linha a linha com sucesso!", vbInformation
End Sub

Public Sub ExecutarSQLiteold(ByVal sPathDestino As String, ByVal sPathSQL As String)
    Dim conn As New ADODB.Connection
    Dim fso As Object
    Dim sSQLConteudo As String
    Dim vComandos As Variant, i As Long
    
    ' 1. Verifica se o arquivo de script existe
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(sPathSQL) Then
        MsgBox "Arquivo de script SQL não encontrado: " & sPathSQL, vbCritical
        Exit Sub
    End If
    
    ' 2. Lê todo o conteúdo do arquivo
    sSQLConteudo = fso.OpenTextFile(sPathSQL, 1).ReadAll
    
    ' 3. Abre a conexão com o banco SQLite
    ' Certifique-se de que a função GeraConexao suporte SQLite
    conn.Open GeraConexao(sPathDestino)
    
    ' 4. Divide e Executa
    ' O SQLite geralmente aceita múltiplos comandos separados por ';'
    ' Mas executar um por um garante que você possa tratar erros individualmente
    vComandos = Split(sSQLConteudo, ";")
    
    On Error Resume Next ' Inicia tratamento de erro para os comandos
    For i = LBound(vComandos) To UBound(vComandos)
        If Trim(vComandos(i)) <> "" Then
            conn.Execute vComandos(i)
            If Err.Number <> 0 Then
                Debug.Print "Erro ao executar comando: " & vComandos(i) & " | Erro: " & Err.Description
                Err.Clear
            End If
        End If
    Next
    On Error GoTo 0
    
    conn.Close
    Set conn = Nothing
    Set fso = Nothing
End Sub

Public Sub GerarArquivoSchemaADO(ByVal sPathOrigem As String, ByVal sPathDestinoSchema As String)
    Dim cat As Object
    Dim tbl As Object, col As Object, idx As Object, idxCol As Object
    Dim connSchema As ADODB.Connection
    Dim rsTb As ADODB.Recordset, rsIx As ADODB.Recordset
    Dim cEXT As String
    Dim rsTabelas As ADODB.Recordset
    Dim rsPragmaColunas As ADODB.Recordset
    Dim rsPragmaIndices As ADODB.Recordset
    Dim sTabela As String
    Dim seq As Long
    Dim sCampos As String
    Dim cmd As ADODB.Command
    Dim connOrigem As ADODB.Connection
    Dim nSize As Long, nPrec As Long, nScale As Long
    Dim sTipoCompleto As String
    Dim sConteudo As String
    Dim vPartes As Variant
    
    
    cEXT = UCase(EXTENSAO(sPathOrigem))
    
    
    ' 1. Conectar ao banco de origem via ADOX
    Set cat = CreateObject("ADOX.Catalog")
    cat.ActiveConnection = GeraConexao(sPathOrigem)
    
    ' 2. Conectar ao banco que guarda o Schema
    Set connSchema = New ADODB.Connection
    connSchema.Open GeraConexao(sPathDestinoSchema)
    
    ' 3. Preparar Recordsets
    Set rsTb = New ADODB.Recordset
    rsTb.Open "SELECT * FROM [myTABLE]", connSchema, adOpenDynamic, adLockOptimistic
    
    Set rsIx = New ADODB.Recordset
    rsIx.Open "SELECT * FROM [myINDEX]", connSchema, adOpenDynamic, adLockOptimistic
    
    ' Limpar dados anteriores
    connSchema.Execute "DELETE FROM [myTABLE] WHERE TABLETIPO='" & cEXT & "'"
    connSchema.Execute "DELETE FROM [myINDEX] WHERE TABLETIPO='" & cEXT & "'"
    
    ' 4. Processar Tabelas e Colunas
    If EArquivoSQLite(sPathOrigem) Then
       Set connOrigem = New ADODB.Connection
       connOrigem.ConnectionString = GeraConexao(sPathOrigem)
    
       On Error Resume Next ' Captura erro de abertura
       connOrigem.Open
    
    ' 1. Instancia o objeto de comando
        Set cmd = New ADODB.Command
        Set cmd.ActiveConnection = connOrigem ' Define a conexão ativa
        
        ' 2. Define o SQL
        cmd.CommandText = "SELECT name FROM sqlite_master WHERE type = 'table' AND SUBSTR(name, 1, 7) <> 'sqlite_'"
        cmd.CommandType = adCmdText
        
        ' 3. Executa e recebe o Recordset
        ' O parâmetro 'Set' é obrigatório aqui
        Set rsTabelas = cmd.Execute()
    
        ' Consulta a lista de tabelas no SQLite
        'Set rsTabelas = connOrigem.Execute("SELECT name FROM sqlite_master WHERE type = 'table' AND SUBSTR(name, 1, 7) <> 'sqlite_'")
        
        Do While Not rsTabelas.EOF
            sTabela = rsTabelas("name").Value
            
            ' --- A. Processar Colunas (Substitui For Each col In tbl.Columns) ---
            Set rsPragmaColunas = connOrigem.Execute("PRAGMA table_info([" & sTabela & "])")
            seq = 1
            
            Do While Not rsPragmaColunas.EOF
                rsTb.AddNew
                rsTb("TableName") = sTabela
                rsTb("SeqNum") = rsPragmaColunas("cid").Value 'seq
                rsTb("FieldName") = rsPragmaColunas("name").Value
                rsTb("DefaultValue") = rsPragmaColunas("dflt_value").Value
                rsTb("notnull") = rsPragmaColunas("notnull").Value
                rsTb("Required") = CBool(rsPragmaColunas("notnull").Value)
                rsTb("AllowZeroLength") = True
                
                'rsTb("") = rsPragmaColunas("").Value
                
                sTipoCompleto = rsPragmaColunas("type").Value
                ' O SQLite retorna o tipo como texto (ex: TEXT, INTEGER, REAL)
                rsTb("TipoNome") = sTipoCompleto 'rsPragmaColunas("type").Value
                nSize = 0
                nPrec = 0
                nScale = 0
                
                
                If InStr(sTipoCompleto, "(") > 0 Then
                    sConteudo = Mid(sTipoCompleto, InStr(sTipoCompleto, "(") + 1, InStr(sTipoCompleto, ")") - InStr(sTipoCompleto, "(") - 1)
                    If InStr(sConteudo, ",") > 0 Then
                        ' Caso: NUMERIC(10,2)
                        vPartes = Split(sConteudo, ",")
                        nPrec = CLng(vPartes(0))
                        nScale = CLng(vPartes(1))
                    Else
                        ' Caso: VARCHAR(200)
                        nSize = CLng(sConteudo)
                    End If
                End If
                            
                
                
                ' Aplica os valores extraídos
                If nSize > 0 Then rsTb("DefinedSize") = nSize
                If nPrec > 0 Then rsTb("NumericPrecision") = nPrec
                If nScale > 0 Then rsTb("NumericScale") = nScale
                
                
                rsTb("TABLETIPO") = cEXT
                rsTb("Status") = "OK"
                
                ' Nota: SQLite não tem propriedades complexas como Jet (Precision, Scale)
                ' da mesma forma, mas você pode salvar o 'notnull' ou 'pk' aqui se desejar
                rsTb.Update
                
                seq = seq + 1
                rsPragmaColunas.MoveNext
            Loop
            rsPragmaColunas.Close
            
            ' --- B. Processar Índices (Substitui For Each idx In tbl.Indexes) ---
            Set rsPragmaIndices = connOrigem.Execute("PRAGMA index_list([" & sTabela & "])")
            
            Do While Not rsPragmaIndices.EOF
                Dim sIdxName As String: sIdxName = rsPragmaIndices("name").Value
                
                rsIx.AddNew
                rsIx("TableName") = sTabela
                rsIx("IndexName") = sIdxName
                rsIx("Unique") = CBool(rsPragmaIndices("unique").Value)
                rsIx("TABLETIPO") = cEXT
                rsIx("Status") = "OK"
                
                ' Para pegar os campos do índice, precisamos de outra query PRAGMA
                Dim rsIdxInfo As ADODB.Recordset
                Set rsIdxInfo = connOrigem.Execute("PRAGMA index_info([" & sIdxName & "])")
                
                sCampos = ""
                Do While Not rsIdxInfo.EOF
                    sCampos = sCampos & "+" & rsIdxInfo("name").Value & ";"
                    rsIdxInfo.MoveNext
                Loop
                rsIdxInfo.Close
                
                rsIx("Fields") = sCampos
                rsIx.Update
                
                rsPragmaIndices.MoveNext
            Loop
            rsPragmaIndices.Close
            
            rsTabelas.MoveNext
        Loop
        rsTabelas.Close
    Else
        For Each tbl In cat.Tables
            If tbl.Type = "TABLE" And Left(tbl.Name, 4) <> "MSys" Then
                
                seq = 1
                For Each col In tbl.Columns
                    rsTb.AddNew
                    rsTb("TableName") = tbl.Name
                    rsTb("SeqNum") = seq
                    rsTb("FieldName") = col.Name
                    rsTb("FieldType") = col.Type
                    
                    ' Campos novos e propriedades avançadas via Properties
                    rsTb("DefinedSize") = col.DefinedSize
                    rsTb("Size") = col.DefinedSize
                    
                    On Error Resume Next
                    rsTb("NumericPrecision") = col.Properties("Precision").Value
                    rsTb("NumericScale") = col.Properties("Scale").Value
                    rsTb("Required") = Not col.Properties("Nullable").Value
                    rsTb("AllowZeroLength") = col.Properties("Jet OLEDB:Allow Zero Length").Value
                    rsTb("DefaultValue") = col.Properties("Default").Value
                    rsTb("Attributes") = col.Attributes
                    rsTb("notnull") = Not col.Properties("Nullable").Value
                    On Error GoTo 0
                    
                    ' Usando sua função existente para definir o nome do tipo
                    rsTb("TipoNome") = GetTipoNome(col.Type)
                    rsTb("TABLETIPO") = cEXT
                    rsTb("Status") = "OK"
                    
                    rsTb.Update
                    seq = seq + 1
                Next
                
                ' 5. Processar Índices
                For Each idx In tbl.Indexes
                    rsIx.AddNew
                    rsIx("TableName") = tbl.Name
                    rsIx("IndexName") = idx.Name
                    rsIx("Primary") = idx.PrimaryKey
                    rsIx("Unique") = idx.Unique
                    rsIx("TABLETIPO") = cEXT
                    rsIx("Status") = "OK"
                    
                    ' Concatenação de campos no formato +CAMPO;
                     sCampos = ""
                    For Each idxCol In idx.Columns
                        sCampos = sCampos & "+" & idxCol.Name & ";"
                    Next
                    rsIx("Fields") = sCampos
                    rsIx.Update
                Next
            End If
        Next
    End If
    ' Limpeza
    rsTb.Close: Set rsTb = Nothing
    rsIx.Close: Set rsIx = Nothing
    connSchema.Close: Set connSchema = Nothing
    Set cat = Nothing
    
    MsgBox "Arquivo de Schema gerado com sucesso via ADO/ADOX!", vbInformation
End Sub


Public Sub GarantirSchemaExistente(ByVal sPathDestinoSchema As String)
    Dim conn As ADODB.Connection
    Dim cCONEXAO As String
   
       
    ValidarOuCriarDestino (sPathDestinoSchema)
    
    If ArquivoExiste(sPathDestinoSchema) Then
        ' 2. Conecta e cria as tabelas
        Set conn = New ADODB.Connection
        conn.Open GeraConexao(sPathDestinoSchema)
        
        If Not TabelaExiste(conn, "myINDEX") Then
           ' Executa os comandos de criação das 3 tabelas
            conn.Execute "CREATE TABLE [myINDEX] ([TableTipo] TEXT(255),[Fields] MEMO, [IndexName] TEXT(255), [Primary] BIT, [Status] TEXT(255), [TableName] TEXT(255), [Unique] BIT);"
        End If
        
        If Not TabelaExiste(conn, "myQUERY") Then
          conn.Execute "CREATE TABLE [myQUERY] ([TableTipo] TEXT(255),[QueryDef] MEMO, [QueryName] TEXT(255), [Status] TEXT(255));"
        End If
        If Not TabelaExiste(conn, "myTABLE") Then
           conn.Execute "CREATE TABLE [myTABLE] ([TableTipo] TEXT(255),[notnull] BIT,[AllowZeroLength] BIT, [Attributes] LONG, [DefaultValue] TEXT(255), [DefinedSize] LONG, [FieldName] TEXT(255), [FieldType] LONG, [NumericPrecision] LONG, [NumericScale] LONG, [Required] BIT, [SeqNum] LONG, [Size] LONG, [Status] TEXT(255), [TableName] TEXT(255), [TipoNome] TEXT(255));"
        End If
        conn.Close
        Set conn = Nothing
    End If
End Sub


Public Sub RecriarBancoDoSchema(ByVal sPathSchema As String, ByVal sPathDestino As String)
    Dim connSchema As ADODB.Connection
    Dim connDestino As ADODB.Connection
    Dim rsTb As ADODB.Recordset
    Dim ssql As String
    Dim sTabelaAtual As String
    Dim sTipoSQL As String
    
    ' 1. Conexões
    Set connSchema = New ADODB.Connection
    connSchema.Open GeraConexao(sPathSchema)
    
    
    ValidarOuCriarDestino (sPathDestino)
    ' Cria o destino se não existir
'    If Dir(sPathDestino) = "" Then
'        Dim cat As Object: Set cat = CreateObject("ADOX.Catalog")
 '       cat.Create GeraConexao(sPathDestino)
 '   End If
    
    Set connDestino = New ADODB.Connection
    connDestino.Open GeraConexao(sPathDestino)
    
    ' 2. Ler todas as tabelas registradas no Schema
    ' Lemos ordenado por TableName e SeqNum para garantir a ordem das colunas
    Set rsTb = New ADODB.Recordset
    rsTb.Open "SELECT * FROM [myTABLE] ORDER BY [TableName], [SeqNum]", connSchema, adOpenStatic
    
    sTabelaAtual = ""
    
    Do While Not rsTb.EOF
        ' Se mudou de tabela, fechamos a anterior (se houver) e iniciamos a nova
        If rsTb("TableName") <> sTabelaAtual Then
            If sTabelaAtual <> "" Then
                ssql = Left(ssql, Len(ssql) - 2) & ");" ' Remove última vírgula e fecha
                connDestino.Execute ssql
            End If
            
            sTabelaAtual = rsTb("TableName")
            ssql = "CREATE TABLE [" & sTabelaAtual & "] ("
        End If
        
        ' Monta a definição do campo
        'sSQL = sSQL & "[" & rsTb("FieldName") & "] " & rsTb("TipoNome") & ", "
        sTipoSQL = ConverterParaSQL(rsTb("FieldType"), rsTb("DefinedSize"))
        ssql = ssql & "[" & rsTb("FieldName") & "] " & sTipoSQL & ", "
        
        rsTb.MoveNext
    Loop
    
    ' Executa a última tabela do loop
    If ssql <> "CREATE TABLE [" & sTabelaAtual & "] (" Then
        ssql = Left(ssql, Len(ssql) - 2) & ");"
        connDestino.Execute ssql
    End If
    
    rsTb.Close
    connSchema.Close
    connDestino.Close
    
    MsgBox "Banco de dados recriado com sucesso a partir do Schema!", vbInformation
End Sub


Private Function ConverterParaSQL(ByVal iFieldType As Long, ByVal iSize As Long) As String
    Select Case iFieldType
        Case 11           ' adBoolean
            ConverterParaSQL = "BIT"
        Case 2            ' adSmallInt
            ConverterParaSQL = "SMALLINT"
        Case 3            ' adInteger
            ConverterParaSQL = "LONG"
        Case 20           ' adBigInt
            ConverterParaSQL = "LONG"
        Case 4            ' adSingle
            ConverterParaSQL = "SINGLE"
        Case 5            ' adDouble
            ConverterParaSQL = "DOUBLE"
        Case 6            ' adCurrency
            ConverterParaSQL = "CURRENCY"
        Case 7            ' adDate
            ConverterParaSQL = "DATETIME"
        Case 131, 139     ' adNumeric, adDecimal
            ConverterParaSQL = "DECIMAL"
        Case 130, 202     ' adWChar, adVarWChar
            ' Se o tamanho for muito grande, vira MEMO, senão TEXT
            If iSize > 255 Or iSize <= 0 Then
                ConverterParaSQL = "MEMO"
            Else
                ConverterParaSQL = "TEXT(" & iSize & ")"
            End If
        Case 128, 204     ' adBinary, adVarBinary
            ConverterParaSQL = "BINARY"
        Case 203, 205     ' adLongVarWChar, adLongVarBinary (Memo/OLE)
            ConverterParaSQL = "MEMO"
        Case Else
            ' Fallback para segurança
            ConverterParaSQL = "TEXT(255)"
    End Select
End Function
Public Sub RecriarIndicesDoSchema(ByVal sPathSchema As String, ByVal sPathDestino As String)
    Dim connSchema As ADODB.Connection, connDestino As ADODB.Connection
    Dim rsIx As ADODB.Recordset
    Dim ssql As String, sCamposFinal As String
    Dim vCampos As Variant, i As Long
    
    Set connSchema = New ADODB.Connection: connSchema.Open GeraConexao(sPathSchema)
    Set connDestino = New ADODB.Connection: connDestino.Open GeraConexao(sPathDestino)
    
    Set rsIx = New ADODB.Recordset
    rsIx.Open "SELECT * FROM [myINDEX]", connSchema, adOpenStatic, adLockReadOnly
    
    Do While Not rsIx.EOF
        ' 1. Limpa a string original para facilitar o split
        ' Remove o + inicial e o ; final
        Dim sRaw As String
        sRaw = Replace(rsIx("Fields"), "+", "")
        If Right(sRaw, 1) = ";" Then sRaw = Left(sRaw, Len(sRaw) - 1)
        
        ' 2. Split pelos pontos e vírgulas
        vCampos = Split(sRaw, ";")
        sCamposFinal = ""
        
        ' 3. Reconstrói os campos com colchetes e vírgulas
        For i = LBound(vCampos) To UBound(vCampos)
            sCamposFinal = sCamposFinal & "[" & vCampos(i) & "]"
            If i < UBound(vCampos) Then sCamposFinal = sCamposFinal & ", "
        Next i
        
        ' 4. Monta o comando
        If rsIx("Primary") = True Then
            ssql = "ALTER TABLE [" & rsIx("TableName") & "] ADD CONSTRAINT [" & rsIx("IndexName") & "] PRIMARY KEY (" & sCamposFinal & ")"
        Else
            Dim sUnique As String: sUnique = IIf(rsIx("Unique") = True, "UNIQUE ", "")
            ssql = "CREATE " & sUnique & "INDEX [" & rsIx("IndexName") & "] ON [" & rsIx("TableName") & "] (" & sCamposFinal & ")"
        End If
        
        ' Execução com log de erro para debug
        On Error Resume Next
        connDestino.Execute ssql
        If Err.Number <> 0 Then
            Debug.Print "Erro no índice " & rsIx("IndexName") & ": " & Err.Description & " | SQL: " & ssql
        End If
        On Error GoTo 0
        
        rsIx.MoveNext
    Loop
    
    rsIx.Close: connSchema.Close: connDestino.Close
End Sub
Private Function TabelaExiste(conn As ADODB.Connection, sTabela As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = conn.OpenSchema(adSchemaTables, Array(Empty, Empty, sTabela, "TABLE"))
    TabelaExiste = Not rs.EOF
    rs.Close
End Function

Private Function CampoExiste(conn As ADODB.Connection, sTabela As String, sCampo As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = conn.OpenSchema(adSchemaColumns, Array(Empty, Empty, sTabela, sCampo))
    CampoExiste = Not rs.EOF
    rs.Close
End Function
Private Function IndiceExiste(conn As ADODB.Connection, sTabela As String, sIndice As String) As Boolean
    Dim rs As ADODB.Recordset
    ' O OpenSchema de Indexes retorna todos os índices da tabela
    Set rs = conn.OpenSchema(adSchemaIndexes, Array(Empty, Empty, sTabela))
    
    IndiceExiste = False
    Do While Not rs.EOF
        If rs("INDEX_NAME") = sIndice Then
            IndiceExiste = True
            Exit Do
        End If
        rs.MoveNext
    Loop
    rs.Close
End Function

Public Function FormatarCamposParaSQL(ByVal sFieldsRaw As String) As String
    Dim vCampos As Variant
    Dim i As Long
    Dim sResult As String
    
    ' 1. Limpeza inicial
    ' Remove o '+' se existir e o ';' final para facilitar o split
    Dim sClean As String
    sClean = Replace(sFieldsRaw, "+", "")
    If Right(sClean, 1) = ";" Then
        sClean = Left(sClean, Len(sClean) - 1)
    End If
    
    ' 2. Divide os campos usando o ';' como delimitador
    vCampos = Split(sClean, ";")
    
    ' 3. Reconstrói com colchetes e vírgulas
    sResult = ""
    For i = LBound(vCampos) To UBound(vCampos)
        If Trim(vCampos(i)) <> "" Then
            sResult = sResult & "[" & Trim(vCampos(i)) & "]"
            ' Adiciona a vírgula se não for o último elemento
            If i < UBound(vCampos) Then
                sResult = sResult & ", "
            End If
        End If
    Next i
    
    FormatarCamposParaSQL = sResult
End Function

Public Sub GerarTabelasAusentesViaSchema(ByVal sTabela As String, ByVal connSchema As ADODB.Connection, ByVal ts As Object)
    Dim rsCampos As ADODB.Recordset
    Dim ssql As String, sCamposDef As String
    
    ' Busca os campos da tabela no schema, ordenados pela sequência
    Set rsCampos = connSchema.Execute("SELECT * FROM [myTABLE] WHERE [TableName]='" & sTabela & "' ORDER BY [SeqNum]")
    
    sCamposDef = ""
    Do While Not rsCampos.EOF
        ' Monta a string: [Campo] Tipo,
        sCamposDef = sCamposDef & "[" & rsCampos("FieldName") & "] " & rsCampos("TipoNome") & ", "
        rsCampos.MoveNext
    Loop
    
    ' Remove a última vírgula e espaço
    If Len(sCamposDef) > 0 Then
        sCamposDef = Left(sCamposDef, Len(sCamposDef) - 2)
        ssql = "CREATE TABLE [" & sTabela & "] (" & sCamposDef & ");"
        ts.WriteLine ssql
    End If
    
    rsCampos.Close
    Set rsCampos = Nothing
End Sub

Public Sub GerarCamposAusentesViaSchema(ByVal sTabela As String, ByVal connSchema As ADODB.Connection, ByVal connDestino As ADODB.Connection, ByVal ts As Object)
    Dim rsSchema As ADODB.Recordset
    
    Set rsSchema = connSchema.Execute("SELECT * FROM [myTABLE] WHERE [TableName]='" & sTabela & "' ORDER BY [SeqNum]")
    
    Do While Not rsSchema.EOF
        ' Verifica se o campo existe no catálogo do banco destino
        If Not CampoExiste(connDestino, sTabela, rsSchema("FieldName")) Then
            ' Campo ausente: gera o script de adição
            ts.WriteLine "ALTER TABLE [" & sTabela & "] ADD COLUMN [" & rsSchema("FieldName") & "] " & rsSchema("TipoNome") & ";"
        End If
        rsSchema.MoveNext
    Loop
    rsSchema.Close
    Set rsSchema = Nothing
End Sub

Public Sub GerarIndicesAusentesViaSchema(ByVal sTabela As String, ByVal connSchema As ADODB.Connection, ByVal connDestino As ADODB.Connection, ByVal ts As Object)
    Dim rsIdx As ADODB.Recordset
    
    Set rsIdx = connSchema.Execute("SELECT * FROM [myINDEX] WHERE [TableName]='" & sTabela & "'")
    
    Do While Not rsIdx.EOF
        ' Verifica se o índice já existe no destino
        If Not IndiceExiste(connDestino, sTabela, rsIdx("IndexName")) Then
            Dim sCampos As String: sCampos = FormatarCamposParaSQL(rsIdx("Fields"))
            Dim ssql As String
            
            ' Se for Primary Key, o tratamento é diferenciado
            If rsIdx("Primary") = True Then
                ssql = "ALTER TABLE [" & sTabela & "] ADD CONSTRAINT [" & rsIdx("IndexName") & "] PRIMARY KEY (" & sCampos & ");"
            Else
                Dim sUnique As String: sUnique = IIf(rsIdx("Unique") = True, "UNIQUE ", "")
                ssql = "CREATE " & sUnique & "INDEX [" & rsIdx("IndexName") & "] ON [" & sTabela & "] (" & sCampos & ");"
            End If
            
            ts.WriteLine ssql
        End If
        rsIdx.MoveNext
    Loop
    rsIdx.Close
    Set rsIdx = Nothing
End Sub


Public Sub GerarScriptAlteracoesViaSchema(ByVal sPathSchema As String, ByVal sPathDestino As String, _
                                         ByVal bExecutar As Boolean)
    Dim connSchema As ADODB.Connection
    Dim connDestino As ADODB.Connection
    Dim rsTabelas As ADODB.Recordset
    Dim sPathSQL As String
    Dim cEXT As String
    Dim fso As Object, ts As Object ' Usando FSO em vez de Open For Output
    
    
    cEXT = UCase(EXTENSAO(sPathDestino))
    
    ' 1. Configurações de arquivo (usando FSO para evitar erros de I/O)
    sPathSQL = sPathDestino & "_atualizacao.sql"
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(sPathSQL, True)
    
    ' 2. Conexões
    Set connSchema = New ADODB.Connection
    connSchema.Open GeraConexao(sPathSchema)
    
    Set connDestino = New ADODB.Connection
    connDestino.Open GeraConexao(sPathDestino)
    
    ' 3. Lê a lista de tabelas registradas no Schema
    Set rsTabelas = connSchema.Execute("SELECT DISTINCT [TableName] FROM [myTABLE]  WHERE TABLETIPO='" & cEXT & "' ORDER BY [TableName]")
    
    Do While Not rsTabelas.EOF
        Dim sTabela As String: sTabela = rsTabelas("TableName")
        
        ' Verifica se a tabela existe no destino
        If Not TabelaExiste(connDestino, sTabela) Then
            ' Se não existe, gera o CREATE TABLE (você precisará de uma função para isso)
            ts.WriteLine "-- Tabela ausente: " & sTabela
            ' GerarCreateTabelaDoSchema sTabela, connSchema, ts
            GerarTabelasAusentesViaSchema sTabela, connSchema, ts
        Else
            ' Se a tabela existe, compara os campos gravados no Schema com o destino
            ts.WriteLine "-- Verificando campos da tabela: " & sTabela
            GerarCamposAusentesViaSchema sTabela, connSchema, connDestino, ts
            
            ts.WriteLine "-- Verificando índices da tabela: " & sTabela
            GerarIndicesAusentesViaSchema sTabela, connSchema, connDestino, ts
        End If
        
        rsTabelas.MoveNext
    Loop
    
    ' 4. Finalização
    ts.Close
    connSchema.Close
    connDestino.Close
    
    If bExecutar Then Call ExecutarSQL(sPathDestino, sPathSQL)
    MsgBox "Script de atualização gerado: " & sPathSQL, vbInformation
End Sub


Public Sub ProcessarPastaCompletaSchema(ByVal sFolder As String)
    Dim fso As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim cEXT As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    
       
    If Not fso.FolderExists(sFolder) Then
        MsgBox "Diretório não encontrado!", vbCritical
        Exit Sub
    End If
    
    Set oFolder = fso.GetFolder(sFolder)
    
    ' Itera por todos os arquivos da pasta
    For Each oFile In oFolder.Files
        ' Verifica se é um arquivo .mdb
        sCaminhoOrigem = oFile.Path
        sNomeArquivo = fso.GetBaseName(oFile.Name)
        cEXT = LCase(fso.GetExtensionName(sCaminhoOrigem))
        If cEXT = "mdb" Or cEXT = "accdb" Or cEXT = "sqlite" Then
        
            sCaminhoDestino = sFolder & "\" & sNomeArquivo & "_schema.mdb"
    
            
            GarantirSchemaExistente sCaminhoDestino
            ' 4. Chama a rotina que criamos anteriormente
            Call GerarArquivoSchemaADO(sCaminhoOrigem, sCaminhoDestino)

        
        
        
        End If
    Next oFile
    
    MsgBox "Processamento da pasta concluído!", vbInformation
    
    Set oFolder = Nothing
    Set fso = Nothing
End Sub
' Função principal que orquestra a geração de SQL
Public Sub GerarScriptsSQL(ByVal sCaminho As String, ByVal bIncluirDados As Boolean)
    
    If EArquivoSQLite(sCaminho) Then
        ' Lógica para SQLite
        ' Aqui você chamaria uma rotina que percorre o Schema do SQLite
        ' e monta as strings de CREATE/INSERT usando colScriptsPendentes.Add
        GerarScriptSQLite sCaminho, bIncluirDados
    Else
        ' Lógica para Access (ADOX)
        GerarScriptAccess sCaminho, bIncluirDados
    End If
    
    ' Após preencher a coleção, você pode exibir no TxtInfo ou salvar em arquivo
    MsgBox "Scripts gerados! Total: " ' & colScriptsPendentes.Count
End Sub

Public Sub ProcessarPastaCompletaSql(ByVal sFolder As String, ByVal lINCLUI As Boolean)
    Dim fso As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim cEXT As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    
       
    If Not fso.FolderExists(sFolder) Then
        MsgBox "Diretório não encontrado!", vbCritical
        Exit Sub
    End If
    
    Set oFolder = fso.GetFolder(sFolder)
    
    ' Itera por todos os arquivos da pasta
    For Each oFile In oFolder.Files
        ' Verifica se é um arquivo .mdb
        sCaminhoOrigem = oFile.Path
        sNomeArquivo = fso.GetBaseName(oFile.Name)
        cEXT = LCase(fso.GetExtensionName(sCaminhoOrigem))
        If cEXT = "mdb" Or cEXT = "accdb" Or cEXT = "sqlite" Then
           GerarScriptsSQL sCaminhoOrigem, lINCLUI '(CheckIncDados.Value = vbChecked)
        End If
    Next oFile
    MsgBox "Processamento da pasta concluído!", vbInformation
    
    Set oFolder = Nothing
    Set fso = Nothing
End Sub
Public Function GetTipoSqlite(ByVal adType As Long) As String
    Select Case adType
        ' Inteiros
        Case 3, 20:          GetTipoSqlite = "INTEGER"
        
        ' Texto
        Case 129, 200, 202:  GetTipoSqlite = "TEXT"
        
        ' Data/Hora
        Case 7, 133, 135:    GetTipoSqlite = "DATETIME"
        
        ' Booleano
        Case 11:             GetTipoSqlite = "BOOLEAN"
        
        ' Reais / Ponto Flutuante
        Case 4, 5:           GetTipoSqlite = "REAL"
        
        ' Binários / Imagens (LongVarBinary / LongBinary)
        ' O valor 205 corresponde a adLongVarBinary
        Case 205, 128:       GetTipoSqlite = "BLOB"
        
        ' Padrão para tipos não listados
        Case Else:           GetTipoSqlite = "TEXT"
    End Select
End Function

Public Sub GerarCreateTabelaSqlite(tblOrigem As Object, fFile As Integer)
    Dim ssql As String
    Dim col As Object
    Dim bPrimeira As Boolean
    
    ' Como a checagem já foi feita na função mãe, aqui só geramos o comando
    ssql = "CREATE TABLE [" & tblOrigem.Name & "] ("
    bPrimeira = True
    
    ' Iteramos sobre as colunas da origem (MDB)
    For Each col In tblOrigem.Columns
        If Not bPrimeira Then ssql = ssql & ", "
        
        ' Converte o tipo ADOX do Access para SQLite
        ssql = ssql & "[" & col.Name & "] " & GetTipoSqlite(col.Type)
        
        bPrimeira = False
    Next
    
    ssql = ssql & ");"
    
    ' Grava o comando no arquivo
    Print #fFile, ssql
End Sub
Public Sub GerarScriptAlteracoesAccessParaSqlite(ByVal sPathOrigem As String, ByVal sPathDestino As String, ByVal bExecutar As Boolean)
    Dim catOrigem As Object
    Dim tblOrigem As Object
    Dim connOrigem As New ADODB.Connection
    Dim connDestino As New ADODB.Connection
    Dim sPathSQL As String, fFile As Integer
    
    ' 1. Preparação
    sPathSQL = sPathDestino & "_Migracao.sql"
    fFile = FreeFile
    Open sPathSQL For Output As #fFile
    
    ' 2. Abrir Conexões
    connOrigem.Open GeraConexao(sPathOrigem)
    connDestino.Open GeraConexao(sPathDestino)
    
    Set catOrigem = CreateObject("ADOX.Catalog")
    catOrigem.ActiveConnection = connOrigem
    
    ' 3. Loop pelas tabelas
    For Each tblOrigem In catOrigem.Tables
        If tblOrigem.Type = "TABLE" And Left(tblOrigem.Name, 4) <> "MSys" Then
            
            ' Verifica no SQLite se a tabela existe
            If Not TabelaExisteNoSQLite(tblOrigem.Name, connDestino) Then
                Print #fFile, "-- Criando tabela ausente: " & tblOrigem.Name
                GerarCreateTabelaSqlite tblOrigem, fFile
            Else
               GerarCamposAusentesSqlite tblOrigem, connDestino, fFile
                Print #fFile, "-- Tabela já existe, ignorando criação: " & tblOrigem.Name
            End If
            
            GerarIndicesSqlite tblOrigem, connDestino, fFile
            
        End If
    Next
    
    Close #fFile
    connOrigem.Close
    connDestino.Close
    
    ' 5. IMPLEMENTAÇÃO DO TERCEIRO PARÂMETRO
    If bExecutar Then
        ' Verifica se o arquivo não está vazio antes de tentar executar
        If FileLen(sPathSQL) > 0 Then
            Call ExecutarSQLite(sPathDestino, sPathSQL)
            MsgBox "Script executado com sucesso no destino!", vbInformation
        Else
            MsgBox "O script está vazio. Nenhuma alteração necessária.", vbInformation
        End If
    Else
        MsgBox "Script gerado com sucesso em: " & sPathSQL, vbInformation
    End If
    
End Sub
Public Sub GerarCamposAusentesSqlite(tblOrigem As Object, connDestino As ADODB.Connection, fFile As Integer)
    Dim colOrigem As Object
    Dim rsCampo As ADODB.Recordset
    Dim ssql As String
    
    ' Percorre cada coluna da tabela origem (Access)
    For Each colOrigem In tblOrigem.Columns
        
        ' Verifica se esta coluna existe no SQLite
        ' O PRAGMA table_info retorna as colunas da tabela no SQLite
        Set rsCampo = connDestino.Execute("PRAGMA table_info([" & tblOrigem.Name & "])")
        
        Dim bCampoExiste As Boolean: bCampoExiste = False
        Do While Not rsCampo.EOF
            If LCase(rsCampo("name").Value) = LCase(colOrigem.Name) Then
                bCampoExiste = True
                Exit Do
            End If
            rsCampo.MoveNext
        Loop
        rsCampo.Close
        
        ' Se o campo NÃO existe no destino, gera o comando ALTER TABLE
        If Not bCampoExiste Then
            ssql = "ALTER TABLE [" & tblOrigem.Name & "] ADD COLUMN [" & colOrigem.Name & "] " & GetTipoSqlite(colOrigem.Type) & ";"
            Print #fFile, ssql
        End If
        
    Next
End Sub
Public Function TabelaExisteNoSQLite(ByVal sTabela As String, ByVal connDestino As ADODB.Connection) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = connDestino.Execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name='" & sTabela & "'")
    
    TabelaExisteNoSQLite = (rs(0).Value > 0)
    rs.Close
End Function
Public Function ObterAssinaturaIndice(idx As Object) As String
    Dim col As Object
    Dim sAssinatura As String
    For Each col In idx.Columns
        sAssinatura = sAssinatura & col.Name & "|"
    Next
    ObterAssinaturaIndice = sAssinatura
End Function
' Função auxiliar para limpar a assinatura (remove colchetes, espaços e normaliza)
Public Function LimparAssinatura(ByVal sCampos As String) As String
    Dim sTemp As String
    sTemp = LCase(sCampos)
    sTemp = Replace(sTemp, "[", "")
    sTemp = Replace(sTemp, "]", "")
    sTemp = Replace(sTemp, " ", "")
    LimparAssinatura = sTemp
End Function
Public Function ExisteIndiceComMesmaAssinatura(tblOrigem As Object, idxOrigem As Object, connDestino As ADODB.Connection) As Boolean
    Dim rs As ADODB.Recordset
    Dim sIdxSql As String
    Dim sAssinaturaOrigem As String
    
    sAssinaturaOrigem = ObterAssinaturaIndice(idxOrigem)
    
    ' Pega todos os índices da tabela no SQLite
    Set rs = connDestino.Execute("PRAGMA index_list([" & tblOrigem.Name & "])")
    
    Do While Not rs.EOF
        ' Pega o SQL de criação do índice (o que mostra as colunas que ele indexa)
        Dim rsInfo As ADODB.Recordset
        Set rsInfo = connDestino.Execute("SELECT sql FROM sqlite_master WHERE name='" & rs("name").Value & "'")
        
        If Not rsInfo.EOF Then
            sIdxSql = LCase(rsInfo("sql").Value)
            ' Verifica se todas as colunas da origem estão na string SQL do índice do destino
            Dim col As Object
            Dim bTodasPresentes As Boolean: bTodasPresentes = True
            For Each col In idxOrigem.Columns
                If InStr(sIdxSql, LCase(col.Name)) = 0 Then
                    bTodasPresentes = False
                    Exit For
                End If
            Next
            
            If bTodasPresentes Then
                ExisteIndiceComMesmaAssinatura = True
                rsInfo.Close
                rs.Close
                Exit Function
            End If
        End If
        rsInfo.Close
        rs.MoveNext
    Loop
    
    rs.Close
    ExisteIndiceComMesmaAssinatura = False
End Function

Public Function GerarNomeUnicoParaIndice(sNomeBase As String, connDestino As ADODB.Connection) As String
    Dim iSeq As Integer
    Dim sNovoNome As String
    Dim rsExiste As ADODB.Recordset
    
    iSeq = 1
    sNovoNome = sNomeBase & "-" & iSeq
    
    ' Loop: verifica em TUDO (tabelas, índices, etc.) se o nome já está em uso
    Do
        ' Removemos o filtro 'type' para checar o banco inteiro
        Set rsExiste = connDestino.Execute("SELECT count(*) FROM sqlite_master WHERE name='" & sNovoNome & "'")
        
        If rsExiste(0).Value = 0 Then
            ' Nome livre de conflitos com qualquer objeto
            Exit Do
        End If
        
        iSeq = iSeq + 1
        sNovoNome = sNomeBase & "-" & iSeq
        rsExiste.Close
    Loop
    
    GerarNomeUnicoParaIndice = sNovoNome
End Function
Public Sub ProcessarPastaCompletaMDBSqlite(ByVal sFolder As String, ByVal bExecutar As Boolean)
    Dim fso As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim sCaminhoMDB As String
    Dim sCaminhoSQLite As String
    Dim sNomeBase As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(sFolder) Then
        MsgBox "Diretório não encontrado!", vbCritical
        Exit Sub
    End If
    
    Set oFolder = fso.GetFolder(sFolder)
    
    ' Itera por todos os arquivos da pasta
    For Each oFile In oFolder.Files
        ' Filtra apenas arquivos Access (.mdb ou .accdb)
        If LCase(fso.GetExtensionName(oFile.Name)) = "mdb" Or _
           LCase(fso.GetExtensionName(oFile.Name)) = "accdb" Then
           
            sCaminhoMDB = oFile.Path
            sNomeBase = fso.GetParentFolderName(sCaminhoMDB) & "\" & fso.GetBaseName(oFile.Name)
            
            ' Constrói o caminho do contraparte (ex: C:\Pasta\vendas.sqlite)
            sCaminhoSQLite = sNomeBase & ".sqlite"
            
            ' Verifica se o par existe antes de chamar a migração
            If fso.FileExists(sCaminhoSQLite) Then
                Debug.Print "Migrando: " & oFile.Name & " -> " & fso.GetFileName(sCaminhoSQLite)
                
                ' Chama sua rotina principal de comparação
                Call GerarScriptAlteracoesAccessParaSqlite(sCaminhoMDB, sCaminhoSQLite, bExecutar)
            Else
                Debug.Print "Contraparte não encontrado para: " & oFile.Name
            End If
            
        End If
    Next oFile
    
    MsgBox "Processamento em lote concluído!", vbInformation
    
    Set oFolder = Nothing
    Set fso = Nothing
End Sub

Public Sub RenomearIndice(ByVal sPathMDB As String, ByVal sTabelaAlvo As String, ByVal sNomeAntigo As String, ByVal sNomeNovo As String)
    Dim conn As Object, cat As Object, tbl As Object, idx As Object
    Dim bTabelaEncontrada As Boolean
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open GeraConexao(sPathMDB)
    
    Set cat = CreateObject("ADOX.Catalog")
    Set cat.ActiveConnection = conn
    
    ' 1. A MESMA LÓGICA DE BUSCA QUE FUNCIONA NA GERARINFO
    bTabelaEncontrada = False
    For Each tbl In cat.Tables
        ' Compara o nome da mesma forma que você faz no relatório
        If tbl.Type = "TABLE" Then
            If LCase(Trim(tbl.Name)) = LCase(Trim(sTabelaAlvo)) Then
                bTabelaEncontrada = True
                Exit For
            End If
        End If
    Next
    
    If Not bTabelaEncontrada Then
        MsgBox "Tabela '" & sTabelaAlvo & "' não encontrada na coleção de tabelas.", vbCritical
        GoTo Finalizar
    End If
    
    ' 2. A MESMA LÓGICA DE BUSCA DE ÍNDICE
    Dim bEncontrouIdx As Boolean: bEncontrouIdx = False
    For Each idx In tbl.Indexes
        If LCase(Trim(idx.Name)) = LCase(Trim(sNomeAntigo)) Then
            bEncontrouIdx = True
            Exit For
        End If
    Next
    
    If Not bEncontrouIdx Then
        MsgBox "Índice '" & sNomeAntigo & "' não encontrado nesta tabela.", vbExclamation
        GoTo Finalizar
    End If
    
    ' 3. COLETA OS DADOS (Campos)
    Dim sCampos As String, col As Object
    For Each col In idx.Columns
        sCampos = sCampos & "[" & col.Name & "], "
    Next
    sCampos = Left(sCampos, Len(sCampos) - 2) ' Remove a última vírgula
    
    ' 4. AÇÃO: DROP E CREATE (SQL DIRETO)
    On Error Resume Next
    conn.Execute "DROP INDEX [" & sNomeAntigo & "] ON [" & sTabelaAlvo & "]"
    
    Dim ssql As String
    ssql = "CREATE " & IIf(idx.Unique, "UNIQUE ", "") & _
           "INDEX [" & sNomeNovo & "] ON [" & sTabelaAlvo & "] (" & sCampos & ")"
    
    conn.Execute ssql
    
    If Err.Number = 0 Then
        MsgBox "Índice renomeado com sucesso via SQL!", vbInformation
    Else
        MsgBox "Erro ao executar SQL: " & Err.Description, vbCritical
    End If
    On Error GoTo 0
    
Finalizar:
    conn.Close
    Set cat = Nothing
    Set conn = Nothing
End Sub

' Esta função limpa o SQL do SQLite e retorna os campos: "[campo1], [campo2]"
Public Function ExtrairCamposDoSqlite(ByVal sSqlIndex As Variant) As String
    If IsNull(sSqlIndex) Then Exit Function
    
    Dim sClean As String
    sClean = CStr(sSqlIndex)
    
    ' 1. Remove quebras de linha e tabulações que podem confundir o processamento
    sClean = Replace(sClean, vbCr, " ")
    sClean = Replace(sClean, vbLf, " ")
    sClean = Replace(sClean, vbTab, " ")
    
    ' 2. Extrai o conteúdo entre parênteses
    Dim iInicio As Integer, iFim As Integer
    iInicio = InStr(sClean, "(")
    iFim = InStrRev(sClean, ")")
    
    If iInicio > 0 And iFim > iInicio Then
        ExtrairCamposDoSqlite = Mid(sClean, iInicio + 1, iFim - iInicio - 1)
    End If
End Function

' Esta função gera a assinatura do índice do MDB para comparar
Public Function GerarAssinaturaMdb(idxMdb As Object) As String
    Dim col As Object, sAssin As String
    For Each col In idxMdb.Columns
        sAssin = sAssin & "[" & LCase(col.Name) & "], "
    Next
    ' Remove a última vírgula e espaço
    If Len(sAssin) > 2 Then sAssin = Left(sAssin, Len(sAssin) - 2)
    GerarAssinaturaMdb = sAssin
End Function

Public Function IndiceJaExiste(tblOrigem As Object, idxOrigem As Object, connDestino As ADODB.Connection) As Boolean
    Dim rs As ADODB.Recordset
    Dim sAssinaturaMdb As String
    Dim sAssinaturaSqlite As String
    
    sAssinaturaMdb = GerarAssinaturaMdb(idxOrigem)
    
    ' Busca os índices da tabela destino no sqlite_master
    Set rs = connDestino.Execute("SELECT sql FROM sqlite_master WHERE type='index' AND tbl_name='" & tblOrigem.Name & "'")
    
    Do While Not rs.EOF
        sAssinaturaSqlite = ExtrairCamposDoSqlite(rs("sql").Value)
        
        ' Compara se a assinatura do MDB está contida na do SQLite
        ' Remove espaços extras para garantir a igualdade
        If Replace(sAssinaturaMdb, " ", "") = Replace(sAssinaturaSqlite, " ", "") Then
            IndiceJaExiste = True
            Exit Do
        End If
        rs.MoveNext
    Loop
    
    rs.Close
    IndiceJaExiste = IndiceJaExiste
End Function
Public Sub GerarIndicesSqlite(tblOrigem As Object, connDestino As ADODB.Connection, fFile As Integer)
    Dim idx As Object, col As Object
    Dim sCamposMdb As String, ssql As String
    Dim bExiste As Boolean
    Dim rs As ADODB.Recordset
    
    For Each idx In tblOrigem.Indexes
        If Not idx.PrimaryKey Then
            ' 1. Assinatura do MDB
            sCamposMdb = ""
            For Each col In idx.Columns
                If sCamposMdb <> "" Then sCamposMdb = sCamposMdb & ", "
                sCamposMdb = sCamposMdb & "[" & col.Name & "]"
            Next
            
            ' 2. Verifica índices existentes no SQLite
            bExiste = False
            'Dim ssql As String
            ssql = "SELECT replace(replace(replace(sql, char(10), ' '), char(13), ' '), char(9), ' ') AS sql " & _
               "FROM sqlite_master WHERE type='index' AND tbl_name='" & tblOrigem.Name & "'"
            Set rs = connDestino.Execute(ssql)
            
            ' Se o Recordset estiver vazio (EOF), o Loop nem começa, bExiste continua False.
            ' O código entende que não existe índice e vai para o passo 3 (CRIAÇÃO).
            Do While Not rs.EOF
                If Not IsNull(rs("sql").Value) Then
                    Dim sSqliteCampos As String
                    Dim SSQQLUSO As String
                    SSQQLUSO = rs("sql").Value
                    sSqliteCampos = ExtrairCamposDoSqlite(SSQQLUSO)
                    sCamposMdb = LimparAssinatura(sCamposMdb)
                    sSqliteCampos = LimparAssinatura(sSqliteCampos)
                    ' Compara assinaturas
                    If Replace(LCase(sCamposMdb), " ", "") = Replace(LCase(sSqliteCampos), " ", "") Then
                        bExiste = True
                        Exit Do
                    End If
                End If
                rs.MoveNext
            Loop
            rs.Close
            
            ' 3. Criação
            If Not bExiste Then
                ssql = "CREATE "
                If idx.Unique Then ssql = ssql & "UNIQUE "
                ssql = ssql & "INDEX [IDX_" & tblOrigem.Name & "_" & idx.Name & "] " & _
                       "ON [" & tblOrigem.Name & "] (" & sCamposMdb & ");"
                
                Print #fFile, ssql
            End If
        End If
    Next
End Sub

Public Sub ProcessarPastaCompletaMDBaccdb(ByVal sFolder As String, ByVal bExecutar As Boolean)
    Dim fso As Object
    Dim oFolder As Object
    Dim oFile As Object
    Dim sCaminhoMDB As String
    Dim sCaminhoaccdb As String
    Dim sNomeBase As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(sFolder) Then
        MsgBox "Diretório não encontrado!", vbCritical
        Exit Sub
    End If
    
    Set oFolder = fso.GetFolder(sFolder)
    
    ' Itera por todos os arquivos da pasta
    For Each oFile In oFolder.Files
        ' Filtra apenas arquivos Access (.mdb ou .accdb)
        If LCase(fso.GetExtensionName(oFile.Name)) = "mdb" Then
            sCaminhoMDB = oFile.Path
            sNomeBase = fso.GetParentFolderName(sCaminhoMDB) & "\" & fso.GetBaseName(oFile.Name)
            
            ' Constrói o caminho do contraparte (ex: C:\Pasta\vendas.sqlite)
            sCaminhoaccdb = sNomeBase & ".accdb"
            
            ' Verifica se o par existe antes de chamar a migração
            If fso.FileExists(sCaminhoaccdb) Then
                Debug.Print "Migrando: " & oFile.Name & " -> " & fso.GetFileName(sCaminhoaccdb)
                
                ' Chama sua rotina principal de comparação
                Call GerarScriptAlteracoes(sCaminhoMDB, sCaminhoaccdb, bExecutar)
            Else
                Debug.Print "Contraparte não encontrado para: " & oFile.Name
            End If
            
        End If
    Next oFile
    
    MsgBox "Processamento em lote concluído!", vbInformation
    
    Set oFolder = Nothing
    Set fso = Nothing
End Sub

