VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form dCompare 
   Caption         =   "Database Comparador/Corretor"
   ClientHeight    =   7908
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   13248
   Icon            =   "dcompare.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7908
   ScaleWidth      =   13248
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdGeraPastaSql 
      Caption         =   "SQL"
      Height          =   372
      Left            =   1200
      TabIndex        =   27
      Top             =   3360
      Width           =   852
   End
   Begin VB.CheckBox CheckEfetua 
      Caption         =   "Efetua  correcao"
      Height          =   252
      Left            =   5160
      TabIndex        =   26
      Top             =   1560
      Width           =   1572
   End
   Begin VB.CommandButton CmdEscdir 
      Caption         =   "..."
      Height          =   372
      Left            =   6360
      TabIndex        =   25
      Top             =   3000
      Width           =   492
   End
   Begin VB.CommandButton CmdGeraPasta 
      Caption         =   "Schema"
      Height          =   372
      Left            =   240
      TabIndex        =   24
      Top             =   3360
      Width           =   852
   End
   Begin VB.TextBox Diretorio 
      Height          =   372
      Left            =   240
      TabIndex        =   23
      Top             =   3000
      Width           =   6012
   End
   Begin VB.CommandButton cmdSelSQL 
      Caption         =   "..."
      Height          =   315
      Left            =   7800
      TabIndex        =   21
      Top             =   2520
      Width           =   495
   End
   Begin VB.TextBox txtSQL 
      Height          =   315
      Left            =   1200
      TabIndex        =   20
      Top             =   2520
      Width           =   6495
   End
   Begin VB.CommandButton CmdSchemaDestino 
      Caption         =   "Recria do Schema"
      Height          =   372
      Left            =   3480
      TabIndex        =   19
      Top             =   1920
      Width           =   1452
   End
   Begin VB.CommandButton CmdSchemaOrigem 
      Caption         =   "Schema"
      Height          =   372
      Left            =   1800
      TabIndex        =   18
      Top             =   480
      Width           =   732
   End
   Begin VB.CommandButton CmdSqlDestino 
      Caption         =   "sql"
      Height          =   372
      Left            =   1080
      TabIndex        =   17
      Top             =   1440
      Width           =   612
   End
   Begin VB.CommandButton CmdSqlOrigem 
      Caption         =   "sql"
      Height          =   372
      Left            =   1080
      TabIndex        =   16
      Top             =   480
      Width           =   612
   End
   Begin VB.CheckBox CheckIncDados 
      Caption         =   "Incluir dados"
      Height          =   252
      Left            =   6960
      TabIndex        =   15
      Top             =   1560
      Width           =   1332
   End
   Begin VB.TextBox txtArquivoDestino 
      Height          =   375
      Left            =   1080
      TabIndex        =   14
      Top             =   1080
      Width           =   7215
   End
   Begin VB.TextBox TxTArquivoOrigem 
      Height          =   375
      Left            =   1080
      TabIndex        =   13
      Top             =   120
      Width           =   7215
   End
   Begin VB.TextBox TxtInfo 
      Height          =   2808
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Top             =   3960
      Width           =   8055
   End
   Begin VB.CommandButton cmdCompactar 
      Caption         =   "Compactar"
      Height          =   372
      Left            =   3600
      TabIndex        =   11
      Top             =   480
      Width           =   1332
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Criar Tabela de Destino"
      Height          =   372
      Left            =   1080
      TabIndex        =   10
      Top             =   1920
      Width           =   2172
   End
   Begin VB.CommandButton Encerrar 
      Caption         =   "Encerrar"
      Height          =   375
      Left            =   9480
      TabIndex        =   9
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton CmdTableInfo 
      Caption         =   "TableInfo"
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton CmdExportarSqlite 
      Caption         =   "Exportar Sqlite"
      Height          =   375
      Left            =   5040
      TabIndex        =   7
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton CmdCompareSchema 
      Caption         =   "Comparar Shema"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   1440
      Width           =   1452
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Compare Arquivos"
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1440
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8640
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdEscolheDestino 
      Caption         =   "Destino"
      Height          =   372
      Left            =   8400
      TabIndex        =   4
      Top             =   1080
      Width           =   852
   End
   Begin VB.CommandButton CmdEscolheOrigem 
      Caption         =   "Origem"
      Height          =   372
      Left            =   8400
      TabIndex        =   3
      Top             =   120
      Width           =   852
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "SqlScript"
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
      Height          =   252
      Left            =   240
      TabIndex        =   22
      Top             =   2520
      Width           =   792
   End
   Begin VB.Label lProgress 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   372
      Left            =   5160
      TabIndex        =   2
      Top             =   1920
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
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   792
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

Private Sub Check1_Click()

End Sub

Private Sub cmdCompactar_Click()
Dim sCaminho As String
    sCaminho = txtArquivoDestino.Text ' Assumindo que o alvo da compactação é o destino
    
    If EArquivoSQLite(sCaminho) Then
        If CompactarSQLite(sCaminho) Then
            MsgBox "SQLite compactado com sucesso (VACUUM executado).", vbInformation
        End If
    Else
        ' Chama a sua função antiga de compactação do Access
        If CompactAccess(sCaminho) Then
            MsgBox "Access compactado com sucesso (JRO executado).", vbInformation
        End If
    End If
End Sub

Private Sub CmdCompareSchema_Click()
GerarScriptAlteracoesViaSchema TxTArquivoOrigem.Text, txtArquivoDestino.Text, (CheckEfetua.Value = vbChecked)
End Sub

Private Sub CmdEscdir_Click()
Dim sCaminho As String
    sCaminho = SelecionarPasta("Selecione a pasta contendo dos arquivos MDB:")
    
    If sCaminho <> "" Then
        Diretorio.Text = sCaminho
    End If
End Sub

Private Sub CmdEscolheDestino_Click()
    On Error GoTo Erro
    
    ' Configura o CommonDialog
    With CommonDialog1
        .Filter = "Bancos de Dados (*.mdb;*.accdb;*.sqlite;*.db;*.db3;*.fossil)|*.mdb;*.accdb;*.sqlite;*.db;*.db3;*.fossil|" & _
                  "Access (*.mdb;*.accdb)|*.mdb;*.accdb|" & _
                  "SQLite (*.sqlite;*.db;*.db3;*.fossil)|*.sqlite;*.db;*.db3;*.fossil"
        .FilterIndex = 1
        .DialogTitle = "Selecione ou crie o ficheiro de Destino"
       '.Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist ' Avisa se sobrescrever
        .filename = ""
        .ShowSave ' Usamos ShowSave em vez de ShowOpen para permitir indicar novos nomes
    End With
    
    ' Grava o caminho no TextBox
    If CommonDialog1.filename <> "" Then
        txtArquivoDestino.Text = CommonDialog1.filename
    End If
    
    If Not ValidarOuCriarDestino(txtArquivoDestino.Text) Then
        MsgBox "O destino é inválido ou não foi criado."
        Exit Sub
    End If
    
    Exit Sub
Erro:
    If Err.Number <> 32755 Then ' Ignora se carregar em Cancelar
        MsgBox "Erro ao selecionar destino: " & Err.Description, vbCritical
    End If
End Sub

Private Sub CmdEscolheOrigem_Click()
    On Error GoTo Erro
    
    CommonDialog1.Filter = "Bancos de Dados (*.mdb;*.accdb;*.sqlite;*.db;*.db3;*.fossil)|*.mdb;*.accdb;*.sqlite;*.db;*.db3;*.fossil|" & _
                  "Access (*.mdb;*.accdb)|*.mdb;*.accdb|" & _
                  "SQLite (*.sqlite;*.db;*.db3;*.fossil)|*.sqlite;*.db;*.db3;*.fossil|" & _
                  "Todos os Ficheiros (*.*)|*.*"
                           
    CommonDialog1.FilterIndex = 1
    CommonDialog1.DialogTitle = "Selecione o arquivo de Origem"
    CommonDialog1.filename = ""
    CommonDialog1.ShowOpen
    
    ' Verifica se o usuário selecionou um arquivo e cancelou
    If CommonDialog1.filename <> "" Then
        TxTArquivoOrigem.Text = CommonDialog1.filename
    End If
    
    Exit Sub
Erro:
    ' Captura o erro caso o usuário aperte Cancelar
    If Err.Number <> 32755 Then ' 32755 é o código do botão cancelar
        MsgBox "Erro ao selecionar arquivo: " & Err.Description, vbCritical
    End If
End Sub

Private Sub CmdExportarSqlite_Click()
    Dim sOrigem As String
    Dim sDestino As String
    
    sOrigem = Trim(TxTArquivoOrigem.Text)
    
    ' 1. Valida se a origem existe
    If sOrigem = "" Or Not ArquivoExiste(sOrigem) Then
        MsgBox "Selecione um arquivo de origem válido!", vbExclamation
        Exit Sub
    End If
    
    ' 2. Define o destino automaticamente se o campo estiver vazio
    sDestino = Trim(txtArquivoDestino.Text)
    If sDestino = "" Then
        ' Troca a extensão para .sqlite mantendo o mesmo caminho e nome
        sDestino = Left(sOrigem, InStrRev(sOrigem, ".")) & "sqlite"
        txtArquivoDestino.Text = sDestino
    End If
    
    ' 3. Executa a conversão
    If MDG("Converter e sobrescrever destino?") Then
        Convert sOrigem, sDestino
    End If
End Sub



Private Sub CmdGeraPasta_Click()
    ProcessarPastaCompletaSchema (Diretorio.Text)
End Sub

Private Sub CmdGeraPastaSql_Click()
   ProcessarPastaCompletaSql Diretorio.Text, (CheckIncDados.Value = vbChecked)
End Sub

Private Sub CmdSchemaDestino_Click()
    Dim sPathSchema As String
    Dim sPathRecriado As String
    Dim sPasta As String
    Dim sNomeBase As String
    Dim sEXT As String
    ' O Schema está onde o usuário apontou em TxTArquivoOrigem
    sPathSchema = Me.TxTArquivoOrigem.Text
    
    If Dir(sPathSchema) = "" Then
        MsgBox "O arquivo de Schema não foi encontrado no caminho especificado!", vbCritical
        Exit Sub
    End If
    
    
    sPathRecriado = Me.txtArquivoDestino.Text
    If Me.txtArquivoDestino.Text = "" Then
        ' Define o nome do destino baseado no arquivo de schema
        sPasta = parsefile(sPathSchema, "C") 'Left(sPathSchema, InStrRev(sPathSchema, "\"))
        sEXT = EXTENSAO(sPathSchema)
        sPathRecriado = sPasta & "_recriado." & sEXT
        sPathRecriado = Replace(sPathRecriado, "_schema", "")
        
    '    sNomeBase = Mid(sPathSchema, InStrRev(sPathSchema, "\") + 1)
        
        ' Remove a extensão para criar o novo nome
     '   sNomeBase = Left(sNomeBase, InStrRev(sNomeBase, ".") - 1)
        
        ' Define o destino como o mesmo nome + _recriado.mdb
      '  sPathRecriado = sPasta & sNomeBase & "_recriado.mdb"
        
        Me.txtArquivoDestino.Text = sPathRecriado
    End If
    
    ' Executa a reconstrução
    Call RecriarBancoDoSchema(sPathSchema, sPathRecriado)
    Call RecriarIndicesDoSchema(sPathSchema, sPathRecriado)
    
End Sub

Private Sub CmdSchemaOrigem_Click()
    Dim sCaminhoOrigem As String
    Dim sCaminhoDestino As String
    Dim sPasta As String
    Dim sNomeArquivo As String
    
    sCaminhoOrigem = Me.TxTArquivoOrigem.Text
    
    If sCaminhoOrigem = "" Then
        MsgBox "Selecione o arquivo de origem primeiro!", vbExclamation
        Exit Sub
    End If
    
    ' 1. Extrai o caminho da pasta e o nome do arquivo (sem a extensão)
    sPasta = Left(sCaminhoOrigem, InStrRev(sCaminhoOrigem, "\"))
    sNomeArquivo = Mid(sCaminhoOrigem, InStrRev(sCaminhoOrigem, "\") + 1)
    
    ' Remove a extensão (.mdb ou .accdb)
    sNomeArquivo = Left(sNomeArquivo, InStrRev(sNomeArquivo, ".") - 1)
    
    ' 2. Monta o caminho do destino com _schema.mdb
    sCaminhoDestino = sPasta & sNomeArquivo & "_schema.mdb"
    
    ' 3. Atualiza o campo de destino (opcional)
    Me.txtArquivoDestino.Text = sCaminhoDestino
    
    GarantirSchemaExistente sCaminhoDestino
    ' 4. Chama a rotina que criamos anteriormente
    Call GerarArquivoSchemaADO(sCaminhoOrigem, sCaminhoDestino)
'    Call RecriarIndicesDoSchema(sPathSchema, sPathRecriado)
End Sub

Private Sub cmdSelSQL_Click()
  On Error Resume Next
    CommonDialog1.Filter = "Script SQL (*.sql)|*.sql"
    CommonDialog1.filename = "correcao_indices.sql"
    CommonDialog1.ShowSave
    If CommonDialog1.filename <> "" Then txtSQL.Text = CommonDialog1.filename
End Sub

Private Sub CmdTableInfo_Click()
' Verifica se um arquivo foi selecionado
    If Trim(TxTArquivoOrigem.Text) = "" Then
        MsgBox "Selecione o arquivo de origem primeiro!", vbExclamation
        Exit Sub
    End If
    
 If EArquivoSQLite(TxTArquivoOrigem.Text) Then
        GerarInfoTabelasSQLite TxTArquivoOrigem.Text, TxtInfo
    Else
        ' Chama sua rotina antiga para Access
        GerarInfoTabelasAccess TxTArquivoOrigem.Text, TxtInfo
    End If
End Sub

Private Sub Command3_Click()
If EArquivoSQLite(txtArquivoDestino.Text) Then
   GerarScriptAlteracoesAccessParaSqlite TxTArquivoOrigem.Text, txtArquivoDestino.Text, (CheckEfetua.Value = vbChecked)
Else
   GerarScriptAlteracoes TxTArquivoOrigem.Text, txtArquivoDestino.Text, (CheckEfetua.Value = vbChecked)
End If
End Sub

Private Sub Command5_Click()
If Not ValidarOuCriarDestino(txtArquivoDestino.Text) Then
        MsgBox "O destino é inválido ou não foi criado."
        Exit Sub
    End If
    
End Sub



Private Sub cmdSqlOrigem_Click()
    ' Gera o script baseado no arquivo selecionado na origem
    GerarScriptsSQL TxTArquivoOrigem.Text, (CheckIncDados.Value = vbChecked)
End Sub

Private Sub cmdSqlDestino_Click()
    ' Gera o script baseado no arquivo selecionado no destino
    GerarScriptsSQL txtArquivoDestino.Text, (CheckIncDados.Value = vbChecked)
End Sub

