VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form dCompare 
   Caption         =   "Database Comparador/Corretor"
   ClientHeight    =   5808
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11892
   Icon            =   "dcompare.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5808
   ScaleWidth      =   11892
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSqlDestino 
      Caption         =   "sql"
      Height          =   252
      Left            =   9600
      TabIndex        =   17
      Top             =   720
      Width           =   612
   End
   Begin VB.CommandButton CmdSqlOrigem 
      Caption         =   "sql"
      Height          =   252
      Left            =   9600
      TabIndex        =   16
      Top             =   240
      Width           =   612
   End
   Begin VB.CheckBox CheckIncDados 
      Caption         =   "Incluir dados"
      Height          =   252
      Left            =   6240
      TabIndex        =   15
      Top             =   1200
      Width           =   1692
   End
   Begin VB.TextBox txtArquivoDestino 
      Height          =   375
      Left            =   1080
      TabIndex        =   14
      Top             =   600
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
   Begin VB.CommandButton CmdTableInfo 
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
   Begin VB.CommandButton CmdEscolheDestino 
      Caption         =   "Destino"
      Height          =   255
      Left            =   8400
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton CmdEscolheOrigem 
      Caption         =   "Origem"
      Height          =   255
      Left            =   8400
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
Public Function EArquivoSQLite(ByVal cCaminho As String) As Boolean
    Dim sExt As String
    sExt = LCase(Right(cCaminho, 6)) ' Pega a extensão final
    
    ' Lista de extensões que o seu sistema aceita como SQLite
    If InStr(cCaminho, ".sqlite") > 0 Or _
       InStr(cCaminho, ".db") > 0 Or _
       InStr(cCaminho, ".db3") > 0 Or _
       InStr(cCaminho, ".fossil") > 0 Then
       
       EArquivoSQLite = True
    Else
       EArquivoSQLite = False
    End If
End Function
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
        .FileName = ""
        .ShowSave ' Usamos ShowSave em vez de ShowOpen para permitir indicar novos nomes
    End With
    
    ' Grava o caminho no TextBox
    If CommonDialog1.FileName <> "" Then
        txtArquivoDestino.Text = CommonDialog1.FileName
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
    CommonDialog1.FileName = ""
    CommonDialog1.ShowOpen
    
    ' Verifica se o usuário selecionou um arquivo e cancelou
    If CommonDialog1.FileName <> "" Then
        TxTArquivoOrigem.Text = CommonDialog1.FileName
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

Public Function MDG(ByVal cMEnSSAGEM As String, Optional cTITULO = "Confirme")
  Dim eRESP As Variant
  MDG = False
  eRESP = MsgBox(cMEnSSAGEM, vbYesNo + vbDefaultButton1, cTITULO)
  If eRESP = vbYes Then
    MDG = True
  End If
End Function

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

Private Sub Command5_Click()
If Not ValidarOuCriarDestino(txtArquivoDestino.Text) Then
        MsgBox "O destino é inválido ou não foi criado."
        Exit Sub
    End If
    
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


Private Sub cmdSqlOrigem_Click()
    ' Gera o script baseado no arquivo selecionado na origem
    GerarScriptsSQL TxTArquivoOrigem.Text, (CheckIncDados.Value = vbChecked)
End Sub

Private Sub cmdSqlDestino_Click()
    ' Gera o script baseado no arquivo selecionado no destino
    GerarScriptsSQL txtArquivoDestino.Text, (CheckIncDados.Value = vbChecked)
End Sub


