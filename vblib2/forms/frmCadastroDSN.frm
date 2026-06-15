VERSION 5.00
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "vbccr18.ocx"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form frmCadastroDSN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assistente de Configuração de DSN"
   ClientHeight    =   6108
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8556
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6108
   ScaleWidth      =   8556
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbConfig 
      Height          =   288
      Left            =   360
      TabIndex        =   23
      Text            =   "Combo1"
      Top             =   360
      Width           =   6012
   End
   Begin VB.CommandButton CmdConfigurarPeloIni 
      Caption         =   "Configurar Pelo INI"
      Height          =   492
      Left            =   6720
      TabIndex        =   20
      Top             =   4800
      Width           =   1572
   End
   Begin VB.ComboBox cboDriver 
      Height          =   288
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   840
      Width           =   4455
   End
   Begin VB.TextBox txtDSN 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   1320
      Width           =   4455
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   1800
      Width           =   4455
   End
   Begin VB.TextBox txtDB 
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   2280
      Width           =   4455
   End
   Begin VB.TextBox txtHost 
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   2760
      Width           =   4455
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox txtPass 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   3480
      Width           =   1935
   End
   Begin VB.TextBox txtPort 
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CheckBox chkUserDSN 
      Caption         =   "Criar apenas para o Usuário Atual (User DSN)"
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   4320
      Width           =   4455
   End
   Begin VB.CommandButton btnSalvar 
      Caption         =   "&Salvar e Registrar DSN"
      Height          =   495
      Left            =   2880
      TabIndex        =   10
      Top             =   4800
      Width           =   2415
   End
   Begin VB.CommandButton btnTestar 
      Caption         =   "&Testar Conexão"
      Height          =   495
      Left            =   720
      TabIndex        =   11
      Top             =   4800
      Width           =   1815
   End
   Begin VBCCR18.CommandButtonW btnclose 
      Height          =   612
      Left            =   7080
      TabIndex        =   19
      Top             =   360
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Retornar"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "frmCadastroDSN.frx":0000
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW cmdRevelar 
      Height          =   612
      Left            =   4320
      TabIndex        =   22
      Top             =   3360
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Caption         =   "Exibir Senha"
      Picture         =   "frmCadastroDSN.frx":059A
      PictureAndCaption=   -1  'True
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW CmdEscdir 
      Height          =   372
      Left            =   7200
      TabIndex        =   24
      Top             =   2280
      Width           =   612
      _ExtentX        =   1080
      _ExtentY        =   656
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "frmCadastroDSN.frx":0A34
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW CmdEscolheOrigem 
      Height          =   372
      Left            =   6480
      TabIndex        =   25
      Top             =   2280
      Width           =   612
      _ExtentX        =   1080
      _ExtentY        =   656
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "frmCadastroDSN.frx":0FCE
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin vbExtra.CommonDialogEx CommonDialog1 
      Left            =   7800
      Top             =   3000
      _ExtentX        =   720
      _ExtentY        =   720
      MaxFileSize     =   255
      FontName        =   ""
   End
   Begin VB.Label lblStatus 
      Caption         =   "Label1"
      Height          =   372
      Left            =   480
      TabIndex        =   21
      Top             =   5520
      Width           =   5892
   End
   Begin VB.Label lblDriver 
      Alignment       =   1  'Right Justify
      Caption         =   "Tipo de Driver:"
      Height          =   252
      Left            =   240
      TabIndex        =   12
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label lblDSN 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome do DSN:"
      Height          =   252
      Left            =   240
      TabIndex        =   13
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Label lblDesc 
      Alignment       =   1  'Right Justify
      Caption         =   "Descrição:"
      Height          =   252
      Left            =   240
      TabIndex        =   14
      Top             =   1800
      Width           =   1500
   End
   Begin VB.Label lblDB 
      Alignment       =   1  'Right Justify
      Caption         =   "Banco / Arquivo:"
      Height          =   252
      Left            =   240
      TabIndex        =   15
      Top             =   2280
      Width           =   1500
   End
   Begin VB.Label lblHost 
      Alignment       =   1  'Right Justify
      Caption         =   "Servidor (Host):"
      Height          =   252
      Left            =   360
      TabIndex        =   16
      Top             =   2760
      Width           =   1500
   End
   Begin VB.Label lblUser 
      Alignment       =   1  'Right Justify
      Caption         =   "Usuário (UID):"
      Height          =   252
      Left            =   240
      TabIndex        =   17
      Top             =   3120
      Width           =   1500
   End
   Begin VB.Label lblPass 
      Alignment       =   1  'Right Justify
      Caption         =   "Senha (PWD):"
      Height          =   252
      Left            =   240
      TabIndex        =   18
      Top             =   3480
      Width           =   1500
   End
   Begin VB.Label lblPort 
      Alignment       =   1  'Right Justify
      Caption         =   "Porta (Opcional):"
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   3840
      Width           =   1500
   End
End
Attribute VB_Name = "frmCadastroDSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnClose_Click()
 Unload Me
End Sub

Private Sub CmdConfigurarPeloIni_Click()
   configuraodbc
End Sub

Private Sub CmdEscdir_Click()
Dim sCaminho As String
'    sCaminho = SelecionarPasta("Selecione a pasta contendo dos arquivos MDB:")
    CommonDialog1.ShowFolder
    sCaminho = CommonDialog1.FolderName
    
   If sCaminho <> "" Then
        txtDB.Text = sCaminho
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
        txtDB.Text = CommonDialog1.filename
    End If
    
        
    Exit Sub
Erro:
    ' Captura o erro caso o usuário aperte Cancelar
    If Err.Number <> 32755 Then ' 32755 é o código do botão cancelar
        MsgBox "Erro ao selecionar arquivo: " & Err.Description, vbCritical
    End If
End Sub

Private Sub cmdRevelar_Click()
  If txtPass.PasswordChar = "*" Then
        txtPass.PasswordChar = ""
        cmdRevelar.Caption = "Ocultar"
    Else
        txtPass.PasswordChar = "*"
        cmdRevelar.Caption = "Ver"
    End If
    txtPass.SetFocus
End Sub

Private Sub Form_Load()
    ' Alimenta a ComboBox com os drivers mapeados na biblioteca mesclada
    With cboDriver
        .AddItem "MDB"
        .AddItem "ACCDB"
        .AddItem "DBF"
        .AddItem "SQLITE"
        .AddItem "MSSQL"
        .AddItem "MYSQL"
        .AddItem "MARIADB"
        .AddItem "PGSQL"
        .AddItem "ORACLE"
        .ListIndex = 0 ' Define MDB como padrão
    End With
    txtPort.Text = ""
    CarregarComboIni
End Sub
Public Sub configuraodbc()
  Dim lCONTINUA As Boolean
  Dim cCFG As String
  Dim aDAD As Variant
  Dim x As Integer
  lCONTINUA = True
  x = 1
  While lCONTINUA
    cCFG = FixStr(PegPath("ODBCCFG", "ODBCCFG" & StrZero(x, 3)))
    If InStr(cCFG, "|") = 0 Then
      lCONTINUA = False
    Else
      aDAD = Split(cCFG, "|")
      AddDSN aDAD(0), aDAD(1), aDAD(2), aDAD(3), True
    End If
    x = x + 1
  Wend

End Sub

Private Sub cmbConfig_Click()
    Dim LinhaCompleta As String
    Dim PosIgual As Long
    Dim DadosDireita As String
    Dim aDAD As Variant
    Dim i As Integer
    
    If cmbConfig.ListIndex = -1 Then Exit Sub
    
    LinhaCompleta = cmbConfig.Text
    PosIgual = InStr(LinhaCompleta, "=")
    
    If PosIgual > 0 Then
        ' Separa apenas a parte com os Pipes "|" (aDAD)
        DadosDireita = Mid(LinhaCompleta, PosIgual + 1)
        aDAD = Split(DadosDireita, "|")
        
        ' Garante que a linha possui os 4 parâmetros obrigatórios da sua estrutura
        If UBound(aDAD) >= 3 Then
            txtDSN.Text = Trim(aDAD(0))          ' nomedsn
            txtDescription.Text = Trim(aDAD(1))  ' descricao
            txtDB.Text = Trim(aDAD(2))           ' banco ou caminho
            
            Dim TIPO As String
            TIPO = UCase(Trim(aDAD(3)))          ' tipo (MDB, SQLITE, DBF...)
            
            ' Percorre o combo dos tipos para posicioná-lo corretamente
            For i = 0 To cboDriver.ListCount - 1
                If UCase(cboDriver.List(i)) = TIPO Then
                    cboDriver.ListIndex = i
                    Exit For
                End If
            Next i
        End If
    End If
End Sub
' --- CARREGA DADOS DO INI UTILIZANDO A SUA LOGICA DO WHILE/WEND ---
Private Sub CarregarComboIni()
    Dim lCONTINUA As Boolean
    Dim cCFG As String
    Dim x As Integer
    
    cmbConfig.Clear
    lCONTINUA = True
    x = 1
    
    While lCONTINUA
        ' Utiliza as suas funções nativas do sistema para extrair a linha do wrpt.ini
        cCFG = FixStr(PegPath("ODBCCFG", "ODBCCFG" & StrZero(x, 3)))
        
        If InStr(cCFG, "|") = 0 Then
            lCONTINUA = False
        Else
            ' Adiciona a string inteira na ComboBox para podermos desmembrá-la no clique
            cmbConfig.AddItem "ODBCCFG" & StrZero(x, 3) & "=" & cCFG
        End If
        x = x + 1
    Wend
    
    ' Se carregou itens, deixa o primeiro selecionado por padrão
    If cmbConfig.ListCount > 0 Then cmbConfig.ListIndex = 0
End Sub
' --- EVENTO AO SAIR DO FOCO DO CAMPO NOME DO DSN ---
Private Sub txtDSN_LostFocus()
    Dim DsnDigitado As String
    Dim i As Integer
    
    DsnDigitado = UCase(Trim$(txtDSN.Text))
    
    ' Se o campo estiver vazio, não faz nada
    If DsnDigitado = "" Then Exit Sub
    
    ' Usamos o parâmetro "Server" ou "Database" para testar se este DSN já tem dados no cofre
    ' Nota: Se no seu cofre o nome gravado for o banco, certifique-se de que digita o mesmo nome no DSN
    If LerDoCofre(DsnDigitado, "Server") <> "" Or LerDoCofre(DsnDigitado, "Database") <> "" Then
        
        ' 1. Recupera a descrição e o banco/arquivo
        txtDescription.Text = LerDoCofre(DsnDigitado, "Description")
        txtDB.Text = LerDoCofre(DsnDigitado, "Database")
        
        ' Se o parâmetro "Database" retornar vazio, tenta buscar por "Server" (bancos relacionais)
        If txtDB.Text = "" Then txtDB.Text = LerDoCofre(DsnDigitado, "Banco")
        
        ' 2. Recupera os dados de conexão com o servidor remoto
        txtHost.Text = LerDoCofre(DsnDigitado, "Server")
        txtUser.Text = LerDoCofre(DsnDigitado, "User")
        txtPass.Text = LerDoCofre(DsnDigitado, "Password")
        txtPort.Text = LerDoCofre(DsnDigitado, "Port")
        
        ' 3. Tenta recuperar e ajustar o Tipo de Driver no ComboBox
        Dim DriverGuardado As String
        DriverGuardado = UCase(LerDoCofre(DsnDigitado, "DriverType"))
        
        If DriverGuardado <> "" Then
            For i = 0 To cboDriver.ListCount - 1
                If UCase(cboDriver.List(i)) = DriverGuardado Then
                    cboDriver.ListIndex = i
                    Exit For
                End If
            Next i
        End If
        
        ' Se o formulário tiver a label de status (lblStatus) igual ao outro form [cite: 64]
        On Error Resume Next
        lblStatus.Caption = "Dados do DSN [" & DsnDigitado & "] carregados do cofre!"
        On Error GoTo 0
        
    Else
        ' Se não encontrar nada no cofre para este nome
        On Error Resume Next
        lblStatus.Caption = "DSN novo ou não localizado no cofre."
        On Error GoTo 0
    End If
End Sub
' Liga/Desliga campos do formulário de acordo com a arquitetura do banco escolhido
Private Sub cboDriver_Click()
    Dim driver As String
    driver = UCase(cboDriver.Text)
    
    ' Bancos de dados baseados em arquivos locais (Não usam rede/servidor)
    If driver = "MDB" Or driver = "ACCDB" Or driver = "DBF" Or driver = "SQLITE" Then
        txtHost.Enabled = False
        txtUser.Enabled = False
        txtPass.Enabled = False
        txtPort.Enabled = False
        txtHost.BackColor = &H8000000F
        txtUser.BackColor = &H8000000F
        txtPass.BackColor = &H8000000F
        txtPort.BackColor = &H8000000F
    Else
        ' Bancos de dados cliente/servidor baseados em rede
        txtHost.Enabled = True
        txtUser.Enabled = True
        txtPass.Enabled = True
        txtPort.Enabled = True
        txtHost.BackColor = &H80000005
        txtUser.BackColor = &H80000005
        txtPass.BackColor = &H80000005
        txtPort.BackColor = &H80000005
        
        ' Sugestão de portas padrões para ajudar o usuário
        Select Case driver
            Case "MYSQL", "MARIADB": txtPort.Text = "3306"
            Case "PGSQL":            txtPort.Text = "5432"
            Case "MSSQL":            txtPort.Text = ""
        End Select
    End If
End Sub

' Ação do Botão Salvar - Valida a tela e invoca o módulo Odbc.bas [cite: 1, 4]
Private Sub btnSalvar_Click()
    If Trim$(txtDSN.Text) = "" Then
        MsgBox "Por favor, defina um nome para o seu DSN.", vbExclamation, "Validação"
        txtDSN.SetFocus
        Exit Sub
    End If
    
    If Trim$(txtDB.Text) = "" Then
        MsgBox "Defina o banco de dados ou o caminho do arquivo destino.", vbExclamation, "Validação"
        txtDB.SetFocus
        Exit Sub
    End If

    Dim blnSucesso As Boolean
    
    ' Chama a rotina mesclada unificada que processa e trata os erros [cite: 4, 9, 25]
    blnSucesso = AddDSN(txtDSN.Text, _
                           txtDescription.Text, _
                           txtDB.Text, _
                           cboDriver.Text, _
                           (chkUserDSN.Value = 1), _
                           txtHost.Text, _
                           txtUser.Text, _
                           txtPass.Text, _
                           txtPort.Text)
                         
    If blnSucesso Then
        If MDG("Gravar no Cofre") Then
            Dim DsnNome As String
            DsnNome = UCase(Trim$(txtDSN.Text))
            
            ' Sincroniza salvando também no cofre de senhas local para redundância
            On Error Resume Next
            Call GravarNoCofre(DsnNome, "Description", txtDescription.Text)
            Call GravarNoCofre(DsnNome, "Database", txtDB.Text)
            Call GravarNoCofre(DsnNome, "Server", txtHost.Text)
            Call GravarNoCofre(DsnNome, "User", txtUser.Text)
            Call GravarNoCofre(DsnNome, "Password", txtPass.Text)
            Call GravarNoCofre(DsnNome, "Port", txtPort.Text)
            Call GravarNoCofre(DsnNome, "DriverType", cboDriver.Text)
            On Error GoTo 0
        End If
        MsgBox "DSN '" & txtDSN.Text & "' registrado e sincronizado com sucesso!", vbInformation, "Sucesso"
    End If
End Sub

' Ação do Botão Testar - Executa um teste de conexão ADO ativo com o DSN recém mapeado
Private Sub btnTestar_Click()
    If Trim$(txtDSN.Text) = "" Then
        MsgBox "Insira o nome do DSN que deseja testar.", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    On Error GoTo ErroConexao
    Dim cn As Object
    Set cn = CreateObject("ADODB.Connection")
    
    Dim strConn As String
    strConn = "DSN=" & txtDSN.Text & ";"
    
    ' Se for um banco remoto, adiciona as credenciais informadas na tela para o teste
    If txtUser.Enabled And Trim$(txtUser.Text) <> "" Then
        strConn = strConn & "UID=" & txtUser.Text & ";PWD=" & txtPass.Text & ";"
    End If
    
    Screen.MousePointer = vbHourglass
    cn.Open strConn
    Screen.MousePointer = vbDefault
    
    MsgBox "Conexão realizada com sucesso via ADO!", vbInformation, "Status da Conexão"
    cn.Close
    Set cn = Nothing
    Exit Sub

ErroConexao:
    Screen.MousePointer = vbDefault
    MsgBox "Erro ao conectar no DSN fornecido:" & vbCrLf & Err.Description, vbCritical, "Falha de Conexão"
End Sub
