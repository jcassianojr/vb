VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerador de Cofre de Senhas - VB96/VB6"
   ClientHeight    =   4644
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   5412
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4644
   ScaleWidth      =   5412
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAtualizar 
      Caption         =   "Atualizar Lista"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   480
      Width           =   1215
   End
   Begin VB.ComboBox cmbBancos 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir Banco"
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmdRevelar 
      Caption         =   "Ver"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton cmdLer 
      Caption         =   "Ler do Cofre"
      Height          =   495
      Left            =   2040
      TabIndex        =   7
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Gravar no Cofre"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox txtPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3000
      Width           =   2655
   End
   Begin VB.TextBox txtUser 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   2400
      Width           =   3375
   End
   Begin VB.TextBox txtServer 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox txtBanco 
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Bancos Salvos:"
      Height          =   195
      Left            =   360
      TabIndex        =   14
      Top             =   540
      Width           =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Password:"
      Height          =   195
      Left            =   360
      TabIndex        =   13
      Top             =   3090
      Width           =   735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "User ID:"
      Height          =   195
      Left            =   360
      TabIndex        =   12
      Top             =   2490
      Width           =   580
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Data Source / Servidor:"
      Height          =   195
      Left            =   360
      TabIndex        =   11
      Top             =   1890
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nome da Seção:"
      Height          =   195
      Left            =   360
      TabIndex        =   10
      Top             =   1290
      Width           =   1200
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' --- EVENTOS DO FORMULÁRIO ---

Private Sub Form_Load()
    ' Preenche a lista de bancos assim que o programa abre
    Call AtualizarLista
End Sub

Private Sub cmdAtualizar_Click()
    Call AtualizarLista
End Sub

Private Sub cmbBancos_Click()
    ' Quando o usuário clica num banco da lista, joga o nome para o campo e lê os dados automaticamente
    txtBanco.Text = cmbBancos.Text
    Call cmdLer_Click
End Sub

' --- BOTÃO GRAVAR (COM VERIFICAÇÃO DE DUPLICIDADE) ---
Private Sub cmdSalvar_Click()
    Dim NomeSecao As String
    Dim Resposta As VbMsgBoxResult
    
    NomeSecao = UCase(Trim(txtBanco.Text))

    If NomeSecao = "" Or Trim(txtServer.Text) = "" Or Trim(txtUser.Text) = "" Then
        MsgBox "Por favor, preencha todos os campos antes de salvar!", vbExclamation, "Aviso"
        Exit Sub
    End If

    ' Verifica se a seção já existe para alertar o usuário
    If LerDoCofre(NomeSecao, "Server") <> "" Then
        Resposta = MsgBox("A seção [" & NomeSecao & "] já existe no cofre." & vbCrLf & _
                          "Deseja substituir as informações atuais?", _
                          vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar Substituição")
        
        If Resposta = vbNo Then Exit Sub
    End If

    ' Grava os dados usando a CryptoAPI nativa
    Call GravarNoCofre(NomeSecao, "Server", Trim(txtServer.Text))
    Call GravarNoCofre(NomeSecao, "User", Trim(txtUser.Text))
    Call GravarNoCofre(NomeSecao, "Password", Trim(txtPassword.Text))

    MsgBox "Dados salvos com sucesso na seção [" & NomeSecao & "]!", vbInformation, "Sucesso"
    
    LimparCampos
    Call AtualizarLista
End Sub

' --- BOTÃO LER ---
Private Sub cmdLer_Click()
    Dim NomeSecao As String
    Dim Servidor As String
    Dim Usuario As String
    Dim Senha As String

    NomeSecao = UCase(Trim(txtBanco.Text))

    If NomeSecao = "" Then
        MsgBox "Por favor, digite ou selecione o 'Nome da Seção' para ler!", vbExclamation, "Aviso"
        txtBanco.SetFocus
        Exit Sub
    End If

    Servidor = LerDoCofre(NomeSecao, "Server")
    Usuario = LerDoCofre(NomeSecao, "User")
    Senha = LerDoCofre(NomeSecao, "Password")

    If Servidor = "" And Usuario = "" Then
        MsgBox "A seção [" & NomeSecao & "] não foi encontrada no arquivo config.dat!", vbCritical, "Erro"
        Exit Sub
    End If

    txtServer.Text = Servidor
    txtUser.Text = Usuario
    txtPassword.Text = Senha

    ' Mantém a senha oculta por padrão após ler
    txtPassword.PasswordChar = "*"
    cmdRevelar.Caption = "Ver"
End Sub

' --- BOTÃO REVELAR / OCULTAR SENHA ---
Private Sub cmdRevelar_Click()
    If txtPassword.PasswordChar = "*" Then
        txtPassword.PasswordChar = ""
        cmdRevelar.Caption = "Ocultar"
    Else
        txtPassword.PasswordChar = "*"
        cmdRevelar.Caption = "Ver"
    End If
    txtPassword.SetFocus
End Sub

' --- BOTÃO EXCLUIR BANCO ---
Private Sub cmdExcluir_Click()
    Dim Resposta As VbMsgBoxResult
    Dim BancoParaExcluir As String
    
    BancoParaExcluir = UCase(Trim(txtBanco.Text))
    
    If BancoParaExcluir = "" Then
        MsgBox "Selecione ou digite o nome do banco que deseja excluir!", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    Resposta = MsgBox("Tem certeza que deseja apagar o banco [" & BancoParaExcluir & "] permanentemente?", _
                      vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar Exclusão")
                      
    If Resposta = vbYes Then
        Call ExcluirBanco(BancoParaExcluir)
        MsgBox "Banco [" & BancoParaExcluir & "] excluído com sucesso!", vbInformation, "Concluído"
        LimparCampos
        Call AtualizarLista
    End If
End Sub

' --- ROTINAS AUXILIARES ---
Private Sub AtualizarLista()
    Call ListarBancosNoCombo(cmbBancos)
End Sub

Private Sub LimparCampos()
    txtBanco.Text = ""
    txtServer.Text = ""
    txtUser.Text = ""
    txtPassword.Text = ""
    txtPassword.PasswordChar = "*"
    cmdRevelar.Caption = "Ver"
    txtBanco.SetFocus
End Sub
