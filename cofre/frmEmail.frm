VERSION 5.00
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "VBCCR18.OCX"
Begin VB.Form frmEmail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerador de Cofre de E-mails - VB96/VB6"
   ClientHeight    =   5796
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   5952
   LinkTopic       =   "frmEmail"
   MaxButton       =   0   'False
   ScaleHeight     =   5796
   ScaleWidth      =   5952
   StartUpPosition =   2  'CenterScreen
   Begin VBCCR18.CommandButtonW cmdAtualizar 
      Height          =   612
      Left            =   3720
      TabIndex        =   15
      Top             =   240
      Width           =   2052
      _ExtentX        =   3620
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Caption         =   "Atualizar Lista"
      Picture         =   "form1.frx":0000
      PictureAndCaption=   -1  'True
      Style           =   1
   End
   Begin VBCCR18.ComboBoxW cmbEmails 
      Height          =   288
      Left            =   1560
      TabIndex        =   14
      Tag             =   "cmbEmails"
      Top             =   360
      Width           =   2052
      _ExtentX        =   3620
      _ExtentY        =   508
      Text            =   "cmbEmails"
   End
   Begin VB.TextBox txtConta 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txtSmtp 
      Height          =   285
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtPorta 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   2160
      Width           =   1200
   End
   Begin VB.TextBox txtSslTls 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   2640
      Width           =   2055
      Text            =   "STARTTLS" ' Valor padrão sugerido
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   3600
      Width           =   2055
   End
   Begin VBCCR18.CommandButtonW cmdRevelar 
      Height          =   612
      Left            =   3960
      TabIndex        =   16
      Top             =   3360
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Caption         =   "Exibir Senha"
      Picture         =   "form1.frx":049A
      PictureAndCaption=   -1  'True
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW cmdSalvar 
      Height          =   612
      Left            =   1680
      TabIndex        =   17
      Top             =   4080
      Width           =   1332
      _ExtentX        =   2350
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Caption         =   "Salvar"
      Picture         =   "form1.frx":0934
      PictureAndCaption=   -1  'True
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW cmdLimpar 
      Height          =   612
      Left            =   240
      TabIndex        =   18
      Top             =   4080
      Width           =   1332
      _ExtentX        =   2350
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Caption         =   "Limpar Campos"
      Picture         =   "form1.frx":0DCE
      PictureAndCaption=   -1  'True
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW cmdExcluir 
      Height          =   612
      Left            =   3120
      TabIndex        =   19
      Top             =   4080
      Width           =   1332
      _ExtentX        =   2350
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Caption         =   "Excluir Conta"
      Picture         =   "form1.frx":1268
      PictureAndCaption=   -1  'True
      Style           =   1
   End
   Begin VB.Label Label7 
      Caption         =   "SSL / TLS:"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Porta SMTP:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Contas Salvas:"
      Height          =   252
      Left            =   240
      TabIndex        =   10
      Top             =   360
      Width           =   1212
   End
   Begin VB.Label Label4 
      Caption         =   "Senha:"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Usuário / Email:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Servidor SMTP:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nome da Conta:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   5760
      Y1              =   996
      Y2              =   996
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   5640
      Y1              =   5100
      Y2              =   5100
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Cofre de e-mails pronto para operação"
      ForeColor       =   &H00800000&
      Height          =   252
      Left            =   240
      TabIndex        =   13
      Top             =   5304
      Width           =   5292
   End
End
Attribute VB_Name = "frmEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call AtualizarLista
End Sub

' --- BOTÃO ATUALIZAR LISTA ---
Private Sub cmdAtualizar_Click()
    Call AtualizarLista
End Sub

' --- EVENTO CLIQUE NA LISTA DE EMAILS ---
Private Sub cmbEmails_Click()
    Dim ContaSelecionada As String
    If cmbEmails.ListIndex = -1 Then Exit Sub
    
    ContaSelecionada = cmbEmails.List(cmbEmails.ListIndex)
    txtConta.Text = ContaSelecionada
    
    ' Lê as configurações do E-mail mapeando a mesma lógica do seu cofre
    txtSmtp.Text = LerDoCofre(ContaSelecionada, "SmtpServer")
    txtPorta.Text = LerDoCofre(ContaSelecionada, "SmtpPort")
    txtSslTls.Text = LerDoCofre(ContaSelecionada, "SslTls")
    txtUser.Text = LerDoCofre(ContaSelecionada, "EmailUser")
    txtPassword.Text = LerDoCofre(ContaSelecionada, "EmailPassword")
    
    lblStatus.Caption = "Dados de [" & ContaSelecionada & "] carregados."
End Sub

' --- BOTÃO SALVAR ---
Private Sub cmdSalvar_Click()
    Dim Conta As String
    Conta = UCase(Trim(txtConta.Text))
    
    If Conta = "" Then
        MsgBox "O nome da conta é obrigatório para salvar!", vbExclamation, "Aviso"
        txtConta.SetFocus
        Exit Sub
    End If
    
    ' Grava os parâmetros no seu arquivo/banco cofre criptografado
    Call GravarNoCofre(Conta, "SmtpServer", txtSmtp.Text)
    Call GravarNoCofre(Conta, "SmtpPort", txtPorta.Text)
    Call GravarNoCofre(Conta, "SslTls", txtSslTls.Text)
    Call GravarNoCofre(Conta, "EmailUser", txtUser.Text)
    Call GravarNoCofre(Conta, "EmailPassword", txtPassword.Text)
    
    lblStatus.Caption = "Conta [" & Conta & "] gravada com sucesso!"
    MsgBox "Dados de e-mail gravados com sucesso no cofre!", vbInformation, "Sucesso"
    
    Call AtualizarLista
    
    ' Seleciona automaticamente o e-mail recém-salvo no combo
    Dim i As Integer
    For i = 0 To cmbEmails.ListCount - 1
        If cmbEmails.List(i) = Conta Then
            cmbEmails.ListIndex = i
            Exit For
        End If
    Next i
End Sub

' --- BOTÃO LIMPAR CAMPOS / NOVO ---
Private Sub cmdLimpar_Click()
    txtConta.Text = ""
    txtSmtp.Text = ""
    txtPorta.Text = ""
    txtSslTls.Text = "STARTTLS"
    txtUser.Text = ""
    txtPassword.Text = ""
    
    cmbEmails.ListIndex = -1
    lblStatus.Caption = "Campos limpos. Pronto para nova configuração de e-mail."
    txtConta.SetFocus
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

' --- BOTÃO EXCLUIR CONTA ---
Private Sub cmdExcluir_Click()
    Dim Resposta As VbMsgBoxResult
    Dim ContaParaExcluir As String
    
    ContaParaExcluir = UCase(Trim(txtConta.Text))
    
    If ContaParaExcluir = "" Then
        MsgBox "Selecione ou digite o nome da conta que deseja excluir!", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    Resposta = MsgBox("Tem certeza que deseja apagar a conta [" & ContaParaExcluir & "] permanentemente?", _
                      vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar Exclusão")
                      
    If Resposta = vbYes Then
        ' Se a sua função ExcluirBanco for genérica para chaves do cofre, pode usá-la.
        ' Caso contrário, adapte para a rotina correta de deleção do seu sistema.
        Call ExcluirBanco(ContaParaExcluir) 
        MsgBox "Conta [" & ContaParaExcluir & "] excluída com sucesso!", vbInformation, "Concluído"
        Call cmdLimpar_Click
        Call AtualizarLista
    End If
End Sub

' --- ROTINAS AUXILIARES ---
Private Sub ProfilerCheck()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim sPasta As String
    sPasta = fso.GetFile(CaminhoArquivoCofre).ParentFolder
    If Not fso.FolderExists(sPasta) Then
        fso.CreateFolder (sPasta)
    End If
    If Not fso.FileExists(CaminhoArquivoCofre) Then
        Dim ts As Object
        Set ts = fso.CreateTextFile(CaminhoArquivoCofre, True)
        ts.Close
    End If
    Set fso = Nothing
End Sub

Private Sub AtualizarLista()
    On Error Resume Next
    Call ProfilerCheck
    ' Passa o novo ComboBox modificado como parâmetro
    Call ListarBancosNoCombo(cmbEmails) 
End Sub