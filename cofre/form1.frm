VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerador de Cofre de Senhas - VB96/VB6"
   ClientHeight    =   5796
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   5412
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5796
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
   Begin VB.TextBox txtBanco 
      Height          =   285
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtPorta 
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Top             =   2160
      Width           =   1200
   End
   Begin VB.TextBox txtOwner 
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   3600
      Width           =   2055
   End
   Begin VB.CommandButton cmdRevelar 
      Caption         =   "Ver"
      Height          =   315
      Left            =   3840
      TabIndex        =   8
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "Salvar"
      Height          =   495
      Left            =   480
      TabIndex        =   9
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar Campos"
      Height          =   495
      Left            =   2040
      TabIndex        =   10
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "Excluir Banco"
      Height          =   495
      Left            =   3600
      TabIndex        =   11
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Owner / Esquema:"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Porta Banco:"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Bancos Salvos:"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Senha:"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Usuario:"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Servidor / IP:"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nome do Banco:"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   5160
      Y1              =   1000
      Y2              =   1000
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   5160
      Y1              =   5100
      Y2              =   5100
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Cofre pronto para operacao"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   5300
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call AtualizarLista
End Sub

' --- BOTÃƒO ATUALIZAR LISTA ---
Private Sub cmdAtualizar_Click()
    Call AtualizarLista
End Sub

' --- EVENTO CLIQUE NA LISTA DE BANCOS ---
Private Sub cmbBancos_Click()
    Dim BancoSelecionado As String
    If cmbBancos.ListIndex = -1 Then Exit Sub
    
    BancoSelecionado = cmbBancos.List(cmbBancos.ListIndex)
    txtBanco.Text = BancoSelecionado
    
    ' LÃª as credenciais tradicionais
    txtServer.Text = LerDoCofre(BancoSelecionado, "Server")
    txtUser.Text = LerDoCofre(BancoSelecionado, "User")
    txtPassword.Text = LerDoCofre(BancoSelecionado, "Password")
    
    ' LÃª os novos parÃ¢metros adicionados
    txtPorta.Text = LerDoCofre(BancoSelecionado, "Port")
    txtOwner.Text = LerDoCofre(BancoSelecionado, "Owner")
    
    lblStatus.Caption = "Dados de [" & BancoSelecionado & "] carregados."
End Sub

' --- BOTÃƒO SALVAR ---
Private Sub cmdSalvar_Click()
    Dim Banco As String
    Banco = UCase(Trim(txtBanco.Text))
    
    If Banco = "" Then
        MsgBox "O nome do banco Ã© obrigatÃ³rio para salvar!", vbExclamation, "Aviso"
        txtBanco.SetFocus
        Exit Sub
    End If
    
    ' Salva os parÃ¢metros tradicionais
    Call GravarNoCofre(Banco, "Server", txtServer.Text)
    Call GravarNoCofre(Banco, "User", txtUser.Text)
    Call GravarNoCofre(Banco, "Password", txtPassword.Text)
    
    ' Salva os dois parÃ¢metros novos
    Call GravarNoCofre(Banco, "Port", txtPorta.Text)
    Call GravarNoCofre(Banco, "Owner", txtOwner.Text)
    
    lblStatus.Caption = "Banco [" & Banco & "] gravado com sucesso!"
    MsgBox "Dados gravados com sucesso no cofre criptografado!", vbInformation, "Sucesso"
    
    Call AtualizarLista
    
    ' Seleciona automaticamente o banco recÃ©m-salvo no combo
    Dim i As Integer
    For i = 0 To cmbBancos.ListCount - 1
        If cmbBancos.List(i) = Banco Then
            cmbBancos.ListIndex = i
            Exit For
        End If
    Next i
End Sub

' --- BOTÃƒO LIMPAR CAMPOS / NOVO ---
Private Sub cmdLimpar_Click()
    txtBanco.Text = ""
    txtServer.Text = ""
    txtPorta.Text = ""
    txtOwner.Text = ""
    txtUser.Text = ""
    txtPassword.Text = ""
    
    cmbBancos.ListIndex = -1
    lblStatus.Caption = "Campos limpos. Pronto para novo cadastro."
    txtBanco.SetFocus
End Sub

' --- BOTÃƒO REVELAR / OCULTAR SENHA ---
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

' --- BOTÃƒO EXCLUIR BANCO ---
Private Sub cmdExcluir_Click()
    Dim Resposta As VbMsgBoxResult
    Dim BancoParaExcluir As String
    
    BancoParaExcluir = UCase(Trim(txtBanco.Text))
    
    If BancoParaExcluir = "" Then
        MsgBox "Selecione ou digite o nome do banco que deseja excluir!", vbExclamation, "Aviso"
        Exit Sub
    End If
    
    Resposta = MsgBox("Tem certeza que deseja apagar o banco [" & BancoParaExcluir & "] permanentemente?", _
                      vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar ExclusÃ£o")
                      
    If Resposta = vbYes Then
        Call ExcluirBanco(BancoParaExcluir)
        MsgBox "Banco [" & BancoParaExcluir & "] excluÃ­do com sucesso!", vbInformation, "ConcluÃ­do"
        Call cmdLimpar_Click
        Call AtualizarLista
    End If
End Sub

' --- ROTINAS AUXILIARES ---
Private Sub ProfilerCheck()
    ' Verifica se a pasta existe antes de listar
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

Private Sub CleanIniGarbage()
    ' Esta rotina opcional limpa chaves vazias se necessÃ¡rio
End Sub

Private Sub AtualizarLista()
    On Error Resume Next
    Call ProfilerCheck
    Call ListarBancosNoCombo(cmbBancos)
End Sub
