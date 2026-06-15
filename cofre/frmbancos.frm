VERSION 5.00
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "vbccr18.ocx"
Begin VB.Form Frmbancos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerador de Cofre de Senhas - VB96/VB6"
   ClientHeight    =   5796
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   6612
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5796
   ScaleWidth      =   6612
   StartUpPosition =   2  'CenterScreen
   Begin VBCCR18.CommandButtonW cmdAtualizar 
      Height          =   612
      Left            =   4080
      TabIndex        =   15
      Top             =   240
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Caption         =   "Atualizar Lista"
      Picture         =   "frmbancos.frx":0000
      PictureAndCaption=   -1  'True
      Style           =   1
   End
   Begin VBCCR18.ComboBoxW cmbBancos 
      Height          =   288
      Left            =   1560
      TabIndex        =   14
      Tag             =   "cmbBancos"
      Top             =   360
      Width           =   2292
      _ExtentX        =   4043
      _ExtentY        =   508
      Text            =   "cmbBancos"
   End
   Begin VB.TextBox txtBanco 
      Height          =   285
      Left            =   1680
      TabIndex        =   0
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txtServer 
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
   Begin VB.TextBox txtOwner 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Top             =   2640
      Width           =   2055
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
      Top             =   3480
      Width           =   1212
      _ExtentX        =   2138
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Caption         =   "Exibir Senha"
      Picture         =   "frmbancos.frx":049A
      PictureAndCaption=   -1  'True
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW cmdSalvar 
      Height          =   612
      Left            =   1680
      TabIndex        =   17
      Top             =   4200
      Width           =   1332
      _ExtentX        =   2350
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Caption         =   "Salvar"
      Picture         =   "frmbancos.frx":0934
      PictureAndCaption=   -1  'True
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW cmdLimpar 
      Height          =   612
      Left            =   240
      TabIndex        =   18
      Top             =   4200
      Width           =   1332
      _ExtentX        =   2350
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Caption         =   "Limpar Campos"
      Picture         =   "frmbancos.frx":0DCE
      PictureAndCaption=   -1  'True
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW cmdExcluir 
      Height          =   612
      Left            =   3120
      TabIndex        =   19
      Top             =   4200
      Width           =   1332
      _ExtentX        =   2350
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Caption         =   "Excluir Banco"
      Picture         =   "frmbancos.frx":1268
      PictureAndCaption=   -1  'True
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW cmddsn 
      Height          =   612
      Left            =   5400
      TabIndex        =   20
      Top             =   1080
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Dsn(ODBC)"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "frmbancos.frx":1702
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW cmdsair 
      Height          =   612
      Left            =   5400
      TabIndex        =   21
      Top             =   240
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
      Picture         =   "frmbancos.frx":185C
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VB.Label Label7 
      Caption         =   "Owner / Esquema:"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Porta Banco:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Bancos Salvos:"
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
      Caption         =   "Usuario:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Servidor / IP:"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Nome do Banco:"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   5280
      Y1              =   996
      Y2              =   996
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   5280
      Y1              =   5100
      Y2              =   5040
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      Caption         =   "Cofre pronto para operacao"
      ForeColor       =   &H00800000&
      Height          =   252
      Left            =   240
      TabIndex        =   13
      Top             =   5304
      Width           =   5292
   End
End
Attribute VB_Name = "Frmbancos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButtonW5_Click()

End Sub

Private Sub cmddsn_Click()
  frmCadastroDSN.Show vbModal, Me
   
End Sub

Private Sub cmdsair_Click()
  Unload Me
End Sub

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
