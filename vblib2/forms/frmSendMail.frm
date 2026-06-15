VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "vbccr18.ocx"
Begin VB.Form FrmSendMail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enviando Email"
   ClientHeight    =   7020
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   10080
   Icon            =   "frmSendMail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   10080
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFromName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   21
      Top             =   120
      Width           =   4815
   End
   Begin VB.TextBox txtPassword 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   5280
      PasswordChar    =   "*"
      TabIndex        =   19
      Top             =   5040
      Width           =   1800
   End
   Begin VB.TextBox txtUsername 
      Height          =   300
      Left            =   2520
      TabIndex        =   17
      Top             =   5040
      Width           =   1800
   End
   Begin VB.CheckBox chkSSL 
      Alignment       =   1  'Right Justify
      Caption         =   "Req. SSL"
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   240
      TabIndex        =   15
      Top             =   5040
      Width           =   1065
   End
   Begin VB.TextBox txt_status 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5520
      Width           =   7515
   End
   Begin VB.TextBox txt_message_text 
      Appearance      =   0  'Flat
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1920
      Width           =   9615
   End
   Begin XPControls.XPButton CmdLimpa 
      Height          =   615
      Left            =   7680
      TabIndex        =   13
      Top             =   5520
      Width           =   1095
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Limpar Informaçoes"
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
   Begin VB.TextBox Txt_Porta 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6600
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "25"
      Top             =   4560
      Width           =   615
   End
   Begin VB.TextBox txt_email_to 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1200
      TabIndex        =   2
      Top             =   840
      Width           =   4815
   End
   Begin VB.TextBox txt_subject 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1200
      TabIndex        =   3
      Top             =   1200
      Width           =   4815
   End
   Begin VB.TextBox txt_attach 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1200
      TabIndex        =   4
      Top             =   1560
      Width           =   4815
   End
   Begin VB.TextBox txt_email_from 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   4815
   End
   Begin VB.TextBox txt_smtp_server 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1680
      TabIndex        =   0
      Top             =   4560
      Width           =   4215
   End
   Begin XPControls.XPButton Command9x 
      Height          =   255
      Left            =   6240
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1560
      Width           =   855
      _ExtentX        =   1503
      _ExtentY        =   445
      Caption         =   "ZIP"
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
   Begin VBCCR18.CommandButtonW retornar 
      Height          =   612
      Left            =   8520
      TabIndex        =   23
      Top             =   120
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
      Picture         =   "frmSendMail.frx":038A
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW CmdEnviar 
      Height          =   612
      Left            =   6120
      TabIndex        =   24
      Top             =   120
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Email"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "frmSendMail.frx":0924
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   0
      TabIndex        =   20
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   6
      Left            =   4440
      TabIndex        =   18
      Top             =   5040
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   5
      Left            =   1680
      TabIndex        =   16
      Top             =   5040
      Width           =   720
   End
   Begin VB.Label Label6 
      Caption         =   "Porta:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6000
      TabIndex        =   12
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Anexos"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1005
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Assunto"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Email Para"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Email"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "SMTP(Servidor/IP)"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   4680
      Width           =   1365
   End
End
Attribute VB_Name = "FrmSendMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Defina aqui o nome da conta que o formulário deve carregar por padrão
Private Const CONTA_PADRAO As String = "PADRAO"

Private Sub cmdCANCELAR_Click()
  Unload Me
End Sub

Private Sub CmdEnviar_Click()
  cdoenviar_Click
End Sub

Private Sub CmdLimpa_Click()
  txt_status.Text = ""
End Sub

Private Sub Command9_Click()
  Dim cOrigem As String
  Dim cDestino As String
  cOrigem = txt_attach.Text
  If InStr(UCase(cOrigem), ".ZIP") > 0 Then
    Alert ("Ja e um zip")
  Else
    cDestino = TrocaExt(cOrigem, "ZIP")
    With New cZipArchive
      .AddFile cOrigem
       .CompressArchive cDestino
    End With
    txt_attach.Text = cDestino
  End If
End Sub

Private Sub cdoenviar_Click()
  Dim RetVal As String
  RetVal = SendMailCDO(Trim$(txt_email_to.Text), _
                       Trim$(txt_subject.Text), _
                       Trim$(txtFromName.Text) & "<" & Trim$(txt_email_from.Text) & ">", _
                       Trim$(txt_message_text.Text), _
                       Trim$(txt_smtp_server.Text), _
                       CInt(Trim$(Txt_Porta.Text)), _
                       Trim$(txtUsername.Text), _
                       Trim$(txtPassword.Text), _
                       Trim$(txt_attach.Text), _
                       CBool(chkSSL.Value))
  txt_status.Text = RetVal
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  txt_status.Text = "Pronto." & vbCrLf
  
  ' --- IDENTIFICAÇÃO DINÂMICA DA CONTA ---
  Dim sContaAtiva As String
  
  ' Tenta obter o nome da conta associada ao usuário/folha atual
  sContaAtiva = Trim(PegPath("EMAIL", UCase(zNOMEFOLHA), ""))
  
  ' Caso o INI retorne vazio, define uma conta padrão para não quebrar a busca no cofre
  If sContaAtiva = "" Then
    sContaAtiva = "PADRAO"
  End If
  
  ' --- CARREGAMENTO INDEPENDENTE ELEMENTO POR ELEMENTO ---
  
  ' 0 - Servidor SMTP
  If Trim(ePASS01(0)) <> "" Then
    txt_smtp_server.Text = ePASS01(0)
  Else
    txt_smtp_server.Text = LerDoCofre(sContaAtiva, "SmtpServer")
  End If

  ' 1 - Porta SMTP
  If Trim(ePASS01(1)) <> "" Then
    Txt_Porta.Text = ePASS01(1)
  Else
    Txt_Porta.Text = LerDoCofre(sContaAtiva, "SmtpPort")
  End If

  ' 2 - Email Remetente (From) e Usuário
  If Trim(ePASS01(2)) <> "" Then
    txt_email_from.Text = ePASS01(2)
    txtUsername.Text = ePASS01(2)
  Else
    txt_email_from.Text = LerDoCofre(sContaAtiva, "EmailUser")
    txtUsername.Text = LerDoCofre(sContaAtiva, "EmailUser")
  End If

  ' Senha (Buscada do cofre se não estiver preenchida na tela)
  If Trim(txtUsername.Text) <> "" And Trim(txtPassword.Text) = "" Then
    txtPassword.Text = LerDoCofre(sContaAtiva, "EmailPassword")
  End If

  ' 3 - Destinatário (To)
  If Trim(ePASS01(3)) <> "" Then
    txt_email_to.Text = ePASS01(3)
  End If

  ' 4 - Assunto (Subject)
  If Trim(ePASS01(4)) <> "" Then
    txt_subject.Text = ePASS01(4)
  End If

  ' 5 - Anexo (Attach)
  If Trim(ePASS01(5)) <> "" Then
    txt_attach.Text = ePASS01(5)
  End If

  ' 6 - Corpo da Mensagem (Message Text)
  If Trim(ePASS01(6)) <> "" Then
    txt_message_text.Text = ePASS01(6)
  End If

  ' --- GERENCIAMENTO DO CHECKBOX SSL ---
  ' Busca a segurança configurada para esta conta específica no cofre
  Dim sSeguranca As String
  sSeguranca = UCase(Trim(LerDoCofre(sContaAtiva, "SslTls")))
  If sSeguranca = "SSL" Or sSeguranca = "TLS" Or sSeguranca = "STARTTLS" Then
      chkSSL.Value = vbChecked
  Else
      chkSSL.Value = vbUnchecked
  End If

  ' --- FALLBACKS SE CONDICIONAIS FAIHAREM ---
  If Len(txt_smtp_server.Text) = 0 Then
    txt_smtp_server.Text = "smtp.com.br"
  End If
  If Len(Txt_Porta.Text) = 0 Then
    Txt_Porta.Text = "25"
  End If

  If ePASS01(7) = True Then
    CmdEnviar_Click
  End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Unload Me
End Sub

Private Sub Lista_Click()
  CmdEnviar_Click
End Sub

Private Sub retornar_Click()
  Unload Me
End Sub

Private Sub txt_email_from_LostFocus()
  If Not CheckEmail(txt_email_from) Then
    txt_email_from.SetFocus
  End If
End Sub

Private Sub txt_email_to_LostFocus()
  If Not CheckEmail(txt_email_to) Then
    txt_email_to.SetFocus
  End If
End Sub

Private Sub txt_porta_KeyPress(KeyAscii As Integer)
  KeyAscii = ValiText(KeyAscii, "#NI")
End Sub

Function campos_checagem() As Boolean
  campos_checagem = False
  If Len(txt_subject.Text) = 0 Then
    If Not MDG("Enviar Sem Assunto") Then
      Exit Function
    End If
  End If
  If Len(txt_email_to.Text) = 0 Then
    Alert ("Sem Destinatario")
    Exit Function
  End If
  If Len(txt_email_from.Text) = 0 Then
    Alert ("Sem Remetente")
    Exit Function
  End If
  If Not CheckEmail(txt_email_from.Text) Then
    txt_status.Text = "Erro: " & txt_email_from.Text & "." & vbCrLf & txt_status.Text & vbCrLf
    Alert ("Email Origem Invalido")
    Exit Function
  End If
  If Not CheckEmail(txt_email_to.Text) Then
    txt_status.Text = "Erro: " & txt_email_to.Text & "." & vbCrLf & txt_status.Text & vbCrLf
    Alert ("Email Destino Invalido")
    Exit Function
  End If
  If Len(txt_smtp_server.Text) = 0 Then
    Alert ("Servidor de Envio Nao Preenchido")
    Exit Function
  End If
  If Not IsWebConnected Then
    If Not MDG("Internet Pode nao estar conectada Continuar") Then
      Exit Function
    End If
  End If
  campos_checagem = True
End Function
