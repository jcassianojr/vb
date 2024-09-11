VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form FrmSendMail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enviando Email"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   315
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   28
      Top             =   120
      Width           =   4815
   End
   Begin VB.TextBox txtPassword 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   5280
      PasswordChar    =   "*"
      TabIndex        =   26
      Top             =   5040
      Width           =   1800
   End
   Begin VB.TextBox txtUsername 
      Height          =   300
      Left            =   2520
      TabIndex        =   24
      Top             =   5040
      Width           =   1800
   End
   Begin VB.CheckBox chkSSL 
      Alignment       =   1  'Right Justify
      Caption         =   "Req. SSL"
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   240
      TabIndex        =   22
      Top             =   5040
      Width           =   1065
   End
   Begin VB.ListBox Lista 
      Height          =   1035
      ItemData        =   "frmSendMail.frx":038A
      Left            =   6240
      List            =   "frmSendMail.frx":0397
      TabIndex        =   18
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton cdoenviar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cdo"
      Height          =   615
      Left            =   2040
      Picture         =   "frmSendMail.frx":03AD
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6240
      Width           =   855
   End
   Begin VB.TextBox txt_status 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
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
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   8760
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   8160
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.CommandButton mapienviar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Mapl"
      Height          =   615
      Left            =   120
      Picture         =   "frmSendMail.frx":0737
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton shellenviar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Shell"
      Height          =   615
      Left            =   1080
      Picture         =   "frmSendMail.frx":0AC1
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6240
      Width           =   855
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
         Size            =   8.25
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
   Begin XPControls.XPButton CmdCancelar 
      Height          =   435
      Left            =   8280
      TabIndex        =   20
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmSendMail.frx":0E4B
      Caption         =   "Retornar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton CmdEnviar 
      Height          =   375
      Left            =   8280
      TabIndex        =   21
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Picture         =   "frmSendMail.frx":13E5
      Caption         =   "EnviarEmail"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton Command9 
      Height          =   255
      Left            =   6240
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1560
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      Caption         =   "ZIP"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   0
      TabIndex        =   27
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
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
      TabIndex        =   25
      Top             =   5040
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
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
      TabIndex        =   23
      Top             =   5040
      Width           =   720
   End
   Begin VB.Label Label7 
      Caption         =   "Enviar Usando"
      Height          =   255
      Left            =   6240
      TabIndex        =   19
      Top             =   120
      Width           =   2055
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
Private Sub CmdCancelar_Click()
  Unload Me
End Sub
Private Sub CmdEnviar_Click()
  Select Case Lista.ListIndex

  Case 0
    mapienviar_Click  ''mapi
  Case 1
    shellenviar_Click  ''shell
  Case 2
    cdoenviar_Click  ''cdo
  End Select

End Sub
Private Sub CmdLimpa_Click()
  txt_status.tEXT = ""
End Sub

Private Sub Command9_Click()
  Dim cORIGEM As String
  Dim cDESTINO As String
  cORIGEM = txt_attach.tEXT
  If InStr(UCase(cORIGEM), ".ZIP") > 0 Then
    Alert ("Ja e um zip")
  Else
    cDESTINO = TrocaExt(cORIGEM, "ZIP")
    With New cZipArchive
      .AddFile cORIGEM
       .CompressArchive cDESTINO
    End With
    txt_attach.tEXT = cDESTINO
  End If

End Sub

Private Sub shellenviar_Click()
  If campos_checagem Then
    SendEmailShell txt_email_to.tEXT, txt_subject.tEXT, txt_message_text.tEXT, txt_attach.tEXT
  End If
End Sub
Private Sub mapienviar_Click()
  If Not campos_checagem Then
    Exit Sub
  End If
  MAPISession1.LogonUI = True
  MAPISession1.SignOn
  MAPIMessages1.sessionID = MAPISession1.sessionID
  MAPIMessages1.Compose

  '////////////////////////////////////////
  '// Informa o email, se enviar para mais de 1 email é separar com(;)e ir informando
  '////////////////////////////////////////
  MAPIMessages1.RecipIndex = 0
  MAPIMessages1.RecipAddress = txt_email_to.tEXT
  MAPIMessages1.RecipType = 1

  '////////////////////////////////////////
  '// Assunto
  '////////////////////////////////////////
  MAPIMessages1.MsgSubject = txt_subject.tEXT
  '////////////////////////////////////////
  '// Conteudo da mens
  '////////////////////////////////////////
  MAPIMessages1.MsgNoteText = txt_message_text.tEXT
  '////////////////////////////////////////
  '// Informa anexo
  '////////////////////////////////////////
  MAPIMessages1.AttachmentIndex = 0
  MAPIMessages1.AttachmentPathName = txt_attach.tEXT
  MAPIMessages1.AttachmentPosition = 0
  '////////////////////////////////////////////////
  '// Se existir outros anexos alterar os valores
  '////////////////////////////////////////////////
  'MAPIMessages1.AttachmentIndex = 1
  'MAPIMessages1.AttachmentPathName = app.path & "\Leia.doc"
  'MAPIMessages1.AttachmentPosition = 1

  MAPIMessages1.Send False
  MAPISession1.SignOff
End Sub


Private Sub cdoenviar_Click()
  Dim RetVal As String
  RetVal = SendMailCDO(Trim$(txt_email_to.tEXT), _
                       Trim$(txt_subject.tEXT), _
                       Trim$(txtFromName.tEXT) & "<" & Trim$(txt_email_from.tEXT) & ">", _
                       Trim$(txt_message_text.tEXT), _
                       Trim$(txt_smtp_server.tEXT), _
                       CInt(Trim$(Txt_Porta.tEXT)), _
                       Trim$(txtUsername.tEXT), _
                       Trim$(txtPassword.tEXT), _
                       Trim$(txt_attach.tEXT), _
                       CBool(chkSSL.Value))
  txt_status.tEXT = RetVal

End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  txt_status.tEXT = "Pronto." & vbCrLf
  txt_smtp_server.tEXT = ePASS01(0)
  Txt_Porta.tEXT = ePASS01(1)
  txt_email_from.tEXT = ePASS01(2)
  txt_email_to.tEXT = ePASS01(3)
  txt_subject.tEXT = ePASS01(4)
  txt_attach.tEXT = ePASS01(5)
  txt_message_text.tEXT = ePASS01(6)
  If Len(txt_smtp_server.tEXT) = 0 Then
    txt_smtp_server.tEXT = PegPath("EMAIL", "SERVER", "stmp..com.br")
  End If
  If Len(Txt_Porta.tEXT) = 0 Then
    Txt_Porta.tEXT = PegPath("EMAIL", "PORTA", "25")
  End If
  If Len(txt_email_from.tEXT) = 0 Then
    txt_email_from.tEXT = PegPath("EMAIL", UCase(zNOMEFOLHA), " ")
    If Len(Trim(txt_email_from.tEXT)) = 0 Then
      txt_email_from.tEXT = PegPath("EMAIL", "FROM", "@.com.br")
    End If
  End If
  Lista.ListIndex = 0

  If ePASS01(7) = True Then
    'enviaresair
  End If
  ''txt
End Sub
Private Sub Form_Unload(Cancel As Integer)
  Unload Me
End Sub
Private Sub Lista_Click()
  CmdEnviar_Click
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
  If Len(txt_subject.tEXT) = 0 Then
    If Not MDG("Enviar Sem Assunto") Then
      Exit Function
    End If
  End If
  If Len(txt_email_to.tEXT) = 0 Then
    Alert ("Sem Destinatario")
    Exit Function
  End If
  If Len(txt_email_from.tEXT) = 0 Then
    Alert ("Sem Remetente")
    Exit Function
  End If
  If Not CheckEmail(txt_email_from.tEXT) Then
    txt_status.tEXT = "Erro: " & txt_email_from.tEXT & "." & vbCrLf & txt_status.tEXT & vbCrLf
    Alert ("Email Origem Invalido")
    Exit Function
  End If
  If Not CheckEmail(txt_email_to.tEXT) Then
    txt_status.tEXT = "Erro: " & txt_email_to.tEXT & "." & vbCrLf & txt_status.tEXT & vbCrLf
    Alert ("Email Destino Invalido")
    Exit Function
  End If
  If Len(txt_smtp_server.tEXT) = 0 Then
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


