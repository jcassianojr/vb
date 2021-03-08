VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
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
      TabIndex        =   31
      Top             =   120
      Width           =   4815
   End
   Begin VB.TextBox txtPassword 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   5280
      PasswordChar    =   "*"
      TabIndex        =   29
      Top             =   5040
      Width           =   1800
   End
   Begin VB.TextBox txtUsername 
      Height          =   300
      Left            =   2520
      TabIndex        =   27
      Top             =   5040
      Width           =   1800
   End
   Begin VB.CheckBox chkSSL 
      Alignment       =   1  'Right Justify
      Caption         =   "Req. SSL"
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   240
      TabIndex        =   25
      Top             =   5040
      Width           =   1065
   End
   Begin VB.ListBox Lista 
      Height          =   1425
      ItemData        =   "frmSendMail.frx":038A
      Left            =   6240
      List            =   "frmSendMail.frx":03A0
      TabIndex        =   21
      Top             =   360
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "PY"
      Height          =   615
      Left            =   4920
      Picture         =   "frmSendMail.frx":03C8
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "vbm"
      Height          =   615
      Left            =   2040
      Picture         =   "frmSendMail.frx":0752
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "cdo"
      Height          =   615
      Left            =   1080
      Picture         =   "frmSendMail.frx":0ADC
      Style           =   1  'Graphical
      TabIndex        =   18
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
      TabIndex        =   17
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
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Mapl"
      Height          =   615
      Left            =   3960
      Picture         =   "frmSendMail.frx":0E66
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Shell"
      Height          =   615
      Left            =   3000
      Picture         =   "frmSendMail.frx":11F0
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton CmdLimpa 
      Caption         =   "Limpar Informaחoes"
      Height          =   615
      Left            =   7680
      TabIndex        =   14
      Top             =   5520
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Winsock"
      Height          =   615
      Left            =   120
      Picture         =   "frmSendMail.frx":157A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6240
      Width           =   855
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
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   9480
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
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
      TabIndex        =   23
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmSendMail.frx":1904
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
      TabIndex        =   24
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Picture         =   "frmSendMail.frx":1E9E
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
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   0
      TabIndex        =   30
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
      TabIndex        =   28
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
      TabIndex        =   26
      Top             =   5040
      Width           =   720
   End
   Begin VB.Label Label7 
      Caption         =   "Enviar Usando"
      Height          =   255
      Left            =   6240
      TabIndex        =   22
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
Dim Response As String

Sub wait_for(winsock_answare As String)
    'FIXIT: Replace 'Left' function with 'Left$' function                                      FixIT90210ae-R9757-R1B8ZE
    Do While Left(Response, 3) <> winsock_answare
        DoEvents
    Loop
    Response = ""
End Sub

Function find_date() As String
    Dim temp As String
    Dim fd_day As String
    Dim fd_month As String
    Dim fd_time As String
    
    fd_day = Format(Date, "Dddd")
    Select Case fd_day
    Case "יום ראשון": fd_day = "Sun, "
    Case "יום שני": fd_day = "Mon, "
    Case "יום שלישי": fd_day = "Tue, "
    Case "יום רביעי": fd_day = "Wed, "
    Case "יום חמישי": fd_day = "Thu, "
    Case "יום שישי": fd_day = "Fri, "
    Case "יום שבת": fd_day = "Sat, "
    End Select
    fd_month = Month(Date)
    Select Case fd_month
    Case 1: fd_month = "Jan "
    Case 2: fd_month = "Feb "
    Case 3: fd_month = "Mar "
    Case 4: fd_month = "Apr "
    Case 5: fd_month = "May "
    Case 6: fd_month = "Jun "
    Case 7: fd_month = "Jul "
    Case 8: fd_month = "Aug "
    Case 9: fd_month = "Sep "
    Case 10: fd_month = "Oct "
    Case 11: fd_month = "Nov "
    Case 12: fd_month = "Dec "
    End Select
    fd_time = Format(Time) & " +0200"
    temp = fd_day & Day(Format(Date)) & " " & fd_month & Year(Format(Date, "dd/mm/yyyy")) & " " & fd_time
    find_date = temp
End Function

Function attach_file(attach_str As String) As String
    Dim s As Integer
    Dim temp As String
   
    s = InStr(1, attach_str, "\")
    temp = attach_str
    Do While s > 0
        temp = Mid(temp, s + 1, Len(temp))
        s = InStr(1, temp, "\")
    Loop
    attach_file = temp
End Function

Function encode_the_file(attach_str As String) As String
    Dim blocksize As Long
    Dim Buffer As String
    Dim s As String
    Dim i As Long
    Dim temp As String
    
    Open attach_str For Binary Access Read As #1
    blocksize = 3
    Do While Not EOF(1)
        Buffer = Space(blocksize)
        Get 1, , Buffer
        s = s & base64_encode_string(Buffer)
        DoEvents
    Loop
    Close #1
    For i = 1 To Len(s) Step 76
        'FIXIT: Replace 'Mid' function with 'Mid$' function                                        FixIT90210ae-R9757-R1B8ZE
        temp = temp & Mid(s, i, 76) & vbCrLf
    Next i
    'FIXIT: Replace 'Mid' function with 'Mid$' function                                        FixIT90210ae-R9757-R1B8ZE
    temp = Mid(temp, 1, Len(temp) - 2)
    encode_the_file = temp
End Function

Sub send_email(email_to As String, email_from As String, subject As String, message_text As String, attach As String)
    Const boundary = "Hapoel_Tel_Aviv"
    
    Dim se_body As String
    Dim se_date As String
    Dim se_from As String
    Dim se_to As String
    Dim se_mime As String
    Dim se_content_type As String
    Dim se_content_type_message As String
    Dim se_content_type_attach As String
    Dim x_mailer As String
    Dim x_oem As String
     
    se_date = "Date: " & find_date
    se_from = "From: " & email_from
    se_to = "To: " & email_to
    subject = "Subject: " & subject
    
    se_mime = "MIME-Version: 1.0"
    se_content_type = "Content-Type: multipart/mixed;" & vbCrLf _
                    & vbTab & "boundary = " & """" & boundary & """"
    
    x_oem = "X-OEM: zubin"
    x_mailer = "X-Mailer: " & """" & "Mailer" & """" & " - by ag v1.0"
    
    se_content_type_message = "This is a multi-part message in MIME format." & vbCrLf _
                            & "--" & boundary & vbCrLf _
                            & "Content-Type: text/plain;" & vbCrLf _
                            & vbTab & "charset=" & """" & "iso-8859-1" & """" & vbCrLf _
                            & "Content-Transfer-Encoding: 7bit"
        
    If Len(txt_attach.tEXT) > 0 Then
        se_content_type_attach = "--" & boundary & vbCrLf _
                               & "Content-Type: application/octet-stream;" & vbCrLf _
                               & vbTab & "name=" & attach_file(txt_attach.tEXT) & vbCrLf _
                               & "Content-Transfer-Encoding: base64" & vbCrLf _
                               & "Content-Disposition: attachment;" & vbCrLf _
                               & vbTab & "filename=" & attach_file(txt_attach.tEXT) & vbCrLf _
                               & vbCrLf _
                               & encode_the_file(txt_attach.tEXT)
    End If
    
    se_body = se_from & vbCrLf _
            & se_to & vbCrLf _
            & subject & vbCrLf _
            & se_date & vbCrLf _
            & se_mime & vbCrLf _
            & x_oem & vbCrLf _
            & x_mailer & vbCrLf _
            & se_content_type & vbCrLf _
            & vbCrLf _
            & se_content_type_message & vbCrLf _
            & vbCrLf _
            & message_text & vbCrLf _
            & vbCrLf _
            & se_content_type_attach & vbCrLf _
            & "." & vbCrLf
    
    txt_status.tEXT = "Enviando Mensagem..." & vbCrLf & txt_status.tEXT & vbCrLf
    'FIXIT: Replace 'Left' function with 'Left$' function                                      FixIT90210ae-R9757-R1B8ZE
    Winsock1.SendData "HELO " & Left(email_from, InStr(1, email_from, "@") - 1) & vbCrLf
    wait_for "250"
    Winsock1.SendData "MAIL FROM: " & email_from & vbCrLf
    wait_for "250"
    Winsock1.SendData "RCPT TO: " & email_to & vbCrLf
    wait_for "250"
    Winsock1.SendData "DATA" & vbCrLf
    wait_for "354"
    Winsock1.SendData se_body
    wait_for "250"
    Winsock1.SendData "QUIT" & vbCrLf
    wait_for "221"
    txt_status.tEXT = "Mensagem Enviada." & vbCrLf & txt_status.tEXT & vbCrLf
    Winsock1.Close
    DoEvents
End Sub

Sub init_me()
    txt_status.tEXT = "Pronto." & vbCrLf
    Response = ""
End Sub

Function form_errors() As Boolean
    form_errors = False
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
    form_errors = True
End Function

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub CmdEnviar_Click()
    Select Case Lista.ListIndex

    Case 0
        Command3_Click
    Case 1
        Command2_Click
    Case 2
        Command1_Click
    Case 3
        Command4_Click
    Case 4
        Command5_Click
    Case 5
        Command6_Click
    End Select
End Sub

Private Sub CmdLimpa_Click()
    txt_status.tEXT = ""
End Sub

Private Sub Command1_Click()
    If form_errors Then
        Screen.MousePointer = vbHourglass
      
        Winsock1.LocalPort = 0
        Winsock1.RemoteHost = txt_smtp_server.tEXT 'smtp.server.com
        Winsock1.RemotePort = Txt_Porta.tEXT     '25
        Winsock1.Connect
      
        txt_status.tEXT = "Conectado em: " & txt_smtp_server.tEXT & "." & vbCrLf & txt_status.tEXT & vbCrLf
        send_email txt_email_to.tEXT, txt_email_from.tEXT, txt_subject.tEXT, txt_message_text.tEXT, txt_attach.tEXT
      
        'connect_to_smtp_server txt_smtp_server
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Command2_Click()
    If form_errors Then
        SendEMail txt_email_to.tEXT, txt_subject.tEXT, txt_message_text.tEXT, txt_attach.tEXT
    End If
End Sub

Private Sub Command3_Click()
    If Not form_errors Then
        Exit Sub
    End If
    MAPISession1.LogonUI = True
    MAPISession1.SignOn
    MAPIMessages1.SessionID = MAPISession1.SessionID
    MAPIMessages1.Compose
     
    '////////////////////////////////////////
    '// Informa o email, se enviar para mais de 1 email י separar com(;)e ir informando
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


Private Sub Command4_Click_old()
    Dim Msg As CDO.Message
    Dim Cof As CDO.Configuration
    Dim Camp
    Set Msg = New CDO.Message
    Set Cof = New CDO.Configuration
    Set Camp = Cof.Fields
    On Error GoTo trataerro
    If form_errors Then
        Screen.MousePointer = vbHourglass
        txt_status.tEXT = "Preparando Coneccao"
        With Camp
            .item(cdoSendUsingMethod) = 2        ' cdoSendUsingPort
            .item(cdoSMTPServer) = txt_smtp_server.tEXT
            .item(cdoSMTPConnectionTimeout) = 20 ' quick timeout
            '.Item(cdoSMTPAuthenticate) = 0
            '.Item(cdoSMTPAuthenticate) = 1
            '.Item(cdoSendUserName) = User.Text
            '.Item(cdoSendPassword) = LCase(PWR.Text)
            .Update
        End With
        txt_status.tEXT = "Preparando Mensagem"
        With Msg
            Set .Configuration = Cof
            .To = txt_email_to.tEXT
            .From = txt_email_from.tEXT
            .subject = txt_subject.tEXT
            .TextBody = txt_message_text.tEXT
            '.CC = CC.Text
            .AddAttachment txt_attach.tEXT
            txt_status.tEXT = "Enviando Mensagem"
            .Send
        End With
        txt_status.tEXT = "Mensagem Enviada"
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
trataerro:
    Select Case Err.Number
    Case Else
        SayErro "Email CDO", True
        Exit Sub
    End Select
End Sub

Private Sub Command4_Click()
   Dim retVal          As String
   retVal = SendMail(Trim$(txt_email_to.tEXT), _
        Trim$(txt_subject.tEXT), _
        Trim$(txtFromName.tEXT) & "<" & Trim$(txt_email_from.tEXT) & ">", _
        Trim$(txt_message_text.tEXT), _
        Trim$(txt_smtp_server.tEXT), _
        CInt(Trim$(Txt_Porta.tEXT)), _
        Trim$(txtUsername.tEXT), _
        Trim$(txtPassword.tEXT), _
        Trim$(txt_attach.tEXT), _
        CBool(chkSSL.Value))
   txt_status.tEXT = retVal

End Sub

Private Sub Command5_Click()
    Dim poSendMail As vbSendMail.clsSendMail
    On Error GoTo trataerro
    If form_errors Then
        Screen.MousePointer = vbHourglass
        txt_status.tEXT = "Preparando Coneccao"
        Set poSendMail = New clsSendMail
        With poSendMail
            .SMTPHostValidation = VALIDATE_HOST_DNS
            .EmailAddressValidation = VALIDATE_SYNTAX
            .Delimiter = ";"
        End With
        With poSendMail
            ' **************************************************************************
            ' Set the basic properties common to all messages to be sent
            ' **************************************************************************
            .SMTPHost = txt_smtp_server.tEXT     ' Required the fist time, optional thereafter
            .From = txt_email_from.tEXT          ' Required the fist time, optional thereafter
            .FromDisplayName = txt_email_from.tEXT ' Optional, saved after first use
            .Message = txt_message_text.tEXT     ' Optional
            .Attachment = Trim(txt_attach.tEXT)  ' Optional, separate multiple entries with delimiter character
            '     .UseAuthentication = bAuthLogin            ' Optional, default = FALSE
            '     .UserName = txtUserName.Text               ' Optional, default = Null String
            '     .Password = txtPassword.Text               ' Optional, default = Null String, value is NOT saved
            If .Connect Then
                txt_status.tEXT = "Enviando Mensagem"
                .Recipient = txt_email_to.tEXT
                .RecipientDisplayName = txt_email_to.tEXT
                .subject = txt_subject.tEXT
                .Send
                .Disconnect
            End If

        End With
        txt_status.tEXT = "Mensagem Enviada"
        Screen.MousePointer = vbDefault
    End If
    Exit Sub
trataerro:
    Select Case Err.Number
    Case Else
        SayErro "Email CDO", True
        Exit Sub
    End Select

End Sub

Private Sub Command6_Click()
    Dim K As Variant
    Dim Retorno As Variant

    If Not form_errors Then
        Exit Sub
    End If

    Set K = CreateObject("Py.SendMail")
    
    K.SMTPServer = txt_smtp_server.tEXT          'txtServidor.Text
    
    
    '    K.Recipient = "" 'txt_email_to.Text 'txtNomePara.Text
    K.To = txt_email_to.tEXT                     'txtEmailPara.Text
    
    '   K.Sender = txt_email_from.Text 'txtNomeDe.Text
    K.From = txt_email_from.tEXT                 'txtEmailDe.Text
    
    'K.CC = txtCc.Text
    'K.BCC = txtBcc.Text
    'K.CC = ""
    'K.BCC = ""
    K.subject = txt_subject.tEXT                 'txtAssunto.Text
    K.Body = txt_message_text.tEXT               'txtMensagem.Text
    
    'K.PureText = chkTextoPuro.Value
    '  K.puretext = 1
    
    'Em caso de autenticaחדo
    'K.User = "teste+graymalkinsoftware.net"
    'K.Password = "1234"
    'K.User = ""
    'K.Password = ""
    
    'For i = 0 To lstAnexos.ListCount - 1
    '    K.AttachFile lstAnexos.List(i)
    'Next i
    K.AttachFile txt_attach.tEXT
    
    'For i = 0 To lstImagens.ListCount - 1
    '    K.InsertImage lstImagens.List(i)
    'Next i
       
    Retorno = K.Send()
    txt_status.tEXT = Retorno
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
    '    CentralizaJanela Me
    Center Me
    init_me
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
    Winsock1.Close
    Unload Me
End Sub

Private Sub List1_Click()
    CmdEnviar_Click
End Sub

Private Sub Lista_Click()
    CmdEnviar_Click
End Sub

'Private Sub mnu_send_Click()
'    If form_errors = False Then
'    connect_to_smtp_server txt_smtp_server
'    End If
'End Sub
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

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Winsock1.GetData Response
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    txt_status.tEXT = "Erro: " & Description & "." & vbCrLf & txt_status.tEXT & vbCrLf
End Sub

Function bin(ByVal dec_number As Integer) As String
    Dim temp As String
    Dim Count As Integer
    Dim X As Integer
    Dim Length As Integer
    
    Count = 0
    temp = ""
    X = 128
    Do While Count < dec_number
        If X > dec_number Then
            temp = temp & "0"
        ElseIf Count + X > dec_number Then
            temp = temp & "0"
        Else
            temp = temp & "1"
            Count = Count + X
        End If
        X = X - (X / 2)
    Loop
    Length = Len(temp)
    For X = (Length + 1) To 8
        temp = temp & 0
    Next X
    bin = temp
End Function

Function dec(ByVal bin_number As String) As Integer
    Dim temp As Integer
    'FIXIT: Declare 'x' with an early-bound data type                                          FixIT90210ae-R1672-R1B8ZE
    Dim X, c As Integer
    Dim s As String
    
    temp = 0
    X = 128
    c = 1
    Do While c <= 8
        'FIXIT: Replace 'Mid' function with 'Mid$' function                                        FixIT90210ae-R9757-R1B8ZE
        s = Mid(bin_number, c, 1)
        If s = "1" Then
            temp = temp + X
        End If
        X = X - (X / 2)
        c = c + 1
    Loop
    dec = temp
End Function

Function base64_alphabet(ByVal dec_num As Integer) As String
    Dim temp As String
    
    Select Case dec_num
    Case 0: temp = "A"
    Case 1: temp = "B"
    Case 2: temp = "C"
    Case 3: temp = "D"
    Case 4: temp = "E"
    Case 5: temp = "F"
    Case 6: temp = "G"
    Case 7: temp = "H"
    Case 8: temp = "I"
    Case 9: temp = "J"
    Case 10: temp = "K"
    Case 11: temp = "L"
    Case 12: temp = "M"
    Case 13: temp = "N"
    Case 14: temp = "O"
    Case 15: temp = "P"
    Case 16: temp = "Q"
    Case 17: temp = "R"
    Case 18: temp = "S"
    Case 19: temp = "T"
    Case 20: temp = "U"
    Case 21: temp = "V"
    Case 22: temp = "W"
    Case 23: temp = "X"
    Case 24: temp = "Y"
    Case 25: temp = "Z"
    Case 26: temp = "a"
    Case 27: temp = "b"
    Case 28: temp = "c"
    Case 29: temp = "d"
    Case 30: temp = "e"
    Case 31: temp = "f"
    Case 32: temp = "g"
    Case 33: temp = "h"
    Case 34: temp = "i"
    Case 35: temp = "j"
    Case 36: temp = "k"
    Case 37: temp = "l"
    Case 38: temp = "m"
    Case 39: temp = "n"
    Case 40: temp = "o"
    Case 41: temp = "p"
    Case 42: temp = "q"
    Case 43: temp = "r"
    Case 44: temp = "s"
    Case 45: temp = "t"
    Case 46: temp = "u"
    Case 47: temp = "v"
    Case 48: temp = "w"
    Case 49: temp = "x"
    Case 50: temp = "y"
    Case 51: temp = "z"
    Case 52: temp = "0"
    Case 53: temp = "1"
    Case 54: temp = "2"
    Case 55: temp = "3"
    Case 56: temp = "4"
    Case 57: temp = "5"
    Case 58: temp = "6"
    Case 59: temp = "7"
    Case 60: temp = "8"
    Case 61: temp = "9"
    Case 62: temp = "+"
    Case 63: temp = "/"
    End Select
    base64_alphabet = temp
End Function

Function base64_encode(ByVal str_24bits As String) As String
    Dim temp As String
    'FIXIT: Declare 'x' and 'v' with an early-bound data type                                  FixIT90210ae-R1672-R1B8ZE
    Dim X, v, w As String
    'FIXIT: Declare 'i' and 'z' with an early-bound data type                                  FixIT90210ae-R1672-R1B8ZE
    Dim i, z, Y As Integer
    Dim bits_6(4) As String
    Dim bits_8(4) As String
    Dim dec_num(4) As Integer
    Dim base64_val(4) As String
    
    X = ""
    v = ""
    w = ""
    temp = ""
    z = Len(str_24bits)
    For i = 1 To z
        'FIXIT: Replace 'Mid' function with 'Mid$' function                                        FixIT90210ae-R9757-R1B8ZE
        w = Mid(str_24bits, i, 1)
        Y = Asc(w)
        v = v & bin(Y)
    Next i
    If z < 3 Then
        For i = (z + 1) To 3
            X = X & "00000000"
        Next i
    End If
    v = v & X
    z = 1
    For i = 1 To 4
        'FIXIT: Replace 'Mid' function with 'Mid$' function                                        FixIT90210ae-R9757-R1B8ZE
        X = Mid(v, z, 6)
        z = z + 6
        bits_6(i) = X
    Next i
    For i = 1 To 4
        bits_8(i) = "00" & bits_6(i)
    Next i
    For i = 1 To 4
        dec_num(i) = dec(bits_8(i))
    Next i
    For i = 1 To 4
        base64_val(i) = base64_alphabet(dec_num(i))
    Next i
    For i = 1 To 4
        temp = temp & base64_val(i)
    Next i
    base64_encode = temp
End Function

Function base64_encode_string(str As String) As String
    Dim temp As String
    'FIXIT: Declare 'i' and 'x' and 'v' with an early-bound data type                          FixIT90210ae-R1672-R1B8ZE
    Dim i, X, v, Y As Long
    'FIXIT: Declare 's' with an early-bound data type                                          FixIT90210ae-R1672-R1B8ZE
    Dim s, u As String
    Dim crlf() As String
    
    X = Len(str)
    i = X / 76
    i = i + 2
    ReDim crlf(i)
    For i = 1 To X Step 3
        'FIXIT: Replace 'Mid' function with 'Mid$' function                                        FixIT90210ae-R9757-R1B8ZE
        s = Mid(str, i, 3)
        u = base64_encode(s)
        temp = temp & u
    Next i
    u = ""
    i = 0
    v = 1
    Y = 1
    X = Len(temp)
    Do While i <= X
        i = i + 1
        If (i Mod 76) = 0 Then
            'FIXIT: Replace 'Mid' function with 'Mid$' function                                        FixIT90210ae-R9757-R1B8ZE
            s = Mid(temp, v, 76)
            crlf(Y) = s
            Y = Y + 1
            v = v + 76
        End If
        'FIXIT: Replace 'Mid' function with 'Mid$' function                                        FixIT90210ae-R9757-R1B8ZE
        crlf(Y) = Mid(temp, v, (X + 1) - v)
    Loop
    For i = 1 To Y - 1
        crlf(i) = crlf(i) & vbCrLf
    Next i
    temp = ""
    For i = 1 To Y
        temp = temp & crlf(i)
    Next i
    base64_encode_string = temp
End Function

Public Function SendMail(sTo As String, sSubject As String, sFrom As String, _
    sBody As String, sSmtpServer As String, iSmtpPort As Integer, _
    sSmtpUser As String, sSmtpPword As String, _
    sFilePath As String, bSmtpSSL As Boolean) As String
      
    On Error GoTo SendMail_Error:
    Dim lobj_cdomsg      As CDO.Message
    Set lobj_cdomsg = New CDO.Message
    lobj_cdomsg.Configuration.Fields(cdoSMTPServer) = sSmtpServer
    lobj_cdomsg.Configuration.Fields(cdoSMTPServerPort) = iSmtpPort
    lobj_cdomsg.Configuration.Fields(cdoSMTPUseSSL) = bSmtpSSL
    lobj_cdomsg.Configuration.Fields(cdoSMTPAuthenticate) = cdoBasic
    lobj_cdomsg.Configuration.Fields(cdoSendUserName) = sSmtpUser
    lobj_cdomsg.Configuration.Fields(cdoSendPassword) = sSmtpPword
    lobj_cdomsg.Configuration.Fields(cdoSMTPConnectionTimeout) = 30
    lobj_cdomsg.Configuration.Fields(cdoSendUsingMethod) = cdoSendUsingPort
    lobj_cdomsg.Configuration.Fields.Update
    lobj_cdomsg.To = sTo
    lobj_cdomsg.From = sFrom
    lobj_cdomsg.subject = sSubject
    lobj_cdomsg.TextBody = sBody
    If Trim$(sFilePath) <> vbNullString Then
        lobj_cdomsg.AddAttachment (sFilePath)
    End If
    lobj_cdomsg.Send
    Set lobj_cdomsg = Nothing
    SendMail = "ok"
    Exit Function
          
SendMail_Error:
    SendMail = Err.Description
End Function

