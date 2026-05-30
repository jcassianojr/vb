Attribute VB_Name = "MyFunctionsMail"
Option Explicit

' --- CONSTANTES DE CONFIGURAÇÃO CDO (Não exigem referências no projeto) ---
Public Const cdoSendUsingPort = 2
Public Const cdoSMTPServer = "http://schemas.microsoft.com/cdo/configuration/smtpserver"
Public Const cdoSMTPServerPort = "http://schemas.microsoft.com/cdo/configuration/smtpserverport"
Public Const cdoSMTPAuthenticate = "http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"
Public Const cdoSMTPUseSSL = "http://schemas.microsoft.com/cdo/configuration/smtpusessl"
Public Const cdoSendUserName = "http://schemas.microsoft.com/cdo/configuration/sendusername"
Public Const cdoSendPassword = "http://schemas.microsoft.com/cdo/configuration/sendpassword"
Public Const cdoSMTPConnectionTimeout = "http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout"
Public Const cdoBasic = 1

' --- FUNÇÃO SENDMAILCDO (Assinatura original preservada) ---
Public Function SendMailCDO(sTo As String, sSubject As String, sFrom As String, _
                            sBody As String, sSmtpServer As String, iSmtpPort As Integer, _
                            sSmtpUser As String, sSmtpPword As String, _
                            sFilePath As String, bSmtpSSL As Boolean) As String

    Dim lobj_cdomsg As Object
    Dim lobj_config As Object
    
    ' Criando os objetos via Late Binding (compatível com 32 e 64 bits)
    On Error GoTo Erro_CDO
    Set lobj_cdomsg = CreateObject("CDO.Message")
    Set lobj_config = CreateObject("CDO.Configuration")
    
    ' Vincula a configuração à mensagem
    Set lobj_cdomsg.Configuration = lobj_config
    
    ' Configurando os campos de conexão
    With lobj_config.Fields
        .item(cdoSMTPServer) = sSmtpServer
        .item(cdoSMTPServerPort) = iSmtpPort
        .item(cdoSMTPUseSSL) = bSmtpSSL
        .item(cdoSMTPAuthenticate) = cdoBasic
        .item(cdoSendUserName) = sSmtpUser
        .item(cdoSendPassword) = sSmtpPword
        .item(cdoSMTPConnectionTimeout) = 60
        .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPort
        .Update
    End With

    ' Preenchendo a mensagem
    lobj_cdomsg.To = sTo
    lobj_cdomsg.From = sFrom
    lobj_cdomsg.subject = sSubject
    lobj_cdomsg.TextBody = sBody
    
    ' Tratamento de anexo
    If Len(Trim$(sFilePath)) > 0 Then
        If Dir(sFilePath) <> "" Then
            lobj_cdomsg.AddAttachment sFilePath
        End If
    End If

    ' Enviando
    lobj_cdomsg.Send
    
    ' Sucesso
    SendMailCDO = "ok"
    
Limpeza:
    Set lobj_cdomsg = Nothing
    Set lobj_config = Nothing
    Exit Function

Erro_CDO:
    SendMailCDO = "Erro: " & Err.Description
    Resume Limpeza
End Function

