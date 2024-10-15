Attribute VB_Name = "MyFunctionsMail"
Option Explicit

Public Function SendMailCDO(sTo As String, sSubject As String, sFrom As String, _
                            sBody As String, sSmtpServer As String, iSmtpPort As Integer, _
                            sSmtpUser As String, sSmtpPword As String, _
                            sFilePath As String, bSmtpSSL As Boolean) As String

  On Error GoTo SendMail_Error:
  Dim lobj_cdomsg As CDO.Message
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
  SendMailCDO = "ok"
  Exit Function

SendMail_Error:
  SendMailCDO = Err.Description
End Function

Public Sub SendEmailShell(ByVal sEmailAddress As String, Optional ByVal sSubject As String = "", Optional ByVal sBody As String = "", Optional ByVal cARQ As String = "")
  Dim nTENTA As Integer
  Dim ret As Integer
  nTENTA = 0
  Rem HotTo use:
  Rem Call ShellExecute(0, "Open", "mailto:email_address?Subject=your_subject?Body=body_message", "", "", 1)
  Rem sBody = sCR & "Dear friend" & sCR & sCR  & "I will..."
  Rem --------------------------------------------------
  ''  ShellExecute 0, "Open", "mailto:" & sEmailAddress & "?Subject=" & sSubject & "?Body=" & sBody, "", "", SW_SHOWNORMAL
  If Len(cARQ) > 0 Then
    ShellExecute 0, "Open", "mailto:" & sEmailAddress & "?Subject=" & cARQ & "&Body=" & sBody, "", "", SW_SHOWNORMAL
    ' Wait until Outlook Express is ready.
    While ret = 0 And nTENTA < 11
      DoEvents
      ret = FindWindow(vbNullString, cARQ)
      nTENTA = nTENTA + 1
    Wend
    ' Send keys Alt-I-A, the zip file name,
    ' two TABs, and Enter.
    SendKeys "%ia" & cARQ & "{TAB}{TAB}{ENTER}"
  Else
    ShellExecute 0, "Open", "mailto:" & sEmailAddress & "?Subject=" & sSubject & "&Body=" & sBody, "", "", SW_SHOWNORMAL
  End If
End Sub



