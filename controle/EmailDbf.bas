Attribute VB_Name = "EmailDbf"
Option Explicit
Public Function MAILENV(cCODIGO, cTexto)
End Function

Public Function xMAILENV(cCODIGO, cTexto)
    ''Dim oDB As ADODB.Connection
    Dim oMAILPARA As ADODB.Recordset
    Dim oMAILERRO As ADODB.Recordset
    Dim cMAILDIR  As String
    Dim cMAILDIR2 As String
    Dim cARQTMP As String
    Dim cCONSQL As String
    Dim aRETU As Variant
    Dim oMAIL As Variant
    Dim nNUMERO As Variant
    Dim poSendMail As vbSendMail.clsSendMail
    Dim cASSUNTO As String
    Dim txt_email_from As String

    On Error GoTo error
    cASSUNTO = cCODIGO

    cMAILDIR = PegPath("PATH", "MAIL")
    cMAILDIR2 = PegPath("PATH", "MAIL2")
    If Not FileExist(cMAILDIR2 & "MAIL.DBF") Then
        If Not FileExist(cMAILDIR2 & "\" & "MAIL.DBF") Then
            Alert "Falta Arquivo Email"
        End If
    End If
    
    cARQTMP = GeraConn(cMAILDIR, "JETFOX")
    aRETU = TipoConn(cARQTMP)
    cARQTMP = aRETU(1)
          
    Dim DB As New ADODB.Connection
    DB.CursorLocation = adUseClient
    DB.ConnectionTimeout = 120
    
    
    DB.Open cARQTMP
    Set oMAILPARA = New ADODB.Recordset
    Set oMAILERRO = New ADODB.Recordset
    
    cCONSQL = "select * from mailpara WHERE erro='" & cCODIGO & "'"
    oMAILPARA.Open cCONSQL, DB, adOpenForwardOnly, adLockReadOnly
    cCONSQL = "select * from mailerro WHERE erro='" & cCODIGO & "'"
    oMAILERRO.Open cCONSQL, DB, adOpenForwardOnly, adLockReadOnly
    
    
    sx_SetEpoch (Year(Date) - 30)
    sx_SetDateFormat BRITISH
    oMAIL = sx_Use(cMAILDIR2 & "MAIL.dbf", "MAIL", READWRITE, SDEFOX)
    sx_GoBottom
    nNUMERO = sx_GetVariant("NUMERO")
    nNUMERO = nNUMERO + 1
    If Not oMAILERRO.EOF Then
        cASSUNTO = FixStr(oMAILERRO("ASSUNTO"))
        If Not oMAILPARA.EOF Then
            ''oMAILPARA.MoveFirst
            While Not oMAILPARA.EOF
                sx_AppendBlank
                If sx_Rlock(sx_RecNo()) Then
                    sx_PutVariant "NUMERO", nNUMERO
                    sx_PutVariant "ERRO", FixStr(cCODIGO)
                    sx_PutVariant "DATA", FixStr(Date)
                    sx_PutVariant "HORA", FixStr(Time)
                    sx_PutVariant "DE", FixStr(zUSER)
                    sx_PutVariant "DESTINO", FixStr(oMAILPARA("DESTINO"))
                    sx_PutVariant "ASSUNTO", cASSUNTO
                    sx_PutVariant "TEXTO", FixStr(cTexto)
                    sx_Commit
                    sx_Unlock sx_RecNo()
                End If
                oMAILPARA.MoveNext
            Wend
        End If
    End If
    sx_CloseAll
    oMAILPARA.Close
    oMAILERRO.Close
    DB.Close
        
    Screen.MousePointer = vbHourglass
    'txt_status.tEXT = "Preparando Coneccao"
    Set poSendMail = New clsSendMail
    With poSendMail
        .SMTPHostValidation = VALIDATE_HOST_DNS
        .EmailAddressValidation = VALIDATE_SYNTAX
        .Delimiter = ";"
    End With
    With poSendMail
        .SMTPHost = PegPath("EMAIL", "SERVER", "stmp..com.br")
        txt_email_from = PegPath("EMAIL", UCase(zNOMEFOLHA), " ")
        If Len(Trim(txt_email_from)) = 0 Then
            txt_email_from = PegPath("EMAIL", "FROM", "@.com.br")
        End If
        .From = txt_email_from
        .FromDisplayName = txt_email_from
        .Message = FixStr(cTexto)
        If .Connect Then
            'txt_status.tEXT = "Enviando Mensagem"
            .Recipient = PegPath("EMAIL", cCODIGO, "")
            .RecipientDisplayName = ""
            .subject = cASSUNTO
            .Send
            .Disconnect
            
        End If
    End With
    Screen.MousePointer = vbDefault

        
        
        
    Exit Function
error:
    Select Case Err.Number
    Case Else
        SayErro "Mail Env"
        Screen.MousePointer = vbDefault
        Resume Next
    End Select

End Function

