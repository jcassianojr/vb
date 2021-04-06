Attribute VB_Name = "EmailDbf"
Option Explicit
'mailenv caso o envio nao se faz max necessario
Public Function MAILENV(cCODIGO, cTexto)
End Function
'xmailenv caso o envio aina e necessario
Public Function xMAILENV(cCODIGO, cTexto, Optional ByVal lEMAILINT = True, Optional ByVal lEMAILEXT = True)
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
    Dim cASSUNTO As String
    Dim txt_email_from As String

    On Error GoTo error
    cASSUNTO = cCODIGO

If lEMAILINT Then
    
    cMAILDIR = PegPath("PATH", "MAIL")
    cMAILDIR2 = PegPath("PATH", "MAIL2")
    If Not FileExist(cMAILDIR2 & "MAIL.DBF") Then
        If Not FileExist(cMAILDIR2 & "\" & "MAIL.DBF") Then
            Alert "Falta Arquivo Email"
        End If
    End If
    nNUMERO = PegMAXSQL(GeraConn(cMAILDIR, "FOX"), "MAIL", "NUMERO", 1) '+1 fica no loop
    
    cARQTMP = GeraConn(cMAILDIR, "JETFOX")
    aRETU = TipoConn(cARQTMP)
    cARQTMP = aRETU(1)
          
    Dim DB As New ADODB.Connection
    DB.CursorLocation = adUseClient
    DB.ConnectionTimeout = 120
    
    
    DB.Open cARQTMP
    Set oMAILPARA = New ADODB.Recordset
    Set oMAILERRO = New ADODB.Recordset
    Set oMAIL = New ADODB.Recordset
    
    VFPSetValues DB
    'set delete on e set null false(permite nao gravar todos os campos da tabela)
    
    cCONSQL = "select * from mailpara WHERE erro='" & cCODIGO & "'"
    oMAILPARA.Open cCONSQL, DB, adOpenForwardOnly, adLockReadOnly
    cCONSQL = "select * from mailerro WHERE erro='" & cCODIGO & "'"
    oMAILERRO.Open cCONSQL, DB, adOpenForwardOnly, adLockReadOnly
    cCONSQL = "select * from mail  WHERE NUMERO=-1" 'usa numero-1 apenas para abrir pois abrir como table demora muito
    oMAIL.Open cCONSQL, DB, adOpenStatic, adLockOptimistic
    
    
   
    If Not oMAILERRO.EOF Then
        cASSUNTO = FixStr(oMAILERRO("ASSUNTO"))
        If Not oMAILPARA.EOF Then
            While Not oMAILPARA.EOF
                nNUMERO = nNUMERO + 1
                oMAIL.AddNew
                oMAIL("NUMERO") = nNUMERO
                oMAIL("ERRO") = FixStr(cCODIGO)
                oMAIL("DATA") = FixStr(Date)
                oMAIL("HORA") = FixStr(Time)
                oMAIL("DE") = FixStr(zUSER)
                oMAIL("DESTINO") = FixStr(oMAILPARA("DESTINO"))
                oMAIL("ASSUNTO") = cASSUNTO
                oMAIL("TEXTO") = FixStr(cTexto)
                oMAILPARA.MoveNext
            Wend
        End If
    End If
    oMAILPARA.Close
    oMAILERRO.Close
    DB.Close
End If
    
'sendmail via cdo
'SendMail(sTo As String, sSubject As String, sFrom As String, _
'    sBody As String, sSmtpServer As String, iSmtpPort As Integer, _
'    sSmtpUser As String, sSmtpPword As String, _
'    sFilePath As String, bSmtpSSL As Boolean) As String
    If lEMAILEXT Then
        Screen.MousePointer = vbHourglass
        SendMailCDO PegPath("EMAIL", cCODIGO, ""), cASSUNTO, PegPath("EMAIL", UCase(zNOMEFOLHA), PegPath("EMAIL", "FROM", "@.com.br")), _
        FixStr(cTexto), PegPath("EMAIL", "SERVER", "stmp..com.br"), FixNum(PegPath("EMAIL", "PORTA", "25")), _
        "", "", _
        "", False
        Screen.MousePointer = vbDefault
    End If
    
Exit Function
error:
    Select Case Err.Number
    Case Else
        SayErro "Mail Env"
        Screen.MousePointer = vbDefault
        Resume Next
    End Select

End Function

