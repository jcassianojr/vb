Attribute VB_Name = "ImagensMain"
Option Explicit
Public objConsole As New clsConsole

Sub Main()
    Cmdiniciar_Click
End Sub

Public Sub Cmdiniciar_Click()
    'DBCONNSTR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\modelos\PECAS\FMP04CPF.MDB;Mode=Share Deny None"
    Dim OBJCONN As ADODB.Connection
    Dim OBJRSGLOB As ADODB.Recordset
    Dim strSQL As String
    Dim DBCONNSTR As String
    Dim iFileNum As Integer
    Dim lFileLength As Long
    Dim abBytes() As Byte
    Dim sTEMPFILE As String
    Dim cCAMPO As String
    Dim cCAMJPG As String
    Dim cCODIGO As String
    Dim rawPath As String
  
    Set OBJCONN = New ADODB.Connection           'Create a new object
    Set OBJRSGLOB = New ADODB.Recordset



    cCAMPO = "IMAGEM"
    DBCONNSTR = PegPath("PATH", "IMGCOM")
    cCAMJPG = PegPath("PATH", "IMGJPG")
    rawPath = PegPath("PATH", "IMGCOM")
  
    ' LÓGICA DE CONSTRUÇÃO DA CONNECTION STRING
    If InStr(1, rawPath, "Provider=", vbTextCompare) > 0 Or _
       InStr(1, rawPath, "Driver=", vbTextCompare) > 0 Then
        ' Se já for uma string de conexão completa, usa como está
        DBCONNSTR = rawPath
    Else
        ' Se for apenas o caminho de um arquivo, identifica a extensão
        Dim ext As String
        ext = LCase(Right(rawPath, 7))           ' Pega o final para extensões como .sqlite

        If EArquivoSQLite(rawPath) Then
            ' Monta string para SQLite3 ODBC [cite: 11]
            DBCONNSTR = "Driver={SQLite3 ODBC Driver};Database=" & rawPath & ";"
        ElseIf InStr(ext, ".accdb") > 0 Then
            ' Monta string para Access 2007-2016 (Engine ACE)
            DBCONNSTR = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & rawPath & ";Persist Security Info=False;"
        ElseIf InStr(ext, ".mdb") > 0 Then
            ' Monta string para Access antigo (Jet)
            DBCONNSTR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & rawPath & ";Mode=Share Deny None;"
        Else
            ' Fallback ou erro caso não reconheça a extensão
            DBCONNSTR = rawPath
        End If
    End If

    'se sql lite verifica se odbc esta instalado
    If InStr(LCase(DBCONNSTR), ".sqlite") > 0 Or InStr(LCase(DBCONNSTR), ".sqlite3") > 0 _
       Or InStr(LCase(DBCONNSTR), ".fossil") > 0 Or InStr(LCase(DBCONNSTR), ".db3") > 0 _
       Or (InStr(LCase(DBCONNSTR), ".db") > 0) Then
        If IsSQLiteDriverInstalled() Then
            ' MsgBox "O driver SQLite ODBC está pronto para uso!", vbInformation
        Else
            MsgBox "Driver SQLite ODBC não encontrado. Por favor, instale o driver para continuar.", vbCritical
            End
        End If
    End If

    ' DBCONNSTR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\modelos\PECAS\FMP04CPF.MDB;Mode=Share Deny None"
    ' cCAMJPG = "d:\modelos\rh\fotoscpf\"
    OBJCONN.Open DBCONNSTR
    strSQL = "SELECT * FROM imagens where numero>0"
    OBJRSGLOB.Open strSQL, OBJCONN, adOpenForwardOnly, adLockOptimistic

    If Not OBJRSGLOB.RecordCount = 0 Then
        Do While (Not OBJRSGLOB.EOF)
            cCODIGO = OBJRSGLOB.Fields("CODIGO")
            Debug.Print cCODIGO
            '      WriteToConsole cCODIGO
            If IsNull(OBJRSGLOB(cCAMPO)) Then
                lFileLength = 0
            Else
                lFileLength = LenB(OBJRSGLOB(cCAMPO))
            End If
            If lFileLength > 1 Then              ' Ja carregado
            Else
                sTEMPFILE = cCAMJPG & Trim(TiraOut(cCODIGO)) + ".JPG"
                If FileExists(sTEMPFILE) Then
                    iFileNum = FreeFile
                    Open sTEMPFILE For Binary Access Read As #iFileNum
                    lFileLength = LOF(iFileNum)
                    ReDim abBytes(lFileLength)
                    Get #iFileNum, , abBytes()
                    Close #iFileNum
                    OBJRSGLOB.Fields(cCAMPO).AppendChunk abBytes()
                    OBJRSGLOB.Update
                End If
            End If
            OBJRSGLOB.MoveNext
        Loop
    End If


End Sub

Public Function TiraOut(ByVal eVAR As Variant) As String
    Dim cTexto As String
    cTexto = FixStr(eVAR)
    cTexto = Replace(cTexto, "-", "")
    cTexto = Replace(cTexto, ",", "")
    cTexto = Replace(cTexto, ".", "")
    cTexto = Replace(cTexto, ":", "")
    cTexto = Replace(cTexto, "/", "")
    cTexto = Replace(cTexto, ";", "")
    cTexto = Replace(cTexto, "*", "")
    cTexto = Replace(cTexto, "(", "")
    cTexto = Replace(cTexto, ")", "")
    TiraOut = cTexto
End Function

Public Function FileExists(ByRef sFileName As String) As Boolean
    On Error Resume Next
    FileExists = PathFileExists(sFileName)
End Function

Public Function FixStr(ByVal eVAR As Variant, _
                       Optional ByVal ePAD As Variant = "", _
                       Optional ByVal coper As String = "", _
                       Optional ByVal nLEN As Integer = 0) As Variant
    On Error GoTo errhandler
    If IsNull(eVAR) Then
        If ePAD <> "" Then
            eVAR = ePAD
        End If
    End If
    FixStr = CStr(eVAR)
    If UCase(eVAR) = "NULL" Then
        FixStr = ePAD
    End If
    If InStr(coper, "TRIM") > 0 Then
        eVAR = Trim(eVAR)
        FixStr = eVAR
    End If
    If nLEN > 0 And Len(eVAR) > nLEN Then
        eVAR = Mid(eVAR, 1, nLEN)
        FixStr = eVAR
    End If
    Exit Function
errhandler:
    FixStr = ""
    Resume Next
End Function

Public Function EArquivoSQLite(ByVal cCaminho As String) As Boolean
  
    cCaminho = LCase(cCaminho)
    ' Lista de extensões que o seu sistema aceita como SQLite
    If InStr(cCaminho, ".sqlite") > 0 Or _
       InStr(cCaminho, ".sqlite3") > 0 Or _
       InStr(cCaminho, ".db") > 0 Or _
       InStr(cCaminho, ".db3") > 0 Or _
       InStr(cCaminho, ".fossil") > 0 Then
       
        EArquivoSQLite = True
    Else
        EArquivoSQLite = False
    End If
End Function


