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
Set OBJCONN = New ADODB.Connection 'Create a new object
Set OBJRSGLOB = New ADODB.Recordset

 'Dim scrbuf      As CONSOLE_SCREEN_BUFFER_INFO
    
'Get the standard output handle
  '  hOutput = GetStdHandle(STD_OUTPUT_HANDLE)
  '  GetConsoleScreenBufferInfo hOutput, scrbuf


cCAMPO = "IMAGEM"
DBCONNSTR = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=d:\modelos\PECAS\FMP04CPF.MDB;Mode=Share Deny None"
cCAMJPG = "d:\modelos\rh\fotoscpf\"
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
         If lFileLength > 1 Then ' Ja carregado
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
Public Function StrToArray(ByVal cGRUPO As String) As Variant
Dim X, nLEN As Integer
Dim aUSO As Variant
Dim cCHAR, eCNV As String
   Select Case cGRUPO
      Case "OEM"
           eCNV = "€‚ƒ„…†‡ˆ‰Š‹Œ‘’“”•–—˜™š›œŸ ¡¢£¤¥¦§¨©ª«¬­®¯àáâãäåæçèéêëìíîïğñòóôõö÷øù³Å¿ÄØ"
      Case "ANSI"
            eCNV = "ÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏĞÑÒÓÔÕÖ×ØÙÚÛÜİŞßàáâãäåæçèéêëìíîïğñòóôõö÷øùúûüışÿ¨¸¥´ªº²³¯¿¦†¬­‡"
      Case "MEUANSI"
           eCNV = "ªºÇçÅåÀàÈèÌìÒòÙùÄäËëÏïÖöÜüÂâÊêÎîÔôÛûÁáÉéÍíÓóÚúİıÃãÕõÑñ|"
      Case "MEUOEM"
           ''Numeral
           eCNV = "¦§"
           ''Cidilha
           eCNV = eCNV & "€‡"
           ''Grau
           eCNV = eCNV & "†"
           ''Crase
           eCNV = eCNV & "·…ÔŠŞã•ë—"
           ''Trema
           eCNV = eCNV & "„Ó‰Ø‹™”š"
           ''cIRCUNFLEXO
           eCNV = eCNV & "¶ƒÒˆ×Œâ“ê–"
           ''Agudo
           eCNV = eCNV & "µ ‚Ö¡à¢é£íì"
           '''Til
           eCNV = eCNV & "ÇÆåä¥¤"
           ''Tracos quadro
           eCNV = eCNV & "İ"
           
           ''"_" Caracter so maiscula ou minuscula usado "_" manter tamanho array
      Case "UACENTO"
           eCNV = "ÇÁÉÍÓÚÀÃÕÂÊÔÄÖÜªº__Å"
      Case "LACENTO"
           eCNV = "çáéíóúàãõâêôäöüªºòù_"
      Case "LACETIR"
           eCNV = "caeiouaaoaeoaouaoou_"
      Case "UACETIR"
           eCNV = "CAEIOUAAOAEOAOUao__A"
      Case Else
           eCNV = cGRUPO
   End Select

nLEN = Len(eCNV)
ReDim aUSO(nLEN)
For X = 1 To nLEN
    cCHAR = Mid(eCNV, X, 1)
    Select Case cCHAR
           Case "ª"
                aUSO(X - 1) = "a."

         Case "º"
                aUSO(X - 1) = "o."

           Case Else
                aUSO(X - 1) = cCHAR

    End Select
Next X
StrToArray = aUSO
End Function

Public Function CharConv(ByVal cTexto As String, ByVal eORI As Variant, ByVal eDES As Variant) As String
Dim nLEN, nTEXTO, X, Y As Integer
Dim aORI, aDES, aTEXTO As Variant
If IsArray(eORI) Then
   aORI = eORI
   aDES = eDES
Else
   aORI = StrToArray(CStr(eORI))
   aDES = StrToArray(CStr(eDES))
End If
aTEXTO = StrToArray(cTexto)
nLEN = UBound(aORI)
nTEXTO = UBound(aTEXTO)
For Y = 0 To nTEXTO
   For X = 0 To nLEN
       If aTEXTO(Y) = aORI(X) Then ''Encerra Analise Para Evitar
          aTEXTO(Y) = aDES(X)       ''Loop de Troca
          Exit For
       End If
   Next
Next
CharConv = ""
For Y = 0 To nTEXTO
   CharConv = CharConv & aTEXTO(Y)
Next Y
End Function
Public Function TiraOut(ByVal eVAR As Variant) As String
Dim cTexto As String
   cTexto = FixStr(eVAR)
   TiraOut = CharConv(cTexto, Array("-", ",", ".", ":", "/", ";", "*", "(", ")"), _
                              Array("", "", "", "", "", "", "", "", ""))
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


