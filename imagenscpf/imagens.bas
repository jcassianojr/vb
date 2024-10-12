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
  Set OBJCONN = New ADODB.Connection  'Create a new object
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
      If lFileLength > 1 Then  ' Ja carregado
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
Public Function FileExists(ByVal PathName As String) As Boolean

On Error Resume Next
Dim Attributes As VbFileAttribute, ErrVal As Long
Attributes = GetAttr(PathName)
ErrVal = Err.Number
On Error GoTo 0
If (Attributes And (vbDirectory Or vbVolume)) = 0 And ErrVal = 0 Then FileExists = True
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


