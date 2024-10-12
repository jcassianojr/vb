Attribute VB_Name = "Module1"
Option Explicit
Public lRETU As Boolean
Public Function tirace2(ByVal cXml As String) As String
'  dim cRemoveTag AS STRING
'  DIM nPos AS INTEGER
  Dim nAscii As Integer
  Dim ltroca As Boolean
  Dim nCont As Integer
  Dim cLETRA As String

  '' cRemoveTag := { ;
  ''    [<?xml version="1.0" encoding="utf-8"?>], ; // Petrobras inventou de usar assim
  ''    [<?xml version="1.0" encoding="ISO-8859-1"?>], ; // Petrobras agora assim
  ''    [<?xml version="1.0" encoding="UTF-8"?>], ; // o mais correto
  ''    [<?xml version="1.00"?>], ;
  ''    [<?xml version="1.0"?>] }''

  ''  cXml := StrTran( cXml, , "" )

  ''  IF ! ["] $ cXml // Pode ser usado aspas simples
  ''     cXml := replace( cXml, ['], ["] )
  ''  ENDIF
  ''  IF Chr(195) $ cXml
  ''     nPos := At( Chr(195), cXml )
  ''     IF Asc( Substr( cXml, nPos + 1 ) ) > 122
  ''        cXml := hb_Utf8ToStr( cXml )
  ''     ENDIF
  ''  ENDIF

  For nCont = 1 To 2
    cXml = Replace(cXml, Chr(26), "")
    cXml = Replace(cXml, Chr(13), "")
    cXml = Replace(cXml, Chr(10), "")
    If InStr(Chr(239) + Chr(187) + Chr(191), Mid(cXml, 1, 1)) Then
      cXml = Mid(cXml, 2)
    End If
    cXml = Replace(cXml, " />", "/>")
    cXml = Replace(cXml, Chr(195) + Chr(173), "i")
    cXml = Replace(cXml, Chr(195) + Chr(135), "C")
    cXml = Replace(cXml, Chr(195) + Chr(141), "I")
    cXml = Replace(cXml, Chr(195) + Chr(163), "a")
    cXml = Replace(cXml, Chr(195) + Chr(167), "c")
    cXml = Replace(cXml, Chr(195) + Chr(161), "a")
    cXml = Replace(cXml, Chr(195) + Chr(131), "A")
    cXml = Replace(cXml, Chr(194) + Chr(186), "o.")
    cXml = Replace(cXml, Chr(195) + Chr(162), "a")
    cXml = Replace(cXml, Chr(195) + Chr(161), "a")
    cXml = Replace(cXml, Chr(195) + Chr(163), "a")
    cXml = Replace(cXml, Chr(195) + Chr(173), "i")
    cXml = Replace(cXml, Chr(195) + Chr(179), "o")
    cXml = Replace(cXml, Chr(195) + Chr(167), "c")
    cXml = Replace(cXml, Chr(195) + Chr(169), "e")
    cXml = Replace(cXml, Chr(195) + Chr(170), "e")
    cXml = Replace(cXml, Chr(195) + Chr(181), "o")
    cXml = Replace(cXml, Chr(195) + Chr(160), "o")
    cXml = Replace(cXml, Chr(195) + Chr(181), "o")
    cXml = Replace(cXml, Chr(195) + Chr(129), "A")
    cXml = Replace(cXml, Chr(226) + Chr(128) + Chr(156), "*")   '// aspas de destaque "cames"
    cXml = Replace(cXml, Chr(226) + Chr(128) + Chr(157), "*")   '// aspas de destaque "cames"
    cXml = Replace(cXml, Chr(195) + Chr(180), "o")
    cXml = Replace(cXml, Chr(195) + Chr(186), "u")
    cXml = Replace(cXml, Chr(195) + Chr(147), "O")
    cXml = Replace(cXml, Chr(226) + Chr(128) + Chr(153), " ")   ' // caixa d'agua
    cXml = Replace(cXml, Chr(226) + Chr(128) + Chr(147), "-")   '// - mesmo
    cXml = Replace(cXml, Chr(194) + Chr(179), "3")   '// m3
    '// so pra corrigir no MySql
    cXml = Replace(cXml, "+" + Chr(129), "A")
    cXml = Replace(cXml, "+" + Chr(137), "E")
    cXml = Replace(cXml, "+" + Chr(131), "A")
    cXml = Replace(cXml, "+" + Chr(135), "C")
    cXml = Replace(cXml, "?" + Chr(167), "c")
    cXml = Replace(cXml, "?" + Chr(163), "a")
    cXml = Replace(cXml, "?" + Chr(173), "i")
    cXml = Replace(cXml, "?" + Chr(131), "A")
    cXml = Replace(cXml, "?" + Chr(161), "a")
    cXml = Replace(cXml, "?" + Chr(141), "I")
    cXml = Replace(cXml, "?" + Chr(135), "C")
    cXml = Replace(cXml, Chr(195) + Chr(156), "a")
    cXml = Replace(cXml, Chr(195) + Chr(159), "A")
    cXml = Replace(cXml, "?" + Chr(129), "A")
    cXml = Replace(cXml, "?" + Chr(137), "E")
    cXml = Replace(cXml, Chr(195) + "?", "C")
    cXml = Replace(cXml, "?" + Chr(149), "O")
    cXml = Replace(cXml, "?" + Chr(154), "U")
    cXml = Replace(cXml, "+" + Chr(170), "o")
    cXml = Replace(cXml, "?" + Chr(128), "A")
    cXml = Replace(cXml, Chr(195) + Chr(166), "e")
    cXml = Replace(cXml, Chr(135) + Chr(227), "ca")
    cXml = Replace(cXml, "n" + Chr(227), "na")
    cXml = Replace(cXml, Chr(162), "o")
    cXml = Replace(cXml, " " + Chr(241) + " ", " ")
    cXml = Replace(cXml, Chr(176), "")   ' graus
    cXml = Replace(cXml, Chr(186), "o")   ' numero
    cXml = Replace(cXml, Chr(220), "U")   ' u com trema
    cXml = Replace(cXml, Chr(170), "")   ' desconhecido
  Next
  For nCont = 1 To Len(cXml)
    cLETRA = Mid(cXml, nCont, 1)
    nAscii = Asc(cLETRA)
    ltroca = True
    If InStr("0123456789", cLETRA) > 0 Then
      ltroca = False
    End If
    If InStr("abcdefghijklmnopqrstuvwxyz", cLETRA) > 0 Then
      ltroca = False
    End If
    If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", cLETRA) > 0 Then
      ltroca = False
    End If
    If InStr(",.:/;%*$@?<>()+-#=:_", cLETRA) > 0 Then
      ltroca = False
    End If
    If nAscii = 231 Then
      cLETRA = "c"
    End If
    If nAscii = 199 Then
      cLETRA = "C"
    End If
    If nAscii = 193 Or nAscii = 194 Or nAscii = 195 Or nAscii = 192 Then
      cLETRA = "A"
    End If
    If nAscii = 224 Or nAscii = 225 Or nAscii = 226 Or nAscii = 227 Or nAscii = 228 Or nAscii = 229 Then
      cLETRA = "a"
    End If
    If nAscii = 242 Or nAscii = 243 Or nAscii = 244 Or nAscii = 245 Or nAscii = 246 Then
      cLETRA = "o"
    End If
    If nAscii = 210 Or nAscii = 211 Or nAscii = 212 Or nAscii = 213 Or nAscii = 214 Then
      cLETRA = "O"
    End If
    If nAscii = 200 Or nAscii = 201 Or nAscii = 202 Or nAscii = 203 Then
      cLETRA = "E"
    End If
    If nAscii = 232 Or nAscii = 233 Or nAscii = 234 Or nAscii = 235 Then
      cLETRA = "e"
    End If
    If nAscii = 236 Or nAscii = 237 Or nAscii = 238 Or nAscii = 239 Then
      cLETRA = "i"
    End If
    If nAscii = 204 Or nAscii = 205 Or nAscii = 206 Or nAscii = 207 Then
      cLETRA = "I"
    End If
    If nAscii = 249 Or nAscii = 250 Or nAscii = 251 Or nAscii = 252 Then
      cLETRA = "u"
    End If
    If nAscii = 217 Or nAscii = 218 Or nAscii = 219 Then
      cLETRA = "U"
    End If
    If nAscii = 128 Then
      cLETRA = "C"
    End If
    If nAscii = 144 Then
      cLETRA = "E"
    End If
    If nAscii = 248 Then
      cLETRA = ""
    End If
    If nAscii = 167 Then
      cLETRA = "o"
    End If
    If ltroca Then
      cXml = Mid(cXml, 1, nCont - 1) + cLETRA + Mid(cXml, nCont + 1)
    End If
  Next
  tirace2 = cXml
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
  If Not IsNull(eVAR) Then
     FixStr = CStr(eVAR)
  End If
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
Public Function TrimNull(ByVal sTxt As String) As String
  Dim Arr() As String
  Arr() = Split(sTxt, Chr$(0))
  If UBound(Arr) >= 0 Then
    TrimNull = Arr(0)
  Else
    TrimNull = sTxt
  End If
End Function
Public Function FileExists(ByVal PathName As String) As Boolean
On Error Resume Next
Dim Attributes As VbFileAttribute, ErrVal As Long
Attributes = GetAttr(PathName)
ErrVal = Err.Number
On Error GoTo 0
If (Attributes And (vbDirectory Or vbVolume)) = 0 And ErrVal = 0 Then FileExists = True
End Function
Public Function TipoDado2(ByVal intType As Integer) As String
  Select Case intType
  Case adSmallInt, adInteger, adSingle, adDouble, adCurrency, adBigInt, adBinary, _
       adDecimal, adNumeric, adTinyInt, adUnsignedBigInt, adUnsignedInt, _
       adUnsignedSmallInt, adUnsignedTinyInt
    TipoDado2 = "N"                          ''E um Numero
  Case adDate, adDBDate, adDBTimeStamp
    TipoDado2 = "D"                          ''Tipos Data
  Case adBoolean
    TipoDado2 = "B"                          ''Tipos Boolean
  Case adVarChar, adLongVarChar, adChar, adLongVarWChar, adVarWChar, adWChar
    TipoDado2 = "S"                          ''Tipos String
  Case adDBTime
    TipoDado2 = "T"                          ''Tipos Time
  Case adBSTR, adChapter, adFileTime, adGUID, adIDispatch, adUserDefined
    TipoDado2 = "X"                          ''Outros Tipos
  Case adEmpty
    TipoDado2 = "U"                          ''uNDEFINIDOS VAZIOS
  Case adError, adIUnknown
    TipoDado2 = "E"                          ''Erros
  Case adLongVarBinary, adVarBinary, adBinary
    TipoDado2 = "O"                          ''Ole
  Case adPropVariant, adVariant   ''Variantes
    TipoDado2 = "V"
  Case Else
    TipoDado2 = "U"
  End Select
End Function
