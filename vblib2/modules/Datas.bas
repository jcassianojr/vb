Attribute VB_Name = "Datas"
Option Explicit

Function IsEndOfMonth(DateToCheck As Date) As Boolean
    IsEndOfMonth = (DateToCheck = DateSerial(Year(DateToCheck), Month(DateToCheck) + 1, 0))
End Function

Function IsEndOfYear(DateToCheck As Date) As Boolean
    IsEndOfYear = (Month(DateToCheck) = 12 And Day(DateToCheck) = 31)
End Function

Function IsEndOfQuarter(DateToCheck As Date) As Boolean
    IsEndOfQuarter = (IsEndOfMonth(DateToCheck) And Month(DateToCheck) Mod 3 = 0)
End Function

'Returns 1/1/
Function YearStart(AsOfDate As Date, Optional RelativeYear As Integer = 0) As Date
    YearStart = DateSerial(Year(AsOfDate) + RelativeYear, 1, 1)
End Function

'Returns 12/31/
Function YearEnd(AsOfDate As Date, Optional RelativeYear As Integer = 0) As Date
    YearEnd = DateSerial(Year(AsOfDate) + RelativeYear, 12, 31)
End Function

'Returns 1/1/, 2/1/, 3/1/, 4/1/, 5/1/, 6/1/, 7/1/, 8/1/, 9/1/, 10/1/, 11/1, 12/1
Function MonthStart(AsOfDate As Date, Optional RelativeMonth As Integer = 0) As Date
    MonthStart = DateSerial(Year(AsOfDate), Month(AsOfDate) + RelativeMonth, 1)
End Function

'Returns 1/31/, 2/28/ or 2/29/, 3/31/, 4/30/, 5/31/, 6/30/, 7/31/, 8/31/, 9/30/, 10/31/, 11/30/, 12/31/
Function MonthEnd(AsOfDate As Date, Optional RelativeMonth As Integer = 0) As Date
    MonthEnd = DateSerial(Year(AsOfDate), Month(AsOfDate) + RelativeMonth + 1, 0)
End Function

'Returns 1/1/, 4/1/, 7/1/, 10/1/
Function QuarterStart(AsOfDate As Date, _
                      Optional RelativeQuarter As Integer = 0, _
                      Optional MonthOfQuarter As Integer = 1) As Date
Dim Mo As Integer
    Mo = Int((Month(AsOfDate) + 2 + (3 * RelativeQuarter)) / 3) * 3
    Mo = Mo - 3 + MonthOfQuarter
    QuarterStart = DateSerial(Year(AsOfDate), Mo, 1)
End Function

'Returns 3/31/, 6/30/, 9/30/, 12/31/
Function QuarterEnd(AsOfDate As Date, _
                    Optional RelativeQuarter As Integer = 0, _
                    Optional MonthOfQuarter As Integer = 3) As Date
Dim Mo As Integer
    Mo = Int((Month(AsOfDate) + 2 + (3 * RelativeQuarter)) / 3) * 3
    Mo = Mo - 3 + MonthOfQuarter + 1
    QuarterEnd = DateSerial(Year(AsOfDate), Mo, 0)
End Function

Public Function DataBranco(ByVal eVAR As Variant) As Boolean
  DataBranco = False
  If IsNull(eVAR) Then
    DataBranco = True
    Exit Function
  End If
  If eVAR = "" Or eVAR = "  /  /    " Or eVAR = "00/00/0000" Or eVAR = Space(8) _
     Or eVAR = "00:00:00" Or eVAR = "0000-00-00" _
     Or eVAR = "0000-00-00 00:00:00" Then
    DataBranco = True
    Exit Function
  End If
  If IsDate(eVAR) Then                         ''Corrige Erro data null voltando 31/12/1899
    If Year(Date) <= 1899 Then
      DataBranco = True
      Exit Function
    End If
  End If
End Function
Public Function Fdata(ByVal Data As Variant, _
                      Optional ByVal cTipoData = "", _
                      Optional ByVal ePAD As Variant, _
                      Optional ByVal cTipoDataNull As String = "", _
                      Optional ByVal cMASCARA As String = "dd/mm/yyyy") As Variant
  Dim dDATA As Date

  '         "\#mm\/dd\/yyyy#"
  '         "\#hh:mm:ss#"
  '         "\#mm\/dd\/yyyy hh:mm:ss#"
  'data padrao
  Fdata = ePAD

  If cTipoData = "D-" Then
     cMASCARA = "yyyy-mm-dd"
  End If
  If IsDate(Data) And Not DataBranco(Data) Then  ''Data <> "00:00:00" And Data <> "0000-00-00" Then
    Fdata = Format(Data, cMASCARA)
  Else
    Select Case cTipoData 'gera null conforme o tipo
    Case "", "D", "DS"
      Fdata = DateSerial(0, 0, 0)
    Case "DD"
      Fdata = dDATA
    Case "DN"
      Fdata = NullDate(cTipoDataNull)
    Case "DC"
      Fdata = Space(8)
    Case "DZ"
      Fdata = ""
    Case "DF"
      Fdata = "  /  /    "
    Case "DH"
      Fdata = Today()
    End Select
  End If
End Function
Public Function MesAnt(ByVal nMES As Integer, ByVal nANO As Long) As Variant
  Dim nMESTMP As Integer
  Dim nANOTMP As Long
  nMESTMP = nMES - 1
  nANOTMP = nANO
  If nMESTMP = 0 Then
    nMESTMP = 12
    nANOTMP = nANOTMP - 1
  End If
  MesAnt = Array(nMESTMP, nANOTMP)
End Function

Public Function NullDate(Optional ByVal cTIPO As String = "", Optional ByVal cARQ As String = "") As Variant
Dim aRETU As Variant
  If Len(cARQ) > 0 Then
    cARQ = GeraConn(cARQ, cTIPO)
    aRETU = TipoConn(cARQ)
    cTIPO = aRETU(2)
  End If
  Select Case cTIPO
  Case "DBF", "ADSCDX", "ADSNTX", "ADSADT", "JETFOX"
    NullDate = "        "
  Case "MYSQL", "MARIADB"
    NullDate = "'0000-00-00'"
  Case Else
    NullDate = Null
  End Select
End Function

Public Function NullDateTime(Optional ByVal cTIPO As String = "", Optional ByVal cARQ As String = "") As Variant
 Dim aRETU As Variant
  If Len(cARQ) > 0 Then
    cARQ = GeraConn(cARQ, cTIPO)
    aRETU = TipoConn(cARQ)
    cTIPO = aRETU(2)
  End If
  Select Case cTIPO
  Case "DBF", "ADSCDX", "ADSNTX", "ADSADT", "JETFOX"
    NullDateTime = "        "
  Case "MYSQL", "MARIADB"
    NullDateTime = "'0000-00-00 00:00:00'"
  Case Else
    NullDateTime = Null
  End Select
End Function
Public Function Today() As Date
  Today = Format(Date, "dd/mm/yyyy")
End Function
Public Function NumToData(ByVal nNUM As Variant) As Date
  Dim dDATA As Date
  Dim cData As String
  cData = Trim(TiraOut(nNUM))
  If Len(cData) = 6 Then                       ''yymmdd
    dDATA = DateSerial(Mid(cData, 1, 2), Mid(cData, 3, 2), Mid(cData, 5, 2))
  End If
  If Len(cData) - 8 Then                       ''8 ''yyyymmdd
    dDATA = DateSerial(Mid(cData, 1, 4), Mid(cData, 5, 2), Mid(cData, 7, 2))
  End If
  NumToData = Fdata(dDATA)
End Function
Public Function DataToLit(ByVal dDATA As Variant, Optional ByVal cTIPO As String = "", Optional ByVal cARQ As String = "") As String
  Dim aRETU As Variant
  Dim cSEP As String
  On Error GoTo trataerro
  cSEP = ""
  dDATA = Fdata(dDATA, "DN")
  Select Case Mid(cTIPO, 1, 1)
  Case "/"
    cSEP = "/"
    cTIPO = Mid(cTIPO, 2)
  Case "-"
    cSEP = "-"
    cTIPO = Mid(cTIPO, 2)
  End Select
  If Len(cARQ) > 0 Then
    cARQ = GeraConn(cARQ, cTIPO)
    aRETU = TipoConn(cARQ)
    cTIPO = aRETU(2)
  End If
  Select Case cTIPO
  Case "ACESSS"                                ''access
    DataToLit = "#" & Year(dDATA) & "/" & Month(dDATA) & "/" & Day(dDATA) & "#"
  Case "CRYSTAL"
    DataToLit = "CDATE(" & Year(dDATA) & "," & Month(dDATA) & "," & Day(dDATA) & ")"
  Case "MYSQL/"
    DataToLit = "'" & Year(dDATA) & "/" & Month(dDATA) & "/" & Day(dDATA) & "'"
  Case "MYSQL", "MYSQL-", "MARIADB"
    DataToLit = "'" & Year(dDATA) & "-" & Month(dDATA) & "-" & Day(dDATA) & "'"
  Case "SQLSERVER"
    DataToLit = "CONVERT(datetime, '" & Format(DateValue(dDATA), "yyyy-mm-dd") & "', 102)"
  Case "ORACLE"
    DataToLit = "to_date('" + Format(dDATA, "dd/mm/yyyy") + "','DD/MM/YYYY')"
  Case "CRYSTAX"
    DataToLit = "DATE(" & Year(dDATA) & "," & Month(dDATA) & "," & Day(dDATA) & ")"
  Case "NDL"                                   ''yyyymmdd
    DataToLit = StrZero(Year(dDATA), 4) & cSEP & StrZero(Month(dDATA), 2) & cSEP & StrZero(Day(dDATA), 2)
  Case "NDC"                                   ''yymmdd
    DataToLit = Mid(StrZero(Year(dDATA), 4), 3) & cSEP & StrZero(Month(dDATA), 2) & cSEP & StrZero(Day(dDATA), 2)
  Case Else
    DataToLit = "'" & Year(dDATA) & "/" & Month(dDATA) & "/" & Day(dDATA) & "'"
  End Select
  Exit Function
trataerro:
  Select Case Err.Number
  Case Else
    SayErro "DatatoLit"
  End Select
End Function

