Attribute VB_Name = "Datas"
Option Explicit

' +--------------------------------------------------------------------
' + VariÃ¡veis Globais Privadas para Cache de DicionÃ¡rios (Performance)
' + Absorvido da lÃ³gica do stdDate.cls
' +--------------------------------------------------------------------
Private p_dictMeses As Object
Private p_dictDias As Object
Private p_dictMesesStr As Object

' +--------------------------------------------------------------------
' + InicializaÃ§Ã£o PreguiÃ§osa (Lazy Load) dos DicionÃ¡rios
' +--------------------------------------------------------------------
Private Sub InitDicts()
    If p_dictMeses Is Nothing Then
        Set p_dictMeses = CreateObject("Scripting.Dictionary")
        p_dictMeses.Add 1, "Janeiro":   p_dictMeses.Add 2, "Fevereiro"
        p_dictMeses.Add 3, "MarÃ§o":     p_dictMeses.Add 4, "Abril"
        p_dictMeses.Add 5, "Maio":      p_dictMeses.Add 6, "Junho"
        p_dictMeses.Add 7, "Julho":     p_dictMeses.Add 8, "Agosto"
        p_dictMeses.Add 9, "Setembro":  p_dictMeses.Add 10, "Outubro"
        p_dictMeses.Add 11, "Novembro": p_dictMeses.Add 12, "Dezembro"
        
        Set p_dictDias = CreateObject("Scripting.Dictionary")
        p_dictDias.Add 1, "Domingo":      p_dictDias.Add 2, "Segunda-Feira"
        p_dictDias.Add 3, "TerÃ§a-Feira":  p_dictDias.Add 4, "Quarta-Feira"
        p_dictDias.Add 5, "Quinta-Feira": p_dictDias.Add 6, "Sexta-Feira"
        p_dictDias.Add 7, "SÃ¡bado"
        
        Set p_dictMesesStr = CreateObject("Scripting.Dictionary")
        ' Cache RÃ¡pido para Mapeamento de Meses em PT e EN (UniversalToDate)
        p_dictMesesStr.Add "JAN", 1: p_dictMesesStr.Add "FEV", 2: p_dictMesesStr.Add "FEB", 2
        p_dictMesesStr.Add "MAR", 3: p_dictMesesStr.Add "ABR", 4: p_dictMesesStr.Add "APR", 4
        p_dictMesesStr.Add "MAI", 5: p_dictMesesStr.Add "MAY", 5: p_dictMesesStr.Add "JUN", 6
        p_dictMesesStr.Add "JUL", 7: p_dictMesesStr.Add "AGO", 8: p_dictMesesStr.Add "AUG", 8
        p_dictMesesStr.Add "SET", 9: p_dictMesesStr.Add "SEP", 9: p_dictMesesStr.Add "OUT", 10
        p_dictMesesStr.Add "OCT", 10: p_dictMesesStr.Add "NOV", 11: p_dictMesesStr.Add "DEZ", 12
        p_dictMesesStr.Add "DEC", 12
    End If
End Sub

Function IsEndOfMonth(DateToCheck As Date) As Boolean
    IsEndOfMonth = (DateToCheck = DateSerial(Year(DateToCheck), Month(DateToCheck) + 1, 0))
End Function

Function IsEndOfYear(DateToCheck As Date) As Boolean
    IsEndOfYear = (Month(DateToCheck) = 12 And Day(DateToCheck) = 31)
End Function

Function IsEndOfQuarter(DateToCheck As Date) As Boolean
    IsEndOfQuarter = (IsEndOfMonth(DateToCheck) And Month(DateToCheck) Mod 3 = 0)
End Function

Function YearStart(AsOfDate As Date, Optional RelativeYear As Integer = 0) As Date
    YearStart = DateSerial(Year(AsOfDate) + RelativeYear, 1, 1)
End Function

Function YearEnd(AsOfDate As Date, Optional RelativeYear As Integer = 0) As Date
    YearEnd = DateSerial(Year(AsOfDate) + RelativeYear, 12, 31)
End Function

Function MonthStart(AsOfDate As Date, Optional RelativeMonth As Integer = 0) As Date
    MonthStart = DateSerial(Year(AsOfDate), Month(AsOfDate) + RelativeMonth, 1)
End Function

Function MonthEnd(AsOfDate As Date, Optional RelativeMonth As Integer = 0) As Date
    MonthEnd = DateSerial(Year(AsOfDate), Month(AsOfDate) + RelativeMonth + 1, 0)
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
  If IsDate(eVAR) Then
    If Year(eVAR) <= 1899 Then
      DataBranco = True
      Exit Function
    End If
  End If
End Function

Public Function Fdata(ByVal Data As Variant, _
                      Optional ByVal cTipoData As String = "", _
                      Optional ByVal ePAD As Variant, _
                      Optional ByVal cTipoDataNull As String = "", _
                      Optional ByVal cMASCARA As String = "dd/mm/yyyy") As Variant
  Dim dDATA As Date
  Fdata = ePAD
  If cTipoData = "D-" Then cMASCARA = "yyyy-mm-dd"
  
  If IsDate(Data) And Not DataBranco(Data) Then
    Fdata = Format(Data, cMASCARA)
  Else
    Select Case cTipoData
    Case "", "D", "DS": Fdata = DateSerial(0, 0, 0)
    Case "DD": Fdata = dDATA
    Case "DN": Fdata = NullDate(cTipoDataNull)
    Case "DC": Fdata = Space(8)
    Case "DZ": Fdata = ""
    Case "DF": Fdata = "  /  /    "
    Case "DH": Fdata = Today()
    End Select
  End If
End Function

Public Function MesAnt(ByVal nMES As Integer, ByVal nANO As Long) As Variant
  Dim nMESTMP As Integer, nANOTMP As Long
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
  Case "DBF", "ADSCDX", "ADSNTX", "ADSADT", "JETFOX": NullDate = "        "
  Case "MYSQL", "MARIADB": NullDate = "'0000-00-00'"
  Case Else: NullDate = Null
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
  Case "DBF", "ADSCDX", "ADSNTX", "ADSADT", "JETFOX": NullDateTime = "        "
  Case "MYSQL", "MARIADB": NullDateTime = "'0000-00-00 00:00:00'"
  Case Else: NullDateTime = Null
  End Select
End Function

Public Function Today() As Date
  Today = Format(Date, "dd/mm/yyyy")
End Function

Public Function NumToData(ByVal nNUM As Variant) As Date
  Dim dDATA As Date, cData As String
  cData = Trim(TiraOut(nNUM))
  If Len(cData) = 6 Then dDATA = DateSerial(Mid(cData, 1, 2), Mid(cData, 3, 2), Mid(cData, 5, 2))
  If Len(cData) = 8 Then dDATA = DateSerial(Mid(cData, 1, 4), Mid(cData, 5, 2), Mid(cData, 7, 2))
  NumToData = Fdata(dDATA)
End Function

' +--------------------------------------------------------------------
' +  FunÃ§Ã£o: UniversalToDate
' +  Aprimorada: SubstituiÃ§Ã£o de Arrays por DicionÃ¡rio (Maior velocidade)
' +--------------------------------------------------------------------
Public Function UniversalToDate(ByVal xData As Variant) As Date
    Dim cData As String, cTemp As String, cLimpa As String
    Dim aParts() As String, i As Integer
    Dim sAno As String, sMes As String, sDia As String, cMesStr As String
    Dim nMes As Integer, nAno As Integer
    
    If VarType(xData) = vbDate Then
        UniversalToDate = CDate(xData)
        Exit Function
    End If
    
    If IsNull(xData) Or IsEmpty(xData) Then
        UniversalToDate = CDate(0)
        Exit Function
    End If
    
    cData = UCase$(Trim$(CStr(xData)))
    If cData = "" Or cData = "NULL" Or cData = "NIL" Or cData = "<NULL>" Or cData = "NUL" Or _
       cData = "/  /" Or cData = "-  -" Or cData = "01/01/1900" Or cData = "0000-00-00" Or Left$(cData, 4) = "0000" Then
        UniversalToDate = CDate(0)
        Exit Function
    End If
    
    cTemp = Replace(Replace(cData, ",", " "), "-", " ")
    Do While InStr(cTemp, "  ") > 0: cTemp = Replace(cTemp, "  ", " "): Loop
    
    aParts = Split(cTemp, " ")
    If UBound(aParts) >= 3 Then
        InitDicts ' Carrega cache de traduÃ§Ã£o
        For i = 0 To UBound(aParts)
            If Len(aParts(i)) >= 3 Then
                cMesStr = Left$(aParts(i), 3)
                nMes = 0
                ' Substitui Loop de Arrays por busca instantÃ¢nea via DicionÃ¡rio
                If p_dictMesesStr.Exists(cMesStr) Then nMes = p_dictMesesStr.Item(cMesStr)
                
                If nMes > 0 Then
                    sMes = Right$("0" & CStr(nMes), 2)
                    If i = 1 And UBound(aParts) >= 4 Then
                        sDia = Right$("0" & aParts(2), 2): sAno = aParts(4)
                    ElseIf i = 2 Then
                        sDia = Right$("0" & aParts(1), 2): sAno = aParts(3)
                        If Len(sAno) = 2 Then sAno = IIf(Val(sAno) < 50, "20" & sAno, "19" & sAno)
                    Else
                        GoTo ProximoBloco
                    End If
                    
                    If Val(sDia) >= 1 And Val(sDia) <= 31 And Val(sAno) >= 1000 And Len(sAno) = 4 Then
                        UniversalToDate = DateSerial(Val(sAno), Val(sMes), Val(sDia))
                        Exit Function
                    End If
                End If
            End If
ProximoBloco:
        Next i
    End If
    
    cTemp = cData
    If InStr(cTemp, " ") > 0 Then cTemp = Trim$(Split(cTemp, " ")(0))
    cTemp = Replace(Replace(cTemp, "-", "/"), ".", "/")
    aParts = Split(cTemp, "/")
    
    If UBound(aParts) = 2 Then
        If Len(aParts(0)) = 4 Then
            sAno = aParts(0): sMes = Right$("0" & aParts(1), 2): sDia = Right$("0" & aParts(2), 2)
        Else
            sDia = Right$("0" & aParts(0), 2): sMes = Right$("0" & aParts(1), 2): sAno = aParts(2)
            If Len(sAno) = 2 Then sAno = IIf(Val(sAno) < 50, "20" & sAno, "19" & sAno)
        End If
        If (sAno & sMes & sDia) = "00000000" Then
            UniversalToDate = CDate(0)
        Else
            UniversalToDate = DateSerial(Val(sAno), Val(sMes), Val(sDia))
        End If
        Exit Function
    End If
    
    cLimpa = Trim$(Replace(Replace(Replace(cData, "/", ""), "-", ""), ".", ""))
    If Len(cLimpa) = 8 And IsNumeric(cLimpa) Then
        If Val(Left$(cLimpa, 4)) > 1900 Then
            UniversalToDate = DateSerial(Val(Left$(cLimpa, 4)), Val(Mid$(cLimpa, 5, 2)), Val(Right$(cLimpa, 2)))
        Else
            UniversalToDate = DateSerial(Val(Right$(cLimpa, 4)), Val(Mid$(cLimpa, 3, 2)), Val(Left$(cLimpa, 2)))
        End If
    ElseIf Len(cLimpa) = 6 And IsNumeric(cLimpa) Then
        nAno = Val(Right$(cLimpa, 2))
        nAno = IIf(nAno < 50, 2000 + nAno, 1900 + nAno)
        UniversalToDate = DateSerial(nAno, Val(Mid$(cLimpa, 3, 2)), Val(Left$(cLimpa, 2)))
    Else
        If IsDate(cData) Then UniversalToDate = CDate(cData) Else UniversalToDate = CDate(0)
    End If
End Function

Public Function C_Data(ByVal mDATA As String) As Date
    C_Data = UniversalToDate(mDATA)
End Function

Public Function DigaData(ByVal mDATA As Variant) As String
    Dim dReal As Date
    dReal = UniversalToDate(mDATA)
    If dReal = CDate(0) Then DigaData = "": Exit Function
    DigaData = Trim$(Str$(Day(dReal))) & " de " & cMes(dReal) & " de " & Trim$(Str$(Year(dReal)))
End Function

Public Function cMes(ByVal mDATA As Variant) As String
    Dim dReal As Date
    dReal = UniversalToDate(mDATA)
    If dReal = CDate(0) Then cMes = "": Exit Function
    cMes = mMES(Month(dReal))
End Function

Public Function cDia(ByVal mDATA As Variant) As String
    Dim dReal As Date
    dReal = UniversalToDate(mDATA)
    If dReal = CDate(0) Then cDia = "": Exit Function
    cDia = DDia(Weekday(dReal))
End Function

' +--------------------------------------------------------------------
' +  FunÃ§Ã£o: mMES (Aprimorada via Dictionary ao invÃ©s de Select Case)
' +--------------------------------------------------------------------
Public Function mMES(ByVal nMES As Integer) As String
    InitDicts
    If p_dictMeses.Exists(nMES) Then
        mMES = p_dictMeses.Item(nMES)
    Else
        mMES = ""
    End If
End Function

' +--------------------------------------------------------------------
' +  FunÃ§Ã£o: DDia (Aprimorada via Dictionary ao invÃ©s de Select Case)
' +--------------------------------------------------------------------
Public Function DDia(ByVal nDIA As Integer) As String
    InitDicts
    If p_dictDias.Exists(nDIA) Then
        DDia = p_dictDias.Item(nDIA)
    Else
        DDia = ""
    End If
End Function

Public Function DataS(ByVal dDATA As Variant) As String
    Dim dReal As Date
    dReal = UniversalToDate(dDATA)
    If dReal = CDate(0) Then DataS = "": Exit Function
    DataS = Format(dReal, "dd/mm/yyyy")
End Function

Public Function DataS2() As String
    DataS2 = Format(Date, "dd/mm/yyyy")
End Function

Public Function DataToLit(ByVal dDATA As Variant, Optional ByVal cTIPO As String = "", Optional ByVal cARQ As String = "") As String
    Dim dReal As Date, aRETU As Variant
    Dim cAno As String, cMes As String, cDia As String
    
    dReal = UniversalToDate(dDATA)
    If dReal = CDate(0) Then DataToLit = "NULL": Exit Function
    
    If Len(cARQ) > 0 Then
        cARQ = GeraConn(cARQ, cTIPO)
        aRETU = TipoConn(cARQ)
        cTIPO = aRETU(2)
    End If
    
    cTIPO = UCase$(Trim$(cTIPO))
    cAno = Trim$(Str$(Year(dReal)))
    cMes = Right$("0" & Trim$(Str$(Month(dReal))), 2)
    cDia = Right$("0" & Trim$(Str$(Day(dReal))), 2)
    
    Select Case cTIPO
        Case "ACCESS", "ACESS": DataToLit = "#" & cAno & "/" & Month(dReal) & "/" & Day(dReal) & "#"
        Case "CRYSTAL": DataToLit = "CDATE(" & cAno & "," & Month(dReal) & "," & Day(dReal) & ")"
        Case "MYSQL/": DataToLit = "'" & cAno & "/" & cMes & "/" & cDia & "'"
        Case "MYSQL", "MYSQL-", "MARIADB": DataToLit = "'" & cAno & "-" & cMes & "-" & cDia & "'"
        Case "SQLSERVER": DataToLit = "CONVERT(datetime, '" & cAno & "-" & cMes & "-" & cDia & "', 102)"
        Case "ORACLE": DataToLit = "to_date('" & cDia & "/" & cMes & "/" & cAno & "','DD/MM/YYYY')"
        Case "POSTGRESQL", "POSTGRES": DataToLit = "DATE '" & cAno & "-" & cMes & "-" & cDia & "'"
        Case "SQLITE", "ADVANTAGE": DataToLit = "'" & cAno & "-" & cMes & "-" & cDia & "'"
        Case "CRYSTAX": DataToLit = "DATE(" & cAno & "," & Month(dReal) & "," & Day(dReal) & ")"
        Case Else: DataToLit = "'" & cAno & "-" & cMes & "-" & cDia & "'"
    End Select
End Function

Function QuarterStart(ByVal AsOfDate As Variant, Optional RelativeQuarter As Integer = 0) As Date
    Dim dReal As Date, NewMonth As Integer
    dReal = UniversalToDate(AsOfDate)
    If dReal = CDate(0) Then dReal = Date
    NewMonth = (((Month(dReal) - 1) \ 3) * 3) + 1 + (RelativeQuarter * 3)
    QuarterStart = DateSerial(Year(dReal), NewMonth, 1)
End Function

Function QuarterEnd(ByVal AsOfDate As Variant, Optional RelativeQuarter As Integer = 0) As Date
    Dim dReal As Date, NewMonth As Integer
    dReal = UniversalToDate(AsOfDate)
    If dReal = CDate(0) Then dReal = Date
    NewMonth = (((Month(dReal) - 1) \ 3) * 3) + 4 + (RelativeQuarter * 3)
    QuarterEnd = DateSerial(Year(dReal), NewMonth, 0)
End Function

Public Function SomaMes(ByVal mDATA As Variant, ByVal mMES As Integer) As Date
    Dim dReal As Date
    dReal = UniversalToDate(mDATA)
    If dReal = CDate(0) Then dReal = Date
    SomaMes = DateAdd("m", mMES, dReal)
End Function

Public Function SomaAno(ByVal mDATA As Variant, ByVal mANO As Integer) As Date
    Dim dReal As Date
    dReal = UniversalToDate(mDATA)
    If dReal = CDate(0) Then dReal = Date
    SomaAno = DateAdd("yyyy", mANO, dReal)
End Function

Public Function Dia_Mes(ByVal nMES As Integer, ByVal nANO As Integer) As Integer
    If nMES >= 1 And nMES <= 12 And nANO > 0 Then
        Dia_Mes = Day(DateSerial(nANO, nMES + 1, 0))
    Else
        Dia_Mes = 0
    End If
End Function

' +--------------------------------------------------------------------
' +  FunÃ§Ã£o: Bissexto
' +  Aprimorada: CÃ¡lculo matemÃ¡tico exato com correÃ§Ã£o baseada no stdDate
' +--------------------------------------------------------------------
Public Function Bissexto(ByVal nANO As Integer) As Boolean
    If nANO <= 0 Then
        Bissexto = False
        Exit Function
    End If
    ' Tratamento para manter retrocompatibilidade com o comportamento do VBA/Excel
    If nANO = 1900 Then 
        Bissexto = True
        Exit Function
    End If
    ' LÃ³gica matemÃ¡tica correta (MÃºltiplo de 4, nÃ£o mÃºltiplo de 100 exceto mÃºltiplos de 400)
    Bissexto = ((nANO Mod 4 = 0 And nANO Mod 100 <> 0) Or nANO Mod 400 = 0)
End Function

Public Function ValiData(ByVal mDATA As Variant) As Boolean
    Dim dReal As Date
    If IsNull(mDATA) Or IsEmpty(mDATA) Then
        ValiData = False
        Exit Function
    End If
    dReal = UniversalToDate(mDATA)
    If dReal = CDate(0) Then ValiData = False Else ValiData = True
End Function

Public Function Inverte(ByVal mDATA As Variant) As String
    Dim dReal As Date, cAno As String, cMes As String, cDia As String
    dReal = UniversalToDate(mDATA)
    If dReal = CDate(0) Then Inverte = "": Exit Function
    cAno = Trim$(Str$(Year(dReal)))
    cMes = Right$("0" & Trim$(Str$(Month(dReal))), 2)
    cDia = Right$("0" & Trim$(Str$(Day(dReal))), 2)
    Inverte = cAno & "-" & cMes & "-" & cDia
End Function

Public Function DataAno(ByVal mDATA As Variant, ByVal mNUME As Integer) As String
    Dim dReal As Date, cAnoCurto As String
    dReal = UniversalToDate(mDATA)
    If dReal = CDate(0) Then dReal = Date
    cAnoCurto = Right$("0" & Trim$(Str$(Year(dReal))), 2)
    DataAno = cAnoCurto & Right$("0" & Trim$(Str$(mNUME)), 2)
End Function

Public Function ValiDataE(ByVal mDATA As Variant, Optional ByVal lExibeMensagem As Boolean = False) As Boolean
    Dim dReal As Date
    dReal = UniversalToDate(mDATA)
    If dReal = CDate(0) Then
        If lExibeMensagem Then MsgBox "Data invÃ¡lida ou incorreta!", vbCritical, "Aviso do Sistema"
        ValiDataE = False
    Else
        ValiDataE = True
    End If
End Function

Public Function D_Extenso(ByVal mDATA As Variant) As String
    Dim dReal As Date
    dReal = UniversalToDate(mDATA)
    If dReal = CDate(0) Then D_Extenso = "": Exit Function
    D_Extenso = cDia(dReal) & ", " & Trim$(Str$(Day(dReal))) & " de " & cMes(dReal) & " de " & Trim$(Str$(Year(dReal)))
End Function