Attribute VB_Name = "modIni"
Option Explicit

' ====================================================================
' DECLARAÃ‡Ã•ES DA API (32/64-BIT)
' Movidas do Myfunctions para manter a compatibilidade moderna
' ====================================================================
#If VBA7 Or Win64 Then
    Public Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As LongPtr, ByVal lpFileName As String) As Long
    Public Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
    Public Declare PtrSafe Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As LongPtr, ByVal lpFileName As String) As Long
#Else
    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
    Public Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
#End If

' ====================================================================
' FUNÃ‡Ã•ES ORIGINAIS DO modIni
' Mantidas para efeito de compatibilidade com chamadas antigas
' ====================================================================
Public Function GetIni(Section As String, Item As String, Default As String, IniName As String) As String
    Dim Retour As String
    Dim Longueur As Long ' Alterado de Integer para Long para alinhar com o retorno da API
    Retour = Space$(255)
    Longueur = GetPrivateProfileString(Section, Item, Default, Retour, 255, IniName)
    GetIni = Left$(Retour, Longueur)
End Function

Public Sub SetIni(Section As String, Item As String, ByVal Value As String, IniName As String)
    Dim ok As Long
    ok = WritePrivateProfileString(Section, Item, Value, IniName)
End Sub

' ====================================================================
' FUNÃ‡Ã•ES MOVIDAS DO Myfunctions
' Mantendo nomes e parÃ¢metros rigorosamente iguais
' ====================================================================
Public Function PegINIVAL(ByVal cARQINI As String, ByVal cGRUPO As String, ByVal cCAMPO As String, Optional ByVal ePAD As String = "") As String
    Dim z As Long
    Dim sCaminho As String * 255
    z = GetPrivateProfileString(cGRUPO, cCAMPO, "", sCaminho, 150, cARQINI)
    PegINIVAL = IIf(Asc(Left(sCaminho, 1)) = Asc("0"), ePAD, Left(sCaminho, z))
End Function

Public Function PegPath(ByVal cGRUPO As String, ByVal cCAMPO As String, Optional ByVal ePAD As String = "", Optional ByVal cARQINI As String = "") As String
    Dim z As Long
    Dim sCaminho As String * 255
    If cARQINI = "" Then
       cARQINI = App.Path & "\" & App.EXEName & ".INI"
    End If
    z = GetPrivateProfileString(cGRUPO, cCAMPO, "", sCaminho, 150, cARQINI)
    If Len(ePAD) = 0 Then
      ePAD = App.Path & "\"
    End If
    PegPath = IIf(Asc(Left(sCaminho, 1)) = Asc("0"), ePAD, Left(sCaminho, z))
End Function

Public Function PegTable(ByVal cGRUPO As String, ByVal cCAMPO As String)
    Dim z As Long
    Dim sCaminho As String * 255
    Dim cTMP As String
    z = GetPrivateProfileString(cGRUPO, cCAMPO, "", sCaminho, 150, App.Path & "\" & App.EXEName & ".INI")
    cTMP = IIf(Asc(Left(sCaminho, 1)) = Asc("0"), App.Path & "\", Left(sCaminho, z))
    PegTable = PegPath("PATH", cTMP)
End Function

' ====================================================================
' IMPLEMENTAÃ‡Ã•ES EQUIVALENTES DO profile.prg (HARBOUR)
' Traduzidas mantendo a resiliÃªncia e as lÃ³gicas de conversÃ£o
' ====================================================================

Public Function ProfileString(ByVal cINIFile As Variant, ByVal cSection As Variant, ByVal cKey As Variant, Optional ByVal cDefault As Variant = "") As String
    If InStr(cINIFile, ".") = 0 Then
        cINIFile = UCase$(Trim$(cINIFile)) & ".INI"
    End If
    
    ' Garante padronizaÃ§Ã£o sem colchetes e em maiÃºsculo
    cSection = UCase$(Replace(Replace(cSection, "[", ""), "]", ""))
    cKey = UCase$(Trim$(cKey))
    
    ProfileString = GetIni(CStr(cSection), CStr(cKey), CStr(cDefault), CStr(cINIFile))
End Function

Public Function SetProfile(ByVal cINIFile As Variant, ByVal cSection As Variant, ByVal cKey As Variant, ByVal xValue As Variant) As Boolean
    Dim cNewValue As String
    
    Select Case VarType(xValue)
        Case vbString
            cNewValue = xValue
        Case vbInteger, vbLong, vbSingle, vbDouble, vbDecimal, vbCurrency
            cNewValue = Trim$(Str$(xValue))
        Case vbBoolean
            cNewValue = IIf(xValue, "1", "0")
        Case vbDate
            ' Armazena datas no formato AAAAMMDD
            cNewValue = Format$(xValue, "yyyymmdd")
        Case Else
            cNewValue = ""
    End Select
    
    cSection = Replace(Replace(cSection, "[", ""), "]", "")
    
    If InStr(cINIFile, ".") = 0 Then
        cINIFile = UCase$(Trim$(cINIFile)) & ".INI"
    End If
    
    ' Salva utilizando a API nativa
    SetProfile = (WritePrivateProfileString(CStr(cSection), CStr(cKey), cNewValue, CStr(cINIFile)) <> 0)
End Function

Public Function ProfileNum(ByVal cFile As Variant, ByVal cSection As Variant, ByVal cKey As Variant, Optional ByVal nDefault As Variant = 0) As Double
    Dim cValue As String
    Dim cDefault As String
    
    cDefault = Trim$(Str$(nDefault))
    cValue = ProfileString(cFile, cSection, cKey, cDefault)
    
    If Trim$(cValue) <> "" Then
        ProfileNum = Val(cValue)
    Else
        ProfileNum = nDefault
    End If
End Function

Public Function ProfileDate(ByVal cFile As Variant, ByVal cSection As Variant, ByVal cKey As Variant, Optional ByVal dDefault As Variant) As Date
    Dim cValue As String
    Dim cDefault As String
    
    If IsMissing(dDefault) Or Not IsDate(dDefault) Then
        dDefault = CDate(0)
    End If
    
    cDefault = Trim$(Format$(dDefault, "yyyymmdd"))
    cValue = ProfileString(cFile, cSection, cKey, cDefault)
    
    If Trim$(cValue) <> "" And Len(Trim$(cValue)) = 8 And IsNumeric(cValue) Then
        ' Converte YYYYMMDD de volta para Date
        ProfileDate = DateSerial(Val(Left$(cValue, 4)), Val(Mid$(cValue, 5, 2)), Val(Right$(cValue, 2)))
    Else
        ProfileDate = dDefault
    End If
End Function

Public Function ProfileLogical(ByVal cINIFile As Variant, ByVal cSection As Variant, ByVal cKey As Variant, Optional ByVal lDefault As Variant = False) As Boolean
    Dim cVAL As String
    Dim cDefaultStr As String
    
    If VarType(lDefault) <> vbBoolean Then lDefault = False
    
    cDefaultStr = IIf(lDefault, "1", "0")
    cVAL = ProfileString(cINIFile, cSection, cKey, cDefaultStr)
    
    ' Utiliza o nÃºcleo da lÃ³gica contida em StrLogic do Myfunctions
    cVAL = UCase$(Trim$(cVAL))
    Select Case cVAL
        Case "S", "SIM", "T", ".T.", "TRUE", "YES", "ON", "Y", "1"
            ProfileLogical = True
        Case Else
            ProfileLogical = False
    End Select
End Function
