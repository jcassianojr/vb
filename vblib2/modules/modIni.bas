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
' FUNÃ‡ÃƒO AUXILIAR (SEGURANÃ‡A E VALIDAÃ‡ÃƒO)
' ====================================================================
Private Sub EnsureFolder(ByVal sFilePath As String)
    ' Garante que o diretÃ³rio do INI existe antes de tentar gravar
    Dim sFolder As String
    Dim p As Long
    p = InStrRev(sFilePath, "\")
    If p > 0 Then
        sFolder = Left$(sFilePath, p - 1)
        If Len(Dir$(sFolder, vbDirectory)) = 0 Then
            On Error Resume Next
            MkDir sFolder
            On Error GoTo 0
        End If
    End If
End Sub

' ====================================================================
' FUNÃ‡Ã•ES ORIGINAIS DO modIni
' ====================================================================
Public Function GetIni(Section As String, Item As String, Default As String, IniName As String) As String
    Dim Retour As String
    Dim Longueur As Long
    ' Buffer ampliado para 1024 caracteres para evitar truncamento
    Retour = Space$(1024)
    Longueur = GetPrivateProfileString(Section, Item, Default, Retour, 1024, IniName)
    
    If Longueur > 0 Then
        GetIni = Left$(Retour, Longueur)
    Else
        GetIni = Default
    End If
End Function

Public Sub SetIni(Section As String, Item As String, ByVal Value As String, IniName As String)
    Dim ok As Long
    EnsureFolder IniName ' Valida diretÃ³rio antes de gravar
    ok = WritePrivateProfileString(Section, Item, Value, IniName)
End Sub

' ====================================================================
' FUNÃ‡Ã•ES MOVIDAS DO Myfunctions
' ====================================================================
Public Function PegINIVAL(ByVal cARQINI As String, ByVal cGRUPO As String, ByVal cCAMPO As String, Optional ByVal ePAD As String = "") As String
    Dim z As Long
    Dim sCaminho As String * 1024
    z = GetPrivateProfileString(cGRUPO, cCAMPO, "", sCaminho, 1024, cARQINI)
    
    ' CorreÃ§Ã£o do bug do "0": Agora valida o retorno 'z' em vez do primeiro caractere
    If z > 0 Then
        PegINIVAL = Left$(sCaminho, z)
    Else
        PegINIVAL = ePAD
    End If
End Function

Public Function PegPath(ByVal cGRUPO As String, ByVal cCAMPO As String, Optional ByVal ePAD As String = "", Optional ByVal cARQINI As String = "") As String
    Dim z As Long
    Dim sCaminho As String * 1024
    If Len(Trim$(cARQINI)) = 0 Then
       cARQINI = App.Path & "\" & App.EXEName & ".INI"
    End If
    z = GetPrivateProfileString(cGRUPO, cCAMPO, "", sCaminho, 1024, cARQINI)
    
    If Len(Trim$(ePAD)) = 0 Then
      ePAD = App.Path & "\"
    End If
    
    If z > 0 Then
        PegPath = Left$(sCaminho, z)
    Else
        PegPath = ePAD
    End If
End Function

Public Function PegTable(ByVal cGRUPO As String, ByVal cCAMPO As String)
    Dim z As Long
    Dim sCaminho As String * 1024
    Dim cTMP As String
    z = GetPrivateProfileString(cGRUPO, cCAMPO, "", sCaminho, 1024, App.Path & "\" & App.EXEName & ".INI")
    
    If z > 0 Then
        cTMP = Left$(sCaminho, z)
    Else
        cTMP = App.Path & "\"
    End If
    
    PegTable = PegPath("PATH", cTMP)
End Function

' ====================================================================
' IMPLEMENTAÃ‡Ã•ES EQUIVALENTES DO profile.prg (HARBOUR)
' ====================================================================

Public Function ProfileString(ByVal cINIFile As Variant, ByVal cSection As Variant, ByVal cKey As Variant, Optional ByVal cDefault As Variant = "") As String
    If InStr(cINIFile, ".") = 0 Then
        cINIFile = UCase$(Trim$(cINIFile)) & ".INI"
    End If
    
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
            ' Trim e Str$ para garantir formato independente de localidade (ponto como separador)
            cNewValue = Trim$(Str$(xValue))
        Case vbBoolean
            cNewValue = IIf(xValue, "1", "0")
        Case vbDate
            cNewValue = Format$(xValue, "yyyymmdd")
        Case Else
            cNewValue = ""
    End Select
    
    cSection = Replace(Replace(cSection, "[", ""), "]", "")
    
    If InStr(cINIFile, ".") = 0 Then
        cINIFile = UCase$(Trim$(cINIFile)) & ".INI"
    End If
    
    EnsureFolder CStr(cINIFile) ' Valida diretÃ³rio antes de gravar
    SetProfile = (WritePrivateProfileString(CStr(cSection), CStr(cKey), cNewValue, CStr(cINIFile)) <> 0)
End Function

Public Function ProfileNum(ByVal cFile As Variant, ByVal cSection As Variant, ByVal cKey As Variant, Optional ByVal nDefault As Variant = 0) As Double
    Dim cValue As String
    Dim cDefault As String
    
    cDefault = Trim$(Str$(nDefault))
    cValue = ProfileString(cFile, cSection, cKey, cDefault)
    
    If Len(Trim$(cValue)) > 0 Then
        ProfileNum = Val(cValue)
    Else
        ProfileNum = nDefault
    End If
End Function


Public Function ProfileLogical(ByVal cINIFile As Variant, ByVal cSection As Variant, ByVal cKey As Variant, Optional ByVal lDefault As Variant = False) As Boolean
    Dim cVAL As String
    Dim cDefaultStr As String
    
    If VarType(lDefault) <> vbBoolean Then lDefault = False
    
    cDefaultStr = IIf(lDefault, "1", "0")
    cVAL = ProfileString(cINIFile, cSection, cKey, cDefaultStr)
    
    cVAL = UCase$(Trim$(cVAL))
    Select Case cVAL
        Case "S", "SIM", "T", ".T.", "TRUE", "YES", "ON", "Y", "1"
            ProfileLogical = True
        Case Else
            ProfileLogical = False
    End Select
End Function


' ====================================================================
' FUNÃ‡ÃƒO AUXILIAR DE DATA BASEADA EM UNIVERSALTODATE (HARBOUR)
' ====================================================================
Private Function CheckIniDate(ByVal cData As String) As Date
    Dim dResult As Date
    Dim cClean As String
    Dim aParts() As String
    
    dResult = CDate(0)
    cData = Trim$(cData)
    
    If Len(cData) = 0 Then
        CheckIniDate = dResult
        Exit Function
    End If
    
    cClean = UCase$(cData)
    If cClean = "NULL" Or cClean = "NIL" Or cClean = "<NULL>" Or cClean = "NUL" Or cClean = "/  /" Or cClean = "-  -" Then
        CheckIniDate = dResult
        Exit Function
    End If

    ' Uniformiza separadores
    cClean = Replace(cClean, "-", "/")
    cClean = Replace(cClean, ".", "/")
    
    ' Tenta converter via funÃ§Ã£o nativa do VB (que jÃ¡ trata muitos casos regionais e ISO)
    If IsDate(cClean) Then
        CheckIniDate = CDate(cClean)
        Exit Function
    End If

    ' Tratamento de separadores explÃ­citos (DD/MM/YYYY ou YYYY/MM/DD)
    If InStr(cClean, "/") > 0 Then
        aParts = Split(cClean, "/")
        If UBound(aParts) = 2 Then ' Se tem 3 partes
            If Len(aParts(0)) = 4 Then
                ' YYYY/MM/DD
                dResult = DateSerial(Val(aParts(0)), Val(aParts(1)), Val(aParts(2)))
            Else
                ' DD/MM/YYYY (ou variaÃ§Ãµes curtas suportadas pelo Serial)
                dResult = DateSerial(Val(aParts(2)), Val(aParts(1)), Val(aParts(0)))
            End If
            CheckIniDate = dResult
            Exit Function
        End If
    End If
    
    ' Tratamento da lÃ³gica original "tudo grudado" (AAAAMMDD ou DDMMYYYY)
    cClean = Replace(cClean, "/", "")
    If Len(cClean) = 8 And IsNumeric(cClean) Then
        ' Testa se os 4 primeiros dÃ­gitos sÃ£o um ano plausÃ­vel (ex: > 1900)
        If Val(Left$(cClean, 4)) > 1900 Then
            ' Formato AAAAMMDD
            dResult = DateSerial(Val(Left$(cClean, 4)), Val(Mid$(cClean, 5, 2)), Val(Right$(cClean, 2)))
        Else
            ' Formato DDMMYYYY
            dResult = DateSerial(Val(Right$(cClean, 4)), Val(Mid$(cClean, 3, 2)), Val(Left$(cClean, 2)))
        End If
    ElseIf Len(cClean) = 6 And IsNumeric(cClean) Then
        ' Formato DDMMYY
        ' O VB6 converte anos 2 dÃ­gitos baseado nas configs regionais, mas DateSerial Ã© manual.
        ' Simula a lÃ³gica de EPOCH (geralmente cortes em 30 ou 50)
        Dim nAno As Integer
        nAno = Val(Right$(cClean, 2))
        If nAno < 50 Then
            nAno = 2000 + nAno
        Else
            nAno = 1900 + nAno
        End If
        dResult = DateSerial(nAno, Val(Mid$(cClean, 3, 2)), Val(Left$(cClean, 2)))
    End If
    
    CheckIniDate = dResult
End Function

' ====================================================================
' FUNÃ‡ÃƒO PROFILEDATE ATUALIZADA
' ====================================================================
Public Function ProfileDate(ByVal cFile As Variant, ByVal cSection As Variant, ByVal cKey As Variant, Optional ByVal dDefault As Variant) As Date
    Dim cValue As String
    Dim cDefault As String
    Dim dResult As Date
    
    If IsMissing(dDefault) Or Not IsDate(dDefault) Then
        dDefault = CDate(0)
    End If
    
    ' Converte o default pro formato INI padrÃ£o (AAAAMMDD) para a busca
    cDefault = Trim$(Format$(dDefault, "yyyymmdd"))
    cValue = ProfileString(cFile, cSection, cKey, cDefault)
    
    If Len(Trim$(cValue)) > 0 Then
        ' Usa a nova funÃ§Ã£o baseada em UniversalToDate para interpretar o retorno
        dResult = CheckIniDate(cValue)
        If dResult = CDate(0) And Trim$(cValue) <> cDefault Then
             ProfileDate = dDefault ' Falhou a conversÃ£o, volta pro default real
        Else
             ProfileDate = dResult
        End If
    Else
        ProfileDate = dDefault
    End If
End Function