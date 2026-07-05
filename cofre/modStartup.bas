Attribute VB_Name = "modstartup"
Option Explicit
#If VBA7 Or Win64 Then
    ' --- BLOCO 64-BIT / TWINBASIC / VBA7 ---
    Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
    Public Declare PtrSafe Function ShellExecuteForExplore Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, lpParameters As Any, lpDirectory As Any, ByVal nShowCmd As LongPtr) As LongPtr
    
    Public Declare PtrSafe Function WinAPI_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As LongPtr) As Long
    Public Declare PtrSafe Function InternetGetConnectedState Lib "wininet" (ByRef dwFlags As LongPtr, ByVal dwReserved As LongPtr) As Long
    Public Declare PtrSafe Function EbExecuteLine Lib "vba6.dll" (ByVal pStringToExec As LongPtr, ByVal Unknownn1 As LongPtr, ByVal Unknownn2 As LongPtr, ByVal fCheckOnly As LongPtr) As Long
    Public Declare PtrSafe Function ReleaseCapture Lib "user32" () As Long
    
    Public Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As LongPtr, ByVal lpFileName As String) As Long
    Public Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
    
    Public Declare PtrSafe Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As LongPtr, ByVal LCType As LongPtr, ByVal lpLCData As String, ByVal cchData As LongPtr) As Long
    Public Declare PtrSafe Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As LongPtr, ByVal LCType As LongPtr, ByVal lpLCData As String) As Long
    
    Public Declare PtrSafe Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
    Public Declare PtrSafe Function OemToChar Lib "user32" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
    Public Declare PtrSafe Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As LongPtr, ByVal lpHelpFile As String, ByVal wCommand As LongPtr, ByVal dwData As LongPtr) As Long
   Public Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" ( _
        ByVal CodePage As Long, _
        ByVal dwFlags As Long, _
        ByVal lpMultiByteStr As LongPtr, _
        ByVal cbMultiByte As Long, _
        ByVal lpWideCharStr As LongPtr, _
        ByVal cchWideChar As Long) As Long

#Else
    ' --- BLOCO 32-BIT CLÁSSICO (VB6) ---
    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Public Declare Function ShellExecuteForExplore Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, lpParameters As Any, lpDirectory As Any, ByVal nShowCmd As Long) As Long
    
    Public Declare Function WinAPI_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Public Declare Function InternetGetConnectedState Lib "wininet" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
    Public Declare Function EbExecuteLine Lib "vba6.dll" (ByVal pStringToExec As Long, ByVal Unknownn1 As Long, ByVal Unknownn2 As Long, ByVal fCheckOnly As Long) As Long
    Public Declare Function ReleaseCapture Lib "user32" () As Long
    
    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
    
    Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
    Public Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Long
    
    Public Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
    Public Declare Function OemToChar Lib "user32" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
    Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
    Public Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

#End If
Public Sub Main()
    Dim Resposta As VbMsgBoxResult
    
    ' Pergunta ao usuário qual cofre ele deseja abrir
    Resposta = MsgBox("Qual cofre você deseja gerenciar?" & vbCrLf & vbCrLf & _
                      "[ Sim ] - Cofre de Bancos de Dados" & vbCrLf & _
                      "[ Não ] - Cofre de Contas de E-mail", _
                      vbQuestion + vbYesNoCancel + vbDefaultButton1, "Selecionar Módulo")
                      
    Select Case Resposta
        Case vbYes
            ' Carrega e exibe o formulário de Bancos de Dados
            Frmbancos.Show
            
        Case vbNo
            ' Carrega e exibe o formulário de E-mails
            frmEmail.Show
            
        Case vbCancel
            ' Encerra a aplicação sem abrir nenhum formulário
            End
    End Select
End Sub
Public Function StrZero(ByVal nNUM, Optional ByVal nLen As Integer = 0)
  Dim cTemp As String
  If nLen = 0 Then
    cTemp = FixStr(nNUM, "0", "TRIM")
    nLen = Trim(cTemp)
  End If
  StrZero = Right(String(nLen, "0") & CStr(nNUM), nLen)
End Function
Public Function FixStr(ByVal eVAR As Variant, _
                       Optional ByVal ePAD As Variant = "", _
                       Optional ByVal coper As String = "", _
                       Optional ByVal nLen As Integer = 0) As Variant
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
  If nLen > 0 And Len(eVAR) > nLen Then
    eVAR = Mid(eVAR, 1, nLen)
    FixStr = eVAR
  End If
  Exit Function
errhandler:
  FixStr = ""
  Resume Next
End Function
Public Function PegPath(ByVal cGRUPO As String, ByVal cCAMPO As String, Optional ByVal ePAD As String = "" _
                       , Optional ByVal cARQINI As String = "") As String
  Dim z As Long
  Dim sCaminho As String * 255
  If cARQINI = "" Then
     cARQINI = App.Path & "\" & App.EXEName & ".INI"
  End If
  z = GetPrivateProfileString(cGRUPO, cCAMPO, "", sCaminho, 150, cARQINI)
  If Len(ePAD) = 0 Then
    ePAD = App.Path & "\"
  End If
  PegPath = IIf(Asc(Left(sCaminho, 1)) = "0", ePAD, Left(sCaminho, z))

End Function

Public Function MDG(ByVal cMEnSSAGEM As String, Optional cTITULO = "Confirme")
  Dim eRESP As Variant
  MDG = False
  eRESP = MsgBox(cMEnSSAGEM, vbYesNo + vbDefaultButton1, cTITULO)
  If eRESP = vbYes Then
    MDG = True
  End If
End Function

