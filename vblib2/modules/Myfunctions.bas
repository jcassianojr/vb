Attribute VB_Name = "Myfunctions"
Option Explicit
'FileConnExist(cARQ)Verifica a existencia de Uma Arquivo
'pegpath(Cgrupo,Ccamp) 'Traz caminho de uma arquivo ini(nomeaplicativo.ini)f
'today() data atualf
'nulldate() data em branco
'fdata() formata data dd/mm/yyy
'GetPrivateProfileString -le ini files
'WritePrivateProfileString - grava ini files
'NetworkUserName Nome usuario rede
'MachineName Nome do equipamento agora na classe osinfo
'dividir divide dois numero com checagem tipo
'multiplicar multiplica dos numero com checagem tipo
'somaarr soma conteudo de uma array com checagem tipo
'mdg (retorna sim falso a uma pergunta
'pegcamini() Traz Caminho com opcao empresa/ano
'busca traz uma string com tamanho formatado
'FIXNUM retorna 0 se nao for numerico
'FIXSTR retorna "" se nao for caracter
'INCDEC aumenta ou diminui um numero
'alert imprime mensagem na tela
'XOREncryption Encriptar Mensagems
'FILECOPY COPIAR ARQUIVOS
'CompactAndRepairDB Compactar e Repar dbf

#If (VBA7 = 0) Then
Private Enum LongPtr
[_]
End Enum
#End If

#If Win64 Then
Private Const NULL_PTR As LongPtr = 0
Private Const PTR_SIZE As Long = 8
#Else
Private Const NULL_PTR As Long = 0
Private Const PTR_SIZE As Long = 4
#End If


Private CurMBTitle As String

'Public Const COLOR_SCROLLBAR = 0  'The Scrollbar color
'Public Const COLOR_BACKGROUND = 1  'Colour of the background with no wallpaper
'Public Const COLOR_ACTIVECAPTION = 2  'Caption of Active Window
'Public Const COLOR_INACTIVECAPTION = 3  'Caption of Inactive window
'Public Const COLOR_MENU = 4  'Menu
'Public Const COLOR_WINDOW = 5  'Windows background
'Public Const COLOR_WINDOWFRAME = 6  'Window frame
'Public Const COLOR_MENUTEXT = 7  'Window Text
'Public Const COLOR_WINDOWTEXT = 8  '3D dark shadow (Win95)
'Public Const COLOR_CAPTIONTEXT = 9  'Text in window caption
'Public Const COLOR_ACTIVEBORDER = 10  'Border of active window
'Public Const COLOR_INACTIVEBORDER = 11  'Border of inactive window
'Public Const COLOR_APPWORKSPACE = 12  'Background of MDI desktop
'Public Const COLOR_HIGHLIGHT = 13  'Selected item background
'Public Const COLOR_HIGHLIGHTTEXT = 14  'Selected menu item
'Public Const COLOR_BTNFACE = 15  'Button
'Public Const COLOR_BTNSHADOW = 16  '3D shading of button
'Public Const COLOR_GRAYTEXT = 17  'Grey text of zero if dithering is used.
'Public Const COLOR_BTNTEXT = 18  'Button text
'Public Const COLOR_INACTIVECAPTIONTEXT = 19  'Text of inactive window
'Public Const COLOR_BTNHIGHLIGHT = 20  '3D highlight of button

'Constante para pegar formato data hora do sistema
Public Const LOCALE_SDECIMAL = &HE               '---------------------------------- AGRUPAMENTO DE DIGITOS NORMAL
Public Const LOCALE_STHOUSAND = &HF              '----------------------------------------- SEPARADOR DECIMAL MONETÁRIO
Public Const LOCALE_SMONDECIMALSEP = &H16        '------------------------------------------ AGRUPAMENTO DE DIGITOS MONETÁRIOS
Public Const LOCALE_SMONTHOUSANDSEP = &H17       '------------------------------------- FORMATO DA DATA
Public Const LOCALE_SSHORTDATE = &H1F
Public Const LOCALE_SDATE = &H1D
Private Const WM_SETTINGCHANGE = &H1A
Private Const HWND_BROADCAST = &HFFFF&
Private Const CONNECT_LAN As Long = &H2
Private Const CONNECT_MODEM As Long = &H1
Private Const CONNECT_PROXY As Long = &H4
Private Const CONNECT_OFFLINE As Long = &H20
Private Const CONNECT_CONFIGURED As Long = &H40
Public Const WM_CLOSE = &H10
Public Const WM_UNDO = &H304
Public Const WM_PASTE = &H302


Private Const FO_COPY = &H2
Private Const FOF_ALLOWUNDO = &H40
Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const CB_FINDSTRING As Long = &H14C
Public Const CB_ERR As Long = (-1)
'Tirar botao maximar minizar
'Public Const WS_MINIMIZEBOX As Long = &H20000
'Public Const WS_MAXIMIZEBOX As Long = &H10000
Public Const GWL_STYLE As Long = (-16)
Public Const HELP_CONTENTS = &H3&

'Private Type SHFILEOPSTRUCT
'  hwnd As Long
'  wFunc As Long
'  pFrom As String
'  pTo As String
'  fFlags As Integer
'  fAnyOperationsAborted As Boolean
'  hNameMappings As Long
'  lpszProgressTitle As String
'End Type

Public Enum EShellShowConstants
  essSW_HIDE = 0
  essSW_MAXIMIZE = 3
  essSW_MINIMIZE = 6
  essSW_SHOWMAXIMIZED = 3
  essSW_SHOWMINIMIZED = 2
  essSW_SHOWNORMAL = 1
  essSW_SHOWNOACTIVATE = 4
  essSW_SHOWNA = 8
  essSW_SHOWMINNOACTIVE = 7
  essSW_SHOWDEFAULT = 10
  essSW_RESTORE = 9
  essSW_SHOW = 5
End Enum


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
#End If



Private Sub ForceSystemDecimalToPeriod()
    ' We MUST use the ANSI API version so it's an ANSI character that's used for the actual decimal character.
    Const LOCALE_SDECIMAL   As Long = &HE&
    Const LOCALE_SGROUPING  As Long = &H10&
    Const Eng_LCID          As Long = 1033&
    Dim s                   As String
    '
    s = String$(GetLocaleInfo(Eng_LCID, LOCALE_SDECIMAL, vbNullString, 0&), 0)
    GetLocaleInfo Eng_LCID, LOCALE_SDECIMAL, s, Len(s)
    If RTrimNull(s) <> "." Then
        SetLocaleInfo Eng_LCID, LOCALE_SDECIMAL, "."
        SetLocaleInfo Eng_LCID, LOCALE_SGROUPING, ","
    End If
End Sub

'Public Function tocar(cARQ)
'  PlaySound (cARQ), ByVal 0&, SND_FILENAME Or SND_ASYNC
'End Function

Public Function IsWebConnected(Optional ByRef ConnType As String) As Boolean
  Dim dwFlags As Long
  Dim WebTest As Boolean
  ConnType = ""
  WebTest = InternetGetConnectedState(CLng(dwFlags), 0&)
  Select Case WebTest
  Case dwFlags And CONNECT_LAN: ConnType = "LAN"
  Case dwFlags And CONNECT_MODEM: ConnType = "Modem"
  Case dwFlags And CONNECT_PROXY: ConnType = "Proxy"
  Case dwFlags And CONNECT_OFFLINE: ConnType = "Offline"
  Case dwFlags And CONNECT_CONFIGURED: ConnType = "Configured"
  'Case dwFlags And CONNECT_RAS: ConnType = "Remote"
  End Select
  IsWebConnected = WebTest
End Function

Public Function ExecutarDLL(ByVal cDLL As String) As Double
' appwiz.cpl adicionar e remover programas
' odbccp32.cpl gerenciador odbc
' sysdm.cpl informacoes do sistema
' outro cpl
' outra dll que possa ser executada via rundll32
'  Dim RetVal
  On Error Resume Next
  ExecutarDLL = shell("rundll32.exe shell32.dll,Control_RunDLL " & cDLL & ",,3", 1)
  
End Function

Public Function ExecuteLine(ByVal Scode As String, Optional ByVal fCheckOnly As Boolean) As Boolean
  ExecuteLine = EbExecuteLine(StrPtr(Scode), 0&, 0&, Abs(fCheckOnly)) = 0
End Function

Public Function ComboLostFocus(ByRef Combo1)
Dim strPartial
Dim i
  With Combo1
    If Len(.tEXT) Then
      'Procura pelo texto digitado
      strPartial = .tEXT
      i = SendMessage(.hwnd, CB_FINDSTRING, -1, ByVal strPartial)
      'Se não achou, retorna      o focus para o Combo
      If i = CB_ERR Then .SetFocus
    End If
  End With
End Function
Public Function pegue2delimitado(ByVal ctmpline As String, ByVal cDelimIni As String, ByVal cdelifim As String) As Variant
Dim nPOS As Integer
Dim nPOS2 As Integer
Dim nPOS3 As Integer
Dim cValor As String
Dim cINICIO As String
nPOS = 0
nPOS2 = 0
nPOS3 = 0
cValor = ""
cINICIO = ""
If Len(ctmpline) = 0 Then
   pegue2delimitado = Array("", "")
   Exit Function
End If
nPOS = InStr(ctmpline, cDelimIni)
If Len(cdelifim) = 0 Then
Else
  If nPOS > 0 Then
     nPOS2 = InStr(nPOS, ctmpline, cdelifim)
  Else
     nPOS2 = InStr(ctmpline, cdelifim)
  End If
End If
  
  If nPOS > 0 And nPOS2 = 0 Then
    If nPOS > 1 Then
      cINICIO = Mid(ctmpline, nPOS - 1)
    End If
    cValor = Mid(ctmpline, nPOS + Len(cDelimIni))
    ctmpline = ""
  End If
  If nPOS > 0 And nPOS2 > 0 Then
    cINICIO = Mid(ctmpline, 1, nPOS + Len(cDelimIni) - 1)
    cValor = Mid(ctmpline, nPOS + Len(cDelimIni))
    nPOS3 = InStr(cValor, cdelifim)
    If nPOS3 > 1 Then
       cValor = Mid(cValor, 1, nPOS3 - 1)
    End If
    
    If Right(cValor, Len(cdelifim)) = cdelifim Then
       cValor = Left(cValor, Len(cValor) - Len(cdelifim))
    End If
    ctmpline = Mid(ctmpline, nPOS2)
    If Left(ctmpline, Len(cdelifim)) = cdelifim Then
        ctmpline = Mid(ctmpline, Len(cdelifim) + 1)
    End If
    If ctmpline = cdelifim Then
       ctmpline = ""
    End If
  End If
pegue2delimitado = Array(cValor, ctmpline, cINICIO)
End Function
Public Function ComboChange(ByRef Combo1)
Dim strPartial
Dim i
Dim strTotal
Dim J
Dim m_bEditFromCode
  With Combo1
    'Procura pelo texto já digitado
    strPartial = .tEXT
    i = SendMessage(.hwnd, CB_FINDSTRING, -1, _
                    ByVal strPartial)

    'Se achou, adiciona o resto do Texto
    If i <> CB_ERR Then
      'Pega o texto inteiro
      strTotal = .List(i)

      'Compute number of unmatched characters
      J = Len(strTotal) - Len(strPartial)
      If J <> 0 Then
        'Adiciona o resto da string encontrada
        m_bEditFromCode = True
        .SelText = Right$(strTotal, J)

        'Marca os caracteres adicionados
        .SelStart = Len(strPartial)
        .SelLength = J
      End If
    End If
  End With
End Function

Public Sub MoveObject(ByRef Obj As Control)
  Screen.MousePointer = vbSizeAll
  SendMessage Obj.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 1
  ReleaseCapture
  Screen.MousePointer = vbDefault
End Sub

Public Function funNumeroPuro(ByVal pNumero) As String
  Dim i As Integer
  pNumero = FixStr(pNumero)
  funNumeroPuro = ""
  For i = 1 To Len(pNumero)
    If InStr("0123456789", Mid(pNumero, i, 1)) > 0 Then
      funNumeroPuro = funNumeroPuro & Mid(pNumero, i, 1)
    End If
  Next
End Function

Public Function FVar(ByVal eVAR As Variant, Optional ByVal cFORM As String = "", Optional ByVal ePAD As Variant)
  If cFORM <> "" Then
    Select Case cFORM
    Case "N"                                 'Numerico
      FVar = FixNum(eVAR, ePAD)
    Case "NI"                                'Numerico Inteiro
      FVar = FixInt(eVAR, ePAD)
    Case "C"                                 'Caracter
      FVar = LTrim(FixStr(eVAR, ePAD))
    Case "CU"                                'Caracter Maisculas
      FVar = UCase(FixStr(eVAR, ePAD))
    Case "CL"                                'Caracter Minusculas
      FVar = LCase(FixStr(eVAR, ePAD))
    Case "CC"                                'Caracter Capitalizado
      eVAR = FixStr(eVAR, ePAD)
      FVar = MMCase(eVAR)
    Case "D", "DN", "DS", "DC", "DF", "DH", "DD", "DZ", "D-"
      FVar = Fdata(eVAR, cFORM, ePAD)
      'Data dd/mm/yyyy
      'se for null
      'D OU DS  DateSerial(0, 0, 0)
      'DN Campo Null
      'DC Space(8)
      'DZ ''
      'DF mascara  "  /  /    "
      'DH data de hoje
      'DD date.
    Case "B", "BF", "BT"                     ''Boleano Boleano se null (BF)false (BT)TRUE
      If VarType(eVAR) = vbBoolean Then
        FVar = eVAR
      Else
        If cFORM = "BF" Or cFORM = "B" Then
          FVar = False
        Else
          FVar = True                  'BT
        End If
      End If
    Case "BSN"                               'Converte boleano SN
      If VarType(eVAR) = vbBoolean Then
        If eVAR Then
          FVar = "S"
        Else
          FVar = "N"
        End If
      Else
        If IsNull(ePAD) Then
          FVar = " "
        Else
          FVar = ePAD
        End If
      End If
    Case "BN"                                'Converte Boleano em Numerico
      If VarType(eVAR) = vbBoolean Then
        If eVAR Then
          FVar = 1
        Else
          FVar = 0
        End If
      Else
        If IsNull(ePAD) Then
          FVar = 0
        Else
          FVar = ePAD
        End If
      End If
    Case "CPF", "CIC"                        '' 111.222.333-44   14 Digitos
      eVAR = TiraOut(eVAR)
      eVAR = Mid(eVAR, 1, 3) & "." & Mid(eVAR, 4, 3) & "." & _
             Mid(eVAR, 7, 3) & "-" & Mid(eVAR, 10, 2)
    Case "CNPJ", "CGC"                       '' 11.222.333/4444-55 18 digitos
      eVAR = TiraOut(eVAR)
      FVar = Mid(eVAR, 1, 2) & "." & Mid(eVAR, 3, 3) & "." & _
             Mid(eVAR, 6, 3) & "/" & Mid(eVAR, 9, 4) & "-" & Mid(eVAR, 13, 2)
    Case "RG"
      FVar = FormataRG(eVAR)
      '            Case "IE" Precisa estado
      '                FVar = FormataIE(eVAR)
    Case "CHAPA"
      eVAR = TiraOut(eVAR)
      FVar = Mid(eVAR, 1, 3) & "-" & Mid(eVAR, 4)
    Case "CEP"                               'Cep NNNNN-NNN
      eVAR = TiraOut(eVAR)
      FVar = Mid(eVAR, 1, 5) & "-" & Mid(eVAR, 6)
    Case "TEL"                               'Cep NNNN-NNNN
      eVAR = TiraOut(eVAR)
      FVar = Mid(eVAR, 1, 4) & "-" & Mid(eVAR, 5)
    Case "H"
      FVar = Format(eVAR, "HH.MM")
    Case "MES"                               'Numerico Mes do Ano 1-12
      FVar = FixInt(eVAR, ePAD)
      If FVar < 1 And FVar > 12 Then
        FVar = 0
      End If
    Case "ANO"                               'Numerico  Ano 1900-2300
      FVar = FixInt(eVAR, ePAD)
      If FVar < 1900 And FVar > 2300 Then
        FVar = 0
      End If

    Case Else                                'Retorna o Campo Sem Formataçao
      FVar = eVAR
    End Select
  Else
    FVar = eVAR
  End If
End Function

Public Function GeraSplit(ByVal aVAR As Variant, Optional ByVal cINI As String = "", _
                          Optional ByVal cMID As String = "", _
                          Optional ByVal cFIM As String = "") As String
  Dim nUSO As Long
  Dim x As Long
  nUSO = UBound(aVAR) - 1
  GeraSplit = cINI & aVAR(0) & cMID
  For x = 1 To nUSO
    GeraSplit = GeraSplit & aVAR(x) & cMID
  Next x
  GeraSplit = GeraSplit & aVAR(nUSO + 1)
  GeraSplit = GeraSplit & cFIM
End Function

Public Function SepSqlOpe(ByVal eEXP As String) As Variant
  Dim aRETU As Variant
  aRETU = Array("", "", "")
  If InStr(eEXP, ",") > 0 Then
    aRETU = Split(eEXP, ",")
  End If
  SepSqlOpe = aRETU
End Function
Public Function CharToLit(aVAL As Variant, Optional ByVal cTIPO As String = "", Optional ByVal cARQ As String = "") As String
Dim aRETU As Variant

  If Len(cARQ) > 0 Then
    cARQ = GeraConn(cARQ, cTIPO)
    aRETU = TipoConn(cARQ)
    cTIPO = aRETU(2)
  End If
  Select Case cTIPO
  Case "MYSQL", "MARIADB"
    CharToLit = GeraSplit(aVAL, "CONCAT(", ",", ")")
  Case Else
    CharToLit = GeraSplit(aVAL, "", " + ", "")
  End Select
  Exit Function
Erro:
  Select Case Err.Number
  Case Else
    SayErro "ChartoLit"
  End Select
End Function


Public Function Alert(ByVal cDIZ As String, Optional ByVal cTITLE As String = "Informacao")
  MsgBox cDIZ, vbOKOnly, cTITLE
End Function

Public Function Busca(ByVal cDIZ As String, ByVal cCAB As String, ByVal cVAL As String, Optional ByVal nLEN As Integer = 0) As String
  Busca = InputBox(cDIZ, cCAB, cVAL)
  If nLEN > 0 Then
     Busca = Left$(Busca, nLEN)
  End If
End Function

Public Function Caminex(ByVal cARQ As String, Optional ByVal nANO As Integer = 0, Optional ByVal nMES As Integer = 0, Optional ByVal nEMP = 1)
  Dim cAno As String
  Dim cEMP As String
  Dim cMes As String
  Dim cTMPVAR As String
  Dim nPOS  As Integer
  Dim nPOS2 As Integer
  nPOS = InStr(cARQ, "{")
  If nPOS > 0 Then
    nPOS2 = InStr(cARQ, "}")
    If nPOS2 > 0 Then
      cTMPVAR = Mid(cARQ, nPOS + 1, nPOS2 - nPOS - 1)
      cTMPVAR = Mid(cARQ, 1, nPOS - 1) + PegPath("PATH", cTMPVAR) + Mid(cARQ, nPOS2 + 1)
      cARQ = cTMPVAR
    End If
  End If

  If nANO = 0 And InStr(cARQ, "[A") > 0 Then
    frmcomp.Show vbModal
    nANO = Val(eRETU01)
    nMES = Val(eRETU02)
    nEMP = Val(eRETU03)
  End If

  If nMES = 0 And InStr(cARQ, "[M") > 0 Then
    frmcomp.Show vbModal
    nANO = Val(eRETU01)
    nMES = Val(eRETU02)
    nEMP = Val(eRETU03)
  End If

  If nEMP = 0 And InStr(cARQ, "[Z") > 0 Then
    frmcomp.Show vbModal
    nANO = Val(eRETU01)
    nMES = Val(eRETU02)
    nEMP = Val(eRETU03)
  End If

  If nANO = 0 And InStr(cARQ, "[A") > 0 Then
    frmcomp.Show vbModal
    nANO = Val(eRETU01)
    nMES = Val(eRETU02)
    nEMP = Val(eRETU03)
  End If

  If InStr(cARQ, "[MM]") > 0 Then
    cMes = StrZero(nMES, 2)
    cARQ = Replace(cARQ, "[MM]", cMes)
  End If
  If InStr(cARQ, "[AAAA]") > 0 Then
    cAno = StrZero(nANO, 4)
    cARQ = Replace(cARQ, "[AAAA]", cAno)
  End If
  If InStr(cARQ, "[AA]") > 0 Then
    cAno = StrZero(nANO, 4)
    cAno = Right(cAno, 2)
    cARQ = Replace(cARQ, "[AA]", cAno)
  End If
  If InStr(cARQ, "[ZZZ]") > 0 Then
    cEMP = StrZero(nEMP, 3)
    cARQ = Replace(cARQ, "[ZZZ]", cEMP)
  End If
  If InStr(cARQ, "[ZZ]") > 0 Then
    cEMP = StrZero(nEMP, 2)
    cARQ = Replace(cARQ, "[ZZ]", cEMP)
  End If
  If InStr(cARQ, "[Z]") > 0 Then
    cEMP = StrZero(nEMP, 1)
    cARQ = Replace(cARQ, "[Z]", cEMP)
  End If
  Caminex = cARQ
End Function
Public Function Dividir(ByVal nVAL As Variant, ByVal nDIV As Variant)
  Dividir = 0
  nVAL = FixNum(nVAL)
  nDIV = FixNum(nDIV)
  If nVAL > 0 And nDIV > 0 Then
    Dividir = nVAL / nDIV
  End If
  
End Function
'Public Function Dividir(ByVal n1 As Variant, ByVal n2 As Variant) As Double
    ' 1. Tratamento preventivo:
    ' Se n1 ou n2 não forem números, o RC6 trata de forma limpa.
    ' Evita o erro de divisão por zero ou erro de tipo em 64-bits.
    
 '   Dim d1 As Double, d2 As Double
    
    ' O New_c.Val garante a conversão correta independente do idioma
    ' (ponto vs vírgula) e da arquitetura (32/64).
  '  d1 = New_c.Val(n1)
   ' d2 = New_c.Val(n2)
    
    ' 2. Lógica de segurança de divisão
    'If d2 <> 0 Then
     '   Dividir = d1 / d2
    'Else
     '   Dividir = 0 ' Ou trate o erro da forma que seu sistema exige
    'End If
'End Function


Public Function FileConnExist(ByVal cARQ As Variant, _
                          Optional ByVal lMES As Boolean = False, _
                          Optional ByVal cMes As String = "Arquivo Não Encontrado ", _
                          Optional ByVal cSQL As String = "")
  Dim cARQUIVO As String
  Dim nFILELEN
  Dim nPOS As Long
  On Error GoTo NotExist
  FileConnExist = False
  cARQUIVO = UCase(Trim(CStr(cARQ)))
  If Len(cARQ) = 0 Then
     FileConnExist = False
     Exit Function
  End If

  If InStr(cARQUIVO, "[") > 0 Then
     
    If InStr(cARQUIVO, "[JET") > 0 Or InStr(cARQUIVO, "[SQLITE]") Or InStr(cARQUIVO, "[ACCDB") > 0 Then
      cARQUIVO = Replace(cARQUIVO, "[JET", "")
      cARQUIVO = Replace(cARQUIVO, "[SQLITE]", "")
      cARQUIVO = Replace(cARQUIVO, "[ACCDB", "")
      cARQUIVO = Replace(cARQUIVO, "MDB]", "")
      cARQUIVO = Replace(cARQUIVO, "]", "")
    Else
      If InStr(cARQUIVO, "[JETFOX]") > 0 Or InStr(cARQUIVO, "[ADS") > 0 Then
        cARQUIVO = Replace(cARQUIVO, "[JETFOX]", "")
        cARQUIVO = Replace(cARQUIVO, "[ADS", "")
        cARQUIVO = Replace(cARQUIVO, "CDX]", "")
        cARQUIVO = Replace(cARQUIVO, "NTX]", "")
        cARQUIVO = Replace(cARQUIVO, "ADT]", "")
        cARQUIVO = Replace(cARQUIVO, "]", "")
        If Len(cSQL) = 0 Then
          Alert ("Fileexit  Nao Passado cSQL Necessario Para DBF")
          FileConnExist = True
          Exit Function
        End If
        cARQUIVO = cARQUIVO & NomeTableSql(cSQL, ".DBF")
      End If

      If InStr(cARQUIVO, "PDX3]") > 0 Or InStr(cARQUIVO, "PDX4]") Or InStr(cARQUIVO, "PDX5]") > 0 Then
        cARQUIVO = Replace(cARQUIVO, "PDX3]", "")
        cARQUIVO = Replace(cARQUIVO, "PDX4]", "")
        cARQUIVO = Replace(cARQUIVO, "PDX5]", "")
        cARQUIVO = Replace(cARQUIVO, "[A16", "")
        cARQUIVO = Replace(cARQUIVO, "[A15", "")
        cARQUIVO = Replace(cARQUIVO, "[A14", "")
        cARQUIVO = Replace(cARQUIVO, "[A12", "")
        cARQUIVO = Replace(cARQUIVO, "[JET", "")
        cARQUIVO = Replace(cARQUIVO, "]", "")
        If Len(cSQL) = 0 Then
          Alert ("Fileexit  Nao Passado cSQL Necessario Para DB")
          FileConnExist = True
          Exit Function
        End If
        cARQUIVO = cARQUIVO & NomeTableSql(cSQL, ".DB")
      End If

      If InStr(cARQUIVO, "[CONN]") > 0 Or InStr(cARQUIVO, "[MYSQL]") > 0 Or InStr(cARQUIVO, "[MARIADB]") > 0 Or _
         InStr(cARQUIVO, "[PGSQL]") > 0 Or InStr(cARQUIVO, "[POSTGRESQL]") > 0 Or InStr(cARQUIVO, "[SQLSERVER]") > 0 Then
        FileConnExist = True
        Exit Function
      End If
      If InStr(cARQUIVO, "[FIREBIRD]") > 0 Or InStr(cARQUIVO, "[INTERBASE]") > 0 Then
        nPOS = InStr(cARQUIVO, "DBNAME=")
        cARQUIVO = Mid(cARQUIVO, nPOS + 7)
        nPOS = InStr(cARQUIVO, ";")
        If nPOS > 0 Then
          cARQUIVO = Mid(cARQUIVO, 1, nPOS - 1)
        End If
        FileConnExist = True
        Exit Function
      End If
      If InStr(cARQUIVO, "[XLSEXT]") > 0 Then
        cARQUIVO = Replace(cARQUIVO, "[XLSEXT]", "")
      End If
      If InStr(cARQUIVO, "[XLS]") > 0 Then
        cARQUIVO = Replace(cARQUIVO, "[XLS]", "")
      End If
    End If
  End If
  If InStr(cARQUIVO, "PROVIDER") > 0 Or InStr(cARQUIVO, ".MARIADB") > 0 Or InStr(cARQUIVO, ".MYSQL") > 0 Then
    FileConnExist = True
    Exit Function
  End If
  If FileExists(cARQUIVO) Then
    
    FileConnExist = True
  End If

  Exit Function
NotExist:
  FileConnExist = False
  If lMES Then
    Alert cMes & cARQUIVO, "Erro Arquivo"
  End If
End Function

Public Function FixInt(ByVal cUSO As Variant, Optional ByVal ePAD As Variant = 0)
  If IsNull(ePAD) Or ePAD = "" Then            'necessario pois pode ser passado outros tipos
    ePAD = 0
  End If
  If IsNull(cUSO) Or cUSO = "" Then
    FixInt = ePAD
    Exit Function
  End If
  If Not IsNumeric(cUSO) Then
    cUSO = funNumeroPuro(cUSO)
  End If
  If IsNull(cUSO) Or cUSO = "" Then            'necessario apos funnumeropuro
    FixInt = ePAD
    Exit Function
  End If
  FixInt = Int(cUSO)
End Function
Public Function senhapos(cPOSTELA)
  Dim cSUBTXT As String
  Dim nCHAR As Integer
  senhapos = ""
  While Len(cPOSTELA) > 2  'no fim pois postela  usa reverse string
    cSUBTXT = Mid(cPOSTELA, Len(cPOSTELA) - 2, 3)
    nCHAR = FixInt(cSUBTXT)
    nCHAR = nCHAR / 2
    nCHAR = nCHAR - 30

    senhapos = senhapos + Chr(nCHAR)
    cPOSTELA = Mid(cPOSTELA, 1, Len(cPOSTELA) - 3)
  Wend
  ''//Alert (senhapos)
End Function


Public Function FixNum(ByVal cUSO As Variant, Optional ByVal ePAD As Variant = 0) As Variant
  Dim nPOS01 As Integer
  Dim nPOS02 As Integer
  If IsNull(ePAD) Then
    ePAD = 0
  End If
  If IsNull(cUSO) Or cUSO = "" Then
    ''cUSO = ePAD
    FixNum = ePAD
    Exit Function
  End If
  If Not IsNumeric(cUSO) Then
    cUSO = CStr(cUSO)
    nPOS01 = InStr(cUSO, ",")
    nPOS02 = InStr(cUSO, ".")
    If nPOS02 = 0 And nPOS01 > 0 Then        'Somente Virgula 99,99
      cUSO = Replace(cUSO, ",", ".")       ''muda virgula 99.99
    End If
    If nPOS01 > 0 And nPOS02 > 0 Then
      If nPOS02 < nPOS01 Then              'Milhar Ponto Decimal Virgula 99.999,99
        cUSO = Replace(cUSO, ".", "")    ''Tira os Pontos 99999,99
        cUSO = Replace(cUSO, ",", ".")   ''muda virgula  99999.99
      Else                                 'Milhar virgula ponto Decimal  99,999.99
        cUSO = Replace(cUSO, ",", "")    ''tira virgula
      End If
    End If
  End If
  FixNum = CDbl(cUSO)
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

'Public Function INCDEC(ByVal eVAR, ByVal nVAL, ByVal nLIMINF, ByVal nLIMSUP)
'  INCDEC = FixNum(eVAR)
'  INCDEC = INCDEC + nVAL
'  If INCDEC < nLIMINF Then INCDEC = nLIMINF
'  If INCDEC > nLIMSUP Then INCDEC = nLIMSUP
'End Function
' Mantemos exatamente a mesma assinatura, preservando os parâmetros que você já usa
Public Function INCDEC(ByRef vValor As Variant, ByVal nIncremento As Variant, Optional ByVal nLimiteMax As Variant, Optional ByVal nLimiteMin As Variant) As Variant
    
    ' 1. Validação estrita usando sua FixNum (que você já confia)
    ' 2. Processamos a lógica mantendo os limites que você precisa
    If FixNum(vValor) And FixNum(nIncremento) Then
        
        vValor = vValor + nIncremento
        
        ' Lógica de limite (se aplicável)
        If Not IsMissing(nLimiteMax) Then
            If vValor > nLimiteMax Then vValor = nLimiteMax
        End If
        
        If Not IsMissing(nLimiteMin) Then
            If vValor < nLimiteMin Then vValor = nLimiteMin
        End If
        
        INCDEC = vValor
    Else
        INCDEC = vValor
    End If
    
End Function

Public Function MDG(ByVal cMEnSSAGEM As String, Optional cTITULO = "Confirme")
  Dim eRESP As Variant
  MDG = False
  eRESP = MsgBox(cMEnSSAGEM, vbYesNo + vbDefaultButton1, cTITULO)
  If eRESP = vbYes Then
    MDG = True
  End If
End Function

Public Function MathOper(ByVal nVAL01 As Variant, ByVal nVAL02 As Variant, _
                         Optional ByVal cOPE As String = "&", _
                         Optional ByVal nROUND As Integer = -1)
  Dim eVAL As Variant
  Dim eTMP As Variant
  eVAL = 0
  Select Case cOPE
  Case "ASCII"
    If Len(nVAL01) > 0 Then
      eVAL = Asc(nVAL01)
    Else
      eVAL = 32
    End If
  Case "EXT"
    eVAL = UCase(EXTENSAO(FixStr(nVAL01, FixStr(nVAL02))))
  Case "NPD"
    eTMP = FixInt(nVAL01)
    If eTMP > 0 Then
      eVAL = Fdata(nVAL02)
    Else
      eVAL = Space(8)
    End If
  Case "PCF"
    If FixStr(nVAL01) = "PC" Then
      eVAL = Format(nVAL02, "###,###,##0")
    Else
      eVAL = Format(nVAL02, "###,###,##0.00")
    End If
  Case "*"
    eVAL = Multiplicar(nVAL01, nVAL02)
  Case "/"
    eVAL = Dividir(nVAL01, nVAL02)
  Case "%"
    eVAL = Dividir(nVAL01, nVAL02)
    eVAL = Multiplicar(eVAL, 100)
  Case "+"
    eVAL = FixNum(nVAL01) + FixNum(nVAL02)
  Case "CNPJCPF"
    Select Case nVAL01
    Case "J"
      eVAL = FVar(nVAL02, "CNPJ")
    Case "F"
      eVAL = FVar(nVAL02, "CPF")
    Case Else
      eVAL = nVAL02
    End Select
  Case "-"
    eVAL = FixNum(nVAL01) - FixNum(nVAL02)
  Case "&"
    eVAL = Concatenar(nVAL01, nVAL02, "")
  Case "& ", "& / ", "&/", "&-", "& - "
    eVAL = Concatenar(nVAL01, nVAL02, cOPE)
  Case Else                                    ''Operacao Nao Definida cONCATENA
    eVAL = FixStr(nVAL01) & FixStr(nVAL02)
  End Select
  If nROUND > 0 Then
    eVAL = Round(eVAL, nROUND)
  End If
  MathOper = eVAL
End Function

Public Function Concatenar(ByVal eVAR01 As Variant, ByVal eVAR02, Optional ByVal cSEP As String = "") As String
  cSEP = Replace(cSEP, "&", "")
  Concatenar = FixStr(eVAR01) & cSEP & FixStr(eVAR02)
End Function

Public Function Multiplicar(ByVal nVAL As Variant, ByVal nMUL As Variant)
  Multiplicar = 0
  nVAL = FixNum(nVAL)
  nMUL = FixNum(nMUL)
  If nVAL > 0 And nMUL > 0 Then
    Multiplicar = nVAL * nMUL
  End If
End Function


Public Function PadRight(ByVal cTEXTO, ByVal nLEN) As String
  cTEXTO = cTEXTO & Space(nLEN)
  cTEXTO = Left(cTEXTO, nLEN)
  PadRight = cTEXTO
End Function

Public Function PadLeft(ByVal cTEXTO, ByVal nLEN) As String
  cTEXTO = Space(nLEN) & cTEXTO
  cTEXTO = Right(cTEXTO, nLEN)
  PadLeft = cTEXTO
End Function

Public Function PegCamini(ByVal cCaminho As String) As String
  PegCamini = Caminex(cCaminho, 0, 0, 0)
End Function
Public Function PegINIVAL(ByVal cARQINI As String, ByVal cGRUPO As String, ByVal cCAMPO As String, Optional ByVal ePAD As String = "") As String
  Dim z As Long
  Dim scaminho As String * 255
  z = GetPrivateProfileString(cGRUPO, cCAMPO, "", scaminho, 150, cARQINI)
  PegINIVAL = IIf(Asc(Left(scaminho, 1)) = "0", ePAD, Left(scaminho, z))

End Function

Public Function PegPath(ByVal cGRUPO As String, ByVal cCAMPO As String, Optional ByVal ePAD As String = "" _
                       , Optional ByVal cARQINI As String = "") As String
  Dim z As Long
  Dim scaminho As String * 255
  If cARQINI = "" Then
     cARQINI = App.Path & "\" & App.EXEName & ".INI"
  End If
  z = GetPrivateProfileString(cGRUPO, cCAMPO, "", scaminho, 150, cARQINI)
  If Len(ePAD) = 0 Then
    ePAD = App.Path & "\"
  End If
  PegPath = IIf(Asc(Left(scaminho, 1)) = "0", ePAD, Left(scaminho, z))

End Function

Public Function PegTable(ByVal cGRUPO As String, ByVal cCAMPO As String)
  Dim z As Long
  Dim scaminho As String * 255
  Dim cTMP As String
  z = GetPrivateProfileString(cGRUPO, cCAMPO, "", scaminho, 150, App.Path + "\" & App.EXEName & ".INI")
  cTMP = IIf(Asc(Left(scaminho, 1)) = "0", App.Path & "\", Left(scaminho, z))
  PegTable = PegPath("PATH", cTMP)
End Function

Public Function MMCase(ByVal texto As String) As String
  Dim sPalavra As String
  Dim iPosIni As Integer
  Dim iPosFim As Integer
  Dim sResultado As String
  iPosIni = 1
  texto = LCase(texto) & " "
  sResultado = ""
  Do Until InStr(iPosIni, texto, " ") = 0
    iPosFim = InStr(iPosIni, texto, " ")
    sPalavra = Mid(texto, iPosIni, iPosFim - iPosIni)
    iPosIni = iPosFim + 1
    If sPalavra <> "de" And sPalavra <> "da" And _
       sPalavra <> "do" And sPalavra <> "das" _
       And sPalavra <> "dos" And sPalavra <> _
       "a" And sPalavra <> "e" Then
      sPalavra = UCase(Left(sPalavra, 1)) & _
                 LCase(Mid(sPalavra, 2))
    End If
    sResultado = sResultado & " " & sPalavra
  Loop
  MMCase = Trim(sResultado)
End Function

Public Sub SayErro(Optional ByVal cERROUSO As String = "", Optional ByVal lMES As Boolean = True)
  Dim nHANDLE As Long
  Dim cARQ As String
  Dim cERRO As String
     Dim osver
  cERRO = ""
  On Error Resume Next
  If Err.Number <> 0 Then
    Screen.MousePointer = vbDefault
    cARQ = zUSER & Format(Now, "_DDMMYYYY_HHMMSS") & ".TXT"

    cERRO = " No.        : " & Err.Number & Chr(13) & Chr(10) & _
          " Fonte      : " & Err.Source & Chr(13) & Chr(10) & _
          " Descricao  : " & Err.Description & Chr(13) & Chr(10) & _
          " Aplicativo : " & App.EXEName & " " & App.Major & "." & App.Minor & ". " & App.Revision & Chr(13) & Chr(10)
    '           " Linha      : " & Erl & Chr(13) & Chr(10) & _


                If Not IsNull(Err.LastDllError) And Err.LastDllError <> 0 Then
    cERRO = cERRO & " DLL        : " & Err.LastDllError & Chr(13) & Chr(10)
  End If
  If IsObject(Screen.ActiveForm) Then
    '' cERRO = cERRO & " Formulario : " & Screen.ActiveForm.Name
    '' If Not IsNull(Screen.ActiveForm.HelpContextID) Then
    ''    cERRO = cERRO & " ID         : " & Screen.ActiveForm.HelpContextID
    '' End If
  End If

  If IsObject(Screen.ActiveControl) Then
    ''cERRO = cERRO & " Em Uso     : " & Screen.ActiveControl
  End If


  'cERRO = cERRO & " Equipamento: " & MachineName & vbCrLf 'abaixo classe
  cERRO = cERRO & " Usuario    : " & NetworkUserName & vbCrLf
  'cERRO = cERRO & " IP         : " & GetIPAddress & vbCrLf  'abaixo classe
  'cERRO = cERRO & " Local      : " & GetIPHostName & vbCrLf 'abaixo classe

  cERRO = cERRO & " Erro       : " & cERROUSO
  If lMES Then
    TimedMsgBox cERRO, , , "Mensagem Sistema"  ''usuario ficavam pendurados travando registros agora com timeout
    'Alert cERRO, "Informacao de Erro "
  End If


  Select Case Err.Number
  Case -2147483647
    cERRO = cERRO & vbCrLf & "not implemented"
  Case -2147483646
    cERRO = cERRO & vbCrLf & "ran out of memory"
  Case -2147483645
    cERRO = cERRO & vbCrLf & "one or more arguments are invalid"
  Case -2147483644
    cERRO = cERRO & vbCrLf & "no such interface supported"
  Case -2147483643
    cERRO = cERRO & vbCrLf & "invalid pointer"
  Case -2147483642
    cERRO = cERRO & vbCrLf & "invalid handle"
  Case -2147483641
    cERRO = cERRO & vbCrLf & "operation aborted"
  Case -2147483640
    cERRO = cERRO & vbCrLf & "unspecified error"
  Case -2147483639
    cERRO = cERRO & vbCrLf & "general access denied error"
  Case -2147483638
    cERRO = cERRO & vbCrLf & "the data necessary to complete this operation is not yet available."
  Case -2147467263
    cERRO = cERRO & vbCrLf & "not implemented"
  Case -2147467262
    cERRO = cERRO & vbCrLf & "no such interface supported"
  Case -2147467261
    cERRO = cERRO & vbCrLf & "invalid pointer"
  Case -2147467260
    cERRO = cERRO & vbCrLf & "operation aborted"
  Case -2147467259
    cERRO = cERRO & vbCrLf & "unspecified error"
  Case -2147467258
    cERRO = cERRO & vbCrLf & "thread local storage failure"
  Case -2147467257
    cERRO = cERRO & vbCrLf & "get shared memory allocator failure"
  Case -2147467256
    cERRO = cERRO & vbCrLf & "get memory allocator failure"
  Case -2147467255
    cERRO = cERRO & vbCrLf & "unable to initialize class cache"
  Case -2147467254
    cERRO = cERRO & vbCrLf & "unable to initialize rpc services"
  Case -2147467253
    cERRO = cERRO & vbCrLf & "cannot set thread local storage channel control"
  Case -2147467252
    cERRO = cERRO & vbCrLf & "could not allocate thread local storage channel control"
  Case -2147467251
    cERRO = cERRO & vbCrLf & "the user supplied memory allocator is unacceptable"
  Case -2147467250
    cERRO = cERRO & vbCrLf & "the ole service mutex already exists"
  Case -2147467249
    cERRO = cERRO & vbCrLf & "the ole service file mapping already exists"
  Case -2147467248
    cERRO = cERRO & vbCrLf & "unable to map view of file for ole service"
  Case -2147467247
    cERRO = cERRO & vbCrLf & "failure attempting to launch ole service"
  Case -2147467246
    cERRO = cERRO & vbCrLf & "there was an attempt to call coinitialize a second time while single threaded"
  Case -2147467245
    cERRO = cERRO & vbCrLf & "a remote activation was necessary but was not allowed"
  Case -2147467244
    cERRO = cERRO & vbCrLf & "a remote activation was necessary but the server name provided was invalid"
  Case -2147467243
    cERRO = cERRO & vbCrLf & "the class is configured to run as a security id different from the caller"
  Case -2147467242
    cERRO = cERRO & vbCrLf & "use of ole1 services requiring dde windows is disabled"
  Case -2147467241
    cERRO = cERRO & vbCrLf & "a runas specification must be \ or simply 0x80004018 -2147467240 the server process could not be started. the pathname may be incorrect."
  Case -2147467239
    cERRO = cERRO & vbCrLf & "the server process could not be started as the configured identity. the pathname may be incorrect or unavailable."
  Case -2147467238
    cERRO = cERRO & vbCrLf & "the server process could not be started because the configured identity is incorrect. check the username and password."
  Case -2147467237
    cERRO = cERRO & vbCrLf & "the client is not allowed to launch this server."
  Case -2147467236
    cERRO = cERRO & vbCrLf & "the service providing this server could not be started."
  Case -2147467235
    cERRO = cERRO & vbCrLf & "this computer was unable to communicate with the computer providing the server."
  Case -2147467234
    cERRO = cERRO & vbCrLf & "the server did not respond after being launched."
  Case -2147467233
    cERRO = cERRO & vbCrLf & "the registration information for this server is inconsistent or incomplete."
  Case -2147467232
    cERRO = cERRO & vbCrLf & "the registration information for this interface is inconsistent or incomplete."
  Case -2147467231
    cERRO = cERRO & vbCrLf & "the operation attempted is not supported."
  Case -2147418113
    cERRO = cERRO & vbCrLf & "catastrophic failure"
  Case -2147024891
    cERRO = cERRO & vbCrLf & "general access denied error"
  Case -2147024890
    cERRO = cERRO & vbCrLf & "invalid handle"
  Case -2147024882
    cERRO = cERRO & vbCrLf & "ran out of memory"
  Case -2147024809
    cERRO = cERRO & vbCrLf & "one or more arguments are invalid"
  Case -2147217920
    cERRO = cERRO & vbCrLf & "invalid accessor"
  Case -2147217919
    cERRO = cERRO & vbCrLf & "creating another row would have exceeded the total number of active rows supported by the rowset"
  Case -2147217918
    cERRO = cERRO & vbCrLf & "unable to write with a read-only accessor"
  Case -2147217917
    cERRO = cERRO & vbCrLf & "given values violate the database schema"
  Case -2147217916
    cERRO = cERRO & vbCrLf & "invalid row handle"
  Case -2147217915
    cERRO = cERRO & vbCrLf & "an object was open"
  Case -2147217914
    cERRO = cERRO & vbCrLf & "invalid chapter"
  Case -2147217913
    cERRO = cERRO & vbCrLf & "a literal value in the command could not be converted to the correct type due to a reason other than data overflow"
  Case -2147217912
    cERRO = cERRO & vbCrLf & "invalid binding info"
  Case -2147217911
    cERRO = cERRO & vbCrLf & "permission denied"
  Case -2147217910
    cERRO = cERRO & vbCrLf & "specified column does not contain bookmarks or chapters"
  Case -2147217909
    cERRO = cERRO & vbCrLf & "some cost limits were rejected"
  Case -2147217908
    cERRO = cERRO & vbCrLf & "no command has been set for the command object"
  Case -2147217907
    cERRO = cERRO & vbCrLf & "unable to find a query plan within the given cost limit"
  Case -2147217906
    cERRO = cERRO & vbCrLf & "invalid bookmark"
  Case -2147217905
    cERRO = cERRO & vbCrLf & "invalid lock mode"
  Case -2147217904
    cERRO = cERRO & vbCrLf & "no value given for one or more required parameters"
  Case -2147217903
    cERRO = cERRO & vbCrLf & "invalid column id"
  Case -2147217902
    cERRO = cERRO & vbCrLf & "invalid ratio"
  Case -2147217901
    cERRO = cERRO & vbCrLf & "invalid value"
  Case -2147217900
    cERRO = cERRO & vbCrLf & "the command contained one or more errors"
  Case -2147217899
    cERRO = cERRO & vbCrLf & "the executing command cannot be canceled"
  Case -2147217898
    cERRO = cERRO & vbCrLf & "the provider does not support the specified dialect"
  Case -2147217897
    cERRO = cERRO & vbCrLf & "a data source with the specified name already exists"
  Case -2147217896
    cERRO = cERRO & vbCrLf & "the rowset was built over a live data feed and cannot be restarted"
  Case -2147217895
    cERRO = cERRO & vbCrLf & "no key matching the described characteristics could be found within the current range"
  Case -2147217894
    cERRO = cERRO & vbCrLf & "ownership of this tree has been given to the provider"
  Case -2147217893
    cERRO = cERRO & vbCrLf & "the provider is unable to determine identity for newly inserted rows"
  Case -2147217892
    cERRO = cERRO & vbCrLf & "no nonzero weights specified for any goals supported, so goal was rejected; current goal was not changed"
  Case -2147217891
    cERRO = cERRO & vbCrLf & "requested conversion is not supported"
  Case -2147217890
    cERRO = cERRO & vbCrLf & "lrowsoffset would position you past either end of the rowset, regardless of the crows value specified; crowsobtained is 0"
  Case -2147217889
    cERRO = cERRO & vbCrLf & "information was requested for a query, and the query was not set"
  Case -2147217888
    cERRO = cERRO & vbCrLf & "provider called a method from irowsetnotify in the consumer and nt"
  Case -2147217887
    cERRO = cERRO & vbCrLf & "errors occurred"
  Case -2147217886
    cERRO = cERRO & vbCrLf & "a non-null controlling iunknown was specified and the object being created does not support aggregation"
  Case -2147217885
    cERRO = cERRO & vbCrLf & "a given hrow referred to a hard- or soft- deleted row"
  Case -2147217884
    cERRO = cERRO & vbCrLf & "the rowset does not support fetching backwards"
  Case -2147217883
    cERRO = cERRO & vbCrLf & "all hrows must be released before new ones can be obtained"
  Case -2147217882
    cERRO = cERRO & vbCrLf & "one of the specified storage flags was not supported"
  Case -2147217881
    cERRO = cERRO & vbCrLf & "the comparison operator was invalid"
  Case -2147217880
    cERRO = cERRO & vbCrLf & "the specified status flag was neither dbcolumnstatus_ok nor dbcolumnstatus_isnull"
  Case -2147217879
    cERRO = cERRO & vbCrLf & "the rowset cannot scroll backwards"
  Case -2147217878
    cERRO = cERRO & vbCrLf & "invalid region handle"
  Case -2147217877
    cERRO = cERRO & vbCrLf & "the specified set of rows was not contiguous to or overlapping the rows in the specified watch region"
  Case -2147217876
    cERRO = cERRO & vbCrLf & "a transition from all* to move* or extend* was specified"
  Case -2147217875
    cERRO = cERRO & vbCrLf & "the specified region is not a proper subregion of the region identified by the given watch region handle"
  Case -2147217874
    cERRO = cERRO & vbCrLf & "the provider does not support multi-statement commands"
  Case -2147217873
    cERRO = cERRO & vbCrLf & "a specified value violated the integrity constraints for a column or table"
  Case -2147217872
    cERRO = cERRO & vbCrLf & "the given type name was unrecognized"
  Case -2147217871
    cERRO = cERRO & vbCrLf & "execution aborted because a resource limit has been reached; no results have been returned"
  Case -2147217870
    cERRO = cERRO & vbCrLf & "cannot clone a command object whose command tree contains a rowset or rowsets"
  Case -2147217869
    cERRO = cERRO & vbCrLf & "cannot represent the current tree as text"
  Case -2147217868
    cERRO = cERRO & vbCrLf & "the specified index already exists"
  Case -2147217867
    cERRO = cERRO & vbCrLf & "the specified index does not exist"
  Case -2147217866
    cERRO = cERRO & vbCrLf & "the specified index was in use"
  Case -2147217865
    cERRO = cERRO & vbCrLf & "the specified table does not exist"
  Case -2147217864
    cERRO = cERRO & vbCrLf & "the rowset was using optimistic concurrency and the value of a column has been changed since it was last read"
  Case -2147217863
    cERRO = cERRO & vbCrLf & "errors were detected during the copy"
  Case -2147217862
    cERRO = cERRO & vbCrLf & "a specified precision was invalid"
  Case -2147217861
    cERRO = cERRO & vbCrLf & "a specified scale was invalid"
  Case -2147217860
    cERRO = cERRO & vbCrLf & "invalid table id"
  Case -2147217859
    cERRO = cERRO & vbCrLf & "a specified type was invalid"
  Case -2147217858
    cERRO = cERRO & vbCrLf & "a column id was occurred more than once in the specification"
  Case -2147217857
    cERRO = cERRO & vbCrLf & "the specified table already exists"
  Case -2147217856
    cERRO = cERRO & vbCrLf & "the specified table was in use"
  Case -2147217855
    cERRO = cERRO & vbCrLf & "the specified locale id was not supported"
  Case -2147217854
    cERRO = cERRO & vbCrLf & "the specified record number is invalid"
  Case -2147217853
    cERRO = cERRO & vbCrLf & "although the bookmark was validly formed, no row could be found to match it"
  Case -2147217852
    cERRO = cERRO & vbCrLf & "the value of a property was invalid"
  Case -2147217851
    cERRO = cERRO & vbCrLf & "the rowset was not chaptered"
  Case -2147217850
    cERRO = cERRO & vbCrLf & "invalid accessor"
  Case -2147217849
    cERRO = cERRO & vbCrLf & "invalid storage flags"
  Case -2147217848
    cERRO = cERRO & vbCrLf & "by-ref accessors are not supported by this provider"
  Case -2147217847
    cERRO = cERRO & vbCrLf & "null accessors are not supported by this provider"
  Case -2147217846
    cERRO = cERRO & vbCrLf & "the command was not prepared"
  Case -2147217845
    cERRO = cERRO & vbCrLf & "the specified accessor was not a parameter accessor"
  Case -2147217844
    cERRO = cERRO & vbCrLf & "the given accessor was write-only"
  Case -2147217843
    cERRO = cERRO & vbCrLf & "authentication failed"
  Case -2147217842
    cERRO = cERRO & vbCrLf & "the change was canceled during notification; no columns are changed"
  Case -2147217841
    cERRO = cERRO & vbCrLf & "the rowset was single-chaptered and the chapter was not released"
  Case -2147217840
    cERRO = cERRO & vbCrLf & "invalid source handle"
  Case -2147217839
    cERRO = cERRO & vbCrLf & "the provider cannot derive parameter info and setparameterinfo has not been called"
  Case -2147217838
    cERRO = cERRO & vbCrLf & "the data source object is already initialized"
  Case -2147217837
    cERRO = cERRO & vbCrLf & "the provider does not support this method"
  Case -2147217836
    cERRO = cERRO & vbCrLf & "the number of rows with pending changes has exceeded the set limit"
  Case -2147217835
    cERRO = cERRO & vbCrLf & "the specified column did not exist"
  Case -2147217834
    cERRO = cERRO & vbCrLf & "there are pending changes on a row with a reference count of zero"
  Case -2147217833
    cERRO = cERRO & vbCrLf & "a literal value in the command overflowed the range of the type of the associated column"
  Case -2147217832
    cERRO = cERRO & vbCrLf & "the supplied hresult was invalid"
  Case -2147217831
    cERRO = cERRO & vbCrLf & "the supplied lookupid was invalid"
  Case -2147217830
    cERRO = cERRO & vbCrLf & "the supplied dynamicerrorid was invalid"
  Case -2147217829
    cERRO = cERRO & vbCrLf & "unable to get visible data for a newly-inserted row that has not yet been updated"
  Case -2147217828
    cERRO = cERRO & vbCrLf & "invalid conversion flag"
  Case -2147217827
    cERRO = cERRO & vbCrLf & "the given parameter name was unrecognized"
  Case -2147217826
    cERRO = cERRO & vbCrLf & "multiple storage objects can not be open simultaneously"
  Case -2147217825
    cERRO = cERRO & vbCrLf & "the requested filter could not be opened"
  Case -2147217824
    cERRO = cERRO & vbCrLf & "the requested order could not be opened"
  Case -2147217823
    cERRO = cERRO & vbCrLf & "bad tuple"
  Case -2147217822
    cERRO = cERRO & vbCrLf & "bad coordinate"
  Case -2147217821
    cERRO = cERRO & vbCrLf & "the given axis was not valid for this dataset"
  Case -2147217820
    cERRO = cERRO & vbCrLf & "one or more of the given cell ordinals was invalid"
  Case -2147217819
    cERRO = cERRO & vbCrLf & "the supplied columnid was invalid"
  Case -2147217817
    cERRO = cERRO & vbCrLf & "the supplied command does not have a dbid"
  Case -2147217816
    cERRO = cERRO & vbCrLf & "the supplied dbid already exists"
  Case -2147217815
    cERRO = cERRO & vbCrLf & "the maximum number of sessions supported by the provider has already been created. the consumer must release one or more currently held sessions before obtaining a new session object"
  Case -2147217806
    cERRO = cERRO & vbCrLf & "the index id is invalid"
  Case -2147217805
    cERRO = cERRO & vbCrLf & "the initialization string does not conform to specification"
  Case -2147217804
    cERRO = cERRO & vbCrLf & "the ole db root enumerator did not return any providers that matched an of the sources_types requested"
  Case -2147217803
    cERRO = cERRO & vbCrLf & "the initialization string specifies a provider which does not match the currently active provider."
  Case -2147217802
    cERRO = cERRO & vbCrLf & "the specified dbid is invalid"
  Case -2147217814
    cERRO = cERRO & vbCrLf & "invalid trustee value"
  Case -2147217813
    cERRO = cERRO & vbCrLf & "the trustee is not for the current data source"
  Case -2147217812
    cERRO = cERRO & vbCrLf & "the trustee does not support memberships/ collections"
  Case -2147217811
    cERRO = cERRO & vbCrLf & "the object is invalid or unknown to the provider"
  Case -2147217810
    cERRO = cERRO & vbCrLf & "no owner exists for the object"
  Case -2147217809
    cERRO = cERRO & vbCrLf & "the access entry list supplied is invalid"
  Case -2147217808
    cERRO = cERRO & vbCrLf & "the trustee supplied as owner is invalid or unknown to the provider"
  Case -2147217807
    cERRO = cERRO & vbCrLf & "the permission supplied in the access entry list is invalid"
  Case -2147217801
    cERRO = cERRO & vbCrLf & "the constrainttype was invalid or not supported by the provider."
  Case -2147217800
    cERRO = cERRO & vbCrLf & "the constrainttype was not constrainttype_foreignkey and cforeignkeycolumns was not zero"
  Case -2147217799
    cERRO = cERRO & vbCrLf & "the deferrability was invalid or the value was not supported by the provider"
  Case -2147217792
    cERRO = cERRO & vbCrLf & "the matchtype was invalid or the value was not supported by the provider"
  Case -2147217782
    cERRO = cERRO & vbCrLf & "the updaterule or deleterule was invalid or the value was not supported by the provider"
  Case -2147217781
    cERRO = cERRO & vbCrLf & "the pconstraintid did not exist in the data source"
  Case -2147217780
    cERRO = cERRO & vbCrLf & "the dwflags was invalid"
  Case -2147217779
    cERRO = cERRO & vbCrLf & "the rguidcolumntype pointed to a guid that does not match the object type of this column or this column was not set"
  Case -2147217778
    cERRO = cERRO & vbCrLf & "the requested url was out-of-scope"
  Case -2147217776
    cERRO = cERRO & vbCrLf & "the provider could not drop the object"
  Case -2147217775
    cERRO = cERRO & vbCrLf & "there is no source row"
  Case -2147217774
    cERRO = cERRO & vbCrLf & "the ole db object represented by this url is locked by one or more other processes"
  Case -2147217773
    cERRO = cERRO & vbCrLf & "the client requested an object type that is only valid for a collection"
  Case -2147217772
    cERRO = cERRO & vbCrLf & "the caller requested write access to a read-only object"
  Case -2147217771
    cERRO = cERRO & vbCrLf & "the provider could not connect to the server for this object"
  Case -2147217770
    cERRO = cERRO & vbCrLf & "the provider could not connect to the server for this object"
  Case -2147217769
    cERRO = cERRO & vbCrLf & "the attempt to bind to the object timed out"
  Case -2147217768
    cERRO = cERRO & vbCrLf & "the provider was unable to create an object at this url because an object named by this url already exists"
  Case -2147217767
    cERRO = cERRO & vbCrLf & "the provider could not drop the object"
  Case -2147217766
    cERRO = cERRO & vbCrLf & "the provider was unable to create an object at this url because the server was out of physical storage"
  End Select

 cERRO = cERRO & vbCrLf & infosistema

 '' Set oSVER = New clsOSInfo
 '' With oSVER
  ''  cERRO = cERRO & vbCrLf & "IP         : " & .IPAddress
  ''  cERRO = cERRO & vbCrLf & "Local      : " & .IPHostName
  ''  cERRO = cERRO & vbCrLf & "Equipamento: " & .MachineName
  ''  cERRO = cERRO & vbCrLf & "Usuario    : " & .UserName
  ''  cERRO = cERRO & vbCrLf & "OS Name    : " & .OSName
  ''  cERRO = cERRO & vbCrLf & "Service Pack ver.: " & .SPVer
  ''  cERRO = cERRO & vbCrLf & "Bitness    : " & .Bitness
  ''  cERRO = cERRO & vbCrLf & "Edition    : " & .Edition
  ''  cERRO = cERRO & vbCrLf & "Family     : " & .Family
 ''   cERRO = cERRO & vbCrLf & "Suite mask : " & .SuiteMask
''    cERRO = cERRO & vbCrLf & "ProductType: " & .ProductType
''    cERRO = cERRO & vbCrLf & "OS - Vista or newer? " & .IsVistaOrLater
''    cERRO = cERRO & vbCrLf & "Major      : " & .Major
 ''   cERRO = cERRO & vbCrLf & "Minor      : " & .Minor
 ''   cERRO = cERRO & vbCrLf & "Major + Minor: " & .MajorMinor
''    cERRO = cERRO & vbCrLf & "Build: " & .Build
''    cERRO = cERRO & vbCrLf & "ReleaseId: " & .ReleaseId
''    cERRO = cERRO & vbCrLf & "Language in dialogues: " & .LangDisplayCode & " " & .LangDisplayName & " " & .LangDisplayNameFull
''    cERRO = cERRO & vbCrLf & "Language of OS inslallation: " & .LangSystemCode & " " & .LangSystemName & " " & .LangSystemNameFull
''    cERRO = cERRO & vbCrLf & "Language for non-Unicode programs: " & .LangNonUnicodeCode & " " & .LangNonUnicodeName & " " & .LangNonUnicodeNameFull
 ''   cERRO = cERRO & vbCrLf & "Process integrity level: " & .IntegrityLevel
 ''   cERRO = cERRO & vbCrLf & "Elevated process? " & .IsElevated
 ''   cERRO = cERRO & vbCrLf & "User group: " & .UserType
 ''   cERRO = cERRO & vbCrLf & "Safe boot? " & .IsSafeBoot & " (" & .SafeBootMode & ")"
''    cERRO = cERRO & vbCrLf & "OEM Codepage: " & .CodepageOEM & " (" & .CodepageOEM_File & ")"
''    cERRO = cERRO & vbCrLf & "ANSI Codepage: " & .CodepageANSI & " (" & .CodepageANSI_File & ")"
    
 '' End With
  


  nHANDLE = FreeFile
  Open cARQ For Output As #nHANDLE
  Print #nHANDLE, cERRO
  Close (nHANDLE)

End If
End Sub

Public Function infosistema()
Dim osver As clsOSInfo
 

    Set osver = New clsOSInfo

    Dim s As String
    With osver
    
    
      ' S = S & vbCrLf & "IP         : " & .IPAddress
      ' S = S & vbCrLf & "Local      : " & .IPHostName
        s = "Equipamento: " & .ComputerName
        s = s & vbCrLf & "Usuario: " & .UserName
        s = s & vbCrLf & "Grupo do Usuario: " & .UserType
        s = s & vbCrLf & "Is in Admin group? " & .IsAdminGroup
        s = s & vbCrLf & "OS Name: " & .OSName
        s = s & vbCrLf & "Service Pack ver.: " & .SPVer
        s = s & vbCrLf & "Is Server? " & .IsServer
        s = s & vbCrLf & "Bitness: " & .Bitness
        s = s & vbCrLf & "Is Win x64: " & .IsWin64
        s = s & vbCrLf & "Is Win x32: " & .IsWin32
        s = s & vbCrLf & "Edition: " & .Edition
        s = s & vbCrLf & "Suite mask: " & .SuiteMask
        s = s & vbCrLf & "ProductType: " & .ProductType
        s = s & vbCrLf & "PlatformID: " & .PlatformID & " (" & .Platform & ")"
        s = s & vbCrLf & "Is Domain controller: " & .IsDomainController
        s = s & vbCrLf & "Is Embedded: " & .IsEmbedded
        s = s & vbCrLf & "Language in dialogues: " & .LangDisplayCode & " " & .LangDisplayName & " " & .LangDisplayNameFull
        s = s & vbCrLf & "Language of OS inslallation: " & .LangSystemCode & " " & .LangSystemName & " " & .LangSystemNameFull
        s = s & vbCrLf & "Language for non-Unicode programs: " & .LangNonUnicodeCode & " " & .LangNonUnicodeName & " " & .LangNonUnicodeNameFull
        s = s & vbCrLf & "ID of default locale: " & .LCID_UserDefault
        s = s & vbCrLf & "File System Case sensitive? " & .IsFileSystemCaseSensitive
        s = s & vbCrLf & "OEM Codepage: " & .CodepageOEM & " (" & .CodepageOEM_File & ")"
        s = s & vbCrLf & "ANSI Codepage: " & .CodepageANSI & " (" & .CodepageANSI_File & ")"

        
        s = s & vbCrLf & "Major: " & .Major
        s = s & vbCrLf & "Minor: " & .Minor
        s = s & vbCrLf & "Major + Minor:         " & .MajorMinor
        s = s & vbCrLf & "Major + Minor (NtDll): " & .MajorMinorNTDLL
        s = s & vbCrLf & "Build: " & .Build
        s = s & vbCrLf & "NT Dll Major.Minor.Rev: " & .NtDllVersion
        s = s & vbCrLf & "Revision: " & .Revision
        s = s & vbCrLf & "ReleaseId: " & .ReleaseId
        s = s & vbCrLf & "DisplayVersion: " & .DisplayVersion
        s = s & vbCrLf & "Process integrity level: " & .IntegrityLevel
        s = s & vbCrLf & "Elevated process? " & .IsElevated
        s = s & vbCrLf & "Is Local system context? " & .IsLocalSystemContext
        s = s & vbCrLf & "User sid of current process owner: " & .SID_CurrentProcess
        s = s & vbCrLf & "Safe boot? " & .IsSafeBoot & " (" & .SafeBootMode & ")"
        s = s & vbCrLf & "Secure Boot supported? " & .SecureBootSupported & " (Enabled? " & .SecureBoot & ")"
        s = s & vbCrLf & "TestSigning: " & .TestSigning
        s = s & vbCrLf & "DebugMode: " & .DebugMode
        s = s & vbCrLf & "CodeIntegrity: " & .CodeIntegrity
        s = s & vbCrLf & "Memory MiB (Free/Total): " & .MemoryFree & "/" & .MemoryTotal & " (Loaded: " & .MemoryLoad & "%)"
        
              s = s & vbCrLf & "OS - XP/Server 2003(R2)? " & .IsWindowsXP
        s = s & vbCrLf & "OS - Vista/Server 2008? " & .IsWindowsVista
        s = s & vbCrLf & "OS - 7/Server 2008R2? " & .IsWindows7
        s = s & vbCrLf & "OS - 8/Server 2012? " & .IsWindows8
        s = s & vbCrLf & "OS - 8.1/Server 2012R2? " & .IsWindows8OrGreater
        s = s & vbCrLf & "OS - 10/Server 2016? " & .IsWindows10
        s = s & vbCrLf & "OS - XP or newer? " & .IsWindowsXPOrGreater
        s = s & vbCrLf & "OS - XP SP3 or newer? " & .IsWindowsXP_SP3OrGreater
        s = s & vbCrLf & "OS - Vista or newer? " & .IsWindowsVistaOrGreater
        s = s & vbCrLf & "OS - 7 or newer? " & .IsWindows7OrGreater
        s = s & vbCrLf & "OS - 8 or newer? " & .IsWindows8OrGreater
        s = s & vbCrLf & "OS - 8.1 or newer? " & .IsWindows8Point1OrGreater
        s = s & vbCrLf & "OS - 10 or newer? " & .IsWindows10OrGreater
        s = s & vbCrLf & "OS - 11 or newer? " & .IsWindows11OrGreater
  
        'Debug.Print s
        infosistema = s
    End With
    
End Function



Public Function ShellEx( _
       ByVal sFile As String, _
       Optional ByVal eShowCmd As EShellShowConstants = essSW_SHOWDEFAULT, _
       Optional ByVal sParameters As String = "", _
       Optional ByVal sDefaultDir As String = "", _
       Optional sOperation As String = "open", _
       Optional Owner As Long = 0 _
     ) As Boolean
  Dim lr As Long
  Dim lErr As Long
  Dim sErr As String
  If (InStr(UCase$(sFile), ".EXE") <> 0) Then
    eShowCmd = 0
  End If
  On Error Resume Next
  If (sParameters = "") And (sDefaultDir = "") Then
    lr = ShellExecuteForExplore(Owner, sOperation, sFile, 0, 0, essSW_SHOWNORMAL)
  Else
    lr = ShellExecute(Owner, sOperation, sFile, sParameters, sDefaultDir, eShowCmd)
  End If
  If (lr < 0) Or (lr > 32) Then
    ShellEx = True
  Else
    ' raise an appropriate error:
    lErr = vbObjectError + 1048 + lr
    Select Case lr
    Case 0
      lErr = 7: sErr = "Sem Memoria"
  '  Case ERROR  'ERROR_FILE_NOT_FOUND
  '    lErr = 53: sErr = "Arquivo nao encontrado"
  '  Case ERROR_PATH_NOT_FOUND
  '    lErr = 76: sErr = "Caminho nao encontrado"
  '  Case ERROR_BAD_FORMAT
  '    sErr = "Executavel invalido ou corrompido"
    Case SE_ERR_ACCESSDENIED
      lErr = 75: sErr = "Caminho ou arquivo sem acesso "
    Case SE_ERR_ASSOCINCOMPLETE
      sErr = "Executavel nao associado a esta extensao."
    Case SE_ERR_DDEBUSY
      lErr = 285: sErr = "O Executavel nao pode ser aberto no momento, tente outra vez."
    Case SE_ERR_DDEFAIL
      lErr = 285: sErr = "Falha na transacao DDE, tente outra vez."
    Case SE_ERR_DDETIMEOUT
      lErr = 286: sErr = "Tempo esgotado, tente outra vez."
    Case SE_ERR_DLLNOTFOUND
      lErr = 48: sErr = "Biblioteca nao encotrada."
    Case SE_ERR_FNF
      lErr = 53: sErr = "Arquivo nao encontrado"
    Case SE_ERR_NOASSOC
      sErr = "Executavel nao associado a esta extensao."
    Case SE_ERR_OOM
      lErr = 7: sErr = "Sem memoria"
    Case SE_ERR_PNF
      lErr = 76: sErr = "Caminho nao encontrado"
    Case SE_ERR_SHARE
      lErr = 75: sErr = "Violacao de Acesso."
    Case Else
      sErr = "Erro nao especificado."
    End Select

  '  Err.Raise lErr, , App.EXEName & ".GShell", sErr
    ShellEx = False
  End If

End Function

'Public Function SomaArr(ByVal aARRAY As Variant, ByVal nITEM As Integer)
'  Dim x As Integer
'  SomaArr = 0
'  For x = 0 To nITEM - 1
'    SomaArr = SomaArr + FixNum(aARRAY(x))
'  Next x
'End Function
Public Function SomaArr(ByRef vArray As Variant)
    ' 1. Verificação de segurança: checa se é um array antes de iterar
    If Not IsArray(vArray) Then Exit Function
    
    Dim vItem As Variant
    Dim dSoma As Double
    
    ' 2. O twinBASIC e o VB6 lidam com 'For Each' em arrays de variantes
    ' de forma idêntica em 32 e 64 bits. É o método mais seguro.
    For Each vItem In vArray
        ' Usamos sua FixNum para garantir que só somamos o que é numérico
        If FixNum(vItem) Then
            dSoma = dSoma + FixNum(vItem) 'CDbl(vItem)
        End If
    Next vItem
    
    SomaArr = dSoma
End Function


Public Function SomaExt(ByVal cARQ As String, Optional ByVal cEXT As String = ".MDB") As String
  If InStr(UCase(cARQ), cEXT) > 0 Then
    SomaExt = cARQ
  Else
    SomaExt = cARQ & cEXT
  End If
End Function

Public Function StrZero(ByVal nNUM, Optional ByVal nLEN As Integer = 0)
  Dim cTemp As String
  If nLEN = 0 Then
    cTemp = FixStr(nNUM, "0", "TRIM")
    nLEN = Trim(cTemp)
  End If
  StrZero = Right(String(nLEN, "0") & CStr(nNUM), nLEN)
End Function

Function Convert2ansi(ByVal in_string As String) As String
  Dim Out_String As String * 256
  Dim t As String
  t = OemToChar(in_string, Out_String)
  Convert2ansi = Out_String
End Function
Function Convert2oem(ByVal in_string As String) As String
  Dim Out_String As String * 256
  Dim t As String
  t = CharToOem(in_string, Out_String)
  Convert2oem = Out_String
End Function
'Public Function ConvOEM(ByVal cTEXTO As String) As String
'  ConvOEM = Convert2oem(cTEXTO)
'End Function

'Public Function ConvOEM2(ByVal cTEXTO As String) As String
'  ConvOEM2 = Convert2oem(cTEXTO)
'End Function

'Public Function ConvAnsi2(ByVal cTEXTO As String) As String
'  ConvAnsi2 = Convert2ansi(cTEXTO)
'End Function

'Public Function ConvAnsi(ByVal texto As String) As String
'  ConvAnsi = Convert2ansi(texto)
'End Function
Public Function Tirace(ByVal texto As String) As String
  Tirace = tirace2(texto)
End Function
Public Function StrToArray(ByVal cGRUPO As String) As Variant
  Dim x As Integer
  Dim nLEN As Integer
  Dim aUSO As Variant
  Dim cCHAR As String
  Dim eCNV As String
  ' utilizando agora Convert2ansi Convert2oem tirace2
  Select Case cGRUPO
  Case Else
    eCNV = cGRUPO
  End Select

  nLEN = Len(eCNV)
  ReDim aUSO(nLEN)
  For x = 1 To nLEN
    cCHAR = Mid(eCNV, x, 1)
    Select Case cCHAR
    Case "ª"
      aUSO(x - 1) = "a."

    Case "º"
      aUSO(x - 1) = "o."

    Case Else
      aUSO(x - 1) = cCHAR

    End Select
  Next x
  StrToArray = aUSO
End Function

Public Function TiraSin(ByVal texto As String, Optional ByVal RemoveUp As Boolean = True) As String
Dim x As Integer
  'https://www.ascii-code.com/pt
  'Caracteres de controle ASCII (código de caractere 0-31)
  For x = 0 To 31 'nao caracteres padrao
    texto = Replace(texto, Chr(x), "")
  Next x
  'Caracteres imprimiveis ASCII (código de caractere 32-127)
  '32 espaco
  For x = 33 To 38                 '33! 34" 35# 36$ 37% 38&
    texto = Replace(texto, Chr(x), "")
  Next x
  For x = 39 To 47                 '39' 40( 41) 42* 43+ 44, 45- 46. 47/
    texto = Replace(texto, Chr(x), "")
  Next x
  '48 a 57 numeros 0-9
  For x = 58 To 64                 '58: 59: 60< 61= 62> 63? 64@
    texto = Replace(texto, Chr(x), "")
  Next x
  '65 a 90 maisculas
  For x = 91 To 96                 '91[ 92\ 93] 94^ 95_ 96`
    texto = Replace(texto, Chr(x), "")
  Next x
  '97 a 122 minusculas
  For x = 123 To 126                '123{ 124| 125} 126~
    texto = Replace(texto, Chr(x), "")
  Next x
  For x = 127 To 127                 '127 del
    texto = Replace(texto, Chr(x), "")
  Next x
  'Códigos ASCII estendidos (código de caracteres 128-255)
  If RemoveUp Then
    For x = 128 To 255
      texto = Replace(texto, Chr(x), "")
    Next x
  Else 'matem caracteres de acentuacao uso tirace caso queira manter sem acentuacao
    For x = 155 To 159
      texto = Replace(texto, Chr(x), "")
    Next x
    For x = 168 To 180
      texto = Replace(texto, Chr(x), "")
    Next x
    For x = 184 To 197
      texto = Replace(texto, Chr(x), "")
    Next x
    For x = 200 To 209
      texto = Replace(texto, Chr(x), "")
    Next x
    For x = 217 To 223
      texto = Replace(texto, Chr(x), "")
    Next x
    For x = 238 To 255
      texto = Replace(texto, Chr(x), "")
    Next x
  End If
  TiraSin = texto
End Function
Function CheckPass(ByVal cTEXTO As String, Optional ByVal lMES As Boolean = True) As Boolean

  Dim lMAIS As Boolean
  Dim lMINUS As Boolean
  Dim lDIG As Boolean
  Dim lSYMBOL As Boolean
  Dim l8DIG As Boolean
  Dim i As Integer

  CheckPass = False
  lMAIS = False
  lMINUS = False
  lDIG = False
  lSYMBOL = False
  l8DIG = False


  For i = 1 To Len(cTEXTO)
    If InStr("0123456789", Mid(cTEXTO, i, 1)) > 0 Then
      lDIG = True
    End If
    If InStr("abcdefghijklmnopqrstuvwxyz", Mid(cTEXTO, i, 1)) > 0 Then
      lMINUS = True
    End If
    If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", Mid(cTEXTO, i, 1)) > 0 Then
      lMAIS = True
    End If
    If InStr("-+_!@#$%^&*., ?", Mid(cTEXTO, i, 1)) > 0 Then
      lSYMBOL = True
    End If
  Next

  If Len(Trim(cTEXTO)) >= 8 Then
    l8DIG = True
  End If


  If lMAIS And lMINUS And lDIG And lSYMBOL And l8DIG Then
    CheckPass = True
  Else
    If lMES Then
      If Not lMAIS Then
        Alert (" Sem maisculas")
      End If
      If Not lMINUS Then
        Alert (" Sem minusclas")
      End If
      If Not lDIG Then
        Alert (" Sem numeros")
      End If
      If Not lSYMBOL Then
        Alert (" Sem simbulos")
      End If
      If Not l8DIG Then
        Alert ("Minimo 8 Digitos")
      End If
    End If
  End If

End Function

Function isnumber(ByVal cLETRA As String) As Boolean
   isnumber = InStr("0123456789", cLETRA) > 0
End Function

Function isstringupper(ByVal cLETRA As String) As Boolean
   isstringupper = InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", cLETRA) > 0
End Function

Function isstringlower(ByVal cLETRA As String) As Boolean
   isstringlower = InStr("abcdefghijklmnopqrstuvwxyz", cLETRA) > 0
End Function

Function issimbolo(ByVal cLETRA As String) As Boolean
   issimbolo = InStr(",.:;$@&*!?<>%+-*/\|#=:_{}[]()'^~´`", cLETRA) > 0
End Function

Function isstring(ByVal cLETRA As String) As Boolean
    isstring = isstringlower(cLETRA) Or isstringupper(cLETRA)
End Function
Public Function tirace2(ByVal cXml As String) As String
  Dim nAscii As Integer
  Dim ltroca As Boolean
  Dim nCont As Integer
  Dim cLETRA As String

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



Public Function XOREncryption(strCodeKey As String, _
                              strDataIn As String) As String

  Dim lonDataPtr As Long
  Dim intXORValue1 As Integer
  Dim intXORValue2 As Integer
  Dim strDataOut As String
  Dim nLENINI As Integer
  Dim cTMPCHAR As String


  strDataOut = ""
  nLENINI = Len(strDataIn)


  For lonDataPtr = 1 To Len(strDataIn)
    intXORValue1 = Asc(Mid$(strDataIn, lonDataPtr, 1))
    intXORValue2 = Asc(Mid$(strCodeKey, _
                            ((lonDataPtr Mod Len(strCodeKey)) + 1), 1))
    cTMPCHAR = Chr(intXORValue1 Xor intXORValue2)
    If Asc(cTMPCHAR) = 0 Or Asc(cTMPCHAR) = 32 Then
      cTMPCHAR = "_"
    End If
    strDataOut = strDataOut + cTMPCHAR

  Next lonDataPtr



  XOREncryption = Mid$(strDataOut, 1, nLENINI)

End Function

Public Sub FocusMe()
  If TypeOf Screen.ActiveControl Is TextBox _
     Or TypeOf Screen.ActiveControl Is ComboBox _
     Or TypeOf Screen.ActiveControl Is XPText Then
    Screen.ActiveControl.SelStart = 0
    Screen.ActiveControl.SelLength = Len(Trim(Screen.ActiveControl.tEXT))
  End If
End Sub
Public Function CharConv(ByVal cTEXTO As String, ByVal eORI As Variant, ByVal eDES As Variant) As String
  Dim nLEN As Integer
  Dim nTEXTO As Integer
  Dim x As Integer
  Dim y As Integer
  Dim aORI As Variant
  Dim aDES As Variant
  Dim aTEXTO As Variant
  If IsArray(eORI) Then
    aORI = eORI
    aDES = eDES
  Else
    aORI = StrToArray(CStr(eORI))
    aDES = StrToArray(CStr(eDES))
  End If
  aTEXTO = StrToArray(cTEXTO)
  nLEN = UBound(aORI)
  nTEXTO = UBound(aTEXTO)
  For y = 0 To nTEXTO
    For x = 0 To nLEN
      If aTEXTO(y) = aORI(x) Then          ''Encerra Analise Para Evitar
        aTEXTO(y) = aDES(x)              ''Loop de Troca
        Exit For
      End If
    Next
  Next
  CharConv = ""
  For y = 0 To nTEXTO
    CharConv = CharConv & aTEXTO(y)
  Next y
End Function
Public Function TiraOut(ByVal eVAR As Variant) As String
  Dim cTEXTO As String
  cTEXTO = FixStr(eVAR)
  TiraOut = CharConv(cTEXTO, Array("-", ",", ".", ":", "/", ";", "*", "(", ")"), _
                     Array("", "", "", "", "", "", "", "", ""))
End Function

Public Function TiraOutNum(ByVal eVAR As Variant) As String  ' Mantem ', . que sao usadas nos numeros
  Dim cTEXTO As String
  cTEXTO = FixStr(eVAR)
  TiraOutNum = CharConv(cTEXTO, Array("-", ":", "/", ";", "*", "(", ")"), _
                        Array("", "", "", "", "", "", ""))
End Function

Public Function TiraOutAlf(ByVal eVAR As Variant) As String
  Dim cTEXTO As String
  cTEXTO = FixStr(eVAR)
  TiraOutAlf = CharConv(cTEXTO, Array("-", ".", ":", "/", ";", "*", "(", ")", _
                                      "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", _
                                      "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", _
                                      "U", "V", "Y", "W", "X", "Z", ","), _
                                      Array("", "", "", "", "", "", "", "", _
                                            "", "", "", "", "", "", "", "", "", "", _
                                            "", "", "", "", "", "", "", "", "", "", _
                                            "", "", "", "", "", "", ""))
End Function

Public Function Extenso(ByVal Valor As Double, _
                        Optional ByVal MoedaPlural As String = "Reais", _
                        Optional ByVal MoedaSingular As String = "Real") As String
  Dim StrValor As String
  Dim Negativo As Boolean
  Dim buf As String
  Dim Parcial As Integer
  Dim Posicao As Integer
  Dim Unidades As Variant
  Dim Dezenas As Variant
  Dim Centenas As Variant
  Dim PotenciasSingular As Variant
  Dim PotenciasPlural As Variant

  Negativo = (Valor < 0)
  Valor = Abs(CDec(Valor))
  If Valor Then
    Unidades = Array(vbNullString, "Um", "Dois", _
                     "Tres", "Quatro", "Cinco", _
                     "Seis", "Sete", "Oito", "Nove", _
                     "Dez", "Onze", "Doze", "Treze", _
                     "Quatorze", "Quinze", "Dezesseis", _
                     "Dezessete", "Dezoito", "Dezenove")
    Dezenas = Array(vbNullString, vbNullString, _
                    "Vinte", "Trinta", "Quarenta", _
                    "Cinquenta", "Sessenta", "Setenta", _
                    "Oitenta", "Noventa")
    Centenas = Array(vbNullString, "Cento", _
                     "Duzentos", "Trezentos", _
                     "Quatrocentos", "Quinhentos", _
                     "Seiscentos", "Setecentos", _
                     "Oitocentos", "Novecentos")
    PotenciasSingular = Array(vbNullString, " Mil", _
                            " Milhao", " Bilhao", _
                            " Trilhao", " Quatrilhao")
    PotenciasPlural = Array(vbNullString, " Mil", _
                          " Milhoes", " Bilhoes", _
                          " Trilhoes", " Quatrilhoes")

    StrValor = Left(Format(Valor, String(18, "0") & _
                                  ".000"), 18)
    For Posicao = 1 To 18 Step 3
      Parcial = Val(Mid(StrValor, Posicao, 3))
      If Parcial Then
        If Parcial = 1 Then
          buf = "Um" & PotenciasSingular((18 - _
                                          Posicao) \ 3)
        ElseIf Parcial = 100 Then
          buf = "Cem" & PotenciasSingular((18 - _
                                           Posicao) \ 3)
        Else
          buf = Centenas(Parcial \ 100)
          Parcial = Parcial Mod 100
          If Parcial <> 0 And buf <> vbNullString Then
            buf = buf & " e "
          End If
          If Parcial < 20 Then
            buf = buf & Unidades(Parcial)
          Else
            buf = buf & Dezenas(Parcial \ 10)
            Parcial = Parcial Mod 10
            If Parcial <> 0 And buf <> vbNullString Then
              buf = buf & " e "
            End If
            buf = buf & Unidades(Parcial)
          End If
          buf = buf & PotenciasPlural((18 - Posicao) \ 3)
        End If
        If buf <> vbNullString Then
          If Extenso <> vbNullString Then
            Parcial = Val(Mid(StrValor, Posicao, 3))
            If Posicao = 16 And (Parcial < 100 Or _
                                 (Parcial Mod 100) = 0) Then
              Extenso = Extenso & " e "
            Else
              Extenso = Extenso & ", "
            End If
          End If
          Extenso = Extenso & buf
        End If
      End If
    Next
    If Extenso <> vbNullString Then
      If Negativo Then
        Extenso = "Menos " & Extenso
      End If
      If Int(Valor) = 1 Then
        Extenso = Extenso & " " & MoedaSingular
      Else
        Extenso = Extenso & " " & MoedaPlural
      End If
    End If
    Parcial = Int((Valor - Int(Valor)) * _
                  100 + 0.1)
    If Parcial Then
      buf = Extenso(Parcial, "Centavos", _
                    "Centavo")
      If Extenso <> vbNullString Then
        Extenso = Extenso & " e "
      End If
      Extenso = Extenso & buf
    End If
  End If
End Function

Public Function Txt2Lin(ByVal cTEXTO As String, Optional ByVal nCOL As Integer = 80) As Variant
  Dim nLIN, x As Integer
  Dim aRETU As Variant
  cTEXTO = FixStr(cTEXTO)
  If nCOL < 1 Then nCOL = 80                   'Evita Erros Divisao
  nLIN = Int(Len(cTEXTO) / nCOL)
  If nLIN * nCOL = Len(cTEXTO) Then            ''Bate com o Multiplo
  Else
    nLIN = nLIN + 1                          ''Soma mais um pois e necesario
  End If
  ReDim aRETU(nLIN)
  For x = 1 To nLIN
    aRETU(x - 1) = Mid(cTEXTO, ((x - 1) * nCOL) + 1, nCOL)
  Next x
  Txt2Lin = aRETU
  eRETU01 = nLIN
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
Public Sub EnableControls(ByVal frm As Form, ByVal enabled_state As Boolean)
  Dim ctl As Control
  ' Examine every control.
  For Each ctl In frm.Controls
    On Error Resume Next
    ctl.Enabled = enabled_state
    On Error GoTo 0
  Next ctl
End Sub


Public Function FixBolNum(ByVal lbol As Boolean) As Integer  'variant
  FixBolNum = 0
  If lbol Then
    FixBolNum = 1
  End If
End Function

Public Function FixNumBol(ByVal nBOL As Integer) As Boolean
  FixNumBol = False
  If nBOL = 1 Then
    FixNumBol = True
  End If
End Function

Public Function FixStrBol(ByVal cCHAR As String) As Boolean
 FixStrBol = StrLogic(cCHAR)
End Function

' +--------------------------------------------------------------------
' +  Função: StrLogic
' +  Objetivo: Converte uma string alfanumérica em valor lógico booleano,
' +            espelhando a mesma inteligência usada no Harbour.
' +  Retorno:  Variant (Pode retornar True, False, ou Null para nulos)
' +--------------------------------------------------------------------
Public Function StrLogic(ByVal cVAL As String) As Variant
    ' Limpa espaços em branco e padroniza em maiúsculo
    cVAL = UCase$(Trim$(cVAL))
    
    Select Case cVAL
        ' Casos que retornam Nulo/Vazio
        Case "", "NULL", "<NULL>", "NIL", "NUL"
            StrLogic = Null
            
        ' Casos que retornam VERDADEIRO (.T.)
        Case "S", "SIM", "T", ".T.", "TRUE", "YES", "ON", "Y", "1"
            StrLogic = True
            
        ' Casos que retornam FALSO (.F.)
        Case "N", "NAO", "F", ".F.", "FALSE", "NO", "OFF", "0"
            StrLogic = False
            
        ' Fallback de segurança (Caso não combine com nenhum)
        Case Else
            StrLogic = False
    End Select
End Function
Public Function Count_Lines_In_File(ByVal strFilePath As String, Optional ByVal lMES As Boolean = True) As Long
    Dim fso As Object
    Dim Stream As Object
    
    ' Mantém a sua validação original usando o seu FileConnExist
    If Not FileConnExist(strFilePath, lMES) Then
        Count_Lines_In_File = -1
        Exit Function
    End If
    
    On Error GoTo TrataErro
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Abre o arquivo em modo leitura (1 = ForReading)
    Set Stream = fso.OpenTextFile(strFilePath, 1, False)
    
    ' TRUQUE DE ALTA PERFORMANCE DO FSO:
    ' Em vez de ler linha por linha em um loop, o .SkipLine pula as linhas na velocidade do Kernel.
    ' .ReadAll joga o ponteiro direto para o fim do arquivo lendo o buffer de uma vez só.
    Stream.ReadAll
    
    ' A propriedade .Line nos dá exatamente o número total de linhas do arquivo
    Count_Lines_In_File = Stream.Line
    
Fim:
    If Not Stream Is Nothing Then Stream.Close
    Set Stream = Nothing
    Set fso = Nothing
    Exit Function

TrataErro:
    Count_Lines_In_File = 0
    Resume Fim
End Function
'Public Function Count_Lines_In_File_old(ByVal strFilePath As String, Optional ByVal lMES As Boolean = True) As Long

'     'delcare variables
'  Dim fileFile As Integer
'  Dim intLinesReadCount As Integer
'  Dim STRBUFFER As String

'  intLinesReadCount = 0
  '     'open file
'  fileFile = FreeFile


 ' If Not FileConnExist(strFilePath, lMES) Then
 '   Count_Lines_In_File = -1
 '   Exit Function
 ' End If

  'loop through file
  'Open strFilePath For Input As #fileFile


  'Do While Not EOF(fileFile)
    'read line
   ' Input #fileFile, STRBUFFER
    'update count
    'intLinesReadCount = intLinesReadCount + 1
  'Loop
  'close file
  'Close fileFile
  'return value
  'Count_Lines_In_File = intLinesReadCount
'End Function
' +--------------------------------------------------------------------
' +  Função: Count_Lines_In_File
' +  Objetivo: Conta linhas em arquivos Gigantescos usando a técnica do
' +            BLOCO ELÁSTICO (Recuo de ponteiro de leitura na emenda).
' +            100% à prova de falhas para os 3 cenários (CR, LF, CRLF).
' +  Retorno:  Long
' +--------------------------------------------------------------------
Public Function Count_Lines_In_Filev2(ByVal cCaminhoArquivo As String) As Long
    On Error GoTo TrataErro
    
    Dim nFile As Integer
    Dim nTamArquivo As Long
    Dim nTamBloco As Long
    Dim abBuffer() As Byte
    
    Dim i As Long
    Dim nLINHAS As Long
    Dim nRestante As Long
    Dim nBytesLidos As Long
    Dim nPonteiroAtual As Long
    Dim bUltimoByteArquivo As Byte
    
    ' Se o arquivo não existir ou estiver vazio
    If Dir(cCaminhoArquivo) = "" Then
        Count_Lines_In_Filev2 = 0
        Exit Function
    End If
    
    nTamArquivo = FileLen(cCaminhoArquivo)
    If nTamArquivo = 0 Then
        Count_Lines_In_Filev2 = 0
        Exit Function
    End If
    
    ' 1 MB de Buffer estático na RAM
    nTamBloco = 1024 * 1024
    ReDim abBuffer(0 To nTamBloco - 1)
    
    nFile = FreeFile
    Open cCaminhoArquivo For Binary Access Read As #nFile
    
    nLINHAS = 0
    nPonteiroAtual = 1 ' No VB6, o Seek do modo Binary começa em 1
    
    Do While nPonteiroAtual <= nTamArquivo
        ' Calcula quantos bytes faltam ler
        nRestante = (nTamArquivo - nPonteiroAtual) + 1
        
        ' Ajusta o tamanho do bloco se for o final do arquivo
        If nRestante < nTamBloco Then
            nTamBloco = nRestante
            ReDim abBuffer(0 To nTamBloco - 1)
        End If
        
        ' Posiciona o ponteiro do arquivo e faz a leitura
        Seek #nFile, nPonteiroAtual
        Get #nFile, , abBuffer
        nBytesLidos = nTamBloco
        
        ' Se for o último pedaço do arquivo, guarda o último byte real
        If (nPonteiroAtual + nBytesLidos - 1) = nTamArquivo Then
            bUltimoByteArquivo = abBuffer(nBytesLidos - 1)
        End If
        
        ' --- AQUI ENTRA A LÓGICA DO BLOCO ELÁSTICO ---
        ' Se o bloco NÃO for o final do arquivo, vamos analisar onde ele cortou
        If (nPonteiroAtual + nBytesLidos - 1) < nTamArquivo Then
            
            ' Caso A: O bloco terminou exatamente em Chr(13) (Pode ser um CRLF cortado)
            If abBuffer(nBytesLidos - 1) = 13 Then
                ' Encolhe o bloco logicamente em 1 byte (ignora o 13 neste turno)
                nBytesLidos = nBytesLidos - 1
                
            ' Caso B: O bloco terminou em Chr(10) e o anterior a ele era Chr(13)
            ' (O par CRLF ficou inteiro dentro do bloco, então está tudo bem)
            ElseIf abBuffer(nBytesLidos - 1) = 10 Then
                If abBuffer(nBytesLidos - 2) = 13 Then
                    ' Par CRLF fechado com sucesso, não mexe em nada.
                Else
                    ' É um LF isolado (Unix), pode processar normal.
                End If
            End If
        End If
        
        ' Varre o bloco ajustado (Sem risco de pegar caractere partido)
        For i = 0 To nBytesLidos - 1
            If abBuffer(i) = 13 Then
                nLINHAS = nLINHAS + 1
            ElseIf abBuffer(i) = 10 Then
                ' Se o anterior imediato foi 13, faz parte do par. Não conta.
                ' Caso contrário, é uma quebra pura Unix (10), conta nova linha.
                If i > 0 Then
                    If abBuffer(i - 1) <> 13 Then
                        nLINHAS = nLINHAS + 1
                    End If
                Else
                    ' Se for o primeiro byte do bloco e for 10, como o bloco elástico
                    ' evitou deixar o 13 no bloco anterior, este 10 é uma linha pura Unix.
                    nLINHAS = nLINHAS + 1
                End If
            End If
        Next i
        
        ' Atualiza o ponteiro real com base nos bytes que nós REALMENTE processamos
        nPonteiroAtual = nPonteiroAtual + nBytesLidos
    Loop
    
    Close #nFile
    nFile = 0
    
    ' Ajuste final da última linha (Caso o arquivo não termine com quebra)
    If nLINHAS > 0 Then
        If bUltimoByteArquivo <> 10 And bUltimoByteArquivo <> 13 Then
            nLINHAS = nLINHAS + 1
        End If
    Else
        If nTamArquivo > 0 Then nLINHAS = 1
    End If
    
    Count_Lines_In_Filev2 = nLINHAS
    Exit Function

TrataErro:
    If nFile > 0 Then Close #nFile
    Count_Lines_In_Filev2 = 0
End Function
'Public Sub OpenUrl(ByVal strURL As String)
'  ShellExecute 0, "Open", strURL, 0&, 0&, SW_SHOWNORMAL
'End Sub

Public Sub OpenUrl(ByVal strURL As String)
    ' 1. Valida se começa com http/https, se não, pode não abrir
    If InStr(1, strURL, "http", vbTextCompare) = 0 Then strURL = "http://" & strURL
    
    ' 2. Executa a abertura no navegador padrão do sistema
    ShellExecute 0&, "Open", strURL, vbNullString, vbNullString, 1 ' SW_SHOWNORMAL
End Sub

Public Function txttopdf(ByVal cOrigem As String, Optional ByVal cDestino As String = "", Optional ByVal cTITULO As String = "", Optional ByVal cAUTOR As String = "") As Boolean
    Dim fso As Object
    Dim streamIn As Object
    Dim cLINHA As String
    Dim pdf As ClsFPDF
    
    ' Validações estruturais padrões do seu sistema
    If Not FileConnExist(cOrigem, True) Then Exit Function
    If Len(cDestino) = 0 Then cDestino = TrocaExt(cOrigem, "PDF")
    
    If FileConnExist(cDestino, False) Then
        Alert ("Arquivo Destino Ja existe")
        Exit Function
    End If
    
    On Error GoTo TrataErro
    
    ' 1. Abre o arquivo de texto em modo Fluxo contínuo (Segurança contra arquivos gigantes)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set streamIn = fso.OpenTextFile(cOrigem, 1, False)
    
    ' 2. Inicializa o motor do FPDF
    Set pdf = New ClsFPDF
    pdf.CreatePDF "P", "mm"
    pdf.AddPage
    pdf.SetFont "Arial", 10
    
    ' 3. Configura a paginação e margens (em milímetros)
    Dim nMargemEsquerda As Single: nMargemEsquerda = 15
    Dim nLinhaAtual As Single: nLinhaAtual = 15
    Dim nAlturaLinha As Single: nAlturaLinha = 5 ' Espaçamento de 5mm entre as linhas
    Dim nLimiteInferior As Single: nLimiteInferior = 275 ' Altura máxima da folha A4 antes de quebrar a página
    
    ' 4. Loop de streaming de performance: lê e grava em tempo real
    Do While Not streamIn.AtEndOfStream
        cLINHA = streamIn.ReadLine
        
        ' Desenha a linha atual de texto no documento PDF
        pdf.Cell cLINHA, nMargemEsquerda, nLinhaAtual
        
        ' Avança o cursor vertical para a próxima linha
        nLinhaAtual = nLinhaAtual + nAlturaLinha
        
        ' QUEBRA DE PÁGINA AUTOMÁTICA: Se o texto atingir o rodapé da folha, abre uma página nova
        If nLinhaAtual > nLimiteInferior Then
            pdf.AddPage
            pdf.SetFont "Arial", 10 ' Reaplica a fonte na nova página
            nLinhaAtual = 15        ' Reinicia o cursor no topo da nova folha
        End If
    Loop
    
    ' 5. Compila e descarrega o arquivo final no disco
    pdf.Output cDestino
    txttopdf = True

Fim:
    ' Destrói os ponteiros e liberta totalmente a memória RAM do computador
    If Not streamIn Is Nothing Then streamIn.Close
    Set streamIn = Nothing
    Set fso = Nothing
    Set pdf = Nothing
    Exit Function

TrataErro:
    Alert "Erro na rotina de geração do PDF: " & Err.Description
    Resume Fim
End Function
'Public Function txttohtml_old(ByVal cORIGEM As String, Optional ByVal cDESTINO As String = "", Optional ByVal cTITULO As String = "", Optional ByVal cAUTOR As String = "")
'  Dim nORIGEM As Integer
'  Dim nDESTINO As Integer
'  Dim STRBUFFER
'  Dim cLINHA As String
'  If Not FileConnExist(cORIGEM, True) Then
'    Exit Function
'  End If
'  If Len(cDESTINO) = 0 Then
'    cDESTINO = TrocaExt(cORIGEM, "HTML")
'  End If
'  If FileConnExist(cDESTINO, False) Then
'    Alert ("Arquivo Destino Ja existe")
'    Exit Function
'  End If
'  If Len(cAUTOR) = 0 Then
'    cAUTOR = zNOMEFOLHA
'  End If
'  If Len(cTITULO) = 0 Then
'    cTITULO = NomeArq(cORIGEM, True)
'  End If
'  nORIGEM = FreeFile
'  nDESTINO = FreeFile + 1
'  Open cDESTINO For Output As #nDESTINO
'  Open cORIGEM For Input As #nORIGEM
'  Print #nDESTINO, "<html>" & Chr(13) & Chr(10)
'  Print #nDESTINO, "<head>" & Chr(13) & Chr(10)
'  Print #nDESTINO, "<meta http-equiv=" & Chr(34) & "Content-Type" & Chr(34) & Chr(13) & Chr(10)
'  Print #nDESTINO, "content=" & Chr(34) & "text/html; charset=iso-8859-1" & Chr(34) & ">" & Chr(13) & Chr(10)
'  Print #nDESTINO, "<meta name=" & Chr(34) & "GENERATOR" & Chr(34) & "content=" & Chr(34) & cAUTOR & Chr(34) & ">" & Chr(13) & Chr(10)
'  Print #nDESTINO, "<title>" & cTITULO & "</title>" & Chr(13) & Chr(10)
'  Print #nDESTINO, "</head>" & Chr(13) & Chr(10)
'  Print #nDESTINO, "<body>" & Chr(13) & Chr(10)
'  Do While Not EOF(nORIGEM)
'    Input #nORIGEM, STRBUFFER
'    cLINHA = STRBUFFER
'    cLINHA = str2html(cLINHA)
'    Print #nDESTINO, cLINHA & "<BR>"
'  Loop
'  Print #nDESTINO, "</body>" & Chr(13) & Chr(10) & "</html>" & Chr(13) & Chr(10)
'  Close nORIGEM
'  Close nDESTINO
'End Function

'Public Function txttoRTF_old(ByVal cORIGEM As String, Optional ByVal cDESTINO As String = "")
'  Dim nORIGEM As Integer
'  Dim nDESTINO As Integer
'  Dim STRBUFFER
'  Dim cLINHA As String
'  If Not FileConnExist(cORIGEM, True) Then
'    Exit Function
'  End If
'  If Len(cDESTINO) = 0 Then
'    cDESTINO = TrocaExt(cORIGEM, "RTF")
'  End If
'  If FileConnExist(cDESTINO, False) Then
'    Alert ("Arquivo Destino Ja existe")
'    Exit Function
'  End If
'  nORIGEM = FreeFile
'  nDESTINO = FreeFile + 1
'  Open cDESTINO For Output As #nDESTINO
'  Open cORIGEM For Input As #nORIGEM
'  Print #nDESTINO, "{\rtf1\ansi\ansicpg1252\deff0\deflang1031{\fonttbl{\f0\fnil\fcharset0 "
'  Print #nDESTINO, "Arial;}}\viewkind4\uc1\pard\f0\fs20 "
'  Do While Not EOF(nORIGEM)
'    Input #nORIGEM, STRBUFFER
'    cLINHA = STRBUFFER
'    Print #nDESTINO, cLINHA & "\par" & Chr(13)
'  Loop
'  Print #nDESTINO, "\par}"
'  Close nORIGEM
'  Close nDESTINO
'End Function
Public Function CharCodesToHTML(ByVal iString As String) As String
  Dim iXml As New MSXML2.DOMDocument60

  With iXml.createTextNode(iString)
    CharCodesToHTML = .Xml
  End With
End Function

Public Function HTMLToCharCodes(ByVal iString As String) As String
  With New MSXML2.DOMDocument60
    .loadXML "<p>" & iString & "</p>"
    HTMLToCharCodes = .selectSingleNode("p").nodeTypedValue
  End With
End Function
Public Function str2html(ByVal cTEXTO As String, Optional ByVal lAnsi As Boolean = False) As String
  If lAnsi Then
    Convert2ansi (cTEXTO)
  End If
  cTEXTO = CharCodesToHTML(cTEXTO)
  str2html = cTEXTO
End Function
Function FileText(ByVal filename As String) As String
  Dim Handle As Integer
  Handle = FreeFile
  Open filename$ For Input As #Handle
  FileText = Input$(LOF(Handle), Handle)
  Close #Handle
End Function
Public Function Html2Str(ByVal cTEXTO As String) As String
  Html2Str = HtmlToText(cTEXTO)  'HTMLToCharCodes(cTexto)
End Function
Public Function HtmlToText(sHTML) As String
  Dim oDoc As HTMLDocument
  Set oDoc = New HTMLDocument
  oDoc.Body.innerHTML = sHTML
  HtmlToText = oDoc.Body.innerText
End Function
Public Function FindInList(ByRef cList As ListBox, sSearch As String) As Long
  Dim sString As String
  Dim ID As Integer

  On Error Resume Next
  Err.Clear
  FindInList = -1

  For ID = 0 To cList.ListCount - 1
    sString = UCase$(cList.List(ID))
    If sString = UCase$(sSearch) Then
      FindInList = ID
      Exit For
    End If
  Next

End Function

Function TiraEspaco(sNome As String) As String
  Const sEspaco As String = " "
  Dim CaracAtual As String
  Dim NomeVerificado As String
  Dim NomeSemEspaco As String
  Dim i As Integer

  For i = 1 To Len(sNome)
    NomeVerificado = Mid(sNome, 1, i)
    CaracAtual = Mid(NomeVerificado, i, 1)   'verificando os caracteres de dois em dois
    If CaracAtual = sEspaco Then
      NomeSemEspaco = Trim$(NomeSemEspaco) & sEspaco
    Else
      NomeSemEspaco = NomeSemEspaco & CaracAtual
    End If
  Next
  TiraEspaco = Trim$(NomeSemEspaco)
End Function

Function FastArraySearch(SearchArray As Variant, SearchPhrase As String) As Long  'String
  Dim Pos As Long
  Dim i As Long
  Dim NumCharsProcessed As Long
 ' dim TXT As String
  FastArraySearch = -1
  Pos = InStr(Join(SearchArray, "§"), SearchPhrase)
  If Pos > 0 Then
    For i = LBound(SearchArray) To UBound(SearchArray)
      NumCharsProcessed = NumCharsProcessed + Len(SearchArray(i)) + 1
      If NumCharsProcessed >= Pos Then
        FastArraySearch = i              ''SearchArray(i)
        Exit Function
      End If
    Next i
  End If
End Function



Public Function NetworkUserName() As String
  Dim iStringLength As Long
  Dim i As Long
  Dim sString As String
  Dim oShell As Object
    On Error Resume Next

NetworkUserName = ""
  sString = String(255, 0)

  iStringLength = Len(sString)
  sString = Space$(iStringLength)

  If WinAPI_GetUserName(sString, CLng(iStringLength)) Then

    i = InStr(sString, Chr(0))
    If i Then
      sString = Left(sString, i - 1)
    End If
    NetworkUserName = Trim(Left$(sString, iStringLength))
  Else
   Set oShell = CreateObject("WScript.Shell")
    If Not oShell Is Nothing Then
        NetworkUserName = oShell.ExpandEnvironmentStrings("%USERNAME%")
    Else
        NetworkUserName = Environ$("USERNAME") ' Fallback simples
    End If
  
  End If

  If NetworkUserName = "" Then
    NetworkUserName = "Usuario"
  End If
  

Set oShell = Nothing
End Function
Public Function MachineName() As String
    ' Usamos o CreateObject global do VBA/twinBASIC, não do RC6.
    ' Isso é compatível com 32 e 64 bits e não depende de DLLs ou TypeLibs.
    Dim oShell As Object
    On Error Resume Next
    Set oShell = CreateObject("WScript.Shell")
    If Not oShell Is Nothing Then
        MachineName = oShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
    Else
        MachineName = Environ$("COMPUTERNAME") ' Fallback simples
    End If
    Set oShell = Nothing
End Function
Public Function WordLen(ByRef tEXT As String) As Long
'tamanho somente dos caracteres normal 65 a 90
  Dim Bytes() As Byte
  Dim i As Long

  Bytes = StrConv(UCase$(tEXT), vbFromUnicode)
  For i = 0 To UBound(Bytes)
    If 65 <= Bytes(i) And Bytes(i) <= 90 Then WordLen = WordLen + 1
  Next
End Function

Private Function ExactlyIn(ByRef Searched As String, ByRef Pattern As String, Optional ByRef Bracket As String = vbNullChar) As Boolean
'Bracket is some character that never occurs in the other arguments.
'pesquisa se tem a string com a mesma quantidade de caracteres da busca
'    MyString = "aaaareeeezzzzz"
'    Print ExactlyIn(MyString, "zz")
'    Print ExactlyIn(MyString, "zzzz")
'    Print ExactlyIn(MyString, "zzzzz")
'    Print ExactlyIn(MyString, "zzzzzz")
  ExactlyIn = Bracket & Searched & Bracket _
              Like "*[!" & Left$(Pattern, 1) & "]" _
            & Pattern _
            & "[!" & Right$(Pattern, 1) & "]*"
End Function

Public Function SameWords(ByRef Text1 As String, ByRef Text2 As String) As Boolean
'compara somente dos caracteres normal 65 a 90
  Dim Bytes() As Byte
  Dim LetterCounts() As Byte
  Dim LetterCountsS1 As String
  Dim LetterCountsS2 As String
  Dim i As Long

  ReDim LetterCounts(65 To 90)
  Bytes = StrConv(UCase$(Text1), vbFromUnicode)
  For i = 0 To UBound(Bytes)
    LetterCounts(Bytes(i)) = LetterCounts(Bytes(i)) + 1
  Next
  LetterCountsS1 = LetterCounts

  ReDim LetterCounts(65 To 90)
  Bytes = StrConv(UCase$(Text2), vbFromUnicode)
  For i = 0 To UBound(Bytes)
    If 65 <= Bytes(i) And Bytes(i) <= 90 Then
      LetterCounts(Bytes(i)) = LetterCounts(Bytes(i)) + 1
    End If
  Next
  LetterCountsS2 = LetterCounts

  SameWords = LetterCountsS1 = LetterCountsS2
End Function



Public Function TimedMsgBox(Prompt As String, Optional ByVal TimeOut As Long = 0, Optional Icon As VbMsgBoxStyle = vbOKOnly, Optional Title As String = vbNullString)

  Dim TimerId As Long
  #If VBA7 Then
     Dim nADDRESSOF As Long
  #End If

  CurMBTitle = Title

  If TimeOut = 0 Then
    TimeOut = 5 * 1000
  Else
    TimeOut = TimeOut * 1000
  End If
  
  
  #If VBA7 Then
    nADDRESSOF = CLng(AddressOf TimeOutMB)
    TimerId = SetTimer(0, 0, TimeOut, nADDRESSOF)
  #Else
     TimerId = SetTimer(0, 0, TimeOut, AddressOf TimeOutMB)
  #End If
  TimedMsgBox = MsgBox(Prompt, Icon, CurMBTitle)
  TimedMsgBox = 0
  KillTimer 0, TimerId

End Function

Private Sub TimeOutMB(hwnd As Long, uMsg As Long, idEvent As Long, dwTime As Long)

  SendMessage FindWindow(vbNullString, CurMBTitle), WM_CLOSE, 0&, 0&

End Sub
Function CheckRegEx(texto As String, padrao As String)

  Dim objRegExp As RegExp

  'cria um objeto expressão regular
  Set objRegExp = New RegExp

  'define o padrão - Pattern
  objRegExp.Pattern = padrao

  'define IgnoreCase
  objRegExp.IgnoreCase = True

  'define a propriedade global
  objRegExp.Global = True

  'verifica se a string pode ser comparada
  CheckRegEx = False
  If (objRegExp.Test(texto) = True) Then
    CheckRegEx = True

    'obtem as coincidencias
    '' Set colMatch = objRegExp.Execute(texto)   'executa a busca


    'For Each objMatch In colMatch
    ''   valor = valor & " padrao encontrado na posição "
    ''   valor = valor & objMatch.FirstIndex & ". o valor é '"
    ''   valor = valor & objMatch.Value & " '." & vbCrLf
    'Next
  Else
    ' valor = "Comparacao falhou !"
  End If

Set objRegExp = Nothing ' Adicione isto para boa gestão de memória

End Function

'Returns True if the user is running 64-bit *VBA*
'   This is *NOT* the same as the Windows bitness;
'   64-bit Windows can (and often does) run 32-bit VBA
Public Function Isvba64() As Boolean
Isvba64 = False
    #If Win64 Then
        Isvba64 = True
    #End If
End Function

Public Function txttodoc(ByVal cOrigem As String, _
                                     Optional ByVal cDestino As String = "", _
                                     Optional ByVal cTITULO As String = "", _
                                     Optional ByVal cAUTOR As String = "") As Boolean
    Dim fso As Object
    Dim streamIn As Object
    Dim streamOut As Object
    Dim cLINHA As String
    Dim strOutputFile As String
    
    ' 1. Validação inicial: se o arquivo de origem não existir, aborta retornando False
    If Dir(cOrigem) = "" Then
        txttodoc = False
        Exit Function
    End If
    
    ' 2. Regra do Destino: Se não foi passado, gera o .doc baseado no arquivo de origem
    If Trim(cDestino) = "" Then
        strOutputFile = Replace(LCase(cOrigem), ".txt", ".doc")
    Else
        strOutputFile = cDestino
    End If
    
    On Error GoTo TrataErro
    
    ' 3. Inicializa o FSO para o processamento em Streaming (Leve e rápido)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set streamIn = fso.OpenTextFile(cOrigem, 1, False)
    Set streamOut = fso.OpenTextFile(strOutputFile, 2, True)
    
    ' 4. Escreve o cabeçalho HTML com suporte a metadados de Título e Autor para o Word
    streamOut.WriteLine "<html>"
    streamOut.WriteLine "<head>"
    streamOut.WriteLine "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
    
    ' Injeta o Autor e Título se eles tiverem sido informados por parâmetro
    If Trim(cAUTOR) <> "" Then
        streamOut.WriteLine "<meta name=""author"" content=""" & cAUTOR & """>"
    End If
    If Trim(cTITULO) <> "" Then
        streamOut.WriteLine "<title>" & cTITULO & "</title>"
    End If
    
    ' Estilização CSS para o Word manter as fontes monoespaçadas e respeitar os espaços em branco do TXT
    streamOut.WriteLine "<style>"
    streamOut.WriteLine "body { font-family: 'Courier New', Courier, monospace; font-size: 10pt; line-height: 1.2; }"
    streamOut.WriteLine "p { margin: 0; padding: 0; white-space: pre; }"
    streamOut.WriteLine "</style>"
    streamOut.WriteLine "</head>"
    streamOut.WriteLine "<body>"
    
    ' 5. Loop de Leitura/Escrita linha por linha
    Do While Not streamIn.AtEndOfStream
        cLINHA = streamIn.ReadLine
        
        ' Escapa caracteres que podem quebrar a sintaxe do HTML
        cLINHA = Replace(cLINHA, "&", "&amp;")
        cLINHA = Replace(cLINHA, "<", "&lt;")
        cLINHA = Replace(cLINHA, ">", "&gt;")
        
        ' Se for linha vazia, adiciona o espaço fixo para o Word não "engolir" a linha
        If Len(Trim(cLINHA)) = 0 Then
            streamOut.WriteLine "<p>&nbsp;</p>"
        Else
            streamOut.WriteLine "<p>" & cLINHA & "</p>"
        End If
    Loop
    
    ' 6. Fecha as tags do documento
    streamOut.WriteLine "</body>"
    streamOut.WriteLine "</html>"
    
    ' Se chegou até aqui sem erros, operação concluída com sucesso!
    txttodoc = True

Fim:
    ' Liberação e fechamento seguro de canais de memória
    If Not streamIn Is Nothing Then streamIn.Close
    If Not streamOut Is Nothing Then streamOut.Close
    Set streamIn = Nothing
    Set streamOut = Nothing
    Set fso = Nothing
    Exit Function

TrataErro:
    ' Em caso de qualquer falha (permissão de escrita, arquivo travado), retorna False
    txttodoc = False
    Resume Fim
End Function
Public Function txttortf(ByVal cOrigem As String, Optional ByVal cDestino As String = "", Optional ByVal cTITULO As String = "", Optional ByVal cAUTOR As String = "") As Boolean
    Dim fso As Object
    Dim streamIn As Object
    Dim streamOut As Object
    Dim cLINHA As String
    Dim cBlocoInfo As String
    
    ' Mantém a sua validação original usando FileConnExist
    If Not FileConnExist(cOrigem, True) Then Exit Function
    
    ' Mantém a sua lógica original para caminhos e extensões padrões
    If Len(cDestino) = 0 Then cDestino = TrocaExt(cOrigem, "RTF")
    
    ' Impede a sobreescrita acidental conforme o seu padrão
    If FileConnExist(cDestino, False) Then
        Alert ("Arquivo Destino Ja existe")
        Exit Function
    End If
    
    ' Mantém os valores padrões vindos do seu sistema, idêntico à sua txttohtml
    If Len(cAUTOR) = 0 Then cAUTOR = zNOMEFOLHA
    If Len(cTITULO) = 0 Then cTITULO = NomeArq(cOrigem, True)
    
    On Error GoTo TrataErro
    
    ' Inicializa o FileSystemObject para leitura e gravação em fluxo (streaming)
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Abre o arquivo de origem para leitura (1 = ForReading)
    Set streamIn = fso.OpenTextFile(cOrigem, 1, False)
    ' Cria o arquivo de destino para escrita (True = Sobreescrever)
    Set streamOut = fso.CreateTextFile(cDestino, True)
    
    ' --- CONSTRUÇÃO DOS METADADOS (BLOCO INFO) ---
    ' Monta a tag {\info} se houver título ou autor preenchidos
    cBlocoInfo = ""
    If Len(cTITULO) > 0 Or Len(cAUTOR) > 0 Then
        cBlocoInfo = "{\info"
        If Len(cTITULO) > 0 Then cBlocoInfo = cBlocoInfo & "{\title " & cTITULO & "}"
        If Len(cAUTOR) > 0 Then cBlocoInfo = cBlocoInfo & "{\author " & cAUTOR & "}"
        cBlocoInfo = cBlocoInfo & "}"
    End If
    
    ' Escreve o cabeçalho de inicialização do formato RTF juntamente com os metadados
    streamOut.Write "{\rtf1\ansi\ansicpg1252\deff0\deflang1031" & cBlocoInfo & "{\fonttbl{\f0\fnil\fcharset0 Arial;}}\viewkind4\uc1\pard\f0\fs20 "
    
    ' Loop de streaming otimizado por linha para arquivos gigantes
    Do While Not streamIn.AtEndOfStream
        cLINHA = streamIn.ReadLine
        
        ' Proteção para que caracteres de controle do RTF presentes no texto original não corrompam o arquivo
        If InStr(cLINHA, "\") > 0 Then cLINHA = Replace(cLINHA, "\", "\\")
        If InStr(cLINHA, "{") > 0 Then cLINHA = Replace(cLINHA, "{", "\{")
        If InStr(cLINHA, "}") > 0 Then cLINHA = Replace(cLINHA, "}", "\}")
        
        ' Grava no arquivo adicionando o controle de fim de parágrafo (\par) do RTF
        streamOut.Write cLINHA & "\par " & vbCrLf
    Loop
    
    ' Escreve a tag de encerramento do documento RTF
    streamOut.Write "\par}"
    
    txttortf = True

Fim:
    ' Garante a desconexão dos arquivos e liberação completa dos objetos de streaming da memória
    If Not streamIn Is Nothing Then streamIn.Close
    If Not streamOut Is Nothing Then streamOut.Close
    Set streamIn = Nothing
    Set streamOut = Nothing
    Set fso = Nothing
    Exit Function

TrataErro:
    Alert "Erro na gravação/leitura do arquivo RTF: " & Err.Description
    Resume Fim
End Function
Public Function txttohtml(ByVal cOrigem As String, Optional ByVal cDestino As String = "", Optional ByVal cTITULO As String = "", Optional ByVal cAUTOR As String = "") As Boolean
    Dim fso As Object
    Dim streamIn As Object
    Dim streamOut As Object
    Dim cLINHA As String
    
    ' Mantém a sua validação original usando FileConnExist
    If Not FileConnExist(cOrigem, True) Then Exit Function
    
    ' Mantém a sua lógica original para caminhos e extensões padrões
    If Len(cDestino) = 0 Then cDestino = TrocaExt(cOrigem, "HTML")
    
    ' Impede a sobreescrita acidental conforme seu código original
    If FileConnExist(cDestino, False) Then
        Alert ("Arquivo Destino Ja existe")
        Exit Function
    End If
    
    ' Mantém os valores padrões vindos das suas variáveis globais e funções de string
    If Len(cAUTOR) = 0 Then cAUTOR = zNOMEFOLHA
    If Len(cTITULO) = 0 Then cTITULO = NomeArq(cOrigem, True)
    
    On Error GoTo TrataErro
    
    ' Inicializa o FileSystemObject para ler e gravar em fluxo de buffer
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Abre o arquivo de origem para leitura (1 = ForReading)
    Set streamIn = fso.OpenTextFile(cOrigem, 1, False)
    ' Cria o arquivo de destino para escrita (True = Sobreescrever)
    Set streamOut = fso.CreateTextFile(cDestino, True)
    
    ' Escreve exatamente a mesma estrutura de cabeçalho que o seu Print original fazia
    streamOut.Write "<html>" & vbCrLf
    streamOut.Write "<head>" & vbCrLf
    streamOut.Write "<meta http-equiv=" & Chr(34) & "Content-Type" & Chr(34) & vbCrLf
    streamOut.Write "content=" & Chr(34) & "text/html; charset=iso-8859-1" & Chr(34) & ">" & vbCrLf
    streamOut.Write "<meta name=" & Chr(34) & "GENERATOR" & Chr(34) & "content=" & Chr(34) & cAUTOR & Chr(34) & ">" & vbCrLf
    streamOut.Write "<title>" & cTITULO & "</title>" & vbCrLf
    streamOut.Write "</head>" & vbCrLf
    streamOut.Write "<body>" & vbCrLf
    
    ' Loop de streaming otimizado linha por linha para arquivos gigantes
    Do While Not streamIn.AtEndOfStream
        cLINHA = streamIn.ReadLine
        
        ' AQUI: Aplica a função original str2html para tratar acentos e caracteres especiais
        cLINHA = str2html(cLINHA)
        
        ' Joga no arquivo destino adicionando a quebra de linha do HTML
        streamOut.Write cLINHA & "<br>" & vbCrLf
    Loop
    
    ' Fecha a estrutura do documento exatamente como o seu original
    streamOut.Write "</body>" & vbCrLf
    streamOut.Write "</html>" & vbCrLf
    
    txttohtml = True

Fim:
    ' Garante o fechamento correto dos arquivos e liberação da RAM
    If Not streamIn Is Nothing Then streamIn.Close
    If Not streamOut Is Nothing Then streamOut.Close
    Set streamIn = Nothing
    Set streamOut = Nothing
    Set fso = Nothing
    Exit Function

TrataErro:
    Alert "Erro na gravação/leitura do arquivo: " & Err.Description
    Resume Fim
End Function
' ==============================================================================
' MOTOR 4: CONVERSÃO DE TXT PARA EXCEL - TRATAMENTO MULTI-DELIMITADORES EM STREAM
' ==============================================================================
Public Function txttoxls(ByVal cOrigem As String, _
                                      Optional ByVal cDestino As String = "", _
                                      Optional ByVal cTITULO As String = "", _
                                      Optional ByVal cAUTOR As String = "", _
                                      Optional ByVal cDELIMITADOR As String = ";") As Boolean
    Dim fso As Object
    Dim streamIn As Object
    Dim streamOut As Object
    Dim cLINHA As String
    Dim strOutputFile As String
    Dim arrCampos() As String
    Dim i As Long
    Dim cDelimReal As String
    Dim bPrimeiraLinha As Boolean
    Dim cValor As String
    
    ' 1. Validação inicial: se o arquivo de origem não existir, aborta
    If Dir(cOrigem) = "" Then
        txttoxls = False
        Exit Function
    End If
    
    ' 2. Regra do Destino: Se não foi passado, gera o .xls baseado no nome do TXT
    If Trim(cDestino) = "" Then
        strOutputFile = Replace(LCase(cOrigem), ".txt", ".xls")
    Else
        strOutputFile = cDestino
    End If
    
    ' 3. TRATAMENTO DOS DELIMITADORES DO SEU FORMULÁRIO (tab, ;,  ,, |, ~#)
    cDELIMITADOR = Trim(LCase(cDELIMITADOR))
    
    Select Case cDELIMITADOR
        Case "<tab>", "tab", Chr(9)
            cDelimReal = Chr(9)   ' Tabulação Real
        Case ""
            cDelimReal = ";"      ' Fallback padrão se vier vazio
        Case Else
            cDelimReal = cDELIMITADOR ' Captura ;, ,, |, ~# exatamente como estão no Form
    End Select
    
    On Error GoTo TrataErro
    
    ' 4. Inicializa o FileSystemObject para processamento leve (Streaming)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set streamIn = fso.OpenTextFile(cOrigem, 1, False)       ' 1 = ForReading
    Set streamOut = fso.OpenTextFile(strOutputFile, 2, True) ' 2 = ForWriting
    
    ' 5. Escreve o cabeçalho HTML formatado para Excel e LibreOffice
    streamOut.WriteLine "<html>"
    streamOut.WriteLine "<head>"
    streamOut.WriteLine "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
    
    ' Metadados do documento
    If Trim(cAUTOR) <> "" Then streamOut.WriteLine "<meta name=""author"" content=""" & cAUTOR & """>"
    If Trim(cTITULO) <> "" Then streamOut.WriteLine "<title>" & cTITULO & "</title>"
    
    ' CSS para forçar linhas e bordas visíveis no Excel/Calc
    streamOut.WriteLine "<style>"
    streamOut.WriteLine "table { border-collapse: collapse; }"
    streamOut.WriteLine "td, th { border: 1px solid #A0A0A0; font-family: Calibri, Arial, sans-serif; font-size: 11pt; padding: 5px; }"
    streamOut.WriteLine "th { background-color: #EFEFEF; font-weight: bold; }"
    streamOut.WriteLine "</style>"
    streamOut.WriteLine "</head>"
    streamOut.WriteLine "<body>"
    streamOut.WriteLine "<table>"
    
    bPrimeiraLinha = True
    
    ' 6. Loop de Performance: Separa as colunas com base no delimitador tratado
    Do While Not streamIn.AtEndOfStream
        cLINHA = streamIn.ReadLine
        
        ' Faz a divisão da linha usando o delimitador correto correspondente ao do Form
        arrCampos = Split(cLINHA, cDelimReal)
        
        streamOut.WriteLine "  <tr>"
        For i = LBound(arrCampos) To UBound(arrCampos)
            
            ' Isola o conteúdo da célula tirando espaços extras
            cValor = Trim(CStr(arrCampos(i)))
            
            ' Aplica a sua função nativa str2html para tratar acentos e caracteres especiais
            cValor = str2html(cValor)
            
            ' Se for a primeira linha do TXT, vira cabeçalho destacado (th)
            If bPrimeiraLinha Then
                streamOut.WriteLine "    <th>" & cValor & "</th>"
            Else
                ' Gravação correta da célula de dados (td)
                streamOut.WriteLine "    <td>" & cValor & "</td>"
            End If
        Next i
        streamOut.WriteLine "  </tr>"
        
        bPrimeiraLinha = False
    Loop
    
    ' 7. Fecha o arquivo de forma íntegra
    streamOut.WriteLine "</table>"
    streamOut.WriteLine "</body>"
    streamOut.WriteLine "</html>"
    
    txttoxls = True

Fim:
    ' Destruição rigorosa de objetos para não prender memória RAM
    If Not streamIn Is Nothing Then streamIn.Close
    If Not streamOut Is Nothing Then streamOut.Close
    Set streamIn = Nothing
    Set streamOut = Nothing
    Set fso = Nothing
    Exit Function

TrataErro:
    txttoxls = False
    Resume Fim
End Function


'Public Function MachineName() As String
    ' O objeto OS do RC6 já contém essa informação formatada
    ' Funciona identicamente em 32 ou 64 bits, sem necessidade de PtrSafe
 '   MachineName = New_c.OS.GetComputerName
'End Function

'Public Function NetworkUserName() As String
    ' Igualmente simples para o nome do usuário
'    NetworkUserName = New_c.OS.UserName
'End Function
