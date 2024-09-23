Attribute VB_Name = "Myfunctions"
'fileexist(cARQ)Verifica a existencia de Uma Arquivo
'pegpath(Cgrupo,Ccamp) 'Traz caminho de uma arquivo ini(nomeaplicativo.ini)
'today() data atual
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

Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private CurMBTitle As String

'constantes para syscolors
'Dim lngReturn As Long
'lngReturn = SetSysColors(1, COLOR_ACTIVECAPTION ,RGB(255 0 0))
Declare Function SetSysColors Lib "user32" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Public Const COLOR_SCROLLBAR = 0  'The Scrollbar color
Public Const COLOR_BACKGROUND = 1  'Colour of the background with no wallpaper
Public Const COLOR_ACTIVECAPTION = 2  'Caption of Active Window
Public Const COLOR_INACTIVECAPTION = 3  'Caption of Inactive window
Public Const COLOR_MENU = 4  'Menu
Public Const COLOR_WINDOW = 5  'Windows background
Public Const COLOR_WINDOWFRAME = 6  'Window frame
Public Const COLOR_MENUTEXT = 7  'Window Text
Public Const COLOR_WINDOWTEXT = 8  '3D dark shadow (Win95)
Public Const COLOR_CAPTIONTEXT = 9  'Text in window caption
Public Const COLOR_ACTIVEBORDER = 10  'Border of active window
Public Const COLOR_INACTIVEBORDER = 11  'Border of inactive window
Public Const COLOR_APPWORKSPACE = 12  'Background of MDI desktop
Public Const COLOR_HIGHLIGHT = 13  'Selected item background
Public Const COLOR_HIGHLIGHTTEXT = 14  'Selected menu item
Public Const COLOR_BTNFACE = 15  'Button
Public Const COLOR_BTNSHADOW = 16  '3D shading of button
Public Const COLOR_GRAYTEXT = 17  'Grey text of zero if dithering is used.
Public Const COLOR_BTNTEXT = 18  'Button text
Public Const COLOR_INACTIVECAPTIONTEXT = 19  'Text of inactive window
Public Const COLOR_BTNHIGHLIGHT = 20  '3D highlight of button

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
Public Const WS_MINIMIZEBOX As Long = &H20000
Public Const WS_MAXIMIZEBOX As Long = &H10000
Public Const GWL_STYLE As Long = (-16)

Private Type SHFILEOPSTRUCT
  hWnd As Long
  wFunc As Long
  pFrom As String
  pTo As String
  fFlags As Integer
  fAnyOperationsAborted As Boolean
  hNameMappings As Long
  lpszProgressTitle As String
End Type

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

Public Declare Function WinAPI_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function InternetGetConnectedState Lib "wininet" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function EbExecuteLine Lib "vba6.dll" (ByVal pStringToExec As Long, ByVal Unknownn1 As Long, ByVal Unknownn2 As Long, ByVal fCheckOnly As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
''Exemplos  find   Call ShellExecute(Me.hwnd, "find", "C:\", vbNullString, vbNullString, SW_SHOWNORMAL)
Public Declare Function ShellExecuteForExplore Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, lpParameters As Any, lpDirectory As Any, ByVal nShowCmd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Declare Function WinAPI_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'Declare Function WinAPI_GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
''Chama O HELP AVULSO
''hh_ variaveis acima
''Exemplos(0,"c:\temp\teste.chm",HH_DISPLAY_TOPIC,20)
''no mdi atribua os shortcurts/helpid no menueditor
'html ja na clscommondialog com enum e funcao ShowHTMLHelp
'Public Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, dwData As Any) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias _
                                                "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                                            ByVal lpKeyName As Any, ByVal lpDefault As String, _
                                                                            ByVal lpReturnedString As String, ByVal nSize As Long, _
                                                                            ByVal lpFileName As String) As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias _
                                                  "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                                                ByVal lpKeyName As Any, ByVal lpString As Any, _
                                                                                ByVal lpFileName As String) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function PlaySound Lib "winmm.dll" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
'Tirar botao maximar minizar
Public Declare Function SetWindowLong Lib "user32" _
                                      Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                                              ByVal nIndex As Long, ByVal dwNewLong As _
                                                                                    Long) As Long
Public Declare Function GetWindowLong Lib "user32" _
                                      Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                                              ByVal nIndex As Long) As Long

'declare para data hora do sistema
Private Declare Function PostMessage Lib "user32" Alias _
                                     "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
                                                     ByVal wParam As Long, ByVal lParam As Long) As Long

Declare Function GetLocaleInfoa Lib "kernel32" Alias _
                               "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCTYPE As Long, _
                                                 ByVal lpLCData As String, ByVal cchData As Long) As Long

Declare Function SetLocaleInfoa Lib "kernel32" Alias _
                               "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCTYPE As Long, _
                                                 ByVal lpLCData As String) As Boolean

Declare Function GetUserDefaultLCID% Lib "kernel32" ()
Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Public Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Public Declare Function OemToChar Lib "user32" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
Private Sub ForceSystemDecimalToPeriod()
    ' We MUST use the ANSI API version so it's an ANSI character that's used for the actual decimal character.
    Const LOCALE_SDECIMAL   As Long = &HE&
    Const LOCALE_SGROUPING  As Long = &H10&
    Const Eng_LCID          As Long = 1033&
    Dim S                   As String
    '
    S = String$(GetLocaleInfoa(Eng_LCID, LOCALE_SDECIMAL, vbNullString, 0&), 0)
    GetLocaleInfoa Eng_LCID, LOCALE_SDECIMAL, S, Len(S)
    If RTrimNull(S) <> "." Then
        SetLocaleInfoa Eng_LCID, LOCALE_SDECIMAL, "."
        SetLocaleInfoa Eng_LCID, LOCALE_SGROUPING, ","
    End If
End Sub

Public Function tocar(cARQ)
  PlaySound (cARQ), ByVal 0&, SND_FILENAME Or SND_ASYNC
End Function

Public Function IsWebConnected(Optional ByRef ConnType As String) As Boolean
  Dim dwFlags As Long
  Dim WebTest As Boolean
  ConnType = ""
  WebTest = InternetGetConnectedState(dwFlags, 0&)
  Select Case WebTest
  Case dwFlags And CONNECT_LAN: ConnType = "LAN"
  Case dwFlags And CONNECT_MODEM: ConnType = "Modem"
  Case dwFlags And CONNECT_PROXY: ConnType = "Proxy"
  Case dwFlags And CONNECT_OFFLINE: ConnType = "Offline"
  Case dwFlags And CONNECT_CONFIGURED: ConnType = "Configured"
  Case dwFlags And CONNECT_RAS: ConnType = "Remote"
  End Select
  IsWebConnected = WebTest
End Function

Public Function ExecutarDLL(ByVal cDLL As String)
' appwiz.cpl adicionar e remover programas
' odbccp32.cpl gerenciador odbc
' sysdm.cpl informacoes do sistema
' outro cpl
' outra dll que possa ser executada via rundll32
  Dim RetVal
  On Error Resume Next
  RetVal = Shell("rundll32.exe shell32.dll,Control_RunDLL " & cDLL & ",,3", 1)
End Function

Public Function ExecuteLine(ByVal Scode As String, Optional ByVal fCheckOnly As Boolean) As Boolean
  ExecuteLine = EbExecuteLine(StrPtr(Scode), 0&, 0&, Abs(fCheckOnly)) = 0
End Function

Public Function ComboLostFocus(ByRef Combo1)
  With Combo1
    If Len(.tEXT) Then
      'Procura pelo texto digitado
      strPartial = .tEXT
      i = SendMessage(.hWnd, CB_FINDSTRING, -1, ByVal strPartial)
      'Se não achou, retorna      o focus para o Combo
      If i = CB_ERR Then .SetFocus
    End If
  End With
End Function

Public Function ComboChange(ByRef Combo1)
  With Combo1
    'Procura pelo texto já digitado
    strPartial = .tEXT
    i = SendMessage(.hWnd, CB_FINDSTRING, -1, _
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
  SendMessage Obj.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 1
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
Public Function SepSqlOpe(ByVal eEXP As String) As Variant
  Dim aRETU As Variant
  aRETU = Array("", "", "")
  If InStr(eEXP, ",") > 0 Then
    aRETU = Split(eEXP, ",")
  End If
  SepSqlOpe = aRETU
End Function
Public Function CharToLit(aVAL As Variant, Optional ByVal cTIPO As String = "", Optional ByVal cARQ As String = "") As String
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
erro:
  Select Case Err.Number
  Case Else
    SayErro "ChartoLit"
  End Select
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
  Case "MYSQL", "MYSQL/"
    DataToLit = "'" & Year(dDATA) & "/" & Month(dDATA) & "/" & Day(dDATA) & "'"
  Case "MYSQL-", "MARIADB"
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

Public Function Alert(ByVal cDIZ As String, Optional ByVal cTITLE As String = "Informacao")
  MsgBox cDIZ, vbOKOnly, cTITLE
End Function

Public Function Busca(ByVal cDIZ As String, ByVal cCAB As String, ByVal cVAL As String, ByVal nLEN As Long) As String
  Busca = InputBox(cDIZ, cCAB, cVAL)
  Busca = Left$(Busca, nLEN)
End Function

Public Function Caminex(ByVal cARQ As String, Optional ByVal nANO As Integer = 0, Optional ByVal nMES As Integer = 0, Optional ByVal nEMP = 1)
  Dim cANO, cEMP, cMES, cTMPVAR As String
  Dim nPOS, nPOS2 As Integer
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
    cMES = StrZero(nMES, 2)
    cARQ = Replace(cARQ, "[MM]", cMES)
  End If
  If InStr(cARQ, "[AAAA]") > 0 Then
    cANO = StrZero(nANO, 4)
    cARQ = Replace(cARQ, "[AAAA]", cANO)
  End If
  If InStr(cARQ, "[AA]") > 0 Then
    cANO = StrZero(nANO, 4)
    cANO = Right(cANO, 2)
    cARQ = Replace(cARQ, "[AA]", cANO)
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
  numdata = Fdata(dDATA)
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

Public Function FileExist(ByVal cARQ As Variant, _
                          Optional ByVal lMES As Boolean = False, _
                          Optional ByVal cMES As String = "Arquivo Não Encontrado ", _
                          Optional ByVal cSQL As String = "")
  Dim cARQUIVO As String
  Dim nFILELEN
  Dim nPOS As Long
  Dim cEXT As String
  On Error GoTo NotExist
  FileExist = False
  cARQUIVO = UCase(Trim(CStr(cARQ)))
  If Len(cARQ) = 0 Then
     FileExist = False
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
          FileExist = True
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
          FileExist = True
          Exit Function
        End If
        cARQUIVO = cARQUIVO & NomeTableSql(cSQL, ".DB")
      End If

      If InStr(cARQUIVO, "[CONN]") > 0 Or InStr(cARQUIVO, "[MYSQL]") > 0 Or InStr(cARQUIVO, "[MARIADB]") > 0 Or _
         InStr(cARQUIVO, "[POSTGRESQL]") > 0 Or InStr(cARQUIVO, "[SQLSERVER]") > 0 Then
        FileExist = True
        Exit Function
      End If
      If InStr(cARQUIVO, "[FIREBIRD]") > 0 Or InStr(cARQUIVO, "[INTERBASE]") > 0 Then
        nPOS = InStr(cARQUIVO, "DBNAME=")
        cARQUIVO = Mid(cARQUIVO, nPOS + 7)
        nPOS = InStr(cARQUIVO, ";")
        If nPOS > 0 Then
          cARQUIVO = Mid(cARQUIVO, 1, nPOS - 1)
        End If
        FileExist = True
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
  If InStr(cARQUIVO, "PROVIDER") > 0 Then
    FileExist = True
    Exit Function
  End If
  If FileExists(cARQUIVO) Then
    
    FileExist = True
  End If

  Exit Function
NotExist:
  FileExist = False
  If lMES Then
    Alert cMES & cARQUIVO, "Erro Arquivo"
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
  Dim nPOS01, nPOS02 As Integer
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

Public Function IncDec(ByVal eVAR, ByVal nVAL, ByVal nLIMINF, ByVal nLIMSUP)
  IncDec = FixNum(eVAR)
  IncDec = IncDec + nVAL
  If IncDec < nLIMINF Then IncDec = nLIMINF
  If IncDec > nLIMSUP Then IncDec = nLIMSUP
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
  'Dim nPOS As Integer
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

Public Function NullDate(Optional ByVal cTIPO As String = "", Optional ByVal cARQ As String = "") As Variant
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

Public Function PegCamini(ByVal cCAMINHO As String) As String
  PegCamini = Caminex(cCAMINHO, 0, 0, 0)
End Function

Public Function PegPath(ByVal cGRUPO As String, ByVal cCAMPO As String, Optional ByVal ePAD As String = "")
  Dim z As Long
  Dim scaminho As String * 255
  z = GetPrivateProfileString(cGRUPO, cCAMPO, "", scaminho, 150, App.Path + "\" & App.EXEName & ".INI")
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
  Dim sPalavra As String, iPosIni As Integer
  Dim iPosFim As Integer, sResultado As String
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



  Set OSVer = New clsOSInfo
  With OSVer
    cERRO = cERRO & vbCrLf & "IP: " & .IPAddress
    cERRO = cERRO & vbCrLf & "Local: " & .IPHostName
    cERRO = cERRO & vbCrLf & "Equipamento: " & .MachineName
    cERRO = cERRO & vbCrLf & "OS Name: " & .OSName
    cERRO = cERRO & vbCrLf & "Service Pack ver.: " & .SPVer
    cERRO = cERRO & vbCrLf & "Bitness: " & .Bitness
    cERRO = cERRO & vbCrLf & "Edition: " & .Edition
    cERRO = cERRO & vbCrLf & "Family: " & .Family
    cERRO = cERRO & vbCrLf & "Suite mask: " & .SuiteMask
    cERRO = cERRO & vbCrLf & "ProductType: " & .ProductType
    cERRO = cERRO & vbCrLf & "OS - Vista or newer? " & .IsVistaOrLater
    cERRO = cERRO & vbCrLf & "Major: " & .Major
    cERRO = cERRO & vbCrLf & "Minor: " & .Minor
    cERRO = cERRO & vbCrLf & "Major + Minor: " & .MajorMinor
    cERRO = cERRO & vbCrLf & "Build: " & .Build
    cERRO = cERRO & vbCrLf & "ReleaseId: " & .ReleaseId
    cERRO = cERRO & vbCrLf & "Language in dialogues: " & .LangDisplayCode & " " & .LangDisplayName & " " & .LangDisplayNameFull
    cERRO = cERRO & vbCrLf & "Language of OS inslallation: " & .LangSystemCode & " " & .LangSystemName & " " & .LangSystemNameFull
    cERRO = cERRO & vbCrLf & "Language for non-Unicode programs: " & .LangNonUnicodeCode & " " & .LangNonUnicodeName & " " & .LangNonUnicodeNameFull
    cERRO = cERRO & vbCrLf & "Process integrity level: " & .IntegrityLevel
    cERRO = cERRO & vbCrLf & "Elevated process? " & .IsElevated
    cERRO = cERRO & vbCrLf & "User group: " & .UserType
    cERRO = cERRO & vbCrLf & "Safe boot? " & .IsSafeBoot & " (" & .SafeBootMode & ")"
    cERRO = cERRO & vbCrLf & "OEM Codepage: " & .CodepageOEM & " (" & .CodepageOEM_File & ")"
    cERRO = cERRO & vbCrLf & "ANSI Codepage: " & .CodepageANSI & " (" & .CodepageANSI_File & ")"
  End With

  nHANDLE = FreeFile
  Open cARQ For Output As #nHANDLE
  Print #nHANDLE, cERRO
  Close (nHANDLE)

End If
End Sub

Public Function ShellEx( _
       ByVal sFile As String, _
       Optional ByVal eShowCmd As EShellShowConstants = essSW_SHOWDEFAULT, _
       Optional ByVal sParameters As String = "", _
       Optional ByVal sDefaultDir As String = "", _
       Optional sOperation As String = "open", _
       Optional Owner As Long = 0 _
     ) As Boolean
  Dim lr As Long
  Dim lErr As Long, sErr As String
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
    Case ERROR_FILE_NOT_FOUND
      lErr = 53: sErr = "Arquivo nao encontrado"
    Case ERROR_PATH_NOT_FOUND
      lErr = 76: sErr = "Caminho nao encontrado"
    Case ERROR_BAD_FORMAT
      sErr = "Executavel invalido ou corrompido"
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

Public Function SomaArr(ByVal aARRAY As Variant, ByVal nITEM As Integer)
  Dim x As Integer
  SomaArr = 0
  For x = 0 To nITEM - 1
    SomaArr = SomaArr + FixNum(aARRAY(x))
  Next x
End Function

Public Function SomaExt(ByVal cARQ As String, Optional ByVal cEXT As String = ".MDB") As String
  If InStr(UCase(cARQ), cEXT) > 0 Then
    SomaExt = cARQ
  Else
    SomaExt = cARQ & cEXT
  End If
End Function

Public Function StrZero(ByVal nNUM, Optional ByVal nLEN As Integer = 0)
  Dim cTEMP
  If nLEN = 0 Then
    cTEMP = FixStr(nNUM, "0", "TRIM")
    nLEN = Trim(cTEMP)
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
Public Function ConvOEM(ByVal texto As String) As String
  ConvOEM = Convert2oem(cTEXTO)  'CharConv(texto, "ANSI", "OEM")
End Function

Public Function ConvOEM2(ByVal texto As String) As String
  ConvOEM2 = Convert2oem(cTEXTO)  ' CharConv(texto, "MEUANSI", "MEUOEM")
End Function

Public Function ConvAnsi2(ByVal texto As String) As String
  ConvAnsi2 = Convert2ansi(texto)  'CharConv(texto, "MEUOEM", "MEUANSI")
End Function

Public Function ConvAnsi(ByVal texto As String) As String
  ConvAnsi = Convert2ansi(texto)  ' CharConv(texto, "OEM", "ANSI")
End Function

Public Function Tirace(ByVal texto As String) As String
' texto = CharConv(texto, "UACENTO", "UACETIR")
' texto = CharConv(texto, "LACENTO", "LACETIR")
  Tirace = tirace2(texto)
End Function
Public Function StrToArray(ByVal cGRUPO As String) As Variant
  Dim x, nLEN As Integer
  Dim aUSO As Variant
  Dim cCHAR, eCNV As String
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

Public Function TiraSin(ByVal texto As String)
  For x = 0 To 31
    texto = Replace(texto, Chr(x), "")
  Next x
  For x = 33 To 38
    texto = Replace(texto, Chr(x), "")
  Next x
  For x = 39 To 47
    texto = Replace(texto, Chr(x), "")
  Next x
  For x = 58 To 64
    texto = Replace(texto, Chr(x), "")
  Next x
  For x = 91 To 96
    texto = Replace(texto, Chr(x), "")
  Next x
  For x = 123 To 127
    texto = Replace(texto, Chr(x), "")
  Next x
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
  TiraSin = texto
End Function
Function CheckPass(ByVal cTEXTO As String, Optional ByVal lMES As Boolean = True) As Boolean

  Dim lMAIS, lMINUS, lDIG, lSYMBOL, l8DIG As Boolean
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


Public Function Today() As Date
  Today = Format(Date, "dd/mm/yyyy")
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
  Dim nLEN, nTEXTO, x, y As Integer
  Dim aORI, aDES, aTEXTO As Variant
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

Public Function Extenso(ByVal valor As Double, _
                        Optional ByVal MoedaPlural As String = "Reais", _
                        Optional ByVal MoedaSingular As String = "Real") As String
  Dim StrValor As String, Negativo As Boolean
  Dim buf As String, Parcial As Integer
  Dim Posicao As Integer, Unidades
  Dim Dezenas, Centenas, PotenciasSingular
  Dim PotenciasPlural

  Negativo = (valor < 0)
  valor = Abs(CDec(valor))
  If valor Then
    Unidades = Array(vbNullString, "Um", "Dois", _
                     "Três", "Quatro", "Cinco", _
                     "Seis", "Sete", "Oito", "Nove", _
                     "Dez", "Onze", "Doze", "Treze", _
                     "Quatorze", "Quinze", "Dezesseis", _
                     "Dezessete", "Dezoito", "Dezenove")
    Dezenas = Array(vbNullString, vbNullString, _
                    "Vinte", "Trinta", "Quarenta", _
                    "Cinqüenta", "Sessenta", "Setenta", _
                    "Oitenta", "Noventa")
    Centenas = Array(vbNullString, "Cento", _
                     "Duzentos", "Trezentos", _
                     "Quatrocentos", "Quinhentos", _
                     "Seiscentos", "Setecentos", _
                     "Oitocentos", "Novecentos")
    PotenciasSingular = Array(vbNullString, " Mil", _
                            " Milhão", " Bilhão", _
                            " Trilhão", " Quatrilhão")
    PotenciasPlural = Array(vbNullString, " Mil", _
                          " Milhões", " Bilhões", _
                          " Trilhões", " Quatrilhões")

    StrValor = Left(Format(valor, String(18, "0") & _
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
      If Int(valor) = 1 Then
        Extenso = Extenso & " " & MoedaSingular
      Else
        Extenso = Extenso & " " & MoedaPlural
      End If
    End If
    Parcial = Int((valor - Int(valor)) * _
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

Public Function MesAnt(ByVal nMES As Integer, ByVal nANO As Long)
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
  Select Case cCHAR
  Case ""
    FixStrBol = Null
  Case "S"
    FixStrBol = True
  Case "N"
    FixStrBol = False
  End Select
End Function

Public Function Count_Lines_In_File(ByVal strFilePath As String, Optional ByVal lMES As Boolean = True) As Long

'     'delcare variables
  Dim fileFile As Integer
  Dim intLinesReadCount As Integer
  Dim STRBUFFER As String

  intLinesReadCount = 0
  '     'open file
  fileFile = FreeFile


  If Not FileExist(strFilePath, lMES) Then
    Count_Lines_In_File = -1
    Exit Function
  End If

  'loop through file
  Open strFilePath For Input As #fileFile


  Do While Not EOF(fileFile)
    'read line
    Input #fileFile, STRBUFFER
    'update count
    intLinesReadCount = intLinesReadCount + 1
  Loop
  'close file
  Close fileFile
  'return value
  Count_Lines_In_File = intLinesReadCount
End Function

Public Sub OpenUrl(ByVal strURL As String)
  ShellExecute 0, "Open", strURL, 0&, 0&, SW_SHOWNORMAL
End Sub
Public Function txttohtml(ByVal cORIGEM As String, Optional ByVal cDESTINO As String = "", Optional ByVal cTITULO As String = "", Optional ByVal cAUTOR As String = "")
  Dim nORIGEM, nDESTINO, STRBUFFER, cLINHA
  If Not FileExist(cORIGEM, True) Then
    Exit Function
  End If
  If Len(cDESTINO) = 0 Then
    cDESTINO = TrocaExt(cORIGEM, "HTML")
  End If
  If FileExist(cDESTINO, False) Then
    Alert ("Arquivo Destino Ja existe")
    Exit Function
  End If
  If Len(cAUTOR) = 0 Then
    cAUTOR = zNOMEFOLHA
  End If
  If Len(cTITULO) = 0 Then
    cTITULO = NomeArq(cORIGEM, True)
  End If
  nORIGEM = FreeFile
  nDESTINO = FreeFile + 1
  Open cDESTINO For Output As #nDESTINO
  Open cORIGEM For Input As #nORIGEM
  Print #nDESTINO, "<html>" & Chr(13) & Chr(10)
  Print #nDESTINO, "<head>" & Chr(13) & Chr(10)
  Print #nDESTINO, "<meta http-equiv=" & Chr(34) & "Content-Type" & Chr(34) & Chr(13) & Chr(10)
  Print #nDESTINO, "content=" & Chr(34) & "text/html; charset=iso-8859-1" & Chr(34) & ">" & Chr(13) & Chr(10)
  Print #nDESTINO, "<meta name=" & Chr(34) & "GENERATOR" & Chr(34) & "content=" & Chr(34) & cAUTOR & Chr(34) & ">" & Chr(13) & Chr(10)
  Print #nDESTINO, "<title>" & cTITULO & "</title>" & Chr(13) & Chr(10)
  Print #nDESTINO, "</head>" & Chr(13) & Chr(10)
  Print #nDESTINO, "<body>" & Chr(13) & Chr(10)
  Do While Not EOF(nORIGEM)
    Input #nORIGEM, STRBUFFER
    cLINHA = STRBUFFER
    cLINHA = str2html(cLINHA)
    Print #nDESTINO, cLINHA & "<BR>"
  Loop
  Print #nDESTINO, "</body>" & Chr(13) & Chr(10) & "</html>" & Chr(13) & Chr(10)
  Close nORIGEM
  Close nDESTINO
End Function

Public Function txttoRTF(ByVal cORIGEM As String, Optional ByVal cDESTINO As String = "")
  Dim nORIGEM, nDESTINO, STRBUFFER, cLINHA
  If Not FileExist(cORIGEM, True) Then
    Exit Function
  End If
  If Len(cDESTINO) = 0 Then
    cDESTINO = TrocaExt(cORIGEM, "RTF")
  End If
  If FileExist(cDESTINO, False) Then
    Alert ("Arquivo Destino Ja existe")
    Exit Function
  End If
  nORIGEM = FreeFile
  nDESTINO = FreeFile + 1
  Open cDESTINO For Output As #nDESTINO
  Open cORIGEM For Input As #nORIGEM
  Print #nDESTINO, "{\rtf1\ansi\ansicpg1252\deff0\deflang1031{\fonttbl{\f0\fnil\fcharset0 "
  Print #nDESTINO, "Arial;}}\viewkind4\uc1\pard\f0\fs20 "
  Do While Not EOF(nORIGEM)
    Input #nORIGEM, STRBUFFER
    cLINHA = STRBUFFER
    Print #nDESTINO, cLINHA & "\par" & Chr(13)
  Loop
  Print #nDESTINO, "\par}"
  Close nORIGEM
  Close nDESTINO
End Function
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
Function FileText(ByVal FileName As String) As String
  Dim Handle As Integer
  Handle = FreeFile
  Open FileName$ For Input As #Handle
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
  Dim id As Integer

  On Error Resume Next
  Err.Clear
  FindInList = -1

  For id = 0 To cList.ListCount - 1
    sString = UCase$(cList.List(id))
    If sString = UCase$(sSearch) Then
      FindInList = id
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
  Dim Pos As Long, i As Long, NumCharsProcessed As Long, TXT As String
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

  sString = String(255, 0)

  iStringLength = Len(sString)
  sString = Space$(iStringLength)

  If WinAPI_GetUserName(sString, iStringLength) Then

    i = InStr(sString, Chr(0))
    If i Then
      sString = Left(sString, i - 1)
    End If
    NetworkUserName = Trim(Left$(sString, iStringLength))
  Else
    NetworkUserName = "Usuario"
  End If

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

  CurMBTitle = Title

  If TimeOut = 0 Then
    TimeOut = 5 * 1000
  Else
    TimeOut = TimeOut * 1000
  End If

  TimerId = SetTimer(0, 0, TimeOut, AddressOf TimeOutMB)
  TimedMsgBox = MsgBox(Prompt, Icon, CurMBTitle)
  TimedMsgBox = 0
  KillTimer 0, TimerId

End Function

Private Sub TimeOutMB(hWnd As Long, uMsg As Long, idEvent As Long, dwTime As Long)

  SendMessage FindWindow(vbNullString, CurMBTitle), WM_CLOSE, 0&, 0&

End Sub
Function CheckRegEx(texto As String, padrao As String)

  Dim objRegExp As RegExp
  Dim objMatch As Match
  Dim colMatch As MatchCollection

  Dim valor As String

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
    ' valor = "Comparação falhou !"
  End If



End Function


