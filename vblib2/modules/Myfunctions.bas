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
    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    Public Declare Function ShellExecuteForExplore Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, lpParameters As Any, lpDirectory As Any, ByVal nShowCmd As Long) As Long
    
    Public Declare Function WinAPI_GetUserName Lib "Advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    Public Declare Function InternetGetConnectedState Lib "wininet" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
    Public Declare Function EbExecuteLine Lib "vba6.dll" (ByVal pStringToExec As Long, ByVal Unknownn1 As Long, ByVal Unknownn2 As Long, ByVal fCheckOnly As Long) As Long
    Public Declare Function ReleaseCapture Lib "user32" () As Long
    
    Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
    Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
    
    Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCTYPE As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
    Public Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCTYPE As Long, ByVal lpLCData As String) As Long
    
    Public Declare Function CharToOem Lib "user32" Alias "CharToOemA" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
    Public Declare Function OemToChar Lib "user32" (ByVal lpszSrc As String, ByVal lpszDst As String) As Long
    Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Long) As Long
#End If



Private Sub ForceSystemDecimalToPeriod()
    ' We MUST use the ANSI API version so it's an ANSI character that's used for the actual decimal character.
    Const LOCALE_SDECIMAL   As Long = &HE&
    Const LOCALE_SGROUPING  As Long = &H10&
    Const Eng_LCID          As Long = 1033&
    Dim s                   As String
    '
172:     s = String$(GetLocaleInfo(Eng_LCID, LOCALE_SDECIMAL, vbNullString, 0&), 0)
173:     GetLocaleInfo Eng_LCID, LOCALE_SDECIMAL, s, Len(s)
174:     If RTrimNull(s) <> "." Then
175:         SetLocaleInfo Eng_LCID, LOCALE_SDECIMAL, "."
176:         SetLocaleInfo Eng_LCID, LOCALE_SGROUPING, ","
177:     End If
End Sub

'Public Function tocar(cARQ)
'  PlaySound (cARQ), ByVal 0&, SND_FILENAME Or SND_ASYNC
'End Function

Public Function IsWebConnected(Optional ByRef ConnType As String) As Boolean
  Dim dwFlags As Long
  Dim WebTest As Boolean
187:   ConnType = ""
188:   WebTest = InternetGetConnectedState(CLng(dwFlags), 0&)
189:   Select Case WebTest
  Case dwFlags And CONNECT_LAN: ConnType = "LAN"
  Case dwFlags And CONNECT_MODEM: ConnType = "Modem"
  Case dwFlags And CONNECT_PROXY: ConnType = "Proxy"
  Case dwFlags And CONNECT_OFFLINE: ConnType = "Offline"
  Case dwFlags And CONNECT_CONFIGURED: ConnType = "Configured"
  'Case dwFlags And CONNECT_RAS: ConnType = "Remote"
196:   End Select
197:   IsWebConnected = WebTest
End Function

Public Function ExecutarDLL(ByVal cDLL As String) As Double
' appwiz.cpl adicionar e remover programas
' odbccp32.cpl gerenciador odbc
' sysdm.cpl informacoes do sistema
' outro cpl
' outra dll que possa ser executada via rundll32
'  Dim RetVal
207:   On Error Resume Next
208:   ExecutarDLL = shell("rundll32.exe shell32.dll,Control_RunDLL " & cDLL & ",,3", 1)
  
End Function

Public Function ExecuteLine(ByVal Scode As String, Optional ByVal fCheckOnly As Boolean) As Boolean
213:   ExecuteLine = EbExecuteLine(StrPtr(Scode), 0&, 0&, Abs(fCheckOnly)) = 0
End Function

Public Function ComboLostFocus(ByRef Combo1)
Dim strPartial
Dim i
219:   With Combo1
220:     If Len(.Text) Then
      'Procura pelo texto digitado
222:       strPartial = .Text
223:       i = SendMessage(.hWnd, CB_FINDSTRING, -1, ByVal strPartial)
      'Se não achou, retorna      o focus para o Combo
225:       If i = CB_ERR Then .SetFocus
226:     End If
227:   End With
End Function
Public Function pegue2delimitado(ByVal ctmpline As String, ByVal cDelimIni As String, ByVal cdelifim As String) As Variant
Dim nPOS As Integer
Dim nPOS2 As Integer
Dim nPOS3 As Integer
Dim cValor As String
Dim cINICIO As String
235: nPOS = 0
236: nPOS2 = 0
237: nPOS3 = 0
238: cValor = ""
239: cINICIO = ""
240: If Len(ctmpline) = 0 Then
241:    pegue2delimitado = Array("", "")
242:    Exit Function
243: End If
244: nPOS = InStr(ctmpline, cDelimIni)
245: If Len(cdelifim) = 0 Then
246: Else
247:   If nPOS > 0 Then
248:      nPOS2 = InStr(nPOS, ctmpline, cdelifim)
249:   Else
250:      nPOS2 = InStr(ctmpline, cdelifim)
251:   End If
252: End If
  
254:   If nPOS > 0 And nPOS2 = 0 Then
255:     If nPOS > 1 Then
256:       cINICIO = Mid(ctmpline, nPOS - 1)
257:     End If
258:     cValor = Mid(ctmpline, nPOS + Len(cDelimIni))
259:     ctmpline = ""
260:   End If
261:   If nPOS > 0 And nPOS2 > 0 Then
262:     cINICIO = Mid(ctmpline, 1, nPOS + Len(cDelimIni) - 1)
263:     cValor = Mid(ctmpline, nPOS + Len(cDelimIni))
264:     nPOS3 = InStr(cValor, cdelifim)
265:     If nPOS3 > 1 Then
266:        cValor = Mid(cValor, 1, nPOS3 - 1)
267:     End If
    
269:     If Right(cValor, Len(cdelifim)) = cdelifim Then
270:        cValor = Left(cValor, Len(cValor) - Len(cdelifim))
271:     End If
272:     ctmpline = Mid(ctmpline, nPOS2)
273:     If Left(ctmpline, Len(cdelifim)) = cdelifim Then
274:         ctmpline = Mid(ctmpline, Len(cdelifim) + 1)
275:     End If
276:     If ctmpline = cdelifim Then
277:        ctmpline = ""
278:     End If
279:   End If
280: pegue2delimitado = Array(cValor, ctmpline, cINICIO)
End Function
Public Function ComboChange(ByRef Combo1)
Dim strPartial
Dim i
Dim strTotal
Dim j
Dim m_bEditFromCode
288:   With Combo1
    'Procura pelo texto já digitado
290:     strPartial = .Text
    i = SendMessage(.hWnd, CB_FINDSTRING, -1, _
                    ByVal strPartial)

    'Se achou, adiciona o resto do Texto
295:     If i <> CB_ERR Then
      'Pega o texto inteiro
297:       strTotal = .List(i)

      'Compute number of unmatched characters
300:       j = Len(strTotal) - Len(strPartial)
301:       If j <> 0 Then
        'Adiciona o resto da string encontrada
303:         m_bEditFromCode = True
304:         .SelText = Right$(strTotal, j)

        'Marca os caracteres adicionados
307:         .SelStart = Len(strPartial)
308:         .SelLength = j
309:       End If
310:     End If
311:   End With
End Function

Public Sub MoveObject(ByRef Obj As Control)
315:   Screen.MousePointer = vbSizeAll
316:   SendMessage Obj.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 1
317:   ReleaseCapture
318:   Screen.MousePointer = vbDefault
End Sub

Public Function funNumeroPuro(ByVal pNumero) As String
  Dim i As Integer
323:   pNumero = FixStr(pNumero)
324:   funNumeroPuro = ""
325:   For i = 1 To Len(pNumero)
326:     If InStr("0123456789", Mid(pNumero, i, 1)) > 0 Then
327:       funNumeroPuro = funNumeroPuro & Mid(pNumero, i, 1)
328:     End If
329:   Next
End Function

Public Function FVar(ByVal eVAR As Variant, Optional ByVal cFORM As String = "", Optional ByVal ePAD As Variant)
333:   If cFORM <> "" Then
334:     Select Case cFORM
    Case "N"                                 'Numerico
336:       FVar = FixNum(eVAR, ePAD)
    Case "NI"                                'Numerico Inteiro
338:       FVar = FixInt(eVAR, ePAD)
    Case "C"                                 'Caracter
340:       FVar = LTrim(FixStr(eVAR, ePAD))
    Case "CU"                                'Caracter Maisculas
342:       FVar = UCase(FixStr(eVAR, ePAD))
    Case "CL"                                'Caracter Minusculas
344:       FVar = LCase(FixStr(eVAR, ePAD))
    Case "CC"                                'Caracter Capitalizado
346:       eVAR = FixStr(eVAR, ePAD)
347:       FVar = MMCase(eVAR)
    Case "D", "DN", "DS", "DC", "DF", "DH", "DD", "DZ", "D-"
349:       FVar = Fdata(eVAR, cFORM, ePAD)
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
360:       If VarType(eVAR) = vbBoolean Then
361:         FVar = eVAR
362:       Else
363:         If cFORM = "BF" Or cFORM = "B" Then
364:           FVar = False
365:         Else
366:           FVar = True                  'BT
367:         End If
368:       End If
    Case "BSN"                               'Converte boleano SN
370:       If VarType(eVAR) = vbBoolean Then
371:         If eVAR Then
372:           FVar = "S"
373:         Else
374:           FVar = "N"
375:         End If
376:       Else
377:         If IsNull(ePAD) Then
378:           FVar = " "
379:         Else
380:           FVar = ePAD
381:         End If
382:       End If
    Case "BN"                                'Converte Boleano em Numerico
384:       If VarType(eVAR) = vbBoolean Then
385:         If eVAR Then
386:           FVar = 1
387:         Else
388:           FVar = 0
389:         End If
390:       Else
391:         If IsNull(ePAD) Then
392:           FVar = 0
393:         Else
394:           FVar = ePAD
395:         End If
396:       End If
    Case "CPF", "CIC"                        '' 111.222.333-44   14 Digitos
398:       eVAR = TiraOut(eVAR)
      eVAR = Mid(eVAR, 1, 3) & "." & Mid(eVAR, 4, 3) & "." & _
             Mid(eVAR, 7, 3) & "-" & Mid(eVAR, 10, 2)
    Case "CNPJ", "CGC"                       '' 11.222.333/4444-55 18 digitos
402:       eVAR = TiraOut(eVAR)
      FVar = Mid(eVAR, 1, 2) & "." & Mid(eVAR, 3, 3) & "." & _
             Mid(eVAR, 6, 3) & "/" & Mid(eVAR, 9, 4) & "-" & Mid(eVAR, 13, 2)
    Case "RG"
406:       FVar = FormataRG(eVAR)
      '            Case "IE" Precisa estado
      '                FVar = FormataIE(eVAR)
    Case "CHAPA"
410:       eVAR = TiraOut(eVAR)
411:       FVar = Mid(eVAR, 1, 3) & "-" & Mid(eVAR, 4)
    Case "CEP"                               'Cep NNNNN-NNN
413:       eVAR = TiraOut(eVAR)
414:       FVar = Mid(eVAR, 1, 5) & "-" & Mid(eVAR, 6)
    Case "TEL"                               'Cep NNNN-NNNN
416:       eVAR = TiraOut(eVAR)
417:       FVar = Mid(eVAR, 1, 4) & "-" & Mid(eVAR, 5)
    Case "H"
419:       FVar = Format(eVAR, "HH.MM")
    Case "MES"                               'Numerico Mes do Ano 1-12
421:       FVar = FixInt(eVAR, ePAD)
422:       If FVar < 1 And FVar > 12 Then
423:         FVar = 0
424:       End If
    Case "ANO"                               'Numerico  Ano 1900-2300
426:       FVar = FixInt(eVAR, ePAD)
427:       If FVar < 1900 And FVar > 2300 Then
428:         FVar = 0
429:       End If

    Case Else                                'Retorna o Campo Sem Formataçao
432:       FVar = eVAR
433:     End Select
434:   Else
435:     FVar = eVAR
436:   End If
End Function

Public Function GeraSplit(ByVal aVAR As Variant, Optional ByVal cINI As String = "", _
                          Optional ByVal cMID As String = "", _
                          Optional ByVal cFIM As String = "") As String
  Dim nUSO As Long
  Dim x As Long
444:   nUSO = UBound(aVAR) - 1
445:   GeraSplit = cINI & aVAR(0) & cMID
446:   For x = 1 To nUSO
447:     GeraSplit = GeraSplit & aVAR(x) & cMID
448:   Next x
449:   GeraSplit = GeraSplit & aVAR(nUSO + 1)
450:   GeraSplit = GeraSplit & cFIM
End Function

Public Function SepSqlOpe(ByVal eEXP As String) As Variant
  Dim aRETU As Variant
455:   aRETU = Array("", "", "")
456:   If InStr(eEXP, ",") > 0 Then
457:     aRETU = Split(eEXP, ",")
458:   End If
459:   SepSqlOpe = aRETU
End Function
Public Function CharToLit(aVAL As Variant, Optional ByVal cTIPO As String = "", Optional ByVal cARQ As String = "") As String
Dim aRETU As Variant

464:   If Len(cARQ) > 0 Then
465:     cARQ = GeraConn(cARQ, cTIPO)
466:     aRETU = TipoConn(cARQ)
467:     cTIPO = aRETU(2)
468:   End If
469:   Select Case cTIPO
  Case "MYSQL", "MARIADB"
471:     CharToLit = GeraSplit(aVAL, "CONCAT(", ",", ")")
  Case Else
473:     CharToLit = GeraSplit(aVAL, "", " + ", "")
474:   End Select
475:   Exit Function
Erro:
477:   Select Case Err.Number
  Case Else
479:     SayErro "ChartoLit"
480:   End Select
End Function


Public Function Alert(ByVal cDIZ As String, Optional ByVal cTITLE As String = "Informacao")
485:   MsgBox cDIZ, vbOKOnly, cTITLE
End Function

Public Function Busca(ByVal cDIZ As String, ByVal cCAB As String, ByVal cVAL As String, Optional ByVal nLEN As Integer = 0) As String
489:   Busca = InputBox(cDIZ, cCAB, cVAL)
490:   If nLEN > 0 Then
491:      Busca = Left$(Busca, nLEN)
492:   End If
End Function

Public Function Caminex(ByVal cARQ As String, Optional ByVal nANO As Integer = 0, Optional ByVal nMES As Integer = 0, Optional ByVal nEMP = 1)
  Dim cAno As String
  Dim cEMP As String
  Dim cMes As String
  Dim cTMPVAR As String
  Dim nPOS  As Integer
  Dim nPOS2 As Integer
502:   nPOS = InStr(cARQ, "{")
503:   If nPOS > 0 Then
504:     nPOS2 = InStr(cARQ, "}")
505:     If nPOS2 > 0 Then
506:       cTMPVAR = Mid(cARQ, nPOS + 1, nPOS2 - nPOS - 1)
507:       cTMPVAR = Mid(cARQ, 1, nPOS - 1) + PegPath("PATH", cTMPVAR) + Mid(cARQ, nPOS2 + 1)
508:       cARQ = cTMPVAR
509:     End If
510:   End If

512:   If nANO = 0 And InStr(cARQ, "[A") > 0 Then
513:     frmcomp.Show vbModal
514:     nANO = Val(eRETU01)
515:     nMES = Val(eRETU02)
516:     nEMP = Val(eRETU03)
517:   End If

519:   If nMES = 0 And InStr(cARQ, "[M") > 0 Then
520:     frmcomp.Show vbModal
521:     nANO = Val(eRETU01)
522:     nMES = Val(eRETU02)
523:     nEMP = Val(eRETU03)
524:   End If

526:   If nEMP = 0 And InStr(cARQ, "[Z") > 0 Then
527:     frmcomp.Show vbModal
528:     nANO = Val(eRETU01)
529:     nMES = Val(eRETU02)
530:     nEMP = Val(eRETU03)
531:   End If

533:   If nANO = 0 And InStr(cARQ, "[A") > 0 Then
534:     frmcomp.Show vbModal
535:     nANO = Val(eRETU01)
536:     nMES = Val(eRETU02)
537:     nEMP = Val(eRETU03)
538:   End If

540:   If InStr(cARQ, "[MM]") > 0 Then
541:     cMes = StrZero(nMES, 2)
542:     cARQ = Replace(cARQ, "[MM]", cMes)
543:   End If
544:   If InStr(cARQ, "[AAAA]") > 0 Then
545:     cAno = StrZero(nANO, 4)
546:     cARQ = Replace(cARQ, "[AAAA]", cAno)
547:   End If
548:   If InStr(cARQ, "[AA]") > 0 Then
549:     cAno = StrZero(nANO, 4)
550:     cAno = Right(cAno, 2)
551:     cARQ = Replace(cARQ, "[AA]", cAno)
552:   End If
553:   If InStr(cARQ, "[ZZZ]") > 0 Then
554:     cEMP = StrZero(nEMP, 3)
555:     cARQ = Replace(cARQ, "[ZZZ]", cEMP)
556:   End If
557:   If InStr(cARQ, "[ZZ]") > 0 Then
558:     cEMP = StrZero(nEMP, 2)
559:     cARQ = Replace(cARQ, "[ZZ]", cEMP)
560:   End If
561:   If InStr(cARQ, "[Z]") > 0 Then
562:     cEMP = StrZero(nEMP, 1)
563:     cARQ = Replace(cARQ, "[Z]", cEMP)
564:   End If
565:   Caminex = cARQ
End Function
Public Function Dividir(ByVal nVAL As Variant, ByVal nDIV As Variant)
568:   Dividir = 0
569:   nVAL = FixNum(nVAL)
570:   nDIV = FixNum(nDIV)
571:   If nVAL > 0 And nDIV > 0 Then
572:     Dividir = nVAL / nDIV
573:   End If
  
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
604:   On Error GoTo NotExist
605:   FileConnExist = False
606:   cARQUIVO = UCase(Trim(CStr(cARQ)))
607:   If Len(cARQ) = 0 Then
608:      FileConnExist = False
609:      Exit Function
610:   End If

612:   If InStr(cARQUIVO, "[") > 0 Then
     
614:     If InStr(cARQUIVO, "[JET") > 0 Or InStr(cARQUIVO, "[SQLITE]") Or InStr(cARQUIVO, "[ACCDB") > 0 Then
615:       cARQUIVO = Replace(cARQUIVO, "[JET", "")
616:       cARQUIVO = Replace(cARQUIVO, "[SQLITE]", "")
617:       cARQUIVO = Replace(cARQUIVO, "[ACCDB", "")
618:       cARQUIVO = Replace(cARQUIVO, "MDB]", "")
619:       cARQUIVO = Replace(cARQUIVO, "]", "")
620:     Else
621:       If InStr(cARQUIVO, "[JETFOX]") > 0 Or InStr(cARQUIVO, "[ADS") > 0 Then
622:         cARQUIVO = Replace(cARQUIVO, "[JETFOX]", "")
623:         cARQUIVO = Replace(cARQUIVO, "[ADS", "")
624:         cARQUIVO = Replace(cARQUIVO, "CDX]", "")
625:         cARQUIVO = Replace(cARQUIVO, "NTX]", "")
626:         cARQUIVO = Replace(cARQUIVO, "ADT]", "")
627:         cARQUIVO = Replace(cARQUIVO, "]", "")
628:         If Len(cSQL) = 0 Then
629:           Alert ("Fileexit  Nao Passado cSQL Necessario Para DBF")
630:           FileConnExist = True
631:           Exit Function
632:         End If
633:         cARQUIVO = cARQUIVO & NomeTableSql(cSQL, ".DBF")
634:       End If

636:       If InStr(cARQUIVO, "PDX3]") > 0 Or InStr(cARQUIVO, "PDX4]") Or InStr(cARQUIVO, "PDX5]") > 0 Then
637:         cARQUIVO = Replace(cARQUIVO, "PDX3]", "")
638:         cARQUIVO = Replace(cARQUIVO, "PDX4]", "")
639:         cARQUIVO = Replace(cARQUIVO, "PDX5]", "")
640:         cARQUIVO = Replace(cARQUIVO, "[A16", "")
641:         cARQUIVO = Replace(cARQUIVO, "[A15", "")
642:         cARQUIVO = Replace(cARQUIVO, "[A14", "")
643:         cARQUIVO = Replace(cARQUIVO, "[A12", "")
644:         cARQUIVO = Replace(cARQUIVO, "[JET", "")
645:         cARQUIVO = Replace(cARQUIVO, "]", "")
646:         If Len(cSQL) = 0 Then
647:           Alert ("Fileexit  Nao Passado cSQL Necessario Para DB")
648:           FileConnExist = True
649:           Exit Function
650:         End If
651:         cARQUIVO = cARQUIVO & NomeTableSql(cSQL, ".DB")
652:       End If

      If InStr(cARQUIVO, "[CONN]") > 0 Or InStr(cARQUIVO, "[MYSQL]") > 0 Or InStr(cARQUIVO, "[MARIADB]") > 0 Or _
         InStr(cARQUIVO, "[PGSQL]") > 0 Or InStr(cARQUIVO, "[POSTGRESQL]") > 0 Or InStr(cARQUIVO, "[SQLSERVER]") > 0 Then
656:         FileConnExist = True
657:         Exit Function
658:       End If
659:       If InStr(cARQUIVO, "[FIREBIRD]") > 0 Or InStr(cARQUIVO, "[INTERBASE]") > 0 Then
660:         nPOS = InStr(cARQUIVO, "DBNAME=")
661:         cARQUIVO = Mid(cARQUIVO, nPOS + 7)
662:         nPOS = InStr(cARQUIVO, ";")
663:         If nPOS > 0 Then
664:           cARQUIVO = Mid(cARQUIVO, 1, nPOS - 1)
665:         End If
666:         FileConnExist = True
667:         Exit Function
668:       End If
669:       If InStr(cARQUIVO, "[XLSEXT]") > 0 Then
670:         cARQUIVO = Replace(cARQUIVO, "[XLSEXT]", "")
671:       End If
672:       If InStr(cARQUIVO, "[XLS]") > 0 Then
673:         cARQUIVO = Replace(cARQUIVO, "[XLS]", "")
674:       End If
675:     End If
676:   End If
677:   If InStr(cARQUIVO, "PROVIDER") > 0 Or InStr(cARQUIVO, ".MARIADB") > 0 Or InStr(cARQUIVO, ".MYSQL") > 0 Then
678:     FileConnExist = True
679:     Exit Function
680:   End If
681:   If FileExists(cARQUIVO) Then
    
683:     FileConnExist = True
684:   End If

686:   Exit Function
NotExist:
688:   FileConnExist = False
689:   If lMES Then
690:     Alert cMes & cARQUIVO, "Erro Arquivo"
691:   End If
End Function

Public Function FixInt(ByVal cUSO As Variant, Optional ByVal ePAD As Variant = 0)
695:   If IsNull(ePAD) Or ePAD = "" Then            'necessario pois pode ser passado outros tipos
696:     ePAD = 0
697:   End If
698:   If IsNull(cUSO) Or cUSO = "" Then
699:     FixInt = ePAD
700:     Exit Function
701:   End If
702:   If Not IsNumeric(cUSO) Then
703:     cUSO = funNumeroPuro(cUSO)
704:   End If
705:   If IsNull(cUSO) Or cUSO = "" Then            'necessario apos funnumeropuro
706:     FixInt = ePAD
707:     Exit Function
708:   End If
709:   FixInt = Int(cUSO)
End Function
Public Function senhapos(cPOSTELA)
  Dim cSUBTXT As String
  Dim nCHAR As Integer
714:   senhapos = ""
715:   While Len(cPOSTELA) > 2  'no fim pois postela  usa reverse string
716:     cSUBTXT = Mid(cPOSTELA, Len(cPOSTELA) - 2, 3)
717:     nCHAR = FixInt(cSUBTXT)
718:     nCHAR = nCHAR / 2
719:     nCHAR = nCHAR - 30

721:     senhapos = senhapos + Chr(nCHAR)
722:     cPOSTELA = Mid(cPOSTELA, 1, Len(cPOSTELA) - 3)
723:   Wend
  ''//Alert (senhapos)
End Function


Public Function FixNum(ByVal cUSO As Variant, Optional ByVal ePAD As Variant = 0) As Variant
  Dim nPOS01 As Integer
  Dim nPOS02 As Integer
731:   If IsNull(ePAD) Then
732:     ePAD = 0
733:   End If
734:   If IsNull(cUSO) Or cUSO = "" Then
    ''cUSO = ePAD
736:     FixNum = ePAD
737:     Exit Function
738:   End If
739:   If Not IsNumeric(cUSO) Then
740:     cUSO = CStr(cUSO)
741:     nPOS01 = InStr(cUSO, ",")
742:     nPOS02 = InStr(cUSO, ".")
743:     If nPOS02 = 0 And nPOS01 > 0 Then        'Somente Virgula 99,99
744:       cUSO = Replace(cUSO, ",", ".")       ''muda virgula 99.99
745:     End If
746:     If nPOS01 > 0 And nPOS02 > 0 Then
747:       If nPOS02 < nPOS01 Then              'Milhar Ponto Decimal Virgula 99.999,99
748:         cUSO = Replace(cUSO, ".", "")    ''Tira os Pontos 99999,99
749:         cUSO = Replace(cUSO, ",", ".")   ''muda virgula  99999.99
750:       Else                                 'Milhar virgula ponto Decimal  99,999.99
751:         cUSO = Replace(cUSO, ",", "")    ''tira virgula
752:       End If
753:     End If
754:   End If
755:   FixNum = CDbl(cUSO)
End Function

Public Function FixStr(ByVal eVAR As Variant, _
                       Optional ByVal ePAD As Variant = "", _
                       Optional ByVal coper As String = "", _
                       Optional ByVal nLEN As Integer = 0) As Variant
762:   On Error GoTo errhandler
763:   If IsNull(eVAR) Then
764:     If ePAD <> "" Then
765:       eVAR = ePAD
766:     End If
767:   End If
768:   If Not IsNull(eVAR) Then
769:      FixStr = CStr(eVAR)
770:   End If
771:   If UCase(eVAR) = "NULL" Then
772:     FixStr = ePAD
773:   End If
774:   If InStr(coper, "TRIM") > 0 Then
775:     eVAR = Trim(eVAR)
776:     FixStr = eVAR
777:   End If
778:   If nLEN > 0 And Len(eVAR) > nLEN Then
779:     eVAR = Mid(eVAR, 1, nLEN)
780:     FixStr = eVAR
781:   End If
782:   Exit Function
errhandler:
784:   FixStr = ""
785:   Resume Next
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
799:     If FixNum(vValor) And FixNum(nIncremento) Then
        
801:         vValor = vValor + nIncremento
        
        ' Lógica de limite (se aplicável)
804:         If Not IsMissing(nLimiteMax) Then
805:             If vValor > nLimiteMax Then vValor = nLimiteMax
806:         End If
        
808:         If Not IsMissing(nLimiteMin) Then
809:             If vValor < nLimiteMin Then vValor = nLimiteMin
810:         End If
        
812:         INCDEC = vValor
813:     Else
814:         INCDEC = vValor
815:     End If
    
End Function

Public Function MDG(ByVal cMEnSSAGEM As String, Optional cTITULO = "Confirme")
  Dim eRESP As Variant
821:   MDG = False
822:   eRESP = MsgBox(cMEnSSAGEM, vbYesNo + vbDefaultButton1, cTITULO)
823:   If eRESP = vbYes Then
824:     MDG = True
825:   End If
End Function

Public Function MathOper(ByVal nVAL01 As Variant, ByVal nVAL02 As Variant, _
                         Optional ByVal cOPE As String = "&", _
                         Optional ByVal nROUND As Integer = -1)
  Dim eVAL As Variant
  Dim eTMP As Variant
833:   eVAL = 0
834:   Select Case cOPE
  Case "ASCII"
836:     If Len(nVAL01) > 0 Then
837:       eVAL = Asc(nVAL01)
838:     Else
839:       eVAL = 32
840:     End If
  Case "EXT"
842:     eVAL = UCase(EXTENSAO(FixStr(nVAL01, FixStr(nVAL02))))
  Case "NPD"
844:     eTMP = FixInt(nVAL01)
845:     If eTMP > 0 Then
846:       eVAL = Fdata(nVAL02)
847:     Else
848:       eVAL = Space(8)
849:     End If
  Case "PCF"
851:     If FixStr(nVAL01) = "PC" Then
852:       eVAL = Format(nVAL02, "###,###,##0")
853:     Else
854:       eVAL = Format(nVAL02, "###,###,##0.00")
855:     End If
  Case "*"
857:     eVAL = Multiplicar(nVAL01, nVAL02)
  Case "/"
859:     eVAL = Dividir(nVAL01, nVAL02)
  Case "%"
861:     eVAL = Dividir(nVAL01, nVAL02)
862:     eVAL = Multiplicar(eVAL, 100)
  Case "+"
864:     eVAL = FixNum(nVAL01) + FixNum(nVAL02)
  Case "CNPJCPF"
866:     Select Case nVAL01
    Case "J"
868:       eVAL = FVar(nVAL02, "CNPJ")
    Case "F"
870:       eVAL = FVar(nVAL02, "CPF")
    Case Else
872:       eVAL = nVAL02
873:     End Select
  Case "-"
875:     eVAL = FixNum(nVAL01) - FixNum(nVAL02)
  Case "&"
877:     eVAL = Concatenar(nVAL01, nVAL02, "")
  Case "& ", "& / ", "&/", "&-", "& - "
879:     eVAL = Concatenar(nVAL01, nVAL02, cOPE)
  Case Else                                    ''Operacao Nao Definida cONCATENA
881:     eVAL = FixStr(nVAL01) & FixStr(nVAL02)
882:   End Select
883:   If nROUND > 0 Then
884:     eVAL = Round(eVAL, nROUND)
885:   End If
886:   MathOper = eVAL
End Function

Public Function Concatenar(ByVal eVAR01 As Variant, ByVal eVAR02, Optional ByVal cSEP As String = "") As String
890:   cSEP = Replace(cSEP, "&", "")
891:   Concatenar = FixStr(eVAR01) & cSEP & FixStr(eVAR02)
End Function

Public Function Multiplicar(ByVal nVAL As Variant, ByVal nMUL As Variant)
895:   Multiplicar = 0
896:   nVAL = FixNum(nVAL)
897:   nMUL = FixNum(nMUL)
898:   If nVAL > 0 And nMUL > 0 Then
899:     Multiplicar = nVAL * nMUL
900:   End If
End Function


Public Function PadRight(ByVal cTEXTO, ByVal nLEN) As String
905:   cTEXTO = cTEXTO & Space(nLEN)
906:   cTEXTO = Left(cTEXTO, nLEN)
907:   PadRight = cTEXTO
End Function

Public Function PadLeft(ByVal cTEXTO, ByVal nLEN) As String
911:   cTEXTO = Space(nLEN) & cTEXTO
912:   cTEXTO = Right(cTEXTO, nLEN)
913:   PadLeft = cTEXTO
End Function

Public Function PegCamini(ByVal cCaminho As String) As String
917:   PegCamini = Caminex(cCaminho, 0, 0, 0)
End Function
Public Function PegINIVAL(ByVal cARQINI As String, ByVal cGRUPO As String, ByVal cCAMPO As String, Optional ByVal ePAD As String = "") As String
  Dim z As Long
  Dim scaminho As String * 255
922:   z = GetPrivateProfileString(cGRUPO, cCAMPO, "", scaminho, 150, cARQINI)
923:   PegINIVAL = IIf(Asc(Left(scaminho, 1)) = "0", ePAD, Left(scaminho, z))

End Function

Public Function PegPath(ByVal cGRUPO As String, ByVal cCAMPO As String, Optional ByVal ePAD As String = "" _
                       , Optional ByVal cARQINI As String = "") As String
  Dim z As Long
  Dim scaminho As String * 255
931:   If cARQINI = "" Then
932:      cARQINI = App.Path & "\" & App.EXEName & ".INI"
933:   End If
934:   z = GetPrivateProfileString(cGRUPO, cCAMPO, "", scaminho, 150, cARQINI)
935:   If Len(ePAD) = 0 Then
936:     ePAD = App.Path & "\"
937:   End If
938:   PegPath = IIf(Asc(Left(scaminho, 1)) = "0", ePAD, Left(scaminho, z))

End Function

Public Function PegTable(ByVal cGRUPO As String, ByVal cCAMPO As String)
  Dim z As Long
  Dim scaminho As String * 255
  Dim cTMP As String
946:   z = GetPrivateProfileString(cGRUPO, cCAMPO, "", scaminho, 150, App.Path + "\" & App.EXEName & ".INI")
947:   cTMP = IIf(Asc(Left(scaminho, 1)) = "0", App.Path & "\", Left(scaminho, z))
948:   PegTable = PegPath("PATH", cTMP)
End Function

Public Function MMCase(ByVal texto As String) As String
  Dim sPalavra As String
  Dim iPosIni As Integer
  Dim iPosFim As Integer
  Dim sResultado As String
956:   iPosIni = 1
957:   texto = LCase(texto) & " "
958:   sResultado = ""
959:   Do Until InStr(iPosIni, texto, " ") = 0
960:     iPosFim = InStr(iPosIni, texto, " ")
961:     sPalavra = Mid(texto, iPosIni, iPosFim - iPosIni)
962:     iPosIni = iPosFim + 1
    If sPalavra <> "de" And sPalavra <> "da" And _
       sPalavra <> "do" And sPalavra <> "das" _
       And sPalavra <> "dos" And sPalavra <> _
       "a" And sPalavra <> "e" Then
      sPalavra = UCase(Left(sPalavra, 1)) & _
                 LCase(Mid(sPalavra, 2))
969:     End If
970:     sResultado = sResultado & " " & sPalavra
971:   Loop
972:   MMCase = Trim(sResultado)
End Function

Public Sub SayErro(Optional ByVal cERROUSO As String = "", Optional ByVal lMES As Boolean = True)
  Dim nHANDLE As Long
  Dim cARQ As String
  Dim cERRO As String
     Dim osver
980:   cERRO = ""
981:   On Error Resume Next
982:   If Err.Number <> 0 Then
983:     Screen.MousePointer = vbDefault
984:     cARQ = zUSER & Format(Now, "_DDMMYYYY_HHMMSS") & ".TXT"

    cERRO = " No.        : " & Err.Number & Chr(13) & Chr(10) & _
          " Fonte      : " & Err.Source & Chr(13) & Chr(10) & _
          " Descricao  : " & Err.Description & Chr(13) & Chr(10) & _
          " Aplicativo : " & App.EXEName & " " & App.Major & "." & App.Minor & ". " & App.Revision & Chr(13) & Chr(10)
    '           " Linha      : " & Erl & Chr(13) & Chr(10) & _


993:                 If Not IsNull(Err.LastDllError) And Err.LastDllError <> 0 Then
994:     cERRO = cERRO & " DLL        : " & Err.LastDllError & Chr(13) & Chr(10)
995:   End If
996:   If IsObject(Screen.ActiveForm) Then
    '' cERRO = cERRO & " Formulario : " & Screen.ActiveForm.Name
    '' If Not IsNull(Screen.ActiveForm.HelpContextID) Then
    ''    cERRO = cERRO & " ID         : " & Screen.ActiveForm.HelpContextID
    '' End If
1001:   End If

1003:   If IsObject(Screen.ActiveControl) Then
    ''cERRO = cERRO & " Em Uso     : " & Screen.ActiveControl
1005:   End If


  'cERRO = cERRO & " Equipamento: " & MachineName & vbCrLf 'abaixo classe
1009:   cERRO = cERRO & " Usuario    : " & NetworkUserName & vbCrLf
  'cERRO = cERRO & " IP         : " & GetIPAddress & vbCrLf  'abaixo classe
  'cERRO = cERRO & " Local      : " & GetIPHostName & vbCrLf 'abaixo classe

1013:   cERRO = cERRO & " Erro       : " & cERROUSO
1014:   If lMES Then
1015:     TimedMsgBox cERRO, , , "Mensagem Sistema"  ''usuario ficavam pendurados travando registros agora com timeout
    'Alert cERRO, "Informacao de Erro "
1017:   End If


1020:   Select Case Err.Number
  Case -2147483647
1022:     cERRO = cERRO & vbCrLf & "not implemented"
  Case -2147483646
1024:     cERRO = cERRO & vbCrLf & "ran out of memory"
  Case -2147483645
1026:     cERRO = cERRO & vbCrLf & "one or more arguments are invalid"
  Case -2147483644
1028:     cERRO = cERRO & vbCrLf & "no such interface supported"
  Case -2147483643
1030:     cERRO = cERRO & vbCrLf & "invalid pointer"
  Case -2147483642
1032:     cERRO = cERRO & vbCrLf & "invalid handle"
  Case -2147483641
1034:     cERRO = cERRO & vbCrLf & "operation aborted"
  Case -2147483640
1036:     cERRO = cERRO & vbCrLf & "unspecified error"
  Case -2147483639
1038:     cERRO = cERRO & vbCrLf & "general access denied error"
  Case -2147483638
1040:     cERRO = cERRO & vbCrLf & "the data necessary to complete this operation is not yet available."
  Case -2147467263
1042:     cERRO = cERRO & vbCrLf & "not implemented"
  Case -2147467262
1044:     cERRO = cERRO & vbCrLf & "no such interface supported"
  Case -2147467261
1046:     cERRO = cERRO & vbCrLf & "invalid pointer"
  Case -2147467260
1048:     cERRO = cERRO & vbCrLf & "operation aborted"
  Case -2147467259
1050:     cERRO = cERRO & vbCrLf & "unspecified error"
  Case -2147467258
1052:     cERRO = cERRO & vbCrLf & "thread local storage failure"
  Case -2147467257
1054:     cERRO = cERRO & vbCrLf & "get shared memory allocator failure"
  Case -2147467256
1056:     cERRO = cERRO & vbCrLf & "get memory allocator failure"
  Case -2147467255
1058:     cERRO = cERRO & vbCrLf & "unable to initialize class cache"
  Case -2147467254
1060:     cERRO = cERRO & vbCrLf & "unable to initialize rpc services"
  Case -2147467253
1062:     cERRO = cERRO & vbCrLf & "cannot set thread local storage channel control"
  Case -2147467252
1064:     cERRO = cERRO & vbCrLf & "could not allocate thread local storage channel control"
  Case -2147467251
1066:     cERRO = cERRO & vbCrLf & "the user supplied memory allocator is unacceptable"
  Case -2147467250
1068:     cERRO = cERRO & vbCrLf & "the ole service mutex already exists"
  Case -2147467249
1070:     cERRO = cERRO & vbCrLf & "the ole service file mapping already exists"
  Case -2147467248
1072:     cERRO = cERRO & vbCrLf & "unable to map view of file for ole service"
  Case -2147467247
1074:     cERRO = cERRO & vbCrLf & "failure attempting to launch ole service"
  Case -2147467246
1076:     cERRO = cERRO & vbCrLf & "there was an attempt to call coinitialize a second time while single threaded"
  Case -2147467245
1078:     cERRO = cERRO & vbCrLf & "a remote activation was necessary but was not allowed"
  Case -2147467244
1080:     cERRO = cERRO & vbCrLf & "a remote activation was necessary but the server name provided was invalid"
  Case -2147467243
1082:     cERRO = cERRO & vbCrLf & "the class is configured to run as a security id different from the caller"
  Case -2147467242
1084:     cERRO = cERRO & vbCrLf & "use of ole1 services requiring dde windows is disabled"
  Case -2147467241
1086:     cERRO = cERRO & vbCrLf & "a runas specification must be \ or simply 0x80004018 -2147467240 the server process could not be started. the pathname may be incorrect."
  Case -2147467239
1088:     cERRO = cERRO & vbCrLf & "the server process could not be started as the configured identity. the pathname may be incorrect or unavailable."
  Case -2147467238
1090:     cERRO = cERRO & vbCrLf & "the server process could not be started because the configured identity is incorrect. check the username and password."
  Case -2147467237
1092:     cERRO = cERRO & vbCrLf & "the client is not allowed to launch this server."
  Case -2147467236
1094:     cERRO = cERRO & vbCrLf & "the service providing this server could not be started."
  Case -2147467235
1096:     cERRO = cERRO & vbCrLf & "this computer was unable to communicate with the computer providing the server."
  Case -2147467234
1098:     cERRO = cERRO & vbCrLf & "the server did not respond after being launched."
  Case -2147467233
1100:     cERRO = cERRO & vbCrLf & "the registration information for this server is inconsistent or incomplete."
  Case -2147467232
1102:     cERRO = cERRO & vbCrLf & "the registration information for this interface is inconsistent or incomplete."
  Case -2147467231
1104:     cERRO = cERRO & vbCrLf & "the operation attempted is not supported."
  Case -2147418113
1106:     cERRO = cERRO & vbCrLf & "catastrophic failure"
  Case -2147024891
1108:     cERRO = cERRO & vbCrLf & "general access denied error"
  Case -2147024890
1110:     cERRO = cERRO & vbCrLf & "invalid handle"
  Case -2147024882
1112:     cERRO = cERRO & vbCrLf & "ran out of memory"
  Case -2147024809
1114:     cERRO = cERRO & vbCrLf & "one or more arguments are invalid"
  Case -2147217920
1116:     cERRO = cERRO & vbCrLf & "invalid accessor"
  Case -2147217919
1118:     cERRO = cERRO & vbCrLf & "creating another row would have exceeded the total number of active rows supported by the rowset"
  Case -2147217918
1120:     cERRO = cERRO & vbCrLf & "unable to write with a read-only accessor"
  Case -2147217917
1122:     cERRO = cERRO & vbCrLf & "given values violate the database schema"
  Case -2147217916
1124:     cERRO = cERRO & vbCrLf & "invalid row handle"
  Case -2147217915
1126:     cERRO = cERRO & vbCrLf & "an object was open"
  Case -2147217914
1128:     cERRO = cERRO & vbCrLf & "invalid chapter"
  Case -2147217913
1130:     cERRO = cERRO & vbCrLf & "a literal value in the command could not be converted to the correct type due to a reason other than data overflow"
  Case -2147217912
1132:     cERRO = cERRO & vbCrLf & "invalid binding info"
  Case -2147217911
1134:     cERRO = cERRO & vbCrLf & "permission denied"
  Case -2147217910
1136:     cERRO = cERRO & vbCrLf & "specified column does not contain bookmarks or chapters"
  Case -2147217909
1138:     cERRO = cERRO & vbCrLf & "some cost limits were rejected"
  Case -2147217908
1140:     cERRO = cERRO & vbCrLf & "no command has been set for the command object"
  Case -2147217907
1142:     cERRO = cERRO & vbCrLf & "unable to find a query plan within the given cost limit"
  Case -2147217906
1144:     cERRO = cERRO & vbCrLf & "invalid bookmark"
  Case -2147217905
1146:     cERRO = cERRO & vbCrLf & "invalid lock mode"
  Case -2147217904
1148:     cERRO = cERRO & vbCrLf & "no value given for one or more required parameters"
  Case -2147217903
1150:     cERRO = cERRO & vbCrLf & "invalid column id"
  Case -2147217902
1152:     cERRO = cERRO & vbCrLf & "invalid ratio"
  Case -2147217901
1154:     cERRO = cERRO & vbCrLf & "invalid value"
  Case -2147217900
1156:     cERRO = cERRO & vbCrLf & "the command contained one or more errors"
  Case -2147217899
1158:     cERRO = cERRO & vbCrLf & "the executing command cannot be canceled"
  Case -2147217898
1160:     cERRO = cERRO & vbCrLf & "the provider does not support the specified dialect"
  Case -2147217897
1162:     cERRO = cERRO & vbCrLf & "a data source with the specified name already exists"
  Case -2147217896
1164:     cERRO = cERRO & vbCrLf & "the rowset was built over a live data feed and cannot be restarted"
  Case -2147217895
1166:     cERRO = cERRO & vbCrLf & "no key matching the described characteristics could be found within the current range"
  Case -2147217894
1168:     cERRO = cERRO & vbCrLf & "ownership of this tree has been given to the provider"
  Case -2147217893
1170:     cERRO = cERRO & vbCrLf & "the provider is unable to determine identity for newly inserted rows"
  Case -2147217892
1172:     cERRO = cERRO & vbCrLf & "no nonzero weights specified for any goals supported, so goal was rejected; current goal was not changed"
  Case -2147217891
1174:     cERRO = cERRO & vbCrLf & "requested conversion is not supported"
  Case -2147217890
1176:     cERRO = cERRO & vbCrLf & "lrowsoffset would position you past either end of the rowset, regardless of the crows value specified; crowsobtained is 0"
  Case -2147217889
1178:     cERRO = cERRO & vbCrLf & "information was requested for a query, and the query was not set"
  Case -2147217888
1180:     cERRO = cERRO & vbCrLf & "provider called a method from irowsetnotify in the consumer and nt"
  Case -2147217887
1182:     cERRO = cERRO & vbCrLf & "errors occurred"
  Case -2147217886
1184:     cERRO = cERRO & vbCrLf & "a non-null controlling iunknown was specified and the object being created does not support aggregation"
  Case -2147217885
1186:     cERRO = cERRO & vbCrLf & "a given hrow referred to a hard- or soft- deleted row"
  Case -2147217884
1188:     cERRO = cERRO & vbCrLf & "the rowset does not support fetching backwards"
  Case -2147217883
1190:     cERRO = cERRO & vbCrLf & "all hrows must be released before new ones can be obtained"
  Case -2147217882
1192:     cERRO = cERRO & vbCrLf & "one of the specified storage flags was not supported"
  Case -2147217881
1194:     cERRO = cERRO & vbCrLf & "the comparison operator was invalid"
  Case -2147217880
1196:     cERRO = cERRO & vbCrLf & "the specified status flag was neither dbcolumnstatus_ok nor dbcolumnstatus_isnull"
  Case -2147217879
1198:     cERRO = cERRO & vbCrLf & "the rowset cannot scroll backwards"
  Case -2147217878
1200:     cERRO = cERRO & vbCrLf & "invalid region handle"
  Case -2147217877
1202:     cERRO = cERRO & vbCrLf & "the specified set of rows was not contiguous to or overlapping the rows in the specified watch region"
  Case -2147217876
1204:     cERRO = cERRO & vbCrLf & "a transition from all* to move* or extend* was specified"
  Case -2147217875
1206:     cERRO = cERRO & vbCrLf & "the specified region is not a proper subregion of the region identified by the given watch region handle"
  Case -2147217874
1208:     cERRO = cERRO & vbCrLf & "the provider does not support multi-statement commands"
  Case -2147217873
1210:     cERRO = cERRO & vbCrLf & "a specified value violated the integrity constraints for a column or table"
  Case -2147217872
1212:     cERRO = cERRO & vbCrLf & "the given type name was unrecognized"
  Case -2147217871
1214:     cERRO = cERRO & vbCrLf & "execution aborted because a resource limit has been reached; no results have been returned"
  Case -2147217870
1216:     cERRO = cERRO & vbCrLf & "cannot clone a command object whose command tree contains a rowset or rowsets"
  Case -2147217869
1218:     cERRO = cERRO & vbCrLf & "cannot represent the current tree as text"
  Case -2147217868
1220:     cERRO = cERRO & vbCrLf & "the specified index already exists"
  Case -2147217867
1222:     cERRO = cERRO & vbCrLf & "the specified index does not exist"
  Case -2147217866
1224:     cERRO = cERRO & vbCrLf & "the specified index was in use"
  Case -2147217865
1226:     cERRO = cERRO & vbCrLf & "the specified table does not exist"
  Case -2147217864
1228:     cERRO = cERRO & vbCrLf & "the rowset was using optimistic concurrency and the value of a column has been changed since it was last read"
  Case -2147217863
1230:     cERRO = cERRO & vbCrLf & "errors were detected during the copy"
  Case -2147217862
1232:     cERRO = cERRO & vbCrLf & "a specified precision was invalid"
  Case -2147217861
1234:     cERRO = cERRO & vbCrLf & "a specified scale was invalid"
  Case -2147217860
1236:     cERRO = cERRO & vbCrLf & "invalid table id"
  Case -2147217859
1238:     cERRO = cERRO & vbCrLf & "a specified type was invalid"
  Case -2147217858
1240:     cERRO = cERRO & vbCrLf & "a column id was occurred more than once in the specification"
  Case -2147217857
1242:     cERRO = cERRO & vbCrLf & "the specified table already exists"
  Case -2147217856
1244:     cERRO = cERRO & vbCrLf & "the specified table was in use"
  Case -2147217855
1246:     cERRO = cERRO & vbCrLf & "the specified locale id was not supported"
  Case -2147217854
1248:     cERRO = cERRO & vbCrLf & "the specified record number is invalid"
  Case -2147217853
1250:     cERRO = cERRO & vbCrLf & "although the bookmark was validly formed, no row could be found to match it"
  Case -2147217852
1252:     cERRO = cERRO & vbCrLf & "the value of a property was invalid"
  Case -2147217851
1254:     cERRO = cERRO & vbCrLf & "the rowset was not chaptered"
  Case -2147217850
1256:     cERRO = cERRO & vbCrLf & "invalid accessor"
  Case -2147217849
1258:     cERRO = cERRO & vbCrLf & "invalid storage flags"
  Case -2147217848
1260:     cERRO = cERRO & vbCrLf & "by-ref accessors are not supported by this provider"
  Case -2147217847
1262:     cERRO = cERRO & vbCrLf & "null accessors are not supported by this provider"
  Case -2147217846
1264:     cERRO = cERRO & vbCrLf & "the command was not prepared"
  Case -2147217845
1266:     cERRO = cERRO & vbCrLf & "the specified accessor was not a parameter accessor"
  Case -2147217844
1268:     cERRO = cERRO & vbCrLf & "the given accessor was write-only"
  Case -2147217843
1270:     cERRO = cERRO & vbCrLf & "authentication failed"
  Case -2147217842
1272:     cERRO = cERRO & vbCrLf & "the change was canceled during notification; no columns are changed"
  Case -2147217841
1274:     cERRO = cERRO & vbCrLf & "the rowset was single-chaptered and the chapter was not released"
  Case -2147217840
1276:     cERRO = cERRO & vbCrLf & "invalid source handle"
  Case -2147217839
1278:     cERRO = cERRO & vbCrLf & "the provider cannot derive parameter info and setparameterinfo has not been called"
  Case -2147217838
1280:     cERRO = cERRO & vbCrLf & "the data source object is already initialized"
  Case -2147217837
1282:     cERRO = cERRO & vbCrLf & "the provider does not support this method"
  Case -2147217836
1284:     cERRO = cERRO & vbCrLf & "the number of rows with pending changes has exceeded the set limit"
  Case -2147217835
1286:     cERRO = cERRO & vbCrLf & "the specified column did not exist"
  Case -2147217834
1288:     cERRO = cERRO & vbCrLf & "there are pending changes on a row with a reference count of zero"
  Case -2147217833
1290:     cERRO = cERRO & vbCrLf & "a literal value in the command overflowed the range of the type of the associated column"
  Case -2147217832
1292:     cERRO = cERRO & vbCrLf & "the supplied hresult was invalid"
  Case -2147217831
1294:     cERRO = cERRO & vbCrLf & "the supplied lookupid was invalid"
  Case -2147217830
1296:     cERRO = cERRO & vbCrLf & "the supplied dynamicerrorid was invalid"
  Case -2147217829
1298:     cERRO = cERRO & vbCrLf & "unable to get visible data for a newly-inserted row that has not yet been updated"
  Case -2147217828
1300:     cERRO = cERRO & vbCrLf & "invalid conversion flag"
  Case -2147217827
1302:     cERRO = cERRO & vbCrLf & "the given parameter name was unrecognized"
  Case -2147217826
1304:     cERRO = cERRO & vbCrLf & "multiple storage objects can not be open simultaneously"
  Case -2147217825
1306:     cERRO = cERRO & vbCrLf & "the requested filter could not be opened"
  Case -2147217824
1308:     cERRO = cERRO & vbCrLf & "the requested order could not be opened"
  Case -2147217823
1310:     cERRO = cERRO & vbCrLf & "bad tuple"
  Case -2147217822
1312:     cERRO = cERRO & vbCrLf & "bad coordinate"
  Case -2147217821
1314:     cERRO = cERRO & vbCrLf & "the given axis was not valid for this dataset"
  Case -2147217820
1316:     cERRO = cERRO & vbCrLf & "one or more of the given cell ordinals was invalid"
  Case -2147217819
1318:     cERRO = cERRO & vbCrLf & "the supplied columnid was invalid"
  Case -2147217817
1320:     cERRO = cERRO & vbCrLf & "the supplied command does not have a dbid"
  Case -2147217816
1322:     cERRO = cERRO & vbCrLf & "the supplied dbid already exists"
  Case -2147217815
1324:     cERRO = cERRO & vbCrLf & "the maximum number of sessions supported by the provider has already been created. the consumer must release one or more currently held sessions before obtaining a new session object"
  Case -2147217806
1326:     cERRO = cERRO & vbCrLf & "the index id is invalid"
  Case -2147217805
1328:     cERRO = cERRO & vbCrLf & "the initialization string does not conform to specification"
  Case -2147217804
1330:     cERRO = cERRO & vbCrLf & "the ole db root enumerator did not return any providers that matched an of the sources_types requested"
  Case -2147217803
1332:     cERRO = cERRO & vbCrLf & "the initialization string specifies a provider which does not match the currently active provider."
  Case -2147217802
1334:     cERRO = cERRO & vbCrLf & "the specified dbid is invalid"
  Case -2147217814
1336:     cERRO = cERRO & vbCrLf & "invalid trustee value"
  Case -2147217813
1338:     cERRO = cERRO & vbCrLf & "the trustee is not for the current data source"
  Case -2147217812
1340:     cERRO = cERRO & vbCrLf & "the trustee does not support memberships/ collections"
  Case -2147217811
1342:     cERRO = cERRO & vbCrLf & "the object is invalid or unknown to the provider"
  Case -2147217810
1344:     cERRO = cERRO & vbCrLf & "no owner exists for the object"
  Case -2147217809
1346:     cERRO = cERRO & vbCrLf & "the access entry list supplied is invalid"
  Case -2147217808
1348:     cERRO = cERRO & vbCrLf & "the trustee supplied as owner is invalid or unknown to the provider"
  Case -2147217807
1350:     cERRO = cERRO & vbCrLf & "the permission supplied in the access entry list is invalid"
  Case -2147217801
1352:     cERRO = cERRO & vbCrLf & "the constrainttype was invalid or not supported by the provider."
  Case -2147217800
1354:     cERRO = cERRO & vbCrLf & "the constrainttype was not constrainttype_foreignkey and cforeignkeycolumns was not zero"
  Case -2147217799
1356:     cERRO = cERRO & vbCrLf & "the deferrability was invalid or the value was not supported by the provider"
  Case -2147217792
1358:     cERRO = cERRO & vbCrLf & "the matchtype was invalid or the value was not supported by the provider"
  Case -2147217782
1360:     cERRO = cERRO & vbCrLf & "the updaterule or deleterule was invalid or the value was not supported by the provider"
  Case -2147217781
1362:     cERRO = cERRO & vbCrLf & "the pconstraintid did not exist in the data source"
  Case -2147217780
1364:     cERRO = cERRO & vbCrLf & "the dwflags was invalid"
  Case -2147217779
1366:     cERRO = cERRO & vbCrLf & "the rguidcolumntype pointed to a guid that does not match the object type of this column or this column was not set"
  Case -2147217778
1368:     cERRO = cERRO & vbCrLf & "the requested url was out-of-scope"
  Case -2147217776
1370:     cERRO = cERRO & vbCrLf & "the provider could not drop the object"
  Case -2147217775
1372:     cERRO = cERRO & vbCrLf & "there is no source row"
  Case -2147217774
1374:     cERRO = cERRO & vbCrLf & "the ole db object represented by this url is locked by one or more other processes"
  Case -2147217773
1376:     cERRO = cERRO & vbCrLf & "the client requested an object type that is only valid for a collection"
  Case -2147217772
1378:     cERRO = cERRO & vbCrLf & "the caller requested write access to a read-only object"
  Case -2147217771
1380:     cERRO = cERRO & vbCrLf & "the provider could not connect to the server for this object"
  Case -2147217770
1382:     cERRO = cERRO & vbCrLf & "the provider could not connect to the server for this object"
  Case -2147217769
1384:     cERRO = cERRO & vbCrLf & "the attempt to bind to the object timed out"
  Case -2147217768
1386:     cERRO = cERRO & vbCrLf & "the provider was unable to create an object at this url because an object named by this url already exists"
  Case -2147217767
1388:     cERRO = cERRO & vbCrLf & "the provider could not drop the object"
  Case -2147217766
1390:     cERRO = cERRO & vbCrLf & "the provider was unable to create an object at this url because the server was out of physical storage"
1391:   End Select

1393:  cERRO = cERRO & vbCrLf & infosistema

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
  


1428:   nHANDLE = FreeFile
1429:   Open cARQ For Output As #nHANDLE
1430:   Print #nHANDLE, cERRO
1431:   Close (nHANDLE)

1433: End If
End Sub

Public Function infosistema()
Dim osver As clsOSInfo
 

1440:     Set osver = New clsOSInfo

    Dim s As String
1443:     With osver
    
    
      ' S = S & vbCrLf & "IP         : " & .IPAddress
      ' S = S & vbCrLf & "Local      : " & .IPHostName
1448:         s = "Equipamento: " & .ComputerName
1449:         s = s & vbCrLf & "Usuario: " & .UserName
1450:         s = s & vbCrLf & "Grupo do Usuario: " & .UserType
1451:         s = s & vbCrLf & "Is in Admin group? " & .IsAdminGroup
1452:         s = s & vbCrLf & "OS Name: " & .OSName
1453:         s = s & vbCrLf & "Service Pack ver.: " & .SPVer
1454:         s = s & vbCrLf & "Is Server? " & .IsServer
1455:         s = s & vbCrLf & "Bitness: " & .Bitness
1456:         s = s & vbCrLf & "Is Win x64: " & .IsWin64
1457:         s = s & vbCrLf & "Is Win x32: " & .IsWin32
1458:         s = s & vbCrLf & "Edition: " & .Edition
1459:         s = s & vbCrLf & "Suite mask: " & .SuiteMask
1460:         s = s & vbCrLf & "ProductType: " & .ProductType
1461:         s = s & vbCrLf & "PlatformID: " & .PlatformID & " (" & .Platform & ")"
1462:         s = s & vbCrLf & "Is Domain controller: " & .IsDomainController
1463:         s = s & vbCrLf & "Is Embedded: " & .IsEmbedded
1464:         s = s & vbCrLf & "Language in dialogues: " & .LangDisplayCode & " " & .LangDisplayName & " " & .LangDisplayNameFull
1465:         s = s & vbCrLf & "Language of OS inslallation: " & .LangSystemCode & " " & .LangSystemName & " " & .LangSystemNameFull
1466:         s = s & vbCrLf & "Language for non-Unicode programs: " & .LangNonUnicodeCode & " " & .LangNonUnicodeName & " " & .LangNonUnicodeNameFull
1467:         s = s & vbCrLf & "ID of default locale: " & .LCID_UserDefault
1468:         s = s & vbCrLf & "File System Case sensitive? " & .IsFileSystemCaseSensitive
1469:         s = s & vbCrLf & "OEM Codepage: " & .CodepageOEM & " (" & .CodepageOEM_File & ")"
1470:         s = s & vbCrLf & "ANSI Codepage: " & .CodepageANSI & " (" & .CodepageANSI_File & ")"

        
1473:         s = s & vbCrLf & "Major: " & .Major
1474:         s = s & vbCrLf & "Minor: " & .Minor
1475:         s = s & vbCrLf & "Major + Minor:         " & .MajorMinor
1476:         s = s & vbCrLf & "Major + Minor (NtDll): " & .MajorMinorNTDLL
1477:         s = s & vbCrLf & "Build: " & .Build
1478:         s = s & vbCrLf & "NT Dll Major.Minor.Rev: " & .NtDllVersion
1479:         s = s & vbCrLf & "Revision: " & .Revision
1480:         s = s & vbCrLf & "ReleaseId: " & .ReleaseId
1481:         s = s & vbCrLf & "DisplayVersion: " & .DisplayVersion
1482:         s = s & vbCrLf & "Process integrity level: " & .IntegrityLevel
1483:         s = s & vbCrLf & "Elevated process? " & .IsElevated
1484:         s = s & vbCrLf & "Is Local system context? " & .IsLocalSystemContext
1485:         s = s & vbCrLf & "User sid of current process owner: " & .SID_CurrentProcess
1486:         s = s & vbCrLf & "Safe boot? " & .IsSafeBoot & " (" & .SafeBootMode & ")"
1487:         s = s & vbCrLf & "Secure Boot supported? " & .SecureBootSupported & " (Enabled? " & .SecureBoot & ")"
1488:         s = s & vbCrLf & "TestSigning: " & .TestSigning
1489:         s = s & vbCrLf & "DebugMode: " & .DebugMode
1490:         s = s & vbCrLf & "CodeIntegrity: " & .CodeIntegrity
1491:         s = s & vbCrLf & "Memory MiB (Free/Total): " & .MemoryFree & "/" & .MemoryTotal & " (Loaded: " & .MemoryLoad & "%)"
        
1493:               s = s & vbCrLf & "OS - XP/Server 2003(R2)? " & .IsWindowsXP
1494:         s = s & vbCrLf & "OS - Vista/Server 2008? " & .IsWindowsVista
1495:         s = s & vbCrLf & "OS - 7/Server 2008R2? " & .IsWindows7
1496:         s = s & vbCrLf & "OS - 8/Server 2012? " & .IsWindows8
1497:         s = s & vbCrLf & "OS - 8.1/Server 2012R2? " & .IsWindows8OrGreater
1498:         s = s & vbCrLf & "OS - 10/Server 2016? " & .IsWindows10
1499:         s = s & vbCrLf & "OS - XP or newer? " & .IsWindowsXPOrGreater
1500:         s = s & vbCrLf & "OS - XP SP3 or newer? " & .IsWindowsXP_SP3OrGreater
1501:         s = s & vbCrLf & "OS - Vista or newer? " & .IsWindowsVistaOrGreater
1502:         s = s & vbCrLf & "OS - 7 or newer? " & .IsWindows7OrGreater
1503:         s = s & vbCrLf & "OS - 8 or newer? " & .IsWindows8OrGreater
1504:         s = s & vbCrLf & "OS - 8.1 or newer? " & .IsWindows8Point1OrGreater
1505:         s = s & vbCrLf & "OS - 10 or newer? " & .IsWindows10OrGreater
1506:         s = s & vbCrLf & "OS - 11 or newer? " & .IsWindows11OrGreater
  
        'Debug.Print s
1509:         infosistema = s
1510:     End With
    
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
1527:   If (InStr(UCase$(sFile), ".EXE") <> 0) Then
1528:     eShowCmd = 0
1529:   End If
1530:   On Error Resume Next
1531:   If (sParameters = "") And (sDefaultDir = "") Then
1532:     lr = ShellExecuteForExplore(Owner, sOperation, sFile, 0, 0, essSW_SHOWNORMAL)
1533:   Else
1534:     lr = ShellExecute(Owner, sOperation, sFile, sParameters, sDefaultDir, eShowCmd)
1535:   End If
1536:   If (lr < 0) Or (lr > 32) Then
1537:     ShellEx = True
1538:   Else
    ' raise an appropriate error:
1540:     lErr = vbObjectError + 1048 + lr
1541:     Select Case lr
    Case 0
1543:       lErr = 7: sErr = "Sem Memoria"
  '  Case ERROR  'ERROR_FILE_NOT_FOUND
  '    lErr = 53: sErr = "Arquivo nao encontrado"
  '  Case ERROR_PATH_NOT_FOUND
  '    lErr = 76: sErr = "Caminho nao encontrado"
  '  Case ERROR_BAD_FORMAT
  '    sErr = "Executavel invalido ou corrompido"
    Case SE_ERR_ACCESSDENIED
1551:       lErr = 75: sErr = "Caminho ou arquivo sem acesso "
    Case SE_ERR_ASSOCINCOMPLETE
1553:       sErr = "Executavel nao associado a esta extensao."
    Case SE_ERR_DDEBUSY
1555:       lErr = 285: sErr = "O Executavel nao pode ser aberto no momento, tente outra vez."
    Case SE_ERR_DDEFAIL
1557:       lErr = 285: sErr = "Falha na transacao DDE, tente outra vez."
    Case SE_ERR_DDETIMEOUT
1559:       lErr = 286: sErr = "Tempo esgotado, tente outra vez."
    Case SE_ERR_DLLNOTFOUND
1561:       lErr = 48: sErr = "Biblioteca nao encotrada."
    Case SE_ERR_FNF
1563:       lErr = 53: sErr = "Arquivo nao encontrado"
    Case SE_ERR_NOASSOC
1565:       sErr = "Executavel nao associado a esta extensao."
    Case SE_ERR_OOM
1567:       lErr = 7: sErr = "Sem memoria"
    Case SE_ERR_PNF
1569:       lErr = 76: sErr = "Caminho nao encontrado"
    Case SE_ERR_SHARE
1571:       lErr = 75: sErr = "Violacao de Acesso."
    Case Else
1573:       sErr = "Erro nao especificado."
1574:     End Select

  '  Err.Raise lErr, , App.EXEName & ".GShell", sErr
1577:     ShellEx = False
1578:   End If

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
1591:     If Not IsArray(vArray) Then Exit Function
    
    Dim vItem As Variant
    Dim dSoma As Double
    
    ' 2. O twinBASIC e o VB6 lidam com 'For Each' em arrays de variantes
    ' de forma idêntica em 32 e 64 bits. É o método mais seguro.
1598:     For Each vItem In vArray
        ' Usamos sua FixNum para garantir que só somamos o que é numérico
1600:         If FixNum(vItem) Then
1601:             dSoma = dSoma + FixNum(vItem) 'CDbl(vItem)
1602:         End If
1603:     Next vItem
    
1605:     SomaArr = dSoma
End Function


Public Function SomaExt(ByVal cARQ As String, Optional ByVal cEXT As String = ".MDB") As String
1610:   If InStr(UCase(cARQ), cEXT) > 0 Then
1611:     SomaExt = cARQ
1612:   Else
1613:     SomaExt = cARQ & cEXT
1614:   End If
End Function

Public Function StrZero(ByVal nNUM, Optional ByVal nLEN As Integer = 0)
  Dim cTemp As String
1619:   If nLEN = 0 Then
1620:     cTemp = FixStr(nNUM, "0", "TRIM")
1621:     nLEN = Trim(cTemp)
1622:   End If
1623:   StrZero = Right(String(nLEN, "0") & CStr(nNUM), nLEN)
End Function

Function Convert2ansi(ByVal in_string As String) As String
  Dim Out_String As String * 256
  Dim t As String
1629:   t = OemToChar(in_string, Out_String)
1630:   Convert2ansi = Out_String
End Function
Function Convert2oem(ByVal in_string As String) As String
  Dim Out_String As String * 256
  Dim t As String
1635:   t = CharToOem(in_string, Out_String)
1636:   Convert2oem = Out_String
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
1654:   Tirace = tirace2(texto)
End Function
Public Function StrToArray(ByVal cGRUPO As String) As Variant
  Dim x As Integer
  Dim nLEN As Integer
  Dim aUSO As Variant
  Dim cCHAR As String
  Dim eCNV As String
  ' utilizando agora Convert2ansi Convert2oem tirace2
1663:   Select Case cGRUPO
  Case Else
1665:     eCNV = cGRUPO
1666:   End Select

1668:   nLEN = Len(eCNV)
  ReDim aUSO(nLEN)
1670:   For x = 1 To nLEN
1671:     cCHAR = Mid(eCNV, x, 1)
1672:     Select Case cCHAR
    Case "ª"
1674:       aUSO(x - 1) = "a."

    Case "º"
1677:       aUSO(x - 1) = "o."

    Case Else
1680:       aUSO(x - 1) = cCHAR

1682:     End Select
1683:   Next x
1684:   StrToArray = aUSO
End Function

Public Function TiraSin(ByVal texto As String, Optional ByVal RemoveUp As Boolean = True) As String
Dim x As Integer
  'https://www.ascii-code.com/pt
  'Caracteres de controle ASCII (código de caractere 0-31)
1691:   For x = 0 To 31 'nao caracteres padrao
1692:     texto = Replace(texto, Chr(x), "")
1693:   Next x
  'Caracteres imprimiveis ASCII (código de caractere 32-127)
  '32 espaco
1696:   For x = 33 To 38                 '33! 34" 35# 36$ 37% 38&
1697:     texto = Replace(texto, Chr(x), "")
1698:   Next x
1699:   For x = 39 To 47                 '39' 40( 41) 42* 43+ 44, 45- 46. 47/
1700:     texto = Replace(texto, Chr(x), "")
1701:   Next x
  '48 a 57 numeros 0-9
1703:   For x = 58 To 64                 '58: 59: 60< 61= 62> 63? 64@
1704:     texto = Replace(texto, Chr(x), "")
1705:   Next x
  '65 a 90 maisculas
1707:   For x = 91 To 96                 '91[ 92\ 93] 94^ 95_ 96`
1708:     texto = Replace(texto, Chr(x), "")
1709:   Next x
  '97 a 122 minusculas
1711:   For x = 123 To 126                '123{ 124| 125} 126~
1712:     texto = Replace(texto, Chr(x), "")
1713:   Next x
1714:   For x = 127 To 127                 '127 del
1715:     texto = Replace(texto, Chr(x), "")
1716:   Next x
  'Códigos ASCII estendidos (código de caracteres 128-255)
1718:   If RemoveUp Then
1719:     For x = 128 To 255
1720:       texto = Replace(texto, Chr(x), "")
1721:     Next x
1722:   Else 'matem caracteres de acentuacao uso tirace caso queira manter sem acentuacao
1723:     For x = 155 To 159
1724:       texto = Replace(texto, Chr(x), "")
1725:     Next x
1726:     For x = 168 To 180
1727:       texto = Replace(texto, Chr(x), "")
1728:     Next x
1729:     For x = 184 To 197
1730:       texto = Replace(texto, Chr(x), "")
1731:     Next x
1732:     For x = 200 To 209
1733:       texto = Replace(texto, Chr(x), "")
1734:     Next x
1735:     For x = 217 To 223
1736:       texto = Replace(texto, Chr(x), "")
1737:     Next x
1738:     For x = 238 To 255
1739:       texto = Replace(texto, Chr(x), "")
1740:     Next x
1741:   End If
1742:   TiraSin = texto
End Function
Function CheckPass(ByVal cTEXTO As String, Optional ByVal lMES As Boolean = True) As Boolean

  Dim lMAIS As Boolean
  Dim lMINUS As Boolean
  Dim lDIG As Boolean
  Dim lSYMBOL As Boolean
  Dim l8DIG As Boolean
  Dim i As Integer

1753:   CheckPass = False
1754:   lMAIS = False
1755:   lMINUS = False
1756:   lDIG = False
1757:   lSYMBOL = False
1758:   l8DIG = False


1761:   For i = 1 To Len(cTEXTO)
1762:     If InStr("0123456789", Mid(cTEXTO, i, 1)) > 0 Then
1763:       lDIG = True
1764:     End If
1765:     If InStr("abcdefghijklmnopqrstuvwxyz", Mid(cTEXTO, i, 1)) > 0 Then
1766:       lMINUS = True
1767:     End If
1768:     If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", Mid(cTEXTO, i, 1)) > 0 Then
1769:       lMAIS = True
1770:     End If
1771:     If InStr("-+_!@#$%^&*., ?", Mid(cTEXTO, i, 1)) > 0 Then
1772:       lSYMBOL = True
1773:     End If
1774:   Next

1776:   If Len(Trim(cTEXTO)) >= 8 Then
1777:     l8DIG = True
1778:   End If


1781:   If lMAIS And lMINUS And lDIG And lSYMBOL And l8DIG Then
1782:     CheckPass = True
1783:   Else
1784:     If lMES Then
1785:       If Not lMAIS Then
1786:         Alert (" Sem maisculas")
1787:       End If
1788:       If Not lMINUS Then
1789:         Alert (" Sem minusclas")
1790:       End If
1791:       If Not lDIG Then
1792:         Alert (" Sem numeros")
1793:       End If
1794:       If Not lSYMBOL Then
1795:         Alert (" Sem simbulos")
1796:       End If
1797:       If Not l8DIG Then
1798:         Alert ("Minimo 8 Digitos")
1799:       End If
1800:     End If
1801:   End If

End Function

Function isnumber(ByVal cLETRA As String) As Boolean
1806:    isnumber = InStr("0123456789", cLETRA) > 0
End Function

Function isstringupper(ByVal cLETRA As String) As Boolean
1810:    isstringupper = InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", cLETRA) > 0
End Function

Function isstringlower(ByVal cLETRA As String) As Boolean
1814:    isstringlower = InStr("abcdefghijklmnopqrstuvwxyz", cLETRA) > 0
End Function

Function issimbolo(ByVal cLETRA As String) As Boolean
1818:    issimbolo = InStr(",.:;$@&*!?<>%+-*/\|#=:_{}[]()'^~´`", cLETRA) > 0
End Function

Function isstring(ByVal cLETRA As String) As Boolean
1822:     isstring = isstringlower(cLETRA) Or isstringupper(cLETRA)
End Function
Public Function tirace2(ByVal cXml As String) As String
  Dim nAscii As Integer
  Dim ltroca As Boolean
  Dim nCont As Integer
  Dim cLETRA As String

1830:   For nCont = 1 To 2
1831:     cXml = Replace(cXml, Chr(26), "")
1832:     cXml = Replace(cXml, Chr(13), "")
1833:     cXml = Replace(cXml, Chr(10), "")
1834:     If InStr(Chr(239) + Chr(187) + Chr(191), Mid(cXml, 1, 1)) Then
1835:       cXml = Mid(cXml, 2)
1836:     End If
1837:     cXml = Replace(cXml, " />", "/>")
1838:     cXml = Replace(cXml, Chr(195) + Chr(173), "i")
1839:     cXml = Replace(cXml, Chr(195) + Chr(135), "C")
1840:     cXml = Replace(cXml, Chr(195) + Chr(141), "I")
1841:     cXml = Replace(cXml, Chr(195) + Chr(163), "a")
1842:     cXml = Replace(cXml, Chr(195) + Chr(167), "c")
1843:     cXml = Replace(cXml, Chr(195) + Chr(161), "a")
1844:     cXml = Replace(cXml, Chr(195) + Chr(131), "A")
1845:     cXml = Replace(cXml, Chr(194) + Chr(186), "o.")
1846:     cXml = Replace(cXml, Chr(195) + Chr(162), "a")
1847:     cXml = Replace(cXml, Chr(195) + Chr(161), "a")
1848:     cXml = Replace(cXml, Chr(195) + Chr(163), "a")
1849:     cXml = Replace(cXml, Chr(195) + Chr(173), "i")
1850:     cXml = Replace(cXml, Chr(195) + Chr(179), "o")
1851:     cXml = Replace(cXml, Chr(195) + Chr(167), "c")
1852:     cXml = Replace(cXml, Chr(195) + Chr(169), "e")
1853:     cXml = Replace(cXml, Chr(195) + Chr(170), "e")
1854:     cXml = Replace(cXml, Chr(195) + Chr(181), "o")
1855:     cXml = Replace(cXml, Chr(195) + Chr(160), "o")
1856:     cXml = Replace(cXml, Chr(195) + Chr(181), "o")
1857:     cXml = Replace(cXml, Chr(195) + Chr(129), "A")
1858:     cXml = Replace(cXml, Chr(226) + Chr(128) + Chr(156), "*")   '// aspas de destaque "cames"
1859:     cXml = Replace(cXml, Chr(226) + Chr(128) + Chr(157), "*")   '// aspas de destaque "cames"
1860:     cXml = Replace(cXml, Chr(195) + Chr(180), "o")
1861:     cXml = Replace(cXml, Chr(195) + Chr(186), "u")
1862:     cXml = Replace(cXml, Chr(195) + Chr(147), "O")
1863:     cXml = Replace(cXml, Chr(226) + Chr(128) + Chr(153), " ")   ' // caixa d'agua
1864:     cXml = Replace(cXml, Chr(226) + Chr(128) + Chr(147), "-")   '// - mesmo
1865:     cXml = Replace(cXml, Chr(194) + Chr(179), "3")   '// m3
    '// so pra corrigir no MySql
1867:     cXml = Replace(cXml, "+" + Chr(129), "A")
1868:     cXml = Replace(cXml, "+" + Chr(137), "E")
1869:     cXml = Replace(cXml, "+" + Chr(131), "A")
1870:     cXml = Replace(cXml, "+" + Chr(135), "C")
1871:     cXml = Replace(cXml, "?" + Chr(167), "c")
1872:     cXml = Replace(cXml, "?" + Chr(163), "a")
1873:     cXml = Replace(cXml, "?" + Chr(173), "i")
1874:     cXml = Replace(cXml, "?" + Chr(131), "A")
1875:     cXml = Replace(cXml, "?" + Chr(161), "a")
1876:     cXml = Replace(cXml, "?" + Chr(141), "I")
1877:     cXml = Replace(cXml, "?" + Chr(135), "C")
1878:     cXml = Replace(cXml, Chr(195) + Chr(156), "a")
1879:     cXml = Replace(cXml, Chr(195) + Chr(159), "A")
1880:     cXml = Replace(cXml, "?" + Chr(129), "A")
1881:     cXml = Replace(cXml, "?" + Chr(137), "E")
1882:     cXml = Replace(cXml, Chr(195) + "?", "C")
1883:     cXml = Replace(cXml, "?" + Chr(149), "O")
1884:     cXml = Replace(cXml, "?" + Chr(154), "U")
1885:     cXml = Replace(cXml, "+" + Chr(170), "o")
1886:     cXml = Replace(cXml, "?" + Chr(128), "A")
1887:     cXml = Replace(cXml, Chr(195) + Chr(166), "e")
1888:     cXml = Replace(cXml, Chr(135) + Chr(227), "ca")
1889:     cXml = Replace(cXml, "n" + Chr(227), "na")
1890:     cXml = Replace(cXml, Chr(162), "o")
1891:     cXml = Replace(cXml, " " + Chr(241) + " ", " ")
1892:     cXml = Replace(cXml, Chr(176), "")   ' graus
1893:     cXml = Replace(cXml, Chr(186), "o")   ' numero
1894:     cXml = Replace(cXml, Chr(220), "U")   ' u com trema
1895:     cXml = Replace(cXml, Chr(170), "")   ' desconhecido
1896:   Next
1897:   For nCont = 1 To Len(cXml)
1898:     cLETRA = Mid(cXml, nCont, 1)
1899:     nAscii = Asc(cLETRA)
1900:     ltroca = True
1901:     If InStr("0123456789", cLETRA) > 0 Then
1902:       ltroca = False
1903:     End If
1904:     If InStr("abcdefghijklmnopqrstuvwxyz", cLETRA) > 0 Then
1905:       ltroca = False
1906:     End If
1907:     If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", cLETRA) > 0 Then
1908:       ltroca = False
1909:     End If
1910:     If InStr(",.:/;%*$@?<>()+-#=:_", cLETRA) > 0 Then
1911:       ltroca = False
1912:     End If
1913:     If nAscii = 231 Then
1914:       cLETRA = "c"
1915:     End If
1916:     If nAscii = 199 Then
1917:       cLETRA = "C"
1918:     End If
1919:     If nAscii = 193 Or nAscii = 194 Or nAscii = 195 Or nAscii = 192 Then
1920:       cLETRA = "A"
1921:     End If
1922:     If nAscii = 224 Or nAscii = 225 Or nAscii = 226 Or nAscii = 227 Or nAscii = 228 Or nAscii = 229 Then
1923:       cLETRA = "a"
1924:     End If
1925:     If nAscii = 242 Or nAscii = 243 Or nAscii = 244 Or nAscii = 245 Or nAscii = 246 Then
1926:       cLETRA = "o"
1927:     End If
1928:     If nAscii = 210 Or nAscii = 211 Or nAscii = 212 Or nAscii = 213 Or nAscii = 214 Then
1929:       cLETRA = "O"
1930:     End If
1931:     If nAscii = 200 Or nAscii = 201 Or nAscii = 202 Or nAscii = 203 Then
1932:       cLETRA = "E"
1933:     End If
1934:     If nAscii = 232 Or nAscii = 233 Or nAscii = 234 Or nAscii = 235 Then
1935:       cLETRA = "e"
1936:     End If
1937:     If nAscii = 236 Or nAscii = 237 Or nAscii = 238 Or nAscii = 239 Then
1938:       cLETRA = "i"
1939:     End If
1940:     If nAscii = 204 Or nAscii = 205 Or nAscii = 206 Or nAscii = 207 Then
1941:       cLETRA = "I"
1942:     End If
1943:     If nAscii = 249 Or nAscii = 250 Or nAscii = 251 Or nAscii = 252 Then
1944:       cLETRA = "u"
1945:     End If
1946:     If nAscii = 217 Or nAscii = 218 Or nAscii = 219 Then
1947:       cLETRA = "U"
1948:     End If
1949:     If nAscii = 128 Then
1950:       cLETRA = "C"
1951:     End If
1952:     If nAscii = 144 Then
1953:       cLETRA = "E"
1954:     End If
1955:     If nAscii = 248 Then
1956:       cLETRA = ""
1957:     End If
1958:     If nAscii = 167 Then
1959:       cLETRA = "o"
1960:     End If
1961:     If ltroca Then
1962:       cXml = Mid(cXml, 1, nCont - 1) + cLETRA + Mid(cXml, nCont + 1)
1963:     End If
1964:   Next
1965:   tirace2 = cXml
End Function



Public Function XOREncryption(strCodeKey As String, _
                              strDataIn As String) As String

  Dim lonDataPtr As Long
  Dim intXORValue1 As Integer
  Dim intXORValue2 As Integer
  Dim strDataOut As String
  Dim nLENINI As Integer
  Dim cTMPCHAR As String


1981:   strDataOut = ""
1982:   nLENINI = Len(strDataIn)


1985:   For lonDataPtr = 1 To Len(strDataIn)
1986:     intXORValue1 = Asc(Mid$(strDataIn, lonDataPtr, 1))
    intXORValue2 = Asc(Mid$(strCodeKey, _
                            ((lonDataPtr Mod Len(strCodeKey)) + 1), 1))
1989:     cTMPCHAR = Chr(intXORValue1 Xor intXORValue2)
1990:     If Asc(cTMPCHAR) = 0 Or Asc(cTMPCHAR) = 32 Then
1991:       cTMPCHAR = "_"
1992:     End If
1993:     strDataOut = strDataOut + cTMPCHAR

1995:   Next lonDataPtr



1999:   XOREncryption = Mid$(strDataOut, 1, nLENINI)

End Function

Public Sub FocusMe()
  If TypeOf Screen.ActiveControl Is TextBox _
     Or TypeOf Screen.ActiveControl Is ComboBox _
     Or TypeOf Screen.ActiveControl Is XPText Then
2007:     Screen.ActiveControl.SelStart = 0
2008:     Screen.ActiveControl.SelLength = Len(Trim(Screen.ActiveControl.Text))
2009:   End If
End Sub
Public Function CharConv(ByVal cTEXTO As String, ByVal eORI As Variant, ByVal eDES As Variant) As String
  Dim nLEN As Integer
  Dim nTEXTO As Integer
  Dim x As Integer
  Dim y As Integer
  Dim aORI As Variant
  Dim aDES As Variant
  Dim aTEXTO As Variant
2019:   If IsArray(eORI) Then
2020:     aORI = eORI
2021:     aDES = eDES
2022:   Else
2023:     aORI = StrToArray(CStr(eORI))
2024:     aDES = StrToArray(CStr(eDES))
2025:   End If
2026:   aTEXTO = StrToArray(cTEXTO)
2027:   nLEN = UBound(aORI)
2028:   nTEXTO = UBound(aTEXTO)
2029:   For y = 0 To nTEXTO
2030:     For x = 0 To nLEN
2031:       If aTEXTO(y) = aORI(x) Then          ''Encerra Analise Para Evitar
2032:         aTEXTO(y) = aDES(x)              ''Loop de Troca
2033:         Exit For
2034:       End If
2035:     Next
2036:   Next
2037:   CharConv = ""
2038:   For y = 0 To nTEXTO
2039:     CharConv = CharConv & aTEXTO(y)
2040:   Next y
End Function
Public Function TiraOut(ByVal eVAR As Variant) As String
  Dim cTEXTO As String
2044:   cTEXTO = FixStr(eVAR)
  TiraOut = CharConv(cTEXTO, Array("-", ",", ".", ":", "/", ";", "*", "(", ")"), _
                     Array("", "", "", "", "", "", "", "", ""))
End Function

Public Function TiraOutNum(ByVal eVAR As Variant) As String  ' Mantem ', . que sao usadas nos numeros
  Dim cTEXTO As String
2051:   cTEXTO = FixStr(eVAR)
  TiraOutNum = CharConv(cTEXTO, Array("-", ":", "/", ";", "*", "(", ")"), _
                        Array("", "", "", "", "", "", ""))
End Function

Public Function TiraOutAlf(ByVal eVAR As Variant) As String
  Dim cTEXTO As String
2058:   cTEXTO = FixStr(eVAR)
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

2083:   Negativo = (valor < 0)
2084:   valor = Abs(CDec(valor))
2085:   If valor Then
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

    StrValor = Left(Format(valor, String(18, "0") & _
                                  ".000"), 18)
2110:     For Posicao = 1 To 18 Step 3
2111:       Parcial = Val(Mid(StrValor, Posicao, 3))
2112:       If Parcial Then
2113:         If Parcial = 1 Then
          buf = "Um" & PotenciasSingular((18 - _
                                          Posicao) \ 3)
2116:         ElseIf Parcial = 100 Then
          buf = "Cem" & PotenciasSingular((18 - _
                                           Posicao) \ 3)
2119:         Else
2120:           buf = Centenas(Parcial \ 100)
2121:           Parcial = Parcial Mod 100
2122:           If Parcial <> 0 And buf <> vbNullString Then
2123:             buf = buf & " e "
2124:           End If
2125:           If Parcial < 20 Then
2126:             buf = buf & Unidades(Parcial)
2127:           Else
2128:             buf = buf & Dezenas(Parcial \ 10)
2129:             Parcial = Parcial Mod 10
2130:             If Parcial <> 0 And buf <> vbNullString Then
2131:               buf = buf & " e "
2132:             End If
2133:             buf = buf & Unidades(Parcial)
2134:           End If
2135:           buf = buf & PotenciasPlural((18 - Posicao) \ 3)
2136:         End If
2137:         If buf <> vbNullString Then
2138:           If Extenso <> vbNullString Then
2139:             Parcial = Val(Mid(StrValor, Posicao, 3))
            If Posicao = 16 And (Parcial < 100 Or _
                                 (Parcial Mod 100) = 0) Then
2142:               Extenso = Extenso & " e "
2143:             Else
2144:               Extenso = Extenso & ", "
2145:             End If
2146:           End If
2147:           Extenso = Extenso & buf
2148:         End If
2149:       End If
2150:     Next
2151:     If Extenso <> vbNullString Then
2152:       If Negativo Then
2153:         Extenso = "Menos " & Extenso
2154:       End If
2155:       If Int(valor) = 1 Then
2156:         Extenso = Extenso & " " & MoedaSingular
2157:       Else
2158:         Extenso = Extenso & " " & MoedaPlural
2159:       End If
2160:     End If
    Parcial = Int((valor - Int(valor)) * _
                  100 + 0.1)
2163:     If Parcial Then
      buf = Extenso(Parcial, "Centavos", _
                    "Centavo")
2166:       If Extenso <> vbNullString Then
2167:         Extenso = Extenso & " e "
2168:       End If
2169:       Extenso = Extenso & buf
2170:     End If
2171:   End If
End Function

Public Function Txt2Lin(ByVal cTEXTO As String, Optional ByVal nCOL As Integer = 80) As Variant
  Dim nLIN, x As Integer
  Dim aRETU As Variant
2177:   cTEXTO = FixStr(cTEXTO)
2178:   If nCOL < 1 Then nCOL = 80                   'Evita Erros Divisao
2179:   nLIN = Int(Len(cTEXTO) / nCOL)
2180:   If nLIN * nCOL = Len(cTEXTO) Then            ''Bate com o Multiplo
2181:   Else
2182:     nLIN = nLIN + 1                          ''Soma mais um pois e necesario
2183:   End If
  ReDim aRETU(nLIN)
2185:   For x = 1 To nLIN
2186:     aRETU(x - 1) = Mid(cTEXTO, ((x - 1) * nCOL) + 1, nCOL)
2187:   Next x
2188:   Txt2Lin = aRETU
2189:   eRETU01 = nLIN
End Function

Public Function TrimNull(ByVal sTxt As String) As String
  Dim Arr() As String
2194:   Arr() = Split(sTxt, Chr$(0))
2195:   If UBound(Arr) >= 0 Then
2196:     TrimNull = Arr(0)
2197:   Else
2198:     TrimNull = sTxt
2199:   End If
End Function
Public Sub EnableControls(ByVal frm As Form, ByVal enabled_state As Boolean)
  Dim ctl As Control
  ' Examine every control.
2204:   For Each ctl In frm.Controls
2205:     On Error Resume Next
2206:     ctl.Enabled = enabled_state
2207:     On Error GoTo 0
2208:   Next ctl
End Sub


Public Function FixBolNum(ByVal lbol As Boolean) As Integer  'variant
2213:   FixBolNum = 0
2214:   If lbol Then
2215:     FixBolNum = 1
2216:   End If
End Function

Public Function FixNumBol(ByVal nBOL As Integer) As Boolean
2220:   FixNumBol = False
2221:   If nBOL = 1 Then
2222:     FixNumBol = True
2223:   End If
End Function

Public Function FixStrBol(ByVal cCHAR As String) As Boolean
2227:  FixStrBol = StrLogic(cCHAR)
End Function

' +--------------------------------------------------------------------
' +  Função: StrLogic
' +  Objetivo: Converte uma string alfanumérica em valor lógico booleano,
' +            espelhando a mesma inteligência usada no Harbour.
' +  Retorno:  Variant (Pode retornar True, False, ou Null para nulos)
' +--------------------------------------------------------------------
Public Function StrLogic(ByVal cVAL As String) As Variant
    ' Limpa espaços em branco e padroniza em maiúsculo
2238:     cVAL = UCase$(Trim$(cVAL))
    
2240:     Select Case cVAL
        ' Casos que retornam Nulo/Vazio
        Case "", "NULL", "<NULL>", "NIL", "NUL"
2243:             StrLogic = Null
            
        ' Casos que retornam VERDADEIRO (.T.)
        Case "S", "SIM", "T", ".T.", "TRUE", "YES", "ON", "Y", "1"
2247:             StrLogic = True
            
        ' Casos que retornam FALSO (.F.)
        Case "N", "NAO", "F", ".F.", "FALSE", "NO", "OFF", "0"
2251:             StrLogic = False
            
        ' Fallback de segurança (Caso não combine com nenhum)
        Case Else
2255:             StrLogic = False
2256:     End Select
End Function
Public Function Count_Lines_In_File(ByVal strFilePath As String, Optional ByVal lMES As Boolean = True) As Long
    Dim fso As Object
    Dim Stream As Object
    
    ' Mantém a sua validação original usando o seu FileConnExist
2263:     If Not FileConnExist(strFilePath, lMES) Then
2264:         Count_Lines_In_File = -1
2265:         Exit Function
2266:     End If
    
2268:     On Error GoTo TrataErro
    
2270:     Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Abre o arquivo em modo leitura (1 = ForReading)
2273:     Set Stream = fso.OpenTextFile(strFilePath, 1, False)
    
    ' TRUQUE DE ALTA PERFORMANCE DO FSO:
    ' Em vez de ler linha por linha em um loop, o .SkipLine pula as linhas na velocidade do Kernel.
    ' .ReadAll joga o ponteiro direto para o fim do arquivo lendo o buffer de uma vez só.
2278:     Stream.ReadAll
    
    ' A propriedade .Line nos dá exatamente o número total de linhas do arquivo
2281:     Count_Lines_In_File = Stream.Line
    
Fim:
2284:     If Not Stream Is Nothing Then Stream.Close
2285:     Set Stream = Nothing
2286:     Set fso = Nothing
2287:     Exit Function

TrataErro:
2290:     Count_Lines_In_File = 0
2291:     Resume Fim
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
2333:     On Error GoTo TrataErro
    
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
2348:     If Dir(cCaminhoArquivo) = "" Then
2349:         Count_Lines_In_Filev2 = 0
2350:         Exit Function
2351:     End If
    
2353:     nTamArquivo = FileLen(cCaminhoArquivo)
2354:     If nTamArquivo = 0 Then
2355:         Count_Lines_In_Filev2 = 0
2356:         Exit Function
2357:     End If
    
    ' 1 MB de Buffer estático na RAM
2360:     nTamBloco = 1024 * 1024
    ReDim abBuffer(0 To nTamBloco - 1)
    
2363:     nFile = FreeFile
2364:     Open cCaminhoArquivo For Binary Access Read As #nFile
    
2366:     nLINHAS = 0
2367:     nPonteiroAtual = 1 ' No VB6, o Seek do modo Binary começa em 1
    
2369:     Do While nPonteiroAtual <= nTamArquivo
        ' Calcula quantos bytes faltam ler
2371:         nRestante = (nTamArquivo - nPonteiroAtual) + 1
        
        ' Ajusta o tamanho do bloco se for o final do arquivo
2374:         If nRestante < nTamBloco Then
2375:             nTamBloco = nRestante
            ReDim abBuffer(0 To nTamBloco - 1)
2377:         End If
        
        ' Posiciona o ponteiro do arquivo e faz a leitura
2380:         Seek #nFile, nPonteiroAtual
2381:         Get #nFile, , abBuffer
2382:         nBytesLidos = nTamBloco
        
        ' Se for o último pedaço do arquivo, guarda o último byte real
2385:         If (nPonteiroAtual + nBytesLidos - 1) = nTamArquivo Then
2386:             bUltimoByteArquivo = abBuffer(nBytesLidos - 1)
2387:         End If
        
        ' --- AQUI ENTRA A LÓGICA DO BLOCO ELÁSTICO ---
        ' Se o bloco NÃO for o final do arquivo, vamos analisar onde ele cortou
2391:         If (nPonteiroAtual + nBytesLidos - 1) < nTamArquivo Then
            
            ' Caso A: O bloco terminou exatamente em Chr(13) (Pode ser um CRLF cortado)
2394:             If abBuffer(nBytesLidos - 1) = 13 Then
                ' Encolhe o bloco logicamente em 1 byte (ignora o 13 neste turno)
2396:                 nBytesLidos = nBytesLidos - 1
                
            ' Caso B: O bloco terminou em Chr(10) e o anterior a ele era Chr(13)
            ' (O par CRLF ficou inteiro dentro do bloco, então está tudo bem)
2400:             ElseIf abBuffer(nBytesLidos - 1) = 10 Then
2401:                 If abBuffer(nBytesLidos - 2) = 13 Then
                    ' Par CRLF fechado com sucesso, não mexe em nada.
2403:                 Else
                    ' É um LF isolado (Unix), pode processar normal.
2405:                 End If
2406:             End If
2407:         End If
        
        ' Varre o bloco ajustado (Sem risco de pegar caractere partido)
2410:         For i = 0 To nBytesLidos - 1
2411:             If abBuffer(i) = 13 Then
2412:                 nLINHAS = nLINHAS + 1
2413:             ElseIf abBuffer(i) = 10 Then
                ' Se o anterior imediato foi 13, faz parte do par. Não conta.
                ' Caso contrário, é uma quebra pura Unix (10), conta nova linha.
2416:                 If i > 0 Then
2417:                     If abBuffer(i - 1) <> 13 Then
2418:                         nLINHAS = nLINHAS + 1
2419:                     End If
2420:                 Else
                    ' Se for o primeiro byte do bloco e for 10, como o bloco elástico
                    ' evitou deixar o 13 no bloco anterior, este 10 é uma linha pura Unix.
2423:                     nLINHAS = nLINHAS + 1
2424:                 End If
2425:             End If
2426:         Next i
        
        ' Atualiza o ponteiro real com base nos bytes que nós REALMENTE processamos
2429:         nPonteiroAtual = nPonteiroAtual + nBytesLidos
2430:     Loop
    
2432:     Close #nFile
2433:     nFile = 0
    
    ' Ajuste final da última linha (Caso o arquivo não termine com quebra)
2436:     If nLINHAS > 0 Then
2437:         If bUltimoByteArquivo <> 10 And bUltimoByteArquivo <> 13 Then
2438:             nLINHAS = nLINHAS + 1
2439:         End If
2440:     Else
2441:         If nTamArquivo > 0 Then nLINHAS = 1
2442:     End If
    
2444:     Count_Lines_In_Filev2 = nLINHAS
2445:     Exit Function

TrataErro:
2448:     If nFile > 0 Then Close #nFile
2449:     Count_Lines_In_Filev2 = 0
End Function
'Public Sub OpenUrl(ByVal strURL As String)
'  ShellExecute 0, "Open", strURL, 0&, 0&, SW_SHOWNORMAL
'End Sub

Public Sub OpenUrl(ByVal strURL As String)
    ' 1. Valida se começa com http/https, se não, pode não abrir
2457:     If InStr(1, strURL, "http", vbTextCompare) = 0 Then strURL = "http://" & strURL
    
    ' 2. Executa a abertura no navegador padrão do sistema
2460:     ShellExecute 0&, "Open", strURL, vbNullString, vbNullString, 1 ' SW_SHOWNORMAL
End Sub

Public Function txttopdf(ByVal cOrigem As String, Optional ByVal cDestino As String = "", Optional ByVal cTITULO As String = "", Optional ByVal cAUTOR As String = "") As Boolean
    Dim fso As Object
    Dim streamIn As Object
    Dim cLINHA As String
    Dim pdf As ClsFPDF
    
    ' Validações estruturais padrões do seu sistema
2470:     If Not FileConnExist(cOrigem, True) Then Exit Function
2471:     If Len(cDestino) = 0 Then cDestino = TrocaExt(cOrigem, "PDF")
    
2473:     If FileConnExist(cDestino, False) Then
2474:         Alert ("Arquivo Destino Ja existe")
2475:         Exit Function
2476:     End If
    
2478:     On Error GoTo TrataErro
    
    ' 1. Abre o arquivo de texto em modo Fluxo contínuo (Segurança contra arquivos gigantes)
2481:     Set fso = CreateObject("Scripting.FileSystemObject")
2482:     Set streamIn = fso.OpenTextFile(cOrigem, 1, False)
    
    ' 2. Inicializa o motor do FPDF
2485:     Set pdf = New ClsFPDF
2486:     pdf.CreatePDF "P", "mm"
2487:     pdf.AddPage
2488:     pdf.SetFont "Arial", 10
    
    ' 3. Configura a paginação e margens (em milímetros)
    Dim nMargemEsquerda As Single: nMargemEsquerda = 15
    Dim nLinhaAtual As Single: nLinhaAtual = 15
    Dim nAlturaLinha As Single: nAlturaLinha = 5 ' Espaçamento de 5mm entre as linhas
    Dim nLimiteInferior As Single: nLimiteInferior = 275 ' Altura máxima da folha A4 antes de quebrar a página
    
    ' 4. Loop de streaming de performance: lê e grava em tempo real
2497:     Do While Not streamIn.AtEndOfStream
2498:         cLINHA = streamIn.ReadLine
        
        ' Desenha a linha atual de texto no documento PDF
2501:         pdf.Cell cLINHA, nMargemEsquerda, nLinhaAtual
        
        ' Avança o cursor vertical para a próxima linha
2504:         nLinhaAtual = nLinhaAtual + nAlturaLinha
        
        ' QUEBRA DE PÁGINA AUTOMÁTICA: Se o texto atingir o rodapé da folha, abre uma página nova
2507:         If nLinhaAtual > nLimiteInferior Then
2508:             pdf.AddPage
2509:             pdf.SetFont "Arial", 10 ' Reaplica a fonte na nova página
2510:             nLinhaAtual = 15        ' Reinicia o cursor no topo da nova folha
2511:         End If
2512:     Loop
    
    ' 5. Compila e descarrega o arquivo final no disco
2515:     pdf.Output cDestino
2516:     txttopdf = True

Fim:
    ' Destrói os ponteiros e liberta totalmente a memória RAM do computador
2520:     If Not streamIn Is Nothing Then streamIn.Close
2521:     Set streamIn = Nothing
2522:     Set fso = Nothing
2523:     Set pdf = Nothing
2524:     Exit Function

TrataErro:
2527:     Alert "Erro na rotina de geração do PDF: " & Err.Description
2528:     Resume Fim
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

2607:   With iXml.createTextNode(iString)
2608:     CharCodesToHTML = .Xml
2609:   End With
End Function

Public Function HTMLToCharCodes(ByVal iString As String) As String
2613:   With New MSXML2.DOMDocument60
2614:     .loadXML "<p>" & iString & "</p>"
2615:     HTMLToCharCodes = .selectSingleNode("p").nodeTypedValue
2616:   End With
End Function
Public Function str2html(ByVal cTEXTO As String, Optional ByVal lAnsi As Boolean = False) As String
2619:   If lAnsi Then
2620:     Convert2ansi (cTEXTO)
2621:   End If
2622:   cTEXTO = CharCodesToHTML(cTEXTO)
2623:   str2html = cTEXTO
End Function
Function FileText(ByVal filename As String) As String
  Dim Handle As Integer
2627:   Handle = FreeFile
2628:   Open filename$ For Input As #Handle
2629:   FileText = Input$(LOF(Handle), Handle)
2630:   Close #Handle
End Function
Public Function Html2Str(ByVal cTEXTO As String) As String
2633:   Html2Str = HtmlToText(cTEXTO)  'HTMLToCharCodes(cTexto)
End Function
Public Function HtmlToText(sHTML) As String
  Dim oDoc As HTMLDocument
2637:   Set oDoc = New HTMLDocument
2638:   oDoc.Body.innerHTML = sHTML
2639:   HtmlToText = oDoc.Body.innerText
End Function
Public Function FindInList(ByRef cList As ListBox, sSearch As String) As Long
  Dim sString As String
  Dim ID As Integer

2645:   On Error Resume Next
2646:   Err.Clear
2647:   FindInList = -1

2649:   For ID = 0 To cList.ListCount - 1
2650:     sString = UCase$(cList.List(ID))
2651:     If sString = UCase$(sSearch) Then
2652:       FindInList = ID
2653:       Exit For
2654:     End If
2655:   Next

End Function

Function TiraEspaco(sNome As String) As String
  Const sEspaco As String = " "
  Dim CaracAtual As String
  Dim NomeVerificado As String
  Dim NomeSemEspaco As String
  Dim i As Integer

2666:   For i = 1 To Len(sNome)
2667:     NomeVerificado = Mid(sNome, 1, i)
2668:     CaracAtual = Mid(NomeVerificado, i, 1)   'verificando os caracteres de dois em dois
2669:     If CaracAtual = sEspaco Then
2670:       NomeSemEspaco = Trim$(NomeSemEspaco) & sEspaco
2671:     Else
2672:       NomeSemEspaco = NomeSemEspaco & CaracAtual
2673:     End If
2674:   Next
2675:   TiraEspaco = Trim$(NomeSemEspaco)
End Function

Function FastArraySearch(SearchArray As Variant, SearchPhrase As String) As Long  'String
  Dim Pos As Long
  Dim i As Long
  Dim NumCharsProcessed As Long
 ' dim TXT As String
2683:   FastArraySearch = -1
2684:   Pos = InStr(Join(SearchArray, "§"), SearchPhrase)
2685:   If Pos > 0 Then
2686:     For i = LBound(SearchArray) To UBound(SearchArray)
2687:       NumCharsProcessed = NumCharsProcessed + Len(SearchArray(i)) + 1
2688:       If NumCharsProcessed >= Pos Then
2689:         FastArraySearch = i              ''SearchArray(i)
2690:         Exit Function
2691:       End If
2692:     Next i
2693:   End If
End Function



Public Function NetworkUserName() As String
  Dim iStringLength As Long
  Dim i As Long
  Dim sString As String
  Dim oShell As Object
2703:     On Error Resume Next

2705: NetworkUserName = ""
2706:   sString = String(255, 0)

2708:   iStringLength = Len(sString)
2709:   sString = Space$(iStringLength)

2711:   If WinAPI_GetUserName(sString, CLng(iStringLength)) Then

2713:     i = InStr(sString, Chr(0))
2714:     If i Then
2715:       sString = Left(sString, i - 1)
2716:     End If
2717:     NetworkUserName = Trim(Left$(sString, iStringLength))
2718:   Else
2719:    Set oShell = CreateObject("WScript.Shell")
2720:     If Not oShell Is Nothing Then
2721:         NetworkUserName = oShell.ExpandEnvironmentStrings("%USERNAME%")
2722:     Else
2723:         NetworkUserName = Environ$("USERNAME") ' Fallback simples
2724:     End If
  
2726:   End If

2728:   If NetworkUserName = "" Then
2729:     NetworkUserName = "Usuario"
2730:   End If
  

2733: Set oShell = Nothing
End Function
Public Function MachineName() As String
    ' Usamos o CreateObject global do VBA/twinBASIC, não do RC6.
    ' Isso é compatível com 32 e 64 bits e não depende de DLLs ou TypeLibs.
    Dim oShell As Object
2739:     On Error Resume Next
2740:     Set oShell = CreateObject("WScript.Shell")
2741:     If Not oShell Is Nothing Then
2742:         MachineName = oShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
2743:     Else
2744:         MachineName = Environ$("COMPUTERNAME") ' Fallback simples
2745:     End If
2746:     Set oShell = Nothing
End Function
Public Function WordLen(ByRef Text As String) As Long
'tamanho somente dos caracteres normal 65 a 90
  Dim Bytes() As Byte
  Dim i As Long

2753:   Bytes = StrConv(UCase$(Text), vbFromUnicode)
2754:   For i = 0 To UBound(Bytes)
2755:     If 65 <= Bytes(i) And Bytes(i) <= 90 Then WordLen = WordLen + 1
2756:   Next
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
2782:   Bytes = StrConv(UCase$(Text1), vbFromUnicode)
2783:   For i = 0 To UBound(Bytes)
2784:     LetterCounts(Bytes(i)) = LetterCounts(Bytes(i)) + 1
2785:   Next
2786:   LetterCountsS1 = LetterCounts

  ReDim LetterCounts(65 To 90)
2789:   Bytes = StrConv(UCase$(Text2), vbFromUnicode)
2790:   For i = 0 To UBound(Bytes)
2791:     If 65 <= Bytes(i) And Bytes(i) <= 90 Then
2792:       LetterCounts(Bytes(i)) = LetterCounts(Bytes(i)) + 1
2793:     End If
2794:   Next
2795:   LetterCountsS2 = LetterCounts

2797:   SameWords = LetterCountsS1 = LetterCountsS2
End Function



Public Function TimedMsgBox(Prompt As String, Optional ByVal TimeOut As Long = 0, Optional Icon As VbMsgBoxStyle = vbOKOnly, Optional Title As String = vbNullString)

  Dim TimerId As Long
  #If VBA7 Then
     Dim nADDRESSOF As Long
  #End If

2809:   CurMBTitle = Title

2811:   If TimeOut = 0 Then
2812:     TimeOut = 5 * 1000
2813:   Else
2814:     TimeOut = TimeOut * 1000
2815:   End If
  
  
  #If VBA7 Then
2819:     nADDRESSOF = CLng(AddressOf TimeOutMB)
2820:     TimerId = SetTimer(0, 0, TimeOut, nADDRESSOF)
  #Else
2822:      TimerId = SetTimer(0, 0, TimeOut, AddressOf TimeOutMB)
  #End If
2824:   TimedMsgBox = MsgBox(Prompt, Icon, CurMBTitle)
2825:   TimedMsgBox = 0
2826:   KillTimer 0, TimerId

End Function

Private Sub TimeOutMB(hWnd As Long, uMsg As Long, idEvent As Long, dwTime As Long)

2832:   SendMessage FindWindow(vbNullString, CurMBTitle), WM_CLOSE, 0&, 0&

End Sub
Function CheckRegEx(texto As String, padrao As String)

  Dim objRegExp As RegExp

  'cria um objeto expressão regular
2840:   Set objRegExp = New RegExp

  'define o padrão - Pattern
2843:   objRegExp.Pattern = padrao

  'define IgnoreCase
2846:   objRegExp.IgnoreCase = True

  'define a propriedade global
2849:   objRegExp.Global = True

  'verifica se a string pode ser comparada
2852:   CheckRegEx = False
2853:   If (objRegExp.Test(texto) = True) Then
2854:     CheckRegEx = True

    'obtem as coincidencias
    '' Set colMatch = objRegExp.Execute(texto)   'executa a busca


    'For Each objMatch In colMatch
    ''   valor = valor & " padrao encontrado na posição "
    ''   valor = valor & objMatch.FirstIndex & ". o valor é '"
    ''   valor = valor & objMatch.Value & " '." & vbCrLf
    'Next
2865:   Else
    ' valor = "Comparacao falhou !"
2867:   End If

2869: Set objRegExp = Nothing ' Adicione isto para boa gestão de memória

End Function

'Returns True if the user is running 64-bit *VBA*
'   This is *NOT* the same as the Windows bitness;
'   64-bit Windows can (and often does) run 32-bit VBA
Public Function Isvba64() As Boolean
2877: Isvba64 = False
    #If Win64 Then
2879:         Isvba64 = True
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
2894:     If Dir(cOrigem) = "" Then
2895:         txttodoc = False
2896:         Exit Function
2897:     End If
    
    ' 2. Regra do Destino: Se não foi passado, gera o .doc baseado no arquivo de origem
2900:     If Trim(cDestino) = "" Then
2901:         strOutputFile = Replace(LCase(cOrigem), ".txt", ".doc")
2902:     Else
2903:         strOutputFile = cDestino
2904:     End If
    
2906:     On Error GoTo TrataErro
    
    ' 3. Inicializa o FSO para o processamento em Streaming (Leve e rápido)
2909:     Set fso = CreateObject("Scripting.FileSystemObject")
2910:     Set streamIn = fso.OpenTextFile(cOrigem, 1, False)
2911:     Set streamOut = fso.OpenTextFile(strOutputFile, 2, True)
    
    ' 4. Escreve o cabeçalho HTML com suporte a metadados de Título e Autor para o Word
2914:     streamOut.WriteLine "<html>"
2915:     streamOut.WriteLine "<head>"
2916:     streamOut.WriteLine "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
    
    ' Injeta o Autor e Título se eles tiverem sido informados por parâmetro
2919:     If Trim(cAUTOR) <> "" Then
2920:         streamOut.WriteLine "<meta name=""author"" content=""" & cAUTOR & """>"
2921:     End If
2922:     If Trim(cTITULO) <> "" Then
2923:         streamOut.WriteLine "<title>" & cTITULO & "</title>"
2924:     End If
    
    ' Estilização CSS para o Word manter as fontes monoespaçadas e respeitar os espaços em branco do TXT
2927:     streamOut.WriteLine "<style>"
2928:     streamOut.WriteLine "body { font-family: 'Courier New', Courier, monospace; font-size: 10pt; line-height: 1.2; }"
2929:     streamOut.WriteLine "p { margin: 0; padding: 0; white-space: pre; }"
2930:     streamOut.WriteLine "</style>"
2931:     streamOut.WriteLine "</head>"
2932:     streamOut.WriteLine "<body>"
    
    ' 5. Loop de Leitura/Escrita linha por linha
2935:     Do While Not streamIn.AtEndOfStream
2936:         cLINHA = streamIn.ReadLine
        
        ' Escapa caracteres que podem quebrar a sintaxe do HTML
2939:         cLINHA = Replace(cLINHA, "&", "&amp;")
2940:         cLINHA = Replace(cLINHA, "<", "&lt;")
2941:         cLINHA = Replace(cLINHA, ">", "&gt;")
        
        ' Se for linha vazia, adiciona o espaço fixo para o Word não "engolir" a linha
2944:         If Len(Trim(cLINHA)) = 0 Then
2945:             streamOut.WriteLine "<p>&nbsp;</p>"
2946:         Else
2947:             streamOut.WriteLine "<p>" & cLINHA & "</p>"
2948:         End If
2949:     Loop
    
    ' 6. Fecha as tags do documento
2952:     streamOut.WriteLine "</body>"
2953:     streamOut.WriteLine "</html>"
    
    ' Se chegou até aqui sem erros, operação concluída com sucesso!
2956:     txttodoc = True

Fim:
    ' Liberação e fechamento seguro de canais de memória
2960:     If Not streamIn Is Nothing Then streamIn.Close
2961:     If Not streamOut Is Nothing Then streamOut.Close
2962:     Set streamIn = Nothing
2963:     Set streamOut = Nothing
2964:     Set fso = Nothing
2965:     Exit Function

TrataErro:
    ' Em caso de qualquer falha (permissão de escrita, arquivo travado), retorna False
2969:     txttodoc = False
2970:     Resume Fim
End Function
Public Function txttortf(ByVal cOrigem As String, Optional ByVal cDestino As String = "", Optional ByVal cTITULO As String = "", Optional ByVal cAUTOR As String = "") As Boolean
    Dim fso As Object
    Dim streamIn As Object
    Dim streamOut As Object
    Dim cLINHA As String
    Dim cBlocoInfo As String
    
    ' Mantém a sua validação original usando FileConnExist
2980:     If Not FileConnExist(cOrigem, True) Then Exit Function
    
    ' Mantém a sua lógica original para caminhos e extensões padrões
2983:     If Len(cDestino) = 0 Then cDestino = TrocaExt(cOrigem, "RTF")
    
    ' Impede a sobreescrita acidental conforme o seu padrão
2986:     If FileConnExist(cDestino, False) Then
2987:         Alert ("Arquivo Destino Ja existe")
2988:         Exit Function
2989:     End If
    
    ' Mantém os valores padrões vindos do seu sistema, idêntico à sua txttohtml
2992:     If Len(cAUTOR) = 0 Then cAUTOR = zNOMEFOLHA
2993:     If Len(cTITULO) = 0 Then cTITULO = NomeArq(cOrigem, True)
    
2995:     On Error GoTo TrataErro
    
    ' Inicializa o FileSystemObject para leitura e gravação em fluxo (streaming)
2998:     Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Abre o arquivo de origem para leitura (1 = ForReading)
3001:     Set streamIn = fso.OpenTextFile(cOrigem, 1, False)
    ' Cria o arquivo de destino para escrita (True = Sobreescrever)
3003:     Set streamOut = fso.CreateTextFile(cDestino, True)
    
    ' --- CONSTRUÇÃO DOS METADADOS (BLOCO INFO) ---
    ' Monta a tag {\info} se houver título ou autor preenchidos
3007:     cBlocoInfo = ""
3008:     If Len(cTITULO) > 0 Or Len(cAUTOR) > 0 Then
3009:         cBlocoInfo = "{\info"
3010:         If Len(cTITULO) > 0 Then cBlocoInfo = cBlocoInfo & "{\title " & cTITULO & "}"
3011:         If Len(cAUTOR) > 0 Then cBlocoInfo = cBlocoInfo & "{\author " & cAUTOR & "}"
3012:         cBlocoInfo = cBlocoInfo & "}"
3013:     End If
    
    ' Escreve o cabeçalho de inicialização do formato RTF juntamente com os metadados
3016:     streamOut.Write "{\rtf1\ansi\ansicpg1252\deff0\deflang1031" & cBlocoInfo & "{\fonttbl{\f0\fnil\fcharset0 Arial;}}\viewkind4\uc1\pard\f0\fs20 "
    
    ' Loop de streaming otimizado por linha para arquivos gigantes
3019:     Do While Not streamIn.AtEndOfStream
3020:         cLINHA = streamIn.ReadLine
        
        ' Proteção para que caracteres de controle do RTF presentes no texto original não corrompam o arquivo
3023:         If InStr(cLINHA, "\") > 0 Then cLINHA = Replace(cLINHA, "\", "\\")
3024:         If InStr(cLINHA, "{") > 0 Then cLINHA = Replace(cLINHA, "{", "\{")
3025:         If InStr(cLINHA, "}") > 0 Then cLINHA = Replace(cLINHA, "}", "\}")
        
        ' Grava no arquivo adicionando o controle de fim de parágrafo (\par) do RTF
3028:         streamOut.Write cLINHA & "\par " & vbCrLf
3029:     Loop
    
    ' Escreve a tag de encerramento do documento RTF
3032:     streamOut.Write "\par}"
    
3034:     txttortf = True

Fim:
    ' Garante a desconexão dos arquivos e liberação completa dos objetos de streaming da memória
3038:     If Not streamIn Is Nothing Then streamIn.Close
3039:     If Not streamOut Is Nothing Then streamOut.Close
3040:     Set streamIn = Nothing
3041:     Set streamOut = Nothing
3042:     Set fso = Nothing
3043:     Exit Function

TrataErro:
3046:     Alert "Erro na gravação/leitura do arquivo RTF: " & Err.Description
3047:     Resume Fim
End Function
Public Function txttohtml(ByVal cOrigem As String, Optional ByVal cDestino As String = "", Optional ByVal cTITULO As String = "", Optional ByVal cAUTOR As String = "") As Boolean
    Dim fso As Object
    Dim streamIn As Object
    Dim streamOut As Object
    Dim cLINHA As String
    
    ' Mantém a sua validação original usando FileConnExist
3056:     If Not FileConnExist(cOrigem, True) Then Exit Function
    
    ' Mantém a sua lógica original para caminhos e extensões padrões
3059:     If Len(cDestino) = 0 Then cDestino = TrocaExt(cOrigem, "HTML")
    
    ' Impede a sobreescrita acidental conforme seu código original
3062:     If FileConnExist(cDestino, False) Then
3063:         Alert ("Arquivo Destino Ja existe")
3064:         Exit Function
3065:     End If
    
    ' Mantém os valores padrões vindos das suas variáveis globais e funções de string
3068:     If Len(cAUTOR) = 0 Then cAUTOR = zNOMEFOLHA
3069:     If Len(cTITULO) = 0 Then cTITULO = NomeArq(cOrigem, True)
    
3071:     On Error GoTo TrataErro
    
    ' Inicializa o FileSystemObject para ler e gravar em fluxo de buffer
3074:     Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Abre o arquivo de origem para leitura (1 = ForReading)
3077:     Set streamIn = fso.OpenTextFile(cOrigem, 1, False)
    ' Cria o arquivo de destino para escrita (True = Sobreescrever)
3079:     Set streamOut = fso.CreateTextFile(cDestino, True)
    
    ' Escreve exatamente a mesma estrutura de cabeçalho que o seu Print original fazia
3082:     streamOut.Write "<html>" & vbCrLf
3083:     streamOut.Write "<head>" & vbCrLf
3084:     streamOut.Write "<meta http-equiv=" & Chr(34) & "Content-Type" & Chr(34) & vbCrLf
3085:     streamOut.Write "content=" & Chr(34) & "text/html; charset=iso-8859-1" & Chr(34) & ">" & vbCrLf
3086:     streamOut.Write "<meta name=" & Chr(34) & "GENERATOR" & Chr(34) & "content=" & Chr(34) & cAUTOR & Chr(34) & ">" & vbCrLf
3087:     streamOut.Write "<title>" & cTITULO & "</title>" & vbCrLf
3088:     streamOut.Write "</head>" & vbCrLf
3089:     streamOut.Write "<body>" & vbCrLf
    
    ' Loop de streaming otimizado linha por linha para arquivos gigantes
3092:     Do While Not streamIn.AtEndOfStream
3093:         cLINHA = streamIn.ReadLine
        
        ' AQUI: Aplica a função original str2html para tratar acentos e caracteres especiais
3096:         cLINHA = str2html(cLINHA)
        
        ' Joga no arquivo destino adicionando a quebra de linha do HTML
3099:         streamOut.Write cLINHA & "<br>" & vbCrLf
3100:     Loop
    
    ' Fecha a estrutura do documento exatamente como o seu original
3103:     streamOut.Write "</body>" & vbCrLf
3104:     streamOut.Write "</html>" & vbCrLf
    
3106:     txttohtml = True

Fim:
    ' Garante o fechamento correto dos arquivos e liberação da RAM
3110:     If Not streamIn Is Nothing Then streamIn.Close
3111:     If Not streamOut Is Nothing Then streamOut.Close
3112:     Set streamIn = Nothing
3113:     Set streamOut = Nothing
3114:     Set fso = Nothing
3115:     Exit Function

TrataErro:
3118:     Alert "Erro na gravação/leitura do arquivo: " & Err.Description
3119:     Resume Fim
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
3141:     If Dir(cOrigem) = "" Then
3142:         txttoxls = False
3143:         Exit Function
3144:     End If
    
    ' 2. Regra do Destino: Se não foi passado, gera o .xls baseado no nome do TXT
3147:     If Trim(cDestino) = "" Then
3148:         strOutputFile = Replace(LCase(cOrigem), ".txt", ".xls")
3149:     Else
3150:         strOutputFile = cDestino
3151:     End If
    
    ' 3. TRATAMENTO DOS DELIMITADORES DO SEU FORMULÁRIO (tab, ;,  ,, |, ~#)
3154:     cDELIMITADOR = Trim(LCase(cDELIMITADOR))
    
3156:     Select Case cDELIMITADOR
        Case "<tab>", "tab", Chr(9)
3158:             cDelimReal = Chr(9)   ' Tabulação Real
        Case ""
3160:             cDelimReal = ";"      ' Fallback padrão se vier vazio
        Case Else
3162:             cDelimReal = cDELIMITADOR ' Captura ;, ,, |, ~# exatamente como estão no Form
3163:     End Select
    
3165:     On Error GoTo TrataErro
    
    ' 4. Inicializa o FileSystemObject para processamento leve (Streaming)
3168:     Set fso = CreateObject("Scripting.FileSystemObject")
3169:     Set streamIn = fso.OpenTextFile(cOrigem, 1, False)       ' 1 = ForReading
3170:     Set streamOut = fso.OpenTextFile(strOutputFile, 2, True) ' 2 = ForWriting
    
    ' 5. Escreve o cabeçalho HTML formatado para Excel e LibreOffice
3173:     streamOut.WriteLine "<html>"
3174:     streamOut.WriteLine "<head>"
3175:     streamOut.WriteLine "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"">"
    
    ' Metadados do documento
3178:     If Trim(cAUTOR) <> "" Then streamOut.WriteLine "<meta name=""author"" content=""" & cAUTOR & """>"
3179:     If Trim(cTITULO) <> "" Then streamOut.WriteLine "<title>" & cTITULO & "</title>"
    
    ' CSS para forçar linhas e bordas visíveis no Excel/Calc
3182:     streamOut.WriteLine "<style>"
3183:     streamOut.WriteLine "table { border-collapse: collapse; }"
3184:     streamOut.WriteLine "td, th { border: 1px solid #A0A0A0; font-family: Calibri, Arial, sans-serif; font-size: 11pt; padding: 5px; }"
3185:     streamOut.WriteLine "th { background-color: #EFEFEF; font-weight: bold; }"
3186:     streamOut.WriteLine "</style>"
3187:     streamOut.WriteLine "</head>"
3188:     streamOut.WriteLine "<body>"
3189:     streamOut.WriteLine "<table>"
    
3191:     bPrimeiraLinha = True
    
    ' 6. Loop de Performance: Separa as colunas com base no delimitador tratado
3194:     Do While Not streamIn.AtEndOfStream
3195:         cLINHA = streamIn.ReadLine
        
        ' Faz a divisão da linha usando o delimitador correto correspondente ao do Form
3198:         arrCampos = Split(cLINHA, cDelimReal)
        
3200:         streamOut.WriteLine "  <tr>"
3201:         For i = LBound(arrCampos) To UBound(arrCampos)
            
            ' Isola o conteúdo da célula tirando espaços extras
3204:             cValor = Trim(CStr(arrCampos(i)))
            
            ' Aplica a sua função nativa str2html para tratar acentos e caracteres especiais
3207:             cValor = str2html(cValor)
            
            ' Se for a primeira linha do TXT, vira cabeçalho destacado (th)
3210:             If bPrimeiraLinha Then
3211:                 streamOut.WriteLine "    <th>" & cValor & "</th>"
3212:             Else
                ' Gravação correta da célula de dados (td)
3214:                 streamOut.WriteLine "    <td>" & cValor & "</td>"
3215:             End If
3216:         Next i
3217:         streamOut.WriteLine "  </tr>"
        
3219:         bPrimeiraLinha = False
3220:     Loop
    
    ' 7. Fecha o arquivo de forma íntegra
3223:     streamOut.WriteLine "</table>"
3224:     streamOut.WriteLine "</body>"
3225:     streamOut.WriteLine "</html>"
    
3227:     txttoxls = True

Fim:
    ' Destruição rigorosa de objetos para não prender memória RAM
3231:     If Not streamIn Is Nothing Then streamIn.Close
3232:     If Not streamOut Is Nothing Then streamOut.Close
3233:     Set streamIn = Nothing
3234:     Set streamOut = Nothing
3235:     Set fso = Nothing
3236:     Exit Function

TrataErro:
3239:     txttoxls = False
3240:     Resume Fim
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
