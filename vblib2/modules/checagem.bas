Attribute VB_Name = "checagem"
Option Explicit
Public Enum TabOrder
        NextControl = 0
        PreviousControl = 1
End Enum
Public Function validaprocesso(ByVal cPROCESSO As String) As Boolean
    Dim cNumero As String, cDD As String, cAno As String
    Dim cJTR As String, cFinal As String
    Dim resto1 As Long, resto2 As Long, resto3 As Long
    
    cPROCESSO = funNumeroPuro(cPROCESSO)
    If Len(cPROCESSO) <> 20 Then
        validaprocesso = False
        Exit Function
    End If
    
    cNumero = Mid(cPROCESSO, 1, 7)
    cDD = Mid(cPROCESSO, 8, 2)
    cAno = Mid(cPROCESSO, 10, 4)
    cJTR = Mid(cPROCESSO, 14, 3)
    cFinal = Mid(cPROCESSO, 17, 4)
    
    resto1 = Val(cNumero) Mod 97
    resto2 = Val(Format(resto1, "00") & cAno & cJTR) Mod 97
    resto3 = Val(Format(resto2, "00") & cFinal & cDD) Mod 97
    
    validaprocesso = (resto3 = 1)
End Function

Public Function formataprocesso(ByVal cPROCESSO As String) As String
    If validaprocesso(cPROCESSO) Then
        formataprocesso = Format(funNumeroPuro(cPROCESSO), "0000000-00.0000.0.00.0000")
    Else
        formataprocesso = cPROCESSO
    End If
End Function
Public Function ValidaPlaca(ByVal cPlaca As String) As Boolean
    Dim objRegExp As Object
    
    cPlaca = UCase(funNumeroPuro(cPlaca))
    If Len(cPlaca) <> 7 Then
        ValidaPlaca = False
        Exit Function
    End If
    
    ' Bloqueia placas genéricas apontadas no Harbour[cite: 5]
    If cPlaca = "XXX9999" Or cPlaca = "XXX999" Then
        ValidaPlaca = False
        Exit Function
    End If
    
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.IgnoreCase = True
    
    ' Padrão (Antigo OU Mercosul) unificado do Harbour[cite: 5]
    objRegExp.Pattern = "^[A-Z]{3}[0-9]{4}$|^[A-Z]{3}[0-9][A-Z][0-9]{2}$"
    
    ValidaPlaca = objRegExp.Test(cPlaca)
    Set objRegExp = Nothing
End Function
Public Function CheckIE(cIE As Variant, cUF As Variant, Optional lMES As Boolean = True, Optional ByVal cPESSOA As String = "J", Optional ByVal lOLD As Boolean = True) As Boolean
    Dim zErro As String
    Dim lTAM As Boolean
    Dim nLen As Integer, cPref As String
    
    cIE = UCase(funNumeroPuro(cIE))
    cUF = UCase(Trim(cUF))
    nLen = Len(cIE)
    
    If cUF = "EX" Or cUF = "XX" Then
        If Left(cIE, 5) = "ISENT" Or cIE = "00000000000000" Then
            CheckIE = True
            If lMES Then MsgBox "Inscricao OK-ISENTO-EXPORTACAO", vbInformation, "Checagem IE"
        Else
            CheckIE = False
            If lMES Then MsgBox "Inscricao Invalida Use ISENTO para exportacao", vbExclamation, "Checagem IE"
        End If
        Exit Function
    End If
    
    If Left(cIE, 5) = "ISENT" Or (InStr(cIE, "NAO") > 0 And InStr(cIE, "CONTRIB") > 0) Then
        CheckIE = True ' Isento ou Não Contribuinte[cite: 6]
        Exit Function
    End If
    
    If cPESSOA <> "J" Then
        CheckIE = True ' Não valida IE de Pessoa Física nesta etapa[cite: 6]
        Exit Function
    End If
    
    If nLen < 8 Then
        CheckIE = False
        If lMES Then MsgBox "IE Inválido - Menos de 8 dígitos", vbExclamation
        Exit Function
    End If
    
    ' Validação de Tamanhos Específicos por Estado[cite: 6]
    lTAM = True
    Select Case cUF
        Case "AC", "DF", "MG": If nLen <> 13 And nLen <> 9 Then lTAM = False
        Case "PE", "RO": If nLen <> 9 And nLen <> 14 Then lTAM = False
        Case "TO": If nLen <> 9 And nLen <> 11 Then lTAM = False
        Case "BA": If nLen <> 8 And nLen <> 9 Then lTAM = False
        Case "SP": If nLen <> 12 And nLen <> 13 Then lTAM = False
        Case "MT": If nLen <> 11 Then lTAM = False
        Case "PR", "RS": If nLen <> 10 Then lTAM = False
        Case "AL", "AM", "AP", "CE", "ES", "GO", "MA", "MS", "PA", "PB", "PI", "RN", "RR", "SC", "SE": If nLen <> 9 Then lTAM = False
        Case "RJ": If nLen <> 8 Then lTAM = False
        Case Else: lTAM = False
    End Select
    
    If Not lTAM Then
        CheckIE = False
        If lMES Then MsgBox "Tamanho da Inscrição Inválido para o Estado " & cUF, vbExclamation
        Exit Function
    End If
    
    ' Validação de Prefixos (2 primeiros dígitos) para Estados Específicos[cite: 6]
    lTAM = True
    cPref = Left(cIE, 2)
    Select Case cUF
        Case "AC": If cPref <> "01" Then lTAM = False
        Case "AL": If cPref <> "24" Then lTAM = False
        Case "AP": If cPref <> "03" Then lTAM = False
        Case "CE": If cPref <> "06" Then lTAM = False
        Case "DF": If cPref <> "07" Then lTAM = False
        Case "GO": If cPref <> "10" And cPref <> "11" And cPref <> "15" Then lTAM = False
        Case "MA": If cPref <> "12" Then lTAM = False
        Case "MT": If cPref <> "28" Then lTAM = False
        Case "PA": If cPref <> "15" Then lTAM = False
        Case "PB": If cPref <> "16" Then lTAM = False
        Case "PE": If cPref <> "18" Then lTAM = False
        Case "PI": If cPref <> "19" Then lTAM = False
        Case "RN": If cPref <> "20" Then lTAM = False
        Case "RR": If cPref <> "24" Then lTAM = False
        Case "SC": If cPref <> "25" Then lTAM = False
        Case "SE": If cPref <> "27" Then lTAM = False
        Case "TO": If cPref <> "29" Then lTAM = False
    End Select
    
    If Not lTAM Then
        CheckIE = False
        If lMES Then MsgBox "Início da Inscrição Inválido (" & cPref & ") para o Estado " & cUF, vbExclamation
        Exit Function
    End If

    CheckIE = True
End Function
Public Function MascararNumGenerico(ByVal cNumero As String, Optional ByVal nExibeInicio As Integer = 2, Optional ByVal nExibeFinal As Integer = 2) As String
    Dim cLimpo As String, nLen As Integer
    cLimpo = Trim(cNumero)
    nLen = Len(cLimpo)
    
    If nLen <= (nExibeInicio + nExibeFinal) Then
        If nLen > 2 Then
            MascararNumGenerico = Left(cLimpo, 1) & String(nLen - 2, "*") & Right(cLimpo, 1)
        Else
            MascararNumGenerico = String(nLen, "*")
        End If
        Exit Function
    End If
    
    MascararNumGenerico = Left(cLimpo, nExibeInicio) & String(nLen - nExibeInicio - nExibeFinal, "*") & Right(cLimpo, nExibeFinal)
End Function

Public Function MascararCPF(ByVal cCPF As String) As String
    Dim cLimpo As String
    cLimpo = funNumeroPuro(cCPF)
    
    If Len(cLimpo) <> 11 Then
        MascararCPF = cCPF
        Exit Function
    End If
    
    MascararCPF = Left(cLimpo, 3) & ".***.***-" & Right(cLimpo, 2)
End Function

Public Function MascararNome(ByVal cNomeRaw As String) As String
    Dim czNomeTrimmed
    Dim aPalavras As Variant
    Dim nQtdPalavras As Integer, i As Integer
    Dim cNomeMascarado As String
    
    czNomeTrimmed = Trim(cNomeRaw)
    aPalavras = Split(Trim(cNomeRaw), " ")
    nQtdPalavras = UBound(aPalavras) - LBound(aPalavras) + 1
    
    If nQtdPalavras <= 1 Then
        MascararNome = cNomeRaw
        Exit Function
    End If
    
    If nQtdPalavras = 2 Then
        MascararNome = aPalavras(0) & " " & String(Len(aPalavras(1)), "*")
        Exit Function
    End If
    
    cNomeMascarado = aPalavras(0) & " "
    For i = 1 To nQtdPalavras - 2
        cNomeMascarado = cNomeMascarado & String(Len(aPalavras(i)), "*") & " "
    Next i
    cNomeMascarado = cNomeMascarado & aPalavras(nQtdPalavras - 1)
    
    MascararNome = cNomeMascarado
End Function
Public Sub GetControlTABFocus(frmParent As Form, Optional TbNextControlOrder As TabOrder = 0)
        Dim TabIndexControl As Long
        Dim ctrlFormControl As Control
        
        TabIndexControl = frmParent.ActiveControl.TabIndex
        If TbNextControlOrder = PreviousControl Then
                If TabIndexControl <= 0 Then
                        TabIndexControl = frmParent.Controls.Count - 1
                Else
                        TabIndexControl = frmParent.ActiveControl.TabIndex - 1
                End If
        ElseIf TbNextControlOrder = NextControl Then
        
                 If TabIndexControl >= frmParent.Controls.Count - 1 Then
                        TabIndexControl = 0
                Else
                        TabIndexControl = frmParent.ActiveControl.TabIndex + 1
                End If
                                
        End If
      
        For Each ctrlFormControl In frmParent.Controls
                If ctrlFormControl.TabIndex = TabIndexControl Then
                        ctrlFormControl.SetFocus
                        Exit For
                End If
        Next
End Sub
Public Function CNPJ_Novo(ByVal pCNPJ As String) As Boolean
    Dim lResult As Boolean
    Dim soma As Integer
    Dim dv As String
    Dim digito As Integer
    Dim num As Integer
    Dim wCGC As String
    Dim i As Integer
    Dim j As Integer
    Dim Validos As String

  lResult = True
  soma = 0
  dv = ""
  digito = 0
  num = 0
  wCGC = pCNPJ  ' iif(ValType(pCNPJ)="U", "", pCNPJ)
  i = 0
  j = 0
  Validos = "0123456789"


  wCGC = Replace(wCGC, ".", "")
  wCGC = Replace(wCGC, "-", "")
  wCGC = Replace(wCGC, "/", "")
  If IsNull(wCGC) Then
    lResult = False
  Else
    If Len(wCGC) < 14 Then
      lResult = False
    Else

      For i = 1 To 12
        If InStr("ABCDEFGHIJKLMNOPQRSTUWYXZ", Mid(wCGC, i, 1)) Then
          Validos = "0123456789ABCDEFGHIJKLMNOPQRSTUWYXZ"
        End If
      Next
      dv = ""
      num = 5
      For j = 1 To 2
        soma = 0
        For i = 1 To 12
          soma = soma + ((Asc(Mid(wCGC, i, 1)) - 48) * num)
          num = num - 1
          If num = 1 Then
            num = 9
          End If
        Next
        If j = 2 Then
          soma = soma + (2 * Val(dv))
        End If
        digito = soma - (Int(soma / 11) * 11)
        If digito = 0 Or digito = 1 Then
          dv = dv + "0"
        Else
          dv = dv + Str(11 - digito)  'str(,1)
        End If
        num = 6
      Next
      If dv <> Mid(wCGC, 13, 2) Then
        lResult = False
      End If
    End If
  End If
  CNPJ_Novo = lResult
End Function



Public Function IsValidIP(ByVal IPAddress As String) As Boolean
    If IPAddress = "localhost" Then IsValidIP = True: Exit Function
    
    Dim SubNets() As String
    SubNets = Split(IPAddress, ".")
    
    ' Verifica se tem exatamente 4 octetos
    If UBound(SubNets) <> 3 Then
        IsValidIP = False
        Exit Function
    End If
    
    Dim i As Integer, valOcteto As Long
    For i = 0 To 3
        If Not IsNumeric(SubNets(i)) Then IsValidIP = False: Exit Function
        valOcteto = CLng(SubNets(i))
        If valOcteto < 0 Or valOcteto > 255 Then IsValidIP = False: Exit Function
    Next i
    
    IsValidIP = True
End Function

Public Function formataTelefone(ByVal cNumero As String) As String
    Dim lDDIBR As Boolean
    cNumero = Trim(cNumero)
    
    ' Ignora formatação de 0800, 0300, 0303, 0500, 0900
    If Left(cNumero, 4) = "0300" Or Left(cNumero, 4) = "0800" Or Left(cNumero, 4) = "0900" Or Left(cNumero, 4) = "0500" Or Left(cNumero, 4) = "0303" Then
        formataTelefone = cNumero
        Exit Function
    End If
    
    ' Trata DDI Brasileiro (+55)[cite: 4]
    If Left(cNumero, 1) = "+" Then
        If Left(cNumero, 3) <> "+55" Then
            formataTelefone = cNumero ' Se for de outro país, retorna como está
            Exit Function
        End If
        lDDIBR = True
        cNumero = Trim(Mid(cNumero, 4))
    End If
    
    ' Limpa toda a pontuação suja que o usuário possa ter digitado
    cNumero = funNumeroPuro(cNumero)
    
    ' Se tiver DDD (10 ou 11 dígitos no total)
    If Len(cNumero) >= 10 Then
        ' Remove o zero inicial do DDD se o usuário digitou (ex: 01199999999)[cite: 4]
        If Left(cNumero, 1) = "0" And Len(cNumero) >= 11 Then
            cNumero = Mid(cNumero, 2)
        End If
        ' Separa o DDD e joga o resto pra FormataTelfon2
        cNumero = "(" & Left(cNumero, 2) & ")" & FormataTelfon2(Mid(cNumero, 3))
    Else
        cNumero = FormataTelfon2(cNumero)
    End If
    
    If lDDIBR Then cNumero = "+55 " & cNumero
    
    formataTelefone = cNumero
End Function

Public Function FormataTelfon2(ByVal cNumero As String) As String
    cNumero = funNumeroPuro(cNumero)
    
    ' Máscara para novo celular (9 dígitos)[cite: 4]
    If Len(cNumero) = 9 Then
        FormataTelfon2 = Left(cNumero, 5) & "-" & Mid(cNumero, 6)
    ' Máscara para telefone padrão (8 dígitos)[cite: 1, 4]
    ElseIf Len(cNumero) = 8 Then
        FormataTelfon2 = Left(cNumero, 4) & "-" & Mid(cNumero, 5)
    ' Máscara para formato muito antigo (7 dígitos)[cite: 1, 4]
    ElseIf Len(cNumero) = 7 Then
        FormataTelfon2 = Left(cNumero, 3) & "-" & Mid(cNumero, 4)
    Else
        FormataTelfon2 = cNumero
    End If
End Function

Public Function FormataRG(ByVal Valor As String, Optional ByVal cTIPO As String = "") As String
    Dim cDAC As String, nPOS As Integer
    FormataRG = Valor
    
    If InStr(UCase(Valor), "ISENT") > 0 Or InStr(UCase(Valor), "RNE") > 0 Or InStr(UCase(Valor), "RIC") > 0 Then
        Exit Function ' Não formata isentos e estrangeiros[cite: 4]
    End If
    
    ' Se for CPF ou CNI digitado no campo de RG, redireciona para a máscara de CPF[cite: 4]
    If InStr(UCase(Valor), "CPF") > 0 Or InStr(UCase(Valor), "CNI") > 0 Or cTIPO = "CPF" Or cTIPO = "CNI" Then
        FormataRG = FormataCPF(Replace(Replace(UCase(Valor), "CPF", ""), "CNI", ""))
        Exit Function
    End If

    If Val(funNumeroPuro(Valor)) = 0 Then Exit Function

    Valor = Trim(Valor)
    nPOS = InStr(Valor, "-")
    
    If nPOS = 0 Then
        Valor = Trim(funNumeroPuro(Valor))
        If Len(Valor) = 9 Then
            cDAC = Mid(Valor, 9, 1)
            Valor = Left(Valor, 8)
        End If
    Else
        cDAC = UCase(Mid(Valor, nPOS + 1, 1))
        Valor = Left(Valor, nPOS - 1)
        If cDAC <> "X" Then cDAC = funNumeroPuro(cDAC)
    End If
    
    Valor = funNumeroPuro(Valor)
    
    Select Case Len(Valor)
        Case 8: FormataRG = Trim(Mid(Valor, 1, 2) & "." & Mid(Valor, 3, 3) & "." & Mid(Valor, 6))
        Case 7: FormataRG = Trim(Mid(Valor, 1, 1) & "." & Mid(Valor, 2, 3) & "." & Mid(Valor, 5))
        Case Else: FormataRG = Valor
    End Select
    
    If cDAC <> "" Then FormataRG = FormataRG & "-" & cDAC
    If Left(FormataRG, 1) = "0" Or Left(FormataRG, 1) = "." Then FormataRG = Mid(FormataRG, 2)
End Function
Public Function checkcfo(ByVal cCFO As String, ByVal cUFORI As String, ByVal cUFDES As String, ByVal cTIPO As String) As Boolean
  Dim cINISP As String
  Dim cINIFORA As String
  Dim cINIEXP As String
  
  
  If cTIPO = "E" Then                          ''Entradas
    cINISP = "1"
    cINIFORA = "2"
    cINIEXP = "3"
  Else                                         ''Saidas
    cINISP = "5"
    cINIFORA = "6"
    cINIEXP = "7"
  End If
  checkcfo = False
  cCFO = TiraOut(cCFO)
  If Len(cCFO) < 4 Then
    Alert ("CFO Precisa ter 4 digitos")
    Exit Function
  End If
  If Len(cUFDES) <> 2 Then
    Alert ("Estado necessita 2 digitos")
    Exit Function
  End If
  If cUFDES = cUFORI Then
    If Mid(cCFO, 1, 1) <> cINISP Then
      Alert "CFO SP Tem que Comecar por " & cINISP
      Exit Function
    End If
  Else
    If cUFDES = "EX" Or cUFDES = "XX" Then
      If Mid(cCFO, 1, 1) <> cINIEXP Then
        Alert "CFO Exportacao Tem que Comecar por " & cINIEXP
        Exit Function
      End If
    Else
      If Mid(cCFO, 1, 1) <> cINIFORA Then
        Alert "CFO Fora SP Tem que Comecar por " & cINIFORA
        Exit Function
      End If
    End If
  End If
  checkcfo = True
End Function

Public Function CheckPIS(ByVal ePIS As Variant, Optional lMES As Boolean = True) As Boolean
  Dim cPIS As String
  Dim ftap As String
  Dim Total As String
  Dim i As Integer
  Dim resto As Integer


  CheckPIS = False
  cPIS = FixStr(ePIS, "", "TRIM", 11)

  If Val(cPIS) = 0 Then
    If lMES Then Alert "Numero do Pis Nao Digitado"
    Exit Function
  End If
  If Len(cPIS) <> 11 Then
    If lMES Then Alert "Numero do Pis Nao Tem 11 Digitos"
    Exit Function
  End If
  If Mid(cPIS, 1, 1) <> "1" And Mid(cPIS, 1, 1) <> "2" Then
    If lMES Then Alert ("Codigo do pis invalido Nao Comeca com 1 OU 2")
    Exit Function
  End If
  If Mid(cPIS, 1, 3) >= "109" And Mid(cPIS, 1, 3) <= "119" Then
    If lMES Then Alert ("'Codigo pertence a contribunte individual da previdencia 109-119")
    Exit Function
  End If

  ftap = "3298765432"
  Total = 0

  For i = 1 To 10
    Total = Total + Val(Mid(cPIS, i, 1)) * Val(Mid(ftap, i, 1))
  Next i

  resto = Int(Total Mod 11)

  If resto <> 0 Then
    resto = 11 - resto
  End If

  If resto <> Val(Mid(cPIS, 11, 1)) Then
    CheckPIS = False
    If lMES Then Alert "Codigo do pis invalido"
    Exit Function
  End If

  CheckPIS = True

End Function


Public Function CheckCNPJ(cCGC As Variant, Optional cTIPO As String = "X", Optional lMES As Boolean = True, Optional cUF As String = "") As Boolean
''cTIPO= M-Matriz F-Filial X-Nao Checar nao mais usado a matriz pode ser diferente 0001 agora
  Dim X As Integer
  Dim aUF As Variant

  CheckCNPJ = False
  cCGC = Trim(TiraOut(cCGC))
  If Len(cCGC) = 0 Then
    If lMES Then
      Alert ("CPNJ Numero Nao Informado Em Branco")
    End If
    Exit Function
  End If
  If cCGC = "ISENTO" Then
    CheckCNPJ = True
    Exit Function
  End If

  If Len(cCGC) < 14 Then
    If lMES Then
      Alert ("CNPJ Numero Com Menos de 14 Digitos")
    End If
    Exit Function
  End If

  If Len(cCGC) > 14 Then
    If lMES Then
      Alert ("CNPJ Numero Com Mais de 14 Digitos")
    End If
    Exit Function
  End If


  For X = 0 To 14
    If cCGC = String(14, CStr(X)) Then
      If lMES Then Alert ("CNPJ Invalido - Sequencia Repetitiva de " + Str(X))
      Exit Function
    End If
  Next X


  If Left(funNumeroPuro(cCGC), 8) = "99999997" Then
    If lMES Then Alert ("CNPJ De referencia")
    CheckCNPJ = False
    Exit Function
  End If


  If Mid(cCGC, 9, 4) = "9999" Then
    If lMES Then Alert ("CNPJ Generico /9999-")
    CheckCNPJ = False
    Exit Function
  End If

  '   a matriz nao precisa ser mais 0001
  '   If cTIPO = "M" Then
  '       If Mid(cTIPO, 9, 4) <> "0001" Then
  '           If lMES Then
  '               Alert ("CNPJ Invalido-Nao e Matriz")
  '           End If
  '           Exit Function
  '       End If
  '   End If

  If Len(cUF) > 0 Then
    aUF = Array("AC", "AL", "AM", "AP", "BA", "CE", "DF", "ES", "GO", "MA", "MG", "MS", "MT", "PA", "PB", "PE", "PI", "PR", "RJ", "RN", "RO", "RR", "RS", "SC", "SE", "SP", "TO", "EX", "XX")
    If FastArraySearch(aUF, cUF) = -1 Then
      If lMES Then
        Alert ("Estado Invalido")
      End If
    End If
  End If

  If CNPJ_Novo(cCGC) Then
    CheckCNPJ = True
  Else
    If lMES Then Alert ("CNPJ Invalido ")
    Exit Function
  End If

  'If Mod11(cCGC, 13, 9) Then
  '    If Mod11(cCGC, 14, 9) Then
  '        CheckCNPJ = True
  '    Else
  '        If lMES Then Alert ("CNPJ Invalido - Cheque 14 Digito - 2 Verificador")
  '        Exit Function
  '    End If
  'Else
  '    If lMES Then Alert ("CNPJ Invalido - Cheque 13 Digito - 1 Verificador")
  '    Exit Function
  'End If


End Function

Public Function Mod11(ByVal campo, ByVal posdv, ByVal pesomax) As Boolean
  Dim dv As Integer
  Dim peso As Integer
  Dim imod As Integer
  Dim rest As Integer
  Mod11 = False
  dv = 0
  peso = 1
  For imod = posdv To 1 Step -1
    dv = dv + peso * Val(Mid(campo, imod, 1))
    peso = peso + 1
    If peso > pesomax Then
      peso = 2
    End If
  Next
  rest = dv Mod 11
  If rest = 0 Then
    Mod11 = True
  Else
    If rest = 1 And Val(Mid(campo, posdv, 1)) = 0 Then
      Mod11 = True
    Else
      Mod11 = False
    End If
  End If
End Function

Function FormataCGC(ByVal pCGC As String) As String
  Dim i As Integer
  Dim wCGC As String
  wCGC = ""
  For i = 1 To Len(pCGC)
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUWYXZ", Mid(pCGC, i, 1)) > 0 Then
      wCGC = wCGC & Mid(pCGC, i, 1)
    End If
  Next
  If Len(wCGC) = 14 Then
    FormataCGC = Left(wCGC, 2) & "." & Mid(wCGC, 3, 3) & "." & Mid(wCGC, 6, 3) _
               & "/" & Mid(wCGC, 9, 4) & "-" & Right(wCGC, 2)
  Else
    FormataCGC = pCGC
  End If
End Function

Public Function CheckCNPJIE(ByVal cPESSOA As String, ByVal CCNPJ As String, ByVal cIE As String, ByVal cUF As String, Optional ByVal lMES As Boolean = True) As Boolean
  Dim nERRO As Integer
  CheckCNPJIE = True
  If cPESSOA = "J" Then
    If Not CheckIE(cIE, cUF, lMES) Then
      CheckCNPJIE = False
      nERRO = 2
    End If
    If Not CheckCNPJ(CCNPJ, , lMES) Then
      CheckCNPJIE = False
      nERRO = 3
    End If
  End If
  If cPESSOA = "F" Then
    If Not CheckCPF(CCNPJ, lMES) Then
      CheckCNPJIE = False
      nERRO = 1
    End If
    If cIE <> "" Then
      If Not CheckRG(cIE, lMES) Then
        If lMES Then
          If Not MDG("Continuar Assim Mesmo") Then
            CheckCNPJIE = False
            nERRO = 4
          End If
        Else
          CheckCNPJIE = False
          nERRO = 4
        End If
      End If
    End If
  End If
  If Len(cUF) = 0 Then
    CheckCNPJIE = False
    If lMES Then
      Alert ("UF-Cidade  Nao Preenchido")
    End If
    nERRO = 5
  End If
  eRETU01 = nERRO
End Function

Function FormataCPF(ByVal pCPF As String) As String
  Dim i As Integer
  Dim wCPF As String
  wCPF = ""
  For i = 1 To Len(pCPF)
    If InStr("0123456789", Mid(pCPF, i, 1)) > 0 Then
      wCPF = wCPF & Mid(pCPF, i, 1)
    End If
  Next
  If Len(wCPF) = 11 Then
    FormataCPF = Left(wCPF, 3) & "." & Mid(wCPF, 4, 3) & "." & Mid(wCPF, 7, 3) _
               & "-" & Right(wCPF, 2)
  Else
    FormataCPF = pCPF
  End If
End Function

Public Function CheckCPF(ByVal xCPF As Variant, Optional ByVal lMES As Boolean = True) As Boolean
  Dim X As Integer
  Dim p1 As String
  CheckCPF = False
  p1 = Trim(TiraOut(xCPF))
  If Len(p1) = 0 Then
    If lMES Then Alert ("CPF Invalido - Em Branco")
    Exit Function
  End If
  If Len(p1) < 11 Then
    If lMES Then Alert ("CPF Invalido - Menos de 11 Digitos")
    Exit Function
  End If
  If Len(p1) > 11 Then
    If lMES Then Alert ("CPF Invalido - Mais de 11 Digitos")
    Exit Function
  End If

  For X = 0 To 9
    If p1 = String(11, CStr(X)) Then
      If lMES Then Alert ("CPF Invalido - Sequencia Repetitiva de " + Str(X))
      Exit Function
    End If
  Next X
  If Mod11(p1, 10, 10) Then
    If Mod11(p1, 11, 11) Then
      CheckCPF = True
    Else
      If lMES Then Alert ("CPF - Invalido")
      If lMES Then Alert ("Cheque 11Ý Digito - 2 Verificador")
      Exit Function
    End If
  Else
    If lMES Then Alert ("CPF - Invalido")
    If lMES Then Alert ("Cheque 10Ý Digito - 1 Verificador")
    Exit Function
  End If
End Function


Public Function CheckRena(ByVal dv_p1, Optional ByVal lMES As Boolean = True) As Boolean
  Dim dv_d3 As Integer
  Dim tam_chav As Integer
  Dim xrdig As String
  Dim y1 As Integer
  Dim dv_base As String
  Dim X As Integer
  Dim y2 As Integer
  Dim dv_p2 As String
  Dim tam_chave As Integer
  CheckRena = False
  dv_p1 = TiraOut(dv_p1)
  If Len(dv_p1) = 0 Then
    If lMES Then Alert ("Renavam Invalido - Em Branco")
    Exit Function
  End If
  ''If Len(dv_p1) < 10 Then
  ''   If lMES Then alert ("Renavam Invalido - Menos 10 Digitos")
  ''   Exit Function
  ''End If
  tam_chave = 10
  dv_p1 = StrZero(Val(dv_p1), tam_chave + 1)
  xrdig = Mid(dv_p1, 11, 1)
  dv_p1 = Mid(dv_p1, 1, 10)
  y1 = 0
  dv_base = "8923456789"
  y1 = 0
  For X = 1 To tam_chave
    dv_d3 = Val(Mid(dv_p1, X, 1))
    y1 = y1 + (dv_d3 * Val(Mid(dv_base, X, 1)))
    ''y2 = mod(y1,11)
    y2 = y1 Mod 11
  Next
  If y2 = 10 Then
    y2 = 0
  End If
  dv_p2 = Trim(Str(y2))
  If dv_p2 <> xrdig Then
    Alert ("Codigo do RENAVAN invalido.")
    Exit Function
  End If
  CheckRena = True
End Function

Function TeclaEnter(KeyIN As Integer) 'Enter vira tab e troca de campo
  TeclaEnter = KeyIN
  If KeyIN = 13 Or KeyIN = 10 Or KeyIN = 40 Then
    SendKeys Chr(9)
    TeclaEnter = 0
    Exit Function
  End If
  If KeyIN = 38 Then
    SendKeys "+" + Chr(9)
    TeclaEnter = 0
    Exit Function
  End If
End Function
'If ((Shift And vbCtrlMask) > 0) Then
'    ' do something
'End If
'If Shift = vbCtrlMask Then
    ' do something
'End If
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'

Function ValiText(KeyIN As Integer, Optional ByVal eTipo As String = "#N", _
                  Optional ByVal Editable As Boolean = True, _
                  Optional ByVal eSUBTIPO As String = "", _
                  Optional ByVal lENTER As Boolean = True)
  Dim ValidateList As String
  Dim KeyOut As Integer
  If lENTER Then
    If TeclaEnter(KeyIN) = 0 Then 'Enter vira tab e troca de campo
      ValiText = 0                'teclaentet usa sendkeys zera valitext=0 para nao virar loop
      Exit Function
    End If
  End If
  Select Case eTipo
  Case "#HR"
    ValidateList = "0123456789.:"
  Case "#NI", "I"                                   ''Numeros Inteiros"
    ValidateList = "0123456789"
  Case "#NI-"                                  ''Numeros Inteiros e Negativos
    ValidateList = "0123456789-"
  Case "#NP", "F"                                  ''Numeros Inteiros e o Ponto
    ValidateList = "0123456789."
  Case "#N-.X", "RG"                                ''Numeros Traco  Ponto X (RG)"
    ValidateList = "0123456789.-X"
  Case "#N-./", "CPF"                                ''Numero - e Ponto (CGC/CNPJ,CPF,CIC)
    ValidateList = "0123456789.-/"
   Case "#N-()+", "FONE"                      'Numero -()+  FONE
    ValidateList = "0123456789-()+"
  Case "#N-", "CEP"                            ''Numero -   CEP
    ValidateList = "0123456789-"
  Case "#D", "N/"                              ''Numero / - Data
    ValidateList = "0123456789/"
  Case "#N", "#N.,"                            ''Numeros . ,
    ValidateList = "0123456789.,"
  Case "#NEG", "#N.,-"                         ''Numeros . ,
    ValidateList = "0123456789.,-"
  Case "#SN"                                   ''Sim Nao
    ValidateList = "SN"
  Case "#C", "#CU", "#CL", "C"                     ''SoLetras SoMaisculas SoMinusculas
    ValidateList = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
  Case "#CN", "#CNU", "#CNL", "CN"                  ''letras numeros
    ValidateList = "01234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ"
  Case "CNPJ", "CGC"                  ''letras numeros
    ValidateList = "01234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ.-/"
  Case "#CNP", "#CNPU", "#CNPL"                ''letras numeros
    ValidateList = ",.01234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ"
  Case Else                                    ''Se nao For um tipo e uma lista
    ValidateList = eTipo
  End Select

  If Editable = True Then
    ValidateList = UCase(ValidateList) & Chr(8)
  Else
    ValidateList = UCase(ValidateList)
  End If

  If InStr(1, ValidateList, UCase(Chr(KeyIN)), 1) > 0 Then
    KeyOut = KeyIN
  Else
    KeyOut = 0
    'Beep
  End If
  If KeyOut > 32 Then
    If eTipo = "#CU" Or eTipo = "#SN" Or eSUBTIPO = "#CU" Or eTipo = "#CNU" Or eSUBTIPO = "#CNU" Or eTipo = "#CNPU" Or eSUBTIPO = "#CNPU" Then
      KeyOut = Asc(UCase(Chr$(KeyOut)))
    End If
    If eTipo = "#CL" Or eSUBTIPO = "#CL" Or eTipo = "#CNL" Or eSUBTIPO = "#CNL" Or eTipo = "#CNPL" Or eSUBTIPO = "#CNPL" Then
      KeyOut = Asc(LCase(Chr$(KeyOut)))
    End If
  End If
  ValiText = KeyOut
End Function

Public Function CheckRG(ByVal Valor As String, Optional ByVal lMES As Boolean = True, Optional ByVal cTIPO As String = "", Optional ByVal dDATANASC As Variant, Optional ByVal cUF As String = "") As Boolean
    Dim D As Integer, soma As Integer, nPOS As Integer
    Dim cDAC As String, zDAC As String, p1 As String, X As Integer
    Dim aPESOS As Variant
    
    CheckRG = True
    zDAC = " "
    
    ' Libera isentos e registros estrangeiros diretamente
    If UCase(cTIPO) = "ISENT" Or InStr(UCase(Valor), "ISENT") > 0 Then Exit Function
    If UCase(cTIPO) = "RNE" Or InStr(UCase(Valor), "RNE") > 0 Then Exit Function
    
    If Len(Trim(Valor)) = 0 Then
        CheckRG = False
        If lMES Then MsgBox "RG/RNE/RIC/CPF/CNI em branco", vbExclamation
        Exit Function
    End If
    
    ' Verifica se é CPF/CNI digitado no campo
    If InStr(UCase(Valor), "CPF") > 0 Then
        cTIPO = "CPF"
        Valor = Replace(UCase(Valor), "CPF", "")
    End If
    
    If CheckCPF(Valor, False) Then
        If cTIPO <> "CPF" And cTIPO <> "CNI" Then
            If lMES Then MsgBox "Preencha o tipo como CPF ou CNI", vbExclamation
        End If
        Exit Function
    End If
    
    ' Verifica RIC
    If InStr(UCase(Valor), "RIC") > 0 Then
        cTIPO = "RIC"
        Valor = Replace(UCase(Valor), "RIC", "")
    End If
    
    If cTIPO = "" Then cTIPO = "RG"
    
    ' Proteção contra sequência repetitiva
    p1 = funNumeroPuro(Valor)
    For X = 0 To 9
        If p1 = String(Len(p1), CStr(X)) Then
            CheckRG = False
            If lMES Then MsgBox "RG Inválido - Sequência Repetitiva de " & CStr(X), vbExclamation
            Exit Function
        End If
    Next X
    
    ' Extração do DAC
    Valor = Replace(Valor, ".", "")
    nPOS = InStr(Valor, "-")
    If nPOS = 0 Then
        cDAC = " "
    Else
        cDAC = UCase(Mid(Valor, nPOS + 1, 1))
        Valor = Left(Valor, nPOS - 1)
    End If
    
    Valor = CStr(Val(funNumeroPuro(Valor)))
    
    ' Validação de Tamanho RG com base na Data de Nascimento e UF
    If Len(Trim(Valor)) <= 7 And cTIPO = "RG" And (cUF = "" Or cUF = "SP") Then
        If IsMissing(dDATANASC) Or IsEmpty(dDATANASC) Or IsNull(dDATANASC) Then
            CheckRG = False
            If lMES Then MsgBox "RG com Menos de 7 Dígitos", vbExclamation
            Exit Function
        ElseIf IsDate(dDATANASC) Then
            If CDate(dDATANASC) > CDate("31/12/1990") Then
                CheckRG = False
                If lMES Then MsgBox "RG com Menos de 7 Dígitos", vbExclamation
                Exit Function
            End If
        End If
    End If
    
    If Len(Trim(Valor)) > 9 And (cTIPO = "RG" Or cTIPO = "") And (cUF = "" Or cUF = "SP") Then
        CheckRG = False
        If lMES Then MsgBox "RG com Mais de 9 Dígitos", vbExclamation
        Exit Function
    End If
    
    ' Cálculo RIC (Padrão 11 dígitos)
    If cTIPO = "RIC" Then
        If Len(Trim(Valor)) <> 11 Then
            CheckRG = False
            If lMES Then MsgBox "RIC não tem 11 Dígitos", vbExclamation
            Exit Function
        Else
            aPESOS = Array(8, 9, 2, 3, 4, 5, 6, 7, 8, 9)
            soma = 0
            For X = 1 To 10
                soma = soma + (Val(Mid(Valor, X, 1)) * aPESOS(X - 1))
            Next X
            D = soma Mod 11
            If D = 10 Then D = 0
            If D <> Val(cDAC) Then
                CheckRG = False
                If lMES Then MsgBox "Dígito de Controle RIC Não Confere", vbExclamation
                Exit Function
            End If
        End If
    End If
    
    ' Cálculo RG (Otimizado com For...Next substituindo n1 a n8)
    If cTIPO = "RG" And Len(Valor) = 8 Then
        Valor = Format(Val(Valor), "00000000")
        soma = 0
        For X = 1 To 8
            soma = soma + (Val(Mid(Valor, X, 1)) * (10 - X))
        Next X
        
        D = soma Mod 11
        
        If cDAC = "X" Or cDAC = " " Then
            If D <> 10 Then
                CheckRG = False
                If lMES Then MsgBox "Dígito de Controle RG Nao Confere", vbExclamation
            End If
        Else
            If D <> Val(cDAC) And D <> 0 Then
                CheckRG = False
                If D = 10 Then zDAC = "X" Else zDAC = CStr(D)
                If lMES Then MsgBox "Dígito de Controle RG Não Confere. Sugerido: " & zDAC, vbExclamation
            End If
        End If
    End If
End Function
Public Function ChecaNumMaiorZero(ByVal nText As Variant, ByVal eTexto As String) As Boolean
'nText = expressa o campo/numero/variavel
  nText = FixNum(nText)
  If (nText) > 0 Then
    ChecaNumMaiorZero = True
    eTexto = ""
  Else
    ChecaNumMaiorZero = False
    Alert (eTexto)
  End If
End Function

Public Function ValidaDados(ByVal KeyAscii As Integer, ByVal eTipo As String _
     , Optional ByVal lENTER As Boolean = True) As Integer

  Dim nTAM As Integer
  Dim nLen As Integer
  Dim txtAtual As String
  Dim selInicio As Integer

  If lENTER Then
    If TeclaEnter(KeyAscii) = 0 Then
      ValidaDados = 0
      Exit Function
    End If
  End If

  Select Case eTipo
  Case "PLACA"
    nTAM = 8
  Case "HORA"
    nTAM = 5
  Case "HORAL"
    nTAM = 8
  Case "CEP"
    nTAM = 9
  Case "TEL"
    nTAM = 15
  Case "DC"
    nTAM = 8
  Case "DL"
    nTAM = 10
  Case "CPF", "CIC"
    nTAM = 14
  Case "CNPJ", "CGC"
    nTAM = 18
  
  ' --- NOVOS TIPOS ADICIONADOS ---
  Case "CARTAO"
    nTAM = 19 ' 16 números + 3 espaços (0000 0000 0000 0000)
  Case "BARRAS"
    nTAM = 14 ' EAN/UPC máximo tratado na Validar_CodigoBarras é 14
  Case "PIS"
    nTAM = 14 ' 000.00000.00-0
  Case "CEI"
    nTAM = 15 ' 00.000.00000/00
  Case "CFOP"
    nTAM = 5  ' 0.000
  Case "RENAVAM"
    nTAM = 11 ' Apenas números
  Case "TITULO"
    nTAM = 12 ' Apenas números
  End Select
  
  txtAtual = Trim(FixStr(Screen.ActiveControl.Text))
  nLen = Len(txtAtual)
  selInicio = Screen.ActiveControl.SelStart
  
  If KeyAscii <> 8 And nLen >= nTAM And _
     KeyAscii <> 13 And Screen.ActiveControl.SelLength = 0 Then
    ValidaDados = 0
    Exit Function
  End If

  If KeyAscii <> 13 And KeyAscii <> 8 And eTipo <> "PLACA" And eTipo <> "TEL" Then
    KeyAscii = ValiText(KeyAscii, "#NI", True)
  End If

  If KeyAscii <> 8 And selInicio = nLen Then
    Select Case eTipo
    Case "HORA", "HORAL"
      If nLen = 0 And KeyAscii > 50 Then KeyAscii = 0
      If (nLen = 3 Or nLen = 6) And KeyAscii > 53 Then KeyAscii = 0
      If nLen = 2 Or (nLen = 5 And eTipo = "HORAL") Then
        Screen.ActiveControl.Text = txtAtual & ":"
        Screen.ActiveControl.SelStart = Len(Screen.ActiveControl.Text)
      End If

    Case "PLACA"
      ' 1. Se estiver na posição 3 (indo para o 4º caractere), injeta o hífen e atualiza o tamanho
      If nLen = 3 Then
        Screen.ActiveControl.Text = txtAtual & "-"
        Screen.ActiveControl.SelStart = Len(Screen.ActiveControl.Text)
        nLen = 4
      End If
      
      ' 2. Regras de validação baseadas no Padrão Antigo / Mercosul
      If nLen < 3 Then
        KeyAscii = ValiText(KeyAscii, "#CU") ' Posições 1 a 3: Obrigatório Letra
      ElseIf nLen = 4 Then
        KeyAscii = ValiText(KeyAscii, "#NI") ' Posição 5: Obrigatório Número
      ElseIf nLen = 5 Then
        KeyAscii = ValiText(KeyAscii, "#CNU") ' Posição 6: Letra (Mercosul) ou Número (Antigo)
      ElseIf nLen > 5 Then
        KeyAscii = ValiText(KeyAscii, "#NI") ' Posições 7 e 8: Obrigatório Número
      End If

    Case "DC", "DL"
      If nLen = 0 And KeyAscii > 51 Then KeyAscii = 0
      If nLen = 3 And KeyAscii > 49 Then KeyAscii = 0
      If nLen = 2 Or nLen = 5 Then
        Screen.ActiveControl.Text = txtAtual & "/"
        Screen.ActiveControl.SelStart = Len(Screen.ActiveControl.Text)
      End If

    Case "CEP"
      If nLen = 5 Then
        Screen.ActiveControl.Text = txtAtual & "-"
        Screen.ActiveControl.SelStart = Len(Screen.ActiveControl.Text)
      End If

    Case "TEL"
    ''  If KeyAscii <> 8 Then KeyAscii = ValiText(KeyAscii, "#N-")
      If KeyAscii <> 8 Then KeyAscii = ValiText(KeyAscii, "FONE")

    Case "CPF", "CIC"
      If nLen = 3 Or nLen = 7 Then
        Screen.ActiveControl.Text = txtAtual & "."
        Screen.ActiveControl.SelStart = Len(Screen.ActiveControl.Text)
      ElseIf nLen = 11 Then
        Screen.ActiveControl.Text = txtAtual & "-"
        Screen.ActiveControl.SelStart = Len(Screen.ActiveControl.Text)
      End If

    Case "CNPJ", "CGC"
      If nLen = 2 Or nLen = 6 Then
        Screen.ActiveControl.Text = txtAtual & "."
        Screen.ActiveControl.SelStart = Len(Screen.ActiveControl.Text)
      ElseIf nLen = 10 Then
        Screen.ActiveControl.Text = txtAtual & "/"
        Screen.ActiveControl.SelStart = Len(Screen.ActiveControl.Text)
      ElseIf nLen = 15 Then
        Screen.ActiveControl.Text = txtAtual & "-"
        Screen.ActiveControl.SelStart = Len(Screen.ActiveControl.Text)
      End If

    ' --- MÁSCARAS PARA OS NOVOS TIPOS ---
    Case "CARTAO"
      If nLen = 4 Or nLen = 9 Or nLen = 14 Then
        Screen.ActiveControl.Text = txtAtual & " "
        Screen.ActiveControl.SelStart = Len(Screen.ActiveControl.Text)
      End If

    Case "PIS"
      If nLen = 3 Or nLen = 9 Then
        Screen.ActiveControl.Text = txtAtual & "."
        Screen.ActiveControl.SelStart = Len(Screen.ActiveControl.Text)
      ElseIf nLen = 12 Then
        Screen.ActiveControl.Text = txtAtual & "-"
        Screen.ActiveControl.SelStart = Len(Screen.ActiveControl.Text)
      End If

    Case "CEI"
      If nLen = 2 Or nLen = 6 Then
        Screen.ActiveControl.Text = txtAtual & "."
        Screen.ActiveControl.SelStart = Len(Screen.ActiveControl.Text)
      ElseIf nLen = 12 Then
        Screen.ActiveControl.Text = txtAtual & "/"
        Screen.ActiveControl.SelStart = Len(Screen.ActiveControl.Text)
      End If

    Case "CFOP"
      If nLen = 1 Then
        Screen.ActiveControl.Text = txtAtual & "."
        Screen.ActiveControl.SelStart = Len(Screen.ActiveControl.Text)
      End If

    End Select
  End If
  
  ValidaDados = KeyAscii
End Function
Public Function ChecaSimNao(ByVal nText As Variant) As String
'nText = expressa o campo/numero/variavel
  nText = UCase(nText)
  If (nText) = "S" Then
    ChecaSimNao = "S"
  Else
    ChecaSimNao = "N"
  End If
End Function

Function FormataCEP(ByVal pCEP As String) As String
  Dim i As Integer
  Dim wCEP As String
  wCEP = ""
  For i = 1 To Len(pCEP)
    If InStr("0123456789", Mid(pCEP, i, 1)) > 0 Then
      wCEP = wCEP & Mid(pCEP, i, 1)
    End If
  Next
  If Len(wCEP) > 8 Then
    FormataCEP = pCEP
  Else
    FormataCEP = Format(wCEP, "00000-000")
  End If
End Function

Function FormataCFOP(ByVal pCFOP As String) As String
  Dim wCFOP As String
  wCFOP = Format(Val(funNumeroPuro(pCFOP)), "0")
  If Len(wCFOP) > 5 Then
    FormataCFOP = "*.**.**"
  ElseIf Len(wCFOP) = 0 Then
    FormataCFOP = ""
  ElseIf Len(wCFOP) = 4 Then
    FormataCFOP = " " & Left(wCFOP, 1) & "." & Mid(wCFOP, 2, 2) & "." & Right(wCFOP, 1)
  Else
    wCFOP = Right("     " & wCFOP, 5)
    FormataCFOP = IIf(Mid(wCFOP, 1, 1) = " ", "  ", Mid(wCFOP, 1, 1) & ".") _
                & IIf(Mid(wCFOP, 2, 2) = "  ", "   ", Mid(wCFOP, 2, 2) & ".") _
                & Mid(wCFOP, 4, 2)
  End If
End Function

Function FormataPlaca(ByVal eCHAPA As String) As String
  Dim cCHAPA As String
  cCHAPA = FixStr(eCHAPA, "")
  cCHAPA = UCase(cCHAPA)
  cCHAPA = Replace(cCHAPA, " ", "")
  cCHAPA = Replace(cCHAPA, "-", "")
  cCHAPA = Replace(cCHAPA, ".", "")
  
  ' Ajuste: Inserir o hífen APENAS se a placa já tiver a parte alfabética completa
  If Len(cCHAPA) > 3 Then
    cCHAPA = Mid(cCHAPA, 1, 3) & "-" & Mid(cCHAPA, 4)
  End If
  
  FormataPlaca = cCHAPA
End Function

Function CheckHora(ByVal eHORA As Variant, Optional ByVal lMES As Boolean = False) As Boolean
  Dim cTMP  As String
  Dim cHORA  As String
  Dim cMIN  As String
  Dim cSEG As String
  CheckHora = False
  cSEG = ""
  cTMP = FixStr(eHORA)
  cHORA = Mid(cTMP, 1, 2)
  cMIN = Mid(cTMP, 4, 2)
  If Len(cTMP) > 5 Then                        ''Hora Longa Segundos
    cSEG = Mid(cTMP, 7, 2)
  End If
  If cHORA <> "" And Val(cHORA) < 25 Then      '' A hora e menor que 24 then
    If cMIN <> "" And Val(cMIN) < 60 Then
      If cSEG = "" Then
        CheckHora = True
      Else
        If Val(cSEG) < 60 Then
          CheckHora = True
        Else
          If lMES Then
            Alert ("Segundos Tem que estar entre 0-60")
          End If
        End If
      End If
    Else
      If lMES Then
        Alert ("Minuto Tem que estar entre 0-60")
      End If
    End If
  Else
    If lMES Then
      Alert ("Hora Tem que estar entre 0-24")
    End If
  End If
End Function

Public Function CheckCTA(ByVal cBANCO As Variant, ByVal cAGENCIA As Variant, ByVal cCONTA As Variant, Optional ByVal lMES As Boolean = True) As Boolean
    Dim eTot As Long, nX As Integer, nPeso As Integer, nResto As Integer
    Dim cDVCalc As String, cDVInformado As String, cCorpoConta As String
    Dim cNumPuro As String
    Dim strBanco As String

    strBanco = Format(Val(funNumeroPuro(CStr(cBANCO))), "000")
    cAGENCIA = funNumeroPuro(CStr(cAGENCIA))

    cNumPuro = UCase(Replace(Replace(CStr(cCONTA), "-", ""), ".", ""))
    If Len(cNumPuro) = 0 Then Exit Function

    cCorpoConta = Left(cNumPuro, Len(cNumPuro) - 1)
    cDVInformado = Right(cNumPuro, 1)

    Select Case strBanco
        Case "001" ' Banco do Brasil
            eTot = 0: nPeso = 2
            For nX = Len(cCorpoConta) To 1 Step -1
                eTot = eTot + (Val(Mid(cCorpoConta, nX, 1)) * nPeso)
                nPeso = nPeso + 1
            Next nX
            nResto = 11 - (eTot Mod 11)
            If nResto = 10 Then cDVCalc = "X" Else If nResto = 11 Then cDVCalc = "0" Else cDVCalc = CStr(nResto)

        Case "104" ' Caixa Econômica Federal
            eTot = 0: nPeso = 2
            cCorpoConta = Right("0000" & cAGENCIA, 4) & cCorpoConta
            For nX = Len(cCorpoConta) To 1 Step -1
                eTot = eTot + (Val(Mid(cCorpoConta, nX, 1)) * nPeso)
                If nPeso = 9 Then nPeso = 2 Else nPeso = nPeso + 1
            Next nX
            cDVCalc = CStr((eTot * 10) Mod 11)
            If cDVCalc = "10" Then cDVCalc = "0"

        Case "260", "077", "336", "403", "323", "197", "290", "655", "212", "318" ' Fintechs
            eTot = 0: nPeso = 2
            For nX = Len(cCorpoConta) To 1 Step -1
                eTot = eTot + (Val(Mid(cCorpoConta, nX, 1)) * nPeso)
                If nPeso = 9 Then nPeso = 2 Else nPeso = nPeso + 1
            Next nX
            nResto = (eTot * 10) Mod 11
            If nResto >= 10 Then cDVCalc = "0" Else cDVCalc = CStr(nResto)

        Case "237" ' Bradesco
            eTot = 0: nPeso = 2
            For nX = Len(cCorpoConta) To 1 Step -1
                eTot = eTot + (Val(Mid(cCorpoConta, nX, 1)) * nPeso)
                If nPeso = 7 Then nPeso = 2 Else nPeso = nPeso + 1
            Next nX
            nResto = 11 - (eTot Mod 11)
            If nResto = 10 Then cDVCalc = "P" Else If nResto = 11 Then cDVCalc = "0" Else cDVCalc = CStr(nResto)

        Case "033" ' Santander
            eTot = 0
            cAGENCIA = Right(Right("0000" & cAGENCIA, 4), 3)
            eTot = eTot + CalcDig(7, Mid(cAGENCIA, 1, 1))
            eTot = eTot + CalcDig(3, Mid(cAGENCIA, 2, 1))
            eTot = eTot + CalcDig(1, Mid(cAGENCIA, 3, 1))
            
            cCorpoConta = Right("00000000" & cCorpoConta, 8)
            eTot = eTot + CalcDig(9, Mid(cCorpoConta, 1, 1))
            eTot = eTot + CalcDig(7, Mid(cCorpoConta, 2, 1))
            eTot = eTot + CalcDig(1, Mid(cCorpoConta, 3, 1))
            eTot = eTot + CalcDig(3, Mid(cCorpoConta, 4, 1))
            eTot = eTot + CalcDig(1, Mid(cCorpoConta, 5, 1))
            eTot = eTot + CalcDig(9, Mid(cCorpoConta, 6, 1))
            eTot = eTot + CalcDig(7, Mid(cCorpoConta, 7, 1))
            eTot = eTot + CalcDig(3, Mid(cCorpoConta, 8, 1))
            
            nResto = eTot Mod 10
            If nResto = 0 Then cDVCalc = "0" Else cDVCalc = CStr(10 - nResto)

        Case "341" ' Itaú
            cDVCalc = DAC10(cAGENCIA & cCorpoConta)

        Case Else
            CheckCTA = True
            Exit Function
    End Select

    If UCase(cDVCalc) <> cDVInformado Then
        CheckCTA = False
        If lMES Then MsgBox "Dígito Verificador Inválido para o Banco " & strBanco, vbExclamation, "Checagem de Conta"
    Else
        CheckCTA = True
    End If
End Function
Public Function CalcDig(ByVal n1, ByVal n2)
  Dim eRETU
  n2 = FixInt(n2)
  eRETU = n1 * n2
  eRETU = StrZero(eRETU)
  eRETU = Right(eRETU, 1)
  eRETU = Val(eRETU)
  CalcDig = eRETU
End Function

Public Function DAC10(ByVal Arg1)
  Dim cNumero As String
  Dim ninicio As Integer
  Dim ntotal As Integer
  Dim ccpoaux As String
  Dim X As Integer
  ninicio = Len(Trim(Arg1)) + 1
  ntotal = 0
  If (ninicio < 2) Then
    ninicio = 2
  End If
  ccpoaux = "0" + Trim(Arg1)
  For X = ninicio To 1 Step -2
    cNumero = Mid(ccpoaux, X, 1)
    ntotal = ntotal + InStr("516273849", cNumero)
    ntotal = ntotal + Val(Mid(ccpoaux, X - 1, 1))
  Next
  DAC10 = Trim(Str(InStr("987654321", Mid(StrZero(ntotal, 3), 3, 1))))
End Function

Public Function CheckDat2(ByVal eDATA As Variant, Optional ByVal lMES As Boolean = True) As Boolean

  CheckDat2 = False
  If Not DataBranco(eDATA) Then
    If IsDate(eDATA) Then
      CheckDat2 = True
    End If
  End If
  If lMES And Not CheckDat2 Then
    Alert ("Data Invalida")
  End If
End Function

Public Function CheckData(ByVal eDATA As Variant, Optional ByVal lMES As Boolean = True) As Boolean

  Dim aMES As Variant
  Dim nDIAMES As Integer
  Dim cData As String
  Dim nANO As Integer
  Dim nMES As Integer
  Dim nDIA As Integer
  ''          (01, 02, 03, 04, 05, 06, 07, 08, 09, 10, 11, 12)
  aMES = Array(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
  cData = Trim(TiraOut(eDATA))
  CheckData = True
  If Len(cData) > 5 Then
    nDIA = Val(Mid(cData, 1, 2))
    nMES = Val(Mid(cData, 3, 2))
    nANO = Val(Mid(cData, 5))
    If Len(cData) = 8 Then
      If nANO < 1900 Or nANO > 2300 Then
        If lMES Then
          Alert ("Ano <1900 or ano >2300")
          CheckData = False
        End If
      End If
    Else
      ''   If nANO < 1 Or nANO > 99 Then
      ''      If Lmes Then
      ''         alert ("Ano <1 or ano >99")
      ''         checkdata = False
      ''      End If
      ''   End If
    End If
    If nMES < 1 Or nMES > 12 Then
      If lMES Then
        Alert ("mes <1 or mes >12")
        CheckData = False
      End If
    Else
      nDIAMES = aMES(nMES - 1)             ''Array comeca com zero
      If nMES = 2 Then
        If Len(cData) = 6 Or Bissexto(nANO) Then
          nDIAMES = 29
        End If
      End If
      If nDIA < 1 Or nDIA > nDIAMES Then
        If lMES Then
          Alert ("dia <1 or dia  > " & CStr(nDIAMES))
          CheckData = False
        End If
      End If
    End If
  Else
    If lMES Then
      Alert ("Data Nao Preenchida")
      CheckData = False
    End If
  End If
End Function


Public Function CheckTitulo(ByVal s As String, Optional ByVal lMES As Boolean = True) As Boolean
  Dim DV1 As Integer
  Dim DV2 As Integer
  Dim i As Integer
  Dim resto As Integer
  Dim resto2 As Integer
  Dim Numero As String

  Numero = ""

  For i = Len(s) To 12
    Numero = Numero + "0"
  Next

  Numero = Numero + s
  CheckTitulo = False

  DV1 = ((CInt(Mid(Numero, 1, 1)) * 2) + _
         (CInt(Mid(Numero, 2, 1)) * 9) + _
         (CInt(Mid(Numero, 3, 1)) * 8) + _
         (CInt(Mid(Numero, 4, 1)) * 7) + _
         (CInt(Mid(Numero, 5, 1)) * 6) + _
         (CInt(Mid(Numero, 6, 1)) * 5) + _
         (CInt(Mid(Numero, 7, 1)) * 4) + _
         (CInt(Mid(Numero, 8, 1)) * 3) + _
         (CInt(Mid(Numero, 9, 1)) * 2))
  resto = DV1 Mod 11

  If resto = 1 Then
    DV1 = 0
  Else
    If resto = 0 Then
      If (((Mid(Numero, 10, 2)) = "01") Or ((Mid(Numero, 10, 2)) = "02")) Then
        DV1 = 1
      Else
        DV1 = 0
      End If
    Else
      DV1 = 11 - resto
    End If
  End If

  DV2 = ((CInt(Mid(Numero, 10, 1)) * 4) + (CInt(Mid(Numero, 11, 1)) * 3) + (DV1 * 2))
  resto2 = DV2 Mod 11

  If resto2 = 1 Then
    DV2 = 0
  Else
    If resto2 = 0 Then
      If (((Mid(Numero, 10, 2)) = "01") Or ((Mid(Numero, 10, 2)) = "02")) Then
        DV2 = 1
      Else
        DV2 = 0
      End If
    Else
      DV2 = 11 - resto2
    End If
  End If

  If Mid(Numero, 12, 2) = (CStr(DV1) + CStr(DV2)) Then
    CheckTitulo = True
  End If

  If lMES Then
    If Not CheckTitulo Then
      Alert ("Numero do Titulo Invalido")
    End If
  End If

End Function

Public Function CheckEmail(ByVal cEMAIL As String, _
                           Optional ByVal lVazio As Boolean = False, _
                           Optional ByVal lMES As Boolean = True) As Boolean

  cEMAIL = FixStr(cEMAIL)
  
  If Len(cEMAIL) = 0 Then
    If lVazio Then
      CheckEmail = True
    Else
      CheckEmail = False
      If lMES Then MsgBox "Email em branco"
    End If
    Exit Function
  End If

  If Left(UCase(cEMAIL), 3) = "NAO" Then
    CheckEmail = True
    Exit Function
  End If

  If Len(cEMAIL) < 5 Then
    CheckEmail = False
    If lMES Then MsgBox "O e-mail digitado tem menos de 5 caracterec!!!"
    Exit Function
  End If

  ' Substitui o loop de contagem de arrobas pelo UBound(Split())
  If UBound(Split(cEMAIL, "@")) <> 1 Then
    CheckEmail = False
    If lMES Then MsgBox "O no. de arrobas (@) do e-mail é inválido!!!"
    Exit Function
  End If

  ' Funções Left e Right substituem o InStr para verificar as pontas
  If Left(cEMAIL, 1) = "@" Then
    CheckEmail = False
    If lMES Then MsgBox "O e-mail foi iniciado com uma arroba (@)!!!"
    Exit Function
  ElseIf Right(cEMAIL, 1) = "@" Then
    CheckEmail = False
    If lMES Then MsgBox "O e-mail termina com uma arroba (@)!!!"
    Exit Function
  End If

  ' Verifica a existência de pelo menos um ponto
  If InStr(cEMAIL, ".") = 0 Then
    CheckEmail = False
    If lMES Then MsgBox "O e-mail é inválido, pois não contém pontos (.)!!!"
    Exit Function
  End If

  If Left(cEMAIL, 1) = "." Then
    CheckEmail = False
    If lMES Then MsgBox "O e-mail foi iniciado com um ponto (.)!!!"
    Exit Function
  ElseIf Right(cEMAIL, 1) = "." Then
    CheckEmail = False
    If lMES Then MsgBox "O e-mail termina com um ponto (.)!!!"
    Exit Function
  ElseIf InStr(InStr(cEMAIL, "@"), cEMAIL, ".") = 0 Then
    CheckEmail = False
    If lMES Then MsgBox "O e-mail não tem nenhum ponto (.) após a arroba (@)!!!"
    Exit Function
  End If

  If InStr(cEMAIL, "..") > InStr(cEMAIL, "@") Then
    CheckEmail = False
    If lMES Then MsgBox "O e-mail contém pontos consecutivos (..) após o arroba (@)!!!"
    Exit Function
  End If

  ' Substitui o loop For...Next de validação de caracteres por RegEx
  Dim objRegExp As Object
  Set objRegExp = CreateObject("VBScript.RegExp")
  objRegExp.IgnoreCase = True
  objRegExp.Global = False
  
  ' O padrão [^...] significa "procurar qualquer coisa que NÃO seja o que está aqui dentro"
  ' Mantivemos a exata mesma lista de caracteres permitidos do seu código original
  objRegExp.Pattern = "[^a-z0-9\-\.\_\/\~\@]"
  
  ' Se o RegEx encontrar qualquer caractere fora da lista permitida, ele falha
  If objRegExp.Test(cEMAIL) Then
    CheckEmail = False
    If lMES Then MsgBox "Foi digitado um caracter inválido no e-mail!!!"
    Set objRegExp = Nothing
    Exit Function
  End If

  Set objRegExp = Nothing
  CheckEmail = True

End Function
Public Function FormataHora(ByVal dTime As Variant, Optional ByVal cCEP As String = ":") As String
  dTime = Format(CStr(Int(dTime)), "00") & cCEP & Format(CStr(Int((dTime - Int(dTime)) * 100)), "00")
  FormataHora = CStr(dTime)
End Function

Public Function checkcep(ByVal cCEP As String, Optional ByVal lMES As Boolean = True) As Boolean
  checkcep = False
  cCEP = funNumeroPuro(cCEP)
  If FixInt(cCEP) = 0 Then
    If lMES Then
      Alert ("Cep Nao digitado")
    End If
    Exit Function
  End If
  If Len(cCEP) <> 8 Then
    If lMES Then
      Alert ("Cep Nao Tem 8 Digitos")
    End If
    Exit Function
  End If
  If Left(cCEP, 5) = "00000" Then
    If lMES Then
      Alert ("Cep Comeca com 00000")
    End If
    Exit Function
  End If
  checkcep = True
End Function

Public Function FormataIE(ByVal Valor As String, ByVal cUF As String, Optional ByVal cPESSOA As String = "J") As String
    Dim cMASK As String
    
    Valor = funNumeroPuro(Valor)
    FormataIE = Valor
    
    If Left(UCase(Valor), 5) = "ISENT" Or (InStr(UCase(Valor), "NAO") > 0 And InStr(UCase(Valor), "CONTRIB") > 0) Then
        Exit Function ' Não formata Isentos ou Não Contribuintes[cite: 6]
    End If
    If cPESSOA = "F" Or Len(Valor) = 0 Or cUF = "EX" Or cUF = "XX" Then
        Exit Function
    End If
    
    ' Máscaras unificadas do VB6 e atualizações do Harbour[cite: 1, 6]
    Select Case cUF
        Case "AC": cMASK = "00|000|000/00000"
        Case "AL", "AP", "MA": cMASK = "000000000"
        Case "AM", "PB", "PI", "SE", "GO", "MS": cMASK = "00|000|0000"
        Case "BA"
            If Len(Valor) = 8 Then cMASK = "00000000" Else cMASK = "000|000|000"
        Case "DF": cMASK = "00|000000|00000"
        Case "MG": cMASK = "000000000|00|00"
        Case "MT": cMASK = "00000000000"
        Case "PA", "CE", "RR": cMASK = "00|0000000"
        Case "PE"
            If Len(Valor) > 9 Then cMASK = "00|0|000|00000000" Else cMASK = "000000000"
        Case "RN"
            If Len(Valor) > 9 Then cMASK = "00|0|000|0000" Else cMASK = "00|000|0000"
        Case "RO"
            If Len(Valor) = 9 Then cMASK = "000|000000" Else cMASK = "00000000000000"
        Case "SC": cMASK = "000|000|000"
        Case "SP": cMASK = "000|000|000|000"
        Case "TO"
            If Len(Valor) = 11 Then cMASK = "00|00|0000000" Else cMASK = "00|000|0000"
        Case "ES", "RJ": cMASK = "000|000|000"
        Case "PR": cMASK = "000|0000000"
        Case "RS": cMASK = "000/0000000"
    End Select
    
    If Len(cMASK) > 0 Then
        Valor = Format(Valor, cMASK)
        FormataIE = Replace(Valor, "|", ".")
    End If
End Function
Public Function validacnh(ByVal eVALOR As Variant) As Boolean
    Dim cCNH As String
    cCNH = funNumeroPuro(CStr(eVALOR))
    
    If Len(cCNH) = 9 Or Len(cCNH) = 11 Then
        validacnh = True
    Else
        validacnh = False
    End If
End Function

Public Function CheckCNHCat(ByVal cCNHCAT As String) As Boolean
    Dim aCAT As Variant
    Dim i As Integer, bValido As Boolean
    
    aCAT = Array("A", "B", "C", "D", "E", "AB", "AC", "AD", "AE", "ACC")
    cCNHCAT = Trim(UCase(cCNHCAT))
    
    If Len(cCNHCAT) = 0 Then
        CheckCNHCat = True
        Exit Function
    End If
    
    bValido = False
    For i = LBound(aCAT) To UBound(aCAT)
        If cCNHCAT = aCAT(i) Then
            bValido = True
            Exit For
        End If
    Next i
    
    CheckCNHCat = bValido
End Function
Public Function FeriadoMovel(pdatPesquisa As Date) As Boolean
  Dim datData As Date
  Dim ano As Integer
  Dim A As Integer
  Dim B As Integer
  Dim C As Integer
  Dim D As Integer
  Dim E As Integer
  Dim F As Integer
  Dim G As Integer
  Dim h As Integer
  Dim i As Integer
  Dim k As Integer
  Dim L As Integer
  Dim m As Integer
  Dim P As Integer
  Dim Q As Integer
  FeriadoMovel = False
  ano = Year(pdatPesquisa)
  A = (ano Mod 19)
  B = Int(ano / 100)
  C = (ano Mod 100)
  D = Int(B / 4)
  E = (B Mod 4)
  F = Int((B + 8) / 25)
  G = Int((B - F + 1) / 3)
  h = ((19 * A + B - D - G + 15) Mod 30)
  i = Int(C / 4)
  k = (C Mod 4)
  L = ((32 + 2 * E + 2 * i - h - k) Mod 7)
  m = Int((A + 11 * h + 22 * L) / 451)
  P = Int((h + L - 7 * m + 114) / 31)
  Q = ((h + L - 7 * m + 114) Mod 31)
  datData = Format(Str(Q + 1) & "/" & Str(P) & "/" & Str(ano), "dd/mm/yyyy")
  If pdatPesquisa = datData Then
    '+ Pascoa
    FeriadoMovel = True
  End If
  If pdatPesquisa = (datData - 47) Then
    '+ Carnaval
    FeriadoMovel = True
  End If
  If pdatPesquisa = (datData - 60) Then
    '+ Corpus Christi
    FeriadoMovel = True
  End If
  If pdatPesquisa = (datData - 2) Then
    '+ Sexta-Feira Santa
    FeriadoMovel = True
  End If
End Function

Public Function chkufcep(ByVal cCEP As String, ByVal cUF As String, Optional ByVal lMES As Boolean = True) As Boolean
    chkufcep = True
    cUF = UCase(Trim(cUF))
    
    If cep2uf(cCEP) <> cUF Then
        chkufcep = False
        If lMES Then MsgBox "CEP: " & cCEP & " não é da UF: " & cUF, vbExclamation, "Validação de CEP/UF"
    End If
End Function

Public Function cep2uf(ByVal cepuso As String) As String
    Dim nCep As Long
    
    cepuso = funNumeroPuro(cepuso)
    nCep = Val(cepuso)
    cep2uf = "EX"
    
    Select Case nCep
        Case 1000000 To 19999999: cep2uf = "SP"
        Case 20000000 To 28999999: cep2uf = "RJ"
        Case 29000000 To 29999999: cep2uf = "ES"
        Case 30000000 To 39999999: cep2uf = "MG"
        Case 40000000 To 48999999: cep2uf = "BA"
        Case 49000000 To 49999999: cep2uf = "SE"
        Case 50000000 To 56999999: cep2uf = "PE"
        Case 57000000 To 57999999: cep2uf = "AL"
        Case 58000000 To 58999999: cep2uf = "PB"
        Case 59000000 To 59999999: cep2uf = "RN"
        Case 60000000 To 63999999: cep2uf = "CE"
        Case 64000000 To 64999999: cep2uf = "PI"
        Case 65000000 To 65999999: cep2uf = "MA"
        Case 66000000 To 68899999: cep2uf = "PA"
        Case 68900000 To 68999999: cep2uf = "AP"
        Case 69000000 To 69299999, 69400000 To 69899999: cep2uf = "AM"
        Case 69300000 To 69399999: cep2uf = "RR"
        Case 69900000 To 69999999: cep2uf = "AC"
        Case 70000000 To 72799999, 73000000 To 73699999: cep2uf = "DF"
        Case 72800000 To 72999999, 73700000 To 76799999: cep2uf = "GO"
        Case 77000000 To 77999999: cep2uf = "TO"
        Case 78000000 To 78899999: cep2uf = "MT"
        Case 78900000 To 78999999: cep2uf = "RO"
        Case 79000000 To 79999999: cep2uf = "MS"
        Case 80000000 To 87999999: cep2uf = "PR"
        Case 88000000 To 89999999: cep2uf = "SC"
        Case 90000000 To 99999999: cep2uf = "RS"
        Case Else: cep2uf = "EX"
    End Select
End Function
Public Function CNPJCPFVAL(ByVal cCGC As String, ByVal cPESSOA As String, Optional ByVal cESTADO As String = "") As Boolean
    Dim lRETU As Boolean
    lRETU = True
    
    cPESSOA = UCase(Trim(cPESSOA))
    
    Select Case cPESSOA
        Case "J" ' CNPJ
            lRETU = CheckCNPJ(cCGC, "X", True, cESTADO)
        Case "F" ' CPF
            lRETU = CheckCPF(cCGC, True)
        Case "C" ' CEI
            lRETU = checkCEI(cCGC)
        Case Else ' Tipo em branco: tenta validar por qualquer um dos três[cite: 2]
            lRETU = CheckCNPJ(cCGC, "X", False, cESTADO) Or CheckCPF(cCGC, False) Or checkCEI(cCGC)
    End Select
    
    CNPJCPFVAL = lRETU
End Function
Public Sub PreencheCboEstados(ByRef cbo)
  cbo.Clear
  cbo.AddItem "AC"
  cbo.AddItem "AL"
  cbo.AddItem "AM"
  cbo.AddItem "AP"
  cbo.AddItem "BA"
  cbo.AddItem "CE"
  cbo.AddItem "DF"
  cbo.AddItem "ES"
  cbo.AddItem "GO"
  cbo.AddItem "MA"
  cbo.AddItem "MG"
  cbo.AddItem "MR"
  cbo.AddItem "MS"
  cbo.AddItem "MT"
  cbo.AddItem "PA"
  cbo.AddItem "PB"
  cbo.AddItem "PE"
  cbo.AddItem "PI"
  cbo.AddItem "PR"
  cbo.AddItem "RJ"
  cbo.AddItem "RN"
  cbo.AddItem "RO"
  cbo.AddItem "RR"
  cbo.AddItem "RS"
  cbo.AddItem "SC"
  cbo.AddItem "SE"
  cbo.AddItem "SP"
  cbo.AddItem "TO"
  cbo.AddItem "EX"
End Sub

Public Function PEGDDD(ByVal cTEL As String) As String
  cTEL = formataTelefone(cTEL)
  PEGDDD = ""
  If InStr(cTEL, "(") > 0 Then
    PEGDDD = Mid(cTEL, 2, 2)
  End If
End Function

Public Function PEGTEL(ByVal cTEL As String) As String
  cTEL = formataTelefone(cTEL)
  PEGTEL = cTEL
  If InStr(cTEL, "(") > 0 Then
    PEGTEL = Mid(cTEL, 5)
  End If
End Function

Public Function PEGPREF(ByVal cTEL As String) As String
'cTEL = tem que ser so o numero telefone antes de chamar usar pegtel se necessario
  PEGPREF = ""
  If InStr(cTEL, "-") > 0 Then
    PEGPREF = Mid(cTEL, 1, InStr(cTEL, "-") - 1)
  End If
End Function

Function checkCEI(ByVal pCEI As String) As Boolean
  Dim nTot As Byte
  Dim cAux As String
  Dim lret As Boolean
  Dim i As Byte
  Dim pNu_cei As String

  checkCEI = False

  pCEI = TiraOut(pCEI)

  nTot = 0
  cAux = ""
  lret = True
  i = 0
  pNu_cei = ""

  pNu_cei = Trim$(pCEI)

  If Len(pNu_cei) < 12 Then
    Exit Function
  End If

  For i = 1 To 11
    nTot = nTot + Val(Mid$(pNu_cei, i, 1)) * Val(Mid$("74185216374", i, 1))
  Next

  cAux = Right$(Str(nTot), 2)
  nTot = Val(Left$(cAux, 1)) + Val(Right$(cAux, 1))
  nTot = IIf(nTot > 9, 0, 10 - nTot)
  lret = IIf(Val(Right$(pNu_cei, 1)) = nTot, True, False)

  If lret = False Then
    MsgBox "Codigo CEI incorreto - Tecle <Enter>"
  Else
    checkCEI = True
  End If

End Function

Public Function Validar_CartaoCredito(ByVal CartaoCredito As String, Optional ByVal lMES As Boolean = False, Optional ByRef cBandeira As String) As Boolean
    Dim nSum As Integer, i As Integer, nDigit As Integer
    Dim lEven As Boolean, nLen As Integer
    Dim cPref2 As Integer, cPref4 As Integer, cPref6 As Long

    CartaoCredito = funNumeroPuro(CartaoCredito)
    nLen = Len(CartaoCredito)
    If nLen = 0 Then Exit Function

    lEven = False
    For i = nLen To 1 Step -1
        nDigit = Val(Mid(CartaoCredito, i, 1))
        If lEven Then
            nDigit = nDigit * 2
            If nDigit > 9 Then nDigit = nDigit - 9
        End If
        nSum = nSum + nDigit
        lEven = Not lEven
    Next

    If (nSum Mod 10 = 0) And (nSum > 0) Then
        cPref2 = Val(Left(CartaoCredito, 2))
        cPref4 = Val(Left(CartaoCredito, 4))
        cPref6 = Val(Left(CartaoCredito, 6))

        Select Case True
            Case (cPref2 = 34 Or cPref2 = 37) And nLen = 15: cBandeira = "AmEx"
            Case Left(CartaoCredito, 1) = "4": cBandeira = "Visa"
            Case (cPref2 >= 51 And cPref2 <= 55) Or (cPref6 >= 222100 And cPref6 <= 272099): cBandeira = "MasterCard"
            Case cPref4 = 6011 Or (Val(Left(CartaoCredito, 3)) >= 644 And Val(Left(CartaoCredito, 3)) <= 649): cBandeira = "Discover"
            Case Val(Left(CartaoCredito, 3)) >= 300 And Val(Left(CartaoCredito, 3)) <= 305: cBandeira = "Carte Blanche"
            Case cPref2 = 30 Or cPref2 = 36 Or cPref2 = 38: cBandeira = "Diners"
            Case cPref4 >= 3528 And cPref4 <= 3589: cBandeira = "JCB"
            Case cPref4 = 4011 Or cPref4 = 4389 Or cPref4 = 4514 Or cPref4 = 4576 Or cPref4 = 5041 Or cPref4 = 5066 Or cPref4 = 5090 Or cPref4 = 6277 Or cPref4 = 6362: cBandeira = "Elo"
            Case (cPref2 = 38 Or cPref2 = 60 Or cPref2 = 62) And nLen > 13: cBandeira = "Hipercard"
            Case Else: cBandeira = "Desconhecida"
        End Select
        
        Validar_CartaoCredito = True
        If lMES Then MsgBox "Bandeira: " & cBandeira, vbInformation, "Validação de Cartão"
    Else
        Validar_CartaoCredito = False
        cBandeira = ""
        If lMES Then MsgBox "Número de Cartão Inválido", vbExclamation, "Validação de Cartão"
    End If
End Function
Public Function Validar_CodigoBarras(ByVal CodigoBarras As String) As Boolean
  Dim bolRetorno As Boolean

  bolRetorno = False

  If Len(CodigoBarras) > 0 Then
    'Retiramos possÝveis mîscaras
    CodigoBarras = funNumeroPuro(CodigoBarras)

    If IsNumeric(CodigoBarras) Then

      Dim strNumeros As String
      Dim strDV As String
      Dim intTamanho As Integer

      strNumeros = Left(CodigoBarras, Len(CodigoBarras) - 1)
      strDV = Right(CodigoBarras, 1)

      intTamanho = Len(strNumeros)

      Select Case intTamanho
      Case 7, 11, 12, 13

        Dim intCont As Integer
        Dim intImpares As Integer
        Dim intPares As Integer
        Dim intResultado As Integer

        For intCont = 1 To intTamanho
          If (intCont And Not -2) = 0 Then
            intPares = intPares + Val(Mid(strNumeros, intCont, 1))
          Else
            intImpares = intImpares + Val(Mid(strNumeros, intCont, 1))
          End If
        Next

        If (intTamanho = 7) Or (intTamanho = 11) Or (intTamanho = 13) Then  'EAN-8 (8d - 1dv) ou UPC-A (12d - 1dv) ou EAN-14 (14d - 1dv)
          intImpares = intImpares * 3 + intPares
          intPares = Int((intImpares + 9) / 10) * 10
          intResultado = intPares - intImpares
        Else                             'EAN-13 (13d - 1dv)
          intPares = intPares * 3 + intImpares
          intImpares = Int((intPares + 9) / 10) * 10
          intResultado = intImpares - intPares
        End If

        bolRetorno = (CStr(intResultado) = strDV)

      End Select

    End If
  End If

  Validar_CodigoBarras = bolRetorno
End Function


