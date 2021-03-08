Attribute VB_Name = "checagem"
Option Explicit

Public Function IsValidIP(Test As String) As Boolean
    Dim SubNets() As String
    Dim i As Integer
    If LCase(Test) = "localhost" Then
        IsValidIP = True
        Exit Function
    End If
    If Len(Test) > 16 Then
        IsValidIP = False
        Exit Function
    End If
    SubNets = Split(Test, ".")
    If UBound(SubNets) > 3 Then
        IsValidIP = False
        Exit Function
    End If
    For i = 0 To 3
        If Not IsNumeric(SubNets(i)) Or SubNets(i) < 0 Or SubNets(i) > 255 Then
            IsValidIP = False
            Exit Function
        End If
    Next
    IsValidIP = True
End Function

Public Function formataTelefone(ByVal cNUMERO) As String
    'Dim cTMP As String
    'formataTelefone ("99-12435678")->(99)1234-5678
    'formataTelefone ("99,12345678")->(99)1234-5678
    'formataTelefone ("099-12345678")->(99)1234-5678
    'formataTelefone ("(099)12345678")->(99)1234-5678
    'formataTelefone ("12345678")->1234-5678
    'formataTelefone ("1234567")->123-4567
    'formataTelefone ("99 12345678")->(99)1234-5678
    'formataTelefone ("099 12345678")->(99)1234-5678
    'formataTelefone ("(99)12345678")->(99)1234-5678
    'formataTelefone ("(99)1234567")->(99)123-4567
    If Mid(cNUMERO, 1, 4) = "0300" Or Mid(cNUMERO, 1, 4) = "0800" Then
        formataTelefone = cNUMERO
        Exit Function
    End If
    If Mid(cNUMERO, 1, 1) = "(" And Mid(cNUMERO, 4, 1) = ")" Then '(99)12345678
        cNUMERO = Mid(cNUMERO, 1, 4) & FormataTelfon2(Mid(cNUMERO, 5))
    End If
    If Mid(cNUMERO, 3, 1) = "-" Then             '99-12345678
        cNUMERO = "(" & Mid(cNUMERO, 1, 2) & ")" & FormataTelfon2(Mid(cNUMERO, 4))
    End If
    If Mid(cNUMERO, 3, 1) = "," Then             '99,12345678
        cNUMERO = "(" & Mid(cNUMERO, 1, 2) & ")" & FormataTelfon2(Mid(cNUMERO, 4))
    End If
    If Mid(cNUMERO, 3, 1) = " " Then             '99 12345678
        cNUMERO = "(" & Mid(cNUMERO, 1, 2) & ")" & FormataTelfon2(Mid(cNUMERO, 4))
    End If
    If Mid(cNUMERO, 4, 1) = "-" And Mid(cNUMERO, 1, 1) = "0" Then '099-12345678
        cNUMERO = "(" & Mid(cNUMERO, 2, 2) & ")" & FormataTelfon2(Mid(cNUMERO, 5))
    End If

    If Mid(cNUMERO, 4, 1) = " " And Mid(cNUMERO, 1, 1) = "0" Then '099 12345678
        cNUMERO = "(" & Mid(cNUMERO, 2, 2) & ")" & FormataTelfon2(Mid(cNUMERO, 5))
    End If
    If Mid(cNUMERO, 4, 1) = "-" And Mid(cNUMERO, 1, 1) = "0" Then '99-12345678
        cNUMERO = "(" & Mid(cNUMERO, 2, 2) & ")" & FormataTelfon2(Mid(cNUMERO, 5))
    End If
    If Mid(cNUMERO, 1, 1) = "(" And Mid(cNUMERO, 2, 1) = "0" And Mid(cNUMERO, 5, 1) = ")" Then '(099)12345678
        cNUMERO = "(" & Mid(cNUMERO, 3, 2) & ")" & FormataTelfon2(Mid(cNUMERO, 6))
    End If
    formataTelefone = FormataTelfon2(cNUMERO)
End Function

Public Function FormataTelfon2(ByVal cNUMERO) As String
    FormataTelfon2 = cNUMERO
    cNUMERO = funNumeroPuro(cNUMERO)
    If Len(cNUMERO) = 8 And Mid(cNUMERO, 4, 1) <> "-" Then '123-4567
        cNUMERO = Mid(cNUMERO, 1, 4) & "-" & Mid(cNUMERO, 5)
        FormataTelfon2 = cNUMERO
    End If
    If Len(cNUMERO) = 7 Then
        cNUMERO = Mid(cNUMERO, 1, 3) & "-" & Mid(cNUMERO, 4)
        FormataTelfon2 = cNUMERO
    End If
End Function

Public Function FormataRG(ByVal Valor)
    'Dim cRETU As String
    Dim nPOS As String
    Dim cDAC As String
    FormataRG = FixStr(Valor)
    cDAC = ""
    If Valor = "ISENTO" Then
        Exit Function
    End If
    Valor = Trim(Valor)
    nPOS = InStr(Valor, "-")
    If nPOS = 0 Then
        Valor = TiraOut(Valor)
        Valor = Trim(Valor)
        If Len(Valor) >= 9 Then
            cDAC = Mid(Valor, 9, 1)
        End If
    Else
        cDAC = Mid(Valor, nPOS + 1, 1)
        Valor = Mid(Valor, 1, nPOS - 1)
        If cDAC = "x" Then cDAC = "X"
        If cDAC <> "X" Then                      ''evita erros como -- -. -/ caracter inves de numero no dac
            cDAC = funNumeroPuro(cDAC)
        End If
    End If
    Valor = funNumeroPuro(Valor)
    Select Case Len(Valor)
        ''Case Is >= 8 ric=11
    Case 8
        FormataRG = Trim(Mid(Valor, 1, 2) + "." + Mid(Valor, 3, 3) + "." + Mid(Valor, 6))
    Case 7
        FormataRG = Trim(Mid(Valor, 1, 1) + "." + Mid(Valor, 2, 3) + "." + Mid(Valor, 5))
    Case Else
        FormataRG = Valor
    End Select
    If cDAC <> "" Then
        FormataRG = FormataRG & "-" & cDAC
    End If
    If Left(FormataRG, 1) = "0" Then
        FormataRG = Mid(FormataRG, 2)
    End If
    If Left(FormataRG, 1) = "." Then
        FormataRG = Mid(FormataRG, 2)
    End If
End Function

Public Function checkcfo(ByVal cCFO As String, ByVal cUFORI As String, ByVal cUFDES As String, ByVal cTIPO As String) As Boolean
    Dim cINISP, cINIFORA, cINIEXP
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
    Dim total As String
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
    total = 0

    For i = 1 To 10
        total = total + Val(Mid(cPIS, i, 1)) * Val(Mid(ftap, i, 1))
    Next i

    resto = Int(total Mod 11)

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

Public Function CheckIE(cIE As Variant, cUF As Variant, Optional lMES As Boolean = True) As Boolean
    Dim nRETU As Variant
    cIE = TiraOut(cIE)
    cUF = FixStr(cUF)
    If cUF = "EX" Or cUF = "XX" Then
        If cIE = "ISENTO" Or cIE = "00000000000000" Then
            If lMES Then Alert "Inscricao OK-ISENTO-EXPORTACAO", "Checagem IE"
            CheckIE = True
        Else
            If lMES Then Alert "Inscricao Invalida Use ISENTO para exportacao", "Checagem IE"
            CheckIE = False
        End If
    Else
        '  //criar nova funcao dll nao mais atende novas checagens alguns estados
        CheckIE = True
    End If
End Function

Public Function CheckCNPJ(cCGC As Variant, Optional cTIPO As String = "X", Optional lMES As Boolean = True, Optional cUF As String = "") As Boolean
    ''cTIPO= M-Matriz F-Filial X-Nao Checar
    Dim X As Integer
    Dim aUF As Variant

    CheckCNPJ = False
    cCGC = Trim(TiraOut(cCGC))
    If Len(cCGC) = 0 Then
        If lMES Then
            Alert ("CPNJ Numero Não Informado Em Branco")
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
            If lMES Then Alert ("CNPJ Invalido - Sequencia Repetitiva de " + str(X))
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


    If cTIPO = "M" Then
        If Mid(cTIPO, 9, 4) <> "0001" Then
            If lMES Then
                Alert ("CNPJ Invalido-Nao e Matriz")
            End If
            Exit Function
        End If
    End If
   
    If Len(cUF) > 0 Then
        aUF = Array("AC", "AL", "AM", "AP", "BA", "CE", "DF", "ES", "GO", "MA", "MG", "MS", "MT", "PA", "PB", "PE", "PI", "PR", "RJ", "RN", "RO", "RR", "RS", "SC", "SE", "SP", "TO", "EX", "XX")
        If FastArraySearch(aUF, cUF) = -1 Then
            If lMES Then
                Alert ("Estado Invalido")
            End If
        End If
    End If
   
   
    If Mod11(cCGC, 13, 9) Then
        If Mod11(cCGC, 14, 9) Then
            CheckCNPJ = True
        Else
            If lMES Then Alert ("CNPJ Invalido - Cheque 14 Digito - 2 Verificador")
            Exit Function
        End If
    Else
        If lMES Then Alert ("CNPJ Invalido - Cheque 13 Digito - 1 Verificador")
        Exit Function
    End If
End Function

Public Function Mod11(ByVal campo, ByVal posdv, ByVal pesomax) As Boolean
    Dim Dv, peso, imod, rest As Integer
    Mod11 = False
    Dv = 0
    peso = 1
    For imod = posdv To 1 Step -1
        Dv = Dv + peso * Val(Mid(campo, imod, 1))
        peso = peso + 1
        If peso > pesomax Then
            peso = 2
        End If
    Next
    rest = Dv Mod 11
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
        If InStr("0123456789", Mid(pCGC, i, 1)) > 0 Then
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

Public Function CheckCNPJIE(ByVal cPESSOA, ByVal CCNPJ, ByVal cIE, ByVal cUF, Optional ByVal lMES As Boolean = True) As Boolean
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
    Dim P1 As String
    CheckCPF = False
    P1 = Trim(TiraOut(xCPF))
    If Len(P1) = 0 Then
        If lMES Then Alert ("CPF Invalido - Em Branco")
        Exit Function
    End If
    If Len(P1) < 11 Then
        If lMES Then Alert ("CPF Invalido - Menos de 11 Digitos")
        Exit Function
    End If
    If Len(P1) > 11 Then
        If lMES Then Alert ("CPF Invalido - Mais de 11 Digitos")
        Exit Function
    End If

    For X = 0 To 9
        If P1 = String(11, CStr(X)) Then
            If lMES Then Alert ("CPF Invalido - Sequencia Repetitiva de " + str(X))
            Exit Function
        End If
    Next X
    If Mod11(P1, 10, 10) Then
        If Mod11(P1, 11, 11) Then
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
    Dim dv_d3, tam_chave, xrdig, y1, dv_base, X, y2, dv_p2
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
    dv_p2 = Trim(str(y2))
    If dv_p2 <> xrdig Then
        Alert ("Codigo do RENAVAN invalido.")
        Exit Function
    End If
    CheckRena = True
End Function

Function TeclaEnter(KeyIN As Integer)
    TeclaEnter = KeyIN
    If KeyIN = 13 Or KeyIN = 40 Then
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

Function ValiText(KeyIN As Integer, Optional ByVal eTipo As String = "#N", _
                  Optional ByVal Editable As Boolean = True, _
                  Optional ByVal eSUBTIPO As String = "", _
                  Optional ByVal lENTER As Boolean = True)
    Dim ValidateList As String
    Dim KeyOut As Integer
    If lENTER Then
        If TeclaEnter(KeyIN) = 0 Then
            ValiText = 0
            Exit Function
        End If
    End If
    Select Case eTipo
    Case "#HR"
        ValidateList = "0123456789.:"
    Case "#NI"                                   ''Numeros Inteiros
        ValidateList = "0123456789"
    Case "#NI-"                                  ''Numeros Inteiros e Negativos
        ValidateList = "0123456789-"
    Case "#NP"                                   ''Numeros Inteiros e o Ponto
        ValidateList = "0123456789."
    Case "#N-.X"                                 ''Numeros Traco  Ponto X (RG)"
        ValidateList = "0123456789.-X"
    Case "#N-./"                                 ''Numero - e Ponto (CGC/CNPJ,CPF,CIC)
        ValidateList = "0123456789.-/"
    Case "#N-"                                   ''Numero e -   (Telefone CEP)
        ValidateList = "0123456789-"
    Case "#D", "N/"                              ''Numero e / - (Data)
        ValidateList = "0123456789/"
    Case "#N", "#N.,"                            ''Numeros . ,
        ValidateList = "0123456789.,"
    Case "#NEG", "#N.,-"                         ''Numeros . ,
        ValidateList = "0123456789.,-"
    Case "#SN"                                   ''Sim Nao
        ValidateList = "SN"
    Case "#C", "#CU", "#CL"                      ''SoLetras SoMaisculas SoMinusculas
        ValidateList = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Case "#CN", "#CNU", "#CNL"                   ''letras numeros
        ValidateList = "01234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ"
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

Public Function CheckRG(ByVal Valor, Optional ByVal lMES As Boolean = True) As Boolean
    Dim n1, n2, n3, n4, n5, n6, n7, n8, d
    Dim soma
    Dim nPOS
    Dim cDAC
    Dim zDAC
    Dim lRETU
    Dim ZNERRO
    Dim ZERRO
    'IF ValType(lMES)#"L"
    'lmes:=.T.
    'End If
    zDAC = " "
    ZNERRO = 0

    CheckRG = True
    Valor = Replace(Valor, ".", "")              'tiraout tambem tira - nao pode ser usada
    If Valor = "ISENTO" Or InStr(Valor, "RNE") > 0 Then 'isento ou registro nacional de estrangeiro
        Exit Function
    End If
    If Len(Valor) = 0 Then
        Exit Function
    End If
    nPOS = InStr(Valor, "-")
    If nPOS = 0 Then
        cDAC = " "
    Else
        cDAC = Mid(Valor, nPOS + 1, 1)
        Valor = Mid(Valor, 1, nPOS - 1)
    End If
    Valor = str(Val(Valor))
    If Len(Trim(Valor)) <= 7 Then
        CheckRG = False
        ZERRO = "RG com Menos de 7 Digitos"
        ZNERRO = 3
        If lMES Then
            Alert (ZERRO)
        End If
        Exit Function
    End If
    Valor = StrZero(Val(Valor), 8)
    n1 = Val(Mid(Valor, 1, 1)) * 9
    n2 = Val(Mid(Valor, 2, 1)) * 8
    n3 = Val(Mid(Valor, 3, 1)) * 7
    n4 = Val(Mid(Valor, 4, 1)) * 6
    n5 = Val(Mid(Valor, 5, 1)) * 5
    n6 = Val(Mid(Valor, 6, 1)) * 4
    n7 = Val(Mid(Valor, 7, 1)) * 3
    n8 = Val(Mid(Valor, 8, 1)) * 2
    soma = n1 + n2 + n3 + n4 + n5 + n6 + n7 + n8
    d = soma Mod 11
    If cDAC = "X" Or cDAC = "x" Or cDAC = " " Then
        If d = 10 Then
            Exit Function
        Else
            zDAC = StrZero(d, 1)
            CheckRG = False
            ZNERRO = 5
            ZERRO = "Digito de Controle RG  " + cDAC + "  Nao Confere sugerido: " + zDAC
            If lMES Then
                Alert (ZERRO)
            End If
        End If
    Else
        If d = Val(cDAC) Or d = 0 Then
        Else
            CheckRG = False
            If d = 10 Then
                zDAC = "X"
                ZERRO = "Digito de Controle RG Nao Confere sugerido: X"
            Else
                zDAC = StrZero(d, 1)
                ZERRO = "Digito de Controle RG Nao Confere sugerido: " + zDAC
            End If
            ZNERRO = 4
            If lMES Then
                Alert (ZERRO)
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
    ''A Posicao comeca com zero
    ''Porem o tamanho comeca com 1
    Dim nTAM As Integer
    Dim nLEN As Integer
    If lENTER Then
        If TeclaEnter(KeyAscii) = 0 Then
            ValidaDados = 0
            Exit Function
        End If
    End If

    ''nLEN = Len(Screen.ActiveControl)
    Select Case eTipo
    Case "PLACA"                                 ''###-####
        nTAM = 8
    Case "HORA"                                  ''##.##
        nTAM = 5
    Case "HORAL"                                 ''##.##.##
        nTAM = 8
    Case "CEP"                                   ''#####-###
        nTAM = 9
    Case "TEL"                                   ''####-####
        nTAM = 9
    Case "DC"                                    '' DD/MM/YY Data curta 8 digitos
        nTAM = 8
    Case "DL"                                    '' DD/MM/YYYY Data Longa 10 digitos
        nTAM = 10
    Case "CPF", "CIC"                            '' 111.222.333-44   14 Digitos
        nTAM = 14
    Case "CNPJ", "CGC"                           '' 11.222.333/4444-55 18 digitos
        nTAM = 18
    End Select
    nLEN = Len(Trim(FixStr(Screen.ActiveControl.Text)))
    If KeyAscii <> 8 And nLEN >= nTAM And _
       KeyAscii <> 13 And Screen.ActiveControl.SelStart > 0 _
       And Screen.ActiveControl.SelLength = 0 Then
        KeyAscii = 0
    End If
    If KeyAscii <> 13 And KeyAscii <> 8 And eTipo <> "PLACA" Then
        KeyAscii = ValiText(KeyAscii, "#NI", True) ''Numeros Inteiros pois o
        ''a funcao poem a barra tracos pontos sozinha abaixo
    End If
    If KeyAscii <> 8 Then
        Select Case eTipo                        ''A posicao de referencia e sempre uma antes do ponto barra traco
        Case "HORA", "HORAL"                     ''PONTO na posicao 3
            If nLEN = 0 Then
                If KeyAscii > 50 Then            ''Hora so pode ser 012
                    KeyAscii = 0
                End If
            End If
            If nLEN = 3 Or nLEN = 6 Then
                If KeyAscii > 53 Then            ''minutos/segundos so pode ser ate 59
                    KeyAscii = 0
                End If
            End If
            If nLEN = 2 Or (nLEN = 5 And eTipo = "HORAL") Then
                Screen.ActiveControl.Text = Screen.ActiveControl.Text & "."
                Screen.ActiveControl.SelStart = Len(Screen.ActiveControl.Text) + 1
                Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 1
            End If
            
        Case "PLACA"                             ''traco na posicao 4
            If nLEN = 3 Then
                Screen.ActiveControl.Text = Screen.ActiveControl.Text & "-"
                Screen.ActiveControl.SelStart = Len(Screen.ActiveControl.Text) + 1
                Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 1
                nLEN = nLEN + 1
            End If
            If nLEN < 3 Then
                KeyAscii = ValiText(KeyAscii, "#CU")
            End If
            If nLEN > 3 Then
                KeyAscii = ValiText(KeyAscii, "#NI")
            End If
            
            
            
        Case "DC", "DL"                          ''Barra Posicao 3 e 6
            If nLEN = 0 Then
                If KeyAscii > 51 Then            ''1 digito dia so pode ser 0123
                    KeyAscii = 0
                End If
            End If
            If nLEN = 3 Then
                If KeyAscii > 49 Then            ''1 digito mes so pode ser 01
                    KeyAscii = 0
                End If
            End If
            If nLEN = 2 Or nLEN = 5 Then
                Screen.ActiveControl.Text = Screen.ActiveControl.Text & "/"
                Screen.ActiveControl.SelStart = Len(Screen.ActiveControl.Text) + 1
                Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 1
            End If
        Case "CEP"                               ''traco na posicao 6
            If nLEN = 5 Then
                Screen.ActiveControl.Text = Screen.ActiveControl.Text & "-"
                Screen.ActiveControl.SelStart = Len(Screen.ActiveControl.Text) + 1
                Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 1
            End If
        Case "TEL"                               ''traco na posicao 5
            If nLEN = 4 Then
                Screen.ActiveControl.Text = Screen.ActiveControl.Text & "-"
                Screen.ActiveControl.SelStart = Len(Screen.ActiveControl.Text) + 1
                Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 1
            End If
        Case "CPF", "CIC"                        ''ponto 4 e 8
            If nLEN = 3 Or nLEN = 7 Then
                Screen.ActiveControl.Text = Screen.ActiveControl.Text & "."
                Screen.ActiveControl.SelStart = Len(Screen.ActiveControl.Text) + 1
                Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 1
            End If
            If nLEN = 11 Then                    ''TRaco na 12
                Screen.ActiveControl.Text = Screen.ActiveControl.Text & "-"
                Screen.ActiveControl.SelStart = Len(Screen.ActiveControl.Text) + 1
                Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 1
            End If
        Case "CNPJ", "CGC"                       ''ponto 3 e 7
            If nLEN = 2 Or nLEN = 6 Then
                Screen.ActiveControl.Text = Screen.ActiveControl.Text & "."
                Screen.ActiveControl.SelStart = Len(Screen.ActiveControl.Text) + 1
                Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 1
            End If
            If nLEN = 10 Then                    ''barra na 11
                Screen.ActiveControl.Text = Screen.ActiveControl.Text & "/"
                Screen.ActiveControl.SelStart = Len(Screen.ActiveControl.Text) + 1
                Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 1
            End If
            If nLEN = 15 Then                    ''TRaco na 16
                Screen.ActiveControl.Text = Screen.ActiveControl.Text & "-"
                Screen.ActiveControl.SelStart = Len(Screen.ActiveControl.Text) + 1
                Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text) + 1
            End If
        End Select
    End If
    ValidaDados = KeyAscii
End Function

Public Function ChecaSimNao(ByVal nText As Variant)
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
    'Dim i As Integer
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

Function FormataPlaca(ByVal eCHAPA) As String
    Dim cCHAPA As String
    cCHAPA = FixStr(eCHAPA, "")
    cCHAPA = UCase(cCHAPA)
    cCHAPA = Replace(cCHAPA, " ", "")
    cCHAPA = Replace(cCHAPA, "-", "")
    cCHAPA = Replace(cCHAPA, ".", "")
    If Len(cCHAPA) > 0 Then
        cCHAPA = Mid(cCHAPA, 1, 3) & "-" & Mid(cCHAPA, 4)
    End If
    FormataPlaca = cCHAPA
End Function

Function CheckHora(ByVal eHORA, Optional ByVal lMES As Boolean = False) As Boolean
    Dim cTMP, cHORA, cMIN, cSEG As String
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

Public Function CheckCTA(ByVal cBANCO, ByVal cAGENCIA, ByVal cCONTA) As Boolean
    Dim eTOt As Long
    Dim nFIM As Integer
    Dim nINI As Integer
    Dim X As Integer
    Dim nRES
    eTOt = 0
    cBANCO = StrZero(FixInt(cBANCO), 3)
    If cBANCO <> "033" And cBANCO <> "237" And cBANCO <> "341" Then
        CheckCTA = True
        Exit Function
    End If
    cCONTA = Replace(cCONTA, " ", "")
    cCONTA = TiraOut(cCONTA)
    cAGENCIA = Replace(cAGENCIA, " ", "")
    cAGENCIA = TiraOut(cAGENCIA)
    Select Case cBANCO
    Case "033"
        If Len(cCONTA) <> 9 Then
            Alert ("Quantidade de Digitos da Conta Diferente de 9")
            CheckCTA = False
            Exit Function
        End If
        cAGENCIA = Right(cAGENCIA, 3)
        eTOt = eTOt + CalcDig(7, Mid(cAGENCIA, 1, 1))
        eTOt = eTOt + CalcDig(3, Mid(cAGENCIA, 2, 1))
        eTOt = eTOt + CalcDig(1, Mid(cAGENCIA, 3, 1))
        eTOt = eTOt + CalcDig(9, Mid(cCONTA, 1, 1))
        eTOt = eTOt + CalcDig(7, Mid(cCONTA, 2, 1))
        eTOt = eTOt + CalcDig(1, Mid(cCONTA, 3, 1))
        eTOt = eTOt + CalcDig(3, Mid(cCONTA, 4, 1))
        eTOt = eTOt + CalcDig(1, Mid(cCONTA, 5, 1))
        eTOt = eTOt + CalcDig(9, Mid(cCONTA, 6, 1))
        eTOt = eTOt + CalcDig(7, Mid(cCONTA, 7, 1))
        eTOt = eTOt + CalcDig(3, Mid(cCONTA, 8, 1))
        eTOt = StrZero(eTOt)
        eTOt = Right(eTOt, 1)
        eTOt = IIf(Val(eTOt) > 0, 10 - Val(eTOt), 0)
        If eTOt <> Val(Mid(cCONTA, 9, 1)) Then
            Alert ("Checagem da Conta n„o Confere")
            CheckCTA = False
            Exit Function
        End If
    Case "341"
        If (DAC10(cAGENCIA + Left(cCONTA, 5)) <> Right(cCONTA, 1)) Then
            Alert ("Checagem da Conta n„o Confere")
            CheckCTA = False
            Exit Function
        End If
    Case "237"
        eTOt = 0
        nFIM = Len(cCONTA)
        nINI = nFIM
        nFIM = nFIM - 1
        For X = 1 To nFIM
            eTOt = eTOt + nINI * Val(Mid(cCONTA, X, 1))
            nINI = nINI - 1
        Next X
        nRES = eTOt Mod 11
        nRES = 11 - nRES
        nRES = IIf(nRES = 10, "P", StrZero(nRES, 1))
        If nRES <> Right(cCONTA, 1) Then
            Alert ("Digito de Controle n„o confere")
            CheckCTA = False
            Exit Function
        End If
    End Select
    CheckCTA = True
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
    Dim cNUMERO
    Dim ninicio
    Dim ntotal
    Dim ccpoaux
    Dim X
    ninicio = Len(Trim(Arg1)) + 1
    ntotal = 0
    If (ninicio < 2) Then
        ninicio = 2
    End If
    ccpoaux = "0" + Trim(Arg1)
    For X = ninicio To 1 Step -2
        cNUMERO = Mid(ccpoaux, X, 1)
        ntotal = ntotal + InStr("516273849", cNUMERO)
        ntotal = ntotal + Val(Mid(ccpoaux, X - 1, 1))
    Next
    DAC10 = Trim(str(InStr("987654321", Mid(StrZero(ntotal, 3), 3, 1))))
End Function

Public Function CheckDat2(ByVal eDATA As Variant, Optional ByVal lMES As Boolean = True)
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

Public Function CheckData(ByVal eDATA As Variant, Optional ByVal lMES As Boolean = True)
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
                If Len(cData) = 6 Or Ano_Bissexto(nANO) Then
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

Public Function Ano_Bissexto(ByVal AnoPass As Integer) As Boolean
    If AnoPass / 4 = AnoPass \ 4 Then            ' Se o resultado da divisão real "X/Y"
        Ano_Bissexto = True                      ' for igual ao resultado da divisão inteira
        Exit Function                            ' "X\Y" então é bissexto.
    End If
End Function

Public Function CheckTitulo(ByVal s As String, Optional ByVal lMES As Boolean = True) As Boolean
    Dim DV1 As Integer
    Dim DV2 As Integer
    Dim i As Integer
    'Dim total As Integer
    Dim resto As Integer
    Dim Resto2 As Integer
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
    Resto2 = DV2 Mod 11
    
    If Resto2 = 1 Then
        DV2 = 0
    Else
        If Resto2 = 0 Then
            If (((Mid(Numero, 10, 2)) = "01") Or ((Mid(Numero, 10, 2)) = "02")) Then
                DV2 = 1
            Else
                DV2 = 0
            End If
        Else
            DV2 = 11 - Resto2
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
    Dim nCharacter As Integer
    Dim Count As Integer
    Dim sLetra As String


    cEMAIL = FixStr(cEMAIL)
    ''MsgBox "O e-mail nao digitado !!"
    If Len(cEMAIL) = 0 Then
        If lVazio Then
            CheckEmail = True
        Else
            CheckEmail = False
            If lMES Then
                MsgBox "Email em branco"
            End If
        End If
        Exit Function
    End If
    
    ''nao possue // nao tem
    If Left(UCase(cEMAIL), 3) = "NAO" Then
        CheckEmail = True
        Exit Function
    End If
    
    If Len(cEMAIL) < 5 Then
        'O e-mail é inválido, pois tem menos de 5 caracteres
        CheckEmail = False
        If lMES Then
            MsgBox "O e-mail digitado tem menos de 5 caracterec!!!"
        End If
        Exit Function
    End If

    'Verificar a existencia de arrobas (@) no e-mail
    For nCharacter = 1 To Len(cEMAIL)
        If Mid(cEMAIL, nCharacter, 1) = "@" Then
            'OPA!!! Achou uma arroba!!!
            'Soma 1 ao contador
            Count = Count + 1
        End If
    Next
    'Verifica o número de arrobas.
    'TEM que ter """UMA""" arroba

    If Count <> 1 Then
        'O e-mail é inválido, pois tem 0 ou mais de 1 arroba
        CheckEmail = False
        If lMES Then
            MsgBox "O no. de arrobas (@) do e-mail é inválido!!!"
        End If
        Exit Function
    Else
        'O e-mail tem 1 arroba.
        'Verificar a posição da arroba

        If InStr(cEMAIL, "@") = 1 Then
            'O e-mail é inválido, pois começa com uma @
            CheckEmail = False
            If lMES Then
                MsgBox "O e-mail foi iniciado com uma arroba (@)!!!"
            End If
            Exit Function

        ElseIf InStr(cEMAIL, "@") = Len(cEMAIL) Then
            'O e-mail é inválido, pois termina com uma @
            CheckEmail = False
            If lMES Then
                MsgBox "O e-mail termina com uma arroba (@)!!!"
            End If
            Exit Function
        End If
    End If

    nCharacter = 0
    Count = 0
    'Verificar a existencia de pontos (.) no e-mail

    For nCharacter = 1 To Len(cEMAIL)
        If Mid(cEMAIL, nCharacter, 1) = "." Then
            'OPA!!! Achou um ponto!!!
            'Soma 1 ao contador
            Count = Count + 1
        End If
    Next
    'Verifica o número de pontos.
    'TEM que ter PELO MENOS UM ponto.

    If Count < 1 Then
        'O e-mail é inválido, pois não tem pontos.
        CheckEmail = False
        If lMES Then
            MsgBox "O e-mail é inválido, pois não contém pontos (.)!!!"
        End If
        Exit Function
    Else
        'O e-mail tem pelo menos 1 ponto.
        'Verificar a posição do ponto:

        If InStr(cEMAIL, ".") = 1 Then
            'O e-mail é inválido, pois começa com um ponto
            CheckEmail = False
            If lMES Then
                MsgBox "O e-mail foi iniciado com um ponto (.)!!!"
            End If
            Exit Function

        ElseIf InStr(cEMAIL, ".") = Len(cEMAIL) Then
            'O e-mail é inválido, pois termina com um ponto.
            CheckEmail = False
            If lMES Then
                MsgBox "O e-mail termina com um ponto (.)!!!"
            End If
            Exit Function

        ElseIf InStr(InStr(cEMAIL, "@"), cEMAIL, ".") = 0 Then
            'O e-mail é inválido, pois termina com um ponto.
            CheckEmail = False
            If lMES Then
                MsgBox "O e-mail não tem nenhum ponto (.) após a arroba (@)!!!"
            End If
            Exit Function
        End If
    End If
    nCharacter = 0
    Count = 0
    'Verifica se o e-mail não tem pontos
    'consecutivos (..) após a arroba (@).

    If InStr(cEMAIL, "..") > InStr(cEMAIL, "@") Then
        'O e-mail é inválido, tem pontos consecutivos após o @.
        CheckEmail = False
        If lMES Then
            MsgBox "O e-mail contém pontos consecutivos (..) após o arroba (@)!!!"
        End If
        Exit Function
    End If

    'Verifica se o e-mail tem caracteres inválidos

    For nCharacter = 1 To Len(cEMAIL)
        sLetra = Mid$(cEMAIL, nCharacter, 1)
        ''-./@_~ Or sLetra Like "[0-9]" Or IsNumeric(sLetra)
        If Not (LCase(sLetra) Like "[a-z]" Or sLetra = "@" Or sLetra = "." Or sLetra = "-" Or sLetra = "_" Or sLetra = "/" Or sLetra = "~" Or IsNumeric(sLetra)) Then
            'O e-mail é inválido, pois tem caracteres inválidos
            CheckEmail = False
            If lMES Then
                MsgBox "Foi digitado um caracter inválido no e-mail!!!"
            End If
            Exit Function
        End If
    Next
    nCharacter = 0
    'Bem, se a verificação chegou até aqui é porque o e-mail é válido, então...
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

Public Function FormataIE(ByVal Valor, ByVal cUF, Optional ByVal cPESSOA As String = "J") As String
    Dim cMASK As String
    Valor = FixStr(Valor, "")
    FormataIE = Valor
    cMASK = ""
    If Valor = "ISENTO" Or Valor = "00000000000000" Then
        Exit Function
    End If
    If cPESSOA = "F" Then
        Exit Function
    End If
    If Len(Valor) = 0 Then
        Exit Function
    End If
    If cUF = "EX" Or cUF = "XX" Then
        Exit Function
    End If
    '| no lugar do ponto depois sofre strtarn
    Select Case cUF
    Case "AC"                                    '"99.999.999/999-99" , 13
        cMASK = "00|000|000/000-00"
    Case "AM"                                    '"99.999.999-9"      ,  9
        cMASK = "00|000|000-0"
    Case "AP"                                    ' "99.999.999-9"      ,  9
        cMASK = "00|000|000-0"
    Case "DF"                                    '"99.999999.999-99"  , 13
        cMASK = "00|000000|000-00"
    Case "FN", "EX", "XX"
        ''cMASK = ""
    Case "MA"                                    '"99.999.999-9"      ,  9
        cMASK = "00|000|000-0"
    Case "MG"                                    '"999.999.999/9999"  , 13
        cMASK = "000|000|000/0000"
    Case "MT"                                    '"9999999999-9"      , 11
        cMASK = "0000000000-0"
    Case "PA"                                    '"99.999.999-9"      ,  9
        cMASK = "00|000|000-0"
    Case "PB"                                    '"99.999.999-9"      ,  9
        cMASK = "00|000|000-0"
    Case "PE"                                    '"99.9.999.9999999-9", 14
        cMASK = "00|0|000|0000000-0"
    Case "PI"                                    '"99.999.999-9"      ,  9
        cMASK = "00|000|000-0"
    Case "RN"                                    '"99.999.999-9"      ,  9
        cMASK = "00|000|000-0"
    Case "RO"                                    '"99999999.99999-9"  , 14
        cMASK = "00000000|00000-0"
    Case "RR"                                    '"99.999.999-9"      ,  9
        cMASK = "00|000|000-0"
    Case "SC"                                    ' "999.999.999"       ,  9
        cMASK = "000|000|000"
    Case "SP"                                    '"999.999.999.999"    , 12
        cMASK = "000|000|000|000"
    Case "TO"                                    '"99.999.999-9        , 9
        cMASK = "00|000|000-0"
    Case "AL"                                    '"99.9.99999-9"      ,  9
        cMASK = "00|0|00000-0"
    Case "BA"                                    '"99.999.999"        ,  8
        cMASK = "00|000|000"
    Case "CE"                                    '"99.999.999-9"      ,  9
        cMASK = "00|000|000-0"
    Case "ES"                                    '"99.999.999-9"      ,  9
        cMASK = "000|000|00-0"
    Case "GO"                                    '"99.999.999-9"      ,  9
        cMASK = "00|000|000-0"
    Case "MS"                                    '"99.999.999-9"      ,  9
        cMASK = "00|000|000-0"
    Case "PR"                                    '"999.99999-99"      , 10
        cMASK = "000|00000-00"
    Case "RJ"                                    '"99.999.99-9"       ,  8
        cMASK = "00|000|00-0"
    Case "RS"                                    '"999/999999-9"      , 10
        cMASK = "000/000000-0"
    Case "SE"                                    '"99.999.999-9"      ,  9
        cMASK = "00|000|000-0"
    End Select


    If Len(cMASK) > 0 Then
        Valor = funNumeroPuro(Valor)
        Valor = Format(Valor, cMASK)
        FormataIE = Replace(Valor, "|", ".")
    End If
End Function

Public Function FeriadoMovel(pdatPesquisa As Date) As Boolean
    Dim datData As Date
    Dim ano As Integer
    Dim A As Integer
    Dim b As Integer
    Dim c As Integer
    Dim d As Integer
    Dim E As Integer
    Dim F As Integer
    Dim G As Integer
    Dim h As Integer
    Dim i As Integer
    Dim K As Integer
    Dim L As Integer
    Dim m As Integer
    Dim P As Integer
    Dim Q As Integer
    FeriadoMovel = False
    ano = Year(pdatPesquisa)
    A = (ano Mod 19)
    b = Int(ano / 100)
    c = (ano Mod 100)
    d = Int(b / 4)
    E = (b Mod 4)
    F = Int((b + 8) / 25)
    G = Int((b - F + 1) / 3)
    h = ((19 * A + b - d - G + 15) Mod 30)
    i = Int(c / 4)
    K = (c Mod 4)
    L = ((32 + 2 * E + 2 * i - h - K) Mod 7)
    m = Int((A + 11 * h + 22 * L) / 451)
    P = Int((h + L - 7 * m + 114) / 31)
    Q = ((h + L - 7 * m + 114) Mod 31)
    datData = Format(str(Q + 1) & "/" & str(P) & "/" & str(ano), "dd/mm/yyyy")
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

Public Function CEP2UF(ByVal cCEPUSO As String, Optional ByVal cUF As String = "ZZ") As String
    Dim CEPUSO As Long
    cCEPUSO = TiraOut(cCEPUSO)
    CEPUSO = Val(cCEPUSO)
    CEP2UF = "EX"
    If CEPUSO >= 1000000 And CEPUSO <= 19999999 Then CEP2UF = "SP"
    If CEPUSO >= 20000000 And CEPUSO <= 28999999 Then CEP2UF = "RJ"
    If CEPUSO >= 29000000 And CEPUSO <= 29999999 Then CEP2UF = "ES"
    If CEPUSO >= 30000000 And CEPUSO <= 39999999 Then CEP2UF = "MG"
    If CEPUSO >= 40000000 And CEPUSO <= 48999999 Then CEP2UF = "BA"
    If CEPUSO >= 49000000 And CEPUSO <= 49999999 Then CEP2UF = "SE"
    If CEPUSO >= 50000000 And CEPUSO <= 56999999 Then CEP2UF = "PE"
    If CEPUSO >= 57000000 And CEPUSO <= 57999999 Then CEP2UF = "AL"
    If CEPUSO >= 58000000 And CEPUSO <= 58999999 Then CEP2UF = "PB"
    If CEPUSO >= 59000000 And CEPUSO <= 59999999 Then CEP2UF = "RN"
    If CEPUSO >= 60000000 And CEPUSO <= 63999999 Then CEP2UF = "CE"
    If CEPUSO >= 64000000 And CEPUSO <= 64999999 Then CEP2UF = "PI"
    If CEPUSO >= 65000000 And CEPUSO <= 65999999 Then CEP2UF = "MA"
    If CEPUSO >= 66000000 And CEPUSO <= 68899999 Then CEP2UF = "PA"
    If CEPUSO >= 68900000 And CEPUSO <= 68999999 Then CEP2UF = "AP"
    If CEPUSO >= 69000000 And CEPUSO <= 69299999 Then CEP2UF = "AM" '1a. faixa amazonas
    If CEPUSO >= 69300000 And CEPUSO <= 69399999 Then CEP2UF = "RR"
    If CEPUSO >= 69400000 And CEPUSO <= 69899999 Then CEP2UF = "AM" '2a. faixa amazonas
    If CEPUSO >= 69900000 And CEPUSO <= 69999999 Then CEP2UF = "AC"
    If CEPUSO >= 70000000 And CEPUSO <= 72799999 Then CEP2UF = "DF" '1a. Faixa Distrito Federal
    If CEPUSO >= 72800000 And CEPUSO <= 72999999 Then CEP2UF = "GO" '1a. faixa de Goias
    If CEPUSO >= 73000000 And CEPUSO <= 73699999 Then CEP2UF = "DF" '2a. faixa distrito federal
    If CEPUSO >= 73700000 And CEPUSO <= 76799999 Then CEP2UF = "GO" '2a. faixa de Goias
    If CEPUSO >= 77000000 And CEPUSO <= 77999999 Then CEP2UF = "TO"
    If CEPUSO >= 78000000 And CEPUSO <= 78899999 Then CEP2UF = "MT"
    If CEPUSO >= 78900000 And CEPUSO <= 78999999 Then CEP2UF = "RO"
    If CEPUSO >= 79000000 And CEPUSO <= 79999999 Then CEP2UF = "MS"
    If CEPUSO >= 80000000 And CEPUSO <= 87999999 Then CEP2UF = "PR"
    If CEPUSO >= 88000000 And CEPUSO <= 89999999 Then CEP2UF = "SC"
    If CEPUSO >= 90000000 And CEPUSO <= 99999999 Then CEP2UF = "RS"
    eRETU01 = CEP2UF
    lRETU = False
    If cUF = CEP2UF Then
        lRETU = True
    End If
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

Public Function PEGDDD(ByVal cTEL As String)
    cTEL = formataTelefone(cTEL)
    PEGDDD = ""
    If InStr(cTEL, "(") > 0 Then
        PEGDDD = Mid(cTEL, 2, 2)
    End If
End Function

Public Function PEGTEL(ByVal cTEL As String)
    cTEL = formataTelefone(cTEL)
    PEGTEL = cTEL
    If InStr(cTEL, "(") > 0 Then
        PEGTEL = Mid(cTEL, 5)
    End If
End Function

Public Function PEGPREF(ByVal cTEL As String)
    'cTEL = tem que ser so o numero telefone antes de chamar usar pegtel se necessario
    PEGPREF = ""
    If InStr(cTEL, "-") > 0 Then
        PEGPREF = Mid(cTEL, 1, InStr(cTEL, "-") - 1)
    End If
End Function

Function checkCEI(ByVal pCEI As String) As Boolean
    Dim nTot     As Byte
    Dim cAux     As String
    Dim lret     As Boolean
    Dim i        As Byte
    Dim pNu_cei  As String

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

    cAux = Right$(str(nTot), 2)
    nTot = Val(Left$(cAux, 1)) + Val(Right$(cAux, 1))
    nTot = IIf(nTot > 9, 0, 10 - nTot)
    lret = IIf(Val(Right$(pNu_cei, 1)) = nTot, True, False)

    If lret = False Then
        MsgBox "Codigo CEI incorreto - Tecle <Enter>"
    Else
        checkCEI = True
    End If

End Function

Public Function Validar_CartaoCredito(ByVal CartaoCredito As String) As Boolean
    Dim bolRetorno As Boolean
  
    bolRetorno = False
  
    If Len(CartaoCredito) > 0 Then
        'Retiramos possÝveis mîscaras
        CartaoCredito = funNumeroPuro(CartaoCredito)
    
        If IsNumeric(CartaoCredito) Then
      
            Dim intCont As Integer, intDigito As Integer, intResultado As Integer
      
            intCont = 1
            intDigito = 0
            intResultado = 0
      
            While intCont <= Len(CartaoCredito)
                If Len(CartaoCredito) Mod 2 = 0 Then
        
                    intDigito = Val(Mid(CartaoCredito, intCont, 1))
          
                    If intCont Mod 2 <> 0 Then
                        intDigito = intDigito * 2
                        If intDigito > 9 Then intDigito = intDigito - 9
                    End If
          
                    intResultado = intResultado + intDigito
                    intCont = intCont + 1
          
                Else
        
                    intDigito = Val(Mid(CartaoCredito, intCont, 1))
          
                    If intCont Mod 2 = 0 Then
                        intDigito = intDigito * 2
                        If intDigito > 9 Then intDigito = intDigito - 9
                    End If
          
                    intResultado = intResultado + intDigito
                    intCont = intCont + 1
          
                End If
            Wend
      
            intResultado = intResultado Mod 10
      
            bolRetorno = (intResultado = 0)
      
        End If
    End If
  
    Validar_CartaoCredito = bolRetorno
End Function

Public Function Validar_CodigoBarras(ByVal CodigoBarras As String) As Boolean
    Dim bolRetorno As Boolean
  
    bolRetorno = False
  
    If Len(CodigoBarras) > 0 Then
        'Retiramos possÝveis mîscaras
        CodigoBarras = funNumeroPuro(CodigoBarras)
    
        If IsNumeric(CodigoBarras) Then
      
            Dim strNumeros As String, strDV As String
            Dim intTamanho As Integer
      
            strNumeros = Left(CodigoBarras, Len(CodigoBarras) - 1)
            strDV = Right(CodigoBarras, 1)
      
            intTamanho = Len(strNumeros)
      
            Select Case intTamanho
            Case 7, 11, 12, 13
        
                Dim intCont As Integer, intImpares As Integer, intPares As Integer, intResultado As Integer
          
                For intCont = 1 To intTamanho
                    If (intCont And Not -2) = 0 Then
                        intPares = intPares + Val(Mid(strNumeros, intCont, 1))
                    Else
                        intImpares = intImpares + Val(Mid(strNumeros, intCont, 1))
                    End If
                Next
          
                If (intTamanho = 7) Or (intTamanho = 11) Or (intTamanho = 13) Then 'EAN-8 (8d - 1dv) ou UPC-A (12d - 1dv) ou EAN-14 (14d - 1dv)
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


