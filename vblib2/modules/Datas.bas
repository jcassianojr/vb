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
    If Year(eVAR) <= 1899 Then
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
  If Len(cData) = 8 Then                       ''8 ''yyyymmdd
    dDATA = DateSerial(Mid(cData, 1, 4), Mid(cData, 5, 2), Mid(cData, 7, 2))
  End If
  NumToData = Fdata(dDATA)
End Function

' +--------------------------------------------------------------------
' +  Função: UniversalToDate
' +  Objetivo: Garantir a leitura correta de datas em múltiplos formatos
' +            (DD/MM/YY, DD-MM-YYYY, DD.MM.YY, AAAAMMDD, YYYY-MM-DD)
' +  Retorno:  Date (Retorna a data correspondente ou 0 se for inválida/vazia)
' +--------------------------------------------------------------------
Public Function UniversalToDate(ByVal xData As Variant) As Date
    Dim cData As String
    Dim cLimpa As String
    Dim nLEN As Integer
    
    ' 1. Se já for do tipo Date nativo do VB6, retorna ela mesma
    If VarType(xData) = vbDate Then
        UniversalToDate = CDate(xData)
        Exit Function
    End If
    
    ' 2. Se for nulo ou não for conversível para string, retorna data vazia (0)
    If IsNull(xData) Or IsEmpty(xData) Then
        UniversalToDate = CDate(0)
        Exit Function
    End If
    
    cData = UCase$(Trim$(CStr(xData)))
    
    If cData = "" Or cData = "01/01/1900" Or cData = "0000-00-00" Or Left$(cData, 4) = "0000" Then
        UniversalToDate = CDate(0)
        Exit Function
    End If
    
    ' 3. Trata o formato ISO comum em bancos de dados (YYYY-MM-DD)
    '    Exemplo: "2026-05-16" ou "2026-05-16 00:00:00"
    If Len(cData) >= 10 Then
        If Mid$(cData, 5, 1) = "-" And Mid$(cData, 8, 1) = "-" Then
            UniversalToDate = DateSerial(Val(Left$(cData, 4)), Val(Mid$(cData, 6, 2)), Val(Mid$(cData, 9, 2)))
            Exit Function
        End If
    End If
    
    ' 4. Remove separadores comuns para analisar a sequência numérica pura
    cLimpa = Replace(cData, "/", "")
    cLimpa = Replace(cLimpa, "-", "")
    cLimpa = Replace(cLimpa, ".", "")
    cLimpa = Trim$(cLimpa)
    
    ' Se houver hora grudada na string limpa, pega apenas os primeiros 6 ou 8 dígitos da data
    If Len(cLimpa) > 8 Then
        cLimpa = Left$(cLimpa, 8)
    End If
    
    nLEN = Len(cLimpa)
    
    ' 5. Identifica o formato pelo tamanho da string resultante
    Select Case nLEN
        Case 8
            ' Pode ser AAAAMMDD (banco/perfil) ou DDMMYYYY (padrão local)
            ' Se os 4 primeiros dígitos formarem um ano válido (ex: > 1900)
            If Val(Left$(cLimpa, 4)) > 1900 Then
                UniversalToDate = DateSerial(Val(Left$(cLimpa, 4)), Val(Mid$(cLimpa, 5, 2)), Val(Right$(cLimpa, 2)))
            Else
                ' Caso contrário, assume DDMMYYYY
                UniversalToDate = DateSerial(Val(Right$(cLimpa, 4)), Val(Mid$(cLimpa, 3, 2)), Val(Left$(cLimpa, 2)))
            End If
            
        Case 6
            ' Formato DDMMYY (Ano com 2 dígitos)
            ' O VB6 resolve o centenário automaticamente com base nas configurações do Windows
            UniversalToDate = DateSerial(Val(Right$(cLimpa, 2)), Val(Mid$(cLimpa, 3, 2)), Val(Left$(cLimpa, 2)))
            
        Case Else
            ' Fallback: Se o VB6 conseguir converter nativamente, usa. Se não, retorna data vazia.
            If IsDate(cData) Then
                UniversalToDate = CDate(cData)
            Else
                UniversalToDate = CDate(0)
            End If
    End Select
End Function

' +--------------------------------------------------------------------
' +  Função: C_Data
' +  Objetivo: Nova função para converter e validar strings de data com segurança.
' +  Retorno:  Date
' +--------------------------------------------------------------------
Public Function C_Data(ByVal mDATA As String) As Date
    C_Data = UniversalToDate(mDATA)
End Function

' +--------------------------------------------------------------------
' +  Função: DigaData
' +  Objetivo: Retorna a data formatada por extenso (DD de Mês de AAAA)
' +--------------------------------------------------------------------
Public Function DigaData(ByVal mDATA As Variant) As String
    Dim dReal As Date
    
    dReal = UniversalToDate(mDATA)
    
    If dReal = CDate(0) Then
        DigaData = ""
        Exit Function
    End If
    
    ' Usamos Trim$ nativo do VB6 combinado com Str$ para evitar erros de tipo
    DigaData = Trim$(Str$(Day(dReal))) & " de " & cMes(dReal) & " de " & Trim$(Str$(Year(dReal)))
End Function

' +--------------------------------------------------------------------
' +  Função: CMes
' +  Objetivo: Retorna o nome do mês por extenso a partir de uma data ou string.
' +--------------------------------------------------------------------
Public Function cMes(ByVal mDATA As Variant) As String
    Dim dReal As Date
    
    ' Intercepta e garante que seja uma data real antes de usar o Month()
    dReal = UniversalToDate(mDATA)
    
    If dReal = CDate(0) Then
        cMes = ""
        Exit Function
    End If
    
    ' CORRIGIDO: Chama MMes (com dois Ms) que existe no seu arquivo Datas.bas
    cMes = mMES(Month(dReal))
End Function

' +--------------------------------------------------------------------
' +  Função: CDia
' +  Objetivo: Retorna o nome do dia da semana por extenso a partir de uma data ou string.
' +--------------------------------------------------------------------
Public Function cDia(ByVal mDATA As Variant) As String
    Dim dReal As Date
    
    ' Intercepta e garante que seja uma data real antes de usar o Weekday()
    dReal = UniversalToDate(mDATA)
    
    If dReal = CDate(0) Then
        cDia = ""
        Exit Function
    End If
    
    ' CORRIGIDO: Chama DDia (com dois Ds) que existe no seu arquivo Datas.bas
    cDia = DDia(Weekday(dReal))
End Function

' +--------------------------------------------------------------------
' +  Função: MMes
' +  Objetivo: Recebe o número do mês (1 a 12) e retorna o nome por extenso.
' +  Retorno:  String
' +--------------------------------------------------------------------
Public Function mMES(ByVal nMES As Integer) As String
    Dim cNomMes As String
    
    cNomMes = ""
    
    Select Case nMES
        Case 1:  cNomMes = "Janeiro"
        Case 2:  cNomMes = "Fevereiro"
        Case 3:  cNomMes = "Março"
        Case 4:  cNomMes = "Abril"
        Case 5:  cNomMes = "Maio"
        Case 6:  cNomMes = "Junho"
        Case 7:  cNomMes = "Julho"
        Case 8:  cNomMes = "Agosto"
        Case 9:  cNomMes = "Setembro"
        Case 10: cNomMes = "Outubro"
        Case 11: cNomMes = "Novembro"
        Case 12: cNomMes = "Dezembro"
    End Select
    
    mMES = cNomMes
End Function

' +--------------------------------------------------------------------
' +  Função: DDia
' +  Objetivo: Recebe o número do dia da semana (1 a 7) e retorna o nome por extenso.
' +            (1 = Domingo, 2 = Segunda-Feira, ..., 7 = Sábado)
' +  Retorno:  String
' +--------------------------------------------------------------------
Public Function DDia(ByVal nDIA As Integer) As String
    Dim cNomDia As String
    
    cNomDia = ""
    
    Select Case nDIA
        Case 1: cNomDia = "Domingo"
        Case 2: cNomDia = "Segunda-Feira"
        Case 3: cNomDia = "Terça-Feira"
        Case 4: cNomDia = "Quarta-Feira"
        Case 5: cNomDia = "Quinta-Feira"
        Case 6: cNomDia = "Sexta-Feira"
        Case 7: cNomDia = "Sábado"
    End Select
    
    DDia = cNomDia
End Function
' +--------------------------------------------------------------------
' +  Função: DataS
' +  Objetivo: Converte uma data (ou string de data) para o formato padrão
' +            texto "DD/MM/AAAA", garantindo tipagem segura.
' +  Retorno:  String
' +--------------------------------------------------------------------
Public Function DataS(ByVal dDATA As Variant) As String
    Dim dReal As Date
    
    ' Garante a conversão para uma data legítima do VB6
    dReal = UniversalToDate(dDATA)
    
    ' Se for uma data vazia/inválida, retorna string vazia
    If dReal = CDate(0) Then
        DataS = ""
        Exit Function
    End If
    
    ' Utiliza a função Format nativa do VB6 de forma segura
    DataS = Format(dReal, "dd/mm/yyyy")
End Function

' +--------------------------------------------------------------------
' +  Função: DataS2
' +  Objetivo: Retorna a data atual do sistema (Date) formatada em string "DD/MM/AAAA".
' +  Retorno:  String
' +--------------------------------------------------------------------
Public Function DataS2() As String
    ' Como usa a data do próprio sistema, basta aplicar o Format diretamente
    DataS2 = Format(Date, "dd/mm/yyyy")
End Function

' +--------------------------------------------------------------------
' +  Função: DataToLit
' +  Objetivo: Converte uma data para uma string literal SQL formatada
' +            mantendo as tags completas padrão do sistema.
' +  Retorno:  String
' +--------------------------------------------------------------------
Public Function DataToLit(ByVal dDATA As Variant, Optional ByVal cTIPO As String = "", Optional ByVal cARQ As String = "") As String
    Dim dReal As Date
    Dim aRETU As Variant
    Dim cAno As String
    Dim cMes As String
    Dim cDia As String
    
    ' Garante a conversão do parâmetro para uma data real do VB6 antes de gerar o SQL
    dReal = UniversalToDate(dDATA)
    
    ' Se a data for vazia ou inválida, retorna uma string vazia ou formato nulo padrão
    If dReal = CDate(0) Then
        DataToLit = "NULL"
        Exit Function
    End If
    
    ' Se foi passado um arquivo de configuração, resolve a conexão e pega o tipo
    If Len(cARQ) > 0 Then
        cARQ = GeraConn(cARQ, cTIPO)
        aRETU = TipoConn(cARQ)
        cTIPO = aRETU(2)
    End If
    
    ' Padroniza a tag em maiúsculo e limpa espaços (SEM reduzir o tamanho)
    cTIPO = UCase$(Trim$(cTIPO))
    
    ' Prepara as strings de Ano, Mês e Dia com zeros à esquerda usando lógica nativa VB6
    cAno = Trim$(Str$(Year(dReal)))
    cMes = Right$("0" & Trim$(Str$(Month(dReal))), 2)
    cDia = Right$("0" & Trim$(Str$(Day(dReal))), 2)
    
    ' Monta a sintaxe correta baseada nas tags longas e oficiais do seu sistema
    Select Case cTIPO
        Case "ACCESS", "ACESS"
            DataToLit = "#" & cAno & "/" & Month(dReal) & "/" & Day(dReal) & "#"
            
        Case "CRYSTAL"
            DataToLit = "CDATE(" & cAno & "," & Month(dReal) & "," & Day(dReal) & ")"
            
        Case "MYSQL/"
            DataToLit = "'" & cAno & "/" & cMes & "/" & cDia & "'"
            
        Case "MYSQL", "MYSQL-", "MARIADB"
            DataToLit = "'" & cAno & "-" & cMes & "-" & cDia & "'"
            
        Case "SQLSERVER"
            DataToLit = "CONVERT(datetime, '" & cAno & "-" & cMes & "-" & cDia & "', 102)"
            
        Case "ORACLE"
            DataToLit = "to_date('" & cDia & "/" & cMes & "/" & cAno & "','DD/MM/YYYY')"
            
        Case "POSTGRESQL", "POSTGRES"
            DataToLit = "DATE '" & cAno & "-" & cMes & "-" & cDia & "'"
            
        Case "SQLITE"
            DataToLit = "'" & cAno & "-" & cMes & "-" & cDia & "'"
            
        Case "CRYSTAX"
            DataToLit = "DATE(" & cAno & "," & Month(dReal) & "," & Day(dReal) & ")"
            
        Case Else
            ' Fallback ANSI SQL seguro (YYYY-MM-DD) para ADVANTAGE ou qualquer outro banco
            DataToLit = "'" & cAno & "-" & cMes & "-" & cDia & "'"
    End Select
End Function
' +--------------------------------------------------------------------
' +  Função: QuarterStart
' +  Objetivo: Retorna o primeiro dia do trimestre com base na data fornecida.
' +--------------------------------------------------------------------
Function QuarterStart(ByVal AsOfDate As Variant, Optional RelativeQuarter As Integer = 0) As Date
    Dim dReal As Date
    Dim NewMonth As Integer
    
    dReal = UniversalToDate(AsOfDate)
    If dReal = CDate(0) Then dReal = Date
    
    ' Calcula o mês inicial do trimestre correspondente
    NewMonth = (((Month(dReal) - 1) \ 3) * 3) + 1 + (RelativeQuarter * 3)
    QuarterStart = DateSerial(Year(dReal), NewMonth, 1)
End Function

' +--------------------------------------------------------------------
' +  Função: QuarterEnd
' +  Objetivo: Retorna o último dia do trimestre com base na data fornecida.
' +--------------------------------------------------------------------
Function QuarterEnd(ByVal AsOfDate As Variant, Optional RelativeQuarter As Integer = 0) As Date
    Dim dReal As Date
    Dim NewMonth As Integer
    
    dReal = UniversalToDate(AsOfDate)
    If dReal = CDate(0) Then dReal = Date
    
    ' Localiza o próximo trimestre e define o dia "0" para pegar o último dia do mês anterior
    NewMonth = (((Month(dReal) - 1) \ 3) * 3) + 4 + (RelativeQuarter * 3)
    QuarterEnd = DateSerial(Year(dReal), NewMonth, 0)
End Function

' +--------------------------------------------------------------------
' +  Função: SomaMes
' +  Objetivo: Adiciona ou subtrai meses de uma data de forma segura.
' +--------------------------------------------------------------------
Public Function SomaMes(ByVal mDATA As Variant, ByVal mMES As Integer) As Date
    Dim dReal As Date
    
    dReal = UniversalToDate(mDATA)
    If dReal = CDate(0) Then dReal = Date
    
    ' Utiliza DateAdd nativo do VB6 para garantir o cálculo correto de virada de anos
    SomaMes = DateAdd("m", mMES, dReal)
End Function

' +--------------------------------------------------------------------
' +  Função: SomaAno
' +  Objetivo: Adiciona ou subtrai anos de uma data de forma segura.
' +--------------------------------------------------------------------
Public Function SomaAno(ByVal mDATA As Variant, ByVal mANO As Integer) As Date
    Dim dReal As Date
    
    dReal = UniversalToDate(mDATA)
    If dReal = CDate(0) Then dReal = Date
    
    ' Utiliza DateAdd nativo do VB6 para somar anos mantendo a consistência do calendário
    SomaAno = DateAdd("yyyy", mANO, dReal)
End Function

' +--------------------------------------------------------------------
' +  Função: Dia_Mes
' +  Objetivo: Retorna a quantidade de dias que um determinado mês possui.
' +            Trata automaticamente anos bissextos.
' +  Retorno:  Integer (Ex: 28, 29, 30 ou 31)
' +--------------------------------------------------------------------
Public Function Dia_Mes(ByVal nMES As Integer, ByVal nANO As Integer) As Integer
    ' No VB6, passar o dia 0 do mês seguinte para o DateSerial
    ' retorna exatamente o último dia do mês desejado.
    If nMES >= 1 And nMES <= 12 And nANO > 0 Then
        Dia_Mes = Day(DateSerial(nANO, nMES + 1, 0))
    Else
        Dia_Mes = 0
    End If
End Function

' +--------------------------------------------------------------------
' +  Função: Bissexto
' +  Objetivo: Verifica se um determinado ano é bissexto.
' +  Retorno:  Boolean (True/False)
' +--------------------------------------------------------------------
Public Function Bissexto(ByVal nANO As Integer) As Boolean
    ' Se o dia 29 de fevereiro existir para o ano informado, ele é bissexto
    If nANO > 0 Then
        Bissexto = (Day(DateSerial(nANO, 3, 0)) = 29)
    Else
        Bissexto = False
    End If
End Function

' +--------------------------------------------------------------------
' +  Função: ValiData
' +  Objetivo: Valida de forma rigorosa se uma string ou variante
' +            corresponde a uma data real e lógica do calendário.
' +  Retorno:  Boolean (True se for válida, False se for inválida)
' +--------------------------------------------------------------------
Public Function ValiData(ByVal mDATA As Variant) As Boolean
    Dim dReal As Date
    
    ' Se for nulo ou vazio, já assume inválida
    If IsNull(mDATA) Or IsEmpty(mDATA) Then
        ValiData = False
        Exit Function
    End If
    
    ' Passa pelo motor de conversão universal
    dReal = UniversalToDate(mDATA)
    
    ' Se retornar a data zerada (0), significa que o formato falhou ou é inválido
    If dReal = CDate(0) Then
        ValiData = False
    Else
        ValiData = True
    End If
End Function

' +--------------------------------------------------------------------
' +  Função: Inverte
' +  Objetivo: Converte uma data para uma string invertida sem barras,
' +            usando hífen como separador (útil para nomes de ficheiros ou chaves).
' +  Retorno:  String (Formato: AAAA-MM-DD)
' +--------------------------------------------------------------------
Public Function Inverte(ByVal mDATA As Variant) As String
    Dim dReal As Date
    Dim cAno As String
    Dim cMes As String
    Dim cDia As String
    
    ' Garante que a entrada seja convertida numa data real do VB6
    dReal = UniversalToDate(mDATA)
    
    ' Se for inválida ou zerada, retorna uma string vazia
    If dReal = CDate(0) Then
        Inverte = ""
        Exit Function
    End If
    
    ' Extrai e padroniza as partes com zeros à esquerda
    cAno = Trim$(Str$(Year(dReal)))
    cMes = Right$("0" & Trim$(Str$(Month(dReal))), 2)
    cDia = Right$("0" & Trim$(Str$(Day(dReal))), 2)
    
    ' Monta a string invertida de forma segura
    Inverte = cAno & "-" & cMes & "-" & cDia
End Function

' +--------------------------------------------------------------------
' +  Função: DataAno
' +  Objetivo: Retorna os dois últimos dígitos do ano da data informada,
' +            concatenados com um sufixo numérico (ex: usado em lógicas de lotes/exercícios).
' +  Retorno:  String
' +--------------------------------------------------------------------
Public Function DataAno(ByVal mDATA As Variant, ByVal mNUME As Integer) As String
    Dim dReal As Date
    Dim cAnoCurto As String
    
    ' Garante a conversão da variante para data legítima
    dReal = UniversalToDate(mDATA)
    
    ' Se a data for inválida, assume a data atual do sistema para não quebrar a lógica
    If dReal = CDate(0) Then
        dReal = Date
    End If
    
    ' Extrai os dois últimos dígitos do ano (ex: 2026 vira "26")
    cAnoCurto = Right$("0" & Trim$(Str$(Year(dReal))), 2)
    
    ' Retorna o ano curto concatenado ao número fornecido formatado com dois dígitos
    DataAno = cAnoCurto & Right$("0" & Trim$(Str$(mNUME)), 2)
End Function
' +--------------------------------------------------------------------
' +  Função: ValiDataE
' +  Objetivo: Valida uma data utilizando a UniversalToDate. Se for inválida,
' +            pode opcionalmente exibir uma mensagem de alerta.
' +  Retorno:  Boolean (True se for válida)
' +--------------------------------------------------------------------
Public Function ValiDataE(ByVal mDATA As Variant, Optional ByVal lExibeMensagem As Boolean = False) As Boolean
    Dim dReal As Date
    
    dReal = UniversalToDate(mDATA)
    
    If dReal = CDate(0) Then
        If lExibeMensagem Then
            MsgBox "Data inválida ou incorreta!", vbCritical, "Aviso do Sistema"
        End If
        ValiDataE = False
    Else
        ValiDataE = True
    End If
End Function

' +--------------------------------------------------------------------
' +  Função: D_Extenso
' +  Objetivo: Retorna a data completa formatada por extenso com o dia da semana.
' +            Exemplo: "Sábado, 16 de Maio de 2026"
' +  Retorno:  String
' +--------------------------------------------------------------------
Public Function D_Extenso(ByVal mDATA As Variant) As String
    Dim dReal As Date
    
    ' Garante a conversão da variante para data legítima do VB6
    dReal = UniversalToDate(mDATA)
    
    ' Se a data for inválida ou zerada, retorna string vazia
    If dReal = CDate(0) Then
        D_Extenso = ""
        Exit Function
    End If
    
    ' Monta o extenso completo utilizando as nossas funções seguras CDia e CMes
    ' Usamos Trim$(Str$()) para evitar incompatibilidade de tipos
    D_Extenso = cDia(dReal) & ", " & Trim$(Str$(Day(dReal))) & " de " & cMes(dReal) & " de " & Trim$(Str$(Year(dReal)))
End Function
