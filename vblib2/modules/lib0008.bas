Attribute VB_Name = "LIB0008"
Option Explicit

Public Function pegdizseqssq(ByVal cDIZPAD, ByVal cARQ, _
                             Optional ByVal nTMPSEQ As Integer = 0, _
                             Optional ByVal nTMPSSQ As Integer = 0, _
                             Optional ByVal nTMPORD As Integer = 0, _
                             Optional ByVal nTMPORS As Integer = 0 _
                           )

  Dim cDIZ As String
  Dim cDIZOPER As String
  Dim cDIZFLUXO As String
  Dim cSQL As String


  cDIZ = cDIZPAD
  cDIZOPER = ""
  cDIZFLUXO = ""
  nTMPSEQ = FixInt(Val(Busca("Nº SEQ", "Incluindo", CStr(nTMPSEQ), 8)))
  nTMPSSQ = FixInt(Val(Busca("Nº SSQ", "Incluindo", CStr(nTMPSSQ), 8)))
  nTMPORD = FixInt(Val(Busca("Nº Ordem Fluxo II", "Incluindo", CStr(nTMPORD), 8)))
  nTMPORS = FixInt(Val(Busca("Nº Sub Ordem Fluxo II", "Incluindo", CStr(nTMPORS), 8)))

  If nTMPSEQ > 0 And nTMPSSQ > 0 Then
    cSQL = "select * from PFS WHERE PF=" & nPF & " AND SEQ=" & nTMPSEQ & " AND SSQ=" & nTMPSSQ
    cDIZOPER = PegUltSQL(cARQ, cSQL, "DESCRI", "")
    If Len(cDIZOPER) > 0 Then
      cDIZ = cDIZOPER
    End If
    If nTMPORD > 0 Then
      cSQL = "select * from PFQSBLEP WHERE PF=" & nPF & " AND SEQ=" & nTMPSEQ & " AND SSQ=" & nTMPSSQ & " AND ITEM=" & nTMPORD
      cDIZFLUXO = PegUltSQL(cARQ, cSQL, "DESCRICAO", "")
      If Len(cDIZFLUXO) > 0 Then
        cDIZ = cDIZFLUXO
      End If
    End If
    If Len(cDIZ) = 0 Then
      cDIZ = cDIZPAD
    End If
  End If
  cDIZ = FixStr(Busca("Digite a Descricao", "Incluindo", cDIZ, 255))
  pegdizseqssq = cDIZ
  eRETU01 = Array(nTMPSEQ, nTMPSSQ, nTMPORD, nTMPORS, cDIZFLUXO, cDIZOPER)

End Function

Public Function GERAFEMEAPAD(ByVal nPRO, ByVal nFAL, ByVal nPFVAL, Optional ByVal cDIZ As String = "")
  On Error Resume Next
  Dim nEFE
  Dim nCAU As Integer
  Dim cSQL
  Dim cSQL2 As String
  Dim cARQTMPPF
  Dim cARQTMPFEM
  Dim cCAMPO As String
  Dim DB As New ADODB.Connection
  Dim DB2 As New ADODB.Connection
  Dim DB3 As New ADODB.Connection
  Dim RSTAB As New ADODB.Recordset
  Dim RSTA3 As New ADODB.Recordset
  Dim RSTA2 As New ADODB.Recordset

  cARQTMPFEM = PegPath("PATH", "FEMEA")

  cARQTMPPF = PegPath("PATH", "PF")
  cSQL = "select * from FEMPF WHERE PRONUM=" & nPRO & " AND FALNUM=" & nFAL & " AND PF=" & nPFVAL
  IncluiSQL cARQTMPFEM, cSQL, 3, Array("PRONUM", "FALNUM", "PF"), Array(nPRO, nFAL, nPFVAL), True, False

  'sequencia duplicacao
  DB.ConnectionTimeout = 120
  DB.Open GeracArq(cARQTMPPF, , False)
  RSTAB.Open "select * from DUPLICAR WHERE TABELA='FEMEA'", DB, adOpenDynamic, adLockReadOnly  'DYNAMIC MOVE FIRST
  'arquivo destino
  DB2.ConnectionTimeout = 120
  DB2.Open GeracArq(cARQTMPFEM, , True)
  'origem duplicaçao
  DB3.ConnectionTimeout = 120
  DB3.Open GeracArq(cARQTMPFEM, , False)
  RSTA3.Open "select * from FEMCAU WHERE PRONUM=" & nPRO & " AND FALNUM=" & nFAL, DB3, adOpenForwardOnly, adLockReadOnly
  RSTA3.MoveFirst
  Do While Not RSTA3.EOF
    nEFE = FixInt(RSTA3("EFENUM"))
    nCAU = FixInt(RSTA3("CAUNUM"))
    cSQL2 = "select * from FEMEA WHERE PRONUM=" & nPRO & " AND FALNUM=" & nFAL & " AND EFENUM=" & nEFE & " AND CAUNUM=" & nCAU & " AND PF=" & nPFVAL
    RSTA2.Open cSQL2, DB2, adOpenDynamic, adLockOptimistic
    If RSTA2.EOF Then
      RSTA2.AddNew
      RSTA2("PF") = nPFVAL
      RSTA2("PRONUM") = nPRO
      RSTA2("FALNUM") = nFAL
      RSTA2("EFENUM") = nEFE
      RSTA2("CAUNUM") = nCAU
      RSTAB.MoveFirst
      Do While Not RSTAB.EOF
        cCAMPO = RSTAB("CAMPO")
        RSTA2(cCAMPO) = RSTA3(cCAMPO)
        RSTAB.MoveNext
      Loop
      If Len(cDIZ) > 0 Then
        RSTA2("PROCESSO") = cDIZ
      End If
      RSTA2.Update
    End If
    RSTA3.MoveNext
  Loop
  RSTAB.Close
  RSTA2.Close
  RSTA3.Close
  Set DB = Nothing
  Set RSTAB = Nothing
  Set DB2 = Nothing
  Set RSTA2 = Nothing
  Set DB3 = Nothing
  Set RSTA3 = Nothing

End Function

Function CHECKPFPG()
  Dim sSQL As String
  Dim cARQPF As String
  Dim cARQPFP As String
  Dim cARQPFG As String
  Dim cARQIE As String
  DizerBarra "Checando Femea/Pre/GP12"
  cARQPF = PegPath("PATH", "PF")
  cARQPFP = PegPath("PATH", "PFP")
  cARQPFG = PegPath("PATH", "PFG")
  cARQIE = PegPath("PATH", "IE")
  'IE Instrucao de embalagem
  sSQL = "select pf,data from IE WHERE PF=" & nPF
  IncluiSQL cARQIE, sSQL, 2, Array("PF", "DATA"), Array(nPF, Today()), True, False
  IncluiSQL cARQPFP, sSQL, 2, Array("PF", "DATA"), Array(nPF, Today()), True, False
  IncluiSQL cARQPFG, sSQL, 2, Array("PF", "DATA"), Array(nPF, Today()), True, False
  'Preliminar gp12
  sSQL = "select pf,revdat from PF WHERE PF=" & nPF
  IncluiSQL cARQPFP, sSQL, 2, Array("PF", "REVDAT"), Array(nPF, Today()), True, False
  IncluiSQL cARQPFG, sSQL, 2, Array("PF", "REVDAT"), Array(nPF, Today()), True, False
  DizerBarra ""
End Function

Function ATUPFPG()
  Dim sSQL As String
  Dim cARQPF As String
  Dim cARQPFP As String
  Dim cARQPFG As String
  Dim aVAL As Variant
  Dim aFOR As Variant
  Dim aCAM As Variant
  Dim aPAD As Variant

  DizerBarra "Gravando Informações Complementares Femea/Pre/GP12"
  cARQPF = PegPath("PATH", "PF")
  cARQPFP = PegPath("PATH", "PFP")
  cARQPFG = PegPath("PATH", "PFG")
  'Preliminar gp12

  'sSQL = "select * from PF WHERE PF=" & nPF

  sSQL = "select PF,CPF,OPCAO,CLIENTE,CLINOME,CODIGO,DESCR,STPFPC,STFEPC,FEMEAG,CLIDES,CLIREV,CLIDAT,"
  sSQL = sSQL & "FILIAL,CLIESP,CODCLIENTE,CODMU011,NOMMU011,CODMU012,NOMMU012,CODMU013,"
  sSQL = sSQL & "NOMMU013 , FEMEAEF, FEMEAED, FEMEAEN, FemeaC, FEMEAF, FEMEAD, FemeaR, FEMEAREV, FEMEAREVD, FEMEAREVD2, PCREVD2"
  sSQL = sSQL & " from PF WHERE PF=" & nPF

  aCAM = Array("CPF", "OPCAO", "CLIENTE", "CLINOME", "CODIGO", _
               "DESCR", "STPFPC", "STFEPC", "FEMEAG", "CLIDES", _
               "CLIREV", "CLIDAT", "FILIAL", "CLIESP", "CODCLIENTE", _
               "CODMU011", "NOMMU011", "CODMU012", "NOMMU012", "CODMU013", _
               "NOMMU013", "FEMEAEF", "FEMEAED", "FEMEAEN", "FEMEAC", _
               "FEMEAF", "FEMEAD", "FEMEAR", "FEMEAREV", "FEMEAREVD", "FEMEAREVD2", "PCREVD2")

  aFOR = Array("C", "NI", "NI", "C", "C", _
               "C", "B", "B", "C", "C", _
               "C", "DN", "C", "C", "C", _
               "C", "C", "C", "C", "C", _
               "C", "NI", "DN", "C", "C", _
               "NI", "DN", "C", "NI", "DN", "DN", "DN")

  aPAD = Array("", 0, 0, "", "", _
               "", False, False, "", "", _
               "", "", "", "", "", _
               "", "", "", "", "", _
               "", 0, "", "", "", _
               0, "", "", 0, "", "", "")

  aVAL = PegSQL(cARQPF, sSQL, 32, aCAM, aFOR, aPAD)

  GrvSQL cARQPFG, sSQL, 32, aCAM, aVAL, aFOR
  GrvSQL cARQPFP, sSQL, 32, aCAM, aVAL, aFOR

  DizerBarra ""

End Function

Public Function importa2(ByVal cDUPARQ, ByVal cDUPSQL, ByVal cORIARQ, ByVal cORISQL, ByVal cDESARQ, ByVal cDESSQL, ByVal cTIPO)
  Dim DB As ADODB.Connection
  Dim DB2 As ADODB.Connection
  Dim DB3 As ADODB.Connection
  Dim RSTAB As ADODB.Recordset
  Dim RSTA2 As ADODB.Recordset
  Dim RSTA3 As ADODB.Recordset
  Dim cCAMPO As String
  Dim cARQOPEN As String


  On Error GoTo errhandler

  Set DB = New ADODB.Connection
  Set RSTAB = New ADODB.Recordset
  Set DB2 = New ADODB.Connection
  Set RSTA2 = New ADODB.Recordset
  Set DB3 = New ADODB.Connection
  Set RSTA3 = New ADODB.Recordset


  'sequencia duplicacao
  DB.ConnectionTimeout = 120
  DB.Open GeracArq(cDUPARQ)                      ', , False) O pf tambem pode ser o arquivo origem destino
  RSTAB.Open cDUPSQL, DB, adOpenDynamic, adLockReadOnly  'MOVE FIRST


  'arquivo destino
  cARQOPEN = GeracArq(cDESARQ)
  DB2.ConnectionTimeout = 120
  DB2.Open cARQOPEN
  RSTA2.Open cDESSQL, DB2, adOpenDynamic, adLockOptimistic

  'arquivo origem
  If cORIARQ = cDESARQ Then
    RSTA3.Open cORISQL, DB2, adOpenDynamic, adLockOptimistic
  Else
    cARQOPEN = GeracArq(cORIARQ)
    DB3.ConnectionTimeout = 120
    DB3.Open cARQOPEN
    RSTA3.Open cORISQL, DB3, adOpenForwardOnly, adLockReadOnly
  End If

  If RSTA3.EOF Then
    Alert ("Sua Selecão nao tem informações para importar")
    RSTAB.Close
    RSTA2.Close
    RSTA3.Close
    DB.Close
    DB2.Close
    If cORIARQ <> cDESARQ Then
      DB3.Close
    End If
    Exit Function
  End If

  While Not RSTA3.EOF

    If cTIPO <> "PFX" Then
      RSTA2.AddNew
    Else
      If Not RSTA2.EOF Then
        RSTA2.MoveFirst
      Else
        RSTA2.AddNew
      End If
    End If

    RSTAB.MoveFirst
    Do While Not RSTAB.EOF
      cCAMPO = RSTAB("CAMPO")
      DizerBarra cCAMPO
      RSTA2(cCAMPO) = RSTA3(cCAMPO)
      RSTAB.MoveNext
    Loop

    If (cTIPO = "PFS" Or cTIPO = "PFI" Or cTIPO = "PFQSBLEP" Or cDESSQL = "PFS" Or cDESSQL = "PFI" Or cDESSQL = "PFQSBLEP") Then
      If IsNull(nSEQ) Or IsEmpty(nSEQ) Then
        nSEQ = RSTA3("SEQ")
      End If
      If IsNull(nSSQ) Or IsEmpty(nSSQ) Then
        nSSQ = RSTA3("SSQ")
      End If
    End If

    If cTIPO = "FEMAX" Then
      RSTA2("PRONUM") = RSTA3("PRONUM")
      RSTA2("FALNUM") = RSTA3("FALNUM")
      RSTA2("EFENUM") = RSTA3("EFENUM")
      RSTA2("CAUNUM") = RSTA3("CAUNUM")
      RSTA2("MUDPAD") = RSTA3("MUDPAD")
      RSTA2("ALTMAN") = RSTA3("ALTMAN")
    End If


    If cTIPO <> "PFX" Then
      RSTA2("PF") = nPF
    End If
    If cTIPO = "PFC" Or cDESSQL = "PFC" Then
      RSTA2("SEQ") = nSEQ
      RSTA2("SSQ") = nSSQ
    End If
    If (cTIPO = "PFS" Or cTIPO = "PFI" Or cDESSQL = "PFS" Or cDESSQL = "PFI") Then
      If IsNull(RSTA2("SEQ")) Or IsEmpty(RSTA2("SEQ")) Or RSTA2("SEQ") = 0 Then
        RSTA2("SEQ") = RSTA3("SEQ")
      End If
      If IsNull(RSTA2("SSQ")) Or IsEmpty(RSTA2("SSQ")) Or RSTA2("SSQ") = 0 Then
        RSTA2("SSQ") = RSTA3("SSQ")
      End If
    End If


    If (cTIPO = "PFCO" Or cDESSQL = "PFCO") And cDESSQL <> "PFC" Then
      If Len(Ctipoent) > 0 Then
        RSTA2("TIPOENT") = Ctipoent
      End If
      If Len(Ccodcomp) > 0 Then
        RSTA2("CODCOMP") = Ccodcomp
      End If
    End If


    RSTA2.Update
    RSTA3.MoveNext

  Wend

  RSTAB.Close
  RSTA2.Close
  RSTA3.Close

  DB.Close
  DB2.Close
  If cORIARQ <> cDESARQ Then
    DB3.Close
  End If
  Set DB = Nothing
  Set RSTAB = Nothing
  Set DB2 = Nothing
  Set RSTA2 = Nothing
  Set DB3 = Nothing
  Set RSTA3 = Nothing
  Exit Function
errhandler:
  Select Case Err.Number
  Case Else
    SayErro "Importar 2"
    Exit Function
  End Select
End Function


