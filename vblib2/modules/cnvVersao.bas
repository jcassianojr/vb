Attribute VB_Name = "CnvVersao"
Option Explicit

Public Function rptimpsys()
End Function

Public Function AcessaMenu(ByVal cMENU As String, ByVal nINDICE As Integer) As Boolean
  Dim sSQL As String
  AcessaMenu = False
  If zUSER = "SOFTEC" Or zUSER = "ADMIN" Then
    AcessaMenu = True
    Exit Function
  End If
  sSQL = "select LIGADO from MENUUSU WHERE MENU='" & cMENU & "' AND INDICE=" & nINDICE & " AND IDUSUARIO=" & zUSERID
  If PegUltSQL(Dbname, sSQL, "LIGADO", False) Then
    AcessaMenu = True
    Exit Function
  End If
  If Not AcessaMenu Then
    Alert "Acesso Bloqueado Menu: " & cMENU & " Indice: " & CStr(nINDICE)
  End If
End Function

Public Function AcessaBtnOld(ByVal cFORM As String, ByVal nINDICE As Integer) As Boolean
  Dim sSQL As String
  AcessaBtnOld = False
  If zUSER = "SOFTEC" Or zUSER = "ADMIN" Then
    AcessaBtnOld = True
    Exit Function
  End If

  sSQL = "SELECT LIGADO FROM USUCAD WHERE FORM = '" & cFORM & "' AND INDICE=" & nINDICE & " AND IDUSUARIO=" & zUSERID
  If PegUltSQL(Dbname, sSQL, "LIGADO", False) Then
    AcessaBtnOld = True
    Exit Function
  End If

  sSQL = "SELECT DISPONIVEL FROM CONTROLE WHERE FORM = '" & cFORM & "' and controle='Toolbar1' AND INDICE=" & nINDICE
  If PegUltSQL(Dbname, sSQL, "DISPONIVEL", False) Then

    AcessaBtnOld = True
    Exit Function
  End If

  If Not AcessaBtnOld Then
    Alert "Acesso Bloqueado Botao: " & cFORM & " Indice: " & CStr(nINDICE)
  End If
End Function

Public Function AcessaBotao(ByVal nFORM As Long, ByVal nBOTAO As Integer) As Boolean
  AcessaBotao = True
End Function

Public Function AcessaForm(ByVal nFORM As Long) As Boolean
  AcessaForm = True
End Function

Public Function GravaLog(Optional ByVal nFORM As Long = 0, Optional ByVal nBOTAO As Long = 0, Optional ByVal Copr As String = "", Optional ByVal cOBS As String = "")
  GravaLog = IncluiSQL(cARQERRO, "LOG", 6, Array("USER", "FORM", "BOTAO", "OPR", "DATA", "OBS"), Array(zUSERID, nFORM, nBOTAO, Copr, Now, cOBS))
End Function

Public Function demitido(ByVal nNUMERO As Long, Optional ByVal lMES As Boolean = True)
  Dim cARQ As String
  Dim sSQL As String
  Dim aRETU As Variant
  Dim nLOOP As Integer

  demitido = False


  For nLOOP = 1 To 6
    Select Case nLOOP
    Case 1
      cARQ = PegPath("PATH", "LOGIXODBC")
      sSQL = "SELECT dat_demis as demitido FROM funcionario WHERE cod_empresa='01'  and cod_vinculo<>99 and num_matricula=" & nNUMERO  ''tenta funcionarios
    Case 2
      cARQ = PegPath("PATH", "LOGIXODBC")
      sSQL = "SELECT dat_demis as demitido FROM funcionario WHERE cod_empresa='05'  and cod_vinculo<>99 and num_matricula=" & nNUMERO  ''tenta funcionarios
    Case 3
      cARQ = PegPath("PATH", "LOGIXODBC")
      sSQL = "SELECT dat_demis as demitido FROM funcionario WHERE cod_empresa='11'  and cod_vinculo<>99 and num_matricula=" & nNUMERO  ''tenta funcionarios
    Case 4
      cARQ = PegPath("PATH", "LOGIXODBC")
      sSQL = "SELECT dat_demis as demitido FROM funcionario WHERE cod_vinculo=50 and num_matricula=" & nNUMERO  ''tenta temporarios
    Case 5
      cARQ = PegPath("PATH", "CADMP04")
      cARQ = GeraConn(cARQ, "JETFOX")
      sSQL = "SELECT DEMITIDO FROM MP04 WHERE TECNICO=" & nNUMERO  ''tenta pela folha
    Case 6
      cARQ = PegPath("PATH", "CADMP0499")
      cARQ = GeraConn(cARQ, "JETFOX")
      sSQL = "SELECT DEMITIDO FROM MP04 WHERE TECNICO=" & nNUMERO  ''tenta pela folha

    End Select

    If nLOOP < 5 And zusalx = "NAO" Then
      lRETU = False
    Else
      aRETU = PegSQL(cARQ, sSQL, 1, Array("demitido"), Array("DN"), Array(Today()), False)
    End If
    If lRETU Then
      If Day(aRETU(0)) > 0 And aRETU(0) <> Today() Then
        Alert ("Funcionario: " & CStr(nNUMERO) & " Troca de Empresa/Cracha ou Demitido em " & aRETU(0))
        demitido = True
        eRETU01 = aRETU(0)
      Else
        demitido = False
        Exit Function
      End If
    End If


  Next nLOOP

  '     cARQ = PegPath("PATH", "LOGIXODBC")
  '       sSQL = "SELECT dat_demis FROM funcionario WHERE (cod_empresa='01' or cod_empresa='05' or cod_empresa='11' ) and cod_vinculo<>99 and num_matricula=" & nNUMERO ''tenta funcionarios
  '       aRETU = PegSQL(cARQ, sSQL, 1, Array("dat_demis"), Array("DN"), Array(Today()), False)
  '       If lRETU Then
  '          If Day(aRETU(0)) > 0 And aRETU(0) <> Today() Then
  '             Alert ("Funcionario: " & CStr(nNUMERO) & " Demitido em " & aRETU(0))
  '             demitido = True
  '             eRETU01 = aRETU(0)
  '          End If
  '       Else
  '          cARQ = PegPath("PATH", "LOGIXODBC")
  '          sSQL = "SELECT dat_demis FROM funcionario WHERE cod_vinculo=50 and num_matricula=" & nNUMERO  ''tenta temporarios
  '          aRETU = PegSQL(cARQ, sSQL, 1, Array("dat_demis"), Array("DN"), Array(Today()), False)
  '          If lRETU Then
  '             If Day(aRETU(0)) > 0 And aRETU(0) <> Today() Then
  '                Alert ("Funcionario: " & CStr(nNUMERO) & " Demitido em " & aRETU(0))
  '                demitido = True
  '                eRETU01 = aRETU(0)
  '             End If
  '          Else
  '                cARQ = PegPath("PATH", "CADMP04")
  '                cARQ = GeraConn(cARQ, "JETFOX")
  '                sSQL = "SELECT DEMITIDO FROM MP04 WHERE TECNICO=" & nNUMERO ''tenta pela folha
  '                aRETU = PegSQL(cARQ, sSQL, 1, Array("DEMITIDO"), Array("DN"), Array(Today()), False)
  '                If lRETU Then
  '                   If Day(aRETU(0)) > 0 And aRETU(0) <> Today() Then
  '                      Alert ("Funcionario: " & CStr(nNUMERO) & " Demitido em " & aRETU(0))
  '                      demitido = True
  '                      eRETU01 = aRETU(0)
  '                   End If
  '                End If
  '          End If
  '       End If
End Function

