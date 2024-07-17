Attribute VB_Name = "GridFuncoes"
Option Explicit

'mudar a cor da coluna
'    Dim i As Integer
'    For i = 1 To .rows - 1
'   .Row = i
'   .Col = 2:   .CellBackColor = &HC0FFFF
'    Next
Public Function LocalizaGri1(ByRef oGRID As Variant, Optional ByVal eBUSCA As Variant = "", _
                             Optional ByVal iBUSCA As Integer = 1, _
                             Optional ByVal lMES As Boolean = True, _
                             Optional ByVal iPos As Integer = 0) As Double
  LocalizaGrid oGRID, eBUSCA, iBUSCA, lMES, iPos
End Function

Public Function SomaGrid(ByRef oGRID As Variant, Optional ByVal nCOL As Integer = 0, Optional ByVal nROWINI As Integer = 0)
  Dim nROWS, nROW, x As Integer
  Dim nVAL, nTMPVAL As Variant
  nROWS = oGRID.Rows - 1                       ''gRID ROW Indice inicia 0
  oGRID.Col = nCOL
  For x = nROWINI To nROWS
    oGRID.Row = x
    nTMPVAL = oGRID
    nVAL = nVAL + FixNum(nTMPVAL)
  Next x
  SomaGrid = nVAL
End Function

Public Function PegUltGrid(ByRef oGRID As Variant, Optional ByVal nCOL As Integer = 0)
  Dim nROWS, nROW As Integer
  nROWS = oGRID.Rows - 1                       ''gRID ROW Indice inicia 0
  oGRID.Col = nCOL
  oGRID.Row = nROWS
  PegUltGrid = oGRID
End Function

Public Function LocalizaGrid(ByRef oGRID As Variant, Optional ByVal eBUSCA As Variant = "", _
                             Optional ByVal iBUSCA As Integer = 1, _
                             Optional ByVal lMES As Boolean = True, _
                             Optional ByVal iPos As Integer = 1) As Double
'Ogrid grid
'eBUSCA Expressao para Busca
'iBUSCA Indice
'lMes exibe mensagem padrao sim/nao
'iPOS Incremento posicional pois geralmente a col0 e o id e nao é buscada
  Dim nROWS, nROW, x As Double
  Dim nTAM As Long
  Dim lFOUND As Boolean
  Dim cCOMPARE As String
  lRETU = True
  LocalizaGrid = oGRID.Row
  nROWS = oGRID.Rows - 1                       ''gRID ROW Indice inicia 0
  nROW = oGRID.Row
  eBUSCA = FixStr(eBUSCA, "")
  iBUSCA = FixNum(iBUSCA)
  If eBUSCA = "" Or iBUSCA = 0 Then
    frmLocalizar.Show vbModal
    eBUSCA = FixStr(eRETU01)
    iBUSCA = FixNum(eRETU02) + iPos          '´´Mais Um Pois 0 e sempre id
  End If
  If iBUSCA > 0 Then iBUSCA = iBUSCA - 1       ''gRID COW Inicia 0
  oGRID.Col = iBUSCA
  nTAM = Len(eBUSCA)
  eBUSCA = UCase(eBUSCA)
  For x = 0 To nROWS
    oGRID.Row = x
    cCOMPARE = UCase(Left(FixStr(oGRID.tEXT), nTAM))
    If cCOMPARE = eBUSCA Then
      lFOUND = True
      oGRID.Col = oGRID.cols - 1
      oGRID.ColSel = 0
      If x > 0 Then
        oGRID.TopRow = x
      End If
      Exit For
    End If
  Next
  If Not lFOUND Then                           ''Volta a linha que estava
    lRETU = False
    oGRID.Row = nROW
  End If
  If lMES And Not lFOUND Then
    Alert "Não Localizado"
  End If
  eLOCALIZA = ""
End Function

Public Sub MontaGrid(ByRef oGRID As Variant, _
                     ByVal nITEM As Integer, _
                     ByVal aTAM As Variant, _
                     ByVal aDIZ As Variant, _
                     ByVal aCAM As Variant, _
                     ByVal cARQ As String, _
                     ByVal cSQL As String, _
                     Optional ByVal aFOR As Variant = 0)

  On Error GoTo errhandler
  Dim x As Integer
  Dim cDIZ As Variant
  Dim oDB
  Dim oRS
  Dim aRETU As Variant
  Dim cARQTMP As String
  Dim eVAR As Variant
  Dim aOPE As Variant
  Dim l3265 As Boolean
  Dim cERRO As String


  l3265 = True


  If Not FileExist(cARQ, True, , cSQL) Then
    Exit Sub
  End If

  aRETU = TipoConn(cARQ, , , False)

  cARQTMP = aRETU(1)

  If aRETU(2) <> "MDB" Then
    cSQL = Replace(cSQL, "[", "")
    cSQL = Replace(cSQL, "]", "")
  End If

  Select Case aRETU(0)
  Case "ADO"
    Set oDB = New ADODB.Connection
    oDB.ConnectionTimeout = 120
    oDB.Open cARQTMP



    Set oRS = New ADODB.Recordset
    oRS.Open cSQL, oDB, adOpenForwardOnly, adLockReadOnly
  End Select


  oGRID.Visible = False
  With oGRID
    .cols = nITEM
    .FixedCols = 0
    .Rows = 1
    .Row = 0
    For x = 0 To nITEM - 1
      .Col = x
      .tEXT = aDIZ(x)
      .ColWidth(x) = aTAM(x)
      If InStr(aCAM(x), "$") > 0 Then
        Select Case Mid(aCAM(x), 1, 1)
        Case "L"
          .ColAlignment(x) = flexAlignLeftCenter
        Case "R"
          .ColAlignment(x) = flexAlignRightCenter
        Case "C"
          .ColAlignment(x) = flexAlignCenterCenter
        End Select
        aCAM(x) = Mid(aCAM(x), InStr(aCAM(x), "$") + 1)
      End If
    Next

    While Not oRS.EOF
      cDIZ = ""
      For x = 0 To nITEM - 1
        aOPE = SepSqlOpe(aCAM(x))
        If aOPE(0) = "" Or aOPE(1) = "" Or aOPE(2) = "" Then  ''Operacao com Campos
          eVAR = oRS(aCAM(x))          ''Grava o Valor Default
        Else
          eVAR = MathOper(oRS(aOPE(1)), oRS(aOPE(2)), aOPE(0))
        End If
        Select Case VarType(eVAR)
        Case vbDate
          If IsNumeric(aFOR) Then
            eVAR = Fdata(eVAR, "DF")  ''Nao tem matriz formatacao
          Else
            If aFOR(x) = "" Then
              eVAR = Fdata(eVAR, "DF")  ''Nao tem formatacao
            End If
          End If
        Case vbBoolean
          If eVAR Then
            eVAR = " * "
          Else
            eVAR = "   "
          End If
        Case vbString
          If Len(eVAR) > 100 Then
            eVAR = Mid(eVAR, 1, 100)
          End If
          eVAR = Replace(eVAR, Chr(13) & Chr(10), " ")
        End Select
        If Not IsNumeric(aFOR) Then      'E Matriz
          If aFOR(x) <> "" Then
            eVAR = Format(eVAR, aFOR(x))
          End If
        End If

        cDIZ = cDIZ & eVAR & vbTab
      Next

      .AddItem cDIZ
      oRS.MoveNext

    Wend

  End With

  oGRID.Visible = True
  oRS.Close
  oDB.Close

  Set oRS = Nothing
  Set oDB = Nothing

  Exit Sub

errhandler:

  cERRO = "Montagrid " & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cSQL & Chr(13) & Chr(10)
  Select Case Err.Number
  Case 424, 3265                               ''Campo Inexistente
    If l3265 Then                            ''So uma vez o erro
      cERRO = cERRO & "Campo:" & aCAM(x) & Chr(13) & Chr(10)
      SayErro cERRO
      l3265 = False
    End If
    cDIZ = cDIZ & " ? " & vbTab
    Resume Next
  Case 3051, 3262
    oGRID.Visible = True
    Alert "Arquivo em Uso Exclusivo Outro Usuario"
    SayErro cERRO
    Exit Sub
  Case Else
    oGRID.Visible = True
    SayErro cERRO
    Exit Sub
  End Select

End Sub

Public Sub MontaGridFast(ByRef oGRID As Variant, _
                         ByVal nITEM As Integer, _
                         ByVal aTAM As Variant, _
                         ByVal aDIZ As Variant, _
                         ByVal aCAM As Variant, _
                         ByVal cARQ As String, _
                         ByVal cSQL As String, _
                         Optional ByVal aFOR As Variant = 0)

  On Error GoTo errhandler
  Dim x As Integer
  Dim cDIZ As Variant
  Dim oDB As New ADODB.Connection
  Dim oRS As New ADODB.Recordset
  Dim oCAM00 As ADODB.Field
  Dim oCAM01 As ADODB.Field
  Dim oCAM02 As ADODB.Field
  Dim oCAM03 As ADODB.Field
  Dim oCAM04 As ADODB.Field
  Dim oCAM05 As ADODB.Field
  Dim oCAM06 As ADODB.Field
  Dim oCAM07 As ADODB.Field
  Dim oCAM08 As ADODB.Field
  Dim oCAM09 As ADODB.Field
  Dim cARQTMP As String
  Dim eVAR As Variant
  Dim l3265 As Boolean
  Dim cERRO As String
  Dim lOPEN As Boolean

  lOPEN = False
  l3265 = True

  If Not FileExist(cARQ, True, , cSQL) Then
    Exit Sub
  End If

  If nITEM > 10 Then
    Alert ("MontaGridFast até 10 itens")
    Exit Sub
  End If

  cARQTMP = GeracArq(cARQ, , False)

  oDB.ConnectionTimeout = 120
  oDB.Open cARQTMP



  lOPEN = True

  oRS.Open cSQL, oDB, adOpenForwardOnly, adLockReadOnly

  oGRID.Visible = False
  With oGRID
    .cols = nITEM
    .FixedCols = 0
    .Rows = 1
    .Row = 0
    For x = 0 To nITEM - 1
      .Col = x
      .tEXT = aDIZ(x)
      .ColWidth(x) = aTAM(x)
      If InStr(aCAM(x), "$") > 0 Then
        Select Case Mid(aCAM(x), 1, 1)
        Case "L"
          .ColAlignment(x) = flexAlignLeftCenter
        Case "R"
          .ColAlignment(x) = flexAlignRightCenter
        Case "C"
          .ColAlignment(x) = flexAlignCenterCenter
        End Select
        aCAM(x) = Mid(aCAM(x), InStr(aCAM(x), "$") + 1)
      End If
    Next


    If Not oRS.EOF Then
      For x = 0 To nITEM - 1
        Select Case x
        Case 0
          Set oCAM00 = oRS(aCAM(0))
        Case 1
          Set oCAM01 = oRS(aCAM(1))
        Case 2
          Set oCAM02 = oRS(aCAM(2))
        Case 3
          Set oCAM03 = oRS(aCAM(3))
        Case 4
          Set oCAM04 = oRS(aCAM(4))
        Case 5
          Set oCAM05 = oRS(aCAM(5))
        Case 6
          Set oCAM06 = oRS(aCAM(6))
        Case 7
          Set oCAM07 = oRS(aCAM(7))
        Case 8
          Set oCAM08 = oRS(aCAM(8))
        Case 9
          Set oCAM09 = oRS(aCAM(9))
        End Select
      Next x
    End If



    While Not oRS.EOF
      cDIZ = ""
      For x = 0 To nITEM - 1
        Select Case x
        Case 0
          eVAR = oCAM00
        Case 1
          eVAR = oCAM01
        Case 2
          eVAR = oCAM02
        Case 3
          eVAR = oCAM03
        Case 4
          eVAR = oCAM04
        Case 5
          eVAR = oCAM05
        Case 6
          eVAR = oCAM06
        Case 7
          eVAR = oCAM07
        Case 8
          eVAR = oCAM08
        Case 9
          eVAR = oCAM09
        End Select



        Select Case VarType(eVAR)
        Case vbDate
          If IsNumeric(aFOR) Then
            eVAR = Fdata(eVAR, "DF")  ''Nao tem matriz formatacao
          Else
            If aFOR(x) = "" Then
              eVAR = Fdata(eVAR, "DF")  ''Nao tem formatacao
            End If
          End If
        Case vbBoolean
          If eVAR Then
            eVAR = " * "
          Else
            eVAR = "   "
          End If
        Case vbString
          If Len(eVAR) > 100 Then
            eVAR = Mid(eVAR, 1, 100)
          End If
          eVAR = Replace(eVAR, Chr(13) & Chr(10), " ")
        End Select
        If Not IsNumeric(aFOR) Then      'E Matriz
          If aFOR(x) <> "" Then
            eVAR = Format(eVAR, aFOR(x))
          End If
        End If

        cDIZ = cDIZ & eVAR & vbTab
      Next

      .AddItem cDIZ
      oRS.MoveNext

    Wend

  End With

  oGRID.Visible = True
  oRS.Close
  oDB.Close

  Set oRS = Nothing
  Set oDB = Nothing

  Exit Sub

errhandler:

  cERRO = "Montagrid " & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cSQL & Chr(13) & Chr(10)
  Select Case Err.Number
  Case 6
    Resume Next
  Case 424, 3265                               ''Campo Inexistente
    If l3265 Then                            ''So uma vez o erro
      cERRO = cERRO & "Campo:" & aCAM(x) & Chr(13) & Chr(10)
      SayErro cERRO
      l3265 = False
    End If
    cDIZ = cDIZ & " ? " & vbTab
    Resume Next
  Case 3051, 3262
    oGRID.Visible = True
    Alert "Arquivo em Uso Exclusivo Outro Usuario"
    If lOPEN Then
      ADOErro oRS.ActiveConnection.Errors, cERRO
    Else
      ADOErro oDB.Errors, cERRO
    End If
    Exit Sub
  Case Else
    oGRID.Visible = True
    If lOPEN Then
      ADOErro oRS.ActiveConnection.Errors, cERRO
    Else
      ADOErro oDB.Errors, cERRO
    End If
    Exit Sub
  End Select

End Sub

Public Sub MontaGridUltra(ByRef oGRID As Variant, _
                          ByVal nITEM As Integer, _
                          ByVal aTAM As Variant, _
                          ByVal aDIZ As Variant, _
                          ByVal aCAM As Variant, _
                          ByVal cARQ As String, _
                          ByVal cSQL As String, _
                          Optional ByVal aFOR As Variant = 0)

  On Error GoTo errhandler
  Dim x As Integer
  Dim cDIZ As Variant
  Dim oDB As New ADODB.Connection
  Dim oRS As New ADODB.Recordset
  Dim cARQTMP As String
  Dim l3265 As Boolean
  Dim cERRO As String
  Dim lOPEN As Boolean

  lOPEN = False
  l3265 = True



  If Not FileExist(cARQ, True, , cSQL) Then
    Exit Sub
  End If

  cARQTMP = GeracArq(cARQ, , False)


  oDB.ConnectionTimeout = 120
  oDB.Open cARQTMP



  lOPEN = True

  oRS.Open cSQL, oDB, adOpenStatic, adLockReadOnly

  oGRID.Visible = False
  With oGRID
    .FixedCols = 0
    .Row = 0
    oGRID.Rows = oRS.RecordCount + 1
    oGRID.cols = nITEM
    For x = 0 To nITEM - 1
      .Col = x
      .ColWidth(x) = aTAM(x)
      If InStr(aCAM(x), "$") > 0 Then
        Select Case Mid(aCAM(x), 1, 1)
        Case "L"
          .ColAlignment(x) = flexAlignLeftCenter
        Case "R"
          .ColAlignment(x) = flexAlignRightCenter
        Case "C"
          .ColAlignment(x) = flexAlignCenterCenter
        End Select
        aCAM(x) = Mid(aCAM(x), InStr(aCAM(x), "$") + 1)
      End If
    Next
  End With



  oRS.MoveFirst

  'define o numero de linhas e colunas e configura o grid
  oGRID.Row = 1
  oGRID.Col = 0
  oGRID.RowSel = oGRID.Rows - 1
  oGRID.ColSel = oGRID.cols - 1

  'estamos usando a propriedade Clip e o método GetString para selecionar uma região do grid
  oGRID.clip = oRS.GetString(adClipString, -1, Chr(9), Chr(13), vbNullString)
  oGRID.Row = 0

  With oGRID
    For x = 0 To nITEM - 1
      .Col = x
      .tEXT = aDIZ(x)
    Next x
  End With
  oGRID.Visible = True




  oRS.Close
  oDB.Close

  Set oRS = Nothing
  Set oDB = Nothing

  Exit Sub

errhandler:

  cERRO = "Montagrid " & Chr(13) & Chr(10) & cARQ & Chr(13) & Chr(10) & cSQL & Chr(13) & Chr(10)
  Select Case Err.Number
  Case 6
    Resume Next
  Case 424, 3265                               ''Campo Inexistente
    If l3265 Then                            ''So uma vez o erro
      cERRO = cERRO & "Campo:" & aCAM(x) & Chr(13) & Chr(10)
      SayErro cERRO
      l3265 = False
    End If
    cDIZ = cDIZ & " ? " & vbTab
    Resume Next
  Case 3051, 3262
    oGRID.Visible = True
    Alert "Arquivo em Uso Exclusivo Outro Usuario"
    If lOPEN Then
      ADOErro oRS.ActiveConnection.Errors, cERRO
    Else
      ADOErro oDB.Errors, cERRO
    End If
    Exit Sub
  Case 3021
    'eof bof
    oGRID.Visible = True
    Exit Sub
  Case Else
    oGRID.Visible = True
    If lOPEN Then
      ADOErro oRS.ActiveConnection.Errors, cERRO
    Else
      ADOErro oDB.Errors, cERRO
    End If
    Exit Sub
  End Select

End Sub

Public Function CloneGrid(ByRef oGRIDORI As Variant, ByRef oGRIDDES As Variant)
  Dim nROWS As Long
  Dim nCOLS As Long
  Dim x As Long
  Dim y As Long
  Dim cLINHA As String
  oGRIDORI.Visible = False
  oGRIDDES.Visible = False
  nROWS = oGRIDORI.Rows - 1
  nCOLS = oGRIDORI.cols - 1
  oGRIDORI.Row = 0
  With oGRIDDES
    .cols = oGRIDORI.cols
    .FixedCols = 0
    .Rows = 1
    .Row = 0
    For x = 0 To nCOLS
      oGRIDORI.Col = x
      .Col = x
      .tEXT = oGRIDORI
      .ColWidth(x) = oGRIDORI.ColWidth(x)
      .ColAlignment(x) = oGRIDORI.ColAlignment(x)
    Next
    For y = 1 To nROWS                       ''0 cabecario
      cLINHA = ""
      oGRIDORI.Row = y
      For x = 0 To nCOLS
        oGRIDORI.Col = x
        cLINHA = cLINHA & oGRIDORI & vbTab
      Next x
      .AddItem cLINHA
    Next y
  End With
  oGRIDORI.Visible = True
  oGRIDDES.Visible = True
End Function


