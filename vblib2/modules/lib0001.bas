Attribute VB_Name = "xoldTool"
Public Sub xmontatoolbar(ByRef vFormtool As Variant, _
                         ByVal VNome As String, _
                         Optional ByVal lLIB As Boolean = False)

  Dim btnx As Button
  Dim DB As ADODB.Connection
  Dim RSUSU As ADODB.Recordset
  Dim RSTAB As ADODB.Recordset

  Dim vCONTROL As String
  On Error Resume Next

  vCONTROL = "Toolbar1"

  vFormtool.ImageList = Nothing
  vFormtool.ImageList = frmPRINCIPAL.ImageList1

  sSQL = "SELECT * FROM CONTROLE WHERE FORM = '" & VNome & "' and controle='" & vCONTROL & "' order by indice"

  Set DB = New ADODB.Connection
  DB.CursorLocation = adUseClient
  DB.Mode = adModeRead

  DB.ConnectionTimeout = 120
  DB.Open GeracArq(Dbname, , False)


  Set RSUSU = New ADODB.Recordset
  Set RSTAB = New ADODB.Recordset
  RSTAB.Open sSQL, DB, adOpenForwardOnly, adLockReadOnly


  Do While Not RSTAB.EOF

    Set btnx = vFormtool.Buttons.Add(, , Trim(RSTAB.Fields("Caption")), tbrDefault, CInt(0 & RSTAB("IMAGEM")))
    btnx.ToolTipText = RSTAB("TOOLTIP")
    btnx.Description = btnx.ToolTipText



    If zUSER = "ADMIN" Or RSTAB("DISPONIVEL") Then  'or llib

      btnx.Enabled = True

    Else

      sSQL = "SELECT LIGADO FROM USUCAD WHERE FORM = '" & VNome & "' AND INDICE=" & RSTAB!indice & " AND IDUSUARIO=" & zUSERID
      RSUSU.Open sSQL, DB, adOpenForwardOnly, adLockReadOnly

      If Not RSUSU.EOF Then

        If RSUSU!ligado Then

          btnx.Enabled = True

        End If

      End If
      RSUSU.Close

    End If

    RSTAB.MoveNext

  Loop
  RSTAB.Close
  DB.Close

  Set btnx = Nothing
  Set RSTAB = Nothing
  Set RSUSU = Nothing
  Set DB = Nothing

End Sub


