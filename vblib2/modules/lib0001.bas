Attribute VB_Name = "xoldTool"
Option Explicit
Public Sub xmontatoolbar(ByRef vFormtool As Variant, _
                         ByVal VNome As String, Optional lLIB As Boolean = False, Optional ByVal vCONTROL As String = "Toolbar1")

  Dim btnx As Button
  Dim DB As ADODB.Connection
  Dim RSUSU As ADODB.Recordset
  Dim RSTAB As ADODB.Recordset
  Dim sSQL As String
  Dim aRETU As Variant
  Dim cCOM As String
  Dim cARQ As String

  'On Error Resume Next
  On Error GoTo errhandler




  vFormtool.ImageList = Nothing
  vFormtool.ImageList = frmPRINCIPAL.ImageList1
  
   cARQ = Dbname
  'Muda so aqui ate usar outros tipos  aqui nas grid esta usando ado fixo
If InStr(cARQ, "[VBSQLITE]") > 0 Or InStr(cARQ, "[SQLITERC6]") > 0 Or InStr(cARQ, "[TC6SQLITE]") > 0 Then
           cARQ = Replace(cARQ, "[VBSQLITE]", "[SQLITE]")
           cARQ = Replace(cARQ, "[SQLITERC6]", "[SQLITE]")
           cARQ = Replace(cARQ, "[TC6SQLITE]", "[SQLITE]")
End If

  
  
 aRETU = TipoConn(cARQ, , False)
 cCOM = aRETU(1)

  sSQL = "SELECT * FROM CONTROLE WHERE FORM = '" & VNome & "' and controle='" & vCONTROL & "' order by indice"

  Set DB = New ADODB.Connection
  
  
  DB.CursorLocation = adUseClient
  DB.Mode = adModeRead

  DB.ConnectionTimeout = 120
  DB.Open cCOM


  Set RSUSU = New ADODB.Recordset
  Set RSTAB = New ADODB.Recordset
  If aRETU(2) = "SQLITE" Then
     RSTAB.Open sSQL, DB, adOpenStatic, adLockReadOnly
  Else
     RSTAB.Open sSQL, DB, adOpenForwardOnly, adLockReadOnly
  End If

  Do While Not RSTAB.EOF

    Set btnx = vFormtool.Buttons.Add(, , Trim(RSTAB.Fields("Caption")), tbrDefault, CInt(0 & RSTAB("IMAGEM")))
    btnx.ToolTipText = RSTAB("TOOLTIP")
    btnx.Description = btnx.ToolTipText



    If zUSER = "ADMIN" Or RSTAB("DISPONIVEL") Then

      btnx.Enabled = True

    Else

      sSQL = "SELECT LIGADO FROM USUCAD WHERE FORM = '" & VNome & "' AND INDICE=" & RSTAB!indice & " AND IDUSUARIO=" & zUSERID
      If aRETU(2) = "SQLITE" Then
         RSUSU.Open sSQL, DB, adOpenStatic, adLockReadOnly
      Else
         RSUSU.Open sSQL, DB, adOpenForwardOnly, adLockReadOnly
      End If

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
  Exit Sub



errhandler:

  Select Case Err.Number
  Case Else
    SayErro "xMontaToolBar"
  End Select
End Sub
