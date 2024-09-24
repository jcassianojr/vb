Attribute VB_Name = "Toolbarlib"
Option Explicit

Public Sub MontaToolbar(ByRef vFormtool As Variant, ByVal nFORM As Long)
'Variaves Locais
  Dim oBTNX As Button
  Dim oDB As ADODB.Connection
  Dim oRS As ADODB.Recordset
  Dim cSQL, cNOME, cTOOL, cCAM As String
  Dim nICONE As Integer


  On Error GoTo errhandler

  vFormtool.ImageList = Nothing
  vFormtool.ImageList = frmPRINCIPAL.ImageList1

  cCAM = PegPath("PATH", "SYSCONF")
  If Not FileExist(cCAM, True) Then
    Exit Sub
  End If

  cSQL = "SELECT * FROM BOTOES WHERE FORM = " & nFORM & " ORDER BY INDICE"
  cCAM = GeracArq(cCAM, , False)
  Set oDB = New ADODB.Connection
  oDB.CursorLocation = adUseClient
  oDB.ConnectionTimeout = 120
  oDB.Open cCAM

  Set oRS = New ADODB.Recordset
  oRS.Open cSQL, oDB, adOpenForwardOnly, adLockReadOnly  ', adLockOptimistic ''adOpenStatic



  While Not oRS.EOF
    cNOME = FixStr(oRS("NOME"), "???", "TRIM")
    cTOOL = FixStr(oRS("TOOLTIP"), "???", "TRIM")
    nICONE = FixNum(oRS("Icone"))
    If nICONE > 0 Then
      Set oBTNX = vFormtool.Buttons.Add(, , cNOME, tbrDefault, nICONE)
      oBTNX.ToolTipText = cTOOL
      oBTNX.Enabled = True
    End If
    oRS.MoveNext
  Wend
  oRS.Close
  oDB.Close
  Set oBTNX = Nothing
  Set oRS = Nothing
  Set oDB = Nothing
  Exit Sub

errhandler:

  Select Case Err.Number
  Case Else
    SayErro "MontaToolBar"
  End Select

End Sub

