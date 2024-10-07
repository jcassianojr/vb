VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form ExpMS03 
   Caption         =   "Transferencias Operacação/Sequencia"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6735
   Icon            =   "expms03.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin XPControls.XPText Text 
      Height          =   315
      Index           =   0
      Left            =   300
      TabIndex        =   7
      Top             =   420
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton CmdEscMs01 
      Height          =   375
      Left            =   3060
      TabIndex        =   4
      Top             =   1560
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "expms03.frx":058A
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton CmdTransfer 
      Height          =   375
      Left            =   3780
      TabIndex        =   2
      Top             =   960
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   661
      Picture         =   "expms03.frx":0B24
      Caption         =   "Transferir Composição"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton CmdCancel 
      Height          =   435
      Left            =   3840
      TabIndex        =   3
      Top             =   420
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   767
      Picture         =   "expms03.frx":10BE
      Caption         =   "Retornar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton CmdEscPfim 
      Height          =   375
      Index           =   0
      Left            =   3060
      TabIndex        =   5
      Top             =   420
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "expms03.frx":1658
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton CmdEscPfim 
      Height          =   375
      Index           =   1
      Left            =   3060
      TabIndex        =   6
      Top             =   900
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "expms03.frx":1BF2
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPText Text 
      Height          =   315
      Index           =   1
      Left            =   360
      TabIndex        =   8
      Top             =   1620
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPText Text 
      Height          =   315
      Index           =   2
      Left            =   300
      TabIndex        =   9
      Top             =   960
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton XPButton1 
      Height          =   375
      Left            =   3780
      TabIndex        =   10
      Top             =   1380
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   661
      Picture         =   "expms03.frx":218C
      Caption         =   "Transferir Sequencia"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton XPButton2 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   661
      Picture         =   "expms03.frx":2726
      Caption         =   "Sincronizar Cod.Interno PFS-->MS06"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton XPButton3 
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   661
      Picture         =   "expms03.frx":2CC0
      Caption         =   "Sincronizar Cod.Interno PF-->MS01"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton XPButton4 
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      Top             =   2160
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   661
      Picture         =   "expms03.frx":325A
      Caption         =   "Sincronizar Cod.Cliente MA01-->PF"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton XPButton5 
      Height          =   375
      Left            =   3360
      TabIndex        =   15
      Top             =   2640
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   661
      Picture         =   "expms03.frx":37F4
      Caption         =   "Sincronizar Pesos PF-->MS01"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton XPButton6 
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   3120
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   661
      Picture         =   "expms03.frx":3D8E
      Caption         =   "Sincronizar  PFS-->Ferramenta"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LblDiz 
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3780
      Width           =   6555
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo Destino PCP/Faturamento"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   1380
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo Final Processo Fabricação"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "ExpMS03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sDESCRI As String

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub CmdEscMs01_Click()
  ePASS01 = "MANA5"
  escms01.Show vbModal, Me
  If lRETU Then
    Text(1).Text = eRETU01
  End If
End Sub

Private Sub CmdEscPfim_Click(Index As Integer)
  escpffim.Show vbModal, Me
  If lRETU Then
    If Index = 0 Then Text(0).Text = eRETU01
    If Index = 1 Then Text(2).Text = eRETU01
  End If
End Sub

Private Sub CmdTransfer_Click()

'* On Error Resume Next
  Dim aPF() As Variant
  Dim dbPF As New ADODB.Connection
  Dim rsPF As New ADODB.Recordset
  Dim sqlPF, sqlMS03 As String
  Dim lnrPF, x As Long
  Dim CON As ADODB.Connection
  Dim rs As ADODB.Recordset
  Dim aCAMPOS As Variant
  Dim nBSEQ, nBSSQ As Long
  Dim nHORAS As Single
  Dim cOPCAO As String
  Dim cARQ As String

  cARQ = PegPath("PATH", "PF")

  x = 0
  dbPF.ConnectionTimeout = 120
  dbPF.Open GeracArq(cARQ, , False)

  If Text(2).Text = "" Then

    sqlPF = "SELECT * FROM PF WHERE CODFINAL='" & Trim(Text(0).Text) & "' AND NOT BLOQUEADO ORDER BY PF"

  Else

    sqlPF = "SELECT * FROM PF WHERE CODFINAL='" & Trim(Text(0).Text) & "' OR CODFINAL='" & Trim(Text(2).Text) & "'AND NOT BLOQUEADO ORDER BY PF"

  End If

  rsPF.Open sqlPF, dbPF, adOpenForwardOnly, adLockReadOnly

  If rsPF.EOF Then
    dbPF.Close
    rsPF.Close
    Set dbPF = Nothing
    Set rsPF = Nothing
    Alert "Produto Final não Disponivel"
    Exit Sub
  End If

  lnrPF = rsPF.RecordCount
  ReDim aPF(lnrPF)

  While Not rsPF.EOF

    aPF(x) = rsPF("PF")
    rsPF.MoveNext
    x = x + 1

  Wend

  rsPF.Close
  Set rsPF = Nothing

  '* Open a connection.
  Set CON = New ADODB.Connection
  CON.ConnectionTimeout = 120
  CON.Open "mana5emp"

  sqlMS03 = "DELETE FROM MS03 where codigo='" & Text(1).Text & "'"
  CON.Execute sqlMS03

  Set rs = New ADODB.Recordset
  'RS.CursorType = adOpenKeyset
  rs.CursorType = adOpenKeyset
  rs.LockType = adLockOptimistic
  rs.Open "MS03", CON, , , adCmdTable

  'aCAMPOS(0Tipoent,1Codcomp,2nomecomp,3qtdde,4preco,5total,6baixac,7bseq,8bssq,9opcao)

  'Materia Prima PF
  x = 0

  For x = 0 To lnrPF - 1

    sqlPF = "SELECT * FROM PF WHERE PF=" & aPF(x)
    rsPF.Open sqlPF, dbPF, adOpenForwardOnly, adLockReadOnly

    If Not rsPF.EOF Then

      While Not rsPF.EOF

        If rsPF("CODMU011") <> "" Then

          PEGDADOS "M", rsPF("CODMU011")
          aCAMPOS = Array("M", rsPF("CODMU011"), eRETU01, rsPF("PESMU011"), eRETU02, 0, "S", 0, 0, "1")
          GRAVAMS03 aCAMPOS, rs

        End If

        If rsPF("CODMU012") <> "" Then

          PEGDADOS "M", rsPF("CODMU012")
          aCAMPOS = Array("M", rsPF("CODMU012"), eRETU01, rsPF("PESMU012"), eRETU02, 0, "S", 0, 0, "2")
          GRAVAMS03 aCAMPOS, rs

        End If

        If rsPF("CODMU013") <> "" Then

          PEGDADOS "M", rsPF("CODMU013")
          aCAMPOS = Array("M", rsPF("CODMU012"), eRETU01, rsPF("PESMU013"), eRETU02, 0, "S", 0, 0, "3")
          GRAVAMS03 aCAMPOS, rs

        End If

        rsPF.MoveNext

      Wend

    End If

    rsPF.Close
    Set rsPF = Nothing

  Next x

  'Componentes pfms03
  x = 0

  For x = 0 To lnrPF - 1

    Debug.Print aPF(x)
    sqlPF = "SELECT * FROM PFMS03 WHERE PF=" & aPF(x) & " ORDER BY SEQ,SSQ"
    rsPF.Open sqlPF, dbPF, adOpenForwardOnly, adLockReadOnly

    If Not rsPF.EOF Then

      While Not rsPF.EOF

        PEGDADOS "C", rsPF("CODCOMP")
        aCAMPOS = Array("C", rsPF("CODCOMP"), eRETU01, rsPF("QTDDE"), eRETU02, 0, rsPF("BAIXAC"), rsPF("SEQ"), rsPF("SSQ"), CStr(rsPF("OPCAO")))
        GRAVAMS03 aCAMPOS, rs
        rsPF.MoveNext

      Wend

    End If

    rsPF.Close
    Set rsPF = Nothing

  Next x

  'Mao de Obra Equipamento/Homem/Terceiros PFS
  x = 0

  For x = 0 To lnrPF - 1

    sqlPF = "SELECT * FROM PFS WHERE PF=" & aPF(x) & " ORDER BY SEQ,SSQ"
    rsPF.Open sqlPF, dbPF, adOpenForwardOnly, adLockReadOnly

    If Not rsPF.EOF Then

      While Not rsPF.EOF

        nBSEQ = rsPF!SEQ
        nBSSQ = rsPF!SSQ
        nHORAS = 0
        cOPCAO = "1"

        If rsPF("PCHORA") > 0 Then

          nHORAS = 1 / rsPF("PCHORA")

        End If

        If Not IsNull(rsPF("OPCAO")) Then

          cOPCAO = CStr(rsPF("OPCAO"))

        End If

        If rsPF("CODMP01") <> "" Then

          PEGDADOS "E", rsPF("CODMP01")
          aCAMPOS = Array("E", rsPF("CODMP01"), eRETU01, nHORAS, eRETU02, 0, "S", nBSEQ, nBSSQ, cOPCAO)
          GRAVAMS03 aCAMPOS, rs

        End If

        If rsPF("CODMP02") <> "" Then

          PEGDADOS "H", rsPF("CODMP02")
          aCAMPOS = Array("H", rsPF("CODMP02"), eRETU01, nHORAS, eRETU02, 0, "S", nBSEQ, nBSSQ, cOPCAO)
          GRAVAMS03 aCAMPOS, rs

        End If

        If rsPF("CODMP02B") <> "" Then

          PEGDADOS "H", rsPF("CODMP02B")
          aCAMPOS = Array("H", rsPF("CODMP02B"), eRETU01, nHORAS, eRETU02, 0, "S", nBSEQ, nBSSQ, cOPCAO)
          GRAVAMS03 aCAMPOS, rs

        End If

        If rsPF("CODMP02C") <> "" Then

          PEGDADOS "H", rsPF("CODMP02C")
          aCAMPOS = Array("H", rsPF("CODMP02C"), eRETU01, nHORAS, eRETU02, 0, "S", nBSEQ, nBSSQ, cOPCAO)
          GRAVAMS03 aCAMPOS, rs

        End If

        If rsPF("CODMP02D") <> "" Then

          PEGDADOS "H", rsPF("CODMP02D")
          aCAMPOS = Array("H", rsPF("CODMP02D"), eRETU01, nHORAS, eRETU02, 0, "S", nBSEQ, nBSSQ, cOPCAO)
          GRAVAMS03 aCAMPOS, rs

        End If

        If rsPF("CODMP03") <> "" Then

          PEGDADOS "T", rsPF("CODMP03")
          aCAMPOS = Array("T", rsPF("CODMP03"), eRETU01, nHORAS, eRETU02, 0, "S", nBSEQ, nBSSQ, cOPCAO)
          GRAVAMS03 aCAMPOS, rs

        End If

        rsPF.MoveNext

      Wend

    End If

    rsPF.Close
    Set rsPF = Nothing

  Next x

  rs.Close
  Set rs = Nothing
  CON.Close
  Set CON = Nothing
  Set dbPF = Nothing
  Alert "Transferencia Concluida"

End Sub

Private Sub GRAVAMS03(aCAMPOS As Variant, rs As Variant)

  sDESCRI = CStr(aCAMPOS(2))
  sDESCRI = LCase(sDESCRI)
  sDESCRI = Left(sDESCRI, 40)
  sDESCRI = Tirace(sDESCRI)

  rs.AddNew
  rs("codigo") = Text(1).Text
  rs("TIPOENT") = aCAMPOS(0)
  rs("CODCOMP") = aCAMPOS(1)
  rs("NOMECOMP") = sDESCRI
  rs("QTDDE") = aCAMPOS(3)
  rs("PRECO") = aCAMPOS(4)
  rs("TOTAL") = aCAMPOS(3) * aCAMPOS(4)
  rs("BAIXAC") = aCAMPOS(6)
  rs("BSEQ") = aCAMPOS(7)
  rs("BSSQ") = aCAMPOS(8)
  rs("OPCAO") = aCAMPOS(9)
  rs.Update
  rs.MoveFirst

End Sub

Private Sub PEGDADOS(cTIP As String, cCODIGO As String)
  Dim cCAMPO As String
  Dim cARQ As String
  Dim sSQL As String
  Dim aRETU As Variant

  Select Case cTIP
  Case "E"
    sSQL = "SELECT * FROM MP01 WHERE CODIGO='" & cCODIGO & "'"
  Case "H"
    sSQL = "SELECT * FROM MP02 WHERE CODIGO='" & cCODIGO & "'"
  Case "T"
    sSQL = "SELECT * FROM MP03 WHERE CODIGO='" & cCODIGO & "'"
  Case "M"
    sSQL = "SELECT * FROM MU01 WHERE CODIGO='" & cCODIGO & "'"
  Case "C"
    sSQL = "SELECT * FROM MT01 WHERE CODIGO='" & cCODIGO & "'"
  Case "S"
    sSQL = "SELECT * FROM MS01 WHERE CODIGO='" & cCODIGO & "'"
  Case "P"
    sSQL = "SELECT * FROM MS01 WHERE CODIGO='" & cCODIGO & "'"
  Case "O"
    sSQL = "SELECT * FROM MW05 WHERE CODIGO='" & cCODIGO & "'"
  End Select
  Select Case cTIP
  Case "E", "H", "T"
    cCAMPO = "VALOR"
  Case "M", "C", "S"
    cCAMPO = "PRECUST"
  Case "P"
    cCAMPO = "CUSTF"
  Case "O"
    cCAMPO = "COTVAL"
  End Select
  cARQ = GeracArq(zMANA5EMP, "JETFOX")
  aRETU = PegSQL(cARQ, sSQL, 3, Array("NOME", cCAMPO, "ULTPRC"), Array("C", "N", "N"), Array("", 0, 0))

  eRETU01 = aRETU(0)
  eRETU02 = aRETU(1)
  If eRETU02 < 0.01 Then
    eRETU02 = aRETU(2)
  End If


End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
End Sub

Private Sub tEXT_KeyPress(Index As Integer, KeyAscii As Integer)

  TeclaEnter KeyAscii

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault

End Sub

Private Sub XPButton1_Click()
'* On Error Resume Next
  Dim aPF() As Variant
  Dim aPFI() As Variant
  Dim dbPF As New ADODB.Connection
  Dim rsPF As New ADODB.Recordset
  Dim sqlPF, sqlms06 As String
  Dim lnrPF, x As Long
  Dim CON As ADODB.Connection
  Dim rs As ADODB.Recordset
  Dim sDESCRI As String
  Dim cARQ As String
  cARQ = PegPath("PATH", "PF")

  x = 0
  dbPF.ConnectionTimeout = 120
  dbPF.Open GeracArq(cARQ, , False)

  If Text(2) = "" Then

    sqlPF = "SELECT * FROM PF WHERE CODFINAL='" & Trim(Text(0).Text) & "' AND NOT BLOQUEADO ORDER BY PF"

  Else

    sqlPF = "SELECT * FROM PF WHERE CODFINAL='" & Trim(Text(0).Text) & "' OR CODFINAL='" & Trim(Text(2).Text) & "'AND NOT BLOQUEADO ORDER BY PF"

  End If

  rsPF.Open sqlPF, dbPF, adOpenForwardOnly, adLockReadOnly

  If Not rsPF.EOF Then

    rsPF.MoveFirst

  Else

    Set dbPF = Nothing
    Set rsPF = Nothing
    Alert "Produto Final não Disponivel"
    Exit Sub

  End If

  lnrPF = rsPF.RecordCount
  ReDim aPF(lnrPF)
  ReDim aPFI(16)

  While Not rsPF.EOF

    aPF(x) = rsPF("PF")
    rsPF.MoveNext
    x = x + 1

  Wend

  rsPF.Close
  Set rsPF = Nothing

  '* Open a connection.
  Set CON = New ADODB.Connection

  CON.ConnectionTimeout = 120
  CON.Open "mana5emp"

  sqlms06 = "DELETE FROM MS06 where codigo='" & Text(1).Text & "'"

  CON.Execute sqlms06

  Set rs = New ADODB.Recordset
  rs.CursorType = adOpenKeyset
  rs.LockType = adLockOptimistic

  rs.Open "MS06", CON, , , adCmdTable

  x = 0

  For x = 0 To lnrPF - 1

    Debug.Print aPF(x)
    sqlPF = "SELECT * FROM PFS WHERE PF=" & aPF(x) & " ORDER BY SEQ,SSQ"
    rsPF.Open sqlPF, dbPF, adOpenForwardOnly, adLockReadOnly

    If Not rsPF.EOF Then

      While Not rsPF.EOF

        aPFI(0) = 0 + rsPF("SEQ")
        aPFI(1) = 0 + rsPF("SSQ")
        aPFI(2) = "" & rsPF("DESCRI")
        aPFI(3) = "" & rsPF("CODMP01")
        aPFI(4) = "" & rsPF("CODMP02")
        aPFI(5) = "" & rsPF("CODMP03")
        aPFI(6) = 0 + rsPF("PCHORA")
        aPFI(7) = "" & rsPF("CODMP02B")
        aPFI(8) = "" & rsPF("CODMP02C")
        aPFI(9) = "" & rsPF("CODMP02D")
        aPFI(10) = 0 + rsPF("FATBAT")

        If IsNumeric(rsPF("FATCORTE")) Then

          If rsPF("fatcorte") > 0 And rsPF("FATBAT") > 0 Then

            aPFI(10) = 0 + (rsPF("FATBAT") / rsPF("fatcorte"))

          End If

        End If

        aPFI(11) = "" & rsPF("FERRAMEN")
        aPFI(12) = "" & rsPF("FERRAME2")
        aPFI(13) = "" & rsPF("FERRAME3")
        aPFI(14) = "" & rsPF("FERRAME4")
        aPFI(15) = "" & rsPF("MONTAGEM")
        rsPF.MoveNext
        rs.AddNew
        sDESCRI = CStr(aPFI(2))
        sDESCRI = LCase(sDESCRI)
        sDESCRI = Left(sDESCRI, 70)
        sDESCRI = Tirace(sDESCRI)
        rs("codigo") = Text(1).Text
        rs("SEQ") = aPFI(0)
        rs("SSQ") = aPFI(1)
        rs("DESCRI") = sDESCRI
        rs("CODMP01") = aPFI(3)
        rs("CODMP02") = aPFI(4)
        rs("CODMP03") = aPFI(5)
        rs("PCHORA") = aPFI(6)
        rs("CODMP02B") = aPFI(7)
        rs("CODMP02C") = aPFI(8)
        rs("CODMP02D") = aPFI(9)
        rs("FATBAT") = IIf(IsNull(aPFI(10)), 0, aPFI(10))
        rs("FERRAMEN") = aPFI(11)
        rs("FERRAME2") = aPFI(12)
        rs("FERRAME3") = aPFI(13)
        rs("FERRAME4") = aPFI(14)

        If Not IsNull(aPFI(15)) Then

          If aPFI(15) = "S" Then

            rs("TIPFEC") = "9"

          End If

        End If

        rs.Update

      Wend

      rs.AddNew
      rs("codigo") = Text(1).Text
      rs("SEQ") = aPFI(0)
      rs("SSQ") = 99
      rs("TIPFEC") = "0"
      rs("CODFEC") = Text(0).Text
      rs("DESCRI") = "Final"
      rs.Update

    End If

    rsPF.Close
    Set rsPF = Nothing

  Next x

  rs.Close
  Set rs = Nothing
  CON.Close
  Set CON = Nothing
  Set dbPF = Nothing
  Alert "Transferencia Concluida"

End Sub

Private Sub XPButton2_Click()
  Dim oCON As New ADODB.Connection
  Dim oRS As New ADODB.Recordset
  Dim cARQ As String
  Dim sSQL As String
  Dim cCODIGO As String
  Dim cCODINT As String
  'Dim cARQSDE As String
  Dim aRETU As Variant
  Dim cARQPF As String
  Dim cLINHA As String
  Dim oTXT As Criatxt
  Set oTXT = New Criatxt
  oTXT.arquivo = "c:\temp\PFS.txt"             ''Nome do Arquivo
  oTXT.ABRIR


  cARQPF = GeracArq(PegPath("PATH", "PF"))
  cARQ = GeraConn(zMANA5EMP, "JETFOX")

  oCON.ConnectionTimeout = 120

  oCON.Open cARQPF
  sSQL = "SELECT A.PF,A.CODIGO,A.CODFINAL,A.CODFISCAL,B.SEQ,B.SSQ,B.CODINT FROM PF AS A,PFS AS B WHERE A.PF=B.PF "
  oRS.Open sSQL, oCON, adOpenDynamic, adLockOptimistic
  While Not oRS.EOF
    nPF = FixInt(oRS("PF"))
    cCODIGO = FixStr(oRS("CODFISCAL"), " ", "TRIM")
    If Len(cCODIGO) = 0 Then
      cCODIGO = FixStr(oRS("CODFINAL"), " ", "TRIM")
    End If
    If Len(cCODIGO) = 0 Then
      cCODIGO = FixStr(oRS("CODIGO"), " ", "TRIM")
    End If
    cCODINT = FixStr(oRS("CODINT"), " ", "TRIM")
    LblDiz.Caption = cCODIGO
    LblDiz.Refresh
    sSQL = "SELECT CODINT,PF FROM MS06 WHERE CODIGO='" & cCODIGO & "' AND SEQ=" & oRS("SEQ") & " AND SSQ=" & oRS("SSQ")
    aRETU = PegSQL(cARQ, sSQL, 2, Array("CODINT", "PF"), Array("C", "N"), Array("", 0))
    If lRETU Then
      If Len(cCODINT) = 0 Then
        If Len(Trim(aRETU(0))) > 0 Then  'pega o do ms06
          oRS("CODINT") = FixStr(aRETU(0))
          oRS.Update
        End If
      Else
        If Len(Trim(aRETU(0))) = 0 Then  'nao tem no ms06
          If Len(Trim(cCODINT)) = 14 Then
            sSQL = "SELECT FROM MS06 WHERE ="
            sSQL = sSQL & PadRight(cCODIGO, 24) + PadLeft(oRS("SEQ"), 3) + PadLeft(oRS("SSQ"), 3)
            GrvSQL GeraConn(zMANA5EMP, "FOX"), sSQL, 2, Array("CODINT", "PF"), Array(cCODINT, nPF), Array("C", "N")  'grvsqlsde
          End If
        End If
        If aRETU(1) = 0 Then             'nao esta gravado o numero do pf
          sSQL = "SELECT FROM MS06 WHERE ="
          sSQL = sSQL & PadRight(cCODIGO, 24) + PadLeft(oRS("SEQ"), 3) + PadLeft(oRS("SSQ"), 3)
          GrvSQL GeraConn(zMANA5EMP, "FOX"), sSQL, 1, Array("PF"), Array(nPF), Array("N")  ''GrvSQLSDE zMANA5EMP
        End If
      End If
    End If
    Debug.Print cCODIGO
    If Len(cCODINT) > 0 Then
      cLINHA = cCODINT & "|" & cCODIGO & "|" & oRS("SEQ") & "/" & oRS("SSQ") & "|" & "|" & "|" & "|"
      oTXT.implinha (cLINHA)
    End If
    oRS.MoveNext
  Wend
  oRS.Close
  oCON.Close
  oTXT.fechar
  Alert "Transferencia Concluida"
End Sub

Private Sub XPButton3_Click()
  Dim oCON As New ADODB.Connection
  Dim oRS As New ADODB.Recordset
  Dim cARQ As String
  Dim sSQL As String
  Dim cCODIGO As String
  Dim cCODINT As String
  Dim aRETU As Variant
  Dim cARQPF As String
  cARQPF = GeracArq(PegPath("PATH", "PF"))
  cARQ = GeraConn(zMANA5EMP, "JETFOX")
  oCON.Open cARQPF
  oRS.Open "SELECT CODIGO,CODIGOINT,CODFISCAL,CODFINAL FROM PF ", oCON, adOpenDynamic, adLockOptimistic
  While Not oRS.EOF
    cCODIGO = FixStr(oRS("CODFISCAL"), " ", "TRIM")
    If Len(cCODIGO) = 0 Then
      cCODIGO = FixStr(oRS("CODFINAL"), " ", "TRIM")
    End If
    If Len(cCODIGO) = 0 Then
      cCODIGO = FixStr(oRS("CODIGO"), " ", "TRIM")
    End If
    LblDiz.Caption = cCODIGO
    LblDiz.Refresh
    cCODINT = ""
    If Not IsNull(oRS("CODIGOINT")) Then
      cCODINT = FixStr(oRS("CODIGOINT"), " ", "TRIM")
    End If
    If Len(cCODINT) = 0 Then
      sSQL = "SELECT NOME,CODIGOINT FROM MS01 WHERE CODIGO='" & cCODIGO & "'"
      aRETU = PegSQL(cARQ, sSQL, 2, Array("NOME", "CODIGOINT"), Array("C", "C"), Array("", ""))
      If lRETU Then
        If Len(cCODINT) = 0 Then
          If Len(Trim(aRETU(1))) > 0 Then  'pega o do ms01
            oRS("CODIGOINT") = aRETU(1)
            oRS.Update
          End If
        Else
          If Len(Trim(aRETU(1))) = 0 Then  'nao tem no ms01
            sSQL = "SELECT FROM MS01 WHERE =" & PadRight(cCODIGO, 24)
            GrvSQL GeraConn(zMANA5EMP, "FOX"), sSQL, 1, Array("CODINT"), Array(cCODINT), Array("C")  ''GrvSQLSDE zMANA5EMP
          End If
        End If
      End If
    End If
    Debug.Print cCODIGO
    oRS.MoveNext
  Wend
  oRS.Close
  oCON.Close
  Alert "Transferencia Concluida"
End Sub

Private Sub XPButton4_Click()
  Dim oCON As New ADODB.Connection
  Dim oRS As New ADODB.Recordset
  Dim cARQ As String
  Dim sSQL As String
  Dim cCODCLI As String
  Dim aRETU As Variant
  Dim nNUMERO As Long
  Dim cARQPF As String
  cARQPF = GeracArq(PegPath("PATH", "PF"))
  cARQ = GeraConn(zMANA5EMP, "JETFOX")
  oCON.Open cARQPF
  oRS.Open "SELECT PF,CLIENTE,CODCLIENTE FROM PF ", oCON, adOpenDynamic, adLockOptimistic
  While Not oRS.EOF
    If IsNull(oRS("CODCLIENTE")) Then
      cCODCLI = ""
    Else
      cCODCLI = FixStr(oRS("CODCLIENTE"), "", "TRIM")
    End If

    LblDiz.Caption = cCODCLI
    LblDiz.Refresh


    nNUMERO = FixInt(oRS("CLIENTE"))
    If Len(cCODCLI) = 0 Then                 'Or (nNUMERO >= 1200 And nNUMERO < 1299)
      sSQL = "SELECT NOME,CODIGO FROM MA01 WHERE NUMERO=" & nNUMERO
      aRETU = PegSQL(cARQ, sSQL, 2, Array("NOME", "CODIGO"), Array("C", "C"), Array("", ""))
      If lRETU Then
        oRS("CODCLIENTE") = aRETU(1)
        oRS.Update
      End If
    End If
    Debug.Print oRS("PF")
    oRS.MoveNext
  Wend
  oRS.Close
  oCON.Close
End Sub

Private Sub XPButton5_Click()
  Dim dbPF As New ADODB.Connection
  Dim rsPF As New ADODB.Recordset
  Dim sqlPF As String
  Dim cCODIGO As String
  Dim cARQPF As String
  Dim SACAREA As Variant
 ' Dim iRETVAL As Variant


  cARQPF = PegPath("PATH", "PF")
  dbPF.Open GeracArq(cARQPF, , False)
  SACAREA = GeraConn(zMANA5EMP, "JETFOX")
  sqlPF = "SELECT PF.CODFINAL, PF.PESLIQ, IE.TARA FROM PF,IE WHERE ( PF.PF = IE.PF ) AND PF.CODFINAL<>"" AND PF.CODFINAL IS NOT NULL"
  rsPF.Open sqlPF, dbPF, adOpenForwardOnly, adLockReadOnly
  If rsPF.EOF Then
    Set dbPF = Nothing
    Set rsPF = Nothing
    Alert "Produtos Finais nao Disponivel PF"
    Exit Sub
  End If
  While Not rsPF.EOF
    cCODIGO = FixStr(rsPF("CODFINAL"))
    LblDiz.Caption = cCODIGO
    LblDiz.Refresh
    'criar opcao futura con recodset e find
    GrvSQL SACAREA, "SELECT PESUNIPF,PESEMBPF FROM MS01 WHERE CODIGO='" & cCODIGO & "'", 2, Array("PESUNIPF", "PESEMBPF"), _
           Array(FixNum(rsPF("PESLIQ")), FixNum(rsPF("TARA"))), Array("N", "N")
    rsPF.MoveNext
  Wend
  rsPF.Close
  dbPF.Close
  Set rsPF = Nothing
  Set dbPF = Nothing

  Alert "Transferencia Concluida"


End Sub

Private Sub XPButton6_Click()
  Dim oCON As New ADODB.Connection
  Dim oRS As New ADODB.Recordset
  Dim cARQ As String
  Dim sSQL As String
  Dim cCODIGO As String
  'Dim cCODINT As String
 ' Dim cARQUSO As String
 ' Dim cARQCDX As String
 ' Dim cALIAS As String
  Dim cARQPF As String
  Dim x As Integer
  Dim cFERRAM As String
  Dim nHANDLE
  Dim cnn As New ADODB.Connection
  Dim rst As New ADODB.Recordset
  Dim cCONTMP As String
  Dim cSQLTMP As String
  Dim cCRITERIO As String
  
  cARQ = PegPath("PATH", "MANA5FER")
  cCONTMP = GeraConn(cARQ, "JETFOX")
  cSQLTMP = "SELECT FERRAM,SEQ,SSQ,PECA,PF,OPERN FROM FERRAM"
  cnn.Open cCONTMP
  VFPSetValues cnn

 'Abre o Recordset
  rst.CursorLocation = adUseClient
  rst.Open cSQLTMP, cnn, adOpenKeyset, adLockOptimistic, adCmdText

  cARQPF = GeracArq(PegPath("PATH", "PF"))
  oCON.Open cARQPF
  sSQL = "SELECT A.CODIGO,A.CODFINAL,A.CODFISCAL,B.SEQ,B.SSQ,A.PF,B.DESCRI,B.FERRAMEN,B.FERRAME2,B.FERRAME3,B.FERRAME4 FROM PF AS A,PFS AS B WHERE A.PF=B.PF "
  oRS.Open sSQL, oCON, adOpenDynamic, adLockOptimistic
  While Not oRS.EOF
    cCODIGO = FixStr(oRS("CODFISCAL"), " ", "TRIM")
    If Len(cCODIGO) = 0 Then
      cCODIGO = FixStr(oRS("CODFINAL"), " ", "TRIM")
    End If
    If Len(cCODIGO) = 0 Then
      cCODIGO = FixStr(oRS("CODIGO"), " ", "TRIM")
    End If
    LblDiz.Caption = cCODIGO
    LblDiz.Refresh
    For x = 1 To 4
      Select Case x
      Case 1
        cFERRAM = FixStr(oRS("FERRAMEN"))
      Case 2
        cFERRAM = FixStr(oRS("FERRAME2"))
      Case 3
        cFERRAM = FixStr(oRS("FERRAME3"))
      Case 4
        cFERRAM = FixStr(oRS("FERRAME4"))
      End Select
      If Len(cFERRAM) > 0 Then
        cCRITERIO = "FERRAM='" & cFERRAM & "'"
        rst.MoveFirst
        rst.Find cCRITERIO, 0, adSearchForward
        If Not rst.EOF Then
            If FixInt(rst("SEQ")) = 0 Then
                rst("SEQ") = FixInt(oRS("SEQ"))
                rst("SSQ") = FixInt(oRS("SSQ"))
            End If
            
            If FixInt(rst("PF")) = 0 Then
                rst("PF") = FixInt(oRS("PF"))
            End If
            
            If Len(Trim(FixStr(rst("PECA")))) = 0 Then
                rst("PECA") = FixStr(oRS("PECA"))
            End If
            
            If Len(Trim(FixStr(rst("OPERN")))) = 0 Then
                rst("OPERN") = FixStr(oRS("OPERN"))
            End If
           rst.Update
        End If
        
        
      End If
    Next x
    oRS.MoveNext
  Wend
  oRS.Close
  oCON.Close
  rst.Close
  cnn.Close

  Alert "Transferencia Concluida"

End Sub

