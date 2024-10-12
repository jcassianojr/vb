VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Object = "{2DA70529-3366-414A-B408-46083BCD481B}#1.8#0"; "VBFLXGRD18.OCX"
Begin VB.Form escmu01 
   Caption         =   "Selecione o Item Desejado:"
   ClientHeight    =   6090
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9690
   Icon            =   "Escmu01.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   9690
   StartUpPosition =   2  'CenterScreen
   Begin VBFLXGRD18.VBFlexGrid Grid 
      Height          =   5415
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   9551
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6090
      Left            =   8220
      TabIndex        =   1
      Top             =   0
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   10742
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin vbExtra.FlexFn FlexFn1 
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   714
      PageNumbersFormat=   ""
      PageNumbersFormatIndex=   0
      BeginProperty PageNumbersFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "escmu01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aORDEM As Variant
Dim aORDES As Variant
Dim cORDEM As String
Dim cSUBWHERE As String
Dim cESCUSO As String
Dim cTABUSO As String

Private Sub CmdEscolher()
  Grid.Col = 0
  eRETU01 = Grid
  Grid.Col = 1
  eRETU02 = Grid
  lRETU = True
  Unload Me
End Sub

Private Sub cmdSair()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub FilRelat()
  On Error Resume Next
  Dim cARQ As String
  Dim cSQL As String
  cARQ = GeraConn(zMANA5EMP, "JETFOX")
  'cSQL = "SELECT CODIGO, NOME,NOM2 FROM " & cTABUSO
  cSQL = "SELECT CODIGO, NOME + NOM2 AS NOME FROM " & cTABUSO

  If Len(cSUBWHERE) > 0 Then
    cSQL = cSQL & " WHERE " & cSUBWHERE
  End If

  cSQL = cSQL & " ORDER BY " & cORDEM
  '     MontaGridFast Grid, 3, Array(2000, 3500, 3500), Array("Codigo", "Nome", ""), _
        Array("L$CODIGO", "L$NOME", "L$NOM2"), cARQ, cSQL
  MontaGridFast Grid, 2, Array(1500, 6000), Array("Codigo", "Nome", ""), _
                Array("L$CODIGO", "L$NOME"), cARQ, cSQL

End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  aORDEM = Array("CODIGO", "NOME")
  aORDES = Array("CODIGO", "NOME")
  cORDEM = "CODIGO"
  cSUBWHERE = ""


  eRETU01 = ""
  eRETU02 = ""
  lRETU = False

  If iMU01 = 1 Then cESCUSO = "escMU01"
  If iMU01 = 2 Then cESCUSO = "escMT01"
  If iMU01 = 3 Then cESCUSO = "escMW05"
  If iMU01 = 4 Then cESCUSO = "escMR01"
  If iMU01 = 5 Then cESCUSO = "escMP03"
  If iMU01 = 6 Then cESCUSO = "escMW07"
  If iMU01 = 7 Then cESCUSO = "escMQ01"


  If iMU01 = 1 Then cTABUSO = "MU01"
  If iMU01 = 2 Then cTABUSO = "MT01"
  If iMU01 = 3 Then cTABUSO = "MW05"
  If iMU01 = 4 Then cTABUSO = "MR01"
  If iMU01 = 5 Then cTABUSO = "MP03"
  If iMU01 = 6 Then cTABUSO = "MW07"
  If iMU01 = 7 Then cTABUSO = "MQ01"

  xmontatoolbar Me.Toolbar1, cESCUSO, True


  FilRelat

End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    CmdEscolher
  End If
  If KeyAscii > 31 And KeyAscii < 123 Then
    LocalizaGrid Grid, Chr(KeyAscii), 1, False
  End If

End Sub

'Private Sub Grid_SelChange()
'    With Grid
'        If .Rows > 2 Then
'            .Col = .Cols - 1
'            .ColSel = 0
'            .TopRow = .Row
'        End If
'    End With
'End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim cARQTMP As String
  Dim sSQL As String
  Dim cCODINC As String
  Dim sButton As String
  sButton = Button
  sButton = Left(UCase(Replace(sButton, "&", "")), 3)
  If Not AcessaBtnOld(cESCUSO, Button.Index) Then
    Exit Sub
  End If


  GravaLog 0, Button.Index, sButton, cESCUSO

  Select Case sButton
  Case "NOV", "INC"
    cCODINC = InputBox("Digite o Codigo", "Inclusão", "__")
    cARQTMP = GeraConn(zMANA5EMP, "JETFOX")
    sSQL = "select codigo from " & cTABUSO & " WHERE CODIGO='" & cCODINC & "'"
    If IncluiSQL(cARQTMP, sSQL, 1, Array("CODIGO"), Array(cCODINC), True, True) Then
      ePASS01 = cCODINC
      If iMU01 = 1 Or iMU01 = 2 Or iMU01 = 4 Or iMU01 = 7 Then
        ePASS02 = cTABUSO
        FrmMtMu.Show vbModal, Me
      End If
      If iMU01 = 5 Then
        FrmMP03.Show vbModal, Me
      End If
      FilRelat
    End If
  Case "EDI", "ALT"
    Grid.Col = 0
    ePASS01 = Grid
    If iMU01 = 1 Or iMU01 = 2 Or iMU01 = 4 Or iMU01 = 7 Then
      ePASS02 = cTABUSO
      FrmMtMu.Show vbModal, Me
    End If
    If iMU01 = 5 Then
      FrmMP03.Show vbModal, Me
    End If
  Case "EXC", "DEL"
    cARQTMP = zMANA5EMP
    cARQTMP = GeraConn(cARQTMP, "JETFOX")
    Grid.Col = 0
    sSQL = "select codigo from " & cTABUSO & " WHERE CODIGO='" & Trim(Grid) & "'"
    If ApagaSQLP(cARQTMP, sSQL) Then
      FilRelat
    End If
  Case "ORD"
    ePASS01 = aORDES
    escOrdem.Show vbModal, Me
    If lRETU Then
      cORDEM = aORDEM(eRETU01)
      FilRelat
    End If
  Case "FIL"
    cSUBWHERE = ""
    If MDG("Usar Filtro Avancado") Then
      aARQUIVOS = Array(zMANA5EMP)
      ''Posicao 12 Nome da Tabela
      ''Posicao 13 Nome da Tabela
      aRELCFG = Array("", "", "", 0, False, _
                      False, "", "", "", "", "", _
                      False, cTABUSO, cTABUSO, "", "")
      FrmFiltro.Show vbModal, Me
      If lRETU Then
        cSUBWHERE = Replace(Replace(eRETU01, "{", ""), "}", "")
      End If
    Else
      ''frmLocalizar.Show vbModal, Me
      ePASS01 = aORDES
      frmLocalizaa.Show vbModal, Me
      If lRETU Then
        cSUBWHERE = MontaFiltro(aORDEM, Array("L%", "L%"), eRETU01, eRETU02)
      End If
    End If
    FilRelat

  Case "IMP"
    cTIPO = "R"
    Select Case iMU01
    Case 1
      zgrp = "MU"
    Case 2
      zgrp = "MT"
    Case 3, 6
      zgrp = "CO"
    Case 4, 6
      zgrp = "MR"
    Case 5
      zgrp = "MP"
    Case 7
      zgrp = "MQ"
    End Select
    escRPT.Show vbModal, Me
  Case "FOT"
    Grid.Col = 0
    zgrp = Grid
    Select Case iMU01
    Case 1
      iImage = 7
      cARQRTF = PegPath("PATH", "IMGMU01")
    Case 2
      iImage = 8
      cARQRTF = PegPath("PATH", "IMGMT01")
    Case 3
      iImage = 9
      cARQRTF = PegPath("PATH", "IMGMW05")
    Case 4
      iImage = 10
      cARQRTF = PegPath("PATH", "IMGMR01")
    Case 5
      Alert "Não Possui Imagens", "Opçao não Disponivel"
      Exit Sub
    Case 6
      iImage = 12
      cARQRTF = PegPath("PATH", "IMGMW07")
    Case 7
      iImage = 13
      cARQRTF = PegPath("PATH", "IMGMQ01")

    End Select
    Load frmIMAGENS
    frmIMAGENS.TXTFIELDS(0).Enabled = False
    frmIMAGENS.Escolher(0).Visible = False
    frmIMAGENS.Show vbModal, Me
  Case "ESC"
    CmdEscolher
  Case "LOC"
    ePASS01 = aORDES
    frmLocalizaa.Show vbModal, Me
    If lRETU Then
      LocalizaGrid Grid, eRETU01, eRETU02, , 1
    End If
  Case "SAI"
    cmdSair
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub


