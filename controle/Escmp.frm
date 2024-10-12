VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Object = "{2DA70529-3366-414A-B408-46083BCD481B}#1.8#0"; "VBFLXGRD18.OCX"
Begin VB.Form escmp 
   Caption         =   "Selecione o  desejado:"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8970
   Icon            =   "Escmp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   8970
   StartUpPosition =   2  'CenterScreen
   Begin VBFLXGRD18.VBFlexGrid Grid 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   9763
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6090
      Left            =   7500
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
Attribute VB_Name = "escmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cESCUSO As String
Dim aORDEM As Variant
Dim aORDES As Variant
Dim cORDEM As String
Dim cSUBWHERE As String
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
  Dim cARQ As String
  Dim sSQL As String
  cARQ = GeraConn(zMANA5EMP, "JETFOX")
  sSQL = "SELECT CODIGO,NOME FROM " & cTABUSO
  If iMU01 = 8 Then
    cARQ = PegPath("PATH", "LOGIXODBC")
    sSQL = "SELECT cod_cent_trab AS codigo,den_cent_trab AS NOME FROM " & cTABUSO
  End If
  If iMU01 = 9 Then
    cARQ = PegPath("PATH", "LOGIXODBC")
    sSQL = "SELECT cod_recur AS codigo,den_recur AS NOME FROM " & cTABUSO
  End If
  If iMU01 = 10 Then
    cARQ = PegPath("PATH", "LOGIXODBC")
    sSQL = "SELECT cod_recur AS codigo,den_recur AS NOME FROM " & cTABUSO
  End If
  If iMU01 = 11 Then
    cARQ = PegPath("PATH", "LOGIXODBC")
    sSQL = " SELECT cod_cOMPON AS codigo,trim(fixstr(des_compon)) AS nome fROM equipamentos_970 WHERE val_logico='S' "
    sSQL = sSQL & " AND cod_grupo_equip<>'110'"
    sSQL = sSQL & " AND cod_grupo_equip<>'310'"
    sSQL = sSQL & " AND cod_grupo_equip<>'410'"
    sSQL = sSQL & " AND cod_grupo_equip<>'710'"
    sSQL = sSQL & " AND cod_grupo_equip<>'416'"
    sSQL = sSQL & " AND cod_grupo_equip<>'426'"
    sSQL = sSQL & " AND cod_grupo_equip<>'TI'"
  End If
  If iMU01 = 12 Then                           'grupo 410 dispositivos
    cARQ = PegPath("PATH", "LOGIXODBC")
    sSQL = " SELECT cod_cOMPON AS codigo,trim(fixstr(des_compon)) AS nome fROM equipamentos_970 WHERE val_logico='S' "
    sSQL = sSQL & " AND cod_grupo_equip='410'"
  End If



  If Len(cSUBWHERE) > 0 Then
    sSQL = sSQL & " WHERE " & cSUBWHERE
  End If
  sSQL = sSQL & " ORDER BY " & cORDEM

  MontaGrid Grid, 2, Array(1800, 4800), Array("Código", "Descrição"), _
            Array("L$CODIGO", "L$NOME"), cARQ, sSQL

End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  aORDEM = Array("CODIGO", "NOME")
  aORDES = Array("CODIGO", "NOME")
  cORDEM = "CODIGO"
  cSUBWHERE = ""
  Select Case iMU01                            '1,2,4,7 livres, 3, 5, 6
  Case 1
    cESCUSO = "escMP01"
    cTABUSO = "MP01"
  Case 2
    cESCUSO = "escMP02"
    cTABUSO = "MP02"
  Case 4
    cESCUSO = "escETI"
    cTABUSO = "ETI"
  Case 7
    cESCUSO = "escMP03"
    cTABUSO = "MP03"
  Case 8
    cESCUSO = "escArranjo"
    cTABUSO = "cent_trabalho_970"
    cORDEM = "cod_cent_trab"
  Case 9
    cESCUSO = "escRecMaq"
    cTABUSO = "recurso_970_maquina"
    cORDEM = "cod_recur"
  Case 10
    cESCUSO = "escRecHom"
    cTABUSO = "recurso_970_homem"
    cORDEM = "cod_recur"
  Case 11                                      '
    cESCUSO = "escEquip"
    cTABUSO = "equipamento"
    cORDEM = "1"
  Case 12                                      '
    cESCUSO = "escLXDispo"
    cTABUSO = "equipamento"
    cORDEM = "1"
  End Select
  xmontatoolbar Me.Toolbar1, cESCUSO, True
  FilRelat
  lRETU = False
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
  Dim cCODINC As String
  Dim cNOMEINC As String
  Dim cARQTMP As String
  Dim sButton As String
  Dim sSQL As String
  sButton = Button
  sButton = Left(UCase(Replace(sButton, "&", "")), 3)

  If Not AcessaBtnOld(cESCUSO, Button.Index) Then
    Exit Sub
  End If

  GravaLog 0, Button.Index, sButton, "escMP"

  Select Case sButton

  Case "NOV", "INC"
    If cTABUSO = "ETI" Or cTABUSO = "MP03" Or cTABUSO = "MP01" Or cTABUSO = "MP02" Then
      cCODINC = InputBox("Digite o Codigo", "Codigo para inclusão", "__")
      cNOMEINC = InputBox("Digite a Descricao", "Descricao do codigo", "__")
      cARQTMP = GeraConn(zMANA5EMP, "JETFOX")
      If cTABUSO = "ETI" Then
        sSQL = "select codigo,nome from ETI WHERE CODIGO='" & cCODINC & "'"
      End If
      If cTABUSO = "MP01" Then
        sSQL = "select codigo,nome from MP01 WHERE CODIGO='" & cCODINC & "'"
      End If
      If cTABUSO = "MP02" Then
        sSQL = "select codigo,nome from MP02 WHERE CODIGO='" & cCODINC & "'"
      End If
      If cTABUSO = "MP03" Then
        sSQL = "select codigo,nome from MP03 WHERE CODIGO='" & cCODINC & "'"
      End If
      If IncluiSQL(cARQTMP, sSQL, 2, Array("CODIGO", "NOME"), Array(cCODINC, cNOMEINC), True, True) Then
        Grid.Col = 0
        If cTABUSO = "ETI" Then
          ePASS01 = cCODINC
          frmETI.Show vbModal, Me
        End If
        If cTABUSO = "MP03" Then
          ePASS01 = cCODINC
          FrmMP03.Show vbModal, Me
        End If
        FilRelat
      End If
    End If
  Case "EXC"
    If cTABUSO = "ETI" Or cTABUSO = "MP03" Or cTABUSO = "MP01" Or cTABUSO = "MP02" Then
      cARQTMP = GeraConn(zMANA5EMP, "JETFOX")
      Grid.Col = 0
      sSQL = "select codigo from " & cTABUSO & " WHERE CODIGO='" & Trim(Grid) & "'"

      If ApagaSQLP(cARQTMP, sSQL) Then
        FilRelat
      End If
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
      ePASS01 = aORDES
      frmLocalizaa.Show vbModal, Me
      If lRETU Then
        cSUBWHERE = MontaFiltro(aORDEM, Array("L%", "L%"), eRETU01, eRETU02)
      End If
    End If
    FilRelat
  Case "IMP"
    cTIPO = "R"
    zgrp = "MP"
    If cTABUSO = "ETI" Then
      zgrp = "ETI"
    End If
    If cTABUSO = "MP01" Then
      zgrp = "MP1"
    End If
    escRPT.Show vbModal, Me
  Case "EDI", "ALT"
    If cTABUSO = "ETI" Then
      Grid.Col = 0
      ePASS01 = Grid
      frmETI.Show vbModal, Me
    End If
    If cTABUSO = "MP03" Then
      Grid.Col = 0
      ePASS01 = Grid
      FrmMP03.Show vbModal, Me
    End If


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


