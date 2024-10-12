VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Object = "{2DA70529-3366-414A-B408-46083BCD481B}#1.8#0"; "VBFLXGRD18.OCX"
Begin VB.Form escMD02 
   Caption         =   "Selecione:"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8760
   Icon            =   "Escmd02.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6090
      Left            =   7290
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
   Begin VBFLXGRD18.VBFlexGrid Grid 
      Height          =   5475
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   9657
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
Attribute VB_Name = "escMD02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aORDEM As Variant
Dim aORDES As Variant
Dim cORDEM As String
Dim cSUBWHERE As String
Dim cARQMD As String

Private Sub CmdEscolher()
  lRETU = True
  Grid.Col = 0
  eRETU01 = Grid
  Grid.Col = 1
  eRETU02 = Grid
  Grid.Col = 2
  eRETU03 = Grid
  Unload Me
End Sub

Private Sub cmdSair()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub FilRelat()
  Dim cSQL As String
  If Len(cSUBWHERE) = 0 Then
    cSQL = "SELECT CODIGO1,DESCRICAO,VALOR FROM MD02 WHERE CODIGO='" & tipocodigo() & "' ORDER BY " & cORDEM
  Else
    cSQL = "SELECT CODIGO1,DESCRICAO,VALOR FROM MD02 WHERE CODIGO='" & tipocodigo() & "' AND " & cSUBWHERE & " ORDER BY " & cORDEM
  End If
  MontaGridFast Grid, 3, Array(1200, 3000, 800), Array("Código", "Descrição", "Valor"), _
                Array("L$CODIGO1", "DESCRICAO", "VALOR"), cARQMD, cSQL
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  cARQMD = zMANA5EMP
  cARQMD = GeraConn(cARQMD, "JETFOX")
  aORDEM = Array("CODIGO1", "DESCRICAO")
  aORDES = Array("CODIGO1", "Descricao")
  cORDEM = "CODIGO1"
  cSUBWHERE = ""
  xmontatoolbar Me.Toolbar1, "escMD02", True
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
  Dim sButton As String
  Dim cARQTMP As String
  Dim sSQL As String
  Dim cDESCRICAO As String
  sButton = Button
  sButton = Left(UCase(Replace(sButton, "&", "")), 3)


  lRETU = False
  If Not AcessaBtnOld("escMD02", Button.Index) Then
    Exit Sub
  End If

  GravaLog 0, Button.Index, sButton, "escMD02"

  Select Case sButton
  Case "EXC", "DEL"
    Grid.Col = 0
    cCONJUNTO = Trim(Grid)
    cARQTMP = GeraConn(zMANA5EMP, "JETFOX")
    sSQL = "select codigo from MD02 WHERE =" & PadRight(tipocodigo(), 12) & cCONJUNTO
    If ApagaSQLP(cARQTMP, sSQL) Then
      FilRelat
    End If
  Case "NOV"
    cCONJUNTO = InputBox("Digite o Codigo", "Inclusão ", "__")
    cDESCRICAO = InputBox("Digite a descricao", "Inclusão ", "__")
    If Len(cCONJUNTO) = 0 Or Len(cDESCRICAO) = 0 Then
      Alert ("Preencha Codigo e Descricao")
      Exit Sub
    End If
    cARQTMP = GeraConn(zMANA5EMP, "JETFOX")
    sSQL = "select codigo from MD02 WHERE CODIGO='" & tipocodigo() & "' AND CODIGO1='" & cCONJUNTO & "'"
    IncluiSQL cARQTMP, sSQL, 3, Array("CODIGO", "CODIGO1", "DESCRICAO"), _
              Array(tipocodigo(), cCONJUNTO, cDESCRICAO), True, True
    FilRelat
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
      aARQUIVOS = Array(cARQMD)
      ''Posicao 12 Nome da Tabela
      ''Posicao 13 Nome da Tabela
      aRELCFG = Array("", "", "", 0, False, _
                      False, "", "", "", "", "", _
                      False, "MD02", "MD02", "", "")
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
    zgrp = "MD"
    escRPT.Show vbModal, Me
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

Private Function tipocodigo()
  tipocodigo = ""
  Select Case iMD02
  Case 1
    tipocodigo = "TIPINS"
  Case 2
    tipocodigo = "SETINS"
  Case 3
    tipocodigo = "CRRS02"
  Case 4
    tipocodigo = "SIMNAO"
  Case 5
    tipocodigo = "MONTAGEM"
  Case 6
    tipocodigo = "MARINS"
  Case 7
    tipocodigo = "APLINS"
  Case 8
    tipocodigo = "GRUPOMAQ"
  Case 9
    tipocodigo = "TIPOEQI"
  End Select
End Function


