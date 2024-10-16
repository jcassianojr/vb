VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Object = "{075212A8-C1CF-444E-939D-F6046CCDBC08}#1.0#0"; "VBFLXGRD18.OCX"
Begin VB.Form escpc 
   Caption         =   "Selecione o pc  desejado:"
   ClientHeight    =   5775
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11445
   Icon            =   "escpc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   11445
   StartUpPosition =   2  'CenterScreen
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   5775
      Left            =   10035
      TabIndex        =   1
      Top             =   0
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   10186
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin VBFLXGRD18.VBFlexGrid Grid 
      Height          =   5115
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   9022
   End
End
Attribute VB_Name = "escpc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cARQPC As String
Dim aORDEM As Variant
Dim aORDES As Variant
Dim cORDEM As String
Dim cSUBWHERE As String

Private Sub CmdLiberar_Click()
  Dim nTMPELA As Integer
  Grid.Col = 0
  nPF = Grid
  nTMPELA = FixInt(PegCampoSQLADO(Sdb, "SELECT PCELANUM FROM PF WHERE PF=" & nPF, "", 0))

  If nTMPELA = zIDFOLHA Then
    Alert ("Voce ja e o Elaborador")
    Exit Sub
  End If
  ePASS01 = "PC"
  frmPFLIB.Show vbModal, Me
  FilRelat
End Sub

Private Sub CmdSair_Click()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub EditPf_Click()
  Grid.Col = 0
  nPF = Grid
  frmPCX.Show vbModal
End Sub

Private Sub FilRelat()
  Dim cARQ As String
  Dim cSQL As String
  cARQ = PegPath("PATH", "PF")
  If Len(cSUBWHERE) > 0 Then
    cSQL = "SELECT codigoint,pclibnum,pclibdate,RESNUM,FEMEAF,bloqueado,PF,STPFPC,STFEPC,codigo,DEscr,resdat,FEMEAD FROM PF " & " WHERE RESNUM>0 AND FEMEAF>0 AND NOT BLOQUEADO AND  " & cSUBWHERE & " ORDER BY " & cORDEM
  Else
    cSQL = "SELECT codigoint,pclibnum,pclibdate,RESNUM,FEMEAF,bloqueado,PF,STPFPC,STFEPC,codigo,DEscr,resdat,FEMEAD FROM PF " & " WHERE RESNUM>0 AND FEMEAF>0 AND NOT BLOQUEADO ORDER BY " & cORDEM
  End If
  MontaGrid Grid, 9, Array(600, 300, 300, 1600, 2800, 1000, 1000, 1000, 1600), Array("PF", "PF", "PC", "Codigo", "Descricao", "PF", "Femea", "PC Lib", "Cod Int"), Array("PF", "STPFPC", "STFEPC", "L$CODIGO", "DESCR", "C$RESDAT", "C$FEMEAD", "C$NPD,PCLIBNUM,PCLIBDATE", "L$CODIGOINT"), cARQ, cSQL
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  cARQPC = PegPath("PATH", "PF")
  aORDEM = Array("PF", "STPFPC", "STFEPC", "CODIGO", "DESCR", "RESDAT", "FEMEAD", "PCLIBDATE", "CODIGOINT")
  aORDES = Array("Nº PF", "Alterado PC", "Alterado PF", "Codigo", "Descricao", "Data liberacao Processo fabricacao", "data liberacao femea", "data liberacao plano de controle", "Codigo Integrado")
  cORDEM = "PF"
  cSUBWHERE = ""

  xmontatoolbar Me.Toolbar1, "escPC", True
  FilRelat
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
  If KeyAscii > 31 And KeyAscii < 123 Then
    LocalizaGrid Grid, Chr(KeyAscii), 1, False
  End If

End Sub

'Private Sub Grid_SelChange()'

' With Grid
'    If .Rows > 2 Then
'        .Col = .Cols - 1
'        .ColSel = 0
'        .TopRow = .Row
'    End If
'End With

'End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim sButton As String
  Dim aORDLIB As Variant
  Dim aDESLIB As Variant
  sButton = Button

  sButton = Left(UCase(Replace(sButton, "&", "")), 3)
  If Not AcessaBtnOld("escPC", Button.Index) Then
    Exit Sub
  End If
  GravaLog 0, Button.Index, sButton, "escPC"
  If sButton = "LIB" Then    'Chama cmd_liberar abaixo aqui atribui  sdb  para a base PF/PFP/PFG 'Or sButton = "TRO" troca agora edicao pf
    aORDLIB = Array("PF", "PFP", "PFG")      'lib e troca sao diferentes
    aDESLIB = Array("Normal", "Preliminar", "GP12")
    ePASS01 = aDESLIB
    escOrdem.Show vbModal, Me
    If lRETU Then
      ePASS01 = aORDLIB(eRETU01)
      Sdb = PegPath("PATH", aORDLIB(eRETU01))
    End If
  End If
  Select Case sButton

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
      aARQUIVOS = Array(cARQPC)
      ''Posicao 12 Nome da Tabela
      ''Posicao 13 Nome da Tabela
      aRELCFG = Array("", "", "", 0, False, _
                      False, "", "", "", "", "", _
                      False, "PF", "PF", "", "")
      FrmFiltro.Show vbModal, Me
      If lRETU Then
        cSUBWHERE = Replace(Replace(eRETU01, "{", ""), "}", "")
      End If
    Else
      ''frmLocalizar.Show vbModal, Me
      ePASS01 = aORDES
      frmLocalizaa.Show vbModal, Me
      If lRETU Then
        cSUBWHERE = MontaFiltro(aORDEM, Array("=", "B", "B", "L%", "L%", "L%", "L%", "L%"), eRETU01, eRETU02)
      End If
    End If
    FilRelat
  Case "IMP"
    cTIPO = "R"
    zgrp = "PC"
    escRPT.Show vbModal, Me
  Case "EDI"
    EditPf_Click
  Case "LOC"
    ePASS01 = aORDES
    frmLocalizaa.Show vbModal, Me
    If lRETU Then
      LocalizaGrid Grid, eRETU01, eRETU02, , 1
    End If

  Case "LIB"
    CmdLiberar_Click
  Case "TRO"
    Alert "Use a opcao pela edicao do PF"
    'Trocapf_Click
  Case "SAI"
    CmdSair_Click
  End Select

End Sub
'Private Sub Trocapf_Click()
'If Not AcessaMenu("frmPFTROCA", 1) Then
'       Exit Sub
'   End If
'    Grid.Col = 0
'    nPF = Grid
'    FrmPFT.Show vbModal
'End Sub
Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub


