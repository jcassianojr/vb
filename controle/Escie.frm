VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form escIE 
   Caption         =   "Selecione o Processo  desejado:"
   ClientHeight    =   5775
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11010
   Icon            =   "Escie.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   11010
   StartUpPosition =   2  'CenterScreen
   Begin vbExtra.FlexFn FlexFn1 
      Height          =   405
      Left            =   120
      TabIndex        =   2
      Top             =   120
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
      Left            =   9540
      TabIndex        =   1
      Top             =   0
      Width           =   1470
      _ExtentX        =   2593
      _ExtentY        =   10186
      ButtonWidth     =   609
      ButtonHeight    =   953
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin VBFLXGRD18.VBFlexGrid Grid 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   9128
      _Version        =   393216
   End
End
Attribute VB_Name = "escIE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aORDEM As Variant
Dim aORDES As Variant
Dim cORDEM As String
Dim cSUBWHERE As String
Dim cARQPF As String
Dim cARQIE As String

Private Sub CmdSair_Click()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub EditPf_Click()
  Grid.Col = 0
  nPF = Grid
  frmIE.Show vbModal
End Sub

Private Sub FilRelat()
  Dim cSQL
  If Len(cSUBWHERE) = 0 Then
    cSQL = "SELECT CODIGOINT,PF,codigo,descr,resdat,RESNUM,BLOQUEADO FROM PF WHERE RESNUM>0 AND NOT BLOQUEADO ORDER BY " & cORDEM
  Else
    cSQL = "SELECT CODIGOINT,PF,codigo,descr,resdat,RESNUM,BLOQUEADO FROM PF WHERE RESNUM>0 AND NOT BLOQUEADO AND " & cSUBWHERE & " ORDER BY " & cORDEM
  End If
  MontaGridFast Grid, 5, Array(600, 1600, 4000, 1000, 1600), Array("PF", "Produto", "Descricao", "Liberada", "Codigo INT"), _
                Array("PF", "L$CODIGO", "L$DESCR", "RESDAT", "L$CODIGOINT"), cARQPF, cSQL
End Sub

Private Sub Form_Load()
  cARQIE = PegPath("PATH", "IE")
  cARQPF = PegPath("PATH", "PF")
  CenterFormToScreen Me
  aORDEM = Array("PF", "CODIGO", "DESCR", "RESDAT", "CODIGOINT")
  aORDES = Array("PF", "Codigo", "Descricao", "Data Liberacao", "Codigo Int")
  cORDEM = "PF"
  cSUBWHERE = ""
  xmontatoolbar Me.Toolbar1, "escIE", True
  FilRelat
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
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
  sButton = Button
  sButton = Left(UCase(Replace(sButton, "&", "")), 3)
  If Not AcessaBtnOld("escIE", Button.Index) Then
    Exit Sub
  End If
  GravaLog 0, Button.Index, sButton, "escIE"
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
      aARQUIVOS = Array(cARQPF)
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
      ePASS01 = aORDES
      frmLocalizaa.Show vbModal, Me
      If lRETU Then
        cSUBWHERE = MontaFiltro(aORDEM, Array("=", "L%", "L%", "L%"), eRETU01, eRETU02)
      End If
    End If
    FilRelat
  Case "EDI"
    EditPf_Click
  Case "COP"
    cARQIMP = "IE"
    FrmImp2.Show vbModal, Me
  Case "LOC"
    ePASS01 = aORDES
    frmLocalizaa.Show vbModal, Me
    If lRETU Then
      LocalizaGrid Grid, eRETU01, eRETU02, , 1
    End If
  Case "IMP"
    cTIPO = "R"
    zgrp = "PF"
    ZGRPSUB = "IE"
    eLOCALIZA = ""
    escRPT.Show vbModal, Me
  Case "IEB"
    cTIPO = "R"
    zgrp = "PF"
    ZGRPSUB = "IE"
    eLOCALIZA = "ITA00148"
    escRPT.Show vbModal, Me
  Case "SAI"
    CmdSair_Click
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub


