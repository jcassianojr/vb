VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form escMP05 
   Caption         =   "Selecione o Item  desejado:"
   ClientHeight    =   6090
   ClientLeft      =   2850
   ClientTop       =   1605
   ClientWidth     =   8880
   Icon            =   "Escmp05.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6090
      Left            =   7410
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
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   9551
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
Attribute VB_Name = "escMP05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aORDEM As Variant
Dim aORDES As Variant
Dim cORDEM As String
Dim cSUBWHERE As String

Private Sub CmdEscolher()
  Grid.Col = 1                                 '* Responsavel
  eRETU01 = Grid
  Grid.Col = 2                                 '* Cargo
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
  Dim cSQL As String
  cARQ = zMANA5EMP
  If Len(cSUBWHERE) > 0 Then
    cSQL = "SELECT DESCRI,RESPON,CARGO FROM MP05 WHERE " & cSUBWHERE & " ORDER BY " & cORDEM
  Else
    cSQL = "SELECT DESCRI,RESPON,CARGO FROM MP05 ORDER BY " & cORDEM
  End If
  cARQ = GeraConn(cARQ, "JETFOX")
  MontaGridFast Grid, 3, Array(3000, 3000, 1500), Array("Descrição", "Responsável", "Cargo"), _
                Array("L$DESCRI", "RESPON", "CARGO"), cARQ, cSQL
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  aORDEM = Array("DESCRI", "RESPON", "CARGO")
  aORDES = Array("Descrição", "Responsável", "Cargo")
  cORDEM = "DESCRI"
  cSUBWHERE = ""
  FilRelat
  xmontatoolbar Me.Toolbar1, "escMP05", True
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
  sButton = Button
  sButton = Left(UCase(Replace(sButton, "&", "")), 3)

  If Not AcessaBtnOld("escMP05", Button.Index) Then
    Exit Sub
  End If
  GravaLog 0, Button.Index, sButton, "escMP05"
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
      aARQUIVOS = Array(zMANA5EMP)
      ''Posicao 12 Nome da Tabela
      ''Posicao 13 Nome da Tabela
      aRELCFG = Array("", "", "", 0, False, _
                      False, "", "", "", "", "", _
                      False, "MP05", "MP05", "", "")
      FrmFiltro.Show vbModal, Me
      If lRETU Then
        cSUBWHERE = Replace(Replace(eRETU01, "{", ""), "}", "")
      End If
    Else
      ePASS01 = aORDES
      frmLocalizaa.Show vbModal, Me
      If lRETU Then
        cSUBWHERE = MontaFiltro(aORDEM, Array("L%", "L%", "L%"), eRETU01, eRETU02)
      End If
    End If
    FilRelat
  Case "IMP"
    cTIPO = "R"
    zgrp = "MP"
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


