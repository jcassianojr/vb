VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form escfrmpre 
   Caption         =   "Selecione o Item Desejado"
   ClientHeight    =   6090
   ClientLeft      =   2850
   ClientTop       =   1605
   ClientWidth     =   8715
   Icon            =   "escfrmpre.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   8715
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6090
      Left            =   7230
      TabIndex        =   1
      Top             =   0
      Width           =   1485
      _ExtentX        =   2619
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
Attribute VB_Name = "escfrmpre"
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
Dim cARQFEMEA As String

Private Sub CmdEscolher()
  Grid.Col = 0
  eRETU01 = Grid
  lRETU = True
  Unload Me
End Sub

Private Sub CmdExcluir()
  Dim sSQL As String
  If Grid.Row > 0 Then
    Grid.Col = 0
    nSEQ = FixInt(Grid)
    sSQL = "SELECT * FROM FEMPRE WHERE SEGGRA=" & nSEQ
    If ApagaSQLP(cARQFEMEA, sSQL, "Apagar Pre Cadastro") Then
      FilRelat
    End If
  End If
End Sub

Private Sub cmdSair()
  lRETU = False
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub FilRelat()
'Dim carq As String
  Dim cSQL As String
  If Len(cSUBWHERE) = 0 Then
    cSQL = "SELECT SEGGRA,TITULO FROM FEMPRE ORDER BY " & cORDEM
  Else
    cSQL = "SELECT SEGGRA,TITULO FROM FEMPRE WHERE " & cSUBWHERE & " ORDER BY " & cORDEM
  End If
  MontaGridFast Grid, 2, Array(400, 4000), Array("Código", "Descrição"), _
                Array("SEGGRA", "L$TITULO"), cARQFEMEA, cSQL
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  cARQPF = PegPath("PATH", "PF")
  cARQFEMEA = PegPath("PATH", "FEMEA")
  aORDEM = Array("SEGGRA", "TITULO")
  aORDES = Array("Código", "Descrição")
  cORDEM = "TITULO"
  cSUBWHERE = ""
  xmontatoolbar Me.Toolbar1, "escfrmpre", True
  FilRelat
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
  If KeyAscii > 31 And KeyAscii < 123 Then
    LocalizaGrid Grid, Chr(KeyAscii), 1, False
  End If
  If KeyAscii = 13 Then
    CmdEscolher
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
  If Not AcessaBtnOld("escFRMPRE", Button.Index) Then
    Exit Sub
  End If

  GravaLog 0, Button.Index, sButton, "escFRMPRE"

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
                      False, "FEMPRE", "FEMPRE", "", "")
      FrmFiltro.Show vbModal, Me
      If lRETU Then
        cSUBWHERE = Replace(Replace(eRETU01, "{", ""), "}", "")
      End If
    Else
      ePASS01 = aORDES
      frmLocalizaa.Show vbModal, Me
      If lRETU Then
        cSUBWHERE = MontaFiltro(aORDEM, Array("=", "L%"), eRETU01, eRETU02)
      End If
    End If
    FilRelat
  Case "EXC"
    CmdExcluir

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


