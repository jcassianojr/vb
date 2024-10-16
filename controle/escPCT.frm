VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Object = "{075212A8-C1CF-444E-939D-F6046CCDBC08}#1.0#0"; "VBFLXGRD18.OCX"
Begin VB.Form escPCT 
   Caption         =   "Selecione Simbolo Desejado"
   ClientHeight    =   6090
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8880
   Icon            =   "escPCT.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VBFLXGRD18.VBFlexGrid Grid 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   9551
   End
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
Attribute VB_Name = "ESCPCT"
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
  Dim sSQL As String

  If Len(cSUBWHERE) = 0 Then
    sSQL = "SELECT IDTIPO,DESCRICAO FROM PCTIPO ORDER BY " & cORDEM
  Else
    sSQL = "SELECT IDTIPO,DESCRICAO FROM PCTIPO WHERE " & cSUBWHERE & " ORDER BY " & cORDEM
  End If
  MontaGridFast Grid, 2, Array(800, 3000), Array("Nº", "Descricao"), _
                Array("IDtipo", "Descricao"), cARQPF, sSQL
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  aORDEM = Array("IDtipo", "Descricao")
  aORDES = Array("Nº", "Descricao")
  cORDEM = "IDTIPO"
  cSUBWHERE = ""

  cARQPF = PegPath("PATH", "PF")
  xmontatoolbar Me.Toolbar1, "escPCT", True
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

  Dim sButton As String
  sButton = Button
  sButton = Left(UCase(Replace(sButton, "&", "")), 3)
  If Not AcessaBtnOld("escPCT", Button.Index) Then
    Exit Sub
  End If

  GravaLog 0, Button.Index, sButton, "escPCT"


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
                      False, "PCTIPO", "PCTIPO", "", "")
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


