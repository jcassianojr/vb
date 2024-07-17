VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form escPFD 
   Caption         =   "Selecione Simbolo Desejado"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8190
   Icon            =   "escpfd.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   9763
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6090
      Left            =   6720
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
      Left            =   240
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
Attribute VB_Name = "escPFD"
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

Private Sub CmdSair_Click()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub FilRelat()
  Dim sSQL As String
  If Len(Zsetor) = 0 Then
    If Len(cSUBWHERE) = 0 Then
      sSQL = "SELECT cqtde,freq,ID FROM PFD ORDER BY " & cORDEM
    Else
      sSQL = "SELECT cqtde,freq,ID FROM PFD WHERE " & cSUBWHERE & " ORDER BY " & cORDEM
    End If
  Else
    If Len(cSUBWHERE) = 0 Then
      sSQL = "SELECT cqtde,freq,ID FROM PFD WHERE SETOR='" & Zsetor & "'" & " ORDER BY " & cORDEM
    Else
      sSQL = "SELECT cqtde,freq,ID FROM PFD WHERE SETOR='" & Zsetor & "' AND " & cSUBWHERE & " ORDER BY " & cORDEM
    End If
  End If
  MontaGridFast Grid, 3, Array(3000, 3000, 0), Array("Quantidade", "Frequencia", "id"), _
                Array("L$Cqtde", "L$freq", "id"), cARQPF, sSQL
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  aORDEM = Array("Quantidade", "Frequencia")
  aORDES = Array("Cqtde", "freq")
  cORDEM = "CQTDE"
  cSUBWHERE = ""
  cARQPF = PegPath("PATH", "PF")
  xmontatoolbar Me.Toolbar1, "escPFD", True
  FilRelat
End Sub

Private Sub Grid_DblClick()
  Grid.Col = 0
  eRETU01 = Grid
  Grid.Col = 1
  eRETU02 = Grid
  lRETU = True
  Unload Me
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Grid_DblClick                            ''escolher
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
  Dim cSQL As String
  sButton = Button
  sButton = Left(UCase(Replace(sButton, "&", "")), 3)

  If Not AcessaBtnOld("escPFD", Button.Index) Then
    Exit Sub
  End If

  GravaLog 0, Button.Index, sButton, "escPFD"


  Select Case sButton
  Case "NOV"
    If IncluiSQL(cARQPF, "PFD", 1, Array("SETOR"), Array(Zsetor), False, False, Array("ID")) Then
      ePASS01 = eRETU01(0)
      frmPFD.Show vbModal, Me
      FilRelat
    End If
  Case "EXC"
    If Grid.Row > 0 Then
      Grid.Col = 2
      cSQL = "SELECT * FROM PFD WHERE ID=" & Grid
      If ApagaSQLP(cARQPF, cSQL) Then
        FilRelat
      End If
    End If
  Case "EDI"
    If Grid.Row > 0 Then
      Grid.Col = 2
      ePASS01 = Grid
      frmPFD.Show vbModal, Me
    End If
  Case "ESC"
    Grid_DblClick
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
                      False, "PFD", "PFD", "", "")
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
  Case "LOC"
    ePASS01 = aORDES
    frmLocalizaa.Show vbModal, Me
    If lRETU Then
      LocalizaGrid Grid, eRETU01, eRETU02, , 1
    End If
  Case "SAI"
    CmdSair_Click

  End Select


End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault

End Sub


