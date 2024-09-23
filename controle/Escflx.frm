VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Object = "{2DA70529-3366-414A-B408-46083BCD481B}#1.8#0"; "VBFLXGRD17.OCX"
Begin VB.Form escFLX 
   Caption         =   "Selecione Simbolo Desejado"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8190
   Icon            =   "Escflx.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   8190
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFields 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   24
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Index           =   0
      Left            =   120
      MaxLength       =   1
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   60
      Width           =   1095
   End
   Begin VBFLXGRD17.VBFlexGrid Grid 
      Height          =   5355
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   9446
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6090
      Left            =   6705
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
   Begin vbExtra.FlexFn FlexFn1 
      Height          =   405
      Left            =   1920
      TabIndex        =   4
      Top             =   360
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
   Begin VB.Label Label1 
      Caption         =   "Ver Simbolo Um Click  - Escolher Duplo Click"
      Height          =   255
      Left            =   2160
      TabIndex        =   3
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "escFLX"
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

Private Sub CmdLocalizar_Click()
  LocalizaGri1 Grid
End Sub

Private Sub CmdSair_Click()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub FilRelat()
  Dim cSQL As String
  Dim nROWS As Integer
  Dim x As Integer
  If Len(cSUBWHERE) = 0 Then
    cSQL = "SELECT NUMERO,LETRA,DESCRICAO FROM CFLX ORDER BY " & cORDEM
  Else
    cSQL = "SELECT NUMERO,LETRA,DESCRICAO FROM CFLX WHERE " & cSUBWHERE & " ORDER BY " & cORDEM
  End If

  MontaGrid Grid, 3, Array(600, 600, 3500), Array("Letra", "Nº", "Descriçao"), _
            Array("C$Letra", "C$NUMERO", "L$Descricao"), cARQPF, cSQL
  grid_Click

  nROWS = Grid.Rows - 1

  Grid.Col = 0
  For x = 1 To nROWS                           '0 e to titulo
    Grid.Row = x
    Grid.RowHeight(x) = 600

    'Grid.CellWidth = 600

    Grid.CellFontBold = True
    Grid.CellFontName = "isoqsymbol"
    Grid.CellFontSize = 18
    Grid.CellAlignment = FlexAlignmentCenterCenter
           
  Next x




End Sub

Private Sub Form_Load()
'Dim X As Integer
'Dim cLETRA As String
  CenterFormToScreen Me
  cARQPF = PegPath("PATH", "PF")
  aORDEM = Array("NUMERO", "Descricao")
  aORDES = Array("Letra", "Descricao")
  cORDEM = "NUMERO"
  cSUBWHERE = ""
  xmontatoolbar Me.Toolbar1, "escFLX", True
  FilRelat
  TXTFIELDS(0).Font = "isoqsymbol"
  lRETU = False
  '    For X = 32 To 96
  '      cLETRA = Chr(X)
  '      IncluiSQL cARQPF, "SELECT * FROM CFLX WHERE NUMERO=" & X, 2, Array("NUMERO", "LETRA"), Array(X, cLETRA), True, False
  '    Next X
End Sub

Private Sub grid_Click()
  Grid.Col = 0
  TXTFIELDS(0).tEXT = Grid
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
  If KeyAscii > 31 And KeyAscii < 123 Then
    LocalizaGrid Grid, Chr(KeyAscii), 1, False
  Else
    grid_Click
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
'    grid_Click

'End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Dim sButton As String
  'Dim cTMP As String
  Dim nTMP As Integer

  sButton = Button
  sButton = Left(UCase(Replace(sButton, "&", "")), 3)

  If Not AcessaBtnOld("escFLX", Button.Index) Then
    Exit Sub
  End If

  GravaLog 0, Button.Index, sButton, "escFLX"
  lRETU = False


  Select Case sButton
  Case "ESC"
    Grid_DblClick
  Case "NOV"
    ePASS01 = "ISOQSYMBOL"
    ePASS02 = True
    frmCharacters.Show vbModal, Me
    If lRETU Then
      nTMP = eRETU02
      IncluiSQL cARQPF, "SELECT * FROM CFLX WHERE NUMERO=" & nTMP, 2, Array("NUMERO", "LETRA"), Array(nTMP, Chr(nTMP)), True, True
      FilRelat
    End If
  Case "EXC"
    If Grid.Row > 0 Then
      Grid.Col = 1
      If ApagaSQLP(cARQPF, "SELECT * FROM CFLX WHERE NUMERO=" & Grid) Then
        FilRelat
      End If
    End If
  Case "EDI"
    If Grid.Row > 0 Then
      Grid.Col = 1
      ePASS01 = Grid
      frmCFLX.Show vbModal, Me
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
      aARQUIVOS = Array(cARQPF)
      ''Posicao 12 Nome da Tabela
      ''Posicao 13 Nome da Tabela
      aRELCFG = Array("", "", "", 0, False, _
                      False, "", "", "", "", "", _
                      False, "CFLX", "CFLX", "", "")
      FrmFiltro.Show vbModal, Me
      If lRETU Then
        cSUBWHERE = Replace(Replace(eRETU01, "{", ""), "}", "")
        FilRelat
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
  Case "IMP"
    cTIPO = "R"
    zgrp = "PF"
    escRPT.Show vbModal, Me
  Case "SAI"
    CmdSair_Click
  End Select

End Sub


