VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Object = "{075212A8-C1CF-444E-939D-F6046CCDBC08}#1.5#0"; "VBFLXGRD18.OCX"
Begin VB.Form escMENU 
   Caption         =   "Selecione o Menu Desejado"
   ClientHeight    =   5772
   ClientLeft      =   168
   ClientTop       =   456
   ClientWidth     =   8880
   Icon            =   "escMENU.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5772
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VBFLXGRD18.VBFlexGrid Grid 
      Height          =   5175
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   6975
      _ExtentX        =   12298
      _ExtentY        =   9123
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   5772
      Left            =   7404
      TabIndex        =   0
      Top             =   0
      Width           =   1476
      _ExtentX        =   2604
      _ExtentY        =   10181
      ButtonWidth     =   614
      ButtonHeight    =   953
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin vbExtra.FlexFn FlexFn1 
      Height          =   396
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   699
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
Attribute VB_Name = "escMENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Apaga_Click()
  Dim eRETU As Variant
  Dim cdb As String
  Dim sSQL As String
  If Not MDG("Apagar Menu", "Confirme Exclusão") Then
    Exit Sub
  End If

  Grid.Col = 0
  cMENU = Grid
  Grid.Col = 1
  iMENU = Grid
  cdb = Dbname

  '******************************************************************
  'MENU
  sSQL = "select * from MENU WHERE MENU='" & cMENU & "' AND INDICE=" & iMENU
  '******************************************************************

  eRETU = ApagaSQL(cdb, sSQL)

  '******************************************************************
  'USUARIO
  sSQL = "select * from MENUUSU WHERE MENU='" & cMENU & "' AND INDICE=" & iMENU
  '******************************************************************

  eRETU = ApagaSQL(cdb, sSQL)
  FilRelat

End Sub

Private Sub cmdsair_Click()

  Screen.MousePointer = vbDefault
  Unload Me

End Sub

Private Sub Edit_Click()

  Grid.Col = 0
  cMENU = Grid
  Grid.Col = 1
  iMENU = FixInt(Grid)
  frmMENU.Show vbModal
  FilRelat

End Sub

Private Sub FilRelat()
  Dim sSQL As String
  On Error Resume Next
  sSQL = "SELECT Menu,INDICE,descricao,CADASTRO FROM MENU ORDER BY MENU,INDICE"
  '       MontaGridFast Grid, 4, Array(1600, 400, 4000, 400), Array("Nome", "I", "Descricao", "Cadastro"), _
          Array("Menu", "Indice", "descricao", "Cadastro"), Dbname, sSQL
  MontaGridUltra Grid, 4, Array(1600, 400, 4000, 400), Array("Nome", "I", "Descricao", "Cadastro"), _
                 Array("Menu", "Indice", "descricao", "Cadastro"), Dbname, sSQL
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  FilRelat
  If FixStr(eLOCALIZA) <> "" Then LocalizaGri1 Grid, eLOCALIZA, 1
  xmontatoolbar Me.Toolbar1, "escMENU", True
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

''End Sub

Private Sub LIB_Click()

  If Grid.Row > 0 Then
    Grid.Col = 0
    cMENU = Grid
    Grid.Col = 1
    iMENU = Grid
    escmnuusr.Show vbModal
    FilRelat
  End If
End Sub

Private Sub Novo_Click()
  IncluiSQL Dbname, "SELECT * FROM MENU WHERE MENU='???' AND INDICE=0", 2, Array("MENU", "INDICE"), _
            Array("???", 0), True, False
  FilRelat
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

  Dim sButton As String
  sButton = Button
  sButton = Left(UCase(Replace(sButton, "&", "")), 3)

  If Not AcessaBtnOld("escMENU", Button.Index) Then
    Exit Sub
  End If
  GravaLog 0, Button.Index, sButton, "escMENU"


  Select Case sButton

  Case "NOV"
    Novo_Click

  Case "EDI"
    Edit_Click

  Case "EXC"
    Apaga_Click

  Case "LIB"
    LIB_Click

  Case "LOC"
    LocalizaGri1 Grid

  Case "SAI"
    cmdsair_Click

  End Select

End Sub


