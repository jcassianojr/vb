VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Object = "{2DA70529-3366-414A-B408-46083BCD481B}#1.8#0"; "VBFLXGRD17.OCX"
Begin VB.Form escBTN 
   Caption         =   "Selecione o Botão Desejado"
   ClientHeight    =   5775
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8880
   Icon            =   "escbtn.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin VBFLXGRD17.VBFlexGrid Grid 
      Height          =   5175
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   9128
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   5775
      Left            =   7410
      TabIndex        =   0
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
   Begin vbExtra.FlexFn FlexFn1 
      Height          =   405
      Left            =   240
      TabIndex        =   1
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
Attribute VB_Name = "escBTN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Apaga_Click()
  Dim cARQ As String
  Dim sSQL As String



  Grid.Col = 0
  cMENU = Grid
  Grid.Col = 1
  cTIPO = Grid
  Grid.Col = 2
  iMENU = Grid
  cARQ = Dbname



  '******************************************************************
  'CONTROLE
  sSQL = "select * from controle WHERE FORM='" & cMENU & "' AND CONTROLE='" & cTIPO & "' AND INDICE=" & iMENU
  '******************************************************************

  If ApagaSQLP(cARQ, sSQL) Then

    '******************************************************************
    'USUARIO
    sSQL = "select * from USUCAD WHERE FORM='" & cMENU & "' AND CONTROLE='" & cTIPO & "' AND INDICE=" & iMENU
    '******************************************************************

    ApagaSQL cARQ, sSQL
    FilRelat
  End If

End Sub

Private Sub cmdPADRAO()
  Dim cARQ, cNOME As String
  cNOME = InputBox("Digite o Nome", "Inclusão Botoes Padrão", " ")
  cNOME = Left$(cNOME, 50)

  cARQ = Dbname


  IncluiSQL cARQ, "CONTROLE", 7, Array("FORM", "INDICE", "IMAGEM", "CAPTION", "TOOLTIP", "CONTROLE", "DISPONIVEL"), _
            Array(cNOME, 1, 1, "&Novo", "Incluir Registro", "Toolbar1", False)
  IncluiSQL cARQ, "CONTROLE", 7, Array("FORM", "INDICE", "IMAGEM", "CAPTION", "TOOLTIP", "CONTROLE", "DISPONIVEL"), _
            Array(cNOME, 2, 13, "&Editar", "Editar Registro", "Toolbar1", False)
  IncluiSQL cARQ, "CONTROLE", 7, Array("FORM", "INDICE", "IMAGEM", "CAPTION", "TOOLTIP", "CONTROLE", "DISPONIVEL"), _
            Array(cNOME, 3, 3, "E&xcluir", "EXcluir Registro", "Toolbar1", False)
  IncluiSQL cARQ, "CONTROLE", 7, Array("FORM", "INDICE", "IMAGEM", "CAPTION", "TOOLTIP", "CONTROLE", "DISPONIVEL"), _
            Array(cNOME, 4, 4, "&Localizar", "Localizar Registro", "Toolbar1", True)
  IncluiSQL cARQ, "CONTROLE", 7, Array("FORM", "INDICE", "IMAGEM", "CAPTION", "TOOLTIP", "CONTROLE", "DISPONIVEL"), _
            Array(cNOME, 5, 6, "&Imprimir", "Imprimir Registro", "Toolbar1", True)
  IncluiSQL cARQ, "CONTROLE", 7, Array("FORM", "INDICE", "IMAGEM", "CAPTION", "TOOLTIP", "CONTROLE", "DISPONIVEL"), _
            Array(cNOME, 6, 11, "Es&colher", "Escolher Registro", "Toolbar1", True)
  IncluiSQL cARQ, "CONTROLE", 7, Array("FORM", "INDICE", "IMAGEM", "CAPTION", "TOOLTIP", "CONTROLE", "DISPONIVEL"), _
            Array(cNOME, 7, 14, "&ordenar", "Ordenar", "Toolbar1", True)
  IncluiSQL cARQ, "CONTROLE", 7, Array("FORM", "INDICE", "IMAGEM", "CAPTION", "TOOLTIP", "CONTROLE", "DISPONIVEL"), _
            Array(cNOME, 8, 15, "&filtrar", "filtrar", "Toolbar1", True)
  IncluiSQL cARQ, "CONTROLE", 7, Array("FORM", "INDICE", "IMAGEM", "CAPTION", "TOOLTIP", "CONTROLE", "DISPONIVEL"), _
            Array(cNOME, 9, 19, "&Sair", "Sair/Retornar", "Toolbar1", True)

  FilRelat

End Sub

Private Sub CmdSair_Click()

  Screen.MousePointer = vbDefault
  Unload Me

End Sub

Private Sub Edit_Click()

  Grid.Col = 0
  cMENU = Grid
  Grid.Col = 1
  cTIPO = Grid
  Grid.Col = 2
  iMENU = Grid
  frmBTN.Show vbModal
  FilRelat

End Sub

Private Sub FilRelat()

  Dim cARQ, cSQL As String

  cARQ = Dbname
  cSQL = "SELECT FORM,CONTROLE,INDICE,IMAGEM,CAPTION,DISPONIVEL FROM CONTROLE ORDER BY FORM,CONTROLE,INDICE"
  MontaGridFast Grid, 6, Array(1600, 1000, 400, 400, 1600, 400), Array("Tela", "Botao", "Ord", "Icone", "Operaçao", "L"), _
                Array("FORM", "CONTROLE", "INDICE", "IMAGEM", "CAPTION", "DISPONIVEL"), cARQ, cSQL
  'MontaGridUltra Grid, 6, Array(1600, 1000, 400, 400, 1600, 400), Array("Tela", "Botao", "Ord", "Icone", "Operaçao", "L"), _
   Array("FORM", "CONTROLE", "INDICE", "IMAGEM", "CAPTION", "DISPONIVEL"), cARQ, cSQL
  'campo boleano

End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  FilRelat
  If FixStr(eLOCALIZA) <> "" Then LocalizaGri1 Grid, eLOCALIZA, 1
  xmontatoolbar Me.Toolbar1, "escBTN", True
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
  If KeyAscii > 31 And KeyAscii < 123 Then
    LocalizaGrid Grid, Chr(KeyAscii), 1, False
  End If
End Sub

'Private Sub Grid_SelChange()
'   With Grid
'      If .Rows > 2 Then
'         .Col = .Cols - 1
'        .ColSel = 0
'       .TopRow = .Row
'   End If
'End With
'End Sub

Private Sub LIB_Click()

  Grid.Col = 0
  cMENU = Grid
  Grid.Col = 1
  cTIPO = Grid
  Grid.Col = 2
  iMENU = Grid
  escbtnusr.Show vbModal

End Sub

Private Sub Novo_Click()
  IncluiSQL Dbname, "CONTROLE", 7, Array("FORM", "INDICE", "IMAGEM", "CAPTION", "TOOLTIP", "CONTROLE", "DISPONIVEL"), _
            Array("esc", 1, 1, "&", " ", "Toolbar1", False)
  FilRelat
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

  Dim sButton As String
  sButton = Button
  sButton = Left(UCase(Replace(sButton, "&", "")), 3)

  If Not AcessaBtnOld("escBTN", Button.Index) Then
    Exit Sub
  End If

  GravaLog 0, Button.Index, sButton, "escBTN"


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

  Case "PAD"
    cmdPADRAO
  Case "SAI"
    CmdSair_Click

  End Select

End Sub


