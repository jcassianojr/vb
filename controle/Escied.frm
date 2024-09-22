VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form escIED 
   Caption         =   "Selecione Simbolo Desejado"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8925
   Icon            =   "Escied.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   8925
   StartUpPosition =   2  'CenterScreen
   Begin VBFLXGRD17.VBFlexGrid Grid 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   9763
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6090
      Left            =   7455
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
Attribute VB_Name = "escIED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cARQPF As String
Dim aORDEM As Variant
Dim aORDES As Variant
Dim cORDEM As String
Dim cSUBWHERE As String

Private Sub Apagar()
'Dim eRETU As Variant
  Dim sSQL As String
  Grid.Col = 3
  iITEM = FixInt(Grid)                         'iTEM E AUTOINCREMENT CHAVE PARA DELECAO
  sSQL = "select * from IED WHERE ITEM=" & iITEM
  If ApagaSQLP(cARQPF, sSQL, "Apagar Descritivo") Then
    FilRelat
  End If
End Sub

Private Function cTIPO()
  cTIPO = " "
  Select Case iMU01

    'grupo 0 embalagem
  Case 1                                       ' Tipo
    cTIPO = "T"
  Case 2                                       ' Fechamento
    cTIPO = "F"
  Case 3                                       ' Identificaçao
    cTIPO = "I"
  Case 4                                       ' Proteçao
    cTIPO = "P"
  Case 5                                       ' Local
    cTIPO = "L"

    'grupo 100 plano controle
  Case 101                                     'Descricao
    cTIPO = "1"
  Case 102                                     'Caracteristica
    cTIPO = "2"
  Case 103                                     'Tolerancia
    cTIPO = "3"
  Case 104                                     'Capacidade Prescisao
    cTIPO = "4"
  Case 105                                     'Carta
    cTIPO = "5"
  Case 106                                     'Plano Reaçao
    cTIPO = "6"
  Case 107                                     'frequencia
    cTIPO = "7"
  Case 108                                     'Especificacao
    cTIPO = "8"

    'grupo 200
  Case 201
    cTIPO = "C"
  Case 202                                     'suptipo PFS
    cTIPO = "X"

    'grupo 300 femea
  Case 301                                     ' Ocorrencia
    cTIPO = "O"
  Case 302                                     'Severidade
    cTIPO = "S"
  Case 303                                     ' Deteçao
    cTIPO = "D"
  Case 304                                     'Carac Atual Preventiva
    cTIPO = "A"
  Case 305                                     'Carac Controle Atual
    cTIPO = "B"
  Case 306                                     ''Sub Tipo Femea
    cTIPO = "Z"

    'grupo 400 qsb
  Case 401                                     ''QSB
    cTIPO = "Q"
  Case 402                                     ''QSB Direito
    cTIPO = "W"
  Case 403                                     ''QSB Esquerdo
    cTIPO = "Y"
  Case 404                                     ''5S
    cTIPO = "9"

  End Select

End Function

Private Sub Editar()
  Grid.Col = 3
  iITEM = Grid
  frmIED.Show vbModal
  FilRelat
End Sub

Private Sub Escolher()
  Grid.Col = 0
  eRETU01 = Grid                               'Tipo
  Grid.Col = 1
  eRETU02 = Grid                               'Descritivo
  Grid.Col = 2
  eRETU03 = Grid                               'Valor
  lRETU = True
  Unload Me
End Sub

Private Sub FilRelat()
  Dim cSQL As String
  If Len(cSUBWHERE) = 0 Then
    If iMU01 = 999 Then
      cSQL = "SELECT * FROM IED  ORDER BY " & cORDEM
    Else
      cSQL = "SELECT * FROM IED WHERE TIPO='" & cTIPO() & "' ORDER BY " & cORDEM
    End If
  Else
    If iMU01 = 999 Then
      cSQL = "SELECT * FROM IED WHERE " & cSUBWHERE & " ORDER BY " & cORDEM
    Else
      cSQL = "SELECT * FROM IED WHERE TIPO='" & cTIPO() & "' AND " & cSUBWHERE & " ORDER BY " & cORDEM
    End If

  End If
  MontaGridFast Grid, 4, Array(400, 5000, 400, 10), Array("Tipo", "Descriçao", "Valor", ""), _
                Array("TIPO", "L$DESCRITIVO", "VALOR", "ITEM"), cARQPF, cSQL

End Sub

Private Sub Form_Load()
'    CentralizaJanela Me
  CenterFormToScreen Me
  aORDEM = Array("TIPO", "DESCRITIVO", "VALOR")
  aORDES = Array("Tipo", "Descriçao", "Valor")
  cORDEM = "TIPO"
  cSUBWHERE = ""
  cARQPF = PegPath("PATH", "PF")
  xmontatoolbar Me.Toolbar1, "escIED", True
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

Private Sub incluir()
  IncluiSQL cARQPF, "IED", 1, Array("TIPO"), Array(cTIPO()), False, False
  FilRelat
End Sub

Private Sub Sair()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

  Dim sButton As String
  sButton = Button
  sButton = Left(UCase(Replace(sButton, "&", "")), 3)

  If Not AcessaBtnOld("escIED", Button.Index) Then
    Exit Sub
  End If

  GravaLog 0, Button.Index, sButton, "escIED"

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
                      False, "IED", "IED", "", "")
      FrmFiltro.Show vbModal, Me
      If lRETU Then
        cSUBWHERE = Replace(Replace(eRETU01, "{", ""), "}", "")
      End If
    Else
      ePASS01 = aORDES
      frmLocalizaa.Show vbModal, Me
      If lRETU Then
        cSUBWHERE = MontaFiltro(aORDEM, Array("L%", "L%", "="), eRETU01, eRETU02)
      End If
    End If
    FilRelat


  Case "IMP"
    'cTIPO = "R"
    zgrp = "PF"
    escRPT.Show vbModal, Me

  Case "NOV"
    incluir

  Case "EDI"
    Editar

  Case "EXC"
    Apagar

  Case "ESC"
    Escolher

  Case "LOC"
    ePASS01 = aORDES
    frmLocalizaa.Show vbModal, Me
    If lRETU Then
      LocalizaGrid Grid, eRETU01, eRETU02, , 1
    End If
  Case "SAI"
    Sair

  End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub


