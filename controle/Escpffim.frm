VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form escpffim 
   Caption         =   "Selecione o produto final  desejado:"
   ClientHeight    =   5775
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8880
   Icon            =   "Escpffim.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   9128
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   5775
      Left            =   7410
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
   Begin vbExtra.FlexFn FlexFn1 
      Height          =   405
      Left            =   0
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
Attribute VB_Name = "escpffim"
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
    Grid.Col = 0
    eRETU01 = Grid
    Grid.Col = 1
    eRETU02 = Grid
    Grid.Col = 2
    eRETU03 = Grid
    Grid.Col = 3
    eRETU04 = Grid
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
    cARQ = PegPath("PATH", "PF")
    If Len(cSUBWHERE) = 0 Then
        cSQL = "SELECT CODFINAL,codigo,DEscr,PF FROM PF WHERE NOT BLOQUEADO ORDER BY " & cORDEM
    Else
        cSQL = "SELECT CODFINAL,codigo,DEscr,PF FROM PF WHERE NOT BLOQUEADO AND " & cSUBWHERE & " ORDER BY " & cORDEM
    End If
    MontaGridFast Grid, 4, Array(1600, 1600, 4000, 200), Array("Final", "Produto", "Descriçao", "PF"), _
        Array("L$CODFINAL", "L$CODIGO", "L$DESCR", "PF"), cARQ, cSQL
End Sub

Private Sub Form_Load()
    CenterFormToScreen Me
    aORDEM = Array("CODFINAL", "CODIGO", "DESCR", "PF ")
    aORDES = Array("Codigo Final", "Codigo", "Descricao", "PF")
    cORDEM = "CODFINAL"
    cSUBWHERE = ""

    lRETU = False
    eRETU01 = ""
    eRETU02 = ""
    eRETU03 = ""
    eRETU04 = 0
    xmontatoolbar Me.Toolbar1, "escPFFIM", True
    FilRelat
End Sub

Private Sub grid_Click()
    With Grid
        If .Rows > 2 Then
            .Col = .Cols - 1
            .ColSel = 0
            .TopRow = .Row
        End If
    End With
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
 '       If .Rows > 2 Then
  '          .Col = .Cols - 1
   '         .ColSel = 0
    '        .TopRow = .Row
     '   End If
   ' End With
'End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim sButton As String
    Dim cARQ As String
    sButton = Button
    sButton = Left(UCase(Replace(sButton, "&", "")), 3)
    
    If Not AcessaBtnOld("escPFFIM", Button.Index) Then
        Exit Sub
    End If
    GravaLog 0, Button.Index, sButton, "escPFFIM"
    
    
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
            cARQ = PegPath("PATH", "PF")
            aARQUIVOS = Array(cARQ)
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
                cSUBWHERE = MontaFiltro(aORDEM, Array("L%", "L%", "L%", "="), eRETU01, eRETU02)
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


