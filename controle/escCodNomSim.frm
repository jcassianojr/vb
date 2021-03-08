VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form EscCodNomSim 
   Caption         =   "Selecione o Desejado"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   210
   ClientWidth     =   8265
   HelpContextID   =   6
   LinkTopic       =   "Form1"
   ScaleHeight     =   7680
   ScaleWidth      =   8265
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   7680
      Left            =   6765
      TabIndex        =   1
      ToolTipText     =   "Botoes de Açoes"
      Top             =   0
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   13547
      ButtonWidth     =   714
      ButtonHeight    =   714
      Appearance      =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Lista de Registros"
      Top             =   480
      Width           =   6525
      _ExtentX        =   11509
      _ExtentY        =   12515
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
Attribute VB_Name = "EscCodNomSim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cARQ As String
Dim cSQL As String
Dim aORDEM As Variant
Dim aORDES As Variant
Dim cORDEM As String

Private Sub Escolher()
    lRETU = True
    Grid.Col = 0
    eRETU01 = Grid
    Grid.Col = 1
    eRETU02 = Grid
    Unload Me
End Sub

Private Sub Filgrid()
    Dim cSQLUSO As String
    cSQLUSO = cSQL & " ORDER BY " & cORDEM
    MontaGridFast Grid, 2, Array(1200, 5000), Array("Codigo", "Dizer"), _
        Array("L$CODIGO", "L$NOME"), cARQ, cSQLUSO
End Sub

Private Sub Form_Load()
    '    CentralizaJanela Me
    Center Me
    cARQ = GeracArq(PegPath("PATH", ePASS01))
    cSQL = ePASS02
    aORDEM = ePASS03
    aORDES = Array("Codigo", "Nome")
    cORDEM = aORDEM(0)
    xmontatoolbar Me.Toolbar1, "escSIM", True
    Filgrid
    lRETU = False
End Sub

Private Sub Grid_DblClick()
    Escolher
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
    If KeyAscii > 31 And KeyAscii < 123 Then
        LocalizaGrid Grid, Chr(KeyAscii), 1, False
    End If
End Sub

Private Sub Grid_SelChange()
    With Grid
        If .Rows > 2 Then
            .Col = .Cols - 1
            .ColSel = 0
            .TopRow = .Row
        End If
    End With
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim sButton As String
    'Dim nINDICE As Integer
    sButton = Button
    sButton = Left(UCase(Replace(sButton, "&", "")), 3)
    If sButton = "SAI" Then
        lRETU = False
        Unload Me
    End If
    Select Case sButton
    Case "ORD"
        ORDENAR
    Case "ESC"
        Escolher
    Case "LOC"
        ePASS01 = aORDES
        frmLocalizaa.Show vbModal, Me
        If lRETU Then
            LocalizaGrid Grid, eRETU01, eRETU02 + 1, , 1
        End If
    End Select
End Sub

Private Sub ORDENAR()
    ePASS01 = aORDES
    escOrdem.Show vbModal, Me
    If lRETU Then
        cORDEM = aORDEM(eRETU01)
        Filgrid
    End If
End Sub


