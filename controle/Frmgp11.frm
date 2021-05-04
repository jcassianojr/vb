VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form frmGp11 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PPAP"
   ClientHeight    =   5580
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   10485
   Icon            =   "Frmgp11.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin vbExtra.FlexFn FlexFn1 
      Height          =   405
      Left            =   8520
      TabIndex        =   31
      Top             =   1800
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
   Begin VB.TextBox txtFields 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   4
      Left            =   600
      TabIndex        =   13
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   29
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox inativo 
      Caption         =   "Inativo"
      Height          =   195
      Left            =   3000
      TabIndex        =   9
      Top             =   960
      Width           =   855
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   10485
      TabIndex        =   8
      Top             =   5580
      Width           =   10485
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   6
      Left            =   5040
      MaxLength       =   40
      TabIndex        =   7
      Top             =   360
      Width           =   5115
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   5
      Left            =   1200
      MaxLength       =   24
      TabIndex        =   6
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   2
      Left            =   4560
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1320
      Width           =   5715
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   840
      TabIndex        =   3
      Text            =   "0"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   3495
      Left            =   120
      TabIndex        =   14
      Top             =   1800
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   6165
      _Version        =   393216
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   8520
      TabIndex        =   17
      Top             =   4560
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "Frmgp11.frx":058A
      Caption         =   "Retornar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton cmdClose 
      Height          =   435
      Left            =   8520
      TabIndex        =   18
      Top             =   4080
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "Frmgp11.frx":0B24
      Caption         =   "Salvar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton CmdNovo 
      Height          =   435
      Left            =   8520
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "Frmgp11.frx":10BE
      Caption         =   "Novo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton CmdEditar 
      Height          =   435
      Left            =   8520
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "Frmgp11.frx":1658
      Caption         =   "Editar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton CmdExcluir 
      Height          =   435
      Left            =   8520
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "Frmgp11.frx":1BF2
      Caption         =   "Apagar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton ESCMS01A 
      Height          =   375
      Index           =   0
      Left            =   2880
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Picture         =   "Frmgp11.frx":218C
      Caption         =   "mana5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton ESCMS01A 
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   360
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      Picture         =   "Frmgp11.frx":2726
      Caption         =   "LX"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton ESCMS01A 
      Height          =   375
      Index           =   2
      Left            =   4080
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Picture         =   "Frmgp11.frx":2CC0
      Caption         =   "MC"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton EscMA01A 
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1320
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "Frmgp11.frx":325A
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton Command1 
      Height          =   375
      Left            =   1560
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   840
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Picture         =   "Frmgp11.frx":37F4
      Caption         =   "PF"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton EscMA01A 
      Height          =   375
      Index           =   1
      Left            =   2760
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1320
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Picture         =   "Frmgp11.frx":3D8E
      Caption         =   "LX"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton Command5 
      Height          =   375
      Left            =   2880
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   360
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "Frmgp11.frx":4328
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton Command3 
      Height          =   375
      Left            =   2040
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1320
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "Frmgp11.frx":48C2
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton EscMA01A 
      Height          =   375
      Index           =   2
      Left            =   3600
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1320
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Picture         =   "Frmgp11.frx":4E5C
      Caption         =   "LX2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblLabels 
      Caption         =   "SSMT"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   6000
      TabIndex        =   16
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "PF"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   21
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Disposição"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   20
      Left            =   4200
      TabIndex        =   11
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Código:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   5
      Left            =   1260
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cliente:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Caption         =   "PPAP:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmGp11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cARQPF As String
Dim cSQL As String
Dim aVAL As Variant
Dim aFOR As Variant
Dim aCAM As Variant
Dim aPAD As Variant
Dim nCAMPOS As Integer
Dim iLOOP As Integer

Private Sub cmdClose_Click()
    On Error Resume Next
    If MDG("Gravar alteraçôes") Then
        For iLOOP = 0 To nCAMPOS - 2
            aVAL(iLOOP) = TXTFIELDS(iLOOP)
        Next iLOOP
        aVAL(7) = FixNumBol(inativo.Value)
        GrvSQL cARQPF, cSQL, nCAMPOS, aCAM, aVAL, aFOR
    End If
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub CmdEditar_Click()


   

    eRETU02 = " Desenho:" & TXTFIELDS(5) & "-" & Replace(TXTFIELDS(6), " ", "_")
    eRETU02 = eRETU02 & " Cliente:" & TXTFIELDS(1) & "-" & Replace(TXTFIELDS(2), " ", "_")
    Grid.Col = 0
    nREVI = Grid
    frmPPAPI.Show vbModal
    'analisar em qual cituacao ira chamar frmppapc frmpapi
    
    
    'If Index = 0 Then
    '    Grid.Col = 0
    '    nREVI = Grid
    '    frmPPAPI.Show vbModal
    'Else
    '    Grid.Col = 0
    '    nREVI = Grid
    '    FrmPPAPC.Show vbModal
    'End If
    Filgrid
End Sub

Private Sub CmdExcluir_Click()
    Dim sSQL As String
    Grid.Col = 0
    nREVI = Grid
    Select Case iPPAP
    Case 1, 11
        sSQL = "select * from PPAPI WHERE PPAP=" & nPPAP & " AND  ITEM=" & nREVI
    Case 2
        sSQL = "select * from PPAFI WHERE PPAP=" & nPPAP & " AND  ITEM=" & nREVI
    Case 3
        sSQL = "select * from PPAGI WHERE PPAP=" & nPPAP & " AND  ITEM=" & nREVI
    End Select
    If ApagaSQLP(cARQPF, sSQL) Then
        Filgrid
    End If
End Sub

Private Sub CmdNovo_Click()
    Dim nITEM            As Long
    Dim sSQL As String
    Select Case iPPAP
    Case 1, 11
        sSQL = "PPAPI WHERE PPAP=" & nPPAP
    Case 2
        sSQL = "PPAFI WHERE PPAP=" & nPPAP
    Case 3
        sSQL = "PPAGI WHERE PPAP=" & nPPAP
    End Select
    nITEM = PegMAXSQL(cARQPF, sSQL, "ITEM", 0)
    nITEM = nITEM + 1
    Select Case iPPAP
    Case 1, 11
        sSQL = "PPAPI"
    Case 2
        sSQL = "PPAFI"
    Case 3
        sSQL = "PPAGI"
    End Select
    sSQL = "SELECT * FROM " & sSQL & " WHERE PPAP=" & nPPAP & "AND  ITEM=" & nITEM
    IncluiSQL cARQPF, sSQL, 3, Array("PPAP", "DATA", "ITEM"), Array(nPPAP, Date, nITEM), True, False
    Filgrid
End Sub

Private Sub Command1_Click()
    Dim aRETU As Variant
    Dim sSQL As String
    escpffim.Show vbModal, Me
    If lRETU Then
        TXTFIELDS(5) = eRETU02
        TXTFIELDS(6) = eRETU03
        TXTFIELDS(4) = eRETU04
        sSQL = "select CONPES from PF WHERE PF=" & eRETU04
        aRETU = PegSQL(cARQPF, sSQL, 1, Array("CONPES"), Array("N"), Array(0))
        TXTFIELDS(11) = aRETU(0)
    End If

End Sub

Private Sub Command3_Click()
    Dim cARQ As String
    Dim aRETU As Variant
    Dim sSQL             As String
    Dim nNUMERO          As Long
    nNUMERO = FixInt(TXTFIELDS(1), 0)
    cARQ = GeraConn(zMANA5EMP, "JETFOX")
    sSQL = "SELECT NOME FROM MA01 WHERE NUMERO=" & nNUMERO
    aRETU = PegSQL(cARQ, sSQL, 1, Array("NOME"), Array("C"), Array(""))
    If lRETU Then
        TXTFIELDS(2) = aRETU(0)
    End If
End Sub

Private Sub Command5_Click()

    Dim cARQ As String
    Dim aRETU As Variant
    Dim sSQL             As String
    Dim cCODIGO          As String
    cCODIGO = FixStr(TXTFIELDS(5), "", "TRIM")
    cARQ = GeraConn(zMANA5EMP, "JETFOX")
    sSQL = "SELECT NOME FROM MS01 WHERE CODIGO='" & cCODIGO & "'"
    aRETU = PegSQL(cARQ, sSQL, 1, Array("NOME"), Array("C"), Array(""))
    If lRETU Then
        TXTFIELDS(6) = aRETU(0)
    End If
End Sub

Private Sub Encerrar_Click()
    If Not MDG("Sair sem gravar") Then
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub EscMA01A_Click(Index As Integer)
    lRETU = False
    eRETU01 = ""
    eRETU02 = ""
    cARQESC = "MA01"
    Select Case Index
    Case 0
        cARQESC = "MA01"
    Case 1
        ePASS01 = "LOGIX"
    Case 2
        ePASS01 = "LOGIX"
        cARQESC = "2DIG"
    End Select
    escNUMNOM.Show vbModal, Me

    If lRETU Then

        frmPPAP.TXTFIELDS(1) = eRETU01
        frmPPAP.TXTFIELDS(2) = eRETU02

    End If


End Sub

Private Sub ESCMS01A_Click(Index As Integer)
    ePASS01 = "MANA5"
    If Index = 1 Then
        ePASS01 = "LOGIX"
    End If
    If Index = 2 Then
        ePASS01 = "MICRO"
    End If
    escms01.Show vbModal, Me
    If lRETU Then
        frmPPAP.TXTFIELDS(5) = eRETU01
        frmPPAP.TXTFIELDS(6) = eRETU02
    End If

End Sub

Private Sub Filgrid()
    Dim sSQL As String
    Select Case iPPAP
    Case 1, 11
        sSQL = "select * from PPAPI WHERE PPAP=" & nPPAP & " ORDER BY ITEM"
    Case 2
        sSQL = "select * from PPAFI WHERE PPAP=" & nPPAP & " ORDER BY ITEM"
    Case 3
        sSQL = "select * from PPAGI WHERE PPAP=" & nPPAP & " ORDER BY ITEM"
    End Select
    MontaGridFast Grid, 4, Array(400, 1200, 400, 2000), Array("Item", "Data", "Dispo", "Obs"), _
        Array("Item", "Data", "Dispo", "OBS"), cARQPF, sSQL
    TXTFIELDS(29) = PegUltGrid(Grid, 2)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
    CenterFormToScreen Me
    cARQPF = PegPath("PATH", "PPAP")
    Select Case iPPAP
    Case 1, 11
        cSQL = "select * from PPAP WHERE PPAP=" & nPPAP
    Case 2
        cSQL = "select * from PPAF WHERE PPAP=" & nPPAP
        Me.Caption = "PPAP Fornecedores"
    Case 3
        cSQL = "select * from PPAG WHERE PPAP=" & nPPAP
        Me.Caption = "GP11"
    End Select

    
    
    nCAMPOS = 8
    aCAM = Array("PPAP", "CLIENTE", "CLINOME", "SSMT", "PF", "CODIGO", "NOME", "INATIVO")
    aFOR = Array("NI", "NI", "C", "C", "NI", "C", "C", "BN")
    aPAD = Array(0, 0, "", "", 0, "", "", False)
    aVAL = PegSQL(cARQPF, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
    For iLOOP = 0 To nCAMPOS - 2
        TXTFIELDS(iLOOP) = aVAL(iLOOP)
    Next iLOOP
    inativo.Value = aVAL(7)

    Filgrid
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault

End Sub

Private Sub txtFields_GotFocus(Index As Integer)
    FocusMe
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
        KeyAscii = ValiText(KeyAscii, "#NI")
    End If
End Sub


