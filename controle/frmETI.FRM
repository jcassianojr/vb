VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.0#0"; "vbccr17.ocx"
Begin VB.Form frmETI 
   Caption         =   "Cadastro ETI"
   ClientHeight    =   6285
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10170
   Icon            =   "frmETI.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   10170
   StartUpPosition =   2  'CenterScreen
   Begin VBCCR17.DTPicker DTPicker1 
      Height          =   375
      Left            =   2160
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Enabled         =   0   'False
      Value           =   44285
      CheckBox        =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Distribuicao"
      Height          =   495
      Left            =   7680
      TabIndex        =   22
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   315
      Index           =   8
      Left            =   4200
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   7
      Left            =   1320
      TabIndex        =   4
      Top             =   1320
      Width           =   6855
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   405
      Index           =   6
      Left            =   5520
      TabIndex        =   8
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton CmdNovaREV 
      Caption         =   "Nova Revisão"
      Height          =   495
      Left            =   4080
      TabIndex        =   18
      Top             =   2520
      Width           =   735
   End
   Begin VB.CommandButton CmdMotRev 
      Caption         =   "Motivos"
      Height          =   375
      Left            =   7680
      TabIndex        =   15
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   5
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2520
      Width           =   615
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   0
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   3
      Top             =   960
      Width           =   6855
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   3
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   6855
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   5
      Top             =   1680
      Width           =   6855
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   4
      Left            =   1320
      TabIndex        =   7
      Top             =   2040
      Width           =   6855
   End
   Begin MSFlexGridLib.MSFlexGrid gridrev 
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   1931
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid gridrevi 
      Height          =   1575
      Left            =   120
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4320
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2778
      _Version        =   393216
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   8280
      TabIndex        =   23
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmETI.frx":058A
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
      Left            =   8280
      TabIndex        =   24
      Top             =   240
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmETI.frx":0B24
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
   Begin XPControls.XPButton ComMotInc 
      Height          =   435
      Left            =   7800
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   4320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmETI.frx":10BE
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
   Begin XPControls.XPButton ComMotAlt 
      Height          =   435
      Left            =   7800
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmETI.frx":1658
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
   Begin XPControls.XPButton ComMotExc 
      Height          =   435
      Left            =   7800
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmETI.frx":1BF2
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
   Begin VB.Label Label 
      Caption         =   "Cognome"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   8
      Left            =   3240
      TabIndex        =   21
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "Superficie"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   20
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "Nº"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   6
      Left            =   5040
      TabIndex        =   19
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label 
      Caption         =   "Revisäo"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "Especificaçäo"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "Salt-Spray"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label 
      Caption         =   "Camada"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label 
      Caption         =   "Codigo"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label 
      Caption         =   "Nome"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "frmETI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cARQ As String
Dim cSQL As String
Dim aVAL As Variant
Dim aFOR As Variant
Dim aCAM As Variant
Dim aPAD As Variant
Dim nCAMPOS As Integer
Dim iLOOP As Integer
Dim cARQPF As String
Dim nNUMERO As Integer
Dim nREV As Integer
Dim nREVI As Integer

Private Sub cmdClose_Click()
    On Error Resume Next
    If MDG("Gravar alteraçôes") Then
        For iLOOP = 0 To nCAMPOS - 2
            aVAL(iLOOP) = TXTFIELDS(iLOOP)
        Next iLOOP
        aVAL(9) = DTPicker1.Value
        GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR
    End If
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub CmdMotRev_Click()
    gridrev.Col = 1
    nREV = FixInt(gridrev)
    ComMotAlt.Enabled = True
    ComMotExc.Enabled = True
    ComMotInc.Enabled = True
    Filgridrevi
End Sub

Private Sub CmdNovaREV_Click()
    If nNUMERO < 1 Then
        Alert ("Preencher numero interno")
        Exit Sub
    End If
    TXTFIELDS(5) = FixInt(TXTFIELDS(5), 0) + 1
    DTPicker1 = Date
    IncluiSQL cARQPF, "SELECT * FROM REV WHERE PF=" & nPF & " and tipo='ETI' and revisao=" & TXTFIELDS(5), 4, Array("PF", "REVISAO", "TIPO", "DatA"), _
        Array(nNUMERO, TXTFIELDS(5), "ETI", Date), True, False
    Filgridrev
End Sub

Private Sub Command1_Click()
    If nREV = 0 Then
        Alert "Revisão Nao Escolhida"
        Exit Sub
    End If
    ePASS01 = "select * from rev WHERE PF=" & nPF & " AND TIPO='ETI' AND REVISAO=" & nREV
    FrmRev.Show vbModal, Me
End Sub

Private Sub ComMotAlt_Click()
    If gridrevi.Row > 0 Then
        gridrevi.Col = 3
        nREVI = FixInt(gridrevi, 0)
        ePASS01 = "select * from revi WHERE PF=" & nNUMERO & " AND TIPO='ETI' AND REVISAO=" & nREV & " AND  ITEM=" & nREVI
        lEDITAR = True
        frmREVI.Show vbModal, Me
        Filgridrevi
    End If
End Sub

Private Sub ComMotExc_Click()
    Dim sSQL As String
    If gridrevi.Row > 0 Then
        gridrevi.Col = 3
        nREVI = FixInt(gridrevi, 0)
        sSQL = "select * from revi WHERE PF=" & nNUMERO & " AND TIPO='ETI' AND REVISAO=" & nREV & " AND  ITEM=" & nREVI
        If ApagaSQLP(cARQPF, sSQL) Then
            Filgridrevi
        End If
    End If
End Sub

Private Sub ComMotInc_Click()
    Dim nITEM As Long
    Dim sSQL As String
    sSQL = "revI WHERE PF=" & nNUMERO & " AND TIPO='ETI' AND REVISAO=" & nREV
    nITEM = FixInt(PegMAXSQL(cARQPF, sSQL, "ITEM", 0)) + 1
    sSQL = "select * from revi WHERE PF=" & nNUMERO & " AND TIPO='ETI' AND REVISAO=" & nREV & " AND  ITEM=" & nITEM
    IncluiSQL cARQPF, sSQL, 4, Array("PF", "REVISAO", "TIPO", "ITEM"), _
        Array(nNUMERO, nREV, "ETI", nITEM), True, False
End Sub

Private Sub Filgridrev()
    Dim sSQL As String
    sSQL = "select PF,REVISAO,TIPO,DATA from rev WHERE PF=" & nNUMERO & " AND TIPO='ETI' ORDER BY REVISAO"
    MontaGridFast gridrev, 4, Array(400, 400, 400, 1200), Array("PF", "REV", "Tipo", "DATA"), _
        Array("PF", "REVISAO", "TIPO", "DATA"), cARQPF, sSQL
End Sub

Private Sub Filgridrevi()
    Dim sSQL As String
    sSQL = "select * from revI WHERE PF=" & nNUMERO & " AND REVISAO=" & nREV & " AND TIPO='ETI'ORDER BY ITEM"
    MontaGridFast gridrevi, 5, Array(400, 400, 400, 400, 2000), Array("PF", "Rev", "Tipo", "Item", "Motivo"), _
        Array("PF", "REVISAO", "TIPO", "ITEM", "MOTIVO"), cARQPF, sSQL
End Sub

Private Sub Encerrar_Click()
    If Not MDG("Sair sem gravar") Then
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
    CenterFormToScreen Me
    cARQPF = PegPath("PATH", "PF")
    cARQ = GeraConn(zMANA5EMP, "SDECDX")
    cSQL = "select * from ETI WHERE CODIGO='" & Trim(ePASS01) & "'"
    nCAMPOS = 10
    aCAM = Array("CODIGO", "CAMADA", "SALTO", "NOME", "ESPE", "REV", "NINTERNO", "SUPE", "COGNOME", "REVDATA")
    aFOR = Array("C", "C", "C", "C", "C", "NI", "NI", "C", "C", "D")
    aPAD = Array("", "", "", "", "", 0, 0, "", "", Date)
    aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
    For iLOOP = 0 To nCAMPOS - 2
        TXTFIELDS(iLOOP) = aVAL(iLOOP)
    Next iLOOP
    If IsDate(aVAL(9)) Then
        DTPicker1.Value = aVAL(9)
    End If
    nNUMERO = FixInt(TXTFIELDS(6), 0)
    ComMotAlt.Enabled = False
    ComMotExc.Enabled = False
    ComMotInc.Enabled = False
    Filgridrev
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 6 Then
        KeyAscii = ValiText(KeyAscii, "#NI")
    End If
End Sub

Private Sub TXTFIELDS_LostFocus(Index As Integer)
    If Index = 6 Then
        nNUMERO = FixInt(TXTFIELDS(6), 0)
    End If
End Sub


