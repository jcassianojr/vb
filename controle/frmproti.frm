VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{7020C36F-09FC-41FE-B822-CDE6FBB321EB}#1.0#0"; "vbccr17.ocx"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Begin VB.Form frmproti 
   Caption         =   "Controle de Prototipos"
   ClientHeight    =   4515
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10170
   Icon            =   "frmproti.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   10170
   StartUpPosition =   2  'CenterScreen
   Begin vbExtra.FlexFn FlexFn1 
      Height          =   405
      Left            =   6000
      TabIndex        =   23
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
   Begin VB.TextBox TXTFIELDS 
      Enabled         =   0   'False
      Height          =   285
      Index           =   7
      Left            =   4080
      TabIndex        =   16
      ToolTipText     =   "Valor da Cotaçao"
      Top             =   1860
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   3
      Left            =   6960
      TabIndex        =   14
      ToolTipText     =   "Telefone/Ramal do Engenheiro"
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Cotação"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   1
      Left            =   3360
      TabIndex        =   11
      ToolTipText     =   "Descriçao do Produto"
      Top             =   720
      Width           =   4695
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   10
      ToolTipText     =   "Nome do Engenheiro"
      Top             =   1440
      Width           =   4095
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   6
      Left            =   1320
      TabIndex        =   9
      Top             =   1080
      Width           =   6375
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Part Number"
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Numero Cliente Engenharia"
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Nome do Cliente Engenharia"
      Top             =   360
      Width           =   4335
   End
   Begin MSFlexGridLib.MSFlexGrid GridAtual 
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Lista de Desenhos"
      Top             =   2280
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   3625
      _Version        =   393216
   End
   Begin XPControls.XPButton cmdClose 
      Height          =   435
      Left            =   8160
      TabIndex        =   17
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmproti.frx":058A
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
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   8160
      TabIndex        =   18
      Top             =   600
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmproti.frx":0B24
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
   Begin XPControls.XPButton novo 
      Height          =   435
      Left            =   8400
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmproti.frx":10BE
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
   Begin XPControls.XPButton Editar 
      Height          =   435
      Left            =   8400
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmproti.frx":1658
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
   Begin XPControls.XPButton excluir 
      Height          =   435
      Left            =   8400
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmproti.frx":1BF2
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
   Begin XPControls.XPButton CmdLocalizar 
      Height          =   375
      Left            =   8400
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Picture         =   "frmproti.frx":218C
      Caption         =   "Localizar"
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
   Begin VBCCR17.DTPicker DTPicker1 
      Height          =   375
      Left            =   1440
      TabIndex        =   24
      Top             =   1800
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Value           =   44285
      CheckBox        =   -1  'True
      AllowUserInput  =   -1  'True
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telefone/Ramal:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   16
      Left            =   5640
      TabIndex        =   15
      Top             =   1440
      Width           =   1200
   End
   Begin VB.Label lblLabels 
      Caption         =   "Valor"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   3
      Left            =   3480
      TabIndex        =   13
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Caption         =   "Engenheiro"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Projeto"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "PartNumber"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cliente:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Nome:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   5
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmproti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cARQDES As String
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
        For iLOOP = 0 To nCAMPOS - 3
            aVAL(iLOOP) = TXTFIELDS(iLOOP)
        Next iLOOP
        aVAL(8) = DTPicker1
        aVAL(9) = FixNumBol(Check1)
        GrvSQL cARQDES, cSQL, nCAMPOS, aCAM, aVAL, aFOR
    End If
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub CmdLocalizar_Click()

    LocalizaGri1 GridAtual

End Sub

Private Sub Editar_Click()

    GridAtual.Col = 0
    cCONJUNTO = FixStr(GridAtual)
    frmprotii.Show vbModal, Me
    Filgrid

End Sub

Private Sub Encerrar_Click()
    If Not MDG("Sair sem gravar") Then
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub excluir_Click()
    Dim sSQL As String
    GridAtual.Col = 0
    cCONJUNTO = FixStr(GridAtual)
    GridAtual.Col = 2
    nPPAP = FixInt(GridAtual)
    '******************************************************************
    'Conjunto
    sSQL = "SELECT * FROM PROTOI WHERE PARTNUMBER='" & cDESENHO & "' AND SSMT='" & cCONJUNTO & "'"
    '******************************************************************
    If ApagaSQLP(cARQDES, sSQL) Then
        Filgrid
        eRETU02 = "GP11:" & Trim(Str(nPPAP)) & Chr(13) & Chr(10)
        eRETU02 = eRETU02 & "Código:" & cDESENHO & Chr(13) & Chr(10)
        eRETU02 = eRETU02 & "SSMT" & cCONJUNTO & Chr(13) & Chr(10)
        MAILENV "PROT0002", eRETU02
    End If

End Sub

Private Sub Filgrid()

    Dim cSQL As String
    cSQL = "SELECT SSMT,FASE,PPAP FROM PROTOI WHERE PARTNUMBER='" & cDESENHO & "'"

    MontaGridFast GridAtual, 3, Array(4000, 1000, 1000), Array("SSMT", "Fase", "GP11"), Array("L$SSMT", "fase", "PPAP"), cARQDES, cSQL

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
    'Dim aRETU As Variant
    CenterFormToScreen Me
    cARQDES = PegPath("PATH", "DESENHO")
    cSQL = "select * from PROTO WHERE PARTNUMBER='" & cDESENHO & "'"
    nCAMPOS = 10
    aCAM = Array("PARTNUMBER", "NOME", "ENGENHEIRO", "ENGRAMAL", "CLIENTE", "CLINOME", "PROJETO", "COTVAL", "COTDATA", "COTACAO")
    aFOR = Array("C", "C", "C", "C", "N", "C", "C", "C", "D", "BN")
    aPAD = Array("", "", "", "", 0, "", "", "", Date, False)
    aVAL = PegSQL(cARQDES, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
    For iLOOP = 0 To nCAMPOS - 3
        TXTFIELDS(iLOOP) = aVAL(iLOOP)
    Next iLOOP
    If IsDate(aVAL(8)) Then
        DTPicker1.Value = aVAL(8)
    End If
    Check1.Value = aVAL(9)
    Filgrid
End Sub

Private Sub Novo_Click()
    On Error Resume Next
    Dim cSSMT As String
    Dim cGP11 As String
    Dim DB As New ADODB.Connection
    Dim RSTAB As New ADODB.Recordset
    Dim sSQL As String
    nREVI = 0
    cSSMT = InputBox("Digite o SSMT", "Inclusão SSMT", "__")
    DB.ConnectionTimeout = 120
    DB.Open GeracArq(cARQDES, , False)
    sSQL = "SELECT * FROM PROTOI WHERE PARTNUMBER='" & cDESENHO & "' AND SSMT='" & cSSMT & "'"
    RSTAB.Open sSQL, DB, adOpenForwardOnly, adLockReadOnly
    If Not RSTAB.EOF Then
        RSTAB.Close
        Alert "SSMT Já Cadastrado"
        Exit Sub
    End If
    RSTAB.Close
    DB.Close
    cGP11 = InputBox("Digite o GP11", "Numero do GP11 - Zero Cria Novo", "0")
    nPF = Val(cGP11)
    If nPF = 0 Then
        DB.ConnectionTimeout = 120
        DB.Open GeracArq(cARQDES)
        RSTAB.Open "PPAG", DB, adOpenDynamic, adLockOptimistic
        RSTAB.AddNew
        RSTAB("CODIGO") = cDESENHO
        RSTAB("NOME") = TXTFIELDS(1)
        RSTAB("CLIENTE") = TXTFIELDS(4)
        RSTAB("CLINOME") = TXTFIELDS(5)
        RSTAB("SSMT") = cSSMT
        nPF = RSTAB("PPAP")
        RSTAB.Update
        RSTAB.Close
        DB.Close
    End If
    IncluiSQL cARQDES, "select * from PROTOI where partnumeber='" & cDESENHO & "' and ssmt='" & cSSMT & " and pf=" & nPF, 3, Array("PARTNUMBER", "SSMT", "PPAP"), Array(cDESENHO, cSSMT, nPF), True, False
    Filgrid
    eRETU02 = "GP11:" & Trim(Str(nPPAP)) & Chr(13) & Chr(10)
    eRETU02 = eRETU02 & "Código:" & cDESENHO & Chr(13) & Chr(10)
    eRETU02 = eRETU02 & "SSMT" & cSSMT & Chr(13) & Chr(10)
    MAILENV "PROT0001", eRETU02

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault

End Sub

