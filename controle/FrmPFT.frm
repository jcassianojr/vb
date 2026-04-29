VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form FrmPFT 
   Caption         =   "Trocar Tipo"
   ClientHeight    =   2190
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   ScaleHeight     =   2190
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdTipo 
      Caption         =   "GP12 pla(N)"
      Height          =   315
      Index           =   6
      Left            =   2460
      TabIndex        =   9
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton CmdTipo 
      Caption         =   "GP12 a(G)ressivo"
      Height          =   315
      Index           =   5
      Left            =   960
      TabIndex        =   8
      Top             =   960
      Width           =   1455
   End
   Begin VB.CommandButton CmdTipo 
      Caption         =   "Pre L(A)ncamento"
      Height          =   315
      Index           =   4
      Left            =   960
      TabIndex        =   6
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton CmdTipo 
      Caption         =   "pre(L)iminar"
      Height          =   315
      Index           =   3
      Left            =   2460
      TabIndex        =   5
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton CmdTipo 
      Caption         =   "p(I)loto"
      Height          =   315
      Index           =   2
      Left            =   3300
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton CmdTipo 
      Caption         =   "(P)rodução"
      Height          =   315
      Index           =   1
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton CmdTipo 
      Caption         =   "p(R)ototipo"
      Height          =   315
      Index           =   0
      Left            =   2160
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   360
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   495
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   4560
      TabIndex        =   10
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   767
      Picture         =   "FrmPFT.frx":0000
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
      Left            =   4560
      TabIndex        =   11
      Top             =   240
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   767
      Picture         =   "FrmPFT.frx":059A
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
   Begin VB.Label CABECARIO 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NORMAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label lblLabels 
      Caption         =   "Tipo"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "FrmPFT"
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

Private Sub cmdClose_Click()
    On Error Resume Next
    If MDG("Gravar alteracoes ") Then
        For iLOOP = 0 To nCAMPOS - 1
            aVAL(iLOOP) = TXTFIELDS(iLOOP)
        Next iLOOP
        GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR
       
        lRETU = True
        eRETU01 = TXTFIELDS(0)

       
    End If
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub CmdTipo_Click(Index As Integer)
    Select Case Index
    Case 0
        TXTFIELDS(0) = "R"
    Case 1
        TXTFIELDS(0) = "P"
    Case 2
        TXTFIELDS(0) = "I"
    Case 3
        TXTFIELDS(0) = "L"
    Case 4
        TXTFIELDS(0) = "A"
    Case 5
        TXTFIELDS(0) = "G"
    Case 6
        TXTFIELDS(0) = "N"
    End Select
End Sub

Private Sub Encerrar_Click()
    If Not MDG("Sair sem gravar") Then
        Exit Sub
    End If
    Screen.MousePointer = vbDefault
    lRETU = False
    eRETU01 = ""
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
    Center Me
    lRETU = False
    eRETU01 = ""
    If InStr(Sdb, "PFP.MDB") Then
        CABECARIO.Caption = "PRELIMINAR"
    End If
    If InStr(Sdb, "PFG.MDB") Then
        CABECARIO.Caption = "GP12"
    End If
    cARQ = Sdb
    cSQL = "select pf,tipo from PF WHERE PF=" & nPF
    nCAMPOS = 1
    aCAM = Array("TIPO")
    aFOR = Array("CU")
    aPAD = Array("P")
    aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
    For iLOOP = 0 To nCAMPOS - 1
        TXTFIELDS(iLOOP) = aVAL(iLOOP)
    Next iLOOP
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

