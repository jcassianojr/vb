VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form frmPFD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PFD"
   ClientHeight    =   2595
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1860
      Width           =   375
   End
   Begin VB.CommandButton Setor 
      Caption         =   "Produção"
      Height          =   315
      Index           =   0
      Left            =   600
      TabIndex        =   7
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Setor 
      Caption         =   "Operador"
      Height          =   315
      Index           =   1
      Left            =   600
      TabIndex        =   6
      Top             =   1980
      Width           =   1575
   End
   Begin VB.CommandButton Setor 
      Caption         =   "Laboratorio"
      Height          =   315
      Index           =   2
      Left            =   2160
      TabIndex        =   5
      Top             =   1560
      Width           =   1575
   End
   Begin VB.CommandButton Setor 
      Caption         =   "Recebimento"
      Height          =   315
      Index           =   3
      Left            =   2160
      TabIndex        =   4
      Top             =   1980
      Width           =   1575
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   3735
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3735
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   4080
      TabIndex        =   10
      Top             =   720
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmpfd.frx":0000
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
      Left            =   4080
      TabIndex        =   11
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmpfd.frx":059A
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
   Begin VB.Label lblLabels 
      Caption         =   "Setores"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Frequencia"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Quantidade"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmPFD"
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
    If MDG("Gravar alteraçôes") Then
        For iLOOP = 0 To nCAMPOS - 1
            aVAL(iLOOP) = txtFields(iLOOP)
        Next iLOOP
        GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR
    End If
    Screen.MousePointer = vbDefault
    Unload Me
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
    Center Me
    cARQ = PegPath("PATH", "PF")
    cSQL = "SELECT * FROM PFD WHERE ID=" & ePASS01
    nCAMPOS = 3
    aCAM = Array("CQTDE", "FREQ", "SETOR")
    aFOR = Array("C", "C", "C")
    aPAD = Array("", "", "")
    aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
    For iLOOP = 0 To nCAMPOS - 1
        txtFields(iLOOP) = aVAL(iLOOP)
    Next iLOOP
End Sub

Private Sub Setor_Click(Index As Integer)

    Select Case Index

    Case 0
        Me.txtFields(2) = "P"

    Case 1
        Me.txtFields(2) = "O"

    Case 2
        Me.txtFields(2) = "L"

    Case 3
        Me.txtFields(2) = "R"

    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

