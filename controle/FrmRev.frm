VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form FrmRev 
   Caption         =   "Revisao"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TXTFIELDS 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   7
      Left            =   120
      TabIndex        =   18
      Text            =   "0"
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox TXTFIELDS 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   8
      Left            =   1080
      TabIndex        =   17
      Text            =   "0"
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox TXTFIELDS 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   9
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   10
      Left            =   2520
      MaxLength       =   24
      TabIndex        =   13
      Top             =   1560
      Width           =   2475
   End
   Begin VB.TextBox TXTFIELDS 
      Height          =   285
      Index           =   4
      Left            =   0
      MaxLength       =   24
      TabIndex        =   12
      Top             =   1560
      Width           =   2475
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   6
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   840
      Width           =   615
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   11
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   840
      Width           =   1275
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Index           =   5
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   840
      Width           =   3675
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   615
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox TXTFIELDS 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   855
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   6120
      TabIndex        =   22
      Top             =   720
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FrmRev.frx":0000
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
      Left            =   6120
      TabIndex        =   23
      Top             =   240
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FrmRev.frx":059A
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
      Caption         =   "Estraviadas"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   9
      Left            =   2160
      TabIndex        =   21
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "Devolvidas"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   8
      Left            =   1080
      TabIndex        =   20
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "Copias"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   19
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Caption         =   "obsoleto"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   6
      Left            =   2520
      TabIndex        =   15
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Caption         =   "Setor"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   14
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Caption         =   "Recebedor"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblLabels 
      Caption         =   "Data"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   7
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Caption         =   "PF:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Caption         =   "REVISAO:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   720
      TabIndex        =   4
      Top             =   0
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "TIPO:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   1560
      TabIndex        =   3
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "FrmRev"
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
Dim i As Integer

Private Sub cmdClose_Click()
    On Error Resume Next
    For i = 0 To nCAMPOS - 1
        aVAL(i) = txtFields(i)
    Next i
    GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Encerrar_Click()
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
    Center Me
    cARQ = PegPath("PATH", "PF")
    nREV = FixInt(nREV, 0)
    nREVI = FixInt(nREVI, 0)
    cSQL = ePASS01
    nCAMPOS = 12
    ', ""
    aCAM = Array("PF", "REVISAO", "TIPO", "DATA", "RECSETOR", "RECNOME", "RECNUM", "COPIAS", "COPIASDEV", "COPIASEXT", "OBSOLETO", "DATAREC")
    aFOR = Array("NI", "NI", "CU", "D", "C", "C", "C", "NI", "NI", "NI", "C", "D")
    aPAD = Array(0, 0, "", "", "", "", "", 0, 0, 0, "", "")
    aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
    For i = 0 To nCAMPOS - 1
        txtFields(i) = aVAL(i)
    Next i
    If FixNum(txtFields(6)) > 0 Then
        txtFields(4).Enabled = False
        txtFields(10).Enabled = False
        txtFields(7).Enabled = False
        txtFields(8).Enabled = False
        txtFields(9).Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
    Case 8, 7, 9
        KeyAscii = ValiText(KeyAscii, "#NI")
    End Select
End Sub

