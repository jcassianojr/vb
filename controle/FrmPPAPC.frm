VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.1#0"; "vbccr18.ocx"
Begin VB.Form FrmPPAPC 
   Caption         =   "PPAPC"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   Icon            =   "FrmPPAPC.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   9360
   StartUpPosition =   2  'CenterScreen
   Begin VBCCR18.DTPicker DTPicker1 
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Value           =   44285
      CheckBox        =   -1  'True
      AllowUserInput  =   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+3ANOS"
      Height          =   375
      Index           =   2
      Left            =   1920
      TabIndex        =   13
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+6 Meses"
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   12
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   240
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+2ANOS"
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   9
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtitem 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      Height          =   1005
      Index           =   2
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1800
      Width           =   8835
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   7320
      TabIndex        =   14
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "FrmPPAPC.frx":058A
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
      Left            =   7320
      TabIndex        =   15
      Top             =   240
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FrmPPAPC.frx":0B24
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
   Begin VBCCR18.DTPicker DTPicker2 
      Height          =   375
      Left            =   2880
      TabIndex        =   17
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Value           =   44285
      CheckBox        =   -1  'True
      AllowUserInput  =   -1  'True
   End
   Begin VBCCR18.DTPicker DTPicker3 
      Height          =   375
      Left            =   4800
      TabIndex        =   18
      Top             =   960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Value           =   44285
      CheckBox        =   -1  'True
      AllowUserInput  =   -1  'True
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   3
      Left            =   2340
      TabIndex        =   11
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "PPAP:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Item"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   1320
      TabIndex        =   7
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Observação:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   6
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Abertura"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   675
   End
   Begin VB.Label lblLabels 
      Caption         =   "Previsto"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   5
      Left            =   2880
      TabIndex        =   4
      Top             =   660
      Width           =   795
   End
   Begin VB.Label lblLabels 
      Caption         =   "Efetuado"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   6
      Left            =   4800
      TabIndex        =   3
      Top             =   660
      Width           =   975
   End
End
Attribute VB_Name = "FrmPPAPC"
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
    For iLOOP = 0 To 2
      aVAL(iLOOP) = TXTFIELDS(iLOOP)
    Next iLOOP
    aVAL(3) = DTPicker1
    aVAL(4) = DTPicker2
    aVAL(5) = DTPicker3
    GrvSQL cARQPF, cSQL, nCAMPOS, aCAM, aVAL, aFOR
  End If
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub Command1_Click(Index As Integer)
  Select Case Index
  Case 0
    DTPicker2 = DTPicker1 + 730
  Case 1
    DTPicker2 = DTPicker1 + 182
  Case 2
    DTPicker2 = DTPicker1 + 1095
  End Select
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
  cARQPF = PegPath("PATH", "PPAP")
  Select Case iPPAP
  Case 1, 11
    cSQL = "select * from PPAPC WHERE ITEM=" & nREVI
  Case 2
    cSQL = "select * from PPAFC WHERE ITEM=" & nREVI
  Case 3
    cSQL = "select * from PPAGC WHERE ITEM=" & nREVI
  End Select

  txtItem.tEXT = nREVI
  nCAMPOS = 6
  aCAM = Array("PPAP", "CODIGO", "OBS", "DATA", "PREVISTO", "EFETUADO")
  aFOR = Array("NI", "C", "C", "", "", "")
  aPAD = Array(0, "", "", Date, Date, Date)
  aVAL = PegSQL(cARQPF, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
  For iLOOP = 0 To 2
    TXTFIELDS(iLOOP) = aVAL(iLOOP)
  Next iLOOP
  If IsDate(aVAL(3)) Then
    DTPicker1.Value = aVAL(3)
  End If
  If IsDate(aVAL(4)) Then
    DTPicker2.Value = aVAL(4)
  End If
  If IsDate(aVAL(5)) Then
    DTPicker3.Value = aVAL(5)
  End If
End Sub

Private Sub xCommand1_Click(Index)
  DTPicker2 = DTPicker1 + 730
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

