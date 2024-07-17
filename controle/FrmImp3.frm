VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.1#0"; "vbccr18.ocx"
Begin VB.Form FrmImp3 
   Caption         =   "Importando Dados"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   Icon            =   "FrmImp3.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Final 
      Caption         =   "Plano Final"
      Height          =   372
      Left            =   240
      TabIndex        =   10
      Top             =   3240
      Width           =   2172
   End
   Begin VB.CheckBox laboratorio 
      Caption         =   "Plano Laboratorio"
      Height          =   372
      Left            =   240
      TabIndex        =   9
      Top             =   2760
      Width           =   2172
   End
   Begin VB.CheckBox Recebimento 
      Caption         =   "Plano de Recebimento"
      Height          =   372
      Left            =   240
      TabIndex        =   8
      Top             =   2400
      Width           =   2172
   End
   Begin VB.CheckBox gp12 
      Caption         =   "GP12"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CheckBox Normal 
      Caption         =   "Producao"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CheckBox preliminar 
      Caption         =   "Pre"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   0
      Text            =   "1"
      Top             =   120
      Width           =   975
   End
   Begin XPControls.XPButton CmdescPF 
      Height          =   375
      Left            =   3840
      TabIndex        =   11
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "FrmImp3.frx":058A
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
   Begin XPControls.XPButton ok 
      Height          =   435
      Left            =   4440
      TabIndex        =   12
      Top             =   240
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FrmImp3.frx":0B24
      Caption         =   "Confirmar"
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
   Begin XPControls.XPButton Cancle 
      Height          =   435
      Left            =   4440
      TabIndex        =   13
      Top             =   720
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FrmImp3.frx":10BE
      Caption         =   "Cancelar"
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
   Begin VBCCR18.SpinBox mes 
      Height          =   615
      Left            =   2760
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2400
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   12648447
      Enabled         =   0   'False
      Min             =   1
      Max             =   3
      Value           =   1
   End
   Begin VBCCR18.SpinBox seq 
      Height          =   495
      Left            =   840
      TabIndex        =   15
      Top             =   720
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   1
      Max             =   999
      Value           =   10
      Increment       =   5
   End
   Begin VBCCR18.SpinBox Ssq 
      Height          =   495
      Left            =   2160
      TabIndex        =   16
      Top             =   720
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Min             =   1
      Max             =   999
      Value           =   10
      Increment       =   5
   End
   Begin VB.Label Label3 
      Caption         =   "SSQ"
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1680
      TabIndex        =   4
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "SEQ"
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   435
   End
   Begin VB.Label LABEL 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   120
      TabIndex        =   2
      Top             =   3960
      Width           =   4932
   End
   Begin VB.Label Label1 
      Caption         =   "Digite o Numero do Processo e  Sequencia de Origem"
      ForeColor       =   &H00C00000&
      Height          =   435
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "FrmImp3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cARQPF As String
Dim cARQPFP As String
Dim cARQPFG As String
Dim cARQFEMEA As String

Private Sub Cancle_Click()
  Unload Me
End Sub

Private Sub CmdescPF_Click()
  escpf.Show vbModal, Me
  If lRETU Then
    tEXT(0).tEXT = eRETU01
  End If
End Sub

Private Sub Final_Click()
  laboratorio.Value = False
  Recebimento.Value = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  SEQ.Value = Str(nSEQ)
  SSQ.Value = Str(nSSQ)
  cARQFEMEA = PegPath("PATH", "FEMEA")
  cARQPF = PegPath("PATH", "PF")
  cARQPFP = PegPath("PATH", "PFP")
  cARQPFG = PegPath("PATH", "PFG")
  If cARQIMP = "PFQSBLEP" Or cARQIMP = "PFI" Then
    Recebimento.Visible = False
    laboratorio.Visible = False
    Final.Visible = False
    mes.Visible = False
    ' maismes.Visible = False
    ' menosmes.Visible = False
  End If
End Sub

Private Sub importar()
  Dim nPFORI, nSEQORI, nSSQORI As Long
  Dim cORIGEM As String
  Dim sSQL As String
  Dim cDUPSQL As String

  nPFORI = Val(tEXT(0))
  nSEQORI = Val(SEQ.Value)
  nSSQORI = Val(SSQ.Value)

  cDUPSQL = "select * from DUPLICAR WHERE TABELA='" & cARQIMP & "'"

  cORIGEM = cARQPF                             'NORMAL
  If preliminar.Value = Checked Then
    cORIGEM = cARQPFP
  End If
  If gp12.Value = Checked Then
    cORIGEM = cARQPFG
  End If

  sSQL = "select * from " & cARQIMP & " WHERE pf=" & nPFORI & " AND SEQ=" & nSEQORI & " AND SSQ=" & nSSQORI
  If Recebimento.Value = Checked Then
    Select Case mes
    Case 1
      sSQL = "select * from PFCO WHERE PF=" & nPFORI & " AND SEQ=10 AND SSQ=10"
    Case 2
      sSQL = "select * from PFCO WHERE PF=" & nPFORI & " AND SEQ=20 AND SSQ=10"
    Case 3
      sSQL = "select * from PFCO WHERE PF=" & nPFORI & " AND SEQ=30 AND SSQ=10"
    End Select
  End If
  If laboratorio.Value = Checked Then
    Select Case mes
    Case 1
      sSQL = "select * from PFCO WHERE PF=" & nPFORI & " AND SEQ=10 AND SSQ=20"
    Case 2
      sSQL = "select * from PFCO WHERE PF=" & nPFORI & " AND SEQ=20 AND SSQ=20"
    Case 3
      sSQL = "select * from PFCO WHERE PF=" & nPFORI & " AND SEQ=30 AND SSQ=20"
    End Select
  End If
  If Final.Value = Checked Then
    sSQL = "select * from PFCO WHERE PF=" & nPFORI & " AND SEQ=99 AND SSQ=99"
  End If
  importa2 cARQPF, cDUPSQL, cORIGEM, sSQL, Sdb, cARQIMP, "PFC"
  'cDUPARQ,cDUPSQL, cORIARQ, cORISQL,  cDESARQ,  cDESSQL, cTIPO
End Sub

Private Sub laboratorio_Click()
  Recebimento.Value = False
  Final.Value = False
End Sub

Private Sub maismes_Click()
  mes = CStr(IncDec(mes, 1, 1, 3))
End Sub

Private Sub menosmes_Click()
  mes = CStr(IncDec(mes, -1, 1, 3))
End Sub

Private Sub ok_Click()
  importar
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub gp12_Click()
  Normal.Value = False
  preliminar.Value = False
End Sub

Private Sub preliminar_Click()
  gp12.Value = False
  Normal.Value = False
End Sub

Private Sub Normal_Click()
  gp12.Value = False
  preliminar.Value = False
End Sub

Private Sub Text_GotFocus(Index As Integer)
  FocusMe
End Sub

Private Sub tEXT_KeyPress(Index As Integer, KeyAscii As Integer)
  KeyAscii = ValiText(KeyAscii, "#NI")
End Sub

