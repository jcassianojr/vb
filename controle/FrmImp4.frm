VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.1#0"; "vbccr18.ocx"
Begin VB.Form frmIMP4 
   Caption         =   "Importando Dados"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6465
   Icon            =   "FrmImp4.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   6465
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "S"
      Height          =   252
      Left            =   1320
      TabIndex        =   16
      Top             =   840
      Width           =   372
   End
   Begin VB.CommandButton Command3 
      Caption         =   "T"
      Height          =   252
      Left            =   960
      TabIndex        =   15
      Top             =   840
      Width           =   372
   End
   Begin VB.CommandButton Command2 
      Caption         =   "C"
      Height          =   252
      Left            =   1320
      TabIndex        =   14
      Top             =   600
      Width           =   372
   End
   Begin VB.CommandButton Command1 
      Caption         =   "M"
      Height          =   252
      Left            =   960
      TabIndex        =   13
      Top             =   600
      Width           =   372
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   372
      Index           =   1
      Left            =   480
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   720
      Width           =   372
   End
   Begin VB.CheckBox Final 
      Caption         =   "Plano Final"
      Height          =   372
      Left            =   240
      TabIndex        =   10
      Top             =   2880
      Width           =   2172
   End
   Begin VB.CheckBox laboratorio 
      Caption         =   "Plano Laboratorio"
      Height          =   372
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   2172
   End
   Begin VB.CheckBox Normal 
      Caption         =   "Plano Producao"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CheckBox gp12 
      Caption         =   "Plano GP12"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CheckBox Recebimento 
      Caption         =   "Plano de Recebimento"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   2295
   End
   Begin VB.CheckBox preliminar 
      Caption         =   "Plano Pre"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox TEXT 
      Height          =   285
      Index           =   1
      Left            =   2400
      MaxLength       =   11
      TabIndex        =   3
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox TEXT 
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
      Left            =   3120
      TabIndex        =   0
      Text            =   "1"
      Top             =   120
      Width           =   975
   End
   Begin XPControls.XPButton CmdescPF 
      Height          =   375
      Left            =   4200
      TabIndex        =   17
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "FrmImp4.frx":058A
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
      Left            =   4680
      TabIndex        =   18
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FrmImp4.frx":0B24
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
      Left            =   4680
      TabIndex        =   19
      Top             =   720
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FrmImp4.frx":10BE
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
   Begin XPControls.XPButton escmu 
      Height          =   375
      Left            =   4200
      TabIndex        =   20
      Top             =   720
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "FrmImp4.frx":1658
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
   Begin VBCCR18.SpinBox mes 
      Height          =   615
      Left            =   2760
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1200
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
   Begin VB.Label lblLabels 
      Caption         =   "Tipo:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   372
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo:"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   2
      Left            =   1800
      TabIndex        =   4
      Top             =   720
      Width           =   612
   End
   Begin VB.Label LABEL 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   6135
   End
   Begin VB.Label Label1 
      Caption         =   "Digite o Numero do Processo e  Codigo de  Origem"
      ForeColor       =   &H00C00000&
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmIMP4"
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

Private Sub importar()
  Dim nPFORI As Long
  Dim sSQL As String
  Dim cORIGEM As String
  Dim cDUPSQL As String

  On Error Resume Next

  nPFORI = Val(tEXT(0))

  'Sequencia Duplicaçao
  cDUPSQL = "select * from DUPLICAR WHERE TABELA='" & cARQIMP & "'"


  'sql Arquivo Origem
  sSQL = "select * from PFCMS03 WHERE PF=" & nPFORI & " AND CODCOMP='" & tEXT(1) & "'"

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



  'Arquivo de origem
  cORIGEM = cARQPF                             'NORMAL
  If preliminar.Value = Checked Then
    cORIGEM = cARQPFP
  End If
  If gp12.Value = Checked Then
    cORIGEM = cARQPFG
  End If

  importa2 cARQPF, cDUPSQL, cORIGEM, sSQL, Sdb, cARQIMP, "PFCO"


End Sub

Private Sub CmdescPF_Click()
  escpf.Show vbModal, Me
  If lRETU Then
    tEXT(0).tEXT = eRETU01
  End If

End Sub

Private Sub Command1_Click()
  TXTFIELDS(1) = "M"
End Sub

Private Sub Command2_Click()
  TXTFIELDS(1) = "C"
End Sub

Private Sub Command3_Click()
  TXTFIELDS(1) = "T"
End Sub

Private Sub Command4_Click()
  TXTFIELDS(1) = "S"
End Sub

Private Sub escmu_Click()
  iMU01 = 0
  Select Case TXTFIELDS(1)
  Case "M"
    iMU01 = 1
  Case "C"
    iMU01 = 2
  Case "T"
    iMU01 = 5
  Case "S"
    iMU01 = 7
  End Select
  If iMU01 > 0 Then
    escmu01.Show vbModal, Me
    If Not lRETU Then Exit Sub
    frmPFMS03.TXTFIELDS(2) = eRETU01
    Command1_Click
  End If

End Sub

Private Sub Final_Click()
  laboratorio.Value = False
  Recebimento.Value = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub laboratorio_Click()
  Recebimento.Value = False
  Final.Value = False
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  cARQFEMEA = PegPath("PATH", "FEMEA")
  cARQPF = PegPath("PATH", "PF")
  cARQPFP = PegPath("PATH", "PFP")
  cARQPFG = PegPath("PATH", "PFG")
  TXTFIELDS(1) = "M"
End Sub



'Private Sub maismes_Click()
'    mes = CStr(IncDec(mes, 1, 1, 3))
'End Sub

'Private Sub menosmes_Click()
'    mes = CStr(IncDec(mes, -1, 1, 3))
'End Sub

Private Sub mes_GotFocus()
  FocusMe
End Sub

Private Sub mes_KeyPress(KeyAscii As Integer)
  KeyAscii = ValiText(KeyAscii, "#NI")
End Sub

Private Sub ok_Click()
  importar
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub Recebimento_Click()
  laboratorio.Value = False
  Final.Value = False
End Sub

Private Sub Text_GotFocus(Index As Integer)
  FocusMe
End Sub

Private Sub tEXT_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index = 0 Then
    KeyAscii = ValiText(KeyAscii, "#NI")
  End If
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

