VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "xpcontrols.ocx"
Begin VB.Form FrmImp2 
   Caption         =   "Importando Dados"
   ClientHeight    =   1755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   1755
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
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
      Index           =   1
      Left            =   3000
      TabIndex        =   4
      Text            =   "1"
      Top             =   720
      Width           =   855
   End
   Begin VB.TextBox Text 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
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
      Left            =   3000
      TabIndex        =   0
      Text            =   "1"
      Top             =   240
      Width           =   855
   End
   Begin XPControls.XPButton ok 
      Height          =   435
      Left            =   4680
      TabIndex        =   5
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FrmImp2.frx":0000
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
      TabIndex        =   6
      Top             =   600
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FrmImp2.frx":059A
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
   Begin XPControls.XPButton CmdescPF 
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   240
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "FrmImp2.frx":0B34
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
      Left            =   3960
      TabIndex        =   8
      Top             =   720
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "FrmImp2.frx":10CE
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
   Begin VB.Label Label1 
      Caption         =   "Digite o Numero do Processo Destino"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2772
   End
   Begin VB.Label LABEL 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   6015
   End
   Begin VB.Label Label1 
      Caption         =   "Digite o Numero do Processo Origem"
      ForeColor       =   &H00C00000&
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   2772
   End
End
Attribute VB_Name = "FrmImp2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cancle_Click()
    Unload Me
End Sub

Private Sub CmdescPF_Click()
    escpf.Show vbModal, Me
    If lRETU Then
        Text(0).Text = eRETU01
    End If
End Sub

Private Sub Command1_Click()
    escpf.Show vbModal, Me
    If lRETU Then
        Text(1).Text = eRETU01
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
    Center Me
End Sub

Private Sub ok_Click()
    Dim nPFORI As Long
    Dim cARQ, cSQL1, cSQL2, cSQL3 As String
    
    cARQ = PegPath("PATH", "PF")
    If cARQIMP = "IE" Then
        cARQ = PegPath("PATH", "IE")
    End If
    
    nPF = Val(Text(1))
    nPFORI = Val(Text(0))
    cSQL1 = "select * from DUPLICAR WHERE TABELA='" & cARQIMP & "'"
    cSQL2 = "select * from " & cARQIMP & " WHERE pf=" & nPF
    cSQL3 = "select * from " & cARQIMP & " WHERE pf=" & nPFORI
    
    importa2 cARQ, cSQL1, cARQ, cSQL3, cARQ, cSQL2, "PFX"
    
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

Private Sub Text_GotFocus(Index As Integer)
    FocusMe
End Sub

Private Sub tEXT_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = ValiText(KeyAscii, "#NI")
End Sub

