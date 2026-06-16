VERSION 5.00
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "vbccr18.ocx"
Begin VB.Form FrmDesI 
   Caption         =   "Controle de Desenhos"
   ClientHeight    =   3732
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   8616
   Icon            =   "FrmDesI.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3732
   ScaleWidth      =   8616
   StartUpPosition =   2  'CenterScreen
   Begin VBCCR18.DTPicker DTPicker2 
      Height          =   375
      Left            =   4560
      TabIndex        =   6
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   656
      Value           =   44285
      CheckBox        =   -1  'True
      AllowUserInput  =   -1  'True
   End
   Begin VBCCR18.DTPicker DTPicker1 
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2773
      _ExtentY        =   656
      Value           =   44285
      CheckBox        =   -1  'True
      AllowUserInput  =   -1  'True
   End
   Begin VB.TextBox txtfields 
      Height          =   285
      Index           =   7
      Left            =   720
      MaxLength       =   24
      TabIndex        =   3
      Top             =   1740
      Width           =   495
   End
   Begin VB.TextBox txtfields 
      Height          =   285
      Index           =   3
      Left            =   7320
      MaxLength       =   24
      TabIndex        =   8
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox txtfields 
      Height          =   285
      Index           =   4
      Left            =   6360
      MaxLength       =   24
      TabIndex        =   7
      Top             =   1680
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Padronizado"
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtfields 
      Height          =   1125
      Index           =   6
      Left            =   120
      MaxLength       =   24
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Top             =   2400
      Width           =   8175
   End
   Begin VB.TextBox txtfields 
      Height          =   285
      Index           =   2
      Left            =   120
      MaxLength       =   24
      TabIndex        =   2
      Top             =   1680
      Width           =   495
   End
   Begin VB.TextBox txtfields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   24
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox txtfields 
      Height          =   285
      Index           =   5
      Left            =   2880
      MaxLength       =   40
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Width           =   3735
   End
   Begin VB.TextBox txtfields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   24
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   2535
   End
   Begin VBCCR18.CommandButtonW Encerrar 
      Height          =   612
      Left            =   6360
      TabIndex        =   21
      Top             =   120
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Retornar"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "FrmDesI.frx":058A
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW cmdClose 
      Height          =   612
      Left            =   7440
      TabIndex        =   22
      Top             =   120
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Salvar"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "FrmDesI.frx":0B24
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VB.Label lblLabels 
      Caption         =   "Car."
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   10
      Left            =   720
      TabIndex        =   20
      Top             =   1500
      Width           =   435
   End
   Begin VB.Label lblLabels 
      Caption         =   "No."
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   19
      Top             =   1500
      Width           =   435
   End
   Begin VB.Label lblLabels 
      Caption         =   "Cdrom No."
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   8
      Left            =   7320
      TabIndex        =   18
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "FitaDat No."
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   7
      Left            =   6240
      TabIndex        =   17
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Observação"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   16
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Data"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   4
      Left            =   1320
      TabIndex        =   15
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "Revisão"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   675
   End
   Begin VB.Label lblLabels 
      Caption         =   "Desenho"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descrição"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   3
      Left            =   2880
      TabIndex        =   12
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Conjunto"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "FrmDesI"
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
    For iLOOP = 0 To nCAMPOS - 4
      aVAL(iLOOP) = TXTFIELDS(iLOOP)
    Next iLOOP
    aVAL(8) = DTPicker1.Value
    aVAL(9) = DTPicker2.Value
    aVAL(10) = FixNumBol(Check1.Value)
    GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR, 2 '0 conjunto 1 desenho chaves
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
  CenterFormToScreen Me
  cARQ = PegPath("PATH", "DESENHO")
  cSQL = "select * from atual WHERE conjunto='" & cCONJUNTO & "' AND DESENHO='" & cDESENHO & "'"
  nCAMPOS = 11
  aCAM = Array("CONJUNTO", "DESENHO", "REV", "CDROM", "FITADAT", "DESCRICAO", "OBS", "REVC", "RECEBIDO", "DATA", "PADRONIZADO")
  aFOR = Array("C", "C", "NI", "NI", "NI", "C", "C", "C", "D", "D", "BN")
  aPAD = Array(cCONJUNTO, cDESENHO, 0, 0, 0, "", "", "", Date, Date, True)
  aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
  For iLOOP = 0 To nCAMPOS - 4
    TXTFIELDS(iLOOP) = aVAL(iLOOP)
  Next iLOOP
  If IsDate(aVAL(8)) Then
    DTPicker1.Value = aVAL(8)
  End If
  If IsDate(aVAL(9)) Then
    DTPicker2.Value = aVAL(9)
  End If
  Check1.Value = aVAL(10)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

