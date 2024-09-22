VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.1#0"; "vbccr18.ocx"
Object = "{66E63055-5A66-4C79-9327-4BC077858695}#9.0#0"; "newtab01.OCX"
Begin VB.Form frmprotii 
   Caption         =   "Controle de Prototipos"
   ClientHeight    =   4785
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9720
   Icon            =   "frmprotii.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   9720
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   120
      Width           =   375
   End
   Begin NewTabCtl.NewTab SSTab1 
      Height          =   3375
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   5953
      ControlJustAdded=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabHeight       =   520
      ControlVersion  =   9
      TabCaption(0)   =   "GP11"
      Tab(0).ControlCount=   3
      Tab(0).Control(0)=   "Text3"
      Tab(0).Control(1)=   "Text1(8)"
      Tab(0).Control(2)=   "Grid"
      TabCaption(1)   =   "Pedido Entrega"
      Tab(1).ControlCount=   25
      Tab(1).Control(0)=   "Text1(3)"
      Tab(1).Control(1)=   "Text1(11)"
      Tab(1).Control(2)=   "Text1(18)"
      Tab(1).Control(3)=   "Text1(6)"
      Tab(1).Control(4)=   "Text1(5)"
      Tab(1).Control(5)=   "Command1"
      Tab(1).Control(6)=   "Command2"
      Tab(1).Control(7)=   "Text1(10)"
      Tab(1).Control(8)=   "Text1(2)"
      Tab(1).Control(9)=   "Text1(7)"
      Tab(1).Control(10)=   "DTPicker5"
      Tab(1).Control(11)=   "DTPicker6"
      Tab(1).Control(12)=   "DTPicker7"
      Tab(1).Control(13)=   "DTPicker8"
      Tab(1).Control(14)=   "lbl(26)"
      Tab(1).Control(15)=   "lbl(25)"
      Tab(1).Control(16)=   "lbl(24)"
      Tab(1).Control(17)=   "lbl(23)"
      Tab(1).Control(18)=   "lbl(19)"
      Tab(1).Control(19)=   "lbl(17)"
      Tab(1).Control(20)=   "lbl(16)"
      Tab(1).Control(21)=   "lbl(15)"
      Tab(1).Control(22)=   "lbl(22)"
      Tab(1).Control(23)=   "lbl(10)"
      Tab(1).Control(24)=   "lbl(9)"
      TabCaption(2)   =   "Desenho Observaçao"
      Tab(2).ControlCount=   10
      Tab(2).Control(0)=   "Text1(15)"
      Tab(2).Control(1)=   "Text1(14)"
      Tab(2).Control(2)=   "Text1(13)"
      Tab(2).Control(3)=   "Text1(12)"
      Tab(2).Control(4)=   "Text1(17)"
      Tab(2).Control(5)=   "Text1(16)"
      Tab(2).Control(6)=   "Text1(9)"
      Tab(2).Control(7)=   "lbl(21)"
      Tab(2).Control(8)=   "lbl(20)"
      Tab(2).Control(9)=   "lbl(18)"
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   -71640
         TabIndex        =   48
         Top             =   1200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFFF&
         Height          =   285
         Left            =   1920
         TabIndex        =   47
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   480
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   -74160
         TabIndex        =   35
         Top             =   720
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   18
         Left            =   -74040
         TabIndex        =   34
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   -73920
         TabIndex        =   33
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   -70320
         TabIndex        =   32
         Top             =   1560
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "(R)emessa"
         Height          =   255
         Left            =   -68640
         TabIndex        =   30
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "(M)ercantil"
         Height          =   255
         Left            =   -67560
         TabIndex        =   29
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   -73560
         TabIndex        =   25
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -69120
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1920
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   -73680
         TabIndex        =   23
         Top             =   2400
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   -73560
         TabIndex        =   20
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   -71400
         TabIndex        =   19
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   -73560
         TabIndex        =   18
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   -71400
         TabIndex        =   17
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   17
         Left            =   -73560
         TabIndex        =   16
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   16
         Left            =   -71400
         TabIndex        =   15
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   1125
         Index           =   9
         Left            =   -73920
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   13
         Top             =   1440
         Width           =   7695
      End
      Begin MSFlexGridLib.MSFlexGrid Grid 
         Height          =   2295
         Left            =   240
         TabIndex        =   46
         Top             =   840
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   4048
         _Version        =   393216
      End
      Begin VBCCR18.DTPicker DTPicker5 
         Height          =   375
         Left            =   -70920
         TabIndex        =   53
         Top             =   1800
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Value           =   44285
         CheckBox        =   -1  'True
         AllowUserInput  =   -1  'True
      End
      Begin VBCCR18.DTPicker DTPicker6 
         Height          =   375
         Left            =   -68640
         TabIndex        =   54
         Top             =   2400
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Value           =   44285
         CheckBox        =   -1  'True
         AllowUserInput  =   -1  'True
      End
      Begin VBCCR18.DTPicker DTPicker7 
         Height          =   375
         Left            =   -71280
         TabIndex        =   55
         Top             =   600
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         Value           =   44285
         CheckBox        =   -1  'True
         AllowUserInput  =   -1  'True
      End
      Begin VBCCR18.DTPicker DTPicker8 
         Height          =   375
         Left            =   -68760
         TabIndex        =   56
         Top             =   600
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
         Caption         =   "Pedido:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   26
         Left            =   -74760
         TabIndex        =   42
         Top             =   720
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entrada"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   25
         Left            =   -72000
         TabIndex        =   41
         Top             =   600
         Width           =   555
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entrega"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   24
         Left            =   -69480
         TabIndex        =   40
         Top             =   600
         Width           =   555
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtdde:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   23
         Left            =   -74760
         TabIndex        =   39
         Top             =   1200
         Width           =   630
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Valor:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   19
         Left            =   -72240
         TabIndex        =   38
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Comprador:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   17
         Left            =   -74760
         TabIndex        =   37
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone/Ramal:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   16
         Left            =   -71640
         TabIndex        =   36
         Top             =   1560
         Width           =   1200
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   15
         Left            =   -71400
         TabIndex        =   31
         Top             =   1920
         Width           =   390
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nota Fiscal NÝ."
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   22
         Left            =   -74760
         TabIndex        =   28
         Top             =   1920
         Width           =   1065
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Local Entrega:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   10
         Left            =   -74760
         TabIndex        =   27
         Top             =   2400
         Width           =   1035
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   9
         Left            =   -69240
         TabIndex        =   26
         Top             =   2400
         Width           =   390
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desenho:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   21
         Left            =   -74520
         TabIndex        =   22
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rev:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   20
         Left            =   -71880
         TabIndex        =   21
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Obs:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   18
         Left            =   -74760
         TabIndex        =   14
         Top             =   1440
         Width           =   630
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Teste (F)uncional"
      Height          =   255
      Left            =   6360
      TabIndex        =   11
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton Command7 
      Caption         =   "(T)ry Out Solda"
      Height          =   255
      Left            =   6360
      TabIndex        =   10
      Top             =   0
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "(M)atching"
      Height          =   255
      Left            =   5400
      TabIndex        =   9
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "(G)amma"
      Height          =   255
      Left            =   5400
      TabIndex        =   8
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "(B)eta"
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "(A)Alfa"
      Height          =   255
      Left            =   4800
      TabIndex        =   6
      Top             =   0
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   1
      Left            =   720
      TabIndex        =   5
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   2415
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   8040
      TabIndex        =   49
      Top             =   600
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmprotii.frx":058A
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
      Left            =   8040
      TabIndex        =   50
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmprotii.frx":0B24
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
   Begin VBCCR18.DTPicker DTPicker3 
      Height          =   375
      Left            =   3720
      TabIndex        =   51
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Value           =   44285
      CheckBox        =   -1  'True
      AllowUserInput  =   -1  'True
   End
   Begin VBCCR18.DTPicker DTPicker4 
      Height          =   375
      Left            =   6240
      TabIndex        =   52
      Top             =   600
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
      Caption         =   "Fase"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   3720
      TabIndex        =   44
      Top             =   120
      Width           =   345
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ssmt Req.:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   11
      Left            =   5520
      TabIndex        =   3
      Top             =   600
      Width           =   780
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ssmt Data:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   7
      Left            =   2880
      TabIndex        =   2
      Top             =   600
      Width           =   780
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ssmt:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Partnumber:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   990
   End
End
Attribute VB_Name = "frmprotii"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cARQDES As String
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
    For iLOOP = 0 To nCAMPOS - 7
      aVAL(iLOOP) = Text1(iLOOP)
    Next iLOOP
    aVAL(19) = DTPicker3
    aVAL(20) = DTPicker4
    aVAL(21) = DTPicker5
    aVAL(22) = DTPicker6
    aVAL(23) = DTPicker7
    aVAL(24) = DTPicker8
    GrvSQL cARQDES, cSQL, nCAMPOS, aCAM, aVAL, aFOR, 2 'pula chaves PARTNUMBER,SSMT
  End If
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub Command1_Click()
  Text1(2) = "R"
End Sub

Private Sub Command2_Click()
  Text1(2) = "M"
End Sub

Private Sub Command3_Click()
  Text1(1) = "A"
End Sub

Private Sub Command4_Click()
  Text1(1) = "B"
End Sub

Private Sub Command5_Click()
  Text1(1) = "G"
End Sub

Private Sub Command6_Click()
  Text1(1) = "M"
End Sub

Private Sub Command7_Click()
  Text1(1) = "T"
End Sub

Private Sub Command8_Click()
  Text1(5) = "F"
End Sub

Private Sub FilRelat()
  Dim cSQL As String
  Dim cARQ As String

  cARQ = PegPath("PATH", "PF")
  cSQL = "select * from PPAGI WHERE PPAP=" & nPPAP & " ORDER BY ITEM"

  MontaGridFast Grid, 4, Array(400, 1200, 400, 2000), Array("Item", "Data", "Dispo", "obs"), _
                Array("ITEM", "C$DATA", "DISPO", "OBS"), cARQ, cSQL
  Text3.tEXT = PegUltGrid(Grid, 2)

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
  nCAMPOS = 25
  cARQDES = PegPath("PATH", "DESENHO")
  aCAM = Array("PARTNUMBER", "SSMT", "NFTIPO", "VALOR", "FASE", _
               "TELEFONE", "COMPRADOR", "ENTREGA", "PPAP", "OBS", _
               "NFNUMERO", "PEDIDO", "RE2", "DESENH2", "REV", _
               "DESENHO", "RE3", "DESENH3", "QTDDE", "SSMTDAT", _
               "SSMTREQ", "NFDATA", "ENTRDATA", "PEDDATA", "PEDENTR")
  aFOR = Array("C", "C", "C", "N", "C", _
               "C", "C", "C", "NI", "C", _
               "C", "C", "C", "C", "C", _
               "C", "C", "C", "NI", "D", _
               "D", "D", "D", "D", "D")
  aPAD = Array("", "", "", 0, "", _
               "", "", "", 0, "", _
               "", "", "", "", "", _
               "", "", "", 0, "", _
               "", "", "", "", "")

  cSQL = "select * from PROTOI WHERE PARTNUMBER='" & cDESENHO & "' AND SSMT='" & cCONJUNTO & "'"
  aVAL = PegSQL(cARQDES, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
  For iLOOP = 0 To nCAMPOS - 7
    Text1(iLOOP) = aVAL(iLOOP)
  Next iLOOP

  If IsDate(aVAL(19)) Then
    DTPicker3 = aVAL(19)
  End If
  If IsDate(aVAL(20)) Then
    DTPicker4 = aVAL(20)
  End If
  If IsDate(aVAL(21)) Then
    DTPicker5 = aVAL(21)
  End If
  If IsDate(aVAL(22)) Then
    DTPicker6 = aVAL(22)
  End If
  If IsDate(aVAL(23)) Then
    DTPicker7 = aVAL(23)
  End If
  If IsDate(aVAL(24)) Then
    DTPicker8 = aVAL(24)
  End If

  nPPAP = aVAL(8)
  FilRelat
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault

End Sub


