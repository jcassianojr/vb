VERSION 5.00
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "vbccr18.ocx"
Begin VB.Form FrmFEMRPNO 
   Caption         =   "Femea Apurado Operacoes"
   ClientHeight    =   2220
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   8436
   Icon            =   "FrmFemRPNO.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2220
   ScaleWidth      =   8436
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox text 
      Alignment       =   1  'Right Justify
      Height          =   330
      Index           =   14
      Left            =   6480
      TabIndex        =   27
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox text 
      Alignment       =   1  'Right Justify
      Height          =   330
      Index           =   13
      Left            =   5640
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox text 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   330
      Index           =   4
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   360
      Width           =   810
   End
   Begin VB.TextBox text 
      Alignment       =   1  'Right Justify
      Height          =   330
      Index           =   12
      Left            =   4800
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1095
      Width           =   735
   End
   Begin VB.TextBox text 
      Alignment       =   1  'Right Justify
      Height          =   330
      Index           =   11
      Left            =   4200
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1080
      Width           =   510
   End
   Begin VB.TextBox text 
      Alignment       =   1  'Right Justify
      Height          =   330
      Index           =   10
      Left            =   3600
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1095
      Width           =   465
   End
   Begin VB.TextBox text 
      Alignment       =   1  'Right Justify
      Height          =   330
      Index           =   9
      Left            =   3000
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1095
      Width           =   450
   End
   Begin VB.TextBox text 
      Alignment       =   1  'Right Justify
      Height          =   330
      Index           =   8
      Left            =   2400
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1095
      Width           =   450
   End
   Begin VB.TextBox text 
      Alignment       =   1  'Right Justify
      Height          =   330
      Index           =   7
      Left            =   1800
      TabIndex        =   10
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1095
      Width           =   465
   End
   Begin VB.TextBox text 
      Alignment       =   1  'Right Justify
      Height          =   330
      Index           =   6
      Left            =   840
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1080
      Width           =   810
   End
   Begin VB.TextBox text 
      Alignment       =   1  'Right Justify
      Height          =   330
      Index           =   5
      Left            =   135
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1095
      Width           =   510
   End
   Begin VB.TextBox text 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   330
      Index           =   1
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox text 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   330
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   360
      Width           =   750
   End
   Begin VB.TextBox text 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   330
      Index           =   2
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   690
   End
   Begin VB.TextBox text 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   330
      Index           =   3
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   810
   End
   Begin VBCCR18.CommandButtonW cmdClose 
      Height          =   612
      Left            =   7320
      TabIndex        =   28
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
      Picture         =   "FrmFemRPNO.frx":058A
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW Encerrar 
      Height          =   612
      Left            =   7320
      TabIndex        =   29
      Top             =   840
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
      Picture         =   "FrmFemRPNO.frx":0B24
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tot RPN"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   9
      Left            =   840
      TabIndex        =   25
      Top             =   840
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proc"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   2
      Left            =   3600
      TabIndex        =   24
      Top             =   120
      Width           =   330
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tot RPN Mes Anteriores"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   15
      Left            =   5040
      TabIndex        =   22
      Top             =   840
      Width           =   1845
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ">100"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   14
      Left            =   4200
      TabIndex        =   21
      Top             =   840
      Width           =   360
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "40-100"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   13
      Left            =   3480
      TabIndex        =   20
      Top             =   855
      Width           =   495
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0-40"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   12
      Left            =   3000
      TabIndex        =   19
      Top             =   855
      Width           =   315
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Maior"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   11
      Left            =   2400
      TabIndex        =   18
      Top             =   855
      Width           =   390
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ">40"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   10
      Left            =   1920
      TabIndex        =   17
      Top             =   840
      Width           =   270
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NÝCausa"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   16
      Top             =   840
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PF"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   960
      TabIndex        =   7
      Top             =   120
      Width           =   195
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seq:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   570
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mes:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   3
      Left            =   1800
      TabIndex        =   4
      Top             =   120
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ano:"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   4
      Left            =   2640
      TabIndex        =   3
      Top             =   120
      Width           =   630
   End
End
Attribute VB_Name = "FrmFEMRPNO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const cFORMID = "Apurado Femea Operacoes"
Dim aCAM As Variant
Dim aFOR As Variant
Dim aVAL As Variant
Dim aPAD As Variant
Dim cARQFEMEA As String
Dim cARQPF As String
Dim cSQLUSO As String
Dim nITEM As Long
Dim nCAMPOS As Long

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Encerrar_Click()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If MDG("Gravar e Sair", "Gravando ") Then
    For nITEM = 0 To nCAMPOS - 1
      aVAL(nITEM) = Text(nITEM)
    Next nITEM
    GrvSQL cARQFEMEA, cSQLUSO, nCAMPOS, aCAM, aVAL, aFOR
    GravaLog 999, 0, "Gravou "
  Else
    GravaLog 999, 0, "Nao Gravou "
  End If
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub Form_Load()

  CenterFormToScreen Me
  ''Configura Help
  Me.Caption = cFORMID
  ''HelpContextID = nFORMID

  ''Arquivo
  cARQPF = PegPath("PATH", "PF")
  cARQFEMEA = PegPath("PATH", "FEMEA")
  If Not FileConnExist(cARQFEMEA, True) Then
    End
  End If

  'Variaveis sql
  cSQLUSO = ePASS01
  nCAMPOS = 15
  aCAM = Array("SEQ", "PF", "MES", "ANO", "PRONUM", _
               "TOTCAU", "TOTRPN", "TOTM40", "MAIRPN", "TOTFX01", _
               "TOTFX02", "TOTFX03", "TOTRPN01", "TOTRPN02", "TOTRPN03")
  aFOR = Array("NI", "NI", "NI", "NI", "NI", _
               "NI", "NI", "NI", "NI", "NI", _
               "NI", "NI", "NI", "NI", "NI")
  aPAD = Array(0, 0, 0, 0, 0, _
               0, 0, 0, 0, 0, _
               0, 0, 0, 0, 0)
  aVAL = PegSQL(cARQFEMEA, cSQLUSO, nCAMPOS, aCAM, aFOR, aPAD)
  For nITEM = 0 To nCAMPOS - 1
    Text(nITEM) = aVAL(nITEM)
  Next nITEM
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


