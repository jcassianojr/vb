VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.1#0"; "vbccr18.ocx"
Begin VB.Form frmMENU 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MENU"
   ClientHeight    =   1620
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   8145
   Icon            =   "frmMENU.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1620
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   3
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Tag             =   "Posicao Geral do Menu"
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtFields 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   3840
      TabIndex        =   2
      TabStop         =   0   'False
      Tag             =   "Posiçao no Menu"
      Text            =   "0"
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   1
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "Descriçao a ser exibida no Menu"
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Height          =   285
      Index           =   0
      Left            =   1200
      MaxLength       =   12
      TabIndex        =   0
      Tag             =   "Nome do Menu"
      Top             =   120
      Width           =   1575
   End
   Begin VBCCR18.SpinBox indice 
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   120
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
      BackColor       =   12648447
      Min             =   1
      Max             =   999
      Value           =   1
   End
   Begin VBCCR18.SpinBox Cadastro 
      Height          =   495
      Left            =   5520
      TabIndex        =   7
      Top             =   120
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
      Value           =   1
   End
   Begin XPControls.XPButton cmdClose 
      Height          =   435
      Left            =   6360
      TabIndex        =   10
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmMENU.frx":058A
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
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   6360
      TabIndex        =   11
      Top             =   600
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmMENU.frx":0B24
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Cadastro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Índice"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   8
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Descricao"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   720
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   120
      Width           =   1035
   End
End
Attribute VB_Name = "frmMENU"
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
    txtFields(1) = indice.Value
    txtFields(3) = Cadastro.Value

    For iLOOP = 0 To nCAMPOS - 1
      aVAL(iLOOP) = txtFields(iLOOP)
    Next iLOOP
    GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR, 2 ' 0 MENU 1 INDICE CHAVE
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
  cARQ = Dbname
  cSQL = "select * from MENU WHERE MENU='" & cMENU & "' AND INDICE=" & iMENU

  nCAMPOS = 4
  aCAM = Array("menu", "indice", "descricao", "cadastro")
  aFOR = Array("C", "NI", "C", "NI")
  aPAD = Array("", 0, "", 0)
  aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
  For iLOOP = 0 To nCAMPOS - 1
    txtFields(iLOOP) = aVAL(iLOOP)
  Next iLOOP
  'ajustar depois para nao precisar txtfiels 3 4
  indice.Value = txtFields(1)
  Cadastro.Value = txtFields(3)


End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

Private Sub txtFields_GotFocus(Index As Integer)
  FocusMe
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index = 2 Then
    KeyAscii = ValiText(KeyAscii, "#NI")
  End If
End Sub

