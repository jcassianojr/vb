VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form frmDica 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Frase do dia"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   5820
   ControlBox      =   0   'False
   Icon            =   "frmDica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   5820
   Begin XPControls.XPButton cmdProxima 
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   3240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Picture         =   "frmDica.frx":058A
      Caption         =   "Proxima"
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   120
      Picture         =   "frmDica.frx":0B24
      ScaleHeight     =   2985
      ScaleWidth      =   705
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.CheckBox chkExibir 
      Caption         =   "&Exibir ao iniciar"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   3360
      Width           =   1695
   End
   Begin XPControls.XPButton cmdOK 
      Height          =   435
      Left            =   3960
      TabIndex        =   7
      Top             =   3240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmDica.frx":0C3E
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
   Begin VB.Label lblAutor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   960
      TabIndex        =   6
      Top             =   2850
      Width           =   4575
   End
   Begin VB.Label lblMensagem 
      BackColor       =   &H8000000E&
      Height          =   1995
      Left            =   1080
      TabIndex        =   5
      Top             =   840
      Width           =   4455
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   2415
      Left            =   720
      TabIndex        =   4
      Top             =   720
      Width           =   4935
   End
   Begin VB.Label lblTitulo 
      BackColor       =   &H80000009&
      Caption         =   "Você sabia..."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmDica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sExibir As String * 255
Dim nTOTREG As Long
Dim cARQ As String
Dim x As Long
Private Sub cmdProxima_Click()
  PegMensagem
End Sub
Private Sub cmdOK_Click()
  sExibir = IIf(chkExibir = vbChecked, "S", "N")
  x = WritePrivateProfileString("CITACAO", zUSER, sExibir, App.Path + "\" & App.EXEName & ".INI")
  Unload Me
End Sub

Private Sub PegMensagem()
  Dim cSQL As String
  Dim aRETU As Variant
  lRETU = False
  While Not lRETU
    cSQL = "SELECT tblCitacao.Titulo, tblCitacao.Citacao, tblCitacao.Titulo, tblAutor.Nome as Autor, " & _
           "tblAutor.Nacionalidade, tblOcupacao.Nome As Ocupacao FROM (tblAutor LEFT JOIN tblOcupacao " & _
           "ON tblAutor.codOcupacao = tblOcupacao.codOcupacao) RIGHT JOIN tblCitacao ON tblAutor.codAutor = " & _
           "tblCitacao.codAutor Where Not tblCitacao.Ocultar and codcitacao=" & RandomNumber(nTOTREG)
    aRETU = PegSQL(cARQ, cSQL, 5, Array("titulo", "citacao", "autor", "ocupacao", "nacionalidade"), _
                   Array("CC", "CC", "CC", "CC", "CC"), _
                   Array("", "", "", "", ""))
    lblTitulo = FixStr(aRETU(0))
    lblMensagem = FixStr(aRETU(1))
    lblAutor = FixStr(aRETU(2), "") & ", " & FixStr(aRETU(3), "") & ", " & FixStr(aRETU(4))
  Wend
End Sub

Private Sub Form_Load()
  cARQ = PegPath("PATH", "CITACAO")
  nTOTREG = 0
  If FileExist(cARQ, True) Then
    chkExibir = vbChecked
    CenterFormToScreen Me
    nTOTREG = PegCountSQL(cARQ, "tblcitacao", "codcitacao", 1)
    PegMensagem
  End If
End Sub

Function RandomNumber(intHighestNumber)
  Randomize
  RandomNumber = Int(intHighestNumber * Rnd) + 1
End Function


