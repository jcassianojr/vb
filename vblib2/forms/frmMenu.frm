VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Editando Menus"
   ClientHeight    =   7005
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   10290
   HelpContextID   =   7
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   10290
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdCancelar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Retornar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   8280
      Picture         =   "frmMenu.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1890
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Height          =   330
      Left            =   6000
      Picture         =   "frmMenu.frx":0694
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Escolher Nº Formulario Lista"
      Top             =   930
      Width           =   390
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   5220
      TabIndex        =   3
      ToolTipText     =   "Nº do Formulario Associado"
      Top             =   930
      Width           =   750
   End
   Begin VB.CommandButton CmdEscForm 
      Height          =   330
      Left            =   4395
      Picture         =   "frmMenu.frx":0A1E
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Escolher Nº Formulario Lista"
      Top             =   900
      Width           =   390
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   3630
      TabIndex        =   2
      ToolTipText     =   "Nº do Formulario Associado"
      Top             =   900
      Width           =   750
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2490
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Indice do Menu"
      Top             =   900
      Width           =   750
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   855
      MaxLength       =   120
      TabIndex        =   4
      ToolTipText     =   "Dizer do Menu"
      Top             =   1785
      Width           =   6375
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   870
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Codigo do Menu"
      Top             =   900
      Width           =   1395
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FF8080&
      Caption         =   "&Gravar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   8280
      Picture         =   "frmMenu.frx":0DA8
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Encerra e Grava"
      Top             =   780
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "Módulo"
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
      Index           =   4
      Left            =   5220
      TabIndex        =   11
      Top             =   660
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Formulário"
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
      Left            =   3630
      TabIndex        =   9
      Top             =   630
      Width           =   1185
   End
   Begin VB.Label Label1 
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
      Left            =   2490
      TabIndex        =   8
      Top             =   630
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Nome"
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
      Left            =   945
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Código"
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
      Left            =   900
      TabIndex        =   6
      Top             =   660
      Width           =   735
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const nFORMID = 1007
Const cFORMID = "Editando Menus"
Dim aCAM, aFOR, aVAL, aPAD As Variant
Dim cARQ, cSQL As String
Dim nITEM, nCAMPOS As Long

Private Sub CmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub
Private Sub Command1_Click()
   eLOCALIZA = Text(4)
   aTABELA = Array(1048, "Modulos/Sistemas Documentaçao", "MODULOS", 8, 1049, "SYSDOC")
   EscNumNome.Show vbModal, Me
   If lRETU Then
      Text(4) = eRETU01
   End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If MDG("Gravar e Sair", "Gravando Menus") Then
        For nITEM = 0 To nCAMPOS - 1
            aVAL(nITEM) = UCase(Text(nITEM))
        Next nITEM
        GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR
        GravaLog nFORMID, 0, "Gravou Menu"
    Else
        GravaLog nFORMID, 0, "Nao Gravou Menu"
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdEscform_Click()
    ePASS01 = FixNum(Text(4))
    eLOCALIZA = Text(3)
    EscForm.Show vbModal, Me
    If lRETU Then
        Text(3) = FixNum(eRETU01)
    End If
   
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    TeclaEnter KeyCode
End Sub
Private Sub Form_Load()
   
    ''Configura Help
    Me.Caption = cFORMID
    HelpContextID = nFORMID
   
    ''Arquivo
    cARQ = PegPath("PATH", "SYSCONF")
    If Not FileExist(cARQ, True) Then
        End
    End If
   
    'Variaveis sql
    cSQL = ePASS01
    nCAMPOS = 5
    aCAM = Array("MENU", "DIZER", "INDICE", "FORM", "MODULO")
    aFOR = Array("C", "C", "N", "N", "N")
    aPAD = Array("", "", 0, 0, 0)
    aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
    For nITEM = 0 To nCAMPOS - 1
        Text(nITEM) = aVAL(nITEM)
    Next nITEM
      
   
End Sub

Private Sub Text_GotFocus(Index As Integer)
FocusMe
End Sub

Private Sub Text_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Or KeyCode = 40 Then
        SendKeys Chr(9)
    End If
    If KeyCode = 38 Then
        SendKeys "+" + Chr(9)
    End If
End Sub

