VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.1#0"; "vbccr18.ocx"
Begin VB.Form frmBTN 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Botoes"
   ClientHeight    =   2760
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frmBTN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "Liberar  Botão para todos Usuários"
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
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Marque Para que todos usuarios acessem"
      Top             =   2160
      Width           =   3765
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   4
      Left            =   7080
      MaxLength       =   12
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   3
      Left            =   6360
      MaxLength       =   12
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1920
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   2
      Left            =   120
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1800
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   5
      Left            =   120
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1080
      Width           =   3375
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   1
      Left            =   1920
      MaxLength       =   50
      TabIndex        =   1
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   6360
      TabIndex        =   13
      Top             =   720
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmBTN.frx":058A
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
      Left            =   6360
      TabIndex        =   14
      Top             =   240
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmBTN.frx":0B24
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
   Begin VBCCR18.SpinBox indice 
      Height          =   495
      Left            =   3480
      TabIndex        =   15
      Top             =   360
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
   Begin VBCCR18.SpinBox icone 
      Height          =   495
      Left            =   4560
      TabIndex        =   16
      Top             =   360
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
   Begin VB.Label Label1 
      Caption         =   "Nome do Botão"
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
      Index           =   5
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   1965
   End
   Begin VB.Label Label1 
      Caption         =   "ToolTip => o que deve aparecer quando o mouse fica sobre o botão"
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
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   6405
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Ícone"
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
      Left            =   4440
      TabIndex        =   10
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
      Left            =   3360
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Controle"
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
      Left            =   1920
      TabIndex        =   8
      Top             =   120
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Formulario"
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
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1035
   End
End
Attribute VB_Name = "frmBTN"
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
    'ajustar depois para nao precisar txtfiels 3 4
    TXTFIELDS(3) = indice.Value
    TXTFIELDS(4) = icone.Value

    For iLOOP = 0 To nCAMPOS - 2
      aVAL(iLOOP) = TXTFIELDS(iLOOP)
    Next iLOOP
    aVAL(6) = FixNumBol(Check2.Value)
    GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR
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
  cSQL = "select * from controle WHERE FORM='" & cMENU & "' AND CONTROLE='" & cTIPO & "' AND INDICE=" & iMENU
  nCAMPOS = 7
  aCAM = Array("FORM", "CONTROLE", "TOOLTIP", "INDICE", "IMAGEM", "CAPTION", "DISPONIVEL")
  aFOR = Array("C", "C", "C", "NI", "NI", "C", "BN")
  aPAD = Array("", "", "", 0, 0, "", False)
  aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
  For iLOOP = 0 To nCAMPOS - 2
    TXTFIELDS(iLOOP) = aVAL(iLOOP)
  Next iLOOP
  Check2.Value = aVAL(6)
  'ajustar depois para nao precisar txtfiels 3 4
  indice.Value = TXTFIELDS(3)
  icone.Value = TXTFIELDS(4)

End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

'Private Sub maisind_Click()
'    txtFields(3) = CStr(IncDec(txtFields(3), 1, 1, 999))
'End Sub

'Private Sub maismes_Click()
'    txtFields(4) = CStr(IncDec(txtFields(4), 1, 1, 999))
'End Sub

'Private Sub menosind_Click()
'    txtFields(3) = CStr(IncDec(txtFields(3), -1, 1, 999))
'End Sub

'Private Sub menosmes_Click()
'    txtFields(4) = CStr(IncDec(txtFields(4), -1, 1, 999))
'End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index = 3 Or Index = 4 Then
    KeyAscii = ValiText(KeyAscii, "#NI")
  End If
End Sub

