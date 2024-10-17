VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.1#0"; "vbccr18.ocx"
Begin VB.Form frmcomp 
   Caption         =   "Competencia"
   ClientHeight    =   3210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5100
   Icon            =   "frmcomp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin VBCCR18.SpinBox mes 
      Height          =   615
      Left            =   1560
      TabIndex        =   4
      Top             =   120
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
      Min             =   1
      Max             =   12
      Value           =   1
   End
   Begin VBCCR18.SpinBox ano 
      Height          =   615
      Left            =   1320
      TabIndex        =   5
      Top             =   840
      Width           =   975
      _ExtentX        =   1720
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
      Min             =   1950
      Max             =   2200
      Value           =   1950
      ThousandsSeparator=   0   'False
   End
   Begin VBCCR18.SpinBox empresa 
      Height          =   615
      Left            =   1560
      TabIndex        =   6
      Top             =   1560
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
      Min             =   1
      Max             =   999
      Value           =   1
      AllowOnlyNumbers=   -1  'True
   End
   Begin VBCCR18.SpinBox quinzena 
      Height          =   615
      Left            =   1560
      TabIndex        =   7
      Top             =   2280
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
      Min             =   1
      Max             =   2
      Value           =   1
   End
   Begin XPControls.XPButton cmdOK 
      Height          =   435
      Left            =   3000
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   240
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmcomp.frx":038A
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
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Quinzena"
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
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Empresa"
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
      Height          =   300
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Ano"
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
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   690
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Mês"
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
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   270
      Width           =   690
   End
End
Attribute VB_Name = "frmcomp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const nFORMID = 1020
Const cFORMID = "Competencia"
Option Explicit
Private Sub ano_GotFocus()
   FocusMe
End Sub

Private Sub cmdOK_Click()
   eRETU01 = FixInt(ano)
   eRETU02 = FixInt(mes)
   eRETU03 = FixInt(empresa)
   eRETU04 = FixInt(quinzena)
   Unload Me
End Sub
Private Sub empresa_GotFocus()
   FocusMe
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   TeclaEnter (KeyCode)
End Sub


Private Sub Form_Load()
   CenterFormToScreen Me
   Me.Caption = cFORMID
   Me.HelpContextID = nFORMID
   ano = zANO
   mes = zMES
   empresa = zEMPRESA
End Sub
Private Sub ano_KeyPress(KeyAscii As Integer)
   KeyAscii = ValiText(KeyAscii, "#NI")
End Sub

Private Sub empresa_KeyPress(KeyAscii As Integer)
   KeyAscii = ValiText(KeyAscii, "#NI")
End Sub
Private Sub mes_GotFocus()
   FocusMe
End Sub
Private Sub QUINZENA_GotFocus()
   FocusMe
End Sub

Private Sub mes_KeyPress(KeyAscii As Integer)
   KeyAscii = ValiText(KeyAscii, "#NI")
End Sub

Private Sub quinzena_KeyPress(KeyAscii As Integer)
   KeyAscii = ValiText(KeyAscii, "#NI")
End Sub
