VERSION 5.00
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "VBCCR18.OCX"
Begin VB.Form frmcomp 
   Caption         =   "Competencia"
   ClientHeight    =   2784
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   5304
   Icon            =   "FRMCOMP.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2784
   ScaleWidth      =   5304
   StartUpPosition =   3  'Windows Default
   Begin VBCCR18.SpinBox mes 
      Height          =   612
      Left            =   1080
      TabIndex        =   4
      Top             =   120
      Width           =   732
      _ExtentX        =   1291
      _ExtentY        =   1080
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.8
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
      Height          =   612
      Left            =   2880
      TabIndex        =   5
      Top             =   120
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.8
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
      Height          =   612
      Left            =   1080
      TabIndex        =   6
      Top             =   960
      Width           =   732
      _ExtentX        =   1291
      _ExtentY        =   1080
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.8
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
      Height          =   612
      Left            =   1080
      TabIndex        =   7
      Top             =   1800
      Width           =   732
      _ExtentX        =   1291
      _ExtentY        =   1080
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.8
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
   Begin VBCCR18.CommandButtonW cmdOK 
      Height          =   612
      Left            =   4080
      TabIndex        =   8
      Top             =   120
      Width           =   972
      _ExtentX        =   1715
      _ExtentY        =   1080
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ImageListAlignment=   1
      Caption         =   "Confimar"
      Alignment       =   0
      VerticalAlignment=   0
      Picture         =   "FRMCOMP.frx":038A
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Quinzena"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   972
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   732
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Ano"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   252
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   696
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Mês"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   300
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   576
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
