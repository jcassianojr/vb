VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form frmREVI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Revisâo"
   ClientHeight    =   3330
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   8340
   Icon            =   "frmREVI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFields 
      Height          =   2115
      Index           =   4
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   960
      Width           =   8055
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   615
   End
   Begin XPControls.XPButton cmdClose 
      Height          =   435
      Left            =   6600
      TabIndex        =   10
      Top             =   240
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmREVI.frx":058A
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
      Left            =   4920
      TabIndex        =   11
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmREVI.frx":0B24
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
   Begin VB.Label lblLabels 
      Caption         =   "MOTIVO:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ITEM:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   3
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Caption         =   "TIPO:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Caption         =   "REVISAO:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "PF:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmREVI"
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
Dim i As Integer

Private Sub cmdClose_Click()
  On Error Resume Next
  For i = 0 To nCAMPOS - 1
    aVAL(i) = txtFields(i)
  Next i
  GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR, 4 'pula chaves "PF", "REVISAO", "TIPO", "ITEM"
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub Encerrar_Click()
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  cARQ = PegPath("PATH", "PF")
  nREV = FixInt(nREV, 0)
  nREVI = FixInt(nREVI, 0)
  If nREV = 0 Or nREVI = 0 Then
    Alert "Revisao Ou Item da Revisao zerado"
    lEDITAR = False
  Else
    cSQL = ePASS01
    nCAMPOS = 5
    aCAM = Array("PF", "REVISAO", "TIPO", "ITEM", "MOTIVO")
    aFOR = Array("NI", "NI", "CU", "NI", "C")
    aPAD = Array(0, 0, "", 0, "")
    aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
    For i = 0 To nCAMPOS - 1
      txtFields(i) = aVAL(i)
    Next i
  End If
  If Not lEDITAR Then
    txtFields(4).TabStop = False
    txtFields(4).Locked = True
    txtFields(4).Enabled = False
    cmdClose.Enabled = False
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

