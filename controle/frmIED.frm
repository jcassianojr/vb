VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form frmIED 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IED"
   ClientHeight    =   3045
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   8910
   Icon            =   "frmIED.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   3
      Left            =   2160
      MaxLength       =   50
      TabIndex        =   7
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox txtFields 
      Height          =   1365
      Index           =   2
      Left            =   960
      MaxLength       =   50
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   960
      Width           =   5595
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   1
      Left            =   960
      MaxLength       =   50
      TabIndex        =   3
      Top             =   600
      Width           =   5595
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   0
      Left            =   720
      MaxLength       =   1
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
   Begin XPControls.XPButton cmdClose 
      Height          =   435
      Left            =   7080
      TabIndex        =   8
      Top             =   240
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmIED.frx":058A
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
      Left            =   7080
      TabIndex        =   9
      Top             =   840
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmIED.frx":0B24
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
      Caption         =   "Detalhes"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "Valor"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descritivo"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "TIPO:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmIED"
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
        For iLOOP = 0 To nCAMPOS - 1
            aVAL(iLOOP) = TXTFIELDS(iLOOP)
        Next iLOOP
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
    'CentralizaJanela Me
    cARQ = PegPath("PATH", "PF")
    cSQL = "select * from IED WHERE ITEM=" & iITEM
    nCAMPOS = 4
    aCAM = Array("TIPO", "DESCRITIVO", "DETALHE", "VALOR")
    aFOR = Array("C", "C", "C", "N")
    aPAD = Array("", "", "", 0)
    aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
    For iLOOP = 0 To nCAMPOS - 1
        TXTFIELDS(iLOOP) = aVAL(iLOOP)
    Next iLOOP
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

