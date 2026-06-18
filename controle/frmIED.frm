VERSION 5.00
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "vbccr18.ocx"
Begin VB.Form frmIED 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IED"
   ClientHeight    =   3036
   ClientLeft      =   1092
   ClientTop       =   336
   ClientWidth     =   8916
   Icon            =   "frmIED.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3036
   ScaleWidth      =   8916
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
   Begin VBCCR18.CommandButtonW Encerrar 
      Height          =   612
      Left            =   7440
      TabIndex        =   8
      Top             =   1080
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
      Picture         =   "frmIED.frx":058A
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VBCCR18.CommandButtonW cmdClose 
      Height          =   612
      Left            =   7440
      TabIndex        =   9
      Top             =   360
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
      Picture         =   "frmIED.frx":0B24
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
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

