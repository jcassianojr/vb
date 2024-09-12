VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form frmRPTUSR 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Direitos de Uso Relatorio"
   ClientHeight    =   2205
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   8340
   Icon            =   "frmRPTUSR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   Begin XPControls.XPButton CmdLibTudo 
      Height          =   315
      Left            =   6000
      TabIndex        =   21
      Top             =   1080
      Width           =   1095
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Liberar Tudo"
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
   Begin XPControls.XPButton CmdAbrirCom 
      Height          =   375
      Index           =   0
      Left            =   6000
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "Abrir com Somente da Extencao"
      Top             =   1440
      Width           =   495
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Ext..."
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
   Begin XPControls.XPButton CmdAbrirCom 
      Height          =   375
      Index           =   1
      Left            =   6480
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Abrir Com Todas Extensoes"
      Top             =   1440
      Width           =   375
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "*..."
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
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   120
      Width           =   1395
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   600
      Width           =   1395
   End
   Begin VB.TextBox txtFields 
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   4500
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1500
      Width           =   1335
   End
   Begin VB.CheckBox chkFields 
      Caption         =   "Editar"
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   5
      Left            =   120
      TabIndex        =   13
      Top             =   1500
      Width           =   915
   End
   Begin VB.CheckBox chkFields 
      Caption         =   "Novo"
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   6
      Left            =   1140
      TabIndex        =   12
      Top             =   1500
      Width           =   975
   End
   Begin VB.CheckBox chkFields 
      Caption         =   "Abrir"
      ForeColor       =   &H00C00000&
      Height          =   285
      Index           =   7
      Left            =   2280
      TabIndex        =   11
      Top             =   1500
      Width           =   1035
   End
   Begin VB.CheckBox chkFields 
      Caption         =   " Salvar Como"
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   4
      Left            =   4560
      TabIndex        =   10
      Top             =   960
      Width           =   1335
   End
   Begin VB.CheckBox chkFields 
      Caption         =   "Salvar"
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   3
      Left            =   3480
      TabIndex        =   9
      Top             =   960
      Width           =   975
   End
   Begin VB.CheckBox chkFields 
      Caption         =   "Visualiza"
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   8
      Top             =   960
      Width           =   975
   End
   Begin VB.CheckBox chkFields 
      Caption         =   "Exportar"
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   1
      Left            =   1140
      TabIndex        =   7
      Top             =   960
      Width           =   975
   End
   Begin VB.CheckBox chkFields 
      Caption         =   "Imprime"
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   915
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   5400
      Locked          =   -1  'True
      MaxLength       =   8
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   600
      Width           =   1395
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   855
   End
   Begin XPControls.XPButton cmdClose 
      Height          =   435
      Left            =   6720
      TabIndex        =   22
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "frmRPTUSR.frx":038A
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
   Begin VB.Label lblLabels 
      Caption         =   "Codigo"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   17
      Top             =   600
      Width           =   735
   End
   Begin VB.Label lblLabels 
      Caption         =   "Abrir Com"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   15
      Top             =   1500
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "RPT:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   4920
      TabIndex        =   4
      Top             =   600
      Width           =   315
   End
   Begin VB.Label lblLabels 
      Caption         =   "GRP:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Caption         =   "USUARIO:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmRPTUSR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim aCAM, aFOR, aVAL, aPAD As Variant
Attribute aFOR.VB_VarUserMemId = 1073938432
Attribute aVAL.VB_VarUserMemId = 1073938432
Attribute aPAD.VB_VarUserMemId = 1073938432
Dim cARQUSR, cSQL As String
Attribute cARQUSR.VB_VarUserMemId = 1073938436
Attribute cSQL.VB_VarUserMemId = 1073938436
Dim nFORMID As Integer
Attribute nFORMID.VB_VarUserMemId = 1073938438
Dim cARQUIVO As String
Attribute cARQUIVO.VB_VarUserMemId = 1073938439
Dim x As Integer
Attribute x.VB_VarUserMemId = 1073938440

Private Sub CmdAbrirCom_Click(Index As Integer)
'Dim cEXTENSAO As String
'Dim nPOS As Long

  eLOCALIZA = txtFields(5)
  ePASS01 = ""
  If Index = 0 Then
    ePASS01 = EXTENSAO(cARQUIVO)
  End If
  EscRptExec.Show vbModal, Me
  If lRETU Then
    txtFields(5) = eRETU02
  End If
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub CmdLibTudo_Click()
  For x = 0 To 7
    chkFields(x).Value = vbChecked
  Next x
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub Form_Load()

  CenterFormToScreen Me

  cSQL = "select * from " & ArqRPTUsr() & " WHERE GRP='" & zgrp & "' AND RPT='" & zRPT & "' AND IDUSUARIO=" & zIDRPTUSR
  txtFields(0) = zIDRPTUSR
  txtFields(1) = ePASS03
  'txtFields(2) = "" codigo
  txtFields(3) = zgrp
  txtFields(4) = zRPT

  cARQUSR = DBWRPT
  aCAM = Array("IMPRIME", "EXPORTA", "VISUALIZA", "SALVARTF", "SALVATXT", _
               "NOVO", "ABRIR", "Editar", "ABRIRCOM")
  aFOR = Array("BN", "BN", "BN", "BN", "BN", _
               "BN", "BN", "BN", "C")
  aPAD = Array(False, False, False, False, False, _
               False, False, False, "")
  aVAL = PegSQL(cARQUSR, cSQL, 9, aCAM, aFOR, aPAD)
  For x = 0 To 7
    chkFields(x) = aVAL(x)
  Next x
  txtFields(5) = aVAL(8)
  nFORMID = 998
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Dim x As Integer
  If MDG("Gravar e Sair", "Gravando ") Then
    For x = 0 To 7
      aVAL(x) = FixNumBol(chkFields(x))
    Next x
    aVAL(8) = txtFields(5)
    GrvSQL cARQUSR, cSQL, 9, aCAM, aVAL, aFOR
    GravaLog nFORMID, 0, "Gravou "
  Else
    GravaLog nFORMID, 0, "Nao Gravou "
  End If
  Screen.MousePointer = vbDefault
End Sub


