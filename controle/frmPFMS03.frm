VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Begin VB.Form frmPFMS03 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Componentes"
   ClientHeight    =   2385
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   8520
   Icon            =   "frmPFMS03.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPControls.XPButton CmdPrd 
      Height          =   255
      Index           =   0
      Left            =   5880
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      Caption         =   "Interno"
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
      Height          =   285
      Index           =   11
      Left            =   4800
      TabIndex        =   9
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   10
      Left            =   4080
      TabIndex        =   8
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Height          =   405
      Index           =   4
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Height          =   405
      Index           =   3
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   7
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   6
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   5
      Left            =   1200
      TabIndex        =   6
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   9
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   8
      Left            =   1080
      MaxLength       =   100
      TabIndex        =   4
      Top             =   1080
      Width           =   6255
   End
   Begin VB.TextBox txtFields 
      Height          =   285
      Index           =   2
      Left            =   1080
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   372
      Index           =   1
      Left            =   2040
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Width           =   372
   End
   Begin VB.TextBox txtFields 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   372
      Index           =   0
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   975
   End
   Begin XPControls.XPButton Encerrar 
      Height          =   435
      Left            =   6720
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmPFMS03.frx":058A
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
      Left            =   6720
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   0
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   767
      Picture         =   "frmPFMS03.frx":0B24
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
   Begin XPControls.XPButton escmu 
      Height          =   375
      Left            =   3600
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "frmPFMS03.frx":10BE
      Caption         =   ""
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
   Begin XPControls.XPButton Command1 
      Height          =   375
      Left            =   4080
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Picture         =   "frmPFMS03.frx":1658
      Caption         =   ""
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
   Begin XPControls.XPButton CmdEditar 
      Height          =   375
      Left            =   4560
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Picture         =   "frmPFMS03.frx":1BF2
      Caption         =   "Editar"
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
   Begin XPControls.XPButton ESCpro 
      Height          =   375
      Left            =   5640
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   600
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Picture         =   "frmPFMS03.frx":218C
      Caption         =   "LX"
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
   Begin XPControls.XPButton CmdPrd 
      Height          =   255
      Index           =   1
      Left            =   5880
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      Caption         =   "Externo"
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
      Caption         =   "Produçâo"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   11
      Left            =   4800
      TabIndex        =   23
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "Opção"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   10
      Left            =   4080
      TabIndex        =   22
      Top             =   1440
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Caption         =   "Ssq"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   9
      Left            =   3720
      TabIndex        =   21
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Caption         =   "Seq"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   8
      Left            =   2520
      TabIndex        =   20
      Top             =   120
      Width           =   375
   End
   Begin VB.Label lblLabels 
      Caption         =   "Baixa"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   7
      Left            =   3360
      TabIndex        =   19
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "Total"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   6
      Left            =   2280
      TabIndex        =   18
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Caption         =   "Preço"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   5
      Left            =   1200
      TabIndex        =   17
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "Quantidade"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Caption         =   "Descrição"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label lblLabels 
      Caption         =   "Codigo:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblLabels 
      Caption         =   "Tipo:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   1440
      TabIndex        =   13
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblLabels 
      Caption         =   "PF:"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "frmPFMS03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cARQPF As String
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
      aVAL(iLOOP) = txtFields(iLOOP)
    Next iLOOP
    GrvSQL cARQPF, cSQL, nCAMPOS, aCAM, aVAL, aFOR, 4    'Pula chaves "PF", "TIPOENT", "CODCOMP", "SEQ", "SSQ",
  End If
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub CmdEditar_Click()
  ePASS01 = txtFields(2).text
  Select Case txtFields(1).text
  Case "T"
    FrmMP03.Show vbModal, Me
  Case Else
    Alert "Disponivel apenas para tipo T"
  End Select
End Sub

Private Sub CmdPrd_Click(Index As Integer)
  Select Case Index
  Case 0
    txtFields(11) = "Int"
  Case 1
    txtFields(11) = "Ext"
  End Select
End Sub

Private Sub Command1_Click()
  Dim cNOME As String
  Dim sSQL As String
  Dim cCODIGO As String
  Dim cARQ As String
  Dim aRETU As Variant


  cCODIGO = FixStr(txtFields(2), "", "TRIM")


  cARQ = GeraConn(zMANA5EMP, "JETFOX")


  Select Case txtFields(1)
  Case "C"
    sSQL = "SELECT NOME,NOM2 FROM MT01 WHERE CODIGO='" & cCODIGO & "'"
    aRETU = PegSQL(cARQ, sSQL, 2, Array("NOME", "NOM2"), Array("C", "C"), Array("", ""))
    If lRETU Then
      cNOME = aRETU(0) + " " + aRETU(1)
    End If
  Case "M"
    sSQL = "SELECT NOME,NOM2 FROM MU01 WHERE CODIGO='" & cCODIGO & "'"
    aRETU = PegSQL(cARQ, sSQL, 2, Array("NOME", "NOM2"), Array("C", "C"), Array("", ""))
    If lRETU Then
      cNOME = aRETU(0) + " " + aRETU(1)
    End If

  Case "T"
    sSQL = "SELECT NOME,NOM2 FROM mP03 WHERE CODIGO='" & cCODIGO & "'"
    aRETU = PegSQL(cARQ, sSQL, 2, Array("NOME", "NOM2"), Array("C", "C"), Array("", ""))
    If lRETU Then
      cNOME = aRETU(1)                     ''Nome completo
    End If
  Case "S"
    sSQL = "SELECT NOME,NOM2 FROM mQ01 WHERE CODIGO='" & cCODIGO & "'"
    aRETU = PegSQL(cARQ, sSQL, 2, Array("NOME", "NOM2"), Array("C", "C"), Array("", ""))
    If lRETU Then
      cNOME = aRETU(0) + " " + aRETU(1)
    End If
    If Len(cNOME) = 0 Then                   ''Tenta Produtos
      sSQL = "SELECT NOME FROM mS01 WHERE CODIGO='" & cCODIGO & "'"
      aRETU = PegSQL(cARQ, sSQL, 2, Array("NOME", "NOM2"), Array("C", "C"), Array("", ""))
      If lRETU Then
        cNOME = aRETU(0) + " " + aRETU(1)
      End If
    End If
  End Select
  txtFields(3) = FixStr(cNOME)
End Sub

Private Sub Encerrar_Click()
  If Not MDG("Sair sem gravar") Then
    Exit Sub
  End If
  Screen.MousePointer = vbDefault
  Unload Me
End Sub

Private Sub escmu_Click()

  iMU01 = 0
  Select Case txtFields(1)
  Case "M"
    iMU01 = 1
  Case "C"
    iMU01 = 2
  Case "T"
    iMU01 = 5
  Case "S"
    iMU01 = 7
  End Select
  If iMU01 > 0 Then
    escmu01.Show vbModal, Me
    If Not lRETU Then Exit Sub
    frmPFMS03.txtFields(2) = eRETU01
    Command1_Click
  End If


End Sub

Private Sub ESCpro_Click()
  Dim cCHAVEBUS As String
  ''2 codigo
  ePASS01 = "LOGIC"                            ''
  escms01.Show vbModal, Me
  If lRETU Then
    txtFields(2) = eRETU01
    txtFields(3) = eRETU02
  End If
  cCHAVEBUS = ""
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
  CenterFormToScreen Me
  cARQPF = PegPath("PATH", "PF")
  cSQL = "select * from PFMS03 WHERE PF=" & nPF & " AND TIPOENT='" & Ctipoent & "' AND CODCOMP='" & Ccodcomp & "' and seq=" & FixInt(nSEQ) & " and ssq=" & FixInt(nSSQ)
  nCAMPOS = 12
  aCAM = Array("PF", "TIPOENT", "CODCOMP", "SEQ", _
               "SSQ", "PRECO", "TOTAL", "BAIXAC", _
               "DESCRI", "QTDE", "OPCAO", "PRDORI")
  aFOR = Array("NI", "C", "C", "NI", _
               "NI", "N", "N", "C", _
               "C", "N", "NI", "C")
  aPAD = Array(0, "", 0, 0, _
               0, 0, 0, "", _
               "", 0, 0, "")

  aVAL = PegSQL(cARQPF, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
  For iLOOP = 0 To nCAMPOS - 1
    txtFields(iLOOP) = aVAL(iLOOP)
  Next iLOOP


End Sub

Private Sub TXTFIELDS_Change(Index As Integer)
  Select Case Index
  Case 4, 5
    txtFields(6) = FixNum(txtFields(4)) * FixNum(txtFields(5))
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault

End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
  Dim KeyPressed As Variant
  Select Case Index
  Case 8, 9, 10
    KeyAscii = ValiText(KeyAscii, "#NI")
  Case 4, 5
    KeyAscii = ValiText(KeyAscii, "#N")
  End Select
End Sub


