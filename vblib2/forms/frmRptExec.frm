VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.3#0"; "VBCCR18.OCX"
Begin VB.Form frmRptExec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de Execuçao Gerenciador Relatorio/Documentos"
   ClientHeight    =   2496
   ClientLeft      =   1092
   ClientTop       =   336
   ClientWidth     =   8244
   HelpContextID   =   7
   Icon            =   "frmRptExec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2496
   ScaleWidth      =   8244
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdescexe 
      Caption         =   "Executavel"
      Height          =   375
      Left            =   3480
      TabIndex        =   12
      ToolTipText     =   "Escolhe um Programa Externo"
      Top             =   720
      Width           =   975
   End
   Begin XPControls.XPButton CmdEscform 
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      ToolTipText     =   "Escolher um Formulario Para Execuçao"
      Top             =   720
      Width           =   975
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Formulario"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton CmdShell 
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      ToolTipText     =   "Utiliza Configurado Sistema Operacional"
      Top             =   720
      Width           =   1215
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "Conf.Sis.Opr."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   3
      Left            =   4800
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Codigo do Menu"
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text 
      DataField       =   "NOME"
      Height          =   885
      Index           =   4
      Left            =   600
      MaxLength       =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      ToolTipText     =   "Dizer do Menu"
      Top             =   1440
      Width           =   7335
   End
   Begin VB.TextBox Text 
      Height          =   285
      Index           =   1
      Left            =   1710
      TabIndex        =   1
      ToolTipText     =   "Nº do Formulario Associado"
      Top             =   360
      Width           =   2955
   End
   Begin VB.TextBox Text 
      DataField       =   "NOME"
      Height          =   285
      Index           =   2
      Left            =   225
      MaxLength       =   120
      TabIndex        =   2
      ToolTipText     =   "Dizer do Menu"
      Top             =   1080
      Width           =   6375
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Codigo do Menu"
      Top             =   360
      Width           =   1335
   End
   Begin XPControls.XPButton CmdEscExt 
      Height          =   375
      Left            =   6180
      TabIndex        =   13
      Top             =   240
      Width           =   375
      _ExtentX        =   656
      _ExtentY        =   656
      Picture         =   "frmRptExec.frx":058A
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VBCCR18.CommandButtonW cmdClose 
      Height          =   612
      Left            =   6960
      TabIndex        =   14
      Top             =   240
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
      Picture         =   "frmRptExec.frx":0B24
      PictureAndCaption=   -1  'True
      WordWrap        =   0   'False
      Style           =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Extensao"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   4
      Left            =   4830
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Obs"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   2
      Left            =   210
      TabIndex        =   7
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "Nome"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   3
      Left            =   1755
      TabIndex        =   5
      Top             =   45
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Executar"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Numero"
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   270
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmRptExec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const nFORMID = 1064
Const cFORMID = "Tipos de Execuçao Gerenciador Formularios/Relatórios"
Dim aCAM As Variant
Dim aFOR As Variant
Dim aVAL As Variant
Dim aPAD As Variant
Dim cARQ As String
Dim cSQL As String
Dim nITEM As Long
Dim nCAMPOS As Long
Option Explicit

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If MDG("Gravar e Sair", "Gravando ") Then
    For nITEM = 0 To nCAMPOS - 1
      aVAL(nITEM) = tEXT(nITEM)
    Next nITEM
    GrvSQL cARQ, cSQL, nCAMPOS, aCAM, aVAL, aFOR, 1 'comeca com 1 numero e chave
    GravaLog nFORMID, 0, "Gravou "
  Else
    GravaLog nFORMID, 0, "Nao Gravou "
  End If
  Screen.MousePointer = vbDefault
End Sub

Private Sub cmdescexe_Click()
  Dim sRECENTFILE As String
  Dim sPath As String
  Dim sFILTER As String
  Dim sFileName As String
  sRECENTFILE = tEXT(2)
  sPath = tEXT(2)
  sFILTER = "Executáveis" & vbNullChar & "*.exe" & vbNullChar & "Todos os Arquivos" & vbNullChar & "*.*"
  sFileName = FileOpen(Me, sFILTER, 1, sRECENTFILE, "ini", sPath, "Open Ini File")
  If Not Len(sFileName) = 0 Then
    tEXT(2) = NomeArq(sFileName)
  End If
End Sub

Private Sub CmdEscExt_Click()
  eLOCALIZA = tEXT(3)
  aTABELA = Array(1061, "Tipos Arquivos (Extensoes)", "TIPOSARQ", 8, 1062, "SYSCONF")
'  EscCodNome.Show vbModal, Me
'  If lRETU Then
'    Text(3) = eRETU02
'  End If
End Sub

Private Sub CmdEscform_Click()
  eLOCALIZA = tEXT(2)
  ePASS01 = 0
  'EscForm.Show vbModal, Me
  'If lRETU Then
  '  Text(2) = FixStr(eRETU04)
  'End If

End Sub

Private Sub CmdShell_Click()
  tEXT(2) = "SHELL"
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode

End Sub

Private Sub Form_Load()

  CenterFormToScreen Me
  ''Configura Help
  Me.Caption = cFORMID
  HelpContextID = nFORMID

  ''Arquivo
  cARQ = PegPath("PATH", "SYSCONF")
  If Not FileConnExist(cARQ, True) Then
    End
  End If

  'Variaveis sql
  cSQL = ePASS01
  nCAMPOS = 5
  aCAM = Array("NUMERO", "NOME", "EXECUTAR", "EXTENSAO", "OBS")
  aFOR = Array("N", "C", "C", "C", "C")
  aPAD = Array(0, "", "", "", "")
  aVAL = PegSQL(cARQ, cSQL, nCAMPOS, aCAM, aFOR, aPAD)
  For nITEM = 0 To nCAMPOS - 1
    tEXT(nITEM) = aVAL(nITEM)
  Next nITEM

End Sub

