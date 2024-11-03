VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Object = "{451B73A5-1563-45D5-A6AC-7B2B7D30B778}#1.1#0"; "BSPrin10.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.1#0"; "vbccr18.ocx"
Object = "{075212A8-C1CF-444E-939D-F6046CCDBC08}#1.0#0"; "VBFLXGRD18.OCX"
Begin VB.Form FrmLstView 
   Caption         =   "Visualizador de LST"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10365
   Icon            =   "FrmLstView.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   10365
   StartUpPosition =   3  'Windows Default
   Begin vbExtra.CommonDialogEx CommonDialogEx1 
      Left            =   7440
      Top             =   1080
      _ExtentX        =   900
      _ExtentY        =   900
      MaxFileSize     =   255
      FontName        =   ""
   End
   Begin XPControls.XPButton XPButton1 
      Height          =   375
      Left            =   2400
      TabIndex        =   29
      Top             =   1200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "teste print grid preview "
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
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   255
      Left            =   3840
      ScaleHeight     =   195
      ScaleWidth      =   1395
      TabIndex        =   28
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VBFLXGRD18.VBFlexGrid Grid 
      Height          =   5175
      Left            =   120
      TabIndex        =   25
      Top             =   1680
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9128
      FixedRows       =   0
      Rows            =   0
      ColWidthMax     =   333
   End
   Begin vbExtra.FlexFn FlexFn1 
      Height          =   405
      Left            =   120
      TabIndex        =   24
      Top             =   1080
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   714
      PageNumbersFormat=   ""
      PageNumbersFormatIndex=   0
      BeginProperty PageNumbersFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton CmdIni 
      Height          =   255
      Index           =   3
      Left            =   9840
      TabIndex        =   6
      ToolTipText     =   "Caminho Inicial"
      Top             =   300
      Width           =   375
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "C2"
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
   Begin XPControls.XPButton CmdIni 
      Height          =   255
      Index           =   2
      Left            =   9840
      TabIndex        =   5
      ToolTipText     =   "Caminho Inicial"
      Top             =   60
      Width           =   375
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "C1"
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
   Begin XPControls.XPButton CmdIni 
      Height          =   255
      Index           =   0
      Left            =   9480
      TabIndex        =   4
      ToolTipText     =   "Caminho Inicial"
      Top             =   60
      Width           =   375
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "CI"
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
   Begin XPControls.XPButton CmdIni 
      Height          =   255
      Index           =   1
      Left            =   9480
      TabIndex        =   3
      ToolTipText     =   "Caminho Servidor"
      Top             =   300
      Width           =   375
      _ExtentX        =   0
      _ExtentY        =   0
      Caption         =   "CS"
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
   Begin VB.TextBox TxtExecutar 
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Text            =   "c:\plugin\lstview\log0290.exe"
      Top             =   600
      Width           =   7515
   End
   Begin XPControls.XPButton Command2 
      Height          =   315
      Left            =   8040
      TabIndex        =   1
      ToolTipText     =   "Renomear extensao do Arquivo de .LST para .TXT"
      Top             =   4680
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      Caption         =   ".LST -> .TXT"
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
   Begin VB.TextBox Txtcaminho 
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Text            =   "c:\temp"
      Top             =   120
      Width           =   7515
   End
   Begin XPControls.XPButton CmdConfImp 
      Height          =   375
      Left            =   8040
      TabIndex        =   7
      Top             =   6000
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Picture         =   "FrmLstView.frx":058A
      Caption         =   "Configurar Impressora"
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
   Begin XPControls.XPButton CmdVisua 
      Height          =   375
      Left            =   8040
      TabIndex        =   8
      Top             =   2040
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Picture         =   "FrmLstView.frx":0B24
      Caption         =   "Visualizar"
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
   Begin XPControls.XPButton CmdSendMail 
      Height          =   375
      Left            =   8040
      TabIndex        =   9
      Top             =   5040
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Picture         =   "FrmLstView.frx":10BE
      Caption         =   "EnviarEmail"
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
   Begin XPControls.XPButton CmdLocaliza 
      Height          =   375
      Left            =   8040
      TabIndex        =   10
      Top             =   2520
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Picture         =   "FrmLstView.frx":1658
      Caption         =   "Localizar"
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
   Begin XPControls.XPButton CmdPrinterPort 
      Height          =   375
      Left            =   8040
      TabIndex        =   11
      Top             =   5520
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Picture         =   "FrmLstView.frx":1BF2
      Caption         =   "Imprimir na porta"
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
   Begin XPControls.XPButton Command5 
      Height          =   375
      Left            =   8040
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Picture         =   "FrmLstView.frx":218C
      Caption         =   "Html"
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
   Begin XPControls.XPButton Command6 
      Height          =   375
      Left            =   8040
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Picture         =   "FrmLstView.frx":2726
      Caption         =   "PDF"
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
   Begin XPControls.XPButton Command10 
      Height          =   375
      Left            =   9240
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Picture         =   "FrmLstView.frx":2AC0
      Caption         =   "RTF"
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
   Begin XPControls.XPButton Command4 
      Height          =   375
      Left            =   8040
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3000
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Picture         =   "FrmLstView.frx":2E5A
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
   Begin XPControls.XPButton CmdAbrirCom 
      Height          =   375
      Left            =   8040
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3480
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Picture         =   "FrmLstView.frx":33F4
      Caption         =   "Abrir Com"
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
   Begin XPControls.XPButton CmdExcluir 
      Height          =   435
      Left            =   8040
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   767
      Picture         =   "FrmLstView.frx":398E
      Caption         =   "Apagar"
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
   Begin XPControls.XPButton btnOpen 
      Height          =   375
      Left            =   8040
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1080
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Picture         =   "FrmLstView.frx":3F28
      Caption         =   "Ver Com"
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
   Begin XPControls.XPButton Command3 
      Height          =   375
      Left            =   7800
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Picture         =   "FrmLstView.frx":44C2
      Caption         =   "Buscar Pasta"
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
      Left            =   8040
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      Picture         =   "FrmLstView.frx":4A5C
      Caption         =   "Atualizar"
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
   Begin XPControls.XPButton CmdFiltra 
      Height          =   675
      Left            =   8040
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3960
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1191
      Picture         =   "FrmLstView.frx":4FF6
      Caption         =   "filtrar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureAlignment=   2
   End
   Begin XPControls.XPButton Command7 
      Height          =   675
      Left            =   8760
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3960
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1191
      Picture         =   "FrmLstView.frx":5590
      Caption         =   "Dia"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureAlignment=   2
   End
   Begin XPControls.XPButton Command8 
      Height          =   675
      Left            =   9480
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3960
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1191
      Picture         =   "FrmLstView.frx":592A
      Caption         =   "limpar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureAlignment=   2
   End
   Begin BSPrinter.PrintPreview PrintPreview1 
      Left            =   6840
      Top             =   1080
      _ExtentX        =   1191
      _ExtentY        =   1191
   End
   Begin VBCCR18.RichTextBox RichTextBox1 
      Height          =   255
      Left            =   5520
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      MultiLine       =   -1  'True
      TextRTF         =   "FrmLstView.frx":5CC4
   End
   Begin XPControls.XPButton Command9 
      Height          =   375
      Left            =   9240
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6960
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      Caption         =   "ZIP"
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
End
Attribute VB_Name = "FrmLstView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cNOMEFOLHA As String
Dim aORDES As Variant
Dim cSUBWHERE As String
Dim nSUBWHERE As Long

Private Sub btnOpen_Click()
  TxtExecutar.Text = OpenArqExt(Me, TxtExecutar.Text, "*", "Executarcom")
End Sub

Private Sub CmdAbrirCom_Click()
  cARQRTF = Txtcaminho.Text & Trim(Grid)
  If FileConnExist(cARQRTF, True) Then
    Call OpenWith(cARQRTF, OAIF_ALLOW_REGISTRATION Or OAIF_EXEC Or OAIF_FORCE_REGISTRATION, CLng(Me.hWnd))
  End If
End Sub

Private Sub CmdConfImp_Click()
 ' FrmPrintSetup.Show vbModal, Me
End Sub

Private Sub CmdExcluir_Click()
  Dim cARQUIVO As String
  Grid.Col = 0
  cARQUIVO = Grid
  If Len(cARQUIVO) > 0 Then
    cARQUIVO = Txtcaminho.Text & Grid
    If MDG("Excluir " & cARQUIVO) Then
      DeleteFile cARQUIVO  'Kill cARQUIVO
      Command1_Click
    End If
  End If
End Sub

Private Sub CmdFiltra_Click()
  ePASS01 = aORDES
  frmLocalizaa.Show vbModal, Me
  If lRETU Then
    cSUBWHERE = eRETU01
    nSUBWHERE = eRETU02
  Else
    cSUBWHERE = ""
    nSUBWHERE = 0
  End If
  Command1_Click
End Sub

Private Sub CmdFiltro_Click()

End Sub


Private Sub CmdIni_Click(Index As Integer)
  Dim cCHAVE As String
  Select Case Index
  Case 0
    cCHAVE = "LSTVIEW"
  Case 1
    cCHAVE = "SERVER"
  Case 2
    cCHAVE = "USO01"
  Case 3
    cCHAVE = "USO02"
  End Select
  Txtcaminho.Text = PegPath(cCHAVE, UCase(cNOMEFOLHA), " ")
  If Len(Trim(Txtcaminho.Text)) = 0 Then
    Txtcaminho.Text = PegPath(cCHAVE, "CAMINHO", "C:\TEMP")
  End If

  If InStr(Txtcaminho.Text, "%USUARIO%") > 0 Then
    Txtcaminho.Text = Replace(Txtcaminho.Text, "%USUARIO%", cNOMEFOLHA)
  End If
  Command1_Click
End Sub

Private Sub CmdLocaliza_Click()
  ePASS01 = aORDES
  frmLocalizaa.Show vbModal, Me
  If lRETU Then
    LocalizaGrid Grid, eRETU01, eRETU02, , 1
  End If
End Sub

Private Sub CmdPrinterPort_Click()
  Grid.Col = 0
  cARQRTF = Txtcaminho.Text & Trim(Grid)
  If Right(UCase(cARQRTF), 4) = ".PSR" Or Right(UCase(cARQRTF), 4) = ".QRP" Then
    Alert "Somente LST/TXT"
    Exit Sub
  End If
  ''aRELCFG
  ''0 Arquivo
  ''1 Abrir como
  ''2 Mensagem Antes Imprimir
  ''3 Nº Utilizacoes
  ''4 Arvore
  ''5 Busca
  ''6 Titulo do Relatorio
  ''7 Caminho
  ''8 Caminh2
  ''9 Caminh3
  ''10 Caminh4
  ''11 pergunta filtro
  ''12 alias da tabela
  ''13 nome da tabela
  ''14 prefiltro
  ''15 sqluso
  ''16 carquso
  ''17 carqbai
  ''18 carqfec
  ''19 carqacu

  aDIREITOS = Array(True, True, True, True, True, True, True, True)
  aRELCFG = Array("", "", "", 0, False, _
                  False, "", "", "", "", "", _
                  False, "", "", "", "")

  FrmTxl.Show vbModal, Me
End Sub

Private Sub CmdSendMail_Click()
  Dim cARQUIVO As String
  Grid.Col = 0
  cARQUIVO = Txtcaminho.Text & Trim(Grid)
  'servidor,porta,from,to,assunto,anexos,mensagem,enviar e sair
  ePASS01 = Array("", _
                  "", _
                  "", _
                  "", _
                  "", _
                  cARQUIVO, _
                  "", _
                  False)
  FrmSendMail.Show vbModal, Me
End Sub

Private Sub CmdVisua_Click()
  Grid_DblClick
End Sub

Private Sub Command1_Click()
  Dim FS As New FileSystemObject
  Dim thisFile As File
  Dim allfiles As Files
  Dim thisFolder As Folder
  Dim cDRIVE As String
  Dim cDIRER As String
  Dim cFileName As String
  Dim cData As Variant
  Dim lINCLUI As Boolean
  Dim cEXTENSAO As String
  On Error Resume Next
  
  

  cDRIVE = Left(Txtcaminho.Text, 1)
  ChDrive (cDRIVE)
  cDIRER = Mid(Txtcaminho.Text, 3)
  ChDir cDIRER

  Grid.Rows = 1
  Grid.Clear
  Grid.cols = 2
  Grid.Col = 0
  
  Grid.AllowUserResizing = FlexAllowUserResizingBoth
  Grid.ColWidthMax = 3001
  Grid.ColAlignment(0) = FlexAlignmentLeftCenter ' FlexAlignmentCENTERCenter FlexAlignmentRIGHTCenter FlexAlignmentLeftCenter
  Grid.ColWidth(0) = 3000
  Grid.ColWidth(1) = 2000
  Grid.Col = 0
 Grid.Text = "Arquivo"
 Grid.Col = 1
 Grid.Text = "Data Hora"
  If Right(Txtcaminho.Text, 1) <> "\" Then
    Txtcaminho.Text = Txtcaminho.Text + "\"
  End If
  Set thisFolder = FS.GetFolder(Txtcaminho.Text)
  Set allfiles = thisFolder.Files
  For Each thisFile In allfiles
    cFileName = UCase(thisFile.Name)
    cData = Trim(thisFile.DateCreated)
    'LST1 LST2
    If IsExtensao(cFileName, "TXT") Or InStr(cFileName, ".LST") > 0 Or IsExtensao(cFileName, "PSR") _
       Or IsExtensao(cFileName, "QRP") Or IsExtensao(cFileName, "HTML") _
       Or IsExtensao(cFileName, "PDF") Or IsExtensao(cFileName, "ZIP") _
       Or IsExtensao(cFileName, "RTF") Or IsExtensao(cFileName, "JPG") Then
      lINCLUI = True
      If nSUBWHERE = 1 Then
        If InStr(cFileName, UCase(cSUBWHERE)) = 0 Then
          lINCLUI = False
        End If
      End If
      If nSUBWHERE = 2 Then
        If InStr(cData, cSUBWHERE) = 0 Then
          lINCLUI = False
        End If
      End If
      If lINCLUI And thisFile.Size > 0 Then
        ' nome sem upper problemas linux
        Grid.AddItem thisFile.Name & vbTab & cData
      End If
    End If
  Next
End Sub

Private Sub Command10_Click()
  Dim cNOME As String
  Dim cORIGEM As String
  Grid.Col = 0
  cNOME = Trim(Grid)
  cORIGEM = Txtcaminho.Text & cNOME
  If IsExtensao(cNOME, "TXT") Or InStr(UCase(cNOME), ".LST") > 0 Then
    txttoRTF cORIGEM, TrocaExt(cORIGEM, "RTF")
    Command1_Click
  Else
    Alert ("Nao e TXT ou Lst")
  End If
End Sub

Private Sub Command2_Click()
  Dim cORIGEM As String
  Dim cDESTINO As String
  Dim cNOME As String
  Dim nRETU As Integer
  Grid.Col = 0
  cNOME = Trim(Grid)
  cORIGEM = Txtcaminho.Text & cNOME
  cDESTINO = Txtcaminho.Text & NomeArq(cNOME, True) & ".txt"
  nRETU = CopyFileWindowsWay(cORIGEM, cDESTINO)
  If nRETU = 0 And MDG("Apagar arquivo lst") Then
    DeleteFile cORIGEM  'Kill cORIGEM
  End If
  Command1_Click
End Sub

Private Sub Command3_Click()

CommonDialogEx1.ShowFolder
Txtcaminho.Text = FixPath(CommonDialogEx1.FolderName)
'relendo os arquivos da pasta escolhida
Command1_Click
End Sub

Private Sub Command4_Click()
  Dim cNOME As String
  Grid.Col = 0
  cNOME = Txtcaminho.Text & Trim(Grid)
  ShellEx cNOME, essSW_SHOWDEFAULT, , , , CLng(Me.hWnd)
End Sub

Private Sub Command5_Click()
  Dim cNOME As String
  Dim cORIGEM As String
  Grid.Col = 0
  cNOME = Trim(Grid)
  cORIGEM = Txtcaminho.Text & cNOME
  If IsExtensao(cNOME, "TXT") Or InStr(UCase(cNOME), ".LST") > 0 Then
    txttohtml cORIGEM, TrocaExt(cORIGEM, "htm")
    Command1_Click
  Else
    Alert ("Nao e TXT ou Lst")
  End If
End Sub

Private Sub Command6_Click()
  Dim cNOME As String
  Grid.Col = 0
  cNOME = Trim(Grid)
  If IsExtensao(cNOME, "TXT") Or InStr(UCase(cNOME), ".LST") > 0 Then
    ePASS01 = Txtcaminho.Text & cNOME
    formConvertToPDF.Show vbModal, Me
    Command1_Click
  Else
    Alert ("Nao e TXT ou Lst")
  End If
End Sub

Private Sub Command7_Click()
  frmInsertDate.Show vbModal, Me
  If lRETU Then
    cSUBWHERE = eRETU01
    nSUBWHERE = 2
    Command1_Click
  End If
End Sub

Private Sub Command8_Click()
  cSUBWHERE = ""
  nSUBWHERE = 0
  Command1_Click
End Sub


Private Sub Command9_Click()
  Dim cNOME As String
  Dim cORIGEM As String
  Dim cDESTINO As String
  Grid.Col = 0
  cNOME = Trim(Grid)
  cORIGEM = Txtcaminho.Text & cNOME
  If IsExtensao(cNOME, "ZIP") Then
    Alert ("Ja e um zip")
  Else
    cDESTINO = TrocaExt(cORIGEM, "ZIP")
    With New cZipArchive
      .AddFile cORIGEM
       .CompressArchive cDESTINO
    End With
    Command1_Click
  End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub
Private Sub Form_Load()
'Call systray(TrayAdd, clng(Me.hwnd), Me.Caption, Me.Icon)
  CenterFormToScreen Me
  aORDES = Array("Relatório", "Data")
  cSUBWHERE = ""
  cNOMEFOLHA = NetworkUserName()
  TxtExecutar.Text = PegPath("LSTVIEW", "VIEW", "C:\plugin\lstview\log0290.exe")
  CmdIni_Click 0
End Sub

Private Sub Grid_DblClick()
  Dim cARQ As String
  Dim cCOMP As String
  Dim cNOME As String
  Dim cEXTENSAO As String
  
  If Grid.Row = 0 Then
     ePASS03 = 0 'ZERA TIPO DE PREVIEW
     Exit Sub
  End If
  
  Grid.Col = 0
  cNOME = Trim(Grid)
  cCOMP = UCase(Trim(Grid))
  cARQRTF = Txtcaminho.Text & cNOME
  
  cEXTENSAO = EXTENSAO(cARQRTF)
  
  Select Case cEXTENSAO
         Case "ZIP"
              Alert ("Sem Visualizacao para Zip")
         Case "PSR"
              cARQ = "rptview " & Txtcaminho.Text & cNOME
              Shell cARQ, vbMaximizedFocus
         Case "QRP"
              cARQ = "qrpview " & Txtcaminho.Text & cNOME
              Shell cARQ, vbMaximizedFocus
         Case "PDF"
              cARQ = Txtcaminho.Text & cNOME
              ShellEx cARQ, essSW_SHOWDEFAULT, , , , CLng(Me.hWnd)
         Case "HTM"
              If MDG("Usar Interno") Then
                 ePASS03 = 3
                 PrintPreview1.ShowPreview
              Else
                cARQ = Txtcaminho.Text & cNOME
                ShellEx cARQ, essSW_SHOWDEFAULT, , , , CLng(Me.hWnd)
              End If
          Case "RTF"
               ePASS03 = 2
               RichTextBox1.LoadFile cARQRTF, RtfLoadSaveFormatRTF
              PrintPreview1.ShowPreview
          Case "JPG"
               ePASS03 = 4
               Picture1.Picture = LoadPicture(cARQRTF)
              PrintPreview1.ShowPreview
         Case Else
            'carq = "log0290 " & Txtcaminho.Text & cNOME
              If MDG("Usar Interno") Then
                 ePASS03 = 1
                 PrintPreview1.ShowPreview
              Else
                cARQ = Trim(TxtExecutar.Text) & " " & Txtcaminho.Text & cNOME
                Shell cARQ, vbMaximizedFocus
              End If
  End Select
  
End Sub
Public Sub PrintPreview1_PrepareReport(Cancel As Boolean)
  If ePASS03 = 1 Then  'TXT
    MyPrintingTXT
  End If
  If ePASS03 = 2 Then  'rtf
    MyPrintingRTF
  End If
  If ePASS03 = 3 Then  'HTML
    MyPrintinghtml
  End If
  If ePASS03 = 4 Then 'JPG
      ePASS01 = Array("Esquerda Superior", "Direita Superior", "Esquera Inferior", "Direita Inferior", "Centralizado")
      escOrdem.Show vbModal, Me
      ePASS02 = FixInt(eRETU01, 0)
      MyPrintingJPG
  End If
  If ePASS03 = 5 Then  'grid
    MyPrintinggrid
  End If
  ePASS03 = 0
End Sub
Public Sub MyPrintingTXT()
  Dim fileFile As Integer
  Dim STRBUFFER As String
  fileFile = FreeFile
  Open cARQRTF For Input As #fileFile
  Do While Not EOF(fileFile)
    'read line
    Input #fileFile, STRBUFFER
    Printer.Print STRBUFFER
  Loop
  Close fileFile
End Sub
Public Sub MyPrintinghtml()
  Dim cTEXTO As String
  Dim cLINHA As String
  Dim LINES() As String
  Dim i As Integer
  cTEXTO = FileText(cARQRTF)
  cTEXTO = HtmlToText(cTEXTO)
  LINES = Split(cLINHA, vbCrLf)
  For i = 0 To UBound(LINES)
    Printer.Print LINES(i)
  Next
End Sub
Public Sub MyPrintingRTF()
  PrinterEx.PrintRichTextBox RichTextBox1
End Sub
Private Sub Txtcaminho_LostFocus()
  Command1_Click
End Sub
Public Sub MyPrintingJPG()
  Select Case ePASS02
   Case 1
       Printer.PaintPicture Picture1.Picture, 0, 0
   Case 1
       Printer.PaintPicture Picture1.Picture, Printer.ScaleWidth - Picture1.ScaleX(Picture1.Picture.Width, vbHimetric, Printer.ScaleMode), 0
  Case 2
    Printer.PaintPicture Picture1.Picture, 0, Printer.ScaleHeight - Picture1.ScaleY(Picture1.Picture.Height, vbHimetric, Printer.ScaleMode)
  Case 3
    Printer.PaintPicture Picture1.Picture, Printer.ScaleWidth - Picture1.ScaleX(Picture1.Picture.Width, vbHimetric, Printer.ScaleMode), Printer.ScaleHeight - Picture1.ScaleY(Picture1.Picture.Height, vbHimetric, Printer.ScaleMode)

  Case 4
    Printer.PaintPicture Picture1.Picture, (Printer.ScaleWidth - Picture1.ScaleX(Picture1.Picture.Width, vbHimetric, Printer.ScaleMode)) / 2, (Printer.ScaleHeight - Picture1.ScaleY(Picture1.Picture.Height, vbHimetric, Printer.ScaleMode)) / 2

  
  Case Else
  Printer.PaintPicture Picture1.Picture, 0, 0
       
 End Select
End Sub
Private Sub XPButton1_Click()
  ePASS03 = 5
  PrintPreview1.ShowPreview
End Sub
Public Sub MyPrintinggrid()
  PrintPreview1.PrintGrid Grid
End Sub

