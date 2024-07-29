VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{F22668DE-E08D-467B-8E41-13900013BD5F}#2.7#0"; "VBextra2.OCX"
Object = "{451B73A5-1563-45D5-A6AC-7B2B7D30B778}#1.1#0"; "BSPrin10.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.1#0"; "vbccr18.ocx"
Object = "{2DA70529-3366-414A-B408-46083BCD481B}#1.8#0"; "VBFLXGRD17.OCX"
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
   Begin VBFLXGRD17.VBFlexGrid Grid 
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
'Dim aORDEM As Variant
Dim aORDES As Variant
'Dim cORDEM As String
Dim cSUBWHERE As String
Dim nSUBWHERE As Long

Private Sub btnOpen_Click()
  TxtExecutar.tEXT = OpenArqExt(Me, TxtExecutar.tEXT, "*", "Executarcom")
End Sub

Private Sub CmdAbrirCom_Click()
  cARQRTF = Txtcaminho.tEXT & Trim(Grid)
  If FileExist(cARQRTF, True) Then
    Call OpenWith(cARQRTF, OAIF_ALLOW_REGISTRATION Or OAIF_EXEC Or OAIF_FORCE_REGISTRATION, Me.hWnd)
  End If
End Sub

Private Sub CmdConfImp_Click()
  FrmPrintSetup.Show vbModal, Me
End Sub

Private Sub CmdExcluir_Click()
  Dim cARQUIVO As String
  Grid.Col = 0
  cARQUIVO = Grid
  If Len(cARQUIVO) > 0 Then
    cARQUIVO = Txtcaminho.tEXT & Grid
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
  Txtcaminho.tEXT = PegPath(cCHAVE, UCase(cNOMEFOLHA), " ")
  If Len(Trim(Txtcaminho.tEXT)) = 0 Then
    Txtcaminho.tEXT = PegPath(cCHAVE, "CAMINHO", "C:\TEMP")
  End If

  If InStr(Txtcaminho.tEXT, "%USUARIO%") > 0 Then
    Txtcaminho.tEXT = Replace(Txtcaminho.tEXT, "%USUARIO%", cNOMEFOLHA)
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
  cARQRTF = Txtcaminho.tEXT & Trim(Grid)
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
  cARQUIVO = Txtcaminho.tEXT & Trim(Grid)
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
  Dim cFILENAME As String
  Dim cData As Variant
  Dim lINCLUI As Boolean
  On Error Resume Next

  cDRIVE = Left(Txtcaminho.tEXT, 1)
  ChDrive (cDRIVE)
  cDIRER = Mid(Txtcaminho.tEXT, 3)
  ChDir cDIRER

  Grid.Rows = 1
  Grid.Clear
  Grid.cols = 2
  Grid.Col = 0
  
  Grid.AllowUserResizing = FlexAllowUserResizingBoth
  Grid.ColWidthMax = 3001
  Grid.ColAlignment(0) = FlexAlignmentLeftCenter
  Grid.ColWidth(0) = 3000
  Grid.ColWidth(1) = 2000
  Grid.Col = 0
 Grid.tEXT = "Arquivo"
 Grid.Col = 1
 Grid.tEXT = "Data Hora"
  If Right(Txtcaminho.tEXT, 1) <> "\" Then
    Txtcaminho.tEXT = Txtcaminho.tEXT + "\"
  End If
  Set thisFolder = FS.GetFolder(Txtcaminho.tEXT)
  Set allfiles = thisFolder.Files
  For Each thisFile In allfiles
    cFILENAME = UCase(thisFile.Name)
    cData = Trim(thisFile.DateCreated)
    If InStr(cFILENAME, ".TXT") > 0 Or InStr(cFILENAME, ".LST") > 0 Or Right(cFILENAME, 4) = ".PSR" _
       Or Right(cFILENAME, 4) = ".QRP" Or Right(cFILENAME, 5) = ".HTML" _
       Or Right(cFILENAME, 4) = ".PDF" Or Right(cFILENAME, 4) = ".ZIP" _
       Or Right(cFILENAME, 4) = ".RTF" Then
      lINCLUI = True
      If nSUBWHERE = 1 Then
        If InStr(cFILENAME, UCase(cSUBWHERE)) = 0 Then
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
  cORIGEM = Txtcaminho.tEXT & cNOME
  If InStr(UCase(cNOME), ".TXT") > 0 Or InStr(UCase(cNOME), ".LST") > 0 Then
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
  cORIGEM = Txtcaminho.tEXT & cNOME
  cDESTINO = Txtcaminho.tEXT & NomeArq(cNOME, True) & ".txt"
  nRETU = CopyFileWindowsWay(cORIGEM, cDESTINO)
  If nRETU = 0 And MDG("Apagar arquivo lst") Then
    DeleteFile cORIGEM  'Kill cORIGEM
  End If
  Command1_Click
End Sub

Private Sub Command3_Click()
  Dim DLG As New clsCommonDialogs
  DLG.BrowseFolder "Escolha uma Pasta", Pastas, False
  If DLG.Cancel = True Then Exit Sub
  ' Mostra o caminho escolhido
  Txtcaminho.tEXT = DLG.Path & "\"
  Command1_Click
End Sub

Private Sub Command4_Click()
''Dim carq As String
  Dim cNOME As String
  Grid.Col = 0
  cNOME = Txtcaminho.tEXT & Trim(Grid)
  ShellEx cNOME, essSW_SHOWDEFAULT, , , , Me.hWnd
End Sub

Private Sub Command5_Click()
  Dim cNOME As String
  Dim cORIGEM As String
  Grid.Col = 0
  cNOME = Trim(Grid)
  cORIGEM = Txtcaminho.tEXT & cNOME
  If InStr(UCase(cNOME), ".TXT") > 0 Or InStr(UCase(cNOME), ".LST") > 0 Then
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
  If InStr(UCase(cNOME), ".TXT") > 0 Or InStr(UCase(cNOME), ".LST") > 0 Then
    ePASS01 = Txtcaminho.tEXT & cNOME
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
  'Dim oZIP As cZip
  Grid.Col = 0
  cNOME = Trim(Grid)
  cORIGEM = Txtcaminho.tEXT & cNOME
  If InStr(UCase(cNOME), ".ZIP") > 0 Then
    Alert ("Ja e um zip")
  Else
    Alert "ainda nao implementado"
    ' Set oZIP = New cZip
    'oZIP.ZipFile = cORIGEM & ".ZIP"
    'oZIP.StoreFolderNames = False
    'oZIP.AddFileSpec (cORIGEM)
    'oZIP.Zip
    'Set oZIP = Nothing
    Command1_Click
  End If
End Sub

Private Sub Form_Load()
'Call systray(TrayAdd, Me.Hwnd, Me.Caption, Me.Icon)
  CenterFormToScreen Me
  aORDES = Array("Relatório", "Data")
  cSUBWHERE = ""
  cNOMEFOLHA = NetworkUserName()
  TxtExecutar.tEXT = PegPath("LSTVIEW", "VIEW", "C:\plugin\lstview\log0290.exe")
  CmdIni_Click 0
End Sub

Private Sub Grid_DblClick()
  Dim cARQ As String
  Dim cCOMP As String
  Dim cNOME As String
  Dim cEXTENSAO As String
  
  If Grid.Row = 0 Then
     Exit Sub
  End If
  
  Grid.Col = 0
  cNOME = Trim(Grid)
  cCOMP = UCase(Trim(Grid))
  cARQRTF = Txtcaminho.tEXT & cNOME
  
  Select Case cEXTENSAO
         Case "ZIP"
              Alert ("Sem Visualizacao para Zip")
         Case "PSR"
              cARQ = "rptview " & Txtcaminho.tEXT & cNOME
              Shell cARQ, vbMaximizedFocus
         Case "QRP"
              cARQ = "qrpview " & Txtcaminho.tEXT & cNOME
              Shell cARQ, vbMaximizedFocus
         Case "PDF"
              cARQ = Txtcaminho.tEXT & cNOME
              ShellEx cARQ, essSW_SHOWDEFAULT, , , , Me.hWnd
         Case "HTM"
              If MDG("Usar Interno") Then
                 ePASS03 = 3
                 PrintPreview1.ShowPreview
              Else
                cARQ = Txtcaminho.tEXT & cNOME
                ShellEx cARQ, essSW_SHOWDEFAULT, , , , Me.hWnd
              End If
          Case "RTF"
               ePASS03 = 2
              PrintPreview1.ShowPreview
         Case Else
            'carq = "log0290 " & Txtcaminho.Text & cNOME
              If MDG("Usar Interno") Then
                 ePASS03 = 1
                 PrintPreview1.ShowPreview
              Else
                cARQ = Trim(TxtExecutar.tEXT) & " " & Txtcaminho.tEXT & cNOME
                Shell cARQ, vbMaximizedFocus
              End If
  End Select
  
End Sub
Private Sub PrintPreview1_PrepareReport(Cancel As Boolean)
  If ePASS03 = 1 Then  'TXT
    MyPrintingTXT
  End If
  If ePASS03 = 2 Then  'rtf
    MyPrintingRTF
  End If
  If ePASS03 = 3 Then  'HTML
    MyPrintinghtml
  End If
End Sub
Public Sub MyPrintingTXT()
  Dim fileFile As Integer
  Dim STRBUFFER As String
  'If Not FileExist(cARQRTF, True) Then 'ja checado na cmdvisual click
  '    Exit Sub
  ' End If
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


