VERSION 5.00
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{451B73A5-1563-45D5-A6AC-7B2B7D30B778}#1.1#0"; "BSPrin10.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.1#0"; "vbccr18.ocx"
Object = "{075212A8-C1CF-444E-939D-F6046CCDBC08}#1.0#0"; "VBFLXGRD18.OCX"
Begin VB.Form FrmTxl 
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11760
   Icon            =   "FrmTxl.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   11760
   StartUpPosition =   3  'Windows Default
   Begin VBFLXGRD18.VBFlexGrid Listview1 
      Height          =   1335
      Left            =   120
      TabIndex        =   31
      Top             =   6000
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   2355
   End
   Begin BSPrinter.PrintPreview PrintPreview1 
      Left            =   10200
      Top             =   2640
      _ExtentX        =   1191
      _ExtentY        =   1191
   End
   Begin VBCCR18.RichTextBox RichTextBox1 
      Height          =   255
      Left            =   10080
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   450
      MultiLine       =   -1  'True
      TextRTF         =   "FrmTxl.frx":058A
   End
   Begin VB.Frame Destino 
      Caption         =   "Destino"
      Height          =   1455
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   9735
      Begin VB.TextBox TxtCaminho 
         Height          =   285
         Left            =   2040
         TabIndex        =   12
         Text            =   "TxtCaminho"
         Top             =   960
         Width           =   7575
      End
      Begin VB.OptionButton OptDestino 
         Caption         =   "&Rede Caminho"
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton OptDestino 
         Caption         =   "&Printer Objeto"
         Height          =   375
         Index           =   6
         Left            =   7320
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton OptDestino 
         Caption         =   "&Arquivo(Acima)"
         Height          =   375
         Index           =   5
         Left            =   5640
         TabIndex        =   9
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton OptDestino 
         Caption         =   "&COM"
         Height          =   375
         Index           =   3
         Left            =   2280
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton OptDestino 
         Caption         =   "&LPT"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VBCCR18.SpinBox LptPort 
         Height          =   375
         Left            =   1080
         TabIndex        =   32
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
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
         Max             =   9
         Value           =   1
      End
      Begin VBCCR18.SpinBox comport 
         Height          =   375
         Left            =   3240
         TabIndex        =   33
         Top             =   240
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
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
         Max             =   9
         Value           =   1
      End
   End
   Begin VB.TextBox TxtArquivo 
      Height          =   405
      Left            =   960
      TabIndex        =   5
      Text            =   "TxtArquivo"
      Top             =   1800
      Width           =   10695
   End
   Begin VB.CheckBox CmdEmail 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Enviar Email"
      Height          =   195
      Left            =   1920
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   1035
      Left            =   4320
      TabIndex        =   4
      Top             =   720
      Width           =   7305
      Begin XPControls.XPButton cmdimp 
         Height          =   675
         Left            =   1680
         TabIndex        =   16
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1191
         Picture         =   "FrmTxl.frx":06F0
         Caption         =   "Imprimir (Destino)"
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
         Height          =   675
         Left            =   4440
         TabIndex        =   17
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1191
         Picture         =   "FrmTxl.frx":0C8A
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
      Begin XPControls.XPButton CmdConfImp 
         Height          =   675
         Left            =   120
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1191
         Picture         =   "FrmTxl.frx":1224
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
      Begin XPControls.XPButton Encerrar 
         Height          =   675
         Left            =   5880
         TabIndex        =   25
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1191
         Picture         =   "FrmTxl.frx":17BE
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
      Begin XPControls.XPButton XPButton1 
         Height          =   675
         Left            =   3120
         TabIndex        =   30
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   1191
         Picture         =   "FrmTxl.frx":1D58
         Caption         =   "Imprimir (Impressora)"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exportar/Salvar Como"
      Height          =   1035
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   4095
      Begin XPControls.XPButton Salvar 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   360
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Picture         =   "FrmTxl.frx":22F2
         Caption         =   "TXT"
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
      Begin XPControls.XPButton Salvar 
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Picture         =   "FrmTxl.frx":268C
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
      Begin XPControls.XPButton Salvar 
         Height          =   375
         Index           =   2
         Left            =   2160
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   360
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Picture         =   "FrmTxl.frx":2C26
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
      Begin XPControls.XPButton Salvar 
         Height          =   375
         Index           =   3
         Left            =   3120
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   360
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Picture         =   "FrmTxl.frx":2FC0
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
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Height          =   675
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   11475
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Label1 - Nome do Relatorio que Estamos Imprimindo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   8775
      End
   End
   Begin XPControls.XPButton CmdFiltro 
      Height          =   435
      Left            =   10080
      TabIndex        =   18
      Top             =   3960
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   767
      Picture         =   "FrmTxl.frx":335A
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
   End
   Begin XPControls.XPButton CmdOrdem 
      Height          =   435
      Left            =   10080
      TabIndex        =   19
      Top             =   4920
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   767
      Picture         =   "FrmTxl.frx":38F4
      Caption         =   "ordernar"
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
      Height          =   315
      Left            =   960
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2280
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Picture         =   "FrmTxl.frx":3E8E
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
      Height          =   315
      Left            =   2040
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      Picture         =   "FrmTxl.frx":4428
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
   Begin XPControls.XPButton CmdShell 
      Height          =   315
      Left            =   3480
      TabIndex        =   28
      Top             =   2280
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      Caption         =   "S"
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
   Begin VB.Label Label2 
      Caption         =   "Arquivo"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1800
      Width           =   615
   End
   Begin VB.Label filtro 
      Caption         =   "filtro"
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   4080
      Width           =   9735
      WordWrap        =   -1  'True
   End
   Begin VB.Label Lblsql 
      Caption         =   "sql"
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   5040
      Width           =   9735
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "FrmTxl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const nFORMID = 1181
Const cFORMID = "Imprimir Relatorios Script de Texto"
Dim oTXT As New Criatxt
Dim nQTDELAY As Long
Dim nLINHAS As Long
Dim nCOLUNAS As Long
Dim cDESTINO As String
Dim aLAY As Variant
Dim oCONN As New ADODB.Connection
Dim oRS As New ADODB.Recordset
Dim cSQL As String
Dim cARQ As String
Dim cSETUP As String
Dim larqtxt As Boolean

Private Sub gerar(ByVal cOPE As String)
  Dim x As Long
  Dim cLINHA As String
  Dim aUSO As Variant
  Dim aREG As Variant
  Dim cVAL As String
  'Dim oREG As Variant
  Dim cREG As String
  cLINHA = ""

  oTXT.nCOL = nCOLUNAS
  oTXT.nROW = nLINHAS
  oTXT.Destino = cDESTINO
  If cOPE = "A" Or cOPE = "P" Then
    oTXT.Destino = "ARQ"
    oTXT.arquivo = TxtArquivo
  End If
  If cDESTINO = "ARQ" Then
    oTXT.arquivo = TxtArquivo
  End If
  If cDESTINO = "IMPNET" Then
    oTXT.caminho = TxtCaminho
  End If
  If cSETUP <> "" Then
    oTXT.impsetup cSETUP
  End If
  oTXT.ABRIR
  cARQ = GeracArq(cARQ, , False)
  oCONN.ConnectionTimeout = 120
  oCONN.Open cARQ
  oRS.Open cSQL, oCONN, adOpenForwardOnly, adLockReadOnly
  While Not oRS.EOF
    For x = 1 To nQTDELAY
      cLINHA = aLAY(x)
      Select Case cLINHA
      Case "[NF]"
        oTXT.novafolha
      Case "[GF]"
        oTXT.gravafolha
      Case "{MN}"
        If Not oRS.EOF Then
          oRS.MoveNext
        End If
      Case ""                              ''EM BRANCO NADA FAZ
      Case Else
        aUSO = Split(cLINHA, "|")
        If Left(aUSO(2), 1) = "{" Then
          If Not oRS.EOF Then
            aUSO(2) = Mid(aUSO(2), 2, Len(aUSO(2)) - 2)
            aREG = Split(aUSO(2), "_")

            cVAL = aREG(2)
            cREG = FixStr(oRS(aREG(1)))
            aUSO(2) = FVar(cREG, cVAL)
          Else
            aUSO(2) = ""
          End If
        End If
        oTXT.gravalincon aUSO(2), aUSO(0), aUSO(1)
      End Select
    Next x
  Wend
  oTXT.fechar
  oRS.Close
  oCONN.Close
End Sub

Private Sub CmdAbrirCom_Click()
  cARQRTF = TxtArquivo.tEXT
  If FileConnExist(cARQRTF, True) Then
    Call OpenWith(cARQRTF, OAIF_ALLOW_REGISTRATION Or OAIF_EXEC Or OAIF_FORCE_REGISTRATION, CLng(Me.hWnd))
  End If
End Sub

Private Sub CmdConfImp_Click()
'  FrmPrintSetup.Show vbModal, Me
End Sub

Private Sub CmdEditar_Click()
  If IsExtensao(TxtArquivo.tEXT, "RTF") Or larqtxt Then
    cARQRTF = TxtArquivo.tEXT
    FrmRtfView.Show vbModal, Me
  End If
End Sub

Private Sub CmdFiltro_Click()
  On Error GoTo errhandler
  Dim cFILTRO As String

  If aRELCFG(11) Then
    ePASS01 = ""
    FrmFiltro.Show vbModal, Me
    filtro = Replace(Replace(eRETU01, "{", ""), "}", "")
  End If
  cFILTRO = FixStr(filtro)
  If Len(aRELCFG(15)) > 0 Then
    cSQL = TrocaSqlWhere(aRELCFG(15), FixStr(filtro))
    Lblsql = cSQL
  End If

  Exit Sub

errhandler:
  Select Case Err.Number
  Case 20515
    SayErro "Filtro"

  Case Else
    SayErro "Filtro"
  End Select

End Sub

Private Sub cmdimp_Click()
  If larqtxt Then
    imptxt  'Aqui e direct print com1,COM2.. lpt1,LPT2.. no pode ser usado preview aqui
    Exit Sub
  End If
  If IsExtensao(TxtArquivo.tEXT, "PDF") Or IsExtensao(TxtArquivo.tEXT, "HTML") Or IsExtensao(TxtArquivo.tEXT, "RTF") Then
    CmdVisua_Click
    Exit Sub
  End If
End Sub

Private Sub imptxt()  'Aqui e direct print com1,COM2.. lpt1,LPT2.. no pode ser usado preview aqui
  Dim cLINHA As String
  Dim fileFile As Integer
  Dim STRBUFFER As String

  If Not FileConnExist(cARQRTF, True) Then
    Exit Sub
  End If
  cLINHA = ""
  fileFile = FreeFile + 1
  oTXT.Destino = cDESTINO
  oTXT.ABRIR
  If cDESTINO = "IMPNET" Then
    oTXT.caminho = TxtCaminho
  End If
  If cSETUP <> "" Then
    oTXT.impsetup cSETUP
  End If
  Open cARQRTF For Input As #fileFile
  Do While Not EOF(fileFile)
    Input #fileFile, STRBUFFER
    cLINHA = STRBUFFER
    oTXT.implinha (cLINHA)
  Loop
  oTXT.fechar
  Close fileFile
End Sub

Private Sub CmdOrdem_Click()
  ePASS01 = cARQ
  ePASS02 = NomeTableSql(cSQL)
  EscArqOrdem.Show vbModal, Me
  If lRETU Then
    cSQL = TrocaSQLOrder(cSQL, FixStr(eRETU01))
    Lblsql = cSQL
  End If
End Sub


Private Sub CmdShell_Click()
  Dim cARQSHELL As String
  cARQSHELL = FixStr(TxtArquivo)
  ShellEx cARQSHELL, essSW_SHOWDEFAULT, , , , CLng(Me.hWnd)
End Sub

Private Sub CmdVisua_Click()
  cARQRTF = TxtArquivo.tEXT
  If Not FileConnExist(cARQRTF, True) Then
    Exit Sub
  End If
  If larqtxt Then
    ePASS03 = 1
    PrintPreview1.ShowPreview
  End If
  If IsExtensao(cARQRTF, "PDF") Then
    ShellEx cARQRTF, essSW_SHOWDEFAULT, , , , CLng(Me.hWnd)
  End If
  If IsExtensao(cARQRTF, "HTML") Then
    ePASS01 = Array("Navegador Externo", "Preview Interno", "Navegador Interno")
    escOrdem.Show vbModal, Me
    eRETU01 = FixInt(eRETU01, 0)
    Select Case eRETU01
    Case 0
      OpenUrl (cARQRTF)
    Case 1
      ePASS03 = 3
      PrintPreview1.ShowPreview
    Case 3
      FrmPreview.Show vbModal, Me
    End Select
  End If
  If IsExtensao(cARQRTF, "RTF") Then
    RichTextBox1.LoadFile cARQRTF, RtfLoadSaveFormatRTF  'rtfRTF
    ePASS03 = 2
    PrintPreview1.ShowPreview
    RichTextBox1.tEXT = ""
  End If
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

Public Sub MyPrintinghtml()
  Dim cTEXTO As String
  Dim cLINHA As String
  Dim LINES() As String
  Dim i As Integer

  ' If Not FileConnExist(cARQRTF, True) Then 'ja checado cmdvisualclick
  '     Exit Sub
  ' End If
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
Public Sub MyPrintingTXT()
  Dim fileFile As Integer
  Dim STRBUFFER As String
  'If Not FileConnExist(cARQRTF, True) Then 'ja checado na cmdvisual click
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
Private Sub Encerrar_Click()
  Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
  
End Sub

Private Sub Form_Load()
  Dim x As Long
  Dim cTMP As String
  Dim cARQINI As String
  CenterFormToScreen Me
  filtro = ""
  TxtCaminho = ""
  TxtCaminho.Enabled = False
  OptDestino(6).Value = True
  montaimp

  If IsExtensao(cARQRTF, "TXT") Or IsExtensao(cARQRTF, "ZPL ") Or IsExtensao(cARQRTF, "MAN") Then
    larqtxt = True
  End If
  If InStr(UCase(cARQRTF), ".LST") > 0 Then    ''.LST .LTS2 .LTS3... pode ser sequencial usando instr inves isextensao
    larqtxt = True
  End If


  cARQINI = TrocaExt(cARQRTF, "INI")
  TxtArquivo = cARQRTF
  
  'Pega configuracoes zpl etiquetas outros
  If FileExists(cARQINI) Then
    
    
    nQTDELAY = PegINIVAL(cARQINI, "CONFIGURACAO", "QTDELAY", 1)
    ReDim aLAY(nQTDELAY + 1)
    For x = 1 To nQTDELAY
      aLAY(x) = PegINIVAL(cARQINI, "LAYOUT", StrZero(x, 3), "")
    Next x
    nCOLUNAS = PegINIVAL(cARQINI, "CONFIGURACAO", "COLUNAS", 80)
    nLINHAS = PegINIVAL(cARQINI, "CONFIGURACAO", "LINHAS", 60)
    cDESTINO = PegINIVAL(cARQINI, "CONFIGURACAO", "DESTINO", "ARQ")
    TxtCaminho = PegINIVAL(cARQINI, "CONFIGURACAO", "CAMINHO", "LPT1")

    If Mid(cDESTINO, 1, 3) = "LPT" Then
       OptDestino(0).Value = True
       LptPort.Value = Mid(cDESTINO, 4)
    End If
    
    If Mid(cDESTINO, 1, 3) = "COM" Then
       OptDestino(3).Value = True
       comport.Value = Mid(cDESTINO, 4)
    End If
    

    Select Case cDESTINO
        Case "ARQ", "ARQUIVO"
          OptDestino(5).Value = True
        Case "PRINTER"
          OptDestino(6).Value = True
        Case "IMPNET"
          OptDestino(7).Value = True
          TxtCaminho.Enabled = True
    End Select


    cSETUP = PegINIVAL(cARQINI, "CONFIGURACAO", "SETUP", "")
    TxtArquivo = PegINIVAL(cARQINI, "CONFIGURACAO", "NOMEARQUIVO", "TXL_" & zUSER & Format(Now, "_DDMMYYYY_HHMMSS") & ".TXT")
    If aRELCFG(14) = "" Then
      aRELCFG(14) = PegINIVAL(cARQINI, "CONFIGURACAO", "FILTRO", 1)
    End If
    filtro = aRELCFG(14)
    If aRELCFG(15) = "" Then
      aRELCFG(15) = PegINIVAL(cARQINI, "CONFIGURACAO", "SQL", 1)
    End If
    cSQL = aRELCFG(15)
    If aARQUIVOS(0) = "" Then
      cTMP = PegINIVAL(cARQINI, "CONFIGURACAO", "ARQ", 1)
      If Left(cTMP, 3) = "[T]" Then
        cTMP = PegTable("TABELAS", cTMP)
      End If
      If Left(cTMP, 3) = "[P]" Then
        cTMP = PegPath("PATH", cTMP)
      End If
      aARQUIVOS(0) = cTMP
    End If
    cARQ = Caminex(aARQUIVOS(0))

  End If


  ''Configura Help
  Me.Caption = cFORMID
  HelpContextID = nFORMID



  If Len(aRELCFG(6)) = 0 Then
    Label1 = aRELCFG(0)
  Else
    Label1 = aRELCFG(6)
  End If
  If Len(Label1.Caption) = 0 Then
    Label1.Caption = cARQRTF
  End If


  ' 0-Novo,1-Abrir,2-Salvar,3-Salvacomo
  ' 4-Editar,5-Visualizar,6-Imprimir,7-Exportar

  If larqtxt Then
    aRELCFG(11) = False
    CmdOrdem.Visible = False
    TxtArquivo.Locked = True
    cDESTINO = "LPT1"
    OptDestino(5).Enabled = False
    OptDestino(0).Value = True
  End If

  If Not aDIREITOS(6) Then cmdimp.Visible = False
  If Not aDIREITOS(6) Then CmdConfImp.Visible = False
  If Not aDIREITOS(5) Then CmdVisua.Visible = False
  If Not aDIREITOS(7) Then Salvar(0).Visible = False
  If Not aDIREITOS(7) Then CmdEmail.Visible = False
  If Not aRELCFG(11) Then CmdFiltro.Visible = False
  If Not aRELCFG(11) Then filtro.Visible = False

  If Not aDIREITOS(4) Then CmdEditar.Visible = False
  If Not aDIREITOS(4) Then CmdShell.Visible = False

  aRELCFG(14) = FixStr(aRELCFG(14))
  If Len(aRELCFG(14)) > 0 Then
  Else
    If aRELCFG(11) Then
      CmdFiltro_Click
    End If
  End If
  If Len(aRELCFG(15)) > 0 Then
    cSQL = TrocaSqlWhere(aRELCFG(15), FixStr(filtro))
  End If

  Lblsql = cSQL

  If Not larqtxt Then
    If aRELCFG(13) = "" Then                 'nome tabela nao preenchido pega do sql
      aRELCFG(13) = NomeTableSql(cSQL)
    End If
  Else
    TxtArquivo.Enabled = False
    '    Salvar(0).Enabled = False
  End If
  PrintPreview1.AuxiliaryButtonVisible = PrintPreview1.PrinterExists("Microsoft Print to PDF")
  PrintPreview1.AuxiliaryButtonToolTipText = "Salvar como PDF"

End Sub

Public Sub PrintPreview1_AuxiliaryButtonClick(UpdateReport As Boolean)
  PrintPreview1.ShowSaveToFile "Microsoft Print to PDF", "*.pdf"
  UpdateReport = False  ' we don't need to update the report in the Print preview window after this action (the default value of UpdateReport parameter is True)
End Sub

Private Sub Listview1_Click()
  Dim nPOS As Long
  Dim cPRINTER As String
  cPRINTER = Listview1
  nPOS = InStr(cPRINTER, "|")
  If nPOS = 0 Then
    cPRINTER = Trim(Mid(cPRINTER, nPOS + 1, 40))
  Else
    cPRINTER = Trim(cPRINTER)
  End If
  SelectPrinter (cPRINTER)
End Sub

Private Sub OptDestino_Click(Index As Integer)
  TxtCaminho.Enabled = False
  If Index = 0 Then
     cDESTINO = "LPT" & LptPort.Value
  End If
  If Index = 3 Then
     cDESTINO = "COM" & comport.Value
  End If
  Select Case Index
      Case 5
        cDESTINO = "ARQ"
      Case 6
        cDESTINO = "PRINTER"
      Case 7
        cDESTINO = "IMPNET"
        TxtCaminho.Enabled = True
  End Select
End Sub

Private Sub Salvar_Click(Index As Integer)
  On Error Resume Next
  Dim cARQUIVO As String
  Dim cEXTENSAO As String
  Dim sFILTER As String

  If larqtxt Then
    cARQUIVO = TxtArquivo
  Else
    Select Case Index
    Case 0
      cEXTENSAO = "TXT"
    Case 1
      cEXTENSAO = "HTML"
    Case 2
      cEXTENSAO = "PDF"
    Case 3
      cEXTENSAO = "RTF"

    End Select

    sFILTER = "Formato (*." & cEXTENSAO & ")" & vbNullChar & "*." & cEXTENSAO
    cARQUIVO = FileSave(Me, sFILTER, 1, cEXTENSAO, TxtArquivo.tEXT, App.Path, "Salvar " & cEXTENSAO & " Como")

    If InStr(cARQUIVO, ".") > 0 Then
      cARQUIVO = Left(cARQUIVO, InStr(cARQUIVO, ".") - 1) + "." & cEXTENSAO
    Else
      cARQUIVO = cARQUIVO & cEXTENSAO
    End If

    If Len(cARQUIVO) > 0 Then
    Else
      Alert ("Arquivo Nao Preenchido")
      Exit Sub
    End If


    TxtArquivo = cARQUIVO

    Select Case Index
    Case 0
      gerar "A"
    Case 1
      TxtArquivo = TrocaExt(cARQUIVO, "TMP")
      gerar "A"
      txttohtml TxtArquivo, cARQUIVO
    Case 2
      TxtArquivo = TrocaExt(cARQUIVO, "TMP")
      gerar "A"
      ePASS01 = TxtArquivo
      formConvertToPDF.Show
    Case 3
      TxtArquivo = TrocaExt(cARQUIVO, "TMP")
      gerar "A"
      txttoRTF TxtArquivo, cARQUIVO
    End Select
  End If

  If CmdEmail.Value = 1 Then
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
  End If

End Sub

Private Sub montaimp()
  Dim xPrinter As Printer
  'Dim sIcone As String
  Dim sDesc As String
  Dim sDriver As String
  Dim idx As Integer
  Dim cPAD As String
  On Error Resume Next

  Listview1.Clear
  Listview1.FixedCols = 0
  Listview1.Rows = 1
  Listview1.Row = 0
  Listview1.cols = 3
  Listview1.ColWidth(0) = 3000
  Listview1.ColWidth(1) = 3000
  Listview1.ColWidth(2) = 3000
  Listview1.Col = 0
  Listview1 = "Impressora"
  Listview1.Col = 1
  Listview1 = "Porta"
  Listview1.Col = 2
  Listview1 = "Driver"


  For Each xPrinter In Printers

    cPAD = " "
    ''If xPrinter = Printer.DeviceName Then 'É a Impressora Padrão
    ''      sIcone = ImageList1.ListImages(1).key
    ''   cPAD = "*"
    ''Else 'Outras Impressoras Instaladas
    ''      sIcone = ImageList1.ListImages(2).key
    ''End If

    sDesc = xPrinter.DeviceName
    sDriver = xPrinter.DriverName

    Listview1.AddItem sDesc & vbTab & xPrinter.Port & vbTab & sDriver
    ''ListView1.ListItems.Add , , sDesc, sIcone, sIcone
    ''Listview1.ListItems.Add , , sDesc
    ''Listview1.ListItems(idx).SubItems(lvCad.ColumnHeaders(2).SubItemIndex) = sDriver

    idx = idx + 1

  Next xPrinter

End Sub

Private Function SelectPrinter(ByVal Nome As String) As Boolean
  Dim x As Printer
  Dim nLEN As Long
  Nome = UCase(Trim(Nome))
  nLEN = Len(Nome)
  For Each x In Printers
    If UCase(Mid(x.DeviceName, 1, nLEN)) = Nome Then
      Set Printer = x
      SelectPrinter = True
      Exit For
    End If
  Next
  SelectPrinter = False
End Function


Private Sub XPButton1_Click()
      ePASS01 = Array("Preview", "Imprimir Direto Impressora", "Preview Interno")
      escOrdem.Show vbModal, Me
      eRETU01 = FixInt(eRETU01, 0)
      Select Case eRETU01
         Case 0
           CmdVisua_Click
         Case 1
           FrmPicturePrinter.Show vbModal, Me
         Case 2
           FrmRtfView.Show vbModal, Me
        End Select
End Sub
