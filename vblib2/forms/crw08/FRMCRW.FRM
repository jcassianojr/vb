VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XpControls.ocx"
Object = "{451B73A5-1563-45D5-A6AC-7B2B7D30B778}#1.1#0"; "BSPrin10.ocx"
Object = "{379157C5-E9BD-43F1-9F83-B037496BED42}#1.1#0"; "vbccr18.ocx"
Begin VB.Form FrmCrw 
   Caption         =   "Crystal Report Relatorio"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11865
   Icon            =   "FRMCRW.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   11865
   WindowState     =   2  'Maximized
   Begin BSPrinter.PrintPreview PrintPreview1 
      Left            =   10080
      Top             =   2520
      _ExtentX        =   1191
      _ExtentY        =   1191
   End
   Begin VBCCR18.RichTextBox RichTextBox1 
      Height          =   375
      Left            =   10200
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      MultiLine       =   -1  'True
      Text            =   "FRMCRW.frx":058A
      TextRTF         =   "FRMCRW.frx":05C2
   End
   Begin XPControls.XPButton CmdShell 
      Height          =   315
      Left            =   10080
      TabIndex        =   29
      Top             =   1920
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
   Begin VB.TextBox TxtArquivo 
      Height          =   285
      Left            =   240
      TabIndex        =   15
      Text            =   "TxtArquivo"
      Top             =   1920
      Width           =   7455
   End
   Begin VB.CheckBox CmdEmail1 
      Caption         =   "Enviar Email (Winsock)"
      Height          =   555
      Left            =   120
      TabIndex        =   14
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox TxtTexto 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   2640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Filtro"
      Top             =   4800
      Width           =   7500
   End
   Begin VB.TextBox TxtDestino 
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   4320
      Width           =   5775
   End
   Begin VB.TextBox TxtAssunto 
      Height          =   375
      Left            =   2640
      TabIndex        =   11
      Top             =   3840
      Width           =   5775
   End
   Begin VB.CheckBox CmdEmail 
      Caption         =   "Enviar Email MAPI"
      Height          =   435
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox TxtComp 
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   3360
      Width           =   6135
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   675
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   11235
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Label1 - Nome do Relatório que Estamos Imprimindo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   90
         TabIndex        =   5
         Top             =   240
         Width           =   11055
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Exportar/Salvar Como"
      Height          =   1155
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   4815
      Begin VB.CommandButton CmdExport 
         Caption         =   "Exportar Outros Formatos"
         Height          =   735
         Left            =   3600
         TabIndex        =   3
         Top             =   240
         Width           =   945
      End
      Begin XPControls.XPButton salvar 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   240
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Picture         =   "FRMCRW.frx":0742
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
      Begin XPControls.XPButton salvar 
         Height          =   375
         Index           =   1
         Left            =   960
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Picture         =   "FRMCRW.frx":0ADC
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
      Begin XPControls.XPButton salvar 
         Height          =   375
         Index           =   2
         Left            =   1800
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   240
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Picture         =   "FRMCRW.frx":0E76
         Caption         =   "doc"
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
      Begin XPControls.XPButton salvar 
         Height          =   375
         Index           =   3
         Left            =   2640
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   240
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Picture         =   "FRMCRW.frx":1410
         Caption         =   "xls"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   1155
      Left            =   5880
      TabIndex        =   1
      Top             =   720
      Width           =   5505
      Begin XPControls.XPButton CmdVisua 
         Height          =   435
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
         Picture         =   "FRMCRW.frx":19AA
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
      Begin XPControls.XPButton cmdimp 
         Height          =   435
         Left            =   1680
         TabIndex        =   17
         Top             =   120
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
         Picture         =   "FRMCRW.frx":1F44
         Caption         =   "Imprimir"
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
         Height          =   435
         Left            =   3240
         TabIndex        =   18
         Top             =   600
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "FRMCRW.frx":24DE
         Caption         =   "Propriedades"
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
         Left            =   3240
         TabIndex        =   19
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   767
         Picture         =   "FRMCRW.frx":2A78
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
      Begin XPControls.XPButton Command2 
         Height          =   435
         Left            =   1680
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   600
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   767
         Picture         =   "FRMCRW.frx":3012
         Caption         =   "selecionar"
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
   Begin VB.TextBox FILTRO 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Filtro"
      Top             =   2280
      Width           =   7500
   End
   Begin Crystal.CrystalReport RPT 
      Left            =   11160
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowRefreshBtn=   -1  'True
   End
   Begin XPControls.XPButton CmdFiltro 
      Height          =   435
      Left            =   7800
      TabIndex        =   20
      Top             =   2400
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   767
      Picture         =   "FRMCRW.frx":35AC
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
   Begin XPControls.XPButton CmdEditar 
      Height          =   315
      Left            =   7680
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1920
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   556
      Picture         =   "FRMCRW.frx":3B46
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
      Left            =   8760
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      Picture         =   "FRMCRW.frx":40E0
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
   Begin XPControls.XPButton CmdMudaFec 
      Height          =   495
      Left            =   8880
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   873
      Picture         =   "FRMCRW.frx":467A
      Caption         =   "Competência "
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
   Begin XPControls.XPButton CmdPreview 
      Height          =   315
      Left            =   10560
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1920
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   556
      Picture         =   "FRMCRW.frx":4A14
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
   Begin VB.Label Label4 
      Caption         =   "Texto"
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1560
      TabIndex        =   9
      Top             =   4680
      Width           =   1035
   End
   Begin VB.Label Label3 
      Caption         =   "Destinatario(s)"
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1500
      TabIndex        =   8
      Top             =   4320
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Assunto"
      ForeColor       =   &H00C00000&
      Height          =   315
      Left            =   1500
      TabIndex        =   7
      Top             =   3720
      Width           =   1095
   End
End
Attribute VB_Name = "FrmCrw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const nFORMID = 1068
Const cFORMID = "Imprimir Relatorios Crystal-OCX"
Option Explicit

Private Sub CmdAbrirCom_Click()
  cARQRTF = TxtArquivo.tEXT
  If FileExist(cARQRTF, True) Then
    Call OpenWith(cARQRTF, OAIF_ALLOW_REGISTRATION Or OAIF_EXEC Or OAIF_FORCE_REGISTRATION, Me.hWnd)
  End If
End Sub

Private Sub CmdConfImp_Click()
 ' FrmPrintSetup.Show vbModal, Me
End Sub

Private Sub CmdEditar_Click()
  If IsExtensao(TxtArquivo.tEXT, "RTF") Or IsExtensao(TxtArquivo.tEXT, "TXT") Then
    cARQRTF = TxtArquivo
    FrmRTf.Show
  End If
End Sub

Private Sub CmdEmail_Click()
  If CmdEmail.Value = 1 Then
    TxtAssunto.Enabled = True
    TxtTexto.Enabled = True
    TxtDestino.Enabled = True
    CmdEmail1.Value = 0
  Else
    TxtAssunto.Enabled = False
    TxtTexto.Enabled = False
    TxtDestino.Enabled = False
  End If
End Sub

Private Sub CmdEmail1_Click()
  If CmdEmail1.Value = 1 Then
    TxtAssunto.Enabled = True
    TxtTexto.Enabled = True
    TxtDestino.Enabled = True
    CmdEmail.Value = 0
  Else
    TxtAssunto.Enabled = False
    TxtTexto.Enabled = False
    TxtDestino.Enabled = False
  End If
End Sub

Private Sub CmdExport_Click()
  Alert ("Em Implantaçao")
End Sub

Private Sub CmdFiltro_Click()
  Dim cFILTRO As String

  If aRELCFG(11) Then
    ePASS01 = "CRYSTAL"
    FrmFiltro.Show vbModal, Me
    FILTRO = eRETU01
  End If
  cFILTRO = CStr(FILTRO.tEXT)

  If Len(cFILTRO) > 0 Then
    RPT.ReplaceSelectionFormula (cFILTRO)
  End If
  Exit Sub

errhandler:

  Select Case Err.Number
  Case 20515
    SayErro "Filtro"
    RPT.ReplaceSelectionFormula ("")
  Case Else
    SayErro "Filtro"
  End Select
End Sub

Private Sub cmdimp_Click()
  On Error Resume Next
  RPT.Destination = crptToWindow
  RPT.PrintReport
End Sub

Private Sub CmdMudaFec_Click()
Dim nARQUIVOS As Integer
Dim x As Integer

  Mana5Fec
  TxtComp.tEXT = aARQFEC(4) & " - " & aARQUIVOS(0)
  With RPT
    nARQUIVOS = UBound(aARQUIVOS)
    For x = 0 To nARQUIVOS
      If Len(aARQUIVOS(x)) > 0 Then
        If Not FileExist(aARQUIVOS(x)) Then
          Alert "Arquivo nao Encontrado: " + aARQUIVOS(x)
          Exit Sub
        End If
        .DataFiles(x) = aARQUIVOS(x)
      End If
    Next x
  End With
End Sub

Private Sub CmdPreview_Click()
  cARQRTF = TxtArquivo.tEXT
  If Not FileExist(cARQRTF, True) Then
    Exit Sub
  End If
  If IsExtensao(cARQRTF, "TXT") Then
    ePASS03 = 1
    PrintPreview1.ShowPreview
  End If
  If IsExtensao(cARQRTF, "PDF") Then
    ShellEx cARQRTF, essSW_SHOWDEFAULT, , , , Me.hWnd
  End If
  If IsExtensao(cARQRTF, "HTML") Then
    'FrmPreview.Show vbModal, Me
    If MDG("Sim->Navegador Nao->Visualizador Interno") Then
      OpenUrl (cARQRTF)
    Else
      ePASS03 = 3
      PrintPreview1.ShowPreview
    End If
  End If
  If IsExtensao(cARQRTF, "RTF") Then
    cARQRTF = TxtArquivo.tEXT
    RichTextBox1.LoadFile cARQRTF, RtfLoadSaveFormatRTF
    ePASS03 = 2
    PrintPreview1.ShowPreview
    RichTextBox1.tEXT = ""
  End If
End Sub
Public Sub MyPrintinghtml()
  Dim cTEXTO As String
  Dim cLINHA As String
  Dim LINES() As String
  Dim i As Integer

  ' If Not FileExist(cARQRTF, True) Then 'ja checado cmdvisualclick
  '     Exit Sub
  ' End If
  cTEXTO = FileText(cARQRTF)
  cTEXTO = HtmlToText(cTEXTO)

  LINES = Split(cLINHA, vbCrLf)

  For i = 0 To UBound(LINES)
    Printer.Print LINES(i)
  Next
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

Public Sub MyPrintingRTF()
  PrinterEx.PrintRichTextBox RichTextBox1
End Sub
Public Sub MyPrintingTXT()
  Dim fileFile As Integer
  Dim STRBUFFER As String
  'If FileExist(cARQRTF, True) Then 'ja checado cmdvisualclick
  '    Exit Sub
  'End If
  fileFile = FreeFile
  Open cARQRTF For Input As #fileFile
  Do While Not EOF(fileFile)
    'read line
    Input #fileFile, STRBUFFER
    Printer.Print STRBUFFER
  Loop
  Close fileFile
End Sub
Private Sub CmdShell_Click()
  Dim cARQSHELL As String
  cARQSHELL = FixStr(TxtArquivo)
  ShellEx cARQSHELL, essSW_SHOWDEFAULT, , , , Me.hWnd
End Sub

Private Sub CmdVisua_Click()
  On Error GoTo erro
  RPT.Destination = crptToWindow
  RPT.PrintReport
  Exit Sub
erro:
  SayErro
End Sub

Private Sub Command2_Click()
  RPT.PrinterSelect
End Sub

Private Sub Encerrar_Click()
  Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  TeclaEnter KeyCode
End Sub

Private Sub Form_Load()
  Dim x, nARQUIVOS As Integer
  Dim cFILTRO As String
  ''Configura Help
  CenterFormToScreen Me
  Me.Caption = cFORMID
  HelpContextID = nFORMID

  TxtArquivo = zUSER & Format(Now, "_DDMMYYYY_HHMMSS") & ".TXT"

  TxtComp.Enabled = False
  CmdMudaFec.Enabled = False

  TxtComp.Visible = False
  CmdMudaFec.Visible = False


  If Len(aARQFEC(0)) > 0 Or Len(aARQFEC(1)) > 0 Or Len(aARQFEC(2)) Or Len(aARQFEC(3)) Then
    TxtComp.Visible = True
    CmdMudaFec.Visible = True
    CmdMudaFec.Enabled = True
    TxtComp.tEXT = aARQFEC(4) & " - " & aARQUIVOS(0)
  End If

  If Len(aRELCFG(6)) = 0 Then
    Label1 = aRELCFG(0)
  Else
    Label1 = aRELCFG(6)
  End If

  ' 0-Novo,1-Abrir,2-Salvar,3-Salvacomo
  ' 4-Editar,5-Visualizar,6-Imprimir,7-Exportar

  If Not aDIREITOS(6) Then cmdimp.Visible = False
  If Not aDIREITOS(6) Then CmdConfImp.Visible = False
  If Not aDIREITOS(5) Then CmdVisua.Visible = False
  If Not aDIREITOS(7) Then salvar(2).Visible = False
  If Not aDIREITOS(7) Then salvar(3).Visible = False
  If Not aDIREITOS(7) Then salvar(0).Visible = False
  If Not aDIREITOS(7) Then salvar(1).Visible = False
  If Not aDIREITOS(7) Then CmdExport.Visible = False

  If Not aDIREITOS(7) Then CmdEmail1.Visible = False
  If Not aDIREITOS(7) Then CmdEmail.Visible = False
  If Not aDIREITOS(7) Then TxtAssunto.Visible = False
  If Not aDIREITOS(7) Then TxtDestino.Visible = False
  If Not aDIREITOS(7) Then TxtTexto.Visible = False
  If Not aDIREITOS(7) Then Label2.Visible = False
  If Not aDIREITOS(7) Then Label3.Visible = False
  If Not aDIREITOS(7) Then Label4.Visible = False

  If Not aDIREITOS(4) Then CmdEditar.Visible = False
  If Not aDIREITOS(4) Then CmdShell.Visible = False


  CmdEmail_Click

  If Not aRELCFG(11) Then CmdFiltro.Visible = False
  If Not aRELCFG(11) Then FILTRO.Visible = False



  With RPT

    .ReportFileName = cARQRTF

    .WindowShowPrintBtn = aDIREITOS(6)
    .WindowShowPrintSetupBtn = aDIREITOS(6)
    .WindowShowExportBtn = aDIREITOS(7)
    .WindowShowGroupTree = aRELCFG(4)
    .WindowShowSearchBtn = aRELCFG(5)
    If Len(aRELCFG(6)) > 0 Then .ReportTitle = aRELCFG(6)



    nARQUIVOS = UBound(aARQUIVOS)

    For x = 0 To nARQUIVOS
      If Len(aARQUIVOS(x)) > 0 Then
        If Not FileExist(aARQUIVOS(x)) Then
          Alert "Arquivo nao Encontrado: " + aARQUIVOS(x)
          Exit Sub
        End If
        .DataFiles(x) = aARQUIVOS(x)
      End If
    Next x

  End With

  aRELCFG(14) = FixStr(aRELCFG(14))
  If Len(aRELCFG(14)) > 0 Then
    cFILTRO = aRELCFG(14)
    FILTRO = cFILTRO

    RPT.ReplaceSelectionFormula (cFILTRO)
  Else
    If aRELCFG(11) Then
      CmdFiltro_Click
    End If
  End If
  PrintPreview1.AuxiliaryButtonVisible = PrintPreview1.PrinterExists("Microsoft Print to PDF")
  PrintPreview1.AuxiliaryButtonToolTipText = "Salvar como PDF"
End Sub
Public Sub PrintPreview1_AuxiliaryButtonClick(UpdateReport As Boolean)
  PrintPreview1.ShowSaveToFile "Microsoft Print to PDF", "*.pdf"
  UpdateReport = False  ' we don't need to update the report in the Print preview window after this action (the default value of UpdateReport parameter is True)
End Sub
Private Sub imprimir_Click()
  RPT.Destination = crptToPrinter
  RPT.PrintReport
End Sub

Private Sub Salvar_Click(Index As Integer)
  Dim cARQUIVO, cEXTENSAO As String
  Dim sFILTER As String

  If CmdEmail.Value = 1 Then
    RPT.Destination = crptMapi
    RPT.EMailSubject = FixStr(TxtAssunto)
    RPT.EMailToList = FixStr(TxtDestino)
    RPT.EMailMessage = FixStr(TxtTexto)
  Else
    RPT.Destination = crptToFile
  End If


  Select Case Index
  Case 0
    RPT.PrintFileType = crptRTF
  Case 1
    RPT.PrintFileType = crptText
  Case 2
    RPT.PrintFileType = crptWinWord
  Case 3
    RPT.PrintFileType = crptExcel50
  End Select



  Select Case Index
  Case 0
    RPT.PrintFileType = crptRTF
    cEXTENSAO = "RTF"
  Case 1
    RPT.PrintFileType = crptText
    cEXTENSAO = "TXT"
  Case 2
    RPT.PrintFileType = crptWinWord
    cEXTENSAO = "DOC"
  Case 3
    RPT.PrintFileType = crptExcel50
    cEXTENSAO = "XLS"
  End Select

  sFILTER = "Formato (*." & cEXTENSAO & ")" & vbNullChar & "*." & cEXTENSAO
  cARQUIVO = FileSave(Me, sFILTER, 1, cEXTENSAO, TxtArquivo.tEXT, App.Path, "Salvar " & cEXTENSAO & " Como")


  If InStr(cARQUIVO, ".") > 0 Then
    cARQUIVO = Left(cARQUIVO, InStr(cARQUIVO, ".") - 1) + "." & cEXTENSAO
  Else
    cARQUIVO = cARQUIVO & cEXTENSAO
  End If

  If Len(cARQUIVO) > 0 Then
    RPT.PrintFileName = cARQUIVO
    RPT.PrintReport
  Else
    Alert ("Arquivo Nao Preenchido")
  End If

  TxtArquivo = cARQUIVO

  If CmdEmail1.Value = 1 Then
    'servidor,porta,from,to,assunto,anexos,mensagem,enviar e sair
    ePASS01 = Array("", _
                    "", _
                    "", _
                    FixStr(TxtDestino), _
                    FixStr(TxtAssunto), _
                    cARQUIVO, _
                    FixStr(TxtTexto), _
                    False)
    FrmSendMail.Show vbModal, Me
  End If
End Sub

Private Sub TxtDestino_LostFocus()
  If Not CheckEmail(TxtDestino) Then
    TxtDestino.SetFocus
  End If
End Sub


